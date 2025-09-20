using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

namespace VNC;

/// <summary>
/// Service for performing OCR (Optical Character Recognition) using Windows built-in OCR capabilities
/// Supports multiple languages including Japanese, Chinese, Korean, and other non-ASCII text
/// </summary>
public static class OcrService
{
    private static OcrEngine? _ocrEngine;
    private static readonly object _engineLock = new object();

    /// <summary>
    /// Initializes the OCR engine with the best available language
    /// </summary>
    private static Task<OcrEngine> GetOcrEngineAsync()
    {
        if (_ocrEngine != null)
            return Task.FromResult(_ocrEngine);

        lock (_engineLock)
        {
            if (_ocrEngine != null)
                return Task.FromResult(_ocrEngine);

            var availableLanguages = OcrEngine.AvailableRecognizerLanguages;

            var preferredLanguages = new[]
            {
                "ja",
                "zh-Hans",
                "zh-Hant",
                "ko",
                "en",
                "auto"
            };

            Windows.Globalization.Language? selectedLanguage = null;

            foreach (var preferredLang in preferredLanguages)
            {
                selectedLanguage = availableLanguages.FirstOrDefault(lang => 
                    lang.LanguageTag.StartsWith(preferredLang, StringComparison.OrdinalIgnoreCase));
                if (selectedLanguage != null)
                    break;
            }

            if (selectedLanguage == null)
            {
                selectedLanguage = availableLanguages.FirstOrDefault(lang => 
                    lang.LanguageTag.StartsWith("en", StringComparison.OrdinalIgnoreCase)) 
                    ?? availableLanguages.FirstOrDefault();
            }

            if (selectedLanguage != null)
            {
                _ocrEngine = OcrEngine.TryCreateFromLanguage(selectedLanguage);
            }

            if (_ocrEngine == null)
            {
                _ocrEngine = OcrEngine.TryCreateFromUserProfileLanguages();
            }

            if (_ocrEngine == null)
            {
                throw new InvalidOperationException("Failed to create OCR engine. OCR may not be available on this system.");
            }

            return Task.FromResult(_ocrEngine);
        }
    }

    /// <summary>
    /// Extracts text from a bitmap using Windows OCR or Ollama based on settings
    /// </summary>
    /// <param name="bitmap">The bitmap to extract text from</param>
    /// <param name="settings">Application settings to determine OCR method</param>
    /// <returns>Extracted text, or empty string if no text found</returns>
    public static async Task<string> ExtractTextFromBitmapAsync(Bitmap bitmap, AppSettings? settings = null)
    {
        if (bitmap == null)
            throw new ArgumentNullException(nameof(bitmap));

        // Load settings if not provided
        settings ??= await SettingsManager.LoadSettingsAsync();

        // Use Ollama for local intelligence, but respect the input mode
        if (settings.Intelligence == "local" && !string.IsNullOrEmpty(settings.SelectedOllamaModel))
        {
            if (settings.Model == "Image")
            {
                // Use Ollama with direct image input
                return await ExtractTextUsingOllamaAsync(bitmap, settings);
            }
            else if (settings.Model == "OCR")
            {
                // Use Windows OCR first, then send text to Ollama for processing
                return await ExtractTextUsingOllamaWithOcrAsync(bitmap, settings);
            }
        }

        // Fall back to Windows OCR
        return await ExtractTextUsingWindowsOcrAsync(bitmap);
    }

    /// <summary>
    /// Extracts text from a bitmap using Windows OCR
    /// </summary>
    /// <param name="bitmap">The bitmap to extract text from</param>
    /// <returns>Extracted text, or empty string if no text found</returns>
    public static async Task<string> ExtractTextUsingWindowsOcrAsync(Bitmap bitmap)
    {
        if (bitmap == null)
            throw new ArgumentNullException(nameof(bitmap));

        try
        {
            // Convert System.Drawing.Bitmap to Windows.Graphics.Imaging.SoftwareBitmap
            var softwareBitmap = await ConvertToSoftwareBitmapAsync(bitmap);
            
            // Get OCR engine
            var ocrEngine = await GetOcrEngineAsync();
            
            // Perform OCR
            var ocrResult = await ocrEngine.RecognizeAsync(softwareBitmap);
            
            // Extract text from result
            var extractedText = ExtractTextFromOcrResult(ocrResult);
            
            return extractedText;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Windows OCR failed: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Extracts text from a bitmap using Ollama
    /// </summary>
    /// <param name="bitmap">The bitmap to extract text from</param>
    /// <param name="settings">Application settings containing Ollama configuration</param>
    /// <returns>Extracted text, or empty string if no text found</returns>
    public static async Task<string> ExtractTextUsingOllamaAsync(Bitmap bitmap, AppSettings settings)
    {
        if (bitmap == null)
            throw new ArgumentNullException(nameof(bitmap));
        
        if (settings == null)
            throw new ArgumentNullException(nameof(settings));

        if (string.IsNullOrEmpty(settings.SelectedOllamaModel))
            throw new ArgumentException("No Ollama model selected in settings", nameof(settings));

        try
        {
            using var ollamaClient = new OllamaClient(settings.OllamaEndpoint);
            
            // Check if Ollama is available
            var isAvailable = await ollamaClient.IsAvailableAsync();
            if (!isAvailable)
            {
                throw new InvalidOperationException($"Ollama is not available at {settings.OllamaEndpoint}");
            }

            // Extract text using Ollama
            var extractedText = await ollamaClient.ExtractTextFromImageAsync(bitmap, settings.SelectedOllamaModel);
            
            return extractedText;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Ollama OCR failed: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Extracts text from a bitmap using Windows OCR, then processes it with Ollama
    /// </summary>
    /// <param name="bitmap">The bitmap to extract text from</param>
    /// <param name="settings">Application settings containing Ollama configuration</param>
    /// <returns>Processed text from Ollama</returns>
    public static async Task<string> ExtractTextUsingOllamaWithOcrAsync(Bitmap bitmap, AppSettings settings)
    {
        if (bitmap == null)
            throw new ArgumentNullException(nameof(bitmap));
        
        if (settings == null)
            throw new ArgumentNullException(nameof(settings));

        if (string.IsNullOrEmpty(settings.SelectedOllamaModel))
            throw new ArgumentException("No Ollama model selected in settings", nameof(settings));

        try
        {
            // Step 1: Extract text using Windows OCR
            var extractedText = await ExtractTextUsingWindowsOcrAsync(bitmap);
            
            if (string.IsNullOrWhiteSpace(extractedText))
            {
                return "No text found in image";
            }

            // Step 2: Process the extracted text with Ollama
            using var ollamaClient = new OllamaClient(settings.OllamaEndpoint);
            
            // Check if Ollama is available
            var isAvailable = await ollamaClient.IsAvailableAsync();
            if (!isAvailable)
            {
                // Fall back to just the OCR text if Ollama is not available
                return extractedText;
            }

            // Process the text with Ollama for translation/enhancement
            var processedText = await ollamaClient.ProcessTextAsync(extractedText, settings.SelectedOllamaModel);
            
            return processedText;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Ollama OCR with text processing failed: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Converts a System.Drawing.Bitmap to Windows.Graphics.Imaging.SoftwareBitmap
    /// </summary>
    private static async Task<SoftwareBitmap> ConvertToSoftwareBitmapAsync(Bitmap bitmap)
    {
        using var memoryStream = new MemoryStream();
        
        // Save bitmap to memory stream as PNG to preserve quality
        bitmap.Save(memoryStream, ImageFormat.Png);
        memoryStream.Position = 0;

        // Create IRandomAccessStream from memory stream
        var randomAccessStream = memoryStream.AsRandomAccessStream();
        
        // Create decoder
        var decoder = await BitmapDecoder.CreateAsync(randomAccessStream);
        
        // Get software bitmap
        var softwareBitmap = await decoder.GetSoftwareBitmapAsync();
        
        // Convert to BGRA8 format if needed (required for OCR)
        if (softwareBitmap.BitmapPixelFormat != BitmapPixelFormat.Bgra8 ||
            softwareBitmap.BitmapAlphaMode != BitmapAlphaMode.Premultiplied)
        {
            softwareBitmap = SoftwareBitmap.Convert(softwareBitmap, BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
        }

        return softwareBitmap;
    }

    /// <summary>
    /// Extracts text from OCR result, preserving layout and structure
    /// </summary>
    private static string ExtractTextFromOcrResult(OcrResult ocrResult)
    {
        if (ocrResult?.Lines == null || !ocrResult.Lines.Any())
            return string.Empty;

        var textBuilder = new StringBuilder();
        
        foreach (var line in ocrResult.Lines)
        {
            if (line.Words != null && line.Words.Any())
            {
                // Combine words in the line with spaces
                var lineText = string.Join(" ", line.Words.Select(word => word.Text));
                textBuilder.AppendLine(lineText);
            }
        }

        return textBuilder.ToString().Trim();
    }

    /// <summary>
    /// Gets information about available OCR languages
    /// </summary>
    /// <returns>List of available language tags</returns>
    public static List<string> GetAvailableLanguages()
    {
        try
        {
            var availableLanguages = OcrEngine.AvailableRecognizerLanguages;
            return availableLanguages.Select(lang => $"{lang.LanguageTag} ({lang.DisplayName})").ToList();
        }
        catch
        {
            return new List<string> { "Unable to retrieve available languages" };
        }
    }

    /// <summary>
    /// Checks if OCR is available on the current system
    /// </summary>
    /// <returns>True if OCR is available, false otherwise</returns>
    public static bool IsOcrAvailable()
    {
        try
        {
            var availableLanguages = OcrEngine.AvailableRecognizerLanguages;
            return availableLanguages.Any();
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Gets detailed OCR result with word positions and confidence (for debugging)
    /// This method always uses Windows OCR for detailed analysis
    /// </summary>
    /// <param name="bitmap">The bitmap to analyze</param>
    /// <returns>Detailed OCR information</returns>
    public static async Task<OcrAnalysisResult> AnalyzeBitmapAsync(Bitmap bitmap)
    {
        if (bitmap == null)
            throw new ArgumentNullException(nameof(bitmap));

        try
        {
            var softwareBitmap = await ConvertToSoftwareBitmapAsync(bitmap);
            var ocrEngine = await GetOcrEngineAsync();
            var ocrResult = await ocrEngine.RecognizeAsync(softwareBitmap);

            var result = new OcrAnalysisResult
            {
                Text = ExtractTextFromOcrResult(ocrResult),
                TextAngle = ocrResult.TextAngle ?? 0,
                Lines = new List<OcrLineInfo>()
            };

            if (ocrResult.Lines != null)
            {
                foreach (var line in ocrResult.Lines)
                {
                    var lineInfo = new OcrLineInfo
                    {
                        Text = string.Join(" ", line.Words?.Select(w => w.Text) ?? Enumerable.Empty<string>()),
                        Words = new List<OcrWordInfo>()
                    };

                    if (line.Words != null)
                    {
                        foreach (var word in line.Words)
                        {
                            lineInfo.Words.Add(new OcrWordInfo
                            {
                                Text = word.Text,
                                BoundingRect = new System.Drawing.Rectangle(
                                    (int)word.BoundingRect.X,
                                    (int)word.BoundingRect.Y,
                                    (int)word.BoundingRect.Width,
                                    (int)word.BoundingRect.Height)
                            });
                        }
                    }

                    result.Lines.Add(lineInfo);
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"OCR analysis failed: {ex.Message}", ex);
        }
    }
}

/// <summary>
/// Detailed OCR analysis result
/// </summary>
public class OcrAnalysisResult
{
    public string Text { get; set; } = string.Empty;
    public double TextAngle { get; set; }
    public List<OcrLineInfo> Lines { get; set; } = new List<OcrLineInfo>();
}

/// <summary>
/// Information about a line of text found by OCR
/// </summary>
public class OcrLineInfo
{
    public string Text { get; set; } = string.Empty;
    public List<OcrWordInfo> Words { get; set; } = new List<OcrWordInfo>();
}

/// <summary>
/// Information about a word found by OCR
/// </summary>
public class OcrWordInfo
{
    public string Text { get; set; } = string.Empty;
    public System.Drawing.Rectangle BoundingRect { get; set; }
}
