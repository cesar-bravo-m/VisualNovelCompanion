using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Documents;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.UI.Xaml.Shapes;
using Microsoft.UI.Xaml.Media;
using Windows.Foundation;
using Windows.Graphics;
using System.Runtime.InteropServices;

namespace VNC;

/// <summary>Represents a word definition entry for the translation table</summary>
public class WordDefinition
{
    public string Word { get; set; } = string.Empty;
    public string Pronunciation { get; set; } = string.Empty;
    public string Translation { get; set; } = string.Empty;
}

/// <summary>Represents an original text with its translation</summary>
public class OriginalTextEntry
{
    public string OriginalText { get; set; } = string.Empty;
    public string Translation { get; set; } = string.Empty;
}

/// <summary>Represents a log entry containing original text and associated word definitions</summary>
public class LogEntry
{
    public string OriginalText { get; set; } = string.Empty;
    public string Word { get; set; } = string.Empty;
    public string Pronunciation { get; set; } = string.Empty;
    public string Meaning { get; set; } = string.Empty;
    public int Frequency { get; set; } = 1;
    public DateTime Timestamp { get; set; } = DateTime.Now;
    public List<string> OriginalTexts { get; set; } = new List<string>();
    public List<OriginalTextEntry> OriginalTextEntries { get; set; } = new List<OriginalTextEntry>();
}

public sealed partial class MainWindow : Window
{
    private IntPtr _hwnd;
    private Microsoft.UI.Windowing.AppWindow _appWindow;
    private Microsoft.UI.Windowing.OverlappedPresenter _presenter;

    private DirectXManager _directXManager;
    private TransparencyManager _transparencyManager;
    private readonly ObservableCollection<WordDefinition> _wordDefinitions = new();
    private readonly ObservableCollection<LogEntry> _logEntries = new();

    private AppSettings _appSettings = new AppSettings();

    // P/Invoke declarations for screen capture
    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

    [DllImport("gdi32.dll")]
    private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

    [DllImport("gdi32.dll")]
    private static extern bool BitBlt(IntPtr hdc, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, uint dwRop);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr hObject);

    private const uint SRCCOPY = 0x00CC0020;

    public const int ICON_SMALL = 0;  
    public const int ICON_BIG = 1;  
    public const int ICON_SMALL2 = 2;  
      
    public const int WM_GETICON = 0x007F;  
    public const int WM_SETICON = 0x0080;  

    [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Auto)]  
    public static extern int SendMessage(IntPtr hWnd, uint msg, int wParam, IntPtr lParam);  

    public MainWindow()
    {
        InitializeComponent();

        // Get window handle and configure window
        _hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
        Microsoft.UI.WindowId windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(_hwnd);
        _appWindow = Microsoft.UI.Windowing.AppWindow.GetFromWindowId(windowId);

        // Configure window presenter
        _presenter = _appWindow.Presenter as Microsoft.UI.Windowing.OverlappedPresenter;
        _presenter.IsResizable = true;
        _presenter.SetBorderAndTitleBar(true, false);

        // Set initial window size and position
        // _appWindow.Resize(new Windows.Graphics.SizeInt32(800, 500));
        // _appWindow.Move(new Windows.Graphics.PointInt32(500, 300));

        // Set window corner preference for Windows 11
        SetWindowCornerPreference();

        // Initialize DirectX rendering
        InitializeDirectX();

        // Initialize transparency and click-through functionality
        InitializeTransparency();

        string sExe = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? "";  
        System.Drawing.Icon? ico = System.Drawing.Icon.ExtractAssociatedIcon(sExe);  
        if (ico != null)
            SendMessage(_hwnd, WM_SETICON, ICON_BIG, ico.Handle);  

        WordDefinitionsDataGrid.ItemsSource = _wordDefinitions;
        LogDataGrid.ItemsSource = _logEntries;

        this.Closed += MainWindow_Closed;

        _ = LoadSettingsAsync();
    }


    private async void SettingsButton_Click(object sender, RoutedEventArgs e)
    {
        var settingsDialog = new SettingsDialog()
        {
            XamlRoot = this.Content.XamlRoot
        };
        
        var result = await settingsDialog.ShowAsync();
        
        // If settings were saved, reload them
        if (result == ContentDialogResult.Primary)
        {
            await LoadSettingsAsync();
        }
    }

    private async Task LoadSettingsAsync()
    {
        try
        {
            _appSettings = await SettingsManager.LoadSettingsAsync();
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Failed to load settings: {ex.Message}";
        }
    }


    private async void LlmAnalysisButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            LlmAnalysisButton.IsEnabled = false;
            LlmAnalysisButton.Content = "Processing...";
            StatusText.Text = "Capturing screenshot from transparent area...";

            // Get the screen coordinates of the SwapChainPanel
            var panelBounds = GetSwapChainPanelScreenBounds();
            if (panelBounds.Width <= 0 || panelBounds.Height <= 0)
            {
                StatusText.Text = "Invalid capture area. Please ensure the window is visible.";
                LlmResultTextBox.Text = "Could not determine the capture area bounds.";
                return;
            }

            // Capture the screen region
            var capturedBitmap = CaptureScreenRegion(panelBounds);
            if (capturedBitmap == null)
            {
                StatusText.Text = "Failed to capture screenshot.";
                LlmResultTextBox.Text = "Screenshot capture failed. Please try again.";
                return;
            }

            StatusText.Text = "Sending image to LLM for analysis...";
            LlmResultTextBox.Text = "Processing...";

            // Reload settings to get the latest values
            _appSettings = await SettingsManager.LoadSettingsAsync();
            
            string apiKey = "";
            string modelName = "google/gemma-3n-E4B-it";
            
            // Use BYOK settings if Intelligence is set to BYOK and API key is provided
            if (_appSettings.Intelligence == "BYOK" && !string.IsNullOrWhiteSpace(_appSettings.TogetherApiKey))
            {
                apiKey = _appSettings.TogetherApiKey;
                modelName = _appSettings.SelectedTogetherModel;
            }
            
            string llmResult;
            
            // Handle different intelligence modes
            if (_appSettings.Intelligence == "managed")
            {
                // Use managed service - only IMAGE mode is supported
                StatusText.Text = "Sending image to managed service for analysis...";
                var base64Image = ConvertBitmapToBase64(capturedBitmap);
                llmResult = await ManagedClient.AnalyzeImageAsync(base64Image, "http://172.178.84.127:8080");
            }
            else
            {
                // Use existing logic for BYOK and local modes
                // Check if OCR mode is enabled
                if (_appSettings.Model == "OCR")
                {
                    StatusText.Text = "Extracting text using OCR...";
                    
                    // Check if OCR is available
                    if (!OcrService.IsOcrAvailable())
                    {
                        StatusText.Text = "OCR is not available on this system.";
                        LlmResultTextBox.Text = "OCR functionality is not available on this system. Please ensure Windows OCR features are installed.";
                        capturedBitmap.Dispose();
                        return;
                    }
                    
                    // Extract text using OCR
                    string extractedText;
                    try
                    {
                        extractedText = await OcrService.ExtractTextFromBitmapAsync(capturedBitmap);
                    }
                    catch (Exception ex)
                    {
                        StatusText.Text = $"OCR failed: {ex.Message}";
                        LlmResultTextBox.Text = $"OCR extraction failed: {ex.Message}";
                        capturedBitmap.Dispose();
                        return;
                    }
                    
                    if (string.IsNullOrWhiteSpace(extractedText))
                    {
                        StatusText.Text = "No text detected in the captured area.";
                        LlmResultTextBox.Text = "No text was detected in the captured area. Please try capturing a different region or switch to Image mode.";
                        capturedBitmap.Dispose();
                        return;
                    }
                    
                    StatusText.Text = "Sending extracted text to LLM for translation...";
                    
                    // Send extracted text to LLM
                    llmResult = await TogetherClient.AnalyzeTextAsync(
                        extractedText,
                        apiKey: apiKey,
                        modelName: modelName
                    );
                }
                else
                {
                    // Traditional image-based analysis
                    var base64Image = ConvertBitmapToBase64(capturedBitmap);
                    
                    llmResult = await TogetherClient.AnalyzeScreenshotAsync(
                        base64Image,
                        apiKey: apiKey,
                        modelName: modelName
                    );
                }
            }

            var wordDefs = ParseWordDefinitions(llmResult ?? "");
            var cleaned = CleanTranslationText(llmResult ?? "No response from LLM", wordDefs);
            var originalText = ExtractOriginalText(llmResult ?? "");

            LlmResultTextBox.Text = cleaned;

            _wordDefinitions.Clear();
            foreach (var d in wordDefs) _wordDefinitions.Add(d);

            // Add entries to log (increment frequency if word already exists)
            if (wordDefs.Count > 0 && !string.IsNullOrWhiteSpace(originalText))
            {
                foreach (var wordDef in wordDefs)
                {
                    // Check if word already exists in log
                    var existingEntry = _logEntries.FirstOrDefault(entry => 
                        string.Equals(entry.Word, wordDef.Word, StringComparison.OrdinalIgnoreCase));
                    
                    if (existingEntry != null)
                    {
                        // Increment frequency for existing word
                        existingEntry.Frequency++;
                        existingEntry.Timestamp = DateTime.Now; // Update timestamp
                        
                        // Add original text if it's not already in the list
                        if (!existingEntry.OriginalTexts.Contains(originalText, StringComparer.OrdinalIgnoreCase))
                        {
                            existingEntry.OriginalTexts.Add(originalText);
                            existingEntry.OriginalTextEntries.Add(new OriginalTextEntry
                            {
                                OriginalText = originalText,
                                Translation = cleaned
                            });
                        }
                    }
                    else
                    {
                        // Add new word to log
                        var newEntry = new LogEntry
                        {
                            OriginalText = originalText,
                            Word = wordDef.Word,
                            Pronunciation = wordDef.Pronunciation,
                            Meaning = wordDef.Translation,
                            Frequency = 1,
                            Timestamp = DateTime.Now
                        };
                        newEntry.OriginalTexts.Add(originalText);
                        newEntry.OriginalTextEntries.Add(new OriginalTextEntry
                        {
                            OriginalText = originalText,
                            Translation = cleaned
                        });
                        _logEntries.Add(newEntry);
                    }
                }
            }

            if (wordDefs.Count > 0)
            {
                WordDefinitionsHeader.Visibility = Visibility.Visible;
                WordTableScrollViewer.Visibility = Visibility.Visible;
            }
            else
            {
                WordDefinitionsHeader.Visibility = Visibility.Collapsed;
                WordTableScrollViewer.Visibility = Visibility.Collapsed;
            }

            StatusText.Text = "Translation complete.";
            capturedBitmap.Dispose();
        }
        catch (Exception ex)
        {
            var dialog = new ContentDialog()
            {
                Title = "LLM Analysis Error",
                Content = $"Failed to analyze image with LLM: {ex.Message}",
                CloseButtonText = "OK",
                XamlRoot = this.Content.XamlRoot
            };

            await dialog.ShowAsync();
            StatusText.Text = "LLM analysis failed.";
            LlmResultTextBox.Text = $"LLM analysis failed: {ex.Message}";
        }
        finally
        {
            LlmAnalysisButton.IsEnabled = true;
            LlmAnalysisButton.Content = "Translate";
        }
    }

    private string ConvertBitmapToBase64(System.Drawing.Bitmap bitmap)
    {
        using var ms = new MemoryStream();
        bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        return $"data:image/png;base64,{Convert.ToBase64String(ms.ToArray())}";
    }

    private System.Drawing.Rectangle GetSwapChainPanelScreenBounds()
    {
        try
        {
            // Get the transform of the SwapChainPanel relative to the window
            var transform = swapChainPanel1.TransformToVisual(null);
            var point = transform.TransformPoint(new Windows.Foundation.Point(0, 0));
            
            // Get window position on screen
            if (!GetWindowRect(_hwnd, out var windowRect))
            {
                return System.Drawing.Rectangle.Empty;
            }

            // Calculate the absolute screen position of the panel
            int screenX = windowRect.Left + (int)point.X;
            int screenY = windowRect.Top + (int)point.Y;
            int width = (int)swapChainPanel1.ActualWidth;
            int height = (int)swapChainPanel1.ActualHeight;

            return new System.Drawing.Rectangle(screenX, screenY, width, height);
        }
        catch
        {
            return System.Drawing.Rectangle.Empty;
        }
    }

    private System.Drawing.Bitmap? CaptureScreenRegion(System.Drawing.Rectangle bounds)
    {
        IntPtr screenDc = IntPtr.Zero;
        IntPtr memoryDc = IntPtr.Zero;
        IntPtr bitmap = IntPtr.Zero;
        IntPtr oldBitmap = IntPtr.Zero;

        try
        {
            // Get the screen DC
            screenDc = GetDC(IntPtr.Zero);
            if (screenDc == IntPtr.Zero)
                return null;

            // Create a memory DC compatible with the screen DC
            memoryDc = CreateCompatibleDC(screenDc);
            if (memoryDc == IntPtr.Zero)
                return null;

            // Create a bitmap compatible with the screen DC
            bitmap = CreateCompatibleBitmap(screenDc, bounds.Width, bounds.Height);
            if (bitmap == IntPtr.Zero)
                return null;

            // Select the bitmap into the memory DC
            oldBitmap = SelectObject(memoryDc, bitmap);

            // Copy the screen region to the bitmap
            if (!BitBlt(memoryDc, 0, 0, bounds.Width, bounds.Height, 
                       screenDc, bounds.X, bounds.Y, SRCCOPY))
                return null;

            // Create a GDI+ Bitmap from the GDI bitmap
            var result = System.Drawing.Image.FromHbitmap(bitmap);
            
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            // Clean up resources
            if (oldBitmap != IntPtr.Zero)
                SelectObject(memoryDc, oldBitmap);
            if (bitmap != IntPtr.Zero)
                DeleteObject(bitmap);
            if (memoryDc != IntPtr.Zero)
                DeleteDC(memoryDc);
            if (screenDc != IntPtr.Zero)
                ReleaseDC(IntPtr.Zero, screenDc);
        }
    }

    private void MainWindow_Closed(object sender, WindowEventArgs args)
    {
        // Clean up DirectX resources
        _directXManager?.Dispose();
    }


    #region Log DataGrid Sorting

    private void LogDataGrid_Sorting(object sender, CommunityToolkit.WinUI.UI.Controls.DataGridColumnEventArgs e)
    {
        var column = e.Column;
        var propertyName = "";
        
        // Get the property name from the column binding
        if (column.Tag is string tag)
        {
            propertyName = tag;
        }
        else if (column is CommunityToolkit.WinUI.UI.Controls.DataGridTextColumn textColumn && 
                 textColumn.Binding is Microsoft.UI.Xaml.Data.Binding binding)
        {
            propertyName = binding.Path?.Path ?? "";
        }

        if (string.IsNullOrEmpty(propertyName))
        {
            // Fallback: determine property name from column header
            switch (column.Header?.ToString())
            {
                case "Word":
                    propertyName = nameof(LogEntry.Word);
                    break;
                case "Pronunciation":
                    propertyName = nameof(LogEntry.Pronunciation);
                    break;
                case "Meaning":
                    propertyName = nameof(LogEntry.Meaning);
                    break;
                case "Frequency":
                    propertyName = nameof(LogEntry.Frequency);
                    break;
                default:
                    return; // Unknown column
            }
        }

        // Determine sort direction
        var sortDirection = column.SortDirection == null || column.SortDirection == CommunityToolkit.WinUI.UI.Controls.DataGridSortDirection.Descending
            ? CommunityToolkit.WinUI.UI.Controls.DataGridSortDirection.Ascending
            : CommunityToolkit.WinUI.UI.Controls.DataGridSortDirection.Descending;

        // Clear other column sort indicators
        foreach (var col in LogDataGrid.Columns)
        {
            if (col != column)
                col.SortDirection = null;
        }

        // Set the sort direction for the clicked column
        column.SortDirection = sortDirection;

        // Sort the collection
        SortLogEntries(propertyName, sortDirection == CommunityToolkit.WinUI.UI.Controls.DataGridSortDirection.Ascending);
    }

    private void SortLogEntries(string propertyName, bool ascending)
    {
        var sortedList = propertyName switch
        {
            nameof(LogEntry.Word) => ascending 
                ? _logEntries.OrderBy(x => x.Word, StringComparer.OrdinalIgnoreCase).ToList()
                : _logEntries.OrderByDescending(x => x.Word, StringComparer.OrdinalIgnoreCase).ToList(),
            nameof(LogEntry.Pronunciation) => ascending
                ? _logEntries.OrderBy(x => x.Pronunciation, StringComparer.OrdinalIgnoreCase).ToList()
                : _logEntries.OrderByDescending(x => x.Pronunciation, StringComparer.OrdinalIgnoreCase).ToList(),
            nameof(LogEntry.Meaning) => ascending
                ? _logEntries.OrderBy(x => x.Meaning, StringComparer.OrdinalIgnoreCase).ToList()
                : _logEntries.OrderByDescending(x => x.Meaning, StringComparer.OrdinalIgnoreCase).ToList(),
            nameof(LogEntry.Frequency) => ascending
                ? _logEntries.OrderBy(x => x.Frequency).ToList()
                : _logEntries.OrderByDescending(x => x.Frequency).ToList(),
            _ => _logEntries.ToList()
        };

        // Update the collection
        _logEntries.Clear();
        foreach (var item in sortedList)
        {
            _logEntries.Add(item);
        }
    }

    #endregion


    #region Log Row Click Handler

    private async void LogDataGrid_DoubleTapped(object sender, Microsoft.UI.Xaml.Input.DoubleTappedRoutedEventArgs e)
    {
        if (LogDataGrid.SelectedItem is LogEntry logEntry)
        {
            await ShowOriginalTextsDialog(logEntry);
        }
    }

    private async Task ShowOriginalTextsDialog(LogEntry logEntry)
    {
        var dialogContent = new StackPanel
        {
            Spacing = 15,
            MinWidth = 500,
            MaxWidth = 700
        };

        // Add each original text with its translation
        for (int i = 0; i < logEntry.OriginalTextEntries.Count; i++)
        {
            var entry = logEntry.OriginalTextEntries[i];
            
            // Container for this entry
            var entryContainer = new Border
            {
                Background = new Microsoft.UI.Xaml.Media.SolidColorBrush(Microsoft.UI.Colors.LightGray) { Opacity = 0.1 },
                CornerRadius = new CornerRadius(5),
                Padding = new Thickness(15, 15, 15, 15),
                Margin = new Thickness(0, 5, 0, 5)
            };

            var entryStack = new StackPanel { Spacing = 8 };

            // Original text with highlighted word
            var originalTextBlock = new RichTextBlock
            {
                TextWrapping = TextWrapping.Wrap,
                FontFamily = new Microsoft.UI.Xaml.Media.FontFamily("Yu Gothic UI, Meiryo UI, Segoe UI, Tahoma"),
                FontSize = 14
            };

            var paragraph = new Paragraph();
            var originalText = entry.OriginalText;
            
            // Find and highlight the word (case-insensitive)
            var wordIndex = originalText.IndexOf(logEntry.Word, StringComparison.OrdinalIgnoreCase);
            if (wordIndex >= 0)
            {
                // Text before the word
                if (wordIndex > 0)
                {
                    paragraph.Inlines.Add(new Run { Text = originalText.Substring(0, wordIndex) });
                }
                
                // The highlighted word
                var highlightedRun = new Run 
                { 
                    Text = originalText.Substring(wordIndex, logEntry.Word.Length),
                    FontWeight = Microsoft.UI.Text.FontWeights.Bold,
                    Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(Microsoft.UI.Colors.DarkBlue)
                };
                paragraph.Inlines.Add(highlightedRun);
                
                // Text after the word
                var remainingIndex = wordIndex + logEntry.Word.Length;
                if (remainingIndex < originalText.Length)
                {
                    paragraph.Inlines.Add(new Run { Text = originalText.Substring(remainingIndex) });
                }
            }
            else
            {
                // If word not found (shouldn't happen), just show the text
                paragraph.Inlines.Add(new Run { Text = originalText });
            }

            originalTextBlock.Blocks.Add(paragraph);
            entryStack.Children.Add(originalTextBlock);

            // Translation
            var translationText = new TextBlock
            {
                Text = $"{entry.Translation}",
                FontStyle = Windows.UI.Text.FontStyle.Italic,
                TextWrapping = TextWrapping.Wrap,
                FontSize = 13
            };
            entryStack.Children.Add(translationText);

            entryContainer.Child = entryStack;
            dialogContent.Children.Add(entryContainer);
        }

        var dialog = new ContentDialog()
        {
            Title = "Word in context",
            Content = new ScrollViewer 
            { 
                Content = dialogContent,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                MaxHeight = 500
            },
            CloseButtonText = "Close",
            XamlRoot = this.Content.XamlRoot
        };

        await dialog.ShowAsync();
    }

    #endregion

    #region Export Functionality

    private async void ExportCsvButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var savePicker = new Windows.Storage.Pickers.FileSavePicker();
            savePicker.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.DocumentsLibrary;
            savePicker.FileTypeChoices.Add("CSV Files", new List<string>() { ".csv" });
            savePicker.SuggestedFileName = $"vocabulary_log_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}";

            // Get the current window's handle for the file picker
            var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WinRT.Interop.InitializeWithWindow.Initialize(savePicker, hwnd);

            var file = await savePicker.PickSaveFileAsync();
            if (file != null)
            {
                await ExportToCsv(file);
                
                var dialog = new ContentDialog()
                {
                    Title = "Export Complete",
                    Content = $"Vocabulary log exported successfully to {file.Path}",
                    CloseButtonText = "OK",
                    XamlRoot = this.Content.XamlRoot
                };
                await dialog.ShowAsync();
            }
        }
        catch (Exception ex)
        {
            var dialog = new ContentDialog()
            {
                Title = "Export Error",
                Content = $"Failed to export CSV: {ex.Message}",
                CloseButtonText = "OK",
                XamlRoot = this.Content.XamlRoot
            };
            await dialog.ShowAsync();
        }
    }

    private async void ExportAnkiButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var savePicker = new Windows.Storage.Pickers.FileSavePicker();
            savePicker.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.DocumentsLibrary;
            savePicker.FileTypeChoices.Add("Text Files", new List<string>() { ".txt" });
            savePicker.SuggestedFileName = $"anki_deck_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}";

            // Get the current window's handle for the file picker
            var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WinRT.Interop.InitializeWithWindow.Initialize(savePicker, hwnd);

            var file = await savePicker.PickSaveFileAsync();
            if (file != null)
            {
                await ExportToAnki(file);
                
                var dialog = new ContentDialog()
                {
                    Title = "Export Complete",
                    Content = $"Anki deck exported successfully to {file.Path}",
                    CloseButtonText = "OK",
                    XamlRoot = this.Content.XamlRoot
                };
                await dialog.ShowAsync();
            }
        }
        catch (Exception ex)
        {
            var dialog = new ContentDialog()
            {
                Title = "Export Error",
                Content = $"Failed to export Anki deck: {ex.Message}",
                CloseButtonText = "OK",
                XamlRoot = this.Content.XamlRoot
            };
            await dialog.ShowAsync();
        }
    }

    private async Task ExportToCsv(Windows.Storage.StorageFile file)
    {
        var csvContent = new StringBuilder();
        
        // Add UTF-8 BOM for Excel compatibility
        csvContent.Append('\uFEFF');
        
        // Add header
        csvContent.AppendLine("Word,Pronunciation,Meaning,Frequency,Original Text");
        
        // Add data rows
        foreach (var entry in _logEntries)
        {
            var originalTextsJoined = string.Join("|", entry.OriginalTexts);
            var line = $"\"{EscapeCsvField(entry.Word)}\",\"{EscapeCsvField(entry.Pronunciation)}\",\"{EscapeCsvField(entry.Meaning)}\",{entry.Frequency},\"{EscapeCsvField(originalTextsJoined)}\"";
            csvContent.AppendLine(line);
        }
        
        // Convert to UTF-8 bytes with BOM
        var utf8Bytes = System.Text.Encoding.UTF8.GetBytes(csvContent.ToString());
        var buffer = Windows.Security.Cryptography.CryptographicBuffer.CreateFromByteArray(utf8Bytes);
        await Windows.Storage.FileIO.WriteBufferAsync(file, buffer);
    }

    private async Task ExportToAnki(Windows.Storage.StorageFile file)
    {
        var ankiContent = new StringBuilder();
        
        // Add UTF-8 BOM for proper Unicode support
        ankiContent.Append('\uFEFF');
        
        foreach (var entry in _logEntries)
        {
            // Front side: Just the word
            var frontSide = entry.Word;
            
            // Back side: Pronunciation + Meaning
            var backSide = $"<b>{EscapeHtml(entry.Pronunciation)}</b><br><br>{EscapeHtml(entry.Meaning)}";
            
            // Anki format with pipe separation: Front|Back
            ankiContent.AppendLine($"{frontSide}|{backSide}");
        }
        
        // Convert to UTF-8 bytes with BOM
        var utf8Bytes = System.Text.Encoding.UTF8.GetBytes(ankiContent.ToString());
        var buffer = Windows.Security.Cryptography.CryptographicBuffer.CreateFromByteArray(utf8Bytes);
        await Windows.Storage.FileIO.WriteBufferAsync(file, buffer);
    }

    private static string EscapeCsvField(string field)
    {
        if (string.IsNullOrEmpty(field))
            return "";
        
        // Escape quotes by doubling them
        return field.Replace("\"", "\"\"");
    }

    private static string EscapeHtml(string text)
    {
        if (string.IsNullOrEmpty(text))
            return "";
        
        return text.Replace("&", "&amp;")
                   .Replace("<", "&lt;")
                   .Replace(">", "&gt;")
                   .Replace("\"", "&quot;")
                   .Replace("'", "&#39;");
    }

    #endregion

    #region Transparency and DirectX Methods

    /// <summary>
    /// Sets the window corner preference for Windows 11 styling.
    /// </summary>
    private void SetWindowCornerPreference()
    {
        int cornerPreference = (int)NativeInterop.DWM_WINDOW_CORNER_PREFERENCE.DWMWCP_DEFAULT;
        NativeInterop.DwmSetWindowAttribute(
            _hwnd,
            (int)NativeInterop.DWMWINDOWATTRIBUTE.DWMWA_WINDOW_CORNER_PREFERENCE,
            ref cornerPreference,
            Marshal.SizeOf(typeof(int)));
    }

    /// <summary>
    /// Initializes DirectX resources and binds the swap chain to the SwapChainPanel.
    /// </summary>
    private void InitializeDirectX()
    {
        _directXManager = new DirectXManager();
        _directXManager.Initialize();

        // Bind swap chain to SwapChainPanel
        if (_directXManager.SwapChain != null)
        {
            var panelNative = WinRT.CastExtensions.As<NativeInterop.ISwapChainPanelNative>(swapChainPanel1);
            panelNative.SetSwapChain(_directXManager.SwapChain.NativePointer);
        }
    }

    /// <summary>
    /// Initializes window transparency and click-through functionality.
    /// </summary>
    private void InitializeTransparency()
    {
        UIElement root = (UIElement)this.Content;
        
        _transparencyManager = new TransparencyManager(
            _hwnd,
            _appWindow,
            _presenter,
            root,
            null!);

        _transparencyManager.InitializeTransparency();
    }

    #endregion

    #region Word Definition Parsing

    private static List<WordDefinition> ParseWordDefinitions(string llmResponse)
    {
        var list = new List<WordDefinition>();
        if (string.IsNullOrWhiteSpace(llmResponse)) return list;

        var pattern = @"^([^:\n]+):\s*\(([^)]+)\)\s*(.+)$";
        var regex = new Regex(pattern, RegexOptions.Multiline | RegexOptions.IgnoreCase);

        foreach (Match m in regex.Matches(llmResponse))
        {
            if (m.Groups.Count >= 4)
            {
                var word = m.Groups[1].Value.Trim();
                var pron = m.Groups[2].Value.Trim();
                var trans = m.Groups[3].Value.Trim();

                if (!string.IsNullOrWhiteSpace(word) &&
                    !string.IsNullOrWhiteSpace(pron) &&
                    !string.IsNullOrWhiteSpace(trans))
                {
                    list.Add(new WordDefinition
                    {
                        Word = word,
                        Pronunciation = pron,
                        Translation = trans
                    });
                }
            }
        }
        return list;
    }

    private static string ExtractOriginalText(string llmResponse)
    {
        if (string.IsNullOrWhiteSpace(llmResponse))
            return string.Empty;

        var match = Regex.Match(llmResponse, @"<original_text>(.*?)</original_text>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        return match.Success ? match.Groups[1].Value.Trim() : string.Empty;
    }

    private static string CleanTranslationText(string llmResponse, List<WordDefinition> defs)
    {
        if (string.IsNullOrWhiteSpace(llmResponse))
            return llmResponse ?? string.Empty;

        var text = llmResponse;

        // Remove content between <original_text> and </original_text> tags
        text = Regex.Replace(text, @"<original_text>.*?</original_text>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);

        // Remove word definitions if any exist
        if (defs.Count > 0)
        {
            foreach (var d in defs)
            {
                var line = $@"^{Regex.Escape(d.Word)}:\s*\({Regex.Escape(d.Pronunciation)}\)\s*{Regex.Escape(d.Translation)}\s*$";
                text = Regex.Replace(text, line, "", RegexOptions.Multiline);
            }
        }

        // Clean up extra whitespace and newlines
        text = Regex.Replace(text, @"\n\s*\n\s*\n", "\n\n").Trim();
        text = text.Replace("</original_text>", "");
        text = text.Replace("<original_text>", "");
        return text;
    }

    #endregion
}
