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
    private SizeInt32 _prevWindowSize;
    private bool _hasPrevWindowSize = false;
    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(nint hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(nint hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, nint dwExtraInfo);

    [DllImport("user32.dll")]
    private static extern nint GetForegroundWindow();

    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;

    private static int GetWindowOuterHeight(nint hwnd)
    {
        if (hwnd == nint.Zero) return 0;
        if (!GetWindowRect(hwnd, out var rc)) return 0;
        return Math.Max(0, rc.Bottom - rc.Top);
    }
    private readonly ObservableCollection<WindowInfo> _availableWindows = new();
    private readonly ObservableCollection<WordDefinition> _wordDefinitions = new();
    private readonly ObservableCollection<LogEntry> _logEntries = new();

    private WindowInfo? _selectedWindow;
    private DispatcherTimer? _windowMonitorTimer;
    private DispatcherTimer? _captureTimer;

    private System.Drawing.Bitmap? _currentScreenshotBitmap;
    private bool _isCapturing = false;

    private bool _isDrawingRectangle = false;
    private bool _isResizingRectangle = false;
    private Point _startPoint;
    private string _currentHandle = "";
    private Rect _selectionRect = Rect.Empty;
    private bool _hasSelection = false;
    private bool _hasDefaultRectangle = false;

    private AppSettings _appSettings = new AppSettings();

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

        IntPtr hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);  
        string sExe = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? "";  
        System.Drawing.Icon? ico = System.Drawing.Icon.ExtractAssociatedIcon(sExe);  
        if (ico != null)
            SendMessage(hWnd, WM_SETICON, ICON_BIG, ico.Handle);  

        WindowsComboBox.ItemsSource = _availableWindows;
        WordDefinitionsDataGrid.ItemsSource = _wordDefinitions;
        LogDataGrid.ItemsSource = _logEntries;

        _windowMonitorTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
        _windowMonitorTimer.Tick += WindowMonitorTimer_Tick;

        _captureTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
        _captureTimer.Tick += CaptureTimer_Tick;

        this.Closed += MainWindow_Closed;


        _ = LoadAvailableWindowsAsync();
        _ = LoadSettingsAsync();
    }

    #region Collapse/Expand
    private void ApplyCollapsed(bool collapsed)
    {
        if (collapsed)
        {
            try
            {
                _prevWindowSize = this.AppWindow.Size;
                _hasPrevWindowSize = true;
            } catch { }
            int targetHeight = 0;
            if (_selectedWindow != null)
            {
                targetHeight = GetWindowOuterHeight(_selectedWindow.Handle);
            }
            if (targetHeight <= 0)
            {
                targetHeight = this.AppWindow.Size.Height;
            }
            try
            {
                const int minHeight = 200;
                targetHeight = Math.Max(minHeight, targetHeight);
                this.AppWindow.Resize(new SizeInt32(820, targetHeight));
            } catch { }
            ScreenshotScrollViewer.Visibility = Visibility.Collapsed;
            SplitterBorder.Visibility = Visibility.Collapsed;
            ScreenshotColumn.Width = new GridLength(0);
            SplitterColumn.Width   = new GridLength(0);
            LlmColumn.Width        = new GridLength(1, GridUnitType.Star);

            Grid.SetColumn(LlmPanel, 0);
            Grid.SetColumnSpan(LlmPanel, 3);
            LlmPanel.HorizontalAlignment = HorizontalAlignment.Stretch;
            LlmPanel.Width = double.NaN;
            ExpandButton.Visibility = Visibility.Visible;
            CollapseButton.Visibility = Visibility.Collapsed;
        }
        else
        {
            try
            {
                if (_hasPrevWindowSize && _prevWindowSize.Width > 0 && _prevWindowSize.Height > 0)
                {
                    this.AppWindow.Resize(_prevWindowSize);
                }
            } catch
            {
            }
            ScreenshotScrollViewer.Visibility = Visibility.Visible;
            SplitterBorder.Visibility = Visibility.Collapsed; // or Visible if you want the line
            ScreenshotColumn.Width = new GridLength(3, GridUnitType.Star);
            SplitterColumn.Width   = new GridLength(5);
            LlmColumn.Width        = new GridLength(1, GridUnitType.Star);

            Grid.SetColumn(LlmPanel, 2);
            Grid.SetColumnSpan(LlmPanel, 1);
            LlmPanel.HorizontalAlignment = HorizontalAlignment.Stretch;
            LlmPanel.Width = double.NaN;
            ExpandButton.Visibility = Visibility.Collapsed;
            CollapseButton.Visibility = Visibility.Visible;
        }
    }
    private void CollapseButton_Click(object sender, RoutedEventArgs e)
    {
        ApplyCollapsed(true);
    }

    private void ExpandButton_Click(object sender, RoutedEventArgs e)
    {
        ApplyCollapsed(false);
    }

    #endregion

    private async void RefreshWindowsButton_Click(object sender, RoutedEventArgs e)
    {
        await LoadAvailableWindowsAsync();
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

    private void StopCaptureButton_Click(object sender, RoutedEventArgs e)
    {
        StopCapture();
    }

    private async void WindowsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isCapturing) StopCapture();

        if (WindowsComboBox.SelectedItem is WindowInfo selectedWindow)
        {
            if (!ScreenshotCapture.IsWindowValid(selectedWindow.Handle))
            {
                await HandleWindowClosed();
                return;
            }

            _selectedWindow = selectedWindow;
            StatusText.Text = $"Selected window: '{selectedWindow.Title}'. Starting capture...";

            if (_windowMonitorTimer is { IsEnabled: false })
                _windowMonitorTimer.Start();

            // Automatically start capture when window is selected
            try
            {
                StartCapture();
            }
            catch (Exception ex)
            {
                var dialog = new ContentDialog()
                {
                    Title = "Capture Error",
                    Content = $"Failed to start capture: {ex.Message}",
                    CloseButtonText = "OK",
                    XamlRoot = this.Content.XamlRoot
                };

                await dialog.ShowAsync();
                StatusText.Text = $"Failed to start capture for '{selectedWindow.Title}'.";
            }
        }
        else
        {
            _selectedWindow = null;
            StatusText.Text = "No window selected. Please select a window to start capture.";
            _windowMonitorTimer?.Stop();
        }
    }

    private async Task LoadAvailableWindowsAsync()
    {
        try
        {
            RefreshWindowsButton.IsEnabled = false;
            RefreshWindowsButton.Content = "...";
            StatusText.Text = "Loading available windows...";

            var windows = await ScreenshotCapture.GetVisibleWindowsWithoutThumbnailsAsync();

            _availableWindows.Clear();
            foreach (var w in windows.Where(w => w.IsValid))
                _availableWindows.Add(w);

            StatusText.Text = _availableWindows.Count == 0
                ? "No windows found. Try opening some applications and refresh again."
                : $"Found {_availableWindows.Count} window(s). Select one to start capture automatically.";
        }
        catch (Exception ex)
        {
            var dialog = new ContentDialog()
            {
                Title = "Error Loading Windows",
                Content = $"Failed to load available windows: {ex.Message}",
                CloseButtonText = "OK",
                XamlRoot = this.Content.XamlRoot
            };
            await dialog.ShowAsync();
            StatusText.Text = "Failed to load windows. Please try again.";
        }
        finally
        {
            RefreshWindowsButton.IsEnabled = true;
            RefreshWindowsButton.Content = "↻";
        }
    }

    private async void WindowMonitorTimer_Tick(object? sender, object e)
    {
        if (_selectedWindow != null &&
            !ScreenshotCapture.IsWindowValid(_selectedWindow.Handle))
        {
            await HandleWindowClosed();
        }
    }

    private async void CaptureTimer_Tick(object? sender, object e)
    {
        if (_selectedWindow == null || !_isCapturing) return;

        try
        {
            if (!ScreenshotCapture.IsWindowValid(_selectedWindow.Handle))
            {
                await HandleWindowClosed();
                return;
            }

            var screenshotResult = await ScreenshotCapture.CaptureWindowWithBitmapAsync(_selectedWindow.Handle);

            if (ScreenshotCapture.IsBitmapMostlyWhite(screenshotResult.Bitmap))
            {
                screenshotResult.Bitmap.Dispose();
                await Task.Delay(50);
                return;
            }

            ScreenshotImage.Source = screenshotResult.BitmapImage;

            _currentScreenshotBitmap?.Dispose();
            _currentScreenshotBitmap = screenshotResult.Bitmap;

            UpdateSelectionCanvasSize();

            StatusText.Text = $"Capturing '{_selectedWindow.Title}'";

            LlmAnalysisButton.IsEnabled = true;
        }
        catch (Exception ex)
        {
            StopCapture();
            StatusText.Text = $"Capture stopped due to error: {ex.Message}";
        }
    }

    private void StartCapture()
    {
        if (_selectedWindow == null || _isCapturing) return;
        SelectWindowTextBlock.Visibility = Visibility.Collapsed;
        WindowsComboBox.Visibility = Visibility.Collapsed;
        RefreshWindowsButton.Visibility = Visibility.Collapsed;
        LlmAnalysisButton.IsEnabled = true;

        _isCapturing = true;
        _captureTimer?.Start();

        // Show the right panel; let the current visual state decide its layout
        LlmPanel.Visibility = Visibility.Visible;

        // Show the Stop button
        StopCaptureButton.Visibility = Visibility.Visible;
        StopCaptureButton.IsEnabled = true;
        StatusText.Text = $"Started auto-capturing from '{_selectedWindow.Title}'.";
    }

    private void StopCapture()
    {
        if (!_isCapturing) return;
        SelectWindowTextBlock.Visibility = Visibility.Visible;
        WindowsComboBox.Visibility = Visibility.Visible;
        RefreshWindowsButton.Visibility = Visibility.Visible;
        LlmAnalysisButton.IsEnabled = false;

        _isCapturing = false;
        _captureTimer?.Stop();

        // Hide LLM panel entirely when not capturing
        LlmPanel.Visibility = Visibility.Collapsed;

        // Hide the Stop button
        StopCaptureButton.Visibility = Visibility.Collapsed;
        StopCaptureButton.IsEnabled = false;
        StatusText.Text = _selectedWindow != null
            ? $"Stopped auto-capturing from '{_selectedWindow.Title}'."
            : "Auto-capture stopped.";
    }

    private async Task HandleWindowClosed()
    {
        _windowMonitorTimer?.Stop();
        StopCapture();

        var closedTitle = _selectedWindow?.Title ?? "Unknown";
        _selectedWindow = null;
        WindowsComboBox.SelectedItem = null;
        ScreenshotImage.Source = null;
        LlmAnalysisButton.IsEnabled = false;
        LlmResultTextBox.Text = string.Empty;

        _wordDefinitions.Clear();
        _logEntries.Clear();
        WordDefinitionsHeader.Visibility = Visibility.Collapsed;
        WordTableScrollViewer.Visibility = Visibility.Collapsed;

        // Reset rectangle state but don't hide it completely
        _hasSelection = false;
        _hasDefaultRectangle = false;
        HideSelection();

        _currentScreenshotBitmap?.Dispose();
        _currentScreenshotBitmap = null;

        StatusText.Text = $"The selected window '{closedTitle}' has been closed. Please select a different window.";

        var dialog = new ContentDialog()
        {
            Title = "Window Closed",
            Content = $"The selected window '{closedTitle}' has been closed. Please select a different window.",
            CloseButtonText = "OK",
            XamlRoot = this.Content.XamlRoot
        };
        await dialog.ShowAsync();
    }

    private async void LlmAnalysisButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentScreenshotBitmap == null)
        {
            StatusText.Text = "No screenshot available for LLM analysis. Please take a screenshot first.";
            return;
        }

        try
        {
            LlmAnalysisButton.IsEnabled = false;
            LlmAnalysisButton.Content = "Processing...";
            StatusText.Text = "Sending image to LLM for analysis...";

            LlmResultTextBox.Text = "Processing...";

            var bitmapToAnalyze = CropBitmapToSelection();
            if (bitmapToAnalyze == null)
            {
                StatusText.Text = "Invalid selection area for LLM analysis.";
                return;
            }

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
                var base64Image = ConvertBitmapToBase64(bitmapToAnalyze);
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
                        return;
                    }
                    
                    // Extract text using OCR
                    string extractedText;
                    try
                    {
                        extractedText = await OcrService.ExtractTextFromBitmapAsync(bitmapToAnalyze);
                    }
                    catch (Exception ex)
                    {
                        StatusText.Text = $"OCR failed: {ex.Message}";
                        LlmResultTextBox.Text = $"OCR extraction failed: {ex.Message}";
                        return;
                    }
                    
                    if (string.IsNullOrWhiteSpace(extractedText))
                    {
                        StatusText.Text = "No text detected in the selected area.";
                        LlmResultTextBox.Text = "No text was detected in the selected area. Please try selecting a different region or switch to Image mode.";
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
                    var base64Image = ConvertBitmapToBase64(bitmapToAnalyze);
                    
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

            if (bitmapToAnalyze != _currentScreenshotBitmap) bitmapToAnalyze.Dispose();
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

    private void MainWindow_Closed(object sender, WindowEventArgs args)
    {
        _windowMonitorTimer?.Stop();
        _windowMonitorTimer = null;

        _captureTimer?.Stop();
        _captureTimer = null;

        _currentScreenshotBitmap?.Dispose();
        _currentScreenshotBitmap = null;
    }

    #region Rectangle Selection

    private void UpdateSelectionCanvasSize()
    {
        if (ScreenshotImage.Source is BitmapImage)
        {
            ScreenshotImage.SizeChanged -= ScreenshotImage_SizeChanged;
            ScreenshotImage.SizeChanged += ScreenshotImage_SizeChanged;
            UpdateCanvasSizeToImage();
        }
    }

    private void ScreenshotImage_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        UpdateCanvasSizeToImage();
    }

    private void UpdateCanvasSizeToImage()
    {
        if (ScreenshotImage.ActualWidth > 0 && ScreenshotImage.ActualHeight > 0)
        {
            SelectionCanvas.Width = ScreenshotImage.ActualWidth;
            SelectionCanvas.Height = ScreenshotImage.ActualHeight;

            // Always show default rectangle if no custom selection exists
            if (!_hasSelection && !_hasDefaultRectangle)
            {
                CreateDefaultRectangle();
            }
            else if (_hasSelection || _hasDefaultRectangle)
            {
                ShowSelectionRectangle();
            }
        }
    }

    private void CreateDefaultRectangle()
    {
        if (SelectionCanvas.Width <= 0 || SelectionCanvas.Height <= 0) return;

        // Create rectangle in lower third of screen with padding (visual novel dialogue position)
        var padding = 20.0;
        var rectHeight = SelectionCanvas.Height / 3.0 - padding; // One third minus padding
        var rectWidth = SelectionCanvas.Width - (2 * padding);
        var rectX = padding;
        var rectY = SelectionCanvas.Height - rectHeight - padding; // Position in lower third

        _selectionRect = new Rect(rectX, rectY, rectWidth, rectHeight);
        _hasDefaultRectangle = true;
        _hasSelection = true; // Treat default rectangle as a selection
        
        ShowSelectionRectangle();
        StatusText.Text = "Default selection area created. You can resize or move it as needed.";
    }

    private void SelectionCanvas_PointerPressed(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        if (sender is not Canvas canvas) return;

        var pos = e.GetCurrentPoint(canvas).Position;

        if (_hasSelection && IsPointInRectangle(pos, _selectionRect))
        {
            // Allow moving/dragging the existing rectangle
            _isDrawingRectangle = false;
            _isResizingRectangle = false;
            _startPoint = pos;
            canvas.CapturePointer(e.Pointer);
        }
        else
        {
            // Create new rectangle (replacing the existing one)
            _isDrawingRectangle = true;
            _isResizingRectangle = false;
            _startPoint = pos;
            _selectionRect = new Rect(pos.X, pos.Y, 0, 0);
            _hasDefaultRectangle = false; // User is creating custom rectangle

            SelectionRectangle.Visibility = Visibility.Visible;
            Canvas.SetLeft(SelectionRectangle, pos.X);
            Canvas.SetTop(SelectionRectangle, pos.Y);
            SelectionRectangle.Width = 0;
            SelectionRectangle.Height = 0;

            canvas.CapturePointer(e.Pointer);
        }
    }

    private void SelectionCanvas_PointerMoved(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        if (sender is not Canvas canvas) return;
        var p = e.GetCurrentPoint(canvas).Position;

        if (_isDrawingRectangle)
        {
            var left = Math.Min(_startPoint.X, p.X);
            var top = Math.Min(_startPoint.Y, p.Y);
            var width = Math.Abs(p.X - _startPoint.X);
            var height = Math.Abs(p.Y - _startPoint.Y);

            _selectionRect = new Rect(left, top, width, height);

            Canvas.SetLeft(SelectionRectangle, left);
            Canvas.SetTop(SelectionRectangle, top);
            SelectionRectangle.Width = width;
            SelectionRectangle.Height = height;
        }
        else if (_isResizingRectangle)
        {
            ResizeRectangle(p);
        }
    }

    private void SelectionCanvas_PointerReleased(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        if (sender is not Canvas canvas) return;

        if (_isDrawingRectangle)
        {
            _isDrawingRectangle = false;
            _hasSelection = _selectionRect.Width > 5 && _selectionRect.Height > 5;

            if (_hasSelection)
            {
                ShowSelectionHandles();
                StatusText.Text = "Selection area defined. LLM analysis will analyze only the selected region.";
            }
            else
            {
                // If rectangle is too small, revert to default rectangle instead of hiding
                CreateDefaultRectangle();
            }
        }

        _isResizingRectangle = false;
        canvas.ReleasePointerCapture(e.Pointer);
    }

    private void Handle_PointerPressed(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        if (sender is Ellipse handle && handle.Tag is string tag)
        {
            _isResizingRectangle = true;
            _currentHandle = tag;
            _startPoint = e.GetCurrentPoint(SelectionCanvas).Position;
            SelectionCanvas.CapturePointer(e.Pointer);
            e.Handled = true;
        }
    }

    private void ResizeRectangle(Point currentPoint)
    {
        var dx = currentPoint.X - _startPoint.X;
        var dy = currentPoint.Y - _startPoint.Y;

        var newRect = _selectionRect;

        switch (_currentHandle)
        {
            case "TopLeft":
                newRect = new Rect(
                    Math.Max(0, _selectionRect.X + dx),
                    Math.Max(0, _selectionRect.Y + dy),
                    Math.Max(10, _selectionRect.Width - dx),
                    Math.Max(10, _selectionRect.Height - dy));
                break;
            case "TopRight":
                newRect = new Rect(
                    _selectionRect.X,
                    Math.Max(0, _selectionRect.Y + dy),
                    Math.Max(10, _selectionRect.Width + dx),
                    Math.Max(10, _selectionRect.Height - dy));
                break;
            case "BottomLeft":
                newRect = new Rect(
                    Math.Max(0, _selectionRect.X + dx),
                    _selectionRect.Y,
                    Math.Max(10, _selectionRect.Width - dx),
                    Math.Max(10, _selectionRect.Height + dy));
                break;
            case "BottomRight":
                newRect = new Rect(
                    _selectionRect.X,
                    _selectionRect.Y,
                    Math.Max(10, _selectionRect.Width + dx),
                    Math.Max(10, _selectionRect.Height + dy));
                break;
        }

        if (newRect.Right <= SelectionCanvas.Width && newRect.Bottom <= SelectionCanvas.Height)
        {
            _selectionRect = newRect;
            _startPoint = currentPoint;
            ShowSelectionRectangle();
        }
    }

    private void ShowSelectionRectangle()
    {
        Canvas.SetLeft(SelectionRectangle, _selectionRect.X);
        Canvas.SetTop(SelectionRectangle, _selectionRect.Y);
        SelectionRectangle.Width = _selectionRect.Width;
        SelectionRectangle.Height = _selectionRect.Height;
        SelectionRectangle.Visibility = Visibility.Visible;

        ShowSelectionHandles();
    }

    private void ShowSelectionHandles()
    {
        Canvas.SetLeft(TopLeftHandle, _selectionRect.X - 5);
        Canvas.SetTop(TopLeftHandle, _selectionRect.Y - 5);
        TopLeftHandle.Visibility = Visibility.Visible;

        Canvas.SetLeft(TopRightHandle, _selectionRect.Right - 5);
        Canvas.SetTop(TopRightHandle, _selectionRect.Y - 5);
        TopRightHandle.Visibility = Visibility.Visible;

        Canvas.SetLeft(BottomLeftHandle, _selectionRect.X - 5);
        Canvas.SetTop(BottomLeftHandle, _selectionRect.Bottom - 5);
        BottomLeftHandle.Visibility = Visibility.Visible;

        Canvas.SetLeft(BottomRightHandle, _selectionRect.Right - 5);
        Canvas.SetTop(BottomRightHandle, _selectionRect.Bottom - 5);
        BottomRightHandle.Visibility = Visibility.Visible;
    }

    private void HideSelection()
    {
        // Never completely hide selection - always maintain a default rectangle
        if (!_hasDefaultRectangle)
        {
            CreateDefaultRectangle();
        }
        else
        {
            // Just hide handles but keep rectangle visible
            TopLeftHandle.Visibility = Visibility.Collapsed;
            TopRightHandle.Visibility = Visibility.Collapsed;
            BottomLeftHandle.Visibility = Visibility.Collapsed;
            BottomRightHandle.Visibility = Visibility.Collapsed;
        }
    }

    private static bool IsPointInRectangle(Point p, Rect r)
        => p.X >= r.X && p.X <= r.Right && p.Y >= r.Y && p.Y <= r.Bottom;

    private System.Drawing.Bitmap? CropBitmapToSelection()
    {
        if (_currentScreenshotBitmap == null || !_hasSelection)
            return _currentScreenshotBitmap;

        try
        {
            var scaleX = (double)_currentScreenshotBitmap.Width / SelectionCanvas.Width;
            var scaleY = (double)_currentScreenshotBitmap.Height / SelectionCanvas.Height;

            var cropX = (int)(_selectionRect.X * scaleX);
            var cropY = (int)(_selectionRect.Y * scaleY);
            var cropW = (int)(_selectionRect.Width * scaleX);
            var cropH = (int)(_selectionRect.Height * scaleY);

            cropX = Math.Max(0, Math.Min(cropX, _currentScreenshotBitmap.Width - 1));
            cropY = Math.Max(0, Math.Min(cropY, _currentScreenshotBitmap.Height - 1));
            cropW = Math.Min(cropW, _currentScreenshotBitmap.Width - cropX);
            cropH = Math.Min(cropH, _currentScreenshotBitmap.Height - cropY);

            if (cropW <= 0 || cropH <= 0) return null;

            var srcRect = new System.Drawing.Rectangle(cropX, cropY, cropW, cropH);
            var dst = new System.Drawing.Bitmap(cropW, cropH);

            using var g = System.Drawing.Graphics.FromImage(dst);
            g.DrawImage(_currentScreenshotBitmap, 0, 0, srcRect, System.Drawing.GraphicsUnit.Pixel);

            return dst;
        }
        catch
        {
            return _currentScreenshotBitmap;
        }
    }

    #endregion

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

    #region Next Button Handler

    private async void NextButton_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedWindow == null) return;

        try
        {
            // Get the current foreground window (our main window)
            var currentWindow = GetForegroundWindow();

            // Get the bounds of the selected window
            if (!GetWindowRect(_selectedWindow.Handle, out var rect))
                return;

            // Calculate the center point of the selected window
            var centerX = (rect.Left + rect.Right) / 2;
            var centerY = (rect.Top + rect.Bottom) / 2;

            // Bring the selected window to foreground
            SetForegroundWindow(_selectedWindow.Handle);

            // Small delay to ensure window is focused
            await Task.Delay(100);

            // Set cursor position to center of window
            SetCursorPos(centerX, centerY);

            // Small delay before clicking
            await Task.Delay(50);

            // Perform left mouse click
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

            // Small delay after clicking
            await Task.Delay(50);

            // Return focus to our main window
            SetForegroundWindow(currentWindow);
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Failed to click window: {ex.Message}";
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
