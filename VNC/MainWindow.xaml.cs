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
using Microsoft.UI.Xaml.Input;
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
    
    // Selection box state
    private bool _isSelectionMode = false;
    private bool _isDrawingSelection = false;
    private Point _selectionStartPoint;
    private Windows.Graphics.PointInt32 _selectionStartScreenPoint; // Actual cursor position when selection started
    private Rect? _selectedArea = null;

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
    private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

    [DllImport("user32.dll")]
    private static extern IntPtr GetDpiForWindow(IntPtr hwnd);

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

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

    // Window management constants
    private const int WM_NCLBUTTONDOWN = 0x00A1;
    private const int HTCAPTION = 0x0002;
    private const int SW_MINIMIZE = 6;
    private const int SW_MAXIMIZE = 3;
    private const int SW_RESTORE = 9;
    private const int SW_SHOWMAXIMIZED = 3;
    private const int GWL_STYLE = -16;
    private const int WS_THICKFRAME = 0x00040000;
    private const int WS_MAXIMIZEBOX = 0x00010000;
    private const int WS_MINIMIZEBOX = 0x00020000;

    public const int ICON_SMALL = 0;  
    public const int ICON_BIG = 1;  
    public const int ICON_SMALL2 = 2;  
      
    public const int WM_GETICON = 0x007F;  
    public const int WM_SETICON = 0x0080;  

    [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Auto)]  
    public static extern int SendMessage(IntPtr hWnd, uint msg, int wParam, IntPtr lParam);

    [DllImport("user32.dll", EntryPoint = "SendMessage")]
    private static extern IntPtr SendMessageForDrag(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool ReleaseCapture();

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll")]
    private static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

    [StructLayout(LayoutKind.Sequential)]
    private struct WINDOWPLACEMENT
    {
        public int length;
        public int flags;
        public int showCmd;
        public System.Drawing.Point ptMinPosition;
        public System.Drawing.Point ptMaxPosition;
        public System.Drawing.Rectangle rcNormalPosition;
    }  

    public MainWindow()
    {
        InitializeComponent();

        // Get window handle and configure window
        _hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
        Microsoft.UI.WindowId windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(_hwnd);
        _appWindow = Microsoft.UI.Windowing.AppWindow.GetFromWindowId(windowId);

        // Configure window presenter - disable native title bar, enable resizing
        _presenter = _appWindow.Presenter as Microsoft.UI.Windowing.OverlappedPresenter;
        _presenter.IsResizable = true;
        _presenter.IsMaximizable = true;
        _presenter.IsMinimizable = true;
        _presenter.SetBorderAndTitleBar(true, false);

        // Set initial window size
        _appWindow.Resize(new Windows.Graphics.SizeInt32(1920, 1080));

        // Ensure window has proper styles for resizing on all edges
        EnsureWindowResizable();

        // Set window corner preference for Windows 11
        SetWindowCornerPreference();
        
        // Update maximize/restore button icon based on current state
        UpdateMaximizeRestoreButton();

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

        // Wire up the GridSplitter columns
        ColumnSplitter.LeftColumn = TransparentColumn;
        ColumnSplitter.RightColumn = LlmColumn;

        // Wire up selection canvas events
        InitializeSelectionHandlers();

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

    private void SelectAreaButton_Click(object sender, RoutedEventArgs e)
    {
        _isSelectionMode = !_isSelectionMode;
        
        if (_isSelectionMode)
        {
            // Entering selection mode - clear any previous selection
            _selectedArea = null;
            SelectionRectangle.Visibility = Visibility.Collapsed;
            
            SelectAreaButton.Content = "Cancel Selection";
            SelectionCanvas.IsHitTestVisible = true;
            LlmAnalysisButton.IsEnabled = false;
        }
        else
        {
            SelectAreaButton.Content = "Select Area";
            SelectionCanvas.IsHitTestVisible = false;
            LlmAnalysisButton.IsEnabled = true;
            
            // If canceling without making a selection, hide the rectangle
            if (!_selectedArea.HasValue)
            {
                SelectionRectangle.Visibility = Visibility.Collapsed;
            }
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
            System.Diagnostics.Debug.WriteLine($"Failed to load settings: {ex.Message}");
        }
    }

    #region Custom Title Bar Handlers
    private void MinimizeButton_Click(object sender, RoutedEventArgs e)
    {
        // Minimize window using Win32 API
        ShowWindow(_hwnd, SW_MINIMIZE);
    }

    private void MaximizeRestoreButton_Click(object sender, RoutedEventArgs e)
    {
        // Toggle between maximized and normal state
        if (IsWindowMaximized())
        {
            ShowWindow(_hwnd, SW_RESTORE);
        }
        else
        {
            ShowWindow(_hwnd, SW_MAXIMIZE);
        }
        UpdateMaximizeRestoreButton();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        this.Close();
    }

    private void UpdateMaximizeRestoreButton()
    {
        if (IsWindowMaximized())
        {
            MaximizeRestoreButton.Content = "\uE923"; // Restore icon
            ToolTipService.SetToolTip(MaximizeRestoreButton, "Restore Down");
        }
        else
        {
            MaximizeRestoreButton.Content = "\uE922"; // Maximize icon
            ToolTipService.SetToolTip(MaximizeRestoreButton, "Maximize");
        }
    }

    private bool IsWindowMaximized()
    {
        WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
        placement.length = Marshal.SizeOf(placement);
        GetWindowPlacement(_hwnd, ref placement);
        return placement.showCmd == SW_SHOWMAXIMIZED;
    }

    #endregion

    #region Selection Box Handlers
    
    private void InitializeSelectionHandlers()
    {
        SelectionCanvas.PointerPressed += SelectionCanvas_PointerPressed;
        SelectionCanvas.PointerMoved += SelectionCanvas_PointerMoved;
        SelectionCanvas.PointerReleased += SelectionCanvas_PointerReleased;
    }
    
    private void SelectionCanvas_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        if (!_isSelectionMode) return;
        
        var point = e.GetCurrentPoint(SelectionCanvas).Position;
        _selectionStartPoint = point;
        _isDrawingSelection = true;
        
        // Capture the actual screen cursor position at selection start
        NativeInterop.GetCursorPos(out _selectionStartScreenPoint);
        
        // Reset and show rectangle
        SelectionRectangle.Visibility = Visibility.Visible;
        Canvas.SetLeft(SelectionRectangle, point.X);
        Canvas.SetTop(SelectionRectangle, point.Y);
        SelectionRectangle.Width = 0;
        SelectionRectangle.Height = 0;
        
        SelectionCanvas.CapturePointer(e.Pointer);
        e.Handled = true;
    }
    
    private void SelectionCanvas_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (!_isDrawingSelection) return;
        
        var currentPoint = e.GetCurrentPoint(SelectionCanvas).Position;
        
        // Calculate the rectangle bounds
        double left = Math.Min(_selectionStartPoint.X, currentPoint.X);
        double top = Math.Min(_selectionStartPoint.Y, currentPoint.Y);
        double width = Math.Abs(currentPoint.X - _selectionStartPoint.X);
        double height = Math.Abs(currentPoint.Y - _selectionStartPoint.Y);
        
        // Update rectangle
        Canvas.SetLeft(SelectionRectangle, left);
        Canvas.SetTop(SelectionRectangle, top);
        SelectionRectangle.Width = width;
        SelectionRectangle.Height = height;
        
        e.Handled = true;
    }
    
    private void SelectionCanvas_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (!_isDrawingSelection) return;
        
        _isDrawingSelection = false;
        SelectionCanvas.ReleasePointerCaptures();
        
        var currentPoint = e.GetCurrentPoint(SelectionCanvas).Position;
        
        // Store the selected area in canvas coordinates
        double left = Math.Min(_selectionStartPoint.X, currentPoint.X);
        double top = Math.Min(_selectionStartPoint.Y, currentPoint.Y);
        double width = Math.Abs(currentPoint.X - _selectionStartPoint.X);
        double height = Math.Abs(currentPoint.Y - _selectionStartPoint.Y);
        
        // Only store if the selection has meaningful size
        if (width > 5 && height > 5)
        {
            _selectedArea = new Rect(left, top, width, height);
            
            System.Diagnostics.Debug.WriteLine($"=== Selection Stored ===");
            System.Diagnostics.Debug.WriteLine($"Canvas coords: ({left}, {top}, {width}, {height})");
            System.Diagnostics.Debug.WriteLine($"Start screen cursor: ({_selectionStartScreenPoint.X}, {_selectionStartScreenPoint.Y})");
            
            // Exit selection mode and enable translate
            _isSelectionMode = false;
            SelectAreaButton.Content = "Select Area";
            SelectionCanvas.IsHitTestVisible = false;
            LlmAnalysisButton.IsEnabled = true;
        }
        else
        {
            // Selection too small, clear it
            SelectionRectangle.Visibility = Visibility.Collapsed;
            _selectedArea = null;
        }
        
        e.Handled = true;
    }
    
    #endregion

    private async void LlmAnalysisButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
        LlmAnalysisButton.IsEnabled = false;
            LlmAnalysisButton.Content = "Processing...";

            // Get the screen coordinates of the capture area (either selected box or full panel)
            var panelBounds = GetCaptureAreaScreenBounds();
            if (panelBounds.Width <= 0 || panelBounds.Height <= 0)
            {
                LlmResultTextBox.Text = "Could not determine the capture area bounds.";
                return;
            }

            // Debug: Log the capture area for verification
            System.Diagnostics.Debug.WriteLine($"Capturing area: X={panelBounds.X}, Y={panelBounds.Y}, W={panelBounds.Width}, H={panelBounds.Height}");
            if (_selectedArea.HasValue)
            {
                System.Diagnostics.Debug.WriteLine($"Selected area (local): X={_selectedArea.Value.X}, Y={_selectedArea.Value.Y}, W={_selectedArea.Value.Width}, H={_selectedArea.Value.Height}");
            }

            // Capture the screen region
            var capturedBitmap = CaptureScreenRegion(panelBounds);
            if (capturedBitmap == null)
            {
                LlmResultTextBox.Text = "Screenshot capture failed. Please try again.";
            return;
        }

            // Debug: Save the captured bitmap to verify what's being captured
            try
            {
                var debugPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    $"debug_capture_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                capturedBitmap.Save(debugPath, System.Drawing.Imaging.ImageFormat.Png);
                System.Diagnostics.Debug.WriteLine($"Debug capture saved to: {debugPath}");
                
                // Also save coordinates to a text file
                var coordsPath = debugPath.Replace(".png", "_coords.txt");
                var coordsText = $"Capture Details:\n" +
                    $"Screen bounds: X={panelBounds.X}, Y={panelBounds.Y}, W={panelBounds.Width}, H={panelBounds.Height}\n" +
                    $"Has selection: {_selectedArea.HasValue}\n";
                if (_selectedArea.HasValue)
                {
                    coordsText += $"Selection canvas coords: X={_selectedArea.Value.X}, Y={_selectedArea.Value.Y}, W={_selectedArea.Value.Width}, H={_selectedArea.Value.Height}\n";
                }
                System.IO.File.WriteAllText(coordsPath, coordsText);
            }
            catch (Exception debugEx)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to save debug capture: {debugEx.Message}");
            }

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
                var base64Image = ConvertBitmapToBase64(capturedBitmap);
                llmResult = await ManagedClient.AnalyzeImageAsync(base64Image, "http://172.178.84.127:8080");
            }
            else
            {
                // Use existing logic for BYOK and local modes
                // Check if OCR mode is enabled
                if (_appSettings.Model == "OCR")
                {
                    // Check if OCR is available
                    if (!OcrService.IsOcrAvailable())
                    {
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
                        LlmResultTextBox.Text = $"OCR extraction failed: {ex.Message}";
                        capturedBitmap.Dispose();
                        return;
                    }
                    
                    if (string.IsNullOrWhiteSpace(extractedText))
                    {
                        LlmResultTextBox.Text = "No text was detected in the captured area. Please try capturing a different region or switch to Image mode.";
                        capturedBitmap.Dispose();
                        return;
                    }
                    
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

    private System.Drawing.Rectangle GetCaptureAreaScreenBounds()
    {
        try
        {
            // If there's a selected area, calculate screen coordinates dynamically
            if (_selectedArea.HasValue)
            {
                // Get current cursor position of where selection START was in canvas coordinates
                var transform = SelectionCanvas.TransformToVisual(null);
                var canvasStartInWindow = transform.TransformPoint(
                    new Windows.Foundation.Point(_selectionStartPoint.X, _selectionStartPoint.Y)
                );
                
                // Get current window position
                if (!GetWindowRect(_hwnd, out var windowRect))
                {
                    return System.Drawing.Rectangle.Empty;
                }
                
                // Calculate where the start point WOULD BE on screen now
                int currentStartScreenX = windowRect.Left + (int)canvasStartInWindow.X;
                int currentStartScreenY = windowRect.Top + (int)canvasStartInWindow.Y;
                
                // Calculate the offset between where we clicked and where the transform says we are
                int offsetX = _selectionStartScreenPoint.X - currentStartScreenX;
                int offsetY = _selectionStartScreenPoint.Y - currentStartScreenY;
                
                // Get the selection top-left in window coordinates
                var selectionTopLeftInWindow = transform.TransformPoint(
                    new Windows.Foundation.Point(_selectedArea.Value.X, _selectedArea.Value.Y)
                );
                
                // Get DPI scaling factor
                uint dpi = (uint)GetDpiForWindow(_hwnd);
                double scale = dpi / 96.0; // 96 is the standard DPI
                
                // Calculate final screen position with offset correction
                int screenX = windowRect.Left + (int)selectionTopLeftInWindow.X + offsetX;
                int screenY = windowRect.Top + (int)selectionTopLeftInWindow.Y + offsetY;
                
                // Apply DPI scaling to width and height
                int width = (int)(_selectedArea.Value.Width * scale);
                int height = (int)(_selectedArea.Value.Height * scale);

                System.Diagnostics.Debug.WriteLine($"=== Capture Area Calculation ===");
                System.Diagnostics.Debug.WriteLine($"DPI: {dpi}, Scale factor: {scale}");
                System.Diagnostics.Debug.WriteLine($"Selection canvas coords: ({_selectedArea.Value.X}, {_selectedArea.Value.Y}, {_selectedArea.Value.Width}, {_selectedArea.Value.Height})");
                System.Diagnostics.Debug.WriteLine($"Start point canvas coords: ({_selectionStartPoint.X}, {_selectionStartPoint.Y})");
                System.Diagnostics.Debug.WriteLine($"Start point in window: ({canvasStartInWindow.X}, {canvasStartInWindow.Y})");
                System.Diagnostics.Debug.WriteLine($"Window rect: ({windowRect.Left}, {windowRect.Top})");
                System.Diagnostics.Debug.WriteLine($"Current start would be at screen: ({currentStartScreenX}, {currentStartScreenY})");
                System.Diagnostics.Debug.WriteLine($"Original start was at screen: ({_selectionStartScreenPoint.X}, {_selectionStartScreenPoint.Y})");
                System.Diagnostics.Debug.WriteLine($"Offset: ({offsetX}, {offsetY})");
                System.Diagnostics.Debug.WriteLine($"Selection top-left in window: ({selectionTopLeftInWindow.X}, {selectionTopLeftInWindow.Y})");
                System.Diagnostics.Debug.WriteLine($"Scaled dimensions: original ({_selectedArea.Value.Width}, {_selectedArea.Value.Height}) -> scaled ({width}, {height})");
                System.Diagnostics.Debug.WriteLine($"Final screen coords: ({screenX}, {screenY}, {width}, {height})");

                return new System.Drawing.Rectangle(screenX, screenY, width, height);
            }
            else
            {
                // Capture the full panel
                System.Diagnostics.Debug.WriteLine($"=== No Selection - Capturing Full Panel ===");
                return GetSwapChainPanelScreenBounds();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"GetCaptureAreaScreenBounds error: {ex.Message}");
            return System.Drawing.Rectangle.Empty;
        }
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
    /// Ensures the window has proper styles for resizing on all edges and corners.
    /// </summary>
    private void EnsureWindowResizable()
    {
        try
        {
            // Get current window style
            int style = GetWindowLong(_hwnd, GWL_STYLE);
            
            // Ensure the window has thick frame (resizable border) and maximize/minimize boxes
            style |= WS_THICKFRAME | WS_MAXIMIZEBOX | WS_MINIMIZEBOX;
            
            // Set the updated style
            SetWindowLong(_hwnd, GWL_STYLE, style);
        }
        catch (Exception ex)
        {
            // Log error if needed, but don't crash
            System.Diagnostics.Debug.WriteLine($"Failed to set window style: {ex.Message}");
        }
    }

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
        
        // Attach resize handlers to the border elements
        AttachResizeHandlers();
        
        // Initialize cursor timer to continuously update cursor
        InitializeCursorTimer();
        
        // Hook window procedure to intercept WM_SETCURSOR
        HookWindowProc();
    }
    
    /// <summary>
    /// Hooks into the window procedure to intercept WM_SETCURSOR messages.
    /// </summary>
    private void HookWindowProc()
    {
        _newWndProc = new NativeInterop.WndProcDelegate(WndProc);
        _oldWndProc = NativeInterop.SetWindowLongPtr(_hwnd, NativeInterop.GWLP_WNDPROC, 
            Marshal.GetFunctionPointerForDelegate(_newWndProc));
    }
    
    /// <summary>
    /// Custom window procedure to intercept cursor changes.
    /// </summary>
    private IntPtr WndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
    {
        if (msg == NativeInterop.WM_SETCURSOR)
        {
            // If we have a custom cursor set, use it
            if (_currentCursor != IntPtr.Zero)
            {
                NativeInterop.SetCursor(_currentCursor);
                return new IntPtr(1); // Return TRUE to prevent default processing
            }
        }
        
        // Call the original window procedure
        return NativeInterop.CallWindowProc(_oldWndProc, hWnd, msg, wParam, lParam);
    }
    
    /// <summary>
    /// Initializes a timer to continuously update the cursor.
    /// </summary>
    private void InitializeCursorTimer()
    {
        _cursorTimer = new DispatcherTimer();
        _cursorTimer.Interval = TimeSpan.FromMilliseconds(10);
        _cursorTimer.Tick += (s, e) =>
        {
            if (_isResizing)
            {
                UpdateCurrentCursor(_resizeDirection);
            }
            else if (_currentHoverDirection != ResizeDirection.None)
            {
                UpdateCurrentCursor(_currentHoverDirection);
            }
            else
            {
                UpdateCurrentCursor(ResizeDirection.None);
            }
        };
        _cursorTimer.Start();
    }
    
    /// <summary>
    /// Updates the current cursor based on direction.
    /// </summary>
    private void UpdateCurrentCursor(ResizeDirection direction)
    {
        IntPtr newCursor = direction switch
        {
            ResizeDirection.Left or ResizeDirection.Right => 
                NativeInterop.LoadCursor(IntPtr.Zero, NativeInterop.IDC_SIZEWE),
            ResizeDirection.Top or ResizeDirection.Bottom => 
                NativeInterop.LoadCursor(IntPtr.Zero, NativeInterop.IDC_SIZENS),
            ResizeDirection.TopLeft or ResizeDirection.BottomRight => 
                NativeInterop.LoadCursor(IntPtr.Zero, NativeInterop.IDC_SIZENWSE),
            ResizeDirection.TopRight or ResizeDirection.BottomLeft => 
                NativeInterop.LoadCursor(IntPtr.Zero, NativeInterop.IDC_SIZENESW),
            _ => IntPtr.Zero // Let default cursor show
        };
        
        if (_currentCursor != newCursor)
        {
            _currentCursor = newCursor;
            if (_currentCursor != IntPtr.Zero)
            {
                NativeInterop.SetCursor(_currentCursor);
            }
        }
    }
    
    /// <summary>
    /// Attaches resize event handlers to the invisible border elements.
    /// </summary>
    private void AttachResizeHandlers()
    {
        // Attach to edges
        ResizeTop.PointerPressed += (s, e) => StartResize(ResizeDirection.Top, e);
        ResizeTop.PointerEntered += (s, e) => { _currentHoverDirection = ResizeDirection.Top; };
        ResizeTop.PointerExited += (s, e) => { _currentHoverDirection = ResizeDirection.None; };
        
        ResizeBottom.PointerPressed += (s, e) => StartResize(ResizeDirection.Bottom, e);
        ResizeBottom.PointerEntered += (s, e) => { _currentHoverDirection = ResizeDirection.Bottom; };
        ResizeBottom.PointerExited += (s, e) => { _currentHoverDirection = ResizeDirection.None; };
        
        ResizeLeft.PointerPressed += (s, e) => StartResize(ResizeDirection.Left, e);
        ResizeLeft.PointerEntered += (s, e) => { _currentHoverDirection = ResizeDirection.Left; };
        ResizeLeft.PointerExited += (s, e) => { _currentHoverDirection = ResizeDirection.None; };
        
        ResizeRight.PointerPressed += (s, e) => StartResize(ResizeDirection.Right, e);
        ResizeRight.PointerEntered += (s, e) => { _currentHoverDirection = ResizeDirection.Right; };
        ResizeRight.PointerExited += (s, e) => { _currentHoverDirection = ResizeDirection.None; };
        
        // Attach to corners
        ResizeTopLeft.PointerPressed += (s, e) => StartResize(ResizeDirection.TopLeft, e);
        ResizeTopLeft.PointerEntered += (s, e) => { _currentHoverDirection = ResizeDirection.TopLeft; };
        ResizeTopLeft.PointerExited += (s, e) => { _currentHoverDirection = ResizeDirection.None; };
        
        ResizeTopRight.PointerPressed += (s, e) => StartResize(ResizeDirection.TopRight, e);
        ResizeTopRight.PointerEntered += (s, e) => { _currentHoverDirection = ResizeDirection.TopRight; };
        ResizeTopRight.PointerExited += (s, e) => { _currentHoverDirection = ResizeDirection.None; };
        
        ResizeBottomLeft.PointerPressed += (s, e) => StartResize(ResizeDirection.BottomLeft, e);
        ResizeBottomLeft.PointerEntered += (s, e) => { _currentHoverDirection = ResizeDirection.BottomLeft; };
        ResizeBottomLeft.PointerExited += (s, e) => { _currentHoverDirection = ResizeDirection.None; };
        
        ResizeBottomRight.PointerPressed += (s, e) => StartResize(ResizeDirection.BottomRight, e);
        ResizeBottomRight.PointerEntered += (s, e) => { _currentHoverDirection = ResizeDirection.BottomRight; };
        ResizeBottomRight.PointerExited += (s, e) => { _currentHoverDirection = ResizeDirection.None; };
        
        // Attach pointer moved and released to root for tracking during resize
        var root = (UIElement)this.Content;
        root.PointerMoved += OnResizePointerMoved;
        root.PointerReleased += OnResizePointerReleased;
    }
    
    
    private enum ResizeDirection
    {
        None,
        Left,
        Right,
        Top,
        Bottom,
        TopLeft,
        TopRight,
        BottomLeft,
        BottomRight
    }
    
    private bool _isResizing = false;
    private ResizeDirection _resizeDirection = ResizeDirection.None;
    private ResizeDirection _currentHoverDirection = ResizeDirection.None;
    private Windows.Graphics.PointInt32 _resizeStartScreenPoint;
    private Windows.Graphics.RectInt32 _windowStartBounds;
    private DispatcherTimer _cursorTimer;
    
    // Window procedure hook
    private IntPtr _oldWndProc = IntPtr.Zero;
    private NativeInterop.WndProcDelegate _newWndProc;
    private IntPtr _currentCursor = IntPtr.Zero;
    
    private void StartResize(ResizeDirection direction, PointerRoutedEventArgs e)
    {
        _isResizing = true;
        _resizeDirection = direction;
        
        // Get screen coordinates for accurate tracking
        NativeInterop.GetCursorPos(out _resizeStartScreenPoint);
        
        // Get current window bounds
        if (GetWindowRect(_hwnd, out RECT rect))
        {
            _windowStartBounds = new Windows.Graphics.RectInt32
            {
                X = rect.Left,
                Y = rect.Top,
                Width = rect.Right - rect.Left,
                Height = rect.Bottom - rect.Top
            };
        }
        
        ((UIElement)this.Content).CapturePointer(e.Pointer);
        e.Handled = true;
    }
    
    private void OnResizePointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (!_isResizing) return;
        
        // Get current screen coordinates
        NativeInterop.GetCursorPos(out Windows.Graphics.PointInt32 currentScreenPos);
        PerformResize(currentScreenPos);
        
        e.Handled = true;
    }
    
    private void OnResizePointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (_isResizing)
        {
            _isResizing = false;
            _resizeDirection = ResizeDirection.None;
            _currentHoverDirection = ResizeDirection.None;
            ((UIElement)sender).ReleasePointerCaptures();
            e.Handled = true;
        }
    }
    
    private void PerformResize(Windows.Graphics.PointInt32 currentScreenPos)
    {
        // Calculate delta from start position in screen coordinates
        var deltaX = currentScreenPos.X - _resizeStartScreenPoint.X;
        var deltaY = currentScreenPos.Y - _resizeStartScreenPoint.Y;

        var newX = _windowStartBounds.X;
        var newY = _windowStartBounds.Y;
        var newWidth = _windowStartBounds.Width;
        var newHeight = _windowStartBounds.Height;

        // Minimum window size
        const int MIN_WIDTH = 400;
        const int MIN_HEIGHT = 300;

        switch (_resizeDirection)
        {
            case ResizeDirection.Left:
                newX = _windowStartBounds.X + deltaX;
                newWidth = _windowStartBounds.Width - deltaX;
                if (newWidth < MIN_WIDTH)
                {
                    newX = _windowStartBounds.X + _windowStartBounds.Width - MIN_WIDTH;
                    newWidth = MIN_WIDTH;
                }
                break;

            case ResizeDirection.Right:
                newWidth = _windowStartBounds.Width + deltaX;
                if (newWidth < MIN_WIDTH) newWidth = MIN_WIDTH;
                break;

            case ResizeDirection.Top:
                newY = _windowStartBounds.Y + deltaY;
                newHeight = _windowStartBounds.Height - deltaY;
                if (newHeight < MIN_HEIGHT)
                {
                    newY = _windowStartBounds.Y + _windowStartBounds.Height - MIN_HEIGHT;
                    newHeight = MIN_HEIGHT;
                }
                break;

            case ResizeDirection.Bottom:
                newHeight = _windowStartBounds.Height + deltaY;
                if (newHeight < MIN_HEIGHT) newHeight = MIN_HEIGHT;
                break;

            case ResizeDirection.TopLeft:
                newX = _windowStartBounds.X + deltaX;
                newY = _windowStartBounds.Y + deltaY;
                newWidth = _windowStartBounds.Width - deltaX;
                newHeight = _windowStartBounds.Height - deltaY;
                
                if (newWidth < MIN_WIDTH)
                {
                    newX = _windowStartBounds.X + _windowStartBounds.Width - MIN_WIDTH;
                    newWidth = MIN_WIDTH;
                }
                if (newHeight < MIN_HEIGHT)
                {
                    newY = _windowStartBounds.Y + _windowStartBounds.Height - MIN_HEIGHT;
                    newHeight = MIN_HEIGHT;
                }
                break;

            case ResizeDirection.TopRight:
                newY = _windowStartBounds.Y + deltaY;
                newWidth = _windowStartBounds.Width + deltaX;
                newHeight = _windowStartBounds.Height - deltaY;
                
                if (newWidth < MIN_WIDTH) newWidth = MIN_WIDTH;
                if (newHeight < MIN_HEIGHT)
                {
                    newY = _windowStartBounds.Y + _windowStartBounds.Height - MIN_HEIGHT;
                    newHeight = MIN_HEIGHT;
                }
                break;

            case ResizeDirection.BottomLeft:
                newX = _windowStartBounds.X + deltaX;
                newWidth = _windowStartBounds.Width - deltaX;
                newHeight = _windowStartBounds.Height + deltaY;
                
                if (newWidth < MIN_WIDTH)
                {
                    newX = _windowStartBounds.X + _windowStartBounds.Width - MIN_WIDTH;
                    newWidth = MIN_WIDTH;
                }
                if (newHeight < MIN_HEIGHT) newHeight = MIN_HEIGHT;
                break;

            case ResizeDirection.BottomRight:
                newWidth = _windowStartBounds.Width + deltaX;
                newHeight = _windowStartBounds.Height + deltaY;
                
                if (newWidth < MIN_WIDTH) newWidth = MIN_WIDTH;
                if (newHeight < MIN_HEIGHT) newHeight = MIN_HEIGHT;
                break;
        }

        // Apply the new bounds
        _appWindow.MoveAndResize(new Windows.Graphics.RectInt32
        {
            X = newX,
            Y = newY,
            Width = newWidth,
            Height = newHeight
        });
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
