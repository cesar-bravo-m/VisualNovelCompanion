using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.UI.Xaml.Media.Imaging;
using Windows.Storage.Streams;

namespace VNC
{
    public class WindowInfo
    {
        public IntPtr Handle { get; set; }
        public string Title { get; set; } = string.Empty;
        public BitmapImage? Thumbnail { get; set; }
        public Rectangle Bounds { get; set; }
        public bool IsValid { get; set; } = true;
    }

    public class ScreenshotCapture
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);

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

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);

        [DllImport("dwmapi.dll")]
        private static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [DllImport("shcore.dll")]
        private static extern int SetProcessDpiAwareness(int value);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("gdi32.dll")]
        private static extern int SetStretchBltMode(IntPtr hdc, int mode);

        [DllImport("gdi32.dll")]
        private static extern bool StretchBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest, 
            IntPtr hdcSrc, int xSrc, int ySrc, int wSrc, int hSrc, uint rop);

        private const uint SRCCOPY = 0x00CC0020;
        private const int SM_CXSCREEN = 0;
        private const int SM_CYSCREEN = 1;
        private const uint PW_CLIENTONLY = 1;
        private const uint PW_RENDERFULLCONTENT = 2;
        private const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;
        private const int SW_RESTORE = 9;
        private const int PROCESS_PER_MONITOR_DPI_AWARE = 2;
        private const int HALFTONE = 4;

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MONITORINFO
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public uint dwFlags;
        }

        static ScreenshotCapture()
        {
            try
            {
                SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE);
            }
            catch
            {
                try
                {
                    SetProcessDPIAware();
                }
                catch
                {
                }
            }
        }

        /// <summary>
        /// Captures a screenshot of the entire screen and returns it as a BitmapImage
        /// </summary>
        /// <returns>BitmapImage containing the screenshot</returns>
        public static async Task<BitmapImage> CaptureScreenshotAsync()
        {
        try
        {
            int screenWidth = GetSystemMetrics(SM_CXSCREEN);
                int screenHeight = GetSystemMetrics(SM_CYSCREEN);

                IntPtr desktopHwnd = GetDesktopWindow();
                IntPtr desktopDc = GetWindowDC(desktopHwnd);

                IntPtr memoryDc = CreateCompatibleDC(desktopDc);
                IntPtr bitmap = CreateCompatibleBitmap(desktopDc, screenWidth, screenHeight);

                // Select the bitmap into the memory device context
                IntPtr oldBitmap = SelectObject(memoryDc, bitmap);

                // Copy the desktop to the bitmap
                BitBlt(memoryDc, 0, 0, screenWidth, screenHeight, desktopDc, 0, 0, SRCCOPY);

                // Create a .NET Bitmap from the Win32 bitmap
                Bitmap screenBitmap = Image.FromHbitmap(bitmap);

                // Convert to BitmapImage for WinUI
                BitmapImage bitmapImage = await ConvertToBitmapImageAsync(screenBitmap);

                // Clean up resources
                SelectObject(memoryDc, oldBitmap);
                DeleteObject(bitmap);
                DeleteDC(memoryDc);
                ReleaseDC(desktopHwnd, desktopDc);
                screenBitmap.Dispose();

                return bitmapImage;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to capture screenshot: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Converts a System.Drawing.Bitmap to a WinUI BitmapImage
        /// </summary>
        /// <param name="bitmap">The bitmap to convert</param>
        /// <returns>BitmapImage for use in WinUI controls</returns>
        private static async Task<BitmapImage> ConvertToBitmapImageAsync(Bitmap bitmap)
        {
            using (var memoryStream = new MemoryStream())
            {
                // Save bitmap to memory stream as PNG
                bitmap.Save(memoryStream, ImageFormat.Png);
                memoryStream.Position = 0;

                // Create BitmapImage and set source from stream
                var bitmapImage = new BitmapImage();
                var randomAccessStream = memoryStream.AsRandomAccessStream();
                
                await bitmapImage.SetSourceAsync(randomAccessStream);
                
                return bitmapImage;
            }
        }

        /// <summary>
        /// Gets the dimensions of the primary screen
        /// </summary>
        /// <returns>Size containing width and height of the screen</returns>
        public static Size GetScreenSize()
        {
            int width = GetSystemMetrics(SM_CXSCREEN);
            int height = GetSystemMetrics(SM_CYSCREEN);
            return new Size(width, height);
        }

        /// <summary>
        /// Enumerates all visible windows and captures thumbnails
        /// </summary>
        /// <returns>List of WindowInfo objects with thumbnails</returns>
        public static async Task<List<WindowInfo>> GetVisibleWindowsAsync()
        {
            var windows = new List<WindowInfo>();
            
            EnumWindows((hWnd, lParam) =>
            {
                if (IsWindowVisible(hWnd) && GetWindowTextLength(hWnd) > 0)
                {
                    var title = GetWindowTitle(hWnd);
                    if (!string.IsNullOrEmpty(title) && !IsSystemWindow(title) && IsValidCaptureWindow(hWnd))
                    {
                        if (GetWindowRect(hWnd, out RECT rect))
                        {
                            // Filter out tiny windows (likely system windows)
                            int width = rect.Right - rect.Left;
                            int height = rect.Bottom - rect.Top;
                            
                            if (width >= 100 && height >= 100) // Minimum reasonable window size
                            {
                                var windowInfo = new WindowInfo
                                {
                                    Handle = hWnd,
                                    Title = title,
                                    Bounds = new Rectangle(rect.Left, rect.Top, width, height)
                                };
                                windows.Add(windowInfo);
                            }
                        }
                    }
                }
                return true;
            }, IntPtr.Zero);

            // Sort windows by title for better user experience
            windows = windows.OrderBy(w => w.Title).ToList();

            // Capture thumbnails for each window (in parallel for better performance)
            var thumbnailTasks = windows.Select(async window =>
            {
                try
                {
                    window.Thumbnail = await CaptureWindowThumbnailAsync(window.Handle);
                    return window;
                }
                catch
                {
                    // If thumbnail capture fails, mark as invalid but keep in list
                    window.IsValid = false;
                    window.Thumbnail = await ConvertToBitmapImageAsync(CreatePlaceholderThumbnail(window.Title));
                    return window;
                }
            });

            await Task.WhenAll(thumbnailTasks);

            return windows;
        }

        /// <summary>
        /// Gets visible windows without capturing thumbnails (for dropdown lists)
        /// </summary>
        public static async Task<List<WindowInfo>> GetVisibleWindowsWithoutThumbnailsAsync()
        {
            var windows = new List<WindowInfo>();
            
            await Task.Run(() =>
            {
                EnumWindows((hWnd, lParam) =>
                {
                    if (IsWindowVisible(hWnd) && GetWindowTextLength(hWnd) > 0)
                    {
                        var title = GetWindowTitle(hWnd);
                        if (!string.IsNullOrEmpty(title) && !IsSystemWindow(title) && IsValidCaptureWindow(hWnd))
                        {
                            if (GetWindowRect(hWnd, out RECT rect))
                            {
                                // Filter out tiny windows (likely system windows)
                                int width = rect.Right - rect.Left;
                                int height = rect.Bottom - rect.Top;
                                
                                if (width >= 100 && height >= 100) // Minimum reasonable window size
                                {
                                    var windowInfo = new WindowInfo
                                    {
                                        Handle = hWnd,
                                        Title = title,
                                        Bounds = new Rectangle(rect.Left, rect.Top, width, height),
                                        IsValid = true,
                                        Thumbnail = null // No thumbnail needed for dropdown
                                    };
                                    windows.Add(windowInfo);
                                }
                            }
                        }
                    }
                    return true;
                }, IntPtr.Zero);
            });

            // Sort windows by title for better user experience
            return windows.OrderBy(w => w.Title).ToList();
        }

        /// <summary>
        /// Checks if a window is valid for capture (not a system window, popup, etc.)
        /// </summary>
        private static bool IsValidCaptureWindow(IntPtr hWnd)
        {
            // Skip windows that are likely system windows or popups
            // This is a more sophisticated check than just title filtering
            
            // Skip if window is our own application
            var currentProcess = System.Diagnostics.Process.GetCurrentProcess();
            if (hWnd == currentProcess.MainWindowHandle)
                return false;

            // Additional checks could be added here for specific window styles, etc.
            return true;
        }

        /// <summary>
        /// Captures a screenshot of a specific window
        /// </summary>
        /// <param name="windowHandle">Handle of the window to capture</param>
        /// <returns>BitmapImage containing the window screenshot</returns>
        public static async Task<BitmapImage> CaptureWindowAsync(IntPtr windowHandle)
        {
            var result = await CaptureWindowWithBitmapAsync(windowHandle);
            return result.BitmapImage;
        }

        /// <summary>
        /// Captures a screenshot of a specific window and returns both BitmapImage and System.Drawing.Bitmap
        /// </summary>
        /// <param name="windowHandle">Handle of the window to capture</param>
        /// <returns>Tuple containing both BitmapImage and Bitmap</returns>
        public static async Task<(BitmapImage BitmapImage, Bitmap Bitmap)> CaptureWindowWithBitmapAsync(IntPtr windowHandle)
        {
            if (!IsWindow(windowHandle))
            {
                throw new InvalidOperationException("Window is no longer valid");
            }

            try
            {
                // Try multiple capture methods for full-size window capture
                // For games, BitBlt often works better than PrintWindow
                Bitmap? windowBitmap = null;
                
                // Method 1: Try BitBlt first (better for games and hardware-accelerated content)
                windowBitmap = await TryBitBltCapture(windowHandle, false);
                
                // Method 2: Try enhanced PrintWindow with full content
                if (windowBitmap == null)
                {
                    windowBitmap = await TryEnhancedPrintWindowCapture(windowHandle, true);
                }
                
                // Method 3: Try standard PrintWindow if enhanced failed
                if (windowBitmap == null)
                {
                    windowBitmap = await TryEnhancedPrintWindowCapture(windowHandle, false);
                }
                
                // Method 4: Try legacy method as last resort
                if (windowBitmap == null)
                {
                    windowBitmap = await TryLegacyCapture(windowHandle);
                }

                if (windowBitmap == null)
                {
                    throw new InvalidOperationException("All capture methods failed");
                }

                // Convert to BitmapImage for WinUI
                BitmapImage bitmapImage = await ConvertToBitmapImageAsync(windowBitmap);
                
                // Return both the BitmapImage and a copy of the original Bitmap
                var bitmapCopy = new Bitmap(windowBitmap);
                windowBitmap.Dispose();

                return (bitmapImage, bitmapCopy);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to capture window: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Legacy capture method for compatibility
        /// </summary>
        private static async Task<Bitmap?> TryLegacyCapture(IntPtr windowHandle)
        {
            try
            {
                // Get window bounds
                if (!GetWindowRect(windowHandle, out RECT rect))
                {
                    return null;
                }

                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;

                if (width <= 0 || height <= 0)
                {
                    return null;
                }

                // Get window device context
                IntPtr windowDc = GetWindowDC(windowHandle);
                if (windowDc == IntPtr.Zero)
                {
                    return null;
                }

                // Create compatible device context and bitmap
                IntPtr memoryDc = CreateCompatibleDC(windowDc);
                IntPtr bitmap = CreateCompatibleBitmap(windowDc, width, height);

                // Select the bitmap into the memory device context
                IntPtr oldBitmap = SelectObject(memoryDc, bitmap);

                // Try PrintWindow first
                bool success = PrintWindow(windowHandle, memoryDc, PW_CLIENTONLY);
                
                if (!success)
                {
                    // Fallback to BitBlt
                    success = BitBlt(memoryDc, 0, 0, width, height, windowDc, 0, 0, SRCCOPY);
                }

                Bitmap? result = null;
                if (success)
                {
                    result = Image.FromHbitmap(bitmap);
                }

                // Clean up resources
                SelectObject(memoryDc, oldBitmap);
                DeleteObject(bitmap);
                DeleteDC(memoryDc);
                ReleaseDC(windowHandle, windowDc);

                return result;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Captures a small thumbnail of a window for selection UI
        /// </summary>
        /// <param name="windowHandle">Handle of the window</param>
        /// <returns>Small BitmapImage thumbnail</returns>
        private static async Task<BitmapImage> CaptureWindowThumbnailAsync(IntPtr windowHandle)
        {
            if (!IsWindow(windowHandle))
            {
                throw new InvalidOperationException("Window is no longer valid");
            }

            try
            {
                // Try multiple capture methods to handle different window types
                Bitmap? thumbnail = null;
                
                // Method 1: Try enhanced PrintWindow with full content (thumbnail size)
                thumbnail = await TryEnhancedPrintWindowThumbnailCapture(windowHandle, true);
                
                // Method 2: Try standard PrintWindow if enhanced failed (thumbnail size)
                if (thumbnail == null)
                {
                    thumbnail = await TryEnhancedPrintWindowThumbnailCapture(windowHandle, false);
                }
                
                // Method 3: Try BitBlt as fallback
                if (thumbnail == null)
                {
                    thumbnail = await TryBitBltCapture(windowHandle, true);
                }
                
                // Method 4: Try screen capture as last resort
                if (thumbnail == null)
                {
                    thumbnail = await TryScreenCapture(windowHandle, true);
                }

                if (thumbnail == null)
                {
                    // Create a placeholder thumbnail
                    thumbnail = CreatePlaceholderThumbnail(GetWindowTitle(windowHandle));
                }

                // Convert to BitmapImage
                BitmapImage bitmapImage = await ConvertToBitmapImageAsync(thumbnail);
                thumbnail.Dispose();

                return bitmapImage;
            }
            catch (Exception ex)
            {
                // Create error placeholder
                var placeholder = CreatePlaceholderThumbnail("Error");
                var bitmapImage = await ConvertToBitmapImageAsync(placeholder);
                placeholder.Dispose();
                return bitmapImage;
            }
        }

        /// <summary>
        /// Enhanced PrintWindow capture with better handling of modern applications
        /// </summary>
        private static async Task<Bitmap?> TryEnhancedPrintWindowCapture(IntPtr windowHandle, bool useFullContent)
        {
            return await TryEnhancedPrintWindowCaptureInternal(windowHandle, useFullContent, false);
        }

        /// <summary>
        /// Enhanced PrintWindow capture for thumbnails
        /// </summary>
        private static async Task<Bitmap?> TryEnhancedPrintWindowThumbnailCapture(IntPtr windowHandle, bool useFullContent)
        {
            return await TryEnhancedPrintWindowCaptureInternal(windowHandle, useFullContent, true);
        }

        /// <summary>
        /// Internal enhanced PrintWindow capture with option for thumbnail or full size
        /// </summary>
        private static async Task<Bitmap?> TryEnhancedPrintWindowCaptureInternal(IntPtr windowHandle, bool useFullContent, bool createThumbnail)
        {
            try
            {
                // Get extended window bounds (includes drop shadows, etc.)
                RECT bounds;
                if (DwmGetWindowAttribute(windowHandle, DWMWA_EXTENDED_FRAME_BOUNDS, out bounds, Marshal.SizeOf<RECT>()) != 0)
                {
                    // Fallback to regular window rect
                    if (!GetWindowRect(windowHandle, out bounds))
                        return null;
                }

                int width = bounds.Right - bounds.Left;
                int height = bounds.Bottom - bounds.Top;

                if (width <= 0 || height <= 0 || width > 4000 || height > 4000)
                    return null;

                // Calculate target size
                int targetWidth = width;
                int targetHeight = height;
                
                if (createThumbnail)
                {
                    // Calculate thumbnail size (max 300px width/height, maintain aspect ratio)
                    const int maxSize = 300;
                    float scale = Math.Min((float)maxSize / width, (float)maxSize / height);
                    targetWidth = Math.Max(1, (int)(width * scale));
                    targetHeight = Math.Max(1, (int)(height * scale));
                }

                // Temporarily restore window if minimized
                bool wasMinimized = IsIconic(windowHandle);
                IntPtr previousForeground = GetForegroundWindow();
                
                if (wasMinimized)
                {
                    ShowWindow(windowHandle, SW_RESTORE);
                    await Task.Delay(100); // Give time for window to restore
                }

                // Get window device context
                IntPtr windowDc = GetWindowDC(windowHandle);
                if (windowDc == IntPtr.Zero)
                    return null;

                // Create memory DC and bitmap for full size capture
                IntPtr memoryDc = CreateCompatibleDC(windowDc);
                IntPtr bitmap = CreateCompatibleBitmap(windowDc, width, height);
                IntPtr oldBitmap = SelectObject(memoryDc, bitmap);

                // Try PrintWindow with appropriate flags
                uint flags = useFullContent ? PW_RENDERFULLCONTENT : PW_CLIENTONLY;
                bool success = PrintWindow(windowHandle, memoryDc, flags);

                Bitmap? result = null;
                if (success)
                {
                    // Create .NET bitmap from Win32 bitmap
                    Bitmap fullBitmap = Image.FromHbitmap(bitmap);
                    
                    if (createThumbnail && (targetWidth != width || targetHeight != height))
                    {
                        // Create thumbnail with high quality scaling
                        result = new Bitmap(targetWidth, targetHeight);
                        using (Graphics g = Graphics.FromImage(result))
                        {
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                            g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                            g.DrawImage(fullBitmap, 0, 0, targetWidth, targetHeight);
                        }
                        fullBitmap.Dispose();
                    }
                    else
                    {
                        // Return full size bitmap
                        result = fullBitmap;
                    }
                }

                // Cleanup
                SelectObject(memoryDc, oldBitmap);
                DeleteObject(bitmap);
                DeleteDC(memoryDc);
                ReleaseDC(windowHandle, windowDc);

                // Restore previous foreground window
                if (wasMinimized && previousForeground != IntPtr.Zero)
                {
                    SetForegroundWindow(previousForeground);
                }

                return result;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// BitBlt capture method as fallback
        /// </summary>
        private static async Task<Bitmap?> TryBitBltCapture(IntPtr windowHandle, bool isThumbnail)
        {
            try
            {
                if (!GetWindowRect(windowHandle, out RECT bounds))
                    return null;

                int width = bounds.Right - bounds.Left;
                int height = bounds.Bottom - bounds.Top;

                if (width <= 0 || height <= 0)
                    return null;

                int targetWidth = width;
                int targetHeight = height;
                
                if (isThumbnail)
                {
                    const int maxSize = 200;
                    float scale = Math.Min((float)maxSize / width, (float)maxSize / height);
                    targetWidth = Math.Max(1, (int)(width * scale));
                    targetHeight = Math.Max(1, (int)(height * scale));
                }

                // Get screen DC
                IntPtr screenDc = GetWindowDC(GetDesktopWindow());
                IntPtr memoryDc = CreateCompatibleDC(screenDc);
                IntPtr bitmap = CreateCompatibleBitmap(screenDc, targetWidth, targetHeight);
                IntPtr oldBitmap = SelectObject(memoryDc, bitmap);

                // Set high quality scaling
                SetStretchBltMode(memoryDc, HALFTONE);

                // Capture from screen coordinates
                bool success = StretchBlt(memoryDc, 0, 0, targetWidth, targetHeight, 
                                        screenDc, bounds.Left, bounds.Top, width, height, SRCCOPY);

                Bitmap? result = null;
                if (success)
                {
                    result = Image.FromHbitmap(bitmap);
                }

                // Cleanup
                SelectObject(memoryDc, oldBitmap);
                DeleteObject(bitmap);
                DeleteDC(memoryDc);
                ReleaseDC(GetDesktopWindow(), screenDc);

                return result;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Screen capture method for windows that don't respond to other methods
        /// </summary>
        private static async Task<Bitmap?> TryScreenCapture(IntPtr windowHandle, bool isThumbnail)
        {
            try
            {
                // This is essentially the same as BitBlt but with different approach
                return await TryBitBltCapture(windowHandle, isThumbnail);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Creates a placeholder thumbnail when capture fails
        /// </summary>
        private static Bitmap CreatePlaceholderThumbnail(string windowTitle)
        {
            const int size = 200;
            var bitmap = new Bitmap(size, size);
            
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill with gradient background
                using (var brush = new System.Drawing.Drawing2D.LinearGradientBrush(
                    new Rectangle(0, 0, size, size), 
                    Color.FromArgb(64, 64, 64), 
                    Color.FromArgb(32, 32, 32), 
                    45f))
                {
                    g.FillRectangle(brush, 0, 0, size, size);
                }

                // Draw border
                using (var pen = new Pen(Color.FromArgb(128, 128, 128), 2))
                {
                    g.DrawRectangle(pen, 1, 1, size - 2, size - 2);
                }

                // Draw window icon
                var iconRect = new Rectangle(size / 2 - 20, size / 2 - 30, 40, 40);
                using (var brush = new SolidBrush(Color.FromArgb(180, 180, 180)))
                {
                    g.FillRectangle(brush, iconRect);
                    g.DrawRectangle(Pens.White, iconRect);
                }

                // Draw title
                if (!string.IsNullOrEmpty(windowTitle))
                {
                    var titleRect = new Rectangle(10, size / 2 + 20, size - 20, 40);
                    using (var font = new Font("Segoe UI", 9, FontStyle.Regular))
                    using (var brush = new SolidBrush(Color.White))
                    {
                        var format = new StringFormat
                        {
                            Alignment = StringAlignment.Center,
                            LineAlignment = StringAlignment.Center,
                            Trimming = StringTrimming.EllipsisCharacter
                        };
                        
                        g.DrawString(windowTitle, font, brush, titleRect, format);
                    }
                }
            }
            
            return bitmap;
        }

        /// <summary>
        /// Checks if a window is still valid
        /// </summary>
        /// <param name="windowHandle">Handle of the window to check</param>
        /// <returns>True if window is still valid</returns>
        public static bool IsWindowValid(IntPtr windowHandle)
        {
            return IsWindow(windowHandle) && IsWindowVisible(windowHandle);
        }

        /// <summary>
        /// Gets the title of a window
        /// </summary>
        /// <param name="hWnd">Window handle</param>
        /// <returns>Window title</returns>
        private static string GetWindowTitle(IntPtr hWnd)
        {
            int length = GetWindowTextLength(hWnd);
            if (length == 0) return string.Empty;

            StringBuilder sb = new StringBuilder(length + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        /// <summary>
        /// Checks if a bitmap is mostly white or blank (indicating a failed capture)
        /// </summary>
        /// <param name="bitmap">The bitmap to analyze</param>
        /// <returns>True if the bitmap is mostly white/blank</returns>
        public static bool IsBitmapMostlyWhite(Bitmap bitmap)
        {
            if (bitmap == null) return true;

            try
            {
                // Quick check: sample a small grid of pixels for efficiency
                int width = bitmap.Width;
                int height = bitmap.Height;
                int whitePixels = 0;
                int totalSamples = 0;

                // Sample every 10th pixel in both directions for speed
                int step = 10;
                
                for (int x = 0; x < width; x += step)
                {
                    for (int y = 0; y < height; y += step)
                    {
                        if (x < width && y < height)
                        {
                            Color pixel = bitmap.GetPixel(x, y);
                            
                            // Consider a pixel "white" if it's very close to white
                            // This accounts for slight compression artifacts or anti-aliasing
                            bool isWhite = pixel.R > 248 && pixel.G > 248 && pixel.B > 248;
                            
                            if (isWhite)
                                whitePixels++;
                            
                            totalSamples++;
                            
                            // Early exit optimization: if we find enough non-white pixels early, it's not a white frame
                            if (totalSamples > 50 && (double)(totalSamples - whitePixels) / totalSamples > 0.1)
                            {
                                return false;
                            }
                        }
                    }
                }

                // If more than 90% of sampled pixels are white, consider it a white frame
                double whitePercentage = totalSamples > 0 ? (double)whitePixels / totalSamples : 1.0;
                return whitePercentage > 0.90;
            }
            catch
            {
                // If we can't analyze the bitmap, assume it's bad
                return true;
            }
        }

        /// <summary>
        /// Checks if a window title indicates a system window that should be filtered out
        /// </summary>
        /// <param name="title">Window title</param>
        /// <returns>True if this is a system window to filter out</returns>
        private static bool IsSystemWindow(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
                return true;

            var systemWindows = new[]
            {
                "Program Manager",
                "Desktop Window Manager", 
                "Windows Input Experience",
                "Microsoft Text Input Application",
                "Windows Shell Experience",
                "Search",
                "Cortana",
                "Microsoft Store",
                "Settings",
                "Windows Security",
                "Notification Area",
                "Action Center",
                "Start",
                "Task Manager",
                "Windows.UI.Core.CoreWindow",
                "ApplicationFrameHost",
                "Shell_TrayWnd",
                "DV2ControlHost",
                "MsgrIMEWindowClass",
                "SysShadow",
                "Button",
                "ToolbarWindow32",
                "Shell_SecondaryTrayWnd",
                "WorkerW",
                "Progman"
            };

            // Check for exact matches or contains
            foreach (var systemWindow in systemWindows)
            {
                if (title.Equals(systemWindow, StringComparison.OrdinalIgnoreCase) ||
                    title.Contains(systemWindow, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            // Filter out windows with suspicious patterns
            if (title.Length < 2 || 
                title.All(char.IsDigit) || 
                title.StartsWith("MSCTFIME", StringComparison.OrdinalIgnoreCase) ||
                title.StartsWith("Default IME", StringComparison.OrdinalIgnoreCase) ||
                title.Contains("GDI+ Window", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }
    }
}
