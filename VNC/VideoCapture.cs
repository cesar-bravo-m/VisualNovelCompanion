using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml.Media.Imaging;
using Windows.Storage.Streams;

namespace VNC
{
    public class VideoCapture : IDisposable
    {
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
        private static extern bool DeleteDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("dwmapi.dll")]
        private static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);

        private const uint PW_RENDERFULLCONTENT = 2;
        private const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private IntPtr _windowHandle;
        private Timer? _captureTimer;
        private bool _isCapturing;
        private int _frameInterval;
        private readonly DispatcherQueue _dispatcherQueue;
        
        public event EventHandler<BitmapImage>? FrameCaptured;
        public event EventHandler<string>? ErrorOccurred;

        public bool IsCapturing => _isCapturing;
        public int FrameRate { get; private set; }

        public VideoCapture(DispatcherQueue dispatcherQueue)
        {
            _dispatcherQueue = dispatcherQueue;
            FrameRate = 30; // Default to 30 FPS
            _frameInterval = 1000 / FrameRate;
        }

        /// <summary>
        /// Starts video capture for the specified window at the given frame rate
        /// </summary>
        /// <param name="windowHandle">Handle of the window to capture</param>
        /// <param name="fps">Frames per second (30 or 60)</param>
        public async Task StartCaptureAsync(IntPtr windowHandle, int fps = 30)
        {
            if (_isCapturing)
            {
                await StopCaptureAsync();
            }

            if (!IsWindow(windowHandle))
            {
                ErrorOccurred?.Invoke(this, "Invalid window handle");
                return;
            }

            _windowHandle = windowHandle;
            FrameRate = fps;
            _frameInterval = 1000 / FrameRate;
            _isCapturing = true;

            // Start the capture timer
            _captureTimer = new Timer(CaptureFrame, null, 0, _frameInterval);
        }

        /// <summary>
        /// Stops video capture
        /// </summary>
        public async Task StopCaptureAsync()
        {
            _isCapturing = false;
            
            if (_captureTimer != null)
            {
                await _captureTimer.DisposeAsync();
                _captureTimer = null;
            }
        }

        /// <summary>
        /// Changes the frame rate during capture
        /// </summary>
        /// <param name="fps">New frames per second (30 or 60)</param>
        public async Task ChangeFrameRateAsync(int fps)
        {
            if (!_isCapturing) return;

            FrameRate = fps;
            _frameInterval = 1000 / FrameRate;

            // Restart timer with new interval
            if (_captureTimer != null)
            {
                await _captureTimer.DisposeAsync();
                _captureTimer = new Timer(CaptureFrame, null, 0, _frameInterval);
            }
        }

        /// <summary>
        /// Captures a single frame for LLM analysis
        /// </summary>
        /// <returns>Bitmap of the current frame</returns>
        public async Task<Bitmap?> CaptureFrameForAnalysisAsync()
        {
            if (!_isCapturing || !IsWindow(_windowHandle))
            {
                return null;
            }

            return await Task.Run(() => CaptureWindowBitmap(_windowHandle));
        }

        private void CaptureFrame(object? state)
        {
            if (!_isCapturing || !IsWindow(_windowHandle))
            {
                _dispatcherQueue.TryEnqueue(() =>
                {
                    ErrorOccurred?.Invoke(this, "Window is no longer valid");
                });
                return;
            }

            try
            {
                var bitmap = CaptureWindowBitmap(_windowHandle);
                if (bitmap != null)
                {
                    // Convert to BitmapImage on UI thread
                    _dispatcherQueue.TryEnqueue(async () =>
                    {
                        try
                        {
                            var bitmapImage = await ConvertToBitmapImageAsync(bitmap);
                            bitmap.Dispose();
                            FrameCaptured?.Invoke(this, bitmapImage);
                        }
                        catch (Exception ex)
                        {
                            ErrorOccurred?.Invoke(this, $"Failed to convert frame: {ex.Message}");
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                _dispatcherQueue.TryEnqueue(() =>
                {
                    ErrorOccurred?.Invoke(this, $"Capture error: {ex.Message}");
                });
            }
        }

        private Bitmap? CaptureWindowBitmap(IntPtr windowHandle)
        {
            try
            {
                // Get extended window bounds
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

                // Get window device context
                IntPtr windowDc = GetWindowDC(windowHandle);
                if (windowDc == IntPtr.Zero)
                    return null;

                // Create memory DC and bitmap
                IntPtr memoryDc = CreateCompatibleDC(windowDc);
                IntPtr bitmap = CreateCompatibleBitmap(windowDc, width, height);
                IntPtr oldBitmap = SelectObject(memoryDc, bitmap);

                // Capture using PrintWindow
                bool success = PrintWindow(windowHandle, memoryDc, PW_RENDERFULLCONTENT);

                Bitmap? result = null;
                if (success)
                {
                    result = Image.FromHbitmap(bitmap);
                }

                // Cleanup
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

        private async Task<BitmapImage> ConvertToBitmapImageAsync(Bitmap bitmap)
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

        public void Dispose()
        {
            _ = StopCaptureAsync();
        }
    }
}
