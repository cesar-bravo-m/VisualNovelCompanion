using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Windows.Foundation;

namespace VNC;

public class TransparencyManager
{
    private readonly IntPtr _hwnd;
    private readonly Microsoft.UI.Windowing.AppWindow _appWindow;
    private readonly Microsoft.UI.Windowing.OverlappedPresenter _presenter;
    private readonly UIElement _rootElement;
    private readonly ToggleSwitch _clickThroughToggle;

    private int _dragStartX, _dragStartY, _windowStartX, _windowStartY;
    private bool _isDragging = false;

    public TransparencyManager(
        IntPtr hwnd,
        Microsoft.UI.Windowing.AppWindow appWindow,
        Microsoft.UI.Windowing.OverlappedPresenter presenter,
        UIElement rootElement,
        ToggleSwitch clickThroughToggle)
    {
        _hwnd = hwnd;
        _appWindow = appWindow;
        _presenter = presenter;
        _rootElement = rootElement;
        _clickThroughToggle = clickThroughToggle;

        AttachEventHandlers();
    }

    /// <summary>
    /// Configures the window for transparency and click-through support.
    /// </summary>
    public void InitializeTransparency()
    {
        // Enable layered window style
        long exStyle = NativeInterop.GetWindowLong(_hwnd, NativeInterop.GWL_EXSTYLE);
        if ((exStyle & NativeInterop.WS_EX_LAYERED) == 0)
        {
            NativeInterop.SetWindowLong(_hwnd, NativeInterop.GWL_EXSTYLE, (IntPtr)(exStyle | NativeInterop.WS_EX_LAYERED));

            // Determine background color based on system theme
            uint backgroundColor = GetSystemBackgroundColor();

            // Set layered window attributes for click-through
            NativeInterop.SetLayeredWindowAttributes(_hwnd, backgroundColor, 255, NativeInterop.LWA_COLORKEY);
        }

        // Remove window chrome (caption and thick frame)
        int style = (int)NativeInterop.GetWindowLong(_hwnd, NativeInterop.GWL_STYLE);
        style = style & ~(NativeInterop.WS_CAPTION | NativeInterop.WS_THICKFRAME);
        NativeInterop.SetWindowLong(_hwnd, NativeInterop.GWL_STYLE, (IntPtr)style);
    }

    /// <summary>
    /// Sets the opacity of the layered window to a percentage value from 0 to 100.
    /// </summary>
    public void SetOpacity(int opacity)
    {
        NativeInterop.SetLayeredWindowAttributes(_hwnd, 0, (byte)(255 * opacity / 100), NativeInterop.LWA_ALPHA);
    }

    /// <summary>
    /// Gets the system background color based on the current Windows theme.
    /// </summary>
    private uint GetSystemBackgroundColor()
    {
        int appsUseLightTheme = 0;
        string keyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
        
        using (RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
        {
            using (RegistryKey key = baseKey.OpenSubKey(keyPath, false))
            {
                if (key != null)
                {
                    appsUseLightTheme = (int)key.GetValue("AppsUseLightTheme", 0);
                }
            }
        }

        return appsUseLightTheme == 1
            ? (uint)System.Drawing.ColorTranslator.ToWin32(System.Drawing.Color.White)
            : (uint)System.Drawing.ColorTranslator.ToWin32(System.Drawing.Color.Black);
    }

    /// <summary>
    /// Attaches pointer event handlers for window dragging and click-through.
    /// </summary>
    private void AttachEventHandlers()
    {
        _rootElement.PointerPressed += OnPointerPressed;
        _rootElement.PointerMoved += OnPointerMoved;
        _rootElement.PointerReleased += OnPointerReleased;
    }

    /// <summary>
    /// Handles pointer press events for window dragging and click-through functionality.
    /// </summary>
    private void OnPointerPressed(object sender, PointerRoutedEventArgs e)
    {
        var properties = e.GetCurrentPoint((UIElement)sender).Properties;
        
        if (properties.IsLeftButtonPressed)
        {
            ((UIElement)sender).CapturePointer(e.Pointer);
            _windowStartX = _appWindow.Position.X;
            _windowStartY = _appWindow.Position.Y;
            
            Windows.Graphics.PointInt32 cursorPos;
            NativeInterop.GetCursorPos(out cursorPos);
            _dragStartX = cursorPos.X;
            _dragStartY = cursorPos.Y;

            // Check if we clicked on the SwapChainPanel for click-through
            if (IsClickOnSwapChainPanel(e, (UIElement)sender))
            {
                if (_clickThroughToggle?.IsOn == true)
                {
                    HandleClickThrough(cursorPos);
                    return;
                }
            }

            _isDragging = true;
        }
        else if (properties.IsRightButtonPressed)
        {
            // Right-click exits the application
            System.Threading.Thread.Sleep(200);
            Application.Current.Exit();
        }
    }

    /// <summary>
    /// Handles pointer movement to drag and reposition the window.
    /// </summary>
    private void OnPointerMoved(object sender, PointerRoutedEventArgs e)
    {
        var properties = e.GetCurrentPoint((UIElement)sender).Properties;
        if (properties.IsLeftButtonPressed && _isDragging)
        {
            Windows.Graphics.PointInt32 cursorPos;
            NativeInterop.GetCursorPos(out cursorPos);
            
            int newX = _windowStartX + (cursorPos.X - _dragStartX);
            int newY = _windowStartY + (cursorPos.Y - _dragStartY);
            
            _appWindow.Move(new Windows.Graphics.PointInt32(newX, newY));
            e.Handled = true;
        }
    }

    /// <summary>
    /// Handles pointer release events by stopping window movement.
    /// </summary>
    private void OnPointerReleased(object sender, PointerRoutedEventArgs e)
    {
        ((UIElement)sender).ReleasePointerCaptures();
        _isDragging = false;
    }

    /// <summary>
    /// Determines if the click occurred on the SwapChainPanel (transparent area).
    /// </summary>
    private bool IsClickOnSwapChainPanel(PointerRoutedEventArgs e, UIElement sender)
    {
        var pointerPoint = e.GetCurrentPoint(sender);
        Point clickPoint = new Point(pointerPoint.Position.X, pointerPoint.Position.Y);
        IEnumerable<UIElement> elementStack = VisualTreeHelper.FindElementsInHostCoordinates(clickPoint, sender);
        
        int index = 0;
        foreach (UIElement element in elementStack)
        {
            if (index == 0 && element.GetType() != typeof(Border))
                return false;
            if (index == 1 && element.GetType() != typeof(SwapChainPanel))
                return false;
            index++;
        }
        
        return index > 1;
    }

    /// <summary>
    /// Handles click-through functionality by forwarding clicks to the underlying window.
    /// </summary>
    private void HandleClickThrough(Windows.Graphics.PointInt32 cursorPos)
    {
        IntPtr underlyingWindow = NativeInterop.WindowFromPoint(cursorPos);

        StringBuilder windowClass = new StringBuilder(260);
        NativeInterop.GetClassName(underlyingWindow, windowClass, windowClass.Capacity);
        IntPtr parentWindow = NativeInterop.GetParent(underlyingWindow);
        StringBuilder parentClass = new StringBuilder(260);
        NativeInterop.GetClassName(parentWindow, parentClass, parentClass.Capacity);

        // Ensure window is always on top before click-through
        if (!_presenter.IsAlwaysOnTop)
        {
            _presenter.IsAlwaysOnTop = true;
        }

        // Briefly disable always-on-top to allow focus switch
        _presenter.IsAlwaysOnTop = false;
        System.Threading.Thread.Sleep(50);

        // Switch focus to underlying window
        NativeInterop.SwitchToThisWindow(underlyingWindow, true);
        System.Threading.Thread.Sleep(50);

        // Send mouse input to the underlying window
        NativeInterop.INPUT[] mouseInput = new NativeInterop.INPUT[1];
        mouseInput[0].type = NativeInterop.INPUT_MOUSE;
        mouseInput[0].inputUnion.mi.dwFlags = NativeInterop.MOUSEEVENTF_LEFTDOWN;
        NativeInterop.SendInput(1, mouseInput, Marshal.SizeOf(mouseInput[0]));
        mouseInput[0].inputUnion.mi.dwFlags = NativeInterop.MOUSEEVENTF_LEFTUP;
        NativeInterop.SendInput(1, mouseInput, Marshal.SizeOf(mouseInput[0]));

        // Don't minimize when clicking on desktop
        // (Commented out to keep overlay always visible)
        // if (windowClass.ToString() == "SysListView32" && parentClass.ToString() == "SHELLDLL_DefView")
        // {
        //     _presenter.Minimize();
        // }

        // Re-enable always-on-top to keep window on top
        _presenter.IsAlwaysOnTop = true;
    }
}

