using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

namespace VNC;

/// <summary>
/// A custom border control for window resizing with appropriate cursor.
/// </summary>
public class ResizeBorder : Control
{
    public ResizeBorder()
    {
        DefaultStyleKey = typeof(ResizeBorder);
        Background = new Microsoft.UI.Xaml.Media.SolidColorBrush(Windows.UI.Color.FromArgb(0, 0, 0, 0)); // Transparent
        
        // Set resize cursor based on direction
        UpdateCursorFromDirection();
    }

    /// <summary>
    /// Gets or sets the resize direction for this border.
    /// </summary>
    public ResizeDirection Direction
    {
        get => (ResizeDirection)GetValue(DirectionProperty);
        set => SetValue(DirectionProperty, value);
    }

    public static readonly DependencyProperty DirectionProperty =
        DependencyProperty.Register(
            nameof(Direction),
            typeof(ResizeDirection),
            typeof(ResizeBorder),
            new PropertyMetadata(ResizeDirection.None, OnDirectionChanged));

    private static void OnDirectionChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is ResizeBorder border)
        {
            border.UpdateCursorFromDirection();
        }
    }

    private void UpdateCursorFromDirection()
    {
        ProtectedCursor = Direction switch
        {
            ResizeDirection.Left or ResizeDirection.Right =>
                InputSystemCursor.Create(InputSystemCursorShape.SizeWestEast),
            ResizeDirection.Top or ResizeDirection.Bottom =>
                InputSystemCursor.Create(InputSystemCursorShape.SizeNorthSouth),
            ResizeDirection.TopLeft or ResizeDirection.BottomRight =>
                InputSystemCursor.Create(InputSystemCursorShape.SizeNorthwestSoutheast),
            ResizeDirection.TopRight or ResizeDirection.BottomLeft =>
                InputSystemCursor.Create(InputSystemCursorShape.SizeNortheastSouthwest),
            _ => InputSystemCursor.Create(InputSystemCursorShape.Arrow)
        };
    }
}

/// <summary>
/// Defines the resize direction for a resize border.
/// </summary>
public enum ResizeDirection
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

