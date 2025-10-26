using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;

namespace VNC;

/// <summary>
/// A custom grid splitter control that allows resizing columns or rows.
/// </summary>
public class GridSplitter : Control
{
    private bool _isDragging = false;
    private double _dragStartX;
    private GridLength _leftColumnStartWidth;
    private GridLength _rightColumnStartWidth;
    private ColumnDefinition? _leftColumn;
    private ColumnDefinition? _rightColumn;

    public GridSplitter()
    {
        DefaultStyleKey = typeof(GridSplitter);
        
        // Set the horizontal resize cursor
        ProtectedCursor = InputSystemCursor.Create(InputSystemCursorShape.SizeWestEast);
        
        // Set default styling
        Background = (Brush)Application.Current.Resources["SystemControlForegroundBaseLowBrush"];
        Width = 5;
        HorizontalAlignment = HorizontalAlignment.Stretch;
        VerticalAlignment = VerticalAlignment.Stretch;
        
        // Attach event handlers
        PointerEntered += OnPointerEntered;
        PointerExited += OnPointerExited;
        PointerPressed += OnPointerPressed;
        PointerMoved += OnPointerMoved;
        PointerReleased += OnPointerReleased;
        ManipulationMode = ManipulationModes.All;
        
        // Force the control to be visible
        Opacity = 1.0;
    }

    /// <summary>
    /// Gets or sets the left column to resize (required).
    /// </summary>
    public ColumnDefinition? LeftColumn
    {
        get => _leftColumn;
        set => _leftColumn = value;
    }

    /// <summary>
    /// Gets or sets the right column to resize (required).
    /// </summary>
    public ColumnDefinition? RightColumn
    {
        get => _rightColumn;
        set => _rightColumn = value;
    }

    /// <summary>
    /// Gets or sets the minimum column width in pixels.
    /// </summary>
    public double MinColumnWidth { get; set; } = 200;

    private void OnPointerEntered(object sender, PointerRoutedEventArgs e)
    {
        // Highlight the splitter on hover
        Background = (Brush)Application.Current.Resources["SystemControlForegroundBaseMediumBrush"];
    }

    private void OnPointerExited(object sender, PointerRoutedEventArgs e)
    {
        if (!_isDragging)
        {
            // Reset background
            Background = (Brush)Application.Current.Resources["SystemControlForegroundBaseLowBrush"];
        }
    }

    private void OnPointerPressed(object sender, PointerRoutedEventArgs e)
    {
        if (LeftColumn == null || RightColumn == null)
            return;

        _isDragging = true;
        
        // Get parent grid
        if (Parent is Grid parentGrid)
        {
            var position = e.GetCurrentPoint(parentGrid).Position;
            _dragStartX = position.X;
            
            // Store the starting widths
            _leftColumnStartWidth = LeftColumn.Width;
            _rightColumnStartWidth = RightColumn.Width;
            
            // Capture pointer for smooth dragging
            CapturePointer(e.Pointer);
            e.Handled = true;
        }
    }

    private void OnPointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (!_isDragging || LeftColumn == null || RightColumn == null)
            return;

        if (Parent is Grid parentGrid)
        {
            var position = e.GetCurrentPoint(parentGrid).Position;
            var deltaX = position.X - _dragStartX;
            
            // Get current actual widths
            var leftActualWidth = LeftColumn.ActualWidth;
            var rightActualWidth = RightColumn.ActualWidth;
            
            // Calculate new widths
            var newLeftWidth = leftActualWidth + deltaX;
            var newRightWidth = rightActualWidth - deltaX;
            
            // Enforce minimum widths
            if (newLeftWidth >= MinColumnWidth && newRightWidth >= MinColumnWidth)
            {
                // Update column widths using star sizing to maintain proportions
                var totalWidth = newLeftWidth + newRightWidth;
                var leftStar = newLeftWidth / totalWidth * 4.0; // Total stars = 4 (3* + 1*)
                var rightStar = newRightWidth / totalWidth * 4.0;
                
                LeftColumn.Width = new GridLength(leftStar, GridUnitType.Star);
                RightColumn.Width = new GridLength(rightStar, GridUnitType.Star);
                
                // Update start position for next move
                _dragStartX = position.X;
            }
            
            e.Handled = true;
        }
    }

    private void OnPointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (_isDragging)
        {
            _isDragging = false;
            ReleasePointerCapture(e.Pointer);
            e.Handled = true;
        }
    }
}

