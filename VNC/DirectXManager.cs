using System;
using Vortice.Direct2D1;
using Vortice.Direct3D;
using Vortice.Direct3D11;
using Vortice.DXGI;

namespace VNC;

public class DirectXManager : IDisposable
{
    private ID2D1Factory _d2dFactory;
    private ID2D1Factory1 _d2dFactory1;
    private ID3D11Device _d3dDevice;
    private ID3D11DeviceContext _d3dDeviceContext;
    private IDXGIDevice1 _dxgiDevice;
    private ID2D1DeviceContext _d2dDeviceContext;
    private IDXGISwapChain1 _swapChain;

    public IDXGISwapChain1 SwapChain => _swapChain;

    /// <summary>
    /// Initializes the DirectX manager and creates all required resources.
    /// </summary>
    public void Initialize()
    {
        CreateD2D1Factory();
        CreateDeviceContext();
        CreateSwapChain(IntPtr.Zero);
    }

    /// <summary>
    /// Creates Direct2D factory instances for rendering 2D graphics.
    /// </summary>
    private void CreateD2D1Factory()
    {
        var options = DebugLevel.Information; // For debugging
        _d2dFactory = D2D1.D2D1CreateFactory<ID2D1Factory>(FactoryType.SingleThreaded, options);
        _d2dFactory1 = _d2dFactory.QueryInterface<ID2D1Factory1>();
    }

    /// <summary>
    /// Creates Direct3D 11 device and Direct2D device context for hardware-accelerated rendering.
    /// </summary>
    private void CreateDeviceContext()
    {
        var creationFlags = DeviceCreationFlags.BgraSupport;
        creationFlags |= DeviceCreationFlags.Debug; // For debugging

        Vortice.Direct3D.FeatureLevel[] featureLevels = new Vortice.Direct3D.FeatureLevel[]
        {
            Vortice.Direct3D.FeatureLevel.Level_11_1,
            Vortice.Direct3D.FeatureLevel.Level_11_0,
            Vortice.Direct3D.FeatureLevel.Level_10_1,
            Vortice.Direct3D.FeatureLevel.Level_10_0,
            Vortice.Direct3D.FeatureLevel.Level_9_3,
            Vortice.Direct3D.FeatureLevel.Level_9_2,
            Vortice.Direct3D.FeatureLevel.Level_9_1
        };

        var result = D3D11.D3D11CreateDevice(
            null,
            DriverType.Hardware,
            creationFlags,
            featureLevels,
            out _d3dDevice,
            out Vortice.Direct3D.FeatureLevel selectedFeatureLevel,
            out _d3dDeviceContext);

        if (result.Success && _d3dDevice != null)
        {
            _dxgiDevice = _d3dDevice.QueryInterface<IDXGIDevice1>();
            if (_d2dFactory1 != null)
            {
                var d2dDevice = _d2dFactory1.CreateDevice(_dxgiDevice);
                if (d2dDevice != null)
                {
                    _d2dDeviceContext = d2dDevice.CreateDeviceContext(DeviceContextOptions.None);
                    d2dDevice.Dispose();
                }
            }
        }
    }

    /// <summary>
    /// Creates a DXGI swap chain for rendering, either for a window handle or for composition with SwapChainPanel.
    /// </summary>
    private void CreateSwapChain(IntPtr hWnd)
    {
        var swapChainDesc = new SwapChainDescription1
        {
            Width = 1,
            Height = 1,
            Format = Format.B8G8R8A8_UNorm,
            Stereo = false,
            SampleDescription = new SampleDescription(1, 0),
            BufferUsage = Vortice.DXGI.Usage.RenderTargetOutput,
            BufferCount = 2,
            Scaling = (hWnd != IntPtr.Zero) ? Scaling.None : Scaling.Stretch,
            SwapEffect = SwapEffect.FlipSequential,
            Flags = SwapChainFlags.None
        };

        var adapter = _dxgiDevice.GetAdapter();
        if (adapter != null)
        {
            using (var factory = adapter.GetParent<IDXGIFactory2>())
            {
                if (factory != null)
                {
                    if (hWnd != IntPtr.Zero)
                    {
                        _swapChain = factory.CreateSwapChainForHwnd(_d3dDevice, hWnd, swapChainDesc);
                    }
                    else
                    {
                        _swapChain = factory.CreateSwapChainForComposition(_d3dDevice, swapChainDesc);
                    }

                    _dxgiDevice.MaximumFrameLatency = 1;
                }
            }
            adapter.Dispose();
        }
    }

    /// <summary>
    /// Disposes all DirectX/Direct2D resources to prevent memory leaks.
    /// </summary>
    public void Dispose()
    {
        _d2dDeviceContext?.Dispose();
        _swapChain?.Dispose();
        _dxgiDevice?.Dispose();
        _d3dDeviceContext?.Dispose();
        _d3dDevice?.Dispose();
        _d2dFactory1?.Dispose();
        _d2dFactory?.Dispose();
    }
}
