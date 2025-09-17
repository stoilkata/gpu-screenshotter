#include <Winrt/Windows.Foundation.h>
#include <Winrt/Windows.Graphics.Capture.h>
#include <Windows.Graphics.Capture.Interop.h>
#include <Windows.Graphics.Directx.Direct3d11.Interop.h>

#include <d3d11.h>

#pragma comment(lib,"windowsapp.lib")

#define EXTERN extern "C" __declspec(dllexport)


namespace WinRT
{
    using namespace winrt::Windows::Graphics;
    using namespace winrt::Windows::Graphics::Capture;
    using namespace winrt::Windows::Graphics::DirectX;
    using namespace winrt::Windows::Graphics::DirectX::Direct3D11;

    struct __declspec(uuid("A9B3D012-3DF2-4EE3-B8D1-8695F457D3C1")) IDirect3DDxgiInterfaceAccess : ::IUnknown
    {
        virtual HRESULT __stdcall GetInterface(GUID const& id, void** object) = 0;
    };
}

namespace Utils
{
    static void GetTextureData(ID3D11Device* pD3DDevice, ID3D11Texture2D* texture, void (*func)(UINT, UINT, void*, UINT))
    {
        ID3D11DeviceContext* d3dContext = nullptr;
        pD3DDevice->GetImmediateContext(&d3dContext);

        // Create intermediate texture
        D3D11_TEXTURE2D_DESC intermediateTextureDesc;
        winrt::com_ptr<ID3D11Texture2D> intermidiateTexture = nullptr;
        {
            texture->GetDesc(&intermediateTextureDesc);

            intermediateTextureDesc.Usage = D3D11_USAGE_STAGING;
            intermediateTextureDesc.BindFlags = 0;
            intermediateTextureDesc.CPUAccessFlags = D3D11_CPU_ACCESS_READ;
            intermediateTextureDesc.MiscFlags = 0;

            winrt::check_hresult(pD3DDevice->CreateTexture2D(&intermediateTextureDesc, NULL, intermidiateTexture.put()));
        }

        // Copy actual texture to intermediate texture
        d3dContext->CopyResource(intermidiateTexture.get(), texture);

        D3D11_MAPPED_SUBRESOURCE resource;
        winrt::check_hresult(d3dContext->Map(intermidiateTexture.get(), NULL, D3D11_MAP_READ, 0, &resource));

        // Code
        func(intermediateTextureDesc.Width, intermediateTextureDesc.Height, resource.pData, resource.RowPitch);

        d3dContext->Unmap(intermidiateTexture.get(), NULL);
    }

#ifdef ENABLE_STORE
#include <roerrorapi.h>
#include <shlobj_core.h>
#include <dwmapi.h>

#pragma comment(lib,"Dwmapi.lib")
    static void ResourceToBitmap(D3D11_MAPPED_SUBRESOURCE* resource, LONG width, LONG height)
    {
        BITMAPINFO lBmpInfo;
        {
            // BMP 32 bpp
            ZeroMemory(&lBmpInfo, sizeof(BITMAPINFO));
            lBmpInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
            lBmpInfo.bmiHeader.biBitCount = 32;
            lBmpInfo.bmiHeader.biCompression = BI_RGB;
            lBmpInfo.bmiHeader.biWidth = width;
            lBmpInfo.bmiHeader.biHeight = height;
            lBmpInfo.bmiHeader.biPlanes = 1;
            lBmpInfo.bmiHeader.biSizeImage = width * height * 4;
        }

        BITMAP bitmap;
        bitmap.bmType = 0;
        bitmap.bmWidth = width;
        bitmap.bmHeight = height;
        bitmap.bmWidthBytes = resource->RowPitch;
        bitmap.bmPlanes = 1;
        bitmap.bmBitsPixel = 32;
        bitmap.bmBits = resource->pData;


        UINT lBmpRowPitch = lBmpInfo.bmiHeader.biWidth * 4;

        std::unique_ptr<BYTE> pBuf(new BYTE[lBmpInfo.bmiHeader.biSizeImage]);

        auto sptr = static_cast<BYTE*>(resource->pData);
        auto dptr = pBuf.get() + lBmpInfo.bmiHeader.biSizeImage - lBmpRowPitch;

        UINT lRowPitch = std::min<UINT>(lBmpRowPitch, resource->RowPitch);

        for (size_t h = 0; h < lBmpInfo.bmiHeader.biHeight; ++h)
        {
            memcpy_s(dptr, lBmpRowPitch, sptr, lRowPitch);
            sptr += resource->RowPitch;
            dptr -= lBmpRowPitch;
        }



        // Save bitmap buffer into the file ScreenShot.bmp
        WCHAR lMyDocPath[MAX_PATH]{};
        winrt::check_hresult(SHGetFolderPath(nullptr, CSIDL_PERSONAL, nullptr, SHGFP_TYPE_CURRENT, lMyDocPath));

        std::wstring lFilePath = L"ScreenShot.bmp";

        FILE* lfile = nullptr;

        if (auto lerr = _wfopen_s(&lfile, lFilePath.c_str(), L"wb"); lerr != 0)
            return;

        if (lfile != nullptr)
        {
            BITMAPFILEHEADER bmpFileHeader;

            bmpFileHeader.bfReserved1 = 0;
            bmpFileHeader.bfReserved2 = 0;
            bmpFileHeader.bfSize = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER) + lBmpInfo.bmiHeader.biSizeImage;
            bmpFileHeader.bfType = 'MB';
            bmpFileHeader.bfOffBits = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER);

            fwrite(&bmpFileHeader, sizeof(BITMAPFILEHEADER), 1, lfile);
            fwrite(&lBmpInfo.bmiHeader, sizeof(BITMAPINFOHEADER), 1, lfile);
            fwrite(pBuf.get(), lBmpInfo.bmiHeader.biSizeImage, 1, lfile);

            fclose(lfile);
        }
    }
#endif // ENABLE_STORE
}


struct WindowCaptureFactory {
    winrt::com_ptr<ID3D11Device> pD3DDevice;
    winrt::com_ptr<ID3D11DeviceContext> pD3DContext;

    WinRT::IDirect3DDevice pDirect3DDevice;
};

struct WindowCapture {
    WinRT::Direct3D11CaptureFramePool framePool = NULL;
    WinRT::GraphicsCaptureSession session = NULL;
    WinRT::Direct3D11CaptureFrame frame = NULL;

    winrt::com_ptr<ID3D11Texture2D> texture;
    WinRT::SizeInt32 texture_size;

    winrt::com_ptr<ID3D11Texture2D> texture_cpu;
    WinRT::SizeInt32 texture_cpu_size;

    D3D11_BOX crop_box;

    D3D11_MAPPED_SUBRESOURCE resource;

    int border_thickness = -1;
    int title_bar_height = -1;
};


typedef WindowCaptureFactory* HCAPTURE;
typedef WindowCapture* HWINDOWCAPTURE;

static HRESULT RefreshTexture(ID3D11Device* device, UINT width, UINT height, ID3D11Texture2D** intermidiateTexture) {

    D3D11_TEXTURE2D_DESC intermediateTextureDesc;
    intermediateTextureDesc.Width = width;
    intermediateTextureDesc.Height = height;
    intermediateTextureDesc.MipLevels = 1;
    intermediateTextureDesc.ArraySize = 1;
    intermediateTextureDesc.Format = DXGI_FORMAT_B8G8R8A8_UNORM;
    intermediateTextureDesc.SampleDesc.Count = 1;
    intermediateTextureDesc.SampleDesc.Quality = 0;
    intermediateTextureDesc.Usage = D3D11_USAGE_STAGING;
    intermediateTextureDesc.BindFlags = 0;
    intermediateTextureDesc.CPUAccessFlags = D3D11_CPU_ACCESS_READ;
    intermediateTextureDesc.MiscFlags = 0;

    return device->CreateTexture2D(&intermediateTextureDesc, NULL, intermidiateTexture);
}
static HRESULT RefreshTexture(ID3D11Device* device, WinRT::SizeInt32 size, ID3D11Texture2D** intermidiateTexture) {
    return RefreshTexture(device, size.Width, size.Height, intermidiateTexture);
}

EXTERN BOOL Initalize(_Out_ HCAPTURE* ppCapture) {

    // Initialize COM
    //winrt::init_apartment();

    if (SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT_SYSTEM_AWARE) == NULL)
        return FALSE;

    if (!WinRT::GraphicsCaptureSession::IsSupported()) {
        *ppCapture = nullptr;
        return FALSE;
    }

    *ppCapture = new WindowCaptureFactory;


    auto pCature = *ppCapture;

    // Create Direct 3D Device & Context
    winrt::check_hresult(D3D11CreateDevice(nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr, D3D11_CREATE_DEVICE_BGRA_SUPPORT, nullptr, 0, D3D11_SDK_VERSION, pCature->pD3DDevice.put(), nullptr, pCature->pD3DContext.put()));

    // Extract IDirect3DDevice
    {
        const auto dxgiDevice = pCature->pD3DDevice.as<IDXGIDevice>();
        winrt::com_ptr<IInspectable> inspectableSurface;
        winrt::check_hresult(CreateDirect3D11DeviceFromDXGIDevice(dxgiDevice.get(), inspectableSurface.put()));
        inspectableSurface.as<WinRT::IDirect3DDevice>(pCature->pDirect3DDevice);
    }

    return TRUE;
}

EXTERN double WHATDPI(HWINDOWCAPTURE hCapture) {
    return hCapture->title_bar_height;
}

EXTERN HRESULT StartWindowCapture(_In_ HCAPTURE hCapture, _In_ HWND hwnd, _Out_ HWINDOWCAPTURE* ppWinCapture) {
    
    HRESULT hr;

    // Get CaptureItem from HWND
    WinRT::GraphicsCaptureItem captureItem = NULL;
    {
        const auto activationFactory = winrt::get_activation_factory<WinRT::GraphicsCaptureItem>();
        auto interopFactory = activationFactory.as<IGraphicsCaptureItemInterop>();
        if (FAILED(hr = interopFactory->CreateForWindow(hwnd, winrt::guid_of<ABI::Windows::Graphics::Capture::IGraphicsCaptureItem>(), reinterpret_cast<void**>(winrt::put_abi(captureItem)))))
            return hr;
    }

    *ppWinCapture = new WindowCapture;

    {
        auto hWinCapture = *ppWinCapture;


        hWinCapture->texture_size = captureItem.Size();

        if ((hr = RefreshTexture(hCapture->pD3DDevice.get(), hWinCapture->texture_size, hWinCapture->texture_cpu.put())) != S_OK)
            return hr;

        {
            hWinCapture->title_bar_height = GetSystemMetrics(SM_CYFRAME) + GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CXPADDEDBORDER);
            hWinCapture->border_thickness = GetSystemMetrics(SM_CYBORDER);

            hWinCapture->texture_cpu_size.Width = hWinCapture->texture_size.Width - hWinCapture->border_thickness + hWinCapture->border_thickness;
            hWinCapture->texture_cpu_size.Height = hWinCapture->texture_size.Height - hWinCapture->title_bar_height + hWinCapture->border_thickness;

            hWinCapture->crop_box.left = (hWinCapture->border_thickness);
            hWinCapture->crop_box.top = (hWinCapture->title_bar_height);
            hWinCapture->crop_box.front = 0;
            hWinCapture->crop_box.right = hWinCapture->border_thickness + hWinCapture->texture_cpu_size.Width;
            hWinCapture->crop_box.bottom = hWinCapture->title_bar_height + hWinCapture->texture_cpu_size.Height;
            hWinCapture->crop_box.back = 1;
        }


        hWinCapture->framePool = WinRT::Direct3D11CaptureFramePool::Create(hCapture->pDirect3DDevice, WinRT::DirectXPixelFormat::B8G8R8A8UIntNormalized, 4, hWinCapture->texture_size);
        hWinCapture->session = hWinCapture->framePool.CreateCaptureSession(captureItem);
        hWinCapture->session.IsBorderRequired(false);
        hWinCapture->session.IsCursorCaptureEnabled(false);
        hWinCapture->session.StartCapture();
    }
    
    return 0;
}

EXTERN WinRT::SizeInt32 GetCaptureSize(_In_ HWINDOWCAPTURE hWinCapture) {
    return hWinCapture->texture_size;
}

EXTERN HRESULT GetFrame(HCAPTURE hCapture, HWINDOWCAPTURE hWinCapture, void (*func)(UINT, UINT, void*, UINT))
{
    hWinCapture->frame = hWinCapture->framePool.TryGetNextFrame();
    if (hWinCapture->frame == nullptr) {
        return ERROR_NO_MORE_FILES;
    }

    HRESULT hr;

    winrt::com_ptr<WinRT::IDirect3DDxgiInterfaceAccess> access{ hWinCapture->frame.Surface().as<WinRT::IDirect3DDxgiInterfaceAccess>() };
    if (FAILED(hr = access->GetInterface(__uuidof(ID3D11Texture2D), hWinCapture->texture.put_void())))
        return hr;
    
    WinRT::SizeInt32 textureSize = hWinCapture->frame.ContentSize();
    if (hWinCapture->texture_cpu == nullptr || textureSize != hWinCapture->texture_size)
    {
        hWinCapture->texture_cpu_size.Width = textureSize.Width - (hWinCapture->border_thickness + hWinCapture->border_thickness);
        hWinCapture->texture_cpu_size.Height = textureSize.Height - (hWinCapture->title_bar_height + hWinCapture->border_thickness);
        
        if (FAILED(hr = RefreshTexture(hCapture->pD3DDevice.get(), hWinCapture->texture_cpu_size, hWinCapture->texture_cpu.put())))
            return hr;


        hWinCapture->crop_box.left = hWinCapture->border_thickness;
        hWinCapture->crop_box.top = hWinCapture->title_bar_height;
        hWinCapture->crop_box.front = 0;
        hWinCapture->crop_box.right = hWinCapture->border_thickness + hWinCapture->texture_cpu_size.Width;
        hWinCapture->crop_box.bottom = hWinCapture->title_bar_height + hWinCapture->texture_cpu_size.Height;
        hWinCapture->crop_box.back = 1;

        hWinCapture->texture_size = textureSize;
    }

    // Copy captured texture to cpu texture
    hCapture->pD3DContext->CopySubresourceRegion(hWinCapture->texture_cpu.get(), 0, 0, 0, 0, hWinCapture->texture.get(), 0, &hWinCapture->crop_box);
    
    if (FAILED(hr = hCapture->pD3DContext->Map(hWinCapture->texture_cpu.get(), NULL, D3D11_MAP_READ, 0, &hWinCapture->resource))){
        return hr;
    }


    // Code
    func(hWinCapture->texture_cpu_size.Width, hWinCapture->texture_cpu_size.Height, hWinCapture->resource.pData, hWinCapture->resource.RowPitch);

    hCapture->pD3DContext->Unmap(hWinCapture->texture_cpu.get(), NULL);
    return ERROR_SUCCESS;
}

EXTERN void CloseWindowCapture(_In_ HWINDOWCAPTURE hWinCapture)
{
    hWinCapture->session.Close();
    hWinCapture->framePool.Close();
    //hWinCapture->frame.Close();

    delete hWinCapture;
}

EXTERN void Cleanup(_In_ HCAPTURE hCapture) {
    hCapture->pDirect3DDevice.Close();
    delete hCapture;
}