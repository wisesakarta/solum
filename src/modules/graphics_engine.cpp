#include "graphics_engine.h"

namespace Graphics
{
    bool Engine::Initialize(HWND hwnd)
    {
        HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, __uuidof(ID2D1Factory1), reinterpret_cast<void**>(&m_pFactory));
        if (FAILED(hr)) return false;

        RECT rc;
        GetClientRect(hwnd, &rc);
        
        hr = m_pFactory->CreateHwndRenderTarget(
            D2D1::RenderTargetProperties(),
            D2D1::HwndRenderTargetProperties(hwnd, D2D1::SizeU(rc.right - rc.left, rc.bottom - rc.top)),
            &m_pRenderTarget
        );
        if (FAILED(hr)) return false;

        m_pRenderTarget->QueryInterface(__uuidof(ID2D1DeviceContext), reinterpret_cast<void**>(&m_pDeviceContext));

        hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(&m_pWriteFactory));
        return SUCCEEDED(hr);
    }

    bool Engine::Resize(UINT width, UINT height)
    {
        if (!m_pRenderTarget)
            return false;
        if (width == 0 || height == 0)
            return true;
        
        // When resizing HwndRenderTarget, DeviceContext remains valid as it wraps the RT
        return SUCCEEDED(m_pRenderTarget->Resize(D2D1::SizeU(width, height)));
    }

    void Engine::Release()
    {
        if (m_pWriteFactory) { m_pWriteFactory->Release(); m_pWriteFactory = nullptr; }
        if (m_pDeviceContext) { m_pDeviceContext->Release(); m_pDeviceContext = nullptr; }
        if (m_pRenderTarget) { m_pRenderTarget->Release(); m_pRenderTarget = nullptr; }
        if (m_pFactory) { m_pFactory->Release(); m_pFactory = nullptr; }
    }
}
