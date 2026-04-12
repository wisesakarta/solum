/*
  Technical Standard
  
  Direct2D graphics engine for high-performance native UI rendering.
  Handles resource lifecycle and provides primitives for hardware-accelerated visuals.
*/

#pragma once
#include <windows.h>
#include <d2d1_1.h>
#include <d2d1effects.h>
#include <dwrite.h>
#include <wincodec.h>

namespace Graphics
{
    class Engine
    {
    public:
        Engine() = default;
        ~Engine() { Release(); }

        bool Initialize(HWND hwnd);
        bool Resize(UINT width, UINT height);
        void Release();

        ID2D1HwndRenderTarget* GetRenderTarget() const { return m_pRenderTarget; }
        ID2D1DeviceContext* GetDeviceContext() const { return m_pDeviceContext; }
        ID2D1Factory1* GetFactory() const { return m_pFactory; }
        IDWriteFactory* GetWriteFactory() const { return m_pWriteFactory; }

        void BeginDraw() { if (m_pRenderTarget) m_pRenderTarget->BeginDraw(); }
        HRESULT EndDraw() { return m_pRenderTarget ? m_pRenderTarget->EndDraw() : S_OK; }
        void Clear(COLORREF color) { 
            if (m_pRenderTarget) {
                m_pRenderTarget->Clear(ColorToD2D(color));
            }
        }

        static D2D1_COLOR_F ColorToD2D(COLORREF color, float opacity = 1.0f) {
            return D2D1::ColorF(GetRValue(color) / 255.0f, GetGValue(color) / 255.0f, GetBValue(color) / 255.0f, opacity);
        }

        void DrawBlurredRect(const D2D1_RECT_F& rect, const D2D1_COLOR_F& color, float radius) {
            if (!m_pDeviceContext) return;
            
            ID2D1Effect* blurEffect = nullptr;
            m_pDeviceContext->CreateEffect(CLSID_D2D1GaussianBlur, &blurEffect);
            if (blurEffect) {
                // Create a bitmap to draw the rect
                ID2D1Bitmap1* bitmap = nullptr;
                D2D1_SIZE_U size = D2D1::SizeU((UINT32)(rect.right - rect.left + radius * 4), (UINT32)(rect.bottom - rect.top + radius * 4));
                m_pDeviceContext->CreateBitmap(size, nullptr, 0, D2D1::BitmapProperties1(D2D1_BITMAP_OPTIONS_TARGET, D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED)), &bitmap);
                
                if (bitmap) {
                    ID2D1Image* oldTarget = nullptr;
                    m_pDeviceContext->GetTarget(&oldTarget);
                    m_pDeviceContext->SetTarget(bitmap);
                    m_pDeviceContext->BeginDraw();
                    m_pDeviceContext->Clear(nullptr);
                    ID2D1SolidColorBrush* brush = nullptr;
                    m_pDeviceContext->CreateSolidColorBrush(color, &brush);
                    if (brush) {
                        m_pDeviceContext->FillRectangle(D2D1::RectF(radius*2, radius*2, rect.right - rect.left + radius*2, rect.bottom - rect.top + radius*2), brush);
                        brush->Release();
                    }
                    m_pDeviceContext->EndDraw();
                    
                    blurEffect->SetInput(0, bitmap);
                    blurEffect->SetValue(D2D1_GAUSSIANBLUR_PROP_STANDARD_DEVIATION, radius);
                    m_pDeviceContext->SetTarget(oldTarget);
                    m_pDeviceContext->DrawImage(blurEffect, D2D1::Point2F(rect.left - radius*2, rect.top - radius*2));
                    
                    bitmap->Release();
                    if (oldTarget) oldTarget->Release();
                }
                blurEffect->Release();
            }
        }

    private:
        ID2D1Factory1* m_pFactory = nullptr;
        ID2D1HwndRenderTarget* m_pRenderTarget = nullptr;
        ID2D1DeviceContext* m_pDeviceContext = nullptr;
        IDWriteFactory* m_pWriteFactory = nullptr;
    };
}
