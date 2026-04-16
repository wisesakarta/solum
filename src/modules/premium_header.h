/*
  Solum
  
  Premium header component using Direct2D for hardware-accelerated 
  visual effects and animations (Liquid Logo Reveal).
*/

#pragma once
#include "graphics_engine.h"
#include "../core/spring_solver.h"
#include <string>

namespace Premium
{
    class Header
    {
    public:
        Header() = default;
        ~Header();

        bool Initialize(Graphics::Engine* pEngine);
        void Update(float dt);
        void Render(const RECT& rect);
        void StartReveal();

    private:
        Graphics::Engine* m_pEngine = nullptr;
        Core::Spring m_revealSpring;
        ID2D1SolidColorBrush* m_pAccentBrush = nullptr;
        ID2D1SolidColorBrush* m_pTextBrush = nullptr;
        IDWriteTextFormat* m_pTextFormat = nullptr;
        
        bool m_isInitialized = false;
        void CreateResources();
    };
}
