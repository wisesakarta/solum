/*
  Technical Standard

  Centralized desktop design-system tokens for visual consistency.
*/

#pragma once

#include <windows.h>

namespace DesignSystem
{
inline constexpr wchar_t kUiFontPrimary[] = L"Akkurat Mono LL";
inline constexpr wchar_t kUiFontFallback[] = L"Consolas";

inline constexpr int kChromeFontPointSize = 9;
inline constexpr int kChromeBandHeightPx = 34;
inline constexpr int kChromeStrokePx = 1;
inline constexpr int kMenuTextPaddingHPx = 0;
inline constexpr int kMenuTextPaddingVPx = 1;

inline constexpr int kCommandBarPaddingHPx = 10;
inline constexpr int kCommandBarPaddingVPx = 4;
inline constexpr int kCommandBarIndentPx = 0;
inline constexpr int kCommandBarHoverInsetPx = 0;

inline constexpr int kTabTextPaddingHPx = 14;
inline constexpr int kTabSeamStrokePx = 1;
inline constexpr int kTabSeparatorInsetYPx = 6;
inline constexpr int kTabSeparatorAlphaPct = 8;
inline constexpr int kTabInnerPaddingHPx = 16;
inline constexpr int kTabInnerPaddingVPx = 6;
inline constexpr int kTabFixedWidthPx = 180;

inline float GetDpiScale(HWND hwnd = nullptr)
{
    const HWND ref = hwnd ? hwnd : GetDesktopWindow();
    HDC hdc = GetDC(ref);
    if (!hdc)
        return 1.0f;
    const float scale = static_cast<float>(GetDeviceCaps(hdc, LOGPIXELSX)) / 96.0f;
    ReleaseDC(ref, hdc);
    return scale > 0.0f ? scale : 1.0f;
}

inline int ScalePx(int logicalPx, HWND hwnd = nullptr)
{
    return static_cast<int>(logicalPx * GetDpiScale(hwnd) + 0.5f);
}

// Color Tokens (Mobile Parity)
namespace Color
{
inline constexpr unsigned long kBlack = 0x000000;        // #000000
inline constexpr unsigned long kDarkBg = 0x121212;       // #121212 (Slate Deep)
inline constexpr unsigned long kDarkInk = 0xD6D6D6;      // #D6D6D6 (Soft Silver)
inline constexpr unsigned long kDarkEdge = 0x8B8B8B;     // #D6D6D6 @ 0.65
inline constexpr unsigned long kDarkMuted = 0xABABAB;    // #D6D6D6 @ 0.80

inline constexpr unsigned long kLightBg = 0xFBFDFD;      // #FDFDFB (Studio Bone)
inline constexpr unsigned long kLightInk = 0x1A1A1A;     // #1A1A1A (Ink Dark)
inline constexpr unsigned long kLightEdge = 0x8F8F8F;     // #1A1A1A @ 0.62
inline constexpr unsigned long kLightMuted = 0x4D4D4D;   // #1A1A1A @ 0.76

inline constexpr unsigned long kAccent = 0x0050FF;       // #FF5000 (Safety Orange - BGR)
}
}
