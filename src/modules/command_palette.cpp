/*
  Technical Standard — Studio Command Palette

  A keyboard-driven command interface rendered via Direct2D + DirectWrite.
  Uses native design-system tokens for visual coherence.

  Key decisions (Karpathy Surgical Mode):
    - SetForegroundWindow for reliable keyboard capture.
    - WM_ACTIVATE(WA_INACTIVE) auto-dismisses on click-outside.
    - Message loop bypass in main.cpp — accelerators are skipped
      when the palette is visible, so WM_CHAR reaches us.
    - All colors derived from DesignSystem::Color via IsDarkMode().
    - Selection uses a thin 3px accent bar, not a full-row flood.
*/
#include "command_palette.h"
#include "design_system.h"
#include "theme.h"
#include "tab_layout.h"
#include "../core/globals.h"
#include "../lang/lang.h"
#include <algorithm>
#include <commctrl.h>

namespace UI {

// ── Layout constants (8dp grid) ──────────────────────────────────
static constexpr int   kPaletteWidth   = 560;
static constexpr float kQueryAreaH     = 52.0f;
static constexpr float kItemH          = 46.0f;
static constexpr float kPadH           = 20.0f;
static constexpr float kBorderW        = 1.0f;
static constexpr int   kMaxVisible     = 8;

std::vector<CommandAction> CommandPalette::s_commands;

bool CommandPalette::RegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc   = WindowProc;
    wc.hInstance      = hInstance;
    wc.lpszClassName  = L"TechnicalStandardCommandPalette";
    wc.hCursor        = LoadCursor(nullptr, IDC_IBEAM);
    wc.hbrBackground  = nullptr;
    wc.style          = CS_DROPSHADOW;
    return ::RegisterClassExW(&wc) != 0;
}

static LRESULT CALLBACK EditSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, [[maybe_unused]] UINT_PTR uIdSubclass, [[maybe_unused]] DWORD_PTR dwRefData)
{
    if (uMsg == WM_KEYDOWN) {
        if (wParam == VK_UP || wParam == VK_DOWN || wParam == VK_RETURN || wParam == VK_ESCAPE) {
            return SendMessageW(GetParent(hWnd), uMsg, wParam, lParam);
        }
    }
    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

HWND CommandPalette::Create(HWND hwndParent)
{
    HWND hwnd = CreateWindowExW(
        WS_EX_TOOLWINDOW,
        L"TechnicalStandardCommandPalette", nullptr,
        WS_POPUP | WS_CLIPCHILDREN,
        0, 0, 0, 0,
        hwndParent, nullptr, GetModuleHandle(nullptr), nullptr);

    if (hwnd) {
        State* s   = new State();
        s->hwndParent = hwndParent;
        s->engine.Initialize(hwnd);
        s->results = s_commands;
        
        s->hwndEdit = CreateWindowExW(
            0, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_LEFT | WS_TABSTOP,
            (int)kPadH, (int)((kQueryAreaH - 24) / 2), kPaletteWidth - (int)(kPadH * 2), 24,
            hwnd, (HMENU)101, GetModuleHandle(nullptr), nullptr);
        SetWindowSubclass(s->hwndEdit, EditSubclassProc, 1, 0);
        const auto &lang = GetLangStrings();
        SendMessageW(s->hwndEdit, EM_SETCUEBANNER, FALSE, (LPARAM)lang.palettePlaceholder.c_str());
        
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(s));
    }
    return hwnd;
}

void CommandPalette::Show(HWND hwnd)
{
    State* s = reinterpret_cast<State*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!s) return;

    s->query.clear();
    s->results        = s_commands;
    s->selectedIndex  = 0;
    s->isClosing      = false;

    int visibleItems = std::min((int)s->results.size(), kMaxVisible);
    if (visibleItems == 0) visibleItems = 1;
    int logicalHeight = (int)(kQueryAreaH + kBorderW + visibleItems * kItemH + 16);

    const float scale = DesignSystem::GetDpiScale(hwnd);
    const int scaledW = DesignSystem::ScalePx(kPaletteWidth, hwnd);
    const int scaledH = DesignSystem::ScalePx(logicalHeight, hwnd);

    if (s->hFontEdit) DeleteObject(s->hFontEdit);
    LOGFONTW lf;
    GetObjectW(TabGetRegularFont(), sizeof(LOGFONTW), &lf);
    lf.lfHeight = -DesignSystem::ScalePx(15, hwnd);
    s->hFontEdit = CreateFontIndirectW(&lf);
    SendMessageW(s->hwndEdit, WM_SETFONT, (WPARAM)s->hFontEdit, FALSE);
    const auto &lang = GetLangStrings();
    SendMessageW(s->hwndEdit, EM_SETCUEBANNER, FALSE, (LPARAM)lang.palettePlaceholder.c_str());
    
    SetWindowPos(s->hwndEdit, nullptr, 
        DesignSystem::ScalePx((int)kPadH, hwnd),
        DesignSystem::ScalePx((int)((kQueryAreaH - 24) / 2), hwnd),
        DesignSystem::ScalePx((int)(kPaletteWidth - kPadH * 2), hwnd),
        DesignSystem::ScalePx(24, hwnd),
        SWP_NOZORDER);

    SetWindowTextW(s->hwndEdit, L"");

    RECT rcParent;
    GetWindowRect(s->hwndParent, &rcParent);
    int x = rcParent.left + (rcParent.right - rcParent.left - scaledW) / 2;
    
    int y = rcParent.top + (int)(80 * scale);
    if (g_hwndEditor && IsWindowVisible(g_hwndEditor)) {
        RECT rcEditor;
        GetWindowRect(g_hwndEditor, &rcEditor);
        y = rcEditor.top;
    }

    SetWindowPos(hwnd, HWND_TOPMOST, x, y, scaledW, scaledH, SWP_SHOWWINDOW);
    SetForegroundWindow(hwnd);
    SetFocus(s->hwndEdit);
    InvalidateRect(hwnd, nullptr, FALSE);
}

void CommandPalette::Hide(HWND hwnd)
{
    State* s = reinterpret_cast<State*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    ShowWindow(hwnd, SW_HIDE);
    if (s && s->hwndParent) {
        SetForegroundWindow(s->hwndParent);
        SetFocus(s->hwndParent);
    }
}

bool CommandPalette::IsVisible(HWND hwnd)
{
    return hwnd && IsWindowVisible(hwnd);
}

void CommandPalette::AddCommand(const std::wstring& label, const std::wstring& desc, UINT id)
{
    s_commands.push_back({ label, desc, id });
}

void CommandPalette::PerformFuzzySearch(HWND hwnd, const std::wstring& query)
{
    State* s = reinterpret_cast<State*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!s) return;

    s->query = query;
    if (query.empty()) {
        s->results = s_commands;
    } else {
        s->results.clear();
        std::wstring lq = query;
        std::transform(lq.begin(), lq.end(), lq.begin(), ::towlower);

        for (auto& cmd : s_commands) {
            std::wstring ll = cmd.label;
            std::transform(ll.begin(), ll.end(), ll.begin(), ::towlower);

            size_t pos = ll.find(lq);
            if (pos != std::wstring::npos) {
                CommandAction r = cmd;
                r.score = (pos == 0) ? 100.0f : 50.0f;
                r.score += 10.0f / (float)ll.length();
                s->results.push_back(r);
            }
        }

        std::sort(s->results.begin(), s->results.end(),
            [](const CommandAction& a, const CommandAction& b) { return a.score > b.score; });
    }
    s->selectedIndex = 0;

    // Resize to fit
    int visibleItems = std::min((int)s->results.size(), kMaxVisible);
    if (visibleItems == 0) visibleItems = 1;
    int logicalHeight = (int)(kQueryAreaH + kBorderW + visibleItems * kItemH + 16);

    const int scaledW = DesignSystem::ScalePx(kPaletteWidth, hwnd);
    const int scaledH = DesignSystem::ScalePx(logicalHeight, hwnd);

    RECT rc;
    GetWindowRect(hwnd, &rc);
    SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, scaledW, scaledH, SWP_NOMOVE | SWP_NOZORDER);

    InvalidateRect(hwnd, nullptr, FALSE);
}

LRESULT CALLBACK CommandPalette::WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    State* s = reinterpret_cast<State*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    switch (msg) {

    case WM_ACTIVATE:
        if (LOWORD(wParam) == WA_INACTIVE) {
            Hide(hwnd);
            return 0;
        }
        break;


    case WM_CTLCOLOREDIT: {
        if (!s) break;
        HDC hdc = (HDC)wParam;
        const bool dark = IsDarkMode();
        COLORREF bg = dark ? DesignSystem::Color::kDarkBg : DesignSystem::Color::kLightBg;
        COLORREF fg = dark ? DesignSystem::Color::kDarkInk : DesignSystem::Color::kLightInk;
        SetTextColor(hdc, fg);
        SetBkColor(hdc, bg);
        if (s->hbrEdit) DeleteObject(s->hbrEdit);
        s->hbrEdit = CreateSolidBrush(bg);
        return (LRESULT)s->hbrEdit;
    }

    case WM_SIZE:
        if (s) {
            s->engine.Resize(LOWORD(lParam), HIWORD(lParam));
            SetWindowPos(s->hwndEdit, nullptr, 
                DesignSystem::ScalePx((int)kPadH, hwnd),
                DesignSystem::ScalePx((int)((kQueryAreaH - 24) / 2), hwnd),
                DesignSystem::ScalePx((int)(kPaletteWidth - kPadH * 2), hwnd),
                DesignSystem::ScalePx(24, hwnd),
                SWP_NOZORDER);
        }
        return 0;

    case WM_COMMAND:
        if (!s) break;
        if (LOWORD(wParam) == 101 && HIWORD(wParam) == EN_CHANGE) {
            int len = GetWindowTextLengthW(s->hwndEdit);
            std::wstring newQuery;
            if (len > 0) {
                newQuery.resize(len + 1);
                GetWindowTextW(s->hwndEdit, &newQuery[0], len + 1);
                newQuery.resize(len);
            }
            PerformFuzzySearch(hwnd, newQuery);
        }
        return 0;

    case WM_KEYDOWN:
        if (!s) break;
        switch (wParam) {
        case VK_ESCAPE:
            Hide(hwnd);
            return 0;
        case VK_RETURN:
            if (s->selectedIndex >= 0 && s->selectedIndex < (int)s->results.size()) {
                UINT cmdId = s->results[s->selectedIndex].commandId;
                Hide(hwnd);
                PostMessageW(s->hwndParent, WM_COMMAND, MAKEWPARAM(cmdId, 0), 0);
            }
            return 0;
        case VK_UP:
            if (!s->results.empty()) {
                s->selectedIndex = (s->selectedIndex > 0)
                    ? s->selectedIndex - 1
                    : (int)s->results.size() - 1;
                InvalidateRect(hwnd, nullptr, FALSE);
            }
            return 0;
        case VK_DOWN:
            if (!s->results.empty()) {
                s->selectedIndex = (s->selectedIndex < (int)s->results.size() - 1)
                    ? s->selectedIndex + 1 : 0;
                InvalidateRect(hwnd, nullptr, FALSE);
            }
            return 0;
        }
        break;



    case WM_PAINT: {
        if (!s) break;
        PAINTSTRUCT ps;
        BeginPaint(hwnd, &ps);

        auto ctx    = s->engine.GetDeviceContext();
        auto dwrite = s->engine.GetWriteFactory();
        if (!ctx || !dwrite) { EndPaint(hwnd, &ps); return 0; }

        ctx->BeginDraw();

        // ── Resolve design-system colors based on current theme ──
        const bool dark = IsDarkMode();
        const D2D1_COLOR_F bgColor    = Graphics::Engine::ColorToD2D(dark ? DesignSystem::Color::kDarkBg : DesignSystem::Color::kLightBg);
        const D2D1_COLOR_F inkColor   = Graphics::Engine::ColorToD2D(dark ? DesignSystem::Color::kDarkInk : DesignSystem::Color::kLightInk);
        const D2D1_COLOR_F mutedColor = Graphics::Engine::ColorToD2D(dark ? DesignSystem::Color::kDarkMuted : DesignSystem::Color::kLightMuted);
        const D2D1_COLOR_F edgeColor  = Graphics::Engine::ColorToD2D(dark ? DesignSystem::Color::kDarkEdge : DesignSystem::Color::kLightEdge, 0.5f);

        // Selected-row background: subtle Emil-style contrast
        const D2D1_COLOR_F selBg = dark
            ? D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.08f)
            : D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.04f);

        // ALWAYS USE LOGICAL DIMENSIONS inside D2D, do NOT use GetClientRect bounds 
        // to prevent UI layout double-scaling bug on High-DPI monitors.
        float w = (float)kPaletteWidth;
        int count = std::min((int)s->results.size(), kMaxVisible);
        if (count == 0) count = 1;
        float h = kQueryAreaH + kBorderW + count * kItemH + 16.0f;

        ctx->Clear(bgColor);

        // Create brushes
        ID2D1SolidColorBrush* brInk    = nullptr;
        ID2D1SolidColorBrush* brMuted  = nullptr;
        ID2D1SolidColorBrush* brEdge   = nullptr;
        ID2D1SolidColorBrush* brSelBg  = nullptr;
        ctx->CreateSolidColorBrush(inkColor,    &brInk);
        ctx->CreateSolidColorBrush(mutedColor,  &brMuted);
        ctx->CreateSolidColorBrush(edgeColor,   &brEdge);
        ctx->CreateSolidColorBrush(selBg,       &brSelBg);

        // Create text formats using design-system font
        IDWriteTextFormat* fmtQuery = nullptr;
        IDWriteTextFormat* fmtLabel = nullptr;
        IDWriteTextFormat* fmtDesc  = nullptr;
        dwrite->CreateTextFormat(DesignSystem::kUiFontPrimary, nullptr,
            DWRITE_FONT_WEIGHT_REGULAR, DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL, 15.0f, L"en-us", &fmtQuery);
        dwrite->CreateTextFormat(DesignSystem::kUiFontPrimary, nullptr,
            DWRITE_FONT_WEIGHT_MEDIUM, DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL, 13.0f, L"en-us", &fmtLabel);
        dwrite->CreateTextFormat(DesignSystem::kUiFontPrimary, nullptr,
            DWRITE_FONT_WEIGHT_REGULAR, DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL, 11.0f, L"en-us", &fmtDesc);

        // Fallback font from design system for environments without primary font.
        if (!fmtQuery)
            dwrite->CreateTextFormat(DesignSystem::kUiFontFallback, nullptr,
                DWRITE_FONT_WEIGHT_REGULAR, DWRITE_FONT_STYLE_NORMAL,
                DWRITE_FONT_STRETCH_NORMAL, 15.0f, L"en-us", &fmtQuery);
        if (!fmtLabel)
            dwrite->CreateTextFormat(DesignSystem::kUiFontFallback, nullptr,
                DWRITE_FONT_WEIGHT_MEDIUM, DWRITE_FONT_STYLE_NORMAL,
                DWRITE_FONT_STRETCH_NORMAL, 13.0f, L"en-us", &fmtLabel);
        if (!fmtDesc)
            dwrite->CreateTextFormat(DesignSystem::kUiFontFallback, nullptr,
                DWRITE_FONT_WEIGHT_REGULAR, DWRITE_FONT_STYLE_NORMAL,
                DWRITE_FONT_STRETCH_NORMAL, 11.0f, L"en-us", &fmtDesc);

        if (brInk && fmtQuery && fmtLabel) {

            // ── Query input field ───────────────────────────────
            // Background is cleared above. The Win32 EDIT control perfectly overlays over D2D
            // rendering native text input dynamically!

            // ── Separator ───────────────────────────────────────
            ctx->DrawLine(
                D2D1::Point2F(0, kQueryAreaH),
                D2D1::Point2F(w, kQueryAreaH),
                brEdge, kBorderW);

            // ── Results list ────────────────────────────────────
            float yBase = kQueryAreaH + 4.0f;
            int count = std::min((int)s->results.size(), kMaxVisible);

            if (count == 0 && !s->query.empty()) {
                D2D1_RECT_F noRes = D2D1::RectF(kPadH, yBase, w - kPadH, yBase + kItemH);
                if (fmtLabel) fmtLabel->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                const auto &lang = GetLangStrings();
                ctx->DrawTextW(lang.paletteNoResults.c_str(),
                    static_cast<UINT32>(lang.paletteNoResults.length()), fmtLabel, noRes, brMuted);
            }

            for (int i = 0; i < count; ++i) {
                auto& cmd = s->results[i];
                bool sel  = (i == s->selectedIndex);
                float top = yBase + i * kItemH;

                if (sel) {
                    // Subtle background (Emil style)
                    D2D1_RECT_F selRect = D2D1::RectF(8.0f, top, w - 8.0f, top + kItemH); // Slightly inset
                    ctx->FillRectangle(selRect, brSelBg);
                }

                // Command label
                D2D1_RECT_F labelRect = D2D1::RectF(kPadH, top + 6, w - kPadH, top + 24);
                ctx->DrawTextW(cmd.label.c_str(), (UINT32)cmd.label.length(),
                    fmtLabel, labelRect, brInk);

                // Description
                if (fmtDesc) {
                    D2D1_RECT_F descRect = D2D1::RectF(kPadH, top + 25, w - kPadH, top + kItemH - 2);
                    ctx->DrawTextW(cmd.description.c_str(), (UINT32)cmd.description.length(),
                        fmtDesc, descRect, brMuted);
                }
            }

            // ── Outer border ────────────────────────────────────
            ctx->DrawRectangle(D2D1::RectF(0, 0, w - 1, h - 1), brEdge, kBorderW);
        }

        // Release resources
        if (fmtQuery) fmtQuery->Release();
        if (fmtLabel) fmtLabel->Release();
        if (fmtDesc)  fmtDesc->Release();
        if (brInk)    brInk->Release();
        if (brMuted)  brMuted->Release();
        if (brEdge)   brEdge->Release();
        if (brSelBg)  brSelBg->Release();

        ctx->EndDraw();
        EndPaint(hwnd, &ps);
        return 0;
    }

    case WM_DESTROY:
        if (s) {
            if (s->hFontEdit) DeleteObject(s->hFontEdit);
            if (s->hbrEdit) DeleteObject(s->hbrEdit);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            delete s;
        }
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

} // namespace UI
