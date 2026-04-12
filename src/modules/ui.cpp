/*
  Technical Standard

  User interface functions for window title, status bar, and control layout.
  Manages window resizing, status bar parts, and UI state synchronization.
*/

#include "ui.h"
#include "core/globals.h"
#include "design_system.h"
#include "tab_layout.h"
#include "lang/lang.h"
#include "editor.h"
#include "file.h"
#include "custom_scrollbar.h"
#include "selection_aura.h"
#include <commctrl.h>
#include <shlwapi.h>
#include <algorithm>

static void CommitStatusPartText(int partIndex, const std::wstring &text)
{
    SendMessageW(g_hwndStatus, SB_SETTEXTW, partIndex | SBT_NOBORDERS, reinterpret_cast<LPARAM>(text.c_str()));
}

static int ScaleMainPx(int px)
{
    HWND dpiSource = g_hwndMain ? g_hwndMain : g_hwndStatus;
    if (!dpiSource)
        return px;

    HDC hdc = GetDC(dpiSource);
    if (!hdc)
        return px;

    const int dpi = GetDeviceCaps(hdc, LOGPIXELSX);
    ReleaseDC(dpiSource, hdc);
    return MulDiv(px, dpi, 96);
}

void UpdateTitle()
{
    const auto &lang = GetLangStrings();
    std::wstring filename = g_state.filePath.empty() ? lang.untitled : PathFindFileNameW(g_state.filePath.c_str());
    if (g_state.largeFileMode)
        filename += lang.statusLargeFile;
    std::wstring title = (g_state.modified ? L"*" : L"") + filename + L" - " + lang.appName;
    SetWindowTextW(g_hwndMain, title.c_str());
}

void UpdateStatus()
{
    if (!g_state.showStatusBar)
    {
        ShowWindow(g_hwndStatus, SW_HIDE);
        return;
    }
    ShowWindow(g_hwndStatus, SW_SHOW);
    auto [line, col] = GetCursorPos();
    const auto &lang = GetLangStrings();
    wchar_t buf[256];
    swprintf(buf, sizeof(buf) / sizeof(wchar_t), L"%s%d%s%d ", lang.statusLn.c_str(), line, lang.statusCol.c_str(), col);
    std::wstring nextStatusTexts[4];
    nextStatusTexts[0] = buf;
    if (g_state.largeFileMode)
        nextStatusTexts[0] += lang.statusLargeFile;
    nextStatusTexts[1] = GetEncodingName(g_state.encoding);
    nextStatusTexts[2] = GetLineEndingName(g_state.lineEnding);
    wsprintfW(buf, L" %d%% ", g_state.zoomLevel);
    nextStatusTexts[3] = buf;

    for (int i = 0; i < 4; i++)
    {
        if (nextStatusTexts[i] == g_statusTexts[i])
            continue;
        g_statusTexts[i] = std::move(nextStatusTexts[i]);
        CommitStatusPartText(i, g_statusTexts[i]);
    }
}

void SetupStatusBarParts()
{
    RECT rc;
    GetClientRect(g_hwndMain, &rc);
    const auto &lang = GetLangStrings();
    HDC hdc = GetDC(g_hwndStatus);
    HFONT hFont = reinterpret_cast<HFONT>(SendMessageW(g_hwndStatus, WM_GETFONT, 0, 0));
    HFONT old = reinterpret_cast<HFONT>(SelectObject(hdc, hFont ? hFont : TabGetRegularFont()));
    auto textW = [&](const std::wstring &s)
    {
        SIZE sz{};
        GetTextExtentPoint32W(hdc, s.c_str(), static_cast<int>(s.size()), &sz);
        return sz.cx + ScaleMainPx(32); // 4 * 8dp grid for robust padding
    };

    const auto padded = [](const std::wstring &text)
    { return L" " + text + L" "; };

    const int wZoom = textW(L" 500% ");
    const int wLE = std::max({textW(padded(lang.lineEndingCRLF)),
                              textW(padded(lang.lineEndingLF)),
                              textW(padded(lang.lineEndingCR))});
    const int wEnc = std::max({textW(padded(lang.encodingUTF8)),
                               textW(padded(lang.encodingUTF8BOM)),
                               textW(padded(lang.encodingUTF16LE)),
                               textW(padded(lang.encodingUTF16BE)),
                               textW(padded(lang.encodingANSI))});
    SelectObject(hdc, old);
    ReleaseDC(g_hwndStatus, hdc);

    const int w = rc.right;
    const int right1 = std::max(0, w - (wEnc + wLE + wZoom));
    const int right2 = std::max(right1, w - (wLE + wZoom));
    const int right3 = std::max(right2, w - wZoom);
    int parts[4] = {right1, right2, right3, -1};
    SendMessageW(g_hwndStatus, SB_SETPARTS, 4, reinterpret_cast<LPARAM>(parts));

    // SB_SETPARTS can clear non-first panes; restore cached texts after relayout.
    for (int i = 0; i < 4; ++i)
        CommitStatusPartText(i, g_statusTexts[i]);
}

void ResizeControls()
{
    RECT rc;
    GetClientRect(g_hwndMain, &rc);
    int topOffset = 0;

    if (g_hwndCommandBar)
    {
        const int commandBarHeight = std::max(30, ScaleMainPx(DesignSystem::kChromeBandHeightPx));
        ShowWindow(g_hwndCommandBar, SW_SHOW);
        MoveWindow(g_hwndCommandBar, 0, 0, rc.right, commandBarHeight, TRUE);
        topOffset += commandBarHeight;
    }

    int tabsH = 0;
    int editorTop = topOffset;
    if (g_hwndTabs && g_state.useTabs)
    {
        tabsH = std::max(30, ScaleMainPx(DesignSystem::kChromeBandHeightPx));
        ShowWindow(g_hwndTabs, SW_SHOW);
        MoveWindow(g_hwndTabs, 0, topOffset, rc.right, tabsH, TRUE);
        editorTop = topOffset + tabsH;
    }
    else if (g_hwndTabs)
    {
        ShowWindow(g_hwndTabs, SW_HIDE);
    }
    int statusH = 0;
    if (g_state.showStatusBar)
    {
        ShowWindow(g_hwndStatus, SW_SHOW);
        SendMessageW(g_hwndStatus, WM_SIZE, 0, 0);
        RECT rs;
        GetWindowRect(g_hwndStatus, &rs);
        statusH = rs.bottom - rs.top;
    }
    else
        ShowWindow(g_hwndStatus, SW_HIDE);
    const int scrollbarWidth = ScaleMainPx(12);
    const int editorWidth = rc.right - scrollbarWidth;
    const int editorHeight = rc.bottom - statusH - editorTop;

    MoveWindow(g_hwndEditor, 0, editorTop, editorWidth, editorHeight, TRUE);
    
    if (g_hwndScrollbar)
    {
        MoveWindow(g_hwndScrollbar, editorWidth, editorTop, scrollbarWidth, editorHeight, TRUE);
        InvalidateRect(g_hwndScrollbar, nullptr, FALSE);
    }

    if (g_hwndSelectionAura)
    {
        MoveWindow(g_hwndSelectionAura, 0, editorTop, editorWidth, editorHeight, TRUE);
    }

    ApplyEditorViewportPadding();
    SetupStatusBarParts();
}
