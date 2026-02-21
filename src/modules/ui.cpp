/*
  Saka Studio & Engineering

  User interface functions for window title, status bar, and control layout.
  Manages window resizing, status bar parts, and UI state synchronization.
*/

#include "ui.h"
#include "core/globals.h"
#include "lang/lang.h"
#include "editor.h"
#include "file.h"
#include <commctrl.h>
#include <shlwapi.h>
#include <algorithm>

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
    g_statusTexts[0] = buf;
    if (g_state.largeFileMode)
        g_statusTexts[0] += lang.statusLargeFile;
    g_statusTexts[1] = GetEncodingName(g_state.encoding);
    g_statusTexts[2] = GetLineEndingName(g_state.lineEnding);
    wsprintfW(buf, L" %d%% ", g_state.zoomLevel);
    g_statusTexts[3] = buf;
    for (int i = 0; i < 4; i++)
        SendMessageW(g_hwndStatus, SB_SETTEXTW, i | SBT_NOBORDERS, reinterpret_cast<LPARAM>(g_statusTexts[i].c_str()));
    InvalidateRect(g_hwndStatus, nullptr, TRUE);
}

void SetupStatusBarParts()
{
    RECT rc;
    GetClientRect(g_hwndMain, &rc);
    HDC hdc = GetDC(g_hwndStatus);
    HFONT hFont = reinterpret_cast<HFONT>(SendMessageW(g_hwndStatus, WM_GETFONT, 0, 0));
    HFONT old = reinterpret_cast<HFONT>(SelectObject(hdc, hFont ? hFont : reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT))));
    auto textW = [&](const wchar_t *s)
    {
        SIZE sz{};
        GetTextExtentPoint32W(hdc, s, static_cast<int>(wcslen(s)), &sz);
        return sz.cx + 24;
    };
    int wZoom = textW(L" 500% ");
    int wLE = textW(L" Windows (CRLF) ");
    int wEnc = textW(L" UTF-8 with BOM ");
    SelectObject(hdc, old);
    ReleaseDC(g_hwndStatus, hdc);
    int w = rc.right;
    int parts[4] = {w - (wEnc + wLE + wZoom), w - (wLE + wZoom), w - wZoom, -1};
    SendMessageW(g_hwndStatus, SB_SETPARTS, 4, reinterpret_cast<LPARAM>(parts));
}

void ResizeControls()
{
    RECT rc;
    GetClientRect(g_hwndMain, &rc);
    int tabsH = 0;
    int editorTop = 0;
    if (g_hwndTabs && g_state.useTabs)
    {
        tabsH = std::max(28, ScaleMainPx(32));
        ShowWindow(g_hwndTabs, SW_SHOW);
        MoveWindow(g_hwndTabs, 0, 0, rc.right, tabsH, TRUE);

        RECT displayRect{0, 0, rc.right, tabsH};
        // Anchor editor to tab display area so active tab can visually attach to page.
        TabCtrl_AdjustRect(g_hwndTabs, FALSE, &displayRect);
        editorTop = std::max(0, static_cast<int>(displayRect.top) - 1);
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
    MoveWindow(g_hwndEditor, 0, editorTop, rc.right, rc.bottom - statusH - editorTop, TRUE);
    ApplyEditorViewportPadding();
    SetupStatusBarParts();
}
