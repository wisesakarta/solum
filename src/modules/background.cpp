/*
  Technical Standard

  Background image rendering with GDI+ support and multiple positioning modes.
  Supports tile, stretch, fit, fill, and nine anchor point positions.
*/

#include "background.h"
#include "core/file_dialog_filters.h"
#include "core/globals.h"
#include "theme.h"
#include "resource.h"
#include "settings.h"
#include "lang/lang.h"
#include "design_system.h"
#include "tab_layout.h"
#include <commdlg.h>
#include <algorithm>

namespace
{
int ScaleBackgroundPx(int px)
{
    return DesignSystem::ScalePx(px, g_hwndMain);
}
}

bool EnsureBackgroundGraphicsReady()
{
    if (g_gdiplusToken != 0)
        return true;

    Gdiplus::GdiplusStartupInput startupInput{};
    const Gdiplus::Status status = Gdiplus::GdiplusStartup(&g_gdiplusToken, &startupInput, nullptr);
    return status == Gdiplus::Ok && g_gdiplusToken != 0;
}

void ShutdownBackgroundGraphics()
{
    if (g_gdiplusToken == 0)
        return;
    Gdiplus::GdiplusShutdown(g_gdiplusToken);
    g_gdiplusToken = 0;
}

Gdiplus::Image* LoadImageFromResource(HMODULE hMod, WORD resId, const wchar_t* resType)
{
    HRSRC hRes = FindResourceW(hMod, MAKEINTRESOURCEW(resId), resType);
    if (!hRes) return nullptr;
    DWORD resSize = SizeofResource(hMod, hRes);
    HGLOBAL hResData = LoadResource(hMod, hRes);
    if (!hResData) return nullptr;
    void* pRes = LockResource(hResData);
    if (!pRes) return nullptr;

    HGLOBAL hBuffer = GlobalAlloc(GMEM_MOVEABLE, resSize);
    if (!hBuffer) return nullptr;
    void* pBuffer = GlobalLock(hBuffer);
    if (!pBuffer) { GlobalFree(hBuffer); return nullptr; }
    memcpy(pBuffer, pRes, resSize);
    GlobalUnlock(hBuffer);

    IStream* pStream = nullptr;
    if (CreateStreamOnHGlobal(hBuffer, TRUE, &pStream) != S_OK)
    {
        GlobalFree(hBuffer);
        return nullptr;
    }

    Gdiplus::Image* pImage = Gdiplus::Image::FromStream(pStream);
    pStream->Release();
    if (pImage && pImage->GetLastStatus() != Gdiplus::Ok)
    {
        delete pImage;
        return nullptr;
    }
    return pImage;
}

void LoadBackgroundImage(const std::wstring &path)
{
    if (g_bgImage)
    {
        delete g_bgImage;
        g_bgImage = nullptr;
    }

    if (!EnsureBackgroundGraphicsReady())
    {
        g_state.background.imagePath = path;
        g_state.background.enabled = false;
        SaveFontSettings();
        InvalidateRect(g_hwndEditor, nullptr, TRUE);
        return;
    }

    g_bgImage = Gdiplus::Image::FromFile(path.c_str());
    if (g_bgImage && g_bgImage->GetLastStatus() != Gdiplus::Ok)
    {
        delete g_bgImage;
        g_bgImage = nullptr;
    }
    g_state.background.imagePath = path;
    g_state.background.enabled = (g_bgImage != nullptr);
    SaveFontSettings();
    InvalidateRect(g_hwndEditor, nullptr, TRUE);
}

void PaintBackground(HDC hdc, const RECT &rc)
{
    if (!g_state.background.enabled || !g_bgImage)
        return;
    Gdiplus::Graphics graphics(hdc);
    graphics.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
    int imgW = g_bgImage->GetWidth();
    int imgH = g_bgImage->GetHeight();
    int winW = rc.right - rc.left;
    int winH = rc.bottom - rc.top;
    Gdiplus::ImageAttributes imgAttr;
    float opacity = g_state.background.opacity / 255.0f;
    Gdiplus::ColorMatrix colorMatrix = {{
        {1.0f, 0.0f, 0.0f, 0.0f, 0.0f},
        {0.0f, 1.0f, 0.0f, 0.0f, 0.0f},
        {0.0f, 0.0f, 1.0f, 0.0f, 0.0f},
        {0.0f, 0.0f, 0.0f, opacity, 0.0f},
        {0.0f, 0.0f, 0.0f, 0.0f, 1.0f},
    }};
    imgAttr.SetColorMatrix(&colorMatrix, Gdiplus::ColorMatrixFlagsDefault, Gdiplus::ColorAdjustTypeBitmap);
    auto drawImage = [&](int x, int y, int w, int h)
    {
        graphics.DrawImage(g_bgImage, Gdiplus::Rect(x, y, w, h), 0, 0, imgW, imgH, Gdiplus::UnitPixel, &imgAttr);
    };
    switch (g_state.background.position)
    {
    case BgPosition::TopLeft:
        drawImage(0, 0, imgW, imgH);
        break;
    case BgPosition::TopCenter:
        drawImage((winW - imgW) / 2, 0, imgW, imgH);
        break;
    case BgPosition::TopRight:
        drawImage(winW - imgW, 0, imgW, imgH);
        break;
    case BgPosition::CenterLeft:
        drawImage(0, (winH - imgH) / 2, imgW, imgH);
        break;
    case BgPosition::Center:
        drawImage((winW - imgW) / 2, (winH - imgH) / 2, imgW, imgH);
        break;
    case BgPosition::CenterRight:
        drawImage(winW - imgW, (winH - imgH) / 2, imgW, imgH);
        break;
    case BgPosition::BottomLeft:
        drawImage(0, winH - imgH, imgW, imgH);
        break;
    case BgPosition::BottomCenter:
        drawImage((winW - imgW) / 2, winH - imgH, imgW, imgH);
        break;
    case BgPosition::BottomRight:
        drawImage(winW - imgW, winH - imgH, imgW, imgH);
        break;
    case BgPosition::Tile:
        for (int y = 0; y < winH; y += imgH)
            for (int x = 0; x < winW; x += imgW)
                drawImage(x, y, imgW, imgH);
        break;
    case BgPosition::Stretch:
        drawImage(0, 0, winW, winH);
        break;
    case BgPosition::Fit:
    {
        float scale = (std::min)(static_cast<float>(winW) / imgW, static_cast<float>(winH) / imgH);
        int newW = static_cast<int>(imgW * scale);
        int newH = static_cast<int>(imgH * scale);
        drawImage((winW - newW) / 2, (winH - newH) / 2, newW, newH);
        break;
    }
    case BgPosition::Fill:
    {
        float scale = (std::max)(static_cast<float>(winW) / imgW, static_cast<float>(winH) / imgH);
        int newW = static_cast<int>(imgW * scale);
        int newH = static_cast<int>(imgH * scale);
        drawImage((winW - newW) / 2, (winH - newH) / 2, newW, newH);
        break;
    }
    }
}

void UpdateBackgroundBitmap(HWND hwnd)
{
    if (!g_state.background.enabled || !g_bgImage)
    {
        if (g_bgBitmap)
        {
            DeleteObject(g_bgBitmap);
            g_bgBitmap = nullptr;
        }
        return;
    }
    RECT rc;
    GetClientRect(hwnd, &rc);
    int w = rc.right - rc.left;
    int h = rc.bottom - rc.top;
    if (w <= 0 || h <= 0)
        return;
    if (g_bgBitmap && g_bgBitmapW == w && g_bgBitmapH == h)
        return;

    HDC hdcScreen = GetDC(hwnd);
    if (!hdcScreen)
        return;

    HDC hdcMem = CreateCompatibleDC(hdcScreen);
    if (!hdcMem)
    {
        ReleaseDC(hwnd, hdcScreen);
        return;
    }

    HBITMAP newBitmap = CreateCompatibleBitmap(hdcScreen, w, h);
    if (!newBitmap)
    {
        DeleteDC(hdcMem);
        ReleaseDC(hwnd, hdcScreen);
        return;
    }

    HBITMAP hOldBmp = reinterpret_cast<HBITMAP>(SelectObject(hdcMem, newBitmap));
    if (!hOldBmp || reinterpret_cast<HGDIOBJ>(hOldBmp) == HGDI_ERROR)
    {
        DeleteObject(newBitmap);
        DeleteDC(hdcMem);
        ReleaseDC(hwnd, hdcScreen);
        return;
    }

    COLORREF bgColor = ThemeColorEditorBackground(IsDarkMode());
    HBRUSH hBrush = CreateSolidBrush(bgColor);
    if (!hBrush)
    {
        SelectObject(hdcMem, hOldBmp);
        DeleteObject(newBitmap);
        DeleteDC(hdcMem);
        ReleaseDC(hwnd, hdcScreen);
        return;
    }
    FillRect(hdcMem, &rc, hBrush);
    DeleteObject(hBrush);
    PaintBackground(hdcMem, rc);
    SelectObject(hdcMem, hOldBmp);
    DeleteDC(hdcMem);
    ReleaseDC(hwnd, hdcScreen);

    if (g_bgBitmap)
    {
        DeleteObject(g_bgBitmap);
        g_bgBitmap = nullptr;
    }
    g_bgBitmap = newBitmap;
    g_bgBitmapW = w;
    g_bgBitmapH = h;
}

void SetBackgroundPosition(BgPosition pos)
{
    g_state.background.position = pos;
    HMENU hMenu = GetMenu(g_hwndMain);
    if (!hMenu)
        return;
    HMENU hViewMenu = GetSubMenu(hMenu, 3);
    if (!hViewMenu)
        return;
    HMENU hBgMenu = GetSubMenu(hViewMenu, 7);
    if (!hBgMenu)
        return;
    HMENU hPosMenu = GetSubMenu(hBgMenu, 4);
    if (!hPosMenu)
        return;
    const UINT commands[] = {
        IDM_VIEW_BG_POS_TOPLEFT,
        IDM_VIEW_BG_POS_TOPCENTER,
        IDM_VIEW_BG_POS_TOPRIGHT,
        IDM_VIEW_BG_POS_CENTERLEFT,
        IDM_VIEW_BG_POS_CENTER,
        IDM_VIEW_BG_POS_CENTERRIGHT,
        IDM_VIEW_BG_POS_BOTTOMLEFT,
        IDM_VIEW_BG_POS_BOTTOMCENTER,
        IDM_VIEW_BG_POS_BOTTOMRIGHT,
        IDM_VIEW_BG_POS_TILE,
        IDM_VIEW_BG_POS_STRETCH,
        IDM_VIEW_BG_POS_FIT,
        IDM_VIEW_BG_POS_FILL};
    for (UINT cmd : commands)
        CheckMenuItem(hPosMenu, cmd, MF_BYCOMMAND | MF_UNCHECKED);
    UINT targetCmd = IDM_VIEW_BG_POS_CENTER;
    switch (pos)
    {
    case BgPosition::TopLeft:
        targetCmd = IDM_VIEW_BG_POS_TOPLEFT;
        break;
    case BgPosition::TopCenter:
        targetCmd = IDM_VIEW_BG_POS_TOPCENTER;
        break;
    case BgPosition::TopRight:
        targetCmd = IDM_VIEW_BG_POS_TOPRIGHT;
        break;
    case BgPosition::CenterLeft:
        targetCmd = IDM_VIEW_BG_POS_CENTERLEFT;
        break;
    case BgPosition::Center:
        targetCmd = IDM_VIEW_BG_POS_CENTER;
        break;
    case BgPosition::CenterRight:
        targetCmd = IDM_VIEW_BG_POS_CENTERRIGHT;
        break;
    case BgPosition::BottomLeft:
        targetCmd = IDM_VIEW_BG_POS_BOTTOMLEFT;
        break;
    case BgPosition::BottomCenter:
        targetCmd = IDM_VIEW_BG_POS_BOTTOMCENTER;
        break;
    case BgPosition::BottomRight:
        targetCmd = IDM_VIEW_BG_POS_BOTTOMRIGHT;
        break;
    case BgPosition::Tile:
        targetCmd = IDM_VIEW_BG_POS_TILE;
        break;
    case BgPosition::Stretch:
        targetCmd = IDM_VIEW_BG_POS_STRETCH;
        break;
    case BgPosition::Fit:
        targetCmd = IDM_VIEW_BG_POS_FIT;
        break;
    case BgPosition::Fill:
        targetCmd = IDM_VIEW_BG_POS_FILL;
        break;
    }
    CheckMenuItem(hPosMenu, targetCmd, MF_BYCOMMAND | MF_CHECKED);
    if (g_bgBitmap)
    {
        DeleteObject(g_bgBitmap);
        g_bgBitmap = nullptr;
    }
    SaveFontSettings();
    InvalidateRect(g_hwndEditor, nullptr, TRUE);
}

void ViewSelectBackground()
{
    const auto &lang = GetLangStrings();
    const std::wstring filter = BuildImageFilesFilter(lang);
    wchar_t path[MAX_PATH] = {0};
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_hwndMain;
    ofn.lpstrFilter = filter.c_str();
    ofn.lpstrFile = path;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
    if (GetOpenFileNameW(&ofn))
        LoadBackgroundImage(path);
}

void ViewClearBackground()
{
    if (g_bgImage)
    {
        delete g_bgImage;
        g_bgImage = nullptr;
    }
    if (g_bgBitmap)
    {
        DeleteObject(g_bgBitmap);
        g_bgBitmap = nullptr;
    }
    g_state.background.enabled = false;
    g_state.background.imagePath.clear();
    ShutdownBackgroundGraphics();
    SaveFontSettings();
    InvalidateRect(g_hwndEditor, nullptr, TRUE);
}

static INT_PTR CALLBACK OpacityDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM)
{
    static HWND hEdit = nullptr;
    switch (msg)
    {
    case WM_INITDIALOG:
    {
        const auto &lang = GetLangStrings();
        const int margin = ScaleBackgroundPx(16);
        const int gap = ScaleBackgroundPx(12);
        const int labelW = ScaleBackgroundPx(110);
        const int labelH = ScaleBackgroundPx(20);
        const int inputW = ScaleBackgroundPx(88);
        const int inputH = ScaleBackgroundPx(24);
        const int topPad = ScaleBackgroundPx(20);
        const int buttonW = ScaleBackgroundPx(84);
        const int buttonH = ScaleBackgroundPx(30);
        const int buttonGap = ScaleBackgroundPx(10);

        RECT rcClient{};
        GetClientRect(hDlg, &rcClient);

        const int rowY = topPad;
        const int labelY = rowY + std::max(0, (inputH - labelH) / 2);
        const int inputX = margin + labelW + gap;
        const int buttonY = rcClient.bottom - ScaleBackgroundPx(20) - buttonH;
        const int cancelX = rcClient.right - margin - buttonW;
        const int okX = cancelX - buttonGap - buttonW;

        HFONT hFont = TabGetRegularFont();
        CreateWindowExW(0, L"STATIC", lang.dialogBgOpacityLabel.c_str(), WS_CHILD | WS_VISIBLE, margin, labelY, labelW, labelH, hDlg, nullptr, nullptr, nullptr);
        hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_NUMBER | ES_RIGHT | ES_AUTOHSCROLL,
                                inputX, rowY, inputW, inputH, hDlg, reinterpret_cast<HMENU>(1001), nullptr, nullptr);
        CreateWindowExW(0, L"BUTTON", lang.dialogOK.c_str(), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                        okX, buttonY, buttonW, buttonH, hDlg, reinterpret_cast<HMENU>(IDOK), nullptr, nullptr);
        CreateWindowExW(0, L"BUTTON", lang.dialogCancel.c_str(), WS_CHILD | WS_VISIBLE,
                        cancelX, buttonY, buttonW, buttonH, hDlg, reinterpret_cast<HMENU>(IDCANCEL), nullptr, nullptr);
        for (HWND h = GetWindow(hDlg, GW_CHILD); h; h = GetWindow(h, GW_HWNDNEXT))
            SendMessageW(h, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
        int pct = g_state.background.opacity * 100 / 255;
        wchar_t buf[32];
        wsprintfW(buf, L"%d", pct);
        SetWindowTextW(hEdit, buf);
        SendMessageW(hEdit, EM_SETSEL, 0, -1);
        SetFocus(hEdit);
        return FALSE;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK)
        {
            wchar_t buf[32];
            GetWindowTextW(hEdit, buf, 32);
            int val = _wtoi(buf);
            val = (val < 0) ? 0 : (val > 100) ? 100
                                              : val;
            g_state.background.opacity = static_cast<BYTE>(val * 255 / 100);
            SaveFontSettings();
            EndDialog(hDlg, IDOK);
            return TRUE;
        }
        if (LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, IDCANCEL);
            return TRUE;
        }
        break;
    case WM_CLOSE:
        EndDialog(hDlg, IDCANCEL);
        return TRUE;
    }
    return FALSE;
}

void ViewBackgroundOpacity()
{
    const auto &lang = GetLangStrings();
    const DWORD style = WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE;
    const DWORD exStyle = WS_EX_DLGMODALFRAME;

    const int margin = ScaleBackgroundPx(16);
    const int gap = ScaleBackgroundPx(12);
    const int labelW = ScaleBackgroundPx(110);
    const int labelH = ScaleBackgroundPx(20);
    const int inputW = ScaleBackgroundPx(88);
    const int inputH = ScaleBackgroundPx(24);
    const int topPad = ScaleBackgroundPx(20);
    const int rowGap = ScaleBackgroundPx(30);
    const int buttonH = ScaleBackgroundPx(30);
    const int bottomPad = ScaleBackgroundPx(20);

    const int clientW = std::max(ScaleBackgroundPx(340), margin + labelW + gap + inputW + margin);
    const int clientH = topPad + std::max(labelH, inputH) + rowGap + buttonH + bottomPad;

    RECT windowRect = {0, 0, clientW, clientH};
    AdjustWindowRectEx(&windowRect, style, FALSE, exStyle);
    const int windowW = windowRect.right - windowRect.left;
    const int windowH = windowRect.bottom - windowRect.top;

    int x = 300;
    int y = 300;
    RECT ownerRect{};
    if (g_hwndMain && GetWindowRect(g_hwndMain, &ownerRect))
    {
        x = ownerRect.left + ((ownerRect.right - ownerRect.left) - windowW) / 2;
        y = ownerRect.top + ((ownerRect.bottom - ownerRect.top) - windowH) / 2;
    }

    HWND hDlg = CreateWindowExW(exStyle, L"#32770", lang.dialogBgOpacity.c_str(),
                                style, x, y, windowW, windowH,
                                g_hwndMain, nullptr, GetModuleHandleW(nullptr), nullptr);
    if (!hDlg)
        return;
    SetWindowLongPtrW(hDlg, DWLP_DLGPROC, reinterpret_cast<LONG_PTR>(OpacityDlgProc));
    OpacityDlgProc(hDlg, WM_INITDIALOG, 0, 0);
    EnableWindow(g_hwndMain, FALSE);
    MSG msg;
    int quitCode = -1;
    for (;;)
    {
        BOOL gm = GetMessageW(&msg, nullptr, 0, 0);
        if (gm == 0)
        {
            quitCode = static_cast<int>(msg.wParam);
            break;
        }
        if (gm < 0)
            break;

        if (!IsWindow(hDlg))
            break;
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_RETURN)
        {
            SendMessageW(hDlg, WM_COMMAND, IDOK, 0);
            break;
        }
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE)
        {
            SendMessageW(hDlg, WM_COMMAND, IDCANCEL, 0);
            break;
        }
        if (!IsDialogMessageW(hDlg, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
        if (!IsWindow(hDlg))
            break;
    }
    EnableWindow(g_hwndMain, TRUE);
    if (IsWindow(hDlg))
        DestroyWindow(hDlg);
    if (g_bgBitmap)
    {
        DeleteObject(g_bgBitmap);
        g_bgBitmap = nullptr;
    }
    InvalidateRect(g_hwndEditor, nullptr, TRUE);
    SetForegroundWindow(g_hwndMain);
    if (quitCode >= 0)
        PostQuitMessage(quitCode);
}
