/*
  Technical Standard

  Global variable declarations shared across all application modules for window handles.
  Contains handles for main window, editor, status bar, dialogs, and GDI resources.
*/

#pragma once

#include <windows.h>
#include <commdlg.h>
#include <gdiplus.h>
#include "types.h"

extern HWND g_hwndMain;
extern HWND g_hwndEditor;
extern HWND g_hwndStatus;
extern HWND g_hwndTabs;
extern HWND g_hwndCommandBar;
extern HWND g_hwndFindDlg;
extern HWND g_hwndScrollbar;
extern HWND g_hwndSelectionAura;
extern HWND g_hwndCommandPalette;
extern HACCEL g_hAccel;
extern AppState g_state;
extern WNDPROC g_origEditorProc;
extern WNDPROC g_origStatusProc;
extern WNDPROC g_origTabsProc;
extern ULONG_PTR g_gdiplusToken;
extern Gdiplus::Image *g_bgImage;
extern HBITMAP g_bgBitmap;
extern int g_bgBitmapW;
extern int g_bgBitmapH;
extern HBRUSH g_hbrStatusDark;
extern HBRUSH g_hbrMenuDark;
extern HBRUSH g_hbrDialogDark;
extern HBRUSH g_hbrDialogEditDark;
extern PAGESETUPDLGW g_pageSetup;
extern std::wstring g_statusTexts[4];
extern HICON g_hCustomIcon;
extern std::wstring g_editorClassName;
extern HMODULE g_hRichEditModule;
