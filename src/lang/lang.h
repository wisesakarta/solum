#pragma once

#include <windows.h>
#include <string>

enum class LangID
{
    EN,
    JA
};

struct LangStrings
{
    std::wstring appName;
    std::wstring appNameWordmark;
    std::wstring untitled;

    std::wstring menuFile;
    std::wstring menuNew;
    std::wstring menuOpen;
    std::wstring menuSave;
    std::wstring menuSaveAs;
    std::wstring menuPrint;
    std::wstring menuPageSetup;
    std::wstring menuExit;
    std::wstring menuRecentFiles;
    std::wstring menuCloseTab;
    std::wstring menuReopenClosedTab;

    std::wstring menuEdit;
    std::wstring menuUndo;
    std::wstring menuRedo;
    std::wstring menuCut;
    std::wstring menuCopy;
    std::wstring menuPaste;
    std::wstring menuDelete;
    std::wstring menuFind;
    std::wstring menuFindNext;
    std::wstring menuFindPrev;
    std::wstring menuReplace;
    std::wstring menuGoTo;
    std::wstring menuSelectAll;
    std::wstring menuTimeDate;

    std::wstring menuFormat;
    std::wstring menuWordWrap;
    std::wstring menuFont;
    std::wstring menuBold;
    std::wstring menuItalic;
    std::wstring menuStrikethrough;

    std::wstring menuView;
    std::wstring menuZoomIn;
    std::wstring menuZoomOut;
    std::wstring menuZoomDefault;
    std::wstring menuStatusBar;
    std::wstring menuDarkMode;
    std::wstring menuBackground;
    std::wstring menuBgSelect;
    std::wstring menuBgClear;
    std::wstring menuBgOpacity;
    std::wstring menuBgPosition;
    std::wstring menuBgPosTopLeft;
    std::wstring menuBgPosTopCenter;
    std::wstring menuBgPosTopRight;
    std::wstring menuBgPosCenterLeft;
    std::wstring menuBgPosCenter;
    std::wstring menuBgPosCenterRight;
    std::wstring menuBgPosBottomLeft;
    std::wstring menuBgPosBottomCenter;
    std::wstring menuBgPosBottomRight;
    std::wstring menuBgPosTile;
    std::wstring menuBgPosStretch;
    std::wstring menuBgPosFit;
    std::wstring menuBgPosFill;
    std::wstring menuTransparency;
    std::wstring menuAlwaysOnTop;

    std::wstring menuHelp;
    std::wstring menuAbout;
    std::wstring menuCheckUpdates;
    std::wstring menuRunBenchmark;

    std::wstring menuLanguage;
    std::wstring menuLangEnglish;
    std::wstring menuLangJapanese;

    std::wstring dialogFind;
    std::wstring dialogFindReplace;
    std::wstring dialogGoTo;
    std::wstring dialogTransparency;
    std::wstring dialogFindLabel;
    std::wstring dialogReplaceLabel;
    std::wstring dialogFindNext;
    std::wstring dialogReplace;
    std::wstring dialogReplaceAll;
    std::wstring dialogClose;
    std::wstring dialogLineNumber;
    std::wstring dialogOK;
    std::wstring dialogCancel;
    std::wstring dialogOpacityLabel;
    std::wstring dialogBgOpacity;
    std::wstring dialogBgOpacityLabel;
    std::wstring filterTextDocuments;
    std::wstring filterIconFiles;
    std::wstring filterImageFiles;
    std::wstring filterAllFiles;
    std::wstring palettePlaceholder;
    std::wstring paletteNoResults;

    std::wstring msgCannotFind;
    std::wstring msgSaveChanges;
    std::wstring msgCannotOpenFile;
    std::wstring msgCannotSaveFile;
    std::wstring msgError;
    std::wstring msgAbout;
    std::wstring msgCannotLoadRichEdit;
    std::wstring msgDisableTabsCloseOthers;
    std::wstring msgLineNumberOutOfRange;
    std::wstring msgFileTooLargeOpen;
    std::wstring msgFileTooLargeSave;
    std::wstring msgIconLoadFailed;
    std::wstring msgUpdateCheckFailed;
    std::wstring msgUpdateParseVersionFailed;
    std::wstring msgUpdateAvailable;
    std::wstring msgCurrentVersion;
    std::wstring msgLatestVersion;
    std::wstring msgOpenDownloadPage;
    std::wstring msgUpToDate;
    std::wstring msgBenchmarkCompletedPass;
    std::wstring msgBenchmarkCompletedWarn;
    std::wstring msgBenchmarkWarnExceededBudget;
    std::wstring msgBenchmarkCompletedFail;
    std::wstring msgBenchmarkFailCaseNotRun;
    std::wstring msgBenchmarkReportSavedTo;
    std::wstring msgBenchmarkOpenReportNow;
    std::wstring msgBenchmarkReportSaveFailed;

    std::wstring statusLn;
    std::wstring statusCol;
    std::wstring statusLargeFile;

    std::wstring encodingUTF8;
    std::wstring encodingUTF8BOM;
    std::wstring encodingUTF16LE;
    std::wstring encodingUTF16BE;
    std::wstring encodingANSI;

    std::wstring lineEndingCRLF;
    std::wstring lineEndingLF;
    std::wstring lineEndingCR;
};

void InitLanguage();
void SetLanguage(LangID lang);
void SaveLanguageSetting();
LangID LoadLanguageSetting();
LangID GetCurrentLanguage();
const LangStrings &GetLangStrings();
const std::wstring &GetString(const std::wstring &key);
