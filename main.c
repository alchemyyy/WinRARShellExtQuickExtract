/*
 * WinRAR Shell Extension
 *
 * Features:
 * - "Extract to <folder>" for single archive files
 * - "Zip to <parent>.zip" for multi-file/folder selections
 * - "Zip each folder separately" for multi-folder selections
 * - "Zip all folders to <parent>.zip" for multi-folder selections
 *
 * Reads supported extensions from WinRAR's registry.
 */

#define WIN32_LEAN_AND_MEAN
#define COBJMACROS
#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <shlwapi.h>
#include <strsafe.h>
#include <commoncontrols.h>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "comctl32.lib")

// Our CLSID
static const CLSID CLSID_WinRARExtract = 
    {0xA1B2C3D4, 0x1234, 0x5678, {0x9A, 0xBC, 0xDE, 0xF0, 0x12, 0x34, 0x56, 0x78}};

static HMODULE g_hModule = NULL;
static LONG g_cRef = 0;

// WinRAR path
static wchar_t g_WinRARPath[MAX_PATH] = L"C:\\Program Files\\WinRAR\\WinRAR.exe";

// Dynamic archive extensions from WinRAR registry
#define MAX_EXTENSIONS 64
static wchar_t g_ArchiveExtensions[MAX_EXTENSIONS][16];
static int g_NumExtensions = 0;

// Max items we can handle in a multi-selection
#define MAX_SELECTED_ITEMS 256

//=============================================================================
// Read extensions from WinRAR's registry
//=============================================================================
static void LoadArchiveExtensions(void)
{
    HKEY hKey, hExtKey;
    DWORD index = 0;
    wchar_t subKeyName[32];
    DWORD subKeyNameLen;
    
    g_NumExtensions = 0;
    
    // Open WinRAR's Setup key
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\WinRAR\\Setup", 0, 
                      KEY_READ, &hKey) != ERROR_SUCCESS)
    {
        return;
    }
    
    // Enumerate subkeys that start with '.'
    while (g_NumExtensions < MAX_EXTENSIONS)
    {
        subKeyNameLen = ARRAYSIZE(subKeyName);
        LSTATUS status = RegEnumKeyExW(hKey, index++, subKeyName, &subKeyNameLen, 
                                        NULL, NULL, NULL, NULL);
        
        if (status != ERROR_SUCCESS)
            break;
        
        // Check if this is an extension (starts with '.')
        if (subKeyName[0] == L'.')
        {
            // Check if this extension is actively associated (Set = 1)
            wchar_t extKeyPath[64];
            StringCchPrintfW(extKeyPath, ARRAYSIZE(extKeyPath), L"Software\\WinRAR\\Setup\\%s", subKeyName);
            
            if (RegOpenKeyExW(HKEY_CURRENT_USER, extKeyPath, 0, KEY_READ, &hExtKey) == ERROR_SUCCESS)
            {
                DWORD setVal = 0;
                DWORD size = sizeof(DWORD);
                RegQueryValueExW(hExtKey, L"Set", NULL, NULL, (LPBYTE)&setVal, &size);
                RegCloseKey(hExtKey);
                
                // Only add if Set = 1 (actively associated with WinRAR)
                if (setVal == 1)
                {
                    wcscpy_s(g_ArchiveExtensions[g_NumExtensions], 16, subKeyName);
                    g_NumExtensions++;
                }
            }
        }
    }
    
    RegCloseKey(hKey);
}

static BOOL IsArchiveFile(const wchar_t* path)
{
    const wchar_t* ext = PathFindExtensionW(path);
    if (!ext || !*ext) return FALSE;
    
    for (int i = 0; i < g_NumExtensions; i++)
    {
        if (_wcsicmp(ext, g_ArchiveExtensions[i]) == 0)
            return TRUE;
    }
    return FALSE;
}

//=============================================================================
// Icon to Bitmap conversion for menu
//=============================================================================
static HBITMAP IconToBitmap(HICON hIcon, int cx, int cy)
{
    if (!hIcon) return NULL;
    
    HDC hdcScreen = GetDC(NULL);
    HDC hdcMem = CreateCompatibleDC(hdcScreen);
    
    BITMAPINFO bmi = {0};
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = cx;
    bmi.bmiHeader.biHeight = -cy;
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    bmi.bmiHeader.biCompression = BI_RGB;
    
    void* pvBits = NULL;
    HBITMAP hBitmap = CreateDIBSection(hdcMem, &bmi, DIB_RGB_COLORS, &pvBits, NULL, 0);
    
    if (hBitmap)
    {
        HBITMAP hOldBmp = SelectObject(hdcMem, hBitmap);
        
        RECT rc = {0, 0, cx, cy};
        FillRect(hdcMem, &rc, (HBRUSH)GetStockObject(BLACK_BRUSH));
        
        DrawIconEx(hdcMem, 0, 0, hIcon, cx, cy, 0, NULL, DI_NORMAL);
        
        SelectObject(hdcMem, hOldBmp);
        
        // Pre-multiply alpha
        if (pvBits)
        {
            BYTE* p = (BYTE*)pvBits;
            for (int i = 0; i < cx * cy; i++)
            {
                BYTE a = p[3];
                if (a < 255)
                {
                    p[0] = (BYTE)((p[0] * a) / 255);
                    p[1] = (BYTE)((p[1] * a) / 255);
                    p[2] = (BYTE)((p[2] * a) / 255);
                }
                p += 4;
            }
        }
    }
    
    DeleteDC(hdcMem);
    ReleaseDC(NULL, hdcScreen);
    
    return hBitmap;
}

static HBITMAP GetWinRARMenuBitmap(void)
{
    static HBITMAP s_hBitmap = NULL;
    
    if (s_hBitmap) return s_hBitmap;
    
    int cx = GetSystemMetrics(SM_CXSMICON);
    int cy = GetSystemMetrics(SM_CYSMICON);
    
    HICON hIcon = NULL;
    ExtractIconExW(g_WinRARPath, 0, NULL, &hIcon, 1);
    
    if (hIcon)
    {
        s_hBitmap = IconToBitmap(hIcon, cx, cy);
        DestroyIcon(hIcon);
    }
    
    return s_hBitmap;
}

//=============================================================================
// Selection types
//=============================================================================
typedef enum {
    SEL_NONE = 0,
    SEL_SINGLE_ARCHIVE,      // Single archive file - show extract option
    SEL_FILES_ONLY,          // Multiple files (no folders) - show zip to single archive
    SEL_FOLDERS_ONLY,        // Multiple folders only - show zip each + zip all options
    SEL_MIXED                // Files and folders mixed - show zip to single archive
} SelectionType;

//=============================================================================
// Context Menu implementation
//=============================================================================
typedef struct {
    IContextMenu3 IContextMenu3_iface;
    IShellExtInit IShellExtInit_iface;
    LONG cRef;

    // For single archive extraction
    wchar_t szFilePath[MAX_PATH];
    wchar_t szFolderName[MAX_PATH];
    wchar_t szDestFolder[MAX_PATH];

    // For multi-selection operations
    wchar_t (*szSelectedPaths)[MAX_PATH];  // Dynamically allocated array
    UINT nSelectedCount;
    UINT nFileCount;
    UINT nFolderCount;
    wchar_t szParentFolder[MAX_PATH];      // Parent folder for naming archives
    wchar_t szParentName[MAX_PATH];        // Just the parent folder name

    SelectionType selType;
} ExtractContextMenu;

static inline ExtractContextMenu* impl_from_IContextMenu3(IContextMenu3* iface) {
    return CONTAINING_RECORD(iface, ExtractContextMenu, IContextMenu3_iface);
}

static inline ExtractContextMenu* impl_from_IShellExtInit(IShellExtInit* iface) {
    return CONTAINING_RECORD(iface, ExtractContextMenu, IShellExtInit_iface);
}

//=============================================================================
// IUnknown
//=============================================================================
static HRESULT STDMETHODCALLTYPE Menu_QueryInterface(IContextMenu3* This, REFIID riid, void** ppv)
{
    ExtractContextMenu* self = impl_from_IContextMenu3(This);
    
    if (IsEqualIID(riid, &IID_IUnknown) || 
        IsEqualIID(riid, &IID_IContextMenu) ||
        IsEqualIID(riid, &IID_IContextMenu2) ||
        IsEqualIID(riid, &IID_IContextMenu3))
    {
        *ppv = &self->IContextMenu3_iface;
        InterlockedIncrement(&self->cRef);
        return S_OK;
    }
    if (IsEqualIID(riid, &IID_IShellExtInit))
    {
        *ppv = &self->IShellExtInit_iface;
        InterlockedIncrement(&self->cRef);
        return S_OK;
    }
    
    *ppv = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE Menu_AddRef(IContextMenu3* This)
{
    ExtractContextMenu* self = impl_from_IContextMenu3(This);
    return InterlockedIncrement(&self->cRef);
}

static ULONG STDMETHODCALLTYPE Menu_Release(IContextMenu3* This)
{
    ExtractContextMenu* self = impl_from_IContextMenu3(This);
    ULONG cRef = InterlockedDecrement(&self->cRef);

    if (cRef == 0)
    {
        if (self->szSelectedPaths)
        {
            HeapFree(GetProcessHeap(), 0, self->szSelectedPaths);
        }
        HeapFree(GetProcessHeap(), 0, self);
        InterlockedDecrement(&g_cRef);
    }
    return cRef;
}

//=============================================================================
// IContextMenu
//=============================================================================
#define IDM_EXTRACT             0
#define IDM_ZIP_TO_SINGLE       1
#define IDM_ZIP_EACH_FOLDER     2
#define IDM_ZIP_ALL_FOLDERS     3

// Find WinRAR's menu position to insert after it
static UINT FindWinRARMenuPosition(HMENU hmenu, UINT defaultPos)
{
    int itemCount = GetMenuItemCount(hmenu);

    for (int i = 0; i < itemCount; i++)
    {
        wchar_t itemText[256] = {0};
        MENUITEMINFOW miiCheck = {0};
        miiCheck.cbSize = sizeof(miiCheck);
        miiCheck.fMask = MIIM_STRING | MIIM_SUBMENU;
        miiCheck.dwTypeData = itemText;
        miiCheck.cch = ARRAYSIZE(itemText);

        if (GetMenuItemInfoW(hmenu, i, TRUE, &miiCheck))
        {
            if (wcsstr(itemText, L"WinRAR") != NULL)
            {
                return i + 1;
            }
        }
    }
    return defaultPos;
}

static HRESULT STDMETHODCALLTYPE Menu_QueryContextMenu(
    IContextMenu3* This, HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    ExtractContextMenu* self = impl_from_IContextMenu3(This);
    UINT cmdCount = 0;

    if (self->selType == SEL_NONE)
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);

    if (uFlags & CMF_DEFAULTONLY)
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);

    UINT insertPos = FindWinRARMenuPosition(hmenu, indexMenu);
    wchar_t menuText[MAX_PATH + 64];
    MENUITEMINFOW mii = {0};
    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_STRING | MIIM_ID | MIIM_STATE | MIIM_BITMAP;
    mii.fState = MFS_ENABLED;
    mii.hbmpItem = GetWinRARMenuBitmap();

    switch (self->selType)
    {
    case SEL_SINGLE_ARCHIVE:
        // Original extract functionality
        StringCchPrintfW(menuText, ARRAYSIZE(menuText), L"Extract to \"%s\\\"", self->szFolderName);
        mii.wID = idCmdFirst + IDM_EXTRACT;
        mii.dwTypeData = menuText;
        InsertMenuItemW(hmenu, insertPos, TRUE, &mii);
        cmdCount = IDM_EXTRACT + 1;
        break;

    case SEL_FILES_ONLY:
    case SEL_MIXED:
        // Zip all selected items to a single archive named after parent folder
        StringCchPrintfW(menuText, ARRAYSIZE(menuText), L"Zip to \"%s.zip\"", self->szParentName);
        mii.wID = idCmdFirst + IDM_ZIP_TO_SINGLE;
        mii.dwTypeData = menuText;
        InsertMenuItemW(hmenu, insertPos, TRUE, &mii);
        cmdCount = IDM_ZIP_TO_SINGLE + 1;
        break;

    case SEL_FOLDERS_ONLY:
        // Option 1: Zip each folder to its own archive
        if (self->nFolderCount > 1)
        {
            StringCchPrintfW(menuText, ARRAYSIZE(menuText), L"Zip each folder separately (%u folders)", self->nFolderCount);
        }
        else
        {
            // Single folder - show the folder name
            wchar_t* folderName = PathFindFileNameW(self->szSelectedPaths[0]);
            StringCchPrintfW(menuText, ARRAYSIZE(menuText), L"Zip \"%s\"", folderName);
        }
        mii.wID = idCmdFirst + IDM_ZIP_EACH_FOLDER;
        mii.dwTypeData = menuText;
        InsertMenuItemW(hmenu, insertPos, TRUE, &mii);

        // Option 2: Zip all folders to single archive (only if multiple folders)
        if (self->nFolderCount > 1)
        {
            StringCchPrintfW(menuText, ARRAYSIZE(menuText), L"Zip all to \"%s.zip\"", self->szParentName);
            mii.wID = idCmdFirst + IDM_ZIP_ALL_FOLDERS;
            mii.dwTypeData = menuText;
            InsertMenuItemW(hmenu, insertPos + 1, TRUE, &mii);
            cmdCount = IDM_ZIP_ALL_FOLDERS + 1;
        }
        else
        {
            cmdCount = IDM_ZIP_EACH_FOLDER + 1;
        }
        break;

    default:
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
    }

    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, cmdCount);
}

// Execute WinRAR with the given command line (non-blocking, shows progress window)
static BOOL ExecuteWinRAR(const wchar_t* cmdLine)
{
    STARTUPINFOW si = {0};
    PROCESS_INFORMATION pi = {0};
    si.cb = sizeof(si);

    // Need a mutable copy of cmdLine for CreateProcessW
    size_t len = wcslen(cmdLine) + 1;
    wchar_t* mutableCmd = HeapAlloc(GetProcessHeap(), 0, len * sizeof(wchar_t));
    if (!mutableCmd) return FALSE;
    wcscpy_s(mutableCmd, len, cmdLine);

    BOOL result = CreateProcessW(NULL, mutableCmd, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi);

    HeapFree(GetProcessHeap(), 0, mutableCmd);

    if (result)
    {
        // Don't wait - let WinRAR run independently with its progress window
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
    }
    return result;
}

// Create a temporary list file for WinRAR with all paths
static BOOL CreateListFile(const wchar_t* listPath, wchar_t (*paths)[MAX_PATH], UINT count)
{
    HANDLE hFile = CreateFileW(listPath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
                               FILE_ATTRIBUTE_TEMPORARY, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return FALSE;

    // Write BOM for Unicode
    BYTE bom[] = {0xFF, 0xFE};
    DWORD written;
    WriteFile(hFile, bom, sizeof(bom), &written, NULL);

    for (UINT i = 0; i < count; i++)
    {
        WriteFile(hFile, paths[i], (DWORD)(wcslen(paths[i]) * sizeof(wchar_t)), &written, NULL);
        WriteFile(hFile, L"\r\n", 4, &written, NULL);
    }

    CloseHandle(hFile);
    return TRUE;
}

static HRESULT STDMETHODCALLTYPE Menu_InvokeCommand(
    IContextMenu3* This, CMINVOKECOMMANDINFO* pici)
{
    ExtractContextMenu* self = impl_from_IContextMenu3(This);

    if (HIWORD(pici->lpVerb) != 0)
        return E_INVALIDARG;

    UINT cmd = LOWORD(pici->lpVerb);
    wchar_t cmdLine[32768];  // Large buffer for multiple file paths
    wchar_t archivePath[MAX_PATH];
    wchar_t listFilePath[MAX_PATH];

    switch (cmd)
    {
    case IDM_EXTRACT:
        // Original extract functionality
        CreateDirectoryW(self->szDestFolder, NULL);
        StringCchPrintfW(cmdLine, ARRAYSIZE(cmdLine),
            L"\"%s\" x \"%s\" \"%s\\\"",
            g_WinRARPath, self->szFilePath, self->szDestFolder);
        break;

    case IDM_ZIP_TO_SINGLE:
        // Zip all selected files/folders to a single archive named after parent folder
        StringCchPrintfW(archivePath, ARRAYSIZE(archivePath),
            L"%s\\%s.zip", self->szParentFolder, self->szParentName);

        // Create temp list file
        GetTempPathW(MAX_PATH, listFilePath);
        StringCchCatW(listFilePath, MAX_PATH, L"winrar_files.lst");

        if (!CreateListFile(listFilePath, self->szSelectedPaths, self->nSelectedCount))
            return E_FAIL;

        // Use -r for recursion (in case folders are selected), -ep1 to keep relative paths
        StringCchPrintfW(cmdLine, ARRAYSIZE(cmdLine),
            L"\"%s\" a -afzip -r -ep1 \"%s\" @\"%s\"",
            g_WinRARPath, archivePath, listFilePath);
        break;

    case IDM_ZIP_EACH_FOLDER:
        // Zip each folder to its own archive concurrently (non-blocking)
        // Each folder's contents go to root of archive (no double folder)
        for (UINT i = 0; i < self->nSelectedCount; i++)
        {
            wchar_t folderPath[MAX_PATH];
            StringCchCopyW(folderPath, MAX_PATH, self->szSelectedPaths[i]);

            // Get folder name for archive name
            wchar_t* folderName = PathFindFileNameW(folderPath);
            wchar_t parentDir[MAX_PATH];
            StringCchCopyW(parentDir, MAX_PATH, folderPath);
            PathRemoveFileSpecW(parentDir);

            StringCchPrintfW(archivePath, ARRAYSIZE(archivePath),
                L"%s\\%s.zip", parentDir, folderName);

            // Use -r for recursion, -ep1 to exclude base folder path
            StringCchPrintfW(cmdLine, ARRAYSIZE(cmdLine),
                L"\"%s\" a -afzip -r -ep1 \"%s\" \"%s\\*\"",
                g_WinRARPath, archivePath, folderPath);

            ExecuteWinRAR(cmdLine);
        }
        return S_OK;

    case IDM_ZIP_ALL_FOLDERS:
        // Zip all folders to a single archive
        StringCchPrintfW(archivePath, ARRAYSIZE(archivePath),
            L"%s\\%s.zip", self->szParentFolder, self->szParentName);

        // Create temp list file
        GetTempPathW(MAX_PATH, listFilePath);
        StringCchCatW(listFilePath, MAX_PATH, L"winrar_folders.lst");

        if (!CreateListFile(listFilePath, self->szSelectedPaths, self->nSelectedCount))
            return E_FAIL;

        // Use -r for recursion, -ep1 to strip parent path but keep folder names
        // Result: FolderA/contents, FolderB/contents in archive
        StringCchPrintfW(cmdLine, ARRAYSIZE(cmdLine),
            L"\"%s\" a -afzip -r -ep1 \"%s\" @\"%s\"",
            g_WinRARPath, archivePath, listFilePath);
        break;

    default:
        return E_INVALIDARG;
    }

    if (!ExecuteWinRAR(cmdLine))
        return E_FAIL;

    return S_OK;
}

static HRESULT STDMETHODCALLTYPE Menu_GetCommandString(
    IContextMenu3* This, UINT_PTR idCmd, UINT uType, UINT* pReserved, LPSTR pszName, UINT cchMax)
{
    const wchar_t* helpTextW = NULL;
    const char* helpTextA = NULL;
    const wchar_t* verbW = NULL;
    const char* verbA = NULL;

    switch (idCmd)
    {
    case IDM_EXTRACT:
        helpTextW = L"Extract archive to folder";
        helpTextA = "Extract archive to folder";
        verbW = L"WinRARExtractTo";
        verbA = "WinRARExtractTo";
        break;
    case IDM_ZIP_TO_SINGLE:
        helpTextW = L"Zip selected items to archive";
        helpTextA = "Zip selected items to archive";
        verbW = L"WinRARZipToSingle";
        verbA = "WinRARZipToSingle";
        break;
    case IDM_ZIP_EACH_FOLDER:
        helpTextW = L"Zip each folder to its own archive";
        helpTextA = "Zip each folder to its own archive";
        verbW = L"WinRARZipEachFolder";
        verbA = "WinRARZipEachFolder";
        break;
    case IDM_ZIP_ALL_FOLDERS:
        helpTextW = L"Zip all folders to single archive";
        helpTextA = "Zip all folders to single archive";
        verbW = L"WinRARZipAllFolders";
        verbA = "WinRARZipAllFolders";
        break;
    default:
        return E_INVALIDARG;
    }

    if (uType == GCS_HELPTEXTW)
    {
        StringCchCopyW((LPWSTR)pszName, cchMax, helpTextW);
        return S_OK;
    }
    else if (uType == GCS_HELPTEXTA)
    {
        StringCchCopyA(pszName, cchMax, helpTextA);
        return S_OK;
    }
    else if (uType == GCS_VERBW)
    {
        StringCchCopyW((LPWSTR)pszName, cchMax, verbW);
        return S_OK;
    }
    else if (uType == GCS_VERBA)
    {
        StringCchCopyA(pszName, cchMax, verbA);
        return S_OK;
    }

    return E_INVALIDARG;
}

//=============================================================================
// IContextMenu2/3
//=============================================================================
static HRESULT STDMETHODCALLTYPE Menu_HandleMenuMsg(
    IContextMenu3* This, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE Menu_HandleMenuMsg2(
    IContextMenu3* This, UINT uMsg, WPARAM wParam, LPARAM lParam, LRESULT* plResult)
{
    if (plResult) *plResult = 0;
    return S_OK;
}

static IContextMenu3Vtbl MenuVtbl = {
    Menu_QueryInterface,
    Menu_AddRef,
    Menu_Release,
    Menu_QueryContextMenu,
    Menu_InvokeCommand,
    Menu_GetCommandString,
    Menu_HandleMenuMsg,
    Menu_HandleMenuMsg2
};

//=============================================================================
// IShellExtInit
//=============================================================================
static HRESULT STDMETHODCALLTYPE Init_QueryInterface(IShellExtInit* This, REFIID riid, void** ppv)
{
    ExtractContextMenu* self = impl_from_IShellExtInit(This);
    return Menu_QueryInterface(&self->IContextMenu3_iface, riid, ppv);
}

static ULONG STDMETHODCALLTYPE Init_AddRef(IShellExtInit* This)
{
    ExtractContextMenu* self = impl_from_IShellExtInit(This);
    return Menu_AddRef(&self->IContextMenu3_iface);
}

static ULONG STDMETHODCALLTYPE Init_Release(IShellExtInit* This)
{
    ExtractContextMenu* self = impl_from_IShellExtInit(This);
    return Menu_Release(&self->IContextMenu3_iface);
}

static HRESULT STDMETHODCALLTYPE Init_Initialize(
    IShellExtInit* This, PCIDLIST_ABSOLUTE pidlFolder, IDataObject* pdtobj, HKEY hkeyProgID)
{
    ExtractContextMenu* self = impl_from_IShellExtInit(This);

    if (!pdtobj) return E_INVALIDARG;

    FORMATETC fmt = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM stg = {0};

    HRESULT hr = IDataObject_GetData(pdtobj, &fmt, &stg);
    if (FAILED(hr)) return hr;

    UINT nFiles = DragQueryFileW((HDROP)stg.hGlobal, 0xFFFFFFFF, NULL, 0);
    if (nFiles == 0)
    {
        ReleaseStgMedium(&stg);
        self->selType = SEL_NONE;
        return S_OK;
    }

    // Get first file path for parent folder detection
    wchar_t firstPath[MAX_PATH];
    DragQueryFileW((HDROP)stg.hGlobal, 0, firstPath, MAX_PATH);

    // Get parent folder
    StringCchCopyW(self->szParentFolder, MAX_PATH, firstPath);
    PathRemoveFileSpecW(self->szParentFolder);

    // Get parent folder name
    wchar_t* parentName = PathFindFileNameW(self->szParentFolder);
    StringCchCopyW(self->szParentName, MAX_PATH, parentName);

    // Single file case - check for archive extraction
    if (nFiles == 1)
    {
        StringCchCopyW(self->szFilePath, MAX_PATH, firstPath);

        if (IsArchiveFile(firstPath))
        {
            self->selType = SEL_SINGLE_ARCHIVE;

            // Extract folder name from filename (without extension)
            wchar_t* fileName = PathFindFileNameW(self->szFilePath);
            StringCchCopyW(self->szFolderName, MAX_PATH, fileName);
            PathRemoveExtensionW(self->szFolderName);

            // Build full destination path: parent folder + archive name (no ext)
            StringCchCopyW(self->szDestFolder, MAX_PATH, self->szFilePath);
            PathRemoveFileSpecW(self->szDestFolder);
            PathAppendW(self->szDestFolder, self->szFolderName);
        }
        else if (PathIsDirectoryW(firstPath))
        {
            // Single folder - treat as folders only
            self->selType = SEL_FOLDERS_ONLY;
            self->nFolderCount = 1;
            self->nFileCount = 0;
            self->nSelectedCount = 1;

            self->szSelectedPaths = HeapAlloc(GetProcessHeap(), 0, sizeof(wchar_t[MAX_PATH]));
            if (self->szSelectedPaths)
            {
                StringCchCopyW(self->szSelectedPaths[0], MAX_PATH, firstPath);
            }
        }
        else
        {
            // Single non-archive file - no menu
            self->selType = SEL_NONE;
        }

        ReleaseStgMedium(&stg);
        return S_OK;
    }

    // Multiple files selected - allocate storage and count types
    UINT maxItems = (nFiles > MAX_SELECTED_ITEMS) ? MAX_SELECTED_ITEMS : nFiles;
    self->szSelectedPaths = HeapAlloc(GetProcessHeap(), 0, maxItems * sizeof(wchar_t[MAX_PATH]));
    if (!self->szSelectedPaths)
    {
        ReleaseStgMedium(&stg);
        return E_OUTOFMEMORY;
    }

    self->nFileCount = 0;
    self->nFolderCount = 0;
    self->nSelectedCount = 0;

    for (UINT i = 0; i < maxItems; i++)
    {
        wchar_t path[MAX_PATH];
        DragQueryFileW((HDROP)stg.hGlobal, i, path, MAX_PATH);

        StringCchCopyW(self->szSelectedPaths[self->nSelectedCount], MAX_PATH, path);
        self->nSelectedCount++;

        if (PathIsDirectoryW(path))
        {
            self->nFolderCount++;
        }
        else
        {
            self->nFileCount++;
        }
    }

    ReleaseStgMedium(&stg);

    // Determine selection type
    if (self->nFileCount > 0 && self->nFolderCount == 0)
    {
        self->selType = SEL_FILES_ONLY;
    }
    else if (self->nFolderCount > 0 && self->nFileCount == 0)
    {
        self->selType = SEL_FOLDERS_ONLY;
    }
    else if (self->nFileCount > 0 && self->nFolderCount > 0)
    {
        self->selType = SEL_MIXED;
    }
    else
    {
        self->selType = SEL_NONE;
    }

    return S_OK;
}

static IShellExtInitVtbl InitVtbl = {
    Init_QueryInterface,
    Init_AddRef,
    Init_Release,
    Init_Initialize
};

//=============================================================================
// Class Factory
//=============================================================================
typedef struct {
    IClassFactory IClassFactory_iface;
} ExtractClassFactory;

static HRESULT STDMETHODCALLTYPE Factory_QueryInterface(IClassFactory* This, REFIID riid, void** ppv)
{
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IClassFactory))
    {
        *ppv = This;
        InterlockedIncrement(&g_cRef);
        return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE Factory_AddRef(IClassFactory* This)
{
    return InterlockedIncrement(&g_cRef);
}

static ULONG STDMETHODCALLTYPE Factory_Release(IClassFactory* This)
{
    return InterlockedDecrement(&g_cRef);
}

static HRESULT STDMETHODCALLTYPE Factory_CreateInstance(
    IClassFactory* This, IUnknown* pUnkOuter, REFIID riid, void** ppv)
{
    ExtractContextMenu* pMenu;

    *ppv = NULL;
    if (pUnkOuter) return CLASS_E_NOAGGREGATION;

    pMenu = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(ExtractContextMenu));
    if (!pMenu) return E_OUTOFMEMORY;

    pMenu->IContextMenu3_iface.lpVtbl = &MenuVtbl;
    pMenu->IShellExtInit_iface.lpVtbl = &InitVtbl;
    pMenu->cRef = 1;
    pMenu->selType = SEL_NONE;
    pMenu->szSelectedPaths = NULL;
    pMenu->nSelectedCount = 0;
    pMenu->nFileCount = 0;
    pMenu->nFolderCount = 0;

    InterlockedIncrement(&g_cRef);

    HRESULT hr = Menu_QueryInterface(&pMenu->IContextMenu3_iface, riid, ppv);
    Menu_Release(&pMenu->IContextMenu3_iface);

    return hr;
}

static HRESULT STDMETHODCALLTYPE Factory_LockServer(IClassFactory* This, BOOL fLock)
{
    if (fLock) InterlockedIncrement(&g_cRef);
    else InterlockedDecrement(&g_cRef);
    return S_OK;
}

static IClassFactoryVtbl FactoryVtbl = {
    Factory_QueryInterface,
    Factory_AddRef,
    Factory_Release,
    Factory_CreateInstance,
    Factory_LockServer
};

static ExtractClassFactory g_ClassFactory = { { &FactoryVtbl } };

//=============================================================================
// DLL Exports
//=============================================================================
BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
    if (fdwReason == DLL_PROCESS_ATTACH)
    {
        g_hModule = hinstDLL;
        DisableThreadLibraryCalls(hinstDLL);
        
        // Load WinRAR path from registry
        HKEY hKey;
        if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WinRAR.exe",
                          0, KEY_READ, &hKey) == ERROR_SUCCESS)
        {
            DWORD size = sizeof(g_WinRARPath);
            RegQueryValueExW(hKey, NULL, NULL, NULL, (LPBYTE)g_WinRARPath, &size);
            RegCloseKey(hKey);
        }
        
        // Load archive extensions from WinRAR's registry
        LoadArchiveExtensions();
    }
    return TRUE;
}

HRESULT STDAPICALLTYPE DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    *ppv = NULL;
    
    if (IsEqualCLSID(rclsid, &CLSID_WinRARExtract))
    {
        return Factory_QueryInterface(&g_ClassFactory.IClassFactory_iface, riid, ppv);
    }
    
    return CLASS_E_CLASSNOTAVAILABLE;
}

HRESULT STDAPICALLTYPE DllCanUnloadNow(void)
{
    return g_cRef == 0 ? S_OK : S_FALSE;
}

HRESULT STDAPICALLTYPE DllRegisterServer(void)
{
    wchar_t dllPath[MAX_PATH];
    HKEY hKey;
    LSTATUS status;

    GetModuleFileNameW(g_hModule, dllPath, MAX_PATH);

    // Reload extensions to make sure we have the latest
    LoadArchiveExtensions();

    // Register CLSID
    status = RegCreateKeyExW(HKEY_LOCAL_MACHINE,
        L"SOFTWARE\\Classes\\CLSID\\{A1B2C3D4-1234-5678-9ABC-DEF012345678}",
        0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL);
    if (status == ERROR_SUCCESS)
    {
        RegSetValueExW(hKey, NULL, 0, REG_SZ, (LPBYTE)L"WinRAR Shell Extension",
                       sizeof(L"WinRAR Shell Extension"));
        RegCloseKey(hKey);
    }

    status = RegCreateKeyExW(HKEY_LOCAL_MACHINE,
        L"SOFTWARE\\Classes\\CLSID\\{A1B2C3D4-1234-5678-9ABC-DEF012345678}\\InProcServer32",
        0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL);
    if (status == ERROR_SUCCESS)
    {
        RegSetValueExW(hKey, NULL, 0, REG_SZ, (LPBYTE)dllPath, (DWORD)(wcslen(dllPath) + 1) * sizeof(wchar_t));
        RegSetValueExW(hKey, L"ThreadingModel", 0, REG_SZ, (LPBYTE)L"Apartment", sizeof(L"Apartment"));
        RegCloseKey(hKey);
    }

    // Register as ContextMenuHandler for each archive type from WinRAR's registry
    for (int i = 0; i < g_NumExtensions; i++)
    {
        wchar_t keyPath[256];
        StringCchPrintfW(keyPath, ARRAYSIZE(keyPath),
            L"SOFTWARE\\Classes\\SystemFileAssociations\\%s\\shellex\\ContextMenuHandlers\\WinRARShellExt",
            g_ArchiveExtensions[i]);

        status = RegCreateKeyExW(HKEY_LOCAL_MACHINE, keyPath, 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL);
        if (status == ERROR_SUCCESS)
        {
            RegSetValueExW(hKey, NULL, 0, REG_SZ,
                           (LPBYTE)L"{A1B2C3D4-1234-5678-9ABC-DEF012345678}",
                           sizeof(L"{A1B2C3D4-1234-5678-9ABC-DEF012345678}"));
            RegCloseKey(hKey);
        }
    }

    // Register for all files (for multi-file zip operations)
    status = RegCreateKeyExW(HKEY_LOCAL_MACHINE,
        L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\WinRARShellExt",
        0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL);
    if (status == ERROR_SUCCESS)
    {
        RegSetValueExW(hKey, NULL, 0, REG_SZ,
                       (LPBYTE)L"{A1B2C3D4-1234-5678-9ABC-DEF012345678}",
                       sizeof(L"{A1B2C3D4-1234-5678-9ABC-DEF012345678}"));
        RegCloseKey(hKey);
    }

    // Register for directories (for folder zip operations)
    status = RegCreateKeyExW(HKEY_LOCAL_MACHINE,
        L"SOFTWARE\\Classes\\Directory\\shellex\\ContextMenuHandlers\\WinRARShellExt",
        0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL);
    if (status == ERROR_SUCCESS)
    {
        RegSetValueExW(hKey, NULL, 0, REG_SZ,
                       (LPBYTE)L"{A1B2C3D4-1234-5678-9ABC-DEF012345678}",
                       sizeof(L"{A1B2C3D4-1234-5678-9ABC-DEF012345678}"));
        RegCloseKey(hKey);
    }

    // Clean up old registrations
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\WinRAR~ExtractTo");
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\WinRARExtractTo");
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\~~~WinRARFlat");
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\WinRARFlat");

    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

    return S_OK;
}

HRESULT STDAPICALLTYPE DllUnregisterServer(void)
{
    // Remove CLSID
    RegDeleteTreeW(HKEY_LOCAL_MACHINE,
        L"SOFTWARE\\Classes\\CLSID\\{A1B2C3D4-1234-5678-9ABC-DEF012345678}");

    // Reload extensions to clean up all registrations
    LoadArchiveExtensions();

    // Remove ContextMenuHandler registrations for each archive type
    for (int i = 0; i < g_NumExtensions; i++)
    {
        wchar_t keyPath[256];
        StringCchPrintfW(keyPath, ARRAYSIZE(keyPath),
            L"SOFTWARE\\Classes\\SystemFileAssociations\\%s\\shellex\\ContextMenuHandlers\\WinRARShellExt",
            g_ArchiveExtensions[i]);
        RegDeleteKeyW(HKEY_LOCAL_MACHINE, keyPath);

        // Also clean up old name
        StringCchPrintfW(keyPath, ARRAYSIZE(keyPath),
            L"SOFTWARE\\Classes\\SystemFileAssociations\\%s\\shellex\\ContextMenuHandlers\\WinRARExtractTo",
            g_ArchiveExtensions[i]);
        RegDeleteKeyW(HKEY_LOCAL_MACHINE, keyPath);
    }

    // Remove all files handler
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\WinRARShellExt");

    // Remove directory handler
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\Directory\\shellex\\ContextMenuHandlers\\WinRARShellExt");

    // Clean up old registrations
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\WinRAR~ExtractTo");
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\WinRARExtractTo");
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\~~~WinRARFlat");
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\WinRARFlat");

    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

    return S_OK;
}
