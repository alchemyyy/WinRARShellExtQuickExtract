/*
 * WinRAR "Extract to Folder" Context Menu
 * 
 * Adds a flat "Extract to <folder>" option with dynamic naming that directly 
 * invokes WinRAR.exe. Reads supported extensions from WinRAR's registry.
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
// Context Menu implementation
//=============================================================================
typedef struct {
    IContextMenu3 IContextMenu3_iface;
    IShellExtInit IShellExtInit_iface;
    LONG cRef;
    wchar_t szFilePath[MAX_PATH];
    wchar_t szFolderName[MAX_PATH];
    wchar_t szDestFolder[MAX_PATH];
    BOOL bIsArchive;
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
        HeapFree(GetProcessHeap(), 0, self);
        InterlockedDecrement(&g_cRef);
    }
    return cRef;
}

//=============================================================================
// IContextMenu
//=============================================================================
#define IDM_EXTRACT 0

static HRESULT STDMETHODCALLTYPE Menu_QueryContextMenu(
    IContextMenu3* This, HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    ExtractContextMenu* self = impl_from_IContextMenu3(This);
    
    if (!self->bIsArchive)
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
    
    if (uFlags & CMF_DEFAULTONLY)
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
    
    // Build menu text with dynamic folder name
    wchar_t menuText[MAX_PATH + 32];
    StringCchPrintfW(menuText, ARRAYSIZE(menuText), L"Extract to \"%s\\\"", self->szFolderName);
    
    // Find WinRAR's position and insert right after it
    UINT insertPos = indexMenu;
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
                insertPos = i + 1;
                break;
            }
        }
    }
    
    // Insert menu item with icon
    MENUITEMINFOW mii = {0};
    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_STRING | MIIM_ID | MIIM_STATE | MIIM_BITMAP;
    mii.wID = idCmdFirst + IDM_EXTRACT;
    mii.dwTypeData = menuText;
    mii.fState = MFS_ENABLED;
    mii.hbmpItem = GetWinRARMenuBitmap();
    
    InsertMenuItemW(hmenu, insertPos, TRUE, &mii);
    
    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, IDM_EXTRACT + 1);
}

static HRESULT STDMETHODCALLTYPE Menu_InvokeCommand(
    IContextMenu3* This, CMINVOKECOMMANDINFO* pici)
{
    ExtractContextMenu* self = impl_from_IContextMenu3(This);
    
    if (HIWORD(pici->lpVerb) != 0)
        return E_INVALIDARG;
    
    if (LOWORD(pici->lpVerb) != IDM_EXTRACT)
        return E_INVALIDARG;
    
    // Create destination folder first
    CreateDirectoryW(self->szDestFolder, NULL);
    
    // Build command line: WinRAR.exe x -ibck "<archive>" "<dest folder>\"
    wchar_t cmdLine[MAX_PATH * 3];
    StringCchPrintfW(cmdLine, ARRAYSIZE(cmdLine), 
        L"\"%s\" x -ibck \"%s\" \"%s\\\"",
        g_WinRARPath, self->szFilePath, self->szDestFolder);
    
    // Execute WinRAR
    STARTUPINFOW si = {0};
    PROCESS_INFORMATION pi = {0};
    si.cb = sizeof(si);
    
    if (CreateProcessW(NULL, cmdLine, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi))
    {
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
        return S_OK;
    }
    
    return E_FAIL;
}

static HRESULT STDMETHODCALLTYPE Menu_GetCommandString(
    IContextMenu3* This, UINT_PTR idCmd, UINT uType, UINT* pReserved, LPSTR pszName, UINT cchMax)
{
    if (idCmd != IDM_EXTRACT)
        return E_INVALIDARG;
    
    if (uType == GCS_HELPTEXTW)
    {
        StringCchCopyW((LPWSTR)pszName, cchMax, L"Extract archive to folder");
        return S_OK;
    }
    else if (uType == GCS_HELPTEXTA)
    {
        StringCchCopyA(pszName, cchMax, "Extract archive to folder");
        return S_OK;
    }
    else if (uType == GCS_VERBW)
    {
        StringCchCopyW((LPWSTR)pszName, cchMax, L"WinRARExtractTo");
        return S_OK;
    }
    else if (uType == GCS_VERBA)
    {
        StringCchCopyA(pszName, cchMax, "WinRARExtractTo");
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
    if (nFiles > 0)
    {
        DragQueryFileW((HDROP)stg.hGlobal, 0, self->szFilePath, MAX_PATH);
        
        self->bIsArchive = IsArchiveFile(self->szFilePath);
        
        if (self->bIsArchive)
        {
            // Extract folder name from filename (without extension)
            wchar_t* fileName = PathFindFileNameW(self->szFilePath);
            StringCchCopyW(self->szFolderName, MAX_PATH, fileName);
            PathRemoveExtensionW(self->szFolderName);
            
            // Build full destination path: parent folder + archive name (no ext)
            StringCchCopyW(self->szDestFolder, MAX_PATH, self->szFilePath);
            PathRemoveFileSpecW(self->szDestFolder);
            PathAppendW(self->szDestFolder, self->szFolderName);
        }
    }
    
    ReleaseStgMedium(&stg);
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
    pMenu->bIsArchive = FALSE;
    
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
        RegSetValueExW(hKey, NULL, 0, REG_SZ, (LPBYTE)L"WinRAR Extract To Folder", 
                       sizeof(L"WinRAR Extract To Folder"));
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
            L"SOFTWARE\\Classes\\SystemFileAssociations\\%s\\shellex\\ContextMenuHandlers\\WinRARExtractTo",
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
            L"SOFTWARE\\Classes\\SystemFileAssociations\\%s\\shellex\\ContextMenuHandlers\\WinRARExtractTo",
            g_ArchiveExtensions[i]);
        RegDeleteKeyW(HKEY_LOCAL_MACHINE, keyPath);
    }
    
    // Clean up old registrations
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\WinRAR~ExtractTo");
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\WinRARExtractTo");
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\~~~WinRARFlat");
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\WinRARFlat");
    
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
    
    return S_OK;
}
