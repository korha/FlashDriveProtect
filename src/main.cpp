//FlashDriveProtect
#define _WIN32_WINNT _WIN32_IE_WINBLUE
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <commctrl.h>
#include <dbt.h>

#ifndef RRF_SUBKEY_WOW6464KEY
#define RRF_SUBKEY_WOW6464KEY 0x00010000
#endif

#ifndef RRF_SUBKEY_WOW6432KEY
#define RRF_SUBKEY_WOW6432KEY 0x00020000
#endif

static constexpr const wchar_t *const g_wGuidClass = L"App::8bb2307a-1bed-7fd5-69c0-791f62ac1182";
static constexpr const wchar_t *const g_wGuidClassResident = L"App::fdee71dc-6bd6-20b8-1c77-f2664c7b67ab";
static constexpr const wchar_t *const g_wHelp =
        L"Command line parameters:\n"
        L"[/command [/mode:x]]\n"
        L"\n"
        L"command:\n"
        L"/drives:ABCZ - vaccinate drives (A-Z);\n"
        L"/drives+ - vaccinate all drives;\n"
        L"/system+ - computer vaccination;\n"
        L"/system- - remove computer vaccination;\n"
        L"/resident - start program hidden and prompt for vaccinating every new drive;\n"
        L"/q - quit from resident app.\n"
        L"\n"
        L"/mode:0 - hide all messages;\n"
        L"/mode:1 - show warning only messages (default);\n"
        L"/mode:2 - show all messages.";
static constexpr const DWORD g_dwDriveCountMax = 'Z'-'A'+1;
static constexpr const DWORD g_dwSectorSizeMin = 512;
static constexpr const DWORD g_dwSectorSizeMax = 4096;
static constexpr const DWORD g_dwBufSize = 1024*1024;        //1 MB
static constexpr const DWORD g_dwBufLen = 56;

enum EMode
{
    eMsgHideAll,
    eMsgWarningOnly,
    eMsgShowAll
};

typedef LONG (WINAPI *PRegGetValue)
(HKEY hkey, LPCWSTR lpSubKey, LPCWSTR lpValue, DWORD dwFlags, LPDWORD pdwType, PVOID pvData, LPDWORD pcbData);
static PRegGetValue RegGetValuePtr;
static wchar_t *g_wVolume, *g_wDrive, *g_wAutorunFile, *g_wRowDrive;
static BYTE *g_pBtZeroes, *g_pBtBuffer;
static WORD g_wWinVer;
static LONG g_iFixWidth, g_iMinHeight;

//-------------------------------------------------------------------------------------------------
#ifdef NDEBUG
#define ___assert___(cond) do{static_cast<void>(sizeof(cond));}while(false)
#else
#define ___assert___(cond) do{if(!(cond)){int i=__LINE__;char h[]="RUNTIME ASSERTION. Line:           "; \
    if(i>=0){char *c=h+35;do{*--c=i%10+'0';i/=10;}while(i>0);}else{h[25]='?';} \
    if(MessageBoxA(nullptr,__FILE__,h,MB_ICONERROR|MB_OKCANCEL)==IDCANCEL)ExitProcess(0);}}while(false)
#endif

static inline bool FCompareMemory(const BYTE *pBuf1, const void *const pBuf2__, DWORD dwSize)
{
    const char *pBuf2 = static_cast<const char*>(pBuf2__);
    while (dwSize--)
        if (*pBuf1++ != *pBuf2++)
            return false;
    return true;
}

static inline bool FCompareMemoryW(const wchar_t *pBuf1, const wchar_t *pBuf2, DWORD dwSize)
{
    while (dwSize--)
        if (*pBuf1++ != *pBuf2++)
            return false;
    return true;
}

static inline void FCopyMemoryWEnd(wchar_t *pDst, const wchar_t *pSrc)
{
    while (*pDst)
        ++pDst;
    while ((*pDst++ = *pSrc++));
}

static bool FCompareMemoryW(const wchar_t *pBuf1, const wchar_t *pBuf2)
{
    while (*pBuf1 == *pBuf2 && *pBuf2)
        ++pBuf1, ++pBuf2;
    return *pBuf1 == *pBuf2;
}

//-------------------------------------------------------------------------------------------------
static bool FWriteDisk(const HANDLE hFile, const DWORD dwSectorSize)
{
    for (DWORD i = 0, dwBytes; i < dwSectorSize; i += g_dwSectorSizeMin)
        if (!(WriteFile(hFile, g_pBtZeroes, g_dwSectorSizeMin, &dwBytes, nullptr) && dwBytes == g_dwSectorSizeMin))
            return false;
    return true;
}

//-------------------------------------------------------------------------------------------------
static void FVaccinateDrive(const wchar_t wDriveLetter, const HWND hWndOwner, const EMode eMode)
{
    ___assert___(g_wVolume && g_wDrive && g_wAutorunFile && g_wRowDrive && g_pBtZeroes && g_pBtBuffer);
    wchar_t wFileSystem[6/*FAT32`*/];
    *g_wDrive = wDriveLetter;
    if (GetDriveTypeW(g_wDrive) == DRIVE_REMOVABLE && GetVolumeInformationW(g_wDrive, nullptr, 0, nullptr, nullptr, nullptr, wFileSystem, 6/*FAT32`*/) &&
            (FCompareMemoryW(wFileSystem, L"FAT32") || FCompareMemoryW(wFileSystem, L"FAT")))
    {
        const wchar_t *wError = nullptr;
        *g_wAutorunFile = wDriveLetter;
        WIN32_FIND_DATA findFileData;
        HANDLE hFile = FindFirstFileW(g_wAutorunFile, &findFileData);
        if (hFile != INVALID_HANDLE_VALUE)
        {
            FindClose(hFile);
            const DWORD dwAttrib = GetFileAttributesW(g_wAutorunFile);
            if (dwAttrib == INVALID_FILE_ATTRIBUTES)
                return;        //flash already vaccinated
            if (dwAttrib & FILE_ATTRIBUTE_DIRECTORY)
            {
                if (!RemoveDirectoryW(g_wAutorunFile))
                {
                    wchar_t wBuf[19];        //"X:\AUT0123.tmp`" + 4 just in case
                    if (!(GetTempFileNameW(g_wDrive, L"AUT", 0, wBuf) && (DeleteFileW(wBuf), MoveFileW(g_wAutorunFile, wBuf))))
                        wError = L"Can't (re)move directory";
                }
            }
            else
            {
                if (dwAttrib & FILE_ATTRIBUTE_READONLY)
                    SetFileAttributesW(g_wAutorunFile, dwAttrib & ~FILE_ATTRIBUTE_READONLY);
                if (!DeleteFileW(g_wAutorunFile))
                    wError = L"Can't delete file";
            }
        }

        if (!wError)
        {
            DWORD dwSectorSize;
            if (GetDiskFreeSpaceW(g_wDrive, nullptr, &dwSectorSize, nullptr, nullptr) && (dwSectorSize == g_dwSectorSizeMin || dwSectorSize == g_dwSectorSizeMax))
            {
                hFile = CreateFileW(g_wAutorunFile, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
                if (hFile != INVALID_HANDLE_VALUE)
                {
                    CloseHandle(hFile);
                    g_wVolume[4] = wDriveLetter;
                    hFile = CreateFileW(g_wVolume, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr);
                    if (hFile != INVALID_HANDLE_VALUE)
                    {
                        DWORD dwSkipByte = 0,        //hide warning
                                dwBlockNum = 0,
                                dwReadWrite;
                        BYTE *pBtBeginBlock = g_pBtBuffer,
                                *pBtSectorWrite = nullptr,        //hide warning
                                *pByteTemp = nullptr;        //hide warning
                        const char btAttribPrev = GetFileAttributesW(g_wAutorunFile),
                                cFileRecord[] = {'U','T','O','R','U','N',' ','I','N','F',btAttribPrev};
                        wError = L"File not found";
                        while (dwBlockNum < 0x40000000/g_dwBufSize/*~1 GB*/ && ReadFile(hFile, g_pBtBuffer, g_dwBufSize, &dwReadWrite, nullptr) && dwReadWrite == g_dwBufSize)
                        {
                            for (DWORD i = 0; i < g_dwBufSize; i += 32/*file record size*/)
                            {
                                pByteTemp = g_pBtBuffer+i;
                                if (*pByteTemp == 'A' && FCompareMemory(pByteTemp+1, cFileRecord, 11/*Name&Attr*/) && FCompareMemory(pByteTemp+26, g_pBtZeroes, 6/*FstClusLO&FileSize*/) &&
                                        *(pByteTemp+20) == 0 && *(pByteTemp+21) == 0)        //FstClusHI
                                {
                                    if (i < dwSectorSize)
                                        pBtBeginBlock += g_dwSectorSizeMax;
                                    dwSkipByte = (dwBlockNum*g_dwBufSize+i)/dwSectorSize*dwSectorSize;
                                    pBtSectorWrite = g_pBtBuffer + i/dwSectorSize*dwSectorSize;
                                    pByteTemp += 11;
                                    *pByteTemp = FILE_ATTRIBUTE_DEVICE | FILE_ATTRIBUTE_HIDDEN;
                                    if (g_wWinVer >= _WIN32_WINNT_VISTA)        //direct disk writing disabled
                                        wError = (SetFilePointer(hFile, 0, nullptr, FILE_BEGIN) != INVALID_SET_FILE_POINTER &&
                                                ReadFile(hFile, pBtBeginBlock, dwSectorSize, &dwReadWrite, nullptr) && dwReadWrite == dwSectorSize &&
                                                SetFilePointer(hFile, 0, nullptr, FILE_BEGIN) != INVALID_SET_FILE_POINTER &&
                                                FWriteDisk(hFile, dwSectorSize)) ?
                                                    nullptr : L"Error processing the first sector";
                                    else
                                    {
                                        if (SetFilePointer(hFile, dwSkipByte, nullptr, FILE_BEGIN) != INVALID_SET_FILE_POINTER)
                                            wError = (WriteFile(hFile, pBtSectorWrite, dwSectorSize, &dwReadWrite, nullptr) && dwReadWrite == dwSectorSize) ? nullptr : L"Can't write drive";
                                        else
                                            wError = L"Error positioning";
                                    }
                                    dwBlockNum = 0x40000000/g_dwBufSize;        //end
                                    break;
                                }
                            }
                            ++dwBlockNum;
                        }
                        CloseHandle(hFile);

                        if (!wError)
                        {
                            if (g_wWinVer >= _WIN32_WINNT_VISTA)
                            {
                                hFile = CreateFileW(g_wVolume, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr);
                                if (hFile != INVALID_HANDLE_VALUE)
                                {
                                    if (!(WriteFile(hFile, pBtBeginBlock, dwSectorSize, &dwReadWrite, nullptr) && dwReadWrite == dwSectorSize))
                                        wError = L"Can't restore first sector";
                                    ___assert___(dwSkipByte);
                                    if (SetFilePointer(hFile, dwSkipByte, nullptr, FILE_BEGIN) != INVALID_SET_FILE_POINTER)
                                    {
                                        if (!(WriteFile(hFile, pBtSectorWrite, dwSectorSize, &dwReadWrite, nullptr) && dwReadWrite == dwSectorSize))
                                            wError = L"Can't write drive";
                                    }
                                    else
                                        wError = L"Error positioning";
                                    CloseHandle(hFile);

                                    //checking for success
                                    if (!wError && GetFileAttributesW(g_wAutorunFile) != INVALID_FILE_ATTRIBUTES)
                                    {
                                        hFile = CreateFileW(g_wVolume, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr);
                                        if (hFile != INVALID_HANDLE_VALUE)
                                        {
                                            if (!FWriteDisk(hFile, dwSectorSize))
                                                wError = L"Can't write drive";
                                            CloseHandle(hFile);
                                            if (!wError)
                                            {
                                                hFile = CreateFileW(g_wVolume, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr);
                                                if (hFile != INVALID_HANDLE_VALUE)
                                                {
                                                    if (!WriteFile(hFile, pBtBeginBlock, dwSectorSize, &dwReadWrite, nullptr) && dwReadWrite == dwSectorSize)
                                                        wError = L"Can't restore first sector";
                                                    if (SetFilePointer(hFile, dwSkipByte, nullptr, FILE_BEGIN) != INVALID_SET_FILE_POINTER)
                                                    {
                                                        ___assert___(pByteTemp);
                                                        *pByteTemp = btAttribPrev;
                                                        if (!(WriteFile(hFile, pBtSectorWrite, dwSectorSize, &dwReadWrite, nullptr) && dwReadWrite == dwSectorSize))
                                                            wError = L"Can't write drive";
                                                    }
                                                    else
                                                        wError = L"Error positioning";
                                                    CloseHandle(hFile);
                                                    if (!wError)
                                                        wError = L"Error set attributes";
                                                }
                                                else
                                                    wError = L"Can't open drive";
                                            }
                                        }
                                        else
                                            wError = L"Can't open drive";
                                    }
                                }
                                else
                                    wError = L"Can't open drive";
                            }
                            else if (GetFileAttributesW(g_wAutorunFile) != INVALID_FILE_ATTRIBUTES)        //checking for success
                            {
                                hFile = CreateFileW(g_wVolume, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr);
                                if (hFile != INVALID_HANDLE_VALUE)
                                {
                                    ___assert___(dwSkipByte);
                                    if (SetFilePointer(hFile, dwSkipByte, nullptr, FILE_BEGIN) != INVALID_SET_FILE_POINTER)
                                    {
                                        ___assert___(pByteTemp);
                                        *pByteTemp = btAttribPrev;
                                        wError = (WriteFile(hFile, pBtSectorWrite, dwSectorSize, &dwReadWrite, nullptr) && dwReadWrite == dwSectorSize) ? L"Error set attributes" : L"Can't write drive";
                                    }
                                    else
                                        wError = L"Error positioning";
                                    CloseHandle(hFile);
                                }
                                else
                                    wError = L"Can't open drive";
                            }
                        }
                    }
                    else
                        wError = L"Can't open drive";
                }
                else
                    wError = L"Can't create file";
            }
            else
                wError = L"Unknown sector size";
        }

        if (wError)
        {
            if (eMode != eMsgHideAll)
                MessageBoxW(hWndOwner, wError, L"FlashDriveProtect", MB_ICONWARNING);
        }
        else if (eMode == eMsgShowAll)
            MessageBoxW(hWndOwner, L"Vaccinated successfully", L"FlashDriveProtect", MB_ICONINFORMATION);
    }
}

//-------------------------------------------------------------------------------------------------
static bool FCheckVaccinateComputer()
{
    DWORD dwValue,
            dwSize = sizeof(DWORD);
    LONG iRes = RegGetValuePtr(HKEY_LOCAL_MACHINE, L"SYSTEM\\CurrentControlSet\\Services\\cdrom", L"AutoRun", RRF_RT_REG_DWORD, nullptr, &dwValue, &dwSize);
    if ((iRes == ERROR_SUCCESS && dwValue == 0) || iRes == ERROR_FILE_NOT_FOUND)
        if (RegGetValuePtr(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer", L"NoDriveTypeAutoRun", RRF_RT_REG_DWORD, nullptr, &dwValue, &dwSize) == ERROR_SUCCESS && dwValue == 0xFF)
        {
            HKEY hKey;
            if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\IniFileMapping\\Autorun.inf", 0, KEY_QUERY_VALUE | KEY_WOW64_64KEY, &hKey) == ERROR_SUCCESS)
            {
                wchar_t wBuf[18*sizeof(wchar_t)];
                dwSize = 18*sizeof(wchar_t);
                iRes = RegGetValuePtr(hKey, nullptr, nullptr, g_wWinVer >= _WIN32_WINNT_WS03 ? (RRF_RT_REG_SZ | RRF_SUBKEY_WOW6464KEY) : RRF_RT_REG_SZ, nullptr, wBuf, &dwSize) == ERROR_SUCCESS && dwSize == 18*sizeof(wchar_t) && FCompareMemoryW(wBuf, L"@SYS:DoesNotExist");
                RegCloseKey(hKey);
                if (iRes && RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\IniFileMapping\\Autorun.inf", 0, KEY_QUERY_VALUE | KEY_WOW64_32KEY, &hKey) == ERROR_SUCCESS)
                {
                    iRes = RegGetValuePtr(hKey, nullptr, nullptr, g_wWinVer >= _WIN32_WINNT_WS03 ? (RRF_RT_REG_SZ | RRF_SUBKEY_WOW6432KEY) : RRF_RT_REG_SZ, nullptr, wBuf, &dwSize) == ERROR_SUCCESS && dwSize == 18*sizeof(wchar_t) && FCompareMemoryW(wBuf, L"@SYS:DoesNotExist");
                    RegCloseKey(hKey);
                    if (iRes)
                    {
                        dwSize = sizeof(wchar_t);
                        return RegGetValuePtr(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\AutoplayHandlers\\CancelAutoplay\\Files", L"*.*", RRF_RT_REG_SZ, nullptr, wBuf, &dwSize) == ERROR_SUCCESS && dwSize == sizeof(wchar_t) && *wBuf == L'\0';
                    }
                }
            }
        }
    return false;
}

//-------------------------------------------------------------------------------------------------
static inline const BYTE *FDwordToByte(const DWORD *const pDword)
{
    return static_cast<const BYTE*>(static_cast<const void*>(pDword));
}

//-------------------------------------------------------------------------------------------------
static void FVaccinateComputer()
{
    DWORD dwValue = 0;
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SYSTEM\\CurrentControlSet\\Services\\cdrom", 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS)
    {
        RegSetValueExW(hKey, L"AutoRun", 0, REG_DWORD, FDwordToByte(&dwValue), sizeof(DWORD));
        RegCloseKey(hKey);
    }
    if (RegCreateKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer", 0, nullptr, 0, KEY_WRITE, nullptr, &hKey, nullptr) == ERROR_SUCCESS)
    {
        RegSetValueExW(hKey, L"NoDriveTypeAutoRun", 0, REG_DWORD, FDwordToByte(&(dwValue = 0xFF)), sizeof(DWORD));
        RegCloseKey(hKey);
    }
    if (RegCreateKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\IniFileMapping\\Autorun.inf", 0, nullptr, 0, KEY_WRITE | KEY_WOW64_64KEY, nullptr, &hKey, nullptr) == ERROR_SUCCESS)
    {
        RegSetValueExW(hKey, nullptr, 0, REG_SZ, static_cast<const BYTE*>(static_cast<const void*>(L"@SYS:DoesNotExist")), 18*sizeof(wchar_t));
        RegCloseKey(hKey);
    }
    if (RegCreateKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\IniFileMapping\\Autorun.inf", 0, nullptr, 0, KEY_WRITE | KEY_WOW64_32KEY, nullptr, &hKey, nullptr) == ERROR_SUCCESS)
    {
        RegSetValueExW(hKey, nullptr, 0, REG_SZ, static_cast<const BYTE*>(static_cast<const void*>(L"@SYS:DoesNotExist")), 18*sizeof(wchar_t));
        RegCloseKey(hKey);
    }
    if (RegCreateKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\AutoplayHandlers\\CancelAutoplay\\Files", 0, nullptr, 0, KEY_WRITE, nullptr, &hKey, nullptr) == ERROR_SUCCESS)
    {
        RegSetValueExW(hKey, L"*.*", 0, REG_SZ, nullptr, 0);
        RegCloseKey(hKey);
    }
}

//-------------------------------------------------------------------------------------------------
static void FUnVaccinateComputer()
{
    HKEY hKey;
    DWORD dwValue = 1;
    if (RegCreateKeyExW(HKEY_LOCAL_MACHINE, L"SYSTEM\\CurrentControlSet\\Services\\cdrom", 0, nullptr, 0, KEY_WRITE, nullptr, &hKey, nullptr) == ERROR_SUCCESS)
    {
        RegSetValueExW(hKey, L"AutoRun", 0, REG_DWORD, FDwordToByte(&dwValue), sizeof(DWORD));
        RegCloseKey(hKey);
    }
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer", 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS)
    {
        RegSetValueExW(hKey, L"NoDriveTypeAutoRun", 0, REG_DWORD, FDwordToByte(&(dwValue = 0)), sizeof(DWORD));
        RegCloseKey(hKey);
    }
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\IniFileMapping\\Autorun.inf", 0, KEY_SET_VALUE | KEY_WOW64_64KEY, &hKey) == ERROR_SUCCESS)
    {
        RegDeleteValueW(hKey, nullptr);
        RegCloseKey(hKey);
    }
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\IniFileMapping\\Autorun.inf", 0, KEY_SET_VALUE | KEY_WOW64_32KEY, &hKey) == ERROR_SUCCESS)
    {
        RegDeleteValueW(hKey, nullptr);
        RegCloseKey(hKey);
    }
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\AutoplayHandlers\\CancelAutoplay\\Files", 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS)
    {
        RegDeleteValueW(hKey, L"*.*");
        RegCloseKey(hKey);
    }
}

//-------------------------------------------------------------------------------------------------
static void FUpdateListDrives(const HWND hWndListDrives)
{
    int iSaveSel[g_dwDriveCountMax];
    const LRESULT iCountSel = SendMessageW(hWndListDrives, LB_GETSELITEMS, g_dwDriveCountMax, reinterpret_cast<LPARAM>(iSaveSel));
    ___assert___(iCountSel >= 0);
    LRESULT i = 0;
    for (; i < iCountSel; ++i)
        iSaveSel[i] = SendMessageW(hWndListDrives, LB_GETITEMDATA, iSaveSel[i], 0) & 0xFFFF;        //replace the item number by the letter

    SendMessageW(hWndListDrives, LB_RESETCONTENT, 0, 0);
    DWORD dwDrives = GetLogicalDrives();
    wchar_t wFileSystem[6/*FAT32`*/];
    *g_wDrive = L'A';
    WIN32_FIND_DATA findFileData;
    do
    {
        if ((dwDrives & 1) && GetDriveTypeW(g_wDrive) == DRIVE_REMOVABLE &&
                GetVolumeInformationW(g_wDrive, g_wRowDrive+5, 12/*max*/, nullptr, nullptr, nullptr, wFileSystem, 6/*FAT32`*/) &&
                (FCompareMemoryW(wFileSystem, L"FAT32") || FCompareMemoryW(wFileSystem, L"FAT")))
        {
            g_wRowDrive[1] = *g_wAutorunFile = *g_wDrive;

            int iIsVaccinated = 0;
            HANDLE hFile = FindFirstFileW(g_wAutorunFile, &findFileData);
            if (hFile != INVALID_HANDLE_VALUE)
            {
                FindClose(hFile);
                if (GetFileAttributesW(g_wAutorunFile) == INVALID_FILE_ATTRIBUTES)
                {
                    FCopyMemoryWEnd(g_wRowDrive, L" [vaccinated]");
                    iIsVaccinated = 0x10000;
                }
            }
            const LRESULT iRes = SendMessageW(hWndListDrives, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(g_wRowDrive));
            ___assert___(iRes >= 0);
            if (iRes >= 0)
            {
                //is selected before
                SendMessageW(hWndListDrives, LB_SETITEMDATA, iRes, iIsVaccinated | *g_wDrive);
                for (i = 0; i < iCountSel; ++i)
                    if (*g_wDrive == iSaveSel[i])
                    {
                        SendMessageW(hWndListDrives, LB_SETSEL, TRUE, iRes);
                        break;
                    }
            }
        }
        dwDrives >>= 1;
    } while (++*g_wDrive <= L'Z');
}

//-------------------------------------------------------------------------------------------------
static LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    enum
    {
        eVaccinateComp = 1,
        eVaccinateDrives,
        eUpdate,
        eHelp
    };

    static constexpr const int iListBoxWidth = 239;
    static constexpr const int iListBoxTop = 43;
    static HFONT hFont;
    static HICON hIconSuccess,
            hIconWarning,
            hIconVaccinate,
            hIconRemove,
            hIconVaccinateBig,
            hIconUpdate,
            hIconHelp;
    static HWND hWndStatusIcon,
            hWndStatusCompVaccination,
            hWndListDrives,
            hWndBtnVaccinateComp;
    static HBRUSH hbrSysBackground;
    static bool bIsVaccinatedComputer,
            bTimerActive = false;

    switch (uMsg)
    {
    case WM_CREATE:
    {
        const HINSTANCE hInst = reinterpret_cast<const CREATESTRUCT*>(lParam)->hInstance;
        if ((hbrSysBackground = GetSysColorBrush(COLOR_BTNFACE)) &&
                (hIconSuccess = static_cast<HICON>(LoadImageW(hInst, L"IDI_ICON2", IMAGE_ICON, 16, 16, 0)))  &&
                (hIconWarning = static_cast<HICON>(LoadImageW(hInst, L"IDI_ICON3", IMAGE_ICON, 16, 16, 0))) &&
                (hIconVaccinate = static_cast<HICON>(LoadImageW(hInst, L"IDI_ICON4", IMAGE_ICON, 16, 16, 0))) &&
                (hIconRemove = static_cast<HICON>(LoadImageW(hInst, L"IDI_ICON5", IMAGE_ICON, 16, 16, 0))) &&
                (hIconVaccinateBig = static_cast<HICON>(LoadImageW(hInst, L"IDI_ICON4", IMAGE_ICON, 32, 32, 0))) &&
                (hIconUpdate = static_cast<HICON>(LoadImageW(hInst, L"IDI_ICON6", IMAGE_ICON, 16, 16, 0))) &&
                (hIconHelp = static_cast<HICON>(LoadImageW(hInst, L"IDI_ICON7", IMAGE_ICON, 16, 16, 0))) &&
                (hWndStatusIcon = CreateWindowExW(0, WC_STATIC, nullptr, WS_CHILD | WS_VISIBLE | SS_ICON | SS_REALSIZEIMAGE, 9, 8, 16, 16, hWnd, nullptr, hInst, nullptr)) &&
                (hWndStatusCompVaccination = CreateWindowExW(0, WC_STATIC, nullptr, WS_CHILD | WS_VISIBLE, 30, 8, 156, 18, hWnd, nullptr, hInst, nullptr)) &&
                (hWndListDrives = CreateWindowExW(0, WC_LISTBOX, nullptr, WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | LBS_EXTENDEDSEL | LBS_NOINTEGRALHEIGHT, 5, iListBoxTop, iListBoxWidth, 70, hWnd, nullptr, hInst, nullptr)) &&
                (hWndBtnVaccinateComp = CreateWindowExW(0, WC_BUTTON, nullptr, WS_CHILD | WS_VISIBLE, 189, 5, 150, 22, hWnd, reinterpret_cast<HMENU>(eVaccinateComp), hInst, nullptr)))
            if (const HWND hWndBtnVaccinateDrive = CreateWindowExW(0, WC_BUTTON, L"Vaccinate", WS_CHILD | WS_VISIBLE, 5+iListBoxWidth+3, 43, 150, 44, hWnd, reinterpret_cast<HMENU>(eVaccinateDrives), hInst, nullptr))
                if (const HWND hWndUpdate = CreateWindowExW(0, WC_BUTTON, nullptr, WS_CHILD | WS_VISIBLE | BS_ICON, 342, 5, 26, 22, hWnd, reinterpret_cast<HMENU>(eUpdate), hInst, nullptr))
                    if (const HWND hWndHelp = CreateWindowExW(0, WC_BUTTON, nullptr, WS_CHILD | WS_VISIBLE | BS_ICON, 371, 5, 26, 22, hWnd, reinterpret_cast<HMENU>(eHelp), hInst, nullptr))
                        if (SendMessageW(hWnd, WM_COMMAND, eUpdate, reinterpret_cast<LPARAM>(hWndUpdate)) == 0)
                        {
                            NONCLIENTMETRICS nonClientMetrics;
                            nonClientMetrics.cbSize = sizeof(NONCLIENTMETRICS);
                            if (SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &nonClientMetrics, 0) &&
                                    (nonClientMetrics.lfMessageFont.lfHeight = -13, (hFont = CreateFontIndirect(&nonClientMetrics.lfMessageFont))))
                            {
                                SendMessageW(hWndStatusIcon, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);
                                SendMessageW(hWndStatusCompVaccination, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);
                                SendMessageW(hWndListDrives, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);
                                SendMessageW(hWndBtnVaccinateComp, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);
                                SendMessageW(hWndBtnVaccinateDrive, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);
                                SendMessageW(hWndUpdate, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);
                                SendMessageW(hWndHelp, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);

                                SendMessageW(hWndBtnVaccinateDrive, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconVaccinateBig));
                                SendMessageW(hWndUpdate, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconUpdate));
                                SendMessageW(hWndHelp, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconHelp));

                                //select first unvaccinate drive
                                const LRESULT iCount = SendMessageW(hWndListDrives, LB_GETCOUNT, 0, 0);
                                ___assert___(iCount >= 0);
                                for (LRESULT i = 0; i < iCount; ++i)
                                    if ((SendMessageW(hWndListDrives, LB_GETITEMDATA, i, 0) | 0xFFFF) == 0xFFFF)
                                    {
                                        SendMessageW(hWndListDrives, LB_SETSEL, TRUE, i);
                                        break;
                                    }
                                SetFocus(hWndBtnVaccinateDrive);
                                return 0;
                            }
                        }
        return -1;
    }
    case WM_COMMAND:
    {
        switch (wParam)
        {
        case eVaccinateComp:
        {
            if (bIsVaccinatedComputer)        //remove vaccine
            {
                if (MessageBoxW(hWnd, L"This operation is not recommended.\nAre you sure?", L"FlashDriveProtect", MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2) == IDYES)
                {
                    FUnVaccinateComputer();
                    bIsVaccinatedComputer = FCheckVaccinateComputer();
                    if (bIsVaccinatedComputer)
                        MessageBoxW(hWnd, L"Failed to remove vaccine", L"FlashDriveProtect", MB_ICONWARNING);
                    else
                    {
                        SendMessageW(hWndStatusIcon, STM_SETICON, reinterpret_cast<WPARAM>(hIconWarning), 0);
                        SetWindowTextW(hWndStatusCompVaccination, L"Computer not vaccinated");
                        SetWindowTextW(hWndBtnVaccinateComp, L"Vaccinate Computer");
                        SendMessageW(hWndBtnVaccinateComp, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconVaccinate));
                    }
                }
            }
            else        //vaccinate
            {
                FVaccinateComputer();
                bIsVaccinatedComputer = FCheckVaccinateComputer();
                if (bIsVaccinatedComputer)
                {
                    SendMessageW(hWndStatusIcon, STM_SETICON, reinterpret_cast<WPARAM>(hIconSuccess), 0);
                    SetWindowTextW(hWndStatusCompVaccination, L"Computer vaccinated");
                    SetWindowTextW(hWndBtnVaccinateComp, L"Remove Vaccine");
                    SendMessageW(hWndBtnVaccinateComp, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconRemove));
                }
                else
                    MessageBoxW(hWnd, L"Failed to vaccinate computer", L"FlashDriveProtect", MB_ICONWARNING);
            }
            break;
        }
        case eVaccinateDrives:
        {
            int iSaveSel[g_dwDriveCountMax];
            const LRESULT iCount = SendMessageW(hWndListDrives, LB_GETSELITEMS, g_dwDriveCountMax, reinterpret_cast<LPARAM>(iSaveSel));
            ___assert___(iCount >= 0);
            for (LRESULT i = 0, iRes; i < iCount; ++i)
            {
                iRes = SendMessageW(hWndListDrives, LB_GETITEMDATA, iSaveSel[i], 0);
                ___assert___(iRes >= 0);
                if ((iRes >> 16) == 0)
                    FVaccinateDrive(iRes, hWnd, eMsgWarningOnly);
            }
            FUpdateListDrives(hWndListDrives);
            break;
        }
        case eUpdate:
        {
            bIsVaccinatedComputer = FCheckVaccinateComputer();
            if (bIsVaccinatedComputer)
            {
                SendMessageW(hWndStatusIcon, STM_SETICON, reinterpret_cast<WPARAM>(hIconSuccess), 0);
                SetWindowTextW(hWndStatusCompVaccination, L"Computer vaccinated");
                SetWindowTextW(hWndBtnVaccinateComp, L"Remove Vaccine");
                SendMessageW(hWndBtnVaccinateComp, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconRemove));
            }
            else
            {
                SendMessageW(hWndStatusIcon, STM_SETICON, reinterpret_cast<WPARAM>(hIconWarning), 0);
                SetWindowTextW(hWndStatusCompVaccination, L"Computer not vaccinated");
                SetWindowTextW(hWndBtnVaccinateComp, L"Vaccinate Computer");
                SendMessageW(hWndBtnVaccinateComp, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconVaccinate));
            }
            FUpdateListDrives(hWndListDrives);
            break;
        }
        case eHelp:
        {
            MessageBoxW(hWnd, g_wHelp, L"FlashDriveProtect", MB_ICONINFORMATION);
            break;
        }
        }
        return 0;
    }
    case WM_GETMINMAXINFO:
    {
        MINMAXINFO *const mmInfo = reinterpret_cast<MINMAXINFO*>(lParam);
        mmInfo->ptMinTrackSize.x = g_iFixWidth;
        mmInfo->ptMaxTrackSize.x = g_iFixWidth;
        mmInfo->ptMinTrackSize.y = g_iMinHeight;
        return 0;
    }
    case WM_SIZE:
    {
        SetWindowPos(hWndListDrives, HWND_TOP, 0, 0, iListBoxWidth, HIWORD(lParam)-iListBoxTop-5, SWP_NOMOVE | SWP_NOZORDER);
        return 0;
    }
    case WM_DEVICECHANGE:
    {
        if (wParam == DBT_DEVNODES_CHANGED && SetTimer(hWnd, 1, 3000, nullptr))
            bTimerActive = true;
        return TRUE;
    }
    case WM_TIMER:
    {
        if (KillTimer(hWnd, 1))
        {
            bTimerActive = false;
            FUpdateListDrives(hWndListDrives);
        }
        return 0;
    }
    case WM_ENDSESSION:
    {
        SendMessageW(hWnd, WM_CLOSE, 0, 0);
        return 0;
    }
    case WM_DESTROY:
    {
        if (bTimerActive)
            KillTimer(hWnd, 1);
        if (hFont)
            DeleteObject(hFont);
        if (hIconSuccess)
            DestroyIcon(hIconSuccess);
        if (hIconWarning)
            DestroyIcon(hIconWarning);
        if (hIconVaccinate)
            DestroyIcon(hIconVaccinate);
        if (hIconRemove)
            DestroyIcon(hIconRemove);
        if (hIconVaccinateBig)
            DestroyIcon(hIconVaccinateBig);
        if (hIconUpdate)
            DestroyIcon(hIconUpdate);
        if (hIconHelp)
            DestroyIcon(hIconHelp);

        const wchar_t *const wAppPath = g_wVolume+g_dwBufLen;
        if (*wAppPath)
        {
            RECT rect;
            if (GetWindowRect(hWnd, &rect))
            {
                const HANDLE hFile = CreateFileW(wAppPath, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
                if (hFile != INVALID_HANDLE_VALUE)
                {
                    DWORD dwBytesWrite;
                    WriteFile(hFile, &rect, sizeof(RECT), &dwBytesWrite, nullptr);
                    CloseHandle(hFile);
                }
            }
        }

        PostQuitMessage(0);
        return 0;
    }
    case WM_CTLCOLORSTATIC:
        if (reinterpret_cast<HWND>(lParam) == hWndStatusCompVaccination)
        {
            SetBkMode(reinterpret_cast<HDC>(wParam), TRANSPARENT);
            SetTextColor(reinterpret_cast<HDC>(wParam), bIsVaccinatedComputer ? RGB(0, 85, 0) : RGB(255, 0, 0));
            return reinterpret_cast<LRESULT>(hbrSysBackground);
        }
    }
    return DefWindowProcW(hWnd, uMsg, wParam, lParam);
}

//-------------------------------------------------------------------------------------------------
static void FVaccinateAllDrives(const EMode eMode)
{
    DWORD dwDrives = GetLogicalDrives();
    wchar_t wDrive = L'A';
    do
    {
        if (dwDrives & 1)
            FVaccinateDrive(wDrive, nullptr, eMode);
        dwDrives >>= 1;
    } while (++wDrive <= L'Z');
}

//-------------------------------------------------------------------------------------------------
static LRESULT CALLBACK WindowProcResident(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    static EMode eMode;
    static bool bTimerActive = false;

    switch (uMsg)
    {
    case WM_CREATE:
    {
        eMode = static_cast<EMode>(reinterpret_cast<size_t>(reinterpret_cast<const CREATESTRUCT*>(lParam)->lpCreateParams));
        FVaccinateAllDrives(eMode);
        return 0;
    }
    case WM_DEVICECHANGE:
    {
        if (wParam == DBT_DEVNODES_CHANGED && SetTimer(hWnd, 1, 3000, nullptr))
            bTimerActive = true;
        return TRUE;
    }
    case WM_TIMER:
    {
        if (KillTimer(hWnd, 1))
        {
            bTimerActive = false;
            FVaccinateAllDrives(eMode);
        }
        return 0;
    }
    case WM_ENDSESSION:
    {
        SendMessageW(hWnd, WM_CLOSE, 0, 0);
        return 0;
    }
    case WM_DESTROY:
    {
        if (bTimerActive)
            KillTimer(hWnd, 1);
        PostQuitMessage(0);
        return 0;
    }
    }
    return DefWindowProcW(hWnd, uMsg, wParam, lParam);
}

//-------------------------------------------------------------------------------------------------
static const wchar_t * FGetArg(EMode *const eMode)
{
    if (wchar_t *wCmdLine = GetCommandLineW())
    {
        while (*wCmdLine == L' ' || *wCmdLine == L'\t')
            ++wCmdLine;
        if (*wCmdLine != L'\0')
        {
            //1st
            if (*wCmdLine++ == L'"')
            {
                while (*wCmdLine != L'\"')
                {
                    if (*wCmdLine == L'\0')
                        return nullptr;
                    ++wCmdLine;
                }
                ++wCmdLine;
                if (*wCmdLine != L' ' && *wCmdLine != L'\t')
                    return nullptr;
            }
            else
                while (*wCmdLine != L' ' && *wCmdLine != L'\t')
                {
                    if (*wCmdLine == L'\0' || *wCmdLine == L'\"')
                        return nullptr;
                    ++wCmdLine;
                }

            //2nd
            do {++wCmdLine;}
            while (*wCmdLine == L' ' || *wCmdLine == L'\t');
            if (*wCmdLine != L'\0')
            {
                const wchar_t *wArg = wCmdLine;
                if (*wCmdLine++ == L'"')
                {
                    while (*wCmdLine != L'\"')
                    {
                        if (*wCmdLine == L'\0')
                            return nullptr;
                        ++wCmdLine;
                    }
                    if (wCmdLine[1] != L' ' && wCmdLine[1] != L'\t' && wCmdLine[1] != L'\0')
                        return nullptr;
                    ++wArg;
                }
                else
                    while (*wCmdLine != L' ' && *wCmdLine != L'\t' && *wCmdLine != L'\0')
                        ++wCmdLine;

                if (*wArg == L'/')
                {
                    ++wArg;
                    *wCmdLine = L'\0';

                    //3rd
                    do {++wCmdLine;}
                    while (*wCmdLine == L' ' || *wCmdLine == L'\t');
                    if (*wCmdLine != L'\0')
                    {
                        const wchar_t *wArg2 = wCmdLine;
                        if (*wCmdLine++ == L'"')
                        {
                            while (*wCmdLine != L'\"')
                            {
                                if (*wCmdLine == L'\0')
                                    return wArg;
                                ++wCmdLine;
                            }
                            if (wCmdLine[1] != L' ' && wCmdLine[1] != L'\t' && wCmdLine[1] != L'\0')
                                return wArg;
                            ++wArg2;
                        }
                        else
                            while (*wCmdLine != L' ' && *wCmdLine != L'\t' && *wCmdLine != L'\0')
                                ++wCmdLine;

                        if (FCompareMemoryW(wArg2, L"/mode:", 6))
                        {
                            *wCmdLine = L'\0';
                            if (wArg2[6] == L'0' && wArg2[7] == L'\0')
                                *eMode = eMsgHideAll;
                            else if (wArg2[6] == L'2' && wArg2[7] == L'\0')
                                *eMode = eMsgShowAll;
                        }
                    }
                    return wArg;
                }
            }
        }
    }
    return nullptr;
}

//-------------------------------------------------------------------------------------------------
static void FGetRegGetValuePtr()
{
    if (const HMODULE hMod = GetModuleHandleW(g_wWinVer >= _WIN32_WINNT_WS03 ? L"advapi32.dll" : L"shlwapi.dll"))
        RegGetValuePtr = reinterpret_cast<PRegGetValue>(GetProcAddress(hMod, g_wWinVer >= _WIN32_WINNT_WS03 ? "RegGetValueW" : "SHRegGetValueW"));
}

//-------------------------------------------------------------------------------------------------
void FMain()
{
    INITCOMMONCONTROLSEX initComCtrlEx;
    initComCtrlEx.dwSize = sizeof(INITCOMMONCONTROLSEX);
    initComCtrlEx.dwICC = ICC_STANDARD_CLASSES;
    if (InitCommonControlsEx(&initComCtrlEx) == TRUE)
    {
        EMode eMode = eMsgWarningOnly;
        const wchar_t *pArg = FGetArg(&eMode);
        const HWND hWndResident = FindWindowW(g_wGuidClassResident, nullptr);
        if (pArg && pArg[0] == L'q' && pArg[1] == L'\0')
        {
            ___assert___(pArg[-1] == L'/');
            if (hWndResident)
                PostMessageW(hWndResident, WM_CLOSE, 0, 0);
            return;
        }
        if (hWndResident)
        {
            if (MessageBoxW(nullptr, L"Program already running in resident mode.\nClose it?",
                            L"FlashDriveProtect", MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2) != IDYES)
                return;
            PostMessageW(hWndResident, WM_CLOSE, 0, 0);
        }

        if (const HANDLE hProcHeap = GetProcessHeap())
        {
            wchar_t wBuf[g_dwBufLen+MAX_PATH];        //g_dwBufLen = "\\.\X:`X:\`X:\AUTORUN.INF`(X:) 12345678901 [vaccinated]`"
            wchar_t *pBuf = wBuf;
            const wchar_t *pSrc = L"\\\\.\\X:\0X:\\\0X:\\AUTORUN.INF\0(X:) ";
            DWORD dwSize = 31;
            while (dwSize--)
                *pBuf++ = *pSrc++;

            g_wVolume = wBuf;        //***
            g_wDrive = wBuf+7;
            g_wAutorunFile = wBuf+11;
            g_wRowDrive = wBuf+26;
            g_wWinVer = GetVersion();
            g_wWinVer = g_wWinVer << 8 | g_wWinVer >> 8;

            BYTE btZeroes[g_dwSectorSizeMin];
            g_pBtZeroes = btZeroes;        //***
            BYTE *pDst = btZeroes;
            dwSize = g_dwSectorSizeMin;
            while (dwSize--)
                *pDst++ = '\0';

            if (pArg)
            {
                ___assert___(pArg[-1] == L'/');
                if (FCompareMemoryW(pArg, L"drives", 6))
                {
                    if (pArg[6] == L':')        //selected drives
                    {
                        if ((g_pBtBuffer = static_cast<BYTE*>(HeapAlloc(hProcHeap, HEAP_NO_SERIALIZE, g_dwBufSize))))
                        {
                            pArg += 7;
                            while (*pArg >= L'A' && *pArg <= L'Z')
                            {
                                FVaccinateDrive(*pArg, nullptr, eMode);
                                ++pArg;
                            }
                            HeapFree(hProcHeap, HEAP_NO_SERIALIZE, g_pBtBuffer);
                        }
                        return;
                    }
                    else if (pArg[6] == L'+' && pArg[7] == L'\0')        //all drives
                    {
                        if ((g_pBtBuffer = static_cast<BYTE*>(HeapAlloc(hProcHeap, HEAP_NO_SERIALIZE, g_dwBufSize))))
                        {
                            FVaccinateAllDrives(eMode);
                            HeapFree(hProcHeap, HEAP_NO_SERIALIZE, g_pBtBuffer);
                        }
                        return;
                    }
                }
                else if (FCompareMemoryW(pArg, L"system", 6))
                {
                    pArg += 6;
                    if (*pArg == L'+' && pArg[1] == L'\0')
                    {
                        FVaccinateComputer();
                        if (eMode != eMsgHideAll)
                        {
                            FGetRegGetValuePtr();
                            if (RegGetValuePtr)
                            {
                                if (FCheckVaccinateComputer())
                                {
                                    if (eMode == eMsgShowAll)
                                        MessageBoxW(nullptr, L"System successfully vaccinated", L"FlashDriveProtect", MB_ICONINFORMATION);
                                }
                                else
                                    MessageBoxW(nullptr, L"Failed to vaccinate computer", L"FlashDriveProtect", MB_ICONWARNING);
                            }
                        }
                        return;
                    }
                    else if (*pArg == L'-' && pArg[1] == L'\0')
                    {
                        FUnVaccinateComputer();
                        if (eMode != eMsgHideAll)
                        {
                            FGetRegGetValuePtr();
                            if (RegGetValuePtr)
                            {
                                if (FCheckVaccinateComputer())
                                    MessageBoxW(nullptr, L"Failed to remove vaccine", L"FlashDriveProtect", MB_ICONWARNING);
                                else if (eMode == eMsgShowAll)
                                    MessageBoxW(nullptr, L"Successfully remove vaccine from system", L"FlashDriveProtect", MB_ICONINFORMATION);
                            }
                        }
                        return;
                    }
                }
                else if (FCompareMemoryW(pArg, L"resident"))
                {
                    if ((g_pBtBuffer = static_cast<BYTE*>(HeapAlloc(hProcHeap, HEAP_NO_SERIALIZE, g_dwBufSize))))
                    {
                        WNDCLASSEXW wndCl;
                        wndCl.cbSize = sizeof(WNDCLASSEX);
                        wndCl.style = 0;
                        wndCl.lpfnWndProc = WindowProcResident;
                        wndCl.cbClsExtra = 0;
                        wndCl.cbWndExtra = 0;
                        wndCl.hInstance = GetModuleHandleW(nullptr);
                        wndCl.hIcon = nullptr;
                        wndCl.hCursor = nullptr;
                        wndCl.hbrBackground = nullptr;
                        wndCl.lpszMenuName = nullptr;
                        wndCl.lpszClassName = g_wGuidClassResident;
                        wndCl.hIconSm = nullptr;

                        if (RegisterClassExW(&wndCl))
                        {
                            if (CreateWindowExW(0, g_wGuidClassResident, nullptr, 0, 0, 0, 0, 0, nullptr, nullptr, wndCl.hInstance, reinterpret_cast<PVOID>(eMode)))
                            {
                                MSG msg;
                                while (GetMessageW(&msg, nullptr, 0, 0) > 0)
                                    DispatchMessageW(&msg);
                            }
                            UnregisterClassW(g_wGuidClassResident, wndCl.hInstance);
                        }
                        HeapFree(hProcHeap, HEAP_NO_SERIALIZE, g_pBtBuffer);
                    }
                    return;
                }
            }

            RECT rect; rect.left = 0; rect.top = 0; rect.right = 402; rect.bottom = 92;
            if (AdjustWindowRectEx(&rect, WS_OVERLAPPEDWINDOW | WS_VISIBLE, FALSE, 0))
            {
                FGetRegGetValuePtr();
                if (RegGetValuePtr && (g_pBtBuffer = static_cast<BYTE*>(HeapAlloc(hProcHeap, HEAP_NO_SERIALIZE, g_dwBufSize))))
                {
                    const HINSTANCE hInst = GetModuleHandleW(nullptr);
                    WNDCLASSEX wndCl;
                    wndCl.cbSize = sizeof(WNDCLASSEX);
                    wndCl.style = 0;
                    wndCl.lpfnWndProc = WindowProc;
                    wndCl.cbClsExtra = 0;
                    wndCl.cbWndExtra = 0;
                    wndCl.hInstance = hInst;
                    wndCl.hIcon = static_cast<HICON>(LoadImageW(hInst, L"IDI_ICON1", IMAGE_ICON, 0, 0, LR_DEFAULTSIZE));
                    wndCl.hCursor = LoadCursorW(nullptr, IDC_ARROW);
                    wndCl.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE+1);
                    wndCl.lpszMenuName = nullptr;
                    wndCl.lpszClassName = g_wGuidClass;
                    wndCl.hIconSm = static_cast<HICON>(LoadImageW(hInst, L"IDI_ICON1", IMAGE_ICON, 16, 16, 0));

                    if (RegisterClassExW(&wndCl))
                    {
                        g_iFixWidth = rect.right-rect.left;
                        g_iMinHeight = rect.bottom-rect.top;

                        rect.left = CW_USEDEFAULT;

                        wchar_t *const wAppPath = wBuf+56;
                        DWORD dwSize = GetModuleFileNameW(nullptr, wAppPath, MAX_PATH+1-4);        //".geo"
                        if (dwSize >= 4 && dwSize < MAX_PATH-4)        //".geo"
                        {
                            wchar_t *wTemp = wAppPath+dwSize;
                            *wTemp++ = L'.';
                            *wTemp++ = L'g';
                            *wTemp++ = L'e';
                            *wTemp++ = L'o';
                            *wTemp = L'\0';
                            const HANDLE hFile = CreateFileW(wAppPath, GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
                            if (hFile != INVALID_HANDLE_VALUE)
                            {
                                LARGE_INTEGER iFileSize;
                                if (GetFileSizeEx(hFile, &iFileSize) && iFileSize.QuadPart == sizeof(RECT) &&
                                        !(ReadFile(hFile, &rect, sizeof(RECT), &dwSize, nullptr) && dwSize == sizeof(RECT)))
                                    rect.left = CW_USEDEFAULT;
                                CloseHandle(hFile);
                            }
                        }
                        else
                            *wAppPath = L'\0';

                        if (rect.left == CW_USEDEFAULT)
                        {
                            rect.top = CW_USEDEFAULT;
                            rect.bottom = 268+(g_iMinHeight-92);
                        }
                        else
                            rect.bottom -= rect.top;

                        if (CreateWindowExW(0, g_wGuidClass, L"FlashDriveProtect", WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                                            rect.left, rect.top, g_iFixWidth, rect.bottom, nullptr, nullptr, hInst, nullptr))
                        {
                            MSG msg;
                            while (GetMessageW(&msg, nullptr, 0, 0) > 0)
                                DispatchMessageW(&msg);
                        }
                        UnregisterClassW(g_wGuidClass, hInst);
                    }
                    HeapFree(hProcHeap, HEAP_NO_SERIALIZE, g_pBtBuffer);
                }
            }
        }
    }
}

//-------------------------------------------------------------------------------------------------
extern "C" int start()
{
    ___assert___(sizeof(wchar_t) == 2);
    ___assert___(g_dwBufSize%g_dwSectorSizeMax == 0 && g_dwBufSize >= g_dwSectorSizeMax*2);

    FMain();
    ExitProcess(0);
    return 0;
}
