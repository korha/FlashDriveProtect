#include <windows.h>
#include <commctrl.h>
#include <dbt.h>
#include <cassert>

#ifndef RRF_SUBKEY_WOW6464KEY
#define RRF_SUBKEY_WOW6464KEY 0x00010000
#endif

#ifndef RRF_SUBKEY_WOW6432KEY
#define RRF_SUBKEY_WOW6432KEY 0x00020000
#endif

static const wchar_t *const g_wGuidClass = L"App::8bb2307a-1bed-7fd5-69c0-791f62ac1182",
*const g_wGuidClassResident = L"App::fdee71dc-6bd6-20b8-1c77-f2664c7b67ab",
wHelp[] = {
    'F','l','a','s','h','D','r','i','v','e','P','r','o','t','e','c','t',' ',
    'v', '.', __DATE__[9], __DATE__[10], '.',
    ((__DATE__[0] == 'O' || __DATE__[0] == 'N' || __DATE__[0] == 'D') ? '1' : '0'),
    (__DATE__[0] == 'F' ? '2' : __DATE__[1] == 'p' ? '4' : __DATE__[2] == 'y' ? '5' :
    __DATE__[2] == 'l' ? '7' : __DATE__[2] == 'g' ? '8' : __DATE__[0] == 'S' ? '9' :
    __DATE__[0] == 'O' ? '0' : __DATE__[0] == 'N' ? '1' : __DATE__[0] == 'D' ? '2' :
    __DATE__[0] == 'J' ? '1' : __DATE__[0] == 'M' ? '3' : __DATE__[1] == 'u' ? '6' : '?'),
    '.', (__DATE__[4] == ' ' ? '0' : __DATE__[4]), __DATE__[5],'\n','\n',
    'C','o','m','m','a','n','d',' ','l','i','n','e',' ','p','a','r','a','m','e','t','e','r','s',':','\n',
    '[','/','c','o','m','m','a','n','d',' ','[','/','m','o','d','e',':','x',']',']','\n',
    '\n',
    'c','o','m','m','a','n','d',':','\n',
    '/','d','r','i','v','e','s',':','A','B','C','Z',' ','-',' ','v','a','c','c','i','n','a','t','e',' ','d','r','i','v','e','s',' ','(','A','-','Z',')',';','\n',
    '/','d','r','i','v','e','s','+',' ','-',' ','v','a','c','c','i','n','a','t','e',' ','a','l','l',' ','d','r','i','v','e','s',';','\n',
    '/','s','y','s','t','e','m','+',' ','-',' ','c','o','m','p','u','t','e','r',' ','v','a','c','c','i','n','a','t','i','o','n',';','\n',
    '/','s','y','s','t','e','m','-',' ','-',' ','r','e','m','o','v','e',' ','c','o','m','p','u','t','e','r',' ','v','a','c','c','i','n','a','t','i','o','n',';','\n',
    '/','r','e','s','i','d','e','n','t',' ','-',' ','s','t','a','r','t',' ','p','r','o','g','r','a','m',' ','h','i','d','d','e','n',' ','a','n','d',' ','p','r','o','m','p','t',' ','f','o','r',' ','v','a','c','c','i','n','a','t','i','n','g',' ','e','v','e','r','y',' ','n','e','w',' ','d','r','i','v','e',';','\n',
    '/','q',' ','-',' ','q','u','i','t',' ','f','r','o','m',' ','r','e','s','i','d','e','n','t',' ','a','p','p','.','\n',
    '\n',
    '/','m','o','d','e',':','0',' ','-',' ','h','i','d','e',' ','a','l','l',' ','m','e','s','s','a','g','e','s',';','\n',
    '/','m','o','d','e',':','1',' ','-',' ','s','h','o','w',' ','w','a','r','n','i','n','g',' ','o','n','l','y',' ','m','e','s','s','a','g','e','s',' ','(','d','e','f','a','u','l','t',')',';','\n',
    '/','m','o','d','e',':','2',' ','-',' ','s','h','o','w',' ','a','l','l',' ','m','e','s','s','a','g','e','s','.','\0'};

enum
{
    eDriveCountMax = 'Z'-'A'+1,
    eSectorSizeMin = 512,
    eSectorSizeMax = 4096,
    eBufSize = 1024*1024,        //1 MB
    eBufLen = 56
};

enum EMode
{
    eMsgHideAll,
    eMsgWarningOnly,
    eMsgShowAll
};

typedef LONG WINAPI (*PRegGetValuePointer)
(HKEY hkey, LPCWSTR lpSubKey, LPCWSTR lpValue, DWORD dwFlags, LPDWORD pdwType, PVOID pvData, LPDWORD pcbData);
static PRegGetValuePointer RegGetValuePointer;
static wchar_t *g_wVolume, *g_wDrive, *g_wAutorunFile, *g_wRowDrive;
static BYTE *g_pBtZeroes, *g_pBtBuffer;
static WORD g_wWinVer;
static LONG g_iFixWidth, g_iMinHeight;

//-------------------------------------------------------------------------------------------------
void fVaccinateDrive(const wchar_t wDriveLetter, const HWND hWndOwner, const EMode eMode)
{
    assert(g_wVolume && g_wDrive && g_wAutorunFile && g_wRowDrive && g_pBtZeroes && g_pBtBuffer);
    wchar_t wFileSystem[6/*FAT32`*/];
    *g_wDrive = wDriveLetter;
    if (GetDriveType(g_wDrive) == DRIVE_REMOVABLE && GetVolumeInformation(g_wDrive, 0, 0, 0, 0, 0, wFileSystem, 6/*FAT32`*/) &&
            (wcscmp(wFileSystem, L"FAT32") == 0 || wcscmp(wFileSystem, L"FAT") == 0))
    {
        const wchar_t *wError = 0;
        *g_wAutorunFile = wDriveLetter;
        WIN32_FIND_DATA findFileData;
        HANDLE hFile = FindFirstFile(g_wAutorunFile, &findFileData);
        if (hFile != INVALID_HANDLE_VALUE)
        {
            FindClose(hFile);
            const DWORD dwAttrib = GetFileAttributes(g_wAutorunFile);
            if (dwAttrib == INVALID_FILE_ATTRIBUTES)
                return;        //flash already vaccinated
            if (dwAttrib & FILE_ATTRIBUTE_DIRECTORY)
            {
                if (!RemoveDirectory(g_wAutorunFile))
                {
                    wchar_t wBuf[19];        //[19 = "X:\AUT0123.tmp`" + 4 if something goes wrong)]
                    if (!(GetTempFileName(g_wDrive, L"AUT", 0, wBuf) && (DeleteFile(wBuf), MoveFile(g_wAutorunFile, wBuf))))
                        wError = L"Can't (re)move directory";
                }
            }
            else
            {
                if (dwAttrib & FILE_ATTRIBUTE_READONLY)
                    SetFileAttributes(g_wAutorunFile, dwAttrib & ~FILE_ATTRIBUTE_READONLY);
                if (!DeleteFile(g_wAutorunFile))
                    wError = L"Can't delete file";
            }
        }

        if (!wError)
        {
            DWORD dwSectorSize;
            if (GetDiskFreeSpace(g_wDrive, 0, &dwSectorSize, 0, 0) && (dwSectorSize == eSectorSizeMin || dwSectorSize == eSectorSizeMax))
            {
                hFile = CreateFile(g_wAutorunFile, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
                if (hFile != INVALID_HANDLE_VALUE)
                {
                    CloseHandle(hFile);
                    g_wVolume[4] = wDriveLetter;
                    hFile = CreateFile(g_wVolume, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0);
                    if (hFile != INVALID_HANDLE_VALUE)
                    {
                        DWORD dwSkipByte = 0/*hide warning*/,
                                dwBlockNum = 0,
                                dwReadWrite;
                        BYTE *pBtBeginBlock = g_pBtBuffer,
                                *pBtSectorWrite = 0/*hide warning*/,
                                *pByteTemp = 0/*hide warning*/;
                        const char btAttribPrev = GetFileAttributes(g_wAutorunFile),
                                cFileRecord[] = {'U','T','O','R','U','N',' ','I','N','F',btAttribPrev};
                        wError = L"File not found";
                        while (dwBlockNum < 0x40000000/eBufSize/*~1 GB*/ && ReadFile(hFile, g_pBtBuffer, eBufSize, &dwReadWrite, 0) && dwReadWrite == eBufSize)
                        {
                            for (DWORD i = 0; i < eBufSize; i += 32/*file record size*/)
                            {
                                pByteTemp = g_pBtBuffer+i;
                                if (*pByteTemp == 'A' && memcmp(pByteTemp+1, cFileRecord, 11/*Name&Attr*/) == 0 && memcmp(pByteTemp+26, g_pBtZeroes, 6/*FstClusLO&FileSize*/) == 0 &&
                                        *(pByteTemp+20) == 0 && *(pByteTemp+21) == 0)        //FstClusHI
                                {
                                    if (i < dwSectorSize)
                                        pBtBeginBlock += eSectorSizeMax;
                                    dwSkipByte = (dwBlockNum*eBufSize+i)/dwSectorSize*dwSectorSize;
                                    pBtSectorWrite = g_pBtBuffer + i/dwSectorSize*dwSectorSize;
                                    pByteTemp += 11;
                                    *pByteTemp = FILE_ATTRIBUTE_DEVICE | FILE_ATTRIBUTE_HIDDEN;
                                    if (g_wWinVer >= _WIN32_WINNT_VISTA)        //direct disk writing disabled
                                        wError = (SetFilePointer(hFile, 0, 0, FILE_BEGIN) != INVALID_SET_FILE_POINTER &&
                                                ReadFile(hFile, pBtBeginBlock, dwSectorSize, &dwReadWrite, 0) && dwReadWrite == dwSectorSize &&
                                                SetFilePointer(hFile, 0, 0, FILE_BEGIN) != INVALID_SET_FILE_POINTER &&
                                                WriteFile(hFile, g_pBtZeroes, dwSectorSize, &dwReadWrite, 0) && dwReadWrite == dwSectorSize) ?
                                                    0 : L"Error processing the first sector";
                                    else
                                    {
                                        if (SetFilePointer(hFile, dwSkipByte, 0, FILE_BEGIN) != INVALID_SET_FILE_POINTER)
                                            wError = (WriteFile(hFile, pBtSectorWrite, dwSectorSize, &dwReadWrite, 0) && dwReadWrite == dwSectorSize) ? 0 : L"Can't write drive";
                                        else
                                            wError = L"Error positioning";
                                    }
                                    dwBlockNum = 0x40000000/eBufSize;        //end
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
                                hFile = CreateFile(g_wVolume, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0);
                                if (hFile != INVALID_HANDLE_VALUE)
                                {
                                    if (!(WriteFile(hFile, pBtBeginBlock, dwSectorSize, &dwReadWrite, 0) && dwReadWrite == dwSectorSize))
                                        wError = L"Can't restore first sector";
                                    assert(dwSkipByte);
                                    if (SetFilePointer(hFile, dwSkipByte, 0, FILE_BEGIN) != INVALID_SET_FILE_POINTER)
                                    {
                                        if (!(WriteFile(hFile, pBtSectorWrite, dwSectorSize, &dwReadWrite, 0) && dwReadWrite == dwSectorSize))
                                            wError = L"Can't write drive";
                                    }
                                    else
                                        wError = L"Error positioning";
                                    CloseHandle(hFile);

                                    //checking for success
                                    if (!wError && GetFileAttributes(g_wAutorunFile) != INVALID_FILE_ATTRIBUTES)
                                    {
                                        hFile = CreateFile(g_wVolume, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0);
                                        if (hFile != INVALID_HANDLE_VALUE)
                                        {
                                            if (!(WriteFile(hFile, g_pBtZeroes, dwSectorSize, &dwReadWrite, 0) && dwReadWrite == dwSectorSize))
                                                wError = L"Can't write drive";
                                            CloseHandle(hFile);
                                            if (!wError)
                                            {
                                                hFile = CreateFile(g_wVolume, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0);
                                                if (hFile != INVALID_HANDLE_VALUE)
                                                {
                                                    if (!WriteFile(hFile, pBtBeginBlock, dwSectorSize, &dwReadWrite, 0) && dwReadWrite == dwSectorSize)
                                                        wError = L"Can't restore first sector";
                                                    if (SetFilePointer(hFile, dwSkipByte, 0, FILE_BEGIN) != INVALID_SET_FILE_POINTER)
                                                    {
                                                        assert(pByteTemp);
                                                        *pByteTemp = btAttribPrev;
                                                        if (!(WriteFile(hFile, pBtSectorWrite, dwSectorSize, &dwReadWrite, 0) && dwReadWrite == dwSectorSize))
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
                            else if (GetFileAttributes(g_wAutorunFile) != INVALID_FILE_ATTRIBUTES)        //checking for success
                            {
                                hFile = CreateFile(g_wVolume, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0);
                                if (hFile != INVALID_HANDLE_VALUE)
                                {
                                    assert(dwSkipByte);
                                    if (SetFilePointer(hFile, dwSkipByte, 0, FILE_BEGIN) != INVALID_SET_FILE_POINTER)
                                    {
                                        assert(pByteTemp);
                                        *pByteTemp = btAttribPrev;
                                        wError = (WriteFile(hFile, pBtSectorWrite, dwSectorSize, &dwReadWrite, 0) && dwReadWrite == dwSectorSize) ? L"Error set attributes" : L"Can't write drive";
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
                MessageBox(hWndOwner, wError, L"FlashDriveProtect", MB_ICONWARNING);
        }
        else if (eMode == eMsgShowAll)
            MessageBox(hWndOwner, L"Vaccinated successfully", L"FlashDriveProtect", MB_ICONINFORMATION);
    }
}

//-------------------------------------------------------------------------------------------------
bool fCheckVaccinateComputer()
{
    DWORD dwValue,
            dwSize = sizeof(DWORD);
    LONG iRes = RegGetValuePointer(HKEY_LOCAL_MACHINE, L"SYSTEM\\CurrentControlSet\\Services\\cdrom", L"AutoRun", RRF_RT_REG_DWORD, 0, &dwValue, &dwSize);
    if ((iRes == ERROR_SUCCESS && dwValue == 0) || iRes == ERROR_FILE_NOT_FOUND)
        if (RegGetValuePointer(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer", L"NoDriveTypeAutoRun", RRF_RT_REG_DWORD, 0, &dwValue, &dwSize) == ERROR_SUCCESS && dwValue == 0xFF)
        {
            HKEY hKey;
            if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\IniFileMapping\\Autorun.inf", 0, KEY_QUERY_VALUE | KEY_WOW64_64KEY, &hKey) == ERROR_SUCCESS)
            {
                wchar_t wBuf[18*sizeof(wchar_t)];
                dwSize = 18*sizeof(wchar_t);
                iRes = RegGetValuePointer(hKey, 0, 0, g_wWinVer >= _WIN32_WINNT_WS03 ? (RRF_RT_REG_SZ | RRF_SUBKEY_WOW6464KEY) : RRF_RT_REG_SZ, 0, wBuf, &dwSize) == ERROR_SUCCESS && dwSize == 18*sizeof(wchar_t) && wcscmp(wBuf, L"@SYS:DoesNotExist") == 0;
                RegCloseKey(hKey);
                if (iRes && RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\IniFileMapping\\Autorun.inf", 0, KEY_QUERY_VALUE | KEY_WOW64_32KEY, &hKey) == ERROR_SUCCESS)
                {
                    iRes = RegGetValuePointer(hKey, 0, 0, g_wWinVer >= _WIN32_WINNT_WS03 ? (RRF_RT_REG_SZ | RRF_SUBKEY_WOW6432KEY) : RRF_RT_REG_SZ, 0, wBuf, &dwSize) == ERROR_SUCCESS && dwSize == 18*sizeof(wchar_t) && wcscmp(wBuf, L"@SYS:DoesNotExist") == 0;
                    RegCloseKey(hKey);
                    if (iRes)
                    {
                        dwSize = sizeof(wchar_t);
                        return RegGetValuePointer(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\AutoplayHandlers\\CancelAutoplay\\Files", L"*.*", RRF_RT_REG_SZ, 0, wBuf, &dwSize) == ERROR_SUCCESS && dwSize == sizeof(wchar_t) && *wBuf == L'\0';
                    }
                }
            }
        }
    return false;
}

//-------------------------------------------------------------------------------------------------
inline const BYTE *fDwordToByte(const DWORD *const pDword)
{
    return static_cast<const BYTE*>(static_cast<const void*>(pDword));
}

//-------------------------------------------------------------------------------------------------
void fVaccinateComputer()
{
    DWORD dwValue = 0;
    HKEY hKey;
    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"SYSTEM\\CurrentControlSet\\Services\\cdrom", 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, L"AutoRun", 0, REG_DWORD, fDwordToByte(&dwValue), sizeof(DWORD));
        RegCloseKey(hKey);
    }
    if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer", 0, 0, 0, KEY_WRITE, 0, &hKey, 0) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, L"NoDriveTypeAutoRun", 0, REG_DWORD, fDwordToByte(&(dwValue = 0xFF)), sizeof(DWORD));
        RegCloseKey(hKey);
    }
    if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\IniFileMapping\\Autorun.inf", 0, 0, 0, KEY_WRITE | KEY_WOW64_64KEY, 0, &hKey, 0) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, 0, 0, REG_SZ, static_cast<const BYTE*>(static_cast<const void*>(L"@SYS:DoesNotExist")), 18*sizeof(wchar_t));
        RegCloseKey(hKey);
    }
    if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\IniFileMapping\\Autorun.inf", 0, 0, 0, KEY_WRITE | KEY_WOW64_32KEY, 0, &hKey, 0) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, 0, 0, REG_SZ, static_cast<const BYTE*>(static_cast<const void*>(L"@SYS:DoesNotExist")), 18*sizeof(wchar_t));
        RegCloseKey(hKey);
    }
    if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\AutoplayHandlers\\CancelAutoplay\\Files", 0, 0, 0, KEY_WRITE, 0, &hKey, 0) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, L"*.*", 0, REG_SZ, 0, 0);
        RegCloseKey(hKey);
    }
}

//-------------------------------------------------------------------------------------------------
void fUnVaccinateComputer()
{
    HKEY hKey;
    DWORD dwValue = 1;
    if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"SYSTEM\\CurrentControlSet\\Services\\cdrom", 0, 0, 0, KEY_WRITE, 0, &hKey, 0) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, L"AutoRun", 0, REG_DWORD, fDwordToByte(&dwValue), sizeof(DWORD));
        RegCloseKey(hKey);
    }
    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer", 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, L"NoDriveTypeAutoRun", 0, REG_DWORD, fDwordToByte(&(dwValue = 0)), sizeof(DWORD));
        RegCloseKey(hKey);
    }
    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\IniFileMapping\\Autorun.inf", 0, KEY_SET_VALUE | KEY_WOW64_64KEY, &hKey) == ERROR_SUCCESS)
    {
        RegDeleteValue(hKey, 0);
        RegCloseKey(hKey);
    }
    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\IniFileMapping\\Autorun.inf", 0, KEY_SET_VALUE | KEY_WOW64_32KEY, &hKey) == ERROR_SUCCESS)
    {
        RegDeleteValue(hKey, 0);
        RegCloseKey(hKey);
    }
    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\AutoplayHandlers\\CancelAutoplay\\Files", 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS)
    {
        RegDeleteValue(hKey, L"*.*");
        RegCloseKey(hKey);
    }
}

//-------------------------------------------------------------------------------------------------
void fUpdateListDrives(const HWND hWndListDrives)
{
    int iSaveSel[eDriveCountMax];
    const LRESULT iCountSel = SendMessage(hWndListDrives, LB_GETSELITEMS, eDriveCountMax, reinterpret_cast<LPARAM>(iSaveSel));
    assert(iCountSel >= 0);
    LRESULT i = 0;
    for (; i < iCountSel; ++i)
        iSaveSel[i] = SendMessage(hWndListDrives, LB_GETITEMDATA, iSaveSel[i], 0) & 0xFFFF;        //replace the item number by the letter

    SendMessage(hWndListDrives, LB_RESETCONTENT, 0, 0);
    DWORD dwDrives = GetLogicalDrives();
    wchar_t wFileSystem[6/*FAT32`*/];
    *g_wDrive = L'A';
    WIN32_FIND_DATA findFileData;
    do
    {
        if ((dwDrives & 1) && GetDriveType(g_wDrive) == DRIVE_REMOVABLE &&
                GetVolumeInformation(g_wDrive, g_wRowDrive+5, 12/*max*/, 0, 0, 0, wFileSystem, 6/*FAT32`*/) &&
                (wcscmp(wFileSystem, L"FAT32") == 0 || wcscmp(wFileSystem, L"FAT") == 0))
        {
            g_wRowDrive[1] = *g_wAutorunFile = *g_wDrive;

            int iIsVaccinated = 0;
            HANDLE hFile = FindFirstFile(g_wAutorunFile, &findFileData);
            if (hFile != INVALID_HANDLE_VALUE)
            {
                FindClose(hFile);
                if (GetFileAttributes(g_wAutorunFile) == INVALID_FILE_ATTRIBUTES)
                {
                    wcscat(g_wRowDrive, L" [vaccinated]");
                    iIsVaccinated = 0x10000;
                }
            }
            const LRESULT iRes = SendMessage(hWndListDrives, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(g_wRowDrive));
            assert(iRes >= 0);
            if (iRes >= 0)
            {
                //is selected before
                SendMessage(hWndListDrives, LB_SETITEMDATA, iRes, iIsVaccinated | *g_wDrive);
                for (i = 0; i < iCountSel; ++i)
                    if (*g_wDrive == iSaveSel[i])
                    {
                        SendMessage(hWndListDrives, LB_SETSEL, TRUE, iRes);
                        break;
                    }
            }
        }
        dwDrives >>= 1;
    } while (++*g_wDrive <= L'Z');
}

//-------------------------------------------------------------------------------------------------
LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    enum
    {
        eVaccinateComp = 1,
        eVaccinateDrives,
        eUpdate,
        eHelp,
        eListBoxWidth = 239,
        eListBoxTop = 43
    };

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
                (hIconSuccess = static_cast<HICON>(LoadImage(hInst, L"IDI_ICON2", IMAGE_ICON, 16, 16, 0)))  &&
                (hIconWarning = static_cast<HICON>(LoadImage(hInst, L"IDI_ICON3", IMAGE_ICON, 16, 16, 0))) &&
                (hIconVaccinate = static_cast<HICON>(LoadImage(hInst, L"IDI_ICON4", IMAGE_ICON, 16, 16, 0))) &&
                (hIconRemove = static_cast<HICON>(LoadImage(hInst, L"IDI_ICON5", IMAGE_ICON, 16, 16, 0))) &&
                (hIconVaccinateBig = static_cast<HICON>(LoadImage(hInst, L"IDI_ICON4", IMAGE_ICON, 32, 32, 0))) &&
                (hIconUpdate = static_cast<HICON>(LoadImage(hInst, L"IDI_ICON6", IMAGE_ICON, 16, 16, 0))) &&
                (hIconHelp = static_cast<HICON>(LoadImage(hInst, L"IDI_ICON7", IMAGE_ICON, 16, 16, 0))) &&
                (hWndStatusIcon = CreateWindowEx(0, WC_STATIC, 0, WS_CHILD | WS_VISIBLE | SS_ICON | SS_REALSIZEIMAGE, 9, 8, 16, 16, hWnd, 0, hInst, 0)) &&
                (hWndStatusCompVaccination = CreateWindowEx(0, WC_STATIC, 0, WS_CHILD | WS_VISIBLE, 30, 8, 156, 18, hWnd, 0, hInst, 0)) &&
                (hWndListDrives = CreateWindowEx(0, WC_LISTBOX, 0, WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | LBS_EXTENDEDSEL | LBS_NOINTEGRALHEIGHT, 5, eListBoxTop, eListBoxWidth, 70, hWnd, 0, hInst, 0)) &&
                (hWndBtnVaccinateComp = CreateWindowEx(0, WC_BUTTON, 0, WS_CHILD | WS_VISIBLE, 189, 5, 150, 22, hWnd, reinterpret_cast<HMENU>(eVaccinateComp), hInst, 0)))
            if (const HWND hWndBtnVaccinateDrive = CreateWindowEx(0, WC_BUTTON, L"Vaccinate", WS_CHILD | WS_VISIBLE, 5+eListBoxWidth+3, 43, 150, 44, hWnd, reinterpret_cast<HMENU>(eVaccinateDrives), hInst, 0))
                if (const HWND hWndUpdate = CreateWindowEx(0, WC_BUTTON, 0, WS_CHILD | WS_VISIBLE | BS_ICON, 342, 5, 26, 22, hWnd, reinterpret_cast<HMENU>(eUpdate), hInst, 0))
                    if (const HWND hWndHelp = CreateWindowEx(0, WC_BUTTON, 0, WS_CHILD | WS_VISIBLE | BS_ICON, 371, 5, 26, 22, hWnd, reinterpret_cast<HMENU>(eHelp), hInst, 0))
                        if (SendMessage(hWnd, WM_COMMAND, eUpdate, reinterpret_cast<LPARAM>(hWndUpdate)) == 0)
                        {
                            NONCLIENTMETRICS nonClientMetrics;
                            nonClientMetrics.cbSize = sizeof(NONCLIENTMETRICS);
                            if (SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &nonClientMetrics, 0))
                            {
                                nonClientMetrics.lfMessageFont.lfHeight = -13;
                                if ((hFont = CreateFontIndirect(&nonClientMetrics.lfMessageFont)))
                                {
                                    SendMessage(hWndStatusIcon, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);
                                    SendMessage(hWndStatusCompVaccination, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);
                                    SendMessage(hWndListDrives, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);
                                    SendMessage(hWndBtnVaccinateComp, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);
                                    SendMessage(hWndBtnVaccinateDrive, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);
                                    SendMessage(hWndUpdate, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);
                                    SendMessage(hWndHelp, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);

                                    SendMessage(hWndBtnVaccinateDrive, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconVaccinateBig));
                                    SendMessage(hWndUpdate, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconUpdate));
                                    SendMessage(hWndHelp, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconHelp));

                                    //select first unvaccinate drive
                                    const LRESULT iCount = SendMessage(hWndListDrives, LB_GETCOUNT, 0, 0);
                                    assert(iCount >= 0);
                                    for (LRESULT i = 0; i < iCount; ++i)
                                        if ((SendMessage(hWndListDrives, LB_GETITEMDATA, i, 0) | 0xFFFF) == 0xFFFF)
                                        {
                                            SendMessage(hWndListDrives, LB_SETSEL, TRUE, i);
                                            break;
                                        }
                                    SetFocus(hWndBtnVaccinateDrive);
                                    return 0;
                                }
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
                if (MessageBox(hWnd, L"This operation is not recommended.\nAre you sure?", L"FlashDriveProtect", MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2) == IDYES)
                {
                    fUnVaccinateComputer();
                    bIsVaccinatedComputer = fCheckVaccinateComputer();
                    if (bIsVaccinatedComputer)
                        MessageBox(hWnd, L"Failed to remove vaccine", L"FlashDriveProtect", MB_ICONWARNING);
                    else
                    {
                        SendMessage(hWndStatusIcon, STM_SETICON, reinterpret_cast<WPARAM>(hIconWarning), 0);
                        SetWindowText(hWndStatusCompVaccination, L"Computer not vaccinated");
                        SetWindowText(hWndBtnVaccinateComp, L"Vaccinate Computer");
                        SendMessage(hWndBtnVaccinateComp, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconVaccinate));
                    }
                }
            }
            else        //vaccinate
            {
                fVaccinateComputer();
                bIsVaccinatedComputer = fCheckVaccinateComputer();
                if (bIsVaccinatedComputer)
                {
                    SendMessage(hWndStatusIcon, STM_SETICON, reinterpret_cast<WPARAM>(hIconSuccess), 0);
                    SetWindowText(hWndStatusCompVaccination, L"Computer vaccinated");
                    SetWindowText(hWndBtnVaccinateComp, L"Remove Vaccine");
                    SendMessage(hWndBtnVaccinateComp, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconRemove));
                }
                else
                    MessageBox(hWnd, L"Failed to vaccinate computer", L"FlashDriveProtect", MB_ICONWARNING);
            }
            break;
        }
        case eVaccinateDrives:
        {
            int iSaveSel[eDriveCountMax];
            const LRESULT iCount = SendMessage(hWndListDrives, LB_GETSELITEMS, eDriveCountMax, reinterpret_cast<LPARAM>(iSaveSel));
            assert(iCount >= 0);
            for (LRESULT i = 0, iRes; i < iCount; ++i)
            {
                iRes = SendMessage(hWndListDrives, LB_GETITEMDATA, iSaveSel[i], 0);
                assert(iRes >= 0);
                if ((iRes >> 16) == 0)
                    fVaccinateDrive(iRes, hWnd, eMsgWarningOnly);
            }
            fUpdateListDrives(hWndListDrives);
            break;
        }
        case eUpdate:
        {
            bIsVaccinatedComputer = fCheckVaccinateComputer();
            if (bIsVaccinatedComputer)
            {
                SendMessage(hWndStatusIcon, STM_SETICON, reinterpret_cast<WPARAM>(hIconSuccess), 0);
                SetWindowText(hWndStatusCompVaccination, L"Computer vaccinated");
                SetWindowText(hWndBtnVaccinateComp, L"Remove Vaccine");
                SendMessage(hWndBtnVaccinateComp, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconRemove));
            }
            else
            {
                SendMessage(hWndStatusIcon, STM_SETICON, reinterpret_cast<WPARAM>(hIconWarning), 0);
                SetWindowText(hWndStatusCompVaccination, L"Computer not vaccinated");
                SetWindowText(hWndBtnVaccinateComp, L"Vaccinate Computer");
                SendMessage(hWndBtnVaccinateComp, BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(hIconVaccinate));
            }
            fUpdateListDrives(hWndListDrives);
            break;
        }
        case eHelp:
        {
            MessageBox(hWnd, wHelp, L"FlashDriveProtect", MB_ICONINFORMATION);
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
        SetWindowPos(hWndListDrives, HWND_TOP, 0, 0, eListBoxWidth, HIWORD(lParam)-eListBoxTop-5, SWP_NOMOVE | SWP_NOZORDER);
        return 0;
    }
    case WM_DEVICECHANGE:
    {
        if (wParam == DBT_DEVNODES_CHANGED && SetTimer(hWnd, 1, 3000, 0))
            bTimerActive = true;
        return TRUE;
    }
    case WM_TIMER:
    {
        if (KillTimer(hWnd, 1))
        {
            bTimerActive = false;
            fUpdateListDrives(hWndListDrives);
        }
        return 0;
    }
    case WM_ENDSESSION:
    {
        SendMessage(hWnd, WM_CLOSE, 0, 0);
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

        const wchar_t *const wAppPath = g_wVolume+eBufLen;
        if (*wAppPath)
        {
            RECT rect;
            if (GetWindowRect(hWnd, &rect))
            {
                const HANDLE hFile = CreateFile(wAppPath, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
                if (hFile != INVALID_HANDLE_VALUE)
                {
                    DWORD dwBytesWrite;
                    WriteFile(hFile, &rect, sizeof(RECT), &dwBytesWrite, 0);
                    CloseHandle(hFile);
                }
            }
        }

        PostQuitMessage(0);
        return 0;
    }
    case WM_CTLCOLORSTATIC:
    {
        if (reinterpret_cast<HWND>(lParam) == hWndStatusCompVaccination)
        {
            SetBkMode(reinterpret_cast<HDC>(wParam), TRANSPARENT);
            SetTextColor(reinterpret_cast<HDC>(wParam), bIsVaccinatedComputer ? RGB(0, 85, 0) : RGB(255, 0, 0));
            return reinterpret_cast<LRESULT>(hbrSysBackground);
        }
    }
    }
    return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

//-------------------------------------------------------------------------------------------------
void fVaccinateAllDrives(const EMode eMode)
{
    DWORD dwDrives = GetLogicalDrives();
    wchar_t wDrive = L'A';
    do
    {
        if (dwDrives & 1)
            fVaccinateDrive(wDrive, 0, eMode);
        dwDrives >>= 1;
    } while (++wDrive <= L'Z');
}

//-------------------------------------------------------------------------------------------------
LRESULT CALLBACK WindowProcResident(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    static EMode eMode;
    static bool bTimerActive = false;
    switch (uMsg)
    {
    case WM_CREATE:
    {
        eMode = static_cast<EMode>(reinterpret_cast<size_t>(reinterpret_cast<const CREATESTRUCT*>(lParam)->lpCreateParams));
        fVaccinateAllDrives(eMode);
        return 0;
    }
    case WM_DEVICECHANGE:
    {
        if (wParam == DBT_DEVNODES_CHANGED && SetTimer(hWnd, 1, 3000, 0))
            bTimerActive = true;
        return TRUE;
    }
    case WM_TIMER:
    {
        if (KillTimer(hWnd, 1))
        {
            bTimerActive = false;
            fVaccinateAllDrives(eMode);
        }
        return 0;
    }
    case WM_ENDSESSION:
    {
        SendMessage(hWnd, WM_CLOSE, 0, 0);
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
    return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

//-------------------------------------------------------------------------------------------------
const wchar_t* fGetArgument(EMode *const eMode)
{
    if (wchar_t *wCmdLine = GetCommandLine())
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
                        return 0;
                    ++wCmdLine;
                }
                ++wCmdLine;
                if (*wCmdLine != L' ' && *wCmdLine != L'\t')
                    return 0;
            }
            else
                while (*wCmdLine != L' ' && *wCmdLine != L'\t')
                {
                    if (*wCmdLine == L'\0' || *wCmdLine == L'\"')
                        return 0;
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
                            return 0;
                        ++wCmdLine;
                    }
                    if (wCmdLine[1] != L' ' && wCmdLine[1] != L'\t' && wCmdLine[1] != L'\0')
                        return 0;
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

                        if (wcsncmp(wArg2, L"/mode:", 6) == 0)
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
    return 0;
}

//-------------------------------------------------------------------------------------------------
void fGetRegGetValuePointer()
{
    if (const HMODULE hMod = GetModuleHandle(g_wWinVer >= _WIN32_WINNT_WS03 ? L"advapi32.dll" : L"shlwapi.dll"))
        RegGetValuePointer = reinterpret_cast<PRegGetValuePointer>(GetProcAddress(hMod, g_wWinVer >= _WIN32_WINNT_WS03 ? "RegGetValueW" : "SHRegGetValueW"));
}

//-------------------------------------------------------------------------------------------------
int main()
{
    assert(sizeof(wchar_t) == 2);
    assert(eBufSize%eSectorSizeMax == 0 && eBufSize >= eSectorSizeMax*2);

    INITCOMMONCONTROLSEX initComCtrlEx;
    initComCtrlEx.dwSize = sizeof(INITCOMMONCONTROLSEX);
    initComCtrlEx.dwICC = ICC_STANDARD_CLASSES;
    if (InitCommonControlsEx(&initComCtrlEx) == TRUE)
    {
        EMode eMode = eMsgWarningOnly;
        const wchar_t *pArg = fGetArgument(&eMode);
        const HWND hWndResident = FindWindow(g_wGuidClassResident, 0);
        if (pArg && *pArg == L'q' && pArg[1] == L'\0')
        {
            assert(pArg[-1] == L'/');
            if (hWndResident)
                PostMessage(hWndResident, WM_CLOSE, 0, 0);
            return 0;
        }
        if (hWndResident)
        {
            if (MessageBox(0, L"Program already running in resident mode.\nClose it?",
                           L"FlashDriveProtect", MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2) != IDYES)
                return 0;
            PostMessage(hWndResident, WM_CLOSE, 0, 0);
        }

        wchar_t wBuf[eBufLen+MAX_PATH] = L"\\\\.\\X:\0X:\\\0X:\\AUTORUN.INF\0(X:) ";//[eBufLen = "\\.\X:`X:\`X:\AUTORUN.INF`(X:) 12345678901 [vaccinated]`"]
        g_wVolume = wBuf;        //***it's ok
        g_wDrive = wBuf+7;
        g_wAutorunFile = wBuf+11;
        g_wRowDrive = wBuf+26;
        g_wWinVer = GetVersion();
        g_wWinVer = g_wWinVer << 8 | g_wWinVer >> 8;

        BYTE btZeroes[eSectorSizeMax];
        g_pBtZeroes = btZeroes;        //***it's ok
        memset(g_pBtZeroes, 0, eSectorSizeMax);

        if (pArg)
        {
            assert(pArg[-1] == L'/');
            if (wcsncmp(pArg, L"drives", 6) == 0)
            {
                if (pArg[6] == L':')        //selected drives
                {
                    if ((g_pBtBuffer = static_cast<BYTE*>(malloc(eBufSize))))
                    {
                        pArg += 7;
                        while (*pArg >= L'A' && *pArg <= L'Z')
                        {
                            fVaccinateDrive(*pArg, 0, eMode);
                            ++pArg;
                        }
                        free(g_pBtBuffer);
                    }
                    return 0;
                }
                else if (pArg[6] == L'+' && pArg[7] == L'\0')        //all drives
                {
                    if ((g_pBtBuffer = static_cast<BYTE*>(malloc(eBufSize))))
                    {
                        fVaccinateAllDrives(eMode);
                        free(g_pBtBuffer);
                    }
                    return 0;
                }
            }
            else if (wcsncmp(pArg, L"system", 6) == 0)
            {
                pArg += 6;
                if (*pArg == L'+' && pArg[1] == L'\0')
                {
                    fVaccinateComputer();
                    if (eMode != eMsgHideAll)
                    {
                        fGetRegGetValuePointer();
                        if (RegGetValuePointer)
                        {
                            if (fCheckVaccinateComputer())
                            {
                                if (eMode == eMsgShowAll)
                                    MessageBox(0, L"System successfully vaccinated", L"FlashDriveProtect", MB_ICONINFORMATION);
                            }
                            else
                                MessageBox(0, L"Failed to vaccinate computer", L"FlashDriveProtect", MB_ICONWARNING);
                        }
                    }
                    return 0;
                }
                else if (*pArg == L'-' && pArg[1] == L'\0')
                {
                    fUnVaccinateComputer();
                    if (eMode != eMsgHideAll)
                    {
                        fGetRegGetValuePointer();
                        if (RegGetValuePointer)
                        {
                            if (fCheckVaccinateComputer())
                                MessageBox(0, L"Failed to remove vaccine", L"FlashDriveProtect", MB_ICONWARNING);
                            else if (eMode == eMsgShowAll)
                                MessageBox(0, L"Successfully remove vaccine from system", L"FlashDriveProtect", MB_ICONINFORMATION);
                        }
                    }
                    return 0;
                }
            }
            else if (wcscmp(pArg, L"resident") == 0)
            {
                if ((g_pBtBuffer = static_cast<BYTE*>(malloc(eBufSize))))
                {
                    WNDCLASSEX wndCl;
                    memset(&wndCl, 0, sizeof(WNDCLASSEX));
                    wndCl.cbSize = sizeof(WNDCLASSEX);
                    wndCl.lpfnWndProc = WindowProcResident;
                    wndCl.hInstance = GetModuleHandle(0);
                    wndCl.lpszClassName = g_wGuidClassResident;

                    if (RegisterClassEx(&wndCl))
                    {
                        if (CreateWindowEx(0, g_wGuidClassResident, 0, 0, 0, 0, 0, 0, 0, 0, wndCl.hInstance, reinterpret_cast<PVOID>(eMode)))
                        {
                            MSG msg;
                            while (GetMessage(&msg, 0, 0, 0) > 0)
                                DispatchMessage(&msg);
                        }
                        UnregisterClass(g_wGuidClassResident, wndCl.hInstance);
                    }
                    free(g_pBtBuffer);
                }
                return 0;
            }
        }

        RECT rect; rect.left = 0; rect.top = 0; rect.right = 402; rect.bottom = 92;
        if (AdjustWindowRectEx(&rect, WS_OVERLAPPEDWINDOW | WS_VISIBLE, FALSE, 0))
        {
            fGetRegGetValuePointer();
            if (RegGetValuePointer && (g_pBtBuffer = static_cast<BYTE*>(malloc(eBufSize))))
            {
                const HINSTANCE hInst = GetModuleHandle(0);
                WNDCLASSEX wndCl;
                wndCl.cbSize = sizeof(WNDCLASSEX);
                wndCl.style = 0;
                wndCl.lpfnWndProc = WindowProc;
                wndCl.cbClsExtra = 0;
                wndCl.cbWndExtra = 0;
                wndCl.hInstance = hInst;
                wndCl.hIcon = static_cast<HICON>(LoadImage(hInst, L"IDI_ICON1", IMAGE_ICON, 0, 0, LR_DEFAULTSIZE));
                wndCl.hCursor = LoadCursor(0, IDC_ARROW);
                wndCl.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE+1);
                wndCl.lpszMenuName = 0;
                wndCl.lpszClassName = g_wGuidClass;
                wndCl.hIconSm = static_cast<HICON>(LoadImage(hInst, L"IDI_ICON1", IMAGE_ICON, 16, 16, 0));

                if (RegisterClassEx(&wndCl))
                {
                    g_iFixWidth = rect.right-rect.left;
                    g_iMinHeight = rect.bottom-rect.top;

                    rect.left = CW_USEDEFAULT;

                    wchar_t *const wAppPath = wBuf+56;
                    DWORD dwSize = GetModuleFileName(0, wAppPath, MAX_PATH+1-4/*.geo*/);
                    if (dwSize >= 4 && dwSize < MAX_PATH-4/*.geo*/)
                    {
                        wcscpy(wAppPath+dwSize, L".geo");
                        const HANDLE hFile = CreateFile(wAppPath, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
                        if (hFile != INVALID_HANDLE_VALUE)
                        {
                            LARGE_INTEGER iFileSize;
                            if (GetFileSizeEx(hFile, &iFileSize) && iFileSize.QuadPart == sizeof(RECT) &&
                                    !(ReadFile(hFile, &rect, sizeof(RECT), &dwSize, 0) && dwSize == sizeof(RECT)))
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

                    if (CreateWindowEx(0, g_wGuidClass, L"FlashDriveProtect", WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                                       rect.left, rect.top, g_iFixWidth, rect.bottom, 0, 0, hInst, 0))
                    {
                        MSG msg;
                        while (GetMessage(&msg, 0, 0, 0) > 0)
                            DispatchMessage(&msg);
                    }
                    UnregisterClass(g_wGuidClass, hInst);
                }
                free(g_pBtBuffer);
            }
        }
    }
    return 0;
}
