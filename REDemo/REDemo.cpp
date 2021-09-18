// REDemo.cpp : Defines the entry point for the application.
//

#define _CRT_SECURE_NO_WARNINGS

#include "framework.h"
#include "REDemo.h"
#include "RichEdit.h"
#include "Richole.h"
#include "commdlg.h"
#include "windowsx.h"
#include <Shlobj.h>

#include <string>
#include <fstream>
#include <chrono>

#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR    lpCmdLine,
    _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: Place code here.

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_REDEMO, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    if (!InitInstance(hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_REDEMO));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int)msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.cbClsExtra = 0;
    wcex.cbWndExtra = 0;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_REDEMO));
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszMenuName = MAKEINTRESOURCEW(IDC_REDEMO);
    wcex.lpszClassName = szWindowClass;
    wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
    hInst = hInstance; // Store instance handle in our global variable

    HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

    if (!hWnd)
    {
        return FALSE;
    }

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    return TRUE;
}

BOOL InsertImage(HWND hRichEdit, LPCTSTR pszImage)
{
    HRESULT hr;
    LPRICHEDITOLE pRichEditOle;
    SendMessage(hRichEdit, EM_GETOLEINTERFACE, 0, (LPARAM)&pRichEditOle);
    if (pRichEditOle == NULL)
    {
        return FALSE;
    }

    LPLOCKBYTES pLockBytes = NULL;
    hr = CreateILockBytesOnHGlobal(NULL, TRUE, &pLockBytes);
    if (FAILED(hr))
    {
        return FALSE;
    }

    LPSTORAGE pStorage;
    hr = StgCreateDocfileOnILockBytes(pLockBytes,
        STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_READWRITE,
        0, &pStorage);
    if (FAILED(hr))
    {
        return FALSE;
    }

    FORMATETC formatEtc;
    formatEtc.cfFormat = 0;
    formatEtc.ptd = NULL;
    formatEtc.dwAspect = DVASPECT_CONTENT;
    formatEtc.lindex = -1;
    formatEtc.tymed = TYMED_NULL;
    LPOLECLIENTSITE pClientSite;
    hr = pRichEditOle->GetClientSite(&pClientSite);
    if (FAILED(hr))
    {
        return FALSE;
    }

    LPUNKNOWN pUnk;
    CLSID clsid = CLSID_NULL;
    hr = OleCreateFromFile(clsid, pszImage, IID_IUnknown, OLERENDER_DRAW,
        &formatEtc, pClientSite, pStorage, (void**)&pUnk);
    pClientSite->Release();
    if (FAILED(hr))
    {
        return FALSE;
    }

    LPOLEOBJECT pObject;
    hr = pUnk->QueryInterface(IID_IOleObject, (void**)&pObject);
    pUnk->Release();
    if (FAILED(hr))
    {
        return FALSE;
    }

    OleSetContainedObject(pObject, TRUE);
    REOBJECT reobject = { sizeof(REOBJECT) };
    hr = pObject->GetUserClassID(&clsid);
    if (FAILED(hr))
    {
        pObject->Release();
        return FALSE;
    }

    reobject.clsid = clsid;
    reobject.cp = REO_CP_SELECTION;
    reobject.dvaspect = DVASPECT_CONTENT;
    reobject.dwFlags = REO_RESIZABLE | REO_BELOWBASELINE;
    reobject.dwUser = 0;
    reobject.poleobj = pObject;
    reobject.polesite = pClientSite;
    reobject.pstg = pStorage;
    SIZEL sizel = { 0 };
    reobject.sizel = sizel;
    SendMessage(hRichEdit, EM_SETSEL, 0, -1);
    DWORD dwStart, dwEnd;
    SendMessage(hRichEdit, EM_GETSEL, (WPARAM)&dwStart, (LPARAM)&dwEnd);
    SendMessage(hRichEdit, EM_SETSEL, dwEnd + 1, dwEnd + 1);
    SendMessage(hRichEdit, EM_REPLACESEL, TRUE, (WPARAM)L"\n");
    hr = pRichEditOle->InsertObject(&reobject);
    pObject->Release();
    pRichEditOle->Release();
    if (FAILED(hr))
    {
        return FALSE;
    }
    return TRUE;
}

BOOL InsertObject(HWND hRichEdit, LPCTSTR pszFileName)
{
    HRESULT hr;
    LPRICHEDITOLE pRichEditOle;
    SendMessage(hRichEdit, EM_GETOLEINTERFACE, 0, (LPARAM)&pRichEditOle);
    if (pRichEditOle == NULL)
    {
        return FALSE;
    }

    LPLOCKBYTES pLockBytes = NULL;
    hr = CreateILockBytesOnHGlobal(NULL, TRUE, &pLockBytes);
    if (FAILED(hr))
    {
        return FALSE;
    }

    LPSTORAGE pStorage;
    hr = StgCreateDocfileOnILockBytes(pLockBytes,
        STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_READWRITE,
        0, &pStorage);
    if (FAILED(hr))
    {
        return FALSE;
    }

    FORMATETC formatEtc;
    formatEtc.cfFormat = 0;
    formatEtc.ptd = NULL;
    formatEtc.dwAspect = DVASPECT_CONTENT;
    formatEtc.lindex = -1;
    formatEtc.tymed = TYMED_NULL;
    LPOLECLIENTSITE pClientSite;
    hr = pRichEditOle->GetClientSite(&pClientSite);
    if (FAILED(hr))
    {
        return FALSE;
    }

    LPUNKNOWN pUnk;
    CLSID clsid = CLSID_NULL;
    hr = OleCreateFromFile(clsid, pszFileName, IID_IUnknown, OLERENDER_DRAW,
        &formatEtc, pClientSite, pStorage, (void**)&pUnk);
    pClientSite->Release();
    if (FAILED(hr))
    {
        return FALSE;
    }

    LPOLEOBJECT pObject;
    hr = pUnk->QueryInterface(IID_IOleObject, (void**)&pObject);
    pUnk->Release();
    if (FAILED(hr))
    {
        return FALSE;
    }

    OleSetContainedObject(pObject, TRUE);
    REOBJECT reobject = { sizeof(REOBJECT) };
    hr = pObject->GetUserClassID(&clsid);
    if (FAILED(hr))
    {
        pObject->Release();
        return FALSE;
    }

    reobject.clsid = clsid;
    reobject.cp = REO_CP_SELECTION;
    reobject.dvaspect = DVASPECT_CONTENT;
    reobject.dwFlags = REO_RESIZABLE | REO_BELOWBASELINE;
    reobject.dwUser = 0;
    reobject.poleobj = pObject;
    reobject.polesite = pClientSite;
    reobject.pstg = pStorage;
    SIZEL sizel = { 0 };
    reobject.sizel = sizel;
    SendMessage(hRichEdit, EM_SETSEL, 0, -1);
    DWORD dwStart, dwEnd;
    SendMessage(hRichEdit, EM_GETSEL, (WPARAM)&dwStart, (LPARAM)&dwEnd);
    SendMessage(hRichEdit, EM_SETSEL, dwEnd + 1, dwEnd + 1);
    SendMessage(hRichEdit, EM_REPLACESEL, TRUE, (WPARAM)L"\n");
    hr = pRichEditOle->InsertObject(&reobject);
    pObject->Release();
    pRichEditOle->Release();
    if (FAILED(hr))
    {
        return FALSE;
    }
    return TRUE;
}

std::wstring SelectFile(HWND hwnd, PCWSTR ext)
{
    OPENFILENAMEW ofn;
    std::wstring name_buf;
    name_buf.resize(1024);

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFile = &name_buf[0];
    // Set lpstrFile[0] to '\0' so that GetOpenFileName does not
    // use the contents of szFile to initialize itself.
    ofn.lpstrFile[0] = '\0';
    ofn.nMaxFile = name_buf.size();
    ofn.lpstrFilter = ext;
    ofn.nFilterIndex = 2;
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.lpstrInitialDir = NULL;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
    if (!GetOpenFileName(&ofn))
    {
        name_buf.clear();
    }

    return name_buf;
}

BOOL RE_InsertBitmapFromOleClipBoard(HWND hwnd, HWND hwndEdit)
{
    HRESULT hr;
    IDataObject* pDataObject;
    LPOLEOBJECT lpOleObject = NULL;
    LPSTORAGE lpStorage = NULL;
    LPOLECLIENTSITE lpOleClientSite = NULL;
    LPLOCKBYTES lpLockBytes = NULL;
    CLIPFORMAT cfFormat = 0;
    LPFORMATETC lpFormatEtc = NULL;
    FORMATETC formatEtc;
    LPRICHEDITOLE lpRichEditOle;
    SCODE sc;
    LPUNKNOWN lpUnknown;
    REOBJECT reobject;
    CLSID clsid;
    SIZEL sizel;
    DWORD dwStart, dwEnd;

    CoInitialize(NULL);
    //if (OpenClipboard(NULL))
    //{
    //    EmptyClipboard();
    //    SetClipboardData(CF_BITMAP, hBitmap);
    //    CloseClipboard();
    //}
    //else
    //    return 0;

    OleGetClipboard(&pDataObject);
    if (pDataObject == NULL)
        return 0;

    // 得到 HBITMAP
    FORMATETC fc;
    fc.cfFormat = CF_BITMAP;
    fc.ptd = NULL;
    fc.dwAspect = DVASPECT_CONTENT;
    fc.lindex = -1;
    fc.tymed = TYMED_GDI;
    STGMEDIUM stg;
    stg.hBitmap = NULL;
    pDataObject->GetData(&fc, &stg);
    HBITMAP hBitmap_ = stg.hBitmap;
    BITMAP bitmap_ = { 0 };
    GetObject(hBitmap_, sizeof(bitmap_), &bitmap_);
    ReleaseStgMedium(&stg);

    HDC screen = GetDC(NULL);
    int hSize = GetDeviceCaps(screen, HORZSIZE);
    int hRes = GetDeviceCaps(screen, HORZRES);
    ReleaseDC(NULL, screen);
    // 每个像素代表多少 mm（毫米)
    double mmsPerPixel = (double)hSize / hRes;

    // 我们将 RichEdit 中插入图片大小限制为最宽 150 个像素宽，高度无限制。
    // 这一步得到 reobject中 sz 成员的 cx 成员的最大值
    // reobject.sz.cx 是以 0.01mm 为单位的。
    int max_cx = 300 * mmsPerPixel * 100;
    double radio = (double)bitmap_.bmHeight / bitmap_.bmWidth;
    int bitmap_cx_ = bitmap_.bmWidth * mmsPerPixel * 100;
    SIZEL sz_ = { 0,0 };
    sz_.cx = bitmap_cx_ > max_cx ? max_cx : bitmap_cx_;
    sz_.cy = (LONG)(sz_.cx * radio);

    sc = CreateILockBytesOnHGlobal(NULL, TRUE, &lpLockBytes);
    if (lpLockBytes == NULL)
        return 0;

    sc = StgCreateDocfileOnILockBytes(lpLockBytes,
        STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_READWRITE, 0, &lpStorage);
    if (lpStorage == NULL)
        return 0;

    lpFormatEtc = &formatEtc;
    lpFormatEtc->cfFormat = cfFormat;
    lpFormatEtc->ptd = NULL;
    lpFormatEtc->dwAspect = DVASPECT_CONTENT;
    lpFormatEtc->lindex = -1;
    lpFormatEtc->tymed = TYMED_NULL;

    SendMessage(hwnd, EM_GETOLEINTERFACE, 0, (LPARAM)&lpRichEditOle);
    if (lpRichEditOle == NULL)
        return 0;

    lpRichEditOle->GetClientSite(&lpOleClientSite);
    if (lpOleClientSite == NULL)
        return 0;

    hr = OleCreateStaticFromData(pDataObject, IID_IUnknown,
        OLERENDER_DRAW, lpFormatEtc, lpOleClientSite, lpStorage, (void**)&lpOleObject);
    if (lpOleObject == NULL)
        return 0;

    lpUnknown = (LPUNKNOWN)lpOleObject;
    lpUnknown->QueryInterface(IID_IOleObject, (void**)&lpOleObject);
    lpUnknown->Release();
    OleSetContainedObject((LPUNKNOWN)lpOleObject, TRUE);

    ZeroMemory(&reobject, sizeof(REOBJECT));
    reobject.cbStruct = sizeof(REOBJECT);

    sc = lpOleObject->GetUserClassID(&clsid);

    // Update Edit Control Text
    OLECHAR* guidString;
    StringFromCLSID(clsid, &guidString);
    std::wstring text;
    int len = Edit_GetTextLength(hwnd);
    text.resize(len+1);
    Edit_GetText(hwndEdit, &text[0], len+1);
    if (text.back() == '\0') text.pop_back();
    text.append(L"\n, clsid is: ").append(guidString).append(L"\n");
    ::CoTaskMemFree(guidString);
    Edit_SetText(hwndEdit, text.c_str());

    reobject.clsid = clsid;
    reobject.cp = REO_CP_SELECTION;
    reobject.dvaspect = DVASPECT_CONTENT;
    reobject.dwFlags = REO_BELOWBASELINE;
    reobject.dwUser = 0;
    reobject.poleobj = lpOleObject;
    reobject.polesite = lpOleClientSite;
    reobject.pstg = lpStorage;

    //sizel.cx = sizel.cy = 0;
    //reobject.sizel = sizel;
    // 设置经过缩放的图片
    reobject.sizel = sz_;

    SendMessage(hwnd, EM_SETSEL, 0, -1);
    SendMessage(hwnd, EM_GETSEL, (WPARAM)&dwStart, (LPARAM)&dwEnd);
    SendMessage(hwnd, EM_SETSEL, dwEnd + 1, dwEnd + 1);

    lpRichEditOle->InsertObject(&reobject);

    lpLockBytes->Release();
    lpStorage->Release();
    pDataObject->Release();
    lpRichEditOle->Release();
    lpOleClientSite->Release();
    lpOleObject->Release();
    CoUninitialize();

    return 1;
}

void errhandler(LPCSTR func, HWND)
{
    MessageBoxA(NULL, func, "errhandler", MB_ICONERROR | MB_OK);
    exit(-1);
}

PBITMAPINFO CreateBitmapInfoStruct(HWND hwnd, HBITMAP hBmp)
{
    BITMAP bmp;
    PBITMAPINFO pbmi;
    WORD    cClrBits;

    // Retrieve the bitmap color format, width, and height.
    if (!GetObject(hBmp, sizeof(BITMAP), (LPSTR)&bmp))
        errhandler("GetObject", hwnd);

    // Convert the color format to a count of bits.
    cClrBits = (WORD)(bmp.bmPlanes * bmp.bmBitsPixel);
    if (cClrBits == 1)
        cClrBits = 1;
    else if (cClrBits <= 4)
        cClrBits = 4;
    else if (cClrBits <= 8)
        cClrBits = 8;
    else if (cClrBits <= 16)
        cClrBits = 16;
    else if (cClrBits <= 24)
        cClrBits = 24;
    else cClrBits = 32;

    // Allocate memory for the BITMAPINFO structure. (This structure
    // contains a BITMAPINFOHEADER structure and an array of RGBQUAD
    // data structures.)

    if (cClrBits < 24)
        pbmi = (PBITMAPINFO)LocalAlloc(LPTR,
            sizeof(BITMAPINFOHEADER) +
            sizeof(RGBQUAD) * (1 << cClrBits));

    // There is no RGBQUAD array for these formats: 24-bit-per-pixel or 32-bit-per-pixel

    else
        pbmi = (PBITMAPINFO)LocalAlloc(LPTR,
            sizeof(BITMAPINFOHEADER));

    // Initialize the fields in the BITMAPINFO structure.

    pbmi->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    pbmi->bmiHeader.biWidth = bmp.bmWidth;
    pbmi->bmiHeader.biHeight = bmp.bmHeight;
    pbmi->bmiHeader.biPlanes = bmp.bmPlanes;
    pbmi->bmiHeader.biBitCount = bmp.bmBitsPixel;
    if (cClrBits < 24)
        pbmi->bmiHeader.biClrUsed = (1 << cClrBits);

    // If the bitmap is not compressed, set the BI_RGB flag.
    pbmi->bmiHeader.biCompression = BI_RGB;

    // Compute the number of bytes in the array of color
    // indices and store the result in biSizeImage.
    // The width must be DWORD aligned unless the bitmap is RLE
    // compressed.
    pbmi->bmiHeader.biSizeImage = ((pbmi->bmiHeader.biWidth * cClrBits + 31) & ~31) / 8
        * pbmi->bmiHeader.biHeight;
    // Set biClrImportant to 0, indicating that all of the
    // device colors are important.
    pbmi->bmiHeader.biClrImportant = 0;
    return pbmi;
}

void CreateBMPFile(HWND hwnd, LPCTSTR pszFile, PBITMAPINFO pbi,
    HBITMAP hBMP, HDC hDC)
{
    HANDLE hf;                  // file handle
    BITMAPFILEHEADER hdr;       // bitmap file-header
    PBITMAPINFOHEADER pbih;     // bitmap info-header
    LPBYTE lpBits;              // memory pointer
    DWORD dwTotal;              // total count of bytes
    DWORD cb;                   // incremental count of bytes
    BYTE* hp;                   // byte pointer
    DWORD dwTmp;

    pbih = (PBITMAPINFOHEADER)pbi;
    lpBits = (LPBYTE)GlobalAlloc(GMEM_FIXED, pbih->biSizeImage);

    if (!lpBits)
        errhandler("GlobalAlloc", hwnd);

    // Retrieve the color table (RGBQUAD array) and the bits
    // (array of palette indices) from the DIB.
    if (!GetDIBits(hDC, hBMP, 0, (WORD)pbih->biHeight, lpBits, pbi,
        DIB_RGB_COLORS))
    {
        errhandler("GetDIBits", hwnd);
    }

    // Create the .BMP file.
    hf = CreateFile(pszFile,
        GENERIC_READ | GENERIC_WRITE,
        (DWORD)0,
        NULL,
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        (HANDLE)NULL);
    if (hf == INVALID_HANDLE_VALUE)
        errhandler("CreateFile", hwnd);
    hdr.bfType = 0x4d42;        // 0x42 = "B" 0x4d = "M"
    // Compute the size of the entire file.
    hdr.bfSize = (DWORD)(sizeof(BITMAPFILEHEADER) +
        pbih->biSize + pbih->biClrUsed
        * sizeof(RGBQUAD) + pbih->biSizeImage);
    hdr.bfReserved1 = 0;
    hdr.bfReserved2 = 0;

    // Compute the offset to the array of color indices.
    hdr.bfOffBits = (DWORD)sizeof(BITMAPFILEHEADER) +
        pbih->biSize + pbih->biClrUsed
        * sizeof(RGBQUAD);

    // Copy the BITMAPFILEHEADER into the .BMP file.
    if (!WriteFile(hf, (LPVOID)&hdr, sizeof(BITMAPFILEHEADER),
        (LPDWORD)&dwTmp, NULL))
    {
        errhandler("WriteFile", hwnd);
    }

    // Copy the BITMAPINFOHEADER and RGBQUAD array into the file.
    if (!WriteFile(hf, (LPVOID)pbih, sizeof(BITMAPINFOHEADER)
        + pbih->biClrUsed * sizeof(RGBQUAD),
        (LPDWORD)&dwTmp, (NULL)))
        errhandler("WriteFile", hwnd);

    // Copy the array of color indices into the .BMP file.
    dwTotal = cb = pbih->biSizeImage;
    hp = lpBits;
    if (!WriteFile(hf, (LPSTR)hp, (int)cb, (LPDWORD)&dwTmp, NULL))
        errhandler("WriteFile", hwnd);

    // Close the .BMP file.
    if (!CloseHandle(hf))
        errhandler("CloseHandle", hwnd);

    // Free memory.
    GlobalFree((HGLOBAL)lpBits);
}

bool MBCSToUnicode(const char* input, std::wstring& output, int code_page)
{
    output.clear();
    int length = ::MultiByteToWideChar(code_page, 0, input, -1, NULL, 0);
    if (length <= 0)
        return false;
    output.resize(length - 1);
    if (output.empty())
        return true;
    ::MultiByteToWideChar(code_page,
        0,
        input,
        -1,
        &output[0],
        static_cast<int>(output.size()));
    return true;
}

BOOL RE_GetBitmap(HWND hwndRichEdit, HWND hwndEdit)
{
    LPRICHEDITOLE lpRichEditOle;
    SendMessage(hwndRichEdit, EM_GETOLEINTERFACE, 0, (LPARAM)&lpRichEditOle);
    if (!lpRichEditOle) return FALSE;
    long count = lpRichEditOle->GetObjectCount();
    std::wstring log(L"There are ");
    log.append(std::to_wstring(count)).append(L" OLEOBJECT in richedit");
    Edit_SetText(hwndEdit, log.c_str());

    for (int i = 0; i < count; ++i)
    {
        REOBJECT reo;
        ZeroMemory(&reo, sizeof(reo));
        reo.cbStruct = sizeof reo;
        SCODE sc = lpRichEditOle->GetObjectW(i, &reo, REO_GETOBJ_ALL_INTERFACES);
        if (sc != S_OK)
        {
            continue;
        }

        // 获取IDataObject信息
        IDataObject* pDataObject = NULL;
        sc = reo.poleobj->QueryInterface(IID_IDataObject, (void**)&pDataObject);

        if (sc != S_OK) break;

        STGMEDIUM stg;
        FORMATETC fc;
        fc.cfFormat = CF_BITMAP;        // Clipboard format = CF_BITMAP
        fc.ptd = NULL;                  // Target Device = Screen
        fc.dwAspect = DVASPECT_CONTENT; // Level of detail = Full content
        fc.lindex = -1;                 // Index = Not applicaple
        fc.tymed = TYMED_GDI;           // 对应CF_BITMAP

        HRESULT hr = pDataObject->GetData(&fc, &stg);
        if (hr != S_OK || stg.hBitmap == NULL)
        {
            break; // 已经找到选中的图片对象，获取信息失败，直接退出
        }

        //BOOL bRet = SaveBmpDataToFile(stg.hBitmap, csFilePath);
        //if (bRet)
        //{
        //    bSavePicSuccess = TRUE;
        //}

        auto now = std::chrono::system_clock::now().time_since_epoch();
        auto seconds = std::chrono::duration_cast<std::chrono::seconds>(now).count();

        std::string home = getenv("USERPROFILE");
        home.append("\\Desktop\\");
        std::wstring name;
        MBCSToUnicode(home.c_str(), name, CP_ACP);
        name.append(L"\\").append(std::to_wstring(seconds)).append(L"-")
            .append(std::to_wstring(i)).append(L".bmp");

        PBITMAPINFO pbmi = CreateBitmapInfoStruct(hwndRichEdit, stg.hBitmap);
        HDC hdc = GetDC(NULL);
        CreateBMPFile(hwndRichEdit, name.c_str(), pbmi, stg.hBitmap, hdc);
        ReleaseDC(NULL, hdc);
    }

    return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE: Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    static HWND hwndRichEdit, hwndEdit, hwndBtnInsertFile, hwndBtnInsertImage;
    static HWND hwndBtnPasteWin32ClipBoard, hwndBtnPasteOleClipBoard, hwndBtnGetConent;
    switch (message)
    {
    case WM_CREATE:
    {
        LoadLibrary(TEXT("Riched20.dll"));
        hwndRichEdit = CreateWindowEx(0, RICHEDIT_CLASS, TEXT("Type here"),
            ES_MULTILINE | WS_VISIBLE | WS_CHILD | WS_BORDER | WS_TABSTOP | WS_VSCROLL | WS_HSCROLL,
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            hWnd, (HMENU)IDR_RE, hInst, NULL);

        hwndEdit = CreateWindowEx(0, L"Edit", L"This edit control is used to logging some messages",
            WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL,
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            hWnd, (HMENU)IDE_EDIT, hInst, NULL);

        hwndBtnInsertFile = CreateWindowEx(0, TEXT("BUTTON"), TEXT("Insert File"),
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            hWnd, (HMENU)IDB_INSERT_FILE, hInst, NULL);
        hwndBtnInsertImage = CreateWindowEx(0, TEXT("BUTTON"), TEXT("Insert Image"),
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            hWnd, (HMENU)IDB_INSERT_IMAGE, hInst, NULL);
        hwndBtnPasteWin32ClipBoard = CreateWindowEx(0, TEXT("BUTTON"), TEXT("Paste From Win32 ClipBoard"),
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON, CW_USEDEFAULT, CW_USEDEFAULT,
            CW_USEDEFAULT, CW_USEDEFAULT, hWnd, (HMENU)IDB_PASTE_WIN32, hInst, NULL);
        hwndBtnPasteOleClipBoard = CreateWindowEx(0, TEXT("BUTTON"), TEXT("Paste From OLE ClipBoard"),
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON, CW_USEDEFAULT, CW_USEDEFAULT,
            CW_USEDEFAULT, CW_USEDEFAULT, hWnd, (HMENU)IDB_PASTE_OLE, hInst, NULL);
        hwndBtnGetConent = CreateWindowEx(0, TEXT("BUTTON"), TEXT("Get RichEdit Content"),
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON, CW_USEDEFAULT, CW_USEDEFAULT,
            CW_USEDEFAULT, CW_USEDEFAULT, hWnd, (HMENU)IDB_GET_CONTENT, hInst, NULL);
        break;
    }
    case WM_SIZE:
    {
        WORD cx = LOWORD(lParam), cy = HIWORD(lParam);
        int btnHeight = 50, vmargin = 10, hmargin = 10, btnWidth = 200, btnWidthWide = 300;
        SetWindowPos(hwndRichEdit, HWND_TOP, 0, btnHeight + vmargin, cx * 2 / 3, cy - btnHeight - vmargin, SWP_SHOWWINDOW);
        SetWindowPos(hwndEdit, HWND_TOP, cx * 2 / 3 + hmargin, +btnHeight + vmargin, cx / 3 - hmargin, cy - btnHeight - vmargin, SWP_SHOWWINDOW);
        SetWindowPos(hwndBtnInsertFile, HWND_TOP, 0, 0, btnWidth, btnHeight, SWP_SHOWWINDOW);
        SetWindowPos(hwndBtnInsertImage, HWND_TOP, btnWidth + hmargin, 0, btnWidth, btnHeight, SWP_SHOWWINDOW);
        SetWindowPos(hwndBtnPasteWin32ClipBoard, HWND_TOP, (btnWidth + hmargin) * 2, 0, btnWidthWide, btnHeight, SWP_SHOWWINDOW);
        SetWindowPos(hwndBtnPasteOleClipBoard, HWND_TOP, btnWidth * 2 + hmargin * 3 + btnWidthWide, 0, btnWidthWide, btnHeight, SWP_SHOWWINDOW);
        SetWindowPos(hwndBtnGetConent, HWND_TOP, btnWidth * 2 + hmargin * 4 + btnWidthWide * 2, 0, btnWidthWide, btnHeight, SWP_SHOWWINDOW);

        break;
    }
    case WM_COMMAND:
    {
        int wmId = LOWORD(wParam);
        // Parse the menu selections:
        switch (wmId)
        {
        case IDM_ABOUT:
            DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
            break;
        case IDM_EXIT:
            DestroyWindow(hWnd);
            break;
        case IDB_INSERT_FILE:
        {
            std::wstring file = SelectFile(hWnd, L"Source\0*.C;*.CXX\0All\0*.*\0");
            if (!file.empty())
            {
                InsertObject(hwndRichEdit, file.c_str());
            }
            break;
        }
        case IDB_INSERT_IMAGE:
        {
            std::wstring image = SelectFile(hWnd, L"Image\0*.png;*.bmp;*.jpg;*.jpeg\0");
            {
                if (!image.empty())
                {
                    InsertImage(hwndRichEdit, image.c_str());
                }
            }
            break;
        }
            // Past From Win32 ClipBoard
        case IDB_PASTE_WIN32:
        {
            //if (IsClipboardFormatAvailable(CF_BITMAP))
            //{
            //    OpenClipboard(NULL);
            //    HGLOBAL hGlobal = GetClipboardData(CF_BITMAP);
            //    if (hGlobal && hGlobal != INVALID_HANDLE_VALUE)
            //    {
            //        void* bitmap = (void*)hGlobal;
            //        if (bitmap)
            //        {
            //            BITMAPINFOHEADER* info = reinterpret_cast<BITMAPINFOHEADER*>(bitmap);
            //            BITMAPFILEHEADER fileHeader = { 0 };
            //            fileHeader.bfType = 0x4D42;
            //            fileHeader.bfOffBits = 54;
            //            //fileHeader.bfSize = (((info->biWidth * info->biBitCount + 31) & ~31) / 8
            //            //    * info->biHeight) + fileHeader.bfOffBits;

            //            std::ofstream file("C:/Users/quanzhuo/Desktop/Test.bmp", std::ios::out | std::ios::binary);
            //            if (file)
            //            {
            //                file.write(reinterpret_cast<char*>(&fileHeader), sizeof(BITMAPFILEHEADER));
            //                file.write(reinterpret_cast<char*>(info), sizeof(BITMAPINFOHEADER));
            //                file.write(reinterpret_cast<char*>(++info), info->biSizeImage);
            //            }
            //        }
            //        //GlobalUnlock(hGlobal);
            //    }
            //    CloseClipboard();
            //}
            break;
        }
        // Paste from ole clipboard
        case IDB_PASTE_OLE:
        {
            RE_InsertBitmapFromOleClipBoard(hwndRichEdit, hwndEdit);
            break;
        }
        case IDB_GET_CONTENT:
        {
            RE_GetBitmap(hwndRichEdit, hwndEdit);
            break;
        }
        // Copy
        case IDM_COPY:
            MessageBox(hWnd, L"Copy Selected", L"Info", MB_OK);
            break;
            // Paste
        case IDM_PASTE:
        {
            break;
        }
        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
        }
    }
    break;
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        // TODO: Add any drawing code that uses hdc here...
        EndPaint(hWnd, &ps);
    }
    break;
    case WM_CONTEXTMENU:
    {
        if ((HWND)wParam == hwndRichEdit)
        {
            BOOL bCanPaste = SendMessage(hwndRichEdit, EM_CANPASTE, 0, 0);
            CHARRANGE cr;
            SendMessage(hwndRichEdit, EM_EXGETSEL, 0, (LPARAM)&cr);
            BOOL bCanCopy = cr.cpMin == cr.cpMax ? FALSE : TRUE;

            int xPos = GET_X_LPARAM(lParam), yPos = GET_Y_LPARAM(lParam);
            HMENU hMenu = CreatePopupMenu();
            AppendMenuW(hMenu, (bCanCopy ? MF_ENABLED : MF_DISABLED) | MF_STRING, 11, L"复制");
            AppendMenuW(hMenu, (bCanPaste ? MF_ENABLED : MF_DISABLED) | MF_STRING, 12, L"粘贴");
            TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_TOPALIGN, xPos, yPos, hWnd, NULL);
        }
        break;
    }
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
