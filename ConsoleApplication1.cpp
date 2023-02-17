// ConsoleApplication1.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <cstdio>
#include <windows.h>
#include <ole2.h>
#include <uiautomation.h>
#include <strsafe.h>
#include <chrono>
#include <wrl.h>
#include <comdef.h>
#include <wil\resource.h>
#include <wil\com.h>
#include <wil\result.h>
#include <wil\winrt.h>
#include <winuser.h>
#include<Windows.h>
#include <string>
#include<vector>
#include <fstream>
#include<comutil.h>
#include<stdio.h>
#include <gdiplus.h>
#pragma comment(lib,"gdiplus.lib")
#include <time.h>
//#include "winrt/Windows.Foundation.Collections.h"
//#include "winrt/windows.management.deployment.h"
//#include "winrt/Microsoft.UI.UIAutomation.h"

using namespace std;

IUIAutomation* _automation;
double WIDTH, HEIGHT, WINDOW_LEFT, WINDOW_TOP;
wstring FOLDER_PATH;
wstring curTime;

void errhandler(string s, HWND hwnd)
{}

PBITMAPINFO CreateBitmapInfoStruct(HWND hwnd, HBITMAP hBmp)
{
    BITMAP bmp;
    PBITMAPINFO pbmi;
    WORD cClrBits;

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
    else
        cClrBits = 32;

    // Allocate memory for the BITMAPINFO structure. (This structure
    // contains a BITMAPINFOHEADER structure and an array of RGBQUAD
    // data structures.)

    if (cClrBits < 24)
        pbmi = (PBITMAPINFO)LocalAlloc(LPTR, sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD) * (1i64 << cClrBits));

    // There is no RGBQUAD array for these formats: 24-bit-per-pixel or 32-bit-per-pixel

    else
        pbmi = (PBITMAPINFO)LocalAlloc(LPTR, sizeof(BITMAPINFOHEADER));

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
    pbmi->bmiHeader.biSizeImage = ((pbmi->bmiHeader.biWidth * cClrBits + 31) & ~31) / 8 * pbmi->bmiHeader.biHeight;
    // Set biClrImportant to 0, indicating that all of the
    // device colors are important.
    pbmi->bmiHeader.biClrImportant = 0;
    return pbmi;
}

// Save the Bitmap to Disk
void CreateBMPFile(HWND hwnd, LPTSTR pszFile, PBITMAPINFO pbi, HBITMAP hBMP, HDC hDC)
{
    HANDLE hf;              // file handle
    BITMAPFILEHEADER hdr;   // bitmap file-header
    PBITMAPINFOHEADER pbih; // bitmap info-header
    LPBYTE lpBits;          // memory pointer
    DWORD dwTotal;          // total count of bytes
    DWORD cb;               // incremental count of bytes
    BYTE* hp;               // byte pointer
    DWORD dwTmp;

    pbih = (PBITMAPINFOHEADER)pbi;
    lpBits = (LPBYTE)GlobalAlloc(GMEM_FIXED, pbih->biSizeImage);

    if (!lpBits)
        errhandler("GlobalAlloc", hwnd);

    // Retrieve the color table (RGBQUAD array) and the bits
    // (array of palette indices) from the DIB.
    if (!GetDIBits(hDC, hBMP, 0, (WORD)pbih->biHeight, lpBits, pbi, DIB_RGB_COLORS))
    {
        errhandler("GetDIBits", hwnd);
    }

    // Create the .BMP file.
    hf = CreateFile(pszFile, GENERIC_READ | GENERIC_WRITE, (DWORD)0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, (HANDLE)NULL);
    if (hf == INVALID_HANDLE_VALUE)
        errhandler("CreateFile", hwnd);
    hdr.bfType = 0x4d42; // 0x42 = "B" 0x4d = "M"
    // Compute the size of the entire file.
    hdr.bfSize = (DWORD)(sizeof(BITMAPFILEHEADER) + pbih->biSize + pbih->biClrUsed * sizeof(RGBQUAD) + pbih->biSizeImage);
    hdr.bfReserved1 = 0;
    hdr.bfReserved2 = 0;

    // Compute the offset to the array of color indices.
    hdr.bfOffBits = (DWORD)sizeof(BITMAPFILEHEADER) + pbih->biSize + pbih->biClrUsed * sizeof(RGBQUAD);

    // Copy the BITMAPFILEHEADER into the .BMP file.
    if (!WriteFile(hf, (LPVOID)&hdr, sizeof(BITMAPFILEHEADER), (LPDWORD)&dwTmp, NULL))
    {
        errhandler("WriteFile", hwnd);
    }

    // Copy the BITMAPINFOHEADER and RGBQUAD array into the file.
    if (!WriteFile(hf, (LPVOID)pbih, sizeof(BITMAPINFOHEADER) + pbih->biClrUsed * sizeof(RGBQUAD), (LPDWORD)&dwTmp, (NULL)))
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

// Create Encoder of Required Image type
int GetEncoderClsid(const WCHAR* format, CLSID* pClsid)
{
    using namespace Gdiplus;

    UINT num = 0;  // number of image encoders
    UINT size = 0; // size of the image encoder array in bytes

    ImageCodecInfo* pImageCodecInfo = NULL;

    GetImageEncodersSize(&num, &size);
    if (size == 0)
        return -1; // Failure

    pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
    if (pImageCodecInfo == NULL)
        return -1; // Failure

    GetImageEncoders(num, size, pImageCodecInfo);

    for (UINT j = 0; j < num; ++j)
    {
        if (wcscmp(pImageCodecInfo[j].MimeType, format) == 0)
        {
            *pClsid = pImageCodecInfo[j].Clsid;
            free(pImageCodecInfo);
            return j; // Success
        }
    }

    free(pImageCodecInfo);
    return -1; // Failure
}

// Convert the Screenshot Bitmap to JPG to reduce size reducing latency of the API Call
void convertToJPG(wstring filepath)
{
    using namespace Gdiplus;

    // Initialize GDI+.
    GdiplusStartupInput gdiplusStartupInput;
    ULONG_PTR gdiplusToken;
    GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);
    wstring wfilename = wstring(filepath.begin(), filepath.end());

    CLSID encoderClsid;
    Status stat;
    Image* image = Image::FromFile(wfilename.c_str());
    wstring dest_path = filepath.substr(0, filepath.size() - 10) + L"\\" + curTime + L".jpg";

    // Get the CLSID of the PNG encoder.
    GetEncoderClsid(L"image/jpeg", &encoderClsid);

    stat = image->Save(wstring(dest_path.begin(), dest_path.end()).c_str(), &encoderClsid, NULL);

    // Clean up
    delete image;
    GdiplusShutdown(gdiplusToken);
}

// Function to Capture a Screenshot
void captureScreenShot(HWND hwnd)
{
    double x1, y1, x2, y2, w, h;

    RECT rect = RECT();
    GetClientRect(hwnd, &rect);
    ClientToScreen(hwnd, reinterpret_cast<POINT*>(&rect));
    ClientToScreen(hwnd, reinterpret_cast<POINT*>(&rect) + 1);

    x1 = rect.left;
    x2 = rect.right;
    y1 = rect.top;
    y2 = rect.bottom;

    WINDOW_LEFT = x1;
    WINDOW_TOP = y1;

    w = x2 - x1;
    h = y2 - y1;

    // Copy App Window to bitmap
    // Using GetDesktopWindow() to get the Entire Screen to avoid black screen in unsupported windows.
    HDC hScreen = GetDC(GetDesktopWindow());
    HDC hDC = CreateCompatibleDC(hScreen);
    HBITMAP hBitmap = CreateCompatibleBitmap(hScreen, (int)w, (int)h);
    HGDIOBJ old_obj = SelectObject(hDC, hBitmap);
    BOOL bRet = BitBlt(hDC, 0, 0, (int)w, (int)h, hScreen, (int)x1, (int)y1, SRCCOPY);

    wchar_t dir[1024];
    GetCurrentDirectory(1024, dir);
    FOLDER_PATH.assign(dir);
    FOLDER_PATH.append(L"\\UIA_Training_Data\\");
    CreateDirectory(FOLDER_PATH.c_str(), nullptr);
    curTime = std::to_wstring(GetTickCount64());
    wstring filename = FOLDER_PATH + L"Bitmap.bmp";
    std::wstring stemp = std::wstring(filename.begin(), filename.end());
    LPCTSTR sw = stemp.c_str();

    PBITMAPINFO pbinfo = CreateBitmapInfoStruct(hwnd, hBitmap);
    CreateBMPFile(hwnd, (LPTSTR)sw, pbinfo, hBitmap, hDC);

    // Convert to JPG
    convertToJPG(filename);

    DeleteFile(filename.c_str());

    // clean up
    SelectObject(hDC, old_obj);
    DeleteDC(hDC);
    ReleaseDC(NULL, hScreen);
    DeleteObject(hBitmap);
}


HWND TryFindWindowNow(DWORD m_processId)
{
    HWND hwnd = ::GetDesktopWindow();
    HWND hwndChild = ::GetWindow(hwnd, GW_CHILD);
    for (; hwndChild != nullptr; hwndChild = ::GetWindow(hwndChild, GW_HWNDNEXT))
    {
        // Check for processId match
        DWORD childProcessId = 0;
        ::GetWindowThreadProcessId(hwndChild, &childProcessId);
        if (childProcessId != m_processId)
        {
            continue;
        }

        // Ignore invisible worker windows
        if (!::IsWindowVisible(hwndChild))
        {
            continue;
        }

        // Ignore console windows
        WCHAR testClassName[64];
        ::GetClassName(hwndChild, testClassName, ARRAYSIZE(testClassName));
        if (::lstrcmpi(testClassName, L"ConsoleWindowClass") == 0)
        {
            continue;
        }

        // Matches all criteria - return it!
        return hwndChild;
    }

    return nullptr;
}

void GenerateUIATree()
{
    wil::com_ptr<IUIAutomationElement> element;
    //DWORD processId = GetForegroundWindow();
    HWND handle = GetForegroundWindow();

    captureScreenShot(handle);

    HRESULT hr = _automation->ElementFromHandle(handle, &element);
    /*if (FAILED(hr))
    {
        wprintf(L"Failed to ElementFromHandle, HR: 0x%08x\n\n", hr);
    }*/
    if (element != NULL)
    {
        wil::com_ptr<IUIAutomationCondition> conditionList;
        wil::unique_variant control_type;
        control_type.vt = VT_I4;
        control_type.lVal = UIA_ListControlTypeId;
        _automation->CreatePropertyCondition(UIA_ControlTypePropertyId, control_type, &conditionList);

        wil::com_ptr<IUIAutomationCondition> conditionButton;
        wil::unique_variant control_type1;
        control_type1.vt = VT_I4;
        control_type1.lVal = UIA_ButtonControlTypeId;
        _automation->CreatePropertyCondition(UIA_ControlTypePropertyId, control_type1, &conditionButton);

        wil::com_ptr<IUIAutomationCondition> conditionImage;
        wil::unique_variant control_type2;
        control_type2.vt = VT_I4;
        control_type2.lVal = UIA_ImageControlTypeId;
        _automation->CreatePropertyCondition(UIA_ControlTypePropertyId, control_type2, &conditionImage);

        wil::com_ptr<IUIAutomationCondition> condition;
        _automation->CreateOrCondition(conditionList.get(), conditionButton.get(), &condition);

        wil::com_ptr<IUIAutomationCondition> finalcondition;
        _automation->CreateOrCondition(conditionImage.get(), condition.get(), &finalcondition);

        wil::com_ptr<IUIAutomationCacheRequest> cacheRequest;
        _automation->CreateCacheRequest(&cacheRequest);
        cacheRequest->AddProperty(UIA_NamePropertyId);
        cacheRequest->AddProperty(UIA_BoundingRectanglePropertyId);
        cacheRequest->AddProperty(UIA_ControlTypePropertyId);

        wil::com_ptr<IUIAutomationElementArray> foundElements;

        element->FindAllBuildCache(TreeScope_Descendants, finalcondition.get(), cacheRequest.get(), &foundElements);

        int length = 0;
        foundElements->get_Length(&length);
        string temp = "";
        vector<string> list;
        for (int i = 0; i < length; ++i)
        {
            temp = "";
            wil::com_ptr<IUIAutomationElement> element;
            foundElements->GetElement(i, &element);
            RECT rectangle;
            BSTR name;
            CONTROLTYPEID id;
            element->get_CachedBoundingRectangle(&rectangle);
            element->get_CachedName(&name);
            element->get_CachedControlType(&id);

            std::wstring ws(name, SysStringLen(name));
            string str = string(ws.begin(), ws.end());


            temp = to_string(id) + " " + str + " " + to_string(rectangle.left) + " " + to_string(rectangle.top) + " " + to_string(rectangle.right) + " " + to_string(rectangle.bottom);

            list.push_back(temp);
        }


        wstring uiaTreePath = FOLDER_PATH + curTime + L".txt";
        ofstream fw(uiaTreePath, std::ofstream::out);

        if (fw.is_open())
        {
            //store array contents to text file
            for (int i = 0; i < list.size(); i++) {
                fw << list[i] << "\n";
            }
            fw.close();

        }
        cout << "Processing finished. Files dumped at location:";
        wcout << uiaTreePath << endl<<endl;
    }
    else
    {
        cout << "No element found" << endl;
    }

}

int ProcessInput(int input)
{
    switch (input)
    {
    case 1:
        cout << "Sleep for 5 seconds. Please change the foreground window for which you want the UIA tree" << endl;
        Sleep(5000);
        GenerateUIATree();
        return 1;
    case 2:
        cout << "Exiting" << endl;
        return 0;
    default:
        cout<<"Incorrect Input: Please try again"<<endl;
        return 1;
    }
}

int __cdecl main()
{
    // Initialize COM before using UI Automation
    HRESULT hr = CoInitialize(NULL);
    /*if (FAILED(hr))
    {
        wprintf(L"CoInitialize failed, HR:0x%08x\n", hr);
    }*/

    hr = CoCreateInstance(__uuidof(CUIAutomation8), NULL,
        CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_automation));

    // Save thsi process handle so that we can make this foreground window
    HWND ownHandle = GetForegroundWindow();
    int ret = 0,input =0;
    do
    {
        ::SetForegroundWindow(ownHandle);
        cout << "Please enter out of these options:" << endl;
        cout << "1. Generate UIA tree and capture screenshot"<<endl;
        cout << "2. Exit the program"<<endl;
        cin >> input;
        ret = ProcessInput(input);
    } while (ret != 0);
 
    return ret;
}
