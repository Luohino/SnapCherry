#include <windows.h>
#include <wincodec.h>
#include <shlobj.h>
#include <initguid.h>
#include <knownfolders.h>
#include <shlwapi.h>
#include <stdio.h>
#include <time.h>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "windowscodecs.lib")
#pragma comment(lib, "shlwapi.lib")

// --- Constants & Defines ---
#define HOTKEY_ID 1
#define TOOLBAR_HEIGHT 60
#define TOOLBAR_WIDTH 450
#define COLOR_BUTTON_SIZE 30
#define TOOL_BUTTON_SIZE 40

typedef enum { TOOL_PEN, TOOL_ERASER } ToolType;

// --- Global State ---
HINSTANCE g_hInstance;
HWND g_hMainWnd = NULL;
HWND g_hOverlayWnd = NULL;
HWND g_hToolbarWnd = NULL;

HBITMAP g_hScreenBmp = NULL;
HBITMAP g_hDrawingBmp = NULL;
HDC g_hScreenDC = NULL;
HDC g_hDrawingDC = NULL;

RECT g_Selection = {0};
BOOL g_IsSelecting = FALSE;
BOOL g_IsCaptured = FALSE;
BOOL g_IsDrawing = FALSE;

COLORREF g_CurrentColor = RGB(255, 0, 0); // Red
int g_PenSize = 4;
ToolType g_CurrentTool = TOOL_PEN;

POINT g_LastMouse = {0};

// --- Forward Declarations ---
LRESULT CALLBACK MainWndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK OverlayWndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK ToolbarWndProc(HWND, UINT, WPARAM, LPARAM);
void RegisterAutostart();
void CaptureScreen();
void SaveScreenshot();
HRESULT SaveBitmapToPNG(HBITMAP hBmp, const WCHAR* filename);

// --- Hotkey & App Logic ---
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    HANDLE hMutex = CreateMutex(NULL, TRUE, "SnapCherryUniqueMutex");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        // If already running, don't show an error, just exit or show a small tray tip
        return 0; 
    }

    g_hInstance = hInstance;
    CoInitialize(NULL);

    RegisterAutostart();

    WNDCLASS wc = {0};
    wc.lpfnWndProc = MainWndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = "SnapCherryMain";
    RegisterClass(&wc);

    g_hMainWnd = CreateWindowEx(0, wc.lpszClassName, "SnapCherry", 0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);

    if (!RegisterHotKey(g_hMainWnd, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, 'S')) {
        // If CTRL+SHIFT+S fails, try falling back to ALT+SHIFT+S
        if (!RegisterHotKey(g_hMainWnd, HOTKEY_ID, MOD_ALT | MOD_SHIFT, 'S')) {
            MessageBox(NULL, "Failed to register Hotkeys (CTRL+SHIFT+S and ALT+SHIFT+S).\n\nPlease make sure another screenshot tool isn't running.", "SnapCherry Error", MB_ICONERROR);
            return 1;
        } else {
            // Notify user of fallback
            MessageBox(NULL, "CTRL+SHIFT+S is already in use.\n\nSnapCherry will use ALT+SHIFT+S instead.", "SnapCherry Information", MB_ICONINFORMATION);
        }
    }

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    UnregisterHotKey(g_hMainWnd, HOTKEY_ID);
    CoUninitialize();
    return 0;
}

void RegisterAutostart() {
    HKEY hKey;
    char path[MAX_PATH];
    GetModuleFileName(NULL, path, MAX_PATH);
    if (RegOpenKeyEx(HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS) {
        RegSetValueEx(hKey, "SnapCherry", 0, REG_SZ, (BYTE*)path, strlen(path) + 1);
        RegCloseKey(hKey);
    }
}

// --- Overlay Window ---
void ShowOverlay() {
    if (g_hOverlayWnd) return;

    // Capture screen first
    CaptureScreen();

    WNDCLASS wc = {0};
    wc.lpfnWndProc = OverlayWndProc;
    wc.hInstance = g_hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_CROSS);
    wc.lpszClassName = "SnapCherryOverlay";
    RegisterClass(&wc);

    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);

    g_hOverlayWnd = CreateWindowEx(WS_EX_TOPMOST | WS_EX_TOOLWINDOW, wc.lpszClassName, "", WS_POPUP | WS_VISIBLE, 0, 0, screenW, screenH, NULL, NULL, g_hInstance, NULL);
    SetForegroundWindow(g_hOverlayWnd);
    SetFocus(g_hOverlayWnd);
}

void ShowToolbar() {
    if (g_hToolbarWnd) return;

    WNDCLASS wc = {0};
    wc.lpfnWndProc = ToolbarWndProc;
    wc.hInstance = g_hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.lpszClassName = "SnapCherryToolbar";
    RegisterClass(&wc);

    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);
    
    int x = (screenW - TOOLBAR_WIDTH) / 2;
    int y = screenH - TOOLBAR_HEIGHT - 40;

    g_hToolbarWnd = CreateWindowEx(WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_TOOLWINDOW, wc.lpszClassName, "", WS_POPUP | WS_VISIBLE, x, y, TOOLBAR_WIDTH, TOOLBAR_HEIGHT, g_hOverlayWnd, NULL, g_hInstance, NULL);
    SetLayeredWindowAttributes(g_hToolbarWnd, 0, 255, LWA_ALPHA);

    HRGN hRgn = CreateRoundRectRgn(0, 0, TOOLBAR_WIDTH, TOOLBAR_HEIGHT, 20, 20);
    SetWindowRgn(g_hToolbarWnd, hRgn, TRUE);
}

// --- Screen Capture ---
void CaptureScreen() {
    int w = GetSystemMetrics(SM_CXSCREEN);
    int h = GetSystemMetrics(SM_CYSCREEN);

    HDC hdcScreen = GetDC(NULL);
    g_hScreenDC = CreateCompatibleDC(hdcScreen);
    g_hScreenBmp = CreateCompatibleBitmap(hdcScreen, w, h);
    SelectObject(g_hScreenDC, g_hScreenBmp);
    BitBlt(g_hScreenDC, 0, 0, w, h, hdcScreen, 0, 0, SRCCOPY);

    g_hDrawingDC = CreateCompatibleDC(hdcScreen);
    g_hDrawingBmp = CreateCompatibleBitmap(hdcScreen, w, h);
    SelectObject(g_hDrawingDC, g_hDrawingBmp);
    
    // Fill drawing DC with transparency (or just copy screen)
    BitBlt(g_hDrawingDC, 0, 0, w, h, g_hScreenDC, 0, 0, SRCCOPY);

    ReleaseDC(NULL, hdcScreen);
}

// --- Window Procedures ---

LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    if (msg == WM_HOTKEY && wp == HOTKEY_ID) {
        ShowOverlay();
        return 0;
    }
    if (msg == WM_DESTROY) {
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProc(hwnd, msg, wp, lp);
}

LRESULT CALLBACK OverlayWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    static POINT ptStart;

    switch (msg) {
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);
            int w = GetSystemMetrics(SM_CXSCREEN);
            int h = GetSystemMetrics(SM_CYSCREEN);

            // Draw captured screen
            BitBlt(hdc, 0, 0, w, h, g_hScreenDC, 0, 0, SRCCOPY);

            if (!g_IsCaptured) {
                // Dimming overlay
                HBRUSH hDimBrush = CreateSolidBrush(RGB(0, 0, 0));
                HDC hMemDC = CreateCompatibleDC(hdc);
                HBITMAP hMemBmp = CreateCompatibleBitmap(hdc, w, h);
                SelectObject(hMemDC, hMemBmp);
                BitBlt(hMemDC, 0, 0, w, h, hdc, 0, 0, SRCCOPY);
                
                RECT r = {0, 0, w, h};
                // Use a simpler dimming: AlphaBlend would be better but let's use a pattern or just fill
                // For simplicity, we can use a transparent window, but here we draw manually
                FillRect(hMemDC, &r, hDimBrush);
                
                BLENDFUNCTION bf = {AC_SRC_OVER, 0, 128, 0};
                AlphaBlend(hdc, 0, 0, w, h, hMemDC, 0, 0, w, h, bf);
                
                DeleteObject(hDimBrush);
                DeleteObject(hMemBmp);
                DeleteDC(hMemDC);

                if (g_IsSelecting) {
                    RECT rcSel = g_Selection;
                    if (rcSel.left > rcSel.right) { int tmp = rcSel.left; rcSel.left = rcSel.right; rcSel.right = tmp; }
                    if (rcSel.top > rcSel.bottom) { int tmp = rcSel.top; rcSel.top = rcSel.bottom; rcSel.bottom = tmp; }
                    
                    // Punch a hole in the dimming
                    BitBlt(hdc, rcSel.left, rcSel.top, rcSel.right - rcSel.left, rcSel.bottom - rcSel.top, g_hScreenDC, rcSel.left, rcSel.top, SRCCOPY);
                    
                    // Draw border
                    HBRUSH hBorder = CreateSolidBrush(RGB(255, 255, 255));
                    FrameRect(hdc, &rcSel, hBorder);
                    DeleteObject(hBorder);
                }
            } else {
                // Drawing mode: Draw the edited bitmap
                BitBlt(hdc, 0, 0, w, h, g_hDrawingDC, 0, 0, SRCCOPY);
            }

            EndPaint(hwnd, &ps);
            return 0;
        }

        case WM_LBUTTONDOWN:
            if (!g_IsCaptured) {
                g_IsSelecting = TRUE;
                ptStart.x = LOWORD(lp);
                ptStart.y = HIWORD(lp);
                g_Selection.left = ptStart.x;
                g_Selection.top = ptStart.y;
            } else {
                g_IsDrawing = TRUE;
                g_LastMouse.x = LOWORD(lp);
                g_LastMouse.y = HIWORD(lp);
            }
            return 0;

        case WM_MOUSEMOVE:
            if (g_IsSelecting) {
                g_Selection.right = LOWORD(lp);
                g_Selection.bottom = HIWORD(lp);
                InvalidateRect(hwnd, NULL, FALSE);
            } else if (g_IsDrawing) {
                POINT pt = {LOWORD(lp), HIWORD(lp)};
                HPEN hPen;
                if (g_CurrentTool == TOOL_PEN) {
                    hPen = CreatePen(PS_SOLID, g_PenSize, g_CurrentColor);
                } else {
                    // Eraser: draw with the original screen pixels
                    // This is tricky. Let's just use a white pen for now or copy from g_hScreenDC
                    hPen = CreatePen(PS_SOLID, g_PenSize * 4, RGB(255, 255, 255)); // Placeholder
                }
                
                if (g_CurrentTool == TOOL_ERASER) {
                    // Actual eraser: copy from original screen
                    HDC hdcScr = g_hScreenDC;
                    int r = g_PenSize * 4;
                    BitBlt(g_hDrawingDC, pt.x - r/2, pt.y - r/2, r, r, hdcScr, pt.x - r/2, pt.y - r/2, SRCCOPY);
                } else {
                    SelectObject(g_hDrawingDC, hPen);
                    MoveToEx(g_hDrawingDC, g_LastMouse.x, g_LastMouse.y, NULL);
                    LineTo(g_hDrawingDC, pt.x, pt.y);
                    DeleteObject(hPen);
                }
                
                g_LastMouse = pt;
                InvalidateRect(hwnd, NULL, FALSE);
            }
            return 0;

        case WM_RBUTTONDOWN:
            // Right click to cancel everything immediately
            if (g_hToolbarWnd) DestroyWindow(g_hToolbarWnd);
            g_hToolbarWnd = NULL;
            DestroyWindow(hwnd);
            g_hOverlayWnd = NULL;
            g_IsCaptured = FALSE;
            g_IsSelecting = FALSE;
            return 0;

        case WM_LBUTTONUP:
            if (g_IsSelecting) {
                g_IsSelecting = FALSE;
                
                // Normalize selection
                if (g_Selection.left > g_Selection.right) { int tmp = g_Selection.left; g_Selection.left = g_Selection.right; g_Selection.right = tmp; }
                if (g_Selection.top > g_Selection.bottom) { int tmp = g_Selection.top; g_Selection.top = g_Selection.bottom; g_Selection.bottom = tmp; }
                
                int w = g_Selection.right - g_Selection.left;
                int h = g_Selection.bottom - g_Selection.top;

                if (w < 5 && h < 5) {
                    // Selection too small, assume full screen or cancel? 
                    // Let's just treat it as a full screen capture if they just click
                    g_Selection.left = 0;
                    g_Selection.top = 0;
                    g_Selection.right = GetSystemMetrics(SM_CXSCREEN);
                    g_Selection.bottom = GetSystemMetrics(SM_CYSCREEN);
                }

                g_IsCaptured = TRUE;
                ShowToolbar();
                InvalidateRect(hwnd, NULL, FALSE);
            }
            g_IsDrawing = FALSE;
            return 0;

        case WM_KEYDOWN:
            if (wp == VK_ESCAPE) {
                DestroyWindow(g_hToolbarWnd);
                g_hToolbarWnd = NULL;
                DestroyWindow(hwnd);
                g_hOverlayWnd = NULL;
                g_IsCaptured = FALSE;
            }
            return 0;

        case WM_DESTROY:
            if (g_hScreenBmp) DeleteObject(g_hScreenBmp);
            if (g_hDrawingBmp) DeleteObject(g_hDrawingBmp);
            if (g_hScreenDC) DeleteDC(g_hScreenDC);
            if (g_hDrawingDC) DeleteDC(g_hDrawingDC);
            g_hScreenBmp = g_hDrawingBmp = NULL;
            g_hScreenDC = g_hDrawingDC = NULL;
            return 0;
    }
    return DefWindowProc(hwnd, msg, wp, lp);
}

// --- Toolbar Window ---
LRESULT CALLBACK ToolbarWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    static COLORREF colors[] = { RGB(255, 0, 0), RGB(0, 255, 0), RGB(0, 0, 255), RGB(255, 255, 0), RGB(0, 0, 0) };
    
    switch (msg) {
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);
            RECT r;
            GetClientRect(hwnd, &r);

            // Draw background (Rounded-like flat)
            HBRUSH hBg = CreateSolidBrush(RGB(30, 30, 30));
            FillRect(hdc, &r, hBg);
            DeleteObject(hBg);

            // Draw color buttons
            for (int i = 0; i < 5; i++) {
                HBRUSH hBrush = CreateSolidBrush(colors[i]);
                RECT rcBtn = { 20 + i * (COLOR_BUTTON_SIZE + 10), 15, 20 + i * (COLOR_BUTTON_SIZE + 10) + COLOR_BUTTON_SIZE, 15 + COLOR_BUTTON_SIZE };
                FillRect(hdc, &rcBtn, hBrush);
                if (g_CurrentColor == colors[i] && g_CurrentTool == TOOL_PEN) {
                    HBRUSH hWhite = CreateSolidBrush(RGB(255, 255, 255));
                    FrameRect(hdc, &rcBtn, hWhite);
                    DeleteObject(hWhite);
                }
                DeleteObject(hBrush);
            }

            // Draw Pen/Eraser icons
            HFONT hFont = CreateFont(20, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, "Arial");
            SelectObject(hdc, hFont);
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkMode(hdc, TRANSPARENT);

            RECT rcPen = { 230, 10, 270, 50 };
            DrawText(hdc, g_CurrentTool == TOOL_PEN ? "[P]" : " P ", -1, &rcPen, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            
            RECT rcEraser = { 280, 10, 320, 50 };
            DrawText(hdc, g_CurrentTool == TOOL_ERASER ? "[E]" : " E ", -1, &rcEraser, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

            RECT rcSave = { 340, 10, 380, 50 };
            DrawText(hdc, "Save", -1, &rcSave, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

            RECT rcCheck = { 390, 10, 430, 50 };
            DrawText(hdc, "Done", -1, &rcCheck, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

            DeleteObject(hFont);
            EndPaint(hwnd, &ps);
            return 0;
        }

        case WM_LBUTTONDOWN: {
            int x = LOWORD(lp);
            int y = HIWORD(lp);

            for (int i = 0; i < 5; i++) {
                RECT rcBtn = { 20 + i * (COLOR_BUTTON_SIZE + 10), 15, 20 + i * (COLOR_BUTTON_SIZE + 10) + COLOR_BUTTON_SIZE, 15 + COLOR_BUTTON_SIZE };
                if (x >= rcBtn.left && x <= rcBtn.right && y >= rcBtn.top && y <= rcBtn.bottom) {
                    g_CurrentColor = colors[i];
                    g_CurrentTool = TOOL_PEN;
                    InvalidateRect(hwnd, NULL, FALSE);
                    return 0;
                }
            }

            if (x >= 230 && x <= 270) { g_CurrentTool = TOOL_PEN; InvalidateRect(hwnd, NULL, FALSE); }
            if (x >= 280 && x <= 320) { g_CurrentTool = TOOL_ERASER; InvalidateRect(hwnd, NULL, FALSE); }
            if ((x >= 340 && x <= 380) || (x >= 390 && x <= 430)) {
                SaveScreenshot();
                DestroyWindow(g_hToolbarWnd);
                g_hToolbarWnd = NULL;
                DestroyWindow(g_hOverlayWnd);
                g_hOverlayWnd = NULL;
                g_IsCaptured = FALSE;
            }
            return 0;
        }

        case WM_KEYDOWN:
            if (wp == VK_ESCAPE) {
                SendMessage(g_hOverlayWnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
            return 0;
    }
    return DefWindowProc(hwnd, msg, wp, lp);
}

// --- Saving Logic ---

void SaveScreenshot() {
    // Crop drawing bitmap to selection
    int w = g_Selection.right - g_Selection.left;
    int h = g_Selection.bottom - g_Selection.top;
    if (w <= 0 || h <= 0) return;

    HDC hdc = GetDC(NULL);
    HDC hMemDC = CreateCompatibleDC(hdc);
    HBITMAP hCropBmp = CreateCompatibleBitmap(hdc, w, h);
    SelectObject(hMemDC, hCropBmp);
    
    BitBlt(hMemDC, 0, 0, w, h, g_hDrawingDC, g_Selection.left, g_Selection.top, SRCCOPY);

    // Get Windows Screenshots folder path
    PWSTR pszPath = NULL;
    char szPath[MAX_PATH] = {0};
    HRESULT hr = SHGetKnownFolderPath(&FOLDERID_Screenshots, 0, NULL, &pszPath);
    
    if (SUCCEEDED(hr)) {
        WideCharToMultiByte(CP_ACP, 0, pszPath, -1, szPath, MAX_PATH, NULL, NULL);
        CoTaskMemFree(pszPath);
    } else {
        // Fallback to Pictures/Screenshots if FOLDERID_Screenshots fails
        hr = SHGetKnownFolderPath(&FOLDERID_Pictures, 0, NULL, &pszPath);
        if (SUCCEEDED(hr)) {
            WideCharToMultiByte(CP_ACP, 0, pszPath, -1, szPath, MAX_PATH, NULL, NULL);
            CoTaskMemFree(pszPath);
            strcat(szPath, "\\Screenshots");
            CreateDirectory(szPath, NULL);
        } else {
            // Last resort: local folder
            strcpy(szPath, "Screenshots");
            CreateDirectory(szPath, NULL);
        }
    }

    time_t t = time(NULL);
    struct tm *tm = localtime(&t);
    char fullPath[MAX_PATH];
    sprintf(fullPath, "%s\\Snap_%04d%02d%02d_%02d%02d%02d.png", szPath, tm->tm_year + 1900, tm->tm_mon + 1, tm->tm_mday, tm->tm_hour, tm->tm_min, tm->tm_sec);
    
    WCHAR wFullPath[MAX_PATH];
    MultiByteToWideChar(CP_ACP, 0, fullPath, -1, wFullPath, MAX_PATH);

    SaveBitmapToPNG(hCropBmp, wFullPath);

    DeleteObject(hCropBmp);
    DeleteDC(hMemDC);
    ReleaseDC(NULL, hdc);
}

HRESULT SaveBitmapToPNG(HBITMAP hBmp, const WCHAR* filename) {
    IWICImagingFactory *pFactory = NULL;
    IWICBitmapEncoder *pEncoder = NULL;
    IWICBitmapFrameEncode *pFrame = NULL;
    IWICStream *pStream = NULL;
    
    HRESULT hr = CoCreateInstance(&CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, &IID_IWICImagingFactory, (LPVOID*)&pFactory);
    if (FAILED(hr)) return hr;

    hr = pFactory->lpVtbl->CreateStream(pFactory, &pStream);
    if (SUCCEEDED(hr)) hr = pStream->lpVtbl->InitializeFromFilename(pStream, filename, GENERIC_WRITE);
    if (SUCCEEDED(hr)) hr = pFactory->lpVtbl->CreateEncoder(pFactory, &GUID_ContainerFormatPng, NULL, &pEncoder);
    if (SUCCEEDED(hr)) hr = pEncoder->lpVtbl->Initialize(pEncoder, (IStream*)pStream, WICBitmapEncoderNoCache);
    if (SUCCEEDED(hr)) hr = pEncoder->lpVtbl->CreateNewFrame(pEncoder, &pFrame, NULL);
    if (SUCCEEDED(hr)) hr = pFrame->lpVtbl->Initialize(pFrame, NULL);

    BITMAP bmp;
    GetObject(hBmp, sizeof(BITMAP), &bmp);
    hr = pFrame->lpVtbl->SetSize(pFrame, bmp.bmWidth, bmp.bmHeight);
    
    WICPixelFormatGUID format = GUID_WICPixelFormat32bppBGRA;
    hr = pFrame->lpVtbl->SetPixelFormat(pFrame, &format);
    
    IWICBitmap *pWICBitmap = NULL;
    hr = pFactory->lpVtbl->CreateBitmapFromHBITMAP(pFactory, hBmp, NULL, WICBitmapIgnoreAlpha, &pWICBitmap);
    if (SUCCEEDED(hr)) hr = pFrame->lpVtbl->WriteSource(pFrame, (IWICBitmapSource*)pWICBitmap, NULL);
    
    if (SUCCEEDED(hr)) hr = pFrame->lpVtbl->Commit(pFrame);
    if (SUCCEEDED(hr)) hr = pEncoder->lpVtbl->Commit(pEncoder);

    if (pWICBitmap) pWICBitmap->lpVtbl->Release(pWICBitmap);
    if (pFrame) pFrame->lpVtbl->Release(pFrame);
    if (pEncoder) pEncoder->lpVtbl->Release(pEncoder);
    if (pStream) pStream->lpVtbl->Release(pStream);
    if (pFactory) pFactory->lpVtbl->Release(pFactory);
    
    return hr;
}
