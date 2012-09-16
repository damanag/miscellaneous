#ifndef _WIN32_WINNT
#  define _WIN32_WINNT 0x500
#endif

#include <windows.h>
#include <crtdbg.h>
#include <tchar.h>
#include <gdiplus.h>
#include <atlstr.h>

int getEncoderClsid(LPCWSTR lpszFormat, CLSID* pClsid);
int saveSmallIconToPng(LPCTSTR lpszExt, LPCTSTR lpszSavePath);

using namespace Gdiplus;

int getEncoderClsid(LPCWSTR lpszFormat, CLSID* pClsid)
{
	_ASSERTE(lpszFormat != NULL);
	_ASSERTE(pClsid != NULL);

	UINT num = 0;	// number of image encoders
	UINT size = 0;	// size of the image encoder array in bytes

	ImageCodecInfo* pImageCodecInfo = NULL;

	GetImageEncodersSize(&num, &size);
	if (size == 0)
		return -1;

	pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
	if (pImageCodecInfo == NULL)
		return -1;

	GetImageEncoders(num, size, pImageCodecInfo);

	for (UINT i=0; i<num; ++i) {
		if (wcscmp(pImageCodecInfo[i].MimeType, lpszFormat) == 0) {
			*pClsid = pImageCodecInfo[i].Clsid;
			free(pImageCodecInfo);
			return i;
		}
	}

	free(pImageCodecInfo);
	return -1;
}

BOOL saveSmallIconToPng(LPCTSTR lpszExt, LPCTSTR lpszFilename)
{
	_ASSERTE(lpszExt != NULL);
	_ASSERTE(lpszFilename != NULL);

	// exist file?
	if (GetFileAttributesW(lpszFilename) != 0xFFFFFFFF)
		return FALSE;
	if (GetLastError() != ERROR_FILE_NOT_FOUND)
		return FALSE;

	// Get the small icon that is registered in Windows.
	CString strDotExt;
	strDotExt.Format(_T(".%s"), lpszExt);

	SHFILEINFO shfi;
	SHGetFileInfo(strDotExt, 0, &shfi, sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_ICON | SHGFI_SHELLICONSIZE | SHGFI_SMALLICON);

	// Draw the icon on bitmap handle.
	HBRUSH hBursh = CreateSolidBrush(RGB(0xFF,0xFF,0xFF));
	HDC hDrawDC = CreateCompatibleDC(NULL);
	HDC hDC = CreateDC(_T("DISPLAY"), NULL, NULL, NULL);
	HBITMAP hBitmap = CreateCompatibleBitmap(hDC, 16, 16);
	HBITMAP hOldBitmap = (HBITMAP)SelectObject(hDrawDC, hBitmap);

	DrawIconEx(hDrawDC, 0, 0, shfi.hIcon, 0, 0, 0, hBursh, DI_NORMAL);
	SelectObject(hDrawDC, hOldBitmap);

	DeleteDC(hDC);
	DeleteDC(hDrawDC);
	DeleteObject(hBursh);
	DestroyIcon(shfi.hIcon);

	// Start GDI+.
	GdiplusStartupInput gdiplusStartupInput;
	ULONG_PTR gdiplusToken;
	GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

	// Get encoder.
	CLSID clsid;
	if (getEncoderClsid(L"image/png", &clsid) < 0) {
		DeleteObject(hBitmap);
		GdiplusShutdown(gdiplusToken);
		return FALSE;
	}

	// Icon bitmap => GDI+ bitmap => file.
	Bitmap* lpBitmap = Bitmap::FromHBITMAP(hBitmap, NULL);
	if (lpBitmap == NULL) {
		DeleteObject(hBitmap);
		GdiplusShutdown(gdiplusToken);
		return FALSE;
	}

	lpBitmap->Save(lpszFilename, &clsid, NULL);
	delete lpBitmap;

	DeleteObject(hBitmap);
	GdiplusShutdown(gdiplusToken);
	return TRUE;
}

int _tmain(int argc, _TCHAR* argv[])
{
	if (argc != 3) {
		printf("Usage: extracticon EXTENSION PNG\n");
		printf("Example: url2png gif gif.png\n");
		return EXIT_FAILURE;
	}

	if (saveSmallIconToPng(argv[1], argv[2]) == FALSE)
		return EXIT_FAILURE;

	return EXIT_SUCCESS;
}
