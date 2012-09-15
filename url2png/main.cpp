#include "stdafx.h"
#include "Url2Png.h"

static const GUID IID_IUrl2Png = { 0x90ce882b, 0x629c, 0x48c6, { 0xb2, 0xa2, 0x54, 0x61, 0xe9, 0xd6, 0x6e, 0xa2 } };

static CComModule _Main;

int _tmain(int argc, TCHAR *argv[])
{
	if (argc != 3) {
		printf("Usage: url2png URL PNG\n");
		printf("Example: url2png http://www.google.com google.png\n");
		return EXIT_FAILURE;
	}

	HRESULT hr = _Main.Init(NULL, ::GetModuleHandle(NULL), &IID_IUrl2Png);
	if (FAILED(hr))
		return EXIT_FAILURE;

	if (!AtlAxWinInit())
		return EXIT_FAILURE;

	CUrl2Png ieSanpWnd;
	RECT rcMain = { 0, 0, 10, 10 };

	ieSanpWnd.m_lpszURI = argv[1];
	ieSanpWnd.m_lpszFileName = argv[2];
	ieSanpWnd.Create(NULL, rcMain, _T("Url2Png"), WS_POPUP);

	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0)) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	_Main.Term();

	return EXIT_SUCCESS;
}
