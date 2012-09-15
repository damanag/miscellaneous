#pragma once

class CUrl2PngEventSink;

class CUrl2Png :
	public CWindowImpl<CUrl2Png>
{
public:
	CUrl2Png();

public:
	BEGIN_MSG_MAP(CUrl2Png)
		MESSAGE_HANDLER(WM_CREATE, OnCreate)
		MESSAGE_HANDLER(WM_SIZE, OnSize)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
	END_MSG_MAP()

public:
	LRESULT OnCreate (UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnSize (UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnDestroy (UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

public:
	BOOL SaveSnapshot(IDispatch* pdisp, VARIANT* purl);

public:
	LPCTSTR m_lpszURI;
	LPCTSTR m_lpszFileName;

protected:
	CComPtr<IWebBrowser2> m_spWebBrowser2;
	CComObject<CUrl2PngEventSink>* m_pEventSink;

	HWND m_hwndWebBrowser;
	DWORD m_dwCookie;
};
