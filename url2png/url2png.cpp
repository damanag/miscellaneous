#include "stdafx.h"
#include "Url2PngEventSink.h"
#include "Url2Png.h"

CUrl2Png::CUrl2Png()
{
	m_hwndWebBrowser = NULL;
	m_dwCookie = 0;
}

LRESULT CUrl2Png::OnCreate(UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	HRESULT hr;
	CComPtr<IUnknown> spUnk;
	RECT oldRect;

	GetClientRect(&oldRect);

	m_hwndWebBrowser = ::CreateWindow(_T(ATLAXWIN_CLASS),
		m_lpszURI,
		WS_CHILD|WS_DISABLED/*|WS_POPUP*/,
		oldRect.top,
		oldRect.left,
		oldRect.right,
		oldRect.bottom,
		m_hWnd,
		NULL,
		::GetModuleHandle(NULL),
		NULL);
	if (m_hwndWebBrowser == NULL)
		return -1;

	hr = AtlAxGetControl(m_hwndWebBrowser, &spUnk);
	if (FAILED(hr) || spUnk == NULL )
		return -1;

	hr = spUnk->QueryInterface(IID_IWebBrowser2, (void**)&m_spWebBrowser2);
	if (FAILED(hr) || m_spWebBrowser2 == NULL)
		return -1;

	hr = CComObject<CUrl2PngEventSink>::CreateInstance(&m_pEventSink);
	if (FAILED(hr) || m_pEventSink == NULL)
		return -1;

	m_pEventSink->m_pUrl2Png = this;

	hr = AtlAdvise(m_spWebBrowser2, m_pEventSink->GetUnknown(), DIID_DWebBrowserEvents2, &m_dwCookie);
	if (FAILED(hr))
		return -1;

	return 0;
}

LRESULT CUrl2Png::OnSize(UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	if (m_hWnd != NULL)
		::MoveWindow(m_hWnd, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);

	return 0;
}

LRESULT CUrl2Png::OnDestroy(UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	HRESULT hr;

	if (m_dwCookie != 0)
		hr = AtlUnadvise(m_spWebBrowser2, DIID_DWebBrowserEvents2, m_dwCookie);

	PostQuitMessage(0);

	return 0;
}

BOOL CUrl2Png::SaveSnapshot(IDispatch* pDisp, VARIANT* purl)
{
	HRESULT hr;

	CComPtr<IDispatch> spDisp;
	CComPtr<IDispatch> spWebDisp;
	CComPtr<IHTMLDocument2> spDoc2;
	CComPtr<IHTMLDocument3> spDoc3;
	CComPtr<IHTMLElement> spElem;
	CComPtr<IHTMLElement2> spElem2;
	CComPtr<IViewObject2> spViewObject2;

	long lBodyHeight;
	long lBodyWidth;
	long lRootHeight;
	long lRootWidth;
	long lHeight;
	long lWidth;

	if (m_spWebBrowser2 != pDisp)
		return false;

	hr = m_spWebBrowser2->get_Document(&spDisp);
	if (FAILED(hr) || spDisp == NULL)
		return true;

	hr = spDisp->QueryInterface(IID_IHTMLDocument2, (void**)&spDoc2);
	if (FAILED(hr) || spDoc2 == NULL)
		return true;

	hr = spDoc2->get_body(&spElem);
	if (FAILED(hr) || spElem == NULL)
		return true;

	hr = spElem->QueryInterface(IID_IHTMLElement2, (void**)&spElem2);
	if (FAILED(hr) || spElem2 == NULL)
		return true;

	hr = spElem2->get_scrollHeight(&lBodyHeight);
	if (FAILED(hr))
		return true;

	hr = spElem2->get_scrollWidth(&lBodyWidth);
	if (FAILED(hr))
		return true;

	spElem.Release();
	spElem2.Release();

	hr = spDisp->QueryInterface(IID_IHTMLDocument3, (void**)&spDoc3);
	if (FAILED(hr) || spDoc3 == NULL)
		return true;

	hr = spDoc3->get_documentElement(&spElem);
	if (FAILED(hr) || spElem == NULL)
		return true;

	hr = spElem->QueryInterface(IID_IHTMLElement2, (void**)&spElem2);
	if (FAILED(hr) || spElem2 == NULL)
		return true;

	hr = spElem2->get_scrollHeight(&lRootHeight);
	if (FAILED(hr))
		return true;

	hr = spElem2->get_scrollWidth(&lRootWidth);
	if (FAILED(hr))
		return true;

	lWidth = lBodyWidth;
	lHeight = lRootHeight > lBodyHeight ? lRootHeight : lBodyHeight;

	MoveWindow(0, 0, lWidth, lHeight, TRUE);
	::MoveWindow(m_hwndWebBrowser, 0, 0, lWidth, lHeight, TRUE);

	hr = m_spWebBrowser2->QueryInterface(IID_IViewObject2, (void**)&spViewObject2);
	if (FAILED(hr) || spViewObject2 == NULL)
		return true;

	HDC hdcMain = GetDC();
	HDC hdcMem = CreateCompatibleDC(hdcMain);
	HBITMAP hBitmap = CreateCompatibleBitmap(hdcMain, lWidth, lHeight);
	SelectObject(hdcMem, hBitmap);

	RECTL rcBounds = { 0, 0, lWidth, lHeight };
	hr = spViewObject2->Draw(DVASPECT_CONTENT, -1, NULL, NULL, hdcMain, hdcMem, &rcBounds, NULL, NULL, 0);
	if (SUCCEEDED(hr))
	{
		CImage image;
		image.Create(lWidth, lHeight, 24);

		CImageDC imageDC(image);
		::BitBlt(imageDC, 0, 0, lWidth, lHeight, hdcMem, 0, 0, SRCCOPY);

		image.Save(m_lpszFileName);
	}

	return true;
}
