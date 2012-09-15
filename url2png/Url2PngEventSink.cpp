#include "stdafx.h"
#include "Url2Png.h"
#include "Url2PngEventSink.h"

CUrl2PngEventSink::CUrl2PngEventSink()
{
	m_pUrl2Png = NULL;
}

STDMETHODIMP CUrl2PngEventSink::GetTypeInfoCount(UINT*)
{
	return E_NOTIMPL;
}

STDMETHODIMP CUrl2PngEventSink::GetTypeInfo(UINT, LCID, ITypeInfo**)
{
	return E_NOTIMPL;
}

STDMETHODIMP CUrl2PngEventSink::GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*)
{
	return E_NOTIMPL;
}

STDMETHODIMP CUrl2PngEventSink::Invoke(DISPID dispid, REFIID, LCID, WORD, DISPPARAMS* pdispparams, VARIANT*, EXCEPINFO*, UINT*)
{
	if (dispid != DISPID_DOCUMENTCOMPLETE)
		return S_OK;

	if (pdispparams->cArgs != 2)
		return S_OK;

	if (pdispparams->rgvarg[0].vt != (VT_VARIANT | VT_BYREF))
		return S_OK;

	if (pdispparams->rgvarg[1].vt != VT_DISPATCH)
		return S_OK;

	if (m_pUrl2Png->SaveSnapshot(pdispparams->rgvarg[1].pdispVal,pdispparams->rgvarg[0].pvarVal))
		m_pUrl2Png->PostMessage(WM_CLOSE);

	return S_OK;
}
