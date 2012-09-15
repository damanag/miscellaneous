#pragma once

class CUrl2Png;

class CUrl2PngEventSink :
	public CComObjectRootEx <CComSingleThreadModel>,
	public IDispatch
{
public:
	CUrl2PngEventSink();

public:
	BEGIN_COM_MAP(CUrl2PngEventSink)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY_IID(DIID_DWebBrowserEvents2, IDispatch)
	END_COM_MAP()

	STDMETHOD(GetTypeInfoCount)(UINT*);
	STDMETHOD(GetTypeInfo)(UINT, LCID, ITypeInfo**);
	STDMETHOD(GetIDsOfNames)(REFIID, LPOLESTR*, UINT, LCID, DISPID*);
	STDMETHOD(Invoke)(DISPID, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*);

public:
	CUrl2Png* m_pUrl2Png;
};
