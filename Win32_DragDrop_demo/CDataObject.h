/* ** ** ** ** ** ** ** ** ** ** **

Win32 OLE Drag&Drop demo

Copyright 2018 Giuseppe Pischedda
All rights reserved

https://www.software-on-demand-ita.com

** ** ** ** ** ** ** ** ** ** ** **/
#pragma once
#include "stdafx.h"


class CDataObject : public IDataObject
{


public:

	static HRESULT STDMETHODCALLTYPE CreateInstance(IDataObject **pDataObject);



	/*
	** IUnknown **
	*/
	HRESULT STDMETHODCALLTYPE QueryInterface(_In_ REFIID riid, _Out_ void   **ppvObject);
	ULONG   STDMETHODCALLTYPE AddRef();
	ULONG   STDMETHODCALLTYPE Release();



    /*
	** IDataObject **
	*/

	HRESULT STDMETHODCALLTYPE  DAdvise(
		_In_  FORMATETC   *pformatetc,
		_In_  DWORD       advf,
		_In_  IAdviseSink *pAdvSink,
		_Out_ DWORD       *pdwConnection
	);


	HRESULT STDMETHODCALLTYPE  DUnadvise(
		_In_ DWORD dwConnection
	);


	HRESULT STDMETHODCALLTYPE EnumDAdvise(
		_Out_ IEnumSTATDATA **ppenumAdvise
	);


	HRESULT STDMETHODCALLTYPE EnumFormatEtc(
		_In_  DWORD          dwDirection,
		_Out_ IEnumFORMATETC **ppenumFormatEtc
	);


	HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(
		_In_  FORMATETC *pformatectIn,
		_Out_ FORMATETC *pformatetcOut
	);


	HRESULT STDMETHODCALLTYPE GetData(
		_In_  FORMATETC *pformatetcIn,
		_Out_ STGMEDIUM *pmedium
	);


	HRESULT STDMETHODCALLTYPE QueryGetData(
		_In_ FORMATETC *pformatetc
	);

	HRESULT STDMETHODCALLTYPE GetDataHere(
		_In_      FORMATETC *pformatetc,
		_Inout_ STGMEDIUM *pmedium
	);


	HRESULT  STDMETHODCALLTYPE SetData(
		_In_ FORMATETC *pformatetc,
		_In_ STGMEDIUM *pmedium,
		_In_ BOOL      fRelease
	);


	
	
	//Set text function for this sample
	void SetText(CString szText);

private:

	//Private ctor and destructor
    CDataObject();
	
	virtual  
	~CDataObject();


	//Other private members
	UINT m_refCount;
	CString m_szText;



};