//=--------------------------------------------------------------------------=
// ADVCON.CPP:	IConnectionPoint and IPropertyNotifySink helpers
//=--------------------------------------------------------------------------=
// Copyright  1997  Microsoft Corporation.  All Rights Reserved.
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF 
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A 
// PARTICULAR PURPOSE.
//=--------------------------------------------------------------------------=
//
#include "stdafx.h"
#include <advcon.h>
#include "ipserver.h"
#include <adbmacrs.h>

//=--------------------------------------------------------------------------=
// AdviseConnect [helper]
//=--------------------------------------------------------------------------=
// Called to make an advisory connection
//
// Parameters:
//	punkNotifier	[in]	- Pointer to IUnknown of notifier object
//	riid			[in]	- IID of notification interface
//	punkNotify		[in]	- Pointer to IUnknown of notifyee objecct
//	pdwCookie		[out]	- Storage of connection cookie
//
// Output:
//	HRESULT			[out]	- Error code
//
// Notes:
//
HRESULT WINAPI AdviseConnect(LPUNKNOWN punkNotifier, REFIID riid, LPUNKNOWN punkNotify, DWORD* pdwCookie, BOOL fAddRef /*= TRUE */)
{
	CHECK_POINTER(punkNotifier);
	CHECK_POINTER(punkNotify);
	CHECK_POINTER(pdwCookie);

	LPCONNECTIONPOINTCONTAINER pCPC;
	HRESULT hr = punkNotifier->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&pCPC);

	RETURN_ON_FAILURE(hr);
	CHECK_POINTER(pCPC);

	LPCONNECTIONPOINT pCP;
	hr = pCPC->FindConnectionPoint(riid, &pCP);
	pCPC->Release();

	RETURN_ON_FAILURE(hr);
	CHECK_POINTER(pCP);

	// Set up advisory connection
	//
	hr = pCP->Advise(punkNotify, pdwCookie);
	pCP->Release();

	// Advise adds ref punkNotify, so if caller does not
	// want this side effect, then remove the ref
	//
	if (!fAddRef && SUCCEEDED(hr))
		punkNotify->Release();

	return hr;
}

//=--------------------------------------------------------------------------=
// AdviseDisconnect	[helper]
//=--------------------------------------------------------------------------=
// Called to disconnect an advisory connection. 
//
// Parameters:
//	punkNotifier	[in]	- Pointer to IUnknown of notifier object
//	riid			[in]	- IID of notification interface
//	punkNotify		[in]	- Pointer to IUnknown of notifyee objecct
//	dwCookie		[in]	- Cookie of a connection made prior to this call
//	
// Output:
//	HRESULT			[out]	- Error code
//
// Notes:
//      
//
HRESULT WINAPI AdviseDisconnect(LPUNKNOWN punkNotifier, REFIID riid, LPUNKNOWN punkNotify, DWORD dwCookie, BOOL fAddRef /*= TRUE*/)
{
	CHECK_POINTER(punkNotifier);

	if (fAddRef)
	{
		if (NULL == punkNotify)
			return E_INVALIDARG;

		CHECK_POINTER(punkNotify);
	}

	// When we call Unadvise, the connection point will Release us.  If we
	// didn't keep the reference count when we called Advise, we need to
	// AddRef now, to keep our reference count consistent.  Note that if
	// the Unadvise fails, then we need to undo this extra AddRef by
	// calling Release before we return.
	//
	if (!fAddRef)
		punkNotify->AddRef();

	LPCONNECTIONPOINTCONTAINER pCPC;
	HRESULT hr = punkNotifier->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&pCPC);

	RETURN_ON_FAILURE(hr);
	CHECK_POINTER(pCPC);

	LPCONNECTIONPOINT pCP;
	hr = pCPC->FindConnectionPoint(riid, &pCP);
	pCPC->Release();

	RETURN_ON_FAILURE(hr);
	CHECK_POINTER(pCP);

	// Remove the advisory connection
	//
	hr = pCP->Unadvise(dwCookie);
	pCP->Release();

	// If we failed, undo the earlier AddRef.
	//
	if (!fAddRef && FAILED(hr))
		punkNotify->Release();

	return hr;
}

/////////////////////////////////////////////////////////////////////////////
// CAdviseConnection
//

CAdviseConnection::CAdviseConnection(void) 
{
	m_guid			= GUID_NULL;
	m_punkNotifier	= NULL;
	m_punkNotify	= NULL;
	m_dwCookie		= NULL;

	m_fAddRefNotify	= FALSE;
	m_fAdvising		= FALSE;
}

CAdviseConnection::CAdviseConnection(REFIID riid) 
{
	m_guid			= riid;
	m_punkNotifier	= NULL;
	m_punkNotify	= NULL;
	m_dwCookie		= NULL;

	m_fAddRefNotify	= FALSE;
	m_fAdvising		= FALSE;
}

CAdviseConnection::~CAdviseConnection() 
{
	Unadvise();
}

HRESULT CAdviseConnection::SetNotifyInterface(REFIID riid)
{
	if (IsAdvising())
		return E_FAIL;

	if (!IS_SAME_GUIDS(riid, m_guid))
	{
		Unadvise();
		m_guid = riid;
	}
	return S_OK;
}

CAdviseConnection& CAdviseConnection::operator=(const CAdviseConnection& other)
{
	if (this != &other)
	{
		Unadvise();
		SetNotifyInterface(other.GetNotifyInterface());

		if (other.IsAdvising())
			Advise(other.m_punkNotifier, other.m_punkNotify, other.m_fAddRefNotify);
	}
	return *this;
}

HRESULT CAdviseConnection::Advise(IUnknown *punkNotifier, IUnknown *punkNotify, BOOL fAddRef /* = TRUE */)
{
	CHECK_POINTER(punkNotify);
	CHECK_POINTER(punkNotifier);

	if (IS_SAME_GUIDS(m_guid, GUID_NULL))
		return E_NOINTERFACE;

	Unadvise();
	HRESULT hr = AdviseConnect(punkNotifier, m_guid, punkNotify, &m_dwCookie, fAddRef);

	if (SUCCEEDED(hr))
	{
		m_punkNotifier = punkNotifier;
		m_punkNotify = punkNotify;
		m_fAdvising = TRUE;
		m_fAddRefNotify = fAddRef;
	}
	return hr;
}

HRESULT CAdviseConnection::Unadvise(void)
{
	HRESULT hr = S_OK;

	if (m_fAdvising) 
	{
		hr = AdviseDisconnect(m_punkNotifier, m_guid, m_punkNotify, m_dwCookie, m_fAddRefNotify);
		
		m_punkNotifier = NULL;
		m_punkNotify = NULL;
		m_dwCookie = NULL;
		m_fAddRefNotify = FALSE;
		m_fAdvising = FALSE;
	}
	return hr;
}

/////////////////////////////////////////////////////////////////////////////
// CConnectionsArray
//

HRESULT CConnectionsArray::Add(REFIID riid, IUnknown *punkNotifier, IUnknown *punkNotify, BOOL fAddRef /* TRUE */)
{
	CHECK_POINTER(punkNotifier);
	CHECK_POINTER(punkNotify);

	long lIndex = Find(riid);
	HRESULT hr = S_OK;

	if (-1 == lIndex)
	{
		CAdviseConnection ne(riid);
		if (!CArray<CAdviseConnection, CAdviseConnection&>::Add(ne))
			return E_OUTOFMEMORY;

		lIndex = GetUpperBound();
	}
	hr = m_items[lIndex].Advise(punkNotifier, punkNotify, fAddRef);
	
	return hr;
}

HRESULT CConnectionsArray::Remove(REFIID riid)
{
	long lIndex = Find(riid);
	HRESULT hr = S_OK;

	if (lIndex >= 0)
	{
		hr = m_items[lIndex].Unadvise();
		DeleteAt(lIndex);
	}
	return hr;
}

int CConnectionsArray::Find(REFIID riid)
{
	long lIndex = GetSize();

	while (lIndex--)
	{
		if (IS_SAME_GUIDS(m_items[lIndex].GetNotifyInterface(), riid))
			break;
	}
	return lIndex;
}

//=--------------------------------------------------------------------------=
// CPropNotifySink::CPropNotifySink
//=--------------------------------------------------------------------------=
// Constructor
//
// Parameters:
//	pPropNotify			[in]	- External IPropertyNotifySink interface
//	dispidPropNotifier	[in]	- DISPID of notifier object
//  fAddRefNotify		[in]    - TRUE to addref notifyee
//	fAddRefNotifier		[in]    - TRUE to addref notifier
//
// Output:
//
// Notes:
//	The DISPID will be passed to the Owner's OnChanged and OnRequestEdit.
//  The owner will use it to distinguish which of it's connected property
//	has changed.  Call GetDispID to get the actual property of the notifier
//	that is changing
//
CPropNotifySink::CPropNotifySink
(
	IPropertyNotifySink *pPropNotify, 
	DISPID dispidPropNotifier, 
	BOOL fAddRefNotify /* = TRUE */, 
	BOOL fAddRefNotifier /* = TRUE */
)
: CAdviseConnection(IID_IPropertyNotifySink)
{
	// This is cached without an AddRef.  Advise and Unadvise
	// will do all the ref'ing work
	//
	m_pPropNotify = pPropNotify;

	// fAddRefNotify(TRUE) should be used with care.  When the notifyee
	// object is implemented by a control, fAddRefNotify should be FALSE 
	// so that the control may be released by the host properly.
	// 
	m_fAddRefNotify = fAddRefNotify;

	// fAddRefSource(TRUE) should to ensure that the notifier does not
	// get released before Unadvise is called
	//
	m_fAddRefNotifier = fAddRefNotifier;

	m_dispidPropNotifier = dispidPropNotifier;
	m_dispidProp = 0;
}

//=--------------------------------------------------------------------------=
// CPropNotifySink::~CPropNotifySink
//=--------------------------------------------------------------------------=
// Destructor
//
// Parameters:
//
// Output:
//
// Notes:
//      
//
CPropNotifySink::~CPropNotifySink()
{
    // This is all we need to clean up
	//
    Unadvise();
}

//=--------------------------------------------------------------------------=
// CPropNotifySink::QueryInterface	[IUnknown]
//=--------------------------------------------------------------------------=
// Returns the interfaces that we support
//
// Parameters:
//	riid		[in]	- IID_* of interface to query
//	ppvObj		[out]	- Storage for interface pointer
//
// Output:
//	NRESULT		[out]	- E_NOINTERFACE if we don't support it
//
// Notes:
//
STDMETHODIMP CPropNotifySink::QueryInterface(REFIID riid, LPVOID FAR* ppvObj)
{
	IUnknown *punk;

	if (IS_SAME_GUIDS(IID_IUnknown, riid))
		punk = (IUnknown *)this;
	else if (IS_SAME_GUIDS(IID_IPropertyNotifySink, riid))
		punk = (IPropertyNotifySink *)this;
	else
		return E_NOINTERFACE;

	punk->AddRef();
	*ppvObj = (void*)punk;

	return S_OK;
}

//=--------------------------------------------------------------------------=
// CPropNotifySink::AddRef	[IUnknown]
//=--------------------------------------------------------------------------=
// Called to add to ref count. We pass this on to the owner.
//
// Parameters:
//
// Output:
//
// Notes:
//	
STDMETHODIMP_(ULONG) CPropNotifySink::AddRef(void)
{
	if (m_pPropNotify && m_fAddRefNotify)
		return m_pPropNotify->AddRef();

	return 0;
}

//=--------------------------------------------------------------------------=
// CPropNotifySink::Release	[IUnkown]
//=--------------------------------------------------------------------------=
// Called to release a ref count.  We pass this on to the owner
//
// Parameters:
//
// Output:
//
// Notes:
//      
//
STDMETHODIMP_(ULONG) CPropNotifySink::Release(void)
{
	if (m_pPropNotify && m_fAddRefNotify)
		return m_pPropNotify->Release();

	return 0;
}

//=--------------------------------------------------------------------------=
// CPropNotifySink::OnChanged	[IPropertyNotifySink]
//=--------------------------------------------------------------------------=
// Passes OnChanged to the owner
//
// Parameters:
//	dispid		[in]	- DISPID of notication object's property
//
// Output:
//	HRESULT		[out]	- Error code
//
// Notes:
//	We notify the owner with the assigned property DISPID, the owner
//	will then call GetDispID to get the actual DISPID of the notication
//	object's property that is changing
//
STDMETHODIMP CPropNotifySink::OnChanged(DISPID dispid)
{
	if (m_pPropNotify)
	{
		m_dispidProp = dispid;
		return m_pPropNotify->OnChanged(m_dispidPropNotifier);
	}
	return S_OK;
}

//=--------------------------------------------------------------------------=
// CPropNotifySink::OnRequestEdit	[IPropertyNotifySink]
//=--------------------------------------------------------------------------=
// Passes OnRequestEdit to the owner
//
// Parameters:
//	dispid		[in]	- DISPID of notication object's property
//
// Output:
//	HRESULT		[out]	- Error code
//
// Notes:
//	We notify the owner with the assigned property DISPID, the owner
//	will then call GetDispID to get the actual DISPID of the notication
//	object's property that is changing
//
STDMETHODIMP CPropNotifySink::OnRequestEdit(DISPID dispid)
{
	if (m_pPropNotify)
	{
		m_dispidProp = dispid;
		return m_pPropNotify->OnRequestEdit(m_dispidPropNotifier);
	}
	return S_OK;
}

//=--------------------------------------------------------------------------=
// CPropNotifySink::Advise
//=--------------------------------------------------------------------------=
// Called by the owner to establish a notification connection to the given
// object whose DISPID is given to the contructor.
//
// Parameters:
//	punkNotifier	[in]	- Notifier object
//
// Output:
//	HRESULT			[out]	- Error code
//
// Notes:
//	This method is designed to be called successively.  It will call Unadvise
//	to property disconnect any previous connection      
//
HRESULT CPropNotifySink::Advise(IUnknown *punkNotifier)
{
	// Nothing to do if notifier is the same
	//
	if (punkNotifier == GetNotifier()) return S_OK;

	// Remove any previous advisory connections
	//
	HRESULT hr = Unadvise();

	if (SUCCEEDED(hr) && NULL != punkNotifier)
	{
		hr = CAdviseConnection::Advise(punkNotifier, (IUnknown*)this);
	
		if (SUCCEEDED(hr))
		{
			// AddRef the notifier only if the caller specified so.
			//
			if (m_fAddRefNotifier)
				GetNotifier()->AddRef();
		}
	}
	return hr;
}

//=--------------------------------------------------------------------------=
// CPropNotifySink::Unadvise
//=--------------------------------------------------------------------------=
// Called to diconnnect from the most recent advisory connection
//
// Parameters:
//	
// Output:
//	HRESULT		[out]	- Error code
//
// Notes:
//	This method is design to be called anytime to ensure a disconnection,       
//	and will only disconnect if it needs to
//
HRESULT CPropNotifySink::Unadvise(void)
{
	HRESULT hr = S_OK;

	if (IsAdvising())
	{
		IUnknown *punkNotifier = GetNotifier();
		hr = CAdviseConnection::Unadvise();

		// Always release references even if unadvise fails
		//
		if (m_fAddRefNotifier)
			punkNotifier->Release();

		m_dispidProp = 0;
	}
	return hr;
}

