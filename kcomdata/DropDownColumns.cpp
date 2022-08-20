// DropDownColumns.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "DropDownColumns.h"
#include "KCOMRichDropDownCtl.h"
#include "KCOMRichComboCtl.h"
#include "DropDownColumn.h"
#include "DropDownRealWnd.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDropDownColumns

IMPLEMENT_DYNCREATE(CDropDownColumns, CCmdTarget)

IMPLEMENT_OLETYPELIB(CDropDownColumns, _tlid, _wVerMajor, _wVerMinor)

CDropDownColumns::CDropDownColumns()
{
	EnableAutomation();
	EnableTypeLib();
	
	m_pDropDownCtrl = NULL;
	m_pComboCtrl = NULL;
}

CDropDownColumns::SetCtrlPtr(CKCOMRichDropDownCtrl * pCtrl)
{
	m_pDropDownCtrl = pCtrl;
}

CDropDownColumns::SetCtrlPtr(CKCOMRichComboCtrl * pCtrl)
{
	m_pComboCtrl = pCtrl;
}

CDropDownColumns::~CDropDownColumns()
{
}


void CDropDownColumns::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(CDropDownColumns, CCmdTarget)
	//{{AFX_MSG_MAP(CDropDownColumns)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CDropDownColumns, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CDropDownColumns)
	DISP_PROPERTY_EX(CDropDownColumns, "Count", GetCount, SetNotSupported, VT_I2)
	DISP_FUNCTION(CDropDownColumns, "Remove", Remove, VT_EMPTY, VTS_VARIANT)
	DISP_FUNCTION(CDropDownColumns, "Add", Add, VT_EMPTY, VTS_VARIANT)
	DISP_FUNCTION(CDropDownColumns, "RemoveAll", RemoveAll, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_PARAM(CDropDownColumns, "Item", GetItem, SetNotSupported, VT_DISPATCH, VTS_VARIANT)
	DISP_DEFVALUE(CDropDownColumns, "Item")
	//}}AFX_DISPATCH_MAP
	DISP_PROPERTY_EX_ID(CDropDownColumns, "_NewEnum", DISPID_NEWENUM, _NewEnum, SetNotSupported, VT_UNKNOWN)
END_DISPATCH_MAP()

// Note: we add support for IID_IDropDownColumns to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {491325C0-3068-11D3-A7FE-0080C8763FA4}
static const IID IID_IDropDownColumns =
{ 0x491325c0, 0x3068, 0x11d3, { 0xa7, 0xfe, 0x0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };

BEGIN_INTERFACE_MAP(CDropDownColumns, CCmdTarget)
	INTERFACE_PART(CDropDownColumns, IID_IDropDownColumns, Dispatch)
END_INTERFACE_MAP()

CDropDownColumns::XEnumVARIANT::XEnumVARIANT()
{
	m_nCurrent = 0;
}

STDMETHODIMP_(ULONG) CDropDownColumns::XEnumVARIANT::AddRef()
{   
	METHOD_PROLOGUE(CDropDownColumns, EnumVARIANT)
	return pThis->ExternalAddRef();
}   

STDMETHODIMP_(ULONG) CDropDownColumns::XEnumVARIANT::Release()
{   
	METHOD_PROLOGUE(CDropDownColumns, EnumVARIANT)
	return pThis->ExternalRelease();
}   

STDMETHODIMP CDropDownColumns::XEnumVARIANT::QueryInterface(REFIID iid, void FAR* FAR* ppvObj)
{   
	METHOD_PROLOGUE(CDropDownColumns, EnumVARIANT)
	return (HRESULT)pThis->ExternalQueryInterface((void FAR*)&iid, ppvObj);
}   

// IEnumVARIANT::Next
// 
STDMETHODIMP CDropDownColumns::XEnumVARIANT::Next(ULONG celt, VARIANT FAR* rgvar, ULONG FAR* pceltFetched)
{
	METHOD_PROLOGUE(CDropDownColumns, EnumVARIANT)

	CDropDownColumn * pItem = NULL;

	// pceltFetched can legally == 0
	//                                           
	if (pceltFetched != NULL)
		*pceltFetched = 0;
	else if (celt > 1)
	{   
		return ResultFromScode(E_INVALIDARG); 
	}

    // Retrieve the next celt elements.
	if (pThis->m_pDropDownCtrl)
	{
		if (m_nCurrent < pThis->m_pDropDownCtrl->m_apColumns.GetSize())
		{
			VariantInit(&rgvar[0]);
			
			pItem = pThis->m_pDropDownCtrl->m_apColumns[m_nCurrent ++];
			rgvar[0].vt = VT_DISPATCH;
			rgvar[0].pdispVal = pItem->GetIDispatch(TRUE) ;
			if (pceltFetched != NULL)
				*pceltFetched = 1;
			
			return NOERROR;
		}
	}
	else
	{
		if (m_nCurrent < pThis->m_pComboCtrl->m_apColumns.GetSize())
		{
			VariantInit(&rgvar[0]);
			
			pItem = pThis->m_pComboCtrl->m_apColumns[m_nCurrent ++];
			rgvar[0].vt = VT_DISPATCH;
			rgvar[0].pdispVal = pItem->GetIDispatch(TRUE) ;
			if (pceltFetched != NULL)
				*pceltFetched = 1;
			
			return NOERROR;
		}
	}

	return S_FALSE;
}

// IEnumVARIANT::Skip
//
STDMETHODIMP CDropDownColumns::XEnumVARIANT::Skip(unsigned long celt)
{
    METHOD_PROLOGUE(CDropDownColumns, EnumVARIANT)

	if (pThis->m_pDropDownCtrl)
	{
		while (m_nCurrent < pThis->m_pDropDownCtrl->m_apColumns.GetSize() && celt --)
		    m_nCurrent ++;
	}
	else
	{
		while (m_nCurrent < pThis->m_pComboCtrl->m_apColumns.GetSize() && celt --)
		    m_nCurrent ++;
	}
    
    return celt == 0 ? NOERROR : S_FALSE;
}

STDMETHODIMP CDropDownColumns::XEnumVARIANT::Reset()
{
    METHOD_PROLOGUE(CDropDownColumns, EnumVARIANT)
 
	m_nCurrent = 0;
	
	return NOERROR;
}

STDMETHODIMP CDropDownColumns::XEnumVARIANT::Clone(IEnumVARIANT FAR* FAR* ppenum)
{
    METHOD_PROLOGUE(CDropDownColumns, EnumVARIANT)   
    CDropDownColumns* p = new CDropDownColumns;
	
	if (pThis->m_pDropDownCtrl)
		p->SetCtrlPtr(pThis->m_pDropDownCtrl);
	else
		p->SetCtrlPtr(pThis->m_pComboCtrl);

    if (p)
    {
        p->m_xEnumVARIANT.m_nCurrent = m_nCurrent ;
        return NOERROR ;    
    }
    else
        return ResultFromScode(E_OUTOFMEMORY);
}

/////////////////////////////////////////////////////////////////////////////
// CDropDownColumns message handlers

LPUNKNOWN CDropDownColumns::_NewEnum()
{
	return GetIDispatch(TRUE);
}

BOOL CDropDownColumns::GetDispatchIID(IID *pIID)
{
	*pIID = IID_IDropDownColumns;

	return TRUE;
}

short CDropDownColumns::GetCount() 
{
	if (m_pDropDownCtrl)
		return m_pDropDownCtrl->m_apColumns.GetSize();
	else
		return m_pComboCtrl->m_apColumns.GetSize();
}

// the col/row index in our controls begins from 1, which is recommended
// by MS.
void CDropDownColumns::Remove(const VARIANT FAR& ColIndex) 
{
	if (m_pDropDownCtrl)
	{
		if(m_pDropDownCtrl->m_nDataMode == 0)
			m_pDropDownCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
		
		ROWCOL nColIndex = GetRowColFromVariant(ColIndex);
		if (nColIndex < 1 || nColIndex > (ROWCOL)m_pDropDownCtrl->m_apColumns.GetSize())
			m_pDropDownCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);
		
		m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
		
		if (m_pDropDownCtrl->m_pDropDownRealWnd->RemoveCols(nColIndex, nColIndex))
		{
			nColIndex = m_pDropDownCtrl->GetColOrdinalFromCol(nColIndex);
			CDropDownColumn * pColNew = m_pDropDownCtrl->m_apColumns[nColIndex];
			delete pColNew;
			m_pDropDownCtrl->m_apColumns.RemoveAt(nColIndex);
		}
		
		m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
		m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
		
		if (m_pDropDownCtrl->m_apColumns.GetSize() == 0)
			m_pDropDownCtrl->m_nColsUsed = 0;
	}
	else if (m_pComboCtrl)
	{
		if(m_pComboCtrl->m_nDataMode == 0)
			m_pComboCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
		
		ROWCOL nColIndex = GetRowColFromVariant(ColIndex);
		if (nColIndex < 1 || nColIndex > (ROWCOL)m_pComboCtrl->m_apColumns.GetSize())
			m_pComboCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);
		
		m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		m_pComboCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
		
		if (m_pComboCtrl->m_pDropDownRealWnd->RemoveCols(nColIndex, nColIndex))
		{
			nColIndex = m_pComboCtrl->GetColOrdinalFromCol(nColIndex);
			CDropDownColumn * pColNew = m_pComboCtrl->m_apColumns[nColIndex];
			delete pColNew;
			m_pComboCtrl->m_apColumns.RemoveAt(nColIndex);
		}
		
		m_pComboCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
		m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
		
		if (m_pComboCtrl->m_apColumns.GetSize() == 0)
			m_pComboCtrl->m_nColsUsed = 0;
	}
}

void CDropDownColumns::Add(const VARIANT FAR& ColIndex) 
{
	if (m_pDropDownCtrl)
	{
		if (m_pDropDownCtrl->m_apColumns.GetSize() == 255)
			m_pDropDownCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_MAX255COLS);
		else if(m_pDropDownCtrl->m_nDataMode == 0)
			m_pDropDownCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
		
		ROWCOL nColIndex = GetRowColFromVariant(ColIndex);
		if (nColIndex < 1 || nColIndex > (ROWCOL)m_pDropDownCtrl->m_apColumns.GetSize() + 1)
			m_pDropDownCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);
		
		m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
		
		if (m_pDropDownCtrl->m_pDropDownRealWnd->InsertCols(nColIndex, 1))
		{
			CDropDownColumn * pColNew = new CDropDownColumn;
			m_pDropDownCtrl->InitColumnObject(nColIndex, pColNew);
			m_pDropDownCtrl->m_apColumns.InsertAt(nColIndex - 1, pColNew);
		}
		
		m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
		m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
	}
	else if (m_pComboCtrl)
	{
		if (m_pComboCtrl->m_apColumns.GetSize() == 255)
			m_pComboCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_MAX255COLS);
		else if(m_pComboCtrl->m_nDataMode == 0)
			m_pComboCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
		
		ROWCOL nColIndex = GetRowColFromVariant(ColIndex);
		if (nColIndex < 1 || nColIndex > (ROWCOL)m_pComboCtrl->m_apColumns.GetSize() + 1)
			m_pComboCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);
		
		m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		m_pComboCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
		
		if (m_pComboCtrl->m_pDropDownRealWnd->InsertCols(nColIndex, 1))
		{
			CDropDownColumn * pColNew = new CDropDownColumn;
			m_pComboCtrl->InitColumnObject(nColIndex, pColNew);
			m_pComboCtrl->m_apColumns.InsertAt(nColIndex - 1, pColNew);
		}
		
		m_pComboCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
		m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
	}
}

void CDropDownColumns::RemoveAll() 
{
	if (m_pDropDownCtrl)
	{
		if(m_pDropDownCtrl->m_nDataMode == 0)
			m_pDropDownCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
		
		m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
		m_pDropDownCtrl->m_pDropDownRealWnd->SetColCount(0);
		m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
		m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
		
		for (int i = 0; i < m_pDropDownCtrl->m_apColumns.GetSize(); i ++)
			delete m_pDropDownCtrl->m_apColumns[i];
		
		m_pDropDownCtrl->m_apColumns.RemoveAll();
		m_pDropDownCtrl->m_nColsUsed = 0;
	}
	else if (m_pComboCtrl)
	{
		if(m_pComboCtrl->m_nDataMode == 0)
			m_pComboCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
		
		m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		m_pComboCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
		m_pComboCtrl->m_pDropDownRealWnd->SetColCount(0);
		m_pComboCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
		m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
		
		for (int i = 0; i < m_pComboCtrl->m_apColumns.GetSize(); i ++)
			delete m_pComboCtrl->m_apColumns[i];
		
		m_pComboCtrl->m_apColumns.RemoveAll();
		m_pComboCtrl->m_nColsUsed = 0;
	}
}

LPDISPATCH CDropDownColumns::GetItem(const VARIANT FAR& ColIndex) 
{
	ROWCOL nColIndex = GetRowColFromVariant(ColIndex);
	
	if (m_pDropDownCtrl)
	{
		if (nColIndex < 1 || nColIndex > (ROWCOL)m_pDropDownCtrl->m_apColumns.GetSize())
		return NULL;

		nColIndex = m_pDropDownCtrl->GetColOrdinalFromCol(nColIndex);
		return m_pDropDownCtrl->m_apColumns[nColIndex]->GetIDispatch(TRUE);
	}
	else if (m_pComboCtrl)
	{
		if (nColIndex < 1 || nColIndex > (ROWCOL)m_pComboCtrl->m_apColumns.GetSize())
		return NULL;

		nColIndex = m_pComboCtrl->GetColOrdinalFromCol(nColIndex);
		return m_pComboCtrl->m_apColumns[nColIndex]->GetIDispatch(TRUE);
	}

	return NULL;
}

ROWCOL CDropDownColumns::GetRowColFromVariant(VARIANT va)
{
	USES_CONVERSION;

	COleVariant vaNew = va;
	CString str;

	vaNew.ChangeType(VT_BSTR);
	str = OLE2T(vaNew.bstrVal);
	if (str.GetLength() == 0)
		return GX_INVALID;

	if (str[0] >= (TCHAR)'0' && str[0] <= (TCHAR)'9')
	{
		vaNew.ChangeType(VT_I4);
		return vaNew.lVal;
	}

	if (m_pDropDownCtrl)
	{
		for (int i = 0; i < m_pDropDownCtrl->m_apColumns.GetSize(); i ++)
		if (!m_pDropDownCtrl->m_apColumns[i]->prop.strName.CompareNoCase(str))
			return m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(m_pDropDownCtrl->m_apColumns[i]
				->prop.nColIndex);
	}
	else if (m_pComboCtrl)
	{
		for (int i = 0; i < m_pComboCtrl->m_apColumns.GetSize(); i ++)
		if (!m_pComboCtrl->m_apColumns[i]->prop.strName.CompareNoCase(str))
			return m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(m_pComboCtrl->m_apColumns[i]
				->prop.nColIndex);
	}
	
	return GX_INVALID;
}