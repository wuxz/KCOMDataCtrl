// Columns.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "GridColumns.h"
#include "KCOMRichGridCtl.h"
#include "GridColumn.h"
#include "GridInner.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGridColumns

IMPLEMENT_DYNCREATE(CGridColumns, CCmdTarget)

IMPLEMENT_OLETYPELIB(CGridColumns, _tlid, _wVerMajor, _wVerMinor)

CGridColumns::CGridColumns()
{
	EnableAutomation();
	EnableTypeLib();
	
	m_pCtrl = NULL;
}

CGridColumns::SetCtrlPtr(CKCOMRichGridCtrl * pCtrl)
{
	m_pCtrl = pCtrl;
}

CGridColumns::~CGridColumns()
{
}


void CGridColumns::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(CGridColumns, CCmdTarget)
	//{{AFX_MSG_MAP(CGridColumns)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CGridColumns, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CGridColumns)
	DISP_PROPERTY_EX(CGridColumns, "Count", GetCount, SetNotSupported, VT_I2)
	DISP_FUNCTION(CGridColumns, "Remove", Remove, VT_EMPTY, VTS_VARIANT)
	DISP_FUNCTION(CGridColumns, "Add", Add, VT_EMPTY, VTS_VARIANT)
	DISP_FUNCTION(CGridColumns, "RemoveAll", RemoveAll, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_PARAM(CGridColumns, "Item", Item, SetNotSupported, VT_DISPATCH, VTS_VARIANT)
	DISP_DEFVALUE(CGridColumns, "Item")
	//}}AFX_DISPATCH_MAP
	DISP_PROPERTY_EX_ID(CGridColumns, "_NewEnum", DISPID_NEWENUM, _NewEnum, SetNotSupported, VT_UNKNOWN)

END_DISPATCH_MAP()

// Note: we add support for IID_IGridColumns to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {61724100-12B3-11D3-A7FE-0080C8763FA4}
static const IID IID_IGridColumns =
{ 0x61724100, 0x12b3, 0x11d3, { 0xa7, 0xfe, 0x0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };

BEGIN_INTERFACE_MAP(CGridColumns, CCmdTarget)
	INTERFACE_PART(CGridColumns, IID_IGridColumns, Dispatch)
	INTERFACE_PART(CGridColumns, IID_IEnumVARIANT, EnumVARIANT)
END_INTERFACE_MAP()


CGridColumns::XEnumVARIANT::XEnumVARIANT()
{
	m_nCurrent = 0;
}

STDMETHODIMP_(ULONG) CGridColumns::XEnumVARIANT::AddRef()
{   
	METHOD_PROLOGUE(CGridColumns, EnumVARIANT)
	return pThis->ExternalAddRef();
}   

STDMETHODIMP_(ULONG) CGridColumns::XEnumVARIANT::Release()
{   
	METHOD_PROLOGUE(CGridColumns, EnumVARIANT)
	return pThis->ExternalRelease();
}   

STDMETHODIMP CGridColumns::XEnumVARIANT::QueryInterface(REFIID iid, void FAR* FAR* ppvObj)
{   
	METHOD_PROLOGUE(CGridColumns, EnumVARIANT)
	return (HRESULT)pThis->ExternalQueryInterface((void FAR*)&iid, ppvObj);
}   

// IEnumVARIANT::Next
// 
STDMETHODIMP CGridColumns::XEnumVARIANT::Next(ULONG celt, VARIANT FAR* rgvar, ULONG FAR* pceltFetched)
{
	METHOD_PROLOGUE(CGridColumns, EnumVARIANT)

	CGridColumn * pItem = NULL;

	// pceltFetched can legally == 0
	//                                           
	if (pceltFetched != NULL)
		*pceltFetched = 0;
	else if (celt > 1)
	{   
		return ResultFromScode(E_INVALIDARG); 
	}

    // Retrieve the next celt elements.
	if (m_nCurrent < pThis->m_pCtrl->m_apColumns.GetSize())
	{
		VariantInit(&rgvar[0]);

		pItem = pThis->m_pCtrl->m_apColumns[m_nCurrent ++];
		rgvar[0].vt = VT_DISPATCH;
		rgvar[0].pdispVal = pItem->GetIDispatch(TRUE) ;
		if (pceltFetched != NULL)
			*pceltFetched = 1;

		return NOERROR;
	}

	return S_FALSE;
}

// IEnumVARIANT::Skip
//
STDMETHODIMP CGridColumns::XEnumVARIANT::Skip(unsigned long celt)
{
    METHOD_PROLOGUE(CGridColumns, EnumVARIANT)

	while (m_nCurrent < pThis->m_pCtrl->m_apColumns.GetSize() && celt --)
        m_nCurrent ++;
    
    return celt == 0 ? NOERROR : S_FALSE;
}

STDMETHODIMP CGridColumns::XEnumVARIANT::Reset()
{
    METHOD_PROLOGUE(CGridColumns, EnumVARIANT)
 
	m_nCurrent = 0;
	
	return NOERROR;
}

STDMETHODIMP CGridColumns::XEnumVARIANT::Clone(IEnumVARIANT FAR* FAR* ppenum)
{
    METHOD_PROLOGUE(CGridColumns, EnumVARIANT)   
    CGridColumns* p = new CGridColumns;
	p->SetCtrlPtr(pThis->m_pCtrl);
    if (p)
    {
        p->m_xEnumVARIANT.m_nCurrent = m_nCurrent ;
        return NOERROR ;    
    }
    else
        return ResultFromScode(E_OUTOFMEMORY);
}

/////////////////////////////////////////////////////////////////////////////
// CGridColumns message handlers

short CGridColumns::GetCount() 
{
	return m_pCtrl->m_apColumns.GetSize();
}

// remember that the index begins from 1

void CGridColumns::Remove(const VARIANT FAR& ColIndex) 
{
	if(m_pCtrl->m_nDataMode == 0)
		m_pCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);

	ROWCOL nColIndex = GetRowColFromVariant(ColIndex);
	if (nColIndex < 1 || nColIndex > (ROWCOL)m_pCtrl->m_apColumns.GetSize())
		m_pCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);
	BOOL bLock = m_pCtrl->m_pGridInner->GetParam()->IsLockReadOnly();
	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(FALSE);

	if (m_pCtrl->m_pGridInner->RemoveCols(nColIndex, nColIndex))
	{
		nColIndex = m_pCtrl->GetColOrdinalFromCol(nColIndex);
		CGridColumn * pColNew = m_pCtrl->m_apColumns[nColIndex];
		delete pColNew;
		m_pCtrl->m_apColumns.RemoveAt(nColIndex);
	}
	
	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(TRUE);

	if (m_pCtrl->m_apColumns.GetSize() == 0)
		m_pCtrl->m_nColsUsed = 0;
}

void CGridColumns::Add(const VARIANT FAR& ColIndex) 
{
	if (m_pCtrl->m_apColumns.GetSize() == 255)
		m_pCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_MAX255COLS);
	else if(m_pCtrl->m_nDataMode == 0)
		m_pCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);

	ROWCOL nColIndex = GetRowColFromVariant(ColIndex);
	if (nColIndex < 1 || nColIndex > (ROWCOL)m_pCtrl->m_apColumns.GetSize() + 1)
		m_pCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);
	BOOL bLock = m_pCtrl->m_pGridInner->GetParam()->IsLockReadOnly();
	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
	
	if (m_pCtrl->m_pGridInner->InsertCols(nColIndex, 1))
	{
		CGridColumn * pColNew = new CGridColumn;
		m_pCtrl->InitColumnObject(nColIndex, pColNew);
		m_pCtrl->m_apColumns.InsertAt(nColIndex - 1, pColNew);
	}
	
	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(TRUE);
}

void CGridColumns::RemoveAll() 
{
	if(m_pCtrl->m_nDataMode == 0)
		m_pCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);
	BOOL bLock = m_pCtrl->m_pGridInner->GetParam()->IsLockReadOnly();
	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
	m_pCtrl->m_pGridInner->SetColCount(0);
	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(TRUE);

	for (int i = 0; i < m_pCtrl->m_apColumns.GetSize(); i ++)
		delete m_pCtrl->m_apColumns[i];

	m_pCtrl->m_apColumns.RemoveAll();
	m_pCtrl->m_nColsUsed = 0;
}

LPUNKNOWN CGridColumns::_NewEnum()
{
	return GetIDispatch(TRUE);
}

ROWCOL CGridColumns::GetRowColFromVariant(VARIANT va)
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

	for (int i = 0; i < m_pCtrl->m_apColumns.GetSize(); i ++)
	{
		if (!m_pCtrl->m_apColumns[i]->prop.strName.CompareNoCase(str))
			return m_pCtrl->m_pGridInner->GetColFromIndex(m_pCtrl->m_apColumns[i]
			->prop.nColIndex);
	}

	return GX_INVALID;
}

LPDISPATCH CGridColumns::Item(const VARIANT FAR& ColIndex) 
{
	ROWCOL nColIndex = GetRowColFromVariant(ColIndex);
	if (nColIndex < 1 || nColIndex > (ROWCOL)m_pCtrl->m_apColumns.GetSize())
		m_pCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	nColIndex = m_pCtrl->GetColOrdinalFromCol(nColIndex);
	return m_pCtrl->m_apColumns[nColIndex]->GetIDispatch(TRUE);
}

BOOL CGridColumns::GetDispatchIID(IID *pIID)
{
	*pIID = IID_IGridColumns;

	return TRUE;
}
