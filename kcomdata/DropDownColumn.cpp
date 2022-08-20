// DropDownColumn.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "DropDownColumn.h"
#include "KCOMRichDropDownCtl.h"
#include "KCOMRichComboCtl.h"
#include "GridInner.h"
#include "DropDownRealWnd.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDropDownColumn

IMPLEMENT_DYNCREATE(CDropDownColumn, CCmdTarget)

IMPLEMENT_OLETYPELIB(CDropDownColumn, _tlid, _wVerMajor, _wVerMinor)

CDropDownColumn::CDropDownColumn()
{
	EnableAutomation();
	EnableTypeLib();
	
	m_pDropDownCtrl = NULL;
	m_pComboCtrl = NULL;
}

CDropDownColumn::SetCtrlPtr(CKCOMRichDropDownCtrl * pCtrl)
{
	m_pDropDownCtrl = pCtrl;
}

CDropDownColumn::SetCtrlPtr(CKCOMRichComboCtrl * pCtrl)
{
	m_pComboCtrl = pCtrl;
}

CDropDownColumn::~CDropDownColumn()
{
}


void CDropDownColumn::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(CDropDownColumn, CCmdTarget)
	//{{AFX_MSG_MAP(CDropDownColumn)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CDropDownColumn, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CDropDownColumn)
	DISP_PROPERTY_EX(CDropDownColumn, "Alignment", GetAlignment, SetAlignment, VT_I4)
	DISP_PROPERTY_EX(CDropDownColumn, "BackColor", GetBackColor, SetBackColor, VT_COLOR)
	DISP_PROPERTY_EX(CDropDownColumn, "Caption", GetCaption, SetCaption, VT_BSTR)
	DISP_PROPERTY_EX(CDropDownColumn, "CaptionAlignment", GetCaptionAlignment, SetCaptionAlignment, VT_I4)
	DISP_PROPERTY_EX(CDropDownColumn, "Case", GetCase, SetCase, VT_I4)
	DISP_PROPERTY_EX(CDropDownColumn, "DataField", GetDataField, SetDataField, VT_BSTR)
	DISP_PROPERTY_EX(CDropDownColumn, "ForeColor", GetForeColor, SetForeColor, VT_COLOR)
	DISP_PROPERTY_EX(CDropDownColumn, "HeadBackColor", GetHeadBackColor, SetHeadBackColor, VT_COLOR)
	DISP_PROPERTY_EX(CDropDownColumn, "HeadForeColor", GetHeadForeColor, SetHeadForeColor, VT_COLOR)
	DISP_PROPERTY_EX(CDropDownColumn, "Name", GetName, SetName, VT_BSTR)
	DISP_PROPERTY_EX(CDropDownColumn, "Visible", GetVisible, SetVisible, VT_BOOL)
	DISP_PROPERTY_EX(CDropDownColumn, "Width", GetWidth, SetWidth, VT_I4)
	DISP_PROPERTY_EX(CDropDownColumn, "Text", GetText, SetNotSupported, VT_BSTR)
	DISP_FUNCTION(CDropDownColumn, "CellText", CellText, VT_BSTR, VTS_VARIANT)
	DISP_DEFVALUE(CDropDownColumn, "Text")
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IDropDownColumn to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {491325C3-3068-11D3-A7FE-0080C8763FA4}
static const IID IID_IDropDownColumn =
{ 0x491325c3, 0x3068, 0x11d3, { 0xa7, 0xfe, 0x0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };

BEGIN_INTERFACE_MAP(CDropDownColumn, CCmdTarget)
	INTERFACE_PART(CDropDownColumn, IID_IDropDownColumn, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDropDownColumn message handlers

BOOL CDropDownColumn::GetDispatchIID(IID *pIID)
{
	*pIID = IID_IDropDownColumn;

	return TRUE;
}

long CDropDownColumn::GetAlignment() 
{
	return prop.nAlignment;
}

void CDropDownColumn::SetAlignment(long nNewValue) 
{
	if (nNewValue < 0 || nNewValue > 2)
		return;

	prop.nAlignment = nNewValue;

	DWORD dwAlign;

	switch (prop.nAlignment)
	{
	case 0:
		dwAlign = DT_LEFT;
		break;

	case 1:
		dwAlign = DT_RIGHT;
		break;

	case 2:
		dwAlign = DT_CENTER;
		break;
	}

	ROWCOL nCol;
	if (m_pDropDownCtrl)
	{
		if (m_pDropDownCtrl->m_pDrawGrid)
		{
			nCol = m_pDropDownCtrl->m_pDrawGrid->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDrawGrid->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetHorizontalAlignment(
				dwAlign));
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
		}
		else if (m_pDropDownCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetHorizontalAlignment(
				dwAlign));
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
	else if (m_pComboCtrl)
	{
		if (m_pComboCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pComboCtrl->m_pDropDownRealWnd->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetHorizontalAlignment(
				dwAlign));
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
}

OLE_COLOR CDropDownColumn::GetBackColor() 
{
	if (prop.clrBack == DEFAULTCOLOR)
	{
		if (m_pDropDownCtrl)
			return m_pDropDownCtrl->GetBackColor();
		else if (m_pComboCtrl)
			return m_pComboCtrl->GetBackColor();
	}

	return prop.clrBack;
}

void CDropDownColumn::SetBackColor(OLE_COLOR nNewValue) 
{
	if (prop.clrBack == nNewValue)
		return;

	prop.clrBack = nNewValue;

	ROWCOL nCol;
	if (m_pDropDownCtrl)
	{
		if (m_pDropDownCtrl->m_pDrawGrid)
		{
			nCol = m_pDropDownCtrl->m_pDrawGrid->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDrawGrid->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetInterior(
				m_pDropDownCtrl->TranslateColor(prop.clrBack)));
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
		}
		else if (m_pDropDownCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetInterior(
				m_pDropDownCtrl->TranslateColor(prop.clrBack)));
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
	else if (m_pComboCtrl)
	{
		if (m_pComboCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pComboCtrl->m_pDropDownRealWnd->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetInterior(
				m_pComboCtrl->TranslateColor(prop.clrBack)));
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
}

BSTR CDropDownColumn::GetCaption() 
{
	return prop.strCaption.AllocSysString();
}

void CDropDownColumn::SetCaption(LPCTSTR lpszNewValue) 
{
	prop.strCaption = lpszNewValue;

	ROWCOL nCol;
	if (m_pDropDownCtrl)
	{
		if (m_pDropDownCtrl->m_pDrawGrid)
		{
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->SetLockReadOnly(FALSE);
	
			nCol = m_pDropDownCtrl->m_pDrawGrid->GetColFromIndex(prop.nColIndex);
			m_pDropDownCtrl->m_pDrawGrid->SetValueRange(CGXRange(0, nCol), prop.strCaption);

			m_pDropDownCtrl->m_pDrawGrid->GetParam()->SetLockReadOnly(TRUE);
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
		}
		else if (m_pDropDownCtrl->m_pDropDownRealWnd)
		{
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
	
			nCol = m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			m_pDropDownCtrl->m_pDropDownRealWnd->SetValueRange(CGXRange(0, nCol), prop.strCaption);

			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
	else if (m_pComboCtrl)
	{
		if (m_pComboCtrl->m_pDropDownRealWnd)
		{
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
	
			nCol = m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			m_pComboCtrl->m_pDropDownRealWnd->SetValueRange(CGXRange(0, nCol), prop.strCaption);

			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
}

long CDropDownColumn::GetCaptionAlignment() 
{
	return prop.nCaptionAlignment;
}

void CDropDownColumn::SetCaptionAlignment(long nNewValue) 
{
	prop.nCaptionAlignment = nNewValue;

	DWORD dwAlign;

	switch (prop.nCaptionAlignment)
	{
	case 0:
		dwAlign = DT_LEFT;
		break;

	case 1:
		dwAlign = DT_RIGHT;
		break;

	case 2:
		dwAlign = DT_CENTER;
		break;

	case 3:
		dwAlign = prop.nAlignment == 0 ? DT_LEFT : (prop.nAlignment == 1 ? DT_RIGHT : DT_CENTER);
		break;
	}

	ROWCOL nCol;
	if (m_pDropDownCtrl)
	{
		if (m_pDropDownCtrl->m_pDrawGrid)
		{
			nCol = m_pDropDownCtrl->m_pDrawGrid->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDrawGrid->RowHeaderStyle().SetHorizontalAlignment(dwAlign);
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
		}
		else if (m_pDropDownCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->RowHeaderStyle().SetHorizontalAlignment(dwAlign);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
	else if (m_pComboCtrl)
	{
		if (m_pComboCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pComboCtrl->m_pDropDownRealWnd->RowHeaderStyle().SetHorizontalAlignment(dwAlign);
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
}

long CDropDownColumn::GetCase() 
{
	return prop.nCase;
}

void CDropDownColumn::SetCase(long nNewValue) 
{
	if (nNewValue < 0 || nNewValue > 2)
		return;

	prop.nCase = nNewValue;

	ROWCOL nCol;
	if (m_pDropDownCtrl)
	{
		if (m_pDropDownCtrl->m_pDrawGrid)
		{
			nCol = m_pDropDownCtrl->m_pDrawGrid->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDrawGrid->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetUserAttribute(
				GX_IDS_UA_CASE, prop.nCase));
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
		}
		else if (m_pDropDownCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetUserAttribute(
				GX_IDS_UA_CASE, prop.nCase));
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
	else if (m_pComboCtrl)
	{
		if (m_pComboCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pComboCtrl->m_pDropDownRealWnd->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetUserAttribute(
				GX_IDS_UA_CASE, prop.nCase));
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
}

BSTR CDropDownColumn::GetDataField() 
{
	return prop.strDataField.AllocSysString();
}

void CDropDownColumn::SetDataField(LPCTSTR lpszNewValue) 
{
	if (!prop.strDataField.CompareNoCase(lpszNewValue))
		return;

	prop.strDataField = lpszNewValue;

	if (m_pDropDownCtrl)
	{
		prop.nDataField = m_pDropDownCtrl->GetFieldFromStr(prop.strDataField);
		
		ROWCOL nCol = m_pDropDownCtrl->GetColOrdinalFromIndex(prop.nColIndex);
		m_pDropDownCtrl->m_apColumnsProp[nCol]->strDataField = prop.strDataField;
		m_pDropDownCtrl->m_apColumnsProp[nCol]->nDataField = prop.nDataField;

		m_pDropDownCtrl->Bind();
	}
	else if (m_pComboCtrl)
	{
		prop.nDataField = m_pComboCtrl->GetFieldFromStr(prop.strDataField);
		
		ROWCOL nCol = m_pComboCtrl->GetColOrdinalFromIndex(prop.nColIndex);
		m_pComboCtrl->m_apColumnsProp[nCol]->strDataField = prop.strDataField;
		m_pComboCtrl->m_apColumnsProp[nCol]->nDataField = prop.nDataField;

		m_pComboCtrl->Bind();
	}
}

OLE_COLOR CDropDownColumn::GetForeColor() 
{
	if (prop.clrFore == DEFAULTCOLOR)
	{
		if (m_pDropDownCtrl)
			return m_pDropDownCtrl->GetForeColor();
		else if (m_pComboCtrl)
			return m_pComboCtrl->GetForeColor();
	}

	return prop.clrFore;
}

void CDropDownColumn::SetForeColor(OLE_COLOR nNewValue) 
{
	if (prop.clrFore == nNewValue)
		return;

	prop.clrFore = nNewValue;

	ROWCOL nCol;
	if (m_pDropDownCtrl)
	{
		if (m_pDropDownCtrl->m_pDrawGrid)
		{
			nCol = m_pDropDownCtrl->m_pDrawGrid->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDrawGrid->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetTextColor(
				m_pDropDownCtrl->TranslateColor(prop.clrFore)));
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
		}
		else if (m_pDropDownCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetTextColor(
				m_pDropDownCtrl->TranslateColor(prop.clrFore)));
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}			
	}
	else if (m_pComboCtrl)
	{
		if (m_pComboCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pComboCtrl->m_pDropDownRealWnd->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetTextColor(
				m_pComboCtrl->TranslateColor(prop.clrFore)));
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}			
	}
}

OLE_COLOR CDropDownColumn::GetHeadBackColor() 
{
	if (prop.clrHeadBack == DEFAULTCOLOR)
	{
		if (m_pDropDownCtrl)
			return m_pDropDownCtrl->m_clrHeadBack;
		else if (m_pComboCtrl)
			return m_pComboCtrl->m_clrHeadBack;
	}

	return prop.clrHeadBack;
}

void CDropDownColumn::SetHeadBackColor(OLE_COLOR nNewValue) 
{
	if (prop.clrHeadBack == nNewValue)
		return;

	prop.clrHeadBack = nNewValue;

	ROWCOL nCol;
	if (m_pDropDownCtrl)
	{
		if (m_pDropDownCtrl->m_pDrawGrid)
		{
			nCol = m_pDropDownCtrl->m_pDrawGrid->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDrawGrid->SetStyleRange(CGXRange(0, nCol), CGXStyle().SetInterior(
				m_pDropDownCtrl->TranslateColor(prop.clrHeadBack)));
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
		}
		else if (m_pDropDownCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->SetStyleRange(CGXRange(0, nCol), CGXStyle().SetInterior(
				m_pDropDownCtrl->TranslateColor(prop.clrHeadBack)));
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}			
	}
	else if (m_pComboCtrl)
	{
		if (m_pComboCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pComboCtrl->m_pDropDownRealWnd->SetStyleRange(CGXRange(0, nCol), CGXStyle().SetInterior(
				m_pComboCtrl->TranslateColor(prop.clrHeadBack)));
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}			
	}
}

OLE_COLOR CDropDownColumn::GetHeadForeColor() 
{
	if (prop.clrHeadFore == DEFAULTCOLOR)
	{
		if (m_pDropDownCtrl)
			return m_pDropDownCtrl->m_clrHeadFore;
		else if (m_pComboCtrl)
			return m_pComboCtrl->m_clrHeadFore;
	}

	return prop.clrHeadFore;
}

void CDropDownColumn::SetHeadForeColor(OLE_COLOR nNewValue) 
{
	if (prop.clrHeadFore == nNewValue)
		return;

	prop.clrHeadFore = nNewValue;

	ROWCOL nCol;
	if (m_pDropDownCtrl)
	{
		if (m_pDropDownCtrl->m_pDrawGrid)
		{
			nCol = m_pDropDownCtrl->m_pDrawGrid->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDrawGrid->SetStyleRange(CGXRange(0, nCol), CGXStyle().SetTextColor(
				m_pDropDownCtrl->TranslateColor(prop.clrHeadFore)));
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
		}
		else if (m_pDropDownCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->SetStyleRange(CGXRange(0, nCol), CGXStyle().SetTextColor(
				m_pDropDownCtrl->TranslateColor(prop.clrHeadFore)));
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}			
	}
	else if (m_pComboCtrl)
	{
		if (m_pComboCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pComboCtrl->m_pDropDownRealWnd->SetStyleRange(CGXRange(0, nCol), CGXStyle().SetTextColor(
				m_pComboCtrl->TranslateColor(prop.clrHeadFore)));
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}			
	}
}

BSTR CDropDownColumn::GetName() 
{
	return prop.strName.AllocSysString();
}

void CDropDownColumn::SetName(LPCTSTR lpszNewValue) 
{
	if (lstrlen(lpszNewValue) == 0 || !prop.strName.CompareNoCase(lpszNewValue))
		return;

	if (m_pDropDownCtrl)
	{
		for (int i = 0; i < m_pDropDownCtrl->m_apColumns.GetSize(); i ++)
			if (!m_pDropDownCtrl->m_apColumns[i]->prop.strName.CompareNoCase(
				lpszNewValue))
				return;
	}
	else if (m_pComboCtrl)
	{
		for (int i = 0; i < m_pComboCtrl->m_apColumns.GetSize(); i ++)
			if (!m_pComboCtrl->m_apColumns[i]->prop.strName.CompareNoCase(
				lpszNewValue))
				return;
	}

	prop.strName = lpszNewValue;
}

BOOL CDropDownColumn::GetVisible() 
{
	return prop.bVisible;
}

void CDropDownColumn::SetVisible(BOOL bNewValue) 
{
	prop.bVisible = bNewValue;

	ROWCOL nCol;
	if (m_pDropDownCtrl)
	{
		if (m_pDropDownCtrl->m_pDrawGrid)
		{
			nCol = m_pDropDownCtrl->m_pDrawGrid->GetColFromIndex(prop.nColIndex);

			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);

			m_pDropDownCtrl->m_pDrawGrid->HideCols(nCol, nCol, !prop.bVisible);

			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
		}
		else if (m_pDropDownCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);

			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);

			m_pDropDownCtrl->m_pDropDownRealWnd->HideCols(nCol, nCol, !prop.bVisible);

			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
	else if (m_pComboCtrl)
	{
		if (m_pComboCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);

			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);

			m_pComboCtrl->m_pDropDownRealWnd->HideCols(nCol, nCol, !prop.bVisible);

			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
}

long CDropDownColumn::GetWidth() 
{
	ROWCOL nCol;
 	
	if (m_pDropDownCtrl)
	{
		if (m_pDropDownCtrl->m_pDrawGrid)
		{
			nCol = m_pDropDownCtrl->m_pDrawGrid->GetColFromIndex(prop.nColIndex);
			prop.nWidth = m_pDropDownCtrl->m_pDrawGrid->GetColWidth(nCol);
		}
		else if (m_pDropDownCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			prop.nWidth = m_pDropDownCtrl->m_pDropDownRealWnd->GetColWidth(nCol);
		}
	}
	else if (m_pComboCtrl)
	{
		if (m_pComboCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
			prop.nWidth = m_pComboCtrl->m_pDropDownRealWnd->GetColWidth(nCol);
		}
	}

	return prop.nWidth;
}

void CDropDownColumn::SetWidth(long nNewValue) 
{
	if (nNewValue < 0)
		return;

	prop.nWidth = nNewValue;

	ROWCOL nCol;
	if (m_pDropDownCtrl)
	{
		if (m_pDropDownCtrl->m_pDrawGrid)
		{
			nCol = m_pDropDownCtrl->m_pDrawGrid->GetColFromIndex(prop.nColIndex);

			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDrawGrid->SetColWidth(nCol, nCol, nNewValue);
			
			m_pDropDownCtrl->m_pDrawGrid->GetParam()->EnableUndo(FALSE);
		}
		else if (m_pDropDownCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);

			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->SetColWidth(nCol, nCol, nNewValue);
			
			m_pDropDownCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
	else if (m_pComboCtrl)
	{
		if (m_pComboCtrl->m_pDropDownRealWnd)
		{
			nCol = m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);

			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
			
			m_pComboCtrl->m_pDropDownRealWnd->SetColWidth(nCol, nCol, nNewValue);
			
			m_pComboCtrl->m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		}
	}
}

// the IDropdownColumn has no datatype property, so the value is just 
// the text
BSTR CDropDownColumn::GetText() 
{
	CString strResult;
	ROWCOL nRow, nCol;

	if (m_pDropDownCtrl)
	{
		m_pDropDownCtrl->m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
		nCol = m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);

		strResult = m_pDropDownCtrl->m_pDropDownRealWnd->GetValueRowCol(nRow, nCol);
	}
	else if (m_pComboCtrl)
	{
		m_pComboCtrl->m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
		nCol = m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);

		strResult = m_pComboCtrl->m_pDropDownRealWnd->GetValueRowCol(nRow, nCol);
	}

	if (prop.nCase == 1)
		strResult.MakeUpper();
	else if (prop.nCase == 2)
		strResult.MakeLower();

	return strResult.AllocSysString();
}

// extract the text from control, for this text is the most latest
BSTR CDropDownColumn::GetCellText(long RowIndex) 
{
	if (m_pDropDownCtrl)
	{
		if ((ROWCOL)RowIndex > m_pDropDownCtrl->m_pDropDownRealWnd->GetRowCount() 
			|| RowIndex	< 1)
			m_pDropDownCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);
		
		CString strResult;
		ROWCOL nCol = m_pDropDownCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
		
		CGXStyle style;
		
		if (m_pDropDownCtrl->m_pDropDownRealWnd->GetStyleRowCol(RowIndex, nCol, style))
			m_pDropDownCtrl->m_pDropDownRealWnd->GetControl(RowIndex, nCol)->GetControlText(
			strResult, RowIndex, nCol, NULL, style);
		
		if (prop.nCase == 1)
			strResult.MakeUpper();
		else if (prop.nCase == 2)
			strResult.MakeLower();
		
		return strResult.AllocSysString();
	}
	else if (m_pComboCtrl)
	{
		if ((ROWCOL)RowIndex > m_pComboCtrl->m_pDropDownRealWnd->GetRowCount() 
			|| RowIndex	< 1)
			m_pComboCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);
		
		CString strResult;
		ROWCOL nCol = m_pComboCtrl->m_pDropDownRealWnd->GetColFromIndex(prop.nColIndex);
		
		CGXStyle style;
		
		if (m_pComboCtrl->m_pDropDownRealWnd->GetStyleRowCol(RowIndex, nCol, style))
			m_pComboCtrl->m_pDropDownRealWnd->GetControl(RowIndex, nCol)->GetControlText(
			strResult, RowIndex, nCol, NULL, style);
		
		if (prop.nCase == 1)
			strResult.MakeUpper();
		else if (prop.nCase == 2)
			strResult.MakeLower();
		
		return strResult.AllocSysString();
	}

	return NULL;
}

BSTR CDropDownColumn::CellText(const VARIANT FAR& Bookmark)
{
	if (m_pDropDownCtrl)
		return GetCellText(m_pDropDownCtrl->GetRowFromBmk(&Bookmark));
	else if (m_pComboCtrl)
		return GetCellText(m_pComboCtrl->GetRowFromBmk(&Bookmark));

	return NULL;
}
