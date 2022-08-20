// Column.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "GridColumn.h"
#include "KCOMRichGridCtl.h"
#include "GridInner.h"
#include "GridColumns.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGridColumn

IMPLEMENT_DYNCREATE(CGridColumn, CCmdTarget)
IMPLEMENT_OLETYPELIB(CGridColumn, _tlid, _wVerMajor, _wVerMinor)

CGridColumn::CGridColumn()
{
	EnableAutomation();
	EnableTypeLib();

	m_pCtrl = NULL;
	m_dropDownWnd = NULL;
}

CGridColumn::SetCtrlPtr(CKCOMRichGridCtrl * pCtrl)
{
	m_pCtrl = pCtrl;
}

CGridColumn::~CGridColumn()
{
}


void CGridColumn::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(CGridColumn, CCmdTarget)
	//{{AFX_MSG_MAP(CGridColumn)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CGridColumn, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CGridColumn)
	DISP_PROPERTY_NOTIFY(CGridColumn, "DropDownWnd", m_dropDownWnd, OnDropDownWndChanged, VT_HANDLE)
	DISP_PROPERTY_EX(CGridColumn, "FieldLen", GetFieldLen, SetFieldLen, VT_I2)
	DISP_PROPERTY_EX(CGridColumn, "AllowSizing", GetAllowSizing, SetAllowSizing, VT_BOOL)
	DISP_PROPERTY_EX(CGridColumn, "HeadForeColor", GetHeadForeColor, SetHeadForeColor, VT_COLOR)
	DISP_PROPERTY_EX(CGridColumn, "HeadBackColor", GetHeadBackColor, SetHeadBackColor, VT_COLOR)
	DISP_PROPERTY_EX(CGridColumn, "DataField", GetDataField, SetDataField, VT_BSTR)
	DISP_PROPERTY_EX(CGridColumn, "Locked", GetLocked, SetLocked, VT_BOOL)
	DISP_PROPERTY_EX(CGridColumn, "Visible", GetVisible, SetVisible, VT_BOOL)
	DISP_PROPERTY_EX(CGridColumn, "Width", GetWidth, SetWidth, VT_I4)
	DISP_PROPERTY_EX(CGridColumn, "DataType", GetDataType, SetDataType, VT_I4)
	DISP_PROPERTY_EX(CGridColumn, "Selected", GetSelected, SetSelected, VT_BOOL)
	DISP_PROPERTY_EX(CGridColumn, "Style", GetStyle, SetStyle, VT_I4)
	DISP_PROPERTY_EX(CGridColumn, "CaptionAlignment", GetCaptionAlignment, SetCaptionAlignment, VT_I4)
	DISP_PROPERTY_EX(CGridColumn, "ColChanged", GetColChanged, SetNotSupported, VT_BOOL)
	DISP_PROPERTY_EX(CGridColumn, "Name", GetName, SetName, VT_BSTR)
	DISP_PROPERTY_EX(CGridColumn, "Nullable", GetNullable, SetNullable, VT_BOOL)
	DISP_PROPERTY_EX(CGridColumn, "Mask", GetMask, SetMask, VT_BSTR)
	DISP_PROPERTY_EX(CGridColumn, "PromptChar", GetPromptChar, SetPromptChar, VT_BSTR)
	DISP_PROPERTY_EX(CGridColumn, "Caption", GetCaption, SetCaption, VT_BSTR)
	DISP_PROPERTY_EX(CGridColumn, "Alignment", GetAlignment, SetAlignment, VT_I4)
	DISP_PROPERTY_EX(CGridColumn, "ForeColor", GetForeColor, SetForeColor, VT_COLOR)
	DISP_PROPERTY_EX(CGridColumn, "BackColor", GetBackColor, SetBackColor, VT_COLOR)
	DISP_PROPERTY_EX(CGridColumn, "Case", GetCase, SetCase, VT_I4)
	DISP_PROPERTY_EX(CGridColumn, "Text", GetText, SetText, VT_BSTR)
	DISP_PROPERTY_EX(CGridColumn, "Value", GetValue, SetValue, VT_VARIANT)
	DISP_PROPERTY_EX(CGridColumn, "PromptInclude", GetPromptInclude, SetPromptInclude, VT_BOOL)
	DISP_PROPERTY_EX(CGridColumn, "ListCount", GetListCount, SetNotSupported, VT_I2)
	DISP_PROPERTY_EX(CGridColumn, "ListIndex", GetListIndex, SetListIndex, VT_I2)
	DISP_PROPERTY_EX(CGridColumn, "ComboBoxStyle", GetComboBoxStyle, SetComboBoxStyle, VT_I4)
	DISP_FUNCTION(CGridColumn, "CellText", CellText, VT_BSTR, VTS_VARIANT)
	DISP_FUNCTION(CGridColumn, "CellValue", CellValue, VT_VARIANT, VTS_VARIANT)
	DISP_FUNCTION(CGridColumn, "IsCellValid", IsCellValid, VT_BOOL, VTS_NONE)
	DISP_FUNCTION(CGridColumn, "AddItem", AddItem, VT_EMPTY, VTS_BSTR VTS_VARIANT)
	DISP_FUNCTION(CGridColumn, "RemoveAll", RemoveAll, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CGridColumn, "RemoveItem", RemoveItem, VT_EMPTY, VTS_I2)
	DISP_PROPERTY_PARAM(CGridColumn, "List", GetList, SetList, VT_BSTR, VTS_I2)
	DISP_DEFVALUE(CGridColumn, "Text")
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IGridColumn to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {61724103-12B3-11D3-A7FE-0080C8763FA4}
static const IID IID_IGridColumn =
{ 0x61724103, 0x12b3, 0x11d3, { 0xa7, 0xfe, 0x0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };

BEGIN_INTERFACE_MAP(CGridColumn, CCmdTarget)
	INTERFACE_PART(CGridColumn, IID_IGridColumn, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CGridColumn message handlers

BSTR CGridColumn::GetCaption() 
{
	return prop.strCaption.AllocSysString();
}

void CGridColumn::SetCaption(LPCTSTR lpszNewValue) 
{
	prop.strCaption = lpszNewValue;

	if (!m_pCtrl->m_pGridInner)
		return;

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);
	BOOL bLock = m_pCtrl->m_pGridInner->GetParam()->IsLockReadOnly();
	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
	
	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);
	m_pCtrl->m_pGridInner->SetValueRange(CGXRange(0, nCol), prop.strCaption);

	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);
}

long CGridColumn::GetAlignment() 
{
	return prop.nAlignment;
}

void CGridColumn::SetAlignment(long nNewValue) 
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

	if (!m_pCtrl->m_pGridInner)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetHorizontalAlignment(
		dwAlign));

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

OLE_COLOR CGridColumn::GetForeColor() 
{
	if (prop.clrFore == DEFAULTCOLOR)
		return m_pCtrl->GetForeColor();

	return prop.clrFore;
}

void CGridColumn::SetForeColor(OLE_COLOR nNewValue) 
{
	if (prop.clrFore == nNewValue)
		return;

	prop.clrFore = nNewValue;

	if (!m_pCtrl->m_pGridInner || prop.clrFore == m_pCtrl->GetForeColor())
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetTextColor(
		m_pCtrl->TranslateColor(prop.clrFore)));

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

OLE_COLOR CGridColumn::GetBackColor() 
{
	if (prop.clrBack == DEFAULTCOLOR)
		return m_pCtrl->GetBackColor();

	return prop.clrBack;
}

void CGridColumn::SetBackColor(OLE_COLOR nNewValue) 
{
	if (prop.clrBack == nNewValue)
		return;

	prop.clrBack = nNewValue;

	if (!m_pCtrl->m_pGridInner)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetInterior(
		m_pCtrl->TranslateColor(prop.clrBack)));

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

long CGridColumn::GetCase() 
{
	return prop.nCase;
}

void CGridColumn::SetCase(long nNewValue) 
{
	if (nNewValue < 0 || nNewValue > 2)
		return;

	prop.nCase = nNewValue;

	if (!m_pCtrl->m_pGridInner)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetUserAttribute(
		GX_IDS_UA_CASE, prop.nCase));

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

short CGridColumn::GetFieldLen() 
{
	return (short)prop.nFieldLen;
}

void CGridColumn::SetFieldLen(short nNewValue) 
{
	if (nNewValue < 0)
		return;

	prop.nFieldLen = nNewValue;

	if (!m_pCtrl->m_pGridInner)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().
		SetMaxLength((short)prop.nFieldLen));

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

BOOL CGridColumn::GetAllowSizing() 
{
	return prop.bAllowSizing;
}

void CGridColumn::SetAllowSizing(BOOL bNewValue) 
{
	prop.bAllowSizing = bNewValue;
}

OLE_COLOR CGridColumn::GetHeadForeColor() 
{
	if (prop.clrHeadFore == DEFAULTCOLOR)
		return m_pCtrl->m_clrHeadFore;

	return prop.clrHeadFore;
}

void CGridColumn::SetHeadForeColor(OLE_COLOR nNewValue) 
{
	if (prop.clrHeadFore == nNewValue)
		return;

	prop.clrHeadFore = nNewValue;

	if (!m_pCtrl->m_pGridInner || prop.clrHeadFore == m_pCtrl->m_clrHeadFore)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->SetStyleRange(CGXRange(0, nCol), CGXStyle().SetTextColor(
		m_pCtrl->TranslateColor(prop.clrHeadFore)));

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

OLE_COLOR CGridColumn::GetHeadBackColor() 
{
	if (prop.clrHeadBack == DEFAULTCOLOR)
		return m_pCtrl->m_clrHeadBack;

	return prop.clrHeadBack;
}

void CGridColumn::SetHeadBackColor(OLE_COLOR nNewValue) 
{
	if (prop.clrHeadBack == nNewValue)
		return;

	prop.clrHeadBack = nNewValue;

	if (!m_pCtrl->m_pGridInner || prop.clrHeadBack == m_pCtrl->m_clrHeadBack)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->SetStyleRange(CGXRange(0, nCol), CGXStyle().SetInterior(
		m_pCtrl->TranslateColor(prop.clrHeadBack)));

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

BSTR CGridColumn::GetDataField() 
{
	return prop.strDataField.AllocSysString();
}

BOOL CGridColumn::GetLocked() 
{
	return prop.bLocked;
}

void CGridColumn::SetLocked(BOOL bNewValue) 
{
	if (prop.bForceLock && !bNewValue)
		return;

	prop.bLocked = bNewValue;

	if (!m_pCtrl->m_pGridInner)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetReadOnly(prop.bLocked));

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

BOOL CGridColumn::GetVisible() 
{
	return prop.bVisible;
}

void CGridColumn::SetVisible(BOOL bNewValue) 
{
	prop.bVisible = bNewValue;

	if (!m_pCtrl->m_pGridInner)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->HideCols(nCol, nCol, !prop.bVisible);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

long CGridColumn::GetWidth() 
{
	if (!m_pCtrl->m_pGridInner)
		return prop.nWidth;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	prop.nWidth = m_pCtrl->m_pGridInner->GetColWidth(nCol);

	return prop.nWidth;
}

void CGridColumn::SetWidth(long nNewValue) 
{
	if (nNewValue < 0 || nNewValue > 15000)
		m_pCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	prop.nWidth = nNewValue;

	if (!m_pCtrl->m_pGridInner)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->SetColWidth(nCol, nCol, nNewValue);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

long CGridColumn::GetDataType() 
{
	return prop.nDataType;
}

void CGridColumn::SetDataField(LPCTSTR lpszNewValue) 
{
	if (!prop.strDataField.CompareNoCase(lpszNewValue))
		return;

	prop.strDataField = lpszNewValue;

	m_pCtrl->COleControl::SetRedraw(FALSE);

	prop.nDataField = m_pCtrl->GetFieldFromStr(prop.strDataField);
	SetDataType(m_pCtrl->GetDataTypeFromStr(prop.strDataField));

	ROWCOL nCol = m_pCtrl->GetColOrdinalFromIndex(prop.nColIndex);
	m_pCtrl->m_apColumnsProp[nCol]->strDataField = prop.strDataField;
	m_pCtrl->m_apColumnsProp[nCol]->nDataField = prop.nDataField;
	
	if (prop.nDataField != -1)
	{
		prop.bForceNoNullable = !(m_pCtrl->m_pColumnInfo[prop.nDataField].dwFlags & DBCOLUMNFLAGS_ISNULLABLE);
		prop.bForceLock = !(m_pCtrl->m_pColumnInfo[prop.nDataField].dwFlags & DBCOLUMNFLAGS_WRITE);
	}
	else
		prop.bForceNoNullable = prop.bForceLock = FALSE;
	
	if (prop.bForceLock)
		prop.bLocked = TRUE;
	SetLocked(prop.bLocked);

	if (prop.bForceNoNullable)
		prop.bNullable = TRUE;
	SetNullable(prop.bNullable);

	m_pCtrl->COleControl::SetRedraw(TRUE);
	m_pCtrl->Bind();
	m_pCtrl->m_pGridInner->RedrawRowCol(CGXRange().SetCols(nCol + 1));
}

// if the datatype is date, we should use the hidden date editbox style cell
// else if the mask is set, we should use the hidden mask editbox style cell
void CGridColumn::SetDataType(long nNewValue) 
{
	if (nNewValue < 0 || nNewValue > 7)
		return;

	prop.nDataType = nNewValue;
	
	if (!m_pCtrl->m_pGridInner || nNewValue < 0 || !m_pCtrl->AmbientUserMode())
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetUserAttribute(
		GX_IDS_UA_DATATYPE, prop.nDataType));

	if (prop.nDataType == 7)
		m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
			KCOM_IDS_CTRL_DATEEDIT).SetUserAttribute(GX_IDS_UA_INPUTMASK, prop.strMask));
	else
		SetStyle(prop.nStyle);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	SetStyle(prop.nDataType);
}

BOOL CGridColumn::GetSelected() 
{
	CRowColArray awCols;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	ROWCOL nColCount = m_pCtrl->m_pGridInner->GetSelectedCols(awCols);
	if (nColCount == 0)
		return FALSE;

	for (ROWCOL i = 0; i < nColCount; i ++)
		if (awCols[i] == nCol)
			return TRUE;

	return FALSE;
}

void CGridColumn::SetSelected(BOOL bNewValue) 
{
	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->SelectRange(CGXRange().SetCols(nCol), bNewValue);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);
}

long CGridColumn::GetStyle() 
{
	return prop.nStyle;
}

// if the datatype is date, we should use the hidden date editbox style cell
// else if the mask is set, we should use the hidden mask editbox style cell
// so the given style takes no effect
void CGridColumn::SetStyle(long nNewValue)
{
	if (nNewValue < 0 || nNewValue > 4 || !m_pCtrl->AmbientUserMode())
		return;

	prop.nStyle = nNewValue;
	if (prop.nDataType == 7)
		return;

	if (!m_pCtrl->m_pGridInner)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	if (prop.nDataType == 7)
		m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
			KCOM_IDS_CTRL_DATEEDIT));
	else if (prop.strMask.GetLength())
		m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
			GX_IDS_CTRL_MASKEDIT));
	else
	{
		switch (prop.nStyle)
		{
		case 0:
			m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
				GX_IDS_CTRL_EDIT));
			break;
			
		case 1:
			m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
				KCOM_IDS_CTRL_EDITBUTTON));
			break;
			
		case 2:
			m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
				GX_IDS_CTRL_CHECKBOXEX));
			break;
			
		case 3:
			if (prop.nComboBoxStyle == 0)
				m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
					GX_IDS_CTRL_COMBOBOX).SetChoiceList(prop.strChoiceList));
			else
				m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
					GX_IDS_CTRL_COMBOLIST).SetChoiceList(prop.strChoiceList));
			break;
			
		case 4:
			m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
				GX_IDS_CTRL_PUSHBTN));
			break;
		}
	}

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

long CGridColumn::GetCaptionAlignment() 
{
	return prop.nCaptionAlignment;
}

void CGridColumn::SetCaptionAlignment(long nNewValue) 
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

	if (!m_pCtrl->m_pGridInner)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->RowHeaderStyle().SetHorizontalAlignment(dwAlign);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

BOOL CGridColumn::GetColChanged() 
{
	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	return m_pCtrl->m_pGridInner->IsColDirty(nCol);
}

BSTR CGridColumn::GetName() 
{
	return prop.strName.AllocSysString();
}

void CGridColumn::SetName(LPCTSTR lpszNewValue) 
{
	if (lstrlen(lpszNewValue) == 0 || !prop.strName.CompareNoCase(lpszNewValue))
		return;

	for (int i = 0; i < m_pCtrl->m_apColumns.GetSize(); i ++)
		if (!m_pCtrl->m_apColumns[i]->prop.strName.CompareNoCase(
			lpszNewValue))
			return;

	prop.strName = lpszNewValue;
}

BOOL CGridColumn::GetNullable() 
{
	return prop.bNullable;
}

void CGridColumn::SetNullable(BOOL bNewValue) 
{
	if (prop.bForceNoNullable && bNewValue)
		return;

	prop.bNullable = bNewValue;
}

BSTR CGridColumn::GetMask() 
{
	return prop.strMask.AllocSysString();
}

// if the datatype is date, we should use the hidden date editbox style cell
// else if the mask is set, we should use the hidden mask editbox style cell
// so the given style takes no effect
void CGridColumn::SetMask(LPCTSTR lpszNewValue) 
{
	prop.strMask = lpszNewValue;

	if (!m_pCtrl->m_pGridInner || !m_pCtrl->AmbientUserMode())
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	if (prop.nDataType != 7)
		m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
			GX_IDS_CTRL_MASKEDIT).SetUserAttribute(GX_IDS_UA_INPUTMASK, prop.strMask));
	else
		m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetUserAttribute(
		GX_IDS_UA_INPUTMASK, prop.strMask));

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	SetStyle(prop.nStyle);
}

BSTR CGridColumn::GetPromptChar() 
{
	return prop.strPromptChar.AllocSysString();
}

void CGridColumn::SetPromptChar(LPCTSTR lpszNewValue) 
{
	if (lstrlen(lpszNewValue) != 1)
		return;

	prop.strPromptChar = lpszNewValue;

	if (!m_pCtrl->m_pGridInner)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetUserAttribute(
		GX_IDS_UA_INPUTPROMPT, prop.strPromptChar));

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

BOOL CGridColumn::GetDispatchIID(IID *pIID)
{
	*pIID = IID_IGridColumn;

	return TRUE;
}

BSTR CGridColumn::GetText() 
{
	ROWCOL nRow, nCol;
	m_pCtrl->m_pGridInner->GetCurrentCell(nRow, nCol);

	return GetCellText(nRow);
}

void CGridColumn::SetText(LPCTSTR lpszNewValue) 
{
	ROWCOL nRow, nCol;

	m_pCtrl->m_pGridInner->GetCurrentCell(nRow, nCol);
	nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);
	BOOL bLock = m_pCtrl->m_pGridInner->GetParam()->IsLockReadOnly();
	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
	
	m_pCtrl->m_pGridInner->SetControlTextRowCol(nRow, nCol, lpszNewValue);

	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(TRUE);
}

VARIANT CGridColumn::GetValue() 
{
	ROWCOL nRow, nCol;
	m_pCtrl->m_pGridInner->GetCurrentCell(nRow, nCol);

	return GetCellValue(nRow);
}

void CGridColumn::SetValue(const VARIANT FAR& newValue) 
{
	COleVariant vaResult = newValue;
	vaResult.ChangeType(VT_BSTR);
	CString strValue = vaResult.bstrVal;

	SetText(strValue);
}

BSTR CGridColumn::CellText(const VARIANT FAR& Bookmark)
{
	return GetCellText(m_pCtrl->GetRowFromBmk(&Bookmark));
}

VARIANT CGridColumn::CellValue(const VARIANT FAR& Bookmark)
{
	return GetCellValue(m_pCtrl->GetRowFromBmk(&Bookmark));
}

BSTR CGridColumn::GetCellText(long RowIndex) 
{
	if ((ROWCOL)RowIndex > m_pCtrl->m_pGridInner->GetRowCount() || RowIndex
		< 1)
		m_pCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	CString strResult;
	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	CGXStyle style;
	
	if (m_pCtrl->m_pGridInner->GetStyleRowCol(RowIndex, nCol, style))
		m_pCtrl->m_pGridInner->GetControl(RowIndex, nCol)->GetControlText(
		strResult, RowIndex, nCol, NULL, style);

	if (prop.nCase == 1)
		strResult.MakeUpper();
	else if (prop.nCase == 2)
		strResult.MakeLower();

	return strResult.AllocSysString();
}

VARIANT CGridColumn::GetCellValue(long RowIndex) 
{
	COleVariant vaResult;

	if ((ROWCOL)RowIndex > m_pCtrl->m_pGridInner->GetRowCount() || RowIndex
		< 1)
		m_pCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	vaResult = m_pCtrl->m_pGridInner->GetValueRowCol(RowIndex, nCol);
	
	// if the date mask is mingguo format, the value will be text type
	if (prop.nDataType != 7 || prop.strMask.Find('g') == -1)
		vaResult.ChangeType(arrVt[prop.nDataType]);

	return vaResult;
}

BOOL CGridColumn::IsCellValid() 
{
	return TRUE;
}

BOOL CGridColumn::GetPromptInclude() 
{
	return prop.bPromptInclude;
}

void CGridColumn::SetPromptInclude(BOOL bNewValue) 
{
	prop.bPromptInclude = bNewValue;

	if (!m_pCtrl->m_pGridInner)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetUserAttribute(
		GX_IDS_UA_PROMPTINCLUDE, prop.bPromptInclude ? _T("T") : _T("F")));

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	m_pCtrl->InvalidateControl();
}

short CGridColumn::GetListCount() 
{
	return m_arChoiceList.GetSize();
}

short CGridColumn::GetListIndex() 
{
	if (prop.nStyle != 3)
		return CB_ERR;

	ROWCOL nRow, nCol;

	m_pCtrl->m_pGridInner->GetCurrentCell(nRow, nCol);
	if (nRow > m_pCtrl->m_pGridInner->GetRowCount() || nRow
		< 1)
		return CB_ERR;
	nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	CString strValue = m_pCtrl->m_pGridInner->GetValueRowCol(nRow, nCol);

	for (int i = 0; i < m_arChoiceList.GetSize(); i ++)
		if (!m_arChoiceList[i].CompareNoCase(strValue))
			return i;

	return CB_ERR;
}

void CGridColumn::SetListIndex(short nNewValue) 
{
	if (prop.nStyle != 3 || nNewValue < 0 || nNewValue >= m_arChoiceList.
		GetSize())
		return;

	ROWCOL nRow, nCol;

	m_pCtrl->m_pGridInner->GetCurrentCell(nRow, nCol);
	if (nRow > m_pCtrl->m_pGridInner->GetRowCount() || nRow < 1)
		return;
	nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);
	BOOL bLock = m_pCtrl->m_pGridInner->GetParam()->IsLockReadOnly();
	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
	
	m_pCtrl->m_pGridInner->SetValueRange(CGXRange(nRow, nCol), m_arChoiceList[nNewValue]);

	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(TRUE);
}

void CGridColumn::AddItem(LPCTSTR Item, const VARIANT FAR& Index) 
{
	int nIndex;

	if (Index.vt == VT_ERROR)
		nIndex = m_arChoiceList.GetSize();
	else
	{
		COleVariant vaIndex = Index;
		vaIndex.ChangeType(VT_I4);

		if (vaIndex.lVal < 0)
			nIndex = 0;
		else if(vaIndex.lVal > m_arChoiceList.GetSize())
			nIndex = m_arChoiceList.GetSize();
		else
			nIndex = vaIndex.lVal;
	}

	m_arChoiceList.InsertAt(nIndex, Item);

	CalcChoiceList();
}

void CGridColumn::RemoveAll() 
{
	m_arChoiceList.RemoveAll();
	CalcChoiceList();
}

void CGridColumn::CalcChoiceList()
{
	ROWCOL nRow, nCol;

	m_pCtrl->m_pGridInner->GetCurrentCell(nRow, nCol);
	if (nRow > m_pCtrl->m_pGridInner->GetRowCount() || nRow < 1)
		return;
	nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	prop.strChoiceList.Empty();

	int i = m_arChoiceList.GetSize(), j;
	if (i > 0)
	{
		prop.strChoiceList += m_arChoiceList[0];
		for (j = 1; j < i; j ++)
		{
			prop.strChoiceList += '\n';
			prop.strChoiceList += m_arChoiceList[j];
		}
	}

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);
	BOOL bLock = m_pCtrl->m_pGridInner->GetParam()->IsLockReadOnly();
	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
	
	m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().
		SetChoiceList(prop.strChoiceList));

	m_pCtrl->m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(TRUE);
}

void CGridColumn::RemoveItem(short Index) 
{
	if (Index < 0 || Index >= m_arChoiceList.GetSize())
		return;

	m_arChoiceList.RemoveAt(Index);

	CalcChoiceList();
}

BSTR CGridColumn::GetList(short Index) 
{
	CString strResult;
	
	if (Index >= 0 && Index < m_arChoiceList.GetSize())
		strResult = m_arChoiceList[Index];

	return strResult.AllocSysString();
}

void CGridColumn::SetList(short Index, LPCTSTR lpszNewValue) 
{
	if (Index < 0 || Index >= m_arChoiceList.GetSize())
		return;
	
	m_arChoiceList[Index] = lpszNewValue;

	CalcChoiceList();
}

void CGridColumn::OnDropDownWndChanged() 
{
	if (!m_dropDownWnd)
		// if external dropdown is removed, mark it
		m_pCtrl->m_bDroppedDown = FALSE;
}

long CGridColumn::GetComboBoxStyle() 
{
	return prop.nComboBoxStyle;
}

void CGridColumn::SetComboBoxStyle(long nNewValue) 
{
	if (nNewValue < 1 || nNewValue > 2 || !m_pCtrl->AmbientUserMode())
		return;

	prop.nComboBoxStyle = nNewValue;
	if (prop.nDataType == 7)
		return;

	if (!m_pCtrl->m_pGridInner)
		return;

	ROWCOL nCol = m_pCtrl->m_pGridInner->GetColFromIndex(prop.nColIndex);

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);

	if (prop.nDataType != 7 && prop.strMask.GetLength() == 0 && prop.nStyle == 3)
	{
		if (prop.nComboBoxStyle == 0)
			m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
				GX_IDS_CTRL_COMBOBOX));
		else
			m_pCtrl->m_pGridInner->SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
				GX_IDS_CTRL_COMBOLIST));
	}

	m_pCtrl->m_pGridInner->GetParam()->EnableUndo(FALSE);
}
