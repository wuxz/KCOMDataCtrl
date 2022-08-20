// KCOMMaskControl.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "KCOMMaskControl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskControl

CKCOMMaskControl::CKCOMMaskControl(CGXGridCore* pGrid, UINT nID, CRuntimeClass* pDataClass) :
	CGXMaskControl(pGrid, nID, pDataClass)
{
}

CKCOMMaskControl::~CKCOMMaskControl()
{
}


BEGIN_MESSAGE_MAP(CKCOMMaskControl, CGXMaskControl)
	//{{AFX_MSG_MAP(CKCOMMaskControl)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskControl message handlers

void CKCOMMaskControl::GetCurrentText(CString &sMaskedResult)
{
	BOOL bPromptInclude = IsPromptInclude();

	if (IsInit())
	{
		if (bPromptInclude)
			m_pMask->GetDisplayString(sMaskedResult);
		else
			m_pMask->GetData(sMaskedResult);
	}
	else
	{
		TRACE0("Warning: GetCurrentText was called for an unitialized control!\n");

		// Control is not initialized - Unable to determine text
		sMaskedResult.Empty();
	}
}

BOOL CKCOMMaskControl::GetValue(CString &strResult)
{
	BOOL bPromptInclude = IsPromptInclude();

	// Reads out the window text and converts it into
	// the plain value which should be stored in the style object.

	if (!IsInit())
		return FALSE;

	if (bPromptInclude)
		m_pMask->GetDisplayString(strResult);
	else
		m_pMask->GetData(strResult);

	return TRUE;
}

void CKCOMMaskControl::SetValue(LPCTSTR pszRawValue)
{
	BOOL bPromptInclude = IsPromptInclude();

	// Convert the value to the text which should be
	// displayed in the current cell and show this
	// text. (don't change the modify flag and do not
	// call OnModifyCell!)

	if (m_hWnd)
	{
		NeedStyle();
		InitMask(m_pMask, *m_pStyle);
		
		if (bPromptInclude)
			m_pMask->SetDisplayString(pszRawValue);
		else
			m_pMask->SetData(pszRawValue);

		CString sText;
		m_pMask->GetDisplayString(sText);
		SetWindowText(sText);
	}
}

BOOL CKCOMMaskControl::GetControlText(CString &strDisplay, ROWCOL nRow, ROWCOL nCol, LPCTSTR pszRawValue, const CGXStyle &style)
{
	BOOL bPromptInclude = IsPromptInclude();

	InitMask(m_pExtMask, style);

	if (pszRawValue == NULL)
		pszRawValue = style.GetValueRef();

	if (m_bDisplayEmpty || _tcslen(pszRawValue) > 0)
	{
		if (bPromptInclude)
		{
			if (m_pExtMask->SetDisplayString(pszRawValue) == -1)
			{
				TRACE(_T("Warning: %s does not fit to mask\n"),
					pszRawValue ? pszRawValue : (LPCTSTR) style.GetValueRef());
				return FALSE;
			}

			m_pExtMask->GetDisplayString(strDisplay);
		}
		else
		{
			if (m_pExtMask->SetData(pszRawValue) == -1)
			{
				TRACE(_T("Warning: %s does not fit to mask\n"),
					pszRawValue ? pszRawValue : (LPCTSTR) style.GetValueRef());
				return FALSE;
			}

			m_pExtMask->GetData(strDisplay);
		}

	}
	else
		strDisplay.Empty();

	return TRUE;
}

BOOL CKCOMMaskControl::IsPromptInclude()
{
	if (m_nRow > Grid()->GetRowCount() || m_nCol > Grid()->GetColCount())
		return TRUE;

	NeedStyle();

	CString strUserAttrib;
	if (m_pStyle->GetIncludeUserAttribute(GX_IDS_UA_PROMPTINCLUDE))
	{
		m_pStyle->GetUserAttribute(GX_IDS_UA_PROMPTINCLUDE, strUserAttrib);
		if (strUserAttrib == _T("F"))
			return FALSE;
	}

	return TRUE;
}

BOOL CKCOMMaskControl::ConvertControlTextToValue(CString &str, ROWCOL nRow, ROWCOL nCol, const CGXStyle *pOldStyle)
{
	CGXStyle* pStyle = NULL;
	BOOL bSuccess = FALSE;

	if (pOldStyle == NULL)
	{
		pStyle = Grid()->CreateStyle();
		Grid()->ComposeStyleRowCol(nRow, nCol, pStyle);
		pOldStyle = pStyle;
	}

	// allow only valid input
	{
		InitMask(m_pExtMask, *pOldStyle);

		BOOL bPromptInclude = IsPromptInclude();

		if (bPromptInclude)
			m_pExtMask->SetDisplayString(str);
		else
			m_pExtMask->SetData(str);

		if (!m_pExtMask->IsValidDisplayString())
		{
			// Display string does not match correcttly, try data only
			if (m_pExtMask->SetData(str) == -1)
				goto exitcv;
		}

		if (bPromptInclude)
			m_pExtMask->GetDisplayString(str);
		else
			m_pExtMask->GetData(str);

		bSuccess = TRUE;
	}

exitcv:
	if (pStyle)
		Grid()->RecycleStyle(pStyle);

	return bSuccess;
}
