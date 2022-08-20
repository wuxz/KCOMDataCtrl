// KCOMEditBtn.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "KCOMEditBtn.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define ComboBtnWidth 15
#define _nComboBtnHeight 18

/////////////////////////////////////////////////////////////////////////////
// CKCOMEditBtn

IMPLEMENT_CONTROL(CKCOMEditBtn, CGXEditControl);
IMPLEMENT_DYNAMIC(CKCOMEditBtn, CGXEditControl);

CKCOMEditBtn::CKCOMEditBtn(CGXGridCore* pGrid, UINT nEditID) 
	: CGXEditControl(pGrid, nEditID)
{
	AddChild(m_pButton = new CGXComboBoxButton(this));
	m_nDefaultHeight = 0;
	m_nExtraFrame = 1;
	m_bSizeToContent = TRUE;        // set this TRUE if combobox button shall only be as large as
	m_bInactiveDrawAllCell = FALSE; // draw text over pushbutton when inactive
	m_bInactiveDrawButton = FALSE;  // draw pushbutton when inactive
}

CKCOMEditBtn::~CKCOMEditBtn()
{
}


BEGIN_MESSAGE_MAP(CKCOMEditBtn, CGXEditControl)
	//{{AFX_MSG_MAP(CKCOMEditBtn)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CKCOMEditBtn message handlers

CRect CKCOMEditBtn::GetCellRect(ROWCOL nRow, ROWCOL nCol, LPRECT rectItem, const CGXStyle* pStyle)
{
	// compute the interior rectangle for the text
	// without buttons and borders

	CRect rect = CGXEditControl::GetCellRect(nRow, nCol, rectItem, pStyle);

	if (!m_bInactiveDrawAllCell
		|| m_bInactiveDrawButton
		|| Grid()->IsCurrentCell(nRow, nCol))
		rect.right -= ComboBtnWidth;

	return rect;
}

CSize CKCOMEditBtn::AddBorders(CSize size, const CGXStyle &style)
{
	size.cx += ComboBtnWidth;

	return CGXEditControl::AddBorders(size, style);
}

void CKCOMEditBtn::Init(ROWCOL nRow, ROWCOL nCol)
{
	CGXEditControl::Init(nRow, nCol);

	// Force drawing of buttons for current cell
	GridWnd()->InvalidateRect(m_pButton->GetRect());
}

CSize CKCOMEditBtn::CalcSize(CDC *pDC, ROWCOL nRow, ROWCOL nCol, const CGXStyle &style, const CGXStyle *pStandardStyle, BOOL bVert)
{
	// Size of edit portion
	CSize sizeEdit = CGXEditControl::CalcSize(pDC, nRow, nCol, style, pStandardStyle, bVert);

	// Determine size of listbox entries
	CSize sizeCombo(0, 0);  

	// Select font
	CFont* pOldFont = LoadFont(pDC, style, pStandardStyle);

	if (bVert)
	{
		// row height
		TEXTMETRIC    tm;
		pDC->GetTextMetrics(&tm);
		sizeCombo.cy = tm.tmHeight + tm.tmExternalLeading + 1;
	}

	if (pOldFont)
		pDC->SelectObject(pOldFont);

	// Add space for borders, left and right margin, 3d frame etc.
	sizeCombo = AddBorders(sizeCombo, style);

	// Return the size which fits for both listbox entries and edit portion
	return CSize(max(sizeCombo.cx, sizeEdit.cx), max(sizeCombo.cy, sizeEdit.cy));
}

void CKCOMEditBtn::OnInitChildren(ROWCOL nRow, ROWCOL nCol, const CRect &rect)
{
	if (m_pButton)
	{
		int dy = rect.Height();

		// initialize combobox button only if
		// not the current cell
		CRect rectBtn(max(rect.left + 1, rect.right - 1 - ComboBtnWidth),
			rect.top,
			rect.right - 1,
			rect.top + dy - 1
			);

		// Determine the default height for the combobox edit part
		const CGXStyle* pStyle = NULL;
		if (IsInit() && m_nRow == nRow && m_nCol == nCol)
		{
			NeedStyle();
			pStyle = m_pStyle;
		}
		else if (m_bInactiveDrawButton)
			pStyle = &Grid()->LookupStyleRowCol(nRow, nCol);

		if (pStyle)
		{
			if (m_nDefaultHeight == 0)
			{
				// Get font metrics
				TEXTMETRIC tm;
				CClientDC dc(GridWnd());
				Grid()->OnGridPrepareDC(&dc);
				CFont* pOldFont = Grid()->LoadFont(&dc, *pStyle);
				dc.GetTextMetrics(&tm);
				if (pOldFont)
					dc.SelectObject(pOldFont);

				m_nDefaultHeight = tm.tmHeight;
			}
		}

		if (m_bSizeToContent)
		{
			if (pStyle)
			{
				dy = min(m_nDefaultHeight+2, rect.Height() - 2*m_nExtraFrame);

				if (pStyle->GetIncludeVerticalAlignment())
				{
					// vertical alignment for single line text
					// is done by computing the window rectangle for the CComboBox
					switch (pStyle->GetVerticalAlignment())
					{
					case DT_VCENTER:
						rectBtn.top += (rect.Height() - dy) / 2 - m_nExtraFrame;
						break;
					case DT_BOTTOM:
						rectBtn.top = rect.bottom - dy - 2 * m_nExtraFrame;
					break;
					}
				}
			}

			rectBtn.bottom = rectBtn.top + dy + 2*m_nExtraFrame;
			rectBtn.bottom--;
		}

		m_pButton->SetRect(rectBtn);
	}
}

BOOL CKCOMEditBtn::GetValue(CString &strResult)
{
	if (!CGXEditControl::GetValue(strResult))
		return FALSE;

	NeedStyle();

	CString strUserAttrib;
	COleCurrency cur;
	if (m_pStyle->GetIncludeUserAttribute(GX_IDS_UA_DATATYPE))
	{
		m_pStyle->GetUserAttribute(GX_IDS_UA_DATATYPE, strUserAttrib);

		switch (atoi(strUserAttrib))
		{
		case 1:
			strResult = (atoi(strResult) == 0) ? _T("0") : _T("-1");
			break;

		case 2:
			strResult.Format("%d", (char)atoi(strResult));
			break;

		case 3:
			strResult.Format("%d", (short)atoi(strResult));
			break;

		case 4:
			strResult.Format("%d", atoi(strResult));
			break;

		case 5:
			strResult.Format("%f", atof(strResult));
			break;

		case 6:
			cur.ParseCurrency(strResult);
			if (cur.GetStatus() == COleCurrency::valid)
				strResult = cur.Format();
			else
				strResult.Empty();

			break;
		}
	}
	
	return TRUE;
}
