// KCOMDrawingTools.cpp: implementation of the CKCOMDrawingTools class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "kcomdata.h"
#include "KCOMDrawingTools.h"
#include "KCOMRichGridCtl.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CKCOMDrawingTools::CKCOMDrawingTools()
{
	m_pCtrl = NULL;
}

CKCOMDrawingTools::~CKCOMDrawingTools()
{

}

void CKCOMDrawingTools::DrawStatic(CGXControl *pControl, CDC *pDC, CRect rect, ROWCOL nRow, ROWCOL nCol, const CGXStyle &style, const CGXStyle *pStandardStyle)
{
/*	CGXGridCore* pGrid = pControl->Grid();
	CGXGridParam* pParam = pGrid->GetParam();

	CString sOutput;

	if (pParam->m_nDisplayExpression & GX_EXPR_DISPLAYALWAYS)
		sOutput = pGrid->GetExpressionRowCol(nRow, nCol);
	else
		pControl->GetControlText(sOutput, nRow, nCol, NULL, style);

//	if (sOutput.IsEmpty())
//	{
//		pControl->CGXControl::Draw(pDC, rect, nRow, nCol, style, pStandardStyle);
//		return;
//	}

	if (rect.right <= rect.left || rect.Width() <= 1 || rect.Height() <= 1)
		return;

	// ------------------ Static Text ---------------------------

	CString str;
	HANDLE hbm = 0;
	
	// check for #bmp only for static cells (and not 
	// for cells where the user can enter the text)

	// the "case" property of the column object
	CString strUserAttrib;
	if (style.GetIncludeUserAttribute(GX_IDS_UA_CASE))
	{
		style.GetUserAttribute(GX_IDS_UA_CASE, strUserAttrib);

		switch (atoi(strUserAttrib))
		{
		case 1:
			sOutput.MakeUpper();
			break;

		case 2:
			sOutput.MakeLower();
			break;
		}
	}
	
	if (!pControl->IsKindOf(RUNTIME_CLASS(CEdit)))
		hbm = GXGetDIBState()->GetPool()->LookupOrCreateDIB(sOutput);

	// clone the constant style object to modify it
	CGXStyle styleClone(style);

	if (hbm)
	{
		pControl->DrawBackground(pDC, rect, styleClone);

		// Draw bitmap

		// Don't subtract margins - only borders and 3d-frame
		// Therefore we call CGXControl::AddBorders
		CRect r = pControl->CGXControl::GetCellRect(nRow, nCol, rect, &styleClone);

		DWORD dtAlign = styleClone.GetHorizontalAlignment() | styleClone.GetVerticalAlignment();

		CGXDIB::Draw(pDC, hbm, r, CGXDIB::none, (UINT) dtAlign);

		// child Controls: spin-buttons, hotspot, combobox btn, ...
		pControl->CGXControl::Draw(pDC, rect, nRow, nCol, styleClone, pStandardStyle);
	}
	else
	{
		// Draw text

		// Select font
		CFont* pOldFont = pControl->LoadFont(pDC, styleClone, pStandardStyle);

		// Rectangle
		CRect rectItem = rect;

		// subtract margins and borders 
		CRect rectText = pControl->GetCellRect(nRow, nCol, rect);

		// Cell-Color
		COLORREF rgbText = styleClone.GetTextColor();
		COLORREF rgbCell = styleClone.GetInteriorRef().GetColor();

		if (pParam->GetProperties()->GetBlackWhite())
		{
			rgbText = RGB(0,0,0);
			rgbCell = RGB(255,255,255);
		}
		
		DWORD dtAlign = styleClone.GetHorizontalAlignment();

		HBITMAP hBmp = NULL;
		short bDrawText = 1;

		// only the grid control need firing this event to draw picture
		if (m_pCtrl && style.GetIncludeUserAttribute(IDS_UA_GRIDTYPE))
		{
			VARIANT vaPic;
			VariantInit(&vaPic);
			OLE_COLOR clrBack = (OLE_COLOR)(styleClone.GetInterior().GetColor());

			m_pCtrl->FireBeforeDrawCell(nRow, nCol, &vaPic, &clrBack, &bDrawText);

			styleClone.SetInterior(m_pCtrl->TranslateColor(clrBack));

			if (vaPic.vt == VT_I4)
				hBmp = (HBITMAP)vaPic.lVal;
		}

		pControl->DrawBackground(pDC, rect, styleClone, TRUE);

		if (hBmp)
		{
			if (!bDrawText)
				sOutput.Empty();

			BITMAPINFO bmpinfo;
			bmpinfo.bmiHeader.biBitCount = 0;
			bmpinfo.bmiHeader.biSize = sizeof(bmpinfo.bmiHeader);
			
			GetDIBits(pDC->m_hDC, hBmp, 0, 0, NULL, &bmpinfo, DIB_RGB_COLORS);
			
			CSize sz;
			sz = pDC->GetOutputTextExtent(sOutput);
			
			HDC memdc =	CreateCompatibleDC(pDC->m_hDC);
			HBITMAP hBmpOld = (HBITMAP)SelectObject(memdc, hBmp);
			
			int nShiftX, nShiftY;
			
			if (bmpinfo.bmiHeader.biHeight < rectText.Height())
				nShiftY = (rectText.Height() - bmpinfo.bmiHeader.biHeight) / 2;
			else
				nShiftY = 0;

			// the cell width is big enough to hold both the text and picture
			if (sz.cx + bmpinfo.bmiHeader.biWidth < rectText.Width() && sOutput.Find(_T("\n"), 0) == -1)
			{
				switch (dtAlign)
				{
				case DT_CENTER:
					nShiftX = (rectText.Width() - sz.cx - bmpinfo.bmiHeader.biWidth) / 2;
					BitBlt(pDC->m_hDC, rectText.left + nShiftX, rectText.top + nShiftY, bmpinfo.bmiHeader.biWidth,
						bmpinfo.bmiHeader.biHeight, memdc, 0, 0, SRCCOPY);
					rectText.left += nShiftX + bmpinfo.bmiHeader.biWidth;
					dtAlign = DT_LEFT;

					break;

				case DT_LEFT:
					BitBlt(pDC->m_hDC, rectText.left, rectText.top + nShiftY, bmpinfo.bmiHeader.biWidth,
						bmpinfo.bmiHeader.biHeight, memdc, 0, 0, SRCCOPY);
					rectText.left += bmpinfo.bmiHeader.biWidth;

					break;

				case DT_RIGHT:
					BitBlt(pDC->m_hDC, rectText.right - bmpinfo.bmiHeader.biWidth, rectText.top + nShiftY, 
						bmpinfo.bmiHeader.biWidth, bmpinfo.bmiHeader.biHeight, memdc, 0, 0, SRCCOPY);
					rectText.right -= bmpinfo.bmiHeader.biWidth;
					
					break;
				}
			}
			// the cell width is not big enough to hold both the text and picture
			// so put the picture at left unless the alighment is right
			else
			{
				int nWidth = min(rectText.Width(), bmpinfo.bmiHeader.biWidth);
				
				switch (dtAlign)
				{
				case DT_CENTER:
				case DT_LEFT:
					BitBlt(pDC->m_hDC, rectText.left, rectText.top + nShiftY, nWidth, bmpinfo.bmiHeader.biHeight, 
						memdc, 0, 0, SRCCOPY);
					rectText.left += nWidth;
					dtAlign = DT_LEFT;

					break;

				case DT_RIGHT:
					BitBlt(pDC->m_hDC, rectText.right - nWidth, rectText.top + nShiftY, nWidth, bmpinfo.bmiHeader.biHeight, 
						memdc, 0, 0, SRCCOPY);
					rectText.right -= nWidth;
					
					break;
				}
			}
			
			SelectObject(memdc, hBmpOld);
			DeleteDC(memdc);
		}
		
		if (_tcslen(sOutput) > 0)
		{
			pDC->SetBkMode(TRANSPARENT);
			pDC->SetBkColor(rgbCell);
			pDC->SetTextColor(rgbText);

			dtAlign |= styleClone.GetVerticalAlignment();

			if (styleClone.GetWrapText())
				dtAlign |= DT_NOPREFIX | DT_WORDBREAK;
			else
				dtAlign |= DT_NOPREFIX | DT_SINGLELINE;

			if (styleClone.GetIncludeFont() 
				&& styleClone.GetFontRef().GetIncludeOrientation() 
				&& styleClone.GetFontRef().GetOrientation() != 0)
			{
				// draw vertical text
				GXDaFTools()->GXDrawRotatedText(pDC, 
					sOutput,
					sOutput.GetLength(),
					rectText,
					(UINT) dtAlign,
					styleClone.GetFontRef().GetOrientation(), 
					&rect);
			}
			else
			{
				// check if ellipses are to be used
				if(!styleClone.GetIncludeUserAttribute(GX_IDS_UA_ELLIPSISTYPE))
					GXDrawTextLikeMultiLineEdit(pDC,
					sOutput,
					sOutput.GetLength(),
					rectText,
					(UINT) dtAlign,
					&rect);
				else
						GXDrawTextLikeMultiLineEditEx(pDC,
					sOutput,
					sOutput.GetLength(),
					rectText,
					(UINT) dtAlign,
					&rect, ( (CGXEllipseUserAttribute&)styleClone.GetUserAttribute(GX_IDS_UA_ELLIPSISTYPE)).GetEllipseValue());
			}
		}

		// child Controls: spin-buttons, hotspot, combobox btn, ...
		pControl->CGXControl::Draw(pDC, rect, nRow, nCol, styleClone, pStandardStyle);

		if (pOldFont)
			pDC->SelectObject(pOldFont);
	}
*/}
