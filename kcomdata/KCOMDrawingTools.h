// KCOMDrawingTools.h: interface for the CKCOMDrawingTools class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_KCOMDRAWINGTOOLS_H__47C76D84_179E_11D3_B446_0080C8F18522__INCLUDED_)
#define AFX_KCOMDRAWINGTOOLS_H__47C76D84_179E_11D3_B446_0080C8F18522__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// do custom drawing static cells works, including text and picture

class CKCOMRichGridCtrl;

class CKCOMDrawingTools : public CGXDrawingAndFormatting
{
public:
	virtual void DrawStatic(CGXControl* pControl, CDC* pDC, CRect rect, ROWCOL nRow, ROWCOL nCol, const CGXStyle& style, const CGXStyle* pStandardStyle);
	CKCOMDrawingTools();

	virtual ~CKCOMDrawingTools();

	CKCOMRichGridCtrl * m_pCtrl;

};

#endif // !defined(AFX_KCOMDRAWINGTOOLS_H__47C76D84_179E_11D3_B446_0080C8F18522__INCLUDED_)
