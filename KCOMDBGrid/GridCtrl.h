#if !defined(AFX_GRIDCTRL_H__519FA702_722C_11D1_ABBA_00A0243D1382__INCLUDED_)
#define AFX_GRIDCTRL_H__519FA702_722C_11D1_ABBA_00A0243D1382__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

/////////////////////////////////////////////////////////////////////////////
// GridCtrl.h : header file
//
// MFC Grid Control header file
//
// Written by Chris Maunder (chrismaunder@codeguru.com)
// Copyright (c) 1998.
//
// The code contained in this file is based on the original
// WorldCom Grid control written by Joe Willcoxson,
//        E-mail:  chinajoe@aol.com
//        URL:  http://users.aol.com/chinajoe
//
// This code may be used in compiled form in any way you desire. This
// file may be redistributed unmodified by any means PROVIDING it is 
// not sold for profit without the authors written consent, and 
// providing that this notice and the authors name and all copyright 
// notices remains intact. If the source code in this file is used in 
// any  commercial application then a statement along the lines of 
// "Portions copyright (c) Chris Maunder, 1998" must be included in 
// the startup banner, "About" box or printed documentation. An email 
// letting me know that you are using it would be nice as well. That's 
// not much to ask considering the amount of work that went into this.
//
// This file is provided "as is" with no expressed or implied warranty.
// The author accepts no liability for any damage/loss of business that
// this product may cause.
//
// Expect bugs!
// 
// Please use and enjoy. Please let me know of any bugs/mods/improvements 
// that you have found/implemented and I will fix/incorporate them into this
// file. 
//
/////////////////////////////////////////////////////////////////////////////

#include "CellRange.h"
#include <afxtempl.h>

// Use C++ exception handling instead of structured.
#undef TRY
#undef CATCH
#undef END_CATCH
#define TRY try
#define CATCH(ex_class, ex_object) catch(ex_class* ex_object)
#define END_CATCH

// Use this as the classname when inserting this control as a custom control
// in the MSVC++ dialog editor
#define GRIDCTRL_CLASSNAME _T("MFCGridCtrl")

#define IsSHIFTpressed() ( (GetKeyState(VK_SHIFT) & (1 << (sizeof(SHORT)*8-1))) != 0   )
#define IsCTRLpressed()  ( (GetKeyState(VK_CONTROL) & (1 << (sizeof(SHORT)*8-1))) != 0 )

// Used for Get/SetItem calls.
typedef struct _GV_ITEM { 
    int      row,col;     // Row and Column of item
    UINT     mask;        // Mask for use in getting/setting cell data
    UINT     state;       // cell state (focus/hilighted etc)
    CString  szText;      // Text in cell
    int      iImage;      // index of the list view item’s icon
    LPARAM   lParam;      // 32-bit value to associate with item 
} GV_ITEM; 

// Grid line selection
#define GVL_NONE                0
#define GVL_HORZ                1
#define GVL_VERT                2
#define GVL_BOTH                3

// Cell data mask
#define GVIF_TEXT               LVIF_TEXT
#define GVIF_IMAGE              LVIF_IMAGE
#define GVIF_STATE              LVIF_STATE

// Cell states
#define GVIS_FOCUSED            0x0001
#define GVIS_SELECTED           0x0002
#define GVIS_FIXED              0x0004  // not yet used

// Cell Searching options
#define GVNI_FOCUSED            0x0001
#define GVNI_SELECTED           0x0002
#define GVNI_FIXED              0x0004  // not yet used

#define GVNI_ABOVE              LVNI_ABOVE
#define GVNI_BELOW              LVNI_BELOW
#define GVNI_TOLEFT             LVNI_TOLEFT
#define GVNI_TORIGHT            LVNI_TORIGHT
#define GVNI_ALL                (LVNI_BELOW|LVNI_TORIGHT|LVNI_TOLEFT)
#define GVNI_AREA               (LVNI_BELOW|LVNI_TORIGHT)

// Hit test values (not yet implemented)
#define GVHT_DATA               0x0000
#define GVHT_TOPLEFT            0x0001
#define GVHT_COLHDR             0x0002
#define GVHT_ROWHDR             0x0004
#define GVHT_COLSIZER           0x0008
#define GVHT_ROWSIZER           0x0010
#define GVHT_LEFT               0x0020
#define GVHT_RIGHT              0x0040
#define GVHT_ABOVE              0x0080
#define GVHT_BELOW              0x0100

typedef struct tagGV_DISPINFO { 
    NMHDR   hdr; 
    GV_ITEM item; 
} GV_DISPINFO;

// Messages sent to the grid's parent (More will be added in future)
#define	GVN_BEGINLABELEDIT      LVN_BEGINLABELEDIT   // LVN_FIRST-5
#define GVN_COLUMNCLICK         LVN_COLUMNCLICK
#define	GVN_DELETEITEM          LVN_DELETEITEM
#define GVN_ENDLABELEDIT        LVN_ENDLABELEDIT     // LVN_FIRST-6
#define GVN_SELCHANGING         LVN_ITEMCHANGING 
#define GVN_SELCHANGED          LVN_ITEMCHANGED



// Each cell contains one of these. Fields "row" and "column" are not stored since we
// will usually have acces to them in other ways, and they are an extra 8 bytes per
// cell that is probably unnecessary.
class CInPlaceEdit;

class CGridCell : public CObject
{
public:
    CGridCell() 
    {
        state = 0;
        szText.Empty();
        iImage = -1;
    }

    UINT     state;       // Cell state (selected/focus etc)
    CString  szText;      // Cell text (or binary data if you wish...)
    int      iImage;      // Index of the list view item’s icon 
}; 


// storage typedef for each row in the grid
typedef CTypedPtrArray<CObArray, CGridCell*> GRID_ROW;

/////////////////////////////////////////////////////////////////////////////
// CGridCtrl window

class CGridCtrl : public CWnd
{
    DECLARE_DYNCREATE(CGridCtrl)

// Construction
public:
    CGridCtrl(int nRows = 0, int nCols = 0, int nFixedRows = 0, int nFixedCols = 0);

    BOOL Create(const RECT& rect, CWnd* parent, UINT nID,
                DWORD dwStyle = WS_CHILD | WS_BORDER | WS_TABSTOP | WS_VISIBLE);

// Attributes
public:
    int  GetRowCount() const                    { return m_nRows; }
    int  GetColCount() const                 { return m_nCols; }
    int  GetFixedRowCount() const               { return m_nFixedRows; }
    int  GetFixedColCount() const            { return m_nFixedCols; }
    BOOL SetRowCount(int nRows = 10);
    BOOL SetColCount(int nCols = 10);
    BOOL SetFixedRowCount(int nFixedRows = 1);
    BOOL SetFixedColCount(int nFixedCols = 1);

    int  GetColWidth(int nCol) const;
    BOOL SetColWidth(int col, int width);
	int GetFixedColWidth() const;
	int GetRowHeight(int nRow) const;
	BOOL SetRowHeight(int nRow, int height);

    void     SetTextColor(COLORREF clr)           { m_crTextColour = clr;             }
    COLORREF GetTextColor() const                 { return m_crTextColour;            }
    void     SetTextBkColor(COLORREF clr)         { m_crTextBkColour = clr;           }
    COLORREF GetTextBkColor() const               { return m_crTextBkColour;          }
    void     SetBkColor(COLORREF clr)             { m_crBkColour = clr;               }
    COLORREF GetBkColor() const                   { return m_crBkColour;              }
    void     SetFixedTextColor(COLORREF clr)      { m_crFixedTextColour = clr;        }
    COLORREF GetFixedTextColor() const            { return m_crFixedTextColour;       }
    void     SetFixedBkColor(COLORREF clr)        { m_crFixedBkColour = clr;          }
    COLORREF GetFixedBkColor() const              { return m_crFixedBkColour;         } 
    void     SetGridColor(COLORREF clr)           { m_crGridColour = clr;             }
    COLORREF GetGridColor() const                 { return m_crGridColour;            }

    CCellID GetFocusCell() const                  { return m_idCurrentCell;           }

    void SetImageList(CImageList* pList)          { m_pImageList = pList;             }
    CImageList* GetImageList() const              { return m_pImageList;              }

    void SetGridLines(int nWhichLines = GVL_BOTH) { m_nGridLines = nWhichLines; 
                                                    if (::IsWindow(GetSafeHwnd())) Invalidate(); }
    int  GetGridLines() const                     { return m_nGridLines;              }

    void SetEditable(BOOL bEditable = TRUE)       { m_bEditable = bEditable;          }
    BOOL IsEditable() const                       { return m_bEditable;               }
    void SetModified(BOOL bModified = TRUE)       { m_bModified = bModified;          }
    BOOL GetModified() const                      { return m_bModified;               }
    void SetListMode(BOOL bEnableListMode = TRUE) { m_bListMode = bEnableListMode;    }
    BOOL GetListMode() const                      { return m_bListMode;               }
    void EnableSelection(BOOL bEnable = TRUE)     { ResetSelectedRange(); m_bEnableSelection = bEnable; ResetSelectedRange(); }
    BOOL IsSelectable() const                     { return m_bEnableSelection;        }
    void SetRowResize(BOOL bResize = TRUE)        { m_bAllowRowResize = bResize;      }
    BOOL GetRowResize() const                     { return m_bAllowRowResize;         }
    void SetColResize(BOOL bResize = TRUE)     { m_bAllowColResize = bResize;   }
    BOOL GetColResize() const                  { return m_bAllowColResize;      }
    void SetDoubleBuffering(BOOL bBuffer = TRUE)  { m_bDoubleBuffer = bBuffer;        }
    BOOL GetDoubleBuffering() const               { return m_bDoubleBuffer;           }

    BOOL IsCellFixed(int nRow, int nCol);

    BOOL   SetItemText(int nRow, int nCol, LPCTSTR str);
    virtual CString GetItemText(int nRow, int nCol);
    int InsertCol(LPCTSTR strHeading, UINT nFormat = DT_CENTER|DT_VCENTER|DT_SINGLELINE,
		int nCol = -1);
    int  InsertRow(int nRow = -1);
    BOOL DeleteRow(int nRow);

// Operations
public:
    BOOL DeleteCol(int nCol);
    BOOL DeleteNonFixedRows();
    BOOL DeleteAllItems();
    BOOL   SetItemImage(int nRow, int nCol, int iImage);
    int    GetItemImage(int nRow, int nCol) const;
    BOOL   SetItemState(int nRow, int nCol, UINT state);
    UINT   GetItemState(int nRow, int nCol) const;
    BOOL   SetColFormat(int nCol, UINT nFormat);
    UINT   GetColFormat(int nCol) const;

    BOOL AutoSizeRow(int nRow);
    BOOL AutoSizeCol(int nCol);
    void AutoSizeRows();
    void AutoSizeCols();
    void AutoSize();
    void ExpandColsToFit();
    void ExpandRowsToFit();
    void ExpandToFit();

    // SetRedraw stops/starts redraws on things like changing the # rows/columns
    // and autosizing, but not for user-intervention such as resizes
    void SetRedraw(BOOL bAllowDraw, BOOL bResetScrollBars = FALSE);
    BOOL RedrawCell(int nRow, int nCol, CDC* pDC = NULL);
    BOOL RedrawCell(const CCellID& cell, CDC* pDC = NULL);
    BOOL RedrawRow(int row);
    BOOL RedrawCol(int col);

    void Print(); 
    BOOL Save(CArchive & ar);
    BOOL Load(CArchive & ar);

    void SetSelectedRange(int nMinRow, int nMinCol, int nMaxRow, int nMaxCol,
                          BOOL bForceRepaint = FALSE);
    BOOL IsValid(int nRow, int nCol) const;

    CCellID GetNextItem(CCellID& cell, int nFlags) const;

// Overrides
    // ClassWizard generated virtual function overrides
    //{{AFX_VIRTUAL(CGridCtrl)
    protected:
    virtual void PreSubclassWindow();
    //}}AFX_VIRTUAL
public:
    virtual void OnBeginPrinting(CDC *pDC, CPrintInfo *pInfo);
    virtual void OnPrint(CDC *pDC, CPrintInfo *pInfo);
    virtual void OnEndPrinting(CDC *pDC, CPrintInfo *pInfo);

// Implementation
public:
    virtual ~CGridCtrl();

protected:
    BOOL RegisterWindowClass();

    int GetSelectedCount() const                  { return m_SelectedCellMap.GetCount(); }
    CCellRange GetSelectedCellRange() const;
    void SetSelectedRange(const CCellRange& Range, BOOL bForceRepaint = FALSE);

    BOOL   SetItem(const GV_ITEM* pItem);
    BOOL   GetItem(GV_ITEM* pItem);

    // Clipboard and cut n' paste operations
    virtual void CutSelectedText();
    BOOL CopyTextFromGrid();
    BOOL PasteTextToGrid(CCellID cell);

    virtual void OnEditCut();
    virtual void OnEditCopy();
    virtual void OnEditPaste();

    CSize GetTextExtent(LPCTSTR str, BOOL bUseHeaderFont = TRUE);
    void EnsureVisible(CCellID &cell)       { EnsureVisible(cell.row, cell.col); }
    void EnsureVisible(int nRow, int nCol);
    BOOL IsCellVisible(int nRow, int nCol) const;
    BOOL IsCellVisible(CCellID cell) const;
    BOOL IsValid(const CCellID& cell) const;
    BOOL IsValid(const CCellRange& range) const;

    BOOL GetCellOrigin(int nRow, int nCol, LPPOINT p) const;
    BOOL GetCellOrigin(const CCellID& cell, LPPOINT p) const;
    BOOL GetCellRect(int nRow, int nCol, LPRECT pRect) const;
    BOOL GetCellRect(const CCellID& cell, LPRECT pRect) const;

    BOOL GetTextRect(const CCellID& cell, LPRECT pRect);
    BOOL GetTextRect(int nRow, int nCol, LPRECT pRect);

    long GetVirtualWidth() const;
    long GetVirtualHeight() const;

    BOOL InvalidateCellRect(const CCellID& cell);
    BOOL InvalidateCellRect(const CCellRange& cellRange);
    void EraseBkgnd(CDC* pDC, const CRect & rc);

    BOOL GetCellRangeRect(const CCellRange& cellRange, LPRECT lpRect) const;

    CGridCell* GetCell(int nRow, int nCol) const;
    BOOL SetCell(int nRow, int nCol, CGridCell* pCell);

    BOOL MouseOverRowResizeArea(CPoint& point) const;
    BOOL MouseOverColResizeArea(CPoint& point) const;

    CCellID GetCellFromPt(CPoint point, BOOL bAllowFixedCellCheck = TRUE) const;
    CCellID GetTopleftNonFixedCell() const;
    CCellRange GetUnobstructedNonFixedCellRange() const;
    CCellRange GetVisibleNonFixedCellRange(LPRECT pRect = NULL) const;

    virtual CCellID SetFocusCell(CCellID cell);
    CCellID SetFocusCell(int nRow, int nCol);

    void ResetSelectedRange();
    void ResetScrollBars();
    int  GetScrollPos32(int nBar, BOOL bGetTrackPos = FALSE);
    BOOL SetScrollPos32(int nBar, int nPos, BOOL bRedraw = TRUE);

// Overrrides
protected:
    // Printing    
    virtual void PrintColHeadings(CDC *pDC, CPrintInfo *pInfo);
    virtual void PrintHeader(CDC *pDC, CPrintInfo *pInfo);
    virtual void PrintFooter(CDC *pDC, CPrintInfo *pInfo);

    // Mouse Clicks
    virtual void  OnFixedColClick(CCellID& cell);
    virtual void  OnFixedRowClick(CCellID& cell);

    // Editing
    virtual CSize GetCellExtent(int nRow, int nCol, CDC* pDC);
    virtual void  OnEndEditCell(int nRow, int nCol, CString str);
    virtual void  OnEditCell(int nRow, int nCol, UINT nChar);
    virtual void  CreateInPlaceEditControl(CRect& rect, DWORD dwStyle, UINT nID,
                                           int nRow, int nCol,
                                           LPCTSTR szText, int nChar);

    // Drawing
    virtual void  OnDraw(CDC* pDC, const CRect& rcBounds, const CRect& rcInvalid);
    virtual BOOL  DrawFixedCell(CDC* pDC, int nRow, int nCol, CRect rect, BOOL bEraseBk=FALSE);
    virtual BOOL  DrawCell(CDC* pDC, int nRow, int nCol, CRect rect, BOOL bEraseBk=FALSE);

// Attributes
protected:
    // General attributes
    COLORREF    m_crTextColour, m_crTextBkColour, m_crBkColour,   // Grid colours
                m_crFixedTextColour, m_crFixedBkColour, m_crGridColour;
    COLORREF    m_crWindowText, m_crWindowColour, m_cr3DFace,     // System colours
                m_crShadow;    

    int         m_nGridLines;
    BOOL        m_bEditable;
    BOOL        m_bModified;
    BOOL        m_bListMode;
    BOOL        m_bAllowDraw;
    BOOL        m_bEnableSelection;
    BOOL        m_bDoubleBuffer;

    // Cell size details
    int         m_nRows, m_nFixedRows, m_nCols, m_nFixedCols;
    CUIntArray  m_arColWidths, m_arColFormat;
	int m_nCellHeight, m_nHeaderHeight;
    int         m_nMargin;
    int         m_nDefCellWidth;
    int         m_nVScrollMax, m_nHScrollMax;

    // Fonts and images
    LOGFONT     m_LogCellFont, m_LogHeaderFont;
    CFont       m_PrinterFont,  // for the printer
                m_CellFont, m_HeaderFont;         // for the grid
    CImageList* m_pImageList;

    // Cell data
    CTypedPtrArray<CObArray, GRID_ROW*> m_RowData;

    // Mouse operations such as cell selection
    int         m_MouseMode;
    CPoint      m_LeftClickDownPoint, m_LastMousePoint;
    CCellID     m_LeftClickDownCell, m_SelectionStartCell;
    CCellID     m_idCurrentCell;
    int         m_nResizeCaptureRange;
    BOOL        m_bAllowRowResize, m_bAllowColResize;
    int         m_nRowsPerWheelNotch;
    CMap<DWORD,DWORD, CCellID, CCellID&> m_SelectedCellMap, m_PrevSelectedCellMap;

    // Printing information
    CSize       m_CharSize;
    int         m_nPageHeight;
    CSize       m_LogicalPageSize,      // Page size in gridctrl units.
                m_PaperSize;            // Page size in device units.

protected:
	LOGFONT GetItemFont(int nRow, int nCol);
    void SelectRow(int nRowIndex);
    void SelectCell(CCellID currentCell);
    void OnSelecting(const CCellID& currentCell);
	void InitCellFont();
	void InitHeaderFont();

    // Generated message map functions
    //{{AFX_MSG(CGridCtrl)
    afx_msg void OnPaint();
    afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
    afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
    afx_msg void OnSize(UINT nType, int cx, int cy);
    afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
    afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
    afx_msg void OnMouseMove(UINT nFlags, CPoint point);
    afx_msg UINT OnGetDlgCode();
    afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
    afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
    afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
    afx_msg BOOL OnEraseBkgnd(CDC* pDC);
    afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
    afx_msg void OnSysColorChange();
	afx_msg void OnUpdateEditCopy(CCmdUI* pCmdUI);
	afx_msg void OnUpdateEditCut(CCmdUI* pCmdUI);
	afx_msg void OnUpdateEditPaste(CCmdUI* pCmdUI);
	//}}AFX_MSG
#if (_MFC_VER >= 0x0421) || (_WIN32_WCE >= 210)
    afx_msg void OnSettingChange(UINT uFlags, LPCTSTR lpszSection);
#endif
#if !defined(_WIN32_WCE) && (_MFC_VER >= 0x0421)
    afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
#endif
    afx_msg void OnEndInPlaceEdit(NMHDR* pNMHDR, LRESULT* pResult);
    DECLARE_MESSAGE_MAP()

    enum eMouseModes { MOUSE_NOTHING, MOUSE_SELECT_ROW,
                       MOUSE_SELECT_CELLS, MOUSE_SCROLLING_CELLS,
                       MOUSE_OVER_ROW_DIVIDE, MOUSE_SIZING_ROW, 
                       MOUSE_OVER_COL_DIVIDE, MOUSE_SIZING_COL,
                       MOUSE_PREPARE_EDIT,
    };
};

inline CGridCell* CGridCtrl::GetCell(int nRow, int nCol) const
{
    if (nRow < 0 || nRow >= m_nRows || nCol < 0 || nCol >= m_nCols) return NULL;

    GRID_ROW* pRow = m_RowData[nRow];
    if (!pRow) return NULL;
    return pRow->GetAt(nCol);
}

inline BOOL CGridCtrl::SetCell(int nRow, int nCol, CGridCell* pCell)
{
    if (nRow < 0 || nRow >= m_nRows || nCol < 0 || nCol >= m_nCols) return FALSE;

    GRID_ROW* pRow = m_RowData[nRow];
    if (!pRow) return FALSE;

    CGridCell * pCellNow = pRow->GetAt(nCol);
	if (!pCell && pCellNow)
		delete pCellNow;
	pRow->SetAt(nCol, pCell);

    return TRUE;
}

inline BOOL CGridCtrl::SetColFormat(int nCol, UINT nFormat)
{
	if (nCol < 0 || nCol >= m_nCols)
		return FALSE;

	m_arColFormat.SetAt(nCol, nFormat);
}

inline UINT CGridCtrl::GetColFormat(int nCol) const
{
    if (nCol < 0 || nCol >= m_nCols)
		return FALSE;
	
	return m_arColFormat.GetAt(nCol);
}
/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GRIDCTRL_H__519FA702_722C_11D1_ABBA_00A0243D1382__INCLUDED_)
