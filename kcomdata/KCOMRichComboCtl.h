#if !defined(AFX_KCOMRICHCOMBOCTL_H__029288DF_2FFF_11D3_B446_0080C8F18522__INCLUDED_)
#define AFX_KCOMRICHCOMBOCTL_H__029288DF_2FFF_11D3_B446_0080C8F18522__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMRichComboCtl.h : Declaration of the CKCOMRichComboCtrl ActiveX Control class.

/////////////////////////////////////////////////////////////////////////////
// CKCOMRichComboCtrl : See KCOMRichComboCtl.cpp for implementation.
class CDropDownColumn;
class CDropDownRealWnd;

#define IDC_EDIT 1001
#define IDC_BUTTON 1002
#define BUTTONWIDTH 18
#define BUTTONHEIGHT 18
#define PIXELRATIO 40

// the edit part
class CComboEdit;

class CKCOMRichComboCtrl : public COleControl
{
	DECLARE_DYNCREATE(CKCOMRichComboCtrl)

// Constructor
public:
	CKCOMRichComboCtrl();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKCOMRichComboCtrl)
	public:
	virtual void OnForeColorChanged();
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	virtual BOOL OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray);
	virtual BOOL OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut);
	virtual BOOL OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip);
	virtual BOOL OnSetExtent(LPSIZEL lpSizeL);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	//}}AFX_VIRTUAL

// Implementation
protected:
	~CKCOMRichComboCtrl();
	
	// setup the mouse hook proc
	void HookMouse(HWND hWnd);

	void ShowDropDownWnd();
	void HideDropDownWnd();
	void OnDropDown();
	
	// bind to datasource
	void Bind();

	// search the text in datasource or manually input data
	ROWCOL SearchText(CString strSearch, BOOL bAutoPosition = TRUE);

	// when one record was deleted from datasource, update the grid
	void DeleteRowFromSrc(ROWCOL nRow);
	
	// update the dropdown window
	void UpdateDropDownWnd();

	// get the bookmark of one row
	void GetBmkOfRow(ROWCOL nRow);

	// when one record was undeleted from datasource, update the grid
	void UndeleteRecordFromHRow(HROW hRow);
	
	// fetch new records to fill the grid
	void FetchNewRows();

	// given hrow, calc the row ordinal
	ROWCOL GetRowFromHRow(HROW hRow);

	// given bookmark, calc the row ordinal
	ROWCOL GetRowFromBmk(BYTE *bmk);

	// init the headerfont property
	void InitHeaderFont();

	// init the dropdown window
	void InitDropDownWnd();

	HRESULT GetRowset();
	BOOL GetColumnInfo();
	void FreeColumnMemory();
	
	// clearup cached interfaces
	void ClearInterfaces();

	// serialize the columns' prop between disk file
	void SerializeColumnProp(CArchive & ar, BOOL bLoad);

	// init the cell font property
	void InitCellFont();

	// update the grid divider line
	void UpdateDivider();

	// the function to decide whether the char is valid
	BOOL IsValidChar(UINT nChar, int nPosition);
	
	// the function to decide whether the given text format is valid
	BOOL IsValidFormat(CString strNewFormat);

	// the function to decide whether the given text format is valid
	BOOL IsValidText(CString strNewText);

	// the function to calc mask signal
	void CalcMask();

protected:
	long m_nStyle;

	// the max length of the text of the editbox
	short m_nMaxLength;

	CString m_strText;

	// the current text format
	CString m_strFormat;

	CFontHolder m_fntCell;
	CFontHolder m_fntHeader;
	long m_nRowHeight;
	long m_nHeaderHeight;
	long m_nDataMode;
	IDataSource * m_spDataSource;
	CString m_strDataMember;
	long m_nDefColWidth;
	long m_nRows;
	long m_nCols;
	long m_nDividerType;
	long m_nDividerStyle;
	ROWCOL m_nCol;
	ROWCOL m_nRow;
	CString m_strFieldSeparator, m_strSeparator;
	CString m_strDataField;
	BOOL m_bColumnHeaders;
	short m_nMaxDropDownItems;
	short m_nMinDropDownItems;
	short m_nListWidth;
	OLE_COLOR m_clrBack;
	BOOL m_bReadOnly;

protected:
	ROWCOL GetRowFromBmk(const VARIANT *pBmk);
	void GetBmkOfRow(ROWCOL nRow, VARIANT *pVaRet);
	ROWCOL GetRowFromIndex(ROWCOL nIndex);
	
	// the dropdown window
	CDropDownRealWnd * m_pDropDownRealWnd;

	// the brush to draw the background
	CBrush * m_pBrhBack;

	// the system mouse hook handle
	HHOOK m_hSystemHook;

	// the number of the manually input columns data
	int m_nColsPreset;

	// the handle of the blob data serialize columns prop
	HGLOBAL m_hBlob;

	// whether the datasource was changed
	BOOL m_bDataSrcChanged;

	// Have we reached the end of the rowset?
	BOOL m_bEndReached;

	CAccessorRowset<CManualAccessor> m_Rowset;
	
	// the used number of the col name digit
	ROWCOL m_nColsUsed;

	// the user input combobox data
	CStringArray m_arContent;

	// the filtered column info
	CArray<ColumnData*, ColumnData*> m_apColumnData;

	// the bookmark data
	CArray<BYTE*, BYTE*> m_apBookmark;

	IRowPosition * m_spRowPosition;
	
	// sink cookie
	DWORD m_dwCookieRN;

	// whether the datasource support bookmark
	BOOL m_bHaveBookmarks;

	int m_nColumns, m_nColumnsBind;

	// the raw column info data
	DBCOLUMNINFO * m_pColumnInfo;
	LPOLESTR m_pStrings;
	
	// the memory to hold the data of current record
	BindData m_data;

	// the size of bookmark
	int m_nBookmarkSize;
	
	// Is the LButton pressed?
	BOOL m_bLButtonDown;

	// the rect to draw the button
	CRect m_rectButton;

	// the current text mask
	CString m_strMask;


	// the signal to mark whether current char is delimiter or data
	BOOL m_bIsDelimiter[255];

	// the pointer of the editbox window
	CComboEdit * m_pEdit;

	// the handle of the using font for editbox
	HFONT m_hFont;

	// the height of the using font
	int m_nFontHeight;
	
	DECLARE_OLECREATE_EX(CKCOMRichComboCtrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CKCOMRichComboCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CKCOMRichComboCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CKCOMRichComboCtrl)		// Type name and misc status

// Message maps
	//{{AFX_MSG(CKCOMRichComboCtrl)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg int OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	//{{AFX_DISPATCH(CKCOMRichComboCtrl)
	OLE_COLOR m_clrHeadBack;
	afx_msg void OnHeadBackColorChanged();
	OLE_COLOR m_clrHeadFore;
	afx_msg void OnHeadForeColorChanged();
	BOOL m_bListAutoPosition;
	afx_msg void OnListAutoPositionChanged();
	BOOL m_bListWidthAutoSize;
	afx_msg void OnListWidthAutoSizeChanged();
	afx_msg OLE_COLOR GetBackColor();
	afx_msg void SetBackColor(OLE_COLOR nNewValue);
	afx_msg BSTR GetFormat();
	afx_msg void SetFormat(LPCTSTR lpszNewValue);
	afx_msg short GetMaxLength();
	afx_msg void SetMaxLength(short nNewValue);
	afx_msg BOOL GetReadOnly();
	afx_msg void SetReadOnly(BOOL bNewValue);
	afx_msg LPUNKNOWN GetDataSource();
	afx_msg void SetDataSource(LPUNKNOWN newValue);
	afx_msg BSTR GetDataMember();
	afx_msg void SetDataMember(LPCTSTR lpszNewValue);
	afx_msg long GetDefColWidth();
	afx_msg void SetDefColWidth(long nNewValue);
	afx_msg long GetRowHeight();
	afx_msg void SetRowHeight(long nNewValue);
	afx_msg long GetDividerType();
	afx_msg void SetDividerType(long nNewValue);
	afx_msg long GetDividerStyle();
	afx_msg void SetDividerStyle(long nNewValue);
	afx_msg long GetDataMode();
	afx_msg void SetDataMode(long nNewValue);
	afx_msg long GetLeftCol();
	afx_msg void SetLeftCol(long nNewValue);
	afx_msg long GetFirstRow();
	afx_msg void SetFirstRow(long nNewValue);
	afx_msg long GetRow();
	afx_msg void SetRow(long nNewValue);
	afx_msg long GetCol();
	afx_msg void SetCol(long nNewValue);
	afx_msg long GetRows();
	afx_msg void SetRows(long nNewValue);
	afx_msg long GetCols();
	afx_msg void SetCols(long nNewValue);
	afx_msg BOOL GetDroppedDown();
	afx_msg void SetDroppedDown(BOOL bNewValue);
	afx_msg BSTR GetDataField();
	afx_msg void SetDataField(LPCTSTR lpszNewValue);
	afx_msg LPDISPATCH GetColumns();
	afx_msg long GetVisibleRows();
	afx_msg long GetVisibleCols();
	afx_msg BOOL GetColumnHeaders();
	afx_msg void SetColumnHeaders(BOOL bNewValue);
	afx_msg LPFONTDISP GetFont();
	afx_msg void SetFont(LPFONTDISP newValue);
	afx_msg LPFONTDISP GetHeadFont();
	afx_msg void SetHeadFont(LPFONTDISP newValue);
	afx_msg BSTR GetFieldSeparator();
	afx_msg void SetFieldSeparator(LPCTSTR lpszNewValue);
	afx_msg short GetMaxDropDownItems();
	afx_msg void SetMaxDropDownItems(short nNewValue);
	afx_msg short GetMinDropDownItems();
	afx_msg void SetMinDropDownItems(short nNewValue);
	afx_msg short GetListWidth();
	afx_msg void SetListWidth(short nNewValue);
	afx_msg long GetHeaderHeight();
	afx_msg void SetHeaderHeight(long nNewValue);
	afx_msg BSTR GetText();
	afx_msg void SetText(LPCTSTR lpszNewValue);
	afx_msg VARIANT GetBookmark();
	afx_msg void SetBookmark(const VARIANT FAR& newValue);
	afx_msg long GetStyle();
	afx_msg void SetStyle(long nNewValue);
	afx_msg void AddItem(LPCTSTR Item, const VARIANT FAR& RowIndex);
	afx_msg void RemoveItem(long RowIndex);
	afx_msg void Scroll(long Rows, long Cols);
	afx_msg void RemoveAll();
	afx_msg void MoveFirst();
	afx_msg void MovePrevious();
	afx_msg void MoveNext();
	afx_msg void MoveLast();
	afx_msg void MoveRecords(long Rows);
	afx_msg void ReBind();
	afx_msg BOOL IsItemInList();
	afx_msg VARIANT RowBookmark(long RowIndex);
	//}}AFX_DISPATCH
	afx_msg short GetCtrlType();
	afx_msg void AboutBox();

	DECLARE_DISPATCH_MAP()

	void FreeBookmarkMemory()
	{
		if (m_apBookmark.GetSize() == 0)
			return;

		int i, nSize = m_apBookmark.GetSize();
		for (i=0; i<nSize; i++)
		{
			delete [] m_apBookmark[i];
		}
		m_apBookmark.RemoveAll();
	}

	void FreeColumnData()
	{
		if (m_apColumnData.GetSize() == 0)
			return;

		int i, nSize = m_apColumnData.GetSize();
		for (i=0; i<nSize; i++)
		{
			delete m_apColumnData[i];
		}
		m_apColumnData.RemoveAll();
	}

// Event maps
	//{{AFX_EVENT(CKCOMRichComboCtrl)
	void FireScroll(short FAR* Cancel)
		{FireEvent(eventidScroll,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireScrollAfter()
		{FireEvent(eventidScrollAfter,EVENT_PARAM(VTS_NONE));}
	void FireCloseUp()
		{FireEvent(eventidCloseUp,EVENT_PARAM(VTS_NONE));}
	void FireDropDown()
		{FireEvent(eventidDropDown,EVENT_PARAM(VTS_NONE));}
	void FirePositionList(LPCTSTR Text)
		{FireEvent(eventidPositionList,EVENT_PARAM(VTS_BSTR), Text);}
	void FireChange()
		{FireEvent(eventidChange,EVENT_PARAM(VTS_NONE));}
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	// load the data from datasource
	void GetCellValue(ROWCOL nRow, ROWCOL nCol, CGXStyle & style);

	// init the grid from props saved
	void InitGridFromProp();

	// init a column object
	void InitColumnObject(int nCol, CDropDownColumn * pCol);

	// get the ordinal in columns array
	ROWCOL GetColOrdinalFromField(int nField);
	ROWCOL GetColOrdinalFromCol(ROWCOL nCol);
	ROWCOL GetColOrdinalFromIndex(ROWCOL nColIndex);

	// get the field ordinal of the datasource
	int GetFieldFromStr(CString str);

	ROWCOL GetRowColFromVariant(VARIANT va);
	
	// the column objects array
	CArray<CDropDownColumn *, CDropDownColumn *> m_apColumns;

	// the saved props of all the columns
	CArray<ColumnProp *, ColumnProp *> m_apColumnsProp;

	enum {
	//{{AFX_DISP_ID(CKCOMRichComboCtrl)
	dispidHeadBackColor = 1L,
	dispidHeadForeColor = 2L,
	dispidBackColor = 5L,
	dispidFormat = 6L,
	dispidMaxLength = 7L,
	dispidReadOnly = 8L,
	dispidDataSource = 9L,
	dispidDataMember = 10L,
	dispidDefColWidth = 11L,
	dispidRowHeight = 12L,
	dispidDividerType = 13L,
	dispidDividerStyle = 14L,
	dispidDataMode = 15L,
	dispidLeftCol = 16L,
	dispidFirstRow = 17L,
	dispidRow = 18L,
	dispidCol = 19L,
	dispidRows = 20L,
	dispidCols = 21L,
	dispidDroppedDown = 22L,
	dispidDataField = 23L,
	dispidColumns = 24L,
	dispidVisibleRows = 25L,
	dispidVisibleCols = 26L,
	dispidColumnHeaders = 27L,
	dispidFont = 28L,
	dispidHeadFont = 29L,
	dispidFieldSeparator = 30L,
	dispidListAutoPosition = 3L,
	dispidListWidthAutoSize = 4L,
	dispidMaxDropDownItems = 31L,
	dispidMinDropDownItems = 32L,
	dispidListWidth = 33L,
	dispidHeaderHeight = 34L,
	dispidText = 35L,
	dispidBookmark = 36L,
	dispidStyle = 37L,
	dispidAddItem = 38L,
	dispidRemoveItem = 39L,
	dispidScroll = 40L,
	dispidRemoveAll = 41L,
	dispidMoveFirst = 42L,
	dispidMovePrevious = 43L,
	dispidMoveNext = 44L,
	dispidMoveLast = 45L,
	dispidMoveRecords = 46L,
	dispidReBind = 47L,
	dispidIsItemInList = 48L,
	dispidRowBookmark = 49L,
	eventidScroll = 1L,
	eventidScrollAfter = 2L,
	eventidCloseUp = 3L,
	eventidDropDown = 4L,
	eventidPositionList = 5L,
	eventidChange = 6L,
	//}}AFX_DISP_ID
	dispidCtrlType = 255L,
	};

	DECLARE_INTERFACE_MAP()

	BEGIN_INTERFACE_PART(RowsetNotify, IRowsetNotify)
		INIT_INTERFACE_PART(CKCOMRichComboCtrl, RowsetNotify)
		STDMETHOD(OnFieldChange)(IRowset *, HROW, ULONG, ULONG[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowChange)(IRowset *, ULONG, const HROW[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowsetChange)(IRowset *, DBREASON, DBEVENTPHASE, BOOL);
	END_INTERFACE_PART(RowsetNotify)

	BEGIN_INTERFACE_PART(DataSourceListener, IDataSourceListener)
		INIT_INTERFACE_PART(CKCOMRichComboCtrl, DataSourceListener)
		STDMETHOD(dataMemberChanged)(BSTR bstrDM);
		STDMETHOD(dataMemberAdded)(BSTR bstrDM);
		STDMETHOD(dataMemberRemoved)(BSTR bstrDM);
	END_INTERFACE_PART(DataSourceListener)

	friend class CDropDownColumn;
	friend class CDropDownColumns;
	friend class CKCOMRichGridPropColumnPage;
	friend class CDropDownRealWnd;
	friend class CComboEdit;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMRICHCOMBOCTL_H__029288DF_2FFF_11D3_B446_0080C8F18522__INCLUDED)
