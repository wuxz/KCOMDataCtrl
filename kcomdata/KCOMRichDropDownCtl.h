#if !defined(AFX_KCOMRICHDROPDOWNCTL_H__029288E4_2FFF_11D3_B446_0080C8F18522__INCLUDED_)
#define AFX_KCOMRICHDROPDOWNCTL_H__029288E4_2FFF_11D3_B446_0080C8F18522__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMRichDropDownCtl.h : Declaration of the CKCOMRichDropDownCtrl ActiveX Control class.

/////////////////////////////////////////////////////////////////////////////
// CKCOMRichDropDownCtrl : See KCOMRichDropDownCtl.cpp for implementation.

class CDropDownColumn;
class CDropDownRealWnd;
class CGridInner;

class CKCOMRichDropDownCtrl : public COleControl
{
	DECLARE_DYNCREATE(CKCOMRichDropDownCtrl)

// Constructor
public:
	CKCOMRichDropDownCtrl();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKCOMRichDropDownCtrl)
	public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	virtual void OnForeColorChanged();
	virtual BOOL OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray);
	virtual BOOL OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut);
	//}}AFX_VIRTUAL

// Implementation
protected:
	~CKCOMRichDropDownCtrl();

protected:
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

protected:
	ROWCOL GetRowFromBmk(const VARIANT *pBmk);
	ROWCOL GetRowFromIndex(ROWCOL nIndex);
	void GetBmkOfRow(ROWCOL nRow, VARIANT *pVaRet);
	void Bind();
	void SearchText(CString strSearch);
	void DeleteRowFromSrc(ROWCOL nRow);
	void UpdateDropDownWnd();
	void Bring(int cx, int cy);
	void HideDropDownWnd(LPCTSTR pCharValue = NULL);
	CString m_strText;
	void OnBringMsg(long wParam, long lParam);
	void GetBmkOfRow(ROWCOL nRow);
	void UndeleteRecordFromHRow(HROW hRow);
	void FetchNewRows();
	ROWCOL GetRowFromHRow(HROW hRow);
	ROWCOL GetRowFromBmk(BYTE *bmk);
	void InitHeaderFont();
	void InitDropDownWnd();
	HRESULT GetRowset();
	BOOL GetColumnInfo();
	void FreeColumnMemory();
	void ClearInterfaces();
	virtual BOOL OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip);
	void SerializeColumnProp(CArchive & ar, BOOL bLoad);
	void InitCellFont();
	void UpdateDivider();
	void InitDrawGrid();
	CGridInner * m_pDrawGrid;

	// the pointer to the dropdown window
	CDropDownRealWnd * m_pDropDownRealWnd;

	int m_nColsPreset;
	HGLOBAL m_hBlob;
	BOOL m_bDataSrcChanged;
	BOOL m_bEndReached;
	CAccessorRowset<CManualAccessor> m_Rowset;
	ROWCOL m_nColsUsed;
	CStringArray m_arContent;
	CArray<ColumnData*, ColumnData*> m_apColumnData;
	CArray<BYTE*, BYTE*> m_apBookmark;
	IRowPosition * m_spRowPosition;
	DWORD m_dwCookieRN;
	BOOL m_bHaveBookmarks;
	int m_nColumns, m_nColumnsBind;
	DBCOLUMNINFO * m_pColumnInfo;
	LPOLESTR m_pStrings;
	BindData m_data;
	int m_nBookmarkSize;

	// the handle of the consumer window
	HWND m_hConsumer;
	
	DECLARE_OLECREATE_EX(CKCOMRichDropDownCtrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CKCOMRichDropDownCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CKCOMRichDropDownCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CKCOMRichDropDownCtrl)		// Type name and misc status

// Message maps
	//{{AFX_MSG(CKCOMRichDropDownCtrl)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	//{{AFX_DISPATCH(CKCOMRichDropDownCtrl)
	OLE_COLOR m_clrHeadBack;
	afx_msg void OnHeadBackColorChanged();
	OLE_COLOR m_clrHeadFore;
	afx_msg void OnHeadForeColorChanged();
	BOOL m_bListAutoPosition;
	afx_msg void OnListAutoPositionChanged();
	BOOL m_bListWidthAutoSize;
	afx_msg void OnListWidthAutoSizeChanged();
	afx_msg LPFONTDISP GetFont();
	afx_msg void SetFont(LPFONTDISP newValue);
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
	afx_msg BSTR GetDataField();
	afx_msg void SetDataField(LPCTSTR lpszNewValue);
	afx_msg LPDISPATCH GetColumns();
	afx_msg long GetVisibleRows();
	afx_msg long GetVisibleCols();
	afx_msg BOOL GetColumnHeaders();
	afx_msg void SetColumnHeaders(BOOL bNewValue);
	afx_msg LPFONTDISP GetHeadFont();
	afx_msg void SetHeadFont(LPFONTDISP newValue);
	afx_msg long GetHeaderHeight();
	afx_msg void SetHeaderHeight(long nNewValue);
	afx_msg LPUNKNOWN GetDataSource();
	afx_msg void SetDataSource(LPUNKNOWN newValue);
	afx_msg BSTR GetDataMember();
	afx_msg void SetDataMember(LPCTSTR lpszNewValue);
	afx_msg BSTR GetFieldSeparator();
	afx_msg void SetFieldSeparator(LPCTSTR lpszNewValue);
	afx_msg short GetMaxDropDownItems();
	afx_msg void SetMaxDropDownItems(short nNewValue);
	afx_msg short GetMinDropDownItems();
	afx_msg void SetMinDropDownItems(short nNewValue);
	afx_msg short GetListWidth();
	afx_msg void SetListWidth(short nNewValue);
	afx_msg OLE_COLOR GetBackColor();
	afx_msg void SetBackColor(OLE_COLOR nNewValue);
	afx_msg VARIANT GetBookmark();
	afx_msg void SetBookmark(const VARIANT FAR& newValue);
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
	//{{AFX_EVENT(CKCOMRichDropDownCtrl)
	void FireScroll(short FAR* Cancel)
		{FireEvent(eventidScroll,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireCloseUp()
		{FireEvent(eventidCloseUp,EVENT_PARAM(VTS_NONE));}
	void FireDropDown()
		{FireEvent(eventidDropDown,EVENT_PARAM(VTS_NONE));}
	void FirePositionList(LPCTSTR Text)
		{FireEvent(eventidPositionList,EVENT_PARAM(VTS_BSTR), Text);}
	void FireScrollAfter()
		{FireEvent(eventidScrollAfter,EVENT_PARAM(VTS_NONE));}
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	BOOL m_bDroppedDown;
	void GetCellValue(ROWCOL nRow, ROWCOL nCol, CGXStyle & style);
	void InitGridFromProp();
	void InitColumnObject(int nCol, CDropDownColumn * pCol);
	ROWCOL GetColOrdinalFromField(int nField);
	ROWCOL GetColOrdinalFromCol(ROWCOL nCol);
	ROWCOL GetColOrdinalFromIndex(ROWCOL nColIndex);
	int GetFieldFromStr(CString str);
	void CreateDrawGrid();
	HWND GetApplicationWindow();
	ROWCOL GetRowColFromVariant(VARIANT va);
	CArray<CDropDownColumn *, CDropDownColumn *> m_apColumns;
	CArray<ColumnProp *, ColumnProp *> m_apColumnsProp;

	enum {
	//{{AFX_DISP_ID(CKCOMRichDropDownCtrl)
	dispidFont = 5L,
	dispidDefColWidth = 6L,
	dispidRowHeight = 7L,
	dispidDividerType = 8L,
	dispidDividerStyle = 9L,
	dispidDataMode = 10L,
	dispidLeftCol = 11L,
	dispidFirstRow = 12L,
	dispidRow = 13L,
	dispidCol = 14L,
	dispidRows = 15L,
	dispidCols = 16L,
	dispidDroppedDown = 17L,
	dispidDataField = 18L,
	dispidColumns = 19L,
	dispidVisibleRows = 20L,
	dispidVisibleCols = 21L,
	dispidHeadBackColor = 1L,
	dispidHeadForeColor = 2L,
	dispidColumnHeaders = 22L,
	dispidHeadFont = 23L,
	dispidHeaderHeight = 24L,
	dispidDataSource = 25L,
	dispidDataMember = 26L,
	dispidFieldSeparator = 27L,
	dispidListAutoPosition = 3L,
	dispidListWidthAutoSize = 4L,
	dispidMaxDropDownItems = 28L,
	dispidMinDropDownItems = 29L,
	dispidListWidth = 30L,
	dispidBackColor = 31L,
	dispidBookmark = 32L,
	dispidAddItem = 33L,
	dispidRemoveItem = 34L,
	dispidScroll = 35L,
	dispidRemoveAll = 36L,
	dispidMoveFirst = 37L,
	dispidMovePrevious = 38L,
	dispidMoveNext = 39L,
	dispidMoveLast = 40L,
	dispidMoveRecords = 41L,
	dispidReBind = 42L,
	dispidRowBookmark = 43L,
	eventidScroll = 1L,
	eventidCloseUp = 2L,
	eventidDropDown = 3L,
	eventidPositionList = 4L,
	eventidScrollAfter = 5L,
	//}}AFX_DISP_ID
	dispidCtrlType = 255L,
	};

	DECLARE_INTERFACE_MAP()

	BEGIN_INTERFACE_PART(RowsetNotify, IRowsetNotify)
		INIT_INTERFACE_PART(CKCOMRichDropDownCtrl, RowsetNotify)
		STDMETHOD(OnFieldChange)(IRowset *, HROW, ULONG, ULONG[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowChange)(IRowset *, ULONG, const HROW[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowsetChange)(IRowset *, DBREASON, DBEVENTPHASE, BOOL);
	END_INTERFACE_PART(RowsetNotify)

	BEGIN_INTERFACE_PART(DataSourceListener, IDataSourceListener)
		INIT_INTERFACE_PART(CKCOMRichDropDownCtrl, DataSourceListener)
		STDMETHOD(dataMemberChanged)(BSTR bstrDM);
		STDMETHOD(dataMemberAdded)(BSTR bstrDM);
		STDMETHOD(dataMemberRemoved)(BSTR bstrDM);
	END_INTERFACE_PART(DataSourceListener)

	friend class CDropDownColumn;
	friend class CDropDownColumns;
	friend class CKCOMRichGridPropColumnPage;
	friend class CDropDownRealWnd;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMRICHDROPDOWNCTL_H__029288E4_2FFF_11D3_B446_0080C8F18522__INCLUDED)
