#if !defined(AFX_KCOMRICHGRIDCTL_H__029288DA_2FFF_11D3_B446_0080C8F18522__INCLUDED_)
#define AFX_KCOMRICHGRIDCTL_H__029288DA_2FFF_11D3_B446_0080C8F18522__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMRichGridCtl.h : Declaration of the CKCOMRichGridCtrl ActiveX Control class.

/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridCtrl : See KCOMRichGridCtl.cpp for implementation.

class CGridInner;
class CGridColumn;

class CKCOMRichGridCtrl : public COleControl
{
	DECLARE_DYNCREATE(CKCOMRichGridCtrl)

// Constructor
public:
	CKCOMRichGridCtrl();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKCOMRichGridCtrl)
	public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	virtual BOOL OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip);
	virtual void OnForeColorChanged();
	virtual BOOL OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray);
	virtual BOOL OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual void OnTextChanged();
	//}}AFX_VIRTUAL

// Implementation
protected:
	CFontHolder m_fntCell;
	CFontHolder m_fntHeader;
	long m_nRowHeight;
	long m_nHeaderHeight;
	long m_nDataMode;
	IDataSource * m_spDataSource;
	CString m_strDataMember;
	long m_nDefColWidth;
	BOOL m_bAllowDelete;
	BOOL m_bAllowAddNew;
	BOOL m_bAllowUpdate;
	BOOL m_bRecordSelectors;
	long m_nRows;
	long m_nCols;
	long m_nCapitonAlignment;
	long m_nDividerType;
	long m_nDividerStyle;
	BOOL m_bAllowRowSizing;
	BOOL m_bAllowColumnSizing;
	BOOL m_bColumnHeaders;
	ROWCOL m_nCol;
	ROWCOL m_nRow;
	BOOL m_bRedraw;
	BOOL m_bAllowColumnMoving;
	BOOL m_bSelectByCell;
	BOOL m_bRowChanged;
	CString m_strFieldSeparator, m_strSeparator;
	OLE_COLOR m_clrBack;
	long m_nFrozenCols;

protected:
	int m_nColsPreset;
	HHOOK m_hSystemHook;
	BOOL m_bDataSrcChanged;
	BOOL m_bEndReached;
	CAccessorRowset<CManualAccessor> m_Rowset;
	ROWCOL m_nColsUsed;
	HGLOBAL m_hBlob;
	CStringArray m_arContent;
	CArray<ColumnData*, ColumnData*> m_apColumnData;
	CArray<BYTE*, BYTE*> m_apBookmark;
	IRowPosition * m_spRowPosition;
	DWORD m_dwCookieRPC, m_dwCookieRN;
	BOOL m_bHaveBookmarks;
	int m_nColumns, m_nColumnsBind;
	DBCOLUMNINFO * m_pColumnInfo;
	LPOLESTR m_pStrings;
	BindData m_data;
	int m_nBookmarkSize;
	BOOL m_bIsPosSync;
	CArray<HROW, HROW> m_arHRowToDelete;
	CArray<ROWCOL, ROWCOL> m_arRowToDelete;

protected:
	ROWCOL GetRowFromBmk(const VARIANT * pBmk);
	void GetBmkOfRow(ROWCOL nRow, VARIANT * pVaRet);
	ROWCOL GetRowFromIndex(ROWCOL nIndex);
	void ExchangeStockProps(CPropExchange* pPX);
	void Bind();
	HWND m_hWndDropDown;
	void HookMouse(HWND hWnd);
	HWND m_hExternalDropDownWnd;
	void OnBringMsg(long wParam, long lParam);
	void HideDropDownWnd();
	BOOL m_bDroppedDown;
	void OnDropDown(ROWCOL nRow, ROWCOL nCol);
	void GetBmkOfRow(ROWCOL nRow);
	void UndeleteRecordFromHRow(HROW hRow, ROWCOL nRow);
	void SerializeColumnProp(CArchive & ar, BOOL bLoad);
	void InitColumnObject(int nCol, CGridColumn * pCol);
	void UpdateDivider();
	void InitGridHeader();
	void InitCellFont();
	void InitHeaderFont();
	BOOL GetColumnInfo();
	HRESULT UpdateRowData(ROWCOL nRow);
	void DeleteRowData(ROWCOL nRow);
	void FreeColumnMemory();
	void ClearInterfaces();
	void UpdateRowFromSrc(ROWCOL nRow);
	void DeleteRowFromSrc(ROWCOL nRow);
	void UpdateCellFromSrc(ROWCOL nRow, ROWCOL nCol);
	void FetchNewRows();
	void GetCellValue(ROWCOL nRow, ROWCOL nCol, CGXStyle & style);
	HRESULT GetRowset();
	ROWCOL GetRowFromBmk(BYTE * bmk);
	ROWCOL GetRowFromHRow(HROW hRow);
	void UpdateControl();
	void SetCurrentCell(ROWCOL nRow, ROWCOL nCol);

	~CKCOMRichGridCtrl();

	DECLARE_OLECREATE_EX(CKCOMRichGridCtrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CKCOMRichGridCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CKCOMRichGridCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CKCOMRichGridCtrl)		// Type name and misc status

// Message maps
	//{{AFX_MSG(CKCOMRichGridCtrl)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg int OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	//{{AFX_DISPATCH(CKCOMRichGridCtrl)
	OLE_COLOR m_clrHeadFore;
	afx_msg void OnHeadForeColorChanged();
	OLE_COLOR m_clrHeadBack;
	afx_msg void OnHeadBackColorChanged();
	OLE_COLOR m_clrDivider;
	afx_msg void OnDividerColorChanged();
	afx_msg long GetCaptionAlignment();
	afx_msg void SetCaptionAlignment(long nNewValue);
	afx_msg long GetDividerType();
	afx_msg void SetDividerType(long nNewValue);
	afx_msg BOOL GetAllowAddNew();
	afx_msg void SetAllowAddNew(BOOL bNewValue);
	afx_msg BOOL GetAllowDelete();
	afx_msg void SetAllowDelete(BOOL bNewValue);
	afx_msg BOOL GetAllowUpdate();
	afx_msg void SetAllowUpdate(BOOL bNewValue);
	afx_msg BOOL GetAllowRowSizing();
	afx_msg void SetAllowRowSizing(BOOL bNewValue);
	afx_msg BOOL GetRecordSelectors();
	afx_msg void SetRecordSelectors(BOOL bNewValue);
	afx_msg BOOL GetAllowColumnSizing();
	afx_msg void SetAllowColumnSizing(BOOL bNewValue);
	afx_msg BOOL GetColumnHeaders();
	afx_msg void SetColumnHeaders(BOOL bNewValue);
	afx_msg LPFONTDISP GetHeadFont();
	afx_msg void SetHeadFont(LPFONTDISP newValue);
	afx_msg LPFONTDISP GetFont();
	afx_msg void SetFont(LPFONTDISP newValue);
	afx_msg long GetCol();
	afx_msg void SetCol(long nNewValue);
	afx_msg long GetRow();
	afx_msg void SetRow(long nNewValue);
	afx_msg long GetDataMode();
	afx_msg void SetDataMode(long nNewValue);
	afx_msg long GetVisibleCols();
	afx_msg long GetVisibleRows();
	afx_msg long GetLeftCol();
	afx_msg void SetLeftCol(long nNewValue);
	afx_msg long GetCols();
	afx_msg void SetCols(long nNewValue);
	afx_msg long GetRows();
	afx_msg void SetRows(long nNewValue);
	afx_msg BSTR GetFieldSeparator();
	afx_msg void SetFieldSeparator(LPCTSTR lpszNewValue);
	afx_msg long GetDividerStyle();
	afx_msg void SetDividerStyle(long nNewValue);
	afx_msg long GetFirstRow();
	afx_msg void SetFirstRow(long nNewValue);
	afx_msg long GetDefColWidth();
	afx_msg void SetDefColWidth(long nNewValue);
	afx_msg long GetRowHeight();
	afx_msg void SetRowHeight(long nNewValue);
	afx_msg BOOL GetRedraw();
	afx_msg void SetRedraw(BOOL bNewValue);
	afx_msg BOOL GetAllowColumnMoving();
	afx_msg void SetAllowColumnMoving(BOOL bNewValue);
	afx_msg BOOL GetSelectByCell();
	afx_msg void SetSelectByCell(BOOL bNewValue);
	afx_msg LPDISPATCH GetColumns();
	afx_msg LPUNKNOWN GetDataSource();
	afx_msg void SetDataSource(LPUNKNOWN newValue);
	afx_msg BSTR GetDataMember();
	afx_msg void SetDataMember(LPCTSTR lpszNewValue);
	afx_msg BOOL GetRowChanged();
	afx_msg void SetRowChanged(BOOL bNewValue);
	afx_msg long GetHeaderHeight();
	afx_msg void SetHeaderHeight(long nNewValue);
	afx_msg OLE_COLOR GetBackColor();
	afx_msg void SetBackColor(OLE_COLOR nNewValue);
	afx_msg VARIANT GetBookmark();
	afx_msg void SetBookmark(const VARIANT FAR& newValue);
	afx_msg LPDISPATCH GetSelBookmarks();
	afx_msg long GetFrozenCols();
	afx_msg void SetFrozenCols(long nNewValue);
	afx_msg void AddItem(LPCTSTR Item, const VARIANT FAR& RowIndex);
	afx_msg void RemoveItem(long RowIndex);
	afx_msg void Scroll(long Rows, long Cols);
	afx_msg void RemoveAll();
	afx_msg void MoveFirst();
	afx_msg void MovePrevious();
	afx_msg void MoveNext();
	afx_msg void MoveLast();
	afx_msg void MoveRecords(long Rows);
	afx_msg BOOL IsAddRow();
	afx_msg void Update();
	afx_msg void CancelUpdate();
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
	//{{AFX_EVENT(CKCOMRichGridCtrl)
	void FireBtnClick()
		{FireEvent(eventidBtnClick,EVENT_PARAM(VTS_NONE));}
	void FireAfterColUpdate(short ColIndex)
		{FireEvent(eventidAfterColUpdate,EVENT_PARAM(VTS_I2), ColIndex);}
	void FireBeforeColUpdate(short ColIndex, LPCTSTR OldValue, short FAR* Cancel)
		{FireEvent(eventidBeforeColUpdate,EVENT_PARAM(VTS_I2  VTS_BSTR  VTS_PI2), ColIndex, OldValue, Cancel);}
	void FireBeforeInsert(short FAR* Cancel)
		{FireEvent(eventidBeforeInsert,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireBeforeUpdate(short FAR* Cancel)
		{FireEvent(eventidBeforeUpdate,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireColResize(short ColIndex, short FAR* Cancel)
		{FireEvent(eventidColResize,EVENT_PARAM(VTS_I2  VTS_PI2), ColIndex, Cancel);}
	void FireRowResize(short FAR* Cancel)
		{FireEvent(eventidRowResize,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireRowColChange()
		{FireEvent(eventidRowColChange,EVENT_PARAM(VTS_NONE));}
	void FireScroll(short FAR* Cancel)
		{FireEvent(eventidScroll,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireChange()
		{FireEvent(eventidChange,EVENT_PARAM(VTS_NONE));}
	void FireAfterDelete(short FAR* RtnDispErrMsg)
		{FireEvent(eventidAfterDelete,EVENT_PARAM(VTS_PI2), RtnDispErrMsg);}
	void FireBeforeDelete(short FAR* Cancel, short FAR* DispPromptMsg)
		{FireEvent(eventidBeforeDelete,EVENT_PARAM(VTS_PI2  VTS_PI2), Cancel, DispPromptMsg);}
	void FireAfterUpdate(short FAR* RtnDispErrMsg)
		{FireEvent(eventidAfterUpdate,EVENT_PARAM(VTS_PI2), RtnDispErrMsg);}
	void FireAfterInsert(short FAR* RtnDispErrMsg)
		{FireEvent(eventidAfterInsert,EVENT_PARAM(VTS_PI2), RtnDispErrMsg);}
	void FireColSwap(short FromCol, short ToCol, short DestCol, short FAR* Cancel)
		{FireEvent(eventidColSwap,EVENT_PARAM(VTS_I2  VTS_I2  VTS_I2  VTS_PI2), FromCol, ToCol, DestCol, Cancel);}
	void FireUpdateError()
		{FireEvent(eventidUpdateError,EVENT_PARAM(VTS_NONE));}
	void FireBeforeRowColChange(short FAR* Cancel)
		{FireEvent(eventidBeforeRowColChange,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireScrollAfter()
		{FireEvent(eventidScrollAfter,EVENT_PARAM(VTS_NONE));}
	void FireBeforeDrawCell(long Row, long Col, VARIANT FAR* Picture, OLE_COLOR* BkColor, short FAR* DrawText)
		{FireEvent(eventidBeforeDrawCell,EVENT_PARAM(VTS_I4  VTS_I4  VTS_PVARIANT  VTS_PCOLOR  VTS_PI2), Row, Col, Picture, BkColor, DrawText);}
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	int GetDataTypeFromStr(CString strField);
	ROWCOL GetColOrdinalFromField(int nField);
	ROWCOL GetColOrdinalFromCol(ROWCOL nCol);
	ROWCOL GetColOrdinalFromIndex(ROWCOL nColIndex);
	int GetFieldFromStr(CString str);
	void CreateDrawGrid();
	HWND GetApplicationWindow();
	void ButtonDblClick(UINT nFlags, CPoint point);
	void ButtonDown(UINT nFlags, CPoint point);
	void ButtonUp(UINT nFlags, CPoint point);
	void InitGridFromProp();
	ROWCOL GetRowColFromVariant(VARIANT va);
	CArray<CGridColumn *, CGridColumn *> m_apColumns;
	CArray<ColumnProp *, ColumnProp *> m_apColumnsProp;
	CGridInner * m_pGridInner;

	enum {
	//{{AFX_DISP_ID(CKCOMRichGridCtrl)
	dispidCaptionAlignment = 4L,
	dispidDividerType = 5L,
	dispidAllowAddNew = 6L,
	dispidAllowDelete = 7L,
	dispidAllowUpdate = 8L,
	dispidAllowRowSizing = 9L,
	dispidRecordSelectors = 10L,
	dispidAllowColumnSizing = 11L,
	dispidColumnHeaders = 12L,
	dispidHeadFont = 13L,
	dispidFont = 14L,
	dispidCol = 15L,
	dispidRow = 16L,
	dispidDataMode = 17L,
	dispidVisibleCols = 18L,
	dispidVisibleRows = 19L,
	dispidLeftCol = 20L,
	dispidCols = 21L,
	dispidRows = 22L,
	dispidFieldSeparator = 23L,
	dispidDividerStyle = 24L,
	dispidFirstRow = 25L,
	dispidDefColWidth = 26L,
	dispidRowHeight = 27L,
	dispidRedraw = 28L,
	dispidAllowColumnMoving = 29L,
	dispidSelectByCell = 30L,
	dispidColumns = 31L,
	dispidDataSource = 32L,
	dispidDataMember = 33L,
	dispidRowChanged = 34L,
	dispidHeaderHeight = 35L,
	dispidHeadForeColor = 1L,
	dispidHeadBackColor = 2L,
	dispidBackColor = 36L,
	dispidDividerColor = 3L,
	dispidBookmark = 37L,
	dispidSelBookmarks = 38L,
	dispidFrozenCols = 39L,
	dispidAddItem = 40L,
	dispidRemoveItem = 41L,
	dispidScroll = 42L,
	dispidRemoveAll = 43L,
	dispidMoveFirst = 44L,
	dispidMovePrevious = 45L,
	dispidMoveNext = 46L,
	dispidMoveLast = 47L,
	dispidMoveRecords = 48L,
	dispidIsAddRow = 49L,
	dispidUpdate = 50L,
	dispidCancelUpdate = 51L,
	dispidReBind = 52L,
	dispidRowBookmark = 53L,
	eventidBtnClick = 1L,
	eventidAfterColUpdate = 2L,
	eventidBeforeColUpdate = 3L,
	eventidBeforeInsert = 4L,
	eventidBeforeUpdate = 5L,
	eventidColResize = 6L,
	eventidRowResize = 7L,
	eventidRowColChange = 8L,
	eventidScroll = 9L,
	eventidChange = 10L,
	eventidAfterDelete = 11L,
	eventidBeforeDelete = 12L,
	eventidAfterUpdate = 13L,
	eventidAfterInsert = 14L,
	eventidColSwap = 15L,
	eventidUpdateError = 16L,
	eventidBeforeRowColChange = 17L,
	eventidScrollAfter = 18L,
	eventidBeforeDrawCell = 19L,
	//}}AFX_DISP_ID
	dispidCtrlType = 255L,
	};

	BEGIN_INTERFACE_PART(RowPositionChange, IRowPositionChange)
		INIT_INTERFACE_PART(CKCOMRichGridCtrl, RowPositionChange)
		STDMETHOD(OnRowPositionChange)(DBREASON, DBEVENTPHASE, BOOL);
	END_INTERFACE_PART(RowPositionChange)

	BEGIN_INTERFACE_PART(RowsetNotify, IRowsetNotify)
		INIT_INTERFACE_PART(CKCOMRichGridCtrl, RowsetNotify)
		STDMETHOD(OnFieldChange)(IRowset *, HROW, ULONG, ULONG[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowChange)(IRowset *, ULONG, const HROW[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowsetChange)(IRowset *, DBREASON, DBEVENTPHASE, BOOL);
	END_INTERFACE_PART(RowsetNotify)

	DECLARE_INTERFACE_MAP()

	BEGIN_INTERFACE_PART(DataSourceListener, IDataSourceListener)
		INIT_INTERFACE_PART(CKCOMRichGridCtrl, DataSourceListener)
		STDMETHOD(dataMemberChanged)(BSTR bstrDM);
		STDMETHOD(dataMemberAdded)(BSTR bstrDM);
		STDMETHOD(dataMemberRemoved)(BSTR bstrDM);
	END_INTERFACE_PART(DataSourceListener)

	friend class CGridColumn;
	friend class CGridColumns;
	friend class CGridInner;
	friend class CKCOMRichGridPropColumnPage;
	friend class CSelBookmarks;
	friend class CKCOMDrawingTools;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMRICHGRIDCTL_H__029288DA_2FFF_11D3_B446_0080C8F18522__INCLUDED)
