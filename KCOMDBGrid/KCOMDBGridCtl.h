#if !defined(AFX_KCOMDBGRIDCTL_H__AC212654_FBE2_11D2_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_KCOMDBGRIDCTL_H__AC212654_FBE2_11D2_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMDBGridCtl.h : Declaration of the CKCOMDBGridCtrl ActiveX Control class.

/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridCtrl : See KCOMDBGridCtl.cpp for implementation.

#include <msdatsrc.h>
#include <afxtempl.h>
#include <atldbcli.h>

struct BindData
{
	BYTE bookmark[4];
	TCHAR bstrData[255][255];
	DWORD dwStatus[255];
};

struct RowData
{
	TCHAR bstrData[255][255];
	DWORD dwStatus[255];
};

struct ColumnData
{
	CString strName;
	int nColumn;
	VARTYPE vt;
};

class CGridInner;

class CKCOMDBGridCtrl : public COleControl
{
	DECLARE_DYNCREATE(CKCOMDBGridCtrl)

// Constructor
public:
	CKCOMDBGridCtrl();
	CAccessorRowset<CManualAccessor> m_Rowset;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKCOMDBGridCtrl)
	public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	virtual BOOL OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip);
	virtual void OnForeColorChanged();
	virtual BOOL OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray);
	virtual BOOL OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut);
	virtual BOOL OnGetDisplayString(DISPID dispid, CString& strValue);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual void OnSetClientSite();
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	BOOL m_bShowRecordSelector;
	BOOL m_bDataSrcChanged;
	IDataSource * m_spDataSource;
	CString m_strDataMember;
	CArray<ColumnData*, ColumnData*> m_apColumnData;
	CArray<BYTE*, BYTE*> m_apBookmark;
	IRowPosition * m_spRowPosition;
	DWORD m_dwCookieRPC, m_dwCookieRN;
	BOOL m_bHaveBookmarks;
	int m_nColumns, m_nColumnsBind;
	DBCOLUMNINFO * m_pColumnInfo;
	LPOLESTR m_pStrings;
	BindData m_data;
	int m_nColumnBindNum[255];
	char m_strColumnHeader[255][40];
	int m_nBookmarkSize;
	BOOL m_bIsPosSync;
	BOOL m_bReadOnly;
	BOOL m_bEndReached;
	CString m_strFieldSeparator;

protected:
	CGridInner * m_pGridInner;

protected:
	CString m_strCaption;
	virtual BOOL IsColEditable(int nCol);
	HINSTANCE m_hResHandle;
	HINSTANCE m_hResPrev;
	HGLOBAL m_hBlob;
	void SerializeBindInfo(CArchive & ar, BOOL bLoad = TRUE);
	BOOL GetColumnInfo();
	long m_nGridLineStyle;
	long m_nRows;
	long m_nCols;
	long m_nDataMode;
	long m_nRow;
	CFontHolder m_fntCell;
	void InitCellFont();
	void InitHeaderFont();
	CFontHolder m_fntHeader;
	long m_nDefColWidth;
	long m_nCol;
	BOOL m_bAllowRowSizing;
	BOOL m_bAllowArrowKeys;
	long m_nRowHeight;
	long m_nHeaderHeight;
	HRESULT UpdateRowData(int nRow);
	BOOL m_bAllowDelete;
	BOOL m_bAllowAddNew;
	HRESULT DeleteRowData(int nRow = 0);
	void FreeColumnMemory();
	void ClearInterfaces();
	void UndeleteRecordFromHRow(HROW hRow);
	void UpdateRowFromSrc(int nRow);
	void DeleteRowFromSrc(int nRow);
	void UpdateCellFromSrc(int nRow, int nCol);
//	HRESULT ModifyCell(int nRow, int nCol, CString str);
	void FetchNewRows();
	void InitGridHeader();
	void GetItemText(int nRow, int nCol);
	HRESULT GetRowset();
	int GetRowFromBmk(BYTE * bmk);
	int GetRowFromHRow(HROW hRow);
	void UpdateControl();
	virtual HINSTANCE GetResourceHandle(LCID lcid);

	~CKCOMDBGridCtrl();

	BEGIN_OLEFACTORY(CKCOMDBGridCtrl)        // Class factory and guid
		virtual BOOL VerifyUserLicense();
		virtual BOOL GetLicenseKey(DWORD, BSTR FAR*);
	END_OLEFACTORY(CKCOMDBGridCtrl)

	DECLARE_OLETYPELIB(CKCOMDBGridCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CKCOMDBGridCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CKCOMDBGridCtrl)		// Type name and misc status

// Message maps
	//{{AFX_MSG(CKCOMDBGridCtrl)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg int OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	//{{AFX_DISPATCH(CKCOMDBGridCtrl)
	OLE_COLOR m_clrHeader;
	afx_msg void OnHeaderBackColorChanged();
	OLE_COLOR m_clrBack;
	afx_msg void OnBackColorChanged();
	afx_msg LPUNKNOWN GetDataSource();
	afx_msg void SetDataSource(LPUNKNOWN newValue);
	afx_msg BSTR GetDataMember();
	afx_msg void SetDataMember(LPCTSTR lpszNewValue);
	afx_msg BOOL GetShowRecordSelector();
	afx_msg void SetShowRecordSelector(BOOL bNewValue);
	afx_msg BOOL GetAllowAddNew();
	afx_msg void SetAllowAddNew(BOOL bNewValue);
	afx_msg BOOL GetAllowDelete();
	afx_msg void SetAllowDelete(BOOL bNewValue);
	afx_msg BOOL GetReadOnly();
	afx_msg void SetReadOnly(BOOL bNewValue);
	afx_msg long GetHeaderHeight();
	afx_msg void SetHeaderHeight(long nNewValue);
	afx_msg long GetRowHeight();
	afx_msg void SetRowHeight(long nNewValue);
	afx_msg BOOL GetAllowArrowKeys();
	afx_msg void SetAllowArrowKeys(BOOL bNewValue);
	afx_msg BOOL GetAllowRowSizing();
	afx_msg void SetAllowRowSizing(BOOL bNewValue);
	afx_msg long GetCol();
	afx_msg void SetCol(long nNewValue);
	afx_msg long GetDefColWidth();
	afx_msg void SetDefColWidth(long nNewValue);
	afx_msg LPFONTDISP GetHeaderFont();
	afx_msg void SetHeaderFont(LPFONTDISP newValue);
	afx_msg LPFONTDISP GetFont();
	afx_msg void SetFont(LPFONTDISP newValue);
	afx_msg long GetRow();
	afx_msg void SetRow(long nNewValue);
	afx_msg long GetVisibleCols();
	afx_msg long GetVisibleRows();
	afx_msg long GetDataMode();
	afx_msg void SetDataMode(long nNewValue);
	afx_msg long GetCols();
	afx_msg void SetCols(long nNewValue);
	afx_msg long GetRows();
	afx_msg void SetRows(long nNewValue);
	afx_msg long GetGridLineStyle();
	afx_msg void SetGridLineStyle(long nNewValue);
	afx_msg BSTR GetFieldSeparator();
	afx_msg void SetFieldSeparator(LPCTSTR lpszNewValue);
	afx_msg BSTR GetCaption();
	afx_msg void SetCaption(LPCTSTR lpszNewValue);
	afx_msg long ColContaining(long x);
	afx_msg long RowContaining(long y);
	afx_msg void Scroll(long Cols, long Rows);
	afx_msg void InsertRow(const VARIANT FAR& RowBefore);
	afx_msg void InsertCol(const VARIANT FAR& ColBefore);
	afx_msg void DeleteRow(const VARIANT FAR& RowIndex);
	afx_msg void DeleteCol(const VARIANT FAR& ColIndex);
	afx_msg long GetColWidth(short ColIndex);
	afx_msg void SetColWidth(short ColIndex, long nNewValue);
	afx_msg BSTR GetCellText(short RowIndex, short ColIndex);
	afx_msg void SetCellText(short RowIndex, short ColIndex, LPCTSTR lpszNewValue);
	afx_msg BSTR GetRowText(short RowIndex);
	afx_msg void SetRowText(short RowIndex, LPCTSTR lpszNewValue);
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

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
	//{{AFX_EVENT(CKCOMDBGridCtrl)
	void FireAfterColEdit(short ColIndex)
		{FireEvent(eventidAfterColEdit,EVENT_PARAM(VTS_I2), ColIndex);}
	void FireAfterColUpdate(short ColIndex)
		{FireEvent(eventidAfterColUpdate,EVENT_PARAM(VTS_I2), ColIndex);}
	void FireAfterDelete()
		{FireEvent(eventidAfterDelete,EVENT_PARAM(VTS_NONE));}
	void FireAfterInsert()
		{FireEvent(eventidAfterInsert,EVENT_PARAM(VTS_NONE));}
	void FireAfterUpdate()
		{FireEvent(eventidAfterUpdate,EVENT_PARAM(VTS_NONE));}
	void FireBeforeColEdit(short ColIndex, short KeyAscii, short FAR* Cancel)
		{FireEvent(eventidBeforeColEdit,EVENT_PARAM(VTS_I2  VTS_I2  VTS_PI2), ColIndex, KeyAscii, Cancel);}
	void FireBeforeDelete(short FAR* Cancel)
		{FireEvent(eventidBeforeDelete,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireBeforeUpdate(short FAR* Cancel)
		{FireEvent(eventidBeforeUpdate,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireChange()
		{FireEvent(eventidChange,EVENT_PARAM(VTS_NONE));}
	void FireColResize(short ColIndex, short FAR* Cancel)
		{FireEvent(eventidColResize,EVENT_PARAM(VTS_I2  VTS_PI2), ColIndex, Cancel);}
	void FireBeforeInsert(short FAR* Cancel)
		{FireEvent(eventidBeforeInsert,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireHeadClick(short ColIndex)
		{FireEvent(eventidHeadClick,EVENT_PARAM(VTS_I2), ColIndex);}
	void FireRolColChange()
		{FireEvent(eventidRolColChange,EVENT_PARAM(VTS_NONE));}
	void FireRowResize(short FAR* Cancel)
		{FireEvent(eventidRowResize,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireScroll(short FAR* Cancel)
		{FireEvent(eventidScroll,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireBeforeColUpdate(short ColIndex, VARIANT FAR* OldValue, short FAR* Cancel)
		{FireEvent(eventidBeforeColUpdate,EVENT_PARAM(VTS_I2  VTS_PVARIANT  VTS_PI2), ColIndex, OldValue, Cancel);}
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	void ScrollToRow(int nRow);
	void SetFocusCell(int nRow, int nCol);
	enum {
	//{{AFX_DISP_ID(CKCOMDBGridCtrl)
	dispidDataSource = 3L,
	dispidDataMember = 4L,
	dispidShowRecordSelector = 5L,
	dispidAllowAddNew = 6L,
	dispidAllowDelete = 7L,
	dispidReadOnly = 8L,
	dispidHeaderHeight = 9L,
	dispidRowHeight = 10L,
	dispidAllowArrowKeys = 11L,
	dispidAllowRowSizing = 12L,
	dispidHeaderBackColor = 1L,
	dispidCol = 13L,
	dispidDefColWidth = 14L,
	dispidHeaderFont = 15L,
	dispidFont = 16L,
	dispidRow = 17L,
	dispidVisibleCols = 18L,
	dispidVisibleRows = 19L,
	dispidDataMode = 20L,
	dispidCols = 21L,
	dispidRows = 22L,
	dispidGridLineStyle = 23L,
	dispidBackColor = 2L,
	dispidFieldSeparator = 24L,
	dispidCaption = 25L,
	dispidColContaining = 26L,
	dispidRowContaining = 27L,
	dispidScroll = 28L,
	dispidInsertRow = 29L,
	dispidInsertCol = 30L,
	dispidDeleteRow = 31L,
	dispidDeleteCol = 32L,
	dispidColWidth = 33L,
	dispidCellText = 34L,
	dispidRowText = 35L,
	eventidAfterColEdit = 1L,
	eventidAfterColUpdate = 2L,
	eventidAfterDelete = 3L,
	eventidAfterInsert = 4L,
	eventidAfterUpdate = 5L,
	eventidBeforeColEdit = 6L,
	eventidBeforeDelete = 7L,
	eventidBeforeUpdate = 8L,
	eventidChange = 9L,
	eventidColResize = 10L,
	eventidBeforeInsert = 11L,
	eventidHeadClick = 12L,
	eventidRolColChange = 13L,
	eventidRowResize = 14L,
	eventidScroll = 15L,
	eventidBeforeColUpdate = 16L,
	//}}AFX_DISP_ID
	};
	BEGIN_INTERFACE_PART(RowPositionChange, IRowPositionChange)
		INIT_INTERFACE_PART(CKCOMDBGridCtrl, RowPositionChange)
		STDMETHOD(OnRowPositionChange)(DBREASON, DBEVENTPHASE, BOOL);
	END_INTERFACE_PART(RowPositionChange)

	BEGIN_INTERFACE_PART(RowsetNotify, IRowsetNotify)
		INIT_INTERFACE_PART(CKCOMDBGridCtrl, RowsetNotify)
		STDMETHOD(OnFieldChange)(IRowset *, HROW, ULONG, ULONG[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowChange)(IRowset *, ULONG, const HROW[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowsetChange)(IRowset *, DBREASON, DBEVENTPHASE, BOOL);
	END_INTERFACE_PART(RowsetNotify)

	DECLARE_INTERFACE_MAP()

	BEGIN_INTERFACE_PART(DataSourceListener, IDataSourceListener)
		INIT_INTERFACE_PART(CKCOMDBGridCtrl, DataSourceListener)
		STDMETHOD(dataMemberChanged)(BSTR bstrDM);
		STDMETHOD(dataMemberAdded)(BSTR bstrDM);
		STDMETHOD(dataMemberRemoved)(BSTR bstrDM);
	END_INTERFACE_PART(DataSourceListener)

	friend class CGridInner;
	friend class CKCOMDBGridColumnsPpg;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMDBGRIDCTL_H__AC212654_FBE2_11D2_A7FE_0080C8763FA4__INCLUDED)
