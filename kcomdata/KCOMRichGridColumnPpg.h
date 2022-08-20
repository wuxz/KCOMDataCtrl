#if !defined(AFX_KCOMRICHGRIDCOLUMNPPG_H__0CFC2334_12A9_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_KCOMRICHGRIDCOLUMNPPG_H__0CFC2334_12A9_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMRichGridColumnPpg.h : Declaration of the CKCOMRichGridPropColumnPage property page class.

////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridPropColumnPage : See KCOMRichGridColumnPpg.cpp for implementation.
#include "GridInPpg.h"

class CKCOMRichGridCtrl;
class CKCOMRichDropDownCtrl;
class CKCOMRichComboCtrl;

class CKCOMRichGridPropColumnPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CKCOMRichGridPropColumnPage)
	DECLARE_OLECREATE_EX(CKCOMRichGridPropColumnPage)

// Constructor
public:
	void InsertCol(ColumnProp * pColSource);
	void CheckRadio(CButton * pButton, int nValue);
	CKCOMRichGridPropColumnPage();
	~CKCOMRichGridPropColumnPage();

// Dialog Data
	//{{AFX_DATA(CKCOMRichGridPropColumnPage)
	enum { IDD = IDD_PROPPAGE_KCOMRICHGRID_COLUMN };
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

protected:
	CArray<ColumnProp *, ColumnProp *> m_apColumnProp;
	CKCOMRichGridCtrl * m_pGridCtrl;
	CKCOMRichDropDownCtrl * m_pDropDownCtrl;
	CKCOMRichComboCtrl * m_pComboCtrl;

	CGridInPpg m_wndGrid;
	int m_nColumns;
	int m_nCurrentCol;

protected:

// Message maps
protected:
	int CalcPropertyOrdinal(int nProperty);
	CString FindUniqueName(CString& strRecm);
	void UpdateStyle(ColumnProp * pCol, ROWCOL nCol);
	void HideControls();
	void InitControlsVars();
	//{{AFX_MSG(CKCOMRichGridPropColumnPage)
	virtual BOOL OnInitDialog();
	afx_msg void OnSelchangeListProperty();
	afx_msg void OnSelendokComboColumn();
	afx_msg void OnButtonBackcolor();
	afx_msg void OnButtonForecolor();
	afx_msg void OnButtonHeadbackcolor();
	afx_msg void OnButtonHeadforecolor();
	afx_msg void OnCheckAllowsizing();
	afx_msg void OnCheckLocked();
	afx_msg void OnCheckNullable();
	afx_msg void OnCheckVisible();
	afx_msg void OnKillfocusEditName();
	afx_msg void OnKillfocusEditCaption();
	afx_msg void OnKillfocusEditFieldlen();
	afx_msg void OnKillfocusEditMask();
	afx_msg void OnKillfocusEditPromptchar();
	afx_msg void OnKillfocusEditWidth();
	afx_msg void OnRadioAlignmentCenter();
	afx_msg void OnRadioAlignmentLeft();
	afx_msg void OnRadioAlignmentRight();
	afx_msg void OnRadioCaptionalignmentAsalignment();
	afx_msg void OnRadioCaptionalignmentCenter();
	afx_msg void OnRadioCaptionalignmentLeft();
	afx_msg void OnRadioCaptionalignmentRight();
	afx_msg void OnRadioCaseLower();
	afx_msg void OnRadioCaseNone();
	afx_msg void OnRadioCaseUpper();
	afx_msg void OnRadioDatatypeBoolean();
	afx_msg void OnRadioDatatypeByte();
	afx_msg void OnRadioDatatypeCurrency();
	afx_msg void OnRadioDatatypeDate();
	afx_msg void OnRadioDatatypeInteger();
	afx_msg void OnRadioDatatypeLong();
	afx_msg void OnRadioDatatypeSingle();
	afx_msg void OnRadioDatatypeText();
	afx_msg void OnRadioStyleButton();
	afx_msg void OnRadioStyleCheckbox();
	afx_msg void OnRadioStyleCombobox();
	afx_msg void OnRadioStyleEdit();
	afx_msg void OnRadioStyleEditbutton();
	afx_msg void OnButtonSetupcombobox();
	afx_msg void OnButtonAddcolumn();
	afx_msg void OnButtonDeletecolumn();
	afx_msg void OnButtonReset();
	afx_msg void OnCheckPromptinclude();
	afx_msg void OnSelendokComboDatafield();
	afx_msg void OnButtonFields();
	afx_msg void OnSelendokComboDatemask();
	afx_msg void OnSelendokComboComboboxstyle();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	
	friend class CGridInPpg;
	friend class CAddColumnDlg;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMRICHGRIDCOLUMNPPG_H__0CFC2334_12A9_11D3_A7FE_0080C8763FA4__INCLUDED)
