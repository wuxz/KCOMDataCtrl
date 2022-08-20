// KCOMMaskCtl.cpp : Implementation of the CKCOMMaskCtrl ActiveX Control class.

#include "stdafx.h"
#include <msstkppg.h>
#include "KCOMMask.h"
#include "KCOMMaskCtl.h"
#include "KCOMMaskPpg.h"
#include "MaskEdit.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CKCOMMaskCtrl, COleControl)

#define IDC_EDIT 1001
#define RELEASE(lpUnk) do \
	{ if ((lpUnk) != NULL) { (lpUnk)->Release(); (lpUnk) = NULL; } } while (0)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CKCOMMaskCtrl, COleControl)
	//{{AFX_MSG_MAP(CKCOMMaskCtrl)
	ON_WM_CREATE()
	ON_WM_CTLCOLOR()
	ON_WM_SETFOCUS()
	ON_WM_MOUSEACTIVATE()
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CKCOMMaskCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CKCOMMaskCtrl)
	DISP_PROPERTY_EX(CKCOMMaskCtrl, "PromptInclude", GetPromptInclude, SetPromptInclude, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMMaskCtrl, "AutoTab", GetAutoTab, SetAutoTab, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMMaskCtrl, "MaxLength", GetMaxLength, SetMaxLength, VT_I2)
	DISP_PROPERTY_EX(CKCOMMaskCtrl, "Mask", GetMask, SetMask, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMMaskCtrl, "SelLength", GetSelLength, SetSelLength, VT_I4)
	DISP_PROPERTY_EX(CKCOMMaskCtrl, "SelStart", GetSelStart, SetSelStart, VT_I4)
	DISP_PROPERTY_EX(CKCOMMaskCtrl, "SelText", GetSelText, SetSelText, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMMaskCtrl, "ClipText", GetClipText, SetClipText, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMMaskCtrl, "PromptChar", GetPromptChar, SetPromptChar, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMMaskCtrl, "Text", GetText, SetText, VT_BSTR)
	DISP_DEFVALUE(CKCOMMaskCtrl, "Text")
	DISP_STOCKFUNC_REFRESH()
	DISP_STOCKPROP_FONT()
	DISP_STOCKPROP_APPEARANCE()
	DISP_STOCKPROP_BACKCOLOR()
	DISP_STOCKPROP_FORECOLOR()
	DISP_STOCKPROP_BORDERSTYLE()
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CKCOMMaskCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CKCOMMaskCtrl, COleControl)
	//{{AFX_EVENT_MAP(CKCOMMaskCtrl)
	EVENT_CUSTOM("Change", FireChange, VTS_NONE)
	EVENT_STOCK_KEYDOWN()
	EVENT_STOCK_KEYPRESS()
	EVENT_STOCK_KEYUP()
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CKCOMMaskCtrl, 3)
	PROPPAGEID(CKCOMMaskPropPage::guid)
	PROPPAGEID(CLSID_StockColorPage)
	PROPPAGEID(CLSID_StockFontPage)
END_PROPPAGEIDS(CKCOMMaskCtrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CKCOMMaskCtrl, "KCOMMASK.KCOMMaskCtrl.1",
	0xd1728e25, 0x4985, 0x11d3, 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CKCOMMaskCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DKCOMMask =
		{ 0xd1728e23, 0x4985, 0x11d3, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const IID BASED_CODE IID_DKCOMMaskEvents =
		{ 0xd1728e24, 0x4985, 0x11d3, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwKCOMMaskOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CKCOMMaskCtrl, IDS_KCOMMASK, _dwKCOMMaskOleMisc)


/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskCtrl::CKCOMMaskCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CKCOMMaskCtrl

BOOL CKCOMMaskCtrl::CKCOMMaskCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Verify that your control follows apartment-model threading rules.
	// Refer to MFC TechNote 64 for more information.
	// If your control does not conform to the apartment-model rules, then
	// you must modify the code below, changing the 6th parameter from
	// afxRegApartmentThreading to 0.

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_KCOMMASK,
			IDB_KCOMMASK,
			afxRegApartmentThreading,
			_dwKCOMMaskOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskCtrl::CKCOMMaskCtrl - Constructor

CKCOMMaskCtrl::CKCOMMaskCtrl()
{
	InitializeIIDs(&IID_DKCOMMask, &IID_DKCOMMaskEvents);

	m_hFont = NULL;
	m_pEdit = NULL;
	m_pBrhBack = NULL;
	m_bIsInitializing = TRUE;

	for (int i = 0; i < 255; i ++)
		m_bIsDelimiter[i] = FALSE;
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskCtrl::~CKCOMMaskCtrl - Destructor

CKCOMMaskCtrl::~CKCOMMaskCtrl()
{
	if (m_pEdit)
	{
		m_pEdit->DestroyWindow();
		delete m_pEdit;
		m_pEdit = NULL;
	}

	IFont * pIfont = InternalGetFont().m_pFont;
	if (m_hFont)
	{
		pIfont->ReleaseHfont(m_hFont);
		m_hFont = NULL;
	}

	if (m_pBrhBack)
	{
		delete m_pBrhBack;
		m_pBrhBack = NULL;
	}
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskCtrl::OnDraw - Drawing function

void CKCOMMaskCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	if (AmbientUserMode())
		return;

	pdc->FillSolidRect(&rcBounds, TranslateColor(GetBackColor()));

	CFont * pOldFont = SelectStockFont(pdc);
	pdc->SetBkColor(TranslateColor(GetBackColor()));
	pdc->SetTextColor(TranslateColor(GetForeColor()));
	pdc->ExtTextOut(rcBounds.left, rcBounds.top, ETO_CLIPPED | ETO_OPAQUE, &rcBounds, 
		m_strMask, NULL);
	pdc->SelectObject(pOldFont);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskCtrl::DoPropExchange - Persistence support

void CKCOMMaskCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));

	ASSERT_POINTER(pPX, CPropExchange);
	ExchangeExtent(pPX);
	ExchangeStockProps(pPX);

	PX_Bool(pPX, "AutoTab", m_bAutoTab, TRUE);
	PX_String(pPX, "Mask", m_strMask, _T(""));
	PX_Short(pPX, "MaxLength", m_nMaxLength, 255);
	PX_String(pPX, "PromptChar", m_strPromptChar, _T("_"));
	PX_Bool(pPX, "PromptInclude", m_bPromptInclude, TRUE);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskCtrl::OnResetState - Reset control to default state

void CKCOMMaskCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	m_strInputMask.Empty();
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskCtrl::AboutBox - Display an "About" box to the user

void CKCOMMaskCtrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_KCOMMASK);
	dlgAbout.DoModal();
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskCtrl message handlers

BOOL CKCOMMaskCtrl::GetPromptInclude() 
{
	return m_bPromptInclude;
}

void CKCOMMaskCtrl::SetPromptInclude(BOOL bNewValue) 
{
	m_bPromptInclude = bNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidPromptInclude);
}

BOOL CKCOMMaskCtrl::GetAutoTab() 
{
	return m_bAutoTab;
}

void CKCOMMaskCtrl::SetAutoTab(BOOL bNewValue) 
{
	m_bAutoTab = bNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidAutoTab);
}

short CKCOMMaskCtrl::GetMaxLength() 
{
	return m_nMaxLength;
}

void CKCOMMaskCtrl::SetMaxLength(short nNewValue) 
{
	if (m_strInputMask.GetLength())
		return;

	if (nNewValue <= 0 || nNewValue > 255)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, AFX_IDP_E_INVALIDPROPERTYVALUE);

	m_nMaxLength = nNewValue;

	if (m_pEdit)
		m_pEdit->SetLimitText(m_nMaxLength);

	SetModifiedFlag();
	BoundPropertyChanged(dispidMaxLength);
}

BSTR CKCOMMaskCtrl::GetMask() 
{
	return m_strMask.AllocSysString();
}

void CKCOMMaskCtrl::SetMask(LPCTSTR lpszNewValue) 
{
	if (!IsValidMask(lpszNewValue))
		return;

	m_strMask = lpszNewValue;
	CalcMask();
	
	if (AmbientUserMode())
	{
		CalcText();
		MoveRight();
	}
	else
		InvalidateControl();

	if (m_strInputMask.GetLength())
	{
		m_nMaxLength = m_strInputMask.GetLength();

		if (m_pEdit)
			m_pEdit->SetLimitText(m_nMaxLength);

		BoundPropertyChanged(dispidMaxLength);
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidMask);
}

long CKCOMMaskCtrl::GetSelLength() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	int nStart = 0, nEnd = 0;
	m_pEdit->GetSel(nStart, nEnd);

	return nEnd - nStart + 1;
}

void CKCOMMaskCtrl::SetSelLength(long nNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	if (nNewValue < 0)
		return;

	int nStart = 0, nEnd = 0;
	m_pEdit->GetSel(nStart, nEnd);
	m_pEdit->SetSel(nStart, nStart + nNewValue - 1);
	MoveRight();
}

long CKCOMMaskCtrl::GetSelStart() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	int nStart = 0, nEnd = 0;
	m_pEdit->GetSel(nStart, nEnd);

	return nStart;
}

void CKCOMMaskCtrl::SetSelStart(long nNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	int nStart = 0, nEnd = 0;
	m_pEdit->GetSel(nStart, nEnd);
	
	if (nNewValue > nEnd)
		return;

	m_pEdit->SetSel(nNewValue, nEnd);
	MoveRight();
}

BSTR CKCOMMaskCtrl::GetSelText() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	CString strResult;
	int nStart = 0, nEnd = 0;

	m_pEdit->GetSel(nStart, nEnd);
	m_pEdit->GetWindowText(strResult);
	strResult = strResult.Mid(nStart, nEnd - nStart + 1);

	return strResult.AllocSysString();
}

void CKCOMMaskCtrl::SetSelText(LPCTSTR lpszNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	m_pEdit->ReplaceSel(lpszNewValue);

	CString strText;
	m_pEdit->GetWindowText(strText);
	SetDisplayText(strText);
	MoveRight();
}

BSTR CKCOMMaskCtrl::GetClipText() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	CString strResult, strText;
	m_pEdit->GetWindowText(strText);

	for (int i = 0; i < strText.GetLength(); i ++)
		if (!m_bIsDelimiter[i])
			strResult += strText[i];

	return strResult.AllocSysString();
}

void CKCOMMaskCtrl::SetClipText(LPCTSTR lpszNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	CString strResult, strText;

	if (m_strInputMask.GetLength() == 0)
	{
		strResult = lpszNewValue;
		strResult = strResult.Left(m_nMaxLength);
		m_pEdit->SetWindowText(strResult);
		
		if (!m_bIsInitializing)
			FireChange();

		return;
	}

	m_pEdit->GetWindowText(strText);
	
	int nTextLen = min(m_nMaxLength, strText.GetLength());

	for (int i = 0, j = 0; i < nTextLen; i ++)
	{
		if (m_bIsDelimiter[i])
		{
			strResult += m_strInputMask[i];
			continue;
		}
		else if (IsValidChar(strText[j], i))
			strResult += strText[j];
		else
			strResult += m_strPromptChar;
		
		j ++;
	}

	if (m_strInputMask.GetLength())
	{
		for (i = nTextLen; i < m_nMaxLength; i ++)
		{
			if (m_bIsDelimiter[i])
				strResult += m_strInputMask[i];
			else
				strResult += m_strPromptChar;
		}
	}

	m_pEdit->SetWindowText(strText);

	if (!m_bIsInitializing)
			FireChange();
}

BSTR CKCOMMaskCtrl::GetPromptChar() 
{
	return m_strPromptChar.AllocSysString();
}

void CKCOMMaskCtrl::SetPromptChar(LPCTSTR lpszNewValue) 
{
	if (lstrlen(lpszNewValue) != 1)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, AFX_IDP_E_INVALIDPROPERTYVALUE);

	m_strPromptChar = lpszNewValue;
	
	if (AmbientUserMode())
		CalcText();

	BoundPropertyChanged(dispidPromptChar);
}

BSTR CKCOMMaskCtrl::GetText() 
{
	if (!AmbientUserMode())
		SetNotSupported();

	CString strResult, strText;
	m_pEdit->GetWindowText(strText);

	for (int i = 0; i < strText.GetLength(); i ++)
		if (m_bPromptInclude || !m_bIsDelimiter[i])
			strResult += strText[i];

	return strResult.AllocSysString();
}

void CKCOMMaskCtrl::SetText(LPCTSTR lpszNewValue) 
{
	if (m_bPromptInclude)
		SetDisplayText(lpszNewValue);
	else
		SetClipText(lpszNewValue);
}

void CKCOMMaskCtrl::OnFontChanged() 
{
	HFONT hFont;
	IFont * pIfont = InternalGetFont().m_pFont;
	
	if (pIfont->get_hFont(&hFont) == S_OK)
	{
		//release previous IFont interface pointer
		if (m_hFont)
			pIfont->ReleaseHfont(m_hFont);

		m_hFont = hFont;

		if (m_pEdit)
			m_pEdit->SendMessage(WM_SETFONT, (WPARAM)m_hFont, MAKELPARAM(TRUE, 0));
	}

	COleControl::OnFontChanged();
}

BOOL CKCOMMaskCtrl::OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip) 
{
	if (AmbientUserMode() && m_pEdit)
		m_pEdit->SetWindowPos(&wndTop, 0, 0, lpRectPos->right - lpRectPos->left, 
		lpRectPos->bottom - lpRectPos->top, SWP_NOZORDER);
	
	return COleControl::OnSetObjectRects(lpRectPos, lpRectClip);
}

BOOL CKCOMMaskCtrl::PreTranslateMessage(MSG* pMsg) 
{
	USHORT nCharShort;
	short nShiftState = _AfxShiftState();

	if (pMsg->hwnd == m_hWnd || ::IsChild(m_hWnd, pMsg->hwnd))
	switch (pMsg->message)
	{
	case WM_KEYDOWN:
		nCharShort = (USHORT)pMsg->wParam;
		FireKeyDown(&nCharShort, nShiftState);
		if (nCharShort == 0)
			return TRUE;

		pMsg->wParam = MAKEWPARAM(nCharShort, HIWORD(pMsg->wParam));

		if (nCharShort == VK_LEFT || nCharShort == VK_RIGHT || nCharShort == VK_UP || 
			nCharShort == VK_DOWN || nCharShort == VK_DELETE || (!m_bAutoTab 
			&& nCharShort == VK_TAB))
		{
			GetFocus()->SendMessage(pMsg->message, pMsg->wParam, pMsg->lParam);
			return TRUE;
		}

		break;

	case WM_KEYUP:
		nCharShort = (USHORT)pMsg->wParam;
		FireKeyUp(&nCharShort, nShiftState);
		if (nCharShort == 0)
			return TRUE;

		pMsg->wParam = MAKEWPARAM(nCharShort, HIWORD(pMsg->wParam));

		if (nCharShort == VK_LEFT || nCharShort == VK_RIGHT || nCharShort == VK_UP || 
			nCharShort == VK_DOWN || nCharShort == VK_DELETE || (!m_bAutoTab 
			&& nCharShort == VK_TAB))
		{
			GetFocus()->SendMessage(pMsg->message, pMsg->wParam, pMsg->lParam);
			return TRUE;
		}
		
		break;

	case WM_CHAR:
		nCharShort = (USHORT)pMsg->wParam;
		FireKeyPress(&nCharShort);
		if (nCharShort == 0)
			return TRUE;

		pMsg->wParam = MAKEWPARAM(nCharShort, HIWORD(pMsg->wParam));
		
		break;
	}
	
	return COleControl::PreTranslateMessage(pMsg);
}

int CKCOMMaskCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (COleControl::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	if (AmbientUserMode())
	{
		CRect rect;
		GetClientRect(&rect);

		m_pEdit = new CMaskEdit(this);
		m_pEdit->Create(ES_AUTOHSCROLL | WS_VISIBLE | WS_CHILD, rect, this, IDC_EDIT);
		CalcMask();
		
		m_bIsInitializing = TRUE;
		CalcText();
		m_bIsInitializing = FALSE;

		if (m_pEdit)
			m_pEdit->SetLimitText(m_nMaxLength);
	}
	
	return 0;
}

HBRUSH CKCOMMaskCtrl::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor) 
{
	if (m_pEdit && pWnd->m_hWnd == m_pEdit->m_hWnd && nCtlColor == CTLCOLOR_EDIT)
	{
		if (m_pBrhBack)
		{
			delete m_pBrhBack;
			m_pBrhBack = NULL;
		}

		m_pBrhBack = new CBrush(TranslateColor(GetBackColor()));
		pDC->SetTextColor(TranslateColor(GetForeColor()));
		pDC->SetBkColor(TranslateColor(GetBackColor()));

		return (HBRUSH)*m_pBrhBack;
	}

	HBRUSH hbr = COleControl::OnCtlColor(pDC, pWnd, nCtlColor);

	return hbr;
}

void CKCOMMaskCtrl::OnSetFocus(CWnd* pOldWnd) 
{
	COleControl::OnSetFocus(pOldWnd);

	if (m_pEdit)
	{
		m_pEdit->SetFocus();
		MoveRight();
	}
}

void CKCOMMaskCtrl::SetDisplayText(CString strText)
{
	CString strResult;
	int i, nTextLen;

	if (m_strInputMask.GetLength() == 0)
	{
		strResult = strText.Left(m_nMaxLength);

		goto exit;
	}

	nTextLen = min(m_nMaxLength, strText.GetLength());

	for (i = 0; i < nTextLen; i ++)
	{
		if (m_bIsDelimiter[i])
			strResult += m_strInputMask[i];
		else if (IsValidChar(strText[i], i))
			strResult += strText[i];
		else
			strResult += m_strPromptChar;
	}

	if (m_strInputMask.GetLength())
	{
		for (i = nTextLen; i < m_nMaxLength; i ++)
		{
			if (m_bIsDelimiter[i])
				strResult += m_strInputMask[i];
			else
				strResult += m_strPromptChar;
		}
	}

exit:

	int nStart = 0, nEnd = 0;
	m_pEdit->GetSel(nStart, nEnd);
	m_pEdit->SetWindowText(strResult);
	m_pEdit->SetSel(nStart, nEnd);
	MoveRight();

	if (!m_bIsInitializing)
		FireChange();
}

void CKCOMMaskCtrl::CalcText()
{
	CString strText, strResult;
	m_pEdit->GetWindowText(strText);

	SetDisplayText(strText);
}

void CKCOMMaskCtrl::MoveRight()
{
	int nStart = 0, nEnd = 0;
	m_pEdit->GetSel(nStart, nEnd);

	if (nEnd > m_nMaxLength || m_bIsDelimiter[nEnd])
	{
		m_pEdit->MoveRight();
		m_pEdit->GetSel(nStart, nEnd);
		if (nEnd > m_nMaxLength || m_bIsDelimiter[nEnd])
			m_pEdit->MoveLeft();
	}
}

BOOL CKCOMMaskCtrl::IsValidMask(CString strMask)
{
	if (strMask.IsEmpty())
		return TRUE;

	if (strMask.GetLength() > 255)
		return FALSE;

	char cCurrChr;

	for (int i = 0; i < strMask.GetLength(); i++)
	{
		cCurrChr = strMask[i];

		if (cCurrChr == _T('#') || cCurrChr == _T('9') || cCurrChr == _T('?') ||
			cCurrChr == _T('C') || cCurrChr == _T('A') || cCurrChr == _T('a'))
			return TRUE;
	}

	return FALSE;
}

BOOL CKCOMMaskCtrl::IsValidChar(TCHAR nChar, int nPosition)
{
	if (nPosition >= m_nMaxLength)
		return FALSE;

	if (m_strInputMask.GetLength() == 0 || m_bIsDelimiter[nPosition])
		return TRUE;

	switch (m_strInputMask[nPosition])
	{
	case _T('#'):
		return nChar >= '0' && nChar <= '9';
		break;

	case _T('9'):
		return nChar == ' ' || (nChar >= '0' && nChar <= '9');
		break;

	case _T('?'):
		return (nChar >= 'a' && nChar <= 'z') ||
			(nChar >= 'A' && nChar <= 'Z');
		break;

	case _T('C'):
		return nChar == ' ' || (nChar >= 'a' && nChar <= 'z') 
			|| (nChar >= 'A' && nChar <= 'Z');
		break;

	case _T('A'):
		return (nChar >= 'a' && nChar <= 'z') ||
			(nChar >= 'A' && nChar <= 'Z') ||
			(nChar >= '0' && nChar <= '9');
		break;

	case _T('a'):
		return (nChar >= 32 && nChar <= 126) ||
			(nChar >= 128 && nChar <= 255);
		break;
	};

	return FALSE;
}

void CKCOMMaskCtrl::CalcMask()
{
	BOOL bDelimiter = FALSE;
	char cCurrChr;
	int i, j;
	CString strInit;

	m_strInputMask.Empty();
	for (i = 0; i < 255; i++)
		m_bIsDelimiter[i] = FALSE;

	if (m_strMask.IsEmpty())
		return;

	for (i = 0, j = 0; i < m_strMask.GetLength(); i++)
	{
		cCurrChr = m_strMask[i];
		if (!bDelimiter && cCurrChr == '/')
		{
			bDelimiter = TRUE;
			continue;
		}

		if (cCurrChr != _T('#') && cCurrChr != _T('9') && cCurrChr != _T('?') && 
			cCurrChr != _T('C') && cCurrChr != _T('A') && cCurrChr != _T('a'))
			bDelimiter = TRUE;

		m_strInputMask += cCurrChr;
		
		if (bDelimiter)
			m_bIsDelimiter[j] = TRUE;

		j ++;
		bDelimiter = FALSE;
	}

	m_nMaxLength = m_strInputMask.GetLength();
}

void CKCOMMaskCtrl::ExchangeStockProps(CPropExchange *pPX)
{
	BOOL bLoading = pPX->IsLoading();
	DWORD dwStockPropMask = GetStockPropMask();
	DWORD dwPersistMask = dwStockPropMask;

	PX_ULong(pPX, _T("_StockProps"), dwPersistMask);

	if (dwStockPropMask & (STOCKPROP_CAPTION | STOCKPROP_TEXT))
	{
		CString strText;

		if (dwPersistMask & (STOCKPROP_CAPTION | STOCKPROP_TEXT))
		{
			if (!bLoading)
				strText = InternalGetText();
			if (dwStockPropMask & STOCKPROP_CAPTION)
				PX_String(pPX, _T("Caption"), strText, _T(""));
			if (dwStockPropMask & STOCKPROP_TEXT)
				PX_String(pPX, _T("Text"), strText, _T(""));
		}
		if (bLoading)
		{
			TRY
				SetText(strText);
			END_TRY
		}
	}

	if (dwStockPropMask & STOCKPROP_FORECOLOR)
	{
		if (dwPersistMask & STOCKPROP_FORECOLOR)
			PX_Color(pPX, _T("ForeColor"), m_clrForeColor, AmbientForeColor());
		else if (bLoading)
			m_clrForeColor = AmbientForeColor();
	}

	if (dwStockPropMask & STOCKPROP_BACKCOLOR)
	{
		if (dwPersistMask & STOCKPROP_BACKCOLOR)
			PX_Color(pPX, _T("BackColor"), m_clrBackColor, GetSysColor(COLOR_WINDOW));
		else if (bLoading)
			m_clrBackColor = GetSysColor(COLOR_WINDOW);
	}

	if (dwStockPropMask & STOCKPROP_FONT)
	{
		LPFONTDISP pFontDispAmbient = AmbientFont();
		BOOL bChanged = TRUE;

		if (dwPersistMask & STOCKPROP_FONT)
			bChanged = PX_Font(pPX, _T("Font"), m_font, NULL, pFontDispAmbient);
		else if (bLoading)
			m_font.InitializeFont(NULL, pFontDispAmbient);

		if (bLoading && bChanged)
			OnFontChanged();

		RELEASE(pFontDispAmbient);
	}

	if (dwStockPropMask & STOCKPROP_BORDERSTYLE)
	{
		short sBorderStyle = m_sBorderStyle ? 1 : 0;

		if (dwPersistMask & STOCKPROP_BORDERSTYLE)
			PX_Short(pPX, _T("BorderStyle"), m_sBorderStyle, 1);
		else if (bLoading)
			m_sBorderStyle = 1;

		if (sBorderStyle != m_sBorderStyle)
			_AfxToggleBorderStyle(this);
	}

	if (dwStockPropMask & STOCKPROP_ENABLED)
	{
		BOOL bEnabled = m_bEnabled;

		if (dwPersistMask & STOCKPROP_ENABLED)
			PX_Bool(pPX, _T("Enabled"), m_bEnabled, TRUE);
		else if (bLoading)
			m_bEnabled = TRUE;

		if ((bEnabled != m_bEnabled) && (m_hWnd != NULL))
			::EnableWindow(m_hWnd, m_bEnabled);
	}

	if (dwStockPropMask & STOCKPROP_APPEARANCE)
	{
		short sAppearance = m_sAppearance ? 1 : 0;

		if (dwPersistMask & STOCKPROP_APPEARANCE)
			PX_Short(pPX, _T("Appearance"), m_sAppearance, 1);
		else if (bLoading)
			m_sAppearance = AmbientAppearance();

		if (sAppearance != m_sAppearance)
			_AfxToggleAppearance(this);
	}
}

BOOL CKCOMMaskCtrl::OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray) 
{
	if (dispid == DISPID_APPEARANCE)
	{
		CString strAppearance;
		strAppearance.LoadString(IDS_APPEARANCE_FLAT);
		pStringArray->Add(strAppearance);
		strAppearance.LoadString(IDS_APPEARANCE_3D);
		pStringArray->Add(strAppearance);

		pCookieArray->Add(0);
		pCookieArray->Add(1);

		return TRUE;
	}
	
	return COleControl::OnGetPredefinedStrings(dispid, pStringArray, pCookieArray);
}

BOOL CKCOMMaskCtrl::OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut) 
{
	if (dispid == DISPID_APPEARANCE)
	{
		VariantInit(lpvarOut);
		lpvarOut->vt = VT_I4;
		lpvarOut->lVal = dwCookie;

		return TRUE;
	}
	
	return COleControl::OnGetPredefinedValue(dispid, dwCookie, lpvarOut);
}

BOOL CKCOMMaskCtrl::OnGetDisplayString(DISPID dispid, CString& strValue) 
{
	if (dispid == DISPID_APPEARANCE)
	{
		strValue.LoadString(GetAppearance() ? IDS_APPEARANCE_3D : IDS_APPEARANCE_FLAT);

		return TRUE;
	}

	return COleControl::OnGetDisplayString(dispid, strValue);
}

int CKCOMMaskCtrl::OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message) 
{
	if (AmbientUserMode())
		OnActivateInPlace (TRUE, NULL);
	
	return COleControl::OnMouseActivate(pDesktopWnd, nHitTest, message);
}
