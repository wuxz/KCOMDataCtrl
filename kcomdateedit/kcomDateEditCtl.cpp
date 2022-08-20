// EditboxCtl.cpp : Implementation of the CKCOMDateEditCtrl ActiveX Control class.

#include "stdafx.h"
#include "KCOMDateEdit.h"
#include "KCOMDateEditCtl.h"
#include "KCOMDateEditPpg.h"
#include "DlgDatePicker.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define RELEASE(lpUnk) do \
	{ if ((lpUnk) != NULL) { (lpUnk)->Release(); (lpUnk) = NULL; } } while (0)

IMPLEMENT_DYNCREATE(CKCOMDateEditCtrl, COleControl)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CKCOMDateEditCtrl, COleControl)
	//{{AFX_MSG_MAP(CKCOMDateEditCtrl)
	ON_WM_SETFOCUS()
	ON_WM_KILLFOCUS()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONDBLCLK()
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CKCOMDateEditCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CKCOMDateEditCtrl)
	DISP_PROPERTY_NOTIFY(CKCOMDateEditCtrl, "ReadOnly", m_bReadOnly, OnReadOnlyChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CKCOMDateEditCtrl, "CanYearNegative", m_bCanYearNegative, OnCanYearNegativeChanged, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMDateEditCtrl, "DateFormat", GetDateFormat, SetDateFormat, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMDateEditCtrl, "DisplayDateFormat", GetDisplayDateFormat, SetDisplayDateFormat, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMDateEditCtrl, "Text", GetText, SetText, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMDateEditCtrl, "TextAlign", GetTextAlign, SetTextAlign, VT_I2)
	DISP_FUNCTION(CKCOMDateEditCtrl, "SetDate", SetDate, VT_EMPTY, VTS_I2 VTS_I2 VTS_I2)
	DISP_DEFVALUE(CKCOMDateEditCtrl, "Text")
	DISP_STOCKPROP_FONT()
	DISP_STOCKPROP_BORDERSTYLE()
	DISP_STOCKPROP_FORECOLOR()
	DISP_STOCKPROP_BACKCOLOR()
	DISP_STOCKPROP_ENABLED()
	DISP_STOCKPROP_APPEARANCE()
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CKCOMDateEditCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CKCOMDateEditCtrl, COleControl)
	//{{AFX_EVENT_MAP(CKCOMDateEditCtrl)
	EVENT_CUSTOM("Change", FireChange, VTS_NONE)
	EVENT_STOCK_CLICK()
	EVENT_STOCK_DBLCLICK()
	EVENT_STOCK_KEYDOWN()
	EVENT_STOCK_KEYPRESS()
	EVENT_STOCK_KEYUP()
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CKCOMDateEditCtrl, 3)
	PROPPAGEID(CKCOMDateEditPropPage::guid)
	PROPPAGEID(CLSID_CFontPropPage)
	PROPPAGEID(CLSID_CColorPropPage)
END_PROPPAGEIDS(CKCOMDateEditCtrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CKCOMDateEditCtrl, "KCOMEDITBOX.EditboxCtrl.1",
	0xf6d0e407, 0x2d34, 0x11d2, 0x9a, 0x46, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CKCOMDateEditCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DEditbox =
		{ 0xf6d0e405, 0x2d34, 0x11d2, { 0x9a, 0x46, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const IID BASED_CODE IID_DEditboxEvents =
		{ 0xf6d0e406, 0x2d34, 0x11d2, { 0x9a, 0x46, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwEditboxOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CKCOMDateEditCtrl, IDS_KCOMDateEdit, _dwEditboxOleMisc)


/////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditCtrl::CKCOMDateEditCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CKCOMDateEditCtrl

BOOL CKCOMDateEditCtrl::CKCOMDateEditCtrlFactory::UpdateRegistry(BOOL bRegister)
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
			IDS_KCOMDateEdit,
			IDB_KCOMDateEdit,
			afxRegApartmentThreading,
			_dwEditboxOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditCtrl::CKCOMDateEditCtrl - Constructor

CKCOMDateEditCtrl::CKCOMDateEditCtrl()
{
	InitializeIIDs(&IID_DEditbox, &IID_DEditboxEvents);
	
	ptCaretPosition.x = ptCaretPosition.y = 0;
	nCurrentPosition = 0;
	m_ForeColor = RGB(0, 0, 0);
	m_BackColor = RGB(255, 255, 255);
	m_pFont = NULL;
	m_bReadOnly = FALSE;
	m_ta = Left;
	m_bMingGuo = FALSE;
	m_bDisplayMingGuo = FALSE;
	m_bFormatMingGuo = FALSE;
	m_bDisplayFormatMingGuo = FALSE;
	nYearFinish = nMonthFinish = nDayFinish = 0;
	m_bTextNull = TRUE;
	m_bCanYearNegative = FALSE;

	strFormat = _T("yyyy/mm/dd");
	strDisplayFormat = _T("yyyy/mm/dd");
	strData.Empty();
	strDisplayData.Empty();
	strCurrentData.Empty();
	year = 0;
	month = 0;
	day = 0;
	m_PrevData.wYear = year;
	m_PrevData.wMonth = month;
	m_PrevData.wDay = day;

	dlgDatePicker.m_pCtrl = this;
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditCtrl::~CKCOMDateEditCtrl - Destructor

CKCOMDateEditCtrl::~CKCOMDateEditCtrl()
{
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditCtrl::OnDraw - Drawing function

void CKCOMDateEditCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	CFont* pFntPrev = SelectStockFont(pdc);
	COLORREF clrBack = TranslateColor(m_BackColor);
	CBrush br(clrBack);
	
	if (!(m_hWnd == NULL || GetFocus() != this))
		strCurrentData = strData;
	else
		strCurrentData = strDisplayData;

	if (strCurrentData.GetLength() == 0 && !m_bMingGuo)
		strCurrentData = "1998-08-07";

	pdc->FillRect(&rcBounds, &br);
	pdc->SetTextColor(TranslateColor(m_ForeColor));
	pdc->SetBkColor(clrBack);
	switch (m_ta)
	{
	case Left:
		pdc->SetTextAlign(TA_LEFT);
		pdc->ExtTextOut(rcBounds.left, rcBounds.top, ETO_CLIPPED, rcBounds, strCurrentData,
			strCurrentData.GetLength(), NULL);
		break;

	case Center:
		pdc->SetTextAlign(TA_CENTER);
		pdc->ExtTextOut(rcBounds.left + rcBounds.Width() / 2, rcBounds.top, 
			ETO_CLIPPED, rcBounds, strCurrentData, strCurrentData.GetLength(), NULL);
		break;

	case Right:
		pdc->SetTextAlign(TA_RIGHT);
		pdc->ExtTextOut(rcBounds.right, rcBounds.top, ETO_CLIPPED, rcBounds, 
			strCurrentData, strCurrentData.GetLength(), NULL);
		break;
	}

	pdc->SelectObject(pFntPrev);
	if (!(m_hWnd == NULL || GetFocus() != this))
		DisplayCaret();
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditCtrl::DoPropExchange - Persistence support

void CKCOMDateEditCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));

	ASSERT_POINTER(pPX, CPropExchange);
	ExchangeExtent(pPX);
	ExchangeStockProps(pPX);

	// TODO: Call PX_ functions for each persistent custom property.
	PX_String(pPX, _T("DateFormat"), strFormat, "yyyy/mm/dd");
	PX_String(pPX, _T("DisplayDateFormat"), strDisplayFormat, "yyyy/mm/dd");
	PX_String(pPX, _T("Text"), m_strText, "");
	PX_Bool(pPX, _T("ReadOnly"), m_bReadOnly, FALSE);
	PX_Short(pPX, _T("TextAlign"), m_ta);
	PX_Bool(pPX, _T("CanYearNegative"), m_bCanYearNegative);
	
	if (pPX->IsLoading())
	{
		SetTextIn(m_strText);
		SetModifiedFlag(FALSE);
		m_ForeColor = GetForeColor();
		m_BackColor = GetBackColor();
	}
}

/////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditCtrl::OnResetState - Reset control to default state

void CKCOMDateEditCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditCtrl::PreCreateWindow - Modify parameters for CreateWindowEx

BOOL CKCOMDateEditCtrl::PreCreateWindow(CREATESTRUCT& cs)
{
	// get the 3d look
	cs.dwExStyle |= WS_EX_CLIENTEDGE;
	return COleControl::PreCreateWindow(cs);
}

/////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditCtrl::OnOcmCommand - Handle command messages

/////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditCtrl message handlers


void CKCOMDateEditCtrl::OnSetFocus(CWnd* pOldWnd) 
{
	CalcText();

	COleControl::OnSetFocus(pOldWnd);
	MoveLeft(1);
}

void CKCOMDateEditCtrl::CalcText()
{
	m_strText.Empty();
	// calc data displayed when setfocus
	nYearLength = nMonthLength = nDayLength = 0;
	nYearPosition[0] = nYearPosition[1] = -1;
	nMonthPosition[0] = nMonthPosition[1] = -1;
	nDayPosition[0] = nDayPosition[1] = -1;
	int i;
	nDelimiters = 0;
	strData.Empty();
	strDisplayData.Empty();

	CString str, str2;
	str2.Empty();
	str.Format(_T("%3d"), abs(year - 1911));
	str.TrimLeft();
	str2 += str;
	str.Format(_T("%02d"), month);
	str2 += str;
	str.Format(_T("%02d"), day);
	str2 += str;
	if (year < 1911)
		str2 = "-" + str2;
	m_bMingGuo = strFormat == "ggggggg" ? TRUE : FALSE;
	m_bDisplayMingGuo = strDisplayFormat == "ggggggg" ? TRUE : FALSE;
	if (m_bMingGuo)
	{
		if (!m_bTextNull)
		{
			strData = str2;
			if(str2.GetLength() < 7)
				strData += "_";
		}
		else
			strData.Empty();
	}
	if (m_bDisplayMingGuo)
	{
		if (!m_bTextNull)
			strDisplayData = str2;
		else
			strDisplayData.Empty();
	}

	for (i = 0; i < 8; i++)
	{
		nDelimiterPosition[i] = 0;
		strDelimiter[i / 2].Empty();
	}

	TCHAR chr;
	TCHAR chrPrev = '/';
	for (i = 0; i < strFormat.GetLength() && !m_bMingGuo; i ++)
	{
		chr = strFormat[i];
		switch (chr)
		{
			case 'y':
			{
				m_bFormatMingGuo = FALSE;
				nYearLength ++;
				// the start of current field
				if (chrPrev != chr)
					nYearPosition[0] = i;
				chrPrev = 'y';
				if (i == strFormat.GetLength() - 1 ||
					strFormat[i + 1] != 'y') // end of year data
				{
					// the length should be logical
					if (nYearLength == 1 || nYearLength == 3)
					{
						strFormat = strFormat.Left(nYearPosition[0] 
							+ nYearLength) + 'y' + strFormat.Right(
							strFormat.GetLength() - nYearLength - 
							nYearPosition[0]);
						i ++;
						nYearLength ++;
					}
					// the end of currend field
					nYearPosition[1] = i;
					CString strYear;
					if (m_bTextNull)
						strYear = "____";
					else
						strYear.Format(_T("%04d"), year);
					strYear = strYear.Right(nYearLength);
					strData += strYear;
					m_strText += m_bTextNull ? "" : strYear;
				}
				break;
			}

			case 'g':
			{
				m_bFormatMingGuo = TRUE;
				nYearLength ++;
				// the start of current field
				if (chrPrev != chr)
					nYearPosition[0] = i;
				chrPrev = 'g';
				if (i == strFormat.GetLength() - 1 ||
					strFormat[i + 1] != 'g') // end of year data
				{
					// the length should be logical
					if (nYearLength == 2)
					{
						strFormat = strFormat.Left(nYearPosition[0] 
							+ nYearLength) + 'g' + strFormat.Right(
							strFormat.GetLength() - nYearLength - 
							nYearPosition[0]);
						i ++;
						nYearLength = 3;
					}
					if (nYearLength == 1)
					{
						strFormat = strFormat.Left(nYearPosition[0] 
							+ nYearLength) + "gg" + strFormat.Right(
							strFormat.GetLength() - nYearLength - 
							nYearPosition[0]);
						i += 2;
						nYearLength = 3;
					}
					// the end of currend field
					nYearPosition[1] = i;
					CString strYear;
					if (m_bTextNull)
						strYear = "___";
					else
						strYear.Format(_T("%03d"), year - 1911);
					strYear = strYear.Right(nYearLength);
					strData += strYear;
					m_strText += m_bTextNull ? "" : strYear;
				}
				break;
			}

			case 'm':
			{
				nMonthLength++;
				if (chrPrev != chr)
					nMonthPosition[0] = i;
				chrPrev = 'm';
				if (i == strFormat.GetLength() - 1 ||
					strFormat[i + 1] != 'm') // end of month data
				{
					nMonthPosition[1] = i;
					CString strMonth;
					if(m_bTextNull)
						strMonth = "__";
					else
						strMonth.Format(_T("%02d"), month);
					strMonth = strMonth.Right(nMonthLength);
					strData += strMonth;
					m_strText += m_bTextNull ? "" : strMonth;
				}
				break;
			}

			case 'd':
			{
				nDayLength++;
				if (chrPrev != chr)
					nDayPosition[0] = i;
				chrPrev = 'd';
				if (i == strFormat.GetLength() - 1 ||
					strFormat[i + 1] != 'd') // end of day data
				{
					nDayPosition[1] = i;
					CString strDay;
					if (m_bTextNull)
						strDay = "__";
					else
						strDay.Format(_T("%02d"), day);
					strDay = strDay.Right(nDayLength);
					strData += strDay;
					m_strText += m_bTextNull ? "" : strDay;
				}
				break;
			}

			// the delimiters
			default:
			{
				strData += chr;
				strDelimiter[nDelimiters] += chr;
				if (chrPrev != '/')
					nDelimiterPosition[2 * nDelimiters] = i;
				chrPrev = '/';

				if (i == strFormat.GetLength() - 1 ||
					strFormat[i + 1] == 'y' ||
						strFormat[i + 1] == 'm' || 
						strFormat[i + 1] == 'd')// end of delimiters
				{
					nDelimiterPosition[2 * nDelimiters + 1] = i;
					nDelimiters ++;
				}
				break;
			}
		}
	}

	// calc data displayed when killfocus
	nYearDisplayLength = nMonthDisplayLength = nDayDisplayLength = 0;

	for (i = 0; i < strDisplayFormat.GetLength() && !m_bDisplayMingGuo; i ++)
	{
		chr = strDisplayFormat[i];
		switch (chr)
		{
			case 'y':
			{
				m_bDisplayFormatMingGuo = FALSE;
				nYearDisplayLength++;
				if (i == strDisplayFormat.GetLength() - 1 ||
					strDisplayFormat[i + 1] != 'y') 
					// end of year data
				{
					if (nYearDisplayLength == 1 || 
						nYearDisplayLength == 3)
					{
						strDisplayFormat = strDisplayFormat.Left(
							+ i + 1) + 'y' + strDisplayFormat.Right(
							strDisplayFormat.GetLength() - 1 - i);
						i ++;
						nYearDisplayLength ++;
					}
					CString strYear;
					if (m_bTextNull)
						strYear = "____";
					else
						strYear.Format(_T("%04d"), year);

					strYear = strYear.Right(nYearDisplayLength);
					strDisplayData += strYear;
				}
				break;
			}

			case 'g':
			{
				m_bDisplayFormatMingGuo = TRUE;
				nYearDisplayLength++;
				if (i == strDisplayFormat.GetLength() - 1 ||
					strDisplayFormat[i + 1] != 'g') 
					// end of year data
				{
					if (nYearDisplayLength == 2)
					{
						strDisplayFormat = strDisplayFormat.Left(
							+ i + 1) + 'g' + strDisplayFormat.Right(
							strDisplayFormat.GetLength() - 1 - i);
						i ++;
						nYearDisplayLength = 3;
					}
					if (nYearDisplayLength == 1)
					{
						strDisplayFormat = strDisplayFormat.Left(
							+ i + 1) + "gg" + strDisplayFormat.Right(
							strDisplayFormat.GetLength() - 1 - i);
						i += 2;
						nYearDisplayLength = 3;
					}

					CString strYear;
					if (m_bTextNull)
						strYear = "___";
					else
					{
						int nMGYear = year - 1911;
						if (nMGYear >= 100 || nMGYear <= -10)
							strYear.Format(_T("%3d"), nMGYear);
						else
							strYear.Format(_T("%02d"), nMGYear);
					}
					strYear = strYear.Right(nYearDisplayLength);
					strDisplayData += strYear;
				}
				break;
			}

			case 'm':
			{
				nMonthDisplayLength++;
				if (i == strDisplayFormat.GetLength() - 1 ||
					strDisplayFormat[i + 1] != 'm') // end of month data
				{
					CString strMonth;
					if (m_bTextNull)
						strMonth = "__";
					else
						strMonth.Format(_T("%02d"), month);
					strMonth = strMonth.Right(nMonthDisplayLength);
					strDisplayData += strMonth;
				}
				break;
			}

			case 'd':
			{
				nDayDisplayLength++;
				if (i == strDisplayFormat.GetLength() - 1 ||
					strDisplayFormat[i + 1] != 'd') // end of day data
				{
					CString strDay;
					if (m_bTextNull)
						strDay = "__";
					else
						strDay.Format(_T("%02d"), day);
					strDay = strDay.Right(nDayDisplayLength);
					strDisplayData += strDay;
				}
				break;
			}

			default:
			{
				strDisplayData += chr;
				break;
			}
		}
	}

	if (!IsFinished())
	{
		year = m_PrevData.wYear;
		month = m_PrevData.wMonth;
		month = month == 0 ? 1 : month;
		day = m_PrevData.wDay;
		day = day == 0 ? 1 : day;
		if (!m_bTextNull)
			nYearFinish = nMonthFinish = nDayFinish = 4;
		else
		{
			nYearFinish = nMonthFinish = nDayFinish = 0;
			nCurrentPosition = 0;
		}
	}
/*
	if (!m_bMingGuo)
	{
		if (!m_bTextNull)
			if (!m_bFormatMingGuo)
				m_strText.Format("%04d%02d%02d", year, month, day);
			else
				m_strText.Format("%03d%02d%02d", year - 1911, month, day);
	}
	else*/ if (m_bMingGuo && !m_bTextNull)
		m_strText.Format(_T("%03d%02d%02d"), year - 1911, month, day);

}

void CKCOMDateEditCtrl::OnKillFocus(CWnd* pNewWnd) 
{
	CalcData();
	CalcText();
	HideCaret();
	// should destroy the caret when lost focus
	DestroyCaret();
	InvalidateControl();
	FireChange();
	
	COleControl::OnKillFocus(pNewWnd);
}

void CKCOMDateEditCtrl::MoveRight(int nSteps)
{
	if (m_bMingGuo)
	{
		nCurrentPosition += nSteps;
		if (strData.IsEmpty() || (strData.Right(1) != '_' && strData.GetLength() < 7))
			strData += CString("_");

		if (nCurrentPosition >= strData.GetLength())
			nCurrentPosition = strData.GetLength() - 1;
		if (nCurrentPosition < 0)
			nCurrentPosition = 0;
		InvalidateControl();
		return;
	}

	if (nSteps > 0 && (nCurrentPosition == nYearPosition[0] + nYearFinish)
		&& nCurrentPosition <= nYearPosition[1])
		return;

	if (nSteps > 0 && (nCurrentPosition == nMonthPosition[0] + nMonthFinish)
		&& nCurrentPosition <= nMonthPosition[1])
		return;
	
	if (nSteps > 0 && (nCurrentPosition == nDayPosition[0] + nDayFinish)
		&& nCurrentPosition <= nDayPosition[1])
		return;

	int nNewPosition = nCurrentPosition + nSteps;
	if (nNewPosition >= strFormat.GetLength())
		nNewPosition = strFormat.GetLength() - 1;

	while (nNewPosition <= strFormat.GetLength())
	{
		if (!( (nNewPosition >= nYearPosition[0] && nNewPosition <= nYearPosition[1])
			|| (nNewPosition >= nMonthPosition[0] && nNewPosition <= nMonthPosition[1])
			|| (nNewPosition >= nDayPosition[0] && nNewPosition <= nDayPosition[1])))
			nNewPosition ++;
		else
			break;
	}
	if (nNewPosition > strFormat.GetLength()) // can not move ahead
	{
		while (nNewPosition >= 0)
		{
			if (!( (nNewPosition >= nYearPosition[0] && nNewPosition <= nYearPosition[1])
				|| (nNewPosition >= nMonthPosition[0] && nNewPosition <= nMonthPosition[1])
				|| (nNewPosition >= nDayPosition[0] && nNewPosition <= nDayPosition[1])))
				nNewPosition --;
			else
				break;
		}
	}
	
	nCurrentPosition = nNewPosition;
	InvalidateControl();
}

void CKCOMDateEditCtrl::MoveLeft(int nSteps)
{
	if (m_bMingGuo)
	{
		nCurrentPosition -= nSteps;
		if (nCurrentPosition >= strData.GetLength())
			nCurrentPosition = strData.GetLength() -1;
		if (nCurrentPosition < 0)
			nCurrentPosition = 0;
		InvalidateControl();
		return;
	}

	if (nSteps < 0 && (nCurrentPosition == nYearPosition[0] + nYearFinish)
		&& nCurrentPosition <= nYearPosition[1])
		return;

	if (nSteps < 0 && (nCurrentPosition == nMonthPosition[0] + nMonthFinish)
		&& nCurrentPosition <= nMonthPosition[1])
		return;
	
	if (nSteps < 0 && (nCurrentPosition == nDayPosition[0] + nDayFinish)
		&& nCurrentPosition <= nDayPosition[1])
		return;

	int nNewPosition = nCurrentPosition - nSteps;
	if (nNewPosition < 0)
		nNewPosition = 0;

	while (nNewPosition >= 0)
	{
		if (!( (nNewPosition >= nYearPosition[0] && nNewPosition <= nYearPosition[1])
			|| (nNewPosition >= nMonthPosition[0] && nNewPosition <= nMonthPosition[1])
			|| (nNewPosition >= nDayPosition[0] && nNewPosition <= nDayPosition[1])))
			nNewPosition --;
		else
			break;
	}
	if (nNewPosition < 0) // can not move left
	{
		while (nNewPosition <= strFormat.GetLength())
		{
			if (!( (nNewPosition >= nYearPosition[0] && nNewPosition <= nYearPosition[1])
				|| (nNewPosition >= nMonthPosition[0] && nNewPosition <= nMonthPosition[1])
				|| (nNewPosition >= nDayPosition[0] && nNewPosition <= nDayPosition[1])))
				nNewPosition ++;
			else
				break;
		}
	}

	nCurrentPosition = nNewPosition;
	InvalidateControl();
}

BSTR CKCOMDateEditCtrl::GetDateFormat() 
{
	CString strResult;
	strResult = strFormat;

	return strResult.AllocSysString();
}

void CKCOMDateEditCtrl::SetDateFormat(LPCTSTR lpszNewValue) 
{
	CString strNewValue(lpszNewValue);
	strNewValue.MakeLower();
	if (IsValidFormat(strNewValue))
	{
		if (strNewValue == "ggggggg")
			m_bMingGuo = TRUE;
		else
			m_bMingGuo = FALSE;

		strFormat = strNewValue;
		CalcText();
		InvalidateControl();
		SetModifiedFlag();
		BoundPropertyChanged(dispidDateFormat);
	}
	else
		AfxMessageBox(IDS_ERROR_INVALIDFORMAT);
}

BOOL CKCOMDateEditCtrl::IsValidFormat(CString strNewFormat)
{
	if (strNewFormat == "ggggggg")
		return TRUE;

	TCHAR chr;
	TCHAR chrPrev = '/';
	int i = 0;
	int nYearPosition = 0, nMonthPosition = 0, nDayPosition = 0;
	BOOL bHasYear = FALSE, bHasMonth = FALSE, 
		bHasDay = FALSE;
	
	for (i = 0; i < strNewFormat.GetLength(); i ++)
	{
		chr = strNewFormat[i];
		switch (chr)
		{
			case 'y':
			{
				if (chrPrev == chr)
				{
					// year length can not exceed 4
					if (++ nYearPosition > 4)
						return FALSE;
				}
				else
				{
					if (bHasYear != TRUE)
					{
						chrPrev = 'y';
						// year format already began
						bHasYear = TRUE;
						if (++ nYearPosition > 4)
							return FALSE;
					}
					else
						return FALSE;
				}
			}
			break;

			case 'g':
			{
				if (chrPrev == chr)
				{
					// year length can not exceed 4
					if (++ nYearPosition > 3)
						return FALSE;
				}
				else
				{
					if (bHasYear != TRUE)
					{
						chrPrev = 'g';
						// year format already began
						bHasYear = TRUE;
						if (++ nYearPosition > 3)
							return FALSE;
					}
					else
						return FALSE;
				}
			}
			break;

			case 'm':
			{
				if (chrPrev == chr)
				{
					// month length can not exceed 2
					if (++ nMonthPosition > 2)
						return FALSE;
				}
				else
				{
					if (bHasMonth != TRUE)
					{
						// month format already began
						bHasMonth = TRUE;
						if (++ nMonthPosition > 2)
							return FALSE;
					}
					else
						return FALSE;
				}
				chrPrev = 'm';
			}
			break;

			case 'd':
			{
				if (chrPrev == chr)
				{
					if (++ nDayPosition > 2)
						return FALSE;
				}
				else
				{
					if (bHasDay != TRUE)
					{
						bHasDay = TRUE;
						if (++ nDayPosition > 2)
							return FALSE;
					}
					else
						return FALSE;
				}
				chrPrev = 'd';
			}
			break;

			default:
			{
				// previous character is delimiter
				chrPrev = '/';
				break;
			}
		}
	}
	// can not be all empty
	if (nYearPosition == 0 && nMonthPosition == 0 && nDayPosition == 0)
		return FALSE;
	else
		return TRUE;
}

BSTR CKCOMDateEditCtrl::GetDisplayDateFormat() 
{
	CString strResult;
	strResult = strDisplayFormat;

	return strResult.AllocSysString();
}

void CKCOMDateEditCtrl::SetDisplayDateFormat(LPCTSTR lpszNewValue) 
{
	CString strNewValue(lpszNewValue);
	strNewValue.MakeLower();
	if (IsValidFormat(strNewValue))
	{
		if (strNewValue == "ggggggg")
			m_bDisplayMingGuo = TRUE;
		else
			m_bDisplayMingGuo = FALSE;
		strDisplayFormat = strNewValue;
		CalcText();
		InvalidateControl();
		SetModifiedFlag();
		BoundPropertyChanged(dispidDisplayDateFormat);
	}
	else
		AfxMessageBox(IDS_ERROR_INVALIDDISPLAYFORMAT);
}

BOOL CKCOMDateEditCtrl::PreTranslateMessage(MSG* pMsg) 
{
	if (pMsg->hwnd == this->m_hWnd && pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_DOWN || pMsg->wParam == VK_UP || 
		pMsg->wParam == VK_LEFT || pMsg->wParam == VK_RIGHT))
	{
		OnKeyDown((int)pMsg->wParam, 1, 1);
		return TRUE;
	}

	return COleControl::PreTranslateMessage(pMsg);
}

void CKCOMDateEditCtrl::OnFontChanged() 
{
	COleControl::OnFontChanged();
}

void CKCOMDateEditCtrl::OnForeColorChanged() 
{
	COleControl::OnForeColorChanged();

	m_ForeColor = GetForeColor();

	BoundPropertyChanged(DISPID_FORECOLOR);
}

void CKCOMDateEditCtrl::DisplayCaret()
{
	HideCaret();

	CDC * pDC = GetDC();

	CFont* pFntPrev = SelectStockFont(pDC);
	// the caret should cover the next character
	ptCaretPosition.x = (pDC->GetOutputTextExtent(strData,
		nCurrentPosition)).cx;
	// the length of the total text
	CSize sz = pDC->GetOutputTextExtent(strData);
	CRect rect;
	GetClientRect(&rect);
	// move the caret to the correct position according to the text
	// alignment
	if (m_ta == Center)
		ptCaretPosition.x += (rect.Width() - sz.cx) / 2;
	if (m_ta == Right)
		ptCaretPosition.x += rect.Width() - sz.cx;

	CSize szCaret;
	// should not exceed the array's range, or the program will crash
	szCaret = pDC->GetOutputTextExtent(nCurrentPosition >= strData.GetLength() ? 'A' : CString(strData[nCurrentPosition]), 1);
	// to show the caret, need creating it first
	::CreateCaret(this->m_hWnd, NULL, szCaret.cx, szCaret.cy);
	ShowCaret();
	SetCaretPos(ptCaretPosition);

	/*	CRect rect(ptCaretPosition, szCaret);
	CBrush br(RGB(125, 125, 125));
	pDC->FillRect(&rect, &br);*/
	
	pDC->SelectObject(pFntPrev);
	ReleaseDC(pDC);
}

// we use the "value" property to exchange data with outside world
BSTR CKCOMDateEditCtrl::GetText() 
{
	CalcData();
	CalcText();
	return m_strText.AllocSysString();
}

void CKCOMDateEditCtrl::SetText(LPCTSTR lpszNewValue) 
{
	if (SetTextIn(lpszNewValue))
	{
		InvalidateControl();
		BoundPropertyChanged(dispidText);
		SetModifiedFlag();
		FireChange();
	}
	else
	{
		SetTextIn("");
		InvalidateControl();
		SetModifiedFlag();
		BoundPropertyChanged(dispidText);
		FireChange();
	}
}

void CKCOMDateEditCtrl::OnReadOnlyChanged() 
{
	InvalidateControl();
	SetModifiedFlag();
	BoundPropertyChanged(dispidReadOnly);
}

short CKCOMDateEditCtrl::GetTextAlign() 
{
	return m_ta;
}

void CKCOMDateEditCtrl::SetTextAlign(short nNewValue) 
{
	m_ta = nNewValue % 3;
	InvalidateControl();

	SetModifiedFlag();
	BoundPropertyChanged(dispidTextAlign);
}

BOOL CKCOMDateEditCtrl::IsFinished()
{
	return (nYearFinish >= nYearLength && nMonthFinish >= nMonthLength
		&& nDayFinish >= nDayLength);
}

void CKCOMDateEditCtrl::OnLButtonDown(UINT nFlags, CPoint point) 
{
	COleControl::OnLButtonDown(nFlags, point);

	int pos = nCurrentPosition;

	CDC * pDC = GetDC();
	pDC->DPtoLP(&point);
	CSize sz = pDC->GetOutputTextExtent(strData);
	CSize sz2 = pDC->GetOutputTextExtent("2");
	CRect rect(CPoint(0, 0), sz);
	if (rect.PtInRect(point))
	{
		int i = (int) (point.x / sz2.cx);
		int j;
		if (i > pos)
			for (j = 0; j <= i - pos; j ++)
				MoveRight(1);
		if (i < nCurrentPosition)
			for (j = 0; j < pos - i; j ++)
				MoveLeft(1);
	}

	ReleaseDC(pDC);
}

void CKCOMDateEditCtrl::SetDate(short nYear, short nMonth, short nDay) 
{
	CalcText();

	COleDateTime date;
	int nNewYear = nYear;
	
	if (m_bMingGuo || m_bFormatMingGuo)
		nNewYear = nYear + 1911;
	date.SetDate(nNewYear, nMonth, nDay);
	if (date.GetStatus() == COleDateTime::valid)
	{
		m_bTextNull = FALSE;
		year = nNewYear;
		month = nMonth;
		day = nDay;
		m_PrevData.wYear = year;
		m_PrevData.wMonth = month;
		m_PrevData.wDay = day;
	}
	
	CalcText();
	InvalidateControl();
}

void CKCOMDateEditCtrl::CalcMingGuo()
{
	int i;
	TCHAR chr;
	CString strYear, strMonth, strDay, strTotal;
	int nNewYear, nNewMonth, nNewDay;
	nNewYear = nNewMonth = nNewDay = 0;
	BOOL bValid = TRUE;
	strYear.Empty();
	strMonth.Empty();
	strDay.Empty();
	strTotal.Empty();

	if (strData.IsEmpty())
		return;
	else
	{
		for (i = 0; i < strData.GetLength(); i++)
		{
			chr = strData[i];
			if ((chr >= '0' && chr <= '9') || (chr == '-' &&
				strTotal.IsEmpty()))
				strTotal += chr;
		}

		if (strTotal.IsEmpty() || (strTotal[0] == '-' && strTotal.
			GetLength() < 6) || (strTotal[0] != '-' && strTotal.
			GetLength() < 5))
			return;
		
		strDay = strTotal.Right(2);
		strMonth = strTotal.Mid(strTotal.GetLength() - 4, 2);
		strYear = strTotal.Left(strTotal.GetLength() - 4);
		
		i = strYear[0] == '-' ? 1 : 0;
		for (; i < strYear.GetLength(); i++)
			nNewYear = nNewYear * 10 + strYear[i] - '0';
		if (strYear[0] == '-')
			nNewYear = -nNewYear;
		for (i = 0; i < 2; i++)
			nNewMonth = nNewMonth * 10 + strMonth[i] - '0';
		for (i = 0; i < 2; i++)
			nNewDay = nNewDay * 10 + strDay[i] - '0';

		COleDateTime date;
		date.SetDate(nNewYear + 1911, nNewMonth, nNewDay);
		if (date.GetStatus() == COleDateTime::valid)
			bValid = TRUE;
		else
			return;
	}

	if (bValid)
		SetDate(nNewYear, nNewMonth, nNewDay);
}

void CKCOMDateEditCtrl::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	COleControl::OnLButtonDblClk(nFlags, point);

	// let the user select a date via a calendar

	if (!IsWindow(dlgDatePicker.m_hWnd))
		dlgDatePicker.Create(CDlgDatePicker::IDD);
	dlgDatePicker.ShowWindow(SW_SHOW);

/*	if (dlgDatePicker.DoModal() == IDOK && !m_bReadOnly)
	{
		year = dlgDatePicker.m_Date.GetYear();
		month = dlgDatePicker.m_Date.GetMonth();
		day = dlgDatePicker.m_Date.GetDay();
		m_bTextNull = FALSE;
		nYearFinish = nMonthFinish = nDayFinish = 4;
		nCurrentPosition = 0;
		CalcText();
		MoveRight();
		FireChange();
		InvalidateControl();
	}*/
}

BOOL CKCOMDateEditCtrl::SetTextIn(CString strNewValue)
{
	CalcText();

	CString strData = strNewValue;
	int i;
	TCHAR chr;
	CString strYear, strMonth, strDay, strTotal;
	int nNewYear, nNewMonth, nNewDay;
	nNewYear = nNewMonth = nNewDay = 0;
	BOOL bValid = TRUE;
	strYear.Empty();
	strMonth.Empty();
	strDay.Empty();
	strTotal.Empty();

	if (strData.IsEmpty() || strData == '_')
	{
		m_bTextNull = TRUE;
		nYearFinish = nMonthFinish = nDayFinish = 0;
		nCurrentPosition = 0;
	}
	else
	{
		for (i = 0; i < strData.GetLength(); i++)
		{
			chr = strData[i];
			if ((chr >= '0' && chr <= '9') || (chr == '-' &&
				strTotal.IsEmpty()))
				strTotal += chr;
		}

		if (strTotal.IsEmpty() || (strTotal[0] == '-' && strTotal.
			GetLength() < 6) || (strTotal[0] != '-' && strTotal.
			GetLength() < 5))
			return FALSE;
		
		strDay = strTotal.Right(2);
		strMonth = strTotal.Mid(strTotal.GetLength() - 4, 2);
		strYear = strTotal.Left(strTotal.GetLength() - 4);
	
		i = strYear[0] == '-' ? 1 : 0;
		if (i == 1 && !m_bMingGuo)
			return FALSE;

		for (; i < strYear.GetLength(); i++)
			nNewYear = nNewYear * 10 + strYear[i] - '0';
		if (strYear[0] == '-')
			nNewYear = -nNewYear;
		for (i = 0; i < 2; i++)
			nNewMonth = nNewMonth * 10 + strMonth[i] - '0';
		for (i = 0; i < 2; i++)
			nNewDay = nNewDay * 10 + strDay[i] - '0';

		COleDateTime date;
		date.SetDate(nNewYear + ((m_bMingGuo || m_bFormatMingGuo)? 
			1911 : 0), nNewMonth, nNewDay);
		if (date.GetStatus() == COleDateTime::valid)
			bValid = TRUE;
		else
			return FALSE;
	}

	if (bValid)
	{
		if (!(strData.IsEmpty() || strData == '_'))
			SetDate(nNewYear, nNewMonth, nNewDay);
		else
		{
			month = day = 1;
			CalcText();
		}

		m_strText = strData;
	}

	return TRUE;
}

void CKCOMDateEditCtrl::OnBackColorChanged() 
{
	COleControl::OnBackColorChanged();

	m_BackColor = GetBackColor();

	BoundPropertyChanged(DISPID_BACKCOLOR);
}

void CKCOMDateEditCtrl::OnCanYearNegativeChanged() 
{
	SetModifiedFlag();
	BoundPropertyChanged(dispidCanYearNegative);
}

void CKCOMDateEditCtrl::OnKeyDownEvent(USHORT nChar, USHORT nShiftState) 
{
	switch (nChar)
	{
		case VK_HOME:
		{
			MoveLeft(nCurrentPosition);
		}
		break;

		case VK_UP:
		case VK_LEFT:
		{
			MoveLeft(1);
		}
		break;

		case VK_END:
		{
			int pos = nCurrentPosition;
			for (int i = 0; i < strData.GetLength() - pos - 1; i++)
				MoveRight(1);
		}
		break;

		case VK_DOWN:
		case VK_RIGHT:
		{
			MoveRight(1);
		}
		break;

		case VK_DELETE:
		{
			OnKeyPressEvent(VK_DELETE);
		}
		break;
	}

	COleControl::OnKeyDownEvent(nChar, nShiftState);
}

void CKCOMDateEditCtrl::OnKeyPressEvent(USHORT nChar) 
{
	BOOL bFinish = TRUE;

	if (m_bReadOnly)
		return;

	if (m_bMingGuo)
	{
		if ((nChar < '0' || nChar > '9') && nChar != 
			VK_BACK && nChar != VK_DELETE && nChar != '/' && nChar != '-')
			return;
		if (nChar == VK_BACK)
		{
			if (nCurrentPosition == 0 && strData.IsEmpty())
				return;

			strData = strData.Left(nCurrentPosition - 1) + strData.
				Right(strData.GetLength() - nCurrentPosition);
			if (strData.IsEmpty() || strData == '_')
				SetText(_T(""));
			if (strData.IsEmpty() || strData.Right(1) != '_')
				strData += CString("_");
			MoveLeft(1);
			InvalidateControl();
			return;
		}

		if (nChar == VK_DELETE)
		{
			strData = strData.Left(nCurrentPosition) + strData.
				Right(strData.GetLength() - nCurrentPosition - 1);
			if (strData.IsEmpty() || strData == '_')
				SetText(_T(""));
			if (strData.IsEmpty() || strData.Right(1) != '_')
				strData += CString("_");
			MoveLeft();
			InvalidateControl();
			return;
		}

		else
			if (strData.IsEmpty())
				strData += (char)nChar;
			else
				strData.SetAt(nCurrentPosition, (char)nChar);
		// move the caret right
		if (strData.IsEmpty() || (strData.Right(1) != '_' && strData.GetLength() < 7))
			strData += CString("_");
		// now the value is accpeted
		if (strData.GetLength() == 7 && strData.Right(1) != '_')
			FireChange();

		MoveRight(1);
		InvalidateControl();
		return;
	}

	// unsupported characters
	if ((nChar < '0' || nChar > '9') && nChar != VK_BACK && nChar !=
		VK_DELETE && nChar != '/' && nChar != '-')
		return;

	COleDateTime temptime;
	CString text;
	text = strData;

	int i = 0;
	int nNewYear = 0, nNewMonth = 0, nNewDay = 0;
	// the validity of new data
	BOOL bVerify = TRUE;

	if ((nChar == VK_BACK || nChar == VK_DELETE) && IsFinished()
		&& !m_bTextNull)
	{
		m_PrevData.wYear = year;
		m_PrevData.wMonth = month;
		m_PrevData.wDay = day;
	}
	if (nChar == VK_BACK || nChar == VK_DELETE)
		bFinish = FALSE;
	int nNewChar = nChar - '0';
	// change the year text
	if (nCurrentPosition >= nYearPosition[0] 
		&& nCurrentPosition <= nYearPosition[1]) 
	{
		// whether the first char is '-'
		BOOL bNegative = FALSE;

		// only the first position can be '-'
		if (nChar == '-' && !(m_bCanYearNegative && m_bFormatMingGuo
			&& nCurrentPosition == nYearPosition[0]))
			return;

		if (!bFinish)
		{
			for (i = nCurrentPosition; i <= nYearPosition[1]; i++)
				strData.SetAt(i, '_');

			nYearFinish = nCurrentPosition - nYearPosition[0];
			if (nYearFinish == 0 && nMonthFinish == 0 && nDayFinish == 0)
			{
				SetText(_T(""));
				return;
			}
			i = 0;
			if (text[nYearPosition[0]] == '-')
			{
				bNegative = TRUE;
				i ++;
			}
			for (; i < nYearFinish; i++)
				nNewYear = nNewYear * 10 + text[nYearPosition[0] + i] - '0';
			if (bNegative)
				nNewYear = -nNewYear;

			year = nNewYear + (m_bFormatMingGuo ? 1911 : 0);
			if (nChar == VK_BACK)
				MoveLeft(1);
			InvalidateControl();
			return;
		}

		// update current position number
		if (nChar != '/')
		{
			i = 0;
			if (text[nYearPosition[0]] == '-' || (nCurrentPosition 
				== nYearPosition[0] && nChar == '-'))
			{
				bNegative = TRUE;
				i ++;
			}
			if (nCurrentPosition == nYearPosition[0] && 
				nChar != '-' && bNegative)
			{
				bNegative = FALSE;
				i --;
			}
			for (; i < nCurrentPosition - nYearPosition[0]; i++)
				nNewYear = nNewYear * 10 + text[nYearPosition[0] + i] - '0';
			if (!bNegative || nCurrentPosition != nYearPosition[0])
				nNewYear = nNewYear * 10 + nNewChar;
			if (nCurrentPosition - nYearPosition[0] == nYearFinish)
				nYearFinish ++;
			if (nYearFinish >= nYearLength)
			{
				for (i = 1; i <= nYearPosition[1] - nCurrentPosition; i++)
					nNewYear = nNewYear * 10 + text[nCurrentPosition + i] - '0';
			}
			if (bNegative)
				nNewYear = -nNewYear;

			nNewYear += m_bFormatMingGuo ? 1911 : 0;
			temptime.SetDate(nNewYear, month, day);
			if (temptime.GetStatus() == COleDateTime::invalid && (
				IsFinished() || (nCurrentPosition == nYearPosition[1]
				&& !m_bTextNull)))
			{
				bVerify = FALSE;
				if (text[nCurrentPosition] == '_')
					nYearFinish --;
			}
			else if (nYearLength != 0)
			{
				int i;
				for (i = 0; i < nYearLength; i++)
					year /= 10;
				for (i = 0; i < nYearLength; i++)
					year *= 10;
				year = nNewYear + (m_bFormatMingGuo ? 0 : year);
			}
		}
		else
		{
			// if the user press '/', I should supplement the space
			// he left, and move to next field.
			i = 0;
			if (text[nYearPosition[0]] == '-')
			{
				bNegative = TRUE;
				i ++;
			}
			for (; i < nYearLength && text[i + nYearPosition[0]] != '_' ; i++)
				nNewYear = nNewYear * 10 + text[nYearPosition[0] + i] - '0';
			if (bNegative)
				nNewYear = -nNewYear;
			nNewYear += m_bFormatMingGuo ? 1911 : 0;

			temptime.SetDate(nNewYear, month, day);
			// now, i is the number of space the the user input.
			if (temptime.GetStatus() == COleDateTime::valid)
			{
				int j;
				for (j = 1; j <= i; j ++)
					strData.SetAt(nYearLength - j + nYearPosition[0], strData[i - j + nYearPosition[0]]);
				nYearFinish = 4;

				j = 0;
				if (bNegative)
				{	
					strData.SetAt(nYearPosition[0], '-');
					j = 1;
				}
				for (; j < nYearLength - i; j ++)
					strData.SetAt(j + nYearPosition[0], '0');
				i = nCurrentPosition;
				for (j = 0; j <= nYearPosition[1] - i; j ++)
					MoveRight(1);
			}
		}
	}

	// change the month text
	else if (nCurrentPosition >= nMonthPosition[0] 
		&& nCurrentPosition <= nMonthPosition[1]) 
	{
		if (!bFinish)
		{
			for (i = nMonthPosition[0]; i <= nMonthPosition[1]; i++)
				strData.SetAt(i, '_');
			nMonthFinish = 0;
			month = 1;
			MoveLeft(nCurrentPosition - nMonthPosition[0]);
			if (nChar == VK_BACK)
				MoveLeft(1);
			InvalidateControl();
			return;
		}

		if (nChar != '/')
		{
			for (i = 0; i < nCurrentPosition - nMonthPosition[0]; i++)
				nNewMonth = nNewMonth * 10 + text[nMonthPosition[0] + i] - '0';
			nNewMonth = nNewMonth * 10 + nNewChar;
			for (i = 1; i <= nMonthPosition[1] - nCurrentPosition; i++)
			{
				if (text[nCurrentPosition + i] == '_')
					nNewMonth = nNewMonth * 10;
				else
					nNewMonth = nNewMonth * 10 + text[nCurrentPosition + i] - '0';
			}

			temptime.SetDate(year, nNewMonth, day);
			if (temptime.GetStatus() == COleDateTime::invalid)
			{
				bVerify = FALSE;
				temptime.SetDate(year, nChar - '0', day);

				if (temptime.GetStatus() == COleDateTime::valid)
				{
					for (i = nMonthPosition[0]; i <= nMonthPosition[1]; i++)
						strData.SetAt(i, '0');

					// if the new month is too big, only keep the lowerst
					// position number.
					nCurrentPosition = i - 1;
					month = nChar - '0';
					nMonthFinish = nMonthLength;
					bVerify = TRUE;
				}
				else if (nChar == '0' && nCurrentPosition == nMonthPosition[0])
				{
					bVerify = TRUE;
					nMonthFinish = max(nMonthFinish, nCurrentPosition - nMonthPosition[0] + 1);
				}
			}

			else if (nMonthLength != 0)
			{
				int i;
				for (i = 0; i < nMonthLength; i++)
					month /= 10;
				for (i = 0; i < nMonthLength; i++)
					month *= 10;
				month += nNewMonth;
				nMonthFinish ++;
			}
		}
		else
		{
			// if the user press '/', I should supplement the space
			// he left, and move to next field.
			for (i = 0; i < nMonthLength && text[i + nMonthPosition[0]] != '_' ; i++)
				nNewMonth = nNewMonth * 10 + text[nMonthPosition[0] + i] - '0';
			temptime.SetDate(year, nNewMonth, day);
			// now, i is the number of space the the user input.
			if (temptime.GetStatus() == COleDateTime::valid)
			{
				int j;
				for (j = 1; j <= i; j ++)
					strData.SetAt(nMonthPosition[0] + nMonthLength - j, strData[i - j + nMonthPosition[0]]);
				nMonthFinish = 4;
				for (j = 0; j < nMonthLength - i; j ++)
					strData.SetAt(j + nMonthPosition[0], '0');
				i = nCurrentPosition;
				for (j = 0; j <= nMonthPosition[1] - i; j ++)
					MoveRight(1);

				month = nNewMonth;
			}
		}
	}

	// change the Day text
	else if (nCurrentPosition >= nDayPosition[0] 
		&& nCurrentPosition <= nDayPosition[1]) 
	{
		if (!bFinish)
		{
			for (i = nDayPosition[0]; i <= nDayPosition[1]; i++)
				strData.SetAt(i, '_');
			nDayFinish = 0;
			day = 1;
			MoveLeft(nCurrentPosition - nDayPosition[0]);
			if (nChar == VK_BACK)
				MoveLeft(1);
			InvalidateControl();
			return;
		}

		if (nChar != '/')
		{
			for (i = 0; i < nCurrentPosition - nDayPosition[0]; i++)
				nNewDay = nNewDay * 10 + text[nDayPosition[0] + i] - '0';
			nNewDay = nNewDay * 10 + nNewChar;
			for (i = 1; i <= nDayPosition[1] - nCurrentPosition; i++)
			{
				if (text[nCurrentPosition + i] == '_')
					nNewDay = nNewDay * 10;
				else
					nNewDay = nNewDay * 10 + text[nCurrentPosition + i] - '0';
			}

			temptime.SetDate(year, month, nNewDay);
			if (temptime.GetStatus() == COleDateTime::invalid)
			{
				bVerify = FALSE;
				// eliminate the lower position number to make it valid
				temptime.SetDate(year, month, (nChar - '0') * 10);
				if (temptime.GetStatus() == COleDateTime::valid && 
					nCurrentPosition == nDayPosition[0])
				{
					for (i = nDayPosition[0]; i <= nDayPosition[1]; i++)
						strData.SetAt(i, '0');

					day = (nChar - '0') * 10;
					nCurrentPosition = nDayPosition[0];
					nDayFinish = nDayLength;
					bVerify = TRUE;
				}
				// if the method above does not work, eliminate the lower
				else
				{
					temptime.SetDate(year, month, nChar - '0');

					if (temptime.GetStatus() == COleDateTime::valid)
					{
						for (i = nDayPosition[0]; i <= nDayPosition[1]; 
							i++)
							strData.SetAt(i, '0');

						nCurrentPosition = i - 1;
						day = nChar - '0';
						nDayFinish = nDayLength;
						bVerify = TRUE;
					}
					else if (nChar == '0' && nCurrentPosition == nDayPosition[0])
					{
						bVerify = TRUE;
						nDayFinish ++;
					}
				}
			}
			else if (nDayLength != 0)
			{
				int i;
				for (i = 0; i < nDayLength; i++)
					day /= 10;
				for (i = 0; i < nDayLength; i++)
					day *= 10;
				day += nNewDay;
				nDayFinish ++;
			}
		}
		else
		{
			// if the user press '/', I should supplement the space
			// he left, and move to next field.
			for (i = 0; i < nDayLength && text[i + nDayPosition[0]] != '_' ; i++)
				nNewDay = nNewDay * 10 + text[nDayPosition[0] + i] - '0';
			temptime.SetDate(year, month, nNewDay);
			// now, i is the number of space the the user input.
			if (temptime.GetStatus() == COleDateTime::valid)
			{
				int j;
				for (j = 1; j <= i; j ++)
					strData.SetAt(nDayLength - j + nDayPosition[0], strData[i - j + nDayPosition[0]]);
				nDayFinish = 4;
				for (j = 0; j < nDayLength - i; j ++)
					strData.SetAt(j + nDayPosition[0], '0');
				i = nCurrentPosition;
				for (j = 0; j <= nDayPosition[1] - i; j ++)
					MoveRight(1);
				
				day = nNewDay;
			}
		}

	}

	if (bVerify)
	{
		if (nChar != '/')
		{
			// update current character
			strData.SetAt(nCurrentPosition, (char)nChar);
			// move the caret right
			MoveRight(1);
			// avoid the delimiters
			MoveRight();
		}

		if (IsFinished())
		{
			m_bTextNull = FALSE;
			FireChange();
		}

		InvalidateControl();
	}
	
	COleControl::OnKeyPressEvent(nChar);
}

void CKCOMDateEditCtrl::CalcData()
{
	if (!IsFinished())
	{
		year = m_PrevData.wYear;
		month = m_PrevData.wMonth;
		day = m_PrevData.wDay;
		if (!m_bTextNull)
			nYearFinish = nMonthFinish = nDayFinish = 4;
		else
		{
			nYearFinish = nMonthFinish = nDayFinish = 0;
			nCurrentPosition = 0;
		}
	}
	if (m_bMingGuo)
		CalcMingGuo();
}

void CKCOMDateEditCtrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_KCOMDATEEDIT);
	dlgAbout.DoModal();
}

void CKCOMDateEditCtrl::ExchangeStockProps(CPropExchange *pPX)
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

BOOL CKCOMDateEditCtrl::OnGetDisplayString(DISPID dispid, CString& strValue) 
{
	if (dispid == DISPID_APPEARANCE)
	{
		strValue.LoadString(GetAppearance() ? IDS_APPEARANCE_3D : IDS_APPEARANCE_FLAT);

		return TRUE;
	}
	
	return COleControl::OnGetDisplayString(dispid, strValue);
}

BOOL CKCOMDateEditCtrl::OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray) 
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

BOOL CKCOMDateEditCtrl::OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut) 
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
