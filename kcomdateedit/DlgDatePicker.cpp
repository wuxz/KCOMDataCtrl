// DlgDatePicker.cpp : implementation file
//

#include "stdafx.h"
#include "KCOMDateEdit.h"
#include "DlgDatePicker.h"
#include "KCOMDateEditCtl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define WIDTH 228
#define HEIGHT 225
#define XMARGIN 2
#define YMARGIN 2
#define COMBOHEIGHT 40
#define CELLWIDTH 31
#define CELLHEIGHT 20
#define BUTHEIGHT 25
#define BUTWIDTH 25
#define COMBOWIDTH 80
#define IDR_BUTTONLEFTYEAR 101
#define IDR_BUTTONLEFTMONTH 102
#define IDR_BUTTONRIGHTYEAR 103
#define IDR_BUTTONRIGHTMONTH 104
#define IDR_COMBOMONTH 105
#define IDR_EDITYEAR 106
#define IDR_SPINYEAR 107
/////////////////////////////////////////////////////////////////////////////
// CDlgDatePicker dialog

CDlgDatePicker::CDlgDatePicker(CWnd* pParent /*=NULL*/)
	: CDialog(CDlgDatePicker::IDD, pParent)
{
	//{{AFX_DATA_INIT(CDlgDatePicker)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	m_pButtonLeftYear = NULL;
	m_pButtonLeftMonth = NULL;
	m_pButtonRightYear = NULL;
	m_pButtonRightMonth = NULL;

	m_strDayName[0] = "Sun";
	m_strDayName[1] = "Mon";
	m_strDayName[2] = "Tue";
	m_strDayName[3] = "Wed";
	m_strDayName[4] = "Thu";
	m_strDayName[5] = "Fri";
	m_strDayName[6] = "Sat";

	m_Date = COleDateTime::GetCurrentTime();
	m_CurrentDate = COleDateTime::GetCurrentTime();
	year = m_CurrentDate.GetYear();
	month = m_CurrentDate.GetMonth();
	day = today = m_CurrentDate.GetDay();
	
	CalcCells();

	m_pFontPrev = NULL;
	brhGray.CreateSolidBrush(CGRAY);
	brhWhite.CreateSolidBrush(CWHITE);
	brhBlue.CreateSolidBrush(CBLUE);

	m_pCtrl = NULL;
}


void CDlgDatePicker::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDlgDatePicker)
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CDlgDatePicker, CDialog)
	//{{AFX_MSG_MAP(CDlgDatePicker)
	ON_WM_PAINT()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONDBLCLK()
	ON_CBN_SELCHANGE(IDC_COMBO_MONTH, OnSelchangeComboMonth)
	ON_NOTIFY(UDN_DELTAPOS, IDC_SPIN_YEAR, OnDeltaposSpinYear)
	ON_EN_CHANGE(IDC_EDIT_YEAR, OnChangeEditYear)
	ON_WM_ACTIVATE()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDR_BUTTONLEFTYEAR, OnLeftYear)
	ON_BN_CLICKED(IDR_BUTTONLEFTMONTH, OnLeftMonth)
	ON_BN_CLICKED(IDR_BUTTONRIGHTYEAR, OnRightYear)
	ON_BN_CLICKED(IDR_BUTTONRIGHTMONTH, OnRightMonth)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDlgDatePicker message handlers

BOOL CDlgDatePicker::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// the default font is too large
	CPaintDC dc(this);
	dc.FillRect(CRect(0, 0, WIDTH, HEIGHT), &brhGray);

	LOGFONT logfont;
	memset(&logfont, 0, sizeof(logfont));
	logfont.lfWeight = FW_NORMAL;
	logfont.lfHeight = -MulDiv(8, GetDeviceCaps(dc, LOGPIXELSY),72);
	lstrcpy(logfont.lfFaceName, _T("MS Sans Serif"));
	m_Font.CreateFontIndirect(&logfont);

	m_ComboMonth.SubclassDlgItem(IDC_COMBO_MONTH, this);
	m_EditYear.SubclassDlgItem(IDC_EDIT_YEAR, this);
	m_SpinYear.SubclassDlgItem(IDC_SPIN_YEAR, this);
	m_SpinYear.SetRange(0, 5000);
	m_ComboMonth.SetCurSel(month - 1);
	m_SpinYear.SetPos(year);
	
	m_pButtonLeftYear = new CButton();
	m_pButtonLeftMonth = new CButton();
	m_pButtonRightYear = new CButton();
	m_pButtonRightMonth = new CButton();

	rect = CRect(0, 0, WIDTH, HEIGHT);
	rect.OffsetRect(200, 200);
	MoveWindow(rect);

	// the right and bottom margin should be large than normal to 
	// avoid the 3d shape of the dialog box
	rectComboMonth = CRect(XMARGIN, YMARGIN, XMARGIN + COMBOWIDTH, 
		YMARGIN + COMBOHEIGHT);
	rectEditYear = CRect(WIDTH - XMARGIN - COMBOWIDTH - 5, YMARGIN, 
		WIDTH - XMARGIN - 5, YMARGIN + COMBOHEIGHT);
	rectEditYear = CRect(WIDTH - XMARGIN - 4, YMARGIN, 
		WIDTH - XMARGIN, YMARGIN + COMBOHEIGHT);

	rectLeftYear = CRect(XMARGIN, HEIGHT - YMARGIN * 3 - BUTHEIGHT, 
		XMARGIN + BUTWIDTH, HEIGHT - YMARGIN * 3);
	rectLeftMonth = rectLeftYear;
	rectLeftMonth.OffsetRect(BUTWIDTH + 2, 0);
	rectRightYear = CRect(WIDTH - 3 * XMARGIN - BUTWIDTH, HEIGHT - YMARGIN
		* 3 - BUTHEIGHT, WIDTH - 3 * XMARGIN, HEIGHT - YMARGIN * 3);
	rectRightMonth = rectRightYear;
	rectRightMonth.OffsetRect(-(BUTWIDTH + 2), 0);

	DrawButton();
	
	m_pCtrl->GetWindowRect(&rect);
	CRect rtSize;
	GetWindowRect(&rtSize);
	int x, y;

	if (GetSystemMetrics(SM_CYSCREEN) - rect.bottom >= rtSize.Height())
		y = rect.bottom;
	else if (rect.top >= rtSize.Height()) 
		y = rect.top - rtSize.Height();
	else
		y = -1;

	if (GetSystemMetrics(SM_CXSCREEN) - rect.left >= rtSize.Width())
		x = rect.left;
	else if (rect.left >= rtSize.Width())
		x = rect.left - rtSize.Width();
	else
		x = -1;

	if (x == -1 || y == -1)
		CenterWindow();
	else
		SetWindowPos(NULL, x, y, 0, 0, SWP_NOZORDER | SWP_NOSIZE);

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CDlgDatePicker::OnPaint() 
{
	CPaintDC dc(this); // device context for painting
	
	m_pFontPrev = dc.SelectObject(&m_Font);

	DrawGrid(dc);
	DrawTitle(dc);

	dc.SelectObject(m_pFontPrev);
}

void CDlgDatePicker::OnLButtonDown(UINT nFlags, CPoint point) 
{
	// to see which cell the user pointed in
	int x = (point.x - XMARGIN) / CELLWIDTH;
	int y = (point.y - (YMARGIN + COMBOHEIGHT)) / CELLHEIGHT - 1;
	if (x >= 0 && x <= 6 && y >= 0 && y <= 5 && Cells[y][x] != 0)
	{
		day = Cells[y][x];
		m_CurrentDate.SetDate(year, month, day);
		CalcCells();
		Invalidate();
	}

	CDialog::OnLButtonDown(nFlags, point);
}

void CDlgDatePicker::PostNcDestroy() 
{
	// the own created font should be deleted before exit
	m_Font.DeleteObject();
	
	// the own created buttons shouled be deleted before exit
	m_pButtonLeftYear->DestroyWindow();
	delete m_pButtonLeftYear;
	m_pButtonLeftYear = NULL;
	m_pButtonLeftMonth->DestroyWindow();
	delete m_pButtonLeftMonth;
	m_pButtonLeftMonth = NULL;
	m_pButtonRightYear->DestroyWindow();
	delete m_pButtonRightYear;
	m_pButtonRightYear = NULL;
	m_pButtonRightMonth->DestroyWindow();
	delete m_pButtonRightMonth;
	m_pButtonRightMonth = NULL;

	CDialog::PostNcDestroy();
}

void CDlgDatePicker::DrawGrid(CPaintDC& dc)
{
	dc.SetTextColor(CBLACK);
	dc.SetTextAlign(TA_CENTER);
	dc.SetBkColor(CGRAY);

	int i, j;
	// draw the vertical lines
	for (i = XMARGIN; i <= XMARGIN + CELLWIDTH * 7; i += CELLWIDTH)
	{
		dc.MoveTo(i, YMARGIN + COMBOHEIGHT);
		dc.LineTo(i, YMARGIN + COMBOHEIGHT + CELLHEIGHT * 7);
	}

	// draw the 7 days name
	for (i = 0; i < 7; i++)
	{
		rect = CRect(XMARGIN + i * CELLWIDTH, YMARGIN + COMBOHEIGHT, XMARGIN + (i + 1) 
			* CELLWIDTH, YMARGIN + COMBOHEIGHT + CELLHEIGHT);
		dc.TextOut(rect.left + CELLWIDTH / 2, rect.top + 2, m_strDayName[i]);
	}

	// draw the horizental lines
	for (i = YMARGIN + COMBOHEIGHT; i <= YMARGIN + COMBOHEIGHT + CELLHEIGHT * 7; i += CELLHEIGHT)
	{
		dc.MoveTo(XMARGIN, i);
		dc.LineTo(XMARGIN + CELLWIDTH * 7, i);
	}

	// draw the cells, fill it and draw the day's number
	for (i = 0; i < 6; i++)
		for (j = 0; j < 7; j ++)
		{
			rect = CRect(XMARGIN + j * CELLWIDTH, YMARGIN + COMBOHEIGHT + (i + 1)* CELLHEIGHT, 
				XMARGIN + (j + 1) * CELLWIDTH, YMARGIN + COMBOHEIGHT + CELLHEIGHT * (i + 2));
			// avoid drawing on the drawn lines
			rect.DeflateRect(1, 1);
			// current selected day uses blue, others use white
			dc.FillRect(&rect, Cells[i][j] == day ? &brhBlue : &brhWhite);
			
			// sunday uses red color to draw text
			if (j == 0)
				dc.SetTextColor(CRED);
			// today used magenta color to draw text
			else if (Cells[i][j] == today)
				dc.SetTextColor(CMAGENTA);
			else
				dc.SetTextColor(CBLACK);
			
			// set the correct back color to avoid color collision
			if (Cells[i][j] == day)
				dc.SetBkColor(CBLUE);
			else
				dc.SetBkColor(CWHITE);

			if (Cells[i][j] != 0)
			{
				CString str;
				str.Format(_T("%2d"), Cells[i][j]);
				dc.TextOut((rect.left + rect.right) / 2, rect.top + 2, str);
			}
		}
}

void CDlgDatePicker::DrawButton()
{
	m_pButtonLeftYear->Create(_T("<<"), WS_CHILD | WS_VISIBLE, 
		rectLeftYear, this, IDR_BUTTONLEFTYEAR);
	m_pButtonLeftMonth->Create(_T("<"), WS_CHILD | WS_VISIBLE, 
		rectLeftMonth, this, IDR_BUTTONLEFTMONTH);
	m_pButtonRightMonth->Create(_T(">"), WS_CHILD | WS_VISIBLE, rectRightMonth, 
		this, IDR_BUTTONRIGHTMONTH);
	m_pButtonRightYear->Create(_T(">>"), WS_CHILD | WS_VISIBLE, 
		rectRightYear, this, IDR_BUTTONRIGHTYEAR);
}

void CDlgDatePicker::DrawTitle(CPaintDC& dc)
{
	dc.SetTextColor(CBLACK);
	dc.SetTextAlign(TA_CENTER);
	dc.SetBkColor(CWHITE);

	// the bottom and bottom margin should be large than normal to 
	// avoid the 3d shape of the dialog box
	CRect rect(rectLeftMonth.right + 3, rectRightMonth.top, 
	rectRightMonth.left - 3, rectRightMonth.bottom - 2);
	dc.FillRect(rect, &brhWhite);
	CString str;
	str.Format(_T("%04d/%02d/%02d"), m_CurrentDate.GetYear(), m_CurrentDate.
		GetMonth(), m_CurrentDate.GetDay());
	dc.TextOut((rect.left + rect.right) / 2, rect.top + 4, str);
}

int CDlgDatePicker::GetDaysInMonth(int year, int month)
{
	if (month == 1 || month ==3 || month == 5 || month == 7 || month
		== 8 || month == 10 || month == 12)
		return 31;
	if (month == 4 || month == 6 || month == 9 || month == 11)
		return 30;
	if (month == 2 && (year % 100 == 0 || (year % 100 != 0 && year % 4 == 0)))
		return 29;
	else
		return 28;
}

void CDlgDatePicker::CalcCells()
{
	COleDateTime dtFirstDay;
	dtFirstDay.SetDate(year, month, 1);
	int nStartColumn = dtFirstDay.GetDayOfWeek() - 1;
	int i, j, k;
	for (i = 0; i < 6; i ++)
		for (j = 0; j < 7; j ++)
			Cells[i][j] = 0;

	int nDaysTotal = GetDaysInMonth(dtFirstDay.GetYear(), dtFirstDay.
		GetMonth());
	j = 0;
	i = nStartColumn;
	for (k = 1; k <= nDaysTotal; k ++)
	{
		Cells[j][i ++] = k;
		if (i == 7)
		{
			j ++;
			i = 0;
		}
	}
}

void CDlgDatePicker::ShiftYear(int nStep)
{
	year += nStep;
	m_CurrentDate.SetDate(year, month, day);
	if (m_CurrentDate.GetStatus() == COleDateTime::invalid)
	{
		// maybe the date is not leap year, but the current month is 2
		// and the current day is 29
		if (day == 29)
		{
			day = 28;
			m_CurrentDate.SetDate(year, month, 28);
		}
		// or the year value exceeds the range of the COleDateTime
		else
		{
			year = 2000;
			m_CurrentDate.SetDate(2000, month, day);
		}
	}
	
	ASSERT(m_CurrentDate.GetStatus() == COleDateTime::valid);
}

void CDlgDatePicker::ShiftMonth(int nStep)
{
	month += nStep;

	while (month > 12)
	{
		year += 1;
		month -= 12;
	}

	while (month < 1)
	{
		year -= 1;
		month += 12;
	}

	m_CurrentDate.SetDate(year, month, day);
	if (m_CurrentDate.GetStatus() == COleDateTime::invalid && 
		day == 29)
	{
		day = 28;
		m_CurrentDate.SetDate(year, month, 28);
	}

	ASSERT(m_CurrentDate.GetStatus() == COleDateTime::valid);
}

void CDlgDatePicker::OnLeftYear()
{
	ShiftYear(-1);
	CalcCells();
	Invalidate();
	m_SpinYear.SetPos(year);
}

void CDlgDatePicker::OnLeftMonth()
{
	ShiftMonth(-1);
	CalcCells();
	Invalidate();
	m_ComboMonth.SetCurSel(month - 1);
}

void CDlgDatePicker::OnRightYear()
{
	ShiftYear(1);
	CalcCells();
	Invalidate();
	m_SpinYear.SetPos(year);
}

void CDlgDatePicker::OnRightMonth()
{
	ShiftMonth(1);
	CalcCells();
	Invalidate();
	m_ComboMonth.SetCurSel(month - 1);
}

void CDlgDatePicker::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	// set the chosen date as the uers's choice, terminate the 
	// dialog box
	CDC * pDC = GetDC();
	pDC->DPtoLP(&point);
	ReleaseDC(pDC);

	int x = (point.x - XMARGIN) / CELLWIDTH;
	int y = (point.y - (YMARGIN + COMBOHEIGHT)) / CELLHEIGHT - 1;
	if (x >= 0 && x <= 6 && y >= 0 && y <= 5 && Cells[y][x] != 0)
	{
		m_Date = m_CurrentDate;
		m_pCtrl->SetDate(m_Date.GetYear(), m_Date.GetMonth(), m_Date.GetDay());
		DestroyWindow();
	}
}

void CDlgDatePicker::OnSelchangeComboMonth() 
{
	ShiftMonth(m_ComboMonth.GetCurSel() + 1 - month);
	CalcCells();
	Invalidate();
}

void CDlgDatePicker::OnDeltaposSpinYear(NMHDR* pNMHDR, LRESULT* pResult) 
{
	NM_UPDOWN* pNMUpDown = (NM_UPDOWN*)pNMHDR;

	if (year + pNMUpDown->iDelta <= 5000)
	{
		ShiftYear(pNMUpDown->iDelta);
		CalcCells();
		Invalidate();
	}

	*pResult = 0;
}

void CDlgDatePicker::OnChangeEditYear() 
{
	if (!IsWindow(m_EditYear.m_hWnd))
		return;

	CString strYear;
	m_EditYear.GetWindowText(strYear);
	int nNewYear = atoi(strYear);

	if (nNewYear <= 5000)
	{
		ShiftYear(nNewYear - year);
		CalcCells();
		Invalidate();
	}
}

void CDlgDatePicker::OnActivate(UINT nState, CWnd* pWndOther, BOOL bMinimized) 
{
	CDialog::OnActivate(nState, pWndOther, bMinimized);
	
	if (nState == WA_INACTIVE)
		DestroyWindow();
}
