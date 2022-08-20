// KCOMDateControl.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "KCOMDateControl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CKCOMDateControl

IMPLEMENT_CONTROL(CKCOMDateControl, CGXEditControl);
IMPLEMENT_DYNAMIC(CKCOMDateControl, CGXEditControl);

CKCOMDateControl::CKCOMDateControl(CGXGridCore* pGrid, UINT nID) : CGXEditControl(pGrid, nID)
{
	nCurrentPosition = 0;

	m_bMingGuo = FALSE;
	m_bFormatMingGuo = FALSE;
	nYearFinish = nMonthFinish = nDayFinish = 0;
	m_bTextNull = TRUE;

	strFormat = _T("yyyy/mm/dd");
	strData.Empty();
	strCurrentData.Empty();
	year = 0;
	month = 0;
	day = 0;
	m_PrevData.wYear = year;
	m_PrevData.wMonth = month;
	m_PrevData.wDay = day;
}

CKCOMDateControl::~CKCOMDateControl()
{
}


BEGIN_MESSAGE_MAP(CKCOMDateControl, CWnd)
	//{{AFX_MSG_MAP(CKCOMDateControl)
	ON_WM_KEYDOWN()
	ON_WM_CHAR()
	ON_WM_KILLFOCUS()
	ON_WM_SETFOCUS()
	ON_WM_LBUTTONDOWN()
	ON_WM_KEYUP()
	ON_WM_SYSCHAR()
	ON_WM_SYSKEYDOWN()
	ON_WM_SYSKEYUP()
	ON_WM_SHOWWINDOW()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CKCOMDateControl message handlers

void CKCOMDateControl::Init(ROWCOL nRow, ROWCOL nCol)
{
	CGXEditControl::Init(nRow, nCol);

	if (!m_hWnd)
		CreateControl();

	NeedStyle();

	nCurrentPosition = 0;

	m_bMingGuo = FALSE;
	m_bFormatMingGuo = FALSE;
	nYearFinish = nMonthFinish = nDayFinish = 0;
	m_bTextNull = TRUE;

	strFormat = _T("yyyy/mm/dd");
	strData.Empty();
	strCurrentData.Empty();
	year = 0;
	month = 0;
	day = 0;
	m_PrevData.wYear = year;
	m_PrevData.wMonth = month;
	m_PrevData.wDay = day;

	CString s = Grid()->GetExpressionRowCol(nRow, nCol);

	// hook for setting the window text
	// optimal override for masked edit
	SetValue(s);

	SetModify(FALSE);
}

BOOL CKCOMDateControl::GetValue(CString &strResult)
{
	if (!CGXEditControl::GetValue(strResult))
		return FALSE;

	CalcData();
	CalcText();
	if (m_bTextNull)
	{
		strResult.Empty();
		return TRUE;
	}

	COleDateTime dt;
	if (m_bMingGuo || m_bFormatMingGuo)
		strResult = strData;
	else
	{
		dt.SetDate(year, month, day);
		if (dt.GetStatus() == COleDateTime::valid)
			strResult = dt.Format();
		else
			strResult.Empty();
	}

	return TRUE;
}

void CKCOMDateControl::SetValue(LPCTSTR pszRawValue)
{
	// Convert the value to the text which should be
	// displayed in the current cell and show this
	// text.

	if (m_hWnd)
	{
		NeedStyle();
		InitMask(*m_pStyle);
	}
	
	COleDateTime dt;
	
	if (_tcslen(pszRawValue) == 0)
		SetTextIn(_T(""));
	else if(m_bMingGuo || m_bFormatMingGuo)
		SetTextIn(pszRawValue);
	else
	{
		dt.ParseDateTime(pszRawValue);
		if (dt.GetStatus() == COleDateTime::valid)
			SetDate(dt.GetYear(), dt.GetMonth(), dt.GetDay());
		else
			SetTextIn(_T(""));
	}
}

BOOL CKCOMDateControl::Store()
{
	// Calls SetStyleRange() and resets the modify flag
	ASSERT(m_pStyle);

	CString sValue;
	if (m_pStyle && GetModify() && GetValue(sValue))
	{
		SetActive(FALSE);
		FreeStyle();

		return Grid()->SetExpressionRowCol(
			m_nRow,  m_nCol,
			sValue,
			GX_INVALIDATE);
			// Cell will be automatically redrawn inactive

		// If text is "=xxxx" this will become a formula cell
	}

	return TRUE;
}

BOOL CKCOMDateControl::OnGridKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	if (!m_hWnd)
		return FALSE;

	BOOL bProcessed = FALSE;

	BOOL bCtl = GetKeyState(VK_CONTROL) & 0x8000;
	BOOL bShift = GetKeyState(VK_SHIFT) & 0x8000;

	// Pass all CTRL keys to the grid
	if (!bCtl)
	{
		switch (nChar)
		{
		case VK_DELETE:
			if (Grid()->GetParam()->GetRangeList()->IsEmpty() // no ranges of cells selected
				&& !IsReadOnly()        // cell is not readonly
				&& !bShift
				&& OnDeleteCell()       // OnDeleteCell notification returns TRUE
				&& OnStartEditing())    // OnStartEditing notification returns TRUE
			{
				SetActive(TRUE);

				// Delete text
				SetText(_T(""));
				SetModify(TRUE);

				bProcessed = TRUE;
			}
			break;

		case VK_END:
			if (OnStartEditing())
			{
				if (!IsInit())
					Init(m_nRow, m_nCol);
				SetActive(TRUE);

				CString s;
				GetValue(s);
				int nLen = _tcslen(s);

				// position caret
//wxz				SetSel(nLen, nLen, FALSE);

				bProcessed = TRUE;
			}
			break;

		case VK_HOME:
			if (OnStartEditing())
			{
				if (!IsInit())
					Init(m_nRow, m_nCol);
				SetActive(TRUE);

				// position caret
//wxz				SetSel(0, 0, FALSE);

				bProcessed = TRUE;
			}
			break;
		}
	}

	if (bProcessed)
	{
		// eventually destroys and creates CEdit with appropriate window style
		UpdateEditStyle();

		if (!Grid()->ScrollCellInView(m_nRow, m_nCol))
			Refresh();

		return TRUE;
	}

	return CGXControl::OnGridKeyDown(nChar, nRepCnt, nFlags);
}

void CKCOMDateControl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	switch (nChar)
	{
		case VK_HOME:
		{
			MoveLeft(nCurrentPosition);
		}
		break;

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

		case VK_RIGHT:
		{
			MoveRight(1);
		}
		break;

		case VK_DELETE:
		{
			OnChar(VK_DELETE, 1, 1);
		}
		break;

		case VK_DOWN:
		case VK_UP:
			Grid()->ProcessKeys(this, WM_KEYDOWN, nChar, nRepCnt, nFlags);
			break;
	}
}

void CKCOMDateControl::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	BOOL bFinish = TRUE;

	if (IsReadOnly())
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
				SetText("");
			if (strData.IsEmpty() || strData.Right(1) != '_')
				strData += CString("_");
			MoveLeft(1);
			SetWindowText(strData);
			SetModify(TRUE);
			return;
		}

		if (nChar == VK_DELETE)
		{
			strData = strData.Left(nCurrentPosition) + strData.
				Right(strData.GetLength() - nCurrentPosition - 1);
			if (strData.IsEmpty() || strData == '_')
				SetText("");
			if (strData.IsEmpty() || strData.Right(1) != '_')
				strData += CString("_");
			MoveLeft();
			SetWindowText(strData);
			SetModify(TRUE);
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

		MoveRight(1);
		SetWindowText(strData);
		SetModify(TRUE);
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
		if (nChar == '-' && !(m_bFormatMingGuo
			&& nCurrentPosition == nYearPosition[0]))
			return;

		if (!bFinish)
		{
			for (i = nCurrentPosition; i <= nYearPosition[1]; i++)
				strData.SetAt(i, '_');

			nYearFinish = nCurrentPosition - nYearPosition[0];
			if (nYearFinish == 0 && nMonthFinish == 0 && nDayFinish == 0)
			{
				SetText("");
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
			SetWindowText(strData);
			SetModify(TRUE);
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
			SetWindowText(strData);
			SetModify(TRUE);
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
				nMonthFinish = max(nMonthFinish, nCurrentPosition - nMonthPosition[0] + 1);
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
			SetWindowText(strData);
			SetModify(TRUE);
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
						nDayFinish = max(nDayFinish, nCurrentPosition - nDayPosition[0] + 1);
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
				nDayFinish = max(nDayFinish, nCurrentPosition - nDayPosition[0] + 1);
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
		}

		SetWindowText(strData);
	}
	
	SetModify(TRUE);
}

void CKCOMDateControl::OnKillFocus(CWnd* pNewWnd) 
{
	CGXEditControl::OnKillFocus(pNewWnd);
}

void CKCOMDateControl::OnSetFocus(CWnd* pOldWnd) 
{
	CWnd::OnSetFocus(pOldWnd);
	UpdateEditStyle();
	
	// TODO: Add your message handler code here
	SetActive(TRUE);
	CalcText();
	MoveLeft(1);
	SetWindowText(strData);
}

void CKCOMDateControl::InitMask(const CGXStyle &style)
{
	CString sMask;
	style.GetUserAttribute(GX_IDS_UA_INPUTMASK, sMask);
	strFormat = _T("yyyy/mm/dd");

	SetDateFormat(sMask);
}

void CKCOMDateControl::CalcText()
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

	CString str, str2;
	str2.Empty();
	str.Format("%3d", abs(year - 1911));
	str.TrimLeft();
	str2 += str;
	str.Format("%02d", month);
	str2 += str;
	str.Format("%02d", day);
	str2 += str;
	if (year < 1911)
		str2 = "-" + str2;
	m_bMingGuo = strFormat == "ggggggg" ? TRUE : FALSE;
	if (m_bMingGuo)
	{
		if (!m_bTextNull)
		{
			strData = str2;
			if(str2.GetLength() < 7)
				strData += "_";
			m_strText = str2;
		}
		else
		{
			strData.Empty();
			m_strText.Empty();
		}
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
						strYear.Format("%04d", year);
					strYear = strYear.Right(nYearLength);
					strData += strYear;
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
						strYear.Format("%03d", year - 1911);
					strYear = strYear.Right(nYearLength);
					strData += strYear;
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
						strMonth.Format("%02d", month);
					strMonth = strMonth.Right(nMonthLength);
					strData += strMonth;
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
						strDay.Format("%02d", day);
					strDay = strDay.Right(nDayLength);
					strData += strDay;
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

	m_strText = strData;
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
	
	if (m_bMingGuo && !m_bTextNull)
		m_strText.Format("%03d%02d%02d", year - 1911, month, day);
}

void CKCOMDateControl::MoveRight(int nSteps)
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
		SetSel(nCurrentPosition, nCurrentPosition);
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
	SetSel(nCurrentPosition, nCurrentPosition);
}

void CKCOMDateControl::MoveLeft(int nSteps)
{
	if (m_bMingGuo)
	{
		nCurrentPosition -= nSteps;
		if (nCurrentPosition >= strData.GetLength())
			nCurrentPosition = strData.GetLength() -1;
		if (nCurrentPosition < 0)
			nCurrentPosition = 0;
		SetSel(nCurrentPosition, nCurrentPosition);
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
	SetSel(nCurrentPosition, nCurrentPosition);
}

void CKCOMDateControl::SetDateFormat(LPCTSTR lpszNewValue) 
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
		SetWindowText(strData);
	}
}

BOOL CKCOMDateControl::IsValidFormat(CString strNewFormat)
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

void CKCOMDateControl::SetText(LPCTSTR lpszNewValue) 
{
	if (!IsInit() || !m_hWnd)
		return;

	if (!IsActive())
		SetActive(TRUE);

	if (SetTextIn(lpszNewValue))
	{
		SetWindowText(strData);
		SetModify(TRUE);
	}
	else
	{
		SetTextIn("");
		SetWindowText(m_strText);
		OnModifyCell();
		SetModify(TRUE);
	}
}

BOOL CKCOMDateControl::IsFinished()
{
	return (nYearFinish >= nYearLength && nMonthFinish >= nMonthLength
		&& nDayFinish >= nDayLength);
}

void CKCOMDateControl::CalcMingGuo()
{
	int i;
	char chr;
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

void CKCOMDateControl::SetDate(short nYear, short nMonth, short nDay) 
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

	if (m_hWnd)
		SetWindowText(strData);
}

BOOL CKCOMDateControl::SetTextIn(CString strNewValue)
{
	CalcText();

	CString strData = strNewValue;
	int i;
	char chr;
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
			else if (chr == '_')
			{
				bValid = FALSE;
				strTotal.Empty();
				break;
			}
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

void CKCOMDateControl::CalcData()
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

void CKCOMDateControl::SetCurrentText(const CString &str)
{
	SetText(str);
}

void CKCOMDateControl::SetWindowText(CString strNewText)
{
	DWORD dwSel;
	if (IsWindow(m_hWnd))
	{
		dwSel = CEdit::GetSel();

		CEdit::SetWindowText(strNewText);
		CEdit::SetSel(dwSel);
	}
}

BOOL CKCOMDateControl::GetControlText(CString &strDisplay, ROWCOL nRow, ROWCOL nCol, LPCTSTR pszRawValue, const CGXStyle &style)
{
	InitMask(style);

	if (pszRawValue == NULL)
		pszRawValue = style.GetValueRef();

	if (_tcslen(pszRawValue) > 0)
	{
		SetValue(pszRawValue);

		strDisplay = m_strText;
	}
	else
		strDisplay.Empty();

	return TRUE;
}

void CKCOMDateControl::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CEdit::OnLButtonDown(nFlags, point);

	int nStart, nEnd;
	GetSel(nStart, nEnd);
	if (nEnd < 0)
		nEnd = 0;
	else if(nEnd >= strFormat.GetLength())
		nEnd = strFormat.GetLength() - 1;

	nCurrentPosition = nEnd;
	SetSel(nEnd, nEnd);
	MoveLeft();
}

BOOL CKCOMDateControl::OnGridChar(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	if (IsReadOnly())
		return FALSE;

	if (IsActive())
	{
		OnChar(nChar, 1, 1);

		return TRUE;
	}
	else
	{
		// discard key if cell is read only
		if (IsReadOnly() || !OnStartEditing())
			return FALSE;

		SetActive(TRUE);

		SetText(_T(""));
		SetModify(TRUE);
		SetSel(0, 0);
		OnChar(nChar, 1, 1);
		Refresh();

		return TRUE;
	}
}

void CKCOMDateControl::Draw(CDC *pDC, CRect rect, ROWCOL nRow, ROWCOL nCol, const CGXStyle &style, const CGXStyle *pStandardStyle)
{
	ASSERT(pDC != NULL && pDC->IsKindOf(RUNTIME_CLASS(CDC)));
	// ASSERTION-> Invalid Device Context ->END
	ASSERT(nRow <= Grid()->GetRowCount() && nCol <= Grid()->GetColCount());
	// ASSERTION-> Cell coordinates out of range ->END

	// Notes on Draw:
	// if (nRow == m_nRow && nCol == m_nCol && m_bIsActive)
	//     initialize CWnd, make it visible and set focus
	// else {
	//     lookup style and draw static, e.g.:
	//     - edit control with GXDrawTextLikeMultiLineEdit
	// }

	ASSERT_VALID(pDC);

	// if (nRow == m_nRow && nCol == m_nCol && m_bIsActive)
	//     initialize CWnd, make it visible and set focus
	if (nRow == m_nRow && nCol == m_nCol && IsActive() && !Grid()->IsPrinting())
	{
		// ------------------ Active Window ---------------------------

		Hide();

		// Font
		if (m_font.GetSafeHandle())
			m_font.DeleteObject();

		CFont* pOldFont = 0;

		CGXFont font(m_pStyle->GetFontRef());
		font.SetSize(font.GetSize()*Grid()->GetZoom()/100);
		font.SetOrientation(0);
		
		// no vertical font for editing the cell
		const LOGFONT& _lf = font.GetLogFontRef();
		if ( !m_font.CreateFontIndirect(&_lf) )
		{
			TRACE1("Failed to create font '%s' in CKCOMDateControl::Draw\n",
				(LPCTSTR) m_pStyle->GetFontRef().GetDescription());
		}
		
		// Cell-Color
		COLORREF rgbText = m_rgbWndHilightText;
		COLORREF rgbCell = m_pStyle->GetInteriorRef().GetColor();

		if (Grid()->GetParam()->GetProperties()->GetBlackWhite())
		{
			rgbText = RGB(0,0,0);
			rgbCell = RGB(255,255,255);
		}

		// Background + Frame
		DrawBackground(pDC, rect, *m_pStyle, TRUE);

		CRect rectText = GetCellRect(nRow, nCol, rect);

		pOldFont = pDC->SelectObject(&m_font);

		TEXTMETRIC    tm;
		pDC->GetTextMetrics(&tm);
		m_dyHeight = tm.tmHeight + tm.tmExternalLeading + 1;
		m_dxMaxCharWidth = tm.tmMaxCharWidth;

		if (pOldFont)
			pDC->SelectObject(pOldFont);

		if (!m_pStyle->GetWrapText()) //  && !m_pStyle->GetAllowEnter())
		{
			// vertical alignment for single line text
			// is done by computing the window rectangle for the CEdit
			if (m_pStyle->GetIncludeVerticalAlignment())
			{
				switch (m_pStyle->GetVerticalAlignment())
				{
				case DT_VCENTER:
					rectText.top += max(0, (rectText.Height() - m_dyHeight) /2);
					break;
				case DT_BOTTOM:
					rectText.top = max(rectText.top, rectText.bottom-m_dyHeight);
				break;
				}
			}
			rectText.bottom = min(rectText.bottom, rectText.top + m_dyHeight - 1);
		}

		DWORD ntAlign = m_pStyle->GetHorizontalAlignment() | m_pStyle->GetVerticalAlignment();

		if (ntAlign & DT_RIGHT)
			rectText.right++;

		SetFont(&m_font);

		MoveWindow(rectText, FALSE);
		m_rgbCell = rgbCell;
		m_ntAlign = ntAlign;

		SetCellColor(rgbCell);
		SetTextAlign((UINT) ntAlign);

		SetActive(TRUE);

		SetReadOnly(IsReadOnly());

		if (GetFocus() == GridWnd())
		{
			Grid()->SetIgnoreFocus(TRUE);
			SetFocus();
			Grid()->SetIgnoreFocus(FALSE);
		}

		ShowWindow(SW_SHOW);

		if (nRow > Grid()->GetFrozenRows() && (Grid()->GetTopRow() > nRow 
			|| nCol > Grid()->GetFrozenCols() && Grid()->GetLeftCol() > nCol))
		// Ensure that the window cannot draw outside the clipping area!
		{
			CRect rectClip;
			if (pDC->GetClipBox(&rectClip) != ERROR)
			{
				CRect r = rect & Grid()->GetGridRect();
				GridWnd()->ClientToScreen(&r);
				ScreenToClient(&r);
				GridWnd()->ClientToScreen(&rectClip);
				ScreenToClient(&rectClip);
				ValidateRect(r); 
				InvalidateRect(rectClip); 
				HideCaret();
			}
		}

		UpdateWindow();

		// child Controls: spin-buttons, hotspot, combobox btn, ...
		CGXControl::Draw(pDC, rect, nRow, nCol, *m_pStyle, pStandardStyle);

		ExcludeClipRect(rectText);
	}
	else
	// else {
	//     lookup style and draw static:
	//     - edit control with GXDrawTextLikeMultiLineEdit
	// }
	{

		// ------------------ Static Text ---------------------------

		CGXStatic::Draw(pDC, rect, nRow, nCol, style, pStandardStyle);
	}
}

void CKCOMDateControl::Hide()
{
	// Hides the CEdit without changing the m_bIsActive flag
	// and without invalidating the screen
	if (m_hWnd && IsWindowVisible())
	{
		Grid()->SetIgnoreFocus(TRUE);
		HideCaret();
		MoveWindow0(FALSE);
		ShowWindow(SW_HIDE);
		Grid()->SetIgnoreFocus(FALSE);
	}
}

void CKCOMDateControl::OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (!Grid()->ProcessKeys(this, WM_KEYUP, nChar, nRepCnt, nFlags))
	{
		CEdit::OnKeyUp(nChar, nRepCnt, nFlags);
		UpdateEditStyle();
	}
}

void CKCOMDateControl::OnSysChar(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (!Grid()->ProcessKeys(this, WM_SYSCHAR, nChar, nRepCnt, nFlags))
		CEdit::OnSysChar(nChar, nRepCnt, nFlags);
}

void CKCOMDateControl::OnSysKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (!Grid()->ProcessKeys(this, WM_SYSKEYDOWN, nChar, nRepCnt, nFlags))
		CEdit::OnSysKeyDown(nChar, nRepCnt, nFlags);
}

void CKCOMDateControl::OnSysKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (!Grid()->ProcessKeys(this, WM_SYSKEYUP, nChar, nRepCnt, nFlags))
		CEdit::OnSysKeyUp(nChar, nRepCnt, nFlags);
}

void CKCOMDateControl::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	if (bShow)
		UpdateEditStyle();

	CEdit::OnShowWindow(bShow, nStatus);
}

BOOL CKCOMDateControl::UpdateEditStyle()
{
	if (m_bUpdateEditStyle || !m_hWnd)
		// semaphor for not recursively re-entering UpdateEditStyle()
		return FALSE;

	// TRACE0("CKCOMDateControl: BEGIN UpdateEditStyle()\n");

	int nLine;
	RECT rect;
	GetClientRect(&rect);

	RECT rectNP;
	DWORD dwEditStyle = CalcEditStyle(&rectNP, m_dyHeight);
	BOOL bChanged = m_dwEditStyle != dwEditStyle;

	int ns, ne;
	GetSel(ns, ne);

	if (bChanged)
	{
		m_bUpdateEditStyle = TRUE;

		BOOL bModify   = GetModify();
		BOOL bReadOnly = (BOOL) (GetStyle() & ES_READONLY);

		MapWindowPoints(GridWnd(), &rect);

		CString s;
		GetWindowText(s);

		CFont *pFont = GetFont();

		HideCaret();
		MoveWindow0(FALSE);

		Grid()->SetIgnoreFocus(TRUE);

		DestroyWindow();

		// TRACE0("CKCOMDateControl: Destroying and recreating EDIT-Window\n");

		VERIFY(CreateControl(m_dwEditStyle = dwEditStyle, &rect));

		// not supported by Wind/U and Mac
		SetRect(&rectNP);
		SetFont(pFont);
		SetWindowText(s);
		SetSel(ns, ne, FALSE);
		SetModify(bModify);
		SetReadOnly(bReadOnly);

		m_bUpdateEditStyle = FALSE;
		nLine = LineFromChar(ne);

		HideCaret();
		UpdateWindow();
		ShowCaret();
		SendMessage(EM_SCROLLCARET, 0, 0);

		SetFocus();

		Grid()->SetIgnoreFocus(FALSE);
	}
	else
	{
		nLine = LineFromChar();
		// not supported by Wind/U and Mac
		SetRectNP(&rectNP);
	}

	int nLast = GetLastVisibleLine(&rect);
	int nFirst = GetFirstVisibleLine();

	// TRACE(_T("CKCOMDateControl: FirstVisible=%d,LineFromChar=%d,LastVisible=%d\n"),
	//  nFirst, nLine, nLast);

	if (nLine > nLast || nLine < nFirst)
	{
		HideCaret();

		if (nLine > nLast)
			LineScroll(nLine - nLast);
		else
			LineScroll(nLine - nFirst);
		SendMessage(EM_SCROLLCARET, 0, 0);
		UpdateWindow();
		ShowCaret();
	}
	else
	{
		SendMessage(EM_SCROLLCARET, 0, 0);
	}

	// TRACE1("CKCOMDateControl: END   UpdateEditStyle() RET %d\n", bChanged);

	return bChanged;
}
