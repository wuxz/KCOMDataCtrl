// KCOMEditControl.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "KCOMEditControl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CKCOMEditControl

CKCOMEditControl::CKCOMEditControl(CGXGridCore* pGrid, UINT nID) :
	CGXEditControl(pGrid, nID)
{
}

CKCOMEditControl::~CKCOMEditControl()
{
}


BEGIN_MESSAGE_MAP(CKCOMEditControl, CGXEditControl)
	//{{AFX_MSG_MAP(CKCOMEditControl)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CKCOMEditControl message handlers

BOOL CKCOMEditControl::GetValue(CString &strResult)
{
	if (!CGXEditControl::GetValue(strResult))
		return FALSE;

	NeedStyle();

	CString strUserAttrib;
	COleCurrency cur;
	if (m_pStyle->GetIncludeUserAttribute(GX_IDS_UA_DATATYPE))
	{
		m_pStyle->GetUserAttribute(GX_IDS_UA_DATATYPE, strUserAttrib);

		switch (atoi(strUserAttrib))
		{
		case 1:
			strResult = (atoi(strResult) == 0) ? _T("0") : _T("-1");
			break;

		case 2:
			strResult.Format("%d", (char)atoi(strResult));
			break;

		case 3:
			strResult.Format("%d", (short)atoi(strResult));
			break;

		case 4:
			strResult.Format("%d", atoi(strResult));
			break;

		case 5:
			strResult.Format("%f", atof(strResult));
			break;

		case 6:
			cur.ParseCurrency(strResult);
			if (cur.GetStatus() == COleCurrency::valid)
				strResult = cur.Format();
			else
				strResult.Empty();

			break;
		}
	}
	
	return TRUE;
}