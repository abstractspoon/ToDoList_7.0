// encolordialog.cpp : implementation file
//

#include "stdafx.h"
#include "encolordialog.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CEnColorDialog

// disable 'Add to Custom Colors' until we upgrade to later
// MFC version where CWinApp::WriteProfileInt is virtual
const BOOL CWINAPP_WRITEPROFILEINT_ISVIRTUAL = FALSE;

/////////////////////////////////////////////////////////////////////////////

IMPLEMENT_DYNAMIC(CEnColorDialog, CColorDialog)

CEnColorDialog::CEnColorDialog(COLORREF clrInit, DWORD dwFlags, CWnd* pParentWnd) :
	CColorDialog(clrInit, dwFlags, pParentWnd)
{
}

CEnColorDialog::~CEnColorDialog()
{
}

BEGIN_MESSAGE_MAP(CEnColorDialog, CColorDialog)
	//{{AFX_MSG_MAP(CEnColorDialog)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////

int CEnColorDialog::DoModal()
{
	if (CWINAPP_WRITEPROFILEINT_ISVIRTUAL)
	{
		// restore previously saved custom colors
		for (int nColor = 0; nColor < 16; nColor++)
		{
			CString sKey;
			sKey.Format(_T("CustomColor%d"), nColor);
			
			COLORREF color = (COLORREF)AfxGetApp()->GetProfileInt(_T("ColorDialog"), sKey, (int)RGB(255, 255, 255));
			m_cc.lpCustColors[nColor] = color;
		}
	}
	
	int nRet = CColorDialog::DoModal();
	
	if (CWINAPP_WRITEPROFILEINT_ISVIRTUAL)
	{
		if (nRet == IDOK)
		{
			// save any custom colors
			COLORREF* pColors = GetSavedCustomColors();
			
			for (int nColor = 0; nColor < 16; nColor++)
			{
				CString sKey;
				sKey.Format(_T("CustomColor%d"), nColor);
				
				int nColorVal = (int)pColors[nColor];
				AfxGetApp()->WriteProfileInt(_T("ColorDialog"), sKey, nColorVal);
			}
		}
	}
	
	return nRet;
}


BOOL CEnColorDialog::OnInitDialog()
{
	BOOL bRes = CColorDialog::OnInitDialog();
	
	if (!CWINAPP_WRITEPROFILEINT_ISVIRTUAL)
	{
		const UINT IDC_ADDTOCUSTOMCOLORS = 0x02C8;
		
		if (GetDlgItem(IDC_ADDTOCUSTOMCOLORS))
			GetDlgItem(IDC_ADDTOCUSTOMCOLORS)->EnableWindow(FALSE);
	}

	return bRes;
}
