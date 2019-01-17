// Localizer.cpp: implementation of the CLocalizer class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Localizer.h"
#include "FileMisc.h"
#include "enstring.h"
#include "enmenu.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

/////////////////////////////////////////////////////////////////////////////
// CLocDialog dialog

class CLocDialog : public CDialog
{
// Construction
public:
	CLocDialog(LPCTSTR szIDTemplate, CWnd* pParent = NULL) : CDialog(szIDTemplate, pParent) {}

	BOOL OnInitDialog()
	{
		CDialog::OnInitDialog();

		// move off screen before showing
		CRect rPos;
		GetWindowRect(rPos);
		rPos.OffsetRect(-30000, -30000);
		MoveWindow(rPos, FALSE);

		return TRUE;
	}

// Implementation
protected:
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus)
	{
		CDialog::OnShowWindow(bShow, nStatus);
		PostMessage(WM_CLOSE);
	}

	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CLocDialog, CDialog)
	ON_WM_SHOWWINDOW()
END_MESSAGE_MAP()

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

#define NEEDSINIT ((ITransText*)-1)

ITransText* CLocalizer::s_pTransText = NEEDSINIT; // we only initialize once even if it fails
BOOL CLocalizer::s_bOwner = FALSE; // did we create or was it attached

BOOL CLocalizer::Initialize(LPCTSTR szDictFile, ITT_TRANSLATEOPTION nOption)
{
	if (s_pTransText == NEEDSINIT)
	{
		// Load TransText module
		CString sTTDll = FileMisc::GetAppFolder() + _T("\\TransText.dll");
		
		if (IsTransTextDll(sTTDll) && (nOption != ITTTO_TRANSLATEONLY || FileMisc::FileExists(szDictFile)))
		{
			s_pTransText = CreateTransTextInterface(sTTDll, szDictFile, nOption);
			ASSERT(s_pTransText);

			CEnString::SetLocalizer(s_pTransText);
			CEnMenu::SetLocalizer(s_pTransText);

			s_bOwner = TRUE;
		}
		else
		{
			s_pTransText = NULL;
			s_bOwner = FALSE;
		}
	}

	return (s_pTransText != NEEDSINIT);
}

BOOL CLocalizer::Initialize(ITransText* pTT)
{
	// check it's valid
	if (!ValidLocalizer(pTT))
		return FALSE;

	// check it's a different interface
	if (pTT == s_pTransText)
		return TRUE;

	// release existing interface
	Release();

	// assign as non-owner
	s_pTransText = pTT;
	s_bOwner = FALSE;

	CEnString::SetLocalizer(s_pTransText);
	CEnMenu::SetLocalizer(s_pTransText);

	return TRUE;
}

BOOL CLocalizer::ValidLocalizer(ITransText* pTT)
{
	return (pTT && (pTT != NEEDSINIT));
}

BOOL CLocalizer::ValidLocalizer()
{
	return ValidLocalizer(s_pTransText);
}

void CLocalizer::Release()
{
	if (ValidLocalizer() && s_bOwner)
		s_pTransText->Release();

	s_pTransText = NEEDSINIT;
}

void CLocalizer::SetTranslationOption(ITT_TRANSLATEOPTION nOption)
{
	if (ValidLocalizer())
		s_pTransText->SetTranslationOption(nOption);
}

ITT_TRANSLATEOPTION CLocalizer::GetTranslationOption()
{
	if (ValidLocalizer())
		return s_pTransText->GetTranslationOption();

	// else
	return ITTTO_TRANSLATEONLY;
}

CString CLocalizer::GetDictionaryPath()
{
	if (ValidLocalizer())
		return s_pTransText->GetDictionaryFile();

	return _T("");
}

CString CLocalizer::GetDictionaryVersion()
{
	if (ValidLocalizer())
		return s_pTransText->GetDictionaryVersion();

	return _T("");
}

BOOL CLocalizer::GetDictionaryVersion(CDWordArray& aVersionParts)
{
	CString sVersion = GetDictionaryVersion();

	return FileMisc::SplitVersionNumber(sVersion, aVersionParts);
}

BOOL CLocalizer::CleanupDictionary(LPCTSTR szMasterDictPath)
{
	if (ValidLocalizer())
		return s_pTransText->CleanupDictionary(szMasterDictPath);

	return FALSE;
}

void CLocalizer::ForceTranslateAllUIElements(UINT nIDFirstStr, UINT nIDLastStr)
{
	CWnd* pAppWnd = AfxGetMainWnd();
	ASSERT(pAppWnd);

	if (pAppWnd == NULL)
		return;

	// we translate dialogs by displaying them
	for (int nDlg = 1; nDlg <= 0x6FFF; nDlg++)
	{
		LPCTSTR szDlgTemplate = MAKEINTRESOURCE(nDlg);
		HRSRC hResource = ::FindResource(NULL, szDlgTemplate, RT_DIALOG);
		
		if (hResource == NULL)
			continue;
		
		CWaitCursor wait;
		CLocDialog dialog(szDlgTemplate, pAppWnd);
		
		dialog.DoModal();
	}

	CWaitCursor wait;

	// we translate menus directly
	for (int nMenu = 1; nMenu <= 0x6FFF; nMenu++)
	{
		LPCTSTR szMenuDef = MAKEINTRESOURCE(nMenu);
		HRSRC hResource = ::FindResource(NULL, szMenuDef, RT_MENU);
		CMenu menu;
		
		if (hResource == NULL)
			continue;
		
		VERIFY(menu.LoadMenu(nMenu));

		CLocalizer::TranslateMenu(menu);
	}

	// we translate tooltips by loading the toolbar and
	// iterating its buttons directly
	CList<UINT, UINT> lstTooltipIDs;

	// dummy tooltipctrl to pass to CLocalizer
	CToolTipCtrl tt;

	for (int nTBar = 1; nTBar <= 0x6FFF; nTBar++)
	{
		LPCTSTR szTBDef = MAKEINTRESOURCE(nTBar);
		HRSRC hResource = ::FindResource(NULL, szTBDef, RT_TOOLBAR);
		CToolBar toolbar;
		
		if (hResource == NULL)
			continue;
		
		VERIFY (toolbar.CreateEx(pAppWnd, 0, WS_CHILD | CBRS_TOOLTIPS) && toolbar.LoadToolBar(nTBar));

		int nBtnCount = toolbar.GetToolBarCtrl().GetButtonCount();
		CString sTip;
		
		for (int nBtn = 0; nBtn < nBtnCount; nBtn++)
		{
			UINT nID = toolbar.GetItemID(nBtn);

			if (sTip.LoadString(nID) && !sTip.IsEmpty())
			{
				// tip text starts at '\n' 
				int nStartTip = sTip.Find('\n');
				
				if (nStartTip != -1) 
					sTip = sTip.Right(sTip.GetLength() - nStartTip - 1);
				
				else // strip '...' if present
					sTip.Replace(_T("."), _T(""));
				
				sTip.TrimLeft();
				sTip.TrimRight();

				if (!sTip.IsEmpty())
				{
					// create tooltip ctrl first time we need it
					if (tt.GetSafeHwnd() == NULL)
						tt.Create(pAppWnd, TTS_ALWAYSTIP);
					
					CLocalizer::TranslateText(sTip, tt);

					// keep track of tooltips we've translated
					// so we don't re-handle them when we translate
					// the rest of the strings below
					lstTooltipIDs.AddTail(nID);
				}
			}
		}
	}

	// any left over strings but not AFX_IDS_
	nIDLastStr = min(nIDLastStr, (AFX_IDS_SCFIRST - 1));
	CString sStr;

	for (UINT nStr = nIDFirstStr; nStr <= nIDLastStr; nStr++)
	{
		// app related strings we are not interested in
		if ((nStr >= 0xE000 && nStr < 0xE006) || // AFX_IDS_
			(nStr >= 0xE700 && nStr < 0xE707) || // ID_INDICATOR_
			(nStr >= 0xE210 && nStr < 0xE230))   // ID_OLE_VERB_
			continue;

		// ignore previously translated tooltips
		if (lstTooltipIDs.Find(nStr))
			continue;

		if (sStr.LoadString(nStr) && !sStr.IsEmpty())
			CLocalizer::TranslateText(sStr);
	}
}

BOOL CLocalizer::TranslateText(CString& sText, HWND hWndRef)
{
	VERIFY (Initialize());

	return CEnString::TranslateString(sText, hWndRef);
}

BOOL CLocalizer::TranslateMenu(HMENU hMenu, HWND hWndRef, BOOL bRecursive)
{
	VERIFY (Initialize());

	if (s_pTransText)
		return s_pTransText->TranslateMenu(hMenu, hWndRef, bRecursive);

	return FALSE;
}

void CLocalizer::UpdateMenu(HWND hWnd)
{
	VERIFY (Initialize());

	if (s_pTransText)
		s_pTransText->UpdateMenu(hWnd);
}

void CLocalizer::EnableTranslation(HWND hWnd, BOOL bEnable)
{
	VERIFY (Initialize());

	if (s_pTransText)
		s_pTransText->EnableTranslation(hWnd, bEnable);
}

void CLocalizer::EnableTranslation(const CWnd& hWnd, BOOL bEnable)
{
	VERIFY (Initialize());

	if (s_pTransText)
		s_pTransText->EnableTranslation(hWnd.GetSafeHwnd(), bEnable);
}

void CLocalizer::EnableTranslation(HMENU hMenu, BOOL bEnable)
{
	VERIFY (Initialize());

	if (s_pTransText)
		s_pTransText->EnableTranslation(hMenu, bEnable);
}

void CLocalizer::EnableTranslation(UINT nMenuID, BOOL bEnable)
{
	VERIFY (Initialize());

	if (s_pTransText)
		s_pTransText->EnableTranslation(nMenuID, bEnable);
}

void CLocalizer::EnableTranslation(UINT nMenuIDFirst, UINT nMenuIDLast, BOOL bEnable)
{
	VERIFY (Initialize());

	if (s_pTransText)
	{
		for (UINT nMenuID = nMenuIDFirst; nMenuID <= nMenuIDLast; nMenuID++)
			s_pTransText->EnableTranslation(nMenuID, bEnable);
	}
}


void CLocalizer::IgnoreString(const CString& sText)
{
	VERIFY (Initialize());

	if (s_pTransText)
		s_pTransText->IgnoreString(sText);
}

void CLocalizer::IgnoreString(UINT nIDText)
{
	IgnoreString(CString(MAKEINTRESOURCE(nIDText)));
}
