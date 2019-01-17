// maskedit.cpp : implementation file
//

#include "stdafx.h"
#include "maskedit.h"
#include "winclasses.h"
#include "misc.h"
#include "clipboard.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CMaskEdit

IMPLEMENT_DYNAMIC(CMaskEdit, CEdit)

CMaskEdit::CMaskEdit(LPCTSTR szMask, DWORD dwFlags) : m_sMask(szMask), m_dwFlags(dwFlags)
{
	SetMask(szMask, dwFlags);
}

CMaskEdit::~CMaskEdit()
{
}

BEGIN_MESSAGE_MAP(CMaskEdit, CEdit)
	//{{AFX_MSG_MAP(CMaskEdit)
	ON_WM_CHAR()
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_SETTEXT, OnSetText)
	ON_MESSAGE(WM_PASTE, OnPaste)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CMaskEdit message handlers

void CMaskEdit::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	BOOL bSkipMaskCheck = FALSE;

	// preprocessing
	switch (nChar)
	{
	case VK_BACK:
		bSkipMaskCheck = TRUE;
		break;

	// something wierd going on with edit control accelerators
	case 3:   // c
	case 22:  // v
	case 24:  // x
		if (CWinClasses::IsEditControl(*this))
		{
			ASSERT (Misc::IsKeyPressed(VK_CONTROL));
			bSkipMaskCheck = TRUE;
		}
		break;

	case VK_RETURN:
		bSkipMaskCheck = (GetStyle() & ES_WANTRETURN);
		break;
		
	case '.':
		if (m_dwFlags & ME_LOCALIZEDECIMAL)
		{
			CString sDecimal = Misc::GetDecimalSeparator();
	
			if (!sDecimal.IsEmpty())
				nChar = sDecimal[0];
			else
				bSkipMaskCheck = TRUE;
		}
		break;
		
	case 'c':
	case 'C':
	case 'v':
	case 'V':
	case 'x':
	case 'X':
		bSkipMaskCheck = Misc::IsKeyPressed(VK_CONTROL);
		break;
	}

	if (!bSkipMaskCheck && !m_sMask.IsEmpty())
	{
		if ((m_dwFlags & ME_EXCLUDE) && m_sMask.Find((char)nChar) != -1)
		{
			return;
		}
		else if (!(m_dwFlags & ME_EXCLUDE) && m_sMask.Find((char)nChar) == -1)
		{
			return;
		}
	}

	DefWindowProc(WM_CHAR, (WPARAM)nChar, MAKELPARAM(nRepCnt, nFlags));
}

void CMaskEdit::SetMask(LPCTSTR szMask, DWORD dwFlags) 
{ 
	m_sMask = szMask; 
	m_dwFlags = dwFlags; 

	if (dwFlags & ME_LOCALIZEDECIMAL)
	{
		CString sDecimal = Misc::GetDecimalSeparator();

		if (!sDecimal.IsEmpty() && sDecimal != _T(".") && m_sMask.Find('.') != -1)
			m_sMask += sDecimal[0];
	}
}

LRESULT CMaskEdit::OnSetText(WPARAM wp, LPARAM lp)
{
	if (!IsValid((LPCTSTR)lp))
		return FALSE;
	
	return Default();
}

LRESULT CMaskEdit::OnPaste(WPARAM wp, LPARAM lp)
{
	CString sText;

	if (!m_sMask.IsEmpty() && CClipboard().GetText(sText))
	{
		sText.TrimLeft();
		sText.TrimRight();

		if (!IsValid(sText))
			return 0L;
	}
	
	return Default();
}

BOOL CMaskEdit::IsValid(LPCTSTR szText)
{
	if (!m_sMask.IsEmpty() && szText && *szText)
	{
		// abandon change at the first invalid char
		TCHAR nChar = *szText;
		
		while (nChar)
		{
			if ((m_dwFlags & ME_EXCLUDE) && m_sMask.Find((char)nChar) != -1)
			{
				return FALSE;
			}
			else if (!(m_dwFlags & ME_EXCLUDE) && m_sMask.Find((char)nChar) == -1)
			{
				return FALSE;
			}

			nChar = *(++szText);
		}
	}

	return TRUE;
}
