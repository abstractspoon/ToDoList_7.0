// colorbutton.cpp : implementation file
//

#include "stdafx.h"
#include "colorbutton.h"
#include "encolordialog.h"
#include "graphicsmisc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CColorButton

CColorButton::CColorButton(BOOL bRoundRect) : m_color(0), m_bRoundRect(bRoundRect)
{
}

CColorButton::~CColorButton()
{
}


BEGIN_MESSAGE_MAP(CColorButton, CCustomButton)
	//{{AFX_MSG_MAP(CColorButton)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CColorButton message handlers

void CColorButton::DoExtraPaint(CDC* pDC, const CRect& rExtra)
{
	int nCornerRadius = m_bRoundRect ? (rExtra.Width() / 4) : 0;
	COLORREF crFill = m_color, crBorder = CLR_NONE;

	if (!IsWindowEnabled())
	{
		crFill = GetSysColor(COLOR_3DFACE);
		crBorder = GetSysColor(COLOR_3DDKSHADOW);
	}
	else
		crBorder = GraphicsMisc::Darker(crFill, 0.5);

	GraphicsMisc::DrawRect(pDC, rExtra, crFill, crBorder, nCornerRadius);
}

void CColorButton::SetColor(COLORREF color) 
{ 
	m_color = color; 

	if (GetSafeHwnd())
		Invalidate();
}

BOOL CColorButton::DoAction()
{
	CEnColorDialog dialog(m_color, CC_FULLOPEN | CC_RGBINIT);

	if (dialog.DoModal() == IDOK)
	{
		m_color = dialog.m_cc.rgbResult;
	
		Invalidate();
		UpdateWindow();
	}

	return FALSE; // always notify parent
}

