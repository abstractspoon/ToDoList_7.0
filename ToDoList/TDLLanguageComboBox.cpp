// TDLLanguageComboBox.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "TDLLanguageComboBox.h"

#include "..\Shared\filemisc.h"
#include "..\Shared\enbitmap.h"
#include "..\Shared\localizer.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////

const CString DEFLANG(_T("English (UK)"));

CString CTDLLanguageComboBox::GetDefaultLanguage()
{
	return DEFLANG;
}

/////////////////////////////////////////////////////////////////////////////
// CTDLLanguageComboBox

CTDLLanguageComboBox::CTDLLanguageComboBox(LPCTSTR szFilter) : m_sSelLanguage(DEFLANG), m_sFilter(szFilter)
{
}

CTDLLanguageComboBox::~CTDLLanguageComboBox()
{
}


BEGIN_MESSAGE_MAP(CTDLLanguageComboBox, CComboBoxEx)
	//{{AFX_MSG_MAP(CTDLLanguageComboBox)
	ON_WM_CREATE()
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTDLLanguageComboBox message handlers

void CTDLLanguageComboBox::PreSubclassWindow() 
{
	CComboBoxEx::PreSubclassWindow();

	CLocalizer::EnableTranslation(*this, FALSE);
	BuildLanguageList();
}

int CTDLLanguageComboBox::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CComboBoxEx::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	CLocalizer::EnableTranslation(*this, FALSE);
	BuildLanguageList();
	
	return 0;
}

CString CTDLLanguageComboBox::GetTranslationFolder()
{
	return FileMisc::GetAppResourceFolder(_T("Resources\\Translations"));
}

void CTDLLanguageComboBox::BuildLanguageList()
{
	if (GetCount())
		return; // already done

	// build the language list from csv files in the Resources\Translations folder
	CString sFolder = GetTranslationFolder();
	CStringArray aFiles;

	int nNumFiles = FileMisc::FindFiles(sFolder, aFiles, FALSE, m_sFilter);
	
	// add english as a default item
	HBITMAP hbmFlag = CEnBitmap::LoadImageResource(IDR_GB_FLAG, _T("GIF"));
	AddString(DEFLANG, hbmFlag, 1);

	// add rest of available languages
	for (int nFile = 0; nFile < nNumFiles; nFile++)
	{
		CString sFileName;
		FileMisc::SplitPath(aFiles[nFile], NULL, NULL, &sFileName, NULL);

		// load icon file
		CString sIconPath(aFiles[nFile]);
		FileMisc::ReplaceExtension(sIconPath, _T("png"));

		hbmFlag = CEnBitmap::LoadImageFile(sIconPath);
		AddString(sFileName, hbmFlag, 0);
	}

	SelectLanguage(m_sSelLanguage);
}

int CTDLLanguageComboBox::AddString(LPCTSTR szLanguage, HBITMAP hbmFlag, DWORD dwItemData)
{
	// create and associate the image list first time around
	if (m_il.GetSafeHandle() == NULL)
	{
		m_il.Create(16, 16, ILC_COLOR32 | ILC_MASK, 1, 1);
		SetImageList(&m_il);
	}

	COMBOBOXEXITEM cbe;

	cbe.mask = CBEIF_IMAGE | CBEIF_SELECTEDIMAGE | CBEIF_TEXT | CBEIF_LPARAM;
	cbe.iItem = GetCount();
	cbe.pszText = (LPTSTR)szLanguage;
	cbe.lParam = dwItemData;
	cbe.iImage = cbe.iSelectedImage = GetCount();

	if (hbmFlag == NULL)
		hbmFlag = CEnBitmap::LoadImageResource(IDR_YOURLANG_FLAG, _T("GIF"));

	CBitmap tmp;
	tmp.Attach(hbmFlag); // will auto cleanup
	m_il.Add(&tmp, CLR_NONE);

	return InsertItem(&cbe);
}

void CTDLLanguageComboBox::SetLanguageFile(LPCTSTR szFile)
{
	if (GetDefaultLanguage() == szFile)
	{
		if (GetSafeHwnd())
			SelectLanguage(szFile);
		else
			m_sSelLanguage = szFile;
	}
	else
	{
		CString sPath = FileMisc::GetFullPath(szFile, FileMisc::GetAppFolder());
		
		if (FileMisc::FileExists(sPath))
		{
			FileMisc::SplitPath(szFile, NULL, NULL, &m_sSelLanguage, NULL);

			if (GetSafeHwnd())
				SelectLanguage(m_sSelLanguage);
		}
	}
}

void CTDLLanguageComboBox::OnDestroy() 
{
	int nSel = GetCurSel();

	if (nSel == -1)
		m_sSelLanguage = DEFLANG;
	else
		GetLBText(nSel, m_sSelLanguage);

	CComboBoxEx::OnDestroy();
}

CString CTDLLanguageComboBox::GetLanguageFile(LPCTSTR szLanguage, LPCTSTR szExt, BOOL bRelative) // static
{
	ASSERT(szLanguage && *szLanguage);

	CString sLanguageFile(GetDefaultLanguage());

	if (GetDefaultLanguage() != szLanguage)
	{
		if (bRelative)
			FileMisc::MakePath(sLanguageFile, NULL, _T(".\\Resources\\Translations"), szLanguage, szExt);
		else
			FileMisc::MakePath(sLanguageFile, NULL, GetTranslationFolder(), szLanguage, szExt);
	}

	return sLanguageFile;
}

CString CTDLLanguageComboBox::GetLanguageFile(BOOL bRelative) const
{
	CString sLanguageFile(GetDefaultLanguage());

	if (GetSafeHwnd())
	{
		int nSel = GetCurSel();

		if (nSel == -1)
			m_sSelLanguage = DEFLANG;
		else
			GetLBText(nSel, m_sSelLanguage);
	}

	LPCTSTR szExt = (m_sFilter.Find(_T(".csv")) != -1) ? _T("csv") : _T("txt");
	return GetLanguageFile(m_sSelLanguage, szExt, bRelative);
}

BOOL CTDLLanguageComboBox::IsDefaultLanguage() const
{
	return (GetLanguageFile() == DEFLANG);
}

BOOL CTDLLanguageComboBox::HasYourLanguage() const
{
	return (FindStringExact(0, _T("YourLanguage")) != -1);
}

int CTDLLanguageComboBox::SelectLanguage(LPCTSTR szLanguage)
{
	int nIndex = GetCount();

	while (nIndex--)
	{
		CString sLanguage;
		GetLBText(nIndex, sLanguage);

		if (sLanguage.CompareNoCase(szLanguage) == 0)
		{
			SetCurSel(nIndex);
			return nIndex;
		}
	}

	if (GetCurSel() == -1)
		SetCurSel(0);

	return 0;
}

BOOL CTDLLanguageComboBox::HasLanguages()
{
	return FileMisc::FolderContainsFiles(GetTranslationFolder(), FALSE, _T("*.csv"));
}
