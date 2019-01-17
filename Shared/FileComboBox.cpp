// FileComboBox.cpp : implementation file
//

#include "stdafx.h"
#include "FileComboBox.h"
#include "FileMisc.h"
#include "enbitmap.h"

// CFileComboBox

CFileComboBox::CFileComboBox(int nEditStyle) 
	: 
	CAutoComboBox(ACBS_ALLOWDELETE | ACBS_ADDTOSTART),
	m_fileEdit(nEditStyle),
	m_imageIcons(16, 16)
{

}

CFileComboBox::~CFileComboBox()
{
}


BEGIN_MESSAGE_MAP(CFileComboBox, CAutoComboBox)
	ON_WM_SIZE()
	ON_REGISTERED_MESSAGE(WM_FEN_BROWSECHANGE, OnFileEditBrowseChange)
	ON_REGISTERED_MESSAGE(WM_FE_GETFILEICON, OnFileEditGetFileIcon)
	ON_CONTROL_REFLECT_EX(CBN_SELCHANGE, OnSelChange)
END_MESSAGE_MAP()



// CFileComboBox message handlers
void CFileComboBox::PreSubclassWindow()
{
	CAutoComboBox::PreSubclassWindow();

	VERIFY (InitFileEdit());
}

BOOL CFileComboBox::PreCreateWindow(CREATESTRUCT& cs)
{
	// only support combos with edit boxes
	if (cs.style & CBS_DROPDOWNLIST)
		return FALSE;

	if (!InitFileEdit())
		return FALSE;

	cs.style |= (CBS_OWNERDRAWFIXED | CBS_HASSTRINGS);

	return CAutoComboBox::PreCreateWindow(cs);
}

void CFileComboBox::OnSize(UINT nType, int cx, int cy)
{
	CAutoComboBox::OnSize(nType, cx, cy);

	ResizeEdit();
}

void CFileComboBox::ResizeEdit()
{
	// resize the edit to better fit the combo
	if (m_fileEdit.GetSafeHwnd())
	{
		CRect rCombo;
		GetClientRect(rCombo);

		CRect rEdit(rCombo);
		rEdit.DeflateRect(2, 3, 1, 2);
		rEdit.right -= GetSystemMetrics(SM_CXVSCROLL);
		m_fileEdit.MoveWindow(rEdit, FALSE);
	}
}

BOOL CFileComboBox::InitFileEdit()
{
	ASSERT (!m_fileEdit.GetSafeHwnd());

	if (!m_fileEdit.SubclassDlgItem(1001, this))
		return FALSE;

	ResizeEdit();
	return TRUE;
}

LRESULT CFileComboBox::OnFileEditBrowseChange(WPARAM wp, LPARAM lp)
{
	ASSERT(wp == 1001);
	ASSERT(lp);

	CString sPath((LPCTSTR)lp);

	if (!sPath.IsEmpty())
	{
		HandleReturnKey();
		ResizeEdit();
	}

	return 0L;
}

BOOL CFileComboBox::OnSelChange()
{
	// this constitutes an edit
	m_bEditChange = TRUE;
	
	return CAutoComboBox::OnSelChange();
}

LRESULT CFileComboBox::OnFileEditGetFileIcon(WPARAM wp, LPARAM lp)
{
	ASSERT(wp == 1001);
	ASSERT(lp);

	// pass to parent
	return GetParent()->SendMessage(WM_FE_GETFILEICON, GetDlgCtrlID(), lp);
}

int CFileComboBox::GetFileList(CStringArray& aFiles) 
{ 
	return GetItems(aFiles); 
}

int CFileComboBox::SetFileList(const CStringArray& aFiles) 
{ 
	int nNumFiles = SetStrings(aFiles);
	
	if (nNumFiles)
		SetCurSel(0);
	else
		m_fileEdit.Invalidate();
	
	return nNumFiles;
}

int CFileComboBox::AddFiles(const CStringArray& aFiles)
{
	// add files in reverse order so that the first one ends up at the top
	int nFile = aFiles.GetSize();
	int nNumAdded = 0;

	while (nFile--)
	{
		if (AddFile(Misc::GetItem(aFiles, nFile)) != -1)
			nNumAdded++;
	}

	return nNumAdded;
}

void CFileComboBox::DrawItemText(CDC& dc, const CRect& rect, int nItem, UINT nItemState,
								DWORD dwItemData, const CString& sItem, BOOL bList)
{
	CRect rText(rect);

	if (bList && !sItem.IsEmpty())
	{
		BOOL bDrawn = FALSE;

		if (m_fileEdit.HasStyle(FES_DISPLAYSIMAGES) && CEnBitmap::IsSupportedImageFile(sItem))
		{
			CString sFullPath(sItem);
			FileMisc::MakeFullPath(sFullPath, m_fileEdit.GetCurrentFolder());
			
			if (m_imageIcons.HasIcon(sFullPath) || m_imageIcons.Add(sFullPath, sFullPath))
			{
				bDrawn = m_imageIcons.Draw(&dc, sFullPath, rect.TopLeft());
			}
		}

		// Fallback/default
		if (!bDrawn)
			m_fileEdit.DrawFileIcon(&dc, sItem, rText);

		rText.left += 20;
	}

	CAutoComboBox::DrawItemText(dc, rText, nItem, nItemState, dwItemData, sItem, bList);
}

BOOL CFileComboBox::DeleteLBItem(int nItem)
{
	// Cache item text for later image removal
	CString sFilename;
	GetLBText(nItem, sFilename);

	if (CAutoComboBox::DeleteLBItem(nItem))
	{
		m_bEditChange = TRUE;

		// update the edit box with the current list selection
		int nLBSel = ::SendMessage(GetListbox(), LB_GETCURSEL, 0, 0);

		if (nLBSel != -1)
		{
			SetCurSel(nLBSel);
		}
		else if (GetCount())
		{
			if (nItem < GetCount())
				SetCurSel(nItem);
			else
				SetCurSel(nItem - 1);
		}

		// delete any image associated with this item
		if (m_imageIcons.HasIcon(sFilename))
			m_imageIcons.Remove(sFilename);

		return TRUE;
	}

	// else
	return FALSE;
}

LRESULT CFileComboBox::OnEditboxMessage(UINT msg, WPARAM wp, LPARAM lp)
{
	// handle 'escape' to cancel deletions
	if ((msg == WM_CHAR) && (wp == VK_ESCAPE))
	{
		m_bEditChange = FALSE;
		ShowDropDown(FALSE);

		return 0L;
	}
	
	// default handling
	return CAutoComboBox::OnEditboxMessage(msg, wp, lp);
}
