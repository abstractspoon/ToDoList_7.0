// TDLCsvImportExportDlg.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "TDLCsvImportExportDlg.h"
#include "tdcstatic.h"

#include "..\shared\misc.h"
#include "..\shared\filemisc.h"
#include "..\shared\fileregister.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CTDLCsvImportExportDlg dialog


CTDLCsvImportExportDlg::CTDLCsvImportExportDlg(const CString& sFilePath, 
											   IPreferences* pPrefs, LPCTSTR szKey, CWnd* pParent /*=NULL*/)
	: CDialog(IDD_CSVIMPORTEXPORT_DIALOG, pParent), 
	m_lcColumnSetup(TRUE), 
	m_eFilePath(FES_NOBROWSE)
{
	VERIFY(DoInit(sFilePath, pPrefs, szKey, NULL));
}

CTDLCsvImportExportDlg::CTDLCsvImportExportDlg(const CString& sFilePath, 
											   const CTDCAttributeArray& aExportAttributes, 
											   IPreferences* pPrefs, LPCTSTR szKey, CWnd* pParent /*=NULL*/)
	: CDialog(IDD_CSVIMPORTEXPORT_DIALOG, pParent), 
	m_lcColumnSetup(FALSE), 
	m_eFilePath(FES_NOBROWSE)
{
	VERIFY(DoInit(sFilePath, pPrefs, szKey, &aExportAttributes));
}

BOOL CTDLCsvImportExportDlg::DoInit(const CString& sFilePath, 
									IPreferences* pPrefs, 
									LPCTSTR szKey, 
									const CTDCAttributeArray* pExportAttributes)
{
	m_sFilePath = sFilePath; 
	m_pPrefs = pPrefs;
	m_sPrefKey.Format(_T("%s\\CsvColumnMapping"), szKey);
	m_sDelim = Misc::GetListSeparator();
	m_bAlwaysExportTaskIDs = TRUE;

	LoadMasterColumnMapping();

	// user mapping
	CTDCCsvColumnMapping aMapping;

	if (pExportAttributes)
	{
		m_bImporting = FALSE;
		m_aExportAttributes.Copy(*pExportAttributes);
		BuildExportColumnMapping(aMapping);
	}
	else
	{
		m_bImporting = TRUE;
		BuildImportColumnMapping(aMapping);
	}
	m_lcColumnSetup.SetColumnMapping(aMapping);

	return TRUE;
}

void CTDLCsvImportExportDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTDLCsvImportExportDlg)
	DDX_Text(pDX, IDC_CSVDELIMITER, m_sDelim);
	DDX_Text(pDX, IDC_CSVFILEPATH, m_sFilePath); 
	DDX_Control(pDX, IDC_CSVFILEPATH, m_eFilePath);
	DDX_Control(pDX, IDC_COLUMNSETUP, m_lcColumnSetup);
	DDX_Check(pDX, IDC_EXPORTTASKIDS, m_bAlwaysExportTaskIDs);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CTDLCsvImportExportDlg, CDialog)
	//{{AFX_MSG_MAP(CTDLCsvImportExportDlg)
	ON_EN_CHANGE(IDC_CSVDELIMITER, OnChangeCsvdelimiter)
	ON_BN_CLICKED(IDC_EXPORTTASKIDS, OnExportTaskids)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTDLCsvImportExportDlg message handlers

BOOL CTDLCsvImportExportDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();

	SetWindowText(CEnString(m_bImporting ? IDS_CSV_IMPORTDLG : IDS_CSV_EXPORTDLG));

	GetDlgItem(IDC_EXPORTTASKIDS)->EnableWindow(!m_bImporting);
	GetDlgItem(IDC_EXPORTTASKIDS)->ShowWindow(m_bImporting ? SW_HIDE : SW_SHOW);

	if (!m_bImporting)
		OnExportTaskids();
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}


int CTDLCsvImportExportDlg::GetColumnMapping(CTDCCsvColumnMapping& aMapping) const 
{ 
	return m_lcColumnSetup.GetColumnMapping(aMapping);
}

CString CTDLCsvImportExportDlg::GetDelimiter() const 
{ 
	// some special cases
	if (m_sDelim == _T("\\t"))
	{
		return _T("\t");
	}
	else if (m_sDelim == _T("\\n"))
	{
		return _T("\n");
	}
	else if (m_sDelim == _T("\\r\\n"))
	{
		return _T("\r\n");
	}

	// else
	return m_sDelim; 
}

void CTDLCsvImportExportDlg::OnChangeCsvdelimiter() 
{
	CString sOldDelim = m_sDelim;
	UpdateData();

	if (!m_sDelim.IsEmpty() && m_bImporting && m_sDelim != sOldDelim)
	{
		CTDCCsvColumnMapping aMapping;
		
		if (BuildImportColumnMapping(aMapping))
			m_lcColumnSetup.SetColumnMapping(aMapping);
	}
}

void CTDLCsvImportExportDlg::BuildDefaultMasterColumnMapping()
{
	m_aMasterColumnMapping.RemoveAll();

	for (int nCol = 0; nCol < ATTRIB_COUNT; nCol++)
	{
		TDC_ATTRIBUTE attrib = ATTRIBUTES[nCol].attrib;
		CEnString sName(ATTRIBUTES[nCol].nAttribResID);

		m_aMasterColumnMapping.Add(CSVCOLUMNMAPPING(sName, attrib));
	}
}

int CTDLCsvImportExportDlg::BuildImportColumnMapping(CTDCCsvColumnMapping& aImportMapping) const
{
	ASSERT (m_bImporting);
	ASSERT(!m_sFilePath.IsEmpty());

	// read first few lines from file
	CStringArray aLines;
	
	if (!FileMisc::LoadFileLines(m_sFilePath, aLines, 5))
		return FALSE;
	
	// look for the first line containing something
	// we can use as column heading names
	CStringArray aColumnHeaders;

	for (int nLine = 0; nLine < 5; nLine++)
	{
		if (Misc::Split(aLines[nLine], aColumnHeaders, GetDelimiter(), TRUE))
			break;
	}

	if (aColumnHeaders.GetSize() == 0)
		return FALSE;
		
	// build column mapping from file attributes
	for (int nCol = 0; nCol < aColumnHeaders.GetSize(); nCol++)
	{
		CString sName = Misc::GetItem(aColumnHeaders, nCol);

		if (sName.IsEmpty())
			sName = CEnString(IDS_UNNAMEDFILTER);

		// try to map text column names to column IDs
		// Note: Must include _all_ columns else the column
		// indices will be out of sync with the data rows
		aImportMapping.Add(CSVCOLUMNMAPPING(sName, GetMasterColumnAttribute(sName)));
	}

	return aImportMapping.GetSize();
}

int CTDLCsvImportExportDlg::BuildExportColumnMapping(CTDCCsvColumnMapping& aExportMapping) const
{
	ASSERT (!m_bImporting);

	// build column mapping from passed in attributes
	for (int nAttrib = 0; nAttrib < m_aExportAttributes.GetSize(); nAttrib++)
	{
		TDC_ATTRIBUTE attrib = m_aExportAttributes[nAttrib];
		ASSERT(attrib != TDCA_NONE);

		// try to map text column names to column IDs
		aExportMapping.Add(CSVCOLUMNMAPPING(GetMasterColumnName(attrib), attrib));
	}

	return aExportMapping.GetSize();
}

void CTDLCsvImportExportDlg::OnOK()
{
	CDialog::OnOK();
	
	// save attribute mapping
	UpdateMasterColumnMappingFromList();
	SaveMasterColumnMapping();
}

int CTDLCsvImportExportDlg::LoadMasterColumnMapping()
{
	BuildDefaultMasterColumnMapping();

	m_bAlwaysExportTaskIDs = m_pPrefs->GetProfileInt(m_sPrefKey, _T("AlwaysExportTaskIDs"), TRUE);
	int nColumns = m_pPrefs->GetProfileInt(m_sPrefKey, _T("ColumnCount"), 0);

	// overwrite with translations unless they are empty names
	for (int nCol = 0; nCol < nColumns; nCol++)
	{
		CString sKey = Misc::MakeKey(_T("ColumnAttrib%d"), nCol);
		TDC_ATTRIBUTE attrib = (TDC_ATTRIBUTE)m_pPrefs->GetProfileInt(m_sPrefKey, sKey, TDCA_NONE);
		
		sKey = Misc::MakeKey(_T("ColumnName%d"), nCol);
		CString sName = m_pPrefs->GetProfileString(m_sPrefKey, sKey);
		
		if (!sName.IsEmpty())
			SetMasterColumnName(attrib, sName);
	}

	// load last used delimiter
	CString sDelim = m_pPrefs->GetProfileString(m_sPrefKey, _T("Delimiter"));

	// Now we export as UTF8, we enforce comma if Excel is the default app for csv
	if (CFileRegister::IsRegisteredApp(_T("csv"), _T("EXCEL.EXE"), TRUE))
	{
		sDelim = _T(",");
	}

	if (!sDelim.IsEmpty())
	{
		m_sDelim = sDelim;

	    if (GetSafeHwnd())
			UpdateData(FALSE);
	}

	return m_aMasterColumnMapping.GetSize();
}

void CTDLCsvImportExportDlg::UpdateMasterColumnMappingFromList()
{
	// get mapping from list ctrl and update names in master mapping
	CTDCCsvColumnMapping aListMapping;
	int nListRows = m_lcColumnSetup.GetColumnMapping(aListMapping);

	for (int nRow = 0; nRow < nListRows; nRow++)
	{
		SetMasterColumnName(aListMapping[nRow].nTDCAttrib, aListMapping[nRow].sColumnName);
	}
}

void CTDLCsvImportExportDlg::SaveMasterColumnMapping() const
{
	m_pPrefs->WriteProfileInt(m_sPrefKey, _T("AlwaysExportTaskIDs"), m_bAlwaysExportTaskIDs);

	int nColumns = m_aMasterColumnMapping.GetSize();
	m_pPrefs->WriteProfileInt(m_sPrefKey, _T("ColumnCount"), nColumns);

	for (int nCol = 0; nCol < nColumns; nCol++)
	{
		CString sKey = Misc::MakeKey(_T("ColumnName%d"), nCol);
		m_pPrefs->WriteProfileString(m_sPrefKey, sKey, m_aMasterColumnMapping[nCol].sColumnName);
		
		sKey = Misc::MakeKey(_T("ColumnAttrib%d"), nCol);
		m_pPrefs->WriteProfileInt(m_sPrefKey, sKey, m_aMasterColumnMapping[nCol].nTDCAttrib);
	}

	// save delimiter if different to default
	if (m_sDelim == Misc::GetListSeparator())
		m_pPrefs->WriteProfileString(m_sPrefKey, _T("Delimiter"), _T(""));
	else
		m_pPrefs->WriteProfileString(m_sPrefKey, _T("Delimiter"), m_sDelim);
}

CString CTDLCsvImportExportDlg::GetMasterColumnName(TDC_ATTRIBUTE attrib) const
{
	int nCol = FindMasterColumn(attrib);
	return (nCol == -1) ? _T("") : m_aMasterColumnMapping[nCol].sColumnName;
}

TDC_ATTRIBUTE CTDLCsvImportExportDlg::GetMasterColumnAttribute(LPCTSTR szColumn) const
{
	int nCol = FindMasterColumn(szColumn);
	return (nCol == -1) ? TDCA_NONE : m_aMasterColumnMapping[nCol].nTDCAttrib;
}

void CTDLCsvImportExportDlg::SetMasterColumnAttribute(LPCTSTR szColumn, TDC_ATTRIBUTE attrib)
{
	// check if attribute is already in use
	int nAttribCol = FindMasterColumn(attrib);
	int nNameCol = FindMasterColumn(szColumn);

	// and clear if it is
	if (nAttribCol != -1 && nAttribCol != nNameCol)
		m_aMasterColumnMapping[nNameCol].nTDCAttrib = TDCA_NONE;

	if (nNameCol != -1)
		m_aMasterColumnMapping[nNameCol].nTDCAttrib = attrib;
}

void CTDLCsvImportExportDlg::SetMasterColumnName(TDC_ATTRIBUTE attrib, LPCTSTR szColumn)
{
	// prevent setting the master mapping to an empty name
	if (szColumn && *szColumn)
	{
		// check if the column name is already in use
		int nNameCol = FindMasterColumn(szColumn);
		int nAttribCol = FindMasterColumn(attrib);

		// and clear if it is
		if (nNameCol != -1 && nNameCol != nAttribCol)
			m_aMasterColumnMapping[nNameCol].sColumnName.Empty();

		// and set new name
		if (nAttribCol != -1)
			m_aMasterColumnMapping[nAttribCol].sColumnName = szColumn;
	}
}

int CTDLCsvImportExportDlg::FindMasterColumn(TDC_ATTRIBUTE attrib) const
{
	int nColumns = m_aMasterColumnMapping.GetSize();

	for (int nCol = 0; nCol < nColumns; nCol++)
	{
		if (m_aMasterColumnMapping[nCol].nTDCAttrib == attrib)
			return nCol;
	}

	// else
	return -1;
}

int CTDLCsvImportExportDlg::FindMasterColumn(LPCTSTR szColumn) const
{
	int nColumns = m_aMasterColumnMapping.GetSize();

	for (int nCol = 0; nCol < nColumns; nCol++)
	{
		if (m_aMasterColumnMapping[nCol].sColumnName.CompareNoCase(szColumn) == 0)
			return nCol;
	}

	// else
	return -1;
}

void CTDLCsvImportExportDlg::OnExportTaskids() 
{
	ASSERT(!m_bImporting);

	UpdateData();

	CTDCCsvColumnMapping aMapping;
	m_lcColumnSetup.GetColumnMapping(aMapping);

	if (!m_bAlwaysExportTaskIDs)
	{
		// Find if these attributes were originally present
		BOOL bWantTaskID = FALSE, bWantParentID = FALSE;

		for (int nAttrib = 0; nAttrib < m_aExportAttributes.GetSize(); nAttrib++)
		{
			if (!bWantTaskID)
				bWantTaskID = (m_aExportAttributes[nAttrib] == TDCA_ID);

			if (!bWantParentID)
				bWantParentID = (m_aExportAttributes[nAttrib] == TDCA_PARENTID);
		}

		// if attribute was not present in original attributes then remove
		int nCol = aMapping.GetSize();

		while (nCol--)
		{
			TDC_ATTRIBUTE attrib = aMapping[nCol].nTDCAttrib;

			if (attrib == TDCA_ID && !bWantTaskID)
				aMapping.RemoveAt(nCol);

			else if (attrib == TDCA_PARENTID && !bWantParentID)
				aMapping.RemoveAt(nCol);
		}
	}
	else // always include
	{
		// find out what's already present
		int nTaskID = -1, nParentTaskID = -1;

		for (int nCol = 0; nCol < aMapping.GetSize() && (nTaskID == -1 || nParentTaskID == -1); nCol++)
		{
			TDC_ATTRIBUTE attrib = aMapping[nCol].nTDCAttrib;

			if (attrib == TDCA_ID)
				nTaskID = nCol;

			else if (attrib == TDCA_PARENTID)
				nParentTaskID = nCol;
		}

		// Add TaskID and/or ParentTaskID if not present
		if (nTaskID == -1)
			aMapping.Add(CSVCOLUMNMAPPING(GetMasterColumnName(TDCA_ID), TDCA_ID)); 

		if (nParentTaskID == -1)
			aMapping.Add(CSVCOLUMNMAPPING(GetMasterColumnName(TDCA_PARENTID), TDCA_PARENTID));
	}
	
	m_lcColumnSetup.SetColumnMapping(aMapping);
}
