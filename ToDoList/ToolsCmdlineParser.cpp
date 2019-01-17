// ToolsCmdlineParser.cpp: implementation of the CToolsCmdlineParser class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ToolsCmdlineParser.h"

#include "..\shared\datehelper.h"
#include "..\shared\misc.h"
#include "..\shared\filemisc.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CMap<CString, LPCTSTR, CLA_TYPE, CLA_TYPE&> CToolsCmdlineParser::s_mapTypes;

CToolsCmdlineParser::CToolsCmdlineParser(LPCTSTR szCmdLine)
{
	// init static map first time only
	if (s_mapTypes.GetCount() == 0)
	{
		s_mapTypes[_T("pathname")] = CLAT_PATHNAME;
		s_mapTypes[_T("filetitle")] = CLAT_FILETITLE;
		s_mapTypes[_T("folder")] = CLAT_FOLDER;
		s_mapTypes[_T("filename")] = CLAT_FILENAME;
		s_mapTypes[_T("seltid")] = CLAT_SELTASKID;
		s_mapTypes[_T("selttitle")] = CLAT_SELTASKTITLE;
		s_mapTypes[_T("userfile")] = CLAT_USERFILE;
		s_mapTypes[_T("userfolder")] = CLAT_USERFOLDER;
		s_mapTypes[_T("usertext")] = CLAT_USERTEXT;
		s_mapTypes[_T("userdate")] = CLAT_USERDATE;
		s_mapTypes[_T("todaysdate")] = CLAT_TODAYSDATE;
		s_mapTypes[_T("todolist")] = CLAT_TODOLIST;
		s_mapTypes[_T("seltextid")] = CLAT_SELTASKEXTID;
		s_mapTypes[_T("seltcomments")] = CLAT_SELTASKCOMMENTS;
		s_mapTypes[_T("seltfile")] = CLAT_SELTASKFILELINK;
		s_mapTypes[_T("seltallocby")] = CLAT_SELTASKALLOCBY;
		s_mapTypes[_T("seltallocto")] = CLAT_SELTASKALLOCTO;
		s_mapTypes[_T("seltcustom")] = CLAT_SELTASKCUSTATTRIB;
	}

	SetCmdLine(szCmdLine);
}

CToolsCmdlineParser::~CToolsCmdlineParser()
{

}

void CToolsCmdlineParser::SetCmdLine(LPCTSTR szCmdLine)
{
	m_sCmdLine = szCmdLine;
	m_aArgs.RemoveAll();

	ParseCmdLine();

	// replace 'todaysdate'
	ReplaceArgument(CLAT_TODAYSDATE, CDateHelper::FormatCurrentDate(TRUE));

	// and 'todolist'
	ReplaceArgument(CLAT_TODOLIST, FileMisc::GetAppFilePath());
}

int CToolsCmdlineParser::GetArguments(CCLArgArray& aArgs) const
{
	aArgs.Copy(m_aArgs);
	return aArgs.GetSize();
}

int CToolsCmdlineParser::GetUserArguments(CCLArgArray& aArgs) const
{
	int nArg = m_aArgs.GetSize();

	// add to start to maintain order
	while (nArg--)
	{
		switch (m_aArgs[nArg].nType)
		{
		case CLAT_USERFILE:
		case CLAT_USERFOLDER:
		case CLAT_USERTEXT:
		case CLAT_USERDATE:
			{
				CMDLINEARG cla = m_aArgs[nArg];
				aArgs.InsertAt(0, cla);
			}
			break;
		}
	}

	return aArgs.GetSize();
}

int CToolsCmdlineParser::GetCustomAttributeArguments(CCLArgArray& aArgs) const
{
	int nArg = m_aArgs.GetSize();

	// add to start to maintain order
	while (nArg--)
	{
		switch (m_aArgs[nArg].nType)
		{
		case CLAT_SELTASKCUSTATTRIB:
			{
				CMDLINEARG cla = m_aArgs[nArg];
				cla.sName.MakeUpper();

				aArgs.InsertAt(0, cla);
			}
			break;
		}
	}

	return aArgs.GetSize();
}

BOOL CToolsCmdlineParser::ReplaceArgument(CLA_TYPE nType, LPCTSTR szValue)
{
	// see if we have that type
	int nArg = m_aArgs.GetSize();
	BOOL bFound = FALSE;

	// and replace all of them
	while (nArg--)
	{
		if (m_aArgs[nArg].nType == nType)
			bFound |= ReplaceArgument(nArg, szValue);
	}

	// not found
	return bFound;
}

BOOL CToolsCmdlineParser::ReplaceArgument(CLA_TYPE nType, DWORD dwValue)
{
	CString sValue;
	sValue.Format(_T("%lu"), dwValue);

	return ReplaceArgument(nType, sValue);
}

BOOL CToolsCmdlineParser::ReplaceArgument(LPCTSTR szName, LPCTSTR szValue)
{
	// see if we have a user item with that name
	CString sName(szName);
	sName.MakeLower();

	int nArg = m_aArgs.GetSize();

	while (nArg--)
	{
		switch (m_aArgs[nArg].nType)
		{
		case CLAT_USERFILE:
		case CLAT_USERFOLDER:
		case CLAT_USERTEXT:
		case CLAT_USERDATE:
			if (sName.CompareNoCase(m_aArgs[nArg].sName) == 0)
				return ReplaceArgument(nArg, szValue);
			break;

		case CLAT_SELTASKCUSTATTRIB:
			if (sName.CompareNoCase(m_aArgs[nArg].sPlaceHolder) == 0)
				return ReplaceArgument(nArg, szValue);
			break;
		}
	}

	// not found
	return FALSE;

}

BOOL CToolsCmdlineParser::IsUserInputRequired() const
{
	// see if we have any 'USER' types
	int nArg = m_aArgs.GetSize();

	while (nArg--)
	{
		switch (m_aArgs[nArg].nType)
		{
		case CLAT_USERFILE:
		case CLAT_USERFOLDER:
		case CLAT_USERTEXT:
		case CLAT_USERDATE:
			return TRUE;
		}
	}

	// not found
	return FALSE;
}

BOOL CToolsCmdlineParser::HasTasklistArgument() const
{
	// see if we have any tasklist related types
	int nArg = m_aArgs.GetSize();

	while (nArg--)
	{
		switch (m_aArgs[nArg].nType)
		{
		case CLAT_PATHNAME:
		case CLAT_FILETITLE:   
		case CLAT_FOLDER:        
		case CLAT_FILENAME:
			return TRUE;
		}
	}

	// not found
	return FALSE;
}

BOOL CToolsCmdlineParser::IsUserInputType(CLA_TYPE nType) const
{
	switch (nType)
	{
	case CLAT_USERFILE:
	case CLAT_USERFOLDER:
	case CLAT_USERTEXT:
	case CLAT_USERDATE:
		return TRUE;
	}
	
	return FALSE;
}

BOOL CToolsCmdlineParser::ReplaceArgument(int nArg, LPCTSTR szValue)
{
	if (nArg < 0 || nArg >= m_aArgs.GetSize())
		return FALSE;

	CMDLINEARG& cla = m_aArgs[nArg];

	if (m_sCmdLine.Replace(cla.sPlaceHolder, szValue))
	{
		// also check if there are any user references to this variable name
		// and replace those too
		if (IsUserInputType(cla.nType))
		{
			int nUserArg = m_aUserArgs.GetSize();

			while (nUserArg--)
			{
				CMDLINEARG& claUser = m_aUserArgs[nUserArg];

				if (claUser.sName.CompareNoCase(cla.sName) == 0)
				{
					m_sCmdLine.Replace(claUser.sPlaceHolder, szValue);

					// and remove
					m_aUserArgs.RemoveAt(nUserArg);
				}
			}
		}

		return TRUE;
	}

	// else
	return FALSE;
}

BOOL CToolsCmdlineParser::HasArgument(CLA_TYPE nType) const
{
	int nArg = m_aArgs.GetSize();

	while (nArg--)
	{
		if (m_aArgs[nArg].nType == nType)
			return TRUE;
	}

	return FALSE;
}

void CToolsCmdlineParser::ParseCmdLine()
{
	CString sCmdLine(m_sCmdLine); // preserve original

	int nDollar = sCmdLine.Find('$');

	while (-1 != nDollar)
	{
		// find opening bracket
		int nOpenFind = sCmdLine.Find('(', nDollar);

		if (nOpenFind == -1)
			break;

		// find closing bracket
		int nCloseFind = sCmdLine.Find(')', nOpenFind);

		if (nCloseFind == -1)
			break;

		// parse variable in the form of (vartype, varname, varlabel, vardefvalue)
		CString sVarArgs = sCmdLine.Mid(nOpenFind + 1, nCloseFind - nOpenFind - 1);
		CString sType, sName, sLabel, sDefValue;

		int nComma1Find = sVarArgs.Find(','), nComma2Find = -1;

		if (nComma1Find != -1)
		{
			sType = sVarArgs.Left(nComma1Find);

			nComma2Find = sVarArgs.Find(',', nComma1Find + 1);

			if (nComma2Find != -1)
			{
				sName = sVarArgs.Mid(nComma1Find + 1, nComma2Find - nComma1Find - 1);

				int nComma3Find = sVarArgs.Find(',', nComma2Find + 1);

				if (nComma3Find != -1)
				{
					// this comma can either be a comma in the label string
					// or the delimeter before 'vardefvalue'
					// we determine which of these it is by looking for double quotes
					int nQuoteStartFind = sVarArgs.Find('\"', nComma2Find + 1);

					// and seeing if they preceed nComma3Find
					if (nQuoteStartFind != -1 && nQuoteStartFind < nComma3Find)
					{
						// now look for closing quotes
						int nQuoteEndFind = sVarArgs.Find('\"', nQuoteStartFind + 1);

						if (nQuoteEndFind != -1) // test for nComma3Find again cos it was previously a false find
							nComma3Find = sVarArgs.Find(',', nQuoteEndFind + 1);
						else
							nComma3Find = -1; // safest thing to do because no end quotes found
					}
						
					if (nComma3Find != -1)
					{
						sLabel = sVarArgs.Mid(nComma2Find + 1, nComma3Find - nComma2Find - 1);
						sDefValue = sVarArgs.Mid(nComma3Find + 1);

						sDefValue.Replace(_T("\""), _T("")); // remove double quotes
					}
					else
					{
						sLabel = sVarArgs.Mid(nComma2Find + 1);
					}
				}
				else
				{
					sLabel = sVarArgs.Mid(nComma2Find + 1);
				}

				sLabel.Replace(_T("\""), _T("")); // remove double quotes
			}
			else
			{
				sName = sVarArgs.Mid(nComma1Find + 1);
			}
		}
		else
		{
			sType = sVarArgs;
		}

		// process
		Misc::Trim(sType);
		Misc::Trim(sLabel);
		Misc::Trim(sDefValue);
	
		Misc::Trim(sName);
		sName.MakeLower();

		CMDLINEARG cla;

		cla.nType = GetType(sType);
		cla.sPlaceHolder = sCmdLine.Mid(nDollar, nCloseFind - nDollar + 1);

		switch (cla.nType)
		{
		// user input types must have a valid name that is not a 'keyword'
		case CLAT_USERFILE:
		case CLAT_USERFOLDER:
		case CLAT_USERTEXT:
		case CLAT_USERDATE:
		case CLAT_SELTASKCUSTATTRIB:      
			ASSERT(!sName.IsEmpty() && GetType(sName) == CLAT_NONE);
			
			if (!sName.IsEmpty() && GetType(sName) == CLAT_NONE)
			{
				cla.sName = sName;
				cla.sLabel = sLabel;
				cla.sDefValue = sDefValue;

				m_aArgs.Add(cla);
			}
			break;

		case CLAT_TODOLIST:      
		case CLAT_PATHNAME:      
		case CLAT_FILETITLE:    
		case CLAT_FOLDER:        
		case CLAT_FILENAME:      
		case CLAT_TODAYSDATE:
		case CLAT_SELTASKID:
		case CLAT_SELTASKTITLE:
		case CLAT_SELTASKEXTID:
		case CLAT_SELTASKCOMMENTS:
		case CLAT_SELTASKFILELINK:
		case CLAT_SELTASKALLOCBY:
		case CLAT_SELTASKALLOCTO:      
			m_aArgs.Add(cla);
			break;

		default: // assume it is USER attribute
			{
				ASSERT (cla.nType == CLAT_NONE); // what else?
	
				cla.sName = sType;
				m_aUserArgs.Add(cla);
			}
			break;
		}

		// next arg
		nDollar = sCmdLine.Find('$', nCloseFind);
	}
}

CLA_TYPE CToolsCmdlineParser::GetType(LPCTSTR szVarType) const
{
	CString sType(szVarType);

	Misc::Trim(sType);
	sType.MakeLower();

	CLA_TYPE nType = CLAT_NONE;
	s_mapTypes.Lookup(sType, nType);

	return nType;
}

BOOL CToolsCmdlineParser::PrepareToolPath(CString& sToolPath)
{
	return sToolPath.Replace(_T("$(todolist)"), FileMisc::GetAppFilePath());
}
