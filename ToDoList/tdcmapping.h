#if !defined(AFX_TDCMAP_H__5951FDE6_508A_4A9D_A55D_D16EB026AEF7__INCLUDED_)
#define AFX_TDCMAP_H__5951FDE6_508A_4A9D_A55D_D16EB026AEF7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// tdcmapping.h : header file

/////////////////////////////////////////////////////////////////////////////

#include "resource.h"
#include "tdcenum.h"

/////////////////////////////////////////////////////////////////////////////

namespace TDC
{
	static TDC_COLUMN MapSortIDToColumn(UINT nSortID) 
	{
		switch (nSortID)
		{
		case ID_SORT_BYALLOCBY:			return TDCC_ALLOCBY;
		case ID_SORT_BYALLOCTO:			return TDCC_ALLOCTO;
		case ID_SORT_BYCATEGORY:		return TDCC_CATEGORY;
		case ID_SORT_BYCOLOR:			return TDCC_COLOR;
		case ID_SORT_BYCOST:			return TDCC_COST;
		case ID_SORT_BYCREATEDBY:		return TDCC_CREATEDBY;
		case ID_SORT_BYCREATIONDATE:	return TDCC_CREATIONDATE;
		case ID_SORT_BYDEPENDENCY:		return TDCC_DEPENDENCY;
		case ID_SORT_BYDONE:			return TDCC_DONE;
		case ID_SORT_BYDONEDATE:		return TDCC_DONEDATE; 
		case ID_SORT_BYDUEDATE:			return TDCC_DUEDATE;
		case ID_SORT_BYEXTERNALID:		return TDCC_EXTERNALID;
		case ID_SORT_BYFILEREF:			return TDCC_FILEREF;
		case ID_SORT_BYFLAG:			return TDCC_FLAG;
		case ID_SORT_BYICON:			return TDCC_ICON;
		case ID_SORT_BYID:				return TDCC_ID;
		case ID_SORT_BYMODIFYDATE:		return TDCC_LASTMOD;
		case ID_SORT_BYNAME:			return TDCC_CLIENT;
		case ID_SORT_BYPATH:			return TDCC_PATH;
		case ID_SORT_BYPERCENT:			return TDCC_PERCENT;
		case ID_SORT_BYPOSITION:		return TDCC_POSITION;
		case ID_SORT_BYPRIORITY:		return TDCC_PRIORITY;
		case ID_SORT_BYRECENTEDIT:		return TDCC_RECENTEDIT;
		case ID_SORT_BYRECURRENCE:		return TDCC_RECURRENCE;
		case ID_SORT_BYREMAINING:		return TDCC_REMAINING;
		case ID_SORT_BYRISK:			return TDCC_RISK;
		case ID_SORT_BYSTARTDATE:		return TDCC_STARTDATE;
		case ID_SORT_BYSTATUS:			return TDCC_STATUS;
		case ID_SORT_BYSUBTASKDONE:		return TDCC_SUBTASKDONE;
		case ID_SORT_BYTAG:				return TDCC_TAGS;
		case ID_SORT_BYTIMEEST:			return TDCC_TIMEEST;
		case ID_SORT_BYTIMESPENT:		return TDCC_TIMESPENT;
		case ID_SORT_BYTIMETRACKING:	return TDCC_TRACKTIME;
		case ID_SORT_BYVERSION:			return TDCC_VERSION;
		case ID_SORT_NONE:				return TDCC_NONE;

		}
		
		// handle custom columns
		if (nSortID >= ID_SORT_BYCUSTOMCOLUMN_FIRST && nSortID <= ID_SORT_BYCUSTOMCOLUMN_LAST)
			return (TDC_COLUMN)(TDCC_CUSTOMCOLUMN_FIRST + (nSortID - ID_SORT_BYCUSTOMCOLUMN_FIRST));
		
		// all else
		return TDCC_NONE;
	}
	
	static UINT MapColumnToSortID(TDC_COLUMN nColumn) 
	{
		switch (nColumn)
		{
		case TDCC_ALLOCBY:		return ID_SORT_BYALLOCBY;
		case TDCC_ALLOCTO:		return ID_SORT_BYALLOCTO;
		case TDCC_CATEGORY:		return ID_SORT_BYCATEGORY;
		case TDCC_CLIENT:		return ID_SORT_BYNAME;
		case TDCC_COLOR:		return ID_SORT_BYCOLOR;
		case TDCC_COST:			return ID_SORT_BYCOST;
		case TDCC_CREATEDBY:	return ID_SORT_BYCREATEDBY;
		case TDCC_CREATIONDATE:	return ID_SORT_BYCREATIONDATE;
		case TDCC_DEPENDENCY:	return ID_SORT_BYDEPENDENCY;
		case TDCC_DONE:			return ID_SORT_BYDONE;
		case TDCC_DONEDATE:		return ID_SORT_BYDONEDATE;
		case TDCC_DUEDATE:		return ID_SORT_BYDUEDATE;
		case TDCC_EXTERNALID:	return ID_SORT_BYEXTERNALID;
		case TDCC_FILEREF:		return ID_SORT_BYFILEREF;
		case TDCC_FLAG:			return ID_SORT_BYFLAG;
		case TDCC_ICON:			return ID_SORT_BYICON;
		case TDCC_ID:			return ID_SORT_BYID;
		case TDCC_LASTMOD:		return ID_SORT_BYMODIFYDATE;
		case TDCC_NONE:			return ID_SORT_NONE;
		case TDCC_PATH:			return ID_SORT_BYPATH;
		case TDCC_PERCENT:		return ID_SORT_BYPERCENT;
		case TDCC_POSITION:		return ID_SORT_BYPOSITION;
		case TDCC_PRIORITY:		return ID_SORT_BYPRIORITY;
		case TDCC_RECENTEDIT:	return ID_SORT_BYRECENTEDIT;
		case TDCC_RECURRENCE:	return ID_SORT_BYRECURRENCE;
		case TDCC_REMAINING:	return ID_SORT_BYREMAINING;
		case TDCC_RISK:			return ID_SORT_BYRISK;
		case TDCC_STARTDATE:	return ID_SORT_BYSTARTDATE;
		case TDCC_STATUS:		return ID_SORT_BYSTATUS;
		case TDCC_SUBTASKDONE:	return ID_SORT_BYSUBTASKDONE; 
		case TDCC_TAGS:			return ID_SORT_BYTAG;
		case TDCC_TIMEEST:		return ID_SORT_BYTIMEEST;
		case TDCC_TIMESPENT:	return ID_SORT_BYTIMESPENT;
		case TDCC_TRACKTIME:	return ID_SORT_BYTIMETRACKING;
		case TDCC_VERSION:		return ID_SORT_BYVERSION;

		}
		
		// handle custom columns
		if (nColumn >= TDCC_CUSTOMCOLUMN_FIRST && nColumn < TDCC_CUSTOMCOLUMN_LAST)
		{
			return (ID_SORT_BYCUSTOMCOLUMN_FIRST + (nColumn - TDCC_CUSTOMCOLUMN_FIRST));
		}
		
		// all else
		return 0;
	}

	static TDC_COLUMN MapAttributeToColumn(TDC_ATTRIBUTE nAttrib) 
	{
		switch (nAttrib)
		{
		case TDCA_ALLOCBY:			return TDCC_ALLOCBY;
		case TDCA_ALLOCTO:			return TDCC_ALLOCTO;
		case TDCA_CATEGORY:			return TDCC_CATEGORY;
		case TDCA_COLOR:			return TDCC_COLOR;
		case TDCA_COST:				return TDCC_COST;
		case TDCA_CREATEDBY:		return TDCC_CREATEDBY;
		case TDCA_CREATIONDATE:		return TDCC_CREATIONDATE;
		case TDCA_DEPENDENCY:		return TDCC_DEPENDENCY;
		case TDCA_DONEDATE:			return TDCC_DONEDATE; 
		case TDCA_DONETIME:			return TDCC_DONETIME;
		case TDCA_DUEDATE:			return TDCC_DUEDATE;
		case TDCA_DUETIME:			return TDCC_DUETIME;
		case TDCA_EXTERNALID:		return TDCC_EXTERNALID;
		case TDCA_FILEREF:			return TDCC_FILEREF;
		case TDCA_FLAG:				return TDCC_FLAG;
		case TDCA_ICON:				return TDCC_ICON;
		case TDCA_ID:				return TDCC_ID;
		case TDCA_LASTMOD:			return TDCC_LASTMOD;
		case TDCA_NONE:				return TDCC_NONE;
		case TDCA_PATH:				return TDCC_PATH;
		case TDCA_PERCENT:			return TDCC_PERCENT;
		case TDCA_POSITION:			return TDCC_POSITION;
		case TDCA_PRIORITY:			return TDCC_PRIORITY;
		case TDCA_RECURRENCE:		return TDCC_RECURRENCE;
		case TDCA_RISK:				return TDCC_RISK;
		case TDCA_STARTDATE:		return TDCC_STARTDATE;
		case TDCA_STARTTIME:		return TDCC_STARTTIME;
		case TDCA_STATUS:			return TDCC_STATUS;
		case TDCA_SUBTASKDONE:		return TDCC_SUBTASKDONE;
		case TDCA_TAGS:				return TDCC_TAGS;
		case TDCA_TASKNAME:			return TDCC_CLIENT;
		case TDCA_TIMEEST:			return TDCC_TIMEEST;
		case TDCA_TIMESPENT:		return TDCC_TIMESPENT;
		case TDCA_VERSION:			return TDCC_VERSION;
		}
		
		// handle custom columns
		if (nAttrib >= TDCA_CUSTOMATTRIB_FIRST && nAttrib <= TDCA_CUSTOMATTRIB_LAST)
		{
			return (TDC_COLUMN)(TDCC_CUSTOMCOLUMN_FIRST + (nAttrib - TDCA_CUSTOMATTRIB_FIRST));
		}
		
		// all else
		return TDCC_NONE;
	}

	static UINT MapAttributeToCtrlID(TDC_ATTRIBUTE nAttrib) 
	{
		// custom columns not supported for now
		// We could have used CTDCCustomAttributeHelper but that's an unwanted dependency
		ASSERT(nAttrib < TDCA_CUSTOMATTRIB_FIRST || nAttrib > TDCA_CUSTOMATTRIB_LAST);
		
		switch (nAttrib)
		{
		case TDCA_ALLOCBY:			return IDC_ALLOCBY;
		case TDCA_ALLOCTO:			return IDC_ALLOCTO;
		case TDCA_CATEGORY:			return IDC_CATEGORY;
		case TDCA_COST:				return IDC_COST;
		case TDCA_DEPENDENCY:		return IDC_DEPENDS;
		case TDCA_DONEDATE:			return IDC_DONEDATE; 
		case TDCA_DONETIME:			return IDC_DONETIME;
		case TDCA_DUEDATE:			return IDC_DUEDATE;
		case TDCA_DUETIME:			return IDC_DUETIME;
		case TDCA_EXTERNALID:		return IDC_EXTERNALID;
		case TDCA_FILEREF:			return IDC_FILEPATH;
		case TDCA_PERCENT:			return IDC_PERCENT;
		case TDCA_PRIORITY:			return IDC_PRIORITY;
		case TDCA_RECURRENCE:		return IDC_RECURRENCE;
		case TDCA_RISK:				return IDC_RISK;
		case TDCA_STARTDATE:		return IDC_STARTDATE;
		case TDCA_STARTTIME:		return IDC_STARTTIME;
		case TDCA_STATUS:			return IDC_STATUS;
		case TDCA_TAGS:				return IDC_TAGS;
		case TDCA_TASKNAME:			return IDC_TASKTREELIST;
		case TDCA_TIMEEST:			return IDC_TIMEEST;
		case TDCA_TIMESPENT:		return IDC_TIMESPENT;
		case TDCA_VERSION:			return IDC_VERSION;
		case TDCA_COLOR:			return IDC_COLOUR;

		// don't have controls
		case TDCA_SUBTASKDONE:
		case TDCA_POSITION:
		case TDCA_PATH:
		case TDCA_NONE:
		case TDCA_FLAG:
		case TDCA_ID:
		case TDCA_LASTMOD:
		case TDCA_CREATEDBY:
		case TDCA_CREATIONDATE:
		case TDCA_ICON:
			break;
		}
		
		return (UINT)-1;
	}
	
	static TDC_ATTRIBUTE MapCtrlIDToAttribute(UINT nCtrlID) 
	{
		switch (nCtrlID)
		{
		case IDC_ALLOCBY:		return TDCA_ALLOCBY;		
		case IDC_ALLOCTO:		return TDCA_ALLOCTO;			
		case IDC_CATEGORY:		return TDCA_CATEGORY;			
		case IDC_COST:			return TDCA_COST;				
		case IDC_DEPENDS:		return TDCA_DEPENDENCY;		
		case IDC_DONEDATE:		return TDCA_DONEDATE;			
		case IDC_DONETIME:		return TDCA_DONETIME;			
		case IDC_DUEDATE:		return TDCA_DUEDATE;			
		case IDC_DUETIME:		return TDCA_DUETIME;			
		case IDC_EXTERNALID:	return TDCA_EXTERNALID;		
		case IDC_FILEPATH:		return TDCA_FILEREF;			
		case IDC_PERCENT:		return TDCA_PERCENT;			
		case IDC_PRIORITY:		return TDCA_PRIORITY;			
		case IDC_RECURRENCE:	return TDCA_RECURRENCE;		
		case IDC_RISK:			return TDCA_RISK;				
		case IDC_STARTDATE:		return TDCA_STARTDATE;		
		case IDC_STARTTIME:		return TDCA_STARTTIME;		
		case IDC_STATUS:		return TDCA_STATUS;			
		case IDC_TAGS:			return TDCA_TAGS;				
		case IDC_TASKTREELIST:	return TDCA_TASKNAME;			
		case IDC_TIMEEST:		return TDCA_TIMEEST;			
		case IDC_TIMESPENT:		return TDCA_TIMESPENT;		
		case IDC_VERSION:		return TDCA_VERSION;			
		case IDC_COLOUR:		return TDCA_COLOR;	
			
		default:
			if (nCtrlID >= IDC_FIRST_CUSTOMDATAFIELD && (nCtrlID <= IDC_LAST_CUSTOMDATAFIELD))
				return TDCA_CUSTOMATTRIB;
			break;
		}
		
		return TDCA_NONE;
	}
	
	static TDC_ATTRIBUTE MapColumnToAttribute(TDC_COLUMN nColumn) 
	{
		switch (nColumn)
		{
		case TDCC_ALLOCBY:		return TDCA_ALLOCBY;
		case TDCC_ALLOCTO:		return TDCA_ALLOCTO;
		case TDCC_CATEGORY:		return TDCA_CATEGORY;
		case TDCC_CLIENT:		return TDCA_TASKNAME;
		case TDCC_COLOR:		return TDCA_COLOR;
		case TDCC_COST:			return TDCA_COST;
		case TDCC_CREATEDBY:	return TDCA_CREATEDBY;
		case TDCC_CREATIONDATE:	return TDCA_CREATIONDATE;
		case TDCC_DEPENDENCY:	return TDCA_DEPENDENCY;
		case TDCC_DONEDATE:		return TDCA_DONEDATE;
		case TDCC_DONETIME:		return TDCA_DONETIME;
		case TDCC_DUEDATE:		return TDCA_DUEDATE;
		case TDCC_DUETIME:		return TDCA_DUETIME;
		case TDCC_EXTERNALID:	return TDCA_EXTERNALID;
		case TDCC_FILEREF:		return TDCA_FILEREF;
		case TDCC_FLAG:			return TDCA_FLAG;
		case TDCC_ICON:			return TDCA_ICON;
		case TDCC_ID:			return TDCA_ID;
		case TDCC_LASTMOD:		return TDCA_LASTMOD;
		case TDCC_NONE:			return TDCA_NONE;
		case TDCC_PARENTID:		return TDCA_PARENTID;
		case TDCC_PATH:			return TDCA_PATH;
		case TDCC_PERCENT:		return TDCA_PERCENT;
		case TDCC_POSITION:		return TDCA_POSITION;
		case TDCC_PRIORITY:		return TDCA_PRIORITY;
		case TDCC_RECURRENCE:	return TDCA_RECURRENCE;
		case TDCC_RISK:			return TDCA_RISK;
		case TDCC_STARTDATE:	return TDCA_STARTDATE;
		case TDCC_STARTTIME:	return TDCA_STARTTIME;
		case TDCC_STATUS:		return TDCA_STATUS;
		case TDCC_SUBTASKDONE:	return TDCA_SUBTASKDONE;
		case TDCC_TAGS:			return TDCA_TAGS;
		case TDCC_TIMEEST:		return TDCA_TIMEEST;
		case TDCC_TIMESPENT:	return TDCA_TIMESPENT;
		case TDCC_VERSION:		return TDCA_VERSION;
		}
		
		// handle custom columns
		if (nColumn >= TDCC_CUSTOMCOLUMN_FIRST && nColumn < TDCC_CUSTOMCOLUMN_LAST)
		{
			return (TDC_ATTRIBUTE)(TDCA_CUSTOMATTRIB_FIRST + (nColumn - TDCC_CUSTOMCOLUMN_FIRST));
		}

		// all else
		return TDCA_NONE;
	}

	static int MapColumnsToAttributes(const CTDCColumnIDArray& aCols, CTDCAttributeArray& aAttrib)
	{
		aAttrib.RemoveAll();
		int nCol = aCols.GetSize();
		
		while (nCol--)
		{
			TDC_ATTRIBUTE att = MapColumnToAttribute(aCols[nCol]);
			
			if (att != TDCA_NONE)
				aAttrib.Add(att);
		}
		
		return aAttrib.GetSize();
	}

	static TDC_OFFSET MapUnitsToDateOffset(int nUnits)
	{
		switch (nUnits)
		{
		case TDCU_DAYS:		return TDCO_DAYS;
		case TDCU_WEEKS:	return TDCO_WEEKS;
		case TDCU_MONTHS:	return TDCO_MONTHS;
		case TDCU_YEARS:	return TDCO_YEARS;
		}

		ASSERT(0);
		return (TDC_OFFSET)0;
	}

	static TDC_ATTRIBUTE MapDateToAttribute(TDC_DATE nDate)
	{
		switch (nDate)
		{
		case TDCD_CREATE:	return TDCA_CREATIONDATE;
		case TDCD_LASTMOD:	return TDCA_LASTMOD;
			
		case TDCD_START:	
		case TDCD_STARTDATE:return TDCA_STARTDATE;
		case TDCD_STARTTIME:return TDCA_STARTTIME;
			
		case TDCD_DUE:		
		case TDCD_DUEDATE:	return TDCA_DUEDATE;
		case TDCD_DUETIME:	return TDCA_DUETIME;	

		case TDCD_DONE:		
		case TDCD_DONEDATE:	return TDCA_DONEDATE;
		case TDCD_DONETIME:	return TDCA_DONETIME;
		}
		
		// else
		ASSERT(0);
		return TDCA_NONE;
	}

	static TDC_DATE MapColumnToDate(TDC_COLUMN nCol)
	{
		switch (nCol)
		{
		case TDCC_LASTMOD:		return TDCD_LASTMOD;
		case TDCC_DUEDATE:		return TDCD_DUE;
		case TDCC_CREATIONDATE:	return TDCD_CREATE;
		case TDCC_STARTDATE:	return TDCD_START;
		case TDCC_DONEDATE:		return TDCD_DONE;

		default:
			if ((nCol >= TDCC_CUSTOMCOLUMN_FIRST) && 
				(nCol < TDCC_CUSTOMCOLUMN_LAST))
			{
				return TDCD_CUSTOM;
			}
		}
		
		// else
		ASSERT(0);
		return TDCD_NONE;
	}
}
#endif // AFX_TDCMAP_H__5951FDE6_508A_4A9D_A55D_D16EB026AEF7__INCLUDED_