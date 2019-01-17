// TimeHelper.cpp: implementation of the TimeHelper class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "TimeHelper.h"
#include "misc.h"

#include <locale.h>
#include <math.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////

const double MINS2HOURS			= 60;
double CTimeHelper::HOURS2DAYS	= 8; // user definable
double CTimeHelper::DAYS2WEEKS	= 5; // user definable
const double WEEKS2MONTHS		= 4.348;
const double MONTHS2YEARS		= 12;
const double FUDGE				= 1e-6;

//////////////////////////////////////////////////////////////////////

// assume working days pivot about 1.30pm
// eg. a standard working day of 8 hours (+1 for lunch)
// starts at 9am (13.50 - 4.5) and 
// ends at 6pm (13.30 + 4.5)
const double MIDDAY			= 13.5;
const double LUNCHSTARTTIME = (MIDDAY - 0.5);
const double LUNCHENDTIME	= (MIDDAY + 0.5);

//////////////////////////////////////////////////////////////////////

CMap<int, int, TCHAR, TCHAR&> CTimeHelper::MAPUNIT2CH; // user definable

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CTimeHelper::CTimeHelper() : m_dHours2Days(HOURS2DAYS), m_dDaysToWeeks(DAYS2WEEKS)
{
}

CTimeHelper::CTimeHelper(double dHoursInDay, double dDaysInWeek) 
	: m_dHours2Days(dHoursInDay), m_dDaysToWeeks(dDaysInWeek)
{
}

double CTimeHelper::GetTimeOnly(double dDate)
{
	return (dDate - floor(dDate));
}

double CTimeHelper::GetTimeOnly(const COleDateTime& date)
{
	return GetTimeOnly(date.m_dt);
}

THU_WORKDAYPERIOD CTimeHelper::GetWorkdayPeriod(const COleDateTime& date) const
{
	double dTime = GetTimeOnly(date);

	if (dTime < GetStartOfWorkday())
	{
		return THU_BEFORE;
	}
	else if (dTime < GetStartOfWorkdayLunch())
	{
		return THU_MORNING;
	}
	else if (dTime < GetEndOfWorkdayLunch())
	{
		return THU_LUNCH;
	}
	else if (dTime < GetEndOfWorkday())
	{
		return THU_AFTERNOON;
	}

	// else
	return THU_AFTER;
}

double CTimeHelper::GetStartOfWorkday(BOOL bInDays) const
{
	double dHours = (LUNCHSTARTTIME - (m_dHours2Days / 2));
	
	return (bInDays ? (dHours / 24) : dHours);
}

double CTimeHelper::GetStartOfWorkdayLunch(BOOL bInDays) const
{
	double dHours = LUNCHSTARTTIME;

	return (bInDays ? (dHours / 24) : dHours);
}

double CTimeHelper::GetEndOfWorkday(BOOL bInDays) const
{
	double dHours = (LUNCHENDTIME + (m_dHours2Days / 2));

	return (bInDays ? (dHours / 24) : dHours);
}

double CTimeHelper::GetEndOfWorkdayLunch(BOOL bInDays) const
{
	double dHours = LUNCHENDTIME;

	return (bInDays ? (dHours / 24) : dHours);
}

void CTimeHelper::CalculatePartWorkdays(const COleDateTime& dtStart, const COleDateTime& dtEnd,
										double& dPartStartDay, double& dPartEndDay, BOOL bInDays) const
{
	// calculate time to end of day
	dPartStartDay = (GetTimeOnly(dtStart) * 24);

	switch (GetWorkdayPeriod(dtStart))
	{
	case THU_BEFORE:
		dPartStartDay = m_dHours2Days;
		break;

	case THU_MORNING:
		dPartStartDay = ((GetStartOfWorkdayLunch(FALSE) - dPartStartDay) + (m_dHours2Days / 2));
		break;

	case THU_LUNCH:
		dPartStartDay = (m_dHours2Days / 2);
		break;

	case THU_AFTERNOON:
		dPartStartDay = (GetEndOfWorkday(FALSE) - dPartStartDay);
		break;

	case THU_AFTER:
		dPartStartDay = 0.0;
		break;

	default:
		ASSERT(0);
		dPartStartDay = 0.0;
		break;
	}

	// calculate time from start of day
	dPartEndDay = (GetTimeOnly(dtEnd) * 24);

	switch (GetWorkdayPeriod(dtEnd))
	{
	case THU_BEFORE:
		dPartEndDay = 0.0;
		break;

	case THU_MORNING:
		dPartEndDay = (dPartEndDay - GetStartOfWorkday(FALSE));
		break;

	case THU_LUNCH:
		dPartEndDay = (m_dHours2Days / 2);
		break;

	case THU_AFTERNOON:
		dPartEndDay = ((dPartEndDay - GetEndOfWorkdayLunch(FALSE)) + (m_dHours2Days / 2));
		break;

	case THU_AFTER:
		dPartEndDay = m_dHours2Days;
		break;

	default:
		ASSERT(0);
		dPartEndDay = 0.0;
		break;
	}

	if (bInDays)
	{
		dPartStartDay /= m_dHours2Days;
		dPartEndDay /= m_dHours2Days;
	}
}

BOOL CTimeHelper::IsValidUnit(int nUnits)
{
	switch (nUnits)
	{
	case THU_MINS:
	case THU_HOURS:
	case THU_DAYS:
	case THU_WEEKS:
	case THU_MONTHS:
	case THU_YEARS:
		return TRUE;
	}

	return FALSE;
}

double CTimeHelper::GetTime(double dTime, int nFromUnits, int nToUnits) const
{
	// sanity check
	if (!IsValidUnit(nFromUnits) || !IsValidUnit(nToUnits))
	{	
		ASSERT(0);
		return 0.0;
	}
	
	if (nFromUnits == nToUnits)
	{
		return dTime;
	}
	else if (Compare(nFromUnits, nToUnits) > 0)
	{
		while (Compare(nFromUnits, nToUnits) > 0)
		{
			switch (nFromUnits)
			{
			case THU_HOURS:
				dTime *= MINS2HOURS;
				nFromUnits = THU_MINS;
				break;
				
			case THU_DAYS:
				dTime *= m_dHours2Days;
				nFromUnits = THU_HOURS;
				break;

			case THU_WEEKS:
				dTime *= m_dDaysToWeeks;
				nFromUnits = THU_DAYS;
				break;
				
			case THU_MONTHS:
				dTime *= WEEKS2MONTHS;
				nFromUnits = THU_WEEKS;
				break;
				
			case THU_YEARS:
				dTime *= MONTHS2YEARS;
				nFromUnits = THU_MONTHS;
				break;
			}
		}
	}
	else // nFromUnits < nToUnits
	{
		while (Compare(nFromUnits, nToUnits) < 0)
		{
			switch (nFromUnits)
			{
			case THU_MINS:
				dTime /= MINS2HOURS;
				nFromUnits = THU_HOURS;
				break;

			case THU_HOURS:
				dTime /= m_dHours2Days;
				nFromUnits = THU_DAYS;
				break;

			case THU_DAYS:
				dTime /= m_dDaysToWeeks;
				nFromUnits = THU_WEEKS;
				break;
				
			case THU_WEEKS:
				dTime /= WEEKS2MONTHS;
				nFromUnits = THU_MONTHS;
				break;
				
			case THU_MONTHS:
				dTime /= MONTHS2YEARS;
				nFromUnits = THU_YEARS;
				break;
			}
		}
	}

	return dTime;
}

CString CTimeHelper::FormatClockTime(const COleDateTime& dtTime, BOOL bIncSeconds, BOOL bISO)
{
	return FormatClockTime(dtTime.GetHour(), dtTime.GetMinute(), dtTime.GetSecond(), bIncSeconds, bISO);
}

CString CTimeHelper::FormatClockTime(int nHour, int nMin, int nSec, BOOL bIncSeconds, BOOL bISO)
{
	CString sFormat;
	LPCTSTR szFormat = NULL;

	if (bISO)
	{
		sFormat = _T("HH:mm");
			
		if (bIncSeconds)
			sFormat += _T(":ss");

		szFormat = sFormat;
	}
	
	CString sTime;
	SYSTEMTIME st = { 0, 0, 0, 0, (WORD)nHour, (WORD)nMin, (WORD)nSec, 0 };
	
	::GetTimeFormat(0, bIncSeconds ? 0 : TIME_NOSECONDS, &st, szFormat, sTime.GetBuffer(50), 49);
	sTime.ReleaseBuffer();

	return sTime;
}

CString CTimeHelper::FormatTime(double dTime, int nDecPlaces) const
{
	return FormatTime(dTime, 0, nDecPlaces);
}

CString CTimeHelper::FormatTime(double dTime, int nUnits, int nDecPlaces) const
{
	// sanity check
	if (nUnits && !IsValidUnit(nUnits))
	{	
		ASSERT(0);
		return _T("");
	}

	CString sTime = Misc::Format(dTime, nDecPlaces);
	TCHAR cUnits;
	
	if (nUnits && MAPUNIT2CH.Lookup(nUnits, cUnits))
	{
		CString sTemp(sTime);
		sTime.Format(_T("%s %c"), sTemp, cUnits);
	}
	
	return sTime;
}

void CTimeHelper::SetUnits(int nUnits, TCHAR cUnits)
{
	// sanity check
	if (!IsValidUnit(nUnits))
	{	
		ASSERT(0);
		return ;
	}

	MAPUNIT2CH[nUnits] = cUnits;
}

void CTimeHelper::SetUnits(int nUnits, LPCTSTR szUnits)
{
	// sanity check
	if (!IsValidUnit(nUnits))
	{	
		ASSERT(0);
		return ;
	}

	if (szUnits && *szUnits)
		SetUnits(nUnits, szUnits[0]);
}

TCHAR CTimeHelper::GetUnits(int nUnits)
{
	// sanity check
	if (!IsValidUnit(nUnits))
	{	
		ASSERT(0);
		return 0;
	}

	// handle first time
	if (MAPUNIT2CH.GetCount() == 0)
	{
		SetUnits(THU_MINS, 'm');	
		SetUnits(THU_HOURS, 'H');	
		SetUnits(THU_DAYS, 'D');	
		SetUnits(THU_WEEKS, 'W');	
		SetUnits(THU_MONTHS, 'M');	
		SetUnits(THU_YEARS, 'Y');	
	}

	TCHAR cUnits = 0;
	MAPUNIT2CH.Lookup(nUnits, cUnits);

	return cUnits;
}

CString CTimeHelper::FormatTimeHMS(double dTime, int nUnitsFrom, BOOL bDecPlaces) const
{
	// sanity check
	if (!IsValidUnit(nUnitsFrom))
	{	
		ASSERT(0);
		return _T("");
	}

	// handle negative times
	BOOL bNegative = (dTime < 0.0);

	if (bNegative)
		dTime = -dTime;

	// convert the time to minutes 
	double dMins = GetTime(dTime, nUnitsFrom, THU_MINS);
	
	// and all the others up to years
	double dHours = dMins / MINS2HOURS;
	double dDays = dHours / m_dHours2Days;
	double dWeeks = dDays / m_dDaysToWeeks;
	double dMonths = dWeeks / WEEKS2MONTHS;
	double dYears = dMonths / MONTHS2YEARS;
	
	CString sTime;
	
	if (dYears >= 1.0)
	{
		sTime = FormatTimeHMS(dYears, THU_YEARS, THU_MONTHS, MONTHS2YEARS, bDecPlaces);
	}
	else if (dMonths >= 1.0)
	{
		sTime = FormatTimeHMS(dMonths, THU_MONTHS, THU_WEEKS, WEEKS2MONTHS, bDecPlaces);
	}
	else if (dWeeks >= 1.0)
	{
		sTime = FormatTimeHMS(dWeeks, THU_WEEKS, THU_DAYS, m_dDaysToWeeks, bDecPlaces);
	}	
	else if (dDays >= 1.0)
	{
		sTime = FormatTimeHMS(dDays, THU_DAYS, THU_HOURS, m_dHours2Days, bDecPlaces);
	}
	else if (dHours >= 1.0)
	{
		sTime = FormatTimeHMS(dHours, THU_HOURS, THU_MINS, MINS2HOURS, bDecPlaces);
	}
	else if (dMins >= 1.0)
	{
		sTime = FormatTimeHMS(dMins, THU_MINS, THU_MINS, 0, FALSE);
	}
	sTime.MakeLower();

	// handle negative times
	if (bNegative)
		sTime = "-" + sTime;
	
	return sTime;
	
}

CString CTimeHelper::FormatTimeHMS(double dTime, int nUnits, int nLeftOverUnits, 
								   double dLeftOverMultiplier, BOOL bDecPlaces)
{
	// sanity check
	if (!IsValidUnit(nUnits) || !IsValidUnit(nLeftOverUnits))
	{	
		ASSERT(0);
		return _T("");
	}

	CString sTime;
	
	if (bDecPlaces)
	{
		double dLeftOver = ((dTime - (int)dTime) * dLeftOverMultiplier + FUDGE);

		// omit second element if zero
		if (((int)dLeftOver) == 0)
		{
			sTime.Format(_T("%d%c"), (int)dTime, GetUnits(nUnits));
		}
		else
		{
			sTime.Format(_T("%d%c%d%c"), (int)dTime, GetUnits(nUnits), 
						(int)dLeftOver, GetUnits(nLeftOverUnits));
		}
	}
	else
	{
		sTime.Format(_T("%d%c"), (int)(dTime + 0.5), GetUnits(nUnits));
	}
	
	return sTime;
}

int CTimeHelper::Compare(int nFromUnits, int nToUnits)
{
	// sanity check
	if (!IsValidUnit(nFromUnits) || !IsValidUnit(nToUnits))
	{	
		ASSERT(0);
		return 0;
	}

	if (nFromUnits == nToUnits)
		return 0;

	switch (nFromUnits)
	{
	case THU_MINS:
		return -1; // less than everything else
	
	case THU_HOURS:
		return (nToUnits == THU_MINS) ? 1 : -1;
	
	case THU_DAYS:
		return (nToUnits == THU_HOURS || nToUnits == THU_MINS) ? 1 : -1;
	
	case THU_WEEKS:
		return (nToUnits == THU_YEARS || nToUnits == THU_MONTHS) ? -1 : 1;
	
	case THU_MONTHS:
		return (nToUnits == THU_YEARS) ? -1 : 1;
	
	case THU_YEARS:
		return 1; // greater than everything else
	}

	// else
	return 0;
}

BOOL CTimeHelper::SetHoursInOneDay(double dHours)
{
	if (dHours <= 0 || dHours > 24)
		return FALSE;

	HOURS2DAYS = dHours;
	return TRUE;
}

BOOL CTimeHelper::SetDaysInOneWeek(double dDays)
{
	if (dDays <= 0 || dDays > 7)
		return FALSE;

	DAYS2WEEKS = dDays;
	return TRUE;
}

BOOL CTimeHelper::DecodeOffset(LPCTSTR szTime, double& dAmount, int& nUnits, BOOL bMustHaveSign)
{
	// sanity checks
	CString sTime(szTime);
	Misc::Trim(sTime);
	
	if (sTime.IsEmpty())
		return FALSE;

	// Sign defaults to +ve
	int nSign = 1;
	TCHAR nFirst = Misc::First(sTime);
	
	switch (nFirst)
	{
	case '+': 
		nSign = 1;  
		Misc::TrimFirst(sTime);
		break;
		
	case '-': 
		nSign = -1; 
		Misc::TrimFirst(sTime);
		break;
		
	default: 
		// Must have sign or be number
		if (bMustHaveSign || ((nFirst < '0') || (nFirst > '9')))				
			return FALSE;
	}

	// Trailing units
	TCHAR nLast = Misc::Last(sTime);
	
	if (IsValidUnit(nLast))
		nUnits = nLast;
	else
		nUnits = THU_HOURS;
	
	// Rest is number (note: ttof ignores any trailing letters)
	dAmount = (nSign * _ttof(sTime));
	return TRUE;
}
