// DateHelper.cpp: implementation of the CDateHelper class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DateHelper.h"
#include "TimeHelper.h"
#include "misc.h"

#include <math.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////

const double ONE_HOUR = (1.0 / 24.0);
const double HALF_HOUR = (ONE_HOUR / 2);
const double END_OF_DAY = ((24 * 60 * 60) - 1) / (24.0 * 60 * 60);

//////////////////////////////////////////////////////////////////////

DWORD CDateHelper::s_dwWeekend = (DHW_SATURDAY | DHW_SUNDAY);

//////////////////////////////////////////////////////////////////////

BOOL CDateHelper::IsValidDayInMonth(int nDay, int nMonth, int nYear)
{
	return (nMonth >= 1 && nMonth <= 12) &&
			(nDay >= 1 && nDay <= GetDaysInMonth(nMonth, nYear));
}

BOOL CDateHelper::IsValidDayOfMonth(int nDOW, int nWhich, int nMonth)
{
	return (nWhich >= 1 && nWhich <= 5) &&
			(nDOW >= 1 && nDOW <= 7) &&
			(nMonth >= 1 && nMonth <= 12);
}

int CDateHelper::CalcDaysFromTo(const COleDateTime& dateFrom, const COleDateTime& dateTo, BOOL bInclusive)
{
	ASSERT(IsDateSet(dateFrom) && IsDateSet(dateTo));
	
	if (IsDateSet(dateFrom) && IsDateSet(dateTo))
	{
		int nDiff = (int)(floor(dateTo) - floor(dateFrom));
		
		return nDiff + (bInclusive ? 1 : 0);
	}

	// else
	return 0;
}

int CDateHelper::CalcDaysFromTo(const COleDateTime& dateFrom, DH_DATE nTo, BOOL bInclusive)
{
	return CalcDaysFromTo(dateFrom, GetDate(nTo), bInclusive);
}

int CDateHelper::CalcDaysFromTo(DH_DATE nFrom, DH_DATE nTo, BOOL bInclusive)
{
	ASSERT (nFrom <= nTo);

	if (nFrom > nTo)
		return 0;
	
	else if (nFrom == nTo)
		return bInclusive ? 1 : 0;

	// else
	return CalcDaysFromTo(GetDate(nFrom), GetDate(nTo), bInclusive);
}

BOOL CDateHelper::DecodeDate(const CString& sDate, double& date, BOOL bAndTime)
{
	COleDateTime dt;

	if (DecodeDate(sDate, dt, bAndTime))
	{
		date = dt;
		return TRUE;
	}

	// else
	return FALSE;
}

BOOL CDateHelper::IsValidUnit(int nUnits)
{
	switch (nUnits)
	{
	case DHU_WEEKDAYS:
	case DHU_DAYS:
	case DHU_WEEKS:
	case DHU_MONTHS:
	case DHU_YEARS:
		return TRUE;
	}

	return FALSE;
}

// external
BOOL CDateHelper::DecodeOffset(LPCTSTR szDate, double& dAmount, int& nUnits, BOOL bMustHaveSign)
{
	return DecodeOffsetEx(szDate, dAmount, nUnits, DHU_DAYS, bMustHaveSign);
}

// internal
BOOL CDateHelper::DecodeOffsetEx(LPCTSTR szDate, double& dAmount, 
								int& nUnits, int nDefUnits, BOOL bMustHaveSign)
{
	// sanity checks
	ASSERT(IsValidUnit(nDefUnits) || (nDefUnits == 0));

	CString sDate(szDate);
	Misc::Trim(sDate);

	if (sDate.IsEmpty())
		return FALSE;

	// Sign defaults to +ve
	int nSign = 1;
	TCHAR nFirst = Misc::First(sDate);

	switch (nFirst)
	{
	case '+': 
		nSign = 1;  
		Misc::TrimFirst(sDate);
		break;

	case '-': 
		nSign = -1; 
		Misc::TrimFirst(sDate);
		break;

	default: 
		// Must have sign or be number
		if (bMustHaveSign || ((nFirst < '0') || (nFirst > '9')))				
			return FALSE;
	}

	// units
	TCHAR nLast = Misc::Last(sDate);

	if (IsValidUnit(nLast))
		nUnits = nLast;
	else
		nUnits = nDefUnits;

	// Rest is number (note: ttof ignores any trailing letters)
	dAmount = (nSign * _ttof(sDate));
	return TRUE;
}

BOOL CDateHelper::DecodeRelativeDate(LPCTSTR szDate, COleDateTime& date, BOOL bForceWeekday, BOOL bMustHaveSign)
{
	// sanity check
	CString sDate(szDate);
	Misc::Trim(sDate);

	if (sDate.IsEmpty())
		return FALSE;

	// clear date so we know if it changed
	ClearDate(date);

	// leading char can define starting date 
	TCHAR nFirst(Misc::First(sDate));

	switch (nFirst)
	{
	case 'T': 			date = GetDate(DHD_TODAY);			break;
	case DHU_WEEKDAYS:	date = GetDate(DHD_TODAY);			break;
	case DHU_DAYS:		date = GetDate(DHD_TODAY);			break;
	case DHU_WEEKS:		date = GetDate(DHD_ENDTHISWEEK);	break;
	case DHU_MONTHS:	date = GetDate(DHD_ENDTHISMONTH);	break;
	case DHU_YEARS:		date = GetDate(DHD_ENDTHISYEAR);	break;
	}

	if (IsDateSet(date))
		Misc::TrimFirst(sDate);
	else
		date = GetDate(DHD_TODAY); // default

	// The rest should be a relative date offset
	if (!sDate.IsEmpty())
	{
		int nUnits = 0;
		double dAmount = 0.0;

		if (!DecodeOffsetEx(sDate, dAmount, nUnits, 0, bMustHaveSign))
			return FALSE;

		// Handle missing units
		if (!nUnits)
		{
			if (IsValidUnit(nFirst))
				nUnits = nFirst;
			else
				nUnits = DHU_DAYS; // default
		}

		if (dAmount != 0.0)
			return OffsetDate(date, (int)dAmount, nUnits, bForceWeekday);
	}

	return TRUE;
}

BOOL CDateHelper::DecodeDate(const CString& sDate, COleDateTime& date, BOOL bAndTime)
{
	// check for valid date string
	if (date.ParseDateTime(sDate, VAR_DATEVALUEONLY))
	{
		if (bAndTime)
			return date.ParseDateTime(sDate, 0);

		// else
		return TRUE;
	}

	// Perhaps it's in ISO format
	if (DecodeISODate(sDate, date, bAndTime))
		return TRUE;

	// all else
	return DecodeLocalShortDate(sDate, date);
}

BOOL CDateHelper::GetTimeT(const COleDateTime& date, time_t& timeT)
{
	time64_t t64;
	
	if (!GetTimeT64(date, t64))
		return FALSE;

#ifdef _DEBUG
	COleDateTime dtCheck = GetDate(t64);
	ASSERT(fabs(dtCheck.m_dt - date.m_dt) < 0.00001);
#endif

	return GetTimeT(t64, timeT);
}

BOOL CDateHelper::GetTimeT(time64_t date, time_t& timeT)
{
	if ((date < 0) || (date > LONG_MAX))
		return FALSE;

	timeT = (time_t)date;
	return TRUE;
}

BOOL CDateHelper::GetTimeT64(const COleDateTime& date, time64_t& timeT)
{
	if (!IsDateSet(date))
		return FALSE;

	SYSTEMTIME st = { 0 };
	
	if (!date.GetAsSystemTime(st))
		return FALSE;

	SystemTimeToT64(&st, &timeT);
	ASSERT(timeT != 0);

	return TRUE;
}

COleDateTime CDateHelper::GetDate(time64_t date)
{
	SYSTEMTIME st = { 0 };
	T64ToSystemTime(&date, &st);

	return COleDateTime(st);
}

#if _MSC_VER < 1400
BOOL CDateHelper::DecodeDate(const CString& sDate, time_t& date, BOOL bAndTime)
{
	time64_t t64 = 0;

	if (!DecodeDate(sDate, t64, bAndTime))
		return FALSE;

	return GetTimeT(t64, date);
}
#endif

BOOL CDateHelper::DecodeDate(const CString& sDate, time64_t& date, BOOL bAndTime)
{
	COleDateTime dt;

	if (!DecodeDate(sDate, dt, bAndTime))
		return FALSE;

	return GetTimeT64(dt, date);
}

BOOL CDateHelper::DecodeISODate(const CString& sDate, COleDateTime& date, BOOL bAndTime)
{
	int nYear = -1, nMonth = -1, nDay = -1, nHour = 0, nMin = 0, nSec = 0;
	const LPCTSTR DATETIME_FMT = _T("%4d-%2d-%2dT%2d:%2d:%2d");

#if _MSC_VER >= 1400
	if (_stscanf_s(sDate, DATETIME_FMT, &nYear, &nMonth, &nDay, &nHour, &nMin, &nSec) >= 3)
#else
	if (_stscanf(sDate, DATETIME_FMT, &nYear, &nMonth, &nDay, &nHour, &nMin, &nSec) >= 3)
#endif
	{
		if (bAndTime)
			return (date.SetDateTime(nYear, nMonth, nDay, nHour, nMin, nSec) == 0);
		else
			return (date.SetDate(nYear, nMonth, nDay) == 0);
	}
			
	return FALSE;
}

BOOL CDateHelper::DecodeLocalShortDate(const CString& sDate, COleDateTime& date)
{
	if (!sDate.IsEmpty())
	{
		// split the string and format by the delimiter
		CString sFormat(Misc::GetShortDateFormat()), sDelim(Misc::GetDateSeparator());
		CStringArray aDateParts, aFmtParts;

		Misc::Split(sDate, aDateParts, sDelim, TRUE);
		Misc::Split(sFormat, aFmtParts, sDelim, TRUE);

		//ASSERT (aDateParts.GetSize() == aFmtParts.GetSize());

		if (aDateParts.GetSize() != aFmtParts.GetSize())
			return FALSE;

		// process the parts, deciphering the format
		int nYear = -1, nMonth = -1, nDay = -1;

		for (int nPart = 0; nPart < aFmtParts.GetSize(); nPart++)
		{
			if (aFmtParts[nPart].FindOneOf(_T("Yy")) != -1)
				nYear = _ttoi(aDateParts[nPart]);

			if (aFmtParts[nPart].FindOneOf(_T("Mm")) != -1)
				nMonth = _ttoi(aDateParts[nPart]);

			if (aFmtParts[nPart].FindOneOf(_T("Dd")) != -1)
				nDay = _ttoi(aDateParts[nPart]);
		}

		if (nYear != -1 && nMonth != -1 && nDay != -1)
			return (date.SetDate(nYear, nMonth, nDay) == 0);
	}

	// else
	ClearDate(date);
	return FALSE;
}

double CDateHelper::GetDate(DH_DATE nDate)
{
	COleDateTime date;

	switch (nDate)
	{
	case DHD_TODAY:
		date = COleDateTime::GetCurrentTime();
		break;

	case DHD_TOMORROW:
		date = GetDate(DHD_TODAY) + 1;
		break;

	case DHD_BEGINTHISWEEK:
		date  = (GetDate(DHD_ENDTHISWEEK) - 6);
		ASSERT(date.GetDayOfWeek() == GetFirstDayOfWeek());
		break;

	case DHD_ENDTHISWEEK:
		{
			// we must get the locale info to find out when this 
			// user's week starts
			date = COleDateTime::GetCurrentTime();
			
			// increment the date until we hit the last day of the week
			// note: we could have kept checking date.GetDayOfWeek but
			// it's a lot of calculation that's just not necessary
			int nLastDOW = GetLastDayOfWeek();
			int nDOW = date.GetDayOfWeek();
			
			while (nDOW != nLastDOW) 
			{
				date += 1;
				nDOW = GetNextDayOfWeek(nDOW);
			}
		}
		break;

	case DHD_ENDNEXTWEEK:
		return GetDate(DHD_ENDTHISWEEK) + 7;

	case DHD_ENDTHISMONTH:
		{
			date = COleDateTime::GetCurrentTime();
			int nThisMonth = date.GetMonth();

			while (date.GetMonth() == nThisMonth)
				date += 20; // much quicker than doing it one day at a time

			date -= date.GetDay(); // because we went into next month
		}
		break;

	case DHD_ENDNEXTMONTH:
		{
			date = GetDate(DHD_ENDTHISMONTH) + 1; // first day of next month
			int nNextMonth = date.GetMonth();

			while (date.GetMonth() == nNextMonth)
				date += 20; // much quicker than doing it one day at a time

			date -= date.GetDay(); // because we went into next month + 1
		}
		break;

	case DHD_ENDTHISYEAR:
		date = COleDateTime::GetCurrentTime(); // for current year
		date = COleDateTime(date.GetYear(), 12, 31, 0, 0, 0);
		break;

	case DHD_ENDNEXTYEAR:
		date = COleDateTime::GetCurrentTime(); // for current year
		date = COleDateTime(date.GetYear() + 1, 12, 31, 0, 0, 0);
		break;

	default:
		ASSERT (0);
		ClearDate(date);
		break;
	}

	return GetDateOnly(date);
}

BOOL CDateHelper::Min(COleDateTime& date, const COleDateTime& dtOther)
{
	if (IsDateSet(date))
	{
		if (IsDateSet(dtOther))
		{
			date.m_dt = min(dtOther.m_dt, date.m_dt);
		}
		// else no change
	}
	else if (IsDateSet(dtOther))
	{
		date = dtOther;
	}

	return IsDateSet(date);
}

BOOL CDateHelper::Max(COleDateTime& date, const COleDateTime& dtOther)
{
	if (IsDateSet(date))
	{
		if (IsDateSet(dtOther))
		{
			date.m_dt = max(dtOther.m_dt, date.m_dt);
		}
		// else no change
	}
	else if (IsDateSet(dtOther))
	{
		date = dtOther;
	}

	return IsDateSet(date);
}

BOOL CDateHelper::IsDateSet(const COleDateTime& date)
{
	return (date.m_status == COleDateTime::valid && date.m_dt != 0.0) ? TRUE : FALSE;
}

void CDateHelper::ClearDate(COleDateTime& date)
{
	date.m_dt = 0.0;
	date.m_status = COleDateTime::null;
}

int CDateHelper::CalcWeekdaysFromTo(const COleDateTime& dateFrom, const COleDateTime& dateTo, BOOL bInclusive)
{
	ASSERT(IsDateSet(dateFrom) && IsDateSet(dateTo));

	int nWeekdays = 0;
	
	if (IsDateSet(dateFrom) && IsDateSet(dateTo))
	{
		COleDateTime dFrom = GetDateOnly(dateFrom);
		COleDateTime dTo = GetDateOnly(dateTo);

		if (bInclusive)
			dTo += 1;

		int nDiff = (int)(double)(dTo - dFrom);

		if (nDiff > 0)
		{
			while (dFrom < dTo)
			{
				int nDOW = dFrom.GetDayOfWeek();

				if (!IsWeekend(nDOW))
					nWeekdays++;

				dFrom += 1;
			}
		}
	}

	return nWeekdays;
}

int CDateHelper::GetISODayOfWeek(const COleDateTime& date) 
{
	int nDOW = date.GetDayOfWeek(); // 1=Sun, 2=Mon, ..., 7=Sat

	// ISO DOWs: 1=Mon, 2=Tue, ..., 7=Sun
	switch (nDOW)
	{
	case 1: /* sun */ return 7;
	case 2: /* mon */ return 1;
	case 3: /* tue */ return 2;
	case 4: /* wed */ return 3;
	case 5: /* thu */ return 4;
	case 6: /* fri */ return 5;
	case 7: /* sat */ return 6;
	}

	ASSERT (0);
	return 1;
}

int CDateHelper::GetFirstDayOfWeek()
{
	TCHAR szFDW[3]; // 2 + NULL

	::GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IFIRSTDAYOFWEEK, szFDW, 2);

	int nFirstDOW = _ttoi(szFDW);

	// we need days to have same order as COleDateTime::GetDayOfWeek()
	// which is 1 (sun) - 7 (sat)
	switch (nFirstDOW)
	{
	case 0: /* mon */ return 2;
	case 1: /* tue */ return 3;
	case 2: /* wed */ return 4;
	case 3: /* thu */ return 5;
	case 4: /* fri */ return 6;
	case 5: /* sat */ return 7;
	case 6: /* sun */ return 1;
	}

	ASSERT (0);
	return 1;
}

int CDateHelper::GetLastDayOfWeek()
{
	switch (GetFirstDayOfWeek())
	{
	case 2: /* mon */ return 1; // sun
	case 3: /* tue */ return 2; // mon
	case 4: /* wed */ return 3; // tue
	case 5: /* thu */ return 4; // wed
	case 6: /* fri */ return 5; // thu
	case 7: /* sat */ return 6; // fri
	case 1: /* sun */ return 7; // sat
	}

	ASSERT (0);
	return 1;
}

int CDateHelper::GetNextDayOfWeek(int nDOW)
{
	switch (nDOW)
	{
	case 2: /* mon */ return 3; // tue
	case 3: /* tue */ return 4; // wed
	case 4: /* wed */ return 5; // thu
	case 5: /* thu */ return 6; // fri
	case 6: /* fri */ return 7; // sat
	case 7: /* sat */ return 1; // sun
	case 1: /* sun */ return 2; // mon
	}

	ASSERT (0);
	return 1;
}

COleDateTime CDateHelper::ToWeekday(const COleDateTime& date, BOOL bForwards)
{
	COleDateTime weekday(date);
	MakeWeekday(weekday, bForwards);

	return weekday;
}

BOOL CDateHelper::MakeWeekday(COleDateTime& date, BOOL bForwards)
{
	ASSERT(IsDateSet(date));

	// check we don't have a 7-day weekend
	if (!HasWeekend() || !IsDateSet(date))
	{
		ASSERT(0);
		return FALSE;
	}

	while (IsWeekend(date))
	{
		if (bForwards)
			date = (GetDateOnly(date).m_dt + 1.0);
		else
			date = (GetDateOnly(date).m_dt - 1.0);
	}

	return TRUE;
}

void CDateHelper::SetWeekendDays(DWORD dwDays)
{
	s_dwWeekend = dwDays;
}

DWORD CDateHelper::GetWeekdays()
{
	return (DHW_EVERYDAY & ~s_dwWeekend);
}

BOOL CDateHelper::IsWeekend(const COleDateTime& date)
{
	ASSERT(IsDateSet(date));
	
	return IsWeekend(date.GetDayOfWeek());
}

BOOL CDateHelper::IsWeekend(double dDate)
{
	return IsWeekend(COleDateTime(dDate));
}

BOOL CDateHelper::IsWeekend(int nDOW)
{
	switch (nDOW)
	{
	case 1:	return (s_dwWeekend & DHW_SUNDAY);
	case 2: return (s_dwWeekend & DHW_MONDAY);
	case 3: return (s_dwWeekend & DHW_TUESDAY);
	case 4: return (s_dwWeekend & DHW_WEDNESDAY);
	case 5: return (s_dwWeekend & DHW_THURSDAY);
	case 6: return (s_dwWeekend & DHW_FRIDAY);
	case 7: return (s_dwWeekend & DHW_SATURDAY);
	}

	ASSERT (0);
	return FALSE;
}

int CDateHelper::GetWeekendDuration()
{
	int nDuration = 0;

	for (int nDOW = 1; nDOW <= 7; nDOW++)
	{
		if (IsWeekend(nDOW))
			nDuration++;
	}

	return nDuration;
}

BOOL CDateHelper::HasWeekend()
{
	return (s_dwWeekend && (GetWeekendDuration() < 7));
}

CString CDateHelper::FormatDate(const COleDateTime& date, DWORD dwFlags)
{
	CString sDate, sTime, sDow;
	
	if (FormatDate(date, dwFlags, sDate, sTime, sDow))
	{
		if (!sDow.IsEmpty())
		{
			sDate = ' ' + sDate;
			sDate = sDow + sDate;
		}

		if (!sTime.IsEmpty())
		{
			sDate += ' ';
			sDate += sTime;
		}
	}

	return sDate;
}

CString CDateHelper::FormatCurrentDate(DWORD dwFlags)
{
	return FormatDate(COleDateTime::GetCurrentTime(), dwFlags);
}

BOOL CDateHelper::FormatDate(const COleDateTime& date, DWORD dwFlags, CString& sDate, CString& sTime, CString& sDow)
{
	if (!IsDateSet(date))
		return FALSE;

	SYSTEMTIME st;

	if (!date.GetAsSystemTime(st))
		return FALSE;

	// Date
	CString sFormat;

	if (dwFlags & DHFD_ISO)
		sFormat = "yyyy-MM-dd";
	else
		sFormat = Misc::GetShortDateFormat();

	::GetDateFormat(0, 0, &st, sFormat, sDate.GetBuffer(50), 49);
	sDate.ReleaseBuffer();

	// Day of week
	if (dwFlags & DHFD_DOW)
		sDow = GetWeekdayName(st.wDayOfWeek + 1, TRUE);

	// Time
	if (dwFlags & DHFD_TIME)
		sTime = CTimeHelper::FormatClockTime(st.wHour, st.wMinute, st.wSecond, !(dwFlags & DHFD_NOSEC), (dwFlags & DHFD_ISO));
	else 
		sTime.Empty();
	
	return TRUE;
}

BOOL CDateHelper::FormatCurrentDate(DWORD dwFlags, CString& sDate, CString& sTime, CString& sDow)
{
	return FormatDate(COleDateTime::GetCurrentTime(), dwFlags, sDate, sTime, sDow);
}

CString CDateHelper::GetWeekdayName(int nWeekday, BOOL bShort)
{
	LCTYPE lct = bShort ? LOCALE_SABBREVDAYNAME1 : LOCALE_SDAYNAME1;
	CString sWeekday;

	// data check
	if (nWeekday < 1 || nWeekday> 7)
		return "";

	switch (nWeekday)
	{
	case 1: lct = bShort ? LOCALE_SABBREVDAYNAME7 : LOCALE_SDAYNAME7; break; // sun
	case 2:	lct = bShort ? LOCALE_SABBREVDAYNAME1 : LOCALE_SDAYNAME1; break; // mon
	case 3:	lct = bShort ? LOCALE_SABBREVDAYNAME2 : LOCALE_SDAYNAME2; break; // tue
	case 4:	lct = bShort ? LOCALE_SABBREVDAYNAME3 : LOCALE_SDAYNAME3; break; // wed
	case 5:	lct = bShort ? LOCALE_SABBREVDAYNAME4 : LOCALE_SDAYNAME4; break; // thu
	case 6:	lct = bShort ? LOCALE_SABBREVDAYNAME5 : LOCALE_SDAYNAME5; break; // fri
	case 7:	lct = bShort ? LOCALE_SABBREVDAYNAME6 : LOCALE_SDAYNAME6; break; // sat
	}
	
	GetLocaleInfo(LOCALE_USER_DEFAULT, lct, sWeekday.GetBuffer(30),	29);
	sWeekday.ReleaseBuffer();

	return sWeekday;
}

int CDateHelper::CalcLongestWeekdayName(CDC* pDC, BOOL bShort)
{
	int nLongestWDWidth = 0;
		
	// figure out the longest day in pixels
	for (int nWD = 1; nWD <= 7; nWD++)
	{
		int nWDWidth = pDC->GetTextExtent(GetWeekdayName(nWD, bShort)).cx;
		nLongestWDWidth = max(nLongestWDWidth, nWDWidth);
	}
	
	return nLongestWDWidth;
}

BOOL CDateHelper::OffsetDate(COleDateTime& date, int nAmount, int nUnits, BOOL bForceWeekday)
{
	// sanity checks
	if (!IsDateSet(date) || !IsValidUnit(nUnits) || !nAmount)
		return FALSE;

	switch (nUnits)
	{
	case DHU_WEEKDAYS:
		{
			BOOL bForwards = (nAmount > 0);
			nAmount = abs(nAmount);

			while (nAmount--)
			{
				date.m_dt += (bForwards ? 1.0 : -1.0);

				// skip weekends
				MakeWeekday(date, bForwards);
			}
		}
		break;

	case DHU_DAYS:
		date += (double)nAmount;
		break;

	case DHU_WEEKS:
		date += nAmount * 7.0;
		break;

	case DHU_MONTHS:
		{
			SYSTEMTIME st;
			date.GetAsSystemTime(st);

			// are we at months end?
			int nDaysInMonth = GetDaysInMonth(st.wMonth, st.wYear);
			BOOL bEndOfMonth = (st.wDay == nDaysInMonth);

			// convert amount to years and months
			st.wYear = (WORD)((int)st.wYear + (nAmount / 12));
			st.wMonth = (WORD)((int)st.wMonth + (nAmount % 12));

			// handle overflow
			if (st.wMonth > 12)
			{
				st.wYear++;
				st.wMonth -= 12;
			}
			else if (st.wMonth < 1)
			{
				st.wYear--;
				st.wMonth += 12;
			}

			// if our start date was the end of the month make
			// sure out end date is too
			nDaysInMonth = GetDaysInMonth(st.wMonth, st.wYear);

			if (bEndOfMonth)
				st.wDay = (WORD)nDaysInMonth;

			else // clip dates to the end of the month
				st.wDay = min(st.wDay, (WORD)nDaysInMonth);

			// update time
			date = COleDateTime(st);
		}
		break;

	case DHU_YEARS:
		{
			SYSTEMTIME st;
			date.GetAsSystemTime(st);

			// update year
			st.wYear = (WORD)((int)st.wYear + nAmount);

			// update time
			date = COleDateTime(st);
		}
		break;
	}

	// does the caller only want weekdays
	return (bForceWeekday ? MakeWeekday(date) : date);
}

void CDateHelper::GetWeekdayNames(BOOL bShort, CStringArray& aWeekDays)
{
	aWeekDays.RemoveAll();

	for (int nDay = 1; nDay <= 7; nDay++)
		aWeekDays.Add(GetWeekdayName(nDay, bShort));
}

int CDateHelper::GetDaysInMonth(const COleDateTime& date)
{
	return GetDaysInMonth(date.GetMonth(), date.GetYear());
}

int CDateHelper::GetDaysInMonth(int nMonth, int nYear)
{
	// data check
	ASSERT(nMonth >= 1 && nMonth <= 12);

	if (nMonth < 1 || nMonth> 12)
		return 0;

	switch (nMonth)
	{
	case 1:  return 31; // jan
	case 2:  return IsLeapYear(nYear) ? 29 : 28; // feb
	case 3:  return 31; // mar
	case 4:  return 30; // apr
	case 5:  return 31; // may
	case 6:  return 30; // jun
	case 7:  return 31; // jul
	case 8:  return 31; // aug
	case 9:  return 30; // sep
	case 10: return 31; // oct
	case 11: return 30; // nov
	case 12: return 31; // dec
	}

	ASSERT(0);
	return 0;
}

BOOL CDateHelper::IsLeapYear(const COleDateTime& date)
{
	return IsLeapYear(date.GetYear());
}

BOOL CDateHelper::IsLeapYear(int nYear)
{
	return ((nYear % 4 == 0) && ((nYear % 100 != 0) || (nYear % 400 == 0)));
}

CString CDateHelper::GetMonthName(int nMonth, BOOL bShort)
{
	LCTYPE lct = LOCALE_SABBREVMONTHNAME1;
	CString sMonth;

	// data check
	if (nMonth < 1 || nMonth> 12)
		return "";

	switch (nMonth)
	{
	case 1:  lct = bShort ? LOCALE_SABBREVMONTHNAME1  : LOCALE_SMONTHNAME1;  break; // jan
	case 2:  lct = bShort ? LOCALE_SABBREVMONTHNAME2  : LOCALE_SMONTHNAME2;  break; // feb
	case 3:  lct = bShort ? LOCALE_SABBREVMONTHNAME3  : LOCALE_SMONTHNAME3;  break; // mar
	case 4:  lct = bShort ? LOCALE_SABBREVMONTHNAME4  : LOCALE_SMONTHNAME4;  break; // apr
	case 5:  lct = bShort ? LOCALE_SABBREVMONTHNAME5  : LOCALE_SMONTHNAME5;  break; // may
	case 6:  lct = bShort ? LOCALE_SABBREVMONTHNAME6  : LOCALE_SMONTHNAME6;  break; // jun
	case 7:  lct = bShort ? LOCALE_SABBREVMONTHNAME7  : LOCALE_SMONTHNAME7;  break; // jul
	case 8:  lct = bShort ? LOCALE_SABBREVMONTHNAME8  : LOCALE_SMONTHNAME8;  break; // aug
	case 9:  lct = bShort ? LOCALE_SABBREVMONTHNAME9  : LOCALE_SMONTHNAME9;  break; // sep
	case 10: lct = bShort ? LOCALE_SABBREVMONTHNAME10 : LOCALE_SMONTHNAME10; break; // oct
	case 11: lct = bShort ? LOCALE_SABBREVMONTHNAME11 : LOCALE_SMONTHNAME11; break; // nov
	case 12: lct = bShort ? LOCALE_SABBREVMONTHNAME12 : LOCALE_SMONTHNAME12; break; // dec
	}

	GetLocaleInfo(LOCALE_USER_DEFAULT, lct, sMonth.GetBuffer(30),	29);
	sMonth.ReleaseBuffer();

	return sMonth;
}

void CDateHelper::GetMonthNames(BOOL bShort, CStringArray& aMonths)
{
	aMonths.RemoveAll();

	for (int nMonth = 1; nMonth <= 12; nMonth++)
		aMonths.Add(GetMonthName(nMonth, bShort));
}

COleDateTime CDateHelper::GetTimeOnly(const COleDateTime& date)
{
	return (date.m_dt - GetDateOnly(date));
}

BOOL CDateHelper::DateHasTime(const COleDateTime& date)
{
	return (GetTimeOnly(date) > 0);
}

COleDateTime CDateHelper::GetDateOnly(const COleDateTime& date)
{
	return floor(date.m_dt);
}

void CDateHelper::SplitDate(const COleDateTime& date, double& dDateOnly, double& dTimeOnly)
{
	dDateOnly = GetDateOnly(date);
	dTimeOnly = GetTimeOnly(date);
}

COleDateTime CDateHelper::MakeDate(const COleDateTime& dtDateOnly, const COleDateTime& dtTimeOnly)
{
	double dDateOnly = GetDateOnly(dtDateOnly);
	double dTimeOnly = GetTimeOnly(dtTimeOnly);
	
	return dDateOnly + dTimeOnly;
}

int CDateHelper::CalcDayOfMonth(int nDOW, int nWhich, int nMonth, int nYear)
{
	// data check
	ASSERT(nMonth >= 1 && nMonth <= 12);
	ASSERT(nDOW >= 1 && nDOW <= 7);
	ASSERT(nWhich >= 1 && nWhich <= 5);

	if (nMonth < 1 || nMonth> 12 || nDOW < 1 || nDOW > 7 || nWhich < 1 || nWhich > 5)
		return -1;

	// start with first day of month
	int nDay = 1;
	COleDateTime date(nYear, nMonth, nDay, 0, 0, 0);

	// get it's day of week
	int nWeekDay = date.GetDayOfWeek();

	// move forwards until we hit the requested day of week
	while (nWeekDay != nDOW)
	{
		nDay++;
		nWeekDay = GetNextDayOfWeek(nWeekDay);
	}

	// add a week at a time
	nWhich--;

	while (nWhich--)
		nDay += 7;

	// if we've gone passed the end of the month
	// deduct a week
	if (nDay > GetDaysInMonth(nMonth, nYear))
		nDay -= 7;

	return nDay;
}

COleDateTime CDateHelper::CalcDate(int nDOW, int nWhich, int nMonth, int nYear)
{
	int nDay = CalcDayOfMonth(nDOW, nWhich, nMonth, nYear);

	if (nDay == -1)
		return COleDateTime((time_t)-1);

	return COleDateTime(nYear, nMonth, nDay, 0, 0, 0);
}

int CDateHelper::CalcWeekofYear(const COleDateTime& date)
{
	// http://en.wikipedia.org/wiki/ISO_week_date#Calculating_the_week_number_of_a_given_date
	//
	// Using ISO weekday numbers (running from 1 for Monday to 7 for Sunday), 
	// subtract the weekday from the ordinal date, then add 10. Divide the result by 7. 
	// Ignore the remainder; the quotient equals the week number.

	int nDayOfYear = date.GetDayOfYear();
	int nISODOW = GetISODayOfWeek(date);

	return ((nDayOfYear - nISODOW + 10) / 7);
}

COleDateTime CDateHelper::GetNearestDay(const COleDateTime& date, BOOL bEnd)
{
	COleDateTime dtDay = GetDateOnly(date);

	if ((date.m_dt - dtDay.m_dt) < 0.5)
	{
		if (bEnd) // end of previous day
		{
			dtDay = GetEndOfPreviousDay(dtDay);
		}
	}
	else if (bEnd)
	{
		dtDay.m_dt += END_OF_DAY;
	}
	else // start
	{
		dtDay.m_dt += 1.0;
	}

	return dtDay;
}

COleDateTime CDateHelper::GetNearestHalfDay(const COleDateTime& date, BOOL bEnd)
{
	COleDateTime dtDay = GetDateOnly(date);

	if ((date.m_dt - dtDay.m_dt) < 0.25)
	{
		if (bEnd) // end of previous day
		{
			dtDay = GetEndOfPreviousDay(dtDay);
		}
	}
	else if ((date.m_dt - dtDay.m_dt) < 0.75)
	{
		dtDay.m_dt += 0.5;
	}
	else if (bEnd)
	{
		dtDay.m_dt += END_OF_DAY;
	}
	else // start
	{
		dtDay.m_dt += 1.0;
	}

	return dtDay;
}

COleDateTime CDateHelper::GetEndOfPreviousDay(const COleDateTime& date)
{
	return (date.m_dt - 1.0 + END_OF_DAY);
}

COleDateTime CDateHelper::GetNearestHour(const COleDateTime& date, BOOL bEnd)
{
	COleDateTime dtDay = GetDateOnly(date);
	COleDateTime dtTest = dtDay.m_dt + HALF_HOUR;

	for (int nHour = 0; nHour < 24; nHour++)
	{
		if (date < dtTest)
		{
			// special case: dragging end and nHour == 0
			if (bEnd && (nHour == 0))
			{
				// end of previous day
				return GetEndOfPreviousDay(dtDay);
			}

			// else return hour 
			return MakeDate(dtDay, nHour * ONE_HOUR);
		}

		// test for next hour
		dtTest.m_dt += ONE_HOUR;
	}

	// handle end of day
	if (bEnd)
		return MakeDate(dtDay, END_OF_DAY);

	// else start of next day
	return (dtDay.m_dt + 1.0);
}

COleDateTime CDateHelper::GetNearestYear(const COleDateTime& date, BOOL bEnd)
{
	COleDateTime dtYear;

	int nYear  = date.GetYear();
	int nMonth = date.GetMonth();

	if (nMonth > 6)
	{
		// beginning of next year
		dtYear.SetDate(nYear+1, 1, 1);
	}
	else
	{
		// beginning of this year
		dtYear.SetDate(nYear, 1, 1);
	}

	ASSERT(IsDateSet(dtYear));

	// handle end - last second of day before
	if (bEnd)
		dtYear = GetEndOfPreviousDay(dtYear);

	return dtYear;
}

COleDateTime CDateHelper::GetNearestHalfYear(const COleDateTime& date, BOOL bEnd)
{
	COleDateTime dtHalfYear;

	int nYear  = date.GetYear();
	int nMonth = date.GetMonth();

	if (nMonth > 9)
	{
		// beginning of next year
		dtHalfYear.SetDate(nYear+1, 1, 1);

	}
	else if (nMonth > 3)
	{
		// beginning of july
		dtHalfYear.SetDate(nYear, 7, 1);
	}
	else
	{
		// beginning of this year
		dtHalfYear.SetDate(nYear, 1, 1);
	}

	ASSERT(IsDateSet(dtHalfYear));

	// handle end - last second of day before
	if (bEnd)
		dtHalfYear = GetEndOfPreviousDay(dtHalfYear);

	return dtHalfYear;
}

COleDateTime CDateHelper::GetNearestQuarter(const COleDateTime& date, BOOL bEnd)
{
	COleDateTime dtQuarter;

	int nYear  = date.GetYear();

	if (date > COleDateTime(nYear, 11, 15, 0, 0, 0))
	{
		// beginning of next year
		dtQuarter.SetDate(nYear+1, 1, 1);

	}
	else if (date > COleDateTime(nYear, 8, 15, 0, 0, 0))
	{
		// beginning of october
		dtQuarter.SetDate(nYear, 10, 1);
	}
	else if (date > COleDateTime(nYear, 5, 15, 0, 0, 0))
	{
		// beginning of july
		dtQuarter.SetDate(nYear, 7, 1);
	}
	else if (date > COleDateTime(nYear, 2, 14, 0, 0, 0))
	{
		// beginning of april
		dtQuarter.SetDate(nYear, 4, 1);
	}
	else
	{
		// beginning of this year
		dtQuarter.SetDate(nYear, 1, 1);
	}

	ASSERT(IsDateSet(dtQuarter));

	// handle end - last second of day before
	if (bEnd)
		dtQuarter = GetEndOfPreviousDay(dtQuarter);

	return dtQuarter;
}

void CDateHelper::IncrementMonth(SYSTEMTIME& st, int nBy)
{
	// convert month/year to int
	int nMonth = st.wMonth;
	int nYear = st.wYear;

	IncrementMonth(nMonth, nYear, nBy);

	// convert back
	st.wMonth = (WORD)nMonth;
	st.wYear = (WORD)nYear;

}

void CDateHelper::IncrementMonth(int& nMonth, int& nYear, int nBy)
{
	BOOL bNext = (nBy >= 0);
	nBy = abs(nBy);

	while (nBy--)
		GetNextMonth(nMonth, nYear, bNext);
}

void CDateHelper::GetNextMonth(int& nMonth, int& nYear, BOOL bNext)
{
	if (bNext)
	{
		if (nMonth == 12)
		{
			nYear++;
			nMonth = 1;
		}
		else
		{
			nMonth++;
		}
	}
	else // previous
	{
		if (nMonth == 1)
		{
			nYear--;
			nMonth = 12;
		}
		else
		{
			nMonth--;
		}
	}
}

COleDateTime CDateHelper::GetNearestMonth(const COleDateTime& date, BOOL bEnd)
{
	COleDateTime dtMonth;

	int nYear  = date.GetYear();
	int nMonth = date.GetMonth();
	int nDay   = date.GetDay();

	if (nDay > 15)
	{
		// start of next month
		GetNextMonth(nMonth, nYear);
	}
	// else start of this month

	dtMonth.SetDate(nYear, nMonth, 1);

	// handle end - last second of day before
	if (bEnd)
		dtMonth = GetEndOfPreviousDay(dtMonth);

	return dtMonth;
}

COleDateTime CDateHelper::GetNearestWeek(const COleDateTime& date, BOOL bEnd)
{
	COleDateTime dtWeek = GetDateOnly(date);

	// work forward until the week changes
	int nWeek = CalcWeekofYear(date);

	do 
	{
		dtWeek.m_dt += 1.0;
	}
	while (CalcWeekofYear(dtWeek) == nWeek);

	// if the number of days added >= 4 then subtract a week
	if (CalcDaysFromTo(date, dtWeek, TRUE) >= 4)
		dtWeek.m_dt -= 7.0;

	// handle end - last second of day before
	if (bEnd)
		dtWeek = GetEndOfPreviousDay(dtWeek);

	return dtWeek;
}

////////////////////////////////////////////////////////////////////////////////
// Original code Copyright (c) Robert Walker, support@tunesmithy.co.uk 

#define SECS_TO_FT_MULT 10000000
#define TIME_T_BASE ((time64_t)11644473600)

void CDateHelper::T64ToFileTime(time64_t *pt, FILETIME *pft)
{
	ASSERT(pft && pt);

	LARGE_INTEGER li = { 0 };  
	
	li.QuadPart = ((*pt) * SECS_TO_FT_MULT);

	pft->dwLowDateTime = li.LowPart;
	pft->dwHighDateTime = li.HighPart;
}

void CDateHelper::FileTimeToT64(FILETIME *pft, time64_t *pt)
{
	ASSERT(pft && pt);

	LARGE_INTEGER li = { 0 }; 
	
	li.LowPart = pft->dwLowDateTime;
	li.HighPart = pft->dwHighDateTime;

	(*pt) = (li.QuadPart / SECS_TO_FT_MULT);
}

void CDateHelper::SystemTimeToT64(SYSTEMTIME *pst, time64_t *pt)
{
	ASSERT(pst && pt);

	FILETIME ft = { 0 }, ftLocal = { 0 };

	SystemTimeToFileTime(pst, &ftLocal);
	LocalFileTimeToFileTime(&ftLocal, &ft);
	FileTimeToT64(&ft, pt);

	(*pt) -= TIME_T_BASE;
}

void CDateHelper::T64ToSystemTime(time64_t *pt, SYSTEMTIME *pst)
{
	ASSERT(pst && pt);

	FILETIME ft = { 0 }, ftLocal = { 0 };
	time64_t t = ((*pt) + TIME_T_BASE);

	T64ToFileTime(&t, &ft);
	FileTimeToLocalFileTime(&ft, &ftLocal);
	FileTimeToSystemTime(&ftLocal, pst);
}

// Original code Copyright (c) Robert Walker, support@tunesmithy.co.uk
////////////////////////////////////////////////////////////////////////////////
