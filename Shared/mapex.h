// mapex.h: interface for the CMapEx class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_MAPEX_H__44E4FC2A_83C2_49EE_A784_4D1584CD5339__INCLUDED_)
#define AFX_MAPEX_H__44E4FC2A_83C2_49EE_A784_4D1584CD5339__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <afxtempl.h>

template <class T>
class CMapStringToContainer : protected CMap<CString, LPCTSTR, T*, T*&>
{
public:
	CMapStringToContainer()
	{
	}

	CMapStringToContainer(UINT hashSize)
	{
		if (hashSize)
			InitHashTable(hashSize);
	}
	
	virtual ~CMapStringToContainer()
	{
		RemoveAll();
	}

	const T* GetMapping(const CString& str) const
	{
		T* pMapping = NULL;

		if (!Lookup(str, pMapping))
			return NULL;

		ASSERT(pMapping);
		return pMapping;
	}

	T* GetAddMapping(const CString& str)
	{
		T* pMapping = NULL;
		
		if (Lookup(str, pMapping))
		{
			ASSERT(pMapping);
			return pMapping;
		}

		pMapping = new T;
		SetAt(str, pMapping);

		return pMapping;
	}

	BOOL IsEmpty() const
	{
		return (!GetCount());
	}

	int GetCount() const
	{
		return CMap<CString, LPCTSTR, T*, T*&>::GetCount();
	}

	void RemoveAll()
	{
		CString str;
		T* pMapping = NULL;

		POSITION pos = GetStartPosition();

		while (pos)
		{
			GetNextAssoc(pos, str, pMapping);

			ASSERT(pMapping);
			ASSERT(!str.IsEmpty());

			delete pMapping;
		}

		CMap<CString, LPCTSTR, T*, T*&>::RemoveAll();
	}
	
	void RemoveKey(const CString& str)
	{
		T* pMapping = NULL;
		
		POSITION pos = GetStartPosition();
		
		if (Lookup(str, pMapping))
		{
			ASSERT(pMapping);
			delete pMapping;

			CMap<CString, LPCTSTR, T*, T*&>::RemoveKey(str);
		}
	}

	// overload new/delete because we've hidden the base class
	void* operator new (size_t size)
	{
		return CMap<CString, LPCTSTR, T*, T*&>::operator new(size);
	}

	void operator delete (void *p)
	{
		CMap<CString, LPCTSTR, T*, T*&>::operator delete(p);
	}

};

typedef CMapStringToContainer<CStringArray> CMapStringToStringArray;

typedef CMapStringToContainer<CMapStringToStringArray> CMapStringToStringArrayMap;

typedef CMapStringToContainer<CMapStringToString> CMapStringToStringMap;

#endif // !defined(AFX_MAPEX_H__44E4FC2A_83C2_49EE_A784_4D1584CD5339__INCLUDED_)
