//=--------------------------------------------------------------------------=
// ADBHELP.CPP:	Implementation of Data Access helpers
//=--------------------------------------------------------------------------=
// Copyright  1997  Microsoft Corporation.  All Rights Reserved.
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF 
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A 
// PARTICULAR PURPOSE.
//=--------------------------------------------------------------------------=
//
#include "stdafx.h"
#include "adbhelp.h"
#include "strutils.h"
#include <adbmacrs.h>

/////////////////////////////////////////////////////////////////////////////
// Helpers
//

// Clear a DBID by properly deallocating its pointers
//
void ClearID(DBID& dbid)
{
	// Free name portion
	//
	if (DBKIND_NAME == dbid.eKind || DBKIND_GUID_NAME == dbid.eKind || DBKIND_PGUID_NAME == dbid.eKind)
	{
		if (NULL != dbid.uName.pwszName)
		{
			CoTaskMemFree((void*)dbid.uName.pwszName);
			dbid.uName.pwszName = NULL;
		}
	}

	// Free PGUID portion
	//
	if (DBKIND_PGUID_PROPID == dbid.eKind || DBKIND_PGUID_NAME == dbid.eKind)
	{
		if (NULL != dbid.uGuid.pguid)
		{
			CoTaskMemFree((void*)dbid.uGuid.pguid);
			dbid.uGuid.pguid = NULL;
		}
	}
}

// Copy DBIDs
//
void CopyID(DBID& dbidDst, const DBID& dbidSrc)
{
	if (&dbidDst == &dbidSrc) return;
	
	ClearID(dbidDst);
	dbidDst.eKind = dbidSrc.eKind;

	// Copy name portion
	//
	if (DBKIND_PROPID == dbidSrc.eKind || DBKIND_GUID_PROPID == dbidSrc.eKind || DBKIND_PGUID_PROPID == dbidSrc.eKind)
		dbidDst.uName.ulPropid = dbidSrc.uName.ulPropid;
	else if (DBKIND_NAME == dbidSrc.eKind || DBKIND_GUID_NAME == dbidSrc.eKind || DBKIND_PGUID_NAME == dbidSrc.eKind)
	{
		if (NULL != dbidSrc.uName.pwszName)
		{
			dbidDst.uName.pwszName = (LPOLESTR)CoTaskMemAlloc((wstrlen(dbidSrc.uName.pwszName)+1) * sizeof dbidSrc.uName.pwszName[0]);
			if (NULL != dbidDst.uName.pwszName)
				wstrcpy(dbidDst.uName.pwszName, dbidSrc.uName.pwszName);
		}
		else
			dbidDst.uName.pwszName = NULL;
	}

	// Copy GUID portion
	//
	if (DBKIND_GUID == dbidSrc.eKind || DBKIND_GUID_PROPID == dbidSrc.eKind || DBKIND_GUID_NAME == dbidSrc.eKind)
		dbidDst.uGuid.guid = dbidSrc.uGuid.guid;
	else if (DBKIND_PGUID_PROPID == dbidSrc.eKind || DBKIND_PGUID_NAME == dbidSrc.eKind)
	{
		if (NULL != dbidSrc.uGuid.pguid)
		{
			dbidDst.uGuid.pguid = (GUID*)CoTaskMemAlloc(sizeof (GUID));
			if (NULL != dbidDst.uGuid.pguid)
				dbidDst.uGuid.pguid[0] = dbidSrc.uGuid.pguid[0];
		}
		else
			dbidDst.uGuid.pguid = NULL;
	}
}

// Compare two DBIDs for equality
//
BOOL IsSameIDs(const DBID& dbid1, const DBID& dbid2)
{
	// Equal if addresses are the same
	//
	if (&dbid1 == &dbid2)
		return TRUE;

	const GUID	*pguid1	= dbid1.uGuid.pguid;
	const GUID	*pguid2	= dbid2.uGuid.pguid;
	DBKIND	eKind1	= dbid1.eKind;
	DBKIND	eKind2	= dbid2.eKind;

	// Since PGUIDs can be compared to GUIDs, we want to simplify
	// the comparison code by converting GUIDs to PGUIDs
	//
	if (DBKIND_GUID_NAME == eKind1 || DBKIND_GUID_PROPID == eKind1)
	{
		eKind1 = DBKIND_GUID_NAME == eKind1 ? DBKIND_PGUID_NAME : DBKIND_PGUID_PROPID;
		pguid1 = &dbid1.uGuid.guid;
	}
	if (DBKIND_GUID_NAME == eKind2 || DBKIND_GUID_PROPID == eKind2)
	{
		eKind2 = DBKIND_GUID_NAME == eKind2 ? DBKIND_PGUID_NAME : DBKIND_PGUID_PROPID;
		pguid2 = &dbid2.uGuid.guid;
	}

	// Must be the same kind
	//
	if (eKind1 == eKind2)
	{
		// Compare name portion
		//
		if (DBKIND_PROPID == eKind1 || DBKIND_PGUID_PROPID == eKind1)
		{
			if (dbid2.uName.ulPropid != dbid1.uName.ulPropid)
				return FALSE;
		}
		else if (DBKIND_NAME == eKind1 || DBKIND_PGUID_NAME == eKind1)
		{
			if (0 != wstrcmp(dbid1.uName.pwszName, dbid2.uName.pwszName))
				return FALSE;
		}

		// Compare GUID portion
		//
		if (DBKIND_GUID == eKind1)
		{
			if (!IS_SAME_GUIDS(dbid2.uGuid.guid, dbid1.uGuid.guid))
				return FALSE;
		}
		else if (DBKIND_PGUID_PROPID == eKind1 || DBKIND_PGUID_NAME == eKind1)
		{
			if (pguid1 != pguid2 && !IS_SAME_GUIDS(*pguid1, *pguid2))
				return FALSE;
		}
		return TRUE;
	}
	return FALSE;
}

// Test for a bookmark column
//
BOOL IsBmkColumn(const DBCOLUMNINFO& col)
{
	return 0 != (col.dwFlags & DBCOLUMNFLAGS_ISBOOKMARK);
}

// Intializes a DBBINDING structure for accessing rows
//
void InitBinding(DBBINDING& dbBinding, ULONG iOrdinal, 
 DBTYPE wType, ULONG cbMaxLen, ULONG obValue, 
 ULONG obLength /*= UNUSED_OFFSET*/, ULONG obStatus /*= UNUSED_OFFSET*/, 
 BYTE bPrecision /*=0*/, BYTE bScale /*=0*/)
{
	ZeroMemory((void*)&dbBinding, sizeof dbBinding);

	// Minimum needed to get column value
	//
	dbBinding.iOrdinal	= iOrdinal;
	dbBinding.obValue	= obValue;
	dbBinding.wType		= wType;
	dbBinding.cbMaxLen	= cbMaxLen;
	dbBinding.dwPart	= DBPART_VALUE;
	
	// Optional arguments
	//
	if (UNUSED_OFFSET != obStatus)
	{
		dbBinding.dwPart |= DBPART_STATUS;
		dbBinding.obStatus = obStatus;
	}
	if (UNUSED_OFFSET != obLength)
	{
		dbBinding.dwPart |= DBPART_LENGTH;
		dbBinding.obLength = obLength;
	}
	// Only for DBTYPE_NUMERIC or DBTYPE_DECIMAL
	//
	if (DBTYPE_NUMERIC == wType || DBTYPE_DECIMAL == wType)
	{
		dbBinding.bPrecision = bPrecision;
		dbBinding.bScale = bScale;
	}
}

/////////////////////////////////////////////////////////////////////////////
// COLUMNINFO 
//
COLUMNINFO::COLUMNINFO()
{
	ZeroMemory((void*)(DBCOLUMNINFO*)this, sizeof(DBCOLUMNINFO));
	VariantInit(&varDefault);
}

COLUMNINFO::~COLUMNINFO()
{
	Clear(FALSE);
}

void COLUMNINFO::Clear(BOOL fZeroMemory /*=TRUE*/)
{
	VariantClear(&varDefault);

	if (pwszName)
		CoTaskMemFree((void*)pwszName);

	if (pTypeInfo)
		pTypeInfo->Release();

	ClearID(columnid);

	if (fZeroMemory)
		ZeroMemory((void*)this, sizeof *this);
}

COLUMNINFO& COLUMNINFO::operator=(const COLUMNINFO& other)
{
	if (this != &other)
	{
		*this = *((DBCOLUMNINFO*)&other);
		VariantCopy(&varDefault, (VARIANT*)&other.varDefault);
	}
	return *this;
}

COLUMNINFO& COLUMNINFO::operator=(const DBCOLUMNINFO& other)
{
	if ((DBCOLUMNINFO*)this != &other)
	{
		Clear();

		if (other.pwszName)
		{
			pwszName = (WCHAR*)CoTaskMemAlloc((wstrlen(other.pwszName)+1) * sizeof pwszName[0]);
			if (pwszName)
				wstrcpy(pwszName, other.pwszName);
		}
		if (other.pTypeInfo)
		{
			pTypeInfo = other.pTypeInfo;
			pTypeInfo->AddRef();
		}
		CopyID(columnid, other.columnid);
		dwFlags = other.dwFlags;
		iOrdinal = other.iOrdinal;
		ulColumnSize = other.ulColumnSize;
		wType = other.wType;
		bPrecision = other.bPrecision;
		bScale = other.bScale;
	}
	return *this;
}

/////////////////////////////////////////////////////////////////////////////
// CBindingsArray
//
HRESULT CBindingsArray::AddBinding(ULONG iOrdinal, 
 DBTYPE wType, ULONG cbMaxLen, ULONG obValue, 
 ULONG obLength /*= UNUSED_OFFSET*/, ULONG obStatus /*= UNUSED_OFFSET*/, 
 BYTE bPrecision /*=0*/, BYTE bScale /*=0*/)
{
	DBBINDING dbBinding;

	InitBinding(dbBinding, iOrdinal, wType, obValue, cbMaxLen, 
	 obStatus, obLength, bPrecision, bScale);
	return AddBinding(dbBinding);
}

