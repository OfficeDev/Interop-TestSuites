

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Tue Aug 18 11:59:11 2015
 */
/* Compiler settings for nspi.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __nspi_h__
#define __nspi_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

/* header files for imported files */
#include "OAIdl.h"
#include "OCIdl.h"

#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_nspi_0000_0000 */
/* [local] */ 

typedef long NTSTATUS;

typedef unsigned long DWORD;



extern RPC_IF_HANDLE __MIDL_itf_nspi_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_nspi_0000_0000_v0_0_s_ifspec;

#ifndef __nspi_INTERFACE_DEFINED__
#define __nspi_INTERFACE_DEFINED__

/* interface nspi */
/* [version][uuid] */ 

typedef /* [public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public][public] */ struct __MIDL_nspi_0001
    {
    BYTE ab[ 16 ];
    } 	FlatUID_r;

typedef struct PropertyTagArray_r
    {
    DWORD cValues;
    /* [length_is][size_is][range] */ DWORD aulPropTag[ 1 ];
    } 	PropertyTagArray_r;

typedef struct Binary_r
    {
    /* [range] */ DWORD cb;
    /* [size_is] */ BYTE *lpb;
    } 	Binary_r;

typedef struct ShortArray_r
    {
    /* [range] */ DWORD cValues;
    /* [size_is] */ short *lpi;
    } 	ShortArray_r;

typedef struct _LongArray_r
    {
    /* [range] */ DWORD cValues;
    /* [size_is] */ long *lpl;
    } 	LongArray_r;

typedef struct _StringArray_r
    {
    /* [range] */ DWORD cValues;
    /* [string][size_is] */ unsigned char **lppszA;
    } 	StringArray_r;

typedef struct _BinaryArray_r
    {
    /* [range] */ DWORD cValues;
    /* [size_is] */ Binary_r *lpbin;
    } 	BinaryArray_r;

typedef struct _FlatUIDArray_r
    {
    /* [range] */ DWORD cValues;
    /* [size_is] */ FlatUID_r **lpguid;
    } 	FlatUIDArray_r;

typedef struct _WStringArray_r
    {
    /* [range] */ DWORD cValues;
    /* [string][size_is] */ wchar_t **lppszW;
    } 	WStringArray_r;

typedef struct _DateTimeArray_r
    {
    /* [range] */ DWORD cValues;
    /* [size_is] */ FILETIME *lpft;
    } 	DateTimeArray_r;

typedef struct _PropertyValue_r PropertyValue_r;

typedef struct _PropertyRow_r
    {
    DWORD Reserved;
    /* [range] */ DWORD cValues;
    /* [size_is] */ PropertyValue_r *lpProps;
    } 	PropertyRow_r;

typedef struct _PropertyRowSet_r
    {
    /* [range] */ DWORD cRows;
    /* [size_is] */ PropertyRow_r aRow[ 1 ];
    } 	PropertyRowSet_r;

typedef struct _Restriction_r Restriction_r;

typedef struct _AndOrRestriction_r
    {
    /* [range] */ DWORD cRes;
    /* [size_is] */ Restriction_r *lpRes;
    } 	AndRestriction_r;

typedef struct _AndOrRestriction_r OrRestriction_r;

typedef struct _NotRestriction_r
    {
    Restriction_r *lpRes;
    } 	NotRestriction_r;

typedef struct _ContentRestriction_r
    {
    DWORD ulFuzzyLevel;
    DWORD ulPropTag;
    PropertyValue_r *lpProp;
    } 	ContentRestriction_r;

typedef struct _BitMaskRestriction_r
    {
    DWORD relBMR;
    DWORD ulPropTag;
    DWORD ulMask;
    } 	BitMaskRestriction_r;

typedef struct _PropertyRestriction_r
    {
    DWORD relop;
    DWORD ulPropTag;
    PropertyValue_r *lpProp;
    } 	PropertyRestriction_r;

typedef struct _ComparePropsRestriction_r
    {
    DWORD relop;
    DWORD ulPropTag1;
    DWORD ulPropTag2;
    } 	ComparePropsRestriction_r;

typedef struct _SubRestriction_r
    {
    DWORD ulSubObject;
    Restriction_r *lpRes;
    } 	SubRestriction_r;

typedef struct _SizeRestriction_r
    {
    DWORD relop;
    DWORD ulPropTag;
    DWORD cb;
    } 	SizeRestriction_r;

typedef struct _ExistRestriction_r
    {
    DWORD ulReserved1;
    DWORD ulPropTag;
    DWORD ulReserved2;
    } 	ExistRestriction_r;

typedef /* [switch_type] */ union _RestrictionUnion_r
    {
    /* [case()] */ AndRestriction_r resAnd;
    /* [case()] */ OrRestriction_r resOr;
    /* [case()] */ NotRestriction_r resNot;
    /* [case()] */ ContentRestriction_r resContent;
    /* [case()] */ PropertyRestriction_r resProperty;
    /* [case()] */ ComparePropsRestriction_r resCompareProps;
    /* [case()] */ BitMaskRestriction_r resBitMask;
    /* [case()] */ SizeRestriction_r resSize;
    /* [case()] */ ExistRestriction_r resExist;
    /* [case()] */ SubRestriction_r resSubRestriction;
    } 	RestrictionUnion_r;

struct _Restriction_r
    {
    DWORD rt;
    /* [switch_is] */ RestrictionUnion_r res;
    } ;
typedef struct PropertyName_r
    {
    FlatUID_r *lpguid;
    DWORD ulReserved;
    long lID;
    } 	PropertyName_r;

typedef struct PropertyNameSet_r
    {
    /* [range] */ DWORD cNames;
    /* [size_is] */ PropertyName_r aNames[ 1 ];
    } 	PropertyNameSet_r;

typedef struct _StringsArray
    {
    /* [range] */ DWORD Count;
    /* [string][size_is] */ unsigned char *Strings[ 1 ];
    } 	StringsArray_r;

typedef struct _WStringsArray
    {
    /* [range] */ DWORD Count;
    /* [string][size_is] */ wchar_t *Strings[ 1 ];
    } 	WStringsArray_r;

typedef struct _STAT
    {
    DWORD SortType;
    DWORD ContainerID;
    DWORD CurrentRec;
    long Delta;
    DWORD NumPos;
    DWORD TotalRecs;
    DWORD CodePage;
    DWORD TemplateLocale;
    DWORD SortLocale;
    } 	STAT;

typedef /* [switch_type] */ union _PV_r
    {
    /* [case()] */ short i;
    /* [case()] */ long l;
    /* [case()] */ unsigned short b;
    /* [case()][string] */ unsigned char *lpszA;
    /* [case()] */ Binary_r bin;
    /* [case()][string] */ wchar_t *lpszW;
    /* [case()] */ FlatUID_r *lpguid;
    /* [case()] */ FILETIME ft;
    /* [case()] */ long err;
    /* [case()] */ ShortArray_r MVi;
    /* [case()] */ LongArray_r MVl;
    /* [case()] */ StringArray_r MVszA;
    /* [case()] */ BinaryArray_r MVbin;
    /* [case()] */ FlatUIDArray_r MVguid;
    /* [case()] */ WStringArray_r MVszW;
    /* [case()] */ DateTimeArray_r MVft;
    /* [case()] */ long lReserved;
    } 	PROP_VAL_UNION;

struct _PropertyValue_r
    {
    DWORD ulPropTag;
    DWORD ulReserved;
    /* [switch_is] */ PROP_VAL_UNION Value;
    } ;
typedef /* [context_handle] */ void *NSPI_HANDLE;

long NspiBind( 
    /* [in] */ handle_t hRpc,
    /* [in] */ DWORD dwFlags,
    /* [in] */ STAT *pStat,
    /* [unique][out][in] */ FlatUID_r *pServerGuid,
    /* [ref][out] */ NSPI_HANDLE *contextHandle);

DWORD NspiUnbind( 
    /* [out][in] */ NSPI_HANDLE *contextHandle,
    /* [in] */ DWORD Reserved);

long NspiUpdateStat( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [out][in] */ STAT *pStat,
    /* [unique][out][in] */ long *plDelta);

long NspiQueryRows( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD dwFlags,
    /* [out][in] */ STAT *pStat,
    /* [range][in] */ DWORD dwETableCount,
    /* [size_is][unique][in] */ DWORD *lpETable,
    /* [in] */ DWORD Count,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [out] */ PropertyRowSet_r **ppRows);

long NspiSeekEntries( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [out][in] */ STAT *pStat,
    /* [in] */ PropertyValue_r *pTarget,
    /* [unique][in] */ PropertyTagArray_r *lpETable,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [out] */ PropertyRowSet_r **ppRows);

long NspiGetMatches( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved1,
    /* [out][in] */ STAT *pStat,
    /* [unique][in] */ PropertyTagArray_r *pReserved,
    /* [in] */ DWORD Reserved2,
    /* [unique][in] */ Restriction_r *Filter,
    /* [unique][in] */ PropertyName_r *lpPropName,
    /* [in] */ DWORD ulRequested,
    /* [out] */ PropertyTagArray_r **ppMIds,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [out] */ PropertyRowSet_r **ppRows);

long NspiResortRestriction( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [out][in] */ STAT *pStat,
    /* [in] */ PropertyTagArray_r *pInMIds,
    /* [out][in] */ PropertyTagArray_r **ppOutMIds);

long NspiDNToMId( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [in] */ StringsArray_r *pNames,
    /* [out] */ PropertyTagArray_r **ppMIds);

long NspiGetPropList( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD dwFlags,
    /* [in] */ DWORD dwMId,
    /* [in] */ DWORD CodePage,
    /* [out] */ PropertyTagArray_r **ppPropTags);

long NspiGetProps( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD dwFlags,
    /* [in] */ STAT *pStat,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [out] */ PropertyRow_r **ppRows);

long NspiCompareMIds( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [in] */ STAT *pStat,
    /* [in] */ DWORD MId1,
    /* [in] */ DWORD MId2,
    /* [out] */ long *plResult);

long NspiModProps( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [in] */ STAT *pStat,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [in] */ PropertyRow_r *pRow);

long NspiGetSpecialTable( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD dwFlags,
    /* [in] */ STAT *pStat,
    /* [out][in] */ DWORD *lpVersion,
    /* [out] */ PropertyRowSet_r **ppRows);

long NspiGetTemplateInfo( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD dwFlags,
    /* [in] */ DWORD ulType,
    /* [string][unique][in] */ unsigned char *pDN,
    /* [in] */ DWORD dwCodePage,
    /* [in] */ DWORD dwLocaleID,
    /* [out] */ PropertyRow_r **ppData);

long NspiModLinkAtt( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD dwFlags,
    /* [in] */ DWORD ulPropTag,
    /* [in] */ DWORD dwMId,
    /* [in] */ BinaryArray_r *lpEntryIds);

long Opnum15NotUsedOnWire( 
    /* [in] */ NSPI_HANDLE Reserved1,
    /* [in] */ DWORD Reserved2,
    /* [in] */ DWORD Reserved3,
    /* [in] */ BinaryArray_r *Reserved4);

long NspiQueryColumns( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [in] */ DWORD dwFlags,
    /* [out] */ PropertyTagArray_r **ppColumns);

long NspiGetNamesFromIDs( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [unique][in] */ FlatUID_r *lpguid,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [out] */ PropertyTagArray_r **ppReturnedPropTags,
    /* [out] */ PropertyNameSet_r **ppNames);

long NspiGetIDsFromNames( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [in] */ DWORD dwFlags,
    /* [range][in] */ DWORD cPropNames,
    /* [size_is][in] */ PropertyName_r **pNames,
    /* [out] */ PropertyTagArray_r **ppPropTags);

long NspiResolveNames( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [in] */ STAT *pStat,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [in] */ StringsArray_r *paStr,
    /* [out] */ PropertyTagArray_r **ppMIds,
    /* [out] */ PropertyRowSet_r **ppRows);

long NspiResolveNamesW( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [in] */ STAT *pStat,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [in] */ WStringsArray_r *paWStr,
    /* [out] */ PropertyTagArray_r **ppMIds,
    /* [out] */ PropertyRowSet_r **ppRows);



extern RPC_IF_HANDLE nspi_v56_0_c_ifspec;
extern RPC_IF_HANDLE nspi_v56_0_s_ifspec;
#endif /* __nspi_INTERFACE_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

void __RPC_USER NSPI_HANDLE_rundown( NSPI_HANDLE );

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


