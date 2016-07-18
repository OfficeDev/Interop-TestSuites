

/* this ALWAYS GENERATED file contains the RPC client stubs */


 /* File created by MIDL compiler version 7.00.0555 */
/* at Tue Aug 20 16:56:39 2013
 */
/* Compiler settings for MS-OXNSPI.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 7.00.0555 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#if !defined(_M_IA64) && !defined(_M_AMD64)


#pragma warning( disable: 4049 )  /* more than 64k source lines */
#if _MSC_VER >= 1200
#pragma warning(push)
#endif

#pragma warning( disable: 4211 )  /* redefine extern to static */
#pragma warning( disable: 4232 )  /* dllimport identity*/
#pragma warning( disable: 4024 )  /* array to pointer mapping*/
#pragma warning( disable: 4100 ) /* unreferenced arguments in x86 call */

#pragma optimize("", off ) 

#include <string.h>

#include "MS-OXNSPI.h"

#define TYPE_FORMAT_STRING_SIZE   1275                              
#define PROC_FORMAT_STRING_SIZE   1337                              
#define EXPR_FORMAT_STRING_SIZE   29                                
#define TRANSMIT_AS_TABLE_SIZE    0            
#define WIRE_MARSHAL_TABLE_SIZE   0            

typedef struct _MS2DOXNSPI_MIDL_TYPE_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ TYPE_FORMAT_STRING_SIZE ];
    } MS2DOXNSPI_MIDL_TYPE_FORMAT_STRING;

typedef struct _MS2DOXNSPI_MIDL_PROC_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ PROC_FORMAT_STRING_SIZE ];
    } MS2DOXNSPI_MIDL_PROC_FORMAT_STRING;

typedef struct _MS2DOXNSPI_MIDL_EXPR_FORMAT_STRING
    {
    long          Pad;
    unsigned char  Format[ EXPR_FORMAT_STRING_SIZE ];
    } MS2DOXNSPI_MIDL_EXPR_FORMAT_STRING;


static const RPC_SYNTAX_IDENTIFIER  _RpcTransferSyntax = 
{{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}};


extern const MS2DOXNSPI_MIDL_TYPE_FORMAT_STRING MS2DOXNSPI__MIDL_TypeFormatString;
extern const MS2DOXNSPI_MIDL_PROC_FORMAT_STRING MS2DOXNSPI__MIDL_ProcFormatString;
extern const MS2DOXNSPI_MIDL_EXPR_FORMAT_STRING MS2DOXNSPI__MIDL_ExprFormatString;

#define GENERIC_BINDING_TABLE_SIZE   0            


/* Standard interface: __MIDL_itf_MS2DOXNSPI_0000_0000, ver. 0.0,
   GUID={0x00000000,0x0000,0x0000,{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00}} */


/* Standard interface: nspi, ver. 56.0,
   GUID={0xF5CC5A18,0x4264,0x101A,{0x8C,0x59,0x08,0x00,0x2B,0x2F,0x84,0x26}} */



static const RPC_CLIENT_INTERFACE nspi___RpcClientInterface =
    {
    sizeof(RPC_CLIENT_INTERFACE),
    {{0xF5CC5A18,0x4264,0x101A,{0x8C,0x59,0x08,0x00,0x2B,0x2F,0x84,0x26}},{56,0}},
    {{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}},
    0,
    0,
    0,
    0,
    0,
    0x00000000
    };
RPC_IF_HANDLE nspi_v56_0_c_ifspec = (RPC_IF_HANDLE)& nspi___RpcClientInterface;

extern const MIDL_STUB_DESC nspi_StubDesc;

static RPC_BINDING_HANDLE nspi__MIDL_AutoBindHandle;


long NspiBind( 
    /* [in] */ handle_t hRpc,
    /* [in] */ DWORD dwFlags,
    /* [in] */ STAT *pStat,
    /* [unique][out][in] */ FlatUID_r *pServerGuid,
    /* [ref][out] */ NSPI_HANDLE *contextHandle)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[0],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


DWORD NspiUnbind( 
    /* [out][in] */ NSPI_HANDLE *contextHandle,
    /* [in] */ DWORD Reserved)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[58],
                  ( unsigned char * )&contextHandle);
    return ( DWORD  )_RetVal.Simple;
    
}


long NspiUpdateStat( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [out][in] */ STAT *pStat,
    /* [unique][out][in] */ long *plDelta)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[106],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


long NspiQueryRows( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD dwFlags,
    /* [out][in] */ STAT *pStat,
    /* [range][in] */ DWORD dwETableCount,
    /* [size_is][unique][in] */ DWORD *lpETable,
    /* [in] */ DWORD Count,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [out] */ PropertyRowSet_r **ppRows)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[166],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


long NspiSeekEntries( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [out][in] */ STAT *pStat,
    /* [in] */ PropertyValue_r *pTarget,
    /* [unique][in] */ PropertyTagArray_r *lpETable,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [out] */ PropertyRowSet_r **ppRows)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[250],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


long NspiGetMatches( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved1,
    /* [out][in] */ STAT *pStat,
    /* [unique][in] */ PropertyTagArray_r *pReserved,
    /* [in] */ DWORD Reserved2,
    /* [unique][in] */ Restriction_r *Filter,
    /* [unique][in] */ PropertyName_r *lpPropName,
    /* [in] */ DWORD ulRequested,
    /* [out] */ PropertyTagArray_r **ppOutMIds,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [out] */ PropertyRowSet_r **ppRows)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[328],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


long NspiResortRestriction( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [out][in] */ STAT *pStat,
    /* [in] */ PropertyTagArray_r *pInMIds,
    /* [out][in] */ PropertyTagArray_r **ppOutMIds)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[430],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


long NspiDNToMId( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [in] */ StringsArray_r *pNames,
    /* [out] */ PropertyTagArray_r **ppOutMIds)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[496],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


long NspiGetPropList( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD dwFlags,
    /* [in] */ DWORD dwMId,
    /* [in] */ DWORD CodePage,
    /* [out] */ PropertyTagArray_r **ppPropTags)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[556],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


long NspiGetProps( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD dwFlags,
    /* [in] */ STAT *pStat,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [out] */ PropertyRow_r **ppRows)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[622],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


long NspiCompareMIds( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [in] */ STAT *pStat,
    /* [in] */ DWORD MId1,
    /* [in] */ DWORD MId2,
    /* [out] */ long *plResult)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[688],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


long NspiModProps( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [in] */ STAT *pStat,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [in] */ PropertyRow_r *pRow)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[760],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


long NspiGetSpecialTable( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD dwFlags,
    /* [in] */ STAT *pStat,
    /* [out][in] */ DWORD *lpVersion,
    /* [out] */ PropertyRowSet_r **ppRows)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[826],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


long NspiGetTemplateInfo( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD dwFlags,
    /* [in] */ DWORD ulType,
    /* [string][unique][in] */ unsigned char *pDN,
    /* [in] */ DWORD dwCodePage,
    /* [in] */ DWORD dwLocaleID,
    /* [out] */ PropertyRow_r **ppData)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[892],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


long NspiModLinkAtt( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD dwFlags,
    /* [in] */ DWORD ulPropTag,
    /* [in] */ DWORD dwMId,
    /* [in] */ BinaryArray_r *lpEntryIds)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[970],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


void Opnum15NotUsedOnWire( 
    /* [in] */ handle_t IDL_handle)
{

    NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[1036],
                  ( unsigned char * )&IDL_handle);
    
}


long NspiQueryColumns( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [in] */ DWORD dwFlags,
    /* [out] */ PropertyTagArray_r **ppColumns)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[1064],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


void Opnum17NotUsedOnWire( 
    /* [in] */ handle_t IDL_handle)
{

    NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[1124],
                  ( unsigned char * )&IDL_handle);
    
}


void Opnum18NotUsedOnWire( 
    /* [in] */ handle_t IDL_handle)
{

    NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[1152],
                  ( unsigned char * )&IDL_handle);
    
}


long NspiResolveNames( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [in] */ STAT *pStat,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [in] */ StringsArray_r *paStr,
    /* [out] */ PropertyTagArray_r **ppMIds,
    /* [out] */ PropertyRowSet_r **ppRows)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[1180],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


long NspiResolveNamesW( 
    /* [in] */ NSPI_HANDLE hRpc,
    /* [in] */ DWORD Reserved,
    /* [in] */ STAT *pStat,
    /* [unique][in] */ PropertyTagArray_r *pPropTags,
    /* [in] */ WStringsArray_r *paWStr,
    /* [out] */ PropertyTagArray_r **ppMIds,
    /* [out] */ PropertyRowSet_r **ppRows)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&nspi_StubDesc,
                  (PFORMAT_STRING) &MS2DOXNSPI__MIDL_ProcFormatString.Format[1258],
                  ( unsigned char * )&hRpc);
    return ( long  )_RetVal.Simple;
    
}


#if !defined(__RPC_WIN32__)
#error  Invalid build platform for this stub.
#endif
#if !(TARGET_IS_NT60_OR_LATER)
#error You need Windows Vista or later to run this stub because it uses these features:
#error   forced complex structure or array, new range semantics.
#error However, your C/C++ compilation flags indicate you intend to run this app on earlier systems.
#error This app will fail with the RPC_X_WRONG_STUB_VERSION error.
#endif


static const MS2DOXNSPI_MIDL_PROC_FORMAT_STRING MS2DOXNSPI__MIDL_ProcFormatString =
    {
        0,
        {

	/* Procedure NspiBind */

			0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/*  2 */	NdrFcLong( 0x0 ),	/* 0 */
/*  6 */	NdrFcShort( 0x0 ),	/* 0 */
/*  8 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 10 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 12 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 14 */	NdrFcShort( 0x94 ),	/* 148 */
/* 16 */	NdrFcShort( 0x84 ),	/* 132 */
/* 18 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x5,		/* 5 */
/* 20 */	0x8,		/* 8 */
			0x41,		/* Ext Flags:  new corr desc, has range on conformance */
/* 22 */	NdrFcShort( 0x0 ),	/* 0 */
/* 24 */	NdrFcShort( 0x0 ),	/* 0 */
/* 26 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 28 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 30 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 32 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter dwFlags */

/* 34 */	NdrFcShort( 0x10a ),	/* Flags:  must free, in, simple ref, */
/* 36 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 38 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter pStat */

/* 40 */	NdrFcShort( 0x1a ),	/* Flags:  must free, in, out, */
/* 42 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 44 */	NdrFcShort( 0x14 ),	/* Type Offset=20 */

	/* Parameter pServerGuid */

/* 46 */	NdrFcShort( 0x110 ),	/* Flags:  out, simple ref, */
/* 48 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 50 */	NdrFcShort( 0x2c ),	/* Type Offset=44 */

	/* Parameter contextHandle */

/* 52 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 54 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 56 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiUnbind */


	/* Return value */

/* 58 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 60 */	NdrFcLong( 0x0 ),	/* 0 */
/* 64 */	NdrFcShort( 0x1 ),	/* 1 */
/* 66 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 68 */	0x30,		/* FC_BIND_CONTEXT */
			0xe0,		/* Ctxt flags:  via ptr, in, out, */
/* 70 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 72 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 74 */	NdrFcShort( 0x40 ),	/* 64 */
/* 76 */	NdrFcShort( 0x40 ),	/* 64 */
/* 78 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x3,		/* 3 */
/* 80 */	0x8,		/* 8 */
			0x41,		/* Ext Flags:  new corr desc, has range on conformance */
/* 82 */	NdrFcShort( 0x0 ),	/* 0 */
/* 84 */	NdrFcShort( 0x0 ),	/* 0 */
/* 86 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter contextHandle */

/* 88 */	NdrFcShort( 0x118 ),	/* Flags:  in, out, simple ref, */
/* 90 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 92 */	NdrFcShort( 0x34 ),	/* Type Offset=52 */

	/* Parameter Reserved */

/* 94 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 96 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 98 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Return value */

/* 100 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 102 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 104 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiUpdateStat */

/* 106 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 108 */	NdrFcLong( 0x0 ),	/* 0 */
/* 112 */	NdrFcShort( 0x2 ),	/* 2 */
/* 114 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 116 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 118 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 120 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 122 */	NdrFcShort( 0x90 ),	/* 144 */
/* 124 */	NdrFcShort( 0x6c ),	/* 108 */
/* 126 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x5,		/* 5 */
/* 128 */	0x8,		/* 8 */
			0x41,		/* Ext Flags:  new corr desc, has range on conformance */
/* 130 */	NdrFcShort( 0x0 ),	/* 0 */
/* 132 */	NdrFcShort( 0x0 ),	/* 0 */
/* 134 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 136 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 138 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 140 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter Reserved */

/* 142 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 144 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 146 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pStat */

/* 148 */	NdrFcShort( 0x11a ),	/* Flags:  must free, in, out, simple ref, */
/* 150 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 152 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter plDelta */

/* 154 */	NdrFcShort( 0x1a ),	/* Flags:  must free, in, out, */
/* 156 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 158 */	NdrFcShort( 0x3c ),	/* Type Offset=60 */

	/* Return value */

/* 160 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 162 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 164 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiQueryRows */

/* 166 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 168 */	NdrFcLong( 0x0 ),	/* 0 */
/* 172 */	NdrFcShort( 0x3 ),	/* 3 */
/* 174 */	NdrFcShort( 0x24 ),	/* x86 Stack size/offset = 36 */
/* 176 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 178 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 180 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 182 */	NdrFcShort( 0x84 ),	/* 132 */
/* 184 */	NdrFcShort( 0x50 ),	/* 80 */
/* 186 */	0x47,		/* Oi2 Flags:  srv must size, clt must size, has return, has ext, */
			0x9,		/* 9 */
/* 188 */	0x8,		/* 8 */
			0x47,		/* Ext Flags:  new corr desc, clt corr check, srv corr check, has range on conformance */
/* 190 */	NdrFcShort( 0x1 ),	/* 1 */
/* 192 */	NdrFcShort( 0x1 ),	/* 1 */
/* 194 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 196 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 198 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 200 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter dwFlags */

/* 202 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 204 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 206 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pStat */

/* 208 */	NdrFcShort( 0x11a ),	/* Flags:  must free, in, out, simple ref, */
/* 210 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 212 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter dwETableCount */

/* 214 */	NdrFcShort( 0x88 ),	/* Flags:  in, by val, */
/* 216 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 218 */	NdrFcShort( 0x40 ),	/* 64 */

	/* Parameter lpETable */

/* 220 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 222 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 224 */	NdrFcShort( 0x4a ),	/* Type Offset=74 */

	/* Parameter Count */

/* 226 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 228 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 230 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pPropTags */

/* 232 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 234 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 236 */	NdrFcShort( 0x64 ),	/* Type Offset=100 */

	/* Parameter ppRows */

/* 238 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 240 */	NdrFcShort( 0x1c ),	/* x86 Stack size/offset = 28 */
/* 242 */	NdrFcShort( 0x96 ),	/* Type Offset=150 */

	/* Return value */

/* 244 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 246 */	NdrFcShort( 0x20 ),	/* x86 Stack size/offset = 32 */
/* 248 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiSeekEntries */

/* 250 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 252 */	NdrFcLong( 0x0 ),	/* 0 */
/* 256 */	NdrFcShort( 0x4 ),	/* 4 */
/* 258 */	NdrFcShort( 0x20 ),	/* x86 Stack size/offset = 32 */
/* 260 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 262 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 264 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 266 */	NdrFcShort( 0x74 ),	/* 116 */
/* 268 */	NdrFcShort( 0x50 ),	/* 80 */
/* 270 */	0x47,		/* Oi2 Flags:  srv must size, clt must size, has return, has ext, */
			0x8,		/* 8 */
/* 272 */	0x8,		/* 8 */
			0x47,		/* Ext Flags:  new corr desc, clt corr check, srv corr check, has range on conformance */
/* 274 */	NdrFcShort( 0x1 ),	/* 1 */
/* 276 */	NdrFcShort( 0x1 ),	/* 1 */
/* 278 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 280 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 282 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 284 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter Reserved */

/* 286 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 288 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 290 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pStat */

/* 292 */	NdrFcShort( 0x11a ),	/* Flags:  must free, in, out, simple ref, */
/* 294 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 296 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter pTarget */

/* 298 */	NdrFcShort( 0x10b ),	/* Flags:  must size, must free, in, simple ref, */
/* 300 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 302 */	NdrFcShort( 0x2da ),	/* Type Offset=730 */

	/* Parameter lpETable */

/* 304 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 306 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 308 */	NdrFcShort( 0x64 ),	/* Type Offset=100 */

	/* Parameter pPropTags */

/* 310 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 312 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 314 */	NdrFcShort( 0x64 ),	/* Type Offset=100 */

	/* Parameter ppRows */

/* 316 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 318 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 320 */	NdrFcShort( 0x96 ),	/* Type Offset=150 */

	/* Return value */

/* 322 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 324 */	NdrFcShort( 0x1c ),	/* x86 Stack size/offset = 28 */
/* 326 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiGetMatches */

/* 328 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 330 */	NdrFcLong( 0x0 ),	/* 0 */
/* 334 */	NdrFcShort( 0x5 ),	/* 5 */
/* 336 */	NdrFcShort( 0x30 ),	/* x86 Stack size/offset = 48 */
/* 338 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 340 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 342 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 344 */	NdrFcShort( 0xf4 ),	/* 244 */
/* 346 */	NdrFcShort( 0x50 ),	/* 80 */
/* 348 */	0x47,		/* Oi2 Flags:  srv must size, clt must size, has return, has ext, */
			0xc,		/* 12 */
/* 350 */	0x8,		/* 8 */
			0x47,		/* Ext Flags:  new corr desc, clt corr check, srv corr check, has range on conformance */
/* 352 */	NdrFcShort( 0x1 ),	/* 1 */
/* 354 */	NdrFcShort( 0x1 ),	/* 1 */
/* 356 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 358 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 360 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 362 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter Reserved1 */

/* 364 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 366 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 368 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pStat */

/* 370 */	NdrFcShort( 0x11a ),	/* Flags:  must free, in, out, simple ref, */
/* 372 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 374 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter pReserved */

/* 376 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 378 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 380 */	NdrFcShort( 0x64 ),	/* Type Offset=100 */

	/* Parameter Reserved2 */

/* 382 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 384 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 386 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter Filter */

/* 388 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 390 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 392 */	NdrFcShort( 0x364 ),	/* Type Offset=868 */

	/* Parameter lpPropName */

/* 394 */	NdrFcShort( 0xa ),	/* Flags:  must free, in, */
/* 396 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 398 */	NdrFcShort( 0x450 ),	/* Type Offset=1104 */

	/* Parameter ulRequested */

/* 400 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 402 */	NdrFcShort( 0x1c ),	/* x86 Stack size/offset = 28 */
/* 404 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter ppOutMIds */

/* 406 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 408 */	NdrFcShort( 0x20 ),	/* x86 Stack size/offset = 32 */
/* 410 */	NdrFcShort( 0x46a ),	/* Type Offset=1130 */

	/* Parameter pPropTags */

/* 412 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 414 */	NdrFcShort( 0x24 ),	/* x86 Stack size/offset = 36 */
/* 416 */	NdrFcShort( 0x64 ),	/* Type Offset=100 */

	/* Parameter ppRows */

/* 418 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 420 */	NdrFcShort( 0x28 ),	/* x86 Stack size/offset = 40 */
/* 422 */	NdrFcShort( 0x96 ),	/* Type Offset=150 */

	/* Return value */

/* 424 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 426 */	NdrFcShort( 0x2c ),	/* x86 Stack size/offset = 44 */
/* 428 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiResortRestriction */

/* 430 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 432 */	NdrFcLong( 0x0 ),	/* 0 */
/* 436 */	NdrFcShort( 0x6 ),	/* 6 */
/* 438 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 440 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 442 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 444 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 446 */	NdrFcShort( 0x74 ),	/* 116 */
/* 448 */	NdrFcShort( 0x50 ),	/* 80 */
/* 450 */	0x47,		/* Oi2 Flags:  srv must size, clt must size, has return, has ext, */
			0x6,		/* 6 */
/* 452 */	0x8,		/* 8 */
			0x47,		/* Ext Flags:  new corr desc, clt corr check, srv corr check, has range on conformance */
/* 454 */	NdrFcShort( 0x1 ),	/* 1 */
/* 456 */	NdrFcShort( 0x1 ),	/* 1 */
/* 458 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 460 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 462 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 464 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter Reserved */

/* 466 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 468 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 470 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pStat */

/* 472 */	NdrFcShort( 0x11a ),	/* Flags:  must free, in, out, simple ref, */
/* 474 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 476 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter pInMIds */

/* 478 */	NdrFcShort( 0x10b ),	/* Flags:  must size, must free, in, simple ref, */
/* 480 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 482 */	NdrFcShort( 0x8e ),	/* Type Offset=142 */

	/* Parameter ppOutMIds */

/* 484 */	NdrFcShort( 0x201b ),	/* Flags:  must size, must free, in, out, srv alloc size=8 */
/* 486 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 488 */	NdrFcShort( 0x46a ),	/* Type Offset=1130 */

	/* Return value */

/* 490 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 492 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 494 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiDNToMId */

/* 496 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 498 */	NdrFcLong( 0x0 ),	/* 0 */
/* 502 */	NdrFcShort( 0x7 ),	/* 7 */
/* 504 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 506 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 508 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 510 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 512 */	NdrFcShort( 0x2c ),	/* 44 */
/* 514 */	NdrFcShort( 0x8 ),	/* 8 */
/* 516 */	0x47,		/* Oi2 Flags:  srv must size, clt must size, has return, has ext, */
			0x5,		/* 5 */
/* 518 */	0x8,		/* 8 */
			0x47,		/* Ext Flags:  new corr desc, clt corr check, srv corr check, has range on conformance */
/* 520 */	NdrFcShort( 0x1 ),	/* 1 */
/* 522 */	NdrFcShort( 0x1 ),	/* 1 */
/* 524 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 526 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 528 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 530 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter Reserved */

/* 532 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 534 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 536 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pNames */

/* 538 */	NdrFcShort( 0x10b ),	/* Flags:  must size, must free, in, simple ref, */
/* 540 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 542 */	NdrFcShort( 0x4a0 ),	/* Type Offset=1184 */

	/* Parameter ppOutMIds */

/* 544 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 546 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 548 */	NdrFcShort( 0x46a ),	/* Type Offset=1130 */

	/* Return value */

/* 550 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 552 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 554 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiGetPropList */

/* 556 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 558 */	NdrFcLong( 0x0 ),	/* 0 */
/* 562 */	NdrFcShort( 0x8 ),	/* 8 */
/* 564 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 566 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 568 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 570 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 572 */	NdrFcShort( 0x3c ),	/* 60 */
/* 574 */	NdrFcShort( 0x8 ),	/* 8 */
/* 576 */	0x45,		/* Oi2 Flags:  srv must size, has return, has ext, */
			0x6,		/* 6 */
/* 578 */	0x8,		/* 8 */
			0x43,		/* Ext Flags:  new corr desc, clt corr check, has range on conformance */
/* 580 */	NdrFcShort( 0x1 ),	/* 1 */
/* 582 */	NdrFcShort( 0x0 ),	/* 0 */
/* 584 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 586 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 588 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 590 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter dwFlags */

/* 592 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 594 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 596 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter dwMId */

/* 598 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 600 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 602 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter CodePage */

/* 604 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 606 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 608 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter ppPropTags */

/* 610 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 612 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 614 */	NdrFcShort( 0x46a ),	/* Type Offset=1130 */

	/* Return value */

/* 616 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 618 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 620 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiGetProps */

/* 622 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 624 */	NdrFcLong( 0x0 ),	/* 0 */
/* 628 */	NdrFcShort( 0x9 ),	/* 9 */
/* 630 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 632 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 634 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 636 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 638 */	NdrFcShort( 0x74 ),	/* 116 */
/* 640 */	NdrFcShort( 0x8 ),	/* 8 */
/* 642 */	0x47,		/* Oi2 Flags:  srv must size, clt must size, has return, has ext, */
			0x6,		/* 6 */
/* 644 */	0x8,		/* 8 */
			0x47,		/* Ext Flags:  new corr desc, clt corr check, srv corr check, has range on conformance */
/* 646 */	NdrFcShort( 0x1 ),	/* 1 */
/* 648 */	NdrFcShort( 0x1 ),	/* 1 */
/* 650 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 652 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 654 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 656 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter dwFlags */

/* 658 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 660 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 662 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pStat */

/* 664 */	NdrFcShort( 0x10a ),	/* Flags:  must free, in, simple ref, */
/* 666 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 668 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter pPropTags */

/* 670 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 672 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 674 */	NdrFcShort( 0x64 ),	/* Type Offset=100 */

	/* Parameter ppRows */

/* 676 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 678 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 680 */	NdrFcShort( 0x4aa ),	/* Type Offset=1194 */

	/* Return value */

/* 682 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 684 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 686 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiCompareMIds */

/* 688 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 690 */	NdrFcLong( 0x0 ),	/* 0 */
/* 694 */	NdrFcShort( 0xa ),	/* 10 */
/* 696 */	NdrFcShort( 0x1c ),	/* x86 Stack size/offset = 28 */
/* 698 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 700 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 702 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 704 */	NdrFcShort( 0x84 ),	/* 132 */
/* 706 */	NdrFcShort( 0x24 ),	/* 36 */
/* 708 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x7,		/* 7 */
/* 710 */	0x8,		/* 8 */
			0x41,		/* Ext Flags:  new corr desc, has range on conformance */
/* 712 */	NdrFcShort( 0x0 ),	/* 0 */
/* 714 */	NdrFcShort( 0x0 ),	/* 0 */
/* 716 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 718 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 720 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 722 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter Reserved */

/* 724 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 726 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 728 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pStat */

/* 730 */	NdrFcShort( 0x10a ),	/* Flags:  must free, in, simple ref, */
/* 732 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 734 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter MId1 */

/* 736 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 738 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 740 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter MId2 */

/* 742 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 744 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 746 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter plResult */

/* 748 */	NdrFcShort( 0x2150 ),	/* Flags:  out, base type, simple ref, srv alloc size=8 */
/* 750 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 752 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Return value */

/* 754 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 756 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 758 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiModProps */

/* 760 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 762 */	NdrFcLong( 0x0 ),	/* 0 */
/* 766 */	NdrFcShort( 0xb ),	/* 11 */
/* 768 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 770 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 772 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 774 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 776 */	NdrFcShort( 0x74 ),	/* 116 */
/* 778 */	NdrFcShort( 0x8 ),	/* 8 */
/* 780 */	0x46,		/* Oi2 Flags:  clt must size, has return, has ext, */
			0x6,		/* 6 */
/* 782 */	0x8,		/* 8 */
			0x45,		/* Ext Flags:  new corr desc, srv corr check, has range on conformance */
/* 784 */	NdrFcShort( 0x0 ),	/* 0 */
/* 786 */	NdrFcShort( 0x1 ),	/* 1 */
/* 788 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 790 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 792 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 794 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter Reserved */

/* 796 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 798 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 800 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pStat */

/* 802 */	NdrFcShort( 0x10a ),	/* Flags:  must free, in, simple ref, */
/* 804 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 806 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter pPropTags */

/* 808 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 810 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 812 */	NdrFcShort( 0x64 ),	/* Type Offset=100 */

	/* Parameter pRow */

/* 814 */	NdrFcShort( 0x10b ),	/* Flags:  must size, must free, in, simple ref, */
/* 816 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 818 */	NdrFcShort( 0x314 ),	/* Type Offset=788 */

	/* Return value */

/* 820 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 822 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 824 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiGetSpecialTable */

/* 826 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 828 */	NdrFcLong( 0x0 ),	/* 0 */
/* 832 */	NdrFcShort( 0xc ),	/* 12 */
/* 834 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 836 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 838 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 840 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 842 */	NdrFcShort( 0x90 ),	/* 144 */
/* 844 */	NdrFcShort( 0x24 ),	/* 36 */
/* 846 */	0x45,		/* Oi2 Flags:  srv must size, has return, has ext, */
			0x6,		/* 6 */
/* 848 */	0x8,		/* 8 */
			0x43,		/* Ext Flags:  new corr desc, clt corr check, has range on conformance */
/* 850 */	NdrFcShort( 0x1 ),	/* 1 */
/* 852 */	NdrFcShort( 0x0 ),	/* 0 */
/* 854 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 856 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 858 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 860 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter dwFlags */

/* 862 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 864 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 866 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pStat */

/* 868 */	NdrFcShort( 0x10a ),	/* Flags:  must free, in, simple ref, */
/* 870 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 872 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter lpVersion */

/* 874 */	NdrFcShort( 0x158 ),	/* Flags:  in, out, base type, simple ref, */
/* 876 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 878 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter ppRows */

/* 880 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 882 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 884 */	NdrFcShort( 0x96 ),	/* Type Offset=150 */

	/* Return value */

/* 886 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 888 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 890 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiGetTemplateInfo */

/* 892 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 894 */	NdrFcLong( 0x0 ),	/* 0 */
/* 898 */	NdrFcShort( 0xd ),	/* 13 */
/* 900 */	NdrFcShort( 0x20 ),	/* x86 Stack size/offset = 32 */
/* 902 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 904 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 906 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 908 */	NdrFcShort( 0x44 ),	/* 68 */
/* 910 */	NdrFcShort( 0x8 ),	/* 8 */
/* 912 */	0x47,		/* Oi2 Flags:  srv must size, clt must size, has return, has ext, */
			0x8,		/* 8 */
/* 914 */	0x8,		/* 8 */
			0x43,		/* Ext Flags:  new corr desc, clt corr check, has range on conformance */
/* 916 */	NdrFcShort( 0x1 ),	/* 1 */
/* 918 */	NdrFcShort( 0x0 ),	/* 0 */
/* 920 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 922 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 924 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 926 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter dwFlags */

/* 928 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 930 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 932 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter ulType */

/* 934 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 936 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 938 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pDN */

/* 940 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 942 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 944 */	NdrFcShort( 0x124 ),	/* Type Offset=292 */

	/* Parameter dwCodePage */

/* 946 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 948 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 950 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter dwLocaleID */

/* 952 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 954 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 956 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter ppData */

/* 958 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 960 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 962 */	NdrFcShort( 0x4aa ),	/* Type Offset=1194 */

	/* Return value */

/* 964 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 966 */	NdrFcShort( 0x1c ),	/* x86 Stack size/offset = 28 */
/* 968 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiModLinkAtt */

/* 970 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 972 */	NdrFcLong( 0x0 ),	/* 0 */
/* 976 */	NdrFcShort( 0xe ),	/* 14 */
/* 978 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 980 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 982 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 984 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 986 */	NdrFcShort( 0x3c ),	/* 60 */
/* 988 */	NdrFcShort( 0x8 ),	/* 8 */
/* 990 */	0x46,		/* Oi2 Flags:  clt must size, has return, has ext, */
			0x6,		/* 6 */
/* 992 */	0x8,		/* 8 */
			0x45,		/* Ext Flags:  new corr desc, srv corr check, has range on conformance */
/* 994 */	NdrFcShort( 0x0 ),	/* 0 */
/* 996 */	NdrFcShort( 0x1 ),	/* 1 */
/* 998 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 1000 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 1002 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 1004 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter dwFlags */

/* 1006 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 1008 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 1010 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter ulPropTag */

/* 1012 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 1014 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 1016 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter dwMId */

/* 1018 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 1020 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 1022 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter lpEntryIds */

/* 1024 */	NdrFcShort( 0x10b ),	/* Flags:  must size, must free, in, simple ref, */
/* 1026 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 1028 */	NdrFcShort( 0x21c ),	/* Type Offset=540 */

	/* Return value */

/* 1030 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 1032 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 1034 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure Opnum15NotUsedOnWire */

/* 1036 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 1038 */	NdrFcLong( 0x0 ),	/* 0 */
/* 1042 */	NdrFcShort( 0xf ),	/* 15 */
/* 1044 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 1046 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 1048 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 1050 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1052 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1054 */	0x40,		/* Oi2 Flags:  has ext, */
			0x0,		/* 0 */
/* 1056 */	0x8,		/* 8 */
			0x41,		/* Ext Flags:  new corr desc, has range on conformance */
/* 1058 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1060 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1062 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Procedure NspiQueryColumns */


	/* Parameter IDL_handle */

/* 1064 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 1066 */	NdrFcLong( 0x0 ),	/* 0 */
/* 1070 */	NdrFcShort( 0x10 ),	/* 16 */
/* 1072 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 1074 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 1076 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 1078 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 1080 */	NdrFcShort( 0x34 ),	/* 52 */
/* 1082 */	NdrFcShort( 0x8 ),	/* 8 */
/* 1084 */	0x45,		/* Oi2 Flags:  srv must size, has return, has ext, */
			0x5,		/* 5 */
/* 1086 */	0x8,		/* 8 */
			0x43,		/* Ext Flags:  new corr desc, clt corr check, has range on conformance */
/* 1088 */	NdrFcShort( 0x1 ),	/* 1 */
/* 1090 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1092 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 1094 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 1096 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 1098 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter Reserved */

/* 1100 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 1102 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 1104 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter dwFlags */

/* 1106 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 1108 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 1110 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter ppColumns */

/* 1112 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 1114 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 1116 */	NdrFcShort( 0x46a ),	/* Type Offset=1130 */

	/* Return value */

/* 1118 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 1120 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 1122 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure Opnum17NotUsedOnWire */

/* 1124 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 1126 */	NdrFcLong( 0x0 ),	/* 0 */
/* 1130 */	NdrFcShort( 0x11 ),	/* 17 */
/* 1132 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 1134 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 1136 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 1138 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1140 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1142 */	0x40,		/* Oi2 Flags:  has ext, */
			0x0,		/* 0 */
/* 1144 */	0x8,		/* 8 */
			0x41,		/* Ext Flags:  new corr desc, has range on conformance */
/* 1146 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1148 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1150 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Procedure Opnum18NotUsedOnWire */


	/* Parameter IDL_handle */

/* 1152 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 1154 */	NdrFcLong( 0x0 ),	/* 0 */
/* 1158 */	NdrFcShort( 0x12 ),	/* 18 */
/* 1160 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 1162 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 1164 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 1166 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1168 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1170 */	0x40,		/* Oi2 Flags:  has ext, */
			0x0,		/* 0 */
/* 1172 */	0x8,		/* 8 */
			0x41,		/* Ext Flags:  new corr desc, has range on conformance */
/* 1174 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1176 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1178 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Procedure NspiResolveNames */


	/* Parameter IDL_handle */

/* 1180 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 1182 */	NdrFcLong( 0x0 ),	/* 0 */
/* 1186 */	NdrFcShort( 0x13 ),	/* 19 */
/* 1188 */	NdrFcShort( 0x20 ),	/* x86 Stack size/offset = 32 */
/* 1190 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 1192 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 1194 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 1196 */	NdrFcShort( 0x74 ),	/* 116 */
/* 1198 */	NdrFcShort( 0x8 ),	/* 8 */
/* 1200 */	0x47,		/* Oi2 Flags:  srv must size, clt must size, has return, has ext, */
			0x8,		/* 8 */
/* 1202 */	0x8,		/* 8 */
			0x47,		/* Ext Flags:  new corr desc, clt corr check, srv corr check, has range on conformance */
/* 1204 */	NdrFcShort( 0x1 ),	/* 1 */
/* 1206 */	NdrFcShort( 0x1 ),	/* 1 */
/* 1208 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 1210 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 1212 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 1214 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter Reserved */

/* 1216 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 1218 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 1220 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pStat */

/* 1222 */	NdrFcShort( 0x10a ),	/* Flags:  must free, in, simple ref, */
/* 1224 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 1226 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter pPropTags */

/* 1228 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 1230 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 1232 */	NdrFcShort( 0x64 ),	/* Type Offset=100 */

	/* Parameter paStr */

/* 1234 */	NdrFcShort( 0x10b ),	/* Flags:  must size, must free, in, simple ref, */
/* 1236 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 1238 */	NdrFcShort( 0x4a0 ),	/* Type Offset=1184 */

	/* Parameter ppMIds */

/* 1240 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 1242 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 1244 */	NdrFcShort( 0x46a ),	/* Type Offset=1130 */

	/* Parameter ppRows */

/* 1246 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 1248 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 1250 */	NdrFcShort( 0x96 ),	/* Type Offset=150 */

	/* Return value */

/* 1252 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 1254 */	NdrFcShort( 0x1c ),	/* x86 Stack size/offset = 28 */
/* 1256 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure NspiResolveNamesW */

/* 1258 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 1260 */	NdrFcLong( 0x0 ),	/* 0 */
/* 1264 */	NdrFcShort( 0x14 ),	/* 20 */
/* 1266 */	NdrFcShort( 0x20 ),	/* x86 Stack size/offset = 32 */
/* 1268 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 1270 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 1272 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 1274 */	NdrFcShort( 0x74 ),	/* 116 */
/* 1276 */	NdrFcShort( 0x8 ),	/* 8 */
/* 1278 */	0x47,		/* Oi2 Flags:  srv must size, clt must size, has return, has ext, */
			0x8,		/* 8 */
/* 1280 */	0x8,		/* 8 */
			0x47,		/* Ext Flags:  new corr desc, clt corr check, srv corr check, has range on conformance */
/* 1282 */	NdrFcShort( 0x1 ),	/* 1 */
/* 1284 */	NdrFcShort( 0x1 ),	/* 1 */
/* 1286 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hRpc */

/* 1288 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 1290 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 1292 */	NdrFcShort( 0x38 ),	/* Type Offset=56 */

	/* Parameter Reserved */

/* 1294 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 1296 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 1298 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pStat */

/* 1300 */	NdrFcShort( 0x10a ),	/* Flags:  must free, in, simple ref, */
/* 1302 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 1304 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter pPropTags */

/* 1306 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 1308 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 1310 */	NdrFcShort( 0x64 ),	/* Type Offset=100 */

	/* Parameter paWStr */

/* 1312 */	NdrFcShort( 0x10b ),	/* Flags:  must size, must free, in, simple ref, */
/* 1314 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 1316 */	NdrFcShort( 0x4f0 ),	/* Type Offset=1264 */

	/* Parameter ppMIds */

/* 1318 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 1320 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 1322 */	NdrFcShort( 0x46a ),	/* Type Offset=1130 */

	/* Parameter ppRows */

/* 1324 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 1326 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 1328 */	NdrFcShort( 0x96 ),	/* Type Offset=150 */

	/* Return value */

/* 1330 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 1332 */	NdrFcShort( 0x1c ),	/* x86 Stack size/offset = 28 */
/* 1334 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

			0x0
        }
    };

static const MS2DOXNSPI_MIDL_TYPE_FORMAT_STRING MS2DOXNSPI__MIDL_TypeFormatString =
    {
        0,
        {
			NdrFcShort( 0x0 ),	/* 0 */
/*  2 */	
			0x11, 0x0,	/* FC_RP */
/*  4 */	NdrFcShort( 0x2 ),	/* Offset= 2 (6) */
/*  6 */	
			0x15,		/* FC_STRUCT */
			0x3,		/* 3 */
/*  8 */	NdrFcShort( 0x24 ),	/* 36 */
/* 10 */	0x8,		/* FC_LONG */
			0x8,		/* FC_LONG */
/* 12 */	0x8,		/* FC_LONG */
			0x8,		/* FC_LONG */
/* 14 */	0x8,		/* FC_LONG */
			0x8,		/* FC_LONG */
/* 16 */	0x8,		/* FC_LONG */
			0x8,		/* FC_LONG */
/* 18 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 20 */	
			0x12, 0x0,	/* FC_UP */
/* 22 */	NdrFcShort( 0x8 ),	/* Offset= 8 (30) */
/* 24 */	
			0x1d,		/* FC_SMFARRAY */
			0x0,		/* 0 */
/* 26 */	NdrFcShort( 0x10 ),	/* 16 */
/* 28 */	0x2,		/* FC_CHAR */
			0x5b,		/* FC_END */
/* 30 */	
			0x15,		/* FC_STRUCT */
			0x0,		/* 0 */
/* 32 */	NdrFcShort( 0x10 ),	/* 16 */
/* 34 */	0x4c,		/* FC_EMBEDDED_COMPLEX */
			0x0,		/* 0 */
/* 36 */	NdrFcShort( 0xfff4 ),	/* Offset= -12 (24) */
/* 38 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 40 */	
			0x11, 0x4,	/* FC_RP [alloced_on_stack] */
/* 42 */	NdrFcShort( 0x2 ),	/* Offset= 2 (44) */
/* 44 */	0x30,		/* FC_BIND_CONTEXT */
			0xa0,		/* Ctxt flags:  via ptr, out, */
/* 46 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 48 */	
			0x11, 0x4,	/* FC_RP [alloced_on_stack] */
/* 50 */	NdrFcShort( 0x2 ),	/* Offset= 2 (52) */
/* 52 */	0x30,		/* FC_BIND_CONTEXT */
			0xe1,		/* Ctxt flags:  via ptr, in, out, can't be null */
/* 54 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 56 */	0x30,		/* FC_BIND_CONTEXT */
			0x41,		/* Ctxt flags:  in, can't be null */
/* 58 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 60 */	
			0x12, 0x8,	/* FC_UP [simple_pointer] */
/* 62 */	0x8,		/* FC_LONG */
			0x5c,		/* FC_PAD */
/* 64 */	0xb7,		/* FC_RANGE */
			0x8,		/* 8 */
/* 66 */	NdrFcLong( 0x0 ),	/* 0 */
/* 70 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 74 */	
			0x12, 0x20,	/* FC_UP [maybenull_sizeis] */
/* 76 */	NdrFcShort( 0x2 ),	/* Offset= 2 (78) */
/* 78 */	
			0x1b,		/* FC_CARRAY */
			0x3,		/* 3 */
/* 80 */	NdrFcShort( 0x4 ),	/* 4 */
/* 82 */	0x29,		/* Corr desc:  parameter, FC_ULONG */
			0x0,		/*  */
/* 84 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 86 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 88 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 90 */	NdrFcLong( 0x0 ),	/* 0 */
/* 94 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 98 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 100 */	
			0x12, 0x0,	/* FC_UP */
/* 102 */	NdrFcShort( 0x28 ),	/* Offset= 40 (142) */
/* 104 */	
			0x1c,		/* FC_CVARRAY */
			0x3,		/* 3 */
/* 106 */	NdrFcShort( 0x4 ),	/* 4 */
/* 108 */	0x9,		/* Corr desc: FC_ULONG */
			0x57,		/* FC_ADD_1 */
/* 110 */	NdrFcShort( 0xfffc ),	/* -4 */
/* 112 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 114 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 116 */	NdrFcLong( 0x0 ),	/* 0 */
/* 120 */	NdrFcLong( 0x186a1 ),	/* 100001 */
/* 124 */	0x9,		/* Corr desc: FC_ULONG */
			0x0,		/*  */
/* 126 */	NdrFcShort( 0xfffc ),	/* -4 */
/* 128 */	NdrFcShort( 0x1 ),	/* Corr flags:  early, */
/* 130 */	0x0 , 
			0x0,		/* 0 */
/* 132 */	NdrFcLong( 0x0 ),	/* 0 */
/* 136 */	NdrFcLong( 0x0 ),	/* 0 */
/* 140 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 142 */	
			0x19,		/* FC_CVSTRUCT */
			0x3,		/* 3 */
/* 144 */	NdrFcShort( 0x4 ),	/* 4 */
/* 146 */	NdrFcShort( 0xffd6 ),	/* Offset= -42 (104) */
/* 148 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 150 */	
			0x11, 0x14,	/* FC_RP [alloced_on_stack] [pointer_deref] */
/* 152 */	NdrFcShort( 0x2 ),	/* Offset= 2 (154) */
/* 154 */	
			0x12, 0x0,	/* FC_UP */
/* 156 */	NdrFcShort( 0x2a8 ),	/* Offset= 680 (836) */
/* 158 */	
			0x2b,		/* FC_NON_ENCAPSULATED_UNION */
			0x8,		/* FC_LONG */
/* 160 */	0x0,		/* Corr desc:  field,  */
			0x5d,		/* FC_EXPR */
/* 162 */	NdrFcShort( 0x0 ),	/* 0 */
/* 164 */	NdrFcShort( 0x1 ),	/* Corr flags:  early, */
/* 166 */	0x0 , 
			0x0,		/* 0 */
/* 168 */	NdrFcLong( 0x0 ),	/* 0 */
/* 172 */	NdrFcLong( 0x0 ),	/* 0 */
/* 176 */	NdrFcShort( 0x2 ),	/* Offset= 2 (178) */
/* 178 */	NdrFcShort( 0x8 ),	/* 8 */
/* 180 */	NdrFcShort( 0x12 ),	/* 18 */
/* 182 */	NdrFcLong( 0x2 ),	/* 2 */
/* 186 */	NdrFcShort( 0x8006 ),	/* Simple arm type: FC_SHORT */
/* 188 */	NdrFcLong( 0x3 ),	/* 3 */
/* 192 */	NdrFcShort( 0x8008 ),	/* Simple arm type: FC_LONG */
/* 194 */	NdrFcLong( 0xb ),	/* 11 */
/* 198 */	NdrFcShort( 0x8006 ),	/* Simple arm type: FC_SHORT */
/* 200 */	NdrFcLong( 0x1e ),	/* 30 */
/* 204 */	NdrFcShort( 0x58 ),	/* Offset= 88 (292) */
/* 206 */	NdrFcLong( 0x102 ),	/* 258 */
/* 210 */	NdrFcShort( 0x6c ),	/* Offset= 108 (318) */
/* 212 */	NdrFcLong( 0x1f ),	/* 31 */
/* 216 */	NdrFcShort( 0x7a ),	/* Offset= 122 (338) */
/* 218 */	NdrFcLong( 0x48 ),	/* 72 */
/* 222 */	NdrFcShort( 0xff36 ),	/* Offset= -202 (20) */
/* 224 */	NdrFcLong( 0x40 ),	/* 64 */
/* 228 */	NdrFcShort( 0x72 ),	/* Offset= 114 (342) */
/* 230 */	NdrFcLong( 0xa ),	/* 10 */
/* 234 */	NdrFcShort( 0x8008 ),	/* Simple arm type: FC_LONG */
/* 236 */	NdrFcLong( 0x1002 ),	/* 4098 */
/* 240 */	NdrFcShort( 0x84 ),	/* Offset= 132 (372) */
/* 242 */	NdrFcLong( 0x1003 ),	/* 4099 */
/* 246 */	NdrFcShort( 0xa8 ),	/* Offset= 168 (414) */
/* 248 */	NdrFcLong( 0x101e ),	/* 4126 */
/* 252 */	NdrFcShort( 0xe0 ),	/* Offset= 224 (476) */
/* 254 */	NdrFcLong( 0x1102 ),	/* 4354 */
/* 258 */	NdrFcShort( 0x11a ),	/* Offset= 282 (540) */
/* 260 */	NdrFcLong( 0x1048 ),	/* 4168 */
/* 264 */	NdrFcShort( 0x152 ),	/* Offset= 338 (602) */
/* 266 */	NdrFcLong( 0x101f ),	/* 4127 */
/* 270 */	NdrFcShort( 0x18a ),	/* Offset= 394 (664) */
/* 272 */	NdrFcLong( 0x1040 ),	/* 4160 */
/* 276 */	NdrFcShort( 0x1b2 ),	/* Offset= 434 (710) */
/* 278 */	NdrFcLong( 0x1 ),	/* 1 */
/* 282 */	NdrFcShort( 0x8008 ),	/* Simple arm type: FC_LONG */
/* 284 */	NdrFcLong( 0xd ),	/* 13 */
/* 288 */	NdrFcShort( 0x8008 ),	/* Simple arm type: FC_LONG */
/* 290 */	NdrFcShort( 0xffff ),	/* Offset= -1 (289) */
/* 292 */	
			0x12, 0x8,	/* FC_UP [simple_pointer] */
/* 294 */	
			0x22,		/* FC_C_CSTRING */
			0x5c,		/* FC_PAD */
/* 296 */	
			0x1b,		/* FC_CARRAY */
			0x0,		/* 0 */
/* 298 */	NdrFcShort( 0x1 ),	/* 1 */
/* 300 */	0x19,		/* Corr desc:  field pointer, FC_ULONG */
			0x0,		/*  */
/* 302 */	NdrFcShort( 0x0 ),	/* 0 */
/* 304 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 306 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 308 */	NdrFcLong( 0x0 ),	/* 0 */
/* 312 */	NdrFcLong( 0x200000 ),	/* 2097152 */
/* 316 */	0x2,		/* FC_CHAR */
			0x5b,		/* FC_END */
/* 318 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 320 */	NdrFcShort( 0x8 ),	/* 8 */
/* 322 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 324 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 326 */	NdrFcShort( 0x4 ),	/* 4 */
/* 328 */	NdrFcShort( 0x4 ),	/* 4 */
/* 330 */	0x12, 0x20,	/* FC_UP [maybenull_sizeis] */
/* 332 */	NdrFcShort( 0xffdc ),	/* Offset= -36 (296) */
/* 334 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 336 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 338 */	
			0x12, 0x8,	/* FC_UP [simple_pointer] */
/* 340 */	
			0x25,		/* FC_C_WSTRING */
			0x5c,		/* FC_PAD */
/* 342 */	
			0x15,		/* FC_STRUCT */
			0x3,		/* 3 */
/* 344 */	NdrFcShort( 0x8 ),	/* 8 */
/* 346 */	0x8,		/* FC_LONG */
			0x8,		/* FC_LONG */
/* 348 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 350 */	
			0x1b,		/* FC_CARRAY */
			0x1,		/* 1 */
/* 352 */	NdrFcShort( 0x2 ),	/* 2 */
/* 354 */	0x19,		/* Corr desc:  field pointer, FC_ULONG */
			0x0,		/*  */
/* 356 */	NdrFcShort( 0x0 ),	/* 0 */
/* 358 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 360 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 362 */	NdrFcLong( 0x0 ),	/* 0 */
/* 366 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 370 */	0x6,		/* FC_SHORT */
			0x5b,		/* FC_END */
/* 372 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 374 */	NdrFcShort( 0x8 ),	/* 8 */
/* 376 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 378 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 380 */	NdrFcShort( 0x4 ),	/* 4 */
/* 382 */	NdrFcShort( 0x4 ),	/* 4 */
/* 384 */	0x12, 0x20,	/* FC_UP [maybenull_sizeis] */
/* 386 */	NdrFcShort( 0xffdc ),	/* Offset= -36 (350) */
/* 388 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 390 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 392 */	
			0x1b,		/* FC_CARRAY */
			0x3,		/* 3 */
/* 394 */	NdrFcShort( 0x4 ),	/* 4 */
/* 396 */	0x19,		/* Corr desc:  field pointer, FC_ULONG */
			0x0,		/*  */
/* 398 */	NdrFcShort( 0x0 ),	/* 0 */
/* 400 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 402 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 404 */	NdrFcLong( 0x0 ),	/* 0 */
/* 408 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 412 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 414 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 416 */	NdrFcShort( 0x8 ),	/* 8 */
/* 418 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 420 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 422 */	NdrFcShort( 0x4 ),	/* 4 */
/* 424 */	NdrFcShort( 0x4 ),	/* 4 */
/* 426 */	0x12, 0x20,	/* FC_UP [maybenull_sizeis] */
/* 428 */	NdrFcShort( 0xffdc ),	/* Offset= -36 (392) */
/* 430 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 432 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 434 */	
			0x1b,		/* FC_CARRAY */
			0x3,		/* 3 */
/* 436 */	NdrFcShort( 0x4 ),	/* 4 */
/* 438 */	0x19,		/* Corr desc:  field pointer, FC_ULONG */
			0x0,		/*  */
/* 440 */	NdrFcShort( 0x0 ),	/* 0 */
/* 442 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 444 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 446 */	NdrFcLong( 0x0 ),	/* 0 */
/* 450 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 454 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 456 */	
			0x48,		/* FC_VARIABLE_REPEAT */
			0x49,		/* FC_FIXED_OFFSET */
/* 458 */	NdrFcShort( 0x4 ),	/* 4 */
/* 460 */	NdrFcShort( 0x0 ),	/* 0 */
/* 462 */	NdrFcShort( 0x1 ),	/* 1 */
/* 464 */	NdrFcShort( 0x0 ),	/* 0 */
/* 466 */	NdrFcShort( 0x0 ),	/* 0 */
/* 468 */	0x12, 0x8,	/* FC_UP [simple_pointer] */
/* 470 */	
			0x22,		/* FC_C_CSTRING */
			0x5c,		/* FC_PAD */
/* 472 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 474 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 476 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 478 */	NdrFcShort( 0x8 ),	/* 8 */
/* 480 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 482 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 484 */	NdrFcShort( 0x4 ),	/* 4 */
/* 486 */	NdrFcShort( 0x4 ),	/* 4 */
/* 488 */	0x12, 0x20,	/* FC_UP [maybenull_sizeis] */
/* 490 */	NdrFcShort( 0xffc8 ),	/* Offset= -56 (434) */
/* 492 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 494 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 496 */	
			0x1b,		/* FC_CARRAY */
			0x3,		/* 3 */
/* 498 */	NdrFcShort( 0x8 ),	/* 8 */
/* 500 */	0x19,		/* Corr desc:  field pointer, FC_ULONG */
			0x0,		/*  */
/* 502 */	NdrFcShort( 0x0 ),	/* 0 */
/* 504 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 506 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 508 */	NdrFcLong( 0x0 ),	/* 0 */
/* 512 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 516 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 518 */	
			0x48,		/* FC_VARIABLE_REPEAT */
			0x49,		/* FC_FIXED_OFFSET */
/* 520 */	NdrFcShort( 0x8 ),	/* 8 */
/* 522 */	NdrFcShort( 0x0 ),	/* 0 */
/* 524 */	NdrFcShort( 0x1 ),	/* 1 */
/* 526 */	NdrFcShort( 0x4 ),	/* 4 */
/* 528 */	NdrFcShort( 0x4 ),	/* 4 */
/* 530 */	0x12, 0x20,	/* FC_UP [maybenull_sizeis] */
/* 532 */	NdrFcShort( 0xff14 ),	/* Offset= -236 (296) */
/* 534 */	
			0x5b,		/* FC_END */

			0x4c,		/* FC_EMBEDDED_COMPLEX */
/* 536 */	0x0,		/* 0 */
			NdrFcShort( 0xff25 ),	/* Offset= -219 (318) */
			0x5b,		/* FC_END */
/* 540 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 542 */	NdrFcShort( 0x8 ),	/* 8 */
/* 544 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 546 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 548 */	NdrFcShort( 0x4 ),	/* 4 */
/* 550 */	NdrFcShort( 0x4 ),	/* 4 */
/* 552 */	0x12, 0x20,	/* FC_UP [maybenull_sizeis] */
/* 554 */	NdrFcShort( 0xffc6 ),	/* Offset= -58 (496) */
/* 556 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 558 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 560 */	
			0x1b,		/* FC_CARRAY */
			0x3,		/* 3 */
/* 562 */	NdrFcShort( 0x4 ),	/* 4 */
/* 564 */	0x19,		/* Corr desc:  field pointer, FC_ULONG */
			0x0,		/*  */
/* 566 */	NdrFcShort( 0x0 ),	/* 0 */
/* 568 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 570 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 572 */	NdrFcLong( 0x0 ),	/* 0 */
/* 576 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 580 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 582 */	
			0x48,		/* FC_VARIABLE_REPEAT */
			0x49,		/* FC_FIXED_OFFSET */
/* 584 */	NdrFcShort( 0x4 ),	/* 4 */
/* 586 */	NdrFcShort( 0x0 ),	/* 0 */
/* 588 */	NdrFcShort( 0x1 ),	/* 1 */
/* 590 */	NdrFcShort( 0x0 ),	/* 0 */
/* 592 */	NdrFcShort( 0x0 ),	/* 0 */
/* 594 */	0x12, 0x0,	/* FC_UP */
/* 596 */	NdrFcShort( 0xfdca ),	/* Offset= -566 (30) */
/* 598 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 600 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 602 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 604 */	NdrFcShort( 0x8 ),	/* 8 */
/* 606 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 608 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 610 */	NdrFcShort( 0x4 ),	/* 4 */
/* 612 */	NdrFcShort( 0x4 ),	/* 4 */
/* 614 */	0x12, 0x20,	/* FC_UP [maybenull_sizeis] */
/* 616 */	NdrFcShort( 0xffc8 ),	/* Offset= -56 (560) */
/* 618 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 620 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 622 */	
			0x1b,		/* FC_CARRAY */
			0x3,		/* 3 */
/* 624 */	NdrFcShort( 0x4 ),	/* 4 */
/* 626 */	0x19,		/* Corr desc:  field pointer, FC_ULONG */
			0x0,		/*  */
/* 628 */	NdrFcShort( 0x0 ),	/* 0 */
/* 630 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 632 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 634 */	NdrFcLong( 0x0 ),	/* 0 */
/* 638 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 642 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 644 */	
			0x48,		/* FC_VARIABLE_REPEAT */
			0x49,		/* FC_FIXED_OFFSET */
/* 646 */	NdrFcShort( 0x4 ),	/* 4 */
/* 648 */	NdrFcShort( 0x0 ),	/* 0 */
/* 650 */	NdrFcShort( 0x1 ),	/* 1 */
/* 652 */	NdrFcShort( 0x0 ),	/* 0 */
/* 654 */	NdrFcShort( 0x0 ),	/* 0 */
/* 656 */	0x12, 0x8,	/* FC_UP [simple_pointer] */
/* 658 */	
			0x25,		/* FC_C_WSTRING */
			0x5c,		/* FC_PAD */
/* 660 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 662 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 664 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 666 */	NdrFcShort( 0x8 ),	/* 8 */
/* 668 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 670 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 672 */	NdrFcShort( 0x4 ),	/* 4 */
/* 674 */	NdrFcShort( 0x4 ),	/* 4 */
/* 676 */	0x12, 0x20,	/* FC_UP [maybenull_sizeis] */
/* 678 */	NdrFcShort( 0xffc8 ),	/* Offset= -56 (622) */
/* 680 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 682 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 684 */	
			0x1b,		/* FC_CARRAY */
			0x3,		/* 3 */
/* 686 */	NdrFcShort( 0x8 ),	/* 8 */
/* 688 */	0x19,		/* Corr desc:  field pointer, FC_ULONG */
			0x0,		/*  */
/* 690 */	NdrFcShort( 0x0 ),	/* 0 */
/* 692 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 694 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 696 */	NdrFcLong( 0x0 ),	/* 0 */
/* 700 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 704 */	0x4c,		/* FC_EMBEDDED_COMPLEX */
			0x0,		/* 0 */
/* 706 */	NdrFcShort( 0xfe94 ),	/* Offset= -364 (342) */
/* 708 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 710 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 712 */	NdrFcShort( 0x8 ),	/* 8 */
/* 714 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 716 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 718 */	NdrFcShort( 0x4 ),	/* 4 */
/* 720 */	NdrFcShort( 0x4 ),	/* 4 */
/* 722 */	0x12, 0x20,	/* FC_UP [maybenull_sizeis] */
/* 724 */	NdrFcShort( 0xffd8 ),	/* Offset= -40 (684) */
/* 726 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 728 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 730 */	
			0x1a,		/* FC_BOGUS_STRUCT */
			0x3,		/* 3 */
/* 732 */	NdrFcShort( 0x10 ),	/* 16 */
/* 734 */	NdrFcShort( 0x0 ),	/* 0 */
/* 736 */	NdrFcShort( 0x0 ),	/* Offset= 0 (736) */
/* 738 */	0x8,		/* FC_LONG */
			0x8,		/* FC_LONG */
/* 740 */	0x4c,		/* FC_EMBEDDED_COMPLEX */
			0x0,		/* 0 */
/* 742 */	NdrFcShort( 0xfdb8 ),	/* Offset= -584 (158) */
/* 744 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 746 */	
			0x21,		/* FC_BOGUS_ARRAY */
			0x3,		/* 3 */
/* 748 */	NdrFcShort( 0x0 ),	/* 0 */
/* 750 */	0x19,		/* Corr desc:  field pointer, FC_ULONG */
			0x0,		/*  */
/* 752 */	NdrFcShort( 0x4 ),	/* 4 */
/* 754 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 756 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 758 */	NdrFcLong( 0x0 ),	/* 0 */
/* 762 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 766 */	NdrFcLong( 0xffffffff ),	/* -1 */
/* 770 */	NdrFcShort( 0x0 ),	/* Corr flags:  */
/* 772 */	0x0 , 
			0x0,		/* 0 */
/* 774 */	NdrFcLong( 0x0 ),	/* 0 */
/* 778 */	NdrFcLong( 0x0 ),	/* 0 */
/* 782 */	0x4c,		/* FC_EMBEDDED_COMPLEX */
			0x0,		/* 0 */
/* 784 */	NdrFcShort( 0xffca ),	/* Offset= -54 (730) */
/* 786 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 788 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 790 */	NdrFcShort( 0xc ),	/* 12 */
/* 792 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 794 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 796 */	NdrFcShort( 0x8 ),	/* 8 */
/* 798 */	NdrFcShort( 0x8 ),	/* 8 */
/* 800 */	0x12, 0x20,	/* FC_UP [maybenull_sizeis] */
/* 802 */	NdrFcShort( 0xffc8 ),	/* Offset= -56 (746) */
/* 804 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 806 */	0x8,		/* FC_LONG */
			0x8,		/* FC_LONG */
/* 808 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 810 */	
			0x1b,		/* FC_CARRAY */
			0x3,		/* 3 */
/* 812 */	NdrFcShort( 0xc ),	/* 12 */
/* 814 */	0x9,		/* Corr desc: FC_ULONG */
			0x0,		/*  */
/* 816 */	NdrFcShort( 0xfffc ),	/* -4 */
/* 818 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 820 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 822 */	NdrFcLong( 0x0 ),	/* 0 */
/* 826 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 830 */	0x4c,		/* FC_EMBEDDED_COMPLEX */
			0x0,		/* 0 */
/* 832 */	NdrFcShort( 0xffd4 ),	/* Offset= -44 (788) */
/* 834 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 836 */	
			0x18,		/* FC_CPSTRUCT */
			0x3,		/* 3 */
/* 838 */	NdrFcShort( 0x4 ),	/* 4 */
/* 840 */	NdrFcShort( 0xffe2 ),	/* Offset= -30 (810) */
/* 842 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 844 */	
			0x48,		/* FC_VARIABLE_REPEAT */
			0x49,		/* FC_FIXED_OFFSET */
/* 846 */	NdrFcShort( 0xc ),	/* 12 */
/* 848 */	NdrFcShort( 0x4 ),	/* 4 */
/* 850 */	NdrFcShort( 0x1 ),	/* 1 */
/* 852 */	NdrFcShort( 0xc ),	/* 12 */
/* 854 */	NdrFcShort( 0xc ),	/* 12 */
/* 856 */	0x12, 0x20,	/* FC_UP [maybenull_sizeis] */
/* 858 */	NdrFcShort( 0xff90 ),	/* Offset= -112 (746) */
/* 860 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 862 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 864 */	
			0x11, 0x0,	/* FC_RP */
/* 866 */	NdrFcShort( 0xff78 ),	/* Offset= -136 (730) */
/* 868 */	
			0x12, 0x0,	/* FC_UP */
/* 870 */	NdrFcShort( 0xdc ),	/* Offset= 220 (1090) */
/* 872 */	
			0x2b,		/* FC_NON_ENCAPSULATED_UNION */
			0x8,		/* FC_LONG */
/* 874 */	0x0,		/* Corr desc:  field,  */
			0x5d,		/* FC_EXPR */
/* 876 */	NdrFcShort( 0x1 ),	/* 1 */
/* 878 */	NdrFcShort( 0x1 ),	/* Corr flags:  early, */
/* 880 */	0x0 , 
			0x0,		/* 0 */
/* 882 */	NdrFcLong( 0x0 ),	/* 0 */
/* 886 */	NdrFcLong( 0x0 ),	/* 0 */
/* 890 */	NdrFcShort( 0x2 ),	/* Offset= 2 (892) */
/* 892 */	NdrFcShort( 0xc ),	/* 12 */
/* 894 */	NdrFcShort( 0xa ),	/* 10 */
/* 896 */	NdrFcLong( 0x0 ),	/* 0 */
/* 900 */	NdrFcShort( 0x64 ),	/* Offset= 100 (1000) */
/* 902 */	NdrFcLong( 0x1 ),	/* 1 */
/* 906 */	NdrFcShort( 0x5e ),	/* Offset= 94 (1000) */
/* 908 */	NdrFcLong( 0x2 ),	/* 2 */
/* 912 */	NdrFcShort( 0x6c ),	/* Offset= 108 (1020) */
/* 914 */	NdrFcLong( 0x3 ),	/* 3 */
/* 918 */	NdrFcShort( 0x7a ),	/* Offset= 122 (1040) */
/* 920 */	NdrFcLong( 0x4 ),	/* 4 */
/* 924 */	NdrFcShort( 0x74 ),	/* Offset= 116 (1040) */
/* 926 */	NdrFcLong( 0x5 ),	/* 5 */
/* 930 */	NdrFcShort( 0x84 ),	/* Offset= 132 (1062) */
/* 932 */	NdrFcLong( 0x6 ),	/* 6 */
/* 936 */	NdrFcShort( 0x7e ),	/* Offset= 126 (1062) */
/* 938 */	NdrFcLong( 0x7 ),	/* 7 */
/* 942 */	NdrFcShort( 0x78 ),	/* Offset= 120 (1062) */
/* 944 */	NdrFcLong( 0x8 ),	/* 8 */
/* 948 */	NdrFcShort( 0x72 ),	/* Offset= 114 (1062) */
/* 950 */	NdrFcLong( 0x9 ),	/* 9 */
/* 954 */	NdrFcShort( 0x74 ),	/* Offset= 116 (1070) */
/* 956 */	NdrFcShort( 0xffff ),	/* Offset= -1 (955) */
/* 958 */	
			0x21,		/* FC_BOGUS_ARRAY */
			0x3,		/* 3 */
/* 960 */	NdrFcShort( 0x0 ),	/* 0 */
/* 962 */	0x19,		/* Corr desc:  field pointer, FC_ULONG */
			0x0,		/*  */
/* 964 */	NdrFcShort( 0x0 ),	/* 0 */
/* 966 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 968 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 970 */	NdrFcLong( 0x0 ),	/* 0 */
/* 974 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 978 */	NdrFcLong( 0xffffffff ),	/* -1 */
/* 982 */	NdrFcShort( 0x0 ),	/* Corr flags:  */
/* 984 */	0x0 , 
			0x0,		/* 0 */
/* 986 */	NdrFcLong( 0x0 ),	/* 0 */
/* 990 */	NdrFcLong( 0x0 ),	/* 0 */
/* 994 */	0x4c,		/* FC_EMBEDDED_COMPLEX */
			0x0,		/* 0 */
/* 996 */	NdrFcShort( 0x5e ),	/* Offset= 94 (1090) */
/* 998 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 1000 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 1002 */	NdrFcShort( 0x8 ),	/* 8 */
/* 1004 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 1006 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 1008 */	NdrFcShort( 0x4 ),	/* 4 */
/* 1010 */	NdrFcShort( 0x4 ),	/* 4 */
/* 1012 */	0x12, 0x20,	/* FC_UP [maybenull_sizeis] */
/* 1014 */	NdrFcShort( 0xffc8 ),	/* Offset= -56 (958) */
/* 1016 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 1018 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 1020 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 1022 */	NdrFcShort( 0x4 ),	/* 4 */
/* 1024 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 1026 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 1028 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1030 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1032 */	0x12, 0x0,	/* FC_UP */
/* 1034 */	NdrFcShort( 0x38 ),	/* Offset= 56 (1090) */
/* 1036 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 1038 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 1040 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 1042 */	NdrFcShort( 0xc ),	/* 12 */
/* 1044 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 1046 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 1048 */	NdrFcShort( 0x8 ),	/* 8 */
/* 1050 */	NdrFcShort( 0x8 ),	/* 8 */
/* 1052 */	0x12, 0x0,	/* FC_UP */
/* 1054 */	NdrFcShort( 0xfebc ),	/* Offset= -324 (730) */
/* 1056 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 1058 */	0x8,		/* FC_LONG */
			0x8,		/* FC_LONG */
/* 1060 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 1062 */	
			0x15,		/* FC_STRUCT */
			0x3,		/* 3 */
/* 1064 */	NdrFcShort( 0xc ),	/* 12 */
/* 1066 */	0x8,		/* FC_LONG */
			0x8,		/* FC_LONG */
/* 1068 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 1070 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 1072 */	NdrFcShort( 0x8 ),	/* 8 */
/* 1074 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 1076 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 1078 */	NdrFcShort( 0x4 ),	/* 4 */
/* 1080 */	NdrFcShort( 0x4 ),	/* 4 */
/* 1082 */	0x12, 0x0,	/* FC_UP */
/* 1084 */	NdrFcShort( 0x6 ),	/* Offset= 6 (1090) */
/* 1086 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 1088 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 1090 */	
			0x1a,		/* FC_BOGUS_STRUCT */
			0x3,		/* 3 */
/* 1092 */	NdrFcShort( 0x10 ),	/* 16 */
/* 1094 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1096 */	NdrFcShort( 0x0 ),	/* Offset= 0 (1096) */
/* 1098 */	0x8,		/* FC_LONG */
			0x4c,		/* FC_EMBEDDED_COMPLEX */
/* 1100 */	0x0,		/* 0 */
			NdrFcShort( 0xff1b ),	/* Offset= -229 (872) */
			0x5b,		/* FC_END */
/* 1104 */	
			0x12, 0x0,	/* FC_UP */
/* 1106 */	NdrFcShort( 0x2 ),	/* Offset= 2 (1108) */
/* 1108 */	
			0x16,		/* FC_PSTRUCT */
			0x3,		/* 3 */
/* 1110 */	NdrFcShort( 0xc ),	/* 12 */
/* 1112 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 1114 */	
			0x46,		/* FC_NO_REPEAT */
			0x5c,		/* FC_PAD */
/* 1116 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1118 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1120 */	0x12, 0x0,	/* FC_UP */
/* 1122 */	NdrFcShort( 0xfbbc ),	/* Offset= -1092 (30) */
/* 1124 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 1126 */	0x8,		/* FC_LONG */
			0x8,		/* FC_LONG */
/* 1128 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 1130 */	
			0x11, 0x14,	/* FC_RP [alloced_on_stack] [pointer_deref] */
/* 1132 */	NdrFcShort( 0xfbf8 ),	/* Offset= -1032 (100) */
/* 1134 */	
			0x11, 0x0,	/* FC_RP */
/* 1136 */	NdrFcShort( 0xfc1e ),	/* Offset= -994 (142) */
/* 1138 */	
			0x11, 0x0,	/* FC_RP */
/* 1140 */	NdrFcShort( 0x2c ),	/* Offset= 44 (1184) */
/* 1142 */	
			0x1b,		/* FC_CARRAY */
			0x3,		/* 3 */
/* 1144 */	NdrFcShort( 0x4 ),	/* 4 */
/* 1146 */	0x9,		/* Corr desc: FC_ULONG */
			0x0,		/*  */
/* 1148 */	NdrFcShort( 0xfffc ),	/* -4 */
/* 1150 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 1152 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 1154 */	NdrFcLong( 0x0 ),	/* 0 */
/* 1158 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 1162 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 1164 */	
			0x48,		/* FC_VARIABLE_REPEAT */
			0x49,		/* FC_FIXED_OFFSET */
/* 1166 */	NdrFcShort( 0x4 ),	/* 4 */
/* 1168 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1170 */	NdrFcShort( 0x1 ),	/* 1 */
/* 1172 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1174 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1176 */	0x12, 0x8,	/* FC_UP [simple_pointer] */
/* 1178 */	
			0x22,		/* FC_C_CSTRING */
			0x5c,		/* FC_PAD */
/* 1180 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 1182 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 1184 */	
			0x1a,		/* FC_BOGUS_STRUCT */
			0x3,		/* 3 */
/* 1186 */	NdrFcShort( 0x4 ),	/* 4 */
/* 1188 */	NdrFcShort( 0xffd2 ),	/* Offset= -46 (1142) */
/* 1190 */	NdrFcShort( 0x0 ),	/* Offset= 0 (1190) */
/* 1192 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 1194 */	
			0x11, 0x14,	/* FC_RP [alloced_on_stack] [pointer_deref] */
/* 1196 */	NdrFcShort( 0x2 ),	/* Offset= 2 (1198) */
/* 1198 */	
			0x12, 0x0,	/* FC_UP */
/* 1200 */	NdrFcShort( 0xfe64 ),	/* Offset= -412 (788) */
/* 1202 */	
			0x11, 0xc,	/* FC_RP [alloced_on_stack] [simple_pointer] */
/* 1204 */	0x8,		/* FC_LONG */
			0x5c,		/* FC_PAD */
/* 1206 */	
			0x11, 0x0,	/* FC_RP */
/* 1208 */	NdrFcShort( 0xfe5c ),	/* Offset= -420 (788) */
/* 1210 */	
			0x11, 0x8,	/* FC_RP [simple_pointer] */
/* 1212 */	0x8,		/* FC_LONG */
			0x5c,		/* FC_PAD */
/* 1214 */	
			0x11, 0x0,	/* FC_RP */
/* 1216 */	NdrFcShort( 0xfd5c ),	/* Offset= -676 (540) */
/* 1218 */	
			0x11, 0x0,	/* FC_RP */
/* 1220 */	NdrFcShort( 0x2c ),	/* Offset= 44 (1264) */
/* 1222 */	
			0x1b,		/* FC_CARRAY */
			0x3,		/* 3 */
/* 1224 */	NdrFcShort( 0x4 ),	/* 4 */
/* 1226 */	0x9,		/* Corr desc: FC_ULONG */
			0x0,		/*  */
/* 1228 */	NdrFcShort( 0xfffc ),	/* -4 */
/* 1230 */	NdrFcShort( 0x11 ),	/* Corr flags:  early, */
/* 1232 */	0x1 , /* correlation range */
			0x0,		/* 0 */
/* 1234 */	NdrFcLong( 0x0 ),	/* 0 */
/* 1238 */	NdrFcLong( 0x186a0 ),	/* 100000 */
/* 1242 */	
			0x4b,		/* FC_PP */
			0x5c,		/* FC_PAD */
/* 1244 */	
			0x48,		/* FC_VARIABLE_REPEAT */
			0x49,		/* FC_FIXED_OFFSET */
/* 1246 */	NdrFcShort( 0x4 ),	/* 4 */
/* 1248 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1250 */	NdrFcShort( 0x1 ),	/* 1 */
/* 1252 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1254 */	NdrFcShort( 0x0 ),	/* 0 */
/* 1256 */	0x12, 0x8,	/* FC_UP [simple_pointer] */
/* 1258 */	
			0x25,		/* FC_C_WSTRING */
			0x5c,		/* FC_PAD */
/* 1260 */	
			0x5b,		/* FC_END */

			0x8,		/* FC_LONG */
/* 1262 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 1264 */	
			0x1a,		/* FC_BOGUS_STRUCT */
			0x3,		/* 3 */
/* 1266 */	NdrFcShort( 0x4 ),	/* 4 */
/* 1268 */	NdrFcShort( 0xffd2 ),	/* Offset= -46 (1222) */
/* 1270 */	NdrFcShort( 0x0 ),	/* Offset= 0 (1270) */
/* 1272 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */

			0x0
        }
    };

static const MS2DOXNSPI_MIDL_EXPR_FORMAT_STRING MS2DOXNSPI__MIDL_ExprFormatString =
    {
        0,
        {
			0x4,		/* FC_EXPR_OPER */
			0x6,		/* OP_UNARY_CAST */
/*  2 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */
/*  4 */	0x4,		/* FC_EXPR_OPER */
			0x1b,		/* OP_AND */
/*  6 */	0x0,		/*  */
			0x0,		/* 0 */
/*  8 */	0x3,		/* FC_EXPR_VAR */
			0x9,		/* FC_ULONG */
/* 10 */	NdrFcShort( 0xfff8 ),	/* -8 */
/* 12 */	0x1,		/* FC_EXPR_CONST32 */
			0x8,		/* FC_LONG */
/* 14 */	NdrFcShort( 0x0 ),	/* 0 */
/* 16 */	NdrFcLong( 0xffff ),	/* 65535 */
/* 20 */	0x4,		/* FC_EXPR_OPER */
			0x6,		/* OP_UNARY_CAST */
/* 22 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */
/* 24 */	0x3,		/* FC_EXPR_VAR */
			0x9,		/* FC_ULONG */
/* 26 */	NdrFcShort( 0xfffc ),	/* -4 */

			0x0
        }
    };

static const unsigned short MS2DOXNSPI__MIDL_ExprFormatStringOffsetTable[] =
{
0,
20,
};

static const NDR_EXPR_DESC  MS2DOXNSPI_ExprDesc = 
 {
MS2DOXNSPI__MIDL_ExprFormatStringOffsetTable,
MS2DOXNSPI__MIDL_ExprFormatString.Format
};

static const unsigned short nspi_FormatStringOffsetTable[] =
    {
    0,
    58,
    106,
    166,
    250,
    328,
    430,
    496,
    556,
    622,
    688,
    760,
    826,
    892,
    970,
    1036,
    1064,
    1124,
    1152,
    1180,
    1258
    };


static const MIDL_STUB_DESC nspi_StubDesc = 
    {
    (void *)& nspi___RpcClientInterface,
    MIDL_user_allocate,
    MIDL_user_free,
    &nspi__MIDL_AutoBindHandle,
    0,
    0,
    0,
    0,
    MS2DOXNSPI__MIDL_TypeFormatString.Format,
    1, /* -error bounds_check flag */
    0x60001, /* Ndr library version */
    0,
    0x700022b, /* MIDL Version 7.0.555 */
    0,
    0,
    0,  /* notify & notify_flag routine table */
    0x1, /* MIDL flag */
    0, /* cs routines */
    0,   /* proxy/server info */
    &MS2DOXNSPI_ExprDesc
    };
#pragma optimize("", on )
#if _MSC_VER >= 1200
#pragma warning(pop)
#endif


#endif /* !defined(_M_IA64) && !defined(_M_AMD64)*/

