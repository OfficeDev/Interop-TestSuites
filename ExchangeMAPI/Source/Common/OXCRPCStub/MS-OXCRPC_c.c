

/* this ALWAYS GENERATED file contains the RPC client stubs */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Fri Jul 22 15:04:54 2016
 */
/* Compiler settings for MS-OXCRPC.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#if !defined(_M_IA64) && !defined(_M_AMD64) && !defined(_ARM_)


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

#include "MS-OXCRPC.h"

#define TYPE_FORMAT_STRING_SIZE   221                               
#define PROC_FORMAT_STRING_SIZE   849                               
#define EXPR_FORMAT_STRING_SIZE   1                                 
#define TRANSMIT_AS_TABLE_SIZE    0            
#define WIRE_MARSHAL_TABLE_SIZE   0            

typedef struct _MS2DOXCRPC_MIDL_TYPE_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ TYPE_FORMAT_STRING_SIZE ];
    } MS2DOXCRPC_MIDL_TYPE_FORMAT_STRING;

typedef struct _MS2DOXCRPC_MIDL_PROC_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ PROC_FORMAT_STRING_SIZE ];
    } MS2DOXCRPC_MIDL_PROC_FORMAT_STRING;

typedef struct _MS2DOXCRPC_MIDL_EXPR_FORMAT_STRING
    {
    long          Pad;
    unsigned char  Format[ EXPR_FORMAT_STRING_SIZE ];
    } MS2DOXCRPC_MIDL_EXPR_FORMAT_STRING;


static const RPC_SYNTAX_IDENTIFIER  _RpcTransferSyntax = 
{{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}};


extern const MS2DOXCRPC_MIDL_TYPE_FORMAT_STRING MS2DOXCRPC__MIDL_TypeFormatString;
extern const MS2DOXCRPC_MIDL_PROC_FORMAT_STRING MS2DOXCRPC__MIDL_ProcFormatString;
extern const MS2DOXCRPC_MIDL_EXPR_FORMAT_STRING MS2DOXCRPC__MIDL_ExprFormatString;

#define GENERIC_BINDING_TABLE_SIZE   0            


/* Standard interface: __MIDL_itf_MS2DOXCRPC_0000_0000, ver. 0.0,
   GUID={0x00000000,0x0000,0x0000,{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00}} */


/* Standard interface: emsmdb, ver. 0.81,
   GUID={0xA4F1DB00,0xCA47,0x1067,{0xB3,0x1F,0x00,0xDD,0x01,0x06,0x62,0xDA}} */



static const RPC_CLIENT_INTERFACE emsmdb___RpcClientInterface =
    {
    sizeof(RPC_CLIENT_INTERFACE),
    {{0xA4F1DB00,0xCA47,0x1067,{0xB3,0x1F,0x00,0xDD,0x01,0x06,0x62,0xDA}},{0,81}},
    {{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}},
    0,
    0,
    0,
    0,
    0,
    0x00000000
    };
RPC_IF_HANDLE emsmdb_v0_81_c_ifspec = (RPC_IF_HANDLE)& emsmdb___RpcClientInterface;

extern const MIDL_STUB_DESC emsmdb_StubDesc;

static RPC_BINDING_HANDLE emsmdb__MIDL_AutoBindHandle;


long __stdcall Opnum0Reserved( 
    /* [in] */ handle_t IDL_handle)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[0],
                  ( unsigned char * )&IDL_handle);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall EcDoDisconnect( 
    /* [ref][out][in] */ CXH *pcxh)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[34],
                  ( unsigned char * )&pcxh);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall Opnum2Reserved( 
    /* [in] */ handle_t IDL_handle)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[76],
                  ( unsigned char * )&IDL_handle);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall Opnum3Reserved( 
    /* [in] */ handle_t IDL_handle)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[110],
                  ( unsigned char * )&IDL_handle);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall EcRRegisterPushNotification( 
    /* [ref][out][in] */ CXH *pcxh,
    /* [in] */ unsigned long iRpc,
    /* [size_is][in] */ unsigned char rgbContext[  ],
    /* [in] */ unsigned short cbContext,
    /* [in] */ unsigned long grbitAdviseBits,
    /* [size_is][in] */ unsigned char rgbCallbackAddress[  ],
    /* [in] */ unsigned short cbCallbackAddress,
    /* [out] */ unsigned long *hNotification)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[144],
                  ( unsigned char * )&pcxh);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall Opnum5Reserved( 
    /* [in] */ handle_t IDL_handle)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[228],
                  ( unsigned char * )&IDL_handle);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall EcDummyRpc( 
    /* [in] */ handle_t hBinding)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[262],
                  ( unsigned char * )&hBinding);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall Opnum7Reserved( 
    /* [in] */ handle_t IDL_handle)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[296],
                  ( unsigned char * )&IDL_handle);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall Opnum8Reserved( 
    /* [in] */ handle_t IDL_handle)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[330],
                  ( unsigned char * )&IDL_handle);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall Opnum9Reserved( 
    /* [in] */ handle_t IDL_handle)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[364],
                  ( unsigned char * )&IDL_handle);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall EcDoConnectEx( 
    /* [in] */ handle_t hBinding,
    /* [ref][out] */ CXH *pcxh,
    /* [string][in] */ unsigned char *szUserDN,
    /* [in] */ unsigned long ulFlags,
    /* [in] */ unsigned long ulConMod,
    /* [in] */ unsigned long cbLimit,
    /* [in] */ unsigned long ulCpid,
    /* [in] */ unsigned long ulLcidString,
    /* [in] */ unsigned long ulLcidSort,
    /* [in] */ unsigned long ulIcxrLink,
    /* [in] */ unsigned short usFCanConvertCodePages,
    /* [out] */ unsigned long *pcmsPollsMax,
    /* [out] */ unsigned long *pcRetry,
    /* [out] */ unsigned long *pcmsRetryDelay,
    /* [out] */ unsigned short *picxr,
    /* [string][out] */ unsigned char **szDNPrefix,
    /* [string][out] */ unsigned char **szDisplayName,
    /* [in] */ unsigned short rgwClientVersion[ 3 ],
    /* [out] */ unsigned short rgwServerVersion[ 3 ],
    /* [out] */ unsigned short rgwBestVersion[ 3 ],
    /* [out][in] */ unsigned long *pulTimeStamp,
    /* [size_is][in] */ unsigned char rgbAuxIn[  ],
    /* [in] */ unsigned long cbAuxIn,
    /* [size_is][length_is][out] */ unsigned char rgbAuxOut[  ],
    /* [out][in] */ SMALL_RANGE_ULONG *pcbAuxOut)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[398],
                  ( unsigned char * )&hBinding);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall EcDoRpcExt2( 
    /* [ref][out][in] */ CXH *pcxh,
    /* [out][in] */ unsigned long *pulFlags,
    /* [size_is][in] */ unsigned char rgbIn[  ],
    /* [in] */ unsigned long cbIn,
    /* [size_is][length_is][out] */ unsigned char rgbOut[  ],
    /* [out][in] */ BIG_RANGE_ULONG *pcbOut,
    /* [size_is][in] */ unsigned char rgbAuxIn[  ],
    /* [in] */ unsigned long cbAuxIn,
    /* [size_is][length_is][out] */ unsigned char rgbAuxOut[  ],
    /* [out][in] */ SMALL_RANGE_ULONG *pcbAuxOut,
    /* [out] */ unsigned long *pulTransTime)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[576],
                  ( unsigned char * )&pcxh);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall Opnum12Reserved( 
    /* [in] */ handle_t IDL_handle)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[678],
                  ( unsigned char * )&IDL_handle);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall Opnum13Reserved( 
    /* [in] */ handle_t IDL_handle)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[712],
                  ( unsigned char * )&IDL_handle);
    return ( long  )_RetVal.Simple;
    
}


long __stdcall EcDoAsyncConnectEx( 
    /* [in] */ CXH cxh,
    /* [ref][out] */ ACXH *pacxh)
{

    CLIENT_CALL_RETURN _RetVal;

    _RetVal = NdrClientCall2(
                  ( PMIDL_STUB_DESC  )&emsmdb_StubDesc,
                  (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[746],
                  ( unsigned char * )&cxh);
    return ( long  )_RetVal.Simple;
    
}


/* Standard interface: asyncemsmdb, ver. 0.1,
   GUID={0x5261574A,0x4572,0x206E,{0xB2,0x68,0x6B,0x19,0x92,0x13,0xB4,0xE4}} */



static const RPC_CLIENT_INTERFACE asyncemsmdb___RpcClientInterface =
    {
    sizeof(RPC_CLIENT_INTERFACE),
    {{0x5261574A,0x4572,0x206E,{0xB2,0x68,0x6B,0x19,0x92,0x13,0xB4,0xE4}},{0,1}},
    {{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}},
    0,
    0,
    0,
    0,
    0,
    0x00000000
    };
RPC_IF_HANDLE asyncemsmdb_v0_1_c_ifspec = (RPC_IF_HANDLE)& asyncemsmdb___RpcClientInterface;

extern const MIDL_STUB_DESC asyncemsmdb_StubDesc;

static RPC_BINDING_HANDLE asyncemsmdb__MIDL_AutoBindHandle;


/* [async] */ void  __stdcall EcDoAsyncWaitEx( 
    /* [in] */ PRPC_ASYNC_STATE EcDoAsyncWaitEx_AsyncHandle,
    /* [in] */ ACXH acxh,
    /* [in] */ unsigned long ulFlagsIn,
    /* [out] */ unsigned long *pulFlagsOut)
{

    NdrAsyncClientCall(
                      ( PMIDL_STUB_DESC  )&asyncemsmdb_StubDesc,
                      (PFORMAT_STRING) &MS2DOXCRPC__MIDL_ProcFormatString.Format[794],
                      ( unsigned char * )&EcDoAsyncWaitEx_AsyncHandle);
    
}


#if !defined(__RPC_WIN32__)
#error  Invalid build platform for this stub.
#endif

#if !(TARGET_IS_NT50_OR_LATER)
#error You need Windows 2000 or later to run this stub because it uses these features:
#error   [async] attribute, /robust command line switch.
#error However, your C/C++ compilation flags indicate you intend to run this app on earlier systems.
#error This app will fail with the RPC_X_WRONG_STUB_VERSION error.
#endif


static const MS2DOXCRPC_MIDL_PROC_FORMAT_STRING MS2DOXCRPC__MIDL_ProcFormatString =
    {
        0,
        {

	/* Procedure Opnum0Reserved */

			0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/*  2 */	NdrFcLong( 0x0 ),	/* 0 */
/*  6 */	NdrFcShort( 0x0 ),	/* 0 */
/*  8 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 10 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 12 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 14 */	NdrFcShort( 0x0 ),	/* 0 */
/* 16 */	NdrFcShort( 0x8 ),	/* 8 */
/* 18 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x1,		/* 1 */
/* 20 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 22 */	NdrFcShort( 0x0 ),	/* 0 */
/* 24 */	NdrFcShort( 0x0 ),	/* 0 */
/* 26 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter IDL_handle */

/* 28 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 30 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 32 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure EcDoDisconnect */


	/* Return value */

/* 34 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 36 */	NdrFcLong( 0x0 ),	/* 0 */
/* 40 */	NdrFcShort( 0x1 ),	/* 1 */
/* 42 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 44 */	0x30,		/* FC_BIND_CONTEXT */
			0xe0,		/* Ctxt flags:  via ptr, in, out, */
/* 46 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 48 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 50 */	NdrFcShort( 0x38 ),	/* 56 */
/* 52 */	NdrFcShort( 0x40 ),	/* 64 */
/* 54 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x2,		/* 2 */
/* 56 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 58 */	NdrFcShort( 0x0 ),	/* 0 */
/* 60 */	NdrFcShort( 0x0 ),	/* 0 */
/* 62 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter pcxh */

/* 64 */	NdrFcShort( 0x118 ),	/* Flags:  in, out, simple ref, */
/* 66 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 68 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Return value */

/* 70 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 72 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 74 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure Opnum2Reserved */

/* 76 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 78 */	NdrFcLong( 0x0 ),	/* 0 */
/* 82 */	NdrFcShort( 0x2 ),	/* 2 */
/* 84 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 86 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 88 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 90 */	NdrFcShort( 0x0 ),	/* 0 */
/* 92 */	NdrFcShort( 0x8 ),	/* 8 */
/* 94 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x1,		/* 1 */
/* 96 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 98 */	NdrFcShort( 0x0 ),	/* 0 */
/* 100 */	NdrFcShort( 0x0 ),	/* 0 */
/* 102 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter IDL_handle */

/* 104 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 106 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 108 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure Opnum3Reserved */


	/* Return value */

/* 110 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 112 */	NdrFcLong( 0x0 ),	/* 0 */
/* 116 */	NdrFcShort( 0x3 ),	/* 3 */
/* 118 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 120 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 122 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 124 */	NdrFcShort( 0x0 ),	/* 0 */
/* 126 */	NdrFcShort( 0x8 ),	/* 8 */
/* 128 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x1,		/* 1 */
/* 130 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 132 */	NdrFcShort( 0x0 ),	/* 0 */
/* 134 */	NdrFcShort( 0x0 ),	/* 0 */
/* 136 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter IDL_handle */

/* 138 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 140 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 142 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure EcRRegisterPushNotification */


	/* Return value */

/* 144 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 146 */	NdrFcLong( 0x0 ),	/* 0 */
/* 150 */	NdrFcShort( 0x4 ),	/* 4 */
/* 152 */	NdrFcShort( 0x24 ),	/* x86 Stack size/offset = 36 */
/* 154 */	0x30,		/* FC_BIND_CONTEXT */
			0xe0,		/* Ctxt flags:  via ptr, in, out, */
/* 156 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 158 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 160 */	NdrFcShort( 0x54 ),	/* 84 */
/* 162 */	NdrFcShort( 0x5c ),	/* 92 */
/* 164 */	0x46,		/* Oi2 Flags:  clt must size, has return, has ext, */
			0x9,		/* 9 */
/* 166 */	0x8,		/* 8 */
			0x5,		/* Ext Flags:  new corr desc, srv corr check, */
/* 168 */	NdrFcShort( 0x0 ),	/* 0 */
/* 170 */	NdrFcShort( 0x1 ),	/* 1 */
/* 172 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter pcxh */

/* 174 */	NdrFcShort( 0x118 ),	/* Flags:  in, out, simple ref, */
/* 176 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 178 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter iRpc */

/* 180 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 182 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 184 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter rgbContext */

/* 186 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 188 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 190 */	NdrFcShort( 0xa ),	/* Type Offset=10 */

	/* Parameter cbContext */

/* 192 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 194 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 196 */	0x6,		/* FC_SHORT */
			0x0,		/* 0 */

	/* Parameter grbitAdviseBits */

/* 198 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 200 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 202 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter rgbCallbackAddress */

/* 204 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 206 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 208 */	NdrFcShort( 0x16 ),	/* Type Offset=22 */

	/* Parameter cbCallbackAddress */

/* 210 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 212 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 214 */	0x6,		/* FC_SHORT */
			0x0,		/* 0 */

	/* Parameter hNotification */

/* 216 */	NdrFcShort( 0x2150 ),	/* Flags:  out, base type, simple ref, srv alloc size=8 */
/* 218 */	NdrFcShort( 0x1c ),	/* x86 Stack size/offset = 28 */
/* 220 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Return value */

/* 222 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 224 */	NdrFcShort( 0x20 ),	/* x86 Stack size/offset = 32 */
/* 226 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure Opnum5Reserved */

/* 228 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 230 */	NdrFcLong( 0x0 ),	/* 0 */
/* 234 */	NdrFcShort( 0x5 ),	/* 5 */
/* 236 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 238 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 240 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 242 */	NdrFcShort( 0x0 ),	/* 0 */
/* 244 */	NdrFcShort( 0x8 ),	/* 8 */
/* 246 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x1,		/* 1 */
/* 248 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 250 */	NdrFcShort( 0x0 ),	/* 0 */
/* 252 */	NdrFcShort( 0x0 ),	/* 0 */
/* 254 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter IDL_handle */

/* 256 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 258 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 260 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure EcDummyRpc */


	/* Return value */

/* 262 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 264 */	NdrFcLong( 0x0 ),	/* 0 */
/* 268 */	NdrFcShort( 0x6 ),	/* 6 */
/* 270 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 272 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 274 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 276 */	NdrFcShort( 0x0 ),	/* 0 */
/* 278 */	NdrFcShort( 0x8 ),	/* 8 */
/* 280 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x1,		/* 1 */
/* 282 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 284 */	NdrFcShort( 0x0 ),	/* 0 */
/* 286 */	NdrFcShort( 0x0 ),	/* 0 */
/* 288 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hBinding */

/* 290 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 292 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 294 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure Opnum7Reserved */


	/* Return value */

/* 296 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 298 */	NdrFcLong( 0x0 ),	/* 0 */
/* 302 */	NdrFcShort( 0x7 ),	/* 7 */
/* 304 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 306 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 308 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 310 */	NdrFcShort( 0x0 ),	/* 0 */
/* 312 */	NdrFcShort( 0x8 ),	/* 8 */
/* 314 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x1,		/* 1 */
/* 316 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 318 */	NdrFcShort( 0x0 ),	/* 0 */
/* 320 */	NdrFcShort( 0x0 ),	/* 0 */
/* 322 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter IDL_handle */

/* 324 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 326 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 328 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure Opnum8Reserved */


	/* Return value */

/* 330 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 332 */	NdrFcLong( 0x0 ),	/* 0 */
/* 336 */	NdrFcShort( 0x8 ),	/* 8 */
/* 338 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 340 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 342 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 344 */	NdrFcShort( 0x0 ),	/* 0 */
/* 346 */	NdrFcShort( 0x8 ),	/* 8 */
/* 348 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x1,		/* 1 */
/* 350 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 352 */	NdrFcShort( 0x0 ),	/* 0 */
/* 354 */	NdrFcShort( 0x0 ),	/* 0 */
/* 356 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter IDL_handle */

/* 358 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 360 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 362 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure Opnum9Reserved */


	/* Return value */

/* 364 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 366 */	NdrFcLong( 0x0 ),	/* 0 */
/* 370 */	NdrFcShort( 0x9 ),	/* 9 */
/* 372 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 374 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 376 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 378 */	NdrFcShort( 0x0 ),	/* 0 */
/* 380 */	NdrFcShort( 0x8 ),	/* 8 */
/* 382 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x1,		/* 1 */
/* 384 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 386 */	NdrFcShort( 0x0 ),	/* 0 */
/* 388 */	NdrFcShort( 0x0 ),	/* 0 */
/* 390 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter IDL_handle */

/* 392 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 394 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 396 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure EcDoConnectEx */


	/* Return value */

/* 398 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 400 */	NdrFcLong( 0x0 ),	/* 0 */
/* 404 */	NdrFcShort( 0xa ),	/* 10 */
/* 406 */	NdrFcShort( 0x68 ),	/* x86 Stack size/offset = 104 */
/* 408 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 410 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 412 */	NdrFcShort( 0x94 ),	/* 148 */
/* 414 */	NdrFcShort( 0x112 ),	/* 274 */
/* 416 */	0x47,		/* Oi2 Flags:  srv must size, clt must size, has return, has ext, */
			0x19,		/* 25 */
/* 418 */	0x8,		/* 8 */
			0x7,		/* Ext Flags:  new corr desc, clt corr check, srv corr check, */
/* 420 */	NdrFcShort( 0x1 ),	/* 1 */
/* 422 */	NdrFcShort( 0x1 ),	/* 1 */
/* 424 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter hBinding */

/* 426 */	NdrFcShort( 0x110 ),	/* Flags:  out, simple ref, */
/* 428 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 430 */	NdrFcShort( 0x2a ),	/* Type Offset=42 */

	/* Parameter pcxh */

/* 432 */	NdrFcShort( 0x10b ),	/* Flags:  must size, must free, in, simple ref, */
/* 434 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 436 */	NdrFcShort( 0x30 ),	/* Type Offset=48 */

	/* Parameter szUserDN */

/* 438 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 440 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 442 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter ulFlags */

/* 444 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 446 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 448 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter ulConMod */

/* 450 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 452 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 454 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter cbLimit */

/* 456 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 458 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 460 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter ulCpid */

/* 462 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 464 */	NdrFcShort( 0x1c ),	/* x86 Stack size/offset = 28 */
/* 466 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter ulLcidString */

/* 468 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 470 */	NdrFcShort( 0x20 ),	/* x86 Stack size/offset = 32 */
/* 472 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter ulLcidSort */

/* 474 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 476 */	NdrFcShort( 0x24 ),	/* x86 Stack size/offset = 36 */
/* 478 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter ulIcxrLink */

/* 480 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 482 */	NdrFcShort( 0x28 ),	/* x86 Stack size/offset = 40 */
/* 484 */	0x6,		/* FC_SHORT */
			0x0,		/* 0 */

	/* Parameter usFCanConvertCodePages */

/* 486 */	NdrFcShort( 0x2150 ),	/* Flags:  out, base type, simple ref, srv alloc size=8 */
/* 488 */	NdrFcShort( 0x2c ),	/* x86 Stack size/offset = 44 */
/* 490 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pcmsPollsMax */

/* 492 */	NdrFcShort( 0x2150 ),	/* Flags:  out, base type, simple ref, srv alloc size=8 */
/* 494 */	NdrFcShort( 0x30 ),	/* x86 Stack size/offset = 48 */
/* 496 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pcRetry */

/* 498 */	NdrFcShort( 0x2150 ),	/* Flags:  out, base type, simple ref, srv alloc size=8 */
/* 500 */	NdrFcShort( 0x34 ),	/* x86 Stack size/offset = 52 */
/* 502 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pcmsRetryDelay */

/* 504 */	NdrFcShort( 0x2150 ),	/* Flags:  out, base type, simple ref, srv alloc size=8 */
/* 506 */	NdrFcShort( 0x38 ),	/* x86 Stack size/offset = 56 */
/* 508 */	0x6,		/* FC_SHORT */
			0x0,		/* 0 */

	/* Parameter picxr */

/* 510 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 512 */	NdrFcShort( 0x3c ),	/* x86 Stack size/offset = 60 */
/* 514 */	NdrFcShort( 0x36 ),	/* Type Offset=54 */

	/* Parameter szDNPrefix */

/* 516 */	NdrFcShort( 0x2013 ),	/* Flags:  must size, must free, out, srv alloc size=8 */
/* 518 */	NdrFcShort( 0x40 ),	/* x86 Stack size/offset = 64 */
/* 520 */	NdrFcShort( 0x36 ),	/* Type Offset=54 */

	/* Parameter szDisplayName */

/* 522 */	NdrFcShort( 0xa ),	/* Flags:  must free, in, */
/* 524 */	NdrFcShort( 0x44 ),	/* x86 Stack size/offset = 68 */
/* 526 */	NdrFcShort( 0x3e ),	/* Type Offset=62 */

	/* Parameter rgwClientVersion */

/* 528 */	NdrFcShort( 0x12 ),	/* Flags:  must free, out, */
/* 530 */	NdrFcShort( 0x48 ),	/* x86 Stack size/offset = 72 */
/* 532 */	NdrFcShort( 0x3e ),	/* Type Offset=62 */

	/* Parameter rgwServerVersion */

/* 534 */	NdrFcShort( 0x12 ),	/* Flags:  must free, out, */
/* 536 */	NdrFcShort( 0x4c ),	/* x86 Stack size/offset = 76 */
/* 538 */	NdrFcShort( 0x3e ),	/* Type Offset=62 */

	/* Parameter rgwBestVersion */

/* 540 */	NdrFcShort( 0x158 ),	/* Flags:  in, out, base type, simple ref, */
/* 542 */	NdrFcShort( 0x50 ),	/* x86 Stack size/offset = 80 */
/* 544 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pulTimeStamp */

/* 546 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 548 */	NdrFcShort( 0x54 ),	/* x86 Stack size/offset = 84 */
/* 550 */	NdrFcShort( 0x48 ),	/* Type Offset=72 */

	/* Parameter rgbAuxIn */

/* 552 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 554 */	NdrFcShort( 0x58 ),	/* x86 Stack size/offset = 88 */
/* 556 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter cbAuxIn */

/* 558 */	NdrFcShort( 0x13 ),	/* Flags:  must size, must free, out, */
/* 560 */	NdrFcShort( 0x5c ),	/* x86 Stack size/offset = 92 */
/* 562 */	NdrFcShort( 0x54 ),	/* Type Offset=84 */

	/* Parameter rgbAuxOut */

/* 564 */	NdrFcShort( 0x11a ),	/* Flags:  must free, in, out, simple ref, */
/* 566 */	NdrFcShort( 0x60 ),	/* x86 Stack size/offset = 96 */
/* 568 */	NdrFcShort( 0x6a ),	/* Type Offset=106 */

	/* Parameter pcbAuxOut */

/* 570 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 572 */	NdrFcShort( 0x64 ),	/* x86 Stack size/offset = 100 */
/* 574 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure EcDoRpcExt2 */


	/* Return value */

/* 576 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 578 */	NdrFcLong( 0x0 ),	/* 0 */
/* 582 */	NdrFcShort( 0xb ),	/* 11 */
/* 584 */	NdrFcShort( 0x30 ),	/* x86 Stack size/offset = 48 */
/* 586 */	0x30,		/* FC_BIND_CONTEXT */
			0xe0,		/* Ctxt flags:  via ptr, in, out, */
/* 588 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 590 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 592 */	NdrFcShort( 0x9c ),	/* 156 */
/* 594 */	NdrFcShort( 0xb0 ),	/* 176 */
/* 596 */	0x47,		/* Oi2 Flags:  srv must size, clt must size, has return, has ext, */
			0xc,		/* 12 */
/* 598 */	0x8,		/* 8 */
			0x7,		/* Ext Flags:  new corr desc, clt corr check, srv corr check, */
/* 600 */	NdrFcShort( 0x1 ),	/* 1 */
/* 602 */	NdrFcShort( 0x1 ),	/* 1 */
/* 604 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter pcxh */

/* 606 */	NdrFcShort( 0x118 ),	/* Flags:  in, out, simple ref, */
/* 608 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 610 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */

	/* Parameter pulFlags */

/* 612 */	NdrFcShort( 0x158 ),	/* Flags:  in, out, base type, simple ref, */
/* 614 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 616 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter rgbIn */

/* 618 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 620 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 622 */	NdrFcShort( 0x74 ),	/* Type Offset=116 */

	/* Parameter cbIn */

/* 624 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 626 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 628 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter rgbOut */

/* 630 */	NdrFcShort( 0x13 ),	/* Flags:  must size, must free, out, */
/* 632 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 634 */	NdrFcShort( 0x80 ),	/* Type Offset=128 */

	/* Parameter pcbOut */

/* 636 */	NdrFcShort( 0x11a ),	/* Flags:  must free, in, out, simple ref, */
/* 638 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 640 */	NdrFcShort( 0x96 ),	/* Type Offset=150 */

	/* Parameter rgbAuxIn */

/* 642 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 644 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 646 */	NdrFcShort( 0xa0 ),	/* Type Offset=160 */

	/* Parameter cbAuxIn */

/* 648 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 650 */	NdrFcShort( 0x1c ),	/* x86 Stack size/offset = 28 */
/* 652 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter rgbAuxOut */

/* 654 */	NdrFcShort( 0x13 ),	/* Flags:  must size, must free, out, */
/* 656 */	NdrFcShort( 0x20 ),	/* x86 Stack size/offset = 32 */
/* 658 */	NdrFcShort( 0xac ),	/* Type Offset=172 */

	/* Parameter pcbAuxOut */

/* 660 */	NdrFcShort( 0x11a ),	/* Flags:  must free, in, out, simple ref, */
/* 662 */	NdrFcShort( 0x24 ),	/* x86 Stack size/offset = 36 */
/* 664 */	NdrFcShort( 0xc2 ),	/* Type Offset=194 */

	/* Parameter pulTransTime */

/* 666 */	NdrFcShort( 0x2150 ),	/* Flags:  out, base type, simple ref, srv alloc size=8 */
/* 668 */	NdrFcShort( 0x28 ),	/* x86 Stack size/offset = 40 */
/* 670 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Return value */

/* 672 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 674 */	NdrFcShort( 0x2c ),	/* x86 Stack size/offset = 44 */
/* 676 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure Opnum12Reserved */

/* 678 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 680 */	NdrFcLong( 0x0 ),	/* 0 */
/* 684 */	NdrFcShort( 0xc ),	/* 12 */
/* 686 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 688 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 690 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 692 */	NdrFcShort( 0x0 ),	/* 0 */
/* 694 */	NdrFcShort( 0x8 ),	/* 8 */
/* 696 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x1,		/* 1 */
/* 698 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 700 */	NdrFcShort( 0x0 ),	/* 0 */
/* 702 */	NdrFcShort( 0x0 ),	/* 0 */
/* 704 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter IDL_handle */

/* 706 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 708 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 710 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure Opnum13Reserved */


	/* Return value */

/* 712 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 714 */	NdrFcLong( 0x0 ),	/* 0 */
/* 718 */	NdrFcShort( 0xd ),	/* 13 */
/* 720 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 722 */	0x32,		/* FC_BIND_PRIMITIVE */
			0x0,		/* 0 */
/* 724 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 726 */	NdrFcShort( 0x0 ),	/* 0 */
/* 728 */	NdrFcShort( 0x8 ),	/* 8 */
/* 730 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x1,		/* 1 */
/* 732 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 734 */	NdrFcShort( 0x0 ),	/* 0 */
/* 736 */	NdrFcShort( 0x0 ),	/* 0 */
/* 738 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter IDL_handle */

/* 740 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 742 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 744 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure EcDoAsyncConnectEx */


	/* Return value */

/* 746 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 748 */	NdrFcLong( 0x0 ),	/* 0 */
/* 752 */	NdrFcShort( 0xe ),	/* 14 */
/* 754 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 756 */	0x30,		/* FC_BIND_CONTEXT */
			0x40,		/* Ctxt flags:  in, */
/* 758 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 760 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 762 */	NdrFcShort( 0x24 ),	/* 36 */
/* 764 */	NdrFcShort( 0x40 ),	/* 64 */
/* 766 */	0x44,		/* Oi2 Flags:  has return, has ext, */
			0x3,		/* 3 */
/* 768 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 770 */	NdrFcShort( 0x0 ),	/* 0 */
/* 772 */	NdrFcShort( 0x0 ),	/* 0 */
/* 774 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter cxh */

/* 776 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 778 */	NdrFcShort( 0x0 ),	/* x86 Stack size/offset = 0 */
/* 780 */	NdrFcShort( 0xcc ),	/* Type Offset=204 */

	/* Parameter pacxh */

/* 782 */	NdrFcShort( 0x110 ),	/* Flags:  out, simple ref, */
/* 784 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 786 */	NdrFcShort( 0xd4 ),	/* Type Offset=212 */

	/* Return value */

/* 788 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 790 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 792 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure EcDoAsyncWaitEx */

/* 794 */	0x0,		/* 0 */
			0x48,		/* Old Flags:  */
/* 796 */	NdrFcLong( 0x0 ),	/* 0 */
/* 800 */	NdrFcShort( 0x0 ),	/* 0 */
/* 802 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 804 */	0x30,		/* FC_BIND_CONTEXT */
			0x44,		/* Ctxt flags:  in, no serialize, */
/* 806 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 808 */	0x1,		/* 1 */
			0x0,		/* 0 */
/* 810 */	NdrFcShort( 0x2c ),	/* 44 */
/* 812 */	NdrFcShort( 0x24 ),	/* 36 */
/* 814 */	0xc4,		/* Oi2 Flags:  has return, has ext, has async handle */
			0x4,		/* 4 */
/* 816 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 818 */	NdrFcShort( 0x0 ),	/* 0 */
/* 820 */	NdrFcShort( 0x0 ),	/* 0 */
/* 822 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter acxh */

/* 824 */	NdrFcShort( 0x8 ),	/* Flags:  in, */
/* 826 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 828 */	NdrFcShort( 0xd8 ),	/* Type Offset=216 */

	/* Parameter ulFlagsIn */

/* 830 */	NdrFcShort( 0x48 ),	/* Flags:  in, base type, */
/* 832 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 834 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Parameter pulFlagsOut */

/* 836 */	NdrFcShort( 0x2150 ),	/* Flags:  out, base type, simple ref, srv alloc size=8 */
/* 838 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 840 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Return value */

/* 842 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 844 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 846 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

			0x0
        }
    };

static const MS2DOXCRPC_MIDL_TYPE_FORMAT_STRING MS2DOXCRPC__MIDL_TypeFormatString =
    {
        0,
        {
			NdrFcShort( 0x0 ),	/* 0 */
/*  2 */	
			0x11, 0x4,	/* FC_RP [alloced_on_stack] */
/*  4 */	NdrFcShort( 0x2 ),	/* Offset= 2 (6) */
/*  6 */	0x30,		/* FC_BIND_CONTEXT */
			0xe1,		/* Ctxt flags:  via ptr, in, out, can't be null */
/*  8 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 10 */	
			0x1b,		/* FC_CARRAY */
			0x0,		/* 0 */
/* 12 */	NdrFcShort( 0x1 ),	/* 1 */
/* 14 */	0x27,		/* Corr desc:  parameter, FC_USHORT */
			0x0,		/*  */
/* 16 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 18 */	NdrFcShort( 0x0 ),	/* Corr flags:  */
/* 20 */	0x2,		/* FC_CHAR */
			0x5b,		/* FC_END */
/* 22 */	
			0x1b,		/* FC_CARRAY */
			0x0,		/* 0 */
/* 24 */	NdrFcShort( 0x1 ),	/* 1 */
/* 26 */	0x27,		/* Corr desc:  parameter, FC_USHORT */
			0x0,		/*  */
/* 28 */	NdrFcShort( 0x18 ),	/* x86 Stack size/offset = 24 */
/* 30 */	NdrFcShort( 0x0 ),	/* Corr flags:  */
/* 32 */	0x2,		/* FC_CHAR */
			0x5b,		/* FC_END */
/* 34 */	
			0x11, 0xc,	/* FC_RP [alloced_on_stack] [simple_pointer] */
/* 36 */	0x8,		/* FC_LONG */
			0x5c,		/* FC_PAD */
/* 38 */	
			0x11, 0x4,	/* FC_RP [alloced_on_stack] */
/* 40 */	NdrFcShort( 0x2 ),	/* Offset= 2 (42) */
/* 42 */	0x30,		/* FC_BIND_CONTEXT */
			0xa0,		/* Ctxt flags:  via ptr, out, */
/* 44 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 46 */	
			0x11, 0x8,	/* FC_RP [simple_pointer] */
/* 48 */	
			0x22,		/* FC_C_CSTRING */
			0x5c,		/* FC_PAD */
/* 50 */	
			0x11, 0xc,	/* FC_RP [alloced_on_stack] [simple_pointer] */
/* 52 */	0x6,		/* FC_SHORT */
			0x5c,		/* FC_PAD */
/* 54 */	
			0x11, 0x14,	/* FC_RP [alloced_on_stack] [pointer_deref] */
/* 56 */	NdrFcShort( 0x2 ),	/* Offset= 2 (58) */
/* 58 */	
			0x12, 0x8,	/* FC_UP [simple_pointer] */
/* 60 */	
			0x22,		/* FC_C_CSTRING */
			0x5c,		/* FC_PAD */
/* 62 */	
			0x1d,		/* FC_SMFARRAY */
			0x1,		/* 1 */
/* 64 */	NdrFcShort( 0x6 ),	/* 6 */
/* 66 */	0x6,		/* FC_SHORT */
			0x5b,		/* FC_END */
/* 68 */	
			0x11, 0x8,	/* FC_RP [simple_pointer] */
/* 70 */	0x8,		/* FC_LONG */
			0x5c,		/* FC_PAD */
/* 72 */	
			0x1b,		/* FC_CARRAY */
			0x0,		/* 0 */
/* 74 */	NdrFcShort( 0x1 ),	/* 1 */
/* 76 */	0x29,		/* Corr desc:  parameter, FC_ULONG */
			0x0,		/*  */
/* 78 */	NdrFcShort( 0x58 ),	/* x86 Stack size/offset = 88 */
/* 80 */	NdrFcShort( 0x0 ),	/* Corr flags:  */
/* 82 */	0x2,		/* FC_CHAR */
			0x5b,		/* FC_END */
/* 84 */	
			0x1c,		/* FC_CVARRAY */
			0x0,		/* 0 */
/* 86 */	NdrFcShort( 0x1 ),	/* 1 */
/* 88 */	0x29,		/* Corr desc:  parameter, FC_ULONG */
			0x54,		/* FC_DEREFERENCE */
/* 90 */	NdrFcShort( 0x60 ),	/* x86 Stack size/offset = 96 */
/* 92 */	NdrFcShort( 0x0 ),	/* Corr flags:  */
/* 94 */	0x29,		/* Corr desc:  parameter, FC_ULONG */
			0x54,		/* FC_DEREFERENCE */
/* 96 */	NdrFcShort( 0x60 ),	/* x86 Stack size/offset = 96 */
/* 98 */	NdrFcShort( 0x0 ),	/* Corr flags:  */
/* 100 */	0x2,		/* FC_CHAR */
			0x5b,		/* FC_END */
/* 102 */	
			0x11, 0x0,	/* FC_RP */
/* 104 */	NdrFcShort( 0x2 ),	/* Offset= 2 (106) */
/* 106 */	0xb7,		/* FC_RANGE */
			0x8,		/* 8 */
/* 108 */	NdrFcLong( 0x0 ),	/* 0 */
/* 112 */	NdrFcLong( 0x1008 ),	/* 4104 */
/* 116 */	
			0x1b,		/* FC_CARRAY */
			0x0,		/* 0 */
/* 118 */	NdrFcShort( 0x1 ),	/* 1 */
/* 120 */	0x29,		/* Corr desc:  parameter, FC_ULONG */
			0x0,		/*  */
/* 122 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 124 */	NdrFcShort( 0x0 ),	/* Corr flags:  */
/* 126 */	0x2,		/* FC_CHAR */
			0x5b,		/* FC_END */
/* 128 */	
			0x1c,		/* FC_CVARRAY */
			0x0,		/* 0 */
/* 130 */	NdrFcShort( 0x1 ),	/* 1 */
/* 132 */	0x29,		/* Corr desc:  parameter, FC_ULONG */
			0x54,		/* FC_DEREFERENCE */
/* 134 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 136 */	NdrFcShort( 0x0 ),	/* Corr flags:  */
/* 138 */	0x29,		/* Corr desc:  parameter, FC_ULONG */
			0x54,		/* FC_DEREFERENCE */
/* 140 */	NdrFcShort( 0x14 ),	/* x86 Stack size/offset = 20 */
/* 142 */	NdrFcShort( 0x0 ),	/* Corr flags:  */
/* 144 */	0x2,		/* FC_CHAR */
			0x5b,		/* FC_END */
/* 146 */	
			0x11, 0x0,	/* FC_RP */
/* 148 */	NdrFcShort( 0x2 ),	/* Offset= 2 (150) */
/* 150 */	0xb7,		/* FC_RANGE */
			0x8,		/* 8 */
/* 152 */	NdrFcLong( 0x0 ),	/* 0 */
/* 156 */	NdrFcLong( 0x40000 ),	/* 262144 */
/* 160 */	
			0x1b,		/* FC_CARRAY */
			0x0,		/* 0 */
/* 162 */	NdrFcShort( 0x1 ),	/* 1 */
/* 164 */	0x29,		/* Corr desc:  parameter, FC_ULONG */
			0x0,		/*  */
/* 166 */	NdrFcShort( 0x1c ),	/* x86 Stack size/offset = 28 */
/* 168 */	NdrFcShort( 0x0 ),	/* Corr flags:  */
/* 170 */	0x2,		/* FC_CHAR */
			0x5b,		/* FC_END */
/* 172 */	
			0x1c,		/* FC_CVARRAY */
			0x0,		/* 0 */
/* 174 */	NdrFcShort( 0x1 ),	/* 1 */
/* 176 */	0x29,		/* Corr desc:  parameter, FC_ULONG */
			0x54,		/* FC_DEREFERENCE */
/* 178 */	NdrFcShort( 0x24 ),	/* x86 Stack size/offset = 36 */
/* 180 */	NdrFcShort( 0x0 ),	/* Corr flags:  */
/* 182 */	0x29,		/* Corr desc:  parameter, FC_ULONG */
			0x54,		/* FC_DEREFERENCE */
/* 184 */	NdrFcShort( 0x24 ),	/* x86 Stack size/offset = 36 */
/* 186 */	NdrFcShort( 0x0 ),	/* Corr flags:  */
/* 188 */	0x2,		/* FC_CHAR */
			0x5b,		/* FC_END */
/* 190 */	
			0x11, 0x0,	/* FC_RP */
/* 192 */	NdrFcShort( 0x2 ),	/* Offset= 2 (194) */
/* 194 */	0xb7,		/* FC_RANGE */
			0x8,		/* 8 */
/* 196 */	NdrFcLong( 0x0 ),	/* 0 */
/* 200 */	NdrFcLong( 0x1008 ),	/* 4104 */
/* 204 */	0x30,		/* FC_BIND_CONTEXT */
			0x41,		/* Ctxt flags:  in, can't be null */
/* 206 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 208 */	
			0x11, 0x4,	/* FC_RP [alloced_on_stack] */
/* 210 */	NdrFcShort( 0x2 ),	/* Offset= 2 (212) */
/* 212 */	0x30,		/* FC_BIND_CONTEXT */
			0xa4,		/* Ctxt flags:  via ptr, out, no serialize, */
/* 214 */	0x1,		/* 1 */
			0x1,		/* 1 */
/* 216 */	0x30,		/* FC_BIND_CONTEXT */
			0x45,		/* Ctxt flags:  in, no serialize, can't be null */
/* 218 */	0x1,		/* 1 */
			0x0,		/* 0 */

			0x0
        }
    };

static const unsigned short emsmdb_FormatStringOffsetTable[] =
    {
    0,
    34,
    76,
    110,
    144,
    228,
    262,
    296,
    330,
    364,
    398,
    576,
    678,
    712,
    746
    };


static const MIDL_STUB_DESC emsmdb_StubDesc = 
    {
    (void *)& emsmdb___RpcClientInterface,
    MIDL_user_allocate,
    MIDL_user_free,
    &emsmdb__MIDL_AutoBindHandle,
    0,
    0,
    0,
    0,
    MS2DOXCRPC__MIDL_TypeFormatString.Format,
    1, /* -error bounds_check flag */
    0x50002, /* Ndr library version */
    0,
    0x800025b, /* MIDL Version 8.0.603 */
    0,
    0,
    0,  /* notify & notify_flag routine table */
    0x1, /* MIDL flag */
    0, /* cs routines */
    0,   /* proxy/server info */
    0
    };

static const unsigned short asyncemsmdb_FormatStringOffsetTable[] =
    {
    794
    };


static const MIDL_STUB_DESC asyncemsmdb_StubDesc = 
    {
    (void *)& asyncemsmdb___RpcClientInterface,
    MIDL_user_allocate,
    MIDL_user_free,
    &asyncemsmdb__MIDL_AutoBindHandle,
    0,
    0,
    0,
    0,
    MS2DOXCRPC__MIDL_TypeFormatString.Format,
    1, /* -error bounds_check flag */
    0x50002, /* Ndr library version */
    0,
    0x800025b, /* MIDL Version 8.0.603 */
    0,
    0,
    0,  /* notify & notify_flag routine table */
    0x1, /* MIDL flag */
    0, /* cs routines */
    0,   /* proxy/server info */
    0
    };
#pragma optimize("", on )
#if _MSC_VER >= 1200
#pragma warning(pop)
#endif


#endif /* !defined(_M_IA64) && !defined(_M_AMD64) && !defined(_ARM_) */

