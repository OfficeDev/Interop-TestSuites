

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


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

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __MS2DOXCRPC_h__
#define __MS2DOXCRPC_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_MS2DOXCRPC_0000_0000 */
/* [local] */ 

typedef /* [context_handle] */ void *CXH;

typedef /* [context_handle_noserialize][context_handle] */ void *ACXH;

typedef /* [range] */ unsigned long BIG_RANGE_ULONG;

typedef /* [range] */ unsigned long SMALL_RANGE_ULONG;



extern RPC_IF_HANDLE __MIDL_itf_MS2DOXCRPC_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_MS2DOXCRPC_0000_0000_v0_0_s_ifspec;

#ifndef __emsmdb_INTERFACE_DEFINED__
#define __emsmdb_INTERFACE_DEFINED__

/* interface emsmdb */
/* [unique][version][uuid] */ 

long __stdcall Opnum0Reserved( 
    /* [in] */ handle_t IDL_handle);

long __stdcall EcDoDisconnect( 
    /* [ref][out][in] */ CXH *pcxh);

long __stdcall Opnum2Reserved( 
    /* [in] */ handle_t IDL_handle);

long __stdcall Opnum3Reserved( 
    /* [in] */ handle_t IDL_handle);

long __stdcall EcRRegisterPushNotification( 
    /* [ref][out][in] */ CXH *pcxh,
    /* [in] */ unsigned long iRpc,
    /* [size_is][in] */ unsigned char rgbContext[  ],
    /* [in] */ unsigned short cbContext,
    /* [in] */ unsigned long grbitAdviseBits,
    /* [size_is][in] */ unsigned char rgbCallbackAddress[  ],
    /* [in] */ unsigned short cbCallbackAddress,
    /* [out] */ unsigned long *hNotification);

long __stdcall Opnum5Reserved( 
    /* [in] */ handle_t IDL_handle);

long __stdcall EcDummyRpc( 
    /* [in] */ handle_t hBinding);

long __stdcall Opnum7Reserved( 
    /* [in] */ handle_t IDL_handle);

long __stdcall Opnum8Reserved( 
    /* [in] */ handle_t IDL_handle);

long __stdcall Opnum9Reserved( 
    /* [in] */ handle_t IDL_handle);

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
    /* [out][in] */ SMALL_RANGE_ULONG *pcbAuxOut);

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
    /* [out] */ unsigned long *pulTransTime);

long __stdcall Opnum12Reserved( 
    /* [in] */ handle_t IDL_handle);

long __stdcall Opnum13Reserved( 
    /* [in] */ handle_t IDL_handle);

long __stdcall EcDoAsyncConnectEx( 
    /* [in] */ CXH cxh,
    /* [ref][out] */ ACXH *pacxh);



extern RPC_IF_HANDLE emsmdb_v0_81_c_ifspec;
extern RPC_IF_HANDLE emsmdb_v0_81_s_ifspec;
#endif /* __emsmdb_INTERFACE_DEFINED__ */

#ifndef __asyncemsmdb_INTERFACE_DEFINED__
#define __asyncemsmdb_INTERFACE_DEFINED__

/* interface asyncemsmdb */
/* [unique][version][uuid] */ 

/* [async] */ void  __stdcall EcDoAsyncWaitEx( 
    /* [in] */ PRPC_ASYNC_STATE EcDoAsyncWaitEx_AsyncHandle,
    /* [in] */ ACXH acxh,
    /* [in] */ unsigned long ulFlagsIn,
    /* [out] */ unsigned long *pulFlagsOut);



extern RPC_IF_HANDLE asyncemsmdb_v0_1_c_ifspec;
extern RPC_IF_HANDLE asyncemsmdb_v0_1_s_ifspec;
#endif /* __asyncemsmdb_INTERFACE_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

void __RPC_USER CXH_rundown( CXH );
void __RPC_USER ACXH_rundown( ACXH );

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


