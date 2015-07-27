/////////////////////////////////////////////////////////////////////////////
 //
 // ms-dtyp.h - Windows Data Types header. Adapted from the Full MS-DTYP IDL,
 //             as of 11 October 2010. Declarations and typedefs are in their
 //             original order.
 //
 // Version: 1.0 12 October 2010.
 //
 // [MS-DTYP]: Windows Data Types
 // http://msdn.microsoft.com/en-us/library/cc230273.aspx
 // Section 5 Appendix A: Full MS-DTYP IDL
 // http://msdn.microsoft.com/en-us/library/cc230300.aspx
 //
 // This has been test compiled without errors with the midl compiler (Version
 // 7.00.0555), as well as Visual Studio 2010 - with and without windows.h
 // included, with the _WIN32_WINNT version constant defines shown below.
 //
 // See also:
 // Windows Data Types
 // http://msdn.microsoft.com/en-us/library/aa383751(VS.85).aspx
 //
 /////////////////////////////////////////////////////////////////////////////
 //
 // If compiling with the MIDL compiler, "MIDL_PASS" must be defined.
 // This can be done by using the midl command line with the /D switch.
 // For example, "midl /D MIDL_PASS filename.idl".
 //
 /////////////////////////////////////////////////////////////////////////////
 #ifndef _WINDOWS_
 //
 // Defined in:
 // Microsoft SDKs\Windows\v7.0a\Include\BaseTsd.h
 //
 #if defined(_WIN64)
     #define __int3264   __int64
 #else
     #define __int3264   __int32
 #endif // (_WIN64)
 #ifndef MIDL_PASS
 typedef unsigned char byte;
 #endif // MIDL_PASS
 #endif // _WINDOWS_

#ifndef ANYSIZE_ARRAY
 //
 // Defined in:
 // Microsoft SDKs\Windows\v7.0a\Include\WinNT.h
 //
 #define ANYSIZE_ARRAY 1
 #endif // ANYSIZE_ARRAY

typedef int BOOL, *PBOOL, *LPBOOL;
 typedef unsigned char BYTE, *PBYTE, *LPBYTE;
 typedef BYTE BOOLEAN, *PBOOLEAN;
 typedef wchar_t WCHAR, *PWCHAR;
 typedef WCHAR* BSTR;
 typedef char CHAR, *PCHAR;
 typedef double DOUBLE;
 typedef unsigned long DWORD, *PDWORD, *LPDWORD;
 typedef unsigned int DWORD32;
 typedef unsigned __int64 DWORD64;
 typedef unsigned __int64 ULONGLONG;
 typedef ULONGLONG DWORDLONG, *PDWORDLONG;
 typedef unsigned long error_status_t;
 typedef float FLOAT;
 typedef unsigned char UCHAR, *PUCHAR;
 typedef short SHORT;

typedef void *HANDLE;  
 typedef DWORD HCALL;
 typedef int INT, *LPINT;
 typedef signed char INT8;
 typedef signed short INT16;
 typedef signed int INT32;
 typedef __int64 INT64;
 typedef const wchar_t* LMCSTR; 
 typedef WCHAR* LMSTR;
 typedef long LONG, *PLONG, *LPLONG;
 typedef INT64 LONGLONG;
 typedef LONG HRESULT;

#ifndef _BASETSD_H_
 //
 // Defined in:
 // Microsoft SDKs\Windows\v7.0a\Include\BaseTsd.h
 //
 #ifdef MIDL_PASS
 typedef [public] __int3264 LONG_PTR;
 typedef [public] unsigned __int3264 ULONG_PTR;
 #else
 typedef __int3264 LONG_PTR;
 typedef unsigned __int3264 ULONG_PTR;
 #endif // MIDL_PASS
 #endif // _BASETSD_H_

typedef signed int LONG32;
 typedef signed __int64 LONG64;
 typedef const char* LPCSTR;

typedef const wchar_t* LPCWSTR;
 typedef char* PSTR, *LPSTR;

typedef wchar_t* LPWSTR, *PWSTR;
 typedef DWORD NET_API_STATUS;
 typedef long NTSTATUS;

#ifdef MIDL_PASS
 typedef [context_handle] void* PCONTEXT_HANDLE; 
 typedef [ref] PCONTEXT_HANDLE* PPCONTEXT_HANDLE;
 #else
 typedef void* PCONTEXT_HANDLE; 
 typedef PCONTEXT_HANDLE* PPCONTEXT_HANDLE;
 #endif // MIDL_PASS

typedef unsigned __int64 QWORD;
 typedef void* RPC_BINDING_HANDLE;
 typedef UCHAR* STRING;

typedef unsigned int UINT;
 typedef unsigned char UINT8;
 typedef unsigned short UINT16;
 typedef unsigned int UINT32;
 typedef unsigned __int64 UINT64;
 typedef unsigned long ULONG, *PULONG;

typedef ULONG_PTR DWORD_PTR;
 typedef ULONG_PTR SIZE_T;
 typedef unsigned int ULONG32;
 typedef unsigned __int64 ULONG64;

#ifndef UNICODE
 //
 // There is a conflict with the "#define UNICODE" directive, which
 // affects the character set the Windows header files treat as default.
 //
 // For example, if you define UNICODE, then GetWindowText will map to
 // GetWindowTextW instead of GetWindowTextA.
 //
 // Note: 'typedef wchar_t WCHAR, *PWCHAR;' appears above, WCHAR being the
 //       appropriate typedef.
 //
 typedef wchar_t UNICODE;
 #endif // UNICODE

typedef unsigned short USHORT;

#ifndef _WINDOWS_
 //
 // Defined in:
 // Microsoft SDKs\Windows\v7.0a\Include\WinNT.h
 //
 typedef void VOID, *PVOID, *LPVOID;
 #endif // _WINDOWS_

typedef unsigned short WORD, *PWORD, *LPWORD;

#ifndef _WINDOWS_
 //
 // Defined in:
 // Microsoft SDKs\Windows\v7.0a\Include\WinBase.h
 //
 typedef struct _FILETIME {
   DWORD dwLowDateTime;
   DWORD dwHighDateTime;
 } FILETIME, 
  *PFILETIME, 
  *LPFILETIME;
 //
 // Defined in:
 // Microsoft SDKs\Windows\v7.0a\Include\RpcDce.h
 //
 typedef struct _GUID {
   unsigned long Data1;
   unsigned short Data2;
   unsigned short Data3;
   byte Data4[8];
 } GUID, 
   UUID, 
  *PGUID;
 //
 // Defined in:
 // Microsoft SDKs\Windows\v7.0a\Include\WinNT.h
 //
 typedef struct _LARGE_INTEGER {
     __int64 QuadPart;
 } LARGE_INTEGER, *PLARGE_INTEGER;
 #endif // _WINDOWS_

typedef DWORD LCID;

typedef struct _RPC_UNICODE_STRING {
   unsigned short Length;
   unsigned short MaximumLength;
 #ifdef MIDL_PASS
   [size_is(MaximumLength/2), length_is(Length/2)] 
 #endif
   WCHAR* Buffer;
 } RPC_UNICODE_STRING, 
  *PRPC_UNICODE_STRING;

#ifndef _WINDOWS_
 //
 // Defined in:
 // Microsoft SDKs\Windows\v7.0a\Include\WinBase.h
 //
 typedef struct _SYSTEMTIME {
   WORD wYear;
   WORD wMonth;
   WORD wDayOfWeek;
   WORD wDay;
   WORD wHour;
   WORD wMinute;
   WORD wSecond;
   WORD wMilliseconds;
 } SYSTEMTIME, 
  *PSYSTEMTIME;
 #endif // _WINDOWS_

typedef struct _UINT128 {
   UINT64 lower;
   UINT64 upper;
 } UINT128, 
  *PUINT128;

#ifndef _WINDOWS_
 //
 // Defined in:
 // Microsoft SDKs\Windows\v7.0a\Include\WinNT.h
 //
 typedef struct _ULARGE_INTEGER {
     unsigned __int64 QuadPart;
 } ULARGE_INTEGER, *PULARGE_INTEGER;
 #endif // _WINDOWS_

typedef struct _RPC_SID_IDENTIFIER_AUTHORITY {
   byte Value[6];
 } RPC_SID_IDENTIFIER_AUTHORITY;

typedef DWORD ACCESS_MASK; 
 typedef ACCESS_MASK *PACCESS_MASK;

typedef DWORD SECURITY_INFORMATION, *PSECURITY_INFORMATION;

typedef struct _RPC_SID {
   unsigned char Revision;
   unsigned char SubAuthorityCount;
   RPC_SID_IDENTIFIER_AUTHORITY IdentifierAuthority;
 #ifdef MIDL_PASS
   [size_is(SubAuthorityCount)] unsigned long SubAuthority[];
 #else
   unsigned long SubAuthority[ANYSIZE_ARRAY];
 #endif // MIDL_PASS
 } RPC_SID, 
  *PRPC_SID;
