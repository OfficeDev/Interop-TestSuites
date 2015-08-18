#include "ws2tcpip.h"
#include "winsock2.h"
#include "winsock.h"
#include "MS-OXCRPC.h"
#pragma   comment(lib,"ws2_32.lib")
#include <fstream>
#include <tchar.h>
 
void* __RPC_USER midl_user_allocate(size_t size);
void __RPC_USER midl_user_free(void* p);

static unsigned long inline Hash(const char *str);
unsigned long inline HandleException(RPC_STATUS status);

static RPC_BINDING_HANDLE m_hBind=NULL;
static CXH m_cxh=NULL;
static SEC_WINNT_AUTH_IDENTITY * m_swai=new SEC_WINNT_AUTH_IDENTITY;
static RPC_SECURITY_QOS_V2_W m_Qos;

unsigned long __stdcall Connect(CXH *pcxh,const char * szUserDN)
{
	unsigned long status = 0;
	unsigned long ulFlags = 0x00000000;
	unsigned long ulConMod =Hash(szUserDN);
	unsigned long cbLimit = 0;
	unsigned long ulCpid = 0x000004E4; 
	unsigned long ulLcidString = 0x00000409; 
	unsigned long ulLcidSort = 0x00000409; 
	unsigned long ulIcxrLink = 0xFFFFFFFF;
	unsigned short usFCanConvertCodePages = 0x01;

	unsigned long cmsPollsMax = 0;
	unsigned long cRetry = 0;
	unsigned long cmsRetryDelay = 0;
	unsigned short iCxr = 0;

	unsigned char * pszDNPrefix = NULL;
	unsigned char * pszDisplayName = NULL;

	unsigned short rgwClientVersion[3];
	unsigned short rgwServerVersion[3];
	unsigned short rgwBestVersion[3];

	rgwClientVersion[0] = 0x000c;
	rgwClientVersion[1] = 0x183e;
	rgwClientVersion[2] = 0x03e8;

	memset(rgwServerVersion, 0, sizeof(rgwServerVersion));
	memset(rgwBestVersion, 0, sizeof(rgwBestVersion));

	unsigned long ulTimeStamp = 0;

	unsigned char * pbAuxIn = NULL;
	unsigned long cbAuxIn = 0;

	unsigned char rgbAuxOut[0x1008];
	unsigned long cbAuxOut = 0x1008;

	memset(rgbAuxOut, 0, sizeof(rgbAuxOut));

	RpcTryExcept 
	{
	
		status = EcDoConnectEx(
			m_hBind,		//[in]
			pcxh,		//[out]
			(unsigned char*)szUserDN, //[in]
			ulFlags,	//[in]
			ulConMod,	//[in]
			cbLimit,	//[in] cbLimit
			ulCpid,		//[in] ulCpid
			ulLcidString,	//[in]
			ulLcidSort,		//[in]
			ulIcxrLink,		//[in]
			usFCanConvertCodePages,	//[in]
			&cmsPollsMax,	//[out]
			&cRetry,		//[out]
			&cmsRetryDelay, //[out]
			&iCxr,			//[out]
			&pszDNPrefix,	//[out]
			&pszDisplayName,//[out]
			rgwClientVersion,//[in]
			rgwServerVersion,//[out]
			rgwBestVersion,	//[out]
			&ulTimeStamp,	//[in,out]
			pbAuxIn,	//[in]
			cbAuxIn,	//[in]
			rgbAuxOut,	//[out]
			&cbAuxOut	//[in,out]
			);
		;
	}
	RpcExcept( HandleException(::RpcExceptionCode()) )
	{
		status = ::RpcExceptionCode();
	} 
	RpcEndExcept;

	return status;
}  

unsigned long __stdcall BindToServer(const char * server, int encryptionMethod, int authenticationServices, const char *seqType, bool rpchUseSsl, const char *rpchAuthScheme, const char *spnStr, const char *options, bool setUuid)
{
	unsigned long status = 0;
	RPC_WSTR sequence = NULL;
	RPC_WSTR endpoint = NULL;
	RPC_WSTR uuid = (RPC_WSTR)L"A4F1DB00-CA47-1067-B31F-00DD010662DA";
	RPC_WSTR serverPrincName = NULL;
	RPC_SECURITY_QOS_V2_W *qos = NULL;

	size_t origsize=strlen(server)+1;
	const size_t newsize=100;
	size_t convertedChars=0;
	wchar_t serverstring[newsize];
	mbstowcs_s(&convertedChars, serverstring, origsize, server, _TRUNCATE);

	const wchar_t *wcServer=serverstring;
	const wchar_t *wcOptions = NULL;

	wchar_t optionsstring[newsize];
	if(options)
	{
		mbstowcs_s(&convertedChars, optionsstring, origsize, options, _TRUNCATE);
		wcOptions = optionsstring;
	}

	if(_stricmp(seqType, "ncacn_http") == 0) // the endpoint is 6001 when protocol sequences is "ncacn_http", as specified in MS-OXCRPC section 2.1
	{
		const wchar_t * defaultSeq = L"ncacn_http";
		const wchar_t * defaultEndp = L"6001";
		rsize_t size = wcslen(defaultSeq)+ 1;
		sequence = (RPC_WSTR) new wchar_t[size];
		wcscpy_s((wchar_t *)sequence, size, defaultSeq);

		size = wcslen(defaultEndp)+ 1;
		endpoint = (RPC_WSTR) new wchar_t[size];
		wcscpy_s((wchar_t *)endpoint, size, defaultEndp);
	}
	else
	{
		origsize=strlen(seqType)+1;
		sequence = (RPC_WSTR) new wchar_t[origsize];
		mbstowcs_s(&convertedChars, (wchar_t *)sequence, origsize, seqType, _TRUNCATE);
	}
	
	if (authenticationServices == RPC_C_AUTHN_GSS_KERBEROS) 
	{
		origsize=strlen(spnStr)+1;
		wchar_t spnstring[newsize];
		mbstowcs_s(&convertedChars, spnstring, origsize, spnStr, _TRUNCATE);
		const wchar_t *wcSPN=spnstring;
		serverPrincName = (RPC_WSTR)wcSPN; 
	} 
	else 
	{ 
		serverPrincName = NULL; 
	}

	if (_stricmp(seqType, "ncacn_http") == 0 && rpchAuthScheme != NULL)
	{
		if (_stricmp(rpchAuthScheme, "Basic") == 0)
		{
			qos = &m_Qos;
			qos->u.HttpCredentials->AuthnSchemes[0] = RPC_C_HTTP_AUTHN_SCHEME_BASIC;
		}
		else if (_stricmp(rpchAuthScheme, "NTLM") == 0)
		{
			qos = &m_Qos;
			qos->u.HttpCredentials->AuthnSchemes[0] = RPC_C_HTTP_AUTHN_SCHEME_NTLM;
		}
		else
		{
			status = 1; // only Basic and NTLM http credential authentication are supported
		}
        
		if (rpchUseSsl)
		{
			qos->u.HttpCredentials->Flags = RPC_C_HTTP_FLAG_USE_SSL | RPC_C_HTTP_FLAG_USE_FIRST_AUTH_SCHEME;
		}
		else
		{
			qos->u.HttpCredentials->Flags = RPC_C_HTTP_FLAG_USE_FIRST_AUTH_SCHEME;	// remove SSL flag
		}
	}

	RPC_BINDING_HANDLE hReturnBinding = NULL;
	RPC_WSTR binding = NULL;
    if (status == 0)
    {
	    if (setUuid)
	    {
		    status = RpcStringBindingCompose(uuid, sequence, (RPC_WSTR)wcServer, endpoint, (RPC_WSTR)wcOptions, &binding);
	    }
	    else
	    {
		    status = RpcStringBindingCompose(NULL, sequence, (RPC_WSTR)wcServer, endpoint, (RPC_WSTR)wcOptions, &binding);
	    }

	    if (status == 0)
	    {
		    status = RpcBindingFromStringBinding( binding, &m_hBind);
			
		    if (status == 0)
		    {
			    status = RpcEpResolveBinding(m_hBind, emsmdb_v0_81_c_ifspec);
			    if (status == 0)
			    {
				    status = RpcBindingSetAuthInfoEx(m_hBind,
												    serverPrincName, // server principal name
												    encryptionMethod, // authentication level
												    authenticationServices,	// authentication service
												    m_swai,	// authorization Identity
												    0,		// authorization service
												    (RPC_SECURITY_QOS *)qos);	// quality of service
			    }
		    }
	    }
    }

	if (binding)
	{
		RpcStringFree(&binding);
	}
	
	if (endpoint)
	{
		delete endpoint;
	}

	if (sequence)
	{
		delete sequence;
	}
	return status;
}

void __stdcall CreateQos( )
{
	unsigned long *auth = new unsigned long[1];
	auth[0] = RPC_C_HTTP_AUTHN_SCHEME_NTLM;//RPC_C_HTTP_AUTHN_SCHEME_BASIC; //RPC_C_HTTP_AUTHN_SCHEME_NTLM;
	RPC_HTTP_TRANSPORT_CREDENTIALS_W *ssl = new RPC_HTTP_TRANSPORT_CREDENTIALS_W();

	ssl->TransportCredentials = m_swai;
	ssl->Flags = RPC_C_HTTP_FLAG_USE_SSL | RPC_C_HTTP_FLAG_USE_FIRST_AUTH_SCHEME;
	ssl->AuthenticationTarget = RPC_C_HTTP_AUTHN_TARGET_SERVER;
	ssl->NumberOfAuthnSchemes = 1;
	ssl->AuthnSchemes = (unsigned long *)auth;

	m_Qos.Version = RPC_C_SECURITY_QOS_VERSION_2;
	m_Qos.Capabilities = RPC_C_QOS_CAPABILITIES_DEFAULT;
	m_Qos.IdentityTracking = RPC_C_QOS_IDENTITY_DYNAMIC;
	m_Qos.ImpersonationType = RPC_C_IMP_LEVEL_IMPERSONATE;
	m_Qos.AdditionalSecurityInfoType = RPC_C_AUTHN_INFO_TYPE_HTTP;
	m_Qos.u.HttpCredentials = ssl;
}

void __stdcall CreateIdentity(const char * domain, const char * username, const char* password)
{
	size_t origsizeofdomain=strlen(domain)+1;
	size_t origsizeofusername=strlen(username)+1;
	size_t origsizeofpassword=strlen(password)+1;
	const size_t newsize=100;
	size_t convertedChars=0;
	wchar_t domainstring[newsize];
	wchar_t usernamestring[newsize];
	wchar_t passwordstring[newsize];
	mbstowcs_s(&convertedChars, domainstring, origsizeofdomain, domain, _TRUNCATE);
	mbstowcs_s(&convertedChars, usernamestring, origsizeofusername, username, _TRUNCATE);
	mbstowcs_s(&convertedChars, passwordstring, origsizeofpassword, password, _TRUNCATE);

	memset(m_swai, 0, sizeof(SEC_WINNT_AUTH_IDENTITY));
	if (domainstring && *domainstring)
	{
		m_swai->Domain = (unsigned short *)_tcsdup((const wchar_t*)domainstring);
		m_swai->DomainLength = (unsigned long)_tcslen((const wchar_t*)domainstring);
		
	}
	if (usernamestring && *usernamestring)
	{
		m_swai->User = (unsigned short *)_tcsdup((const wchar_t*)usernamestring);
		m_swai->UserLength = (unsigned long)_tcslen((const wchar_t*)usernamestring);
	}	
	if (passwordstring && *passwordstring)
	{
		m_swai->Password = (unsigned short *)_tcsdup((const wchar_t*)passwordstring);
		m_swai->PasswordLength = (unsigned long)_tcslen((const wchar_t*)passwordstring);
	}
	m_swai->Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;

	CreateQos();
} 

handle_t __stdcall GetBindHandle( )
{
	return m_hBind;	
}


// Memory allocation function for RPC.
void* __RPC_USER midl_user_allocate(size_t size)
{
    return malloc(size);
}

// Memory deallocation function for RPC.
void __RPC_USER midl_user_free(void* p)
{
    free(p);
}

PVOID CreateRpcAsyncHandle()
{
		RPC_ASYNC_STATE *pAsync = new RPC_ASYNC_STATE() ; 
			unsigned long status = RpcAsyncInitializeHandle(pAsync,sizeof(RPC_ASYNC_STATE));
				if(status)
				{
							return NULL;
				}
					pAsync->UserInfo = NULL;
						pAsync->NotificationType = RpcNotificationTypeNone;
							return pAsync;
}

static unsigned long inline Hash(const char *str)
{
	unsigned long value ;
	unsigned long i;
	unsigned long len;
	if(!str) return 0;
	len = strlen(str);
	for(value=0x238F13AF * len ,i =0;i<len;i++)
		value = (value +(str[i]<< (i*5%24)));
	return (1103515243 * value +12345);
}

unsigned long inline HandleException(RPC_STATUS status)
{
	if ((status & 0xc0000000) == 0xc0000000)
		return EXCEPTION_CONTINUE_SEARCH;
	else
		return EXCEPTION_EXECUTE_HANDLER;
}

/// <summary>
/// Asynchronous call that the server will not complete until there are pending events on the Session Context.
/// </summary>
/// <param name="acxh">A unique value to be used as an ACXH</param>
/// <param name="ulFlagsIn">Unused.Reserved for future use.Client MUST pass a value of 0x00000000.</param>
/// <param name="waitSecondThreshold">Indicates the threshold of waiting time in second</param>
/// <param name="makeEvent">It indicates whether client sends an event to server. 
/// True means client send an event to server. False means not.</param>
/// <param name="pulFlagsOut">Output flags for the client.</param>
/// <returns>If success, it returns 0, else returns the error code</returns>
long  __stdcall EcDoAsyncWaitExWrap(
    ACXH acxh, 
    unsigned long ulFlagsIn, 
    unsigned long waitSecondThreshold, 
    BOOL makeEvent, 
    unsigned long * pulFlagsOut
    )
{
    RPC_ASYNC_STATE async ; 
    RPC_STATUS stat;
    RPC_STATUS status;
    long reply = 0;
    unsigned int waitTime = 0;

    // Invoke RpcAsyncInitializeHandle function to initialize the RPC_ASYNC_STATE structure to be used to make an asynchronous call.
    status = RpcAsyncInitializeHandle(&async, sizeof(async));
    if(!status)
    {
        async.UserInfo = NULL;
        async.NotificationType = RpcNotificationTypeNone;
        EcDoAsyncWaitEx(&async,acxh, ulFlagsIn, pulFlagsOut);

        // Wait until the status of asynchronous remote procedure call is NOT RPC_S_ASYNC_CALL_PENDING or wait time reaches the wait second threshold.
        while(waitTime < waitSecondThreshold)
        {
            // Invoke RpcAsyncGetCallStatus function to determine the current status of an asynchronous remote call.
            stat = RpcAsyncGetCallStatus(&async);
            if (stat != RPC_S_ASYNC_CALL_PENDING )
                break;
            // 1000: means to wait 1000 millisecond (1 second)
            Sleep(1000);
            waitTime++;
        }
        // Invoke RpcAsyncCompleteCall function to complete an asynchronous remote procedure call.
        stat = RpcAsyncCompleteCall(&async,&reply);
        if(stat)
        {
            reply = 0x0000ffff; // error code, indicates failure in RpcAsyncCompleteCall
        }
    }
    else
    {
        reply = 0x000fffff; // error code, indicates failure in RpcAsyncInitializeHandle
    }
    return reply;
}

/// <summary>
/// The method EcRRegisterPushNotificationWrap registers a callback address with the server for a Session Context.
/// </summary>
/// <param name="pcxh">On input, the client MUST pass a valid CXH that was created by calling EcDoConnectEx</param>
/// <param name="family">Ip address family</param>
/// <param name="ip">Ip address</param>
/// <param name="port">Port number of a socket</param>
/// <param name="rgbContext">This parameter contains opaque client-generated context data that is sent back to the client at the callback address</param>
/// <param name="cbContext">This parameter contains the size of the opaque client context data that is passed in parameter rgbContext.</param>
/// <param name="hNotification">If the call completes successfully, this output parameter will contain a handle to the notification callback on the server</param>
/// <returns>If success, it returns 0, else returns the error code</returns>
long  __stdcall EcRRegisterPushNotificationWrap(
    CXH *pcxh,
    unsigned short family,
    char * ip,
    unsigned short port,
    unsigned char * rgbContext,
    unsigned short cbContext,
    unsigned long* hNotification
    )
{
    long status = 0xff; // error code, indicates failure in EcRRegisterPushNotification
    struct sockaddr_in saServer;
    struct sockaddr_storage ss;
    int sslen = sizeof(ss);
    unsigned long iRpc = 0; // The server MUST completely ignore this value. The client MUST pass a value of 0x00000000.(Refer to [MS-OXCRPC], section 3.1.4.5)
    unsigned long grbitAdviseBits = 0xFFFFFFFFul; // This parameter MUST be 0xFFFFFFFF.(Refer to [MS-OXCRPC], section 3.1.4.5)
    unsigned char* rgbCallbackAddress = NULL; // Contains the callback address for the server to use to notify the client of a pending event.(Refer to [MS-OXCRPC], section 3.1.4.5)
    unsigned short cbCallbackAddress = 0; // The length of the callback address.(Refer to [MS-OXCRPC], section 3.1.4.5)

    // Initialize structure sockaddr_storage
    WSAStringToAddressA(ip, AF_INET6, NULL, (struct sockaddr*)&ss, &sslen);
    ((struct sockaddr_in6 *)&ss)->sin6_port = port;
    ((struct sockaddr_in6 *)&ss)->sin6_family = family;

    // Initialize structure sockaddr_in
    saServer.sin_family = family;
    saServer.sin_addr.s_addr = (inet_addr(ip));
    saServer.sin_port = htons(port);
    memset(saServer.sin_zero, 0, sizeof(saServer.sin_zero));

    // The server supports the address families AF_INET and AF_INET6 for a callback address.(Refer to [MS-OXCRPC], section 3.1.4.5)
    if(family == AF_INET6)
    {
        cbCallbackAddress = sizeof(sockaddr_in6);
        rgbCallbackAddress = (unsigned char*)malloc(cbCallbackAddress);
        memcpy(rgbCallbackAddress, &ss, cbCallbackAddress);
    }
    else
    {
        cbCallbackAddress = sizeof(sockaddr_in);
        rgbCallbackAddress = (unsigned char*)malloc(cbCallbackAddress);
        memcpy(rgbCallbackAddress, &saServer, cbCallbackAddress);
    }

    // Invoke EcRRegisterPushNotification to register a callback address with the server for a Session Context.
    status = EcRRegisterPushNotification(
        pcxh,
        iRpc,
        rgbContext,
        cbContext,
        grbitAdviseBits,
        rgbCallbackAddress,
        cbCallbackAddress,
        hNotification);

    return status;
}