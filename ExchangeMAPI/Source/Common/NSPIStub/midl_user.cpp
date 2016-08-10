#include "ws2tcpip.h"
#include "winsock2.h"
#include "winsock.h"
#include "MS-OXNSPI.h"
#pragma   comment(lib,"ws2_32.lib")
#include <fstream>
#include <tchar.h>
 
/// <summary>
/// The RPC binding handle to be created.
/// </summary>
static RPC_BINDING_HANDLE m_hBind=NULL;

/// <summary>
/// User Identity used in the RPC binding request.
/// </summary>
static SEC_WINNT_AUTH_IDENTITY * m_swai=new SEC_WINNT_AUTH_IDENTITY;

/// <summary>
/// Security quality-of-service settings used in the RPC binding request.
/// </summary>
static RPC_SECURITY_QOS_V2_W m_Qos;

/// <summary>
/// This method binds client to RPC server.
/// </summary>
/// <param name="server">Representation of a network address of server.</param>
/// <param name="encryptionMethod">The encryption method in this call.</param>
/// <param name="authenticationServices">Authentication service to use.</param>
/// <param name="seqType">Transport sequence type.</param>
/// <param name="rpchUseSsl">True to use RPC over HTTP with SSL, false to use RPC over HTTP without SSL.</param>
/// <param name="rpchAuthScheme">The authentication scheme used in the http authentication for RPC over HTTP. This value can be "Basic" or "NTLM".</param>
/// <param name="spnStr">Service Principal Name (SPN) string used in Kerberos SSP.</param>
/// <param name="options">Proxy attribute.</param>
/// <param name="setUuid">True to set PFC_OBJECT_UUID (0x80) field of RPC header, false to not set this field.</param>
/// <returns>Binding status. The non-zero return value indicates failed binding.</returns>
unsigned long __stdcall BindToServer(const char * server, int encryptionMethod, int authenticationServices, const char *seqType, bool rpchUseSsl, const char *rpchAuthScheme, const char *spnStr, const char *options, bool setUuid)
{
	unsigned long status = 0;
	RPC_WSTR sequence = NULL;
	RPC_WSTR endpoint = NULL;
	RPC_WSTR uuid = (RPC_WSTR)L"F5CC5A18-4264-101A-8C59-08002B2F8426"; // The uuid is specified in MS-OXNSPI section 2.1.
	RPC_WSTR serverPrincName = NULL;
	RPC_SECURITY_QOS_V2_W *qos = NULL;

	size_t origsize=strlen(server)+1;
	const size_t newsize=100;
	size_t convertedChars=0;
	wchar_t serverstring[newsize];
	mbstowcs_s(&convertedChars, serverstring, origsize, server, _TRUNCATE); // Convert server network address parameter from multibyte characters to wide characters.

	const wchar_t *wcServer=serverstring;
	const wchar_t *wcOptions = NULL;

	wchar_t optionsstring[newsize];
	if(options)
	{
		mbstowcs_s(&convertedChars, optionsstring, origsize, options, _TRUNCATE); 
		wcOptions = optionsstring;
	}

	if(_stricmp(seqType, "ncacn_http") == 0) 
	{
		const wchar_t * defaultSeq = L"ncacn_http";
		const wchar_t * defaultEndp = L"6004"; // The endpoint is 6004 when protocol sequences is "ncacn_http", which is a well-known endpoint used by MS-OXNSPI.
		rsize_t size = wcslen(defaultSeq)+ 1;
		sequence = (RPC_WSTR) new wchar_t[size];
		wcscpy_s((wchar_t *)sequence, size, defaultSeq); // Create wide characters transport sequence parameter. 

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
			status = 1; // Only Basic and NTLM http credential authentication are supported.
		}
        
		if (rpchUseSsl)
		{
			qos->u.HttpCredentials->Flags = RPC_C_HTTP_FLAG_USE_SSL | RPC_C_HTTP_FLAG_USE_FIRST_AUTH_SCHEME;
		}
		else
		{
			qos->u.HttpCredentials->Flags = RPC_C_HTTP_FLAG_USE_FIRST_AUTH_SCHEME;	// Remove SSL flag.
		}
	}

	RPC_BINDING_HANDLE hReturnBinding = NULL;
	RPC_WSTR binding = NULL;
    if (status == 0)
    {
		// Create a string representation binding handle.
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
		    status = RpcBindingFromStringBinding( binding, &m_hBind); // Get a binding handle from the created string representation binding handle.
			
		    if (status == 0)
		    {
			    status = RpcEpResolveBinding(m_hBind, nspi_v56_0_c_ifspec); // Resolve the created partially-bound server binding handle into a fully-bound server binding handle.
			    if (status == 0)
			    {
					// Set the authentication, authorization, and security quality-of-service information of the created binding handle.
				    status = RpcBindingSetAuthInfoEx(m_hBind,
												    serverPrincName, // Server principal name.
												    encryptionMethod, // Authentication level.
												    authenticationServices,	// Authentication service.
												    m_swai,	// Authorization Identity.
												    0,		// Authorization service.
												    (RPC_SECURITY_QOS *)qos);	// Security quality-of-service settings.
			    }
		    }
	    }
    }

	if (binding)
	{
		RpcStringFree(&binding); // Free the string representation binding handle allocated by the RPC run-time library.
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

/// <summary>
/// This method creates the security quality-of-service settings information used in the RPC binding request.
/// </summary>
void __stdcall CreateQos( )
{
	RPC_HTTP_TRANSPORT_CREDENTIALS_W *ssl = new RPC_HTTP_TRANSPORT_CREDENTIALS_W(); // This structure contains additional credentials to authenticate to an RPC proxy server when using RPC/HTTP.
	ssl->TransportCredentials = m_swai; // Add the created user identity as transport credential.
	ssl->Flags = RPC_C_HTTP_FLAG_USE_SSL | RPC_C_HTTP_FLAG_USE_FIRST_AUTH_SCHEME; // These two flags are set by default to indicate SSL is used and the first scheme in the AuthnSchemes array will be used.
	ssl->AuthenticationTarget = RPC_C_HTTP_AUTHN_TARGET_SERVER; // The default authentication target is set to 1 means the authentication against the RPC Proxy.
	ssl->NumberOfAuthnSchemes = 1; // The number of elements in the AuthnScheme array.
	unsigned long *auth = new unsigned long[1];
	auth[0] = RPC_C_HTTP_AUTHN_SCHEME_NTLM; // The default authentication scheme is NTLM.
	ssl->AuthnSchemes = (unsigned long *)auth; // Set an array of authentication schemes the client is willing to use.

	m_Qos.Version = RPC_C_SECURITY_QOS_VERSION_2; // The default version is set to 2.
	m_Qos.Capabilities = RPC_C_QOS_CAPABILITIES_DEFAULT; // The default capabilities is set to 0 means no provider-specific capabilities are needed.
	m_Qos.IdentityTracking = RPC_C_QOS_IDENTITY_DYNAMIC; // The default context tracking mode is set to 1 means context is revised whenever the ModifiedId in the client's token is changed.
	m_Qos.ImpersonationType = RPC_C_IMP_LEVEL_IMPERSONATE; // The default impersonation level is set to 3 means server can impersonate the client's security context on its local system, but not on remote systems.
	m_Qos.AdditionalSecurityInfoType = RPC_C_AUTHN_INFO_TYPE_HTTP; // This field is set 1 means the HttpCredentials member of the u union in this structure points to a RPC_HTTP_TRANSPORT_CREDENTIALS structure.
	m_Qos.u.HttpCredentials = ssl; // Additional credentials to pass to RPC run-time library. Used only when the AdditionalSecurityInfoType member is set to RPC_C_AUTHN_INFO_TYPE_HTTP.
}

/// <summary>
/// Create SEC_WINNT_AUTH_IDENTITY structure that enables passing a particular user name and password to the RPC run-time library for the purpose of authentication.
/// </summary>
/// <param name="domain">The domain or workgroup name.</param>
/// <param name="userName">The user name.</param>
/// <param name="password">The user's password in the domain or workgroup.</param>
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
	mbstowcs_s(&convertedChars, domainstring, origsizeofdomain, domain, _TRUNCATE); // Convert domain parameter from multibyte characters to wide characters.
	mbstowcs_s(&convertedChars, usernamestring, origsizeofusername, username, _TRUNCATE); 
	mbstowcs_s(&convertedChars, passwordstring, origsizeofpassword, password, _TRUNCATE); 

	memset(m_swai, 0, sizeof(SEC_WINNT_AUTH_IDENTITY)); // Set all fields of the buffer to store the m_swai structure to 0.
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

	// The default encoding of the created SEC_WINNT_AUTH_IDENTITY structure is Unicode.
	m_swai->Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE; 

	// Create the security quality-of-service settings information.
	CreateQos();
} 

/// <summary>
/// Return the current RPC binding handle.
/// </summary>
/// <returns>Current RPC binding handle.</returns>
handle_t __stdcall GetBindHandle( )
{
	return m_hBind;	
}

/// <summary>
/// Memory allocation function for RPC.
/// </summary>
void* __RPC_USER midl_user_allocate(size_t size)
{
    return malloc(size);
}

/// <summary>
/// Memory deallocation function for RPC.
/// </summary>
void __RPC_USER midl_user_free(void* p)
{
    free(p);
}