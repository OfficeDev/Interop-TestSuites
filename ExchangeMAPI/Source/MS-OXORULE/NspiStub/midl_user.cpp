#include "nspi.h"

/// <summary>
/// This method binds client to RPC server.
/// </summary>
/// <param name="serverName">Representation of a network address of server.</param>
/// <param name="userName">The user name.</param>
/// <param name="domain">The domain or workgroup name.</param>
/// <param name="password">The user's password in the domain or workgroup.</param>
/// <param name="stringBinding">The RPC binding handle string to be free.</param>
/// <returns>The created RPC binding handle.</returns>
RPC_BINDING_HANDLE __stdcall CreateRpcBinding(unsigned short* server, unsigned short* userName, unsigned short* domain, unsigned short* password, unsigned short *StringBinding)
{	
	unsigned long status;
	RPC_BINDING_HANDLE BindingHandle;
	wchar_t * ServerName = server;
	SEC_WINNT_AUTH_IDENTITY AuthIdentity;
	AuthIdentity.User = userName;
	AuthIdentity.UserLength = wcslen(userName);
	AuthIdentity.Domain = domain;
	AuthIdentity.DomainLength = wcslen(domain);
	AuthIdentity.Password = password;
	AuthIdentity.PasswordLength = wcslen(password);
	AuthIdentity.Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;

	status = RpcStringBindingCompose(NULL,  // Object UUID
		L"ncacn_ip_tcp",           // Protocol sequence to use
		ServerName, // Server DNS or Netbios Name
		NULL,
		NULL,
		&StringBinding);

	if(status != RPC_S_OK)
	{
		return 0;
	}

	// Error checking omitted. If no error, proceed below status =99;
	status = RpcBindingFromStringBinding(StringBinding, &BindingHandle);

	if(status != RPC_S_OK)
	{
		return 0;
	}

	status = RpcBindingSetAuthInfo(
		BindingHandle,
		ServerName,
		RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
		RPC_C_AUTHN_WINNT,
		&AuthIdentity,
		RPC_C_AUTHN_WINNT);

	return BindingHandle;
}

/// <summary>
/// Destroy the created RPC binding handle.
/// </summary>
/// <param name="bindingHandle">The RPC binding handle to be destroyed.</param>
/// <param name="stringBinding">The RPC binding handle string to be free.</param>
/// <returns>Status of RPC binding handle free. The non-zero return value indicates failed to free RPC binding handle.</returns>
unsigned long __stdcall FreeRpcBinding(RPC_BINDING_HANDLE bindingHandle,unsigned short* stringBinding)
{
	unsigned long status = RpcStringFree(&stringBinding); 

	status = RpcBindingFree(&bindingHandle);

	return status;
}