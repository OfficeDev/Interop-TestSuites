//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.Protocols.TestTools.Messages.Marshaling;

    /// <summary>
    /// The NspiInterop class exposes the methods of the unmanaged library.
    /// </summary>
    public static class NspiInterop
    {
        /// <summary>
        /// The library compiled for native code to accomplish RPC communication with server.
        /// </summary>
        private const string DllName = "MS-OXORULE_NspiStub.dll";

        #region Unmanaged RPC Calls
        /// <summary>
        /// The NspiBind method initiates a session between a client and the server.
        /// </summary>
        /// <param name="ptrRpc">An RPC binding handle parameter.</param>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A pointer to a STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="ptrServerGuid">The value NULL or a pointer to a GUID value that is associated with the specific server.</param>
        /// <param name="contextHandle">An NSPI context handle.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(DllName, EntryPoint = "NspiBind", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiBind(IntPtr ptrRpc, uint flags, ref STAT stat, IntPtr ptrServerGuid, ref IntPtr contextHandle);

        /// <summary>
        /// The NspiUnbind method destroys the context handle. No other action is taken.
        /// </summary>
        /// <param name="contextHandle">The NSPI context handle to be destroyed.</param>
        /// <param name="reserved">A DWORD value reserved for future use. This property is ignored by the server.</param>
        /// <returns>The server returns a DWORD value that specifies the return status of the method.</returns>
        [DllImport(DllName, EntryPoint = "NspiUnbind", CallingConvention = CallingConvention.Cdecl)]
        public static extern uint NspiUnbind(ref IntPtr contextHandle, uint reserved);

        /// <summary>
        /// The NspiQueryRows method returns a number of rows from a specified table to the client.
        /// </summary>
        /// <param name="contextHandle">An NSPI context handle.</param>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="tableCount">A DWORD value that contains the number values in the input parameter lpETable. 
        /// This value is limited to 100,000.</param>
        /// <param name="table">An array of DWORD values, representing an Explicit Table.</param>
        /// <param name="count">A DWORD value that contains the number of rows the client is requesting.</param>
        /// <param name="ptrPropTags">The value NULL or a reference to a PropertyTagArray_r value, 
        /// containing a list of the proptags of the properties that the client requires to be returned for each row returned.</param>
        /// <param name="ptrRows">A reference to a PropertyRowSet_r value,
        /// Contains the address book container rows that the server returns in response to the request.c</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(DllName, EntryPoint = "NspiQueryRows", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiQueryRows(IntPtr contextHandle, uint flags, ref STAT stat, uint tableCount, [Size("dwETableCount")]uint[] table, uint count, IntPtr ptrPropTags, out IntPtr ptrRows);

        /// <summary>
        /// This method binds client to RPC server.
        /// </summary>
        /// <param name="serverName">Representation of a network address of server.</param>
        /// <param name="userName">The user name.</param>
        /// <param name="domain">The domain or workgroup name.</param>
        /// <param name="password">The user's password in the domain or workgroup.</param>
        /// <param name="stringBinding">The RPC binding handle string to be free.</param>
        /// <returns>The created RPC binding handle.</returns>
        [DllImport(DllName, CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Ansi, ExactSpelling = true)]
        public static extern IntPtr CreateRpcBinding(IntPtr serverName, IntPtr userName, IntPtr domain, IntPtr password, IntPtr stringBinding);

        /// <summary>
        /// Destroy the created RPC binding handle.
        /// </summary>
        /// <param name="bindingHandle">The RPC binding handle to be destroyed.</param>
        /// <param name="stringBinding">The RPC binding handle string to be free.</param>
        /// <returns>Status of RPC binding handle free. The non-zero return value indicates failed to free RPC binding handle.</returns>
        [DllImport(DllName, CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Ansi, ExactSpelling = true)]
        public static extern uint FreeRpcBinding(IntPtr bindingHandle, IntPtr stringBinding);

        #endregion
    }
}
