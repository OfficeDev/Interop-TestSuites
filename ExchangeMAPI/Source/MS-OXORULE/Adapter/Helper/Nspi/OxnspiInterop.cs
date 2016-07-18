namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.Protocols.TestTools.Messages.Marshaling;

    /// <summary>
    /// Class OxnspiInterop exposes the methods of the unmanaged library  by importing the DLL RPC_RuntimeDllName.
    /// </summary>
    public class OxnspiInterop
    {
        /// <summary>
        /// The RPC run time dll name.
        /// </summary>
        private const string RPCRuntimeDllName = "rpcrt4.dll";

        /// <summary>
        /// The MS-OXNSPI Stub dll to be called.
        /// </summary>
        private const string MSOXNSPIDLL = "MS-OXNSPI_Stub";

        #region Unmanaged RPC Calls
        /// <summary>
        /// The NspiBind method initiates a session between a client and the server.
        /// </summary>
        /// <param name="ptrRpc">An RPC binding handle parameter.</param>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A pointer to a STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="ptrServerGuid">The value NULL or a pointer to a GUID value that is associated with the specific server.</param>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiBind", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiBind(IntPtr ptrRpc, uint flags, ref STAT stat, IntPtr ptrServerGuid, ref IntPtr contextHandle);

        /// <summary>
        /// The NspiUnbind method destroys the context handle. No other action is taken.
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="reserved">A DWORD value reserved for future use. This property is ignored by the server.</param>
        /// <returns>The server returns a DWORD value that specifies the return status of the method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiUnbind", CallingConvention = CallingConvention.Cdecl)]
        public static extern uint NspiUnbind(ref IntPtr contextHandle, uint reserved);

        /// <summary>
        /// The NspiGetSpecialTable method returns the rows of a special table to the client. 
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A pointer to a STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="version">A reference to a DWORD. On input, it holds the value of the version number of
        /// the address book hierarchy table that the client has.</param>
        /// <param name="ptrRows">A PropertyRowSet_r structure. On return, it holds the rows for the table that the client is requesting.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiGetSpecialTable", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiGetSpecialTable(IntPtr contextHandle, uint flags, ref STAT stat, ref uint version, out IntPtr ptrRows);

        /// <summary>
        /// The NspiUpdateStat method updates the STAT block that represents the position in a table 
        /// to reflect positioning changes requested by the client.
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="reserved">A DWORD value. Reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A STAT block describing a logical position in a specific address book container.</param>
        /// <param name="ptrDelta">The value NULL or a pointer to a LONG value that indicates movement 
        /// within the address book container specified by the input parameter stat.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiUpdateStat", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiUpdateStat(IntPtr contextHandle, uint reserved, ref STAT stat, IntPtr ptrDelta);

        /// <summary>
        /// The NspiQueryColumns method returns a list of all the properties that the server is aware of. 
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="ptrColumns">A reference to a PropertyTagArray_r structure. On return, it contains a list of proptags.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiQueryColumns", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiQueryColumns(IntPtr contextHandle, uint reserved, uint flags, out IntPtr ptrColumns);

        /// <summary>
        /// The NspiGetPropList method returns a list of all the properties that have values on a specified object.
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="mid">A DWORD value that contains a Minimal Entry ID.</param>
        /// <param name="codePage">The code page in which the client wants the server to express string values properties.</param>
        /// <param name="ptrPropTags">A PropertyTagArray_r value. On return, it holds a list of properties.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiGetPropList", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiGetPropList(IntPtr contextHandle, uint flags, uint mid, uint codePage, out IntPtr ptrPropTags);

        /// <summary>
        /// The NspiGetProps method returns an address book row that contains a set of the properties
        /// and values that exist on an object.
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="ptrPropTags">The value NULL or a reference to a PropertyTagArray_r value. 
        /// It contains a list of the proptags of the properties that the client wants to be returned.</param>
        /// <param name="ptrRows">A reference to a PropertyRow_r value. 
        /// It contains the address book container row the server returns in response to the request.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiGetProps", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiGetProps(IntPtr contextHandle, uint flags, ref STAT stat, IntPtr ptrPropTags, out IntPtr ptrRows);

        /// <summary>
        /// The NspiQueryRows method returns a number of rows from a specified table to the client.
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
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
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiQueryRows", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiQueryRows(IntPtr contextHandle, uint flags, ref STAT stat, uint tableCount, [Size("dwETableCount")]uint[] table, uint count, IntPtr ptrPropTags, out IntPtr ptrRows);

        /// <summary>
        /// The NspiSeekEntries method searches for and sets the logical position in a specific table
        /// to the first entry greater than or equal to a specified value. 
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="reserved">A DWORD value that is reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="ptrTarget">A PropertyValue_r value holding the value that is being sought.</param>
        /// <param name="table">The value NULL or a PropertyTagArray_r value. 
        /// It holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="ptrPropTags">Contains list of the proptags of the columns 
        /// that the client wants to be returned for each row returned.</param>
        /// <param name="ptrRows">It contains the address book container rows the server returns in response to the request.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiSeekEntries", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiSeekEntries(IntPtr contextHandle, uint reserved, ref STAT stat, IntPtr ptrTarget, IntPtr table, IntPtr ptrPropTags, out IntPtr ptrRows);

        /// <summary>
        /// The NspiGetMatches method returns an Explicit Table. 
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="reserved">A DWORD value reserved for future use.</param>
        /// <param name="stat">A STAT block describing a logical position in a specific address book container.</param>
        /// <param name="ptrReserved">A PropertyTagArray_r reserved for future use.</param>
        /// <param name="reserved2">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="ptrFilter">The value NULL or a reference to Restriction_r array value. It holds a logical restriction to apply to the rows in the address book container specified in the stat parameter.</param>
        /// <param name="ptrPropName">The value NULL or a PropertyName_r value. It holds the property to be opened as a restricted address book container.</param>
        /// <param name="requested">A DWORD value. It contains the maximum number of rows to return in a restricted address book container.</param>
        /// <param name="ptrOutMids">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="ptrPropTags">The value NULL or a reference to a PropertyTagArray_r value. It contains a list of the proptags of the columns that the client wants to be returned for each row returned.</param>
        /// <param name="ptrRows">A reference to a PropertyRowSet_r value. It contains the address book container rows the server returns in response to the request.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiGetMatches", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiGetMatches(IntPtr contextHandle, uint reserved, ref STAT stat, IntPtr ptrReserved, uint reserved2, IntPtr ptrFilter, IntPtr ptrPropName, uint requested, out IntPtr ptrOutMids, IntPtr ptrPropTags, out IntPtr ptrRows);

        /// <summary>
        /// The NspiResortRestriction method applies to a sort order to the objects in a restricted address book container.
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A reference to a STAT block describing a logical position in a specific address book container.</param>
        /// <param name="ptrInmids">A reference to a PropertyTagArray_r array. It holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="ptrOutMIds">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs 
        /// that comprise a restricted address book container.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiResortRestriction", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiResortRestriction(IntPtr contextHandle, uint reserved, ref STAT stat, IntPtr ptrInmids, ref IntPtr ptrOutMIds);

        /// <summary>
        /// The NspiCompareMIds method compares the position in an address book container of two objects 
        /// identified by Minimal Entry ID and returns the value of the comparison.
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="mid1">The mid1 is a DWORD value containing a Minimal Entry ID.</param>
        /// <param name="mid2">The mid2 is a DWORD value containing a Minimal Entry ID.</param>
        /// <param name="result">A DWORD value. On return, it contains the result of the comparison.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiCompareMIds", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiCompareMIds(IntPtr contextHandle, uint reserved, ref STAT stat, uint mid1, uint mid2, out int result);

        /// <summary>
        /// The NspiDNToMId method maps a set of DN to a set of Minimal Entry ID.
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="ptrNames">A StringsArray_r value. It holds a list of strings that contain DNs.</param>
        /// <param name="ptrMIds">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiDNToMId", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiDNToMId(IntPtr contextHandle, uint reserved, IntPtr ptrNames, out IntPtr ptrMIds);

        /// <summary>
        /// The NspiModProps method is used to modify the properties of an object in the address book. 
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="reserved">A DWORD value reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="ptrPropTags">The value NULL or a reference to a PropertyTagArray_r. 
        /// It contains a list of the proptags of the columns from which the client requests all the values to be removed. </param>
        /// <param name="ptrRow">A PropertyRow_r value. It contains an address book row.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiModProps", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiModProps(IntPtr contextHandle, uint reserved, ref STAT stat, IntPtr ptrPropTags, IntPtr ptrRow);

        /// <summary>
        /// The NspiModLinkAtt method modifies the values of a specific property of a specific row in the address book.
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="propTag">A DWORD value. It contains the proptag of the property that the client wants to modify.</param>
        /// <param name="mid">A DWORD value that contains the Minimal Entry ID of the address book row that the client wants to modify.</param>
        /// <param name="ptrEntryIds">A BinaryArray value. It contains a list of EntryIDs to be used to modify the requested property on the requested address book row.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiModLinkAtt", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiModLinkAtt(IntPtr contextHandle, uint flags, uint propTag, uint mid, IntPtr ptrEntryIds);

        /// <summary>
        /// The NspiResolveNames method takes a set of string values in an 8-bit character set 
        /// and performs ANR on those strings. 
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="reserved">A DWORD reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="ptrPropTags">The value NULL or a reference to a PropertyTagArray_r value containing a list of the proptags of the columns 
        /// that the client requests to be returned for each row returned.</param>
        /// <param name="ptrStrArray">A StringsArray_r value. It specifies the values on which the client is requesting the server to do ANR.</param>
        /// <param name="ptrMIds">A PropertyTagArray_r value. On return, it contains a list of Minimal Entry IDs that match the array of strings.</param>
        /// <param name="ptrRows">A reference to a PropertyRowSet_r value. 
        /// It contains the address book container rows that the server returns in response to the request.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiResolveNames", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiResolveNames(IntPtr contextHandle, uint reserved, [StaticSize(1, StaticSizeMode.Elements)] STAT[] stat, IntPtr ptrPropTags, IntPtr ptrStrArray, out IntPtr ptrMIds, out IntPtr ptrRows);

        /// <summary>
        /// The NspiResolveNamesW method takes a set of string values in the Unicode character set 
        /// and performs ANR on those strings. 
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="reserved">A DWORD value that is reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="ptrPropTags">The value NULL or a reference to a PropertyTagArray_r containing a list of the proptags of the columns 
        /// that the client requests to be returned for each row returned.</param>
        /// <param name="ptrWStrArray">A WStringsArray_r value. It specifies the values on which the client is requesting the server to perform ANR.</param>
        /// <param name="ptrMIds">A PropertyTagArray_r value. On return, it contains a list of Minimal Entry IDs that match the array of strings.</param>
        /// <param name="ptrRows">A reference to a PropertyRowSet_r structure. 
        /// It contains the address book container rows that the server returns in response to the request.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiResolveNamesW", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiResolveNamesW(IntPtr contextHandle, uint reserved, [StaticSize(1, StaticSizeMode.Elements)] STAT[] stat, IntPtr ptrPropTags, IntPtr ptrWStrArray, out IntPtr ptrMIds, out IntPtr ptrRows);

        /// <summary>
        /// The NspiGetTemplateInfo method returns the information about template objects.
        /// </summary>
        /// <param name="contextHandle">An RPC context handle.</param>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="type">A DWORD value. It specifies the display type of the template for which the information is requested.</param>
        /// <param name="dn">The value NULL or the DN of the template requested. The value is NULL-terminated.</param>
        /// <param name="codePage">A DWORD value. It specifies the code page of the template for which the information is requested.</param>
        /// <param name="localeID">A DWORD value. It specifies the LCID of the template for which the information is requested.</param>
        /// <param name="ptrData">A reference to a PropertyRow_r value. On return, it contains the information requested.</param>
        /// <returns>Status of NSPI method.</returns>
        [DllImport(MSOXNSPIDLL, EntryPoint = "NspiGetTemplateInfo", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NspiGetTemplateInfo(IntPtr contextHandle, uint flags, uint type, string dn, uint codePage, uint localeID, out IntPtr ptrData);

        /// <summary>
        /// This method binds client to RPC server.
        /// </summary>
        /// <param name="serverName">Representation of a network address of server.</param>
        /// <param name="encryptionMethod">The encryption method in this call.</param>
        /// <param name="authnSvc">Authentication service to use.</param>
        /// <param name="seqType">Transport sequence type.</param>
        /// <param name="rpchUseSsl">True to use RPC over HTTP with SSL, false to use RPC over HTTP without SSL.</param>
        /// <param name="rpchAuthScheme">The authentication scheme used in the http authentication for RPC over HTTP. This value can be "Basic" or "NTLM".</param>
        /// <param name="spnStr">Service Principal Name (SPN) string used in Kerberos SSP.</param>
        /// <param name="options">Proxy attribute.</param>
        /// <param name="setUuid">True to set PFC_OBJECT_UUID (0x80) field of RPC header, false to not set this field.</param>
        /// <returns>Binding status. The non-zero return value indicates failed binding.</returns>
        [DllImport(MSOXNSPIDLL, CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Ansi, ExactSpelling = true)]
        public static extern uint BindToServer(string serverName, uint encryptionMethod, uint authnSvc, string seqType, [MarshalAs(UnmanagedType.Bool)]bool rpchUseSsl, string rpchAuthScheme, string spnStr, string options, [MarshalAs(UnmanagedType.Bool)]bool setUuid);

        /// <summary>
        /// Create SEC_WINNT_AUTH_IDENTITY structure in native codes that enables passing a 
        /// particular user name and password to the run-time library for the purpose of authentication.
        /// </summary>
        /// <param name="domain">The domain or workgroup name.</param>
        /// <param name="userName">The user name.</param>
        /// <param name="password">The user's password in the domain or workgroup.</param>
        [DllImport(MSOXNSPIDLL, CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Ansi, ExactSpelling = true)]
        public static extern void CreateIdentity(string domain, string userName, string password);

        /// <summary>
        /// Return the current binding handle.
        /// </summary>
        /// <returns>Current binding handle.</returns>
        [DllImport(MSOXNSPIDLL, CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Ansi, ExactSpelling = true)]
        public static extern IntPtr GetBindHandle();
        #endregion

        #region RPC Runtime Methods
        /// <summary>
        /// Free handle
        /// </summary>
        /// <param name="binding">A pointer to the server binding handle</param>
        /// <returns>Returns an integer value to indicate call success or failure.</returns>
        [DllImport(RPCRuntimeDllName)]
        internal static extern uint RpcBindingFree(
            ref IntPtr binding);
        #endregion
    }
}