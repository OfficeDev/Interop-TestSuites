//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OXNSPI Adapter Interface
    /// </summary>
    public interface IMS_OXNSPIAdapter : IAdapter
    {
        /// <summary>
        /// The NspiBind method initiates a session between a client and the server.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A pointer to a STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="serverGuid">The value NULL or a pointer to a GUID value that is associated with the specific server.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiBind(uint flags, STAT stat, ref FlatUID_r? serverGuid, bool needRetry = true);

        /// <summary>
        /// The NspiUnbind method destroys the context handle. No other action is taken.
        /// </summary>
        /// <param name="reserved">A DWORD [MS-DTYP] value reserved for future use. This property is ignored by the server.</param>
        /// <returns>A DWORD value that specifies the return status of the method.</returns>
        uint NspiUnbind(uint reserved);

        /// <summary>
        /// The NspiGetSpecialTable method returns the rows of a special table to the client. 
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A pointer to a STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="version">A reference to a DWORD. On input, it holds the value of the version number of
        /// the address book hierarchy table that the client has. On output, it holds the version of the server's address book hierarchy table.</param>
        /// <param name="rows">A PropertyRowSet_r structure. On return, it holds the rows for the table that the client is requesting.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiGetSpecialTable(uint flags, ref STAT stat, ref uint version, out PropertyRowSet_r? rows, bool needRetry = true);

        /// <summary>
        /// The NspiUpdateStat method updates the STAT block that represents position in a table 
        /// to reflect positioning changes requested by the client.
        /// </summary>
        /// <param name="reserved">A DWORD value. Reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A pointer to a STAT block describing a logical position in a specific address book container.</param>
        /// <param name="delta">The value NULL or a pointer to a LONG value that indicates movement 
        /// within the address book container specified by the input parameter stat.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiUpdateStat(uint reserved, ref STAT stat, ref int? delta, bool needRetry = true);

        /// <summary>
        /// The NspiQueryColumns method returns a list of all the properties that the server is aware of. 
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="columns">A PropertyTagArray_r structure that contains a list of proptags.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiQueryColumns(uint reserved, uint flags, out PropertyTagArray_r? columns, bool needRetry = true);

        /// <summary>
        /// The NspiGetPropList method returns a list of all the properties that have values on a specified object.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="mid">A DWORD value that contains a Minimal Entry ID.</param>
        /// <param name="codePage">The code page in which the client wants the server to express string values properties.</param>
        /// <param name="propTags">A PropertyTagArray_r value. On return, it holds a list of properties.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiGetPropList(uint flags, uint mid, uint codePage, out PropertyTagArray_r? propTags, bool needRetry = true);

        /// <summary>
        /// The NspiGetProps method returns an address book row that contains a set of the properties
        /// and values that exist on an object.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r value. 
        /// It contains a list of the proptags of the properties that the client wants to be returned.</param>
        /// <param name="rows">A reference to a PropertyRow_r value. 
        /// It contains the address book container row the server returns in response to the request.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiGetProps(uint flags, STAT stat, PropertyTagArray_r? propTags, out PropertyRow_r? rows, bool needRetry = true);

        /// <summary>
        /// The NspiQueryRows method returns to the client a number of rows from a specified table.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="tableCount">A DWORD value that contains the number values in the input parameter table. 
        /// This value is limited to 100,000.</param>
        /// <param name="table">An array of DWORD values, representing an Explicit Table.</param>
        /// <param name="count">A DWORD value that contains the number of rows the client is requesting.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r value, 
        /// containing a list of the proptags of the properties that the client requires to be returned for each row returned.</param>
        /// <param name="rows">A nullable PropertyRowSet_r value, containing the address book container rows that the server returns in response to the request.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiQueryRows(uint flags, ref STAT stat, uint tableCount, uint[] table, uint count, PropertyTagArray_r? propTags, out PropertyRowSet_r? rows, bool needRetry = true);

        /// <summary>
        /// The NspiSeekEntries method searches for and sets the logical position in a specific table
        /// to the first entry greater than or equal to a specified value. 
        /// </summary>
        /// <param name="reserved">A DWORD value that is reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="target">A PropertyValue_r value holding the value that is being sought.</param>
        /// <param name="table">The value NULL or a PropertyTagArray_r value. 
        /// It holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="propTags">It contains a list of the proptags of the columns 
        /// that the client wants to be returned for each row returned.</param>
        /// <param name="rows">It contains the address book container rows the server returns in response to the request.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiSeekEntries(uint reserved, ref STAT stat, PropertyValue_r target, PropertyTagArray_r? table, PropertyTagArray_r? propTags, out PropertyRowSet_r? rows, bool needRetry = true);

        /// <summary>
        /// The NspiGetMatches method returns an Explicit Table. 
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use.</param>
        /// <param name="stat">A STAT block describing a logical position in a specific address book container.</param>
        /// <param name="proReserved">A PropertyTagArray_r reserved for future use.</param>
        /// <param name="reserved2">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="filter">The value NULL or a Restriction_r value. 
        /// It holds a logical restriction to apply to the rows in the address book container specified in the stat parameter.</param>
        /// <param name="propName">The value NULL or a PropertyName_r value. 
        /// It holds the property to be opened as a restricted address book container.</param>
        /// <param name="requested">A DWORD value. It contains the maximum number of rows to return in a restricted address book container.</param>
        /// <param name="outMids">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r value. 
        /// It contains a list of the proptags of the columns that the client wants to be returned for each row returned.</param>
        /// <param name="rows">A reference to a PropertyRowSet_r value. It contains the address book container rows the server returns in response to the request.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiGetMatches(uint reserved, ref STAT stat, PropertyTagArray_r? proReserved, uint reserved2, Restriction_r? filter, PropertyName_r? propName, uint requested, out PropertyTagArray_r? outMids, PropertyTagArray_r? propTags, out PropertyRowSet_r? rows, bool needRetry = true);

        /// <summary>
        /// The NspiResortRestriction method applies a sort order to the objects in a restricted address book container.
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A reference to a STAT block describing a logical position in a specific address book container.</param>
        /// <param name="proInMIds">A PropertyTagArray_r value. It holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="outMIds">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs 
        /// that comprise a restricted address book container.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiResortRestriction(uint reserved, ref STAT stat, PropertyTagArray_r proInMIds, ref PropertyTagArray_r? outMIds, bool needRetry = true);

        /// <summary>
        /// The NspiCompareMIds method compares the position in an address book container of two objects 
        /// identified by Minimal Entry ID and returns the value of the comparison.
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="mid1">The mid1 is a DWORD value containing a Minimal Entry ID.</param>
        /// <param name="mid2">The mid2 is a DWORD value containing a Minimal Entry ID.</param>
        /// <param name="results">A DWORD value. On return, it contains the result of the comparison.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiCompareMIds(uint reserved, STAT stat, uint mid1, uint mid2, out int results, bool needRetry = true);

        /// <summary>
        /// The NspiDNToMId method maps a set of DN to a set of Minimal Entry ID.
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="names">A StringsArray_r value. It holds a list of strings that contain DNs.</param>
        /// <param name="mids">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiDNToMId(uint reserved, StringsArray_r names, out PropertyTagArray_r? mids, bool needRetry = true);

        /// <summary>
        /// The NspiModProps method is used to modify the properties of an object in the address book. 
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r. 
        /// it contains a list of the proptags of the columns from which the client requests all the values to be removed.</param>
        /// <param name="row">A PropertyRow_r value. It contains an address book row.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiModProps(uint reserved, STAT stat, PropertyTagArray_r? propTags, PropertyRow_r row, bool needRetry = true);

        /// <summary>
        /// The NspiModLinkAtt method modifies the values of a specific property of a specific row in the address book.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="propTag">A DWORD value. It contains the proptag of the property that the client wants to modify.</param>
        /// <param name="mid">A DWORD value that contains the Minimal Entry ID of the address book row that the client wants to modify.</param>
        /// <param name="entryIds">A BinaryArray value. It contains a list of EntryIDs to be used to modify the requested property on the requested address book row.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiModLinkAtt(uint flags, uint propTag, uint mid, BinaryArray_r entryIds, bool needRetry = true);

        /// <summary>
        /// The NspiResolveNames method takes a set of string values in an 8-bit character set and performs ANR on those strings. 
        /// The NspiResolveNames method taking string values in an 8-bit character set is not supported when mapi_http transport is used. 
        /// </summary>
        /// <param name="reserved">A DWORD reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r value containing a list of the proptags of the columns 
        /// that the client requests to be returned for each row returned.</param>
        /// <param name="stringArray">A StringsArray_r value. It specifies the values on which the client is requesting the server to do ANR.</param>
        /// <param name="mids">A PropertyTagArray_r value. On return, it contains a list of Minimal Entry IDs that match the array of strings.</param>
        /// <param name="rows">A reference to a PropertyRowSet_r value. 
        /// It contains the address book container rows that the server returns in response to the request.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiResolveNames(uint reserved, STAT stat, PropertyTagArray_r? propTags, StringsArray_r? stringArray, out PropertyTagArray_r? mids, out PropertyRowSet_r? rows, bool needRetry = true);

        /// <summary>
        /// The NspiResolveNamesW method takes a set of string values in the Unicode character set and performs ANR on those strings. 
        /// </summary>
        /// <param name="reserved">A DWORD value that is reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r containing a list of the proptags of the columns 
        /// that the client requests to be returned for each row returned.</param>
        /// <param name="wstr">A WStringsArray_r value. It specifies the values on which the client is requesting the server to perform ANR.</param>
        /// <param name="mids">A PropertyTagArray_r value. On return, it contains a list of Minimal Entry IDs that match the array of strings.</param>
        /// <param name="rowOfResolveNamesW">A reference to a PropertyRowSet_r structure. 
        /// It contains the address book container rows that the server returns in response to the request.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiResolveNamesW(uint reserved, STAT stat, PropertyTagArray_r? propTags, WStringsArray_r? wstr, out PropertyTagArray_r? mids, out PropertyRowSet_r? rowOfResolveNamesW, bool needRetry = true);

        /// <summary>
        /// The NspiGetTemplateInfo method returns information about template objects.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="type">A DWORD value. It specifies the display type of the template for which the information is requested.</param>
        /// <param name="dn">The value NULL or the DN of the template requested. The value is NULL-terminated.</param>
        /// <param name="codePage">A DWORD value. It specifies the code page of the template for which the information is requested.</param>
        /// <param name="localeID">A DWORD value. It specifies the LCID of the template for which the information is requested.</param>
        /// <param name="data">A PropertyRow_r value. On return, it contains the information requested.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        ErrorCodeValue NspiGetTemplateInfo(uint flags, uint type, string dn, uint codePage, uint localeID, out PropertyRow_r? data, bool needRetry = true);
    }
}