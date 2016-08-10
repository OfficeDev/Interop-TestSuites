namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System.Collections.Generic;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// The interface of MS-OXCMAPIHTTP Adapter class.
    /// </summary>
    public interface IMS_OXCMAPIHTTPAdapter : IAdapter
    {
        #region Mailbox Server Endpoint.

        /// <summary>
        /// This method is used to establish a Session Context with the server with specified user.
        /// </summary>
        /// <param name="userName">The user name used to connect with server.</param>
        /// <param name="password">The password used to connect with server.</param>
        /// <param name="userDN">The UserESSDN used to connect with server.</param>
        /// <param name="cookies">Cookies used to identify the Session Context. </param>
        /// <param name="responseBody">The response body of the Connect request type.</param>
        /// <param name="webHeaderCollection">The web headers of the Connect request type.</param>
        /// <param name="httpStatus">The HTTP call response status.</param>
        /// <returns>The status code of the Connect request type.</returns>
        uint Connect(string userName, string password, string userDN, ref CookieCollection cookies, out MailboxResponseBodyBase responseBody, ref WebHeaderCollection webHeaderCollection, out HttpStatusCode httpStatus);

        /// <summary>
        /// This method is used by the client to delete a Session Context with the server.
        /// </summary>
        /// <param name="responseBody">The response body of the Disconnect request type.</param>
        /// <returns>The status code of the Disconnect request type.</returns>
        uint Disconnect(out MailboxResponseBodyBase responseBody);

        /// <summary>
        /// This method is used by the client to send remote operation requests to the server with specified cookies.
        /// </summary>
        /// <param name="requestBody">The request body of the Execute request type.</param>
        /// <param name="cookies">Cookies used to identify the Session Context.</param>
        /// <param name="httpHeaders">The request and response headers of the Execute request type.</param>
        /// <param name="responseBody">The response body of the Execute request type.</param>
        /// <param name="metatags">The meta tags in the response body buffer.</param>
        /// <returns>The status code of the Execute request type.</returns>
        uint Execute(ExecuteRequestBody requestBody, CookieCollection cookies, ref WebHeaderCollection httpHeaders, out MailboxResponseBodyBase responseBody, out List<string> metatags);

        /// <summary>
        /// This method is used by the client to request that the server notify the client when a processing request that takes an extended amount of time completes.
        /// </summary>
        /// <param name="notificationWaitRequestBody">The request body of the NotificationWait request type.</param>
        /// <param name="httpHeaders">The request and response header of the NotificationWait request type.</param>
        /// <param name="responseBody">The response body of the NotificationWait request type.</param>
        /// <param name="metatags">The meta tags of the NotificationWait request type.</param>
        /// <param name="additionalHeader">The additional headers in the Notification request type response.</param>
        /// <returns>The status code of the NotificationWait request type.</returns>
        uint NotificationWait(NotificationWaitRequestBody notificationWaitRequestBody, ref WebHeaderCollection httpHeaders, out MailboxResponseBodyBase responseBody, out List<string> metatags, out Dictionary<string, string> additionalHeader);

        /// <summary>
        /// This method allows a client to determine whether a server's endpoint is reachable and operational.
        /// </summary>
        /// <param name="endpoint">The endpoint used by PING request.</param>
        /// <param name="metatags">The meta tags in the response body of the Ping request type.</param>
        /// <param name="headers">The request and response header of the PING request type.</param>
        /// <returns>The status code of the PING request type.</returns>
        uint PING(ServerEndpoint endpoint, out List<string> metatags, out WebHeaderCollection headers);
        #endregion

        #region Address Book Server Endpoint.

        /// <summary>
        /// This method is used by the client to establish a Session Context with the Address Book Server.
        /// </summary>
        /// <param name="bindRequestBody">The bind request type request body.</param>
        /// <param name="responseCode">The value of X-ResponseCode header of the bind response.</param>
        /// <returns>The response body of bind request type.</returns>
        BindResponseBody Bind(BindRequestBody bindRequestBody, out int responseCode);

        /// <summary>
        /// This method is used by the client to delete a Session Context with the Address Book Server.
        /// </summary>
        /// <param name="unbindRequestBody">The unbind request type request body.</param>
        /// <returns>The response body of unbind request type.</returns>
        UnbindResponseBody Unbind(UnbindRequestBody unbindRequestBody);

        /// <summary>
        /// This method is used by the client to compare the position of two objects in an address book container.
        /// </summary>
        /// <param name="compareMIdsRequestBody">The CompareMinIds request type request body.</param>
        /// <returns>The response body of the CompareMinIds request type.</returns>
        CompareMinIdsResponseBody CompareMinIds(CompareMinIdsRequestBody compareMIdsRequestBody);

        /// <summary>
        /// This method is used by the client to map a set of distinguished names to a set of Minimal Entry IDs.
        /// </summary>
        /// <param name="distinguishedNameToMIdRequestBody">The DnToMinId request type request body.</param>
        /// <returns>The response body of the DnToMinId request type.</returns>
        DnToMinIdResponseBody DnToMinId(DNToMinIdRequestBody distinguishedNameToMIdRequestBody);

        /// <summary>
        /// This method is used by the client to get an Explicit Table, in which the rows are determined by the specified criteria.
        /// </summary>
        /// <param name="getMatchesRequestBody">The GetMatches request type request body.</param>
        /// <returns>The response body of the GetMatches request type.</returns>
        GetMatchesResponseBody GetMatches(GetMatchesRequestBody getMatchesRequestBody);

        /// <summary>
        /// This method is used by the client to get a list of all of the properties that have values on an object.
        /// </summary>
        /// <param name="getPropListRequestBody">The GetPropList request type request body.</param>
        /// <returns>The response body of the GetPropList request type.</returns>
        GetPropListResponseBody GetPropList(GetPropListRequestBody getPropListRequestBody);

        /// <summary>
        /// This method is used by the client to get specific properties on an object.
        /// </summary>
        /// <param name="getPropsRequestBody">The GetProps request type request body.</param>
        /// <returns>The response body of the GetProps request type.</returns>
        GetPropsResponseBody GetProps(GetPropsRequestBody getPropsRequestBody);

        /// <summary>
        /// This method is used by the client to get specific properties on an object.
        /// </summary>
        /// <param name="getPropsRequestBody">The GetProps request type request body.</param>
        /// <param name="responseCodeHeader">The value of X-ResponseCode header</param>
        /// <returns>The response body of the GetProps request type.</returns>
        GetPropsResponseBody GetProps(GetPropsRequestBody getPropsRequestBody, out uint responseCodeHeader);

        /// <summary>
        /// This method is used by the client to get a special table, which can be either an address book hierarchy table or an address creation table.
        /// </summary>
        /// <param name="getSpecialTableRequestBody">The GetSpecialTable request type request body.</param>
        /// <returns>The response body of the GetSpecialTable request type.</returns>
        GetSpecialTableResponseBody GetSpecialTable(GetSpecialTableRequestBody getSpecialTableRequestBody);

        /// <summary>
        /// This method is used by the client to get information about a template that is used by the address book.
        /// </summary>
        /// <param name="getTemplateInfoRequestBody">The GetTemplateInfo request type request body.</param>
        /// <returns>The response body of the GetTemplateInfo request type.</returns>
        GetTemplateInfoResponseBody GetTemplateInfo(GetTemplateInfoRequestBody getTemplateInfoRequestBody);

        /// <summary>
        /// This method is used by the client to modify a specific property of a row in the address book.
        /// </summary>
        /// <param name="modLinkAttRequestBody">The ModLinkATT request type request body.</param>
        /// <returns>The response body of the ModLinkAtt request type.</returns>
        ModLinkAttResponseBody ModLinkAtt(ModLinkAttRequestBody modLinkAttRequestBody);

        /// <summary>
        /// This method is used by the client to modify the specific properties of an Address Book object.
        /// </summary>
        /// <param name="modPropsRequestBody">The ModProps request type request body.</param>
        /// <returns>The response body of the ModProps request type.</returns>
        ModPropsResponseBody ModProps(ModPropsRequestBody modPropsRequestBody);

        /// <summary>
        /// This method is used by the client to get a number of rows from the specified Explicit Table.
        /// </summary>
        /// <param name="queryRowsRequestBody">The QueryRows request type request body.</param>
        /// <returns>The response body of QueryRows request type.</returns>
        QueryRowsResponseBody QueryRows(QueryRowsRequestBody queryRowsRequestBody);

        /// <summary>
        /// This method is used by the client to get a list of all the properties that exist in the address book.
        /// </summary>
        /// <param name="queryColumnsRequestBody">The QueryColumns request type request body.</param>
        /// <returns>The response body of QueryColumns request type.</returns>
        QueryColumnsResponseBody QueryColumns(QueryColumnsRequestBody queryColumnsRequestBody);

        /// <summary>
        /// This method is used by the client to perform ambiguous name resolution(ANR).
        /// </summary>
        /// <param name="resolveNamesRequestBody">The ResolveNames request type request body.</param>
        /// <returns>The response body of the ResolveNames request type.</returns>
        ResolveNamesResponseBody ResolveNames(ResolveNamesRequestBody resolveNamesRequestBody);

        /// <summary>
        /// This method is used by the client to sort the objects in the restricted address book container.
        /// </summary>
        /// <param name="resortRestrictionRequestBody">The ResortRestriction request type request body.</param>
        /// <returns>The response body of the ResortRestriction request type.</returns>
        ResortRestrictionResponseBody ResortRestriction(ResortRestrictionRequestBody resortRestrictionRequestBody);

        /// <summary>
        /// This method is used by the client to search for and set the logical position in a specific table to the first entry greater than or equal to a specified value.
        /// </summary>
        /// <param name="seekEntriesRequestBody">The SeekEntries request type request body.</param>
        /// <returns>The response body of SeekEntries request type.</returns>
        SeekEntriesResponseBody SeekEntries(SeekEntriesRequestBody seekEntriesRequestBody);

        /// <summary>
        /// This method is used by the client to update the STAT structure to reflect the client's changes.
        /// </summary>
        /// <param name="updateStatRequestBody">The UpdateStat request type request body.</param>
        /// <returns>The response body of UpdateStat request type.</returns>
        UpdateStatResponseBody UpdateStat(UpdateStatRequestBody updateStatRequestBody);

        /// <summary>
        /// This method is used by the client to get the Uniform Resource Locator (URL) of the specified mailbox server endpoint.
        /// </summary>
        /// <param name="getMailboxUrlRequestBody">The GetMailboxUrl request type request body.</param>
        /// <returns>The response body of the GetMailboxUrl request type.</returns>
        GetMailboxUrlResponseBody GetMailboxUrl(GetMailboxUrlRequestBody getMailboxUrlRequestBody);

        /// <summary>
        /// This method is used by the client to get the URL of the specified address book server endpoint.
        /// </summary>
        /// <param name="getAddressBookUrlRequestBody">The GetAddressBookUrl request type request body.</param>
        /// <returns>The response body of GetAddressBookUrl request type.</returns>
        GetAddressBookUrlResponseBody GetAddressBookUrl(GetAddressBookUrlRequestBody getAddressBookUrlRequestBody);
        #endregion
    }
}