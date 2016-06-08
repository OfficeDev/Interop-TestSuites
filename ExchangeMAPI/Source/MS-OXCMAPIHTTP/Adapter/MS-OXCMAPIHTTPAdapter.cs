namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// Adapter class of MS-OXCMAPIHTTP.
    /// </summary>
    public partial class MS_OXCMAPIHTTPAdapter : ManagedAdapterBase, IMS_OXCMAPIHTTPAdapter
    {
        #region Variables

        /// <summary>
        /// Whether the common configuration file has been imported.
        /// </summary>
        private static bool commonConfigImported = false;

        /// <summary>
        /// The Mailbox userName which can be used by client to connect to the SUT.
        /// </summary>
        private string userName;

        /// <summary>
        /// The user password which can be used by client to access to the SUT.
        /// </summary>
        private string password;

        /// <summary>
        /// The name of domain where the server belongs to.
        /// </summary>
        private string domainName;

        /// <summary>
        /// The URI that the client can use to connect to a mailbox via MAPIHTTP.
        /// </summary>
        private string mailStoreUrl;

        /// <summary>
        /// The URI that the client can use to connect to a NSPI server via MAPIHTTP.
        /// </summary>
        private string addressBookUrl;

        #endregion Variables

        #region Initialize TestSuite.

        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Pass ITestSite into adapter and make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXCMAPIHTTP";
            if (!commonConfigImported)
            {
                Common.MergeConfiguration(this.Site);
                commonConfigImported = true;
            }

            this.domainName = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.userName = Common.GetConfigurationPropertyValue("AdminUserName", this.Site);
            this.password = Common.GetConfigurationPropertyValue("AdminUserPassword", this.Site);

            AdapterHelper.SessionContextCookies = new CookieCollection();
            AdapterHelper.ClientInstance = string.Empty;
            AdapterHelper.Counter = 0;
        }

        #endregion Initialize TestSuite

        #region MS-OXCMAPIHTTPAdapter Members

        #region Mailbox Server Endpoint

        /// <summary>
        /// This method is used to establish a Session Context with the server with specified user.
        /// </summary>
        /// <param name="userName">The UserName used to connect with server.</param>
        /// <param name="password">The password used to connect with server.</param>
        /// <param name="userDN">The UserESSDN used to connect with server.</param>
        /// <param name="cookies">Cookies used to identify the Session Context.</param>
        /// <param name="responseBody">The response body of the Connect request type.</param>
        /// <param name="webHeaderCollection">The web headers of the Connect request type.</param>
        /// <param name="httpStatus">The HTTP call response status.</param>
        /// <returns>The status code of the Connect request type.</returns>
        public uint Connect(string userName, string password, string userDN, ref CookieCollection cookies, out MailboxResponseBodyBase responseBody, ref WebHeaderCollection webHeaderCollection, out HttpStatusCode httpStatus)
        {
            responseBody = null;
            byte[] rgbAuxIn = new byte[] { };
            byte[] rawBuffer;

            // Prepare the connect request body.
            ConnectRequestBody connectRequestBody = new ConnectRequestBody();
            connectRequestBody.UserDN = userDN;
            connectRequestBody.Flags = ConstValues.ConnectionFlag;
            connectRequestBody.Cpid = ConstValues.CodePageId;
            connectRequestBody.LcidString = ConstValues.DefaultLocale;
            connectRequestBody.LcidSort = ConstValues.DefaultLocale;
            connectRequestBody.AuxiliaryBufferSize = (uint)rgbAuxIn.Length;
            connectRequestBody.AuxiliaryBuffer = rgbAuxIn;
                        
            // Send the HTTP request and get the HTTP response.
            HttpWebResponse response = this.SendMAPIHttpRequest(userName, password, connectRequestBody, ServerEndpoint.MailboxServerEndpoint, cookies, webHeaderCollection, out rawBuffer);
            webHeaderCollection = response.Headers;
            httpStatus = response.StatusCode;
            uint responseCode = AdapterHelper.GetFinalResponseCode(response.Headers["X-ResponseCode"]);
            
            if (httpStatus != HttpStatusCode.OK)
            {
                return 0;
            }

            // Read the HTTP response buffer and parse the response to correct format.
            CommonResponse commonResponse = CommonResponse.ParseCommonResponse(rawBuffer);
            
            if (responseCode == 0)
            {
                Site.Assert.IsNotNull(commonResponse.ResponseBodyRawData, "The response body should contains data.");
                uint statusCode = BitConverter.ToUInt32(commonResponse.ResponseBodyRawData, 0);
                if (statusCode == 0)
                {
                    // Connect succeeded when the StatusCode field equals zero.
                    ConnectSuccessResponseBody responseSuccess = ConnectSuccessResponseBody.Parse(commonResponse.ResponseBodyRawData);
                    responseBody = responseSuccess;

                    this.VerifyConnectSuccessResponseBody(responseSuccess);
                }
                
                this.VerifyHTTPS(response);
                this.VerifyAuthentication(response);
                this.VerifyAutoDiscover(httpStatus, ServerEndpoint.MailboxServerEndpoint);
                this.VerifyHTTPHeaders(response.Headers);
                this.VerifyAdditionalHeaders(commonResponse.AdditionalHeaders);
                this.VerifyConnectResponse(response);
                this.VerifyConnectOrBindResponse(response.Headers);
                this.VerifyRequestTypesForMailboxServerEndpoint(response.Headers, commonResponse);
                this.VerifyResponseMetaTags(commonResponse.MetaTags);
            }

            this.VerifyRespondingToAllRequestTypeRequests(response, commonResponse, responseCode);
            response.GetResponseStream().Close();
            AdapterHelper.SessionContextCookies = response.Cookies;
            cookies = response.Cookies;
            return responseCode;
        }

        /// <summary>
        /// This method is used by the client to delete a Session Context with the server.
        /// </summary>
        /// <param name="responseBody">The response body of the Disconnect request type.</param>
        /// <returns>The status code of the Disconnect request type.</returns>
        public uint Disconnect(out MailboxResponseBodyBase responseBody)
        {
            responseBody = null;
            byte[] rawBuffer;
            
            // Prepare the disconnect request body.
            DisconnectRequestBody disconnectRequestBody = new DisconnectRequestBody();
            byte[] rgbAuxIn = new byte[] { };
            disconnectRequestBody.AuxiliaryBufferSize = (uint)rgbAuxIn.Length;
            disconnectRequestBody.AuxiliaryBuffer = rgbAuxIn;
            WebHeaderCollection webHeaderCollection = AdapterHelper.InitializeHTTPHeader(RequestType.Disconnect, AdapterHelper.ClientInstance, AdapterHelper.Counter); 

            // Send the disconnect HTTP request and get the response.
            HttpWebResponse response = this.SendMAPIHttpRequest(this.userName, this.password, disconnectRequestBody, ServerEndpoint.MailboxServerEndpoint, AdapterHelper.SessionContextCookies, webHeaderCollection, out rawBuffer);
            uint responseCode = AdapterHelper.GetFinalResponseCode(response.Headers["X-ResponseCode"]);

            // Read the HTTP response buffer and parse the response to correct format.
            CommonResponse commonResponse = CommonResponse.ParseCommonResponse(rawBuffer);
            if (responseCode == 0)
            {
                Site.Assert.IsNotNull(commonResponse.ResponseBodyRawData, "The response body should contains data.");
                uint statusCode = BitConverter.ToUInt32(commonResponse.ResponseBodyRawData, 0);
                if (statusCode == 0)
                {
                    // Disconnect succeeded when the StatusCode field equals zero.
                    DisconnectSuccessResponseBody responseSuccess = DisconnectSuccessResponseBody.Parse(commonResponse.ResponseBodyRawData);
                    responseBody = responseSuccess;

                    this.VerifyDisconnectSuccessResponseBody(responseSuccess);
                }
                
                this.VerifyHTTPS(response);
                this.VerifyAuthentication(response);
                this.VerifyAutoDiscover(response.StatusCode, ServerEndpoint.MailboxServerEndpoint);
                this.VerifyHTTPHeaders(response.Headers);
                this.VerifyAdditionalHeaders(commonResponse.AdditionalHeaders);
                this.VerifyDisconnectResponse(response);
                this.VerifyRequestTypesForMailboxServerEndpoint(response.Headers, commonResponse);
                this.VerifyResponseMetaTags(commonResponse.MetaTags);
            }

            this.VerifyContentTypeHeader(response.Headers);
            this.VerifyRespondingToAllRequestTypeRequests(response, commonResponse, responseCode);
            response.GetResponseStream().Close();
            AdapterHelper.SessionContextCookies = response.Cookies;
            return responseCode;
        }

        /// <summary>
        /// This method is used by the client to send remote operation requests to the server with specified cookies.
        /// </summary>
        /// <param name="requestBody">The request body of the Execute request type.</param>
        /// <param name="cookies">Cookies used to identify the Session Context.</param>
        /// <param name="httpHeaders">The request and response header of the Execute request type.</param>
        /// <param name="responseBody">The response body of the Execute request type.</param>
        /// <param name="metatags">The meta tags in the response body buffer.</param>
        /// <returns>The status code of the Execute request type.</returns>
        public uint Execute(ExecuteRequestBody requestBody, CookieCollection cookies, ref WebHeaderCollection httpHeaders, out MailboxResponseBodyBase responseBody, out List<string> metatags)
        {
            responseBody = null;
            metatags = null;
            byte[] rawBuffer;

            // Send the execute HTTP request and get the response.
            HttpWebResponse response = this.SendMAPIHttpRequest(this.userName, this.password, requestBody, ServerEndpoint.MailboxServerEndpoint, cookies, httpHeaders, out rawBuffer);
            httpHeaders = response.Headers;
            uint responseCode = AdapterHelper.GetFinalResponseCode(response.Headers["X-ResponseCode"]);
             
            // Read the HTTP response buffer and parse the response to correct format.
            CommonResponse commonResponse = CommonResponse.ParseCommonResponse(rawBuffer);
            metatags = commonResponse.MetaTags;
            if (responseCode == 0)
            {
                Site.Assert.IsNotNull(commonResponse.ResponseBodyRawData, "The response body should contains data.");
                uint statusCode = BitConverter.ToUInt32(commonResponse.ResponseBodyRawData, 0);
                if (statusCode == 0)
                {
                    // Execute request type executed successfully when the StatusCode field equals zero.
                    ExecuteSuccessResponseBody responseSuccess = ExecuteSuccessResponseBody.Parse(commonResponse.ResponseBodyRawData);
                    responseBody = responseSuccess;

                    this.VerifyHTTPS(response);
                    this.VerifyExecuteSuccessResponseBody(responseSuccess);
                }

                this.VerifyHTTPHeaders(response.Headers);
                this.VerifyAuthentication(response);
                this.VerifyAutoDiscover(response.StatusCode, ServerEndpoint.MailboxServerEndpoint);
                this.VerifyAdditionalHeaders(commonResponse.AdditionalHeaders);
                this.VerifyRequestTypesForMailboxServerEndpoint(response.Headers, commonResponse);
                this.VerifyResponseMetaTags(commonResponse.MetaTags);
            }

            this.VerifyContentTypeHeader(response.Headers);
            this.VerifyRespondingToAllRequestTypeRequests(response, commonResponse, responseCode);
            response.GetResponseStream().Close();
            AdapterHelper.SessionContextCookies = response.Cookies;
            return responseCode;
        }

        /// <summary>
        /// This method is used by the client to request that the server notify the client when a processing request that takes an extended amount of time completes.
        /// </summary>
        /// <param name="notificationWaitRequestBody">The request body of the NotificationWait request type.</param>
        /// <param name="httpHeaders">The request and response header of the NotificationWait request type.</param>
        /// <param name="responseBody">The response body of the NotificationWait request type.</param>
        /// <param name="metatags">The meta tags of the NotificationWait request type.</param>
        /// <param name="additionalHeader">The additional headers in the Notification request type response.</param>
        /// <returns>The status code of the NotificationWait request type.</returns>
        public uint NotificationWait(NotificationWaitRequestBody notificationWaitRequestBody, ref WebHeaderCollection httpHeaders, out MailboxResponseBodyBase responseBody, out List<string> metatags, out Dictionary<string, string> additionalHeader)
        {
            responseBody = null;
            byte[] rawBuffer;

            // Send the NotificationWait HTTP request and get the response.
            HttpWebResponse response = this.SendMAPIHttpRequest(this.userName, this.password, notificationWaitRequestBody, ServerEndpoint.MailboxServerEndpoint, AdapterHelper.SessionContextCookies, httpHeaders, out rawBuffer);
            httpHeaders = response.Headers;
            uint responseCode = AdapterHelper.GetFinalResponseCode(response.Headers["X-ResponseCode"]);

            // Read the HTTP response buffer and parse the response to correct format.
            CommonResponse commonResponse = CommonResponse.ParseCommonResponse(rawBuffer);
            metatags = commonResponse.MetaTags;
            additionalHeader = commonResponse.AdditionalHeaders;
            if (responseCode == 0)
            {
                Site.Assert.IsNotNull(commonResponse.ResponseBodyRawData, "The response body should contains data.");
                uint statusCode = BitConverter.ToUInt32(commonResponse.ResponseBodyRawData, 0);
                if (statusCode == 0)
                {
                    // Send the NotificationWait request succeeded when the StatusCode field equals zero.
                    NotificationWaitSuccessResponseBody responseSuccess = NotificationWaitSuccessResponseBody.Parse(commonResponse.ResponseBodyRawData);
                    responseBody = responseSuccess;

                    this.VerifyNotificationWaitSuccessResponseBody(responseSuccess);
                }
               
                this.VerifyHTTPS(response);
                this.VerifyAuthentication(response);
                this.VerifyAutoDiscover(response.StatusCode, ServerEndpoint.MailboxServerEndpoint);
                this.VerifyHTTPHeaders(response.Headers);
                this.VerifyAdditionalHeaders(commonResponse.AdditionalHeaders);
                this.VerifyRequestTypesForMailboxServerEndpoint(response.Headers, commonResponse);
                this.VerifyResponseMetaTags(commonResponse.MetaTags);
                this.VerifyNotificationWaitRequestType(response.Headers);
            }

            this.VerifyContentTypeHeader(response.Headers);
            this.VerifyRespondingToAllRequestTypeRequests(response, commonResponse, responseCode);
            response.GetResponseStream().Close();
            return responseCode;
        }

        /// <summary>
        /// This method allows a client to determine whether a server's endpoint is reachable and operational.
        /// </summary>
        /// <param name="endpoint">The endpoint used by PING request.</param>
        /// <param name="metatags">The meta tags in the response body of the Ping request type.</param>
        /// <param name="headers">The request and response header of the PING request type.</param>
        /// <returns>The status code of the PING request type.</returns>
        public uint PING(ServerEndpoint endpoint, out List<string> metatags, out WebHeaderCollection headers)
        {
            metatags = null;
            byte[] rawBuffer;
            WebHeaderCollection webHeaderCollection = AdapterHelper.InitializeHTTPHeader(RequestType.PING, AdapterHelper.ClientInstance, AdapterHelper.Counter);

            // Send the PING HTTP request and get the response.
            HttpWebResponse response = this.SendMAPIHttpRequest(this.userName, this.password, null, endpoint, AdapterHelper.SessionContextCookies, webHeaderCollection, out rawBuffer);
            headers = response.Headers;

            // Read the HTTP response buffer and parse the response to correct format.
            uint responseCode = AdapterHelper.GetFinalResponseCode(response.Headers["X-ResponseCode"]);

            CommonResponse commonResponse = CommonResponse.ParseCommonResponse(rawBuffer);
            if (responseCode == 0)
            {
                // PING succeeded when the response code equals zero and PING request has no response body.
                metatags = commonResponse.MetaTags;
                this.VerifyHTTPS(response);
                this.VerifyAuthentication(response);
                this.VerifyAutoDiscover(response.StatusCode, endpoint);
                this.VerifyHTTPHeaders(response.Headers);
                this.VerifyAdditionalHeaders(commonResponse.AdditionalHeaders);
                this.VerifyPINGRequestType(commonResponse, endpoint, responseCode);
                this.VerifyResponseMetaTags(commonResponse.MetaTags);
            }

            this.VerifyContentTypeHeader(response.Headers);
            this.VerifyRespondingToAllRequestTypeRequests(response, commonResponse, responseCode);
            response.GetResponseStream().Close();
            return responseCode;
        }
        #endregion

        #region Address Book Server Endpoint
        
        /// <summary>
        /// This method is used by the client to establish a Session Context with the Address Book Server.
        /// </summary>
        /// <param name="bindRequestBody">The bind request type request body.</param>
        /// <param name="responseCode">The value of X-ResponseCode header of the bind response.</param>
        /// <returns>The response body of bind request type.</returns>
        public BindResponseBody Bind(BindRequestBody bindRequestBody, out int responseCode)
        {
            byte[] rawBuffer = null;
            CommonResponse commonResponse = null;
            BindResponseBody bindResponseBody = null;
            AdapterHelper.Counter = 1;
            AdapterHelper.ClientInstance = Guid.NewGuid().ToString();
            WebHeaderCollection webHeaderCollection = AdapterHelper.InitializeHTTPHeader(RequestType.Bind, AdapterHelper.ClientInstance, AdapterHelper.Counter);

            // Send the Execute HTTP request and get the response.
            HttpWebResponse response = this.SendMAPIHttpRequest(this.userName, this.password, bindRequestBody, ServerEndpoint.AddressBookServerEndpoint, AdapterHelper.SessionContextCookies, webHeaderCollection, out rawBuffer);
            responseCode = (int)AdapterHelper.GetFinalResponseCode(response.Headers["X-ResponseCode"]);
          
            // Read the HTTP response buffer and parse the response to correct format.
            if (responseCode == 0)
            {
                commonResponse = CommonResponse.ParseCommonResponse(rawBuffer);
                Site.Assert.IsNotNull(commonResponse.ResponseBodyRawData, "The response body should contains data.");
                bindResponseBody = BindResponseBody.Parse(commonResponse.ResponseBodyRawData);
                this.VerifyBindResponseBody(bindResponseBody);
                this.VerifyAutoDiscover(response.StatusCode, ServerEndpoint.AddressBookServerEndpoint);
            }

            this.VerifyAuthentication(response);
            response.GetResponseStream().Close();
            AdapterHelper.SessionContextCookies = response.Cookies;

            return bindResponseBody;
        }

        /// <summary>
        /// This method is used by the client to delete a Session Context with the Address Book Server.
        /// </summary>
        /// <param name="unbindRequestBody">The unbind request type request body.</param>
        /// <returns>The response body of unbind request type.</returns>
        public UnbindResponseBody Unbind(UnbindRequestBody unbindRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(unbindRequestBody, RequestType.Unbind);
            AdapterHelper.SessionContextCookies = new CookieCollection();
            UnbindResponseBody unbindResponseBody = UnbindResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyUnbindResponseBody(unbindResponseBody);
         
            return unbindResponseBody;
        }

        /// <summary>
        /// This method is used by the client to compare the position of two objects in an address book container.
        /// </summary>
        /// <param name="compareMIdsRequestBody">The CompareMinIds request type request body.</param>
        /// <returns>The response body of the CompareMinIds request type.</returns>
        public CompareMinIdsResponseBody CompareMinIds(CompareMinIdsRequestBody compareMIdsRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(compareMIdsRequestBody, RequestType.CompareMIds);
            CompareMinIdsResponseBody compareMinIdsResponseBody = CompareMinIdsResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyComapreMinIdsResponsebody(compareMinIdsResponseBody);
 
            return compareMinIdsResponseBody;
        }

        /// <summary>
        /// This method is used by the client to map a set of distinguished names to a set of Minimal Entry IDs.
        /// </summary>
        /// <param name="distinguishedNameToMIdRequestBody">The DnToMinId request type request body.</param>
        /// <returns>The response body of the DnToMinId request type.</returns>
        public DnToMinIdResponseBody DnToMinId(DNToMinIdRequestBody distinguishedNameToMIdRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(distinguishedNameToMIdRequestBody, RequestType.DNToMId);
            DnToMinIdResponseBody distinguishedNameToMinIdResponseBody = DnToMinIdResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyDnToMinIdResponseBody(distinguishedNameToMinIdResponseBody);

            return distinguishedNameToMinIdResponseBody;
        }

        /// <summary>
        /// This method is used by the client to get an Explicit Table, in which the rows are determined by the specified criteria.
        /// </summary>
        /// <param name="getMatchesRequestBody">The GetMatches request type request body.</param>
        /// <returns>The response body of the GetMatches request type.</returns>
        public GetMatchesResponseBody GetMatches(GetMatchesRequestBody getMatchesRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(getMatchesRequestBody, RequestType.GetMatches);
            GetMatchesResponseBody getMatchesResponseBody = GetMatchesResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyGetMatchsResponseBody(getMatchesRequestBody, getMatchesResponseBody);

            if (getMatchesResponseBody.HasColumnsAndRows)
            {
                foreach (AddressBookPropertyRow row in getMatchesResponseBody.RowData)
                {
                    this.VerifyAddressBookPropertyRowStructure(row);
                }
            }

            return getMatchesResponseBody;
        }

        /// <summary>
        /// This method is used by the client to get a list of all of the properties that have values on an object.
        /// </summary>
        /// <param name="getPropListRequestBody">The GetPropList request type request body.</param>
        /// <returns>The response body of the GetPropList request type.</returns>
        public GetPropListResponseBody GetPropList(GetPropListRequestBody getPropListRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(getPropListRequestBody, RequestType.GetPropList);
            GetPropListResponseBody getPropListResponseBody = GetPropListResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyGetPropListResponseBody(getPropListResponseBody);

            return getPropListResponseBody;
        }

        /// <summary>
        /// This method is used by the client to get specific properties on an object.
        /// </summary>
        /// <param name="getPropsRequestBody">The GetProps request type request body.</param>
        /// <returns>The response body of the GetProps request type.</returns>
        public GetPropsResponseBody GetProps(GetPropsRequestBody getPropsRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(getPropsRequestBody, RequestType.GetProps);
            GetPropsResponseBody getPropsResponseBody = GetPropsResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyGetPropsResponseBody(getPropsResponseBody);

            return getPropsResponseBody;
        }

        /// <summary>
        /// This method is used by the client to get specific properties on an object.
        /// </summary>
        /// <param name="getPropsRequestBody">The GetProps request type request body.</param>
        /// <param name="responseCodeHeader">The value of X-ResponseCode header</param>
        /// <returns>The response body of the GetProps request type.</returns>
        public GetPropsResponseBody GetProps(GetPropsRequestBody getPropsRequestBody,out uint responseCodeHeader)
        {
            byte[] rawBuffer;
            CommonResponse commonResponse = null;
            WebHeaderCollection webHeaderCollection = AdapterHelper.InitializeHTTPHeader(RequestType.GetProps, AdapterHelper.ClientInstance, AdapterHelper.Counter);

            // Send the Execute HTTP request and get the response.
            HttpWebResponse response = this.SendMAPIHttpRequest(this.userName, this.password, getPropsRequestBody, ServerEndpoint.AddressBookServerEndpoint, AdapterHelper.SessionContextCookies, webHeaderCollection, out rawBuffer);
            responseCodeHeader = AdapterHelper.GetFinalResponseCode(response.Headers["X-ResponseCode"]);

            // Read the HTTP response buffer and parse the response to correct format.
            commonResponse = CommonResponse.ParseCommonResponse(rawBuffer);
 
            response.GetResponseStream().Close();
            AdapterHelper.SessionContextCookies = response.Cookies;
            GetPropsResponseBody getPropsResponseBody = new GetPropsResponseBody();
            if (commonResponse.ResponseBodyRawData.Length > 0)
            {
                getPropsResponseBody= GetPropsResponseBody.Parse(commonResponse.ResponseBodyRawData);
            }
               
            return getPropsResponseBody;
        }


        /// <summary>
        /// This method is used by the client to get a special table, which can be either an address book hierarchy table or an address creation table.
        /// </summary>
        /// <param name="getSpecialTableRequestBody">The GetSpecialTable request type request body.</param>
        /// <returns>The response body of the GetSpecialTable request type.</returns>
        public GetSpecialTableResponseBody GetSpecialTable(GetSpecialTableRequestBody getSpecialTableRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(getSpecialTableRequestBody, RequestType.GetSpecialTable);
            GetSpecialTableResponseBody getSpecialTableResponseBody = GetSpecialTableResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyGetSpecialTableResponseBody(getSpecialTableResponseBody);

            return getSpecialTableResponseBody;
        }

        /// <summary>
        /// This method is used by the client to get information about a template that is used by the address book.
        /// </summary>
        /// <param name="getTemplateInfoRequestBody">The GetTemplateInfo request type request body.</param>
        /// <returns>The response body of the GetTemplateInfo request type.</returns>
        public GetTemplateInfoResponseBody GetTemplateInfo(GetTemplateInfoRequestBody getTemplateInfoRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(getTemplateInfoRequestBody, RequestType.GetTemplateInfo);
            GetTemplateInfoResponseBody getTemplateInfoResponseBody = GetTemplateInfoResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyGetTemplateInfoResponseBody(getTemplateInfoResponseBody);

            return getTemplateInfoResponseBody;
        }

        /// <summary>
        /// This method is used by the client to modify a specific property of a row in the address book.
        /// </summary>
        /// <param name="modLinkAttRequestBody">The ModLinkAtt request type request body.</param>
        /// <returns>The response body of the ModLinkAtt request type.</returns>
        public ModLinkAttResponseBody ModLinkAtt(ModLinkAttRequestBody modLinkAttRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(modLinkAttRequestBody, RequestType.ModLinkAtt);
            ModLinkAttResponseBody modLinkAttResponseBody = ModLinkAttResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyModLinkAttResponseBody(modLinkAttResponseBody);

            return modLinkAttResponseBody;
        }

        /// <summary>
        /// This method is used by the client to modify the specific properties of an Address Book object.
        /// </summary>
        /// <param name="modPropsRequestBody">The ModProps request type request body.</param>
        /// <returns>The response body of the ModProps request type.</returns>
        public ModPropsResponseBody ModProps(ModPropsRequestBody modPropsRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(modPropsRequestBody, RequestType.ModProps);
            ModPropsResponseBody modPropsResponseBody = ModPropsResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyModPropsResponseBody(modPropsResponseBody);

            return modPropsResponseBody;
        }

        /// <summary>
        /// This method is used by the client to get a number of rows from the specified Explicit Table.
        /// </summary>
        /// <param name="queryRowsRequestBody">The QueryRows request type request body.</param>
        /// <returns>The response body of QueryRows request type</returns>
        public QueryRowsResponseBody QueryRows(QueryRowsRequestBody queryRowsRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(queryRowsRequestBody, RequestType.QueryRows);
            QueryRowsResponseBody queryRowsResponseBody = QueryRowsResponseBody.Parse(commonResponse.ResponseBodyRawData);

            this.VerifyQueryRowsResponseBody(queryRowsResponseBody, queryRowsRequestBody);
            if (queryRowsResponseBody.HasColumnsAndRows)
            {
                foreach (AddressBookPropertyRow row in queryRowsResponseBody.RowData)
                {
                    this.VerifyAddressBookPropertyRowStructure(row);
                   
                    if (row.Flag == 0x0)
                    {
                        for (int i = 0; i < row.ValueArray.Length; i++)
                        {
                            if (queryRowsRequestBody.Columns.PropertyTags[i].PropertyType != 0x0000)
                            {
                                this.VerifyAddressBookPropertyValueStructure(row.ValueArray[i]);
                            }
                        }
                    }
                    else
                    {
                        for (int j = 0; j < row.ValueArray.Length; j++)
                        {
                            this.VerifyAddressBookFlaggedPropertyValueStructure((AddressBookFlaggedPropertyValue)row.ValueArray[j]);
                        }
                    }
                }

                this.VerifyLargePropertyTagArrayStructure(queryRowsResponseBody.Columns.Value);
            }
            
            return queryRowsResponseBody;
        }

        /// <summary>
        /// This method is used by the client to get a list of all the properties that exist in the address book.
        /// </summary>
        /// <param name="queryColumnsRequestBody">The QueryColumns request type request body.</param>
        /// <returns>The response body of QueryColumns request type.</returns>
        public QueryColumnsResponseBody QueryColumns(QueryColumnsRequestBody queryColumnsRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(queryColumnsRequestBody, RequestType.QueryColumns);
            QueryColumnsResponseBody queryColumnsResponseBody = QueryColumnsResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyQueryColumnsResponseBody(queryColumnsResponseBody);
            this.VerifyLargePropertyTagArrayStructure(queryColumnsResponseBody.Columns.Value);

            return queryColumnsResponseBody;
        }

        /// <summary>
        /// This method is used by the client to perform ambiguous name resolution(ANR).
        /// </summary>
        /// <param name="resolveNamesRequestBody">The ResolveNames request type request body.</param>
        /// <returns>The response body of the ResolveNames request type.</returns>
        public ResolveNamesResponseBody ResolveNames(ResolveNamesRequestBody resolveNamesRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(resolveNamesRequestBody, RequestType.ResolveNames);
            ResolveNamesResponseBody resolveNamesResponseBody = ResolveNamesResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyResolveNamesResponseBody(resolveNamesResponseBody);
            if (resolveNamesResponseBody.HasRowsAndPropertyTags)
            {
                foreach (AddressBookPropertyRow row in resolveNamesResponseBody.RowData)
                {
                    this.VerifyAddressBookPropertyRowStructure(row);
                }

                this.VerifyLargePropertyTagArrayStructure(resolveNamesResponseBody.PropertyTags.Value);
            }

            return resolveNamesResponseBody;
        }

        /// <summary>
        /// This method is used by the client to sort the objects in the restricted address book container.
        /// </summary>
        /// <param name="resortRestrictionRequestBody">The ResortRestriction request type request body.</param>
        /// <returns>The response body of the ResortRestriction request type.</returns>
        public ResortRestrictionResponseBody ResortRestriction(ResortRestrictionRequestBody resortRestrictionRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(resortRestrictionRequestBody, RequestType.ResortRestriction);
            ResortRestrictionResponseBody resortRestrictionResponseBody = ResortRestrictionResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyResortRestrictionResponseBody(resortRestrictionResponseBody);

            return resortRestrictionResponseBody;
        }

        /// <summary>
        /// This method is used by the client to search for and set the logical position in a specific table to the first entry greater than or equal to a specified value.
        /// </summary>
        /// <param name="seekEntriesRequestBody">The SeekEntries request type request body.</param>
        /// <returns>The response body of SeekEntries request type.</returns>
        public SeekEntriesResponseBody SeekEntries(SeekEntriesRequestBody seekEntriesRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(seekEntriesRequestBody, RequestType.SeekEntries);
            SeekEntriesResponseBody seekEntriesResponseBody = SeekEntriesResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifySeekEntriesResponseBody(seekEntriesResponseBody);
            if (seekEntriesResponseBody.HasColumnsAndRows)
            {
                foreach (AddressBookPropertyRow row in seekEntriesResponseBody.RowData)
                {
                    this.VerifyAddressBookPropertyRowStructure(row);
                }

                this.VerifyLargePropertyTagArrayStructure(seekEntriesResponseBody.Columns.Value);
            }

            return seekEntriesResponseBody;
        }

        /// <summary>
        /// This method is used by the client to update the STAT structure to reflect the client's changes.
        /// </summary>
        /// <param name="updateStatRequestBody">The UpdateStat request type request body.</param>
        /// <returns>The response body of UpdateStat request type.</returns>
        public UpdateStatResponseBody UpdateStat(UpdateStatRequestBody updateStatRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(updateStatRequestBody, RequestType.UpdateStat);
            UpdateStatResponseBody updateStatResponseBody = UpdateStatResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyUpdateStatResponseBody(updateStatResponseBody);

            return updateStatResponseBody;
        }

        /// <summary>
        /// This method is used by the client to get the Uniform Resource Locator (URL) of the specified mailbox server endpoint.
        /// </summary>
        /// <param name="getMailboxUrlRequestBody">The GetMailboxUrl request type request body.</param>
        /// <returns>The response body of the GetMailboxUrl request type.</returns>
        public GetMailboxUrlResponseBody GetMailboxUrl(GetMailboxUrlRequestBody getMailboxUrlRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(getMailboxUrlRequestBody, RequestType.GetMailboxUrl, cookieChange: false);
            GetMailboxUrlResponseBody getMailboxUrlResponseBody = GetMailboxUrlResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyGetMailboxUrlResponseBody(getMailboxUrlResponseBody);

            return getMailboxUrlResponseBody;
        }

        /// <summary>
        /// This method is used by the client to get the URL of the specified address book server endpoint.
        /// </summary>
        /// <param name="getAddressBookUrlRequestBody">The GetAddressBookUrl request type request body.</param>
        /// <returns>The response body of GetAddressBookUrl request type.</returns>
        public GetAddressBookUrlResponseBody GetAddressBookUrl(GetAddressBookUrlRequestBody getAddressBookUrlRequestBody)
        {
            CommonResponse commonResponse = this.SendAddressBookRequest(getAddressBookUrlRequestBody, RequestType.GetAddressBookUrl, cookieChange: false);
            GetAddressBookUrlResponseBody getAddressBookUrlResponseBody = GetAddressBookUrlResponseBody.Parse(commonResponse.ResponseBodyRawData);
            this.VerifyGetAddressBookUrlResponseBody(getAddressBookUrlResponseBody);

            return getAddressBookUrlResponseBody;
        }

        #endregion

        /// <summary>
        /// Send the request to address book server endpoint. 
        /// </summary>
        /// <param name="requestBody">The request body.</param>
        /// <param name="requestType">The type of the request.</param>
        /// <param name="cookieChange">If the session context cookie changed.</param>
        /// <returns>The common response.</returns>
        private CommonResponse SendAddressBookRequest(IRequestBody requestBody, RequestType requestType, bool cookieChange = true)
        {
            byte[] rawBuffer;
            CommonResponse commonResponse = null;
            WebHeaderCollection webHeaderCollection = AdapterHelper.InitializeHTTPHeader(requestType, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            
            // Send the Execute HTTP request and get the response.
            HttpWebResponse response = this.SendMAPIHttpRequest(this.userName, this.password, requestBody, ServerEndpoint.AddressBookServerEndpoint, AdapterHelper.SessionContextCookies, webHeaderCollection, out rawBuffer);
            uint responseCode = AdapterHelper.GetFinalResponseCode(response.Headers["X-ResponseCode"]);
            this.Site.Assert.AreEqual<uint>(0, responseCode, "The request to the address book server should be executed successfully!");

            // Read the HTTP response buffer and parse the response to correct format.
            commonResponse = CommonResponse.ParseCommonResponse(rawBuffer);
            Site.Assert.IsNotNull(commonResponse.ResponseBodyRawData, "The response body should contains data.");
            this.VerifyRequestTypesForAddressBookServerEndpoint(response.Headers, commonResponse);
            this.VerifyAutoDiscover(response.StatusCode, ServerEndpoint.AddressBookServerEndpoint);
            this.VerifyAuthentication(response);

            response.GetResponseStream().Close();
            if (cookieChange)
            {
                AdapterHelper.SessionContextCookies = response.Cookies;  
            }

            return commonResponse;
        }

        /// <summary>
        /// This method is used to send the http request.
        /// </summary>
        /// <param name="userName">The user name used to connect with server.</param>
        /// <param name="password">The password used to connect with server.</param>
        /// <param name="requestBody">The request body.</param>
        /// <param name="endpoint">The endpoint which the request would be send to.</param>
        /// <param name="cookies">Cookies used to identify the Session Context.</param>
        /// <param name="resquestHeaders">The specified request header used by the request.</param>
        /// <param name="rawBuffer">The raw buffer of the response.</param>
        /// <returns>The response of the request.</returns>
        private HttpWebResponse SendMAPIHttpRequest(string userName, string password, IRequestBody requestBody, ServerEndpoint endpoint, CookieCollection cookies, WebHeaderCollection resquestHeaders, out byte[] rawBuffer)
        {
            rawBuffer = null;
            HttpWebResponse response = null;
            AdapterHelper.Counter++;
            string url = string.Empty;
            if (endpoint == ServerEndpoint.MailboxServerEndpoint)
            {
                if (string.IsNullOrEmpty(this.mailStoreUrl))
                {
                    this.GetEndpointUrl();
                }

                url = this.mailStoreUrl;
            }
            else
            {
                if (string.IsNullOrEmpty(this.addressBookUrl))
                {
                    this.GetEndpointUrl();
                }

                url = this.addressBookUrl;
            }

            System.Net.ServicePointManager.ServerCertificateValidationCallback =
            new System.Net.Security.RemoteCertificateValidationCallback(Common.ValidateServerCertificate);
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.CookieContainer = new CookieContainer();
            request.Method = "POST";
            request.ProtocolVersion = HttpVersion.Version11;
            request.ContentType = "application/mapi-http";
            request.Credentials = new NetworkCredential(userName, password, this.domainName);
            request.Headers.Add(resquestHeaders);
            request.CookieContainer.Add(cookies);
            request.Timeout = System.Threading.Timeout.Infinite;

            byte[] buffer = null;
            if (requestBody != null)
            {
                buffer = requestBody.Serialize();
                request.ContentLength = buffer.Length;
            }
            else
            {
                request.ContentLength = 0;
            }   

            if (requestBody != null)
            {
                using (Stream stream = request.GetRequestStream())
                {
                    stream.Write(buffer, 0, buffer.Length);
                }
            }

            try
            {
                response = request.GetResponse() as HttpWebResponse;

                // Read the HTTP response buffer and parse the response to correct format.
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    rawBuffer = this.ReadHttpResponse(response);
                }
            }
            catch (WebException ex)
            {
                this.Site.Log.Add(
                    LogEntryKind.Comment, 
                    "A WebException happened when connecting the server, The exception is {0}.",
                    ex.Message);
                return (HttpWebResponse)ex.Response;
            }

            return response;
        }

        /// <summary>
        /// Get the mailbox server endpoint URI and the address book server endpoint URI by Autodiscover.
        /// </summary>
        private void GetEndpointUrl()
        {
            string originalServerName = Common.GetConfigurationPropertyValue("SutComputerName", this.Site);
            string transportSequence = Common.GetConfigurationPropertyValue("TransportSeq", this.Site);
            string requestURL = Common.GetConfigurationPropertyValue("AutoDiscoverUrlFormat", this.Site);
            requestURL = Regex.Replace(requestURL, @"\[ServerName\]", originalServerName, RegexOptions.IgnoreCase);

            AutoDiscoverProperties autoDiscoverProperties = AutoDiscover.GetAutoDiscoverProperties(this.Site, originalServerName, this.userName, this.domainName, requestURL, transportSequence.ToLower());
            this.mailStoreUrl = autoDiscoverProperties.PrivateMailStoreUrl;
            this.addressBookUrl = autoDiscoverProperties.AddressBookUrl;
        }

        /// <summary>
        /// Read the HTTP response and get the response bytes.
        /// </summary>
        /// <param name="response">The HTTP response need be read.</param>
        /// <returns>The response bytes.</returns>
        private byte[] ReadHttpResponse(HttpWebResponse response)
        {
            Stream respStream = response.GetResponseStream();
            List<byte> responseBytesList = new List<byte>();

            int read;
            do
            {
                read = respStream.ReadByte();
                if (read != -1)
                {
                    byte singleByte = (byte)read;
                    responseBytesList.Add(singleByte);
                }
            }
            while (read != -1);

            return responseBytesList.ToArray();
        }

        #endregion MS-OXCMAPIHTTPAdapter Members
    }
}