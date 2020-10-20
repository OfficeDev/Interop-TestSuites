namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Net;
    using System.Net.Security;
    using System.Reflection;
    using System.Security;
    using System.Security.Cryptography.X509Certificates;
    using System.Security.Policy;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Xml;
    using System.Xml.Linq;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.MS_ASWBXML;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Provides all the methods defined in MS-ASCMD. 
    /// </summary>
    public partial class ActiveSyncClient
    {
        #region Private Fields
        /// <summary>
        /// MS_ASWBXMLSyntheticImplementation instance field
        /// </summary>
        private MS_ASWBXML msaswbxmlImplementation;

        /// <summary>
        /// string specified prefixOfURI, HTTP or HTTPS
        /// </summary>
        private string prefixOfURI;

        /// <summary>
        /// Current command name
        /// </summary>
        private CommandName commandName;

        /// <summary>
        /// The last XML request
        /// </summary>
        private IXPathNavigable lastRawRequestXml;

        /// <summary>
        /// The last XML response
        /// </summary>
        private IXPathNavigable lastRawResponseXml;

        /// <summary>
        /// An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.
        /// </summary>
        private ITestSite site;
        #endregion

        #region Contructors
        /// <summary>
        /// Initializes a new instance of the ActiveSyncClient class.
        /// </summary>
        /// <param name="testSite">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public ActiveSyncClient(ITestSite testSite)
        {
            this.site = testSite;
            this.Domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            this.DeviceID = Common.GetConfigurationPropertyValue("DeviceID", testSite);
            this.DeviceType = Common.GetConfigurationPropertyValue("DeviceType", testSite);
            this.PolicyKey = string.Empty;
            this.Host = Common.GetConfigurationPropertyValue("SutComputerName", testSite);
            this.Locale = int.Parse(Common.GetConfigurationPropertyValue("Locale", testSite));
            string activeSyncProtocolVersion = Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", testSite);
            string queryValueType = Common.GetConfigurationPropertyValue("HeaderEncodingType", testSite);
            string transportType = Common.GetConfigurationPropertyValue("TransportType", testSite);

            this.ActiveSyncProtocolVersion = Common.ConvertActiveSyncProtocolVersion(activeSyncProtocolVersion, testSite);

            if (string.Equals(queryValueType, "Base64", StringComparison.CurrentCultureIgnoreCase))
            {
                this.QueryValueType = QueryValueType.Base64;
            }
            else if (string.Equals(queryValueType, "PlainText", StringComparison.CurrentCultureIgnoreCase))
            {
                this.QueryValueType = QueryValueType.PlainText;
            }
            else
            {
                testSite.Assert.Fail(queryValueType + " is not a valid value of HeaderEncodingType property, the value should be Base64 or PlainText.");
            }

            if (string.Equals(transportType, "HTTP", StringComparison.CurrentCultureIgnoreCase))
            {
                this.ProtocolTransportType = ProtocolTransportType.HTTP;
            }
            else if (string.Equals(transportType, "HTTPS", StringComparison.CurrentCultureIgnoreCase))
            {
                this.ProtocolTransportType = ProtocolTransportType.HTTPS;
            }
            else
            {
                testSite.Assert.Fail(transportType + " is not a valid value of TransportType property, the value should be HTTP or HTTPS.");
            }

            this.prefixOfURI = transportType;
            this.ContentType = null;
            this.AcceptEncoding = null;
            this.UserAgent = null;
            this.AcceptMultiPart = null;
            this.msaswbxmlImplementation = new MS_ASWBXML(this.site);
            this.validationResult = true;
        }

        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the last XML request
        /// </summary>
        public IXPathNavigable LastRawRequestXml
        {
            get
            {
                return this.lastRawRequestXml;
            }
        }

        /// <summary>
        /// Gets the last XML response
        /// </summary>
        public IXPathNavigable LastRawResponseXml
        {
            get
            {
                return this.lastRawResponseXml;
            }
        }

        /// <summary>
        /// Gets or sets the ActiveSync protocol version
        /// </summary>
        public string ActiveSyncProtocolVersion { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether multipart is accepted in response for ItemOperation
        /// "T" means the client is requesting that the server return content in multipart format.
        /// "F" means the client is requesting that the server return content in inline format.
        /// </summary>
        public string AcceptMultiPart { get; set; }

        /// <summary>
        /// Gets or sets the host name
        /// </summary>
        public string Host { get; set; }

        /// <summary>
        /// Gets or sets the accept language code
        /// </summary>
        public string AcceptLanguage { get; set; }

        /// <summary>
        /// Gets or sets the local code
        /// </summary>
        public int Locale { get; set; }

        /// <summary>
        /// Gets or sets the device id
        /// </summary>
        public string DeviceID { get; set; }

        /// <summary>
        /// Gets or sets the device type
        /// </summary>
        public string DeviceType { get; set; }

        /// <summary>
        /// Gets or sets the policy key
        /// </summary>
        public string PolicyKey { get; set; }

        /// <summary>
        /// Gets or sets the user name
        /// </summary>
        public string UserName { get; set; }

        /// <summary>
        /// Gets or sets the password
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// Gets or sets the domain
        /// </summary>
        public string Domain { get; set; }

        /// <summary>
        /// Gets or sets ProtocolTransportType
        /// </summary>
        public ProtocolTransportType ProtocolTransportType { get; set; }

        /// <summary>
        /// Gets or sets Query Value Type
        /// </summary>
        public QueryValueType QueryValueType { get; set; }

        /// <summary>
        /// Gets or sets the Content-Type request header
        /// </summary>
        public string ContentType { get; set; }

        /// <summary>
        /// Gets or sets the AcceptEncoding request header
        /// </summary>
        public string AcceptEncoding { get; set; }

        /// <summary>
        /// Gets or sets the User-Agent request header
        /// </summary>
        public string UserAgent { get; set; }

        /// <summary>
        /// Gets or sets string specified prefixOfURI, HTTP or HTTPS
        /// </summary>
        public string PrefixOfURI
        {
            get
            {
                return this.prefixOfURI.ToUpper(CultureInfo.InvariantCulture);
            }

            set
            {
                this.prefixOfURI = (value ?? "https").ToLower(CultureInfo.InvariantCulture);
            }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Sends a plain text request.
        /// </summary>
        /// <param name="cmdName">The name of the command to send</param>
        /// <param name="parameters">The command parameters</param>
        /// <param name="request">The plain text request</param>
        /// <returns>The plain text response.</returns>
        public SendStringResponse SendStringRequest(CommandName cmdName, IDictionary<CmdParameterName, object> parameters, string request)
        {
            if (Common.GetSutVersion(this.site) == SutVersion.ExchangeServer2007)
            {
                this.site.Assume.AreNotEqual<string>("140", this.ActiveSyncProtocolVersion, "Exchange Server 2007 does not support ActiveSync protocol version 14.0.");

                this.site.Assume.AreNotEqual<string>("141", this.ActiveSyncProtocolVersion, "Exchange Server 2007 does not support ActiveSync protocol version 14.1");
            }

            ActiveSyncRawRequest rawRequest = ConfigureRawRequestCommandSetting(parameters, request);
            rawRequest.CommandName = cmdName;
            ActiveSyncRawResponse rawResponse;
            if (CommandName.Autodiscover == cmdName)
            {
                rawRequest.ContentType = this.ContentType ?? GetContenTypeString(ContentTypeEnum.Xml);

                rawResponse = this.SendAutodiscoverRequest(rawRequest, this.Host);
            }
            else
            {
                if (cmdName == CommandName.Ping && request.Contains("TestPlainText"))
                {
                    rawRequest.ContentType = "text/plain";
                }

                rawResponse = this.SendRequest(rawRequest);
            }

            if (rawResponse != null)
            {
                return ConvertRawResponse<SendStringResponse>(rawResponse);
            }
            else
            {
                return new SendStringResponse();
            }
        }

        /// <summary>
        /// Sends an Autodiscover request.
        /// </summary>
        /// <param name="request">An AutodiscoverRequest object that contains the request information.</param>
        /// <param name="contentType">Content Type that indicate the body's format</param>
        /// <returns>An Autodiscover object</returns>
        public AutodiscoverResponse Autodiscover(AutodiscoverRequest request, ContentTypeEnum contentType)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.Autodiscover;
            rawRequest.ContentType = GetContenTypeString(contentType);
            ActiveSyncRawResponse rawResponse = this.SendAutodiscoverRequest(rawRequest, this.Host);
            return ConvertRawResponse<AutodiscoverResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a Sync request.
        /// </summary>
        /// <param name="request">A SyncRequest object that contains the request information.</param>
        /// <returns>A Sync object</returns>
        public SyncResponse Sync(SyncRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.Sync;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<SyncResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a Sync request.
        /// </summary>
        /// <param name="request">A SyncRequest object that contains the request information.</param>
        /// <param name="isResyncNeeded">A boolean value indicate whether need to re-sync when the response contains MoreAvailable.</param>
        /// <returns>A SyncResponse object</returns>
        public SyncResponse Sync(SyncRequest request, bool isResyncNeeded)
        {
            if (!isResyncNeeded)
            {
                return this.Sync(request);
            }

            SyncResponse syncResponse;
            List<Response.SyncCollectionsCollectionCommandsAdd> commandsAdd = new List<Response.SyncCollectionsCollectionCommandsAdd>();
            List<Response.SyncCollectionsCollectionCommandsChange> commandsChange = new List<Response.SyncCollectionsCollectionCommandsChange>();
            List<Response.SyncCollectionsCollectionCommandsDelete> commandsDelete = new List<Response.SyncCollectionsCollectionCommandsDelete>();
            List<Response.SyncCollectionsCollectionCommandsSoftDelete> commandsSoftDelete = new List<Response.SyncCollectionsCollectionCommandsSoftDelete>();

            string commandsString = null;

            do
            {
                ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
                rawRequest.CommandName = CommandName.Sync;
                ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
                syncResponse = ConvertRawResponse<SyncResponse>(rawResponse);

                if (syncResponse.ResponseData.Item == null)
                {
                    return syncResponse;
                }

                if (syncResponse.ResponseData.Item.GetType() == typeof(Response.SyncCollections))
                {
                    Response.SyncCollectionsCollection collection = ((Response.SyncCollections)syncResponse.ResponseData.Item).Collection[0];

                    for (int i = 0; i < collection.ItemsElementName.Length; i++)
                    {
                        if (collection.ItemsElementName[i] == Response.ItemsChoiceType10.Status && !collection.Items[i].ToString().Equals("1"))
                        {
                            return syncResponse;
                        }
                        else if (collection.ItemsElementName[i] == Response.ItemsChoiceType10.SyncKey)
                        {
                            request.RequestData.Collections[0].SyncKey = collection.Items[i].ToString();
                        }
                        else if (collection.ItemsElementName[i] == Response.ItemsChoiceType10.Commands)
                        {
                            if (((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Add != null && ((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Add.Length > 0)
                            {
                                for (int j = 0; j < ((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Add.Length; j++)
                                {
                                    commandsAdd.Add(((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Add[j]);
                                }
                            }

                            if (((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Change != null && ((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Change.Length > 0)
                            {
                                for (int j = 0; j < ((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Change.Length; j++)
                                {
                                    commandsChange.Add(((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Change[j]);
                                }
                            }

                            if (((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Delete != null && ((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Delete.Length > 0)
                            {
                                for (int j = 0; j < ((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Delete.Length; j++)
                                {
                                    commandsDelete.Add(((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Delete[j]);
                                }
                            }

                            if (((Response.SyncCollectionsCollectionCommands)collection.Items[i]).SoftDelete != null && ((Response.SyncCollectionsCollectionCommands)collection.Items[i]).SoftDelete.Length > 0)
                            {
                                for (int j = 0; j < ((Response.SyncCollectionsCollectionCommands)collection.Items[i]).SoftDelete.Length; j++)
                                {
                                    commandsSoftDelete.Add(((Response.SyncCollectionsCollectionCommands)collection.Items[i]).SoftDelete[j]);
                                }
                            }

                            ((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Add = commandsAdd.ToArray();
                            ((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Change = commandsChange.ToArray();
                            ((Response.SyncCollectionsCollectionCommands)collection.Items[i]).Delete = commandsDelete.ToArray();
                            ((Response.SyncCollectionsCollectionCommands)collection.Items[i]).SoftDelete = commandsSoftDelete.ToArray();
                        }
                    }

                    ((Response.SyncCollections)syncResponse.ResponseData.Item).Collection[0] = collection;

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(syncResponse.ResponseDataXML);
                    XmlNodeList commandNodes = xmlDoc.GetElementsByTagName("Commands");
                    if (commandNodes.Count > 0)
                    {
                        for (int i = 0; i < commandNodes.Count; i++)
                        {
                            commandsString = commandsString + commandNodes[i].InnerXml;
                        }
                    }
                }
            }
            while (syncResponse.ResponseDataXML.Contains("<MoreAvailable />"));

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(syncResponse.ResponseDataXML);

            if (xmlDocument.GetElementsByTagName("Commands").Count > 0)
            {
                xmlDocument.GetElementsByTagName("Commands")[0].InnerXml = commandsString;
                syncResponse.ResponseDataXML = xmlDocument.OuterXml;
            }

            return syncResponse;
        }

        /// <summary>
        /// Find an email with specific subject.
        /// </summary>
        /// <param name="request">A SyncRequest object that contains the request information.</param>
        /// <param name="subject">The subject of the email to find</param>
        /// <param name="isRetryNeeded">A boolean whether need retry</param>
        /// <returns>The email with specific subject</returns>
        public DataStructures.Sync SyncEmail(SyncRequest request, string subject, bool isRetryNeeded)
        {
            SyncResponse syncResponse;

            if (!isRetryNeeded)
            {
                syncResponse = this.Sync(request, true);
                return FindEmail(syncResponse, subject);
            }

            DataStructures.Sync item = null;
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.site));
            int upperBound = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.site));

            while (counter < upperBound)
            {
                System.Threading.Thread.Sleep(waitTime);

                syncResponse = this.Sync(request, true);

                item = FindEmail(syncResponse, subject);

                if (item != null)
                {
                    break;
                }

                counter++;
            }

            if (item == null)
            {
                this.site.Assert.Fail("Can't find the email with subject {0} after retrying {1} times.", subject, counter);
            }

            return item;
        }

        /// <summary>
        /// Sends an HTTP OPTIONS command.
        /// </summary>
        /// <returns>HTTP OPTIONS response</returns>
        public OptionsResponse Options()
        {
            ActiveSyncRawResponse rawResponse = this.SendOptionRequest();
            OptionsResponse highLevelResponse = new OptionsResponse
            {
                StatusCode = rawResponse.StatusCode,
                StatusDescription = rawResponse.StatusDescription,
                Headers = rawResponse.Headers
            };
            return highLevelResponse;
        }

        /// <summary>
        /// Sends a SendMail request.
        /// </summary>
        /// <param name="request">A SendMailRequest object that contains the request information.</param>
        /// <returns>A SendMailResponse object</returns>
        public SendMailResponse SendMail(SendMailRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.SendMail;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<SendMailResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a GetAttachment request.
        /// </summary>
        /// <param name="request">A GetAttachmentRequest object that contains the request information.</param>
        /// <returns>A GetAttachmentResponse object</returns>
        public GetAttachmentResponse GetAttachment(GetAttachmentRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.GetAttachment;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<GetAttachmentResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a FolderSync request.
        /// </summary>
        /// <param name="request">A FolderSyncRequest object that contains the request information.</param>
        /// <returns>A FolderSyncResponse object</returns>
        public FolderSyncResponse FolderSync(FolderSyncRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.FolderSync;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<FolderSyncResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a FolderCreate request.
        /// </summary>
        /// <param name="request">A FolderCreateRequest object that contains the request information.</param>
        /// <returns>A FolderCreateResponse object</returns>
        public FolderCreateResponse FolderCreate(FolderCreateRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.FolderCreate;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<FolderCreateResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a FolderDelete request.
        /// </summary>
        /// <param name="request">A FolderDeleteRequest object that contains the request information.</param>
        /// <returns>A FolderDeleteResponse object</returns>
        public FolderDeleteResponse FolderDelete(FolderDeleteRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.FolderDelete;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<FolderDeleteResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a FolderUpdate request.
        /// </summary>
        /// <param name="request">A FolderUpdateRequest object that contains the request information.</param>
        /// <returns>A FolderUpdateResponse object</returns>
        public FolderUpdateResponse FolderUpdate(FolderUpdateRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.FolderUpdate;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<FolderUpdateResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a MoveItems request.
        /// </summary>
        /// <param name="request">A MoveItemsRequest object that contains the request information.</param>
        /// <returns>A MoveItemsResponse object</returns>
        public MoveItemsResponse MoveItems(MoveItemsRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.MoveItems;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<MoveItemsResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a GetHierarchy request.
        /// </summary>
        /// <returns>A GetHierarchyResponse object.</returns>
        public GetHierarchyResponse GetHierarchy()
        {
            object request = null;
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.GetHierarchy;
            rawRequest.HttpMethod = "POST";
            rawRequest.ContentType = GetContenTypeString(ContentTypeEnum.Wbxml);
            rawRequest.HttpRequestBody = string.Empty;

            rawRequest.SetCommandParameters(new Dictionary<CmdParameterName, object>());
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<GetHierarchyResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a GetItemEstimate request.
        /// </summary>
        /// <param name="request">A GetItemEstimateRequest object that contains the request information.</param>
        /// <returns>A GetItemEstimateResponse object</returns>
        public GetItemEstimateResponse GetItemEstimate(GetItemEstimateRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.GetItemEstimate;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<GetItemEstimateResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a MeetingResponse request.
        /// </summary>
        /// <param name="request">A MeetingResponseRequest object that contains the request information.</param>
        /// <returns>A MeetingResponseResponse object</returns>
        public MeetingResponseResponse MeetingResponse(MeetingResponseRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.MeetingResponse;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<MeetingResponseResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a Search request.
        /// </summary>
        /// <param name="request">A SearchRequest object that contains the request information.</param>
        /// <returns>A SearchResponse object</returns>
        public SearchResponse Search(SearchRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.Search;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<SearchResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a Search request with loop.
        /// </summary>
        /// <param name="request">A SearchRequest object that contains the request information.</param>
        /// <param name="isLoopNeeded">A boolean whether need the loop</param>
        /// <returns>A SearchResponse object</returns>
        public SearchResponse Search(SearchRequest request, bool isLoopNeeded)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            SearchResponse searchResponse = new SearchResponse();
            rawRequest.CommandName = CommandName.Search;
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.site));
            int upperBound = 1;

            if (isLoopNeeded)
            {
                upperBound = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.site));
            }

            while (counter < upperBound)
            {
                // Wait for the item received. 
                System.Threading.Thread.Sleep(waitTime);

                ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
                searchResponse = ConvertRawResponse<SearchResponse>(rawResponse);

                this.site.Assert.IsNotNull(searchResponse.ResponseData, "The Search should not be null!");

                if (!string.Equals(searchResponse.ResponseData.Status, "10", StringComparison.CurrentCultureIgnoreCase))
                {
                    this.site.Assert.AreEqual<string>("1", searchResponse.ResponseData.Status, "As a child element of Search, the Status element should be 1 which means success.");
                    this.site.Assert.IsNotNull(searchResponse.ResponseData.Response, "The SearchResponse element should not be null!");
                    this.site.Assert.IsNotNull(searchResponse.ResponseData.Response.Store, "The Store element should not be null!");
                    this.site.Assert.IsNotNull(searchResponse.ResponseData.Response.Store.Result, "The Result element in search response should not be null!");

                    if (searchResponse.ResponseData.Response.Store.Result.Length > 0 && searchResponse.ResponseData.Response.Store.Result[0].Class != null)
                    {
                        break;
                    }
                }

                counter++;
            }

            return searchResponse;
        }

        /// <summary>
        /// Sends a Search request with loop and the count of the items that expected to be found.
        /// </summary>
        /// <param name="request">A SearchRequest object that contains the request information.</param>
        /// <param name="isLoopNeeded">A boolean whether need the loop.</param>
        /// <param name="itemsCount">The count of the items that expected to be found.</param>
        /// <returns>A SearchResponse object</returns>
        public SearchResponse Search(SearchRequest request, bool isLoopNeeded, int itemsCount)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            SearchResponse searchResponse = new SearchResponse();
            rawRequest.CommandName = CommandName.Search;
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.site));
            int upperBound = 1;

            if (isLoopNeeded)
            {
                upperBound = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.site));
            }

            while (counter < upperBound)
            {
                // Wait for the item received. 
                System.Threading.Thread.Sleep(waitTime);

                ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
                searchResponse = ConvertRawResponse<SearchResponse>(rawResponse);

                this.site.Assert.IsNotNull(searchResponse.ResponseData, "The Search should not be null.");
                if (!string.Equals(searchResponse.ResponseData.Status, "10", StringComparison.CurrentCultureIgnoreCase))
                {
                    this.site.Assert.AreEqual<string>("1", searchResponse.ResponseData.Status, "As a child element of Search, the Status element should be 1 which means success.");
                    this.site.Assert.IsNotNull(searchResponse.ResponseData.Response, "The SearchResponse element should not be null.");
                    this.site.Assert.IsNotNull(searchResponse.ResponseData.Response.Store, "The Store element should not be null.");
                    this.site.Assert.IsNotNull(searchResponse.ResponseData.Response.Store.Result, "The Result element in search response should not be null.");

                    if (itemsCount == 0 && searchResponse.ResponseData.Response.Store.Result.Length == 1)
                    {
                        this.site.Assert.IsTrue(string.IsNullOrEmpty(searchResponse.ResponseData.Response.Store.Result[0].Class), "No item is expected to be found, the Class element should be null.");
                        this.site.Assert.IsTrue(string.IsNullOrEmpty(searchResponse.ResponseData.Response.Store.Result[0].CollectionId), "No item is expected to be found, the CollectionId element should be null.");
                        this.site.Assert.IsTrue(string.IsNullOrEmpty(searchResponse.ResponseData.Response.Store.Result[0].LongId), "No item is expected to be found, the LongId element should be null.");
                        this.site.Assert.IsNull(searchResponse.ResponseData.Response.Store.Result[0].Properties, "No item is expected to be found, the Properties element should be null.");

                        return searchResponse;
                    }

                    if (itemsCount > 0 && searchResponse.ResponseData.Response.Store.Result.Length > itemsCount)
                    {
                        this.site.Assert.Fail("The number of Result element should not exceed {0}, the actual value is {1}.", itemsCount, searchResponse.ResponseData.Response.Store.Result.Length);
                    }

                    if (itemsCount > 0 && searchResponse.ResponseData.Response.Store.Result.Length == itemsCount)
                    {
                        bool isEmptyResult = false;

                        for (int i = 0; i < itemsCount; i++)
                        {
                            if (string.IsNullOrEmpty(searchResponse.ResponseData.Response.Store.Result[i].Class))
                            {
                                isEmptyResult = true;
                            }
                        }

                        if (!isEmptyResult)
                        {
                            break;
                        }
                    }
                }

                counter++;
            }

            if (itemsCount > 0 && searchResponse.ResponseData.Response.Store.Result.Length == 1 && string.IsNullOrEmpty(searchResponse.ResponseData.Response.Store.Result[0].Class))
            {
                this.site.Assert.Fail("The number of non-empty Result element should be {0}, the actual value is 0.", itemsCount);
            }

            this.site.Assert.AreEqual<int>(itemsCount, searchResponse.ResponseData.Response.Store.Result.Length, "The number of Result element should be {0}.", itemsCount);

            return searchResponse;
        }

        /// <summary>
        /// Sends a Settings request.
        /// </summary>
        /// <param name="request">A SettingsRequest object that contains the request information.</param>
        /// <returns>A SettingsResponse object</returns>
        public SettingsResponse Settings(SettingsRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.Settings;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<SettingsResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a SmartForward request.
        /// </summary>
        /// <param name="request">A SmartForwardRequest object that contains the request information.</param>
        /// <returns>A SmartForwardResponse object</returns>
        public SmartForwardResponse SmartForward(SmartForwardRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.SmartForward;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<SmartForwardResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a SmartReply request.
        /// </summary>
        /// <param name="request">A SmartReplyRequest object that contains the request information.</param>
        /// <returns>A SmartReplyResponse object</returns>
        public SmartReplyResponse SmartReply(SmartReplyRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.SmartReply;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<SmartReplyResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a Ping request.
        /// </summary>
        /// <param name="request">A PingRequest object that contains the request information.</param>
        /// <returns>A PingResponse object</returns>
        public PingResponse Ping(PingRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.Ping;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<PingResponse>(rawResponse);
        }

        /// <summary>
        /// Sends an ItemOperations request.
        /// </summary>
        /// <param name="request">An ItemOperationsRequest object that contains the request information.</param>
        /// <param name="deliveryMethod">Delivery method specifies what kind of response is accepted.</param>
        /// <returns>An ItemOperationsResponse object</returns>
        public ItemOperationsResponse ItemOperations(ItemOperationsRequest request, DeliveryMethodForFetch deliveryMethod)
        {
            if (deliveryMethod == DeliveryMethodForFetch.MultiPart)
            {
                this.AcceptMultiPart = "T";

                if (this.QueryValueType == QueryValueType.Base64)
                {
                    // If command parameters is not specified. Create a new one.
                    if (request.CommandParameters == null)
                    {
                        request.SetCommandParameters(
                            new Dictionary<CmdParameterName, object>
                            {
                                {
                                    CmdParameterName.Options, 0
                                }
                            });
                    }

                    // If the second bit of command parameter option is not set, set it.
                    if (request.CommandParameters.Keys.Contains(CmdParameterName.Options) && (((int)request.CommandParameters[CmdParameterName.Options]) & 0x02) != 0x02)
                    {
                        request.CommandParameters[CmdParameterName.Options] = ((int)request.CommandParameters[CmdParameterName.Options]) | 0x02;
                    }
                }
                else
                {
                    request.SetCommandParameters(null);
                }
            }
            else
            {
                this.AcceptMultiPart = "F";
            }

            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.ItemOperations;

            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            ItemOperationsResponse response = ConvertRawResponse<ItemOperationsResponse>(rawResponse);
            this.AcceptMultiPart = null;

            // Parse metadata
            if (rawResponse != null && rawResponse.Headers != null && rawResponse.Headers.HasKeys() && rawResponse.Headers["Content-Type"].StartsWith("application/vnd.ms-sync.multipart", StringComparison.CurrentCultureIgnoreCase))
            {
                FieldInfo metadataField = typeof(ItemOperationsResponse).GetField("metadata", BindingFlags.NonPublic | BindingFlags.Instance);
                if (metadataField != null)
                {
                    metadataField.SetValue(response, new MultipartMetadata(ReadMetadata(response.RawBody)));
                }
            }

            return response;
        }

        /// <summary>
        /// Sends a Provision request.
        /// </summary>
        /// <param name="request">A ProvisionRequest object that contains the request information.</param>
        /// <returns>A ProvisionResponse object</returns>
        public ProvisionResponse Provision(ProvisionRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.Provision;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);

            // If overwirte for extent response type, just update this invoke.
            return ConvertRawResponse<ProvisionResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a ResolveRecipients request.
        /// </summary>
        /// <param name="request">A ResolveRecipientsRequest object that contains the request information.</param>
        /// <returns>A ResolveRecipientsResponse object</returns>
        public ResolveRecipientsResponse ResolveRecipients(ResolveRecipientsRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.ResolveRecipients;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<ResolveRecipientsResponse>(rawResponse);
        }

        /// <summary>
        /// Sends a ValidateCert request.
        /// </summary>
        /// <param name="request">A ValidateCertRequest object that contains the request information.</param>
        /// <returns>A ValidateCertResponse object</returns>
        public ValidateCertResponse ValidateCert(ValidateCertRequest request)
        {
            ActiveSyncRawRequest rawRequest = ConfigCmdRequest(request);
            rawRequest.CommandName = CommandName.ValidateCert;
            ActiveSyncRawResponse rawResponse = this.SendRequest(rawRequest);
            return ConvertRawResponse<ValidateCertResponse>(rawResponse);
        }

        /// <summary>
        /// Gets the MS_ASWBXML Instance.
        /// </summary>
        /// <returns>An MS_ASWBXML instance</returns>
        public MS_ASWBXML GetMSASWBXMLImplementationInstance()
        {
            return this.msaswbxmlImplementation;
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Verify the remote Secure Sockets Layer (SSL) certificate used for authentication.
        /// </summary>
        /// <param name="sender">An object that contains state information for this validation.</param>
        /// <param name="certificate">The certificate used to authenticate the remote party.</param>
        /// <param name="chain">The chain of certificate authorities associated with the remote certificate.</param>
        /// <param name="sslPolicyErrors">One or more errors associated with the remote certificate.</param>
        /// <returns>A boolean value that determines whether the specified certificate is accepted for authentication.</returns>
        private static bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            SslPolicyErrors errors = sslPolicyErrors;

            if ((errors & SslPolicyErrors.RemoteCertificateNameMismatch) == SslPolicyErrors.RemoteCertificateNameMismatch)
            {
                Zone zone = Zone.CreateFromUrl(((HttpWebRequest)sender).RequestUri.ToString());
                if (zone.SecurityZone == SecurityZone.Intranet || zone.SecurityZone == SecurityZone.MyComputer)
                {
                    errors -= SslPolicyErrors.RemoteCertificateNameMismatch;
                }
            }

            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) == SslPolicyErrors.RemoteCertificateChainErrors)
            {
                if (chain != null && chain.ChainStatus.Length != 0)
                {
                    foreach (X509ChainStatus status in chain.ChainStatus)
                    {
                        // Self-signed certificates have the issuer in the subject field. 
                        if ((certificate.Subject == certificate.Issuer) && (status.Status == X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else if (status.Status != X509ChainStatusFlags.NoError)
                        {
                            // If there are any other errors in the certificate chain, the certificate is invalid, the method returns false.
                            return false;
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are untrusted root errors for self-signed certificates. 
                // These certificates are valid.
                errors -= SslPolicyErrors.RemoteCertificateChainErrors;
            }

            return errors == SslPolicyErrors.None;
        }

        /// <summary>
        /// Configure the command request
        /// </summary>
        /// <typeparam name="T">The type parameter, defined in the ActiveSyncDataStructure</typeparam>
        /// <param name="highLevelRequest">The request instance</param>
        /// <returns>An ActiveSyncRawRequest object</returns>
        private static ActiveSyncRawRequest ConfigCmdRequest<T>(T highLevelRequest)
        {
            if (null == highLevelRequest)
            {
                return new ActiveSyncRawRequest();
            }

            object cmdParas = GetSpecifiedPropertyValueByName(highLevelRequest, "CommandParameters");
            object requestbody = InvokeSpecifiedMethod(highLevelRequest, "GetRequestDataSerializedXML");
            ActiveSyncRawRequest rawRequest = ConfigureRawRequestCommandSetting((IDictionary<CmdParameterName, object>)cmdParas, requestbody.ToString());
            return rawRequest;
        }

        /// <summary>
        /// Extract a content type string based on ContentTypeEnum
        /// </summary>
        /// <param name="contentType">Specified ContentTypeEnum</param>
        /// <returns>The content type string</returns>
        private static string GetContenTypeString(ContentTypeEnum contentType)
        {
            string contenTypeString = string.Empty;
            switch (contentType)
            {
                case ContentTypeEnum.Wbxml:
                    {
                        contenTypeString = @"application/vnd.ms-sync.wbxml";
                        break;
                    }

                case ContentTypeEnum.Xml:
                    {
                        contenTypeString = @"text/xml";
                        break;
                    }

                case ContentTypeEnum.Html:
                    {
                        contenTypeString = @"text/html";
                        break;
                    }
            }

            return contenTypeString;
        }

        /// <summary>
        /// Configure the command setting based on the command parameters 
        /// </summary>
        /// <param name="parameters">The command parameters</param>
        /// <param name="requestDataXML">The request XML string</param>
        /// <returns>An ActiveSyncRawRequest object</returns>
        private static ActiveSyncRawRequest ConfigureRawRequestCommandSetting(IDictionary<CmdParameterName, object> parameters, string requestDataXML)
        {
            ActiveSyncRawRequest rawRequest = new ActiveSyncRawRequest
            {
                HttpMethod = "POST",
                ContentType = GetContenTypeString(ContentTypeEnum.Wbxml),
                HttpRequestBody = requestDataXML
            };
            rawRequest.SetCommandParameters(parameters ?? new Dictionary<CmdParameterName, object>());
            return rawRequest;
        }

        /// <summary>
        /// Convert the ActiveSyncCmdRawResponse to the type T specified response
        /// </summary>
        /// <typeparam name="T">The type T, which is defined in the ActiveSyncDataStructure</typeparam>
        /// <param name="rawResponse">An ActiveSyncRawResponse object</param>
        /// <returns>The type T instance</returns>
        private static T ConvertRawResponse<T>(ActiveSyncRawResponse rawResponse)
        {
            T highLResponse = Activator.CreateInstance<T>();
            SetSpecifiedPropertyValueByName(highLResponse, "Headers", rawResponse.Headers);
            SetSpecifiedPropertyValueByName(highLResponse, "RawBody", rawResponse.AsHttpRawBody);
            SetSpecifiedPropertyValueByName(highLResponse, "ResponseDataXML", rawResponse.DecodedAsHttpBody);
            SetSpecifiedPropertyValueByName(highLResponse, "StatusCode", rawResponse.StatusCode);
            SetSpecifiedPropertyValueByName(highLResponse, "StatusDescription", rawResponse.StatusDescription);
            if (!(highLResponse is SendStringResponse) && !(highLResponse is GetAttachmentResponse))
            {
                InvokeSpecifiedMethod(highLResponse, "DeserializeResponseData");
            }

            return highLResponse;
        }

        /// <summary>
        /// Set a value in the target object using the specified property name
        /// </summary>
        /// <param name="targetObject">The target object</param>
        /// <param name="propertyName">The property name</param>
        /// <param name="value">The property value</param>
        private static void SetSpecifiedPropertyValueByName(object targetObject, string propertyName, object value)
        {
            if (string.IsNullOrEmpty(propertyName) || null == value || null == targetObject)
            {
                return;
            }

            PropertyInfo matchedProperty = targetObject.GetType().GetProperty(propertyName);

            if (matchedProperty != null)
            {
                matchedProperty.SetValue(targetObject, value, null);
            }
        }

        /// <summary>
        /// Get a value in the target object using the specified property name
        /// </summary>
        /// <param name="targetObject">The target object</param>
        /// <param name="propertyName">The property name value</param>
        /// <returns>The property value</returns>
        private static object GetSpecifiedPropertyValueByName(object targetObject, string propertyName)
        {
            if (string.IsNullOrEmpty(propertyName) || null == targetObject)
            {
                return null;
            }

            PropertyInfo matchedProperty = targetObject.GetType().GetProperty(propertyName);

            object value = null;
            if (matchedProperty != null)
            {
                value = matchedProperty.GetValue(targetObject, null);
            }

            return value;
        }

        /// <summary>
        /// Invoke a method in the target object name
        /// </summary>
        /// <param name="targetObject">The target object</param>
        /// <param name="methodName">The invoke method name</param>
        /// <returns>The method return value</returns>
        private static object InvokeSpecifiedMethod(object targetObject, string methodName)
        {
            if (string.IsNullOrEmpty(methodName) || null == targetObject)
            {
                return new object();
            }

            Type currentType = targetObject.GetType();
            MethodInfo matchedMethod = currentType.GetMethod(methodName);
            object invokedValue = matchedMethod.Invoke(targetObject, null);
            return invokedValue;
        }

        /// <summary>
        /// Reads metadata from multipart response
        /// </summary>
        /// <param name="bytes">The byte array that contains the multipart response</param>
        /// <returns>An integer array contains the metadata information</returns>
        private static int[] ReadMetadata(byte[] bytes)
        {
            if (bytes.Length < 12)
            {
                return null;
            }

            List<int> metadata = new List<int>();

            int numbersOfParts = bytes[0] | (bytes[1] << 8) | (bytes[2] << 16) | (bytes[3] << 24);
            metadata.Add(numbersOfParts);

            int startIndexOfWBXML;
            int countOfWBXML;
            for (int i = 4; i < numbersOfParts * 8; i += 8)
            {
                startIndexOfWBXML = bytes[i] | (bytes[i + 1] << 8) | (bytes[i + 2] << 16) | (bytes[i + 3] << 24);
                countOfWBXML = bytes[i + 4] | (bytes[i + 5] << 8) | (bytes[i + 6] << 16) | (bytes[i + 7] << 24);

                metadata.Add(startIndexOfWBXML);
                metadata.Add(countOfWBXML);
            }

            return metadata.ToArray();
        }

        /// <summary>
        /// Encode the command parameter
        /// </summary>
        /// <param name="data">The ActiveSyncCmdRawRequest data contains the command parameters</param>
        /// <returns>A IEnumerable byte sequence</returns>
        private static IEnumerable<byte> EncodeCmdParater(ActiveSyncRawRequest data)
        {
            List<byte> cmdParameterCode = new List<byte> { };
            foreach (KeyValuePair<CmdParameterName, object> cmdParameter in data.CommandParameters)
            {
                cmdParameterCode.Add(Convert.ToByte(cmdParameter.Key));

                byte[] arrValue;
                if (cmdParameter.Key == CmdParameterName.Options)
                {
                    arrValue = new byte[] { (byte)((int)cmdParameter.Value & 0xff) };
                }
                else
                {
                    arrValue = Encoding.UTF8.GetBytes(cmdParameter.Value.ToString());
                }

                cmdParameterCode.Add(Convert.ToByte(arrValue.Length));
                cmdParameterCode.AddRange(arrValue);
            }

            return cmdParameterCode;
        }

        /// <summary>
        /// Accept all the Certificate
        /// </summary>
        private static void AcceptAllCertificate()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            ServicePointManager.ServerCertificateValidationCallback
                = new RemoteCertificateValidationCallback(ValidateServerCertificate);
        }

        /// <summary>
        /// Read all the byte from the HttpWebResponse body
        /// </summary>
        /// <param name="rawRespone">An HttpWebResponse instance</param>
        /// <returns>A byte array of the HttpWebResponse body</returns>
        private static byte[] ReadRawBodyFromStream(HttpWebResponse rawRespone)
        {
            // Read rawresponse's rawbody
            Stream stream = rawRespone.GetResponseStream();

            int bodylength = (int)rawRespone.ContentLength;

            BinaryReader binReader = new BinaryReader(stream);
            return binReader.ReadBytes(bodylength);
        }

        /// <summary>
        /// Convert string data to XML data
        /// </summary>
        /// <param name="stringData">The string date to convert</param>
        /// <returns>The XML data</returns>
        private static IXPathNavigable ConvertStringToXml(string stringData)
        {
            if (!string.IsNullOrEmpty(stringData))
            {
                XmlDocument xmlDocOfReadRequest = new XmlDocument();
                xmlDocOfReadRequest.LoadXml(stringData);
                return xmlDocOfReadRequest.DocumentElement;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Find an email with specific subject.
        /// </summary>
        /// <param name="syncResponse">The sync response</param>
        /// <param name="subject">The subject of the email to find</param>
        /// <returns>The email with specific subject</returns>
        private static DataStructures.Sync FindEmail(SyncResponse syncResponse, string subject)
        {
            DataStructures.SyncStore syncStore = Common.LoadSyncResponse(syncResponse);

            DataStructures.Sync item = null;

            if (syncStore != null && syncStore.AddElements != null)
            {
                foreach (DataStructures.Sync syncItem in syncStore.AddElements)
                {
                    if (syncItem.Email.Subject == subject)
                    {
                        item = syncItem;
                        break;
                    }
                }
            }

            if (syncStore != null && syncStore.ChangeElements != null)
            {
                foreach (DataStructures.Sync syncItem in syncStore.ChangeElements)
                {
                    if (syncItem.Email.Subject == subject)
                    {
                        item = syncItem;
                        break;
                    }
                }
            }

            return item;
        }

        /// <summary>
        /// Extract base64 encoding query string
        /// </summary>
        /// <param name="data">An ActiveSyncRawRequest object</param>
        /// <returns>A base64 encoding query string</returns>
        private string Base64EncodedQuery(ActiveSyncRawRequest data)
        {
            List<byte> queryCode = new List<byte>
            {
                Convert.ToByte(this.ActiveSyncProtocolVersion),
                Convert.ToByte(data.CommandName)
            };

            byte[] locale = new byte[2];

            Array.Copy(BitConverter.GetBytes(this.Locale), locale, 2);

            queryCode.AddRange(locale);

            byte[] deviceID = Encoding.UTF8.GetBytes(this.DeviceID);
            queryCode.Add(Convert.ToByte(deviceID.Length));
            queryCode.AddRange(deviceID);

            uint policyKeyNumber;
            if (!string.IsNullOrEmpty(this.PolicyKey) && uint.TryParse(this.PolicyKey, out policyKeyNumber))
            {
                byte[] policyKey = BitConverter.GetBytes(policyKeyNumber);
                queryCode.Add(Convert.ToByte(policyKey.Length));
                queryCode.AddRange(policyKey);
            }
            else
            {
                queryCode.Add(Convert.ToByte(0));
            }

            byte[] deviceType = Encoding.UTF8.GetBytes(this.DeviceType);
            queryCode.Add(Convert.ToByte(deviceType.Length));
            queryCode.AddRange(deviceType);

            queryCode.AddRange(EncodeCmdParater(data));

            return Convert.ToBase64String(queryCode.ToArray());
        }

        /// <summary>
        /// Convert the HttpWebResponse to ActiveSyncCmdRawResponse
        /// </summary>
        /// <param name="rawResponse">The HttpWebResponse returned by the server</param>
        /// <returns>An ActiveSyncRawResponse object</returns>
        private ActiveSyncRawResponse ReadRawResponse(HttpWebResponse rawResponse)
        {
            this.lastRawResponseXml = null;
            string rawResponseString = string.Empty;
            byte[] byteTemp = null;
            ActiveSyncRawResponse response = new ActiveSyncRawResponse
            {
                StatusCode = rawResponse.StatusCode,
                StatusDescription = rawResponse.StatusDescription,
            };
            response.SetWebHeader(rawResponse.Headers);

            string transferEncoding = response.Headers["Transfer-Encoding"];

            if (transferEncoding != null && string.Equals(transferEncoding, "chunked"))
            {
                // Read chunked response stream
                byteTemp = this.ReadChunkedHttpResponse(rawResponse);
            }
            else
            {
                // Read non-chunked response stream
                byteTemp = ReadRawBodyFromStream(rawResponse);
            }

            response.AsHttpRawBody = byteTemp;

            // Decode the Body
            if (byteTemp != null && byteTemp.Length > 0)
            {
                rawResponseString = this.DecodeBody(byteTemp, rawResponse.ContentType, rawResponse.CharacterSet);
            }

            response.DecodedAsHttpBody = rawResponseString;
            Trace.TraceInformation("Response:\r\n===========================================\r\n" + rawResponseString + "\r\n===========================================\r\n");

            if (this.commandName == CommandName.GetAttachment)
            {
                this.ValidateResponseSchema(string.Empty, this.site);
            }
            else if (rawResponseString.Contains(@"<Picture xmlns=""GAL"">") || rawResponseString.Contains(@"<Picture>"))
            {
                string rawResponseString2 = this.RemovePictureData(rawResponseString, 0, 0);
                this.ValidateResponseSchema(rawResponseString2, this.site);
                this.lastRawResponseXml = ConvertStringToXml(rawResponseString2);
                response.DecodedAsHttpBody = rawResponseString2;
            }
            else
            {
                this.ValidateResponseSchema(rawResponseString, this.site);
                this.lastRawResponseXml = ConvertStringToXml(rawResponseString);
            }

            return response;
        }

        /// <summary>
        /// Decode the body from a WBXML binary array to a raw XML string
        /// </summary>
        /// <param name="rawData">A WBXML binary format byte array</param>
        /// <param name="contentType">The HTTP content type</param>
        /// <param name="characterSet">The charset</param>
        /// <returns>An XML format string</returns>
        private string DecodeBody(byte[] rawData, string contentType, string characterSet)
        {
            // Decode the RawBody according to the contentType
            string decodedata;

            if (contentType.StartsWith("application/vnd.ms-sync.wbxml", StringComparison.OrdinalIgnoreCase))
            {
                if (null == this.msaswbxmlImplementation)
                {
                    throw new MissingFieldException("The wbxmlSyntheticImplementation is not specified");
                }

                this.msaswbxmlImplementation = new MS_ASWBXML(this.site);
                decodedata = this.msaswbxmlImplementation.DecodeToXml(rawData);
            }
            else if (contentType.StartsWith("application/vnd.ms-sync.multipart", StringComparison.CurrentCultureIgnoreCase))
            {
                int[] metadata = ReadMetadata(rawData);

                if (null == this.msaswbxmlImplementation)
                {
                    throw new MissingFieldException("The wbxmlSyntheticImplementation is not specified");
                }

                byte[] wbxmlBytes = new byte[metadata[2]];
                Array.Copy(rawData, metadata[1], wbxmlBytes, 0, metadata[2]);
                this.msaswbxmlImplementation = new MS_ASWBXML(this.site);
                decodedata = this.msaswbxmlImplementation.DecodeToXml(wbxmlBytes);
            }
            else
            {
                Encoding encoder;
                characterSet = characterSet.ToLower(CultureInfo.CurrentCulture);
                switch (characterSet)
                {
                    case "utf-7":
                        {
                            encoder = Encoding.UTF7;
                            break;
                        }

                    case "utf-8":
                        {
                            encoder = Encoding.UTF8;
                            break;
                        }

                    case "unicode":
                        {
                            encoder = Encoding.Unicode;
                            break;
                        }

                    case "ascii":
                    default:
                        {
                            encoder = Encoding.ASCII;
                            break;
                        }
                }

                decodedata = encoder.GetString(rawData);
            }

            return decodedata;
        }

        /// <summary>
        /// Replace data element binary content with size value
        /// </summary>
        /// <param name="rawResponseString">The original raw response string returned from server</param>
        /// <param name="start">The data tag start position</param>
        /// <param name="end">The data tag end position</param>
        /// <returns>The rawResponseString with replaced data</returns>
        private string RemovePictureData(string rawResponseString, int start, int end)
        {
            int startIndex = rawResponseString.IndexOf("<Data>", start, StringComparison.CurrentCulture);
            if (end < rawResponseString.Length)
            {
                int endIndex = rawResponseString.IndexOf("</Data>", end, StringComparison.CurrentCulture);
                if (startIndex > 0 && endIndex > 0)
                {
                    int length = endIndex - startIndex - 6;
                    string dataString = rawResponseString.Substring(startIndex + 6, length);
                    string replaceString = length.ToString();
                    string changeString = rawResponseString.Replace(dataString, replaceString);

                    // After replace photo data string, the </Data> position is also changed.
                    int changeEndIndex = endIndex + 7 - length + replaceString.Length;
                    return this.RemovePictureData(changeString, startIndex + 1, changeEndIndex);
                }
            }
            else
            {
                return rawResponseString;
            }

            return rawResponseString;
        }

        /// <summary>
        /// Using POST method to send all commands except AutoDiscover command defined in MS-ASCMD.
        /// </summary>
        /// <param name="requestData">The request data</param>
        /// <returns>The response data from the server</returns>
        private ActiveSyncRawResponse SendRequest(ActiveSyncRawRequest requestData)
        {
            if (Common.GetSutVersion(this.site) == SutVersion.ExchangeServer2007)
            {
                this.site.Assume.AreNotEqual<string>("140", this.ActiveSyncProtocolVersion, "Exchange Server 2007 does not support ActiveSync protocol version 14.0.");

                this.site.Assume.AreNotEqual<string>("141", this.ActiveSyncProtocolVersion, "Exchange Server 2007 does not support ActiveSync protocol version 14.1");
            }

            this.lastRawRequestXml = null;
            this.commandName = requestData.CommandName;
            string url = this.GetQueryString(this.QueryValueType, requestData);
            AcceptAllCertificate();

            // Generate an HttpWebReuqest
            HttpWebRequest httpWebRequest = this.GetHttpWebRequest(requestData, url);

            // Add headers for plainText Query
            if (QueryValueType.PlainText == this.QueryValueType)
            {
                httpWebRequest.Headers.Add("Accept-Language", this.AcceptLanguage);
                double plainTextFormatVersion = double.Parse(this.ActiveSyncProtocolVersion) / 10;
                httpWebRequest.Headers.Add("MS-ASProtocolVersion", plainTextFormatVersion.ToString("00.0"));

                // Add X-MS-PolicyKey if need.
                if (!string.IsNullOrEmpty(this.PolicyKey))
                {
                    httpWebRequest.Headers.Add("X-MS-PolicyKey", this.PolicyKey);
                }

                // Add header for multipart request.
                if (this.AcceptMultiPart != null)
                {
                    httpWebRequest.Headers.Add("MS-ASAcceptMultiPart", this.AcceptMultiPart);
                }
            }

            if (this.AcceptEncoding != null)
            {
                httpWebRequest.Headers.Add(HttpRequestHeader.AcceptEncoding, "gzip");
            }

            if (this.UserAgent != null)
            {
                httpWebRequest.UserAgent = this.UserAgent;
            }

            byte[] bodyByte;

            if (string.IsNullOrEmpty(requestData.HttpRequestBody))
            {
                bodyByte = null;
                Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + "(Empty request body)" + "\r\n===========================================\r\n");
            }
            else if (this.ActiveSyncProtocolVersion.StartsWith("12", StringComparison.CurrentCultureIgnoreCase) && (requestData.CommandName == CommandName.SendMail || requestData.CommandName == CommandName.SmartForward || requestData.CommandName == CommandName.SmartReply))
            {
                httpWebRequest.ContentType = @"message/rfc822";
                Regex regEx = new Regex(@"\<!\[CDATA\[(.+)?\]\]\>", RegexOptions.Singleline);
                Match match = regEx.Match(requestData.HttpRequestBody);
                bodyByte = Encoding.UTF8.GetBytes(match.Groups[1].Value);
                Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + match.Groups[1].Value + "\r\n===========================================\r\n");
            }
            else if (requestData.CommandName == CommandName.Sync)
            {
                XmlElement xmlElement = (XmlElement)ConvertStringToXml(requestData.HttpRequestBody);

                if (xmlElement != null)
                {
                    if (xmlElement.HasChildNodes)
                    {
                        for (int i = 0; i < xmlElement.ChildNodes.Count; i++)
                        {
                            if (xmlElement.ChildNodes[i].HasChildNodes && xmlElement.ChildNodes[i].Name.Equals("Collections", StringComparison.CurrentCultureIgnoreCase))
                            {
                                for (int j = 0; j < xmlElement.ChildNodes[i].ChildNodes.Count; j++)
                                {
                                    if (xmlElement.ChildNodes[i].ChildNodes[j].HasChildNodes && xmlElement.ChildNodes[i].ChildNodes[j].Name.Equals("Collection", StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        for (int k = 0; k < xmlElement.ChildNodes[i].ChildNodes[j].ChildNodes.Count; k++)
                                        {
                                            if (xmlElement.ChildNodes[i].ChildNodes[j].ChildNodes[k].HasChildNodes && xmlElement.ChildNodes[i].ChildNodes[j].ChildNodes[k].Name.Equals("Supported", StringComparison.CurrentCultureIgnoreCase))
                                            {
                                                foreach (XmlNode node in xmlElement.ChildNodes[i].ChildNodes[j].ChildNodes[k].ChildNodes)
                                                {
                                                    node.RemoveAll();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    this.lastRawRequestXml = xmlElement;

                    if (xmlElement.PreviousSibling != null && !string.IsNullOrEmpty(xmlElement.PreviousSibling.OuterXml))
                    {
                        if (!string.IsNullOrEmpty(xmlElement.OuterXml))
                        {
                            Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + xmlElement.PreviousSibling.OuterXml + "\r\n" + XElement.Parse(xmlElement.OuterXml).ToString() + "\r\n===========================================\r\n");
                        }
                        else
                        {
                            Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + xmlElement.PreviousSibling.OuterXml + "\r\n" + xmlElement.OuterXml + "\r\n===========================================\r\n");
                        }

                        bodyByte = this.EncodeBody(xmlElement.PreviousSibling.OuterXml + xmlElement.OuterXml, httpWebRequest.ContentType);
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(xmlElement.OuterXml))
                        {
                            Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + XElement.Parse(xmlElement.OuterXml).ToString() + "\r\n===========================================\r\n");
                        }
                        else
                        {
                            Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + xmlElement.OuterXml + "\r\n===========================================\r\n");
                        }

                        bodyByte = this.EncodeBody(xmlElement.OuterXml, httpWebRequest.ContentType);
                    }
                }
                else
                {
                    this.lastRawRequestXml = null;
                    bodyByte = null;
                }
            }
            else if (requestData.CommandName == CommandName.ItemOperations)
            {
                XmlElement xmlElement = (XmlElement)ConvertStringToXml(requestData.HttpRequestBody);

                if (xmlElement != null)
                {
                    for (int i = 0; i < xmlElement.ChildNodes.Count; i++)
                    {
                        if (xmlElement.ChildNodes[i].Name.Equals("Fetch", StringComparison.CurrentCultureIgnoreCase))
                        {
                            for (int j = 0; j < xmlElement.ChildNodes[i].ChildNodes.Count; j++)
                            {
                                if (xmlElement.ChildNodes[i].ChildNodes[j].Name.Equals("Options", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    for (int k = 0; k < xmlElement.ChildNodes[i].ChildNodes[j].ChildNodes.Count; k++)
                                    {
                                        if (xmlElement.ChildNodes[i].ChildNodes[j].ChildNodes[k].Name.Equals("Schema", StringComparison.CurrentCultureIgnoreCase))
                                        {
                                            foreach (XmlNode node in xmlElement.ChildNodes[i].ChildNodes[j].ChildNodes[k].ChildNodes)
                                            {
                                                node.RemoveAll();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    this.lastRawRequestXml = xmlElement;

                    if (xmlElement.PreviousSibling != null && !string.IsNullOrEmpty(xmlElement.PreviousSibling.OuterXml))
                    {
                        if (!string.IsNullOrEmpty(xmlElement.OuterXml))
                        {
                            Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + xmlElement.PreviousSibling.OuterXml + "\r\n" + XElement.Parse(xmlElement.OuterXml).ToString() + "\r\n===========================================\r\n");
                        }
                        else
                        {
                            Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + xmlElement.PreviousSibling.OuterXml + "\r\n" + xmlElement.OuterXml + "\r\n===========================================\r\n");
                        }

                        bodyByte = this.EncodeBody(xmlElement.PreviousSibling.OuterXml + xmlElement.OuterXml, httpWebRequest.ContentType);
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(xmlElement.OuterXml))
                        {
                            Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + XElement.Parse(xmlElement.OuterXml).ToString() + "\r\n===========================================\r\n");
                        }
                        else
                        {
                            Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + xmlElement.OuterXml + "\r\n===========================================\r\n");
                        }

                        bodyByte = this.EncodeBody(xmlElement.OuterXml, httpWebRequest.ContentType);
                    }
                }
                else
                {
                    this.lastRawRequestXml = null;
                    bodyByte = null;
                }
            }
            else
            {
                this.lastRawRequestXml = ConvertStringToXml(requestData.HttpRequestBody);

                if (!string.IsNullOrEmpty(requestData.HttpRequestBody))
                {
                    Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + XElement.Parse(requestData.HttpRequestBody).ToString() + "\r\n===========================================\r\n");
                }
                else
                {
                    Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + requestData.HttpRequestBody + "\r\n===========================================\r\n");
                }

                bodyByte = this.EncodeBody(requestData.HttpRequestBody, httpWebRequest.ContentType);
            }

            if (null != bodyByte && bodyByte.Length > 0)
            {
                httpWebRequest.ContentLength = bodyByte.Length;
                httpWebRequest.GetRequestStream().Write(bodyByte, 0, bodyByte.Length);
                httpWebRequest.GetRequestStream().Close();
            }
            else
            {
                httpWebRequest.ContentType = string.Empty;
            }

            HttpWebResponse httpWebRawResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            ActiveSyncRawResponse responsedata = this.ReadRawResponse(httpWebRawResponse);
            return responsedata;
        }

        /// <summary>
        /// Send the Autodiscover command defined in the MS-ASCMD
        /// </summary>
        /// <param name="data">The ActiveSyncCmdRawRequest data instance</param>
        /// <param name="autoDiscoverServerName">The auto discover server url</param>
        /// <returns>The response data from the server</returns>
        private ActiveSyncRawResponse SendAutodiscoverRequest(ActiveSyncRawRequest data, string autoDiscoverServerName)
        {
            if (Common.GetSutVersion(this.site) == SutVersion.ExchangeServer2007)
            {
                this.site.Assume.AreNotEqual<string>("140", this.ActiveSyncProtocolVersion, "Exchange Server 2007 does not support ActiveSync protocol version 14.0.");

                this.site.Assume.AreNotEqual<string>("141", this.ActiveSyncProtocolVersion, "Exchange Server 2007 does not support ActiveSync protocol version 14.1");
            }

            this.lastRawRequestXml = null;
            string url = string.Format(@"{0}://{1}/{2}", this.prefixOfURI, autoDiscoverServerName, Common.GetConfigurationPropertyValue("AutodiscoverEndPoint", this.site));
            AcceptAllCertificate();

            HttpWebRequest httpWebRequest = this.GetHttpWebRequest(data, url);
            this.lastRawRequestXml = ConvertStringToXml(data.HttpRequestBody);

            if (!string.IsNullOrEmpty(data.HttpRequestBody))
            {
                Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + XElement.Parse(data.HttpRequestBody).ToString() + "\r\n===========================================\r\n");
            }
            else
            {
                Trace.TraceInformation("Request(URL : " + url + "):\r\n===========================================\r\n" + data.HttpRequestBody + "\r\n===========================================\r\n");
            }

            // Get Response
            // Get the requestStream and write the body to the Server
            byte[] bodyByte = this.EncodeBody(data.HttpRequestBody, httpWebRequest.ContentType);

            // Write stream to Server
            if (null != bodyByte && bodyByte.Length > 0)
            {
                httpWebRequest.ContentLength = bodyByte.Length;
                httpWebRequest.GetRequestStream().Write(bodyByte, 0, bodyByte.Length);
                httpWebRequest.GetRequestStream().Close();
            }

            // Get response from Server
            HttpWebResponse httpWebRawResponse = (HttpWebResponse)httpWebRequest.GetResponse();

            // Read RawWebResponse
            ActiveSyncRawResponse responseTemp = this.ReadRawResponse(httpWebRawResponse);

            return responseTemp;
        }

        /// <summary>
        /// Send the MS-ASCMD commands using HTTP Option method
        /// </summary>
        /// <returns>The response data from the server</returns>
        private ActiveSyncRawResponse SendOptionRequest()
        {
            if (Common.GetSutVersion(this.site) == SutVersion.ExchangeServer2007)
            {
                this.site.Assume.AreNotEqual<string>("140", this.ActiveSyncProtocolVersion, "Exchange Server 2007 does not support ActiveSync protocol version 14.0.");

                this.site.Assume.AreNotEqual<string>("141", this.ActiveSyncProtocolVersion, "Exchange Server 2007 does not support ActiveSync protocol version 14.1");
            }

            string url = string.Format(@"{0}://{1}/{2}", this.prefixOfURI, this.Host, Common.GetConfigurationPropertyValue("ActiveSyncEndPoint", this.site));

            HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            AcceptAllCertificate();
            httpWebRequest.Method = "OPTIONS";
            CredentialCache cache = new CredentialCache
            {
                {
                    new Uri(url), "Basic", new NetworkCredential(this.UserName, this.Password, this.Domain)
                }
            };

            httpWebRequest.Credentials = cache;
            httpWebRequest.PreAuthenticate = true;

            HttpWebResponse httpWebRawResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            ActiveSyncRawResponse optionsRawResponse = new ActiveSyncRawResponse
            {
                StatusCode = httpWebRawResponse.StatusCode,
                StatusDescription = httpWebRawResponse.StatusDescription,
            };
            optionsRawResponse.SetWebHeader(httpWebRawResponse.Headers);
            return optionsRawResponse;
        }

        /// <summary>
        /// Encode the XML string to the WBXML format when the content type contains application/vnd.ms-sync.wbxml
        /// </summary>
        /// <param name="body">The XML string</param>
        /// <param name="contentType">An HTTP contentType</param>
        /// <returns>The byte array which is defined in MS-WBXML</returns>
        private byte[] EncodeBody(string body, string contentType)
        {
            byte[] rawData;
            if (string.IsNullOrEmpty(body))
            {
                return new byte[0];
            }

            // Encode according to the contentType
            if (contentType.StartsWith("application/vnd.ms-sync.wbxml", StringComparison.OrdinalIgnoreCase))
            {
                if (null == this.msaswbxmlImplementation)
                {
                    throw new MissingFieldException("The wbxmlSyntheticImplementation is not specified");
                }

                this.msaswbxmlImplementation = new MS_ASWBXML(this.site);
                rawData = this.msaswbxmlImplementation.EncodeToWBXML(body);
            }
            else
            {
                rawData = Encoding.UTF8.GetBytes(body);
            }

            return rawData;
        }

        /// <summary>
        /// Construct an HttpWebRequest based on the ActiveSyncCmdRawRequest data
        /// </summary>
        /// <param name="requestdata">The ActiveSyncCmdRawRequest data</param>
        /// <param name="url">Http request URL</param>
        /// <returns>An HttpWebRequest request instance</returns>
        private HttpWebRequest GetHttpWebRequest(ActiveSyncRawRequest requestdata, string url)
        {
            // Config request headers and other contents
            HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            httpWebRequest.Method = requestdata.HttpMethod;
            httpWebRequest.ContentType = requestdata.ContentType;
            CredentialCache cache = new CredentialCache
            {
                {
                    new Uri(url), "Basic", new NetworkCredential(this.UserName, this.Password, this.Domain)
                }
            };

            httpWebRequest.Credentials = cache;
            httpWebRequest.ContentLength = 0;
            httpWebRequest.PreAuthenticate = true;
            return httpWebRequest;
        }

        /// <summary>
        /// Get the query string based on the QueryValueType
        /// </summary>
        /// <param name="queryvalueType">The query string type, either Base64 or PlainText</param>
        /// <param name="requestdata">The request data contains the query string</param>
        /// <returns>A query string</returns>
        private string GetQueryString(QueryValueType queryvalueType, ActiveSyncRawRequest requestdata)
        {
            string url;
            if (requestdata.CommandName == CommandName.GetHierarchy)
            {
                queryvalueType = QueryValueType.PlainText;
            }

            switch (queryvalueType)
            {
                case QueryValueType.PlainText:
                    {
                        url = this.GetPlainTextQueryString(requestdata);
                        break;
                    }

                case QueryValueType.Base64:
                    {
                        url = this.GetBase64QueryString(requestdata);
                        break;
                    }

                default:
                    throw new InvalidOperationException(string.Format("Not supported encode header type{0}", queryvalueType.ToString()));
            }

            return url;
        }

        /// <summary>
        /// Extract plain base64 query string from ActiveSyncCmdRawRequest data
        /// </summary>
        /// <param name="requestdata">The specified ActiveSyncCmdRawRequest instance</param>
        /// <returns>A base64 query format query string</returns>
        private string GetBase64QueryString(ActiveSyncRawRequest requestdata)
        {
            string url = string.Format(@"{0}://{1}/{2}", this.prefixOfURI, this.Host, Common.GetConfigurationPropertyValue("ActiveSyncEndPoint", this.site));
            if ("POST".Equals(requestdata.HttpMethod))
            {
                url += "?" + this.Base64EncodedQuery(requestdata);
            }

            return url;
        }

        /// <summary>
        /// Extract plain text format query string from ActiveSyncCmdRawRequest data
        /// </summary>
        /// <param name="requestdata">The specified ActiveSyncCmdRawRequest instance</param>
        /// <returns>A plain text format query string</returns>
        private string GetPlainTextQueryString(ActiveSyncRawRequest requestdata)
        {
            StringBuilder strBuilder = new StringBuilder();
            strBuilder.AppendFormat(@"{0}://{1}/{2}", this.prefixOfURI, this.Host, Common.GetConfigurationPropertyValue("ActiveSyncEndPoint", this.site));
            strBuilder.AppendFormat(@"?Cmd={0}", requestdata.CommandName);
            strBuilder.AppendFormat(@"&User={0}", this.UserName);
            strBuilder.AppendFormat(@"&DeviceId={0}", this.DeviceID);
            strBuilder.AppendFormat(@"&DeviceType={0}", this.DeviceType);

            // Add command parameters if existed
            if (requestdata.CommandParameters != null && requestdata.CommandParameters.Count > 0)
            {
                foreach (KeyValuePair<CmdParameterName, object> cmdParaItem in requestdata.CommandParameters)
                {
                    strBuilder.AppendFormat(@"&{0}={1}", cmdParaItem.Key, cmdParaItem.Value);
                }
            }

            return strBuilder.ToString();
        }

        /// <summary>
        /// Get the binary data from chunked http response.
        /// </summary>
        /// <param name="response">The structure of the response data.</param>
        /// <returns>Returns the binary format of response data.</returns>
        private byte[] ReadChunkedHttpResponse(HttpWebResponse response)
        {
            Stream responseStream = response.GetResponseStream();
            List<byte> responseBytesList = new List<byte>();

            int read;
            do
            {
                read = responseStream.ReadByte();
                if (read != -1)
                {
                    byte singleByte = (byte)read;
                    responseBytesList.Add(singleByte);
                }
            }
            while (read != -1);

            return responseBytesList.ToArray();
        }

        #endregion
    }
}