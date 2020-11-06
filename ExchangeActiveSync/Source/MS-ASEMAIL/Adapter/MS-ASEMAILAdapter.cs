namespace Microsoft.Protocols.TestSuites.MS_ASEMAIL
{
    using System;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-ASEMAIL.
    /// </summary>
    public partial class MS_ASEMAILAdapter : ManagedAdapterBase, IMS_ASEMAILAdapter
    {
        #region private field
        /// <summary>
        /// Active synctive client
        /// </summary>
        private ActiveSyncClient activeSyncClient;
        #endregion

        #region IMS_ASEMAILAdapter Properties
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT.
        /// </summary>
        public IXPathNavigable LastRawRequestXml
        {
            get { return this.activeSyncClient.LastRawRequestXml; }
        }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT.
        /// </summary>
        public IXPathNavigable LastRawResponseXml
        {
            get { return this.activeSyncClient.LastRawResponseXml; }
        }
        #endregion

        #region IMS_ASEMAILAdapter Commands
        /// <summary>
        /// Sync data from the server.
        /// </summary>
        /// <param name="syncRequest">The request for sync operation.</param>
        /// <returns>The sync result which is returned from server.</returns>
        public SyncStore Sync(SyncRequest syncRequest)
        {
            SyncResponse response = this.activeSyncClient.Sync(syncRequest, true);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");
            SyncStore result = Common.LoadSyncResponse(response);
            this.VerifyTransport();
            this.VerifySyncCommand(result);
            this.VerifyWBXMLCapture();
            return result;
        }

        /// <summary>
        /// Sync data from the server with an invalid sync request which contains additional element.
        /// </summary>
        /// <param name="syncRequest">The request for sync operation.</param>
        /// <param name="addElement">Additional element insert into normal sync request.</param>
        /// <param name="insertTag">Insert tag shows where the additional element should inserted.</param>
        /// <returns>The sync result which is returned from server.</returns>
        public SendStringResponse InvalidSync(SyncRequest syncRequest, string addElement, string insertTag)
        {
            string syncXmlRequest = syncRequest.GetRequestDataSerializedXML();
            string changedSyncXmlRequest = syncXmlRequest.Insert(syncXmlRequest.IndexOf(insertTag, StringComparison.CurrentCulture), addElement);
            SendStringResponse result = this.activeSyncClient.SendStringRequest(CommandName.Sync, null, changedSyncXmlRequest);
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// Sends MIME-formatted e-mail messages to the server.
        /// </summary>
        /// <param name="sendMailRequest">The request for SendMail operation.</param>
        /// <returns>The SendMail response which is returned from the server.</returns>
        public SendMailResponse SendMail(SendMailRequest sendMailRequest)
        {
            SendMailResponse response = this.activeSyncClient.SendMail(sendMailRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");
            this.VerifyTransport();
            return response;
        }

        /// <summary>
        /// MeetingResponse for accepting or declining a MeetingRequest.
        /// </summary>
        /// <param name="meetingResponseRequest">The request for meeting.</param>
        /// <returns>The meeting response which is returned from server.</returns>
        public MeetingResponseResponse MeetingResponse(MeetingResponseRequest meetingResponseRequest)
        {
            MeetingResponseResponse response = this.activeSyncClient.MeetingResponse(meetingResponseRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");
            this.VerifyTransport();
            return response;
        }

        /// <summary>
        /// Search items on server.
        /// </summary>
        /// <param name="searchRequest">The request for search operation.</param>
        /// <returns>The search response which is returned from the server.</returns>
        public SearchResponse Search(SearchRequest searchRequest)
        {
            SearchResponse response = this.activeSyncClient.Search(searchRequest, true);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            SearchStore store = Common.LoadSearchResponse(response, Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site));
            this.VerifyTransport();
            this.VerifySearchCommand(store);
            this.VerifyWBXMLCapture();
            return response;
        }

        /// <summary>
        /// Find items on server.
        /// </summary>
        /// <param name="findRequest">The request for find operation.</param>
        /// <returns>The find response which is returned from the server.</returns>
        public FindResponse Find(FindRequest findRequest)
        {
            FindResponse response = this.activeSyncClient.Find(findRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");
            this.VerifyTransport();
            this.VerifyFindCommand(response);
            this.VerifyWBXMLCapture();
            return response;
        }

        /// <summary>
        /// Search data on the server with an invalid Search request which contains an E-mail Class element.
        /// </summary>
        /// <param name="searchRequest">The request for search operation.</param>
        /// <param name="emailClassElement">The email class element.</param>
        /// <returns>The search response which is returned from server.</returns>
        public SendStringResponse InvalidSearch(SearchRequest searchRequest, string emailClassElement)
        {
            string searchXmlRequest = searchRequest.GetRequestDataSerializedXML();

            // Insert email class element to search command request
            string changedSearchXmlRequest = searchXmlRequest.Insert(searchXmlRequest.LastIndexOf("</And>", StringComparison.CurrentCulture), emailClassElement);
            SendStringResponse result = this.activeSyncClient.SendStringRequest(CommandName.Search, null, changedSearchXmlRequest);
            this.VerifyTransport();
            return result;
        }

        /// <summary>
        /// Fetch all information about exchange object.
        /// </summary>
        /// <param name="itemOperationsRequest">The request for itemOperations.</param>
        /// <returns>The ItemOperations result which is returned from server.</returns>
        public ItemOperationsStore ItemOperations(ItemOperationsRequest itemOperationsRequest)
        {
            ItemOperationsResponse response = this.activeSyncClient.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");
            ItemOperationsStore result = Common.LoadItemOperationsResponse(response);
            this.VerifyTransport();
            this.VerifyItemOperations(result);
            this.VerifyWBXMLCapture();
            return result;
        }

        /// <summary>
        /// Reply to messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="smartReplyRequest">The request for SmartReply operation.</param>
        /// <returns>The SmartReply response which is returned from the server.</returns>
        public SmartReplyResponse SmartReply(SmartReplyRequest smartReplyRequest)
        {
            SmartReplyResponse response = this.activeSyncClient.SmartReply(smartReplyRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");
            this.VerifyTransport();
            return response;
        }

        /// <summary>
        /// Forwards messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="smartForwardRequest">The request for SmartForward operation.</param>
        /// <returns>The SmartForward response which is returned from the server.</returns>
        public SmartForwardResponse SmartForward(SmartForwardRequest smartForwardRequest)
        {
            SmartForwardResponse response = this.activeSyncClient.SmartForward(smartForwardRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");
            this.VerifyTransport();
            return response;
        }

        /// <summary>
        /// Synchronizes the collection hierarchy. 
        /// </summary>
        /// <param name="folderSyncRequest">A FolderSyncRequest object that contains the request information.</param>
        /// <returns>The FolderSync response which is returned from the server.</returns>
        public FolderSyncResponse FolderSync(FolderSyncRequest folderSyncRequest)
        {
            FolderSyncResponse response = this.activeSyncClient.FolderSync(folderSyncRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");
            this.VerifyTransport();
            return response;
        }

        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-ASEMAIL";

            // Merge the configuration
            Common.MergeConfiguration(testSite);

            this.activeSyncClient = new ActiveSyncClient(testSite)
            {
                UserName = Common.GetConfigurationPropertyValue("User1Name", testSite),
                Password = Common.GetConfigurationPropertyValue("User1Password", testSite)
            };
        }
        #endregion

        /// <summary>
        /// Change user to call active sync operation.
        /// </summary>
        /// <param name="userName">The name of a user.</param>
        /// <param name="userPassword">The password of a user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        public void SwitchUser(string userName, string userPassword, string userDomain)
        {
            this.activeSyncClient.UserName = userName;
            this.activeSyncClient.Password = userPassword;
            this.activeSyncClient.Domain = userDomain;
        }
    }
}