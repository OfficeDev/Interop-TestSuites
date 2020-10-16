namespace Microsoft.Protocols.TestSuites.MS_ASRM
{
    using System.Xml.XPath;
    using Common.DataStructures;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// Adapter class of MS-ASRM.
    /// </summary>
    public partial class MS_ASRMAdapter : ManagedAdapterBase, IMS_ASRMAdapter
    {
        #region Variables
        /// <summary>
        /// The instance of ActiveSync client.
        /// </summary>
        private ActiveSyncClient activeSyncClient;

        #endregion

        #region IMS_ASRMAdapter Properties
        /// <summary>
        /// Gets the XML request sent to protocol SUT.
        /// </summary>
        public IXPathNavigable LastRawRequestXml
        {
            get { return this.activeSyncClient.LastRawRequestXml; }
        }

        /// <summary>
        /// Gets the XML response received from protocol SUT.
        /// </summary>
        public IXPathNavigable LastRawResponseXml
        {
            get { return this.activeSyncClient.LastRawResponseXml; }
        }
        #endregion

        #region Initialize TestSuite
        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(TestTools.ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-ASRM";

            // Merge the common configuration
            Common.MergeConfiguration(testSite);
            this.activeSyncClient = new ActiveSyncClient(testSite)
            {
                AcceptLanguage = "en-us",
                UserName = Common.GetConfigurationPropertyValue("User1Name", testSite),
                Password = Common.GetConfigurationPropertyValue("User1Password", testSite)
            };
        }
        #endregion

        #region MS-ASRMAdapter Members
        /// <summary>
        /// Sync data from the server.
        /// </summary>
        /// <param name="syncRequest">The request for Sync command.</param>
        /// <returns>The sync result which is returned from server.</returns>
        public SyncStore Sync(SyncRequest syncRequest)
        {
            SyncResponse response = this.activeSyncClient.Sync(syncRequest, true);
            Site.Assert.IsNotNull(response, "If the command is successful, the response should not be null.");
            SyncStore result = Common.LoadSyncResponse(response);
            return result;
        }

        /// <summary>
        /// Find an e-mail with specific subject.
        /// </summary>
        /// <param name="request">The request for Sync command.</param>
        /// <param name="subject">The subject of the e-mail to find.</param>
        /// <param name="isRetryNeeded">A boolean value specifies whether need retry.</param>
        /// <returns>The Sync result.</returns>
        public Sync SyncEmail(SyncRequest request, string subject, bool isRetryNeeded)
        {
            Sync sync = this.activeSyncClient.SyncEmail(request, subject, isRetryNeeded);
            this.VerifyTransport();
            this.VerifyWBXMLCapture();
            this.VerifySyncResponse(sync);
            return sync;
        }

        /// <summary>
        /// Fetch all information about exchange object.
        /// </summary>
        /// <param name="itemOperationsRequest">The request for ItemOperations command.</param>
        /// <returns>The ItemOperations result which is returned from server.</returns>
        public ItemOperationsStore ItemOperations(ItemOperationsRequest itemOperationsRequest)
        {
            ItemOperationsResponse response = this.activeSyncClient.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Site.Assert.IsNotNull(response, "If the command is successful, the response should not be null.");
            ItemOperationsStore result = Common.LoadItemOperationsResponse(response);
            this.VerifyTransport();
            this.VerifyWBXMLCapture();
            this.VerifyItemOperationsResponse(result);
            return result;
        }

        /// <summary>
        /// Search items on server.
        /// </summary>
        /// <param name="searchRequest">The request for Search command.</param>
        /// <returns>The Search result which is returned from server.</returns>
        public SearchStore Search(SearchRequest searchRequest)
        {
            SearchResponse response = this.activeSyncClient.Search(searchRequest, true);
            Site.Assert.IsNotNull(response, "If the command is successful, the response should not be null.");
            SearchStore result = Common.LoadSearchResponse(response, Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site));
            this.VerifyTransport();
            this.VerifyWBXMLCapture();
            this.VerifySearchResponse(result);
            return result;
        }

        /// <summary>
        /// Synchronize the collection hierarchy.
        /// </summary>
        /// <param name="request">The request for FolderSync command.</param>
        /// <returns>The FolderSync response which is returned from server.</returns>
        public FolderSyncResponse FolderSync(FolderSyncRequest request)
        {
            FolderSyncResponse response = this.activeSyncClient.FolderSync(request);
            Site.Assert.IsNotNull(response, "If the command is successful, the response should not be null.");
            return response;
        }

        /// <summary>
        /// Gets the RightsManagementInformation by Settings command.
        /// </summary>
        /// <returns>The Settings response which is returned from server.</returns>
        public SettingsResponse Settings()
        {
            SettingsRequest request = new SettingsRequest();
            Request.SettingsRightsManagementInformation settingsInformation = new Request.SettingsRightsManagementInformation();
            Request.SettingsUserInformation setUser = new Request.SettingsUserInformation { Item = string.Empty };
            settingsInformation.Get = string.Empty;
            request.RequestData.RightsManagementInformation = settingsInformation;
            request.RequestData.UserInformation = setUser;
            SettingsResponse response = this.activeSyncClient.Settings(request);
            Site.Assert.IsNotNull(response, "If the command is successful, the response should not be null.");
            this.VerifyWBXMLCapture();
            this.VerifySettingsResponse(response);
            return response;
        }

        /// <summary>
        /// Sends MIME-formatted e-mail messages to the server.
        /// </summary>
        /// <param name="sendMailRequest">The request for SendMail command.</param>
        /// <returns>The SendMail response which is returned from server.</returns>
        public SendMailResponse SendMail(SendMailRequest sendMailRequest)
        {
            SendMailResponse response = this.activeSyncClient.SendMail(sendMailRequest);
            Site.Assert.IsNotNull(response, "If the command is successful, the response should not be null.");

            return response;
        }

        /// <summary>
        /// Reply to messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="smartReplyRequest">The request for SmartReply command.</param>
        /// <returns>The SmartReply response which is returned from server.</returns>
        public SmartReplyResponse SmartReply(SmartReplyRequest smartReplyRequest)
        {
            SmartReplyResponse response = this.activeSyncClient.SmartReply(smartReplyRequest);
            Site.Assert.IsNotNull(response, "If the command is successful, the response should not be null.");
            return response;
        }

        /// <summary>
        /// Forwards messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="smartForwardRequest">The request for SmartForward command.</param>
        /// <returns>The SmartForward response which is returned from server.</returns>
        public SmartForwardResponse SmartForward(SmartForwardRequest smartForwardRequest)
        {
            SmartForwardResponse response = this.activeSyncClient.SmartForward(smartForwardRequest);
            Site.Assert.IsNotNull(response, "If the command is successful, the response should not be null.");
            return response;
        }

        /// <summary>
        /// Change user to call ActiveSync command.
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
        #endregion
    }
}