namespace Microsoft.Protocols.TestSuites.MS_ASPROV
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-ASPROV.
    /// </summary>
    public partial class MS_ASPROVAdapter : ManagedAdapterBase, IMS_ASPROVAdapter
    {
        /// <summary>
        /// The instance of ActiveSync client.
        /// </summary>
        private ActiveSyncClient activeSyncClient;

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

        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-ASPROV";

            // Merge the common configuration
            Common.MergeConfiguration(testSite);

            this.activeSyncClient = new ActiveSyncClient(testSite)
            {
                UserName = Common.GetConfigurationPropertyValue("User1Name", testSite),
                Password = Common.GetConfigurationPropertyValue("User1Password", testSite)
            };
        }

        /// <summary>
        /// Change the user authentication.
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

        /// <summary>
        /// Apply the specified PolicyKey.
        /// </summary>
        /// <param name="appliedPolicyKey">The policy key to apply.</param>
        public void ApplyPolicyKey(string appliedPolicyKey)
        {
            this.activeSyncClient.PolicyKey = appliedPolicyKey;
        }

        /// <summary>
        /// Apply the specified DeviceType.
        /// </summary>
        /// <param name="appliedDeviceType">The device type to apply.</param>
        public void ApplyDeviceType(string appliedDeviceType)
        {
            this.activeSyncClient.DeviceType = appliedDeviceType;
        }

        /// <summary>
        /// Request the security policy settings that the administrator sets from the server.
        /// </summary>
        /// <param name="provisionRequest">The request of Provision command.</param>
        /// <returns>The response of Provision command.</returns>
        public ProvisionResponse Provision(ProvisionRequest provisionRequest)
        {
            ProvisionResponse provisionResponse = this.activeSyncClient.Provision(provisionRequest);
            Site.Assert.IsNotNull(provisionResponse, "The Provision response returned from server should not be null.");

            // Verify adapter requirements about Provision.
            this.VerifyProvisionCommandRequirements(provisionResponse);
            this.VerifyWBXMLRequirements();
            return provisionResponse;
        }

        /// <summary>
        /// Synchronizes the changes in a collection between the client and the server by sending SyncRequest object.
        /// </summary>
        /// <param name="syncRequest">A SyncRequest object that contains the request information.</param>
        /// <returns>A SyncStore object.</returns>
        public SyncStore Sync(SyncRequest syncRequest)
        {
            SyncResponse response = this.activeSyncClient.Sync(syncRequest, true);
            Site.Assert.IsNotNull(response, "If the Sync command executes successfully, the response from server should not be null.");

            SyncStore syncResult = Common.LoadSyncResponse(response);
            return syncResult;
        }

        /// <summary>
        /// Find an email with specific subject.
        /// </summary>
        /// <param name="syncRequest">A SyncRequest object that contains the request information.</param>
        /// <param name="subject">The subject of the email to find.</param>
        /// <param name="isRetryNeeded">A boolean indicating whether need retry.</param>
        /// <returns>The email with specific subject.</returns>
        public Sync SyncEmail(SyncRequest syncRequest, string subject, bool isRetryNeeded)
        {
            Sync syncResult = this.activeSyncClient.SyncEmail(syncRequest, subject, isRetryNeeded);
            Site.Assert.IsNotNull(syncResult, "If the Sync command executes successfully, the response from server should not be null.");

            return syncResult;
        }

        /// <summary>
        /// Synchronizes the collection hierarchy from server.
        /// </summary>
        /// <param name="folderSyncRequest">The request of FolderSync command.</param>
        /// <returns>The response of FolderSync command.</returns>
        public FolderSyncResponse FolderSync(FolderSyncRequest folderSyncRequest)
        {
            FolderSyncResponse folderSyncResponse = this.activeSyncClient.FolderSync(folderSyncRequest);
            Site.Assert.IsNotNull(folderSyncResponse, "The FolderSync response returned from server should not be null.");
            return folderSyncResponse;
        }

        /// <summary>
        /// Send string request of Provision command to the server and get the Provision response.
        /// </summary>
        /// <param name="provisionRequest">The string request of Provision command.</param>
        /// <returns>The response of Provision command.</returns>
        public ProvisionResponse SendProvisionStringRequest(string provisionRequest)
        {
            SendStringResponse provisionStringResponse = this.activeSyncClient.SendStringRequest(CommandName.Provision, null, provisionRequest);
            Site.Assert.IsNotNull(provisionStringResponse, "The SendStringRequest response returned from server should not be null.");

            // Convert the SendStringResponse to ProvisionResponse.
            ProvisionResponse provisionResponse = new ProvisionResponse
            {
                ResponseDataXML = provisionStringResponse.ResponseDataXML
            };
            provisionResponse.DeserializeResponseData();

            return provisionResponse;
        }
    }
}