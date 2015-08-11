namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using System.Collections.Generic;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-OXWSFOLD.
    /// </summary>
    public partial class MS_OXWSFOLDAdapter : ManagedAdapterBase, IMS_OXWSFOLDAdapter
    {
        #region Fields

        /// <summary>
        /// Exchange Web Service instance.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;

        #endregion

        #region IMS_OXWSFOLDAdapter Properties
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT.
        /// </summary>
        public IXPathNavigable LastRawRequestXml
        {
            get { return this.exchangeServiceBinding.LastRawRequestXml; }
        }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT.
        /// </summary>
        public IXPathNavigable LastRawResponseXml
        {
            get { return this.exchangeServiceBinding.LastRawResponseXml; }
        }

        #endregion

        #region Initialize TestSuite
        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Pass ITestSite to adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(TestTools.ITestSite testSite)
        {
            // Initialize.
            base.Initialize(testSite);

            testSite.DefaultProtocolDocShortName = "MS-OXWSFOLD";

            // Merge the common configuration into local configuration.
            Common.MergeConfiguration(testSite);

            // Get the parameters from configuration files.
            string userName = Common.GetConfigurationPropertyValue("User1Name", testSite);
            string password = Common.GetConfigurationPropertyValue("User1Password", testSite);
            string domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            string urlFormat = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            // Initialize service.
            this.exchangeServiceBinding = new ExchangeServiceBinding(urlFormat, userName, password, domain, testSite);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
        }

        #endregion

        #region IMS_OXWSFOLDAdapter Operations
        /// <summary>
        /// Copy one folder into another one.
        /// </summary>
        /// <param name="request">Request of CopyFolder operation.</param>
        /// <returns>Response of CopyFolder operation.</returns>
        public CopyFolderResponseType CopyFolder(CopyFolderType request)
        {
            // Send the request and get the response.
            CopyFolderResponseType response = this.exchangeServiceBinding.CopyFolder(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            if (ResponseClassType.Success == response.ResponseMessages.Items[0].ResponseClass)
            {
                this.VerifyCopyFolderResponse(this.exchangeServiceBinding.IsSchemaValidated);
                this.VerifyAllRelatedRequirements(this.exchangeServiceBinding.IsSchemaValidated, response);
            }

            // Verify transport type related requirement.
            this.VerifyTransportType();

            // Verify soap version.
            this.VerifySoapVersion();

            return response;
        }

        /// <summary>
        /// Create a new folder within a specific folder.
        /// </summary>
        /// <param name="request">Request of CreateFolder operation.</param>
        /// <returns>Response of CreateFolder operation.</returns>
        public CreateFolderResponseType CreateFolder(CreateFolderType request)
        {
            // Send the request and get the response.
            CreateFolderResponseType response = this.exchangeServiceBinding.CreateFolder(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            if (ResponseClassType.Success == response.ResponseMessages.Items[0].ResponseClass)
            {
                this.VerifyCreateFolderResponse(this.exchangeServiceBinding.IsSchemaValidated, response);
                this.VerifyAllRelatedRequirements(this.exchangeServiceBinding.IsSchemaValidated, response);
            }

            // Verify transport type related requirement.
            this.VerifyTransportType();

            // Verify soap version.
            this.VerifySoapVersion();

            return response;
        }

        /// <summary>
        /// Create a managed folder in server, the folder should be added in mailbox by server administrator in advance.
        /// </summary>
        /// <param name="request">Request of CreateManagedFolder operation.</param>
        /// <returns>Response of CreateManagedFolder operation.</returns>
        public CreateManagedFolderResponseType CreateManagedFolder(CreateManagedFolderRequestType request)
        {
            // Send the request and get the response.
            CreateManagedFolderResponseType response = this.exchangeServiceBinding.CreateManagedFolder(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            if (ResponseClassType.Success == response.ResponseMessages.Items[0].ResponseClass)
            {
                this.VerifyCreateManagedFolderResponse(this.exchangeServiceBinding.IsSchemaValidated);
                this.VerifyAllRelatedRequirements(this.exchangeServiceBinding.IsSchemaValidated, response);
            }

            // Verify transport type related requirement.
            this.VerifyTransportType();

            // Verify soap version.
            this.VerifySoapVersion();

            return response;
        }

        /// <summary>
        /// Delete a folder from mailbox.
        /// </summary>
        /// <param name="request">Request DeleteFolder operation.</param>
        /// <returns>Response of DeleteFolder operation.</returns>
        public DeleteFolderResponseType DeleteFolder(DeleteFolderType request)
        {
            // Send the request and get the response.
            DeleteFolderResponseType response = this.exchangeServiceBinding.DeleteFolder(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            if (ResponseClassType.Success == response.ResponseMessages.Items[0].ResponseClass)
            {
                this.VerifyDeleteFolderResponse(this.exchangeServiceBinding.IsSchemaValidated);
                this.VerifyAllRelatedRequirements(this.exchangeServiceBinding.IsSchemaValidated, response);
            }

            // Verify transport type related requirement.
            this.VerifyTransportType();

            // Verify soap version.
            this.VerifySoapVersion();

            return response;
        }

        /// <summary>
        /// Get folders, Calendar folders, Contacts folders, Tasks folders, and search folders.
        /// </summary>
        /// <param name="request">Request of GetFolder operation.</param>
        /// <returns>Response of GetFolder operation.</returns>
        public GetFolderResponseType GetFolder(GetFolderType request)
        {
            // Send the request and get the response.
            GetFolderResponseType response = this.exchangeServiceBinding.GetFolder(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            if (ResponseClassType.Success == response.ResponseMessages.Items[0].ResponseClass)
            {
                this.VerifyGetFolderResponse(response, this.exchangeServiceBinding.IsSchemaValidated);
                this.VerifyAllRelatedRequirements(this.exchangeServiceBinding.IsSchemaValidated, response);
            }

            // Verify transport type related requirement.
            this.VerifyTransportType();

            // Verify soap version.
            this.VerifySoapVersion();

            return response;
        }

        /// <summary>
        /// Empty identified folders and can be used to delete the subfolders of the specified folder.
        /// </summary>
        /// <param name="request">Request of EmptyFolder operation.</param>
        /// <returns>Response of EmptyFolder operation.</returns>
        public EmptyFolderResponseType EmptyFolder(EmptyFolderType request)
        {
            // Send the request and get the response.
            EmptyFolderResponseType response = this.exchangeServiceBinding.EmptyFolder(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            if (ResponseClassType.Success == response.ResponseMessages.Items[0].ResponseClass)
            {
                this.VerifyEmptyFolderResponse(this.exchangeServiceBinding.IsSchemaValidated);
                this.VerifyAllRelatedRequirements(this.exchangeServiceBinding.IsSchemaValidated, response);
            }

            // Verify transport type related requirement.
            this.VerifyTransportType();

            // Verify soap version.
            this.VerifySoapVersion();

            return response;
        }

        /// <summary>
        /// Move folders from a specified parent folder to another parent folder.
        /// </summary>
        /// <param name="request">Request of MoveFolder operation.</param>
        /// <returns>Response of MoveFolder operation.</returns>
        public MoveFolderResponseType MoveFolder(MoveFolderType request)
        {
            // Send the request and get the response.
            MoveFolderResponseType response = this.exchangeServiceBinding.MoveFolder(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            if (ResponseClassType.Success == response.ResponseMessages.Items[0].ResponseClass)
            {
                this.VerifyMoveFolderResponse(this.exchangeServiceBinding.IsSchemaValidated);
                this.VerifyAllRelatedRequirements(this.exchangeServiceBinding.IsSchemaValidated, response);
            }

            // Verify transport type related requirement.
            this.VerifyTransportType();

            // Verify soap version.
            this.VerifySoapVersion();

            return response;
        }

        /// <summary>
        /// Modify properties of an existing folder in the server store.
        /// </summary>
        /// <param name="request">Request of UpdateFolder.</param>
        /// <returns>Response of UpdateFolder operation.</returns>
        public UpdateFolderResponseType UpdateFolder(UpdateFolderType request)
        {
            // Send the request and get the response.
            UpdateFolderResponseType response = this.exchangeServiceBinding.UpdateFolder(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            if (ResponseClassType.Success == response.ResponseMessages.Items[0].ResponseClass)
            {
                this.VerifyUpdateFolderResponse(this.exchangeServiceBinding.IsSchemaValidated);
                this.VerifyAllRelatedRequirements(this.exchangeServiceBinding.IsSchemaValidated, response);
            }

            // Verify transport type related requirement.
            this.VerifyTransportType();

            // Verify soap version.
            this.VerifySoapVersion();

            return response;
        }

        /// <summary>
        /// Switch the current user to the new one, with the identity of the new user to communicate with Exchange Web Service.
        /// </summary>
        /// <param name="userName">The name of a user</param>
        /// <param name="userPassword">The password of a user</param>
        /// <param name="userDomain">The domain, in which a user is</param>
        public void SwitchUser(string userName, string userPassword, string userDomain)
        {
            this.Initialize(this.Site);
            this.exchangeServiceBinding.Credentials = new System.Net.NetworkCredential(userName, userPassword, userDomain);
        }

        /// <summary>
        /// Configure the SOAP header before calling operations.
        /// </summary>
        /// <param name="headerValues">Specify the values for SOAP header.</param>
        public void ConfigureSOAPHeader(Dictionary<string, object> headerValues)
        {
            Common.ConfigureSOAPHeader(headerValues, this.exchangeServiceBinding);
        }
        #endregion
    }
}