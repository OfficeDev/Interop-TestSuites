namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OXWSUSRCFG SUT control adapter implementation. 
    /// </summary>
    public class MS_OXWSUSRCFGSUTControlAdapter : ManagedAdapterBase, IMS_OXWSUSRCFGSUTControlAdapter
    {
        #region Fields
        /// <summary>
        /// An instance of ExchangeServiceBinding.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;
        #endregion

        #region Initialize TestSuite
        /// <summary>
        /// Initialize some variables overridden.
        /// </summary>
        /// <param name="testSite">The instance of ITestSite Class.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            string userName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            string password = Common.GetConfigurationPropertyValue("User1Password", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            string url = Common.GetConfigurationPropertyValue("ServiceUrl", this.Site);

            this.exchangeServiceBinding = new ExchangeServiceBinding(url, userName, password, domain, this.Site);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, this.Site);
        }
        #endregion

        #region IMS_OXWSUSRCFGSUTControlAdapter Operations
        /// <summary>
        /// Log on to a mailbox with a specified user account and create an user configuration objects on inbox folder.
        /// </summary>
        /// <param name="userName">The name of the user used to communicate with server</param>
        /// <param name="password">The password of the user used to communicate with server</param>
        /// <param name="domain">The domain of the user used to communicate with server</param>
        /// <param name="userConfigurationName">Name of the user configuration object.</param>
        /// <returns>If succeed, return true; otherwise, return false.</returns>
        public bool CreateUserConfiguration(string userName, string password, string domain, string userConfigurationName)
        {
            userConfigurationName = userConfigurationName.Replace("_", string.Empty);
            this.exchangeServiceBinding.Credentials = new System.Net.NetworkCredential(userName, password, domain);
            CreateUserConfigurationType request = new CreateUserConfigurationType();
            request.UserConfiguration = new UserConfigurationType();
            request.UserConfiguration.UserConfigurationName = new UserConfigurationNameType();
            request.UserConfiguration.UserConfigurationName.Name = userConfigurationName;
            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            distinguishedFolderId.Id = DistinguishedFolderIdNameType.inbox;
            request.UserConfiguration.UserConfigurationName.Item = distinguishedFolderId;

            CreateUserConfigurationResponseType response = this.exchangeServiceBinding.CreateUserConfiguration(request);

            return response.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success
                && response.ResponseMessages.Items[0].ResponseCode == ResponseCodeType.NoError ? true : false;
        }
        #endregion
    }
}
