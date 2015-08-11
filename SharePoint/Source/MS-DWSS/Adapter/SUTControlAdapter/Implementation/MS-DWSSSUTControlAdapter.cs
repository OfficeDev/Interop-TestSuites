namespace Microsoft.Protocols.TestSuites.MS_DWSS
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using System.Net;
    using System.Web.Services.Protocols;

    /// <summary>
    /// The implement of IMS_DWSSSUTControlAdapter interface.
    /// </summary>
    public class MS_DWSSSUTControlAdapter : ManagedAdapterBase, IMS_DWSSSUTControlAdapter
    {
        /// <summary>
        /// The instance of MS-LISTSWS proxy class.
        /// </summary>
        private ListsSoap listsService;

        /// <summary>
        /// Initialize SUT control adapter.
        /// </summary>
        /// <param name="testSite">The test site instance associated with the current adapter.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            this.listsService = new ListsSoap();

            string domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            string userName = Common.GetConfigurationPropertyValue("UserName", testSite);
            string userPassword = Common.GetConfigurationPropertyValue("Password", testSite);
            this.listsService.Credentials = new NetworkCredential(userName, userPassword, domain);

            SoapVersion soapVersion = Common.GetConfigurationPropertyValue<SoapVersion>("SoapVersion", this.Site);
            switch (soapVersion)
            {
                case SoapVersion.SOAP11:
                    {
                        this.listsService.SoapVersion = SoapProtocolVersion.Soap11;
                        break;
                    }

                default:
                    {
                        this.listsService.SoapVersion = SoapProtocolVersion.Soap12;
                        break;
                    }
            }

            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            if (transport == TransportProtocol.HTTPS)
            {
                Common.AcceptServerCertificate();
            }
        }

        #region Implement the IMS_DWSSSUTControlAdapter.
        
        /// <summary>
        /// Creates a list in site collection.
        /// </summary>
        /// <param name="listName">The name of list that will be created in site collection.</param>
        /// <param name="templateID">A 32-bit integer that specifies the list template to use.</param>
        /// <param name="baseUrl">The site URL for connecting with the specified Document Workspace site.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully, 
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        public bool AddList(string listName, int templateID, string baseUrl)
        {
            this.RedirectBaseUrl(baseUrl);

            GetListCollectionResponseGetListCollectionResult getListResult = this.listsService.GetListCollection();

            // Check whether the specified list already exists in site collection.
            bool listIsExit = false;

            foreach (ListDefinitionCT list in getListResult.Lists)
            {
                string title = list.Title;
                if (title == listName)
                {
                    listIsExit = true;
                    break;
                }
            }

            // If the specified list does not exist in site collection, then create a new list.
            if (listIsExit == false)
            {
                // A 32-bit integer that specifies the list template to use.
                AddListResponseAddListResult result = this.listsService.AddList(listName, string.Empty, templateID);

                if (result.List != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Delete the specified list in the base site.
        /// </summary>
        /// <param name="listName">The name of list which will be deleted.</param>
        /// <param name="baseUrl">The site URL for connecting with the specified Document Workspace site.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully,
        /// TRUE means the list is deleted, FALSE means the list does not exist.</returns>
        public bool DeleteList(string listName, string baseUrl)
        {
            this.RedirectBaseUrl(baseUrl);

            GetListCollectionResponseGetListCollectionResult getListResult = this.listsService.GetListCollection();

            // Check whether the specified list already exists in site collection.
            bool listIsExit = false;

            foreach (ListDefinitionCT list in getListResult.Lists)
            {
                string title = list.Title;
                if (title == listName)
                {
                    listIsExit = true;
                    break;
                }
            }

            if (listIsExit)
            {
                // Delete the list if exist.
                this.listsService.DeleteList(listName);
                return true;
            }

            return false;
        }

        #endregion

        /// <summary>
        /// Redirect the base URL of the List Web service.
        /// </summary>
        /// <param name="baseUrl">The site URL for connecting with the specified Document Workspace Site.</param>
        private void RedirectBaseUrl(string baseUrl)
        {
            if (string.IsNullOrEmpty(baseUrl))
            {
                throw new System.ArgumentException("The base URL should not be null or empty!");
            }

            this.listsService.Url = string.Format("{0}{1}", baseUrl.TrimEnd('/'), Common.GetConfigurationPropertyValue("LISTSSuffix", this.Site));
        }
    }
}