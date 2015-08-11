namespace Microsoft.Protocols.TestSuites.MS_VERSS
{
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The implement of IMS-LISTSWSSUTControlAdapter interface.
    /// If the interface methods involve list or file, the file is added in the list, 
    /// the list is generated in the site collection and the site collection is configured as 
    /// the SiteCollectionName property in the MS-VERSS_TestSuite.deployment.ptfconfig file.
    /// </summary>
    public class MS_LISTSWSSUTControlAdapter : ManagedAdapterBase, IMS_LISTSWSSUTControlAdapter
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
            AdapterHelper.Initialize(testSite);

            this.listsService = new ListsSoap();

            string transportType = Common.GetConfigurationPropertyValue("TransportType", testSite);
            this.listsService.Url = Common.GetConfigurationPropertyValue("MSLISTSWSServiceUrl", testSite);

            string domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            string userName = Common.GetConfigurationPropertyValue("UserName", testSite);
            string userPassword = Common.GetConfigurationPropertyValue("Password", testSite);
            this.listsService.Credentials = new NetworkCredential(userName, userPassword, domain);

            SoapVersion soapVersion = Common.GetConfigurationPropertyValue<SoapVersion>("SoapVersion", this.Site);
            switch (soapVersion)
            {
                case SoapVersion.SOAP11:
                    this.listsService.SoapVersion = SoapProtocolVersion.Soap11;
                    break;
                default:
                    this.listsService.SoapVersion = SoapProtocolVersion.Soap12;
                    break;
            }

            if (string.Compare(transportType, "https", true, System.Globalization.CultureInfo.CurrentCulture) == 0)
            {
                Common.AcceptServerCertificate();
            }
        }

        #region Implement the ILISTSWSSUTControlAdapter.
        /// <summary>
        /// Creates a list in site collection.
        /// </summary>
        /// <param name="listName">The name of list that will be created in site collection.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully, 
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        public bool AddList(string listName)
        {
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
                // A 32-bit integer that specifies the list template to use. The 101 specifies Document Library.
                AddListResponseAddListResult result = this.listsService.AddList(listName, string.Empty, 101);

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
        /// Delete the specified list in site collection.
        /// </summary>
        /// <param name="listName">The name of list which will be deleted.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully,
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        public bool DeleteList(string listName)
        {
            try
            {
                this.listsService.DeleteList(listName);
                return true;
            }
            catch (System.Xml.Schema.XmlSchemaValidationException)
            {
                return false;
            }
        }

        /// <summary>
        /// Check in file to a document library.
        /// </summary>
        /// <param name="pageUrl">The URL of the file to be checked in.</param>
        /// <param name="comments">A string containing check-in comments.</param>
        /// <param name="checkInType">A string representation of the values:
        /// 0 means check in minor version, 1 means check in major version or 2 means check in as overwrite.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully,
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        public bool CheckInFile(Uri pageUrl, string comments, string checkInType)
        {
            return this.listsService.CheckInFile(pageUrl.AbsoluteUri, comments, checkInType);
        }

        /// <summary>
        /// Check out a file in a document library.
        /// </summary>
        /// <param name="pageUrl">The URL of the file to be checked out.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully,
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        public bool CheckoutFile(Uri pageUrl)
        {
            return this.listsService.CheckOutFile(pageUrl.AbsoluteUri, bool.FalseString, null);
        }

        /// <summary>
        /// Get the id of a specified list.
        /// </summary>
        /// <param name="listName">The specified list name.</param>
        /// <returns>The string value indicates the id of the specified list.</returns>
        public string GetListID(string listName)
        {
            GetListResponseGetListResult listXml = this.listsService.GetList(listName);

            if (listXml.List.ID != null)
            {
                return listXml.List.ID;
            }
            else
            {
                return string.Empty;
            }
        }

        #endregion
    }
}