//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_VIEWSS
{
    using System.Net;
    using System.Web.Services.Protocols;
    using Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUT Control Adapter class of MS-VIEWSS
    /// </summary>
    public class MS_VIEWSSSUTControlAdapter : ManagedAdapterBase, IMS_VIEWSSSUTControlAdapter
    {
        #region Variables

        /// <summary>
        /// The proxy class for lists service.
        /// </summary>
        private ListsSoap listProxy;

        #endregion Variables

        #region Initialize SUT Control Adapter
        /// <summary>
        /// Initialize the adapter instance.
        /// </summary>
        /// <param name="testSite">A parameter represents a ITestSite instance which is used to get/operate current test suite context.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            // Initialize the proxy.
            this.listProxy = new ListsSoap();

            // Get target service fully qualified URL
            this.listProxy.Url = Common.GetConfigurationPropertyValue("ListsServiceUrl", this.Site);
            string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            string password = Common.GetConfigurationPropertyValue("Password", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.listProxy.Credentials = new NetworkCredential(userName, password, domain);

            // Set soap version according to the configuration file.
            this.SetSoapVersion();

            if (TransportProtocol.HTTPS == Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site))
            {
                Common.AcceptServerCertificate();
            }
        }

        #endregion

        #region Implement IMS_VIEWSSSUTControlAdapter

        /// <summary>
        /// Implement the GetItemsCount method for getting the count of the list items in the specified view.
        /// </summary>
        /// <param name="listGuid">A specified list GUID in the server.</param>
        /// <param name="viewGuid">A specified view GUID in the server.</param>
        /// <returns>The count of the list items in the specified view.</returns>
        public int GetItemsCount(string listGuid, string viewGuid)
        {
            int itemCount = 0;
            try
            {
                GetListItemsResponseGetListItemsResult getItemsResult
                    = this.listProxy.GetListItems(listGuid, viewGuid, null, null, null, null, null);
                itemCount = getItemsResult.listitems.data.Any.Length;
            }
            catch (SoapException soapException)
            {
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    @"There is an exception generated when calling [GetItemsCount] method:\r\n{0}, check parameters are correct or not.",
                    soapException.Message);
                throw;
            }

            return itemCount;
        }

        /// <summary>
        /// A method retrieves the List GUID when a list name is provided.
        /// </summary>
        /// <param name="listDisplayName">A specified list display name on the server.</param>
        /// <returns>The GUID of the list.</returns>
        public string GetListGuidByName(string listDisplayName)
        {
            GetListResponseGetListResult listResult;

            try
            {
                listResult = this.listProxy.GetList(listDisplayName);
            }
            catch (SoapException soapEx)
            {
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    @"There is an exception generated when calling [GetListGuidByName] method:\r\n{0}, check the parameter is correct or not.",
                    soapEx.Message);
                throw;
            }

            this.Site.Assert.IsNotNull(listResult, "The \"listResult\" in the response of \"GetList\" operation should not be null.");
            this.Site.Assert.IsNotNull(listResult.List, "The \"listResult.List\" in the response of \"GetList\" operation should not be null.");
            this.Site.Assert.IsNotNull(listResult.List.ID, "The \"listResult.List.ID\" in the response of \"GetList\" operation should not be null.");

            string listGUID = listResult.List.ID.Trim('{', '}');
            return listGUID;
        }

        /// <summary>
        /// A method retrieves default view.
        /// </summary>
        /// <param name="listDisplayName">A specified list name in the server.</param>
        /// <returns>Name of the default view for the specified list.</returns>
        public string GetListAndView(string listDisplayName)
        {
            string originalDefaultViewName = null;
            GetListAndViewResponseGetListAndViewResult result = null;
            try
            {
                result = this.listProxy.GetListAndView(listDisplayName, string.Empty);
            }
            catch (SoapException soapEx)
            {
                Site.Log.Add(
                    LogEntryKind.Debug,
                    @"There is an exception generated when calling [GetListAndView] method:\r\n{0}, check the parameter is correct or not.",
                    soapEx.Message);

                throw;
            }

            Site.Assert.IsNotNull(result, "If the GetListAndView operation executes successfully, the response should not be null.");
            if (result.ListAndView != null)
            {
                if (result.ListAndView.View != null)
                {
                    Site.Assert.AreEqual<string>(
                        "true", 
                        result.ListAndView.View.DefaultView.ToLower(), 
                        @"The view returned in server response should be default view of the specified list as the 'viewName' was set to empty.");

                    originalDefaultViewName = result.ListAndView.View.Name;
                }
            }

            return originalDefaultViewName;
        }
        #endregion

        #region Private helper methods

        /// <summary>
        /// Set the SOAP version according to the SoapVersion property.
        /// </summary>
        private void SetSoapVersion()
        {
            SoapVersion soapVersion = Common.GetConfigurationPropertyValue<SoapVersion>("SoapVersion", this.Site);

            switch (soapVersion)
            {
                case SoapVersion.SOAP11:
                    {
                        this.listProxy.SoapVersion = SoapProtocolVersion.Soap11;
                        break;
                    }

                default:
                    {
                        this.listProxy.SoapVersion = SoapProtocolVersion.Soap12;
                        break;
                    }
            }
        }

        #endregion
    }
}
