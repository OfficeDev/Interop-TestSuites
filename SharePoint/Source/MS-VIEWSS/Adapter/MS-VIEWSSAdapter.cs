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
    using System.Text.RegularExpressions;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The Adapter class of MS-VIEWSS
    /// </summary>
    public partial class MS_VIEWSSAdapter : ManagedAdapterBase, IMS_VIEWSSAdapter
    {
        #region Variables

        /// <summary>
        /// The proxy class.
        /// </summary>
        private ViewsSoap viewssProxy;

        #endregion Variables

        #region Initialize TestSuite

        /// <summary>
        /// Overrides IAdapter's Initialize method, to set default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">A parameter represents a ITestSite instance which is used to get/operate current test suite context.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            // Set the protocol name of current test suite
            testSite.DefaultProtocolDocShortName = "MS-VIEWSS";

            // Load Common configuration
            this.LoadCommonConfiguration();

            // Check whether the common properties are valid.
            Common.CheckCommonProperties(this.Site, true);

            // Merge the SHOULD MAY config files in the test site.
            Common.MergeSHOULDMAYConfig(this.Site);

            // Initialize the proxy.
            this.viewssProxy = Proxy.CreateProxy<ViewsSoap>(this.Site, false, true, false);

            // Get target service fully qualified URL
            this.viewssProxy.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);
            string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            string password = Common.GetConfigurationPropertyValue("Password", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.viewssProxy.Credentials = new NetworkCredential(userName, password, domain);

            // Set SOAP version according to the configuration file.
            this.SetSoapVersion();

            if (TransportProtocol.HTTPS == Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site))
            {
                Common.AcceptServerCertificate();
            }

            // Configure the service timeout.
            int soapTimeOut = Common.GetConfigurationPropertyValue<int>("ServiceTimeOut", this.Site);

            // 60000 means the SOAP Timeout in minute.
            this.viewssProxy.Timeout = soapTimeOut * 60000;
        }

        #endregion

        #region Implement IMS_VIEWSSAdapter

        /// <summary>
        /// This operation is used to create a list view for the specified list.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        /// <param name="viewFields">Specify the fields included in a list view.</param>
        /// <param name="query">Include the information that affects how a list view displays the data.</param>
        /// <param name="rowLimit">Specify whether a list supports displaying items page-by-page, and the count of items a list view displays per page.</param>
        /// <param name="type">Specify the type of a list view.</param>
        /// <param name="makeViewDefault">Specify whether to make the list view the default list view for the specified list.</param>
        /// <returns>The result returns a list view that the type is BriefViewDefinition, if the operation succeeds.</returns>
        public AddViewResponseAddViewResult AddView(
           string listName,
           string viewName,
           AddViewViewFields viewFields,
           AddViewQuery query,
           AddViewRowLimit rowLimit,
           string type,
           bool makeViewDefault)
        {
            AddViewResponseAddViewResult addViewResult;

            try
            {
                addViewResult = this.viewssProxy.AddView(
                        listName,
                        viewName,
                        viewFields,
                        query,
                        rowLimit,
                        type,
                        makeViewDefault);

                // Used to validate the transport requirements.
                this.CaptureTransportRelatedRequirements();

                // Used to validate AddViewResult schema requirements.
                this.ValidateAddViewResult(addViewResult);
            }
            catch (SoapException soapException)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [AddView] method:\r\n{0}",
                                soapException.Detail.InnerXml);

                // Used to validate the transport requirements.
                this.CaptureTransportRelatedRequirements();

                // Used to validate the schema of SoapFault.
                this.ValidateSOAPFaultDetails(soapException.Detail);

                throw;
            }

            return addViewResult;
        }

        /// <summary>
        /// This operation is used to delete the specified list view of the specified list.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        public void DeleteView(string listName, string viewName)
        {
            try
            {
                this.viewssProxy.DeleteView(listName, viewName);

                // Used to validate the transport requirements
                this.CaptureTransportRelatedRequirements();

                // Used to validate DeleteView result requirements
                this.ValidateDeleteViewResult();
            }
            catch (SoapException soapException)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [DeleteView] method:\r\n{0}",
                                soapException.Detail.InnerXml);

                // Used to validate the transport requirements
                this.CaptureTransportRelatedRequirements();

                // Used to validate the schema of SoapFault.
                this.ValidateSOAPFaultDetails(soapException.Detail);

                throw;
            }
        }

        /// <summary>
        /// This operation is used to obtain details of a specified list view of the specified list.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        /// <returns>The result returns a list view that the type is BriefViewDefinition, if the operation succeeds.</returns>
        public GetViewResponseGetViewResult GetView(string listName, string viewName)
        {
            GetViewResponseGetViewResult getViewResult;

            try
            {
                getViewResult = this.viewssProxy.GetView(listName, viewName);

                // Used to validate the transport requirements
                this.CaptureTransportRelatedRequirements();

                // Used to validate the schema validation related requirements.
                this.ValidateGetViewResult(getViewResult);
            }
            catch (SoapException soapException)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [GetView] method:\r\n{0}",
                                soapException.Detail.InnerXml);

                // Used to validate the transport requirements
                this.CaptureTransportRelatedRequirements();

                // Used to validate the schema of SoapFault.
                this.ValidateSOAPFaultDetails(soapException.Detail);

                throw;
            }

            return getViewResult;
        }

        /// <summary>
        /// This operation is used to retrieve the collection of list views of a specified list.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <returns>The result returns a collection of View elements of the specified list if the operation succeeds.</returns>
        public GetViewCollectionResponseGetViewCollectionResult GetViewCollection(string listName)
        {
            GetViewCollectionResponseGetViewCollectionResult getViewCollectionResult;

            try
            {
                getViewCollectionResult = this.viewssProxy.GetViewCollection(listName);

                // Used to validate the transport requirements.
                this.CaptureTransportRelatedRequirements();

                // Used to validate the GetViewCollection schema requirements
                this.ValidateGetViewCollectionResult(getViewCollectionResult);
            }
            catch (SoapException soapException)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [GetViewCollection] method:\r\n{0}",
                                soapException.Detail.InnerXml);

                // Used to validate the transport requirements.
                this.CaptureTransportRelatedRequirements();

                // Used to validate the schema of SoapFault.
                this.ValidateSOAPFaultDetails(soapException.Detail);

                throw;
            }

            return getViewCollectionResult;
        }

        /// <summary>
        /// This operation is used to obtain details of a specified list view of the specified list, including display properties in CAML and HTML.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        /// <returns>The result returns the details of a specified list view of the specified list if the operation succeeds.</returns>
        public GetViewHtmlResponseGetViewHtmlResult GetViewHtml(string listName, string viewName)
        {
            GetViewHtmlResponseGetViewHtmlResult getViewHtmlResult;
            try
            {
                getViewHtmlResult = this.viewssProxy.GetViewHtml(listName, viewName);

                // Used to validate the transport requirements.
                this.CaptureTransportRelatedRequirements();

                // Used to validate the GetViewHtml schema requirements.
                this.ValidateGetViewHtmlResult(getViewHtmlResult);
            }
            catch (SoapException soapException)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [GetViewHtml] method:\r\n{0}",
                                soapException.Detail.InnerXml);

                // Used to validate the transport requirements.
                this.CaptureTransportRelatedRequirements();

                // Used to validate the schema of SoapFault.
                this.ValidateSOAPFaultDetails(soapException.Detail);
                throw;
            }

            return getViewHtmlResult;
        }

        /// <summary>
        /// This operation is used to update the specified list view, without the display properties.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        /// <param name="viewProperties">Specify the properties of a list view on the server.</param>
        /// <param name="query">Include the information that affects how a list view displays the data.</param>
        /// <param name="viewFields">Specify the fields included in a list view.</param>
        /// <param name="aggregations">The type of the aggregation.</param> 
        /// <param name="formats">Specify the row and column formatting of a list view.</param>
        /// <param name="rowLimit">Specify whether a list supports displaying items page-by-page, and the count of items a list view displays per page.</param>
        /// <returns>The result returns a list view that the type is BriefViewDefinition, if the operation succeeds.</returns>
        public UpdateViewResponseUpdateViewResult UpdateView(
            string listName,
            string viewName,
            UpdateViewViewProperties viewProperties,
            UpdateViewQuery query,
            UpdateViewViewFields viewFields,
            UpdateViewAggregations aggregations,
            UpdateViewFormats formats,
            UpdateViewRowLimit rowLimit)
        {
            UpdateViewResponseUpdateViewResult updateViewResponse;

            try
            {
                updateViewResponse = this.viewssProxy.UpdateView(
                    listName,
                    viewName,
                    viewProperties,
                    query,
                    viewFields,
                    aggregations,
                    formats,
                    rowLimit);

                // Used to validate the transport requirements.
                this.CaptureTransportRelatedRequirements();

                // Used to validate the UpdateView schema requirements.
                this.ValidateUpdateViewResult(updateViewResponse);
            }
            catch (SoapException soapException)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [updateView] method:\r\n{0}",
                                soapException.Detail.InnerXml);

                // Used to validate the transport requirements.
                this.CaptureTransportRelatedRequirements();

                // Used to validate the schema of SoapFault.
                this.ValidateSOAPFaultDetails(soapException.Detail);

                throw;
            }

            return updateViewResponse;
        }

        /// <summary>
        /// This operation is used to update a list view for a specified list, including display properties in CAML and HTML.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        /// <param name="viewProperties">Specify the properties of a list view on the server.</param>
        /// <param name="toolbar">Specify the rendering of the toolbar of a list.</param>
        /// <param name="viewHeader">Specify the rendering of the header, or the top of a list view page.</param>
        /// <param name="viewBody">Specify the rendering of the main, or the middle portion of a list view page.</param>
        /// <param name="viewFooter">Specify the rendering of the footer, or the bottom of a list view page.</param>
        /// <param name="viewEmpty">Specify the message to be displayed when no items are in a list view.</param>
        /// <param name="rowLimitExceeded">Specify rendering of additional items when the number of items exceeds the value.</param>
        /// <param name="query">Include the information that affects how a list view displays the data.</param>
        /// <param name="viewFields">Specify the fields included in a list view.</param>
        /// <param name="aggregations">The type of the aggregation.</param> 
        /// <param name="formats">Specify the row and column formatting of a list view.</param>
        /// <param name="rowLimit">Specify whether a list supports displaying items page-by-page, and the count of items a list view displays per page.</param>
        /// <returns>The result returns a View that the type is ViewDefinition if the operation succeeds</returns>
        public UpdateViewHtmlResponseUpdateViewHtmlResult UpdateViewHtml(
            string listName,
            string viewName,
            UpdateViewHtmlViewProperties viewProperties,
            UpdateViewHtmlToolbar toolbar,
            UpdateViewHtmlViewHeader viewHeader,
            UpdateViewHtmlViewBody viewBody,
            UpdateViewHtmlViewFooter viewFooter,
            UpdateViewHtmlViewEmpty viewEmpty,
            UpdateViewHtmlRowLimitExceeded rowLimitExceeded,
            UpdateViewHtmlQuery query,
            UpdateViewHtmlViewFields viewFields,
            UpdateViewHtmlAggregations aggregations,
            UpdateViewHtmlFormats formats,
            UpdateViewHtmlRowLimit rowLimit)
        {
            UpdateViewHtmlResponseUpdateViewHtmlResult updateViewHtmlResult;

            try 
            {
                updateViewHtmlResult = this.viewssProxy.UpdateViewHtml(
                listName,
                viewName,
                viewProperties,
                toolbar,
                viewHeader,
                viewBody,
                viewFooter,
                viewEmpty,
                rowLimitExceeded,
                query,
                viewFields,
                aggregations,
                formats,
                rowLimit);

                // Used to validate the transport requirements.
                this.CaptureTransportRelatedRequirements();

                // Used to validate the UpdateViewHtml schema requirements.
                this.ValidateUpdateViewHtmlResult(updateViewHtmlResult);
            }
            catch (SoapException soapException)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [UpdateViewHtml] method:\r\n{0}",
                                soapException.Detail.InnerXml);

                // Used to validate the transport requirements.
                this.CaptureTransportRelatedRequirements();

                // Used to validate the schema of SoapFault.
                this.ValidateSOAPFaultDetails(soapException.Detail);

                throw;
            }

            return updateViewHtmlResult;
        }

        /// <summary>
        /// This operation is used to obtain details of a specified list view of the specified list, including display properties in CAML and HTML.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        /// <param name="viewProperties">Specify the properties of a list view on the server.</param>
        /// <param name="toolbar">Specify the rendering of the toolbar of a list.</param>
        /// <param name="viewHeader">Specify the rendering of the header, or the top of a list view page.</param>
        /// <param name="viewBody">Specify the rendering of the main, or the middle portion of a list view page.</param>
        /// <param name="viewFooter">Specify the rendering of the footer, or the bottom of a list view page.</param>
        /// <param name="viewEmpty">Specify the message to be displayed when no items are in a list view.</param>
        /// <param name="rowLimitExceeded">Specify rendering of additional items when the number of items exceeds the value.</param>
        /// <param name="query">Include the information that affects how a list view displays the data.</param>
        /// <param name="viewFields">Specify the fields included in a list view.</param>
        /// <param name="aggregations">The type of the aggregation.</param> 
        /// <param name="formats">Specify the row and column formatting of a list view.</param>
        /// <param name="rowLimit">Specify whether a list supports displaying items page-by-page, and the count of items a list view displays per page.</param>
        /// <param name="openApplicationExtension">Specify what kind of application to use to edit the view.</param>
        /// <returns>The result returns a View that the type is ViewDefinition if the operation succeeds</returns>
        public UpdateViewHtml2ResponseUpdateViewHtml2Result UpdateViewHtml2(
            string listName,
            string viewName,
            UpdateViewHtml2ViewProperties viewProperties,
            UpdateViewHtml2Toolbar toolbar,
            UpdateViewHtml2ViewHeader viewHeader,
            UpdateViewHtml2ViewBody viewBody,
            UpdateViewHtml2ViewFooter viewFooter,
            UpdateViewHtml2ViewEmpty viewEmpty,
            UpdateViewHtml2RowLimitExceeded rowLimitExceeded,
            UpdateViewHtml2Query query,
            UpdateViewHtml2ViewFields viewFields,
            UpdateViewHtml2Aggregations aggregations,
            UpdateViewHtml2Formats formats,
            UpdateViewHtml2RowLimit rowLimit,
            string openApplicationExtension)
        {
            UpdateViewHtml2ResponseUpdateViewHtml2Result updateViewHtml2Result;

            try
            {
                updateViewHtml2Result = this.viewssProxy.UpdateViewHtml2(
                listName,
                viewName,
                viewProperties,
                toolbar,
                viewHeader,
                viewBody,
                viewFooter,
                viewEmpty,
                rowLimitExceeded,
                query,
                viewFields,
                aggregations,
                formats,
                rowLimit,
                openApplicationExtension);

                // Used to validate the transport requirements.
                this.CaptureTransportRelatedRequirements();

                // Used to validate the UpdateViewHtml2 schema requirements.
                this.ValidateUpdateViewHtml2Result(updateViewHtml2Result);
            }
            catch (SoapException soapException)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [UpdateViewHtml2] method:\r\n{0}",
                                soapException.Detail.InnerXml);

                // Used to validate the transport requirements.
                this.CaptureTransportRelatedRequirements();

                // Used to validate the schema of SoapFault.
                this.ValidateSOAPFaultDetails(soapException.Detail);

                throw;
            }

            return updateViewHtml2Result;
        }

        #endregion

        #region Private helper methods
        /// <summary>
        /// A method used to load Common Configuration
        /// </summary>
        private void LoadCommonConfiguration()
        {
            // Get a specified property value from ptfconfig file.
            string conmmonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);

            // Execute the merge the common configuration
            Common.MergeGlobalConfig(conmmonConfigFileName, this.Site);
        }

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
                        this.viewssProxy.SoapVersion = SoapProtocolVersion.Soap11;
                        break;
                    }

                default:
                    {
                        this.viewssProxy.SoapVersion = SoapProtocolVersion.Soap12;
                        break;
                    }
            }
        }

        /// <summary>
        /// Used to validate the schema of error code
        /// </summary>
        /// <param name="errorcode">The error code returned from server.</param>
        /// <returns>If the error code conforms to the regular expression then the returned value will be True, if the error code does not conform to the regex then the returned value will be False.</returns>
        private bool IsErrorCodeHexadecimal(string errorcode)
        {
            // The hexadecimal representation of a 4-byte result code.
            string pattern = @"^0x[0-9A-F]{8}$";
            Regex regex = new Regex(pattern, RegexOptions.Compiled);
            return regex.IsMatch(errorcode);
        }

        /// <summary>
        /// Used to judge whether a string is a server relative URL.
        /// </summary>
        /// <param name="url">A string indicates a URL.</param>
        /// <returns>If the input url is a server relative URL then return True, otherwise return False.</returns>
        private bool IsServerRelativeUrl(string url)
        {
            string siteCollectionName = Common.GetConfigurationPropertyValue("SiteCollectionName", Site);
            string listName = Common.GetConfigurationPropertyValue("DisplayListName", Site);
            string pattern = @"^/sites/" + siteCollectionName + @"/Lists/" + listName + @"/[\w\s]+\.aspx$";
            Regex regex = new Regex(pattern.ToLower(), RegexOptions.Compiled);

            return regex.IsMatch(url.ToLower());
        }
        #endregion
    }
}