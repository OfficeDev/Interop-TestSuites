//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter's implementation
    /// </summary>
    public partial class MS_WEBSSAdapter : ManagedAdapterBase, IMS_WEBSSAdapter
    {
        /// <summary>
        /// SOAP service.
        /// </summary>
        private WebsSoap service;

        #region Initialize TestSuite

        /// <summary>
        /// Initializes the services of MS-WEBSS with the specified transport type, soap version and user authentication provided by the test case.
        /// </summary>
        /// <param name="transportType">The type of the connection</param>
        /// <param name="soapVersion">The soap version of the protocol message</param>
        /// <param name="userAuthentication">a user authenticated</param>
        /// <param name="serverRelativeUrl"> a Server related URL</param>
        public void InitializeService(TransportProtocol transportType, SoapProtocolVersion soapVersion, UserAuthentication userAuthentication, string serverRelativeUrl)
        {
            this.service = Proxy.CreateProxy<WebsSoap>(this.Site);
            this.service.SoapVersion = soapVersion;

            // select transport protocol
            switch (transportType)
            {
                case TransportProtocol.HTTP:
                    this.service.Url = serverRelativeUrl;
                    break;
                default: 
                    this.service.Url = serverRelativeUrl;

                    // when request URL include HTTPS prefix, avoid closing base connection.
                    // local client will accept all certificate after execute this function. 
                    Common.AcceptServerCertificate();

                    break;
            }

            this.service.Credentials = AdapterHelper.ConfigureCredential(userAuthentication);

            // Configure the service timeout.
            int soapTimeOut = Common.GetConfigurationPropertyValue<int>("ServiceTimeOut", this.Site);

            // 60000 means the configure SOAP Timeout is in minute.
            this.service.Timeout = soapTimeOut * 60000;
        }

        /// <summary>
        /// Initialize the services of WEBSS with the specified transport type, soap version and user authentication provided by the test case.
        /// </summary>
        /// <param name="userAuthentication">a user authenticated</param>
        public void InitializeService(UserAuthentication userAuthentication)
        {
            SoapVersion soapVersion = Common.GetConfigurationPropertyValue<SoapVersion>("SoapVersion", this.Site);

            switch (soapVersion)
            {
                case SoapVersion.SOAP11:
                    {
                        this.InitializeService(AdapterHelper.GetTransportType(), SoapProtocolVersion.Soap11, userAuthentication, Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site));
                        break;
                    }

                default:
                    {
                        this.InitializeService(AdapterHelper.GetTransportType(), SoapProtocolVersion.Soap12, userAuthentication, Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site));
                        break;
                    }
            }
        }

        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-WEBSS";
            AdapterHelper.Initialize(testSite);

            // Merge the common configuration
            string conmmonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);
            Common.MergeGlobalConfig(conmmonConfigFileName, this.Site);

            Common.CheckCommonProperties(this.Site, true);

            // Merge SHOULD/MAY configuration
            Common.MergeSHOULDMAYConfig(this.Site);
        }
        #endregion

        /// <summary>
        /// This operation obtains properties of the object referenced by the specified URL.
        /// </summary>
        /// <param name="objectUrl">The URL of the object to retrieve information.</param>
        /// <returns>The Object Id of GetObjectIdFromUrl.</returns>
        public GetObjectIdFromUrlResponseGetObjectIdFromUrlResult GetObjectIdFromUrl(string objectUrl)
        {
            GetObjectIdFromUrlResponseGetObjectIdFromUrlResult getObjectIdFromUrlResult = null;

            getObjectIdFromUrlResult = this.service.GetObjectIdFromUrl(objectUrl);
            this.ValidateGetObjectIdFromUrl(getObjectIdFromUrlResult);
            this.CaptureTransportRelatedRequirements();

            return getObjectIdFromUrlResult;
        }

        /// <summary>
        /// This operation is used to create a new content type on the context site
        /// </summary>
        /// <param name="displayName">displayName means the XML encoded name of content type to be created.</param>
        /// <param name="parentType">parentType is used to indicate the ID of a content type from which the content type to be created will inherit.</param>
        /// <param name="newFields">newFields is the container for a list of existing fields to be included in the content type.</param>
        /// <param name="contentTypeProperties">contentTypeProperties is the container for properties to set on the content type.</param>
        /// <returns>The ID of Created ContentType.</returns>
        public string CreateContentType(string displayName, string parentType, AddOrUpdateFieldsDefinition newFields, CreateContentTypeContentTypeProperties contentTypeProperties)
        {
            string contentTypeId = this.service.CreateContentType(displayName, parentType, newFields, contentTypeProperties);

            this.ValidateCreateContentType();
            this.CaptureTransportRelatedRequirements();

            return contentTypeId;
        }

        /// <summary>
        /// This operation is used to enable customization of the specified cascading style sheet (CSS) for the context site.
        /// </summary>
        /// <param name="cssFile">cssFile specifies the name of one of the CSS files that resides in the default, central location on the server.</param>
        public void CustomizeCss(string cssFile)
        {
            try
            {
                this.service.CustomizeCss(cssFile);
                this.CaptureTransportRelatedRequirements();
            }
            catch (SoapException exp)
            {
                // Capture requirements of detail complex type.
                this.ValidateDetail(exp.Detail);
                this.CaptureTransportRelatedRequirements();

                throw;
            }
        }

        /// <summary>
        /// This operation is used to remove a given content type from the site.
        /// </summary>
        /// <param name="contentTypeId">contentTypeId is the content type ID of the content type that is to be removed from the site.</param>
        /// <returns>The result of DeleteContentType.</returns>
        public DeleteContentTypeResponseDeleteContentTypeResult DeleteContentType(string contentTypeId)
        {
            DeleteContentTypeResponseDeleteContentTypeResult result = new DeleteContentTypeResponseDeleteContentTypeResult();

            result = this.service.DeleteContentType(contentTypeId);

            this.ValidateDeleteContentType();
            this.CaptureTransportRelatedRequirements();

            return result;
        }

        /// <summary>
        /// This operation is used to get a list of activated features on the site and on the parent site collection.
        /// </summary>
        /// <returns>The result of GetActivedFeatures.</returns>
        public string GetActivatedFeatures()
        {
            string result = string.Empty;

            try
            {
                result = this.service.GetActivatedFeatures();

                // Capture requirements of GetActivatedFeatures operation.
                this.ValidateGetActivatedFeatures();
                this.CaptureTransportRelatedRequirements();

                return result;
            }
            catch (SoapException exp)
            {
                // Capture requirements of detail complex type.
                this.ValidateDetail(exp.Detail);
                this.CaptureTransportRelatedRequirements();

                throw;
            }
        }

        /// <summary>
        /// This operation is used to get a list of the titles and URLs of all sites in the site collection.
        /// </summary>
        /// <returns>The result of GetAllSubWebCollection.</returns>
        public GetAllSubWebCollectionResponseGetAllSubWebCollectionResult GetAllSubWebCollection()
        {
            GetAllSubWebCollectionResponseGetAllSubWebCollectionResult result = new GetAllSubWebCollectionResponseGetAllSubWebCollectionResult();

            try
            {
                result = this.service.GetAllSubWebCollection();

                // Capture requirements of WebDefinition complex type.
                if (result.Webs.Length > 0)
                {
                    this.ValidateWebDefinitionForSubWebCollection();
                }

                // Capture requirements of GetAllSubWebCollection operation.
                this.ValidateGetAllSubWebCollection(result);
                this.CaptureTransportRelatedRequirements();

                return result;
            }
            catch (SoapException exp)
            {
                // Capture requirements of detail complex type.
                this.ValidateDetail(exp.Detail);
                this.CaptureTransportRelatedRequirements();

                throw;
            }
        }

        /// <summary>
        /// This operation is used to get the collection of column definitions for all the columns available on the context site. 
        /// </summary>
        /// <returns>The result of GetColumns.</returns>
        public GetColumnsResponseGetColumnsResult GetColumns()
        {
            GetColumnsResponseGetColumnsResult getColumnsResponseGetColumnsResult = null;
            getColumnsResponseGetColumnsResult = this.service.GetColumns();

            this.CaptureTransportRelatedRequirements();

            return getColumnsResponseGetColumnsResult;
        }

        /// <summary>
        /// This operation is used to get the collection of column definitions for all the columns available on the context site. 
        /// </summary>
        /// <param name="contentTypeId">contentTypeId is the ID of the content type to be returned.</param>
        /// <returns>The result of GetContentType.</returns>
        public GetContentTypeResponseGetContentTypeResult GetContentType(string contentTypeId)
        {
            GetContentTypeResponseGetContentTypeResult result = new GetContentTypeResponseGetContentTypeResult();
            result = this.service.GetContentType(contentTypeId);

            this.CaptureTransportRelatedRequirements();

            return result;
        }

        /// <summary>
        /// This method retrieves all content types for a given context site.
        /// </summary>
        /// <returns>The result of GetContentTypes.</returns>
        public GetContentTypesResponseGetContentTypesResult GetContentTypes()
        {
            GetContentTypesResponseGetContentTypesResult result = new GetContentTypesResponseGetContentTypesResult();

            result = this.service.GetContentTypes();

            this.ValidateGetContentTypes();
            this.CaptureTransportRelatedRequirements();

            return result;
        }

        /// <summary>
        /// This operation is used to get the customization status (also known as the ghosted status) of the specified page or file.
        /// </summary>
        /// <param name="fileUrl">fileUrl is the URL of a page or file on the protocol server.</param>
        /// <returns>The result of GetCustomizedPageStatus.</returns>
        public CustomizedPageStatus GetCustomizedPageStatus(string fileUrl)
        {
            CustomizedPageStatus customizedPageStatus = CustomizedPageStatus.None;

            customizedPageStatus = this.service.GetCustomizedPageStatus(fileUrl);
            this.ValidateGetCustomizedPageStatus(customizedPageStatus);
            this.CaptureTransportRelatedRequirements();

            return customizedPageStatus;
        }

        /// <summary>
        /// This operation is used to get the collection of list template definitions for the context site.
        /// </summary>
        /// <returns>The result of GetListTemplates.</returns>
        public GetListTemplatesResponseGetListTemplatesResult GetListTemplates()
        {
            GetListTemplatesResponseGetListTemplatesResult result = null;
            result = this.service.GetListTemplates();

            this.ValidateGetListTemplates();
            this.CaptureTransportRelatedRequirements();

            return result;
        }

        /// <summary>
        /// This operation is used to get the Title, URL, Description, Language, and theme properties of the specified site.
        /// </summary>
        /// <param name="webUrl">WebUrl is a string that contains the absolute URL of the site.</param>
        /// <returns>The result of GetWeb.</returns>
        public GetWebResponseGetWebResult GetWeb(string webUrl)
        {
            GetWebResponseGetWebResult getWebResult = new GetWebResponseGetWebResult();

            try
            {
                getWebResult = this.service.GetWeb(webUrl);

                // Capture requirements of WebDefinition complex type.
                if (getWebResult.Web != null)
                {
                    this.ValidateWebDefinition(getWebResult.Web);
                }

                // Capture requirements of GetWeb operation.
                this.ValidateGetWeb(getWebResult);
                this.CaptureTransportRelatedRequirements();

                return getWebResult;
            }
            catch (SoapException exp)
            {
                // Capture requirements of detail complex type.
                this.ValidateDetail(exp.Detail);
                this.CaptureTransportRelatedRequirements();

                throw;
            }
        }

        /// <summary>
        /// This operation is used to get the Title and URL properties of all immediate child sites of the context site. 
        /// </summary>
        /// <returns>The result of GetWebCollection.</returns>
        public GetWebCollectionResponseGetWebCollectionResult GetWebCollection()
        {
            GetWebCollectionResponseGetWebCollectionResult result = new GetWebCollectionResponseGetWebCollectionResult();

            try
            {
                result = this.service.GetWebCollection();

                // Capture requirements of WebDefinition complex type.
                if (result.Webs.Length > 0)
                {
                    this.ValidateWebDefinitionForSubWebCollection();
                }

                this.ValidateGetWebCollection();
                this.CaptureTransportRelatedRequirements();
            }
            catch (SoapException exp)
            {
                // Capture requirements of detail complex type.
                this.ValidateDetail(exp.Detail);
                this.CaptureTransportRelatedRequirements();

                throw;
            }

            return result;
        }

        /// <summary>
        /// This operation is used to remove an XML document in the XML document collection of a site content type.
        /// </summary>
        /// <param name="contentTypeId">contentTypeID is the content type ID of the site content type to be modified</param>
        /// <param name="documentUri">documentUri is the namespace URI of the XML document of the site content type to remove</param>
        /// <returns>The result of RemoveContentTypeXmlDocument.</returns>
        public RemoveContentTypeXmlDocumentResponseRemoveContentTypeXmlDocumentResult RemoveContentTypeXmlDocument(string contentTypeId, string documentUri)
        {
            RemoveContentTypeXmlDocumentResponseRemoveContentTypeXmlDocumentResult result = new RemoveContentTypeXmlDocumentResponseRemoveContentTypeXmlDocumentResult();

            result = this.service.RemoveContentTypeXmlDocument(contentTypeId, documentUri);

            this.ValidateRemoveContentTypeXmlDocument();
            this.CaptureTransportRelatedRequirements();

            return result;
        }

        /// <summary>
        /// This operation is used to revert all pages within the context site to their original states. 
        /// </summary>
        public void RevertAllFileContentStreams()
        {
            this.service.RevertAllFileContentStreams();

            this.ValidateRevertAllFileContentStreams();
            this.CaptureTransportRelatedRequirements();
        }

        /// <summary>
        /// This operation is used to revert the customizations of the context site defined by the given CSS file and return those customizations to the default style. 
        /// </summary>
        /// <param name="cssFile">cssFile specifies the name of one of the CSS files that resides in the default.</param>
        public void RevertCss(string cssFile)
        {
            try
            {
                this.service.RevertCss(cssFile);

                // Capture requirements of RevertCss operation.
                this.ValidateRevertCss();
                this.CaptureTransportRelatedRequirements();
            }
            catch (SoapException exp)
            {
                // Capture requirements of detail complex type.
                this.ValidateDetail(exp.Detail);
                this.CaptureTransportRelatedRequirements();

                throw;
            }
        }

        /// <summary>
        /// This operation is used to revert the specified page within the context site to its original state.
        /// </summary>
        /// <param name="fileUrl">fileUrl is a string that contains the URL of the page.</param>
        public void RevertFileContentStream(string fileUrl)
        {
            this.service.RevertFileContentStream(fileUrl);

            this.ValidateRevertFileContentStream();
            this.CaptureTransportRelatedRequirements();
        }

        /// <summary>
        /// This operation is used to perform the following operation on the context site and all child sites within its hierarchy
        /// <ul>
        ///     <li>Adding one or more specified new columns</li>
        ///     <li>Updating one or more specified existing columns</li>
        ///     <li>Deleting one or more specified existing columns</li>
        /// </ul>
        /// </summary>
        /// <param name="newFields">newFields is an XML element that represents the collection of columns to be added to the context site and all child sites within its hierarchy.</param>
        /// <param name="updateFields">updateFields is an XML element that represents the collection of columns to be updated on the context site and all child sites within its hierarchy</param>
        /// <param name="deleteFields">deleteFields is an XML element that represents the collection of columns to be deleted from the context site and all child sites within its hierarchy</param>
        /// <returns>The result of UpdateColumns.</returns>
        public UpdateColumnsResponseUpdateColumnsResult UpdateColumns(UpdateColumnsMethod[] newFields, UpdateColumnsMethod1[] updateFields, UpdateColumnsMethod2[] deleteFields)
        {
            UpdateColumnsResponseUpdateColumnsResult updateColumnsResponseUpdateColumnsResult = null;

            updateColumnsResponseUpdateColumnsResult = this.service.UpdateColumns(newFields, updateFields, deleteFields);

            this.ValidateUpdateColumns();
            this.CaptureTransportRelatedRequirements();

            return updateColumnsResponseUpdateColumnsResult;
        }

        /// <summary>
        /// This operation is used to update a content type on the context site.
        /// </summary>
        /// <param name="contentTypeId">contentTypeID is the ID of the content type to be updated.</param>
        /// <param name="contentTypeProperties">properties is the container for properties to set on the content type.</param>
        /// <param name="newFields">newFields is the container for a list of existing fields to be included in the content type.</param>
        /// <param name="updateFields">updateFields is the container for a list of fields to be updated on the content type.</param>
        /// <param name="deleteFields">deleteFields is the container for a list of fields to be updated on the content type.</param>
        /// <returns>The result of UpdateContentType.</returns>
        public UpdateContentTypeResponseUpdateContentTypeResult UpdateContentType(string contentTypeId, UpdateContentTypeContentTypeProperties contentTypeProperties, AddOrUpdateFieldsDefinition newFields, AddOrUpdateFieldsDefinition updateFields, DeleteFieldsDefinition deleteFields)
        {
            UpdateContentTypeResponseUpdateContentTypeResult result = new UpdateContentTypeResponseUpdateContentTypeResult();

            result = this.service.UpdateContentType(contentTypeId, contentTypeProperties, newFields, updateFields, deleteFields);

            this.ValidateUpdateContentType();
            this.CaptureTransportRelatedRequirements();

            return result;
        }

        /// <summary>
        /// This operation is used to add or update an XML document in the XML Document collection of a site content type.
        /// </summary>
        /// <param name="contentTypeId">contentTypeID is the content type ID of the site content type to be modified.</param>
        /// <param name="newDocument">newDocument is the XML document to be added to the site content type XML document collection.</param>
        /// <returns>The result of UpdateContentTypeXmlDocument.</returns>
        public UpdateContentTypeXmlDocumentResponseUpdateContentTypeXmlDocumentResult UpdateContentTypeXmlDocument(string contentTypeId, XmlElement newDocument)
        {
            UpdateContentTypeXmlDocumentResponseUpdateContentTypeXmlDocumentResult result = new UpdateContentTypeXmlDocumentResponseUpdateContentTypeXmlDocumentResult();

            result = this.service.UpdateContentTypeXmlDocument(contentTypeId, newDocument);

            this.ValidateUpdateContentTypeXmlDocument();
            this.CaptureTransportRelatedRequirements();

            return result;
        }

        /// <summary>
        /// This operation is used to get the URL of the parent site of the specified URL.
        /// </summary>
        /// <param name="pageUrl">PageUrl is a URL use to get its parent page URL.</param>
        /// <returns>The result of WebUrlFromPageUrl.</returns>
        public string WebUrlFromPageUrl(string pageUrl)
        {
            string webUrlFromPageUrl = this.service.WebUrlFromPageUrl(pageUrl).ToLower();

            this.ValidateWebUrlFromPageUrl();
            this.CaptureTransportRelatedRequirements();

            return webUrlFromPageUrl;
        }
    }
}