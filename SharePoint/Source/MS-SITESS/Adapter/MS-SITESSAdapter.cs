namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-SITESS.
    /// </summary>
    public partial class MS_SITESSAdapter : ManagedAdapterBase, IMS_SITESSAdapter
    {
        #region Variable
        /// <summary>
        /// Instance of the web service.
        /// </summary>
        private SitesSoap service;

        /// <summary>
        /// Represents the transport protocol of the web service.
        /// </summary>
        private TransportType transportProtocol;
        #endregion

        #region Initialize TestSuite
        /// <summary>
        /// Overrides IAdapter's Initialize(),to set testSite.DefaultProtocolDocShortName.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter,Make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-SITESS";
            string globalConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", testSite);
            AdapterHelper.Initialize(testSite);
            Common.MergeGlobalConfig(globalConfigFileName, testSite);
            
            Common.CheckCommonProperties(this.Site, true);

            Common.MergeSHOULDMAYConfig(testSite);
        }
        #endregion

        #region Implement ISITESSAdapter
        /// <summary>
        /// This operation is used to export the content related to a site to the solution gallery.
        /// </summary>
        /// <param name="solutionFileName">The name of the solution file that will be created. </param>
        /// <param name="title">The name of the solution.</param>
        /// <param name="description">Detailed information that describes the solution.</param>
        /// <param name="fullReuseExportMode">Specify the scope of data that needs to be exported. </param>
        /// <param name="includeWebContent">Specify whether the solution needs to include the contents of all lists and document libraries in the site. </param>
        /// <returns>The result of ExportSolution operation which contains the site-collection relative URL of the created solution file.</returns>
        public string ExportSolution(string solutionFileName, string title, string description, bool fullReuseExportMode, bool includeWebContent)
        {
            // Check whether ExportSolution operation succeeds.
            string exportSolutionResult = null;
            exportSolutionResult = this.service.ExportSolution(solutionFileName, title, description, fullReuseExportMode, includeWebContent);

            // Verify the ExportSolutionResponse structure.
            this.VerifyExportSolution(exportSolutionResult);
            return exportSolutionResult;
        }

        /// <summary>
        /// This operation is used to export the content related to a site into one or more content migration package files.
        /// </summary>
        /// <param name="jobName">Specifies the operation. </param>
        /// <param name="webUrl">The URL of the site to export.</param>
        /// <param name="dataPath">The full path of the location on the server where the content migration package file(s) are saved. </param>
        /// <param name="includeSubwebs">Specifies whether to include the subsite. </param>
        /// <param name="includeUserSecurity">Specifies whether to include access control list (ACL), security group and membership group information.</param>
        /// <param name="overWrite">Specifies whether to overwrite the content migration package file(s) if they exist. </param>
        /// <param name="cabSize">Indicates the suggested size in megabytes for the content migration package file(s). </param>
        /// <returns>The result of ExportWeb operation that contains an error code which indicates whether the operation is succeed or an error type.</returns>
        public int ExportWeb(string jobName, string webUrl, string dataPath, bool includeSubwebs, bool includeUserSecurity, bool overWrite, int cabSize)
        {
            // Check whether ExportWeb operation succeeds.
            int exportWebResult = 0;
            string exportUrl = webUrl;
            string sutComputerName = Common.GetConfigurationPropertyValue(Constants.SutComputerName, this.Site);

            // The server name in this URL must be in lowercase as specified in the Open Specification section 3.1.4.2.2.1.
            exportUrl = exportUrl.Replace(sutComputerName, sutComputerName.ToLower(System.Globalization.CultureInfo.CurrentCulture));

            // SharePoint services 2007 returns an error indicates the URL is not accessible if the transport portion is not in lowercase.
            if (exportUrl.StartsWith("HTTP:", StringComparison.CurrentCultureIgnoreCase))
            {
                exportUrl = exportUrl.Substring(5);
                exportUrl = "http:" + exportUrl;
            }
            else if (exportUrl.StartsWith("HTTPS:", StringComparison.CurrentCultureIgnoreCase))
            {
                exportUrl = exportUrl.Substring(6);
                exportUrl = "https:" + exportUrl;
            }

            exportWebResult = this.service.ExportWeb(jobName, exportUrl, dataPath, includeSubwebs, includeUserSecurity, overWrite, cabSize);

            // Verify the ExportWebResponse structure.
            this.VerifyExportWeb(exportWebResult);
            return exportWebResult;
        }

        /// <summary>
        /// This operation is used to export a workflow template as a site solution to the specified document library.
        /// </summary>
        /// <param name="solutionFileName">The name of the solution file that will be created.</param>
        /// <param name="title">The name of the solution.</param>
        /// <param name="description">Detailed information that describes the solution.</param>
        /// <param name="workflowTemplateName">The name of the workflow template that is to be exported.</param>
        /// <param name="destinationListUrl">The server-relative URL of the document library in which the solution file needs to be created. </param>
        /// <returns>The result of ExportWorkflowTemplate operation which contains the site-relative URL of the created solution file.</returns>
        public string ExportWorkflowTemplate(string solutionFileName, string title, string description, string workflowTemplateName, string destinationListUrl)
        {
            // Check whether ExportWorkflowTemplate operation succeeds.
            string exportWorkflowTemResult = null;
            exportWorkflowTemResult = this.service.ExportWorkflowTemplate(solutionFileName, title, description, workflowTemplateName, destinationListUrl);

            // Legal Uri should not contain space but '%20' instead. 
            if (exportWorkflowTemResult.Contains(" "))
            {
                exportWorkflowTemResult = exportWorkflowTemResult.Replace(" ", "%20");
            }

            // Verify the ExportWorkflowTemplateResponse structure.
            this.VerifyExportWorkflowTemplate(exportWorkflowTemResult);

            return exportWorkflowTemResult;
        }

        /// <summary>
        /// This operation is used to retrieve information about the site collection.
        /// </summary>
        /// <param name="siteUrl">Specifies the absolute URL (Uniform Resource Locator) of a site collection or of a location within a site collection. </param>
        /// <returns>The result of GetSite operation that represents information about the site collection.</returns>
        public string GetSite(string siteUrl)
        {
            // Check whether GetSite operation succeeds.
            string getSiteResult = null;
            getSiteResult = this.service.GetSite(siteUrl);

            // Verify the GetSiteResponse structure.
            this.VerifyGetSite(getSiteResult);
            return getSiteResult;
        }

        /// <summary>
        /// This operation is used to retrieve information about the collection of available site templates.
        /// </summary>
        /// <param name="lcid">Specifies the language code identifier (LCID).</param>
        /// <param name="templateList">SiteTemplates list.</param>
        /// <returns>The result of GetSiteTemplates operation which contains an error code to indicate whether the operation is succeed or an error type.</returns>
        public uint GetSiteTemplates(uint lcid, out Template[] templateList)
        {
            // Check whether GetSiteTemplates operation succeeds.
            uint getSiteTemplateResult = 1;
            templateList = null;
            getSiteTemplateResult = this.service.GetSiteTemplates(lcid, out templateList);

            // Verify the GetSiteTemplatesResponse structure.
            this.VerifyGetSiteTemplates(templateList);
            return getSiteTemplateResult;
        }

        /// <summary>
        /// This operation is used to request renewal of an expired security validation, also known as a message digest. 
        /// </summary>
        /// <returns>The result of GetUpdatedFormDigest operation which contains a security validation token generated by the server.</returns>
        public string GetUpdatedFormDigest()
        {
            // Check whether GetUpdatedFormDigest operation succeeds.
            string getUpdateFormDigestResult = null;
            getUpdateFormDigestResult = this.service.GetUpdatedFormDigest();

            // Verify the GetUpdatedFormDigestResponse structure.
            this.VerifyGetUpdatedFormDigest(getUpdateFormDigestResult);
            return getUpdateFormDigestResult;
        }

        /// <summary>
        /// This operation is used to request renewal of an expired security validation token, also known as a message digest, and the new security validation token’s expiration time.
        /// </summary>
        /// <param name="url">Specify a page URL with which the returned security validation token information is associated.</param>
        /// <returns>The result of GetUpdatedFormDigestInformation operation that contains a security validation token generated by the protocol server, the security validation token’s expiration time in seconds, and other information.</returns>
        public FormDigestInformation GetUpdatedFormDigestInformation(string url)
        {
            // Check whether GetUpdatedFormDigestInformation operation succeeds.
            FormDigestInformation getUpdatedFormDigestInfoResult = null;
            getUpdatedFormDigestInfoResult = this.service.GetUpdatedFormDigestInformation(url);

            // Verify the GetUpdatedFormDigestInformationResponse structure.
            this.VerifyGetUpdatedFormDigestInformation(getUpdatedFormDigestInfoResult);
            return getUpdatedFormDigestInfoResult;
        }

        /// <summary>
        /// This operation is used to import a site from one or more content migration package files to a specified URL.
        /// </summary>
        /// <param name="jobName">Specifies the operation.</param>
        /// <param name="webUrl">The URL of the resulting Web site. </param>
        /// <param name="dataFiles">The URLs of the content migration package files on the server that the server imports to create the resulting Web site.</param>
        /// <param name="logPath">The URL where the server places files describing the progress or status of the operation. </param>
        /// <param name="includeUserSecurity">Specifies whether or not to include ACL, security group and membership group information in the resulting Web site. </param>
        /// <param name="overWrite">Specifies whether or not to overwrite existing files at the location specified by logPath. </param>
        /// <returns>The result of ImportWeb operation that contains an error code which indicates whether the operation is succeed or an error type.</returns>
        public int ImportWeb(string jobName, string webUrl, string[] dataFiles, string logPath, bool includeUserSecurity, bool overWrite)
        {
            // Check whether ImportWeb operation succeeds.
            int importWebResult = -1;
            importWebResult = this.service.ImportWeb(jobName, webUrl, dataFiles, logPath, includeUserSecurity, overWrite);

            // Verify the ImportWebResponse structure.
            this.VerifyImportWeb(importWebResult);
            return importWebResult;
        }

        /// <summary>
        /// This operation is used to create a new subsite of the current site 
        /// </summary>
        /// <param name="url">The site-relative URL of the subsite to be created.</param>
        /// <param name="title">The display name of the subsite to be created.</param>
        /// <param name="description">Description of the subsite to be created.</param>
        /// <param name="templateName">The name of an available site template to be used for the subsite to be created. </param>
        /// <param name="language">An LCID that specifies the language of the user interface of the subsite to be created. </param>
        /// <param name="languageSpecified">Whether language specified</param>
        /// <param name="locale">An LCID that specifies the display format for numbers, dates, times, and currencies in the subsite to be created.</param>
        /// <param name="localeSpecified">Whether locale specified.</param>
        /// <param name="collationLocale">An LCID that specifies the collation order to use in the subsite to be created. </param>
        /// <param name="collationLocaleSpecified">Whether collationLocale specified.</param>
        /// <param name="uniquePermissions">Specifies whether the subsite to be created uses its own set of permissions or parent site.</param>
        /// <param name="uniquePermissionsSpecified">Whether uniquePermissions specified.</param>
        /// <param name="anonymous">Whether the anonymous authentication is to be allowed for the subsite to be created. </param>
        /// <param name="anonymousSpecified">Whether anonymous specified.</param>
        /// <param name="presence">Whether the online presence information is to be enabled for the subsite to be created.</param>
        /// <param name="presenceSpecified">Whether presence specified.</param>
        /// <returns>The result of CreateWeb operation that contains the fully qualified URL to the subsite which was successfully created.</returns>
        public CreateWebResponseCreateWebResult CreateWeb(string url, string title, string description, string templateName, uint language, bool languageSpecified, uint locale, bool localeSpecified, uint collationLocale, bool collationLocaleSpecified, bool uniquePermissions, bool uniquePermissionsSpecified, bool anonymous, bool anonymousSpecified, bool presence, bool presenceSpecified)
        {
            // Check whether CreateWeb operation succeeds.
            CreateWebResponseCreateWebResult createWebResult = null;
            createWebResult = this.service.CreateWeb(url, title, description, templateName, language, languageSpecified, locale, localeSpecified, collationLocale, collationLocaleSpecified, uniquePermissions, uniquePermissionsSpecified, anonymous, anonymousSpecified, presence, presenceSpecified);

            // Verify the CreateWebResponse structure.
            this.VerifyCreateWeb(createWebResult);
            return createWebResult;
        }

        /// <summary>
        /// Delete an existing subsite of the current site. 
        /// </summary>
        /// <param name="url">The site-relative URL of the subsite to be deleted.</param>
        public void DeleteWeb(string url)
        {
            // Check whether DeleteWeb operation succeeds.
            this.service.DeleteWeb(url);

            // Verify requirements related with DeleteWeb operation.
            this.VerifyDeleteWeb();
        }

        /// <summary>
        /// This operation is used to validate whether the specified URLs are valid script safe URLs for the current site.
        /// </summary>
        /// <param name="urls">An array of string contains all URLs that need to be validated. </param>
        /// <returns>An array of boolean contains results for validating the URLs </returns>
        public bool[] IsScriptSafeUrl(string[] urls)
        {
            // Check whether IsScriptSafe operation succeeds.
            bool[] isScriptSafeUrResult = null;
            isScriptSafeUrResult = this.service.IsScriptSafeUrl(urls);

            // Verify the IsScriptSafeResponse structure.
            this.VerifyIsScriptSafeUrl(isScriptSafeUrResult);
            return isScriptSafeUrResult;
        }
        #endregion

        #region Initialize the web service
        /// <summary>
        /// This operation is used to initialize sites service with authority information.
        /// </summary>
        /// <param name="userAuthentication">This parameter used to assign the authentication information of web service.</param>
        public void InitializeWebService(UserAuthenticationOption userAuthentication)
        {
            this.service = Proxy.CreateProxy<SitesSoap>(this.Site);

            // Configure the soap service timeout.
            string soapServiceTimeOut = Common.GetConfigurationPropertyValue(Constants.SoapServiceTimeOut, this.Site);
            this.service.Timeout = Convert.ToInt32(soapServiceTimeOut) * 1000;

            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            switch (transport)
            {
                case TransportProtocol.HTTP:
                    this.transportProtocol = TransportType.HTTP;
                    break;
                default:
                    this.transportProtocol = TransportType.HTTPS;
                    break;
            }

            SoapVersion soapVersion = Common.GetConfigurationPropertyValue<SoapVersion>("SoapVersion", this.Site);
            switch (soapVersion)
            {
                case SoapVersion.SOAP11:
                    this.service.SoapVersion = SoapProtocolVersion.Soap11;
                    break;
                default:
                    this.service.SoapVersion = SoapProtocolVersion.Soap12;
                    break;
            }

            this.service.Url = Common.GetConfigurationPropertyValue(Constants.ServiceUrl, this.Site);

            if (this.transportProtocol == TransportType.HTTPS)
            {
                Common.AcceptServerCertificate();
            }

            this.SetLoginUser(userAuthentication);
        }
        #endregion

        /// <summary>
        /// Select associated user account to login the server according to different user authentications. 
        /// </summary>
        /// <param name="userAuthentication">Assign the authentication information of web service.</param>
        private void SetLoginUser(UserAuthenticationOption userAuthentication)
        {
            string domain = string.Empty;
            string userName = string.Empty;
            string userPassword = string.Empty;

            switch (userAuthentication)
            {
                case UserAuthenticationOption.Authenticated:
                    domain = Common.GetConfigurationPropertyValue(Constants.Domain, this.Site);
                    userName = Common.GetConfigurationPropertyValue(Constants.UserName, this.Site);
                    userPassword = Common.GetConfigurationPropertyValue(Constants.Password, this.Site);
                    break;
                case UserAuthenticationOption.Unauthenticated:
                    domain = Common.GetConfigurationPropertyValue(Constants.UnauthorizedUserDomain, this.Site);
                    userName = Common.GetConfigurationPropertyValue(Constants.UnauthorizedUserName, this.Site);
                    userPassword = Common.GetConfigurationPropertyValue(Constants.UnauthorizedUserPassword, this.Site);
                    break;
            }

            this.service.Credentials = new NetworkCredential(userName, userPassword, domain);
        }
    }
}