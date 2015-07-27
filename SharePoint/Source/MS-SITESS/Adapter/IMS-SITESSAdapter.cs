//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Provides means of interacting with proxy class, and implemented by MS_SITESSAdapter.
    /// </summary>
    public interface IMS_SITESSAdapter : IAdapter
    {
        /// <summary>
        /// This operation is used to initialize Sites this.Service with authority information.
        /// </summary>
        /// <param name="userAuthentication">This parameter used to assign the authentication information of web service.</param>
        void InitializeWebService(UserAuthenticationOption userAuthentication);

        /// <summary>
        /// This operation is used to export the content related to a site to the solution gallery.
        /// </summary>
        /// <param name="solutionFileName">The name of the solution file that will be created. </param>
        /// <param name="title">The name of the solution.</param>
        /// <param name="description">Detailed information that describes the solution.</param>
        /// <param name="fullReuseExportMode">Specify the scope of data that needs to be exported. </param>
        /// <param name="includeWebContent">Specify whether the solution needs to include the contents of all lists and document libraries in the site. </param>
        /// <returns>The result of ExportSolution operation which contains the site-collection relative URL of the created solution file.</returns>
        string ExportSolution(string solutionFileName, string title, string description, bool fullReuseExportMode, bool includeWebContent);

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
        int ExportWeb(string jobName, string webUrl, string dataPath, bool includeSubwebs, bool includeUserSecurity, bool overWrite, int cabSize);

        /// <summary>
        /// This operation is used to export a workflow template as a site solution to the specified document library.
        /// </summary>
        /// <param name="solutionFileName">The name of the solution file that will be created.</param>
        /// <param name="title">The name of the solution.</param>
        /// <param name="description">Detailed information that describes the solution.</param>
        /// <param name="workflowTemplateName">The name of the workflow template that is to be exported.</param>
        /// <param name="destinationListUrl">The server-relative URL of the document library in which the solution file needs to be created. </param>
        /// <returns>The result of ExportWorkflowTemplate operation which contains the site-relative URL of the created solution file.</returns>
        string ExportWorkflowTemplate(string solutionFileName, string title, string description, string workflowTemplateName, string destinationListUrl);

        /// <summary>
        /// This operation is used to retrieve information about the site collection.
        /// </summary>
        /// <param name="siteUrl">Specifies the absolute URL (Uniform Resource Locator) of a site collection or of a location within a site collection. </param>
        /// <returns>The result of GetSite operation that represents information about the site collection. </returns>
        string GetSite(string siteUrl);

        /// <summary>
        /// This operation is used to retrieve information about the collection of available site templates.
        /// </summary>
        /// <param name="lcid">Specifies the language code identifier (LCID).</param>
        /// <param name="templateList">SiteTemplates list.</param>
        /// <returns>The result of GetSiteTemplates operation which contains an error code to indicate whether the operation is succeed or an error type.</returns>
        uint GetSiteTemplates(uint lcid, out Template[] templateList);

        /// <summary>
        /// This operation is used to request renewal of an expired security validation, also known as a message digest. 
        /// </summary>
        /// <returns>The result of GetUpdatedFormDigest operation which contains a security validation token generated by the server.</returns>
        string GetUpdatedFormDigest();

        /// <summary>
        /// This operation is used to request renewal of an expired security validation token, also known as a message digest, and the new security validation token’s expiration time.
        /// </summary>
        /// <param name="url">Specify a page URL with which the returned security validation token information is associated.</param>
        /// <returns>The result of GetUpdatedFormDigestInformation operation that contains a security validation token generated by the protocol server, the security validation token’s expiration time in seconds, and other information.</returns>
        FormDigestInformation GetUpdatedFormDigestInformation(string url);

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
        int ImportWeb(string jobName, string webUrl, string[] dataFiles, string logPath, bool includeUserSecurity, bool overWrite);

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
        CreateWebResponseCreateWebResult CreateWeb(string url, string title, string description, string templateName, uint language, bool languageSpecified, uint locale, bool localeSpecified, uint collationLocale, bool collationLocaleSpecified, bool uniquePermissions, bool uniquePermissionsSpecified, bool anonymous, bool anonymousSpecified, bool presence, bool presenceSpecified);

        /// <summary>
        /// Delete an existing subsite of the current site. 
        /// </summary>
        /// <param name="url">The site-relative URL of the subsite to be deleted.</param>
        void DeleteWeb(string url);
    }
}