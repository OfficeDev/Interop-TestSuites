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
    /// The SUT Control Adapter interface.
    /// </summary>
    public interface IMS_SITESSSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Remove all files in a document library, which is used as the store location for the files exported by the ExportWeb or ExportWorkflowTemplate operation.
        /// </summary>
        /// <param name="siteName">Specify the site name portion in the document library Url, such as "http://.../(siteName)/(webName)/(documentLibraryName)", from where all files should be removed. This value is optional if the document library is created under the root site of the site collection.</param>
        /// <param name="webName">Specify the web name portion in the document library Url, such as "http://.../(siteName)/(webName)/(documentLibraryName)", from where all files should be removed. This value is optional if the document library is created under the root site of the site collection.</param>
        /// <param name="documentLibraryName">Specify the document library name from where all files should be removed. This value should be the same value as the ValidLibraryName property in PTFConfig file.</param>
        /// <returns>A boolean value set to true if the operation to remove all files was successful, false otherwise.</returns>
        [MethodHelp("Remove all files in the specified document library(documentLibraryName), which is used as the store location for the files exported by the ExportWeb and ExportWorkflowTemplate operations," +
            " and the document library should have the same value as the ValidLibraryName property in the MS-SITESS_TestSuite.deployment.ptfconfig file." +
            " The site portion(siteName) and web portion(webName) of the path of the document library are optional if the document library was created under the root site of the site collection." +
            " The return value should be a Boolean value that indicates whether the operation to remove all files was run successfully," +
            " TRUE means the operation was run successfully, FALSE means the operation failed.")]
        bool EmptyDocumentLibrary(string siteName, string webName, string documentLibraryName);

        /// <summary>
        /// Remove all files in solution gallery, which is the store location for the files exported by the ExportSolution operation.
        /// </summary>
        /// <param name="siteName">Specify the site name portion in the solution gallery Url, such as "http://.../(siteName)/(webName)/(solutiongalleryName)", from where solution files should be removed. This value is optional if the solution gallery is created under the root site of the site collection.</param>
        /// <param name="webName">Specify the web name portion in the solution gallery Url, such as "http://.../(siteName)/(webName)/(solutiongalleryName)", from where solution files should be removed. This value is optional if the solution gallery is created under the root site of the site collection.</param>
        /// <param name="solutionGalleryName">Specify the solution gallery name from where all solution files should be removed. This value should be the same value as the SolutionGalleryName property in PTFConfig file.</param>
        /// <returns>A boolean value set to true if the operation to remove all solution files was successful, false otherwise.</returns>
        [MethodHelp("Remove all solutions in the specified solution gallery(solutionGalleryName), which is the store location for the files exported by the ExportSolution operation," +
            " and the solution gallery should have the same values as the SolutionGalleryName property in the MS-SITESS_TestSuite.deployment.ptfconfig file." +
            " For the SharePoint server, this path of the solution gallery cannot be changed." +
            " The site portion(siteName) and web portion(webName) of the path of the solution gallery are optional if the solution gallery was under the root site of the site collection." +
            " The return value should be a Boolean value that indicates whether the operation to remove all solution files was run successfully," +
            " TRUE means the operation was run successfully, FALSE means the operation failed.")]
        bool RemoveAllSolution(string siteName, string webName, string solutionGalleryName);

        /// <summary>
        /// Remove a web with the specified web name. 
        /// </summary>
        /// <param name="webName">Specify the name of the web that will be removed. This value should be the same value as the value of the SubsiteToBeImport property in PTFConfig file.</param>
        /// <returns>A boolean value set to true if the operation to remove the web was successful, false otherwise.</returns>
        [MethodHelp("Remove a web with the specified web name(webName)." +
            " The return value should be a Boolean value that indicates whether the operation to remove the web was run successfully," +
            " TRUE means the operation was run successfully, FALSE means the operation failed.")]
        bool RemoveWeb(string webName);

        /// <summary>
        /// Get specified content file names including extension name of a list on a web, which is chosen to store the exported files.
        /// </summary>
        /// <param name="siteName">Specify the site name portion in the list Url, such as "http://.../(siteName)/(webName)/(listName)". This value is optional if the list is created under the root site of the site collection.</param>
        /// <param name="webName">Specify the web name portion in the list Url, such as "http://.../(siteName)/(webName)/(listName)". This value is optional if the list is created under the root site of the site collection.</param>
        /// <param name="listName">Specify the name of the list from where the content file names should be got. This value should be the same value as the ValidLibraryName property in PTFConfig file when getting the files exported by the ExportWeb and ExportWorkflowTemplate operations, 
        /// or the SolutionGalleryName property when getting the files exported by the ExportSolution operation.</param>
        /// <param name="fileName">Specify the query condition, only the files in the list whose name contain the fileName will be returned. other files will be ignored.</param>
        /// <returns>A string that contains the file names including extension name in the list contents in which each file name should be separated by ';', for example, "filename1;filename2".</returns>
        [MethodHelp("Get specified content file names including extension name of a list on a web, which is chosen to store the exported files," +
            " and should be the same value as the ValidLibraryName property in the MS-SITESS_TestSuite.deployment.ptfconfig file when getting the files exported by the ExportWeb and ExportWorkflowTemplate operations," +
            " or the SolutionGalleryName property when getting the files exported by the ExportSolution operation." +
            " The site portion(siteName) and web portion(webName) should be set as empty if the list is created under the root site of the site collection." +
            " The file portion(fileName) specified the query condition, only the files in the list whose name contain the fileName will be returned. other files will be ignored." +
            " The return value should be a string that contains the exported file names, including extension name, which are separated by ';', for example, \"filename1;filename2\".")]
        string GetDocumentLibraryFileNames(string siteName, string webName, string listName, string fileName);

        /// <summary>
        /// Get properties of the specified web which is created by the CreateWeb operation if it is not null, otherwise, the script will get the properties of the parent web.
        /// </summary>
        /// <param name="siteName">Specify the site name to which the created web belong. This value should be the same value as the SiteName property in PTFConfig file.</param>
        /// <param name="webName">Specify the name of the web. This value should be the same value as the web name portion in the input parameter webUrl of the CreateWeb operation.</param>
        /// <returns>A string that contains the related properties of the web, in which each property should be separated by ';', and the property name and value should be separated by ':'. The returned value should contain the following properties: language, locale, currentUser, permissions, defaultLanguage, anonymous, presence.</returns>
        [MethodHelp("Get properties of the specified web(webName) which is created by the CreateWeb operation if it is not null, otherwise, the script will get the properties of the parent web. " +
            " The site portion(siteName) should be the same value as the SiteName property in the MS-SITESS_TestSuite.deployment.ptfconfig file and web portion(webName) should be the same value as the input parameter webName of the CreateWeb operation." +
            " The return value should be a string that contains the properties, in which each property should be separated by ';', and the property name and value should be separated by ':'," +
            " for example, \"propertyname1:propertyvalue1;propertyname2:propertyvalue2\". The returned value should contain the following properties: " +
            " language: a LCID which value refers to the input parameter 'language' of the CreateWeb operation;" +
            " locale: a locale identifier whose value refers to the input parameter 'locale' in the CreateWeb operation;" +
            " currentUser: the name of the user who has currently logged in;" +
            " permissions: the names of the users who have permissions to the specified website;" +
            " defaultLanguage: a LCID whose value should be the same value as the DefaultLCID property in the MS-SITESS_TestSuite.deployment.ptfconfig file." +
            " anonymous: A boolean value refers to the input parameter 'anonymous' in the CreateWeb operation." +
            " presence: A boolean value refers to the input parameter 'presence' in the CreateWeb operation.")]
        string GetWebProperties(string siteName, string webName);

        /// <summary>
        /// Get the site collection identifier, a Globally unique identifier (GUID) that identifies the site collection.
        /// </summary>
        /// <returns>A string in form of a GUID that specifies the site collection identifier of the site collection which name is preconfigured as the SiteCollectionName property in PTFConfig file.</returns>
        [MethodHelp("Get the site collection identifier, a GUID that identifies the site collection whose name is preconfigured as the SiteCollectionName property in the MS-SITESS_TestSuite.deployment.ptfconfig file. " +
            " The return value should be a string that specifies the GUID that identifies the site collection.")]
        string GetSiteGuid();

        /// <summary>
        /// Set user code to be enabled or disabled on the site collection. User code is managed code that can be uploaded to a site by a site collection administrator.
        /// </summary>
        /// <param name="enable">Indicate whether user code will be enabled for the site collection.</param>
        /// <returns>A boolean value set to true if the user code was enabled, false otherwise.</returns>
        [MethodHelp("Set the user code to be enabled or disabled on the site collection. User code is managed code that can be uploaded to a site by a site collection administrator." +
            " The return value should be a Boolean value that indicates whether user code was enabled successfully," +
            " TRUE means user code was enabled, FALSE means user code was disabled.")]
        bool SetUserCodeEnabled(bool enable);

        /// <summary>
        /// Set the time period in seconds after which the security validation (also named form digest) will expire, which is used to help prevent security attacks where a user unknowingly posts data to a server.
        /// </summary>
        /// <param name="timeout">Specify the timeout value of form digest to be set.</param>
        /// <returns>An integer that specifies the form digest timeout value after this set operation.</returns>
        [MethodHelp("Set the time period(timeout) in seconds after which the security validation will expire. The security validation is generated by the FormDigest control to help prevent the type of attack whereby a user is tricked into the posting data to the server without knowing it." +
            " The return value should be an integer that specifies the input time period, for example, \"60\".")]
        int SetFormDigestTimeout(int timeout);

        /// <summary>
        /// Get a form on a web page and Post the form with the digest value.
        /// </summary>
        /// <param name="digest">Specify the digest value received from server. This value should be the same digest string returned by the GetUpdatedFormDigest or GetUpdatedFormDigestInformation operation.</param>
        /// <param name="webPageUrl">Specify the URL for an uploaded web page that contains a form to be posted. This value should be the same value as the WebPageUrl Property in PTFConfig file.</param>
        /// <returns>A string that specifies the response returned by server, which should contain the keyword indicates whether the web form is posted successfully.</returns>
        [MethodHelp("Post a web form with digest value(digest) set. The digest value should be the same as the digest string returned by the GetUpdatedFormDigest or GetUpdatedFormDigestInformation operation." +
            " If the digest value is not expired, the server will return a response that indicates the web form is posted successfully. Otherwise, the server will return the error information." +
            " The return value should be a string that contains the keyword in the response which indicates whether the web form is posted successfully. For example, If the digest value is not expired," +
            " a string \"Your input\" which is the same value as the constant string 'PostWebFormResponse' defined in Constants.cs should be contained in the return value. Otherwise," +
            " a string \"The security validation for this page has timed out\" which is the same value as the constant string 'TimeOutInformationForSP2007AndSP2010' or 'TimeOutInformationForSP2013' defined in Constants.cs should be contained.")]
        string PostWebForm(string digest, string webPageUrl);

        /// <summary>
        /// Fetch status code from the log file generated by the server for the ExportWeb and ImportWeb operations, which extension name is specified as SNT for SharePoint server.
        /// </summary>
        /// <param name="siteName">Specify the site name portion in the Url where the file stored, such as "http://.../(siteName)/(webName)/(listName)". This value is optional if the list is created under the root site of the site collection.</param>
        /// <param name="webName">Specify the web name portion in the Url where the file stored, such as "http://.../(siteName)/(webName)/(listName)". This value is optional if the list is created under the root site of the site collection.</param>
        /// <param name="listName">Specify the name of the list from where the file stored. This value should be the same value as the ValidLibraryName property in PTFConfig file.</param> 
        /// <param name="fileName">Specify the file name of the log file. This value should be the same value as the input parameter jobName for ExportWeb or ImportWeb operation, with the extension name should be the same value as the constant string SntExtension, for example, "exportweb.snt".</param>
        /// <returns>A non-empty string that indicates the success of the operation or an error code.</returns>
        [MethodHelp("Fetch the status code from the log file(fileName) generated by the server under the specified location(listName) for the ExportWeb and ImportWeb operations, whose extension name is specified as SNT for SharePoint server." +
            " The site portion(siteName) and web portion(webName) of the path of the log file are optional if the list was created under the root site of the site collection." +
            " The return value should be a non-empty string that indicates the success of the operation or an error code. In a Microsoft product, a string starting with the number 0 indicates success, and anything else indicates failure. For example, '0,Export Web Soap Success.' for success, and '4,Invalid Export URL' for failure.")]
        string GetStatusCode(string siteName, string webName, string listName, string fileName);
    }
}