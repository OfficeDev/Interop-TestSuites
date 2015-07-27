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
    /// <summary>
    /// Constants used in the SITESS.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// The computer name of SUT.
        /// </summary>
        public const string SutComputerName = "SutComputerName";

        /// <summary>
        /// The site collection path in the related URL.
        /// </summary>
        public const string SiteCollectionPath = "SiteCollectionPath";

        /// <summary>
        /// The URL of site collection which is used by this test suites.
        /// </summary>
        public const string SiteCollectionUrl = "SiteCollectionUrl";

        /// <summary>
        /// The name of the site created for this test suite.
        /// </summary>
        public const string SiteName = "SiteName";

        /// <summary>
        /// The URL of the site created for this test suite.
        /// </summary>
        public const string SiteUrl = "SiteUrl";

        /// <summary>
        /// The domain name is used to login server.
        /// </summary>
        public const string Domain = "Domain";

        /// <summary>
        /// The user name is used to login server.
        /// </summary>
        public const string UserName = "UserName";

        /// <summary>
        /// The password is used to login server.
        /// </summary>
        public const string Password = "Password";

        /// <summary>
        /// Indicate the transport type! It should be http or https, case insensitive.
        /// </summary>
        public const string TransportType = "TransportType";

        /// <summary>
        /// The URL of the protocol service endpoint.
        /// </summary>
        public const string ServiceUrl = "ServiceUrl";

        /// <summary>
        /// The name of the server product installed on the SUT machine, which contains the system under test(SUT).
        /// </summary>
        public const string SutVersion = "SutVersion";

        /// <summary>
        /// The version of SOAP protocol used to encode and decode requests/responses.
        /// </summary>
        public const string SoapVersion = "SoapVersion";

        /// <summary>
        /// Unauthenticated domain.
        /// </summary>
        public const string UnauthorizedUserDomain = "UnauthorizedUserDomain";

        /// <summary>
        /// Unauthenticated user name.
        /// </summary>
        public const string UnauthorizedUserName = "UnauthorizedUserName";

        /// <summary>
        /// Unauthenticated password.
        /// </summary>
        public const string UnauthorizedUserPassword = "UnauthorizedUserPassword";

        /// <summary>
        /// A valid LCID specifies the language package which is installed on the server.
        /// </summary>
        public const string ValidLCID = "ValidLCID";

        /// <summary>
        /// A valid LCID specifies the language package which is not installed on the server.
        /// </summary>
        public const string NotInstalledLCID = "NotInstalledLCID";

        /// <summary>
        /// An LCID that specifies the language in which the server was originally installed. 
        /// </summary>
        public const string DefaultLCID = "DefaultLCID";

        /// <summary>
        /// A time period in seconds after which the security validation (also named form digest) will expire, which is used to help prevent security attacks where a user unknowingly posts data to a server.
        /// </summary>
        public const string ExpireTimePeriodBySecond = "ExpireTimePeriodBySecond";

        /// <summary>
        /// The default expire time of security validation is 30 minutes for Microsoft product.
        /// </summary>
        public const string DefaultExpireTimePeriod = "DefaultExpireTimePeriod";

        /// <summary>
        /// A time period in seconds for a synchronous XML Web service request to wait the MS-SITESS web service to response.
        /// </summary>
        public const string SoapServiceTimeOut = "SoapServiceTimeOut";

        /// <summary>
        /// A time period in seconds to wait the server to generate all the exported files and log file for the ExportWeb, ExportSolution and ExportWorkflowTemplate operations.
        /// </summary>
        public const string ExportWaitTime = "ExportWaitTime";

        /// <summary>
        /// A time period in seconds to wait the server to generate the result in the log file for the ImportWeb operation.
        /// </summary>
        public const string ImportWebWaitTime = "ImportWebWaitTime";

        /// <summary>
        /// The maximum number of retries if the server could not generate the expected result in the preconfigured time period for some unknown reasons, for example, limit of server resources. If the server still does not complete this sequence after repeating, the operation is considered as failed.
        /// </summary>
        public const string ExportRepeatTime = "ExportRepeatTime";

        /// <summary>
        /// The maximum number of retries if the server could not generate the status code for the ImportWeb operation in the preconfigured time period for some unknown reasons, for example, limit of server resources. If the server still does not complete this sequence after repeating, the operation is considered as failed.
        /// </summary>
        public const string ImportWebRepeatTime = "ImportWebRepeatTime";

        /// <summary>
        /// The URL of a sub site to be exported on which should be no file uploaded.
        /// </summary>
        public const string NormalSubsiteUrl = "NormalSubsiteUrl";

        /// <summary>
        /// The URL of a sub site to be exported on which the special file “Testdata.txt” should be uploaded.
        /// </summary>
        public const string SpecialSubsiteUrl = "SpecialSubsiteUrl";

        /// <summary>
        /// A random string that corresponds to the name of a non-existent computer in the test environment.
        /// </summary>
        public const string NonExistentImportUrl = "NonExistentImportUrl";

        /// <summary>
        /// A random string that corresponds to the name of a non-existent sub site under the site of this test suite.
        /// </summary>
        public const string NonExistentSiteName = "NonExistentSiteName";

        /// <summary>
        /// The name of the document library used as the store location.
        /// </summary>
        public const string ValidLibraryName = "ValidLibraryName";

        /// <summary>
        /// The name of the document library which cannot be used as the store location. 
        /// </summary>
        public const string InvalidLibraryName = "InvalidLibraryName";

        /// <summary>
        /// TThe store location of the exported workflow. 
        /// </summary>
        public const string SolutionGalleryName = "SolutionGalleryName";

        /// <summary>
        /// The name of a declarative workflow template on the server, which is a definition of operations, the sequence of operations, constraints, and timing for a specific process, and can be used to create a new workflow.
        /// </summary>
        public const string WorkflowTemplateName = "WorkflowTemplateName";

        /// <summary>
        /// The name of the workflow template which does not exist.
        /// </summary>
        public const string InvalidWorkflowTemplateName = "InvalidWorkflowTemplateName";

        /// <summary>
        /// The full path of the location on the server where the exported files for the ExportWeb and ExportWorkflowTemplate operations are saved.
        /// </summary>
        public const string DataPath = "DataPath";

        /// <summary>
        /// The URL for an uploaded web page contains a form to be posted. The value of the property is case sensitive.
        /// </summary>
        public const string WebPageUrl = "WebPageUrl";

        /// <summary>
        /// Windows SharePoint Services 3.0 SP3.
        /// </summary>
        public const string WindowsSharePointServices3 = "WindowsSharePointServices3";

        /// <summary>
        /// Microsoft SharePoint Foundation 2010 SP1.
        /// </summary>
        public const string SharePointFoundation2010 = "SharePointFoundation2010";

        /// <summary>
        /// Microsoft SharePoint Foundation 2013 Preview.
        /// </summary>
        public const string SharePointFoundation2013 = "SharePointFoundation2013";

        /// <summary>
        /// Microsoft Office SharePoint Server 2007 SP3.
        /// </summary>
        public const string SharePointServer2007 = "SharePointServer2007";

        /// <summary>
        /// Microsoft SharePoint Server 2010 SP1.
        /// </summary>
        public const string SharePointServer2010 = "SharePointServer2010";

        /// <summary>
        /// Microsoft SharePoint Server 2013 Preview.
        /// </summary>
        public const string SharePointServer2013 = "SharePointServer2013";

        /// <summary>
        /// The name of the property which specifies the language code identifier (LCID) for the display language of the specified subsite got by SUT control adapter.
        /// </summary>
        public const string SubSitePropertyLanguage = "SubSitePropertyLanguage";

        /// <summary>
        /// The name of the property which specifies the locale value of the specified sub site got by SUT control adapter.
        /// </summary>
        public const string SubSitePropertyLocale = "SubSitePropertyLocale";

        /// <summary>
        /// The name of the property which specifies the current user name login on the specified sub site got by SUT control adapter.
        /// </summary>
        public const string SubSitePropertyCurrentUser = "SubSitePropertyCurrentUser";

        /// <summary>
        /// The name of the property which specifies the user name in the collection of permissions for the specified sub site got by SUT control adapter.
        /// </summary>
        public const string SubSitePropertyUserNameInPermissions = "SubSitePropertyUserNameInPermissions";

        /// <summary>
        /// The name of the property which specifies the server's default language got by SUT control adapter.
        /// </summary>
        public const string SubSitePropertyDefaultLanguage = "SubSitePropertyDefaultLanguage";

        /// <summary>
        /// The name of the property which specifies the level of access for anonymous users on the specified sub site got by SUT control adapter.
        /// </summary>
        public const string SubSitePropertyAnonymous = "SubSitePropertyAnonymous";

        /// <summary>
        /// JobName for ExportWeb operation.
        /// </summary>
        public const string ExportJobName = "ExportSite";

        /// <summary>
        /// JobName for ImportWeb operation.
        /// </summary>
        public const string ImportJobName = "ImportSite";

        /// <summary>
        /// Extension for log file for ExportWeb and ImportWeb operations.
        /// </summary>
        public const string SntExtension = ".snt";

        /// <summary>
        /// Extension for content migration package files exported by the ExportWeb operation.
        /// </summary>
        public const string CmpExtension = ".cmp";

        /// <summary>
        /// Extension for solution files exported by the ExportSolution and ExportWorkflowTemplate operations.
        /// </summary>
        public const string WspExtension = ".wsp";

        /// <summary>
        /// Title used in CreateWeb operation.
        /// </summary>
        public const string WebTitle = "webTitle";

        /// <summary>
        /// Description used in CreateWeb operation.
        /// </summary>
        public const string WebDescription = "webDescription";

        /// <summary>
        /// Title used in operations: ExportSolution and ExportWorkflowTemplate.
        /// </summary>
        public const string SolutionTitle = "solutionTitle";

        /// <summary>
        /// Description used in operations: ExportSolution and ExportWorkflowTemplate.
        /// </summary>
        public const string SolutionDescription = "solutionDescription";

        /// <summary>
        /// A string should be contained in the response when the uploaded web form which contains an input textbox is posted successfully.
        /// </summary>
        public const string PostWebFormResponse = "Your input";

        /// <summary>
        /// A string should be contained in the security validation error information for posting the web form has timed out on SharePoint Server 2007 and 2010.
        /// </summary>
        public const string TimeOutInformationForSP2007AndSP2010 = "The security validation for this page has timed out.";

        /// <summary>
        /// A string should be contained in the security validation error information for posting the web form has timed out on SharePoint Server 2013.
        /// </summary>
        public const string TimeOutInformationForSP2013 = "The remote server returned an error: (403) Forbidden.";

        /// <summary>
        /// The char value for item split.
        /// </summary>
        public const char ItemSpliter = ';';

        /// <summary>
        /// The char value for key and value split.
        /// </summary>
        public const char KeySpliter = ':';
    }
}