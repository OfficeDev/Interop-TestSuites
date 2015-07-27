//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Linq;
    using System.Web;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This class is used to get the WOPI resource URL
    /// </summary>
    public class TokenAndRequestUrlHelper : HelperBase
    {   
        /// <summary>
        /// A string represents the folder level pattern for WOPI resource URL.
        /// </summary>
        private static string folderLevelPattern = @"HTTP://server/<...>/wopi*/folders/<id>?access_token=<token>";

        /// <summary>
        /// A string represents the file level pattern for WOPI resource URL.
        /// </summary>
        private static string fileLevelPattern = @"HTTP://server/<...>/wopi*/files/<id>?access_token=<token>";

        /// <summary>
        /// A bool value represents whether the TokenAndRequestUrlHelper has been initialized. The value 'true' means it has been initialized.
        /// </summary>
        private static bool hasInitializedTokenAndRequestUrlHelper = false;

        /// <summary>
        /// A ITestSite type value represents the current ITestSite instance of test suite.
        /// </summary>
        private static ITestSite currentTestSite;

        /// <summary>
        /// Prevents a default instance of the TokenAndRequestUrlHelper class from being created
        /// </summary>
        private TokenAndRequestUrlHelper()
        {
        }

        /// <summary>
        /// Gets or sets a string value represents the default domain name which the user belong to. 
        /// </summary>
        public static string DefaultDomain { get; set; }

        /// <summary>
        /// Gets or sets a string value represents the default user name. 
        /// </summary>
        public static string DefaultUserName { get; set; }

        /// <summary>
        /// Gets or sets a string value represents the default password name of the default user.
        /// </summary>
        public static string DefaultPassword { get; set; }

        /// <summary>
        /// Gets or sets a TransportProtocol instance represents the current http transport the test suite uses.
        /// </summary>
        public static TransportProtocol CurrentHttpTransport { get; set; }

        /// <summary>
        /// Gets a value indicating whether the TokenAndRequestUrlHelper has been initialized. The value 'true' means it has been initialized.
        /// </summary>
        public static bool HasInitialized
        {
            get
            {
                return hasInitializedTokenAndRequestUrlHelper;
            }
        }

        /// <summary>
        /// A method is used to initialize the TokenAndRequestUrlHelper helper.
        /// </summary>
        /// <param name="managedCodeSutConrollerAdapterInstance">A parameter represents the IMS_WOPISUTManageCodeControlAdapter type instance, it is used to get WOPI root resource URL.</param>
        /// <param name="testSiteInstance">A parameter represents an ITestSite instance which is used to get test suite context.</param>
        public static void InitializeHelper(IMS_WOPIManagedCodeSUTControlAdapter managedCodeSutConrollerAdapterInstance, ITestSite testSiteInstance)
        {
            HelperBase.CheckInputParameterNullOrEmpty<IMS_WOPIManagedCodeSUTControlAdapter>(managedCodeSutConrollerAdapterInstance, "managedCodeSutConrollerAdapterInstance", "InitializeHelper");
            HelperBase.CheckInputParameterNullOrEmpty<ITestSite>(testSiteInstance, "testSiteInstance", "InitializeHelper");

            if (!hasInitializedTokenAndRequestUrlHelper)
            {
                if (string.IsNullOrEmpty(DefaultUserName))
                {
                    DefaultUserName = Common.GetConfigurationPropertyValue("UserName", testSiteInstance);
                    DefaultDomain = Common.GetConfigurationPropertyValue("Domain", testSiteInstance);
                    DefaultPassword = Common.GetConfigurationPropertyValue("Password", testSiteInstance);
                }

                if (null == currentTestSite)
                {
                    currentTestSite = testSiteInstance;
                }

                CurrentHttpTransport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", testSiteInstance);
                hasInitializedTokenAndRequestUrlHelper = true;
            }
        }
 
        /// <summary>
        /// A method is used to get token value.
        /// </summary>
        /// <param name="wopiResourceUrl">A parameter represents the WOPI resource URL, which is contain token information.</param>
        /// <returns>A return value represents the token value.</returns>
        public static string GetTokenValueFromWOPIResourceUrl(string wopiResourceUrl)
        {
            if (string.IsNullOrEmpty(wopiResourceUrl))
            {
                throw new ArgumentNullException("wopiResourceUrl");
            }

            Dictionary<string, string> wopiSrcAndTokenValues = GetWOPISrcAndTokenValueFromWOPIResourceUrl(wopiResourceUrl);
            return wopiSrcAndTokenValues["Token"];
        }

        /// <summary>
        /// A method is used to get folder children's level URL.
        /// </summary>
        /// <param name="wopiRootResourceUrl">A parameter represents the WOPI resource URL for the folder level.</param>
        /// <param name="subResourceUrlType">A parameter represents the WOPI sub resource URL for the folder level.</param>
        /// <returns>A return value represents the folder children's level URL.</returns>
        public static string GetSubResourceUrl(string wopiRootResourceUrl, WOPISubResourceUrlType subResourceUrlType)
        {
            #region Verify parameter
            if (string.IsNullOrEmpty(wopiRootResourceUrl))
            {
                throw new ArgumentNullException("wopiRootResourceUrl");
            }

            Uri currentUri = null;
            if (!Uri.TryCreate(wopiRootResourceUrl, UriKind.Absolute, out currentUri))
            {
                throw new ArgumentException("It must be a valid absolute URL", "wopiRootResourceUrl");
            }

            string expectedIncludeStringValue = string.Empty;
            string expectedPatternValue = string.Empty;
            string expectedSubResourceUrlPostfix = string.Empty;
            WOPIRootResourceUrlType expectedRootResourceUrlType = WOPIRootResourceUrlType.FileLevel;
            switch (subResourceUrlType)
            {
                case WOPISubResourceUrlType.FolderChildrenLevel:
                    {
                        expectedIncludeStringValue = @"/folders/";
                        expectedPatternValue = folderLevelPattern;
                        expectedSubResourceUrlPostfix = @"/children";
                        expectedRootResourceUrlType = WOPIRootResourceUrlType.FolderLevel;
                        break;
                    }

                case WOPISubResourceUrlType.FileContentsLevel:
                    {
                        expectedIncludeStringValue = @"/files/";
                        expectedPatternValue = fileLevelPattern;
                         expectedSubResourceUrlPostfix = @"/contents";
                        expectedRootResourceUrlType = WOPIRootResourceUrlType.FileLevel;
                        break;
                    }

                default:
                    {
                       throw new InvalidOperationException(string.Format(@"The test suite only supports [{0}] and [{1}] two WOPI sub resource URL formats.", WOPISubResourceUrlType.FileContentsLevel, WOPISubResourceUrlType.FolderChildrenLevel));
                    }
            }

            if (wopiRootResourceUrl.IndexOf(expectedIncludeStringValue, StringComparison.OrdinalIgnoreCase) < 0)
            {
                string errorMsg = string.Format(
                           @"To getting the [{0}] sub resource URL, the WOPI root resource URL must be [{1}] type, and its format must be [{2}] format.",
                           subResourceUrlType,
                           expectedRootResourceUrlType,
                           expectedPatternValue);

                HelperBase.AppendLogs(typeof(TokenAndRequestUrlHelper), errorMsg);
                throw new InvalidOperationException(errorMsg);
            }
           
            string pathValueOfUrl = currentUri.AbsolutePath;
            if (pathValueOfUrl.EndsWith(@"/contents", StringComparison.OrdinalIgnoreCase)
                || pathValueOfUrl.EndsWith(@"/children", StringComparison.OrdinalIgnoreCase))
            {
                string errorMsg = string.Format(
                          @"The URL value has been WOPI sub resource URL:[{0}]",
                          wopiRootResourceUrl);
                HelperBase.AppendLogs(typeof(TokenAndRequestUrlHelper), errorMsg);
                throw new InvalidOperationException(errorMsg);
            }

            #endregion 

            Dictionary<string, string> wopiSrcAndTokenValues = GetWOPISrcAndTokenValueFromWOPIResourceUrl(wopiRootResourceUrl);
            string wopiSrcValue = wopiSrcAndTokenValues["WOPISrc"];
            string tokenValue = wopiSrcAndTokenValues["Token"];

            // Construct the sub WOPI resource URL.
            string wopiSubResourceUrl = string.Format(
                                        @"{0}{1}{2}{3}",
                                        wopiSrcValue,
                                        expectedSubResourceUrlPostfix,
                                        @"?access_token=",
                                        tokenValue);

            return wopiSubResourceUrl;
        }

        /// <summary>
        /// A method is used to get the WOPIsrc value and the token value.
        /// </summary>
        /// <param name="wopiResourceUrl">A parameter represents the WOPI resource URL</param>
        /// <returns>A return value represents a name-value pairs collection which includes the WOPIsrc value and the token value. The token value can be get by "Token" key, the WOPIsrc value can be get by "WOPISrc" key.</returns>
        protected static Dictionary<string, string> GetWOPISrcAndTokenValueFromWOPIResourceUrl(string wopiResourceUrl)
        {
            if (string.IsNullOrEmpty(wopiResourceUrl))
            {
                throw new ArgumentNullException("wopiResourceUrl");
            }

            Uri currentUri = null;
            if (!Uri.TryCreate(wopiResourceUrl, UriKind.Absolute, out currentUri))
            {
                throw new ArgumentException("It must be a valid absolute URL", "wopiResourceUrl");
            }

            #region verify the URL.

            NameValueCollection queryParameterValues = HttpUtility.ParseQueryString(currentUri.Query);
            if (null == queryParameterValues || 0 == queryParameterValues.Count)
            {
                string errorMsg = string.Format("The WOPI resource URL must contain the URL query parameter. current URL:[{0}] \r\n", wopiResourceUrl);
                HelperBase.AppendLogs(typeof(TokenAndRequestUrlHelper), errorMsg);
                throw new InvalidOperationException(errorMsg);
            }

            string expectedQueryParameter = @"access_token";

            // Verify the query parameters whether contain "access_token"
            if (!queryParameterValues.AllKeys.Any(Founder => Founder.Equals(expectedQueryParameter, StringComparison.OrdinalIgnoreCase)))
            {
                string errorMsg = string.Format("The WOPI resource URL must contain [access_token] parameter. current URL:[{0}] \r\n", wopiResourceUrl);
                HelperBase.AppendLogs(typeof(TokenAndRequestUrlHelper), errorMsg);
                throw new InvalidOperationException(errorMsg);
            }

            #endregion 

            string wopiSrcUrl = currentUri.GetLeftPart(UriPartial.Path);
            Dictionary<string, string> wopiSrcAndTokenValues = new Dictionary<string, string>();
            wopiSrcAndTokenValues.Add("WOPISrc", wopiSrcUrl);
            wopiSrcAndTokenValues.Add("Token", queryParameterValues[expectedQueryParameter]);
            return wopiSrcAndTokenValues;
        }
    }
}