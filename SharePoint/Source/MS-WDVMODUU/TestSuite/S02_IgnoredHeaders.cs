namespace Microsoft.Protocols.TestSuites.MS_WDVMODUU
{
    using System.Collections.Specialized;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    
    /// <summary>
    /// Traditional tests for IgnoredHeaders scenario.
    /// </summary>
    [TestClass]
    public class S02_IgnoredHeaders : TestSuiteBase
    {
        #region ClassInitialize method

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.Initialize(testContext);
        }

        #endregion

        #region ClassCleanup method

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestSuiteBase.Cleanup();
        }

        #endregion

        #region IgnoredHeaders test cases

        /// <summary>
        /// This test case is used to partially test that the server ignores some headers in HTTP GET request.
        /// </summary>
        [TestCategory("MSWDVMODUU"), TestMethod()]
        public void MSWDVMODUU_S02_TC01_IgnoredHeaders_Get()
        {
            // This test case is used to partially test that the server ignores following headers in HTTP GET request.
            //     Moss-Uid
            //     Moss-Did
            //     Moss-VerFrom
            //     Moss-CBFile
            //     MS-Set-Repl-Uid
            //     MS-BinDiff
            //     X-Office-Version
            // This test case calls the private help method "CompareHttpResponses_Get" for each above ignored header to partially test ignored headers.
            // This help method has two input parameters, one is for the ignored header name, and another is for the value of the ignored header. 
            // The method "CompareHttpResponses_Get" will call HTTP GET method twice, in the first time the HTTP GET request includes the ignored header, 
            // and in the second time the HTTP GET request does NOT include the ignored header. 
            // And then the help method compare the key data in the two HTTP responses, if they are same then the help method return true, else return false.

            // Get the request URI from the property "Server_NewFile001Uri".
            string requestUri = Common.GetConfigurationPropertyValue("Server_NewFile001Uri", this.Site);

            // Test the ignored header "Moss-Uid" in HTTP GET method.
            // Call private help method "CompareHttpResponses_Get" with the ignored header "Moss-Uid" and its value.
            // And capture MS-WDVMODUU_R80 if the method "CompareHttpResponses_Get" return true.
            bool doesCaptureR80 = false;
            doesCaptureR80 = this.CompareHttpResponses_Get(requestUri, "Moss-Uid", "{E6AA0E42-D27C-4FD8-89C6-EDB73AB1C741}");
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureR80,
                80,
                @"[In Moss-Uid Header] [The reply of implementation does be the same whether Moss-Uid header is included in the request or not.]");

            // Test the ignored header "Moss-Did" in HTTP GET method.
            // Call private help method "CompareHttpResponses_Get" with the ignored header "Moss-Did" and its value.
            // And capture MS-WDVMODUU_R82 if the method "CompareHttpResponses_Get" return true.
            bool doesCaptureR82 = false;
            doesCaptureR82 = this.CompareHttpResponses_Get(requestUri, "Moss-Did", "{A5660731-CA3B-8654-B786-76A540E7AD34}");
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureR82,
                82,
                @"[In Moss-Did Header] [The reply of implementation does be the same whether Moss-Did header is included in the request or not.]");

            // Test the ignored header "Moss-VerFrom" in HTTP GET method.
            // Call private help method "CompareHttpResponses_Get" with the ignored header "Moss-VerFrom" and its value.
            // And capture MS-WDVMODUU_R84 if the method "CompareHttpResponses_Get" return true.
            bool doesCaptureR84 = false;
            doesCaptureR84 = this.CompareHttpResponses_Get(requestUri, "Moss-VerFrom", "1");
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureR84,
                84,
                @"[In Moss-VerFrom Header] [The reply of implementation does be the same whether Moss-VerFrom header is included in the request or not.]");

            // Test the ignored header "Moss-CBFile" in HTTP GET method.
            // Call private help method "CompareHttpResponses_Get" with the ignored header "Moss-CBFile" and its value.
            // And capture MS-WDVMODUU_R86 if the method "CompareHttpResponses_Get" return true.
            bool doesCaptureR86 = false;
            doesCaptureR86 = this.CompareHttpResponses_Get(requestUri, "Moss-CBFile", "0");
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureR86,
                86,
                @"[In Moss-CBFile Header] [The reply of implementation does be the same whether Moss-CBFile header is included in the request or not.]");

            // Test the ignored header "MS-Set-Repl-Uid" in HTTP GET method.
            // Call private help method "CompareHttpResponses_Get" with the ignored header "MS-Set-Repl-Uid" and its value.
            // And capture MS-WDVMODUU_R88 if the method "CompareHttpResponses_Get" return true.
            bool doesCaptureR88 = false;
            doesCaptureR88 = this.CompareHttpResponses_Get(requestUri, "MS-Set-Repl-Uid", "rid:{E819DFCB-DB60-49D7-A70E-51E31F5344BE}");
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureR88,
                88,
                @"[In MS-Set-Repl-Uid Header] [The reply of implementation does be the same whether MS-Set-Repl-Uid header is included in the request or not.]");

            // Test the ignored header "MS-BinDiff" in HTTP GET method.
            // Call private help method "CompareHttpResponses_Get" with the ignored header "MS-BinDiff" and its value.
            // And capture MS-WDVMODUU_R90 if the method "CompareHttpResponses_Get" return true.
            bool doesCaptureR90 = false;
            doesCaptureR90 = this.CompareHttpResponses_Get(requestUri, "MS-BinDiff", "1.0");
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureR90,
                90,
                @"[In MS-BinDiff Header] [The reply of implementation does be the same whether MS-BinDiff header is included in the request or not.]");

            // Test the ignored header "X-Office-Version" in HTTP GET method.
            // Call private help method "CompareHttpResponses_Get" with the ignored header "X-Office-Version" and its value.
            // And capture MS-WDVMODUU_R92 if the method "CompareHttpResponses_Get" return true.
            bool doesCaptureR92 = false;
            doesCaptureR92 = this.CompareHttpResponses_Get(requestUri, "X-Office-Version", "12.0.6234");
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureR92,
                92,
                @"[In X-Office-Version Header] [The reply of implementation does be the same whether X-Office-Version header is included in the request or not.]");
        }

        /// <summary>
        /// This test case is used to partially test that the server ignores some headers in HTTP PUT request.
        /// </summary>
        [TestCategory("MSWDVMODUU"), TestMethod()]
        public void MSWDVMODUU_S02_TC02_IgnoredHeaders_Put()
        {
            // This test case is used to partially test that the server ignores following headers in HTTP PUT request.
            //     Moss-Uid
            //     Moss-Did
            //     Moss-VerFrom
            //     Moss-CBFile
            //     MS-Set-Repl-Uid
            //     X-Office-Version
            // This test case calls the private help method "CompareHttpResponses_Put" for each above ignored header to partially test ignored headers.
            // This help method has two input parameters, one is for the ignored header name, and another is for the value of the ignored header. 
            // The method "CompareHttpResponses_Put" will call HTTP PUT method and DELETE method twice, PUT method will upload a test file to the server, 
            // and the DELETE method will remove the test file from the server. 
            // In the first time the HTTP PUT and DELETE requests do NOT include the ignored header, and in the second time the HTTP PUT and DELETE requests include the ignored header. 
            // And then the help method compare the key data in the HTTP responses, if they are same then the help method return true, else return false.

            // Get the request URI from the property "Server_TestTxtFileUri_Put".
            string requestUri = Common.GetConfigurationPropertyValue("Server_TestTxtFileUri_Put", this.Site);

            // Get file name from the property "Client_TestTxtFileName", and read its content into byte array.
            byte[] bytesTxtFile = GetLocalFileContent(Common.GetConfigurationPropertyValue("Client_TestTxtFileName", this.Site));

            // Test the ignored header "Moss-Uid" in HTTP PUT method.
            // Call private help method "CompareHttpResponses_Put" with the ignored header "Moss-Uid" and its value.
            // And capture MS-WDVMODUU_R80 if the method "CompareHttpResponses_Put" return true.
            bool doesCaptureR80 = false;
            doesCaptureR80 = this.CompareHttpResponses_Put(requestUri, bytesTxtFile, "Moss-Uid", "{E6AA0E42-D27C-4FD8-89C6-EDB73AB1C741}");
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureR80,
                80,
                @"[In Moss-Uid Header] [The reply of implementation does be the same whether Moss-Uid header is included in the request or not.]");

            // Test the ignored header "Moss-Did" in HTTP PUT method.
            // Call private help method "CompareHttpResponses_Put" with the ignored header "Moss-Did" and its value.
            // And capture MS-WDVMODUU_R82 if the method "CompareHttpResponses_Put" return true.
            bool doesCaptureR82 = false;
            doesCaptureR82 = this.CompareHttpResponses_Put(requestUri, bytesTxtFile, "Moss-Did", "{A5660731-CA3B-8654-B786-76A540E7AD34}");
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureR82,
                82,
                @"[In Moss-Did Header] [The reply of implementation does be the same whether Moss-Did header is included in the request or not.]");

            // Test the ignored header "Moss-VerFrom" in HTTP PUT method.
            // Call private help method "CompareHttpResponses_Put" with the ignored header "Moss-VerFrom" and its value.
            // And capture MS-WDVMODUU_R84 if the method "CompareHttpResponses_Put" return true.            
            bool doesCaptureR84 = false;
            doesCaptureR84 = this.CompareHttpResponses_Put(requestUri, bytesTxtFile, "Moss-VerFrom", "1");
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureR84,
                84,
                @"[In Moss-VerFrom Header] [The reply of implementation does be the same whether Moss-VerFrom header is included in the request or not.]");

            // Test the ignored header "Moss-CBFile" in HTTP PUT method.
            // Call private help method "CompareHttpResponses_Put" with the ignored header "Moss-CBFile" and its value.
            // And capture MS-WDVMODUU_R86 if the method "CompareHttpResponses_Put" return true.             
            bool doesCaptureR86 = false;
            doesCaptureR86 = this.CompareHttpResponses_Put(requestUri, bytesTxtFile, "Moss-CBFile", bytesTxtFile.Length.ToString());
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureR86,
                86,
                @"[In Moss-CBFile Header] [The reply of implementation does be the same whether Moss-CBFile header is included in the request or not.]");

            // Test the ignored header "MS-Set-Repl-Uid" in HTTP PUT method.
            // Call private help method "CompareHttpResponses_Put" with the ignored header "MS-Set-Repl-Uid" and its value.
            // And capture MS-WDVMODUU_R88 if the method "CompareHttpResponses_Put" return true.    
            bool doesCaptureR88 = false;
            doesCaptureR88 = this.CompareHttpResponses_Put(requestUri, bytesTxtFile, "MS-Set-Repl-Uid", "rid:{E819DFCB-DB60-49D7-A70E-51E31F5344BE}");
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureR88,
                88,
                @"[In MS-Set-Repl-Uid Header] [The reply of implementation does be the same whether MS-Set-Repl-Uid header is included in the request or not.]");

            // Test the ignored header "X-Office-Version" in HTTP PUT method.
            // Call private help method "CompareHttpResponses_Put" with the ignored header "X-Office-Version" and its value.
            // And capture MS-WDVMODUU_R92 if the method "CompareHttpResponses_Put" return true.                
            bool doesCaptureR92 = false;
            doesCaptureR92 = this.CompareHttpResponses_Put(requestUri, bytesTxtFile, "X-Office-Version", "12.0.6234");
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureR92,
                92,
                @"[In X-Office-Version Header] [The reply of implementation does be the same whether X-Office-Version header is included in the request or not.]");
        }

        /// <summary>
        /// This test case is used to verify that, in Windows SharePoint Services 3.0, the reply of implementation does be the same whether "SyncMan []" is included in the User-Agent Header or not.
        /// </summary>
        [TestCategory("MSWDVMODUU"), TestMethod()]
        public void MSWDVMODUU_S02_TC03_IgnoredHeaders_UserAgent()
        {
            if (Common.IsRequirementEnabled(128, this.Site))
            {
                string requestUri = Common.GetConfigurationPropertyValue("Server_NewFile001Uri", this.Site);
                
                // Call HTTP GET method with User-Agent header that include "SyncMan []" comments.
                WDVMODUUResponse responseForGetRequestWithSyncManComment = null;
                NameValueCollection headersCollectionWithSyncManComment = this.GetHttpHeadersCollection(null, null);
                headersCollectionWithSyncManComment.Add("User-Agent", "Microsoft Office/12.0 (Windows NT 5.2; SyncMan 12.0.6234; Pro)");
                headersCollectionWithSyncManComment.Add("X-Office-Version", "12.0.6234");
                responseForGetRequestWithSyncManComment = this.Adapter.Get(requestUri, headersCollectionWithSyncManComment);

                // Call HTTP GET method with User-Agent header that does not include "SyncMan []" comments.
                WDVMODUUResponse responseForGetRequestWithoutSyncManComment = null;
                NameValueCollection headersCollectionWithoutSyncManComment = this.GetHttpHeadersCollection(null, null);
                headersCollectionWithoutSyncManComment.Add("User-Agent", "Microsoft Office/12.0 (Windows NT 5.2; Pro)");
                headersCollectionWithoutSyncManComment.Add("X-Office-Version", "12.0.6234");
                responseForGetRequestWithoutSyncManComment = this.Adapter.Get(requestUri, headersCollectionWithoutSyncManComment);

                #region Compare the two above responses of HTTP GET method
                string errorLog = string.Empty;

                // Compare status code values in the two response.
                bool isSameStatusCode = false;
                if (responseForGetRequestWithSyncManComment.StatusCode == responseForGetRequestWithoutSyncManComment.StatusCode)
                {
                    isSameStatusCode = true;
                }
                else
                {
                    isSameStatusCode = false;
                    errorLog = "Test Failed: The status codes in two response are not same! \r\n";
                    errorLog += string.Format("responseForGetRequestWithSyncManComment.StatusCode={0}\r\n", responseForGetRequestWithSyncManComment.StatusCode);
                    errorLog += string.Format("responseForGetRequestWithoutSyncManComment.StatusCode={0}\r\n", responseForGetRequestWithoutSyncManComment.StatusCode);
                    this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
                }

                // Compare status description values in the two response.
                bool isSameStatusDescription = false;
                if (string.Compare(responseForGetRequestWithSyncManComment.StatusDescription, responseForGetRequestWithoutSyncManComment.StatusDescription, true) == 0)
                {
                    isSameStatusDescription = true;
                }
                else
                {
                    isSameStatusDescription = false;
                    errorLog = "Test Failed: The status description in two response are not same! \r\n";
                    errorLog += string.Format("responseForGetRequestWithSyncManComment.StatusDescription={0}\r\n", responseForGetRequestWithSyncManComment.StatusDescription);
                    errorLog += string.Format("responseForGetRequestWithoutSyncManComment.StatusDescription={0}\r\n", responseForGetRequestWithoutSyncManComment.StatusDescription);
                    this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
                }

                // Compare content length values in the two response.
                bool isSameContentLength = false;
                if (responseForGetRequestWithSyncManComment.ContentLength == responseForGetRequestWithoutSyncManComment.ContentLength)
                {
                    isSameContentLength = true;
                }
                else
                {
                    isSameContentLength = false;
                    errorLog = "Test Failed: The content length in two response are not same! \r\n";
                    errorLog += string.Format("responseForGetRequestWithSyncManComment.ContentLength={0}\r\n", responseForGetRequestWithSyncManComment.ContentLength);
                    errorLog += string.Format("responseForGetRequestWithoutSyncManComment.ContentLength={0}\r\n", responseForGetRequestWithoutSyncManComment.ContentLength);
                    this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
                }

                // Compare content type values in the two response.
                bool isSameContentType = false;
                if (string.Compare(responseForGetRequestWithSyncManComment.ContentType, responseForGetRequestWithoutSyncManComment.ContentType, true) == 0)
                {
                    isSameContentType = true;
                }
                else
                {
                    isSameContentType = false;
                    errorLog = "Test Failed: The content type in two response are not same! \r\n";
                    errorLog += string.Format("responseForGetRequestWithSyncManComment.ContentType={0}\r\n", responseForGetRequestWithSyncManComment.ContentType);
                    errorLog += string.Format("responseForGetRequestWithoutSyncManComment.ContentType={0}\r\n", responseForGetRequestWithoutSyncManComment.ContentType);
                    this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
                }

                // Compare body data values in the two response.
                bool isSameBodyData = false;
                if (responseForGetRequestWithSyncManComment.BodyData == responseForGetRequestWithoutSyncManComment.BodyData)
                {
                    isSameBodyData = true;
                }
                else
                {
                    isSameBodyData = false;
                    errorLog = "Test Failed: The body data in two response are not same! \r\n";
                    errorLog += string.Format("responseForGetRequestWithSyncManComment.BodyData={0}\r\n", responseForGetRequestWithSyncManComment.BodyData);
                    errorLog += string.Format("responseForGetRequestWithoutSyncManComment.BodyData={0}\r\n", responseForGetRequestWithoutSyncManComment.BodyData);
                    this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
                }
                #endregion Compare the two above responses of HTTP GET method

                // Capture MS-WDVMODUU_R128, if the key data in the above two responses of HTTP GET method are same.
                bool doesCaptureR128 = false;
                if (isSameStatusCode && isSameStatusDescription && isSameContentLength && isSameContentType && isSameBodyData)
                {
                    doesCaptureR128 = true;
                }

                this.Site.CaptureRequirementIfIsTrue(
                    doesCaptureR128,
                    128,
                    @"[In Appendix B: Product Behavior] Implementation does reply the same response whether ""SyncMan []"" is included in the User-Agent Header or not.[In Appendix B: Product Behavior] <9> Section 2.2.1.9:  Servers running Windows SharePoint Services 3.0 ignore comments [""SyncMan[]""] of this value in the User-Agent Header.");
            }
            else
            {
                this.Site.Assume.Inconclusive("Test is executed only when R128Enabled is set to true.");
            }
        }

        #endregion

        #region Help methods in Scenario 02

        /// <summary>
        /// Get the HTTP collection that main HTTP headers and their values are included. 
        /// If the special HTTP header and its value are set, then the special HTTP header and the value is also included the HTTP collection
        /// </summary>
        /// <param name="specialHttpHeaderName">The special HTTP header name.</param>
        /// <param name="specialHttpHeaderValue">The special HTTP header value.</param>
        /// <returns>Return the HTTP collection that is constructed in this method. </returns>
        private NameValueCollection GetHttpHeadersCollection(string specialHttpHeaderName, string specialHttpHeaderValue)
        {
            NameValueCollection headersCollection = new NameValueCollection();
            headersCollection.Add("Cache-Control", "no-cache");
            headersCollection.Add("ContentType", "text/xml");
            headersCollection.Add("Depth", "0");
            headersCollection.Add("Pragma", "no-cache");
            headersCollection.Add("ProtocolVersion", "HTTP/1.1");
            if ((specialHttpHeaderName != null) && (specialHttpHeaderValue != null)
                && (specialHttpHeaderName != string.Empty) && (specialHttpHeaderValue != string.Empty))
            {
                headersCollection.Add(specialHttpHeaderName, specialHttpHeaderValue);
            }
            else
            {
                string errorInfo = "In GetHttpHeadersCollection, the input parameters are not correct!";
                if (specialHttpHeaderName != null)
                {
                    errorInfo += "\r\n specialHttpHeaderName = " + specialHttpHeaderName;
                }
                else
                {
                    errorInfo += "\r\n specialHttpHeaderName is null! ";
                }

                if (specialHttpHeaderValue != null)
                {
                    errorInfo += "\r\n specialHttpHeaderValue = " + specialHttpHeaderValue;
                }
                else
                {
                    errorInfo += "\r\n specialHttpHeaderValue is null! ";
                }

                this.Site.Log.Add(LogEntryKind.TestError, errorInfo);
            }

            return headersCollection;
        }

        /// <summary>
        /// The method will call HTTP GET method twice, in the first time the HTTP GET request includes the ignored header, 
        /// and in the second time the HTTP GET request does NOT include the ignored header. 
        /// And then the help method compare the key data in the two HTTP responses, if they are same then the help method return true, else return false.
        /// </summary>
        /// <param name="requestUri">The request URI that is used in HTTP GET method</param>
        /// <param name="ignoredHttpHeaderName">The ignored header name.</param>
        /// <param name="ignoredHttpHeaderValue">The ignored header value.</param>
        /// <returns>Return true if the key data are same by comparing the responses, else return false.</returns>
        private bool CompareHttpResponses_Get(string requestUri, string ignoredHttpHeaderName, string ignoredHttpHeaderValue)
        {
            this.Site.Assert.IsNotNull(requestUri, "In CompareHttpResponses_Get method, the request URI should not be null!");
            this.Site.Assert.IsTrue(requestUri != string.Empty, "In CompareHttpResponses_Get method, the request URI should not be empty string!");
            this.Site.Assert.IsNotNull(ignoredHttpHeaderName, "In CompareHttpResponses_Get method, the ignored header name should not be null!");
            this.Site.Assert.IsNotNull(ignoredHttpHeaderValue, "In CompareHttpResponses_Get method, the ignored header value should not be null!");
            this.Site.Assert.IsTrue(ignoredHttpHeaderName != string.Empty, "In CompareHttpResponses_Get method, the ignored header name should not be empty string!");
            this.Site.Assert.IsTrue(ignoredHttpHeaderValue != string.Empty, "In CompareHttpResponses_Get method, the ignored header value should not be empty string!");

            // Call HTTP GET method with ignored header.
            WDVMODUUResponse responseForGetRequestWithIgnoredHeader = null;
            NameValueCollection headersCollectionWithIgnoredHeader = this.GetHttpHeadersCollection(ignoredHttpHeaderName, ignoredHttpHeaderValue);
            responseForGetRequestWithIgnoredHeader = this.Adapter.Get(requestUri, headersCollectionWithIgnoredHeader);

            // Call HTTP GET method without ignored header.
            WDVMODUUResponse responseForGetRequestWithoutIgnoredHeader = null;
            NameValueCollection headersCollectionWithoutIgnoredHeader = this.GetHttpHeadersCollection(null, null);
            responseForGetRequestWithoutIgnoredHeader = this.Adapter.Get(requestUri, headersCollectionWithoutIgnoredHeader);

            string errorLog = string.Empty;
            this.Site.Log.Add(TestTools.LogEntryKind.Comment, "The ignored header is {0}. The ignored header value is {1}.", ignoredHttpHeaderName, ignoredHttpHeaderValue);

            #region Compare the two responses of HTTP GET method
            this.Site.Log.Add(TestTools.LogEntryKind.Comment, "Compare the two responses of HTTP GET method.");

            // Compare status code values in the two response.
            bool isSameStatusCode = false;
            if (responseForGetRequestWithIgnoredHeader.StatusCode == responseForGetRequestWithoutIgnoredHeader.StatusCode)
            {
                isSameStatusCode = true;
            }
            else
            {
                isSameStatusCode = false;
                errorLog = "Test Failed: The status codes in two response are not same! \r\n";
                errorLog += string.Format("responseForGetRequestWithIgnoredHeader.StatusCode={0}\r\n", (int)responseForGetRequestWithIgnoredHeader.StatusCode);
                errorLog += string.Format("responseForGetRequestWithoutIgnoredHeader.StatusCode={0}\r\n", (int)responseForGetRequestWithoutIgnoredHeader.StatusCode);
                this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
            }

            // Compare status description values in the two response.
            bool isSameStatusDescription = false;
            if (string.Compare(responseForGetRequestWithIgnoredHeader.StatusDescription, responseForGetRequestWithoutIgnoredHeader.StatusDescription, true) == 0)
            {
                isSameStatusDescription = true;
            }
            else
            {
                isSameStatusDescription = false;
                errorLog = "Test Failed: The status description in two response are not same! \r\n";
                errorLog += string.Format("responseForGetRequestWithIgnoredHeader.StatusDescription={0}\r\n", responseForGetRequestWithIgnoredHeader.StatusDescription);
                errorLog += string.Format("responseForGetRequestWithoutIgnoredHeader.StatusDescription={0}\r\n", responseForGetRequestWithoutIgnoredHeader.StatusDescription);
                this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
            }

            // Compare content length values in the two response.
            bool isSameContentLength = false;
            if (responseForGetRequestWithIgnoredHeader.ContentLength == responseForGetRequestWithoutIgnoredHeader.ContentLength)
            {
                isSameContentLength = true;
            }
            else
            {
                isSameContentLength = false;
                errorLog = "Test Failed: The content length in two response are not same! \r\n";
                errorLog += string.Format("responseForGetRequestWithIgnoredHeader.ContentLength={0}\r\n", responseForGetRequestWithIgnoredHeader.ContentLength);
                errorLog += string.Format("responseForGetRequestWithoutIgnoredHeader.ContentLength={0}\r\n", responseForGetRequestWithoutIgnoredHeader.ContentLength);
                this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
            }

            // Compare content type values in the two response.
            bool isSameContentType = false;
            if (string.Compare(responseForGetRequestWithIgnoredHeader.ContentType, responseForGetRequestWithoutIgnoredHeader.ContentType, true) == 0)
            {
                isSameContentType = true;
            }
            else
            {
                isSameContentType = false;
                errorLog = "Test Failed: The content type in two response are not same! \r\n";
                errorLog += string.Format("responseForGetRequestWithIgnoredHeader.ContentType={0}\r\n", responseForGetRequestWithIgnoredHeader.ContentType);
                errorLog += string.Format("responseForGetRequestWithoutIgnoredHeader.ContentType={0}\r\n", responseForGetRequestWithoutIgnoredHeader.ContentType);
                this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
            }

            // Compare body data values in the two response.
            bool isSameBodyData = false;
            if (responseForGetRequestWithIgnoredHeader.BodyData == responseForGetRequestWithoutIgnoredHeader.BodyData)
            {
                isSameBodyData = true;
            }
            else
            {
                isSameBodyData = false;
                errorLog = "Test Failed: The body data in two response are not same! \r\n";
                errorLog += string.Format("responseForGetRequestWithIgnoredHeader.BodyData={0}\r\n", responseForGetRequestWithIgnoredHeader.BodyData);
                errorLog += string.Format("responseForGetRequestWithoutIgnoredHeader.BodyData={0}\r\n", responseForGetRequestWithoutIgnoredHeader.BodyData);
                this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
            }
            #endregion Compare the two responses of HTTP GET method

            // Judge if the two response are same.
            bool isSameResponse = false;
            if (isSameStatusCode && isSameStatusDescription && isSameContentLength && isSameContentType && isSameBodyData)
            {
                isSameResponse = true;
            }

            return isSameResponse;
        }

        /// <summary>
        /// The method "CompareHttpResponses_Put" will call HTTP PUT method and DELETE method twice, PUT method will upload a test file to the server, 
        /// and the DELETE method will remove the test file from the server. 
        /// In the first time the HTTP PUT and DELETE requests do NOT include the ignored header, and in the second time the HTTP PUT and DELETE requests include the ignored header. 
        /// And then the help method compare the key data in the HTTP responses, if they are same then the help method return true, else return false.
        /// </summary>
        /// <param name="requestUri">The request URI that is used in HTTP PUT and DELETE method</param>
        /// <param name="bytesTxtFile">The file content that will be used in HTTP body</param>
        /// <param name="ignoredHttpHeaderName">The ignored header name.</param>
        /// <param name="ignoredHttpHeaderValue">The ignored header value.</param>
        /// <returns>Return true if the key data are same by comparing the responses, else return false.</returns>
        private bool CompareHttpResponses_Put(string requestUri, byte[] bytesTxtFile, string ignoredHttpHeaderName, string ignoredHttpHeaderValue)
        {
            this.Site.Assert.IsNotNull(ignoredHttpHeaderName, "In CompareHttpResponses_Put method, the ignored header name should not be null!");
            this.Site.Assert.IsNotNull(ignoredHttpHeaderValue, "In CompareHttpResponses_Put method, the ignored header value should not be null!");
            this.Site.Assert.IsTrue(ignoredHttpHeaderName != string.Empty, "In CompareHttpResponses_Put method, the ignored header name should not be empty string!");
            this.Site.Assert.IsTrue(ignoredHttpHeaderValue != string.Empty, "In CompareHttpResponses_Put method, the ignored header value should not be empty string!");

            WDVMODUUResponse responseForPutRequestWithIgnoredHeader = null;
            WDVMODUUResponse responseForPutRequestWithoutIgnoredHeader = null;

            WDVMODUUResponse responseForDeleteRequestWithIgnoredHeader = null;
            WDVMODUUResponse responseForDeleteRequestWithoutIgnoredHeader = null;

            NameValueCollection headersCollectionWithIgnoredHeader = this.GetHttpHeadersCollection(ignoredHttpHeaderName, ignoredHttpHeaderValue);
            NameValueCollection headersCollectionWithoutIgnoredHeader = this.GetHttpHeadersCollection(null, null);

            // Call HTTP PUT method without ignored header.
            responseForPutRequestWithoutIgnoredHeader = this.Adapter.Put(requestUri, bytesTxtFile, headersCollectionWithoutIgnoredHeader);
            this.ArrayListForDeleteFile.Add(requestUri);

            // Call HTTP DELETE method without ignored header.
            responseForDeleteRequestWithoutIgnoredHeader = this.Adapter.Delete(requestUri, headersCollectionWithoutIgnoredHeader);
            this.RemoveFileUriFromDeleteList(requestUri);

            // Call HTTP PUT method with ignored header.
            responseForPutRequestWithIgnoredHeader = this.Adapter.Put(requestUri, bytesTxtFile, headersCollectionWithIgnoredHeader);
            this.ArrayListForDeleteFile.Add(requestUri);

            // Call HTTP DELETE method with ignored header.
            responseForDeleteRequestWithIgnoredHeader = this.Adapter.Delete(requestUri, headersCollectionWithIgnoredHeader);
            this.RemoveFileUriFromDeleteList(requestUri);

            string errorLog = string.Empty;
            this.Site.Log.Add(TestTools.LogEntryKind.Comment, "The ignored header is {0}. The ignored header value is {1}.", ignoredHttpHeaderName, ignoredHttpHeaderValue);

            #region Compare the two responses of HTTP PUT method
            this.Site.Log.Add(TestTools.LogEntryKind.Comment, "Compare the two responses of HTTP PUT method.");

            // Compare status code values in the two response.
            bool isSameStatusCode_Put = false;
            if (responseForPutRequestWithIgnoredHeader.StatusCode == responseForPutRequestWithoutIgnoredHeader.StatusCode)
            {
                isSameStatusCode_Put = true;
            }
            else
            {
                isSameStatusCode_Put = false;
                errorLog = "Test Failed: The status codes in two response are not same! \r\n";
                errorLog += string.Format("responseForPutRequestWithIgnoredHeader.StatusCode={0}\r\n", (int)responseForPutRequestWithIgnoredHeader.StatusCode);
                errorLog += string.Format("responseForPutRequestWithoutIgnoredHeader.StatusCode={0}\r\n", (int)responseForPutRequestWithoutIgnoredHeader.StatusCode);
                this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
            }

            // Compare status description values in the two response.
            bool isSameStatusDescription_Put = false;
            if (string.Compare(responseForPutRequestWithIgnoredHeader.StatusDescription, responseForPutRequestWithoutIgnoredHeader.StatusDescription, true) == 0)
            {
                isSameStatusDescription_Put = true;
            }
            else
            {
                isSameStatusDescription_Put = false;
                errorLog = "Test Failed: The status description in two response are not same! \r\n";
                errorLog += string.Format("responseForPutRequestWithIgnoredHeader.StatusDescription={0}\r\n", responseForPutRequestWithIgnoredHeader.StatusDescription);
                errorLog += string.Format("responseForPutRequestWithoutIgnoredHeader.StatusDescription={0}\r\n", responseForPutRequestWithoutIgnoredHeader.StatusDescription);
                this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
            }

            // Judge if the two response for PUT methods are same.
            bool isSameResponse_Put = false;
            if (isSameStatusCode_Put && isSameStatusDescription_Put)
            {
                isSameResponse_Put = true;
            }

            #endregion Compare the two responses of HTTP PUT method

            #region Compare the two responses of HTTP DELETE method
            this.Site.Log.Add(TestTools.LogEntryKind.Comment, "Compare the two responses of HTTP DELETE method.");

            // Compare status code values in the two response.
            bool isSameStatusCode_Delete = false;
            if (responseForDeleteRequestWithIgnoredHeader.StatusCode == responseForDeleteRequestWithoutIgnoredHeader.StatusCode)
            {
                isSameStatusCode_Delete = true;
            }
            else
            {
                isSameStatusCode_Delete = false;
                errorLog = "Test Failed: The status codes in two response are not same! \r\n";
                errorLog += string.Format("responseForDeleteRequestWithIgnoredHeader.StatusCode={0}\r\n", (int)responseForDeleteRequestWithIgnoredHeader.StatusCode);
                errorLog += string.Format("responseForDeleteRequestWithoutIgnoredHeader.StatusCode={0}\r\n", (int)responseForDeleteRequestWithoutIgnoredHeader.StatusCode);
                this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
            }

            // Compare status description values in the two response.
            bool isSameStatusDescription_Delete = false;
            if (string.Compare(responseForDeleteRequestWithIgnoredHeader.StatusDescription, responseForDeleteRequestWithoutIgnoredHeader.StatusDescription, true) == 0)
            {
                isSameStatusDescription_Delete = true;
            }
            else
            {
                isSameStatusDescription_Delete = false;
                errorLog = "Test Failed: The status description in two response are not same! \r\n";
                errorLog += string.Format("responseForDeleteRequestWithIgnoredHeader.StatusDescription={0}\r\n", responseForDeleteRequestWithIgnoredHeader.StatusDescription);
                errorLog += string.Format("responseForDeleteRequestWithoutIgnoredHeader.StatusDescription={0}\r\n", responseForDeleteRequestWithoutIgnoredHeader.StatusDescription);
                this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, errorLog);
            }

            // Judge if the two response for PUT methods are same.
            bool isSameResponse_Delete = false;
            if (isSameStatusCode_Delete && isSameStatusDescription_Delete)
            {
                isSameResponse_Delete = true;
            }

            #endregion Compare the two responses of HTTP DELETE method

            // Judge if the responses in PUT and DELETE methods are same.
            bool isSameResponse = false;
            if (isSameResponse_Put && isSameResponse_Delete)
            {
                isSameResponse = true;
            }

            return isSameResponse;
        }

        #endregion Help methods in Scenario 02
    }
}