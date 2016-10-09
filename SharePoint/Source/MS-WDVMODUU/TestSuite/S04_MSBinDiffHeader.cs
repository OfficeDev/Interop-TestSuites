namespace Microsoft.Protocols.TestSuites.MS_WDVMODUU
{
    using System;
    using System.Collections.Specialized;
    using System.IO;
    using System.Net;
    using System.Text.RegularExpressions;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Traditional tests for "MSBinDiffHeader" scenario.
    /// </summary>
    [TestClass]
    public class S04_MSBinDiffHeader : TestSuiteBase
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

        #region MSBinDiffHeader test cases

        /// <summary>
        /// This test case is used to verify the "MS-BinDiff header" in the response to an HTTP PUT request.
        /// </summary>
        [TestCategory("MSWDVMODUU"), TestMethod()]
        public void MSWDVMODUU_S04_TC01_MSBinDiffHeader_Put()
        {
            // Get the request URI from the property "Server_TestTxtFileUri_Put".
            string requestUri = Common.GetConfigurationPropertyValue("Server_TestTxtFileUri_Put", this.Site);

            // Get file name from the property "Client_TestTxtFileName", and read its content into byte array.
            byte[] bytesTxtFile = GetLocalFileContent(Common.GetConfigurationPropertyValue("Client_TestTxtFileName", this.Site));

            // Construct the request headers.
            NameValueCollection headersCollection = new NameValueCollection();
            headersCollection.Add("MS-BinDiff", "1.0");
            headersCollection.Add("Cache-Control", "no-cache");
            headersCollection.Add("Pragma", "no-cache");
            headersCollection.Add("ProtocolVersion", "HTTP/1.1");


            // Call HTTP PUT method to upload the file to the server.
            WDVMODUUResponse response = null;
            HttpWebResponse httpWebResponse = null;

            try
            {
                response = this.Adapter.Put(requestUri, bytesTxtFile, headersCollection);
                this.ArrayListForDeleteFile.Add(requestUri);
                this.Site.Assert.Fail("Failed: The virus file should not be put successfully! \r\n The last request is:\r\n {0} The last response is: \r\n {1}", this.Adapter.LastRawRequest, this.Adapter.LastRawResponse);
            }
            catch (WebException webException)
            {
                this.Site.Assert.IsNotNull(webException.Response, "The 'Response' in the caught web exception should not be null!");
                httpWebResponse = (HttpWebResponse)webException.Response;
                if (httpWebResponse.StatusCode != HttpStatusCode.UnsupportedMediaType)
                {
                    // The expected web exception is "Unsupported Media Type", if the caught web exception is not "Unsupported Media Type", then throw the web exception to the framework.
                    throw;
                }

                response = new WDVMODUUResponse();
                response.ReserveResponseData(httpWebResponse, this.Site);
            }

            this.Site.Log.Add(
                TestTools.LogEntryKind.Comment,
                string.Format("In Method 'MSWDVMODUU_S04_TC01_MSBinDiffHeader_Put', the request URI is {0}. In the response, the status code is '{1}'; the status description is '{2}'.", requestUri, response.StatusCode, response.StatusDescription));

            #region Verify Requirements

            // Confirm the server fails the request and respond with a message containing HTTP status code "415 UNSUPPORTED MEDIA TYPE", and then capture MS-WDVMODUU_R901.
            bool isResponseStatusCode415UNSUPPORTEDMEDIATYPE = false;
            if ((response.StatusCode == HttpStatusCode.UnsupportedMediaType)
                && (string.Compare(response.StatusDescription, "UNSUPPORTED MEDIA TYPE", true) == 0))
            {
                isResponseStatusCode415UNSUPPORTEDMEDIATYPE = true;
            }

            if (isResponseStatusCode415UNSUPPORTEDMEDIATYPE == false)
            {
                // Log some information to help users to know why the response does not include "409 CONFLICT". 
                string helpDebugInformation = @"The status code in the HTTP response is not ""415 UNSUPPORTED MEDIA TYPE""\r\n ";
                this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, helpDebugInformation);
            }

            this.Site.CaptureRequirementIfIsTrue(
                isResponseStatusCode415UNSUPPORTEDMEDIATYPE,
                901,
                @"[In MS-BinDiff Header] If the MS-BinDiff header is included in an HTTP PUT request, the server MUST fail the request and response with a message containing HTTP status code ""415 UNSUPPORTED MEDIA TYPE"".");

            #endregion
        }
        #endregion
    }
}