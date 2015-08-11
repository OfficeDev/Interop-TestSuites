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
    /// Traditional tests for "XVirusInfectedHeader" scenario.
    /// </summary>
    [TestClass]
    public class S01_XVirusInfectedHeader : TestSuiteBase
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

        #region XVirusInfectedHeader test cases

        /// <summary>
        /// This test case is used to verify the "X-Virus-Infected header" in the response to an HTTP GET request.
        /// </summary>
        [TestCategory("MSWDVMODUU"), TestMethod()]
        public void MSWDVMODUU_S01_TC01_XVirusInfectedHeader_Get()
        {
            // Prerequisites for this test case:
            //     The server has installed valid virus scanner software, and a fake virus file that can be detected 
            //     by the virus scanner software has existed in a URI under the server. 
            //     The URI should be set as the value property "Server_FakeVirusInfectedFileUri_Get" in PTF configuration file.

            // Get the URI of the fake virus file in the server from the property "Server_FakeVirusInfectedFileUri_Get".
            string requestUri = Common.GetConfigurationPropertyValue("Server_FakeVirusInfectedFileUri_Get", this.Site);

            // Construct the request headers.
            NameValueCollection headersCollection = new NameValueCollection();
            headersCollection.Add("Cache-Control", "no-cache");
            headersCollection.Add("Depth", "0");
            headersCollection.Add("Pragma", "no-cache");
            headersCollection.Add("ProtocolVersion", "HTTP/1.1");
            headersCollection.Add("Translate", "f");

            // Call HTTP GET method with the URI of the fake virus file in the server.
            WDVMODUUResponse response = null;
            HttpWebResponse httpWebResponse = null;
            try
            {
                response = this.Adapter.Get(requestUri, headersCollection);
                this.Site.Assert.Fail("Failed: The virus file should not be get successfully! \r\n The last request is:\r\n {0} The last response is: \r\n {1}", this.Adapter.LastRawRequest, this.Adapter.LastRawResponse);
            }
            catch (WebException webException)
            {
                this.Site.Assert.IsNotNull(webException.Response, "The 'Response' in the caught web exception should not be null!");
                httpWebResponse = (HttpWebResponse)webException.Response;
                if (httpWebResponse.StatusCode != HttpStatusCode.Conflict)
                {
                    // The expected web exception is "Conflict", if the caught web exception is not "Conflict", then throw the web exception to the framework.
                    throw;
                }

                response = new WDVMODUUResponse();
                response.ReserveResponseData(httpWebResponse, this.Site);
            }

            this.Site.Log.Add(
                TestTools.LogEntryKind.Comment,
                string.Format("In Method 'MSWDVMODUU_S01_TC01_XVirusInfectedHeader_Get', the request URI is {0}. In the response, the status code is '{1}'; the status description is '{2}'.", requestUri, response.StatusCode, response.StatusDescription));

            #region Verify Requirements

            // Confirm the server fails the request and respond with a message containing HTTP status code "409 CONFLICT", and then capture MS-WDVMODUU_R77.
            bool isResponseStatusCode409CONFLICT = false;
            if ((response.StatusCode == HttpStatusCode.Conflict) && (string.Compare(response.StatusDescription, "CONFLICT", true) == 0))
            {
                isResponseStatusCode409CONFLICT = true;
            }

            if (isResponseStatusCode409CONFLICT == false)
            {
                // Log some information to help users to know why the response does not include "409 CONFLICT". 
                string helpDebugInformation = @"The status code in the HTTP response is not ""409 CONFLICT"", make sure following conditions are matched, and see more help information under section 1.2.1 in the test suite specification: \r\n ";
                helpDebugInformation += @"1. A valid virus scanner software has been installed in the protocol server, and the virus scanner software must be active when a document is downloaded from the protocol server. \r\n";
                helpDebugInformation += @"2. The fake virus infected file that can be detected by the valid virus scanner software exists in the URI under the protocol server, as specified by the PTF property ""Server_FakeVirusInfectedFileUri_Get"".";
                this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, helpDebugInformation);
            }

            this.Site.CaptureRequirementIfIsTrue(
                isResponseStatusCode409CONFLICT,
                77,
                @"[In X-Virus-Infected Header] If this [X-Virus-Infected] header is returned by a WebDAV server in response to an HTTP [PUT or a] GET request, the server MUST fail the request and respond with a message containing HTTP status code ""409 CONFLICT"".");

            // Confirm the server does not return the fake virus file in the response with "409 CONFLICT" error condition, and then capture MS-WDVMODUU_R78.
            bool isContentBodyEmpty = false;
            if (response.ContentLength == 0)
            {
                isContentBodyEmpty = true;
            }

            this.Site.CaptureRequirementIfIsTrue(
                isResponseStatusCode409CONFLICT && isContentBodyEmpty,
                78,
                @"[In X-Virus-Infected Header] [If X-Virus-Infected Header is returned] The server MUST NOT return the infected file to the client following a GET request ""409 CONFLICT"" error condition.");

            // Confirm the server returns the X-Virus-Infected header in the response, and then capture MS-WDVMODUU_R74.
            bool isXVirusInfectedHeaderReturned = false;
            foreach (string headerName in response.HttpHeaders.AllKeys)
            {
                if (string.Compare(headerName, "X-Virus-Infected", true) == 0)
                {
                    isXVirusInfectedHeaderReturned = true;
                    break;
                }
            }

            this.Site.CaptureRequirementIfIsTrue(
               isXVirusInfectedHeaderReturned,
               74,
               @"[In X-Virus-Infected Header] A WebDAV server returns the X-Virus-Infected header in response to an HTTP GET [or a PUT] request to indicate that the requested file is infected with a virus.");

            // Get the X-Virus-Infected header and its value.
            string[] headerAndValues = response.HttpHeaders.ToString().Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            string virusInfectedHeaderAndValue = string.Empty;
            foreach (string headerAndValue in headerAndValues)
            {
                if (headerAndValue.ToLower().IndexOf("x-virus-infected") == 0)
                {
                    virusInfectedHeaderAndValue = headerAndValue;
                    break;
                }
            }

            this.Site.Assert.IsTrue(virusInfectedHeaderAndValue != string.Empty, "The X-Virus-Infected header and its value should not be empty string!");

            // Define a regular expression pattern for the syntax: "x-virus-infected" ":" Virus-Name Virus-Name = 1*TEXT
            // "x-virus-infected" is matching for the header name: "x-virus-infected"
            // ":" is matching the colon which must exist. 
            // "\s*" is matching the leading white spaces
            // ".*" is matching the name characters other than \r\n 
            string virusHeaderPattern = string.Format(@"x-virus-infected:\s*.*");

            // Check the format of X-Virus-Infected header and its value in the response.
            Site.Log.Add(TestTools.LogEntryKind.Comment, "The X-Virus-Infected header and its value is {0}", virusInfectedHeaderAndValue);
            bool isRightXVirusHeader = Regex.IsMatch(virusInfectedHeaderAndValue, virusHeaderPattern, RegexOptions.IgnoreCase);
            Site.Assert.IsTrue(isRightXVirusHeader, "The format of X-Virus-Infected header should be correct.");

            // Capture MS-WDVMODUU_R7, MS-WDVMODUU_R4, and MS-WDVMODUU_R5, if the format of X-Virus-Infected header and its value is correct.
            this.Site.CaptureRequirementIfIsTrue(
               isXVirusInfectedHeaderReturned && isRightXVirusHeader,
               4,
               @"[In MODUU Extension Headers] The extension headers in this protocol conform to the form and behavior of other custom HTTP 1.1 headers, as specified in [RFC2616] section 4.2.");

            this.Site.CaptureRequirementIfIsTrue(
               isXVirusInfectedHeaderReturned && isRightXVirusHeader,
               5,
               @"[In MODUU Extension Headers] They [The extension headers] are consistent with the WebDAV verbs and headers, as specified in [RFC2518] sections 8 and 9.");

            this.Site.CaptureRequirementIfIsTrue(
                isXVirusInfectedHeaderReturned && isRightXVirusHeader,
                7,
                @"[In X-Virus-Infected Header] If returned, the X-Virus-Infected header MUST take the following form:
X-Virus-Infected Header = ""x-virus-infected"" "":"" Virus-Name
Virus-Name = 1*TEXT");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify the "X-Virus-Infected header" in the response to an HTTP PUT request.
        /// </summary>
        [TestCategory("MSWDVMODUU"), TestMethod()]
        public void MSWDVMODUU_S01_TC02_XVirusInfectedHeader_Put()
        {
            // Prerequisites for this test case:
            //     The server has installed valid virus scanner software.
            //     And in the client test environment, under the local output path there is a fake virus file that can be detected by the valid virus scanner software. 
            //     The fake virus file name should be set as the value of property "Client_FakeVirusInfectedFileName" in PTF configuration file.

            // Get the URI of the fake virus file from the property "Server_FakeVirusInfectedFileUri_Put", 
            // the URI will be used as Request-URI in the HTTP PUT method.
            string requestUri = Common.GetConfigurationPropertyValue("Server_FakeVirusInfectedFileUri_Put", this.Site);

            // The fake virus file is under the local output path, and its name is the value of property "Client_FakeVirusInfectedFileName" in PTF configuration file.
            string fakeVirusInfectedFileName = Common.GetConfigurationPropertyValue("Client_FakeVirusInfectedFileName", this.Site);

            // Assert the fake virus infected file is existed in the local folder.
            if (File.Exists(fakeVirusInfectedFileName) == false)
            {
                this.Site.Assert.Fail("The file '{0}' was not found in local output path(such as: <Solution Directory>\\TestSuite\\Resources\\), prepare a fake virus infected file under the local output path, and make sure the name of the file is same as the value of PTF property \"Client_FakeVirusInfectedFileName\".", fakeVirusInfectedFileName);
            }

            // Read the contents of fake virus file in the client. 
            byte[] bytes = GetLocalFileContent(fakeVirusInfectedFileName);

            // Construct the request headers.
            NameValueCollection headersCollection = new NameValueCollection();
            headersCollection.Add("Cache-Control", "no-cache");
            headersCollection.Add("Pragma", "no-cache");
            headersCollection.Add("ProtocolVersion", "HTTP/1.1");

            // Call HTTP PUT method to upload the fake virus file to the server.
            WDVMODUUResponse response = null;
            HttpWebResponse httpWebResponse = null;
            try
            {
                response = this.Adapter.Put(requestUri, bytes, headersCollection);
                this.ArrayListForDeleteFile.Add(requestUri);
                this.Site.Assert.Fail("Failed: The virus file should not be put successfully! \r\n The last request is:\r\n {0} The last response is: \r\n {1}", this.Adapter.LastRawRequest, this.Adapter.LastRawResponse);
            }
            catch (WebException webException)
            {
                this.Site.Assert.IsNotNull(webException.Response, "The 'Response' in the caught web exception should not be null!");
                httpWebResponse = (HttpWebResponse)webException.Response;
                if (httpWebResponse.StatusCode != HttpStatusCode.Conflict)
                {
                    // The expected web exception is "Conflict", if the caught web exception is not "Conflict", then throw the web exception to the framework.
                    throw;
                }

                response = new WDVMODUUResponse();
                response.ReserveResponseData(httpWebResponse, this.Site);
            }

            this.Site.Log.Add(
                TestTools.LogEntryKind.Comment,
                string.Format("In Method 'MSWDVMODUU_S01_TC01_XVirusInfectedHeader_Put', the request URI is {0}. In the response, the status code is '{1}'; the status description is '{2}'.", requestUri, response.StatusCode, response.StatusDescription));

            #region Verify Requirements

            // Confirm the server fails the request and respond with a message containing HTTP status code "409 CONFLICT", and then capture MS-WDVMODUU_R76.
            bool isResponseStatusCode409CONFLICT = false;
            if ((response.StatusCode == HttpStatusCode.Conflict)
                && (string.Compare(response.StatusDescription, "CONFLICT", true) == 0))
            {
                isResponseStatusCode409CONFLICT = true;
            }

            if (isResponseStatusCode409CONFLICT == false)
            {
                // Log some information to help users to know why the response does not include "409 CONFLICT". 
                string helpDebugInformation = @"The status code in the HTTP response is not ""409 CONFLICT"", make sure following conditions are matched, and see more help information under section 1.2.1 in the test suite specification: \r\n ";
                helpDebugInformation += @"1. A valid virus scanner software has been installed in the protocol server, and virus scanner software must be active when a document is uploaded to the protocol server. \r\n";
                helpDebugInformation += @"2. The fake virus infected file " + fakeVirusInfectedFileName + " can be detected by the valid virus scanner software in the protocol server. \r\n";
                helpDebugInformation += @"3. The URI in the PTF property ""Server_FakeVirusInfectedFileUri_Put"" is an addressable path in the protocol server that the test case has permission to upload a file.";
                this.Site.Log.Add(TestTools.LogEntryKind.TestFailed, helpDebugInformation);
            }

            this.Site.CaptureRequirementIfIsTrue(
                isResponseStatusCode409CONFLICT,
                76,
                @"[In X-Virus-Infected Header] If this [X-Virus-Infected ] header is returned by a WebDAV server in response to an HTTP PUT[ or a GET] request, the server MUST fail the request and respond with a message containing HTTP status code ""409 CONFLICT"".");

            // Confirm the server returns the X-Virus-Infected header in the response, and then capture MS-WDVMODUU_R75.
            bool isXVirusInfectedHeaderReturned = false;
            foreach (string headerName in response.HttpHeaders.AllKeys)
            {
                if (string.Compare(headerName, "X-Virus-Infected", true) == 0)
                {
                    isXVirusInfectedHeaderReturned = true;
                    break;
                }
            }

            this.Site.CaptureRequirementIfIsTrue(
                isXVirusInfectedHeaderReturned,
                75,
                @"[In X-Virus-Infected Header] A WebDAV server returns the X-Virus-Infected header in response to an HTTP [GET or a ]PUT request to indicate that the requested file is infected with a virus.");

            // Get the X-Virus-Infected header and its value into a string.
            string[] headerAndValues = response.HttpHeaders.ToString().Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            string virusInfectedHeaderAndValue = string.Empty;
            foreach (string headerAndValue in headerAndValues)
            {
                if (headerAndValue.ToLower().IndexOf("x-virus-infected") == 0)
                {
                    virusInfectedHeaderAndValue = headerAndValue;
                    break;
                }
            }

            this.Site.Assert.IsTrue(virusInfectedHeaderAndValue != string.Empty, "The X-Virus-Infected header and its value should not be empty string!");

            // Define a regular expression pattern for the syntax: "x-virus-infected" ":" Virus-Name Virus-Name = 1*TEXT
            // "x-virus-infected" is matching for the header name: "x-virus-infected"
            // ":" is matching the colon which must exist. 
            // "\s*" is matching the leading white spaces
            // ".*" is matching the name characters other than \r\n 
            string virusHeaderPattern = string.Format(@"^x-virus-infected:\s*.*");

            // Check the format of X-Virus-Infected header and its value in the response.
            Site.Log.Add(TestTools.LogEntryKind.Comment, "The X-Virus-Infected header and its value is {0}", virusInfectedHeaderAndValue);
            bool isRightXVirusHeader = Regex.IsMatch(virusInfectedHeaderAndValue, virusHeaderPattern, RegexOptions.IgnoreCase);
            Site.Assert.IsTrue(isRightXVirusHeader, "The format of X-Virus-Infected header should be correct.");

            // Capture MS-WDVMODUU_R7, MS-WDVMODUU_R4, and MS-WDVMODUU_R5, if the format of X-Virus-Infected header and its value is correct.
            this.Site.CaptureRequirementIfIsTrue(
               isXVirusInfectedHeaderReturned && isRightXVirusHeader,
               4,
               @"[In MODUU Extension Headers] The extension headers in this protocol conform to the form and behavior of other custom HTTP 1.1 headers, as specified in [RFC2616] section 4.2.");

            this.Site.CaptureRequirementIfIsTrue(
               isXVirusInfectedHeaderReturned && isRightXVirusHeader,
               5,
               @"[In MODUU Extension Headers] They [The extension headers] are consistent with the WebDAV verbs and headers, as specified in [RFC2518] sections 8 and 9.");

            this.Site.CaptureRequirementIfIsTrue(
                isXVirusInfectedHeaderReturned && isRightXVirusHeader,
                7,
                @"[In X-Virus-Infected Header] If returned, the X-Virus-Infected header MUST take the following form:
X-Virus-Infected Header = ""x-virus-infected"" "":"" Virus-Name
Virus-Name = 1*TEXT");

            #endregion
        }
        #endregion
    }
}