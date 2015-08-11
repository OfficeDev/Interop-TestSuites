namespace Microsoft.Protocols.TestSuites.MS_ASHTTP
{
    using System.Collections.Generic;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test HTTP POST with negative response.
    /// </summary>
    [TestClass]
    public class S02_HTTPPOSTNegative : TestSuiteBase
    {
        #region Class initialize and clean up
        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test cases
        /// <summary>
        /// This test case is intended to validate the 400 Bad Request HTTP POST status code.
        /// </summary>
        [TestCategory("MSASHTTP"), TestMethod()]
        public void MSASHTTP_S02_TC01_Verify400StatusCode()
        {
            #region Call ConfigureRequestPrefixFields to set the ActiveSyncProtocolVersion to 191.
            IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();
            requestPrefix.Add(HTTPPOSTRequestPrefixField.ActiveSyncProtocolVersion, "191");
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
            #endregion

            #region Synchronize the collection hierarchy via FolderSync command.
            string statusCode = string.Empty;
            HttpWebResponse httpWebResponse = null;
            string folderSyncRequest = Common.CreateFolderSyncRequest("0").GetRequestDataSerializedXML();

            try
            {
                // Call HTTP POST using FolderSync command to synchronize the collection hierarchy.
                this.HTTPAdapter.HTTPPOST(CommandName.FolderSync, null, folderSyncRequest);
                Site.Assert.Fail("The server should throw 400 Bad Request exception.");
            }
            catch (WebException exception)
            {
                Site.Log.Add(LogEntryKind.Debug, "Caught exception message is:" + exception.Message.ToString());
                httpWebResponse = (HttpWebResponse)exception.Response;
                statusCode = TestSuiteHelper.GetStatusCodeFromException(exception);
            }

            bool is400StatusCode = (httpWebResponse.StatusCode == HttpStatusCode.BadRequest) && statusCode.Equals("400");
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R176");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R176
            // If the caught status code is 400 Bad Request, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                is400StatusCode,
                176,
                @"[In Status Line] [Status code] 400 Bad Request [is described as] the request could not be understood by the server due to malformed syntax.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R188");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R188
            // If the caught status code is 400 Bad Request, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                is400StatusCode,
                188,
                @"[In Status Line] In the case of other malformed requests, the server returns status code 400.");

            #region Synchronize the collection hierarchy again via FolderSync command.
            try
            {
                // Call HTTP POST using FolderSync command to synchronize the collection hierarchy.
                this.HTTPAdapter.HTTPPOST(CommandName.FolderSync, null, folderSyncRequest);
                Site.Assert.Fail("The server should throw 400 Bad Request exception.");
            }
            catch (WebException exception)
            {
                Site.Log.Add(LogEntryKind.Debug, "Caught exception message is:" + exception.Message.ToString());
                httpWebResponse = (HttpWebResponse)exception.Response;
                statusCode = TestSuiteHelper.GetStatusCodeFromException(exception);
            }
            finally
            {
                // Reset the ActiveSyncProtocolVersion.
                string activeSyncProtocolVersion = Common.ConvertActiveSyncProtocolVersion(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), Site);

                requestPrefix[HTTPPOSTRequestPrefixField.ActiveSyncProtocolVersion] = activeSyncProtocolVersion;
                this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
            }

            is400StatusCode = (httpWebResponse.StatusCode == HttpStatusCode.BadRequest) && statusCode.Equals("400");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R177");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R177
            // If the caught status code is 400 Bad Request, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                is400StatusCode,
                177,
                @"[In Status Line] [Status code] 400 Bad Request [is described as] if the client repeats the request without modifications, then the same error [400 Bad Request] occurs.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the 401 Unauthorized HTTP POST status code.
        /// </summary>
        [TestCategory("MSASHTTP"), TestMethod()]
        public void MSASHTTP_S02_TC02_Verify401StatusCode()
        {
            #region Call FolderSync command with no credential.
            HttpWebResponse httpWebResponse = null;
            string statusCode = string.Empty;
            try
            {
                // Switch the user to an invalid user.
                this.InvalidUserInformation.UserName = null;
                this.InvalidUserInformation.UserPassword = null;
                this.InvalidUserInformation.UserDomain = null;
                this.SwitchUser(this.InvalidUserInformation, false);

                // Call HTTP POST using FolderSync command  to synchronize the collection hierarchy.
                this.CallFolderSyncCommand();
                Site.Assert.Fail("The server should throw 401 Unauthorized exception.");
            }
            catch (WebException exception)
            {
                Site.Log.Add(LogEntryKind.Debug, "Caught exception message is:" + exception.Message.ToString());
                httpWebResponse = (HttpWebResponse)exception.Response;
                statusCode = TestSuiteHelper.GetStatusCodeFromException(exception);
            }
            finally
            {
                // Reset the user credential.
                this.SwitchUser(this.UserOneInformation, false);
            }

            bool is401StatusCode = (httpWebResponse.StatusCode == HttpStatusCode.Unauthorized) && statusCode.Equals("401");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R178");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R178
            // If the caught status code is 401 Unauthorized, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                is401StatusCode,
                178,
                @"[In Status Line] [Status code] 401 Unauthorized [is described as] the resource requires authorization [or authorization was refused].");
            #endregion

            #region Call FolderSync command with invalid credential.
            try
            {
                // Switch the user to an invalid user.
                this.InvalidUserInformation.UserName = this.UserOneInformation.UserName;
                this.InvalidUserInformation.UserPassword = this.UserOneInformation.UserPassword + "InvalidPassword";
                this.InvalidUserInformation.UserDomain = this.UserOneInformation.UserDomain;
                this.SwitchUser(this.InvalidUserInformation, false);

                // Call HTTP POST using FolderSync command to synchronize the collection hierarchy.
                this.CallFolderSyncCommand();
                Site.Assert.Fail("The server should throw 401 Unauthorized exception.");
            }
            catch (WebException exception)
            {
                Site.Log.Add(LogEntryKind.Debug, "Caught exception message is:" + exception.Message.ToString());
                httpWebResponse = (HttpWebResponse)exception.Response;
                statusCode = TestSuiteHelper.GetStatusCodeFromException(exception);
            }
            finally
            {
                // Reset the user credential.
                this.SwitchUser(this.UserOneInformation, false);
            }

            is401StatusCode = (httpWebResponse.StatusCode == HttpStatusCode.Unauthorized) && statusCode.Equals("401");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R523");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R523
            // If the caught status code is 401 Unauthorized, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                is401StatusCode,
                523,
                @"[In Status Line] [Status code] 401 Unauthorized [is described as] the [resource requires authorization or] authorization was refused.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the 403 Forbidden HTTP POST status code.
        /// </summary>
        [TestCategory("MSASHTTP"), TestMethod()]
        public void MSASHTTP_S02_TC03_Verify403StatusCode()
        {
            Site.Assume.AreEqual<string>("HTTPS", Common.GetConfigurationPropertyValue("TransportType", this.Site), "The status code 403 is verified in HTTPS transport type.");

            #region Call SUT Control adapter to enable the SSL setting.
            string sutComputerName = Common.GetConfigurationPropertyValue("SutComputerName", Site);
            string userName = Common.GetConfigurationPropertyValue("User4Name", Site);
            string password = Common.GetConfigurationPropertyValue("User4Password", Site);

            // Use the user who is in Administrator group to enable the SSL setting.
            bool isSSLUpdated = this.HTTPSUTControlAdapter.ConfigureSSLSetting(sutComputerName, userName, password, Common.GetConfigurationPropertyValue("Domain", Site), true);
            Site.Assert.IsTrue(isSSLUpdated, "The SSL setting of protocol web service should be enabled.");
            #endregion

            #region Call FolderSync command with HTTP transport type.
            string statusCode = string.Empty;
            HttpWebResponse httpWebResponse = null;
            string folderSyncRequestBody = Common.CreateFolderSyncRequest("0").GetRequestDataSerializedXML();
            Dictionary<HTTPPOSTRequestPrefixField, string> requestPrefixFields = new Dictionary<HTTPPOSTRequestPrefixField, string>();
            try
            {
                // Change the prefix of URI to make it disagree with the configuration of server.
                requestPrefixFields.Add(HTTPPOSTRequestPrefixField.PrefixOfURI, ProtocolTransportType.HTTP.ToString());
                this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);

                // Call HTTP POST using FolderSync command to synchronize the collection hierarchy.
                this.HTTPAdapter.HTTPPOST(CommandName.FolderSync, null, folderSyncRequestBody);
                Site.Assert.Fail("The server should throw 403 Forbidden exception.");
            }
            catch (WebException exception)
            {
                Site.Log.Add(LogEntryKind.Debug, "Caught exception message is:" + exception.Message.ToString());
                httpWebResponse = (HttpWebResponse)exception.Response;
                statusCode = TestSuiteHelper.GetStatusCodeFromException(exception);
            }
            finally
            {
                // Reset the PrefixOfURI.
                requestPrefixFields[HTTPPOSTRequestPrefixField.PrefixOfURI] = Common.GetConfigurationPropertyValue("TransportType", this.Site);
                this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);

                // Reset the SSL setting.
                isSSLUpdated = this.HTTPSUTControlAdapter.ConfigureSSLSetting(sutComputerName, userName, password, Common.GetConfigurationPropertyValue("Domain", this.Site), false);
                Site.Assert.IsTrue(isSSLUpdated, "The SSL setting of protocol web service should be disabled.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R180");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R180
            // If the caught status code is 403 Forbidden, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                (httpWebResponse.StatusCode == HttpStatusCode.Forbidden) && statusCode.Equals("403"),
                180,
                @"[In Status Line] [Status code] 403 Forbidden [is described as] the user is not enabled for ActiveSync synchronization.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the 404 Not Found HTTP POST status code.
        /// </summary>
        [TestCategory("MSASHTTP"), TestMethod()]
        public void MSASHTTP_S02_TC04_Verify404StatusCode()
        {
            #region Call FolderSync command with invalid URI.
            string statusCode = string.Empty;
            HttpWebResponse httpWebResponse = null;
            string folderSyncRequestBody = Common.CreateFolderSyncRequest("0").GetRequestDataSerializedXML();
            string sutComputerName = Common.GetConfigurationPropertyValue("SutComputerName", Site);
            Dictionary<HTTPPOSTRequestPrefixField, string> requestPrefixFields = new Dictionary<HTTPPOSTRequestPrefixField, string>();
            try
            {
                // Change the Host property to make the URI unsupported.
                requestPrefixFields.Add(HTTPPOSTRequestPrefixField.Host, sutComputerName + "/NotSupport");
                this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);

                // Call HTTP POST using FolderSync command to synchronize the collection hierarchy.
                this.HTTPAdapter.HTTPPOST(CommandName.FolderSync, null, folderSyncRequestBody);
                Site.Assert.Fail("The server should throw 404 Not Found exception.");
            }
            catch (WebException exception)
            {
                Site.Log.Add(LogEntryKind.Debug, "Caught exception message is:" + exception.Message.ToString());
                httpWebResponse = (HttpWebResponse)exception.Response;
                statusCode = TestSuiteHelper.GetStatusCodeFromException(exception);
            }
            finally
            {
                // Reset the Host.
                requestPrefixFields[HTTPPOSTRequestPrefixField.Host] = sutComputerName;
                this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);
            }

            bool is404StatusCode = (httpWebResponse.StatusCode == HttpStatusCode.NotFound) && statusCode.Equals("404");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R182");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R182
            // If the caught status code is 404 Not Found, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                is404StatusCode,
                182,
                @"[In Status Line] [Status code] 404 Not Found [is described as] the specified URI could not be found [or the server is not a valid server with ActiveSync].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R441");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R441
            // If the caught status code is 404 Not Found, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                is404StatusCode,
                441,
                @"[In Status Line] [Status code] 404 Not Found [is described as] [the specified URI could not be found or] the server is not a valid server with ActiveSync.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the 501 Not Implemented HTTP POST status code.
        /// </summary>
        [TestCategory("MSASHTTP"), TestMethod()]
        public void MSASHTTP_S02_TC05_Verify501StatusCode()
        {
            #region Call a command which is not supported by server.
            string statusCode = string.Empty;
            HttpWebResponse httpWebResponse = null;
            string folderSyncRequestBody = Common.CreateFolderSyncRequest("0").GetRequestDataSerializedXML();
            Dictionary<HTTPPOSTRequestPrefixField, string> requestPrefixFields = new Dictionary<HTTPPOSTRequestPrefixField, string>
            {
                {
                    HTTPPOSTRequestPrefixField.QueryValueType, QueryValueType.PlainText.ToString()
                }
            };
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);

            try
            {
                // Call HTTP POST using a NotExist command.
                this.HTTPAdapter.HTTPPOST(CommandName.NotExist, null, folderSyncRequestBody);
                Site.Assert.Fail("The server should throw 501 Not Implemented exception.");
            }
            catch (WebException exception)
            {
                Site.Log.Add(LogEntryKind.Debug, "Caught exception message is:" + exception.Message.ToString());
                httpWebResponse = (HttpWebResponse)exception.Response;
                statusCode = TestSuiteHelper.GetStatusCodeFromException(exception);
            }
            finally
            {
                requestPrefixFields[HTTPPOSTRequestPrefixField.QueryValueType] = Common.GetConfigurationPropertyValue("HeaderEncodingType", this.Site);
                this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);
            }

            bool is501StatusCode = (httpWebResponse.StatusCode == HttpStatusCode.NotImplemented) && statusCode.Equals("501");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R186");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R186
            // If the caught status code is 501 Not Implemented, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                is501StatusCode,
                186,
                @"[In Status Line] [Status code] 501 Not Implemented [is described as] the server does not support the functionality that is required to fulfill the request.");

            if (Common.IsRequirementEnabled(465, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R465");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R465
                // If the caught status code is 501 Not Implemented, this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    is501StatusCode,
                    465,
                    @"[In Appendix A: Product Behavior] Implementation does return 501 Not Implemented status code when implementation does not recognize the request method [or is not able to support it for any resource].(Exchange 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion
    }
}