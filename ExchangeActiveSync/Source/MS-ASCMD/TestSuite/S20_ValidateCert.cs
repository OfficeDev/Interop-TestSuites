namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test the ValidateCert command.
    /// </summary>
    [TestClass]
    public class S20_ValidateCert : TestSuiteBase
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

        #region Test Cases

        /// <summary>
        /// This test case is used to verify the status value will be 1, when certificate validation is successfully completed for the ValidateCert command.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S20_TC01_ValidateCert_Success()
        {
            #region Switch to User9 mail account, the inbox had received a S/MIME mail.

            this.SwitchUser(this.User9Information);

            #endregion

            #region Call method Sync to synchronize changes of Inbox folder in User9's mailbox between the client and the server, and get the body content of the email item.

            string emailSubject = Common.GetConfigurationPropertyValue("MIMEMailSubject", this.Site);

            Request.BodyPreference bodyPreference = new Request.BodyPreference
            {
                AllOrNone = false,
                AllOrNoneSpecified = true,
                TruncationSize = 4294967295,
                TruncationSizeSpecified = true,
                Type = 4
            };

            Request.Options option = new Request.Options
            {
                Items = new object[] { (byte)2, bodyPreference, (byte)8 },
                ItemsElementName = new Request.ItemsChoiceType1[]
                {
                    Request.ItemsChoiceType1.MIMESupport,
                    Request.ItemsChoiceType1.BodyPreference, Request.ItemsChoiceType1.MIMETruncation
                }
            };

            SyncResponse syncResponse = this.CheckEmail(this.User9Information.InboxCollectionId, emailSubject, new Request.Options[] { option });
            Response.Body mailBody = null;
            Response.SyncCollectionsCollectionCommandsAddApplicationData applicationData = TestSuiteBase.GetAddApplicationData(syncResponse, Response.ItemsChoiceType8.Subject1, emailSubject);
            for (int i = 0; i < applicationData.ItemsElementName.Length; i++)
            {
                if (applicationData.ItemsElementName[i] == Response.ItemsChoiceType8.Body)
                {
                    mailBody = applicationData.Items[i] as Response.Body;
                    break;
                }
            }

            Site.Assert.IsNotNull(mailBody, "The body of the received email should not be null.");

            string specifiedString = "MIME-Version: 1.0";
            string body = mailBody.Data.Substring(
                mailBody.Data.IndexOf(specifiedString, StringComparison.CurrentCultureIgnoreCase) + specifiedString.Length);
            body = body.Replace("\r", string.Empty);
            body = body.Replace("\n", string.Empty);

            Request.ValidateCert validateCert = new Request.ValidateCert
            {
                CheckCrl = "1",
                Certificates = new byte[][] { System.Text.Encoding.Default.GetBytes(body) }
            };

            ValidateCertRequest validateRequest = new ValidateCertRequest { RequestData = validateCert };

            ValidateCertResponse validateResponse = this.CMDAdapter.ValidateCert(validateRequest);

            XmlNodeList status = this.GetValidateCertStatusCode(validateResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4474");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4474
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                status[0].InnerText,
                4474,
                @"[In Status(ValidateCert)] [When the scope is Global], [the cause of the status value 1 is] Server successfully completed command.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5387");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5387
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                status[0].InnerText,
                5387,
                @"[In Status(ValidateCert)] A value of 1 indicates success.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the status value is 3, when the signature in the certificate is invalid.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S20_TC02_ValidateCert_InvalidSignature()
        {
            #region User calls ValidateCert command to verify the certificate with invalid signature
            Request.ValidateCert validateCert = new Request.ValidateCert
            {
                CheckCrl = "1",
                Certificates = new byte[][] { Convert.FromBase64String("TUlJQ1lqQ0NBY3VnQXdJQkFnSVVZR3M4alpiWDBWeGpPYnU0bncwhQ==") }
            };

            ValidateCertRequest validateRequest = new ValidateCertRequest { RequestData = validateCert };

            ValidateCertResponse validateResponse = this.CMDAdapter.ValidateCert(validateRequest);

            XmlNodeList status = this.GetValidateCertStatusCode(validateResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4479");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4479
            Site.CaptureRequirementIfAreEqual<string>(
                "3",
                status[0].InnerText,
                4479,
                @"[In Status(ValidateCert)] [When the scope is Item], [the cause of the status value 3 is] The signature in the certificate is invalid.");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify the status value is 2, when ValidateCert command request contains more than 100 Certificate elements.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S20_TC03_ValidateCert_LimitingSize()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5671, this.Site), "[In Appendix A: Product Behavior] Implementation does limit the number of elements in command requests and return the specified error if the limit is exceeded. (<118> Section 3.1.5.8: Update Rollup 6 for Exchange 2010 SP2 and Exchange 2013 do limit the number of elements in command requests.)");
            
            #region User calls ValidateCert command to verify the number of certificate element is more than 100
            Request.ValidateCert validateCert = new Request.ValidateCert
            {
                CheckCrl = "1",
                Certificates = new byte[101][]
            };

            for (int i = 0; i < validateCert.Certificates.Length; i++)
            {
                validateCert.Certificates[i] = Convert.FromBase64String("TUlJQ1lqQ0NBY3VnQXdJQkFnSVVZR3M4alpiWDBWeGpPYnU0bncwhQ==");
            }

            Site.Assert.IsTrue(validateCert.Certificates.Length > 100, "The number of certificate element is more than 100.");

            ValidateCertRequest validateRequest = new ValidateCertRequest { RequestData = validateCert };

            ValidateCertResponse validateResponse = this.CMDAdapter.ValidateCert(validateRequest);

            XmlNodeList status = this.GetValidateCertStatusCode(validateResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5662");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5662
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                status[0].InnerText,
                5662,
                @"[In Limiting Size of Command Requests] In ValidateCert (section 2.2.2.20) command request, when the limit value of Certificate element is bigger than 100 (minimum 1, maximum 2,147,483,647), the error returned by server is Status element (section 2.2.3.162.17) value of 2.");

            #endregion
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Get ValidateCert response status which returned by the ValidateCert operation.
        /// </summary>
        /// <param name="response">The ValidateCert response data</param>
        /// <returns>The Status code</returns>
        private XmlNodeList GetValidateCertStatusCode(ValidateCertResponse response)
        {
            string xmlResponse = response.ResponseDataXML;
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlResponse);
            XmlNamespaceManager xmlNameSpaceManager = new XmlNamespaceManager(doc.NameTable);
            xmlNameSpaceManager.AddNamespace("e", "ValidateCert");
            XmlNodeList status = doc.SelectNodes("/e:ValidateCert/e:Certificate/e:Status", xmlNameSpaceManager);

            if (status != null && status.Count == 0)
            {
                status = doc.SelectNodes("/e:ValidateCert/e:Status", xmlNameSpaceManager);
            }

            Site.Assert.IsTrue(status.Count != 0, "The server response should contain a Status element.");
            return status;
        }

        #endregion
    }
}