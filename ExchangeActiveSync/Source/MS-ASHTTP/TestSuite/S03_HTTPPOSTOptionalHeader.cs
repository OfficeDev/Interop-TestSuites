namespace Microsoft.Protocols.TestSuites.MS_ASHTTP
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test optional headers of HTTP POST command.
    /// </summary>
    [TestClass]
    public class S03_HTTPPOSTOptionalHeader : TestSuiteBase
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
        /// This test case is intended to validate the MS-ASAcceptMultiPart optional header in HTTP POST request.
        /// </summary>
        [TestCategory("MSASHTTP"), TestMethod()]
        public void MSASHTTP_S03_TC01_SetASAcceptMultiPartRequestHeader()
        {
            #region Call SendMail command to send email to User2.
            // Call ConfigureRequestPrefixFields to set the QueryValueType to PlainText.
            IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();
            requestPrefix.Add(HTTPPOSTRequestPrefixField.QueryValueType, QueryValueType.PlainText.ToString());
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

            // Call FolderSync command to synchronize the collection hierarchy.
            this.CallFolderSyncCommand();

            string sendMailSubject = Common.GenerateResourceName(Site, "SendMail");
            string userOneMailboxAddress = Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain);
            string userTwoMailboxAddress = Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain);

            // Call SendMail command.
            this.CallSendMailCommand(userOneMailboxAddress, userTwoMailboxAddress, sendMailSubject, null);
            #endregion

            #region Get the received email.
            // Call ConfigureRequestPrefixFields to switch the credential to User2 and synchronize the collection hierarchy.
            this.SwitchUser(this.UserTwoInformation, true);
            this.AddCreatedItemToCollection("User2", this.UserTwoInformation.InboxCollectionId, sendMailSubject);

            // Call Sync command to get the received email.
            string itemServerId = this.LoopToSyncItem(this.UserTwoInformation.InboxCollectionId, sendMailSubject, true);
            #endregion

            #region Call ItemOperation command with setting MS-ASAcceptMultiPart header to "T".
            // Call ConfigureRequestPrefixFields to set MS-ASAcceptMultiPart header to "T".
            requestPrefix.Add(HTTPPOSTRequestPrefixField.AcceptMultiPart, "T");
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

            // Call ItemOperation command to fetch the received email.
            SendStringResponse itemOperationResponse = this.CallItemOperationsCommand(this.UserTwoInformation.InboxCollectionId, itemServerId, false);
            Site.Assert.IsNotNull(itemOperationResponse.Headers["Content-Type"], "The Content-Type header should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R154");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R154
            // The content is in multipart, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                "application/vnd.ms-sync.multipart",
                itemOperationResponse.Headers["Content-Type"],
                154,
                @"[In MS-ASAcceptMultiPart] If this [MS-ASAcceptMultiPart] header is present and the value is 'T', the client is requesting that the server return content in multipart format.");
            #endregion

            #region Call ItemOperation command with setting MS-ASAcceptMultiPart header to "F".
            // Call ConfigureRequestPrefixFields to change the MS-ASAcceptMultiPart header to "F".
            requestPrefix[HTTPPOSTRequestPrefixField.AcceptMultiPart] = "F";
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

            // Call ItemOperation command to fetch the received email.
            itemOperationResponse = this.CallItemOperationsCommand(this.UserTwoInformation.InboxCollectionId, itemServerId, false);
            Site.Assert.IsNotNull(itemOperationResponse.Headers["Content-Type"], "The Content-Type header should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R440");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R440
            // The content is not in multipart, so this requirement can be captured.
            Site.CaptureRequirementIfAreNotEqual<string>(
                "application/vnd.ms-sync.multipart",
                itemOperationResponse.Headers["Content-Type"],
                440,
                @"[In MS-ASAcceptMultiPart] If the [MS-ASAcceptMultiPart] header [is not present, or] is present and set to 'F', the client is requesting that the server return content in inline format.");
            #endregion

            #region Call ItemOperation command with setting MS-ASAcceptMultiPart header to null.
            // Call ConfigureRequestPrefixFields to change the MS-ASAcceptMultiPart header to null.
            requestPrefix[HTTPPOSTRequestPrefixField.AcceptMultiPart] = null;
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

            // Call ItemOperation command to fetch the received email.
            itemOperationResponse = this.CallItemOperationsCommand(this.UserTwoInformation.InboxCollectionId, itemServerId, false);
            Site.Assert.IsNotNull(itemOperationResponse.Headers["Content-Type"], "The Content-Type header should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R155");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R155
            // The content is not in multipart, so this requirement can be captured.
            Site.CaptureRequirementIfAreNotEqual<string>(
                "application/vnd.ms-sync.multipart",
                itemOperationResponse.Headers["Content-Type"],
                155,
                @"[In MS-ASAcceptMultiPart] If the [MS-ASAcceptMultiPart] header is not present [, or is present and set to 'F'], the client is requesting that the server return content in inline format.");
            #endregion

            #region Reset the query value type and credential.
            requestPrefix[HTTPPOSTRequestPrefixField.QueryValueType] = Common.GetConfigurationPropertyValue("HeaderEncodingType", this.Site);
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
            this.SwitchUser(this.UserOneInformation, false);
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the User-Agent optional header in HTTP POST request.
        /// </summary>
        [TestCategory("MSASHTTP"), TestMethod()]
        public void MSASHTTP_S03_TC02_SetUserAgentRequestHeader()
        {
            #region Call ConfigureRequestPrefixFields to add the User-Agent header.
            string folderSyncRequestBody = Common.CreateFolderSyncRequest("0").GetRequestDataSerializedXML();
            Dictionary<HTTPPOSTRequestPrefixField, string> requestPrefixFields = new Dictionary<HTTPPOSTRequestPrefixField, string>
            {
                {
                    HTTPPOSTRequestPrefixField.UserAgent, "ASOM"
                }
            };

            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);
            #endregion

            #region Call FolderSync command.
            SendStringResponse folderSyncResponse = HTTPAdapter.HTTPPOST(CommandName.FolderSync, null, folderSyncRequestBody);

            // Check the command is executed successfully.
            this.CheckResponseStatus(folderSyncResponse.ResponseDataXML);
            #endregion

            #region Reset the User-Agent header.
            requestPrefixFields[HTTPPOSTRequestPrefixField.UserAgent] = null;
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the X-MS-PolicyKey optional header and Policy key optional field in HTTP POST request.
        /// </summary>
        [TestCategory("MSASHTTP"), TestMethod()]
        public void MSASHTTP_S03_TC03_SetPolicyKeyRequestHeader()
        {
            #region Change the query value type to PlainText.
            // Call ConfigureRequestPrefixFields to set the QueryValueType to PlainText.
            IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();
            requestPrefix.Add(HTTPPOSTRequestPrefixField.QueryValueType, QueryValueType.PlainText.ToString());
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
            #endregion

            #region Call Provision command without setting X-MS-PolicyKey header.
            SendStringResponse provisionResponse = this.CallProvisionCommand(string.Empty);

            // Get the policy key from the response of Provision command.
            string policyKey = TestSuiteHelper.GetPolicyKeyFromSendString(provisionResponse);
            #endregion

            #region Call Provision command with setting X-MS-PolicyKey header of the PlainText encoded query value type.
            // Set the X-MS-PolicyKey header.
            requestPrefix.Add(HTTPPOSTRequestPrefixField.PolicyKey, policyKey);
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

            this.CallProvisionCommand(policyKey);

            // Reset the X-MS-PolicyKey header.
            requestPrefix[HTTPPOSTRequestPrefixField.PolicyKey] = string.Empty;
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
            #endregion

            #region Change the query value type to Base64.
            // Call ConfigureRequestPrefixFields to set the QueryValueType to Base64.
            requestPrefix[HTTPPOSTRequestPrefixField.QueryValueType] = QueryValueType.Base64.ToString();
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
            #endregion

            #region Call Provision command without setting Policy key field.
            provisionResponse = this.CallProvisionCommand(string.Empty);

            // Get the policy key from the response of Provision command.
            policyKey = TestSuiteHelper.GetPolicyKeyFromSendString(provisionResponse);
            #endregion

            #region Call Provision command with setting Policy key field of the base64 encoded query value type.
            // Set the Policy key field.
            requestPrefix[HTTPPOSTRequestPrefixField.PolicyKey] = policyKey;
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

            this.CallProvisionCommand(policyKey);

            // Reset the Policy key field.
            requestPrefix[HTTPPOSTRequestPrefixField.PolicyKey] = string.Empty;
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
            #endregion

            #region Reset the query value type.
            requestPrefix[HTTPPOSTRequestPrefixField.QueryValueType] = Common.GetConfigurationPropertyValue("HeaderEncodingType", this.Site);
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
            #endregion
        }
        #endregion
    }
}