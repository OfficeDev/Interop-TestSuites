namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test the GetAttachment command.
    /// </summary>
    [TestClass]
    public class S06_GetAttachment : TestSuiteBase
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
        /// This test case is used to verify the requirements related to a successful GetAttachment command.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S06_TC01_GetAttachment_Success()
        {
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 14.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
           
            #region Send a mail with normal attachment
            string subject = Common.GenerateResourceName(Site, "NormalAttachment_Subject");
            string body = Common.GenerateResourceName(Site, "NormalAttachment_Body");
            this.SendEmailWithAttachment(subject, body);
            #endregion

            this.SwitchUser(this.User2Information);
            SyncResponse syncResponse = this.GetMailItem(this.User2Information.InboxCollectionId, subject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, subject);

            Response.AttachmentsAttachment[] attachments = this.GetEmailAttachments(syncResponse, subject);
            Site.Assert.IsTrue(attachments != null && attachments.Length == 1, "The email should contain a single attachment.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5501");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5501
            // The attachment name in the sent email is "number1.jpg", so if there is only one attachment in the email and the attachment name is matched, this requirement can be covered.
            Site.CaptureRequirementIfAreEqual<string>(
                "number1.jpg",
                attachments[0].DisplayName,
                5501,
                @"[In GetAttachment] Instead, an Attachment element ([MS-ASAIRS] section 2.2.2.2) is included for each attachment.");

            #region Call GetAttachment command to fetch attachment
            string attachmentFileReference = attachments[0].FileReference;
            IDictionary<CmdParameterName, object> parameters = new Dictionary<CmdParameterName, object>();
            parameters.Add(CmdParameterName.AttachmentName, attachmentFileReference);

            GetAttachmentRequest request = new GetAttachmentRequest();
            request.SetCommandParameters(parameters);
            GetAttachmentResponse response = this.CMDAdapter.GetAttachment(request);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R157");

            // Verify MS-ASCMD requirement: MS-ASCMD_R157
            // The GetAttachment command response data xml only contains the size information of the attachment, if it is not null and includes the size information, this requirement can be covered.
            bool isVerifyR157 = !string.IsNullOrEmpty(response.ResponseDataXML) && Convert.ToInt32(response.ResponseDataXML) > 0;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR157,
                157,
                @"[In GetAttachment] The GetAttachment command retrieves an email attachment from the server.<2>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R160");

            // Verify MS-ASCMD requirement: MS-ASCMD_R160
            // In ExchangeCommonConfiguration.deployment.ptfconfig, HTTP/HTTPS has been specified, so GetAttachment command is issued within the HTTP POST command.
            Site.CaptureRequirement(
                160,
                @"[In GetAttachment] This command [GetAttachment] is issued within the HTTP POST command.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if the GetAttachment command is used to retrieve an attachment that has been deleted on the server, a 500 status code is returned.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S06_TC02_GetAttachment_Status500()
        {
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 14.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            
            // Send a mail with normal attachment
            string subject = Common.GenerateResourceName(Site, "NormalAttachment_Subject");
            string body = Common.GenerateResourceName(Site, "NormalAttachment_Body");
            this.SendEmailWithAttachment(subject, body);

            this.SwitchUser(this.User2Information);
            SyncResponse syncResponse = this.CheckEmail(this.User2Information.InboxCollectionId, subject, null);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, subject);

            Response.AttachmentsAttachment[] attachments = this.GetEmailAttachments(syncResponse, subject);
            Site.Assert.IsTrue(attachments != null && attachments.Length == 1, "The email should contain a single attachment.");

            GetAttachmentRequest getAttachmentRequest = new GetAttachmentRequest();

            getAttachmentRequest.SetCommandParameters(new Dictionary<CmdParameterName, object>
            {
                {
                    CmdParameterName.AttachmentName, attachments[0].FileReference
                }
            });

            GetAttachmentResponse getAttachmentResponse = this.CMDAdapter.GetAttachment(getAttachmentRequest);
            Site.Assert.AreEqual<string>("OK", getAttachmentResponse.StatusDescription, "The attachment should be retrieved successfully.");

            // Delete the email in the Inbox folder.
            syncResponse = this.SyncChanges(this.User2Information.InboxCollectionId);

            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", subject);
            SyncRequest syncRequest = TestSuiteBase.CreateSyncDeleteRequest(this.LastSyncKey, this.User2Information.InboxCollectionId, serverId);
            this.Sync(syncRequest);

            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNull(TestSuiteBase.FindServerId(syncResponse, "Subject", subject), "The email should be deleted.");
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, subject);

            syncResponse = this.SyncChanges(this.User2Information.DeletedItemsCollectionId);
            Site.Assert.IsNotNull(TestSuiteBase.FindServerId(syncResponse, "Subject", subject), "The deleted email should be in the DeletedItems folder.");
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.DeletedItemsCollectionId, subject);

            try
            {
                // Call GetAttachment command again to fetch the deleted attachment.
                this.CMDAdapter.GetAttachment(getAttachmentRequest);
                Site.Assert.Fail("If the GetAttachment command is used to retrieve an attachment that has been deleted on the server, a 500 status code should be returned in the HTTP POST response.");
            }
            catch (System.Net.WebException exception)
            {
                int statusCode = ((System.Net.HttpWebResponse)exception.Response).StatusCode.GetHashCode();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R166");

                // Verify MS-ASCMD requirement: MS-ASCMD_R166
                Site.CaptureRequirementIfAreEqual<int>(
                    500,
                    statusCode,
                    166,
                    @"[In GetAttachment] If the GetAttachment command is used to retrieve an attachment that has been deleted on the server, a 500 status code is returned in the HTTP POST response.");
            }
        }
        #endregion
    }
}