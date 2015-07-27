//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASRM
{
    using System.Globalization;
    using Common.DataStructures;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the SmartReply command.
    /// </summary>
    [TestClass]
    public class S05_SmartReply : TestSuiteBase
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
        public static void ClassCleanUp()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region MSASRM_S05_TC01_SmartReply_EditAllowed_False_NoReplaceMime
        /// <summary>
        /// This test case is designed to test when EditAllowed is set to false and composemail:ReplaceMime is not present in a SmartReply request, the server will add the original rights-managed email message as an attachment to the new message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC01_SmartReply_EditAllowed_False_NoReplaceMime()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with EditAllowed is set to false.
            string templateID = this.GetTemplateID("MSASRM_EditExport_NotAllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ReplyAllowed, "The ReplyAllowed element in expected rights-managed e-mail message should be true.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ModifyRecipientsAllowed, "The ModifyRecipientsAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account, calls SmartReply method without ReplaceMime in request to reply the received email to User1.
            string replySubject = string.Format("Re: {0}", subject);

            string replyMime = Common.CreatePlainTextMime(
                    Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                    Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                    string.Empty,
                    string.Empty,
                    replySubject,
                    Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and get the rights-managed replied e-mail message
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R104");

            // Verify MS-ASRM requirement: MS-ASRM_R104
            // The original e-mail item has ReplyAllowed element set to true.
            Site.CaptureRequirementIfIsNotNull(
                item.Email,
                104,
                @"[In ReplyAllowed] The value is TRUE (1) if the user can reply to the e-mail message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R375");

            // Verify MS-ASRM requirement: MS-ASRM_R375
            // The SmartReply command request included the same TemplateID element with that in SendMail command request.
            // The original e-mail item has ModifyRecipientsAllowed element set to true, to allow the recipient list modified, and so that the server creates the reply message.
            Site.CaptureRequirementIfIsNotNull(
                item.Email,
                375,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartReply command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template: If the TemplateID element is included in the command request, the server does the following:] If the original message is protected and the specified TemplateID value is the same as the TemplateID value on the original message, the server proceeds to step 4 [The server compares the recipients (1) on the original message to the recipients (1) sent by the client within the new message. The server verifies that the recipient (1) list on the new message aligns with the granted permissions, as specified in the following table[section 3.2.5.1]. If permissions allow it, the server creates the reply message].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R387");

            // Verify MS-ASRM requirement: MS-ASRM_R387
            // The SmartReply command request included the same TemplateID element with that in SendMail command request.
            // Only the sender received the e-mail message.
            Site.CaptureRequirementIfIsNotNull(
                item.Email,
                387,
                @"[In Handling SmartForward and SmartReply Requests] If ReplyAllowed is set to TRUE, the server will reply to exactly one recipient (1).");

            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R470");

            // Verify MS-ASRM requirement: MS-ASRM_R470
            // The original e-mail item has EditAllowed element set to false, and the ReplaceMime element is not include in SmartReply command request.
            // If the response contains Attachments element as not null, represents the server will add the original rights-managed email message as an attachment to the new message.
            Site.CaptureRequirementIfIsNotNull(
                item.Email.Attachments,
                470,
                @"[In EditAllowed] When EditAllowed is set to FALSE and composemail:ReplaceMime ([MS-ASCMD] section 2.2.3.135) is not present in a SmartReply request, the server will add the original rights-managed email message as an attachment to the new message.");

            #endregion
        }
        #endregion

        #region MSASRM_S05_TC02_SmartReply_EditAllowed_False_ReplaceMime
        /// <summary>
        /// This test case is designed to test when EditAllowed is set to false and composemail:ReplaceMime is present in a SmartReply request, the server will not attach the original rights-managed email message as an attachment.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC02_SmartReply_EditAllowed_False_ReplaceMime()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with EditAllowed is set to false.
            string templateID = this.GetTemplateID("MSASRM_EditExport_NotAllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be false.");
            #endregion

            #region The client logs on User2's account and calls SmartReply method with ReplaceMime in request to reply the received email to User1.
            string replySubject = string.Format("Re: {0}", subject);

            string replyMime = Common.CreatePlainTextMime(
                    Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                    Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                    string.Empty,
                    string.Empty,
                    replySubject,
                    Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            replyRequest.RequestData.ReplaceMime = string.Empty;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");

            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox.
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R471");

            // Verify MS-ASRM requirement: MS-ASRM_R471
            // The original e-mail item has EditAllowed element set to false, and the ReplaceMime element is included in SmartReply command request.
            // If the response contains Attachments element as null, represents the server will not attach the original rights-managed email message as an attachment to the new message.
            Site.CaptureRequirementIfIsNull(
                item.Email.Attachments,
                471,
                @"[In EditAllowed] Conversely, if [EditAllowed is set to FALSE and]composemail:ReplaceMime is present[in a SmartReply command], the server will not attach the original rights-managed email message as an attachment.");

            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");

            #endregion
        }
        #endregion

        #region MSASRM_S05_TC03_SmartReply_ExportAllowed_True
        /// <summary>
        /// This test case is designed to test when the user reply the e-mail message and the ExportAllowed is true, the user can remove the IRM protection.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC03_SmartReply_ExportAllowed_True()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account and calls Settings command to get a templateID with ExportAllowed set to true.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, then switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with TemplateID set to "00000000-0000-0000-0000-000000000000" in request to reply the received email to User1.
            string replySubject = string.Format("Re: {0}", subject);

            string replyMime = Common.CreatePlainTextMime(
                    Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                    Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                    string.Empty,
                    string.Empty,
                    replySubject,
                    Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = "00000000-0000-0000-0000-000000000000";
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");

            this.SwitchUser(this.UserOneInformation, false);
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and gets the e-mail message which removed the IRM protection.
            Sync repliedItem = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(repliedItem, "The returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R343");

            // Verify MS-ASRM requirement: MS-ASRM_R343
            // The ExportAllowed element is set to TRUE in original message, if the RightsManagementLicense element of new message is null, this requirement can be verified.
            Site.CaptureRequirementIfIsNull(
                repliedItem.Email.RightsManagementLicense,
                343,
                @"[In ExportAllowed] The value is TRUE (1) if the user can remove the IRM protection when the user replies to the e-mail message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R349");

            // Verify MS-ASRM requirement: MS-ASRM_R349
            // The ExportAllowed element is set to TRUE and the TemplateID on the new message is set to "00000000-0000-0000-0000-000000000000", if the RightsManagementLicense element of new message is null, this requirement can be verified.
            Site.CaptureRequirementIfIsNull(
                repliedItem.Email.RightsManagementLicense,
                349,
                @"[In TemplateID (SendMail, SmartForward, SmartReply)] If a rights-managed e-mail message is replied to using the SmartReply command, the IRM protection is removed from the outgoing message if the following conditions are true:
 The original rights policy template has the ExportAllowed element set to TRUE.
 The TemplateID on the new message is set to the ""No Restriction"" template (TemplateID value ""00000000-0000-0000-0000-000000000000"").");
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R169");

            // Verify MS-ASRM requirement: MS-ASRM_R169
            // The ExportAllowed element is set to TRUE and the TemplateID on the new message is set to "00000000-0000-0000-0000-000000000000", if the RightsManagementLicense element of original message is not null, this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item.Email.RightsManagementLicense,
                169,
                @"[In TemplateID (SendMail, SmartForward, SmartReply)] [if the following conditions are true:
 The original rights policy template has the ExportAllowed element set to TRUE.
 The TemplateID on the new message is set to the ""No Restriction"" template (TemplateID value ""00000000-0000-0000-0000-000000000000"")]The original message retains its IRM protection.");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC04_SmartReplyAll_ExportAllowed_True
        /// <summary>
        /// This test case is designed to test when the user replies the e-mail message to all recipients and the ExportAllowed is true, the user can remove the IRM protection.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC04_SmartReplyAll_ExportAllowed_True()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account and calls Settings command to get a templateID with ExportAllowed set to true.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2 and cc User3, switches to User3 and User2 to do FolderSync
            string subject = this.SendMailAndFolderSync(templateID, false, this.UserThreeInformation);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ReplyAllAllowed, "The ReplyAllAllowed element in expected rights-managed e-mail message should be true.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account and calls SmartReply method with TemplateID set to "00000000-0000-0000-0000-000000000000" in request to reply the received email to User1 and cc User3.
            string replySubject = string.Format("Re: {0}", subject);

            string replyMime = Common.CreatePlainTextMime(
                    Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                    Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                    Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                    string.Empty,
                    replySubject,
                    Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = "00000000-0000-0000-0000-000000000000";
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");

            this.SwitchUser(this.UserThreeInformation, true);
            TestSuiteBase.AddCreatedItemToCollection(this.UserThreeInformation, this.UserThreeInformation.InboxCollectionId, replySubject);
            this.SwitchUser(this.UserOneInformation, false);
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and gets the e-mail message which removed the IRM protection.
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in returned item should be null.");

            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and gets the e-mail message which removed the IRM protection.
            this.SwitchUser(this.UserThreeInformation, false);
            item = this.SyncEmail(replySubject, this.UserThreeInformation.InboxCollectionId, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R99");

            // Verify MS-ASRM requirement: MS-ASRM_R99
            // If the replied message arrives in both User1 and User3's mailbox, this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item,
                99,
                @"[In ReplyAllAllowed] The value is TRUE (1) if the user can reply to all of the recipients (1) of the e-mail message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R344");

            // Verify MS-ASRM requirement: MS-ASRM_R344
            // The ExportAllowed element is set to TRUE in original message, if the RightsManagementLicense element of new message is null, this requirement can be verified.
            Site.CaptureRequirementIfIsNull(
                item.Email.RightsManagementLicense,
                344,
                @"[In ExportAllowed] The value is TRUE (1) if the user can remove the IRM protection when the user replies all to the e-mail message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R400");

            // Verify MS-ASRM requirement: MS-ASRM_R400
            // This requirement can be captured after checking the message arrives in both User1 and User3's mailbox
            Site.CaptureRequirement(
                400,
                @"[In Handling SmartForward and SmartReply Requests] If ReplyAllAllowed is set to TRUE, the server will reply to all original recipients (1).");

            #endregion
        }
        #endregion

        #region MSASRM_S05_TC05_SmartReply_ReplyAllowed_True_ReplyAllAllowed_False_Status172
        /// <summary>
        /// This test case is designed to test when the user replies the e-mail message, the ReplyAllowed is true and the ReplyAllAllowed is false, the user cannot reply to all and will receive a response with status 172.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC05_SmartReply_ReplyAllowed_True_ReplyAllAllowed_False_Status172()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account and calls Settings command to get a templateID with ReplyAllAllowed set to false.
            string templateID = this.GetTemplateID("MSASRM_ViewReply_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2 and cc User3, then switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, this.UserThreeInformation);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ReplyAllowed, "The ReplyAllowed element in expected rights-managed e-mail message should be true.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ReplyAllAllowed, "The ReplyAllAllowed element in expected rights-managed e-mail message should be false.");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with the TemplateID in request to reply the received email to User1 and cc User3, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);

            string replyMime = Common.CreatePlainTextMime(
                    Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                    Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                    Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                    string.Empty,
                    replySubject,
                    Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.IsNotNull(smartReplyResponse.ResponseData, "The SmartReply element should not be null.");
            Site.Assert.AreEqual<string>("172", smartReplyResponse.ResponseData.Status, "The Status of SmartReply should be 172.");
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives.
            this.SwitchUser(this.UserOneInformation, false);

            // Get the new added email item
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, false);
            Site.Assert.IsNull(item, "The returned item should be null.");
            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and checks if the e-mail message arrives.
            this.SwitchUser(this.UserThreeInformation, true);

            // Get the new added email item
            item = this.SyncEmail(replySubject, this.UserThreeInformation.InboxCollectionId, true, false);
            Site.Assert.IsNull(item, "The returned item should be null.");
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R653");

            // Verify MS-ASRM requirement: MS-ASRM_R653
            // The original e-mail item has ReplyAllowed element set to true, ReplyAllAllowed set to false, the recipient reply the email to all and receive a response with status 172, thus this requirement can be verified.
            Site.CaptureRequirement(
                653,
                @"[In Handling SmartForward and SmartReply Requests] [If ReplyAllowed is set to TRUE and ReplyAllAllowed is set to FALSE, the server will reply to exactly one recipient (1), the sender of the original message.] All other SmartReply command requests are restricted and error out with a composemail:Status value of 172.");
        }
        #endregion

        #region MSASRM_S05_TC06_SmartReply_ReplyAllowed_False_ReplyAllAllowed_True_Status172
        /// <summary>
        /// This test case is designed to test when the user reply the e-mail message, the ReplyAllowed is false and the ReplyAllAllowed is true, the user cannot reply to the sender and will receive a response with status 172.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC06_SmartReply_ReplyAllowed_False_ReplyAllAllowed_True_Status172()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account and calls Settings command to get a templateID with ReplyAllowed set to false.
            string templateID = this.GetTemplateID("MSASRM_ViewReplyAll_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, then switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, this.UserThreeInformation);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ReplyAllowed, "The ReplyAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ReplyAllAllowed, "The ReplyAllAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with the TemplateID in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);

            string replyMime = Common.CreatePlainTextMime(
                    Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                    Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                    string.Empty,
                    string.Empty,
                    replySubject,
                    Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.IsNotNull(smartReplyResponse.ResponseData, "The SmartReply element should not be null.");
            Site.Assert.AreEqual<string>("172", smartReplyResponse.ResponseData.Status, "The Status of SmartReply should be 172.");
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives.
            this.SwitchUser(this.UserOneInformation, false);

            // Get the new added email item
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, false);
            Site.Assert.IsNull(item, "The returned item should be null.");
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R654");

            // The original e-mail item has ReplyAllowed element set to false, ReplyAllAllowed set to true, the recipient replies the email to the sender, and receives a response with status 172, thus this requirement can be verified.
            Site.CaptureRequirement(
                654,
                @"[In Handling SmartForward and SmartReply Requests] [If ReplyAllowed is set to FALSE and ReplyAllAllowed is set to TRUE, the server will reply to exactly all original recipients (1). Whether the sender chooses to include themselves in the reply message is optional.] All other SmartReply command requests are restricted and error out with a composemail:Status value of 172.");
        }
        #endregion

        #region MSASRM_S05_TC07_SmartReply_ModifyRecipientsAllowed_False
        /// <summary>
        /// This test case is designed to test when the user reply the e-mail message and the ModifyRecipientsAllowed is false, the user cannot change the recipients in SmartReply command.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC07_SmartReply_ModifyRecipientsAllowed_False()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account and calls Settings command to get a templateID with ModifyRecipientsAllowed set to false.
            string templateID = this.GetTemplateID("MSASRM_View_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, then switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ModifyRecipientsAllowed, "The ModifyRecipientsAllowed element in expected rights-managed e-mail message should be false.");
            #endregion

            #region The client logs on User2's account and calls SmartReply method with TemplateID in request to reply the received email to another account:User3.
            string replySubject = string.Format("Re: {0}", subject);

            string replyMime = Common.CreatePlainTextMime(
                    Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                    Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                    string.Empty,
                    string.Empty,
                    replySubject,
                    Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.IsNotNull(smartReplyResponse.ResponseData, "The SmartReply element should not be null.");
            Site.Assert.AreEqual<string>("172", smartReplyResponse.ResponseData.Status, "The Status of SmartReply should be 172.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R329");

            // The original e-mail item has ModifyRecipientsAllowed element set to false, and receive a SmartReply response with status 172, thus this requirement can be verified.
            Site.CaptureRequirement(
                329,
                @"[In ModifyRecipientsAllowed] otherwise[the user cannot modify the recipient (1) list], FALSE (0).");
     
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R652");

            // The original e-mail item has ModifyRecipientsAllowed element set to false, and receive a SmartReply response with status 172, thus this requirement can be verified.
            Site.CaptureRequirement(
                652,
                @"[In Handling SmartForward and SmartReply Requests] [If ReplyAllowed and ReplyAllAllowed are set to TRUE, the server will reply to exactly one recipient (1) or all original recipients (1). Whether the sender chooses to include themselves in the reply message is optional.] All other SmartReply command requests are restricted and error out with a composemail:Status value of 172.");

            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and checks if the e-mail message arrives.
            this.SwitchUser(this.UserThreeInformation, true);

            // Get the new added email item
            item = this.SyncEmail(replySubject, this.UserThreeInformation.InboxCollectionId, true, false);
            Site.Assert.IsNull(item, "The returned item should be null.");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC08_SmartReply_ModifyRecipientsAllowed_True
        /// <summary>
        /// This test case is designed to test when the user replies the e-mail message and the ModifyRecipientsAllowed is true, the user can change the recipients in SmartReply command.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC08_SmartReply_ModifyRecipientsAllowed_True()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account and calls Settings command to get a templateID with ModifyRecipientsAllowed set to true.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ModifyRecipientsAllowed, "The ModifyRecipientsAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account and calls SmartReply method with TemplateID in request to reply the received email to another account:User3.
            string replySubject = string.Format("Re: {0}", subject);

            string replyMime = Common.CreatePlainTextMime(
                    Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                    Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                    string.Empty,
                    string.Empty,
                    replySubject,
                    Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and checks if the e-mail message arrives.
            this.SwitchUser(this.UserThreeInformation, true);
            TestSuiteBase.AddCreatedItemToCollection(this.UserThreeInformation, this.UserThreeInformation.InboxCollectionId, replySubject);

            // Get the new added email item
            item = this.SyncEmail(replySubject, this.UserThreeInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R67");

            // The original e-mail item has ModifyRecipientsAllowed element set to true, the recipient replies the email and the User3 can receive the email, thus this requirement can be verified.
            Site.CaptureRequirement(
                67,
                @"[In ModifyRecipientsAllowed] The value is TRUE (1) if the user can modify the recipient (1) list.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R379");

            // The original e-mail item has ModifyRecipientsAllowed element set to true, the recipient reply the email and the User3 can receive the email, thus this requirement can be verified.
            Site.CaptureRequirement(
                379,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartReply command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template: If the TemplateID element is included in the command request, the server does the following:] If the TemplateID value corresponds to a template on the server, the server proceeds to step 4[The server compares the recipients (1) on the original message to the recipients (1) sent by the client within the new message. The server verifies that the recipient (1) list on the new message aligns with the granted permissions, as specified in the following table. If permissions allow it, the server creates the reply message]. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R383");

            // The original e-mail item has ModifyRecipientsAllowed element set to true, the recipient reply the email and the sender can receive the email, thus this requirement can be verified.
            Site.CaptureRequirement(
                383,
                @"[In Handling SmartForward and SmartReply Requests] If ModifyRecipientsAllowed is set to TRUE, there is no restrictions for server to perform reply or reply all enforcement.");
        }
        #endregion

        #region MSASRM_S05_TC09_SmartReply_ReplyAllowed_True_ReplyAllAllowed_False
        /// <summary>
        /// This test case is designed to test when the user replies the e-mail message, the ReplyAllowed is true and the ReplyAllAllowed is false, the user can reply to the sender.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC09_SmartReply_ReplyAllowed_True_ReplyAllAllowed_False()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account and calls Settings command to get a templateID with ReplyAllAllowed set to false.
            string templateID = this.GetTemplateID("MSASRM_ViewReply_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2 and cc User3, then switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, this.UserThreeInformation);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ReplyAllowed, "The ReplyAllowed element in expected rights-managed e-mail message should be true.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ReplyAllAllowed, "The ReplyAllAllowed element in expected rights-managed e-mail message should be false.");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with the TemplateID in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);

            string replyMime = Common.CreatePlainTextMime(
                    Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                    Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                    string.Empty,
                    string.Empty,
                    replySubject,
                    Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives.
            this.SwitchUser(this.UserOneInformation, false);

            // Get the new added email item
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);

            Site.Assert.IsNotNull(item, "The returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R333");

            // The original e-mail item has ReplyAllowed element set to true, ReplyAllAllowed set to false, the recipient replies the email and the sender can receive the email, thus this requirement can be verified.
            Site.CaptureRequirement(
                333,
                @"[In ReplyAllAllowed] otherwise[the user cannot reply to all of the recipients (1) of the e-mail message], FALSE (0).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R389");

            // The original e-mail item has ReplyAllowed element set to true, ReplyAllAllowed set to false, the recipient reply the email and the sender can receive the email, thus this requirement can be verified.
            Site.CaptureRequirement(
                389,
                @"[In Handling SmartForward and SmartReply Requests] If ReplyAllowed is set to TRUE and ReplyAllAllowed is set to FALSE, the server will reply to exactly one recipient (1), the sender of the original message. ");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC10_SmartReply_ReplyAllowed_False_ReplyAllAllowed_True
        /// <summary>
        /// This test case is designed to test when the user reply the e-mail message, the ReplyAllowed is false and the ReplyAllAllowed is true, the user can reply to all the recipients.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC10_SmartReply_ReplyAllowed_False_ReplyAllAllowed_True()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account and calls Settings command to get a templateID with ReplyAllowed set to false.
            string templateID = this.GetTemplateID("MSASRM_ViewReplyAll_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, then switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, this.UserThreeInformation);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ReplyAllowed, "The ReplyAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ReplyAllAllowed, "The ReplyAllAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with the TemplateID in request to reply the received email to User1 and cc User3, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);

            string replyMime = Common.CreatePlainTextMime(
                    Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                    Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                    Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                    string.Empty,
                    replySubject,
                    Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives.
            this.SwitchUser(this.UserOneInformation, false);

            // Get the new added email item
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and checks if the e-mail message arrives.
            this.SwitchUser(this.UserThreeInformation, true);

            // Get the new added email item
            item = this.SyncEmail(replySubject, this.UserThreeInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserThreeInformation, this.UserThreeInformation.InboxCollectionId, replySubject);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R390");

            // The original e-mail item has ReplyAllowed element set to false, ReplyAllAllowed set to true, the recipient reply the email to all, all original recipients can receive the email, thus this requirement can be verified.
            Site.CaptureRequirement(
                390,
                @"[In Handling SmartForward and SmartReply Requests] If ReplyAllowed is set to FALSE and ReplyAllAllowed is set to TRUE, the server will reply to exactly all original recipients (1). ");
        }
        #endregion

        #region MSASRM_S05_TC11_SmartReply_ReplyAllowed_False_ReplyAllAllowed_False
        /// <summary>
        /// This test case is designed to test when ReplyAllowed and ReplyAllAllowed are set to FALSE, SmartReply command request error out with a composemail:Status value of 172.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC11_SmartReply_ReplyAllowed_False_ReplyAllAllowed_False()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account and calls Settings command to get a templateID with all rights denied except view rights.
            string templateID = this.GetTemplateID("MSASRM_View_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ReplyAllowed, "The ReplyAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ReplyAllAllowed, "The ReplyAllAllowed element in expected rights-managed e-mail message should be false.");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with the TemplateID in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);

            string replyMime = Common.CreatePlainTextMime(
                    Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                    Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                    string.Empty,
                    string.Empty,
                    replySubject,
                    Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.IsNotNull(smartReplyResponse.ResponseData, "The SmartReply element should not be null.");
            Site.Assert.AreEqual<string>("172", smartReplyResponse.ResponseData.Status, "The Status of SmartReply should be 172.");
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives.
            this.SwitchUser(this.UserOneInformation, false);

            // Get the new added email item
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, false);
            Site.Assert.IsNull(item, "The returned item should be null.");
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R334");

            // The original e-mail item has ReplyAllowed element set to false, the recipient reply the email to sender, thus this requirement can be verified.
            Site.CaptureRequirement(
                334,
                @"[In ReplyAllowed] otherwise[the user cannot reply to the e-mail message], FALSE (0).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R392");

            // The original e-mail item has ReplyAllowed element and ReplyAllAllowed set to false, the recipient reply the email, and receive a response with status 172, thus this requirement can be verified.
            Site.CaptureRequirement(
                392,
                @"[In Handling SmartForward and SmartReply Requests] If ReplyAllowed and ReplyAllAllowed are set to FALSE, SmartReply command request error out with a composemail:Status value of 172.");
        }
        #endregion

        #region MSASRM_S05_TC12_SmartReply_Status171
        /// <summary>
        /// This test case is designed to test when the request included an invalid TemplateID value, the server returns status value 171 in a SmartReply command response.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC12_SmartReply_Status171()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account and calls Settings command to get a templateID with all rights allowed.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with an invalid TemplateID in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);

            string replyMime = Common.CreatePlainTextMime(
                    Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                    Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                    string.Empty,
                    string.Empty,
                    replySubject,
                    Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = "invalidTemplateID";
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.IsNotNull(smartReplyResponse.ResponseData, "The SmartReply element should not be null.");
            Site.Assert.AreEqual<string>("171", smartReplyResponse.ResponseData.Status, "The Status of SmartReply should be 171.");
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the message arrives.
            this.SwitchUser(this.UserOneInformation, false);

            // Get the new added email item
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, false);
            Site.Assert.IsNull(item, "The returned item should be null.");
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R200");

            // The server returns Status value 171 in a SmartReply response, thus this requirement can be verified.
            Site.CaptureRequirement(
                200,
                @"[In Enforcing Rights Policy Template Settings] The server returns Status value 171 in a SmartReply command response if the request included an invalid TemplateID value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R374");

            // The server returns Status value 171 in a SmartReply response, thus this requirement can be verified.
            Site.CaptureRequirement(
                374,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartReply command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template: If the TemplateID element is included in the command request, the server does the following:] If the template does not exist on the server, the server fails the request and returns a composemail:Status value of 171. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R378");

            // The server returns Status value 171 in a SmartReply response, thus this requirement can be verified.
            Site.CaptureRequirement(
                378,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartReply command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template: If the TemplateID element is included in the command request, the server does the following:] If the TemplateID value does not correspond to a template on the server, the server fails the request and returns a composemail:Status value of 171; ");
        }
        #endregion

        #region MSASRM_S05_TC13_SmartReply_NotProtected_OriginalMessage_NotProtected
        /// <summary>
        /// This test case is designed to test when the original message being replied has no rights management restrictions and no TemplateID element is included in the command request, the server sends the new message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC13_SmartReply_NotProtected_OriginalMessage_NotProtected()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls SendMail command without a templateID to send a message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(null, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should be null.");
            
            #endregion

            #region The client logs on User2's account, calls SmartReply method to reply the received email to User1.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and gets the e-mail message
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R366");

            // Verify MS-ASRM requirement: MS-ASRM_R366
            // The original e-mail item has no IRM protection, and no TemplateID element is included in the SmartReply command request.
            // If the new message arrives, this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item.Email,
                366,
                @"[In Handling SmartForward and SmartReply Requests] When the client sends the server a SmartReply command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template:
1.	If no TemplateID element is included in the command request, the server proceeds as follows: 
If the original message being replied to has no rights management restrictions, the server proceeds to step 6[The server sends the new message].");

            #endregion
        }
        #endregion

        #region MSASRM_S05_TC14_SmartReply_HTTP_Status168
        /// <summary>
        /// This test case is designed to test when the connection to the server does not use SSL, the server fails the request and returns composemail:Status value 168.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC14_SmartReply_HTTP_Status168()
        {
            Site.Assume.AreEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Implementation does consider the XML body of the command request to be invalid, if the protocol version specified by in the command request is not 14.1.");
            Site.Assume.AreEqual<string>("HTTP", Common.GetConfigurationPropertyValue("TransportType", this.Site).ToUpper(CultureInfo.CurrentCulture), "This test case is designed to run under HTTP");

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command without the templateID to send a rights-managed e-mail message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(null, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to false to synchronize changes of Inbox folder in User2's mailbox, and gets the e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, false, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should be null.");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with the TemplateID in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R371");

            // Verify MS-ASRM requirement: MS-ASRM_R371
            Site.CaptureRequirementIfAreEqual<string>(
                "168",
                smartReplyResponse.ResponseData.Status,
                371,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartReply command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template: If any of the following are true, the server fails the request and returns composemail:Status value 168:] The connection to the server does not use SSL.");

            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, false);
            Site.Assert.IsNull(item, "The returned item should be null.");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC15_SmartReply_Protected_OriginalMessage_NotProtected
        /// <summary>
        /// This test case is designed to test when the original message being replied has no rights management restrictions and the TemplateID element is included in the command request, the server protects the new outgoing message with the specified rights policy template.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC15_SmartReply_Protected_OriginalMessage_NotProtected()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command without the templateID to send a rights-managed e-mail message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(null, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to false to synchronize changes of Inbox folder in User2's mailbox, and gets the e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, false, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should be null.");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with the TemplateID in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R373");

            // Verify MS-ASRM requirement: MS-ASRM_R373
            // The original e-mail item has no IRM protection, and the TemplateID element is included in the SmartReply command request,
            // If the new message has IRM protection, this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item.Email.RightsManagementLicense,
                373,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartReply command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template:] If the TemplateID element is included in the command request, the server does the following: If the original message is not protected, the server proceeds to step 5[If the message has a TemplateID element, the server protects the new outgoing message with the specified rights policy template.].");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC16_SmartReply_Protected_OriginalMessage_Protected_DifferentTemplate
        /// <summary>
        /// This test case is designed to test when the original message being replied has rights management restrictions and the TemplateID element is included in the command request, the server protects the new outgoing message with the specified rights policy template.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC16_SmartReply_Protected_OriginalMessage_Protected_DifferentTemplate()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            #endregion

            #region The client logs on User2's account, calls Settings command to get another templateID with reply rights allowed.
            templateID = this.GetTemplateID("MSASRM_EditExport_NotAllowedTemplate");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with the TemplateID in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R377");

            // Verify MS-ASRM requirement: MS-ASRM_R377
            Site.CaptureRequirementIfAreEqual<string>(
                "MSASRM_EditExport_NotAllowedTemplate",
                item.Email.RightsManagementLicense.TemplateName,
                377,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartReply command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template: If the TemplateID element is included in the command request, the server does the following:] If the original message is protected and the specified TemplateID value is different than the TemplateID value on the original message, the server verifies that the new TemplateID value exists on the server.");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC17_SmartReply_Export_True_Edit_True_ReplaceMime_SameTemplate
        /// <summary>
        /// This test case is designed to test when ExportAllowed, EditAllowed are set to true and with ReplaceMime element included in SmartReply request, the server will send the message when it uses the same TemplateID with the original message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC17_SmartReply_Export_True_Edit_True_ReplaceMime_SameTemplate()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be true.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with the same TemplateID and ReplaceMime element in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            replyRequest.RequestData.ReplaceMime = string.Empty;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R253");

            // Verify MS-ASRM requirement: MS-ASRM_R253
            // The original e-mail item has ExportAllowed and EditAllowed element set to TRUE, and composemail:ReplaceMime is included in the SmartReply command request with the same TemplateID in the original message,
            // If the new message arrives, this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item,
                253,
                @"[In Handling SmartForward and SmartReply Requests] If ExportAllowed, EditAllowed and composemail:ReplaceMime are set to TRUE, the server will send the message when it uses the same TemplateID with the original message.");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC18_SmartReply_Export_True_Edit_True_ReplaceMime_DifferentTemplate
        /// <summary>
        /// This test case is designed to test when ExportAllowed, EditAllowed are set to true and with ReplaceMime element included in SmartReply request, the server will send the message when it uses a different TemplateID with the original message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC18_SmartReply_Export_True_Edit_True_ReplaceMime_DifferentTemplate()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be true.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account, calls Settings command to get another templateID with reply rights allowed.
            templateID = this.GetTemplateID("MSASRMReplyAll_NotAllowedTemplate");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with a new TemplateID and ReplaceMime element in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            replyRequest.RequestData.ReplaceMime = string.Empty;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R254");

            // Verify MS-ASRM requirement: MS-ASRM_R254
            // The original e-mail item has ExportAllowed and EditAllowed element set to TRUE, and composemail:ReplaceMime is included in the SmartReply command request with a different TemplateID in the original message,
            // If the new message arrives, this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item,
                254,
                @"[In Handling SmartForward and SmartReply Requests] If ExportAllowed, EditAllowed and composemail:ReplaceMime are set to TRUE, the server will send the message when it uses a different TemplateID with the original message.");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC19_SmartReply_Export_False_Edit_True_ReplaceMime_SameTemplate
        /// <summary>
        /// This test case is designed to test when ExportAllowed is set to false, EditAllowed is set to true and with ReplaceMime element included in SmartReply request, the server will send the message when it uses the same TemplateID with the original message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC19_SmartReply_Export_False_Edit_True_ReplaceMime_SameTemplate()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed except export rights.
            string templateID = this.GetTemplateID("MSASRM_Export_NotAllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with the same TemplateID and ReplaceMime element in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            replyRequest.RequestData.ReplaceMime = string.Empty;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R257");

            // Verify MS-ASRM requirement: MS-ASRM_R257
            // The original e-mail item has ExportAllowed element set to FALSE and EditAllowed element set to TRUE, and composemail:ReplaceMime is included in the SmartReply command request with the same TemplateID in the original message,
            // If the new message arrives, this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item,
                257,
                @"[In Handling SmartForward and SmartReply Requests] If ExportAllowed is set to FALSE, EditAllowed and composemail:ReplaceMime are set to TRUE, the server will send the message when it uses the same TemplateID with the original message.");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC20_SmartReply_Export_False_Edit_True_ReplaceMime_DifferentTemplate
        /// <summary>
        /// This test case is designed to test when ExportAllowed is set to false, EditAllowed is set to true and with ReplaceMime element included in SmartReply request, the server will send the message when it uses a different TemplateID with the original message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC20_SmartReply_Export_False_Edit_True_ReplaceMime_DifferentTemplate()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed except export rights.
            string templateID = this.GetTemplateID("MSASRM_Export_NotAllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account, calls Settings command to get another templateID with reply rights allowed.
            templateID = this.GetTemplateID("MSASRMReplyAll_NotAllowedTemplate");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with a new TemplateID and ReplaceMime element in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            replyRequest.RequestData.ReplaceMime = string.Empty;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R258");

            // Verify MS-ASRM requirement: MS-ASRM_R258
            // The original e-mail item has ExportAllowed element set to FALSE and EditAllowed element set to TRUE, and composemail:ReplaceMime is included in the SmartReply command request with a different TemplateID in the original message,
            // If the new message arrives, this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item,
                258,
                @"[In Handling SmartForward and SmartReply Requests] If ExportAllowed is set to FALSE, EditAllowed and composemail:ReplaceMime are set to TRUE, the server will send the message when it uses a different TemplateID with the original message.");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC21_SmartReply_Export_False_Edit_False_ReplaceMime_DifferentTemplate
        /// <summary>
        /// This test case is designed to test when ExportAllowed and EditAllowed are set to false and with ReplaceMime element included in SmartReply request, the server will send the message when it uses a different TemplateID with the original message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC21_SmartReply_Export_False_Edit_False_ReplaceMime_DifferentTemplate()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed except Edit and Export rights.
            string templateID = this.GetTemplateID("MSASRM_EditExport_NotAllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be false.");
            #endregion

            #region The client logs on User2's account, calls Settings command to get another templateID with reply rights allowed.
            templateID = this.GetTemplateID("MSASRMReplyAll_NotAllowedTemplate");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with a new TemplateID and ReplaceMime element in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            replyRequest.RequestData.ReplaceMime = string.Empty;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R260");

            // Verify MS-ASRM requirement: MS-ASRM_R260
            // The original e-mail item has ExportAllowed and EditAllowed elements set to FALSE, and composemail:ReplaceMime is included in the SmartReply command request with a different TemplateID in the original message,
            // If the new message arrives, this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item,
                260,
                @"[In Handling SmartForward and SmartReply Requests] If ExportAllowed and EditAllowed are set to FALSE, composemail:ReplaceMime is set to TRUE, the server will send the message when it uses a different TemplateID with the original message.");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC22_SmartReply_Export_True_Edit_True_NoReplaceMime_SameTemplate
        /// <summary>
        /// This test case is designed to test when ExportAllowed and EditAllowed are set to true and without ReplaceMime element included in SmartReply request, the server will send the message including the original message as inline content when it uses the same TemplateID with the original message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC22_SmartReply_Export_True_Edit_True_NoReplaceMime_SameTemplate()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be true.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be true.");
            Site.Assert.IsNotNull(item.Email.Body, "The Body element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsNotNull(item.Email.Body.Data, "The Data element in expected rights-managed e-mail message should not be null.");
            string originalContent = item.Email.Body.Data;
            #endregion

            #region The client logs on User2's account, calls SmartReply method with the same TemplateID in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsNotNull(item.Email.Body, "The Body element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsNotNull(item.Email.Body.Data, "The Data element in expected rights-managed e-mail message should not be null.");
            string repliedContent = item.Email.Body.Data;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R261");

            // Verify MS-ASRM requirement: MS-ASRM_R261
            // Check if the content of the original message is included within the replied message
            // The original e-mail item has ExportAllowed and EditAllowed elements set to TRUE, and composemail:ReplaceMime is not included in the SmartReply command request with the same TemplateID in the original message,
            // If the new message contains the content of original message, this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                repliedContent.Contains(originalContent),
                261,
                @"[In Handling SmartForward and SmartReply Requests] If ExportAllowed and EditAllowed are set to TRUE, composemail:ReplaceMime is set to FALSE, the server will send the message including the original message as inline content when it uses the same TemplateID with the original message.");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC23_SmartReply_Export_True_Edit_True_NoReplaceMime_DifferentTemplate
        /// <summary>
        /// This test case is designed to test when ExportAllowed and EditAllowed is set to true and without ReplaceMime element included in SmartReply request, the server will send the message including the original message as inline content when it uses a different TemplateID with the original message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC23_SmartReply_Export_True_Edit_True_NoReplaceMime_DifferentTemplate()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be true.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be true.");
            Site.Assert.IsNotNull(item.Email.Body, "The Body element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsNotNull(item.Email.Body.Data, "The Data element in expected rights-managed e-mail message should not be null.");
            string originalContent = item.Email.Body.Data;
            #endregion

            #region The client logs on User2's account, calls Settings command to get another templateID with reply rights allowed.
            templateID = this.GetTemplateID("MSASRMReplyAll_NotAllowedTemplate");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with a new TemplateID in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsNotNull(item.Email.Body, "The Body element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsNotNull(item.Email.Body.Data, "The Data element in expected rights-managed e-mail message should not be null.");
            string repliedContent = item.Email.Body.Data;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R262");

            // Verify MS-ASRM requirement: MS-ASRM_R262
            // The original e-mail item has ExportAllowed and EditAllowed elements set to TRUE, and composemail:ReplaceMime is not included in the SmartReply command request with a different TemplateID in the original message,
            // If the new message contains the content of original message, this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                repliedContent.Contains(originalContent),
                262,
                @"[In Handling SmartForward and SmartReply Requests] If ExportAllowed and EditAllowed are set to TRUE, composemail:ReplaceMime is set to FALSE, the server will send the message including the original message as inline content when it uses a different TemplateID with the original message.");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC24_SmartReply_Export_False_Edit_True_NoReplaceMime_SameTemplate
        /// <summary>
        /// This test case is designed to test when ExportAllowed is set to false, EditAllowed is set to true and without ReplaceMime element included in SmartReply request, the server will send the message, including the original message as inline content when it uses the same TemplateID with the original message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC24_SmartReply_Export_False_Edit_True_NoReplaceMime_SameTemplate()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed except Export rights.
            string templateID = this.GetTemplateID("MSASRM_Export_NotAllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be true.");
            Site.Assert.IsNotNull(item.Email.Body, "The Body element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsNotNull(item.Email.Body.Data, "The Data element in expected rights-managed e-mail message should not be null.");
            string originalContent = item.Email.Body.Data;
            #endregion

            #region The client logs on User2's account, calls SmartReply method with the same TemplateID in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsNotNull(item.Email.Body, "The Body element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsNotNull(item.Email.Body.Data, "The Data element in expected rights-managed e-mail message should not be null.");
            string repliedContent = item.Email.Body.Data;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R265");

            // Verify MS-ASRM requirement: MS-ASRM_R265
            // The original e-mail item has ExportAllowed set to false, EditAllowed set to true and composemail:ReplaceMime is not included in the SmartReply command request with the same TemplateID in the original message,
            // If the new message contains the content of original message, this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                repliedContent.Contains(originalContent),
                265,
                @"[In Handling SmartForward and SmartReply Requests] If ExportAllowed and composemail:ReplaceMime are set to FALSE, EditAllowed is set to TRUE, the server will send the message, including the original message as inline content when it uses the same TemplateID with the original message.");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC25_SmartReply_Export_False_Edit_True_NoReplaceMime_DifferentTemplate
        /// <summary>
        /// This test case is designed to test when ExportAllowed is set to false, EditAllowed is set to true and without ReplaceMime element included in SmartReply request, the server will send the message, including the original message as an attachment when it uses a different TemplateID with the original message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC25_SmartReply_Export_False_Edit_True_NoReplaceMime_DifferentTemplate()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed except Export rights.
            string templateID = this.GetTemplateID("MSASRM_Export_NotAllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account, calls Settings command to get another templateID with reply rights allowed.
            templateID = this.GetTemplateID("MSASRMReplyAll_NotAllowedTemplate");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with a new TemplateID in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R266");

            // Verify MS-ASRM requirement: MS-ASRM_R266
            // The original e-mail item has ExportAllowed element set to FALSE and EditAllowed element set to TRUE, and composemail:ReplaceMime is not included in the SmartReply command request with a different TemplateID in the original message,
            // If attachment is not null, this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item.Email.Attachments,
                266,
                @"[In Handling SmartForward and SmartReply Requests] If ExportAllowed and composemail:ReplaceMime are set to FALSE, EditAllowed is set to TRUE, the server will send the message, including the original message as an attachment when it uses a different TemplateID with the original message.");
            #endregion
        }
        #endregion

        #region MSASRM_S05_TC26_SmartReply_Export_False_Edit_False_NoReplaceMime_DifferentTemplate
        /// <summary>
        /// This test case is designed to test when ExportAllowed and EditAllowed are set to false and without ReplaceMime element included in SmartReply request, the server will send the message, including the original message as an attachment when it uses a different TemplateID with the original message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S05_TC26_SmartReply_Export_False_Edit_False_NoReplaceMime_DifferentTemplate()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed except Edit and Export rights.
            string templateID = this.GetTemplateID("MSASRM_EditExport_NotAllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be false.");
            #endregion

            #region The client logs on User2's account, calls Settings command to get another templateID with reply rights allowed.
            templateID = this.GetTemplateID("MSASRMReplyAll_NotAllowedTemplate");
            #endregion

            #region The client logs on User2's account, calls SmartReply method with a new TemplateID in request to reply the received email to User1, and checks the response of SmartReply command.
            string replySubject = string.Format("Re: {0}", subject);
            string replyMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain),
                string.Empty,
                string.Empty,
                replySubject,
                Common.GenerateResourceName(Site, "reply: body"));

            SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, replyMime);
            replyRequest.RequestData.TemplateID = templateID;
            SmartReplyResponse smartReplyResponse = this.ASRMAdapter.SmartReply(replyRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartReplyResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            TestSuiteBase.AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.InboxCollectionId, replySubject);
            #endregion

            #region The client logs on User1's account, calls Sync command to synchronize changes of Inbox folder in User1's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(replySubject, this.UserOneInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R268");

            // Verify MS-ASRM requirement: MS-ASRM_R268
            // The original e-mail item has ExportAllowed and EditAllowed elements set to FALSE, and composemail:ReplaceMime is not included in the SmartReply command request with a different TemplateID in the original message,
            // If attachment is not null, this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item.Email.Attachments,
                268,
                @"[In Handling SmartForward and SmartReply Requests] If ExportAllowed, EditAllowed and composemail:ReplaceMime are set to FALSE, the server will send the message, including the original message as an attachment when it uses a different TemplateID with the original message.");
            #endregion
        }
        #endregion
    }
}
