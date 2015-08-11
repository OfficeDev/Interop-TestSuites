namespace Microsoft.Protocols.TestSuites.MS_ASRM
{
    using System;
    using System.Globalization;
    using Common.DataStructures;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;    

    /// <summary>
    /// This scenario is designed to test the SmartForward command.
    /// </summary>
    [TestClass]
    public class S04_SmartForward : TestSuiteBase
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

        #region MSASRM_S04_TC01_SmartForward_EditAllowed_False_NoReplaceMime
        /// <summary>
        /// This test case is designed to test when EditAllowed is set to false and composemail:ReplaceMime is not present in a SmartForward request, the server will add the original rights-managed email message as an attachment to the new message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S04_TC01_SmartForward_EditAllowed_False_NoReplaceMime()
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
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ForwardAllowed, "The ForwardAllowed element in expected rights-managed e-mail message should be true.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ModifyRecipientsAllowed, "The ModifyRecipientsAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account, calls SmartForward method without ReplaceMime in request to forward the received email to User3.
            string forwardSubject = string.Format("FW: {0}", subject);
            string forwardMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                string.Empty,
                string.Empty,
                forwardSubject,
                Common.GenerateResourceName(Site, "forward: body"));

            SmartForwardRequest forwardRequest = TestSuiteHelper.CreateSmartForwardRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, forwardMime);
            forwardRequest.RequestData.TemplateID = templateID;
            SmartForwardResponse smartForwardResponse = this.ASRMAdapter.SmartForward(forwardRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartForwardResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and get the rights-managed e-mail message
            this.SwitchUser(this.UserThreeInformation, true);
            TestSuiteBase.AddCreatedItemToCollection(this.UserThreeInformation, this.UserThreeInformation.InboxCollectionId, forwardSubject);
            item = this.SyncEmail(forwardSubject, this.UserThreeInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R62");

            // Verify MS-ASRM requirement: MS-ASRM_R62
            // The original e-mail item has ForwardAllowed element set to true.
            Site.CaptureRequirementIfIsNotNull(
                item.Email,
                62,
                @"[In ForwardAllowed] The value is TRUE (1) if the user can forward the e-mail message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R359");

            // Verify MS-ASRM requirement: MS-ASRM_R359
            // The SmartForward command request included the same TemplateID element with that in SendMail command request.
            // The original e-mail item has ModifyRecipientsAllowed element set to true, to allow the recipient list modified, and so that the server creates the forwarded message.
            Site.CaptureRequirementIfIsNotNull(
                item.Email,
                359,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartForward command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template: If the TemplateID element is included in the command request, the server does the following:] If the original message is protected and the specified TemplateID value is the same as the TemplateID value on the original message, the server proceeds to step 4 [The server compares the recipients (1) on the original message to the recipients (1) sent by the client within the new message. The server verifies that the recipient (1) list on the new message aligns with the granted permissions, as specified in the following table[section 3.2.5.1]. If permissions allow it, the server creates the forwarded message].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R363");

            // Verify MS-ASRM requirement: MS-ASRM_R363
            // The SmartForward command request included the TemplateID element returned in Settings command response.
            // The original e-mail item has ModifyRecipientsAllowed element set to true, to allow the recipient list modified, and so that the server creates the forwarded message.
            Site.CaptureRequirementIfIsNotNull(
                item.Email,
                363,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartForward command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template: If the TemplateID element is included in the command request, the server does the following:] If the TemplateID value corresponds to a template on the server, the server proceeds to step 4[The server compares the recipients (1) on the original message to the recipients (1) sent by the client within the new message. The server verifies that the recipient (1) list on the new message aligns with the granted permissions, as specified in the following table. If permissions allow it, the server creates the forwarded message].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R384");

            // Verify MS-ASRM requirement: MS-ASRM_R384
            // The original e-mail item has ForwardAllowed element set to true.
            Site.CaptureRequirementIfIsNotNull(
                item.Email,
                384,
                @"[In Handling SmartForward and SmartReply Requests] If ForwardAllowed is set to TRUE, there is no restrictions for server to perform forward enforcement.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R267");

            // Verify MS-ASRM requirement: MS-ASRM_R267
            // The original e-mail item has EditAllowed element and ExportAllowed element set to false, with the ReplaceMime element is not include in SmartForward command request, and includes the same TemplateID element with that in SendMail command request.
            // If the response contains Attachments element as not null, represents the server will add the original rights-managed email message as an attachment to the new message.
            Site.CaptureRequirementIfIsNotNull(
                item.Email.Attachments,
                267,
                @"[In Handling SmartForward and SmartReply Requests] If ExportAllowed, EditAllowed and composemail:ReplaceMime are set to FALSE, the server will send the message, including the original message as an attachment when it uses the same TemplateID with the original message.");

            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R41");

            // Verify MS-ASRM requirement: MS-ASRM_R41
            // The original e-mail item has EditAllowed element set to false, and the ReplaceMime element is not include in SmartForward command request.
            // If the response contains Attachments element as not null, represents the server will add the original rights-managed email message as an attachment to the new message.
            Site.CaptureRequirementIfIsNotNull(
                item.Email.Attachments,
                41,
                @"[In EditAllowed] When EditAllowed is set to FALSE and composemail:ReplaceMime ([MS-ASCMD] section 2.2.3.135) is not present in a SmartForward request, the server will add the original rights-managed email message as an attachment to the new message.");

            #endregion
        }
        #endregion

        #region MSASRM_S04_TC02_SmartForward_EditAllowed_False_ReplaceMime
        /// <summary>
        /// This test case is designed to test when EditAllowed is set to false and composemail:ReplaceMime is present in a SmartForward request, the server will not attach the original rights-managed email message as an attachment.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S04_TC02_SmartForward_EditAllowed_False_ReplaceMime()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with EditAllowed set to false.
            string templateID = this.GetTemplateID("MSASRM_EditExport_NotAllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.EditAllowed, "The EditAllowed element in expected rights-managed e-mail message should be false.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ForwardAllowed, "The ForwardAllowed element in expected rights-managed e-mail message should be true.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ModifyRecipientsAllowed, "The ModifyRecipientsAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account, calls SmartForward method with ReplaceMime in request to forward the received email to User3.
            string forwardSubject = string.Format("FW: {0}", subject);
            string forwardMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                string.Empty,
                string.Empty,
                forwardSubject,
                Common.GenerateResourceName(Site, "forward: body"));

            SmartForwardRequest forwardRequest = TestSuiteHelper.CreateSmartForwardRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, forwardMime);
            forwardRequest.RequestData.TemplateID = templateID;
            forwardRequest.RequestData.ReplaceMime = string.Empty;
            SmartForwardResponse smartForwardResponse = this.ASRMAdapter.SmartForward(forwardRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartForwardResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and get the rights-managed e-mail message
            this.SwitchUser(this.UserThreeInformation, true);
            TestSuiteBase.AddCreatedItemToCollection(this.UserThreeInformation, this.UserThreeInformation.InboxCollectionId, forwardSubject);
            item = this.SyncEmail(forwardSubject, this.UserThreeInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R42");

            // Verify MS-ASRM requirement: MS-ASRM_R42
            // The original e-mail item has EditAllowed element set to false, and the ReplaceMime element is included in SmartForward command request.
            // If the response contains Attachments element as null, represents the server will not attach the original rights-managed email message as an attachment to the new message.
            Site.CaptureRequirementIfIsNull(
                item.Email.Attachments,
                42,
                @"[In EditAllowed] Conversely, if [EditAllowed is set to FALSE and]composemail:ReplaceMime is present[in a SmartForward command], the server will not attach the original rights-managed email message as an attachment.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R259");

            // Verify MS-ASRM requirement: MS-ASRM_R259
            // The original e-mail item has EditAllowed element set to false, with the ReplaceMime element included in SmartForward command request, and includes the same TemplateID element with that in SendMail command request.
            Site.CaptureRequirementIfIsNotNull(
                item.Email,
                259,
                @"[In Handling SmartForward and SmartReply Requests] If ExportAllowed and EditAllowed are set to FALSE, composemail:ReplaceMime is set to TRUE, the server will send the message when it uses the same TemplateID with the original message.");

            #endregion
        }
        #endregion

        #region MSASRM_S04_TC03_SmartForward_ExportAllowed_True
        /// <summary>
        /// This test case is designed to test when the user forwards the e-mail message and the ExportAllowed is true, the user can remove the IRM protection.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S04_TC03_SmartForward_ExportAllowed_True()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with ExportAllowed set to true.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account, calls SmartForward method with TemplateID set to "00000000-0000-0000-0000-000000000000" in request to forward the received email to User3.
            string forwardSubject = string.Format("FW: {0}", subject);
            string forwardMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                string.Empty,
                string.Empty,
                forwardSubject,
                Common.GenerateResourceName(Site, "forward: body"));

            SmartForwardRequest forwardRequest = TestSuiteHelper.CreateSmartForwardRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, forwardMime);
            forwardRequest.RequestData.TemplateID = "00000000-0000-0000-0000-000000000000";
            SmartForwardResponse smartForwardResponse = this.ASRMAdapter.SmartForward(forwardRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartForwardResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and gets the e-mail message which removed the IRM protection.
            this.SwitchUser(this.UserThreeInformation, true);
            TestSuiteBase.AddCreatedItemToCollection(this.UserThreeInformation, this.UserThreeInformation.InboxCollectionId, forwardSubject);
            Sync forwardItem = this.SyncEmail(forwardSubject, this.UserThreeInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(forwardItem, "The returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R48");

            // Verify MS-ASRM requirement: MS-ASRM_R48
            // The original e-mail item has ExportAllowed element set to true, and TemplateID in SmartForward request set to "00000000-0000-0000-0000-000000000000"
            // If the forwarded e-mail message contains RightsManagementLicense element as null, represents IRM protection is removed from the forwarded e-mail message.
            Site.CaptureRequirementIfIsNull(
                forwardItem.Email.RightsManagementLicense,
                48,
                @"[In ExportAllowed] The value is TRUE (1) if the user can remove the IRM protection when the user forwards the e-mail message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R51");

            // Verify MS-ASRM requirement: MS-ASRM_R51
            // If the RightsManagementLicense element is null, represent the IRM protection is removed from the message.
            Site.CaptureRequirementIfIsNull(
                forwardItem.Email.RightsManagementLicense,
                51,
                @"[In ExportAllowed] If both of the conditions[The original rights policy template has the ExportAllowed element set to TRUE; The TemplateID on the new message is set to the ""No Restriction"" template (TemplateID value ""00000000-0000-0000-0000-000000000000"")] are true, the IRM protection is removed from the outgoing message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R309");

            // Verify MS-ASRM requirement: MS-ASRM_R309
            // If the RightsManagementLicense element is null, represent the IRM protection is removed from the message.
            Site.CaptureRequirementIfIsNull(
                forwardItem.Email.RightsManagementLicense,
                309,
                @"[In TemplateID (SendMail, SmartForward, SmartReply)] If a rights-managed e-mail message is forwarded using the SmartForward command, the IRM protection is removed from the outgoing message if the following conditions are true:
 The original rights policy template has the ExportAllowed element set to TRUE.
 The TemplateID on the new message is set to the ""No Restriction"" template (TemplateID value ""00000000-0000-0000-0000-000000000000"").");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R52");

            // Verify MS-ASRM requirement: MS-ASRM_R52
            // If the RightsManagementLicense element is not null, represent original message retains the IRM protection.
            Site.CaptureRequirementIfIsNotNull(
                item.Email.RightsManagementLicense,
                52,
                @"[In ExportAllowed] [If The original rights policy template has the ExportAllowed element set to TRUE; The TemplateID on the new message is set to the ""No Restriction"" template (TemplateID value ""00000000-0000-0000-0000-000000000000"")]The original message retains its IRM protection.");

            #endregion
        }
        #endregion

        #region MSASRM_S04_TC04_SmartForward_ForwardAllowed_False
        /// <summary>
        /// This test case is designed to test when the user forwards the e-mail message and the ForwardAllowed is false, the user cannot forward the message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S04_TC04_SmartForward_ForwardAllowed_False()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with ForwardAllowed set to false.
            string templateID = this.GetTemplateID("MSASRM_View_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsFalse(item.Email.RightsManagementLicense.ForwardAllowed, "The ForwardAllowed element in expected rights-managed e-mail message should be false.");
            #endregion

            #region The client logs on User2's account, calls SmartForward method to forward the received email to User3.
            string forwardSubject = string.Format("FW: {0}", subject);
            string forwardMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                string.Empty,
                string.Empty,
                forwardSubject,
                Common.GenerateResourceName(Site, "forward: body"));

            SmartForwardRequest forwardRequest = TestSuiteHelper.CreateSmartForwardRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, forwardMime);
            forwardRequest.RequestData.TemplateID = templateID;
            SmartForwardResponse smartForwardResponse = this.ASRMAdapter.SmartForward(forwardRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R386");

            // Verify MS-ASRM requirement: MS-ASRM_R386
            Site.CaptureRequirementIfAreEqual<string>(
                "172",
                smartForwardResponse.ResponseData.Status,
                386,
                @"[In Handling SmartForward and SmartReply Requests] If ForwardAllowed is set to FALSE, SmartForward command requests are restricted and return a composemail:Status value of 172.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R201");

            // Verify MS-ASRM requirement: MS-ASRM_R201
            Site.CaptureRequirementIfAreEqual<string>(
                "172",
                smartForwardResponse.ResponseData.Status,
                201,
                @"[In Enforcing Rights Policy Template Settings] The server returns Status value 172 if the client tries to perform an action on a rights-managed e-mail message that is prohibited by the rights policy template.");

            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and checks if the message arrives.
            this.SwitchUser(this.UserThreeInformation, true);
            item = this.SyncEmail(forwardSubject, this.UserThreeInformation.InboxCollectionId, true, false);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R328");

            // Verify MS-ASRM requirement: MS-ASRM_R328
            // ForwardAllowed element is verified as False in previous step.
            // If the item is null, represent the user cannot forward the e-mail message.
            Site.CaptureRequirementIfIsNull(
                item,
                328,
                @"[In ForwardAllowed] otherwise[the user cannot forward the e-mail message], FALSE (0).");

            #endregion
        }
        #endregion

        #region MSASRM_S04_TC05_SmartForward_Status171
        /// <summary>
        /// This test case is designed to test when the request included an invalid TemplateID value, the server returns status value 171 in a SmartForward command response.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S04_TC05_SmartForward_Status171()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed.
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

            #region The client logs on User2's account, calls SmartForward method with an invalid TemplateID in request to forward the received email to User3, and checks the response of SmartForward command.
            string forwardSubject = string.Format("FW: {0}", subject);
            string forwardMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                string.Empty,
                string.Empty,
                forwardSubject,
                Common.GenerateResourceName(Site, "forward: body"));

            SmartForwardRequest forwardRequest = TestSuiteHelper.CreateSmartForwardRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, forwardMime);
            forwardRequest.RequestData.TemplateID = "invalidTemplateID";
            SmartForwardResponse smartForwardResponse = this.ASRMAdapter.SmartForward(forwardRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R356");

            // Verify MS-ASRM requirement: MS-ASRM_R356
            Site.CaptureRequirementIfAreEqual<string>(
                "171",
                smartForwardResponse.ResponseData.Status,
                356,
                @"[In Enforcing Rights Policy Template Settings] The server returns Status value 171 in a SmartForward command response if the request included an invalid TemplateID value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R358");

            // Verify MS-ASRM requirement: MS-ASRM_R358
            Site.CaptureRequirementIfAreEqual<string>(
                "171",
                smartForwardResponse.ResponseData.Status,
                358,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartForward command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template: If the TemplateID element is included in the command request, the server does the following:] If the template does not exist on the server, the server fails the request and returns a composemail:Status value of 171.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R362");

            // Verify MS-ASRM requirement: MS-ASRM_R362
            Site.CaptureRequirementIfAreEqual<string>(
                "171",
                smartForwardResponse.ResponseData.Status,
                362,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartForward command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template: If the TemplateID element is included in the command request, the server does the following:] If the TemplateID value does not correspond to a template on the server, the server fails the request and returns a composemail:Status value of 171;");

            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and checks if the message arrives.
            this.SwitchUser(this.UserThreeInformation, true);
            item = this.SyncEmail(forwardSubject, this.UserThreeInformation.InboxCollectionId, true, false);
            Site.Assert.IsNull(item, "The returned item should be null.");
            #endregion
        }
        #endregion

        #region MSASRM_S04_TC06_SmartForward_NotProtected_OriginalMessage_NotProtected
        /// <summary>
        /// This test case is designed to test when the original message being forwarded has no rights management restrictions and no TemplateID element is included in the command request, the server sends the new message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S04_TC06_SmartForward_NotProtected_OriginalMessage_NotProtected()
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

            #region The client logs on User2's account, calls SmartForward method to forward the received email to User3.
            string forwardSubject = string.Format("FW: {0}", subject);
            string forwardMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                string.Empty,
                string.Empty,
                forwardSubject,
                Common.GenerateResourceName(Site, "forward: body"));

            SmartForwardRequest forwardRequest = TestSuiteHelper.CreateSmartForwardRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, forwardMime);
            SmartForwardResponse smartForwardResponse = this.ASRMAdapter.SmartForward(forwardRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartForwardResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and gets the e-mail message
            this.SwitchUser(this.UserThreeInformation, true);
            TestSuiteBase.AddCreatedItemToCollection(this.UserThreeInformation, this.UserThreeInformation.InboxCollectionId, forwardSubject);
            item = this.SyncEmail(forwardSubject, this.UserThreeInformation.InboxCollectionId, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R217");

            // Verify MS-ASRM requirement: MS-ASRM_R217
            // If the message is not null, represents the server sends the new message.
            Site.CaptureRequirementIfIsNotNull(
                item.Email,
                217,
                @"[In Handling SmartForward and SmartReply Requests] When the client sends the server a SmartForward command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template:1.	If no TemplateID element is included in the command request, the server proceeds as follows: If the original message being forwarded has no rights management restrictions, the server proceeds to step 6[The server sends the new message].");

            #endregion
        }
        #endregion

        #region MSASRM_S04_TC07_SmartForward_NotProtected_OriginalMessage_Protected
        /// <summary>
        /// This test case is designed to test when the original message being forwarded has rights management restrictions and no TemplateID element is included in the command request, the server replaces the body of the message with boilerplate text.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S04_TC07_SmartForward_NotProtected_OriginalMessage_Protected()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed.
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

            #region The client logs on User2's account, calls SmartForward method without a TemplateID in request to forward the received email to User3.
            string forwardSubject = string.Format("FW: {0}", subject);
            string forwardMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                string.Empty,
                string.Empty,
                forwardSubject,
                Common.GenerateResourceName(Site, "forward: body"));

            SmartForwardRequest forwardRequest = TestSuiteHelper.CreateSmartForwardRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, forwardMime);
            SmartForwardResponse smartForwardResponse = this.ASRMAdapter.SmartForward(forwardRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartForwardResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and gets the e-mail message
            this.SwitchUser(this.UserThreeInformation, true);
            TestSuiteBase.AddCreatedItemToCollection(this.UserThreeInformation, this.UserThreeInformation.InboxCollectionId, forwardSubject);
            item = this.SyncEmail(forwardSubject, this.UserThreeInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");

            Site.Assert.IsNotNull(item.Email.Attachments, "The Attachments in returned message should not be null.");
            Site.Assert.IsNotNull(item.Email.Attachments.Attachment, "The Attachment in returned message should not be null.");
            Site.Assert.IsTrue(item.Email.Attachments.Attachment.Length >= 1, "There should be at least 1 attachment in returned message");
            Site.Assert.IsNotNull(item.Email.Attachments.Attachment[0], "The first Attachment in returned message should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R219");

            // Verify MS-ASRM requirement: MS-ASRM_R219
            Site.CaptureRequirementIfIsTrue(
                item.Email.Attachments.Attachment[0].DisplayName.EndsWith(".rpmsg", StringComparison.CurrentCulture),
                219,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartForward command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template:
1. If no TemplateID element is included in the command request, the server proceeds as follows: ]If the original message had rights management restrictions, the rights-managed e-mail message is added as an .rpmsg attachment as specified in [MS-OXORMMS].");
            #endregion
        }
        #endregion

        #region MSASRM_S04_TC08_SmartForward_HTTP_Status168
        /// <summary>
        /// This test case is designed to test when the connection to the server does not use SSL, the server fails the request and returns composemail:Status value 168.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S04_TC08_SmartForward_HTTP_Status168()
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

            #region The client logs on User2's account, calls SmartForward method with the TemplateID in request to forward the received email to User3, and checks the response of SmartForward command.
            string forwardSubject = string.Format("FW: {0}", subject);
            string forwardMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                string.Empty,
                string.Empty,
                forwardSubject,
                Common.GenerateResourceName(Site, "forward: body"));

            SmartForwardRequest forwardRequest = TestSuiteHelper.CreateSmartForwardRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, forwardMime);
            forwardRequest.RequestData.TemplateID = templateID;
            SmartForwardResponse smartForwardResponse = this.ASRMAdapter.SmartForward(forwardRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R222");

            // Verify MS-ASRM requirement: MS-ASRM_R222
            Site.CaptureRequirementIfAreEqual<string>(
                "168",
                smartForwardResponse.ResponseData.Status,
                222,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartForward command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template: If any of the following are true, the server fails the request and returns composemail:Status value 168:] The connection to the server does not use SSL.");

            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and checks if the e-mail message arrives
            this.SwitchUser(this.UserThreeInformation, true);
            item = this.SyncEmail(forwardSubject, this.UserThreeInformation.InboxCollectionId, true, false);
            Site.Assert.IsNull(item, "The returned item should be null.");
            #endregion
        }
        #endregion

        #region MSASRM_S04_TC09_SmartForward_Protected_OriginalMessage_NotProtected
        /// <summary>
        /// This test case is designed to test when the original message being forwarded has no rights management restrictions and the TemplateID element is included in the command request, the server protects the new outgoing message with the specified rights policy template.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S04_TC09_SmartForward_Protected_OriginalMessage_NotProtected()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command without the templateID to send a rights-managed e-mail message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(null, false, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the e-mail message.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should be null.");
            #endregion

            #region The client logs on User2's account, calls SmartForward method with the TemplateID in request to forward the received email to User3.
            string forwardSubject = string.Format("FW: {0}", subject);
            string forwardMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                string.Empty,
                string.Empty,
                forwardSubject,
                Common.GenerateResourceName(Site, "forward: body"));

            SmartForwardRequest forwardRequest = TestSuiteHelper.CreateSmartForwardRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, forwardMime);
            forwardRequest.RequestData.TemplateID = templateID;
            SmartForwardResponse smartForwardResponse = this.ASRMAdapter.SmartForward(forwardRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartForwardResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and gets the rights-managed e-mail message.
            this.SwitchUser(this.UserThreeInformation, true);
            TestSuiteBase.AddCreatedItemToCollection(this.UserThreeInformation, this.UserThreeInformation.InboxCollectionId, forwardSubject);
            item = this.SyncEmail(forwardSubject, this.UserThreeInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R224");

            // Verify MS-ASRM requirement: MS-ASRM_R224
            // If the RightsManagementLicense is not null, represents the new message has IRM protection.
            Site.CaptureRequirementIfIsNotNull(
                item.Email.RightsManagementLicense,
                224,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartForward command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template:] If the TemplateID element is included in the command request, the server does the following: If the original message is not protected, the server proceeds to step 5[If the message has a TemplateID element, the server protects the new outgoing message with the specified rights policy template.].");

            #endregion
        }
        #endregion

        #region MSASRM_S04_TC10_SmartForward_Protected_OriginalMessage_Protected_DifferentTemplate
        /// <summary>
        /// This test case is designed to test when the original message is protected and the specified TemplateID value is different from the TemplateID value on the original message, the server verifies that the new TemplateID value exists on the server.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S04_TC10_SmartForward_Protected_OriginalMessage_Protected_DifferentTemplate()
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

            #region The client logs on User2's account, calls Settings command to get another templateID with forward rights allowed.
            templateID = this.GetTemplateID("MSASRM_EditExport_NotAllowedTemplate");
            #endregion

            #region The client logs on User2's account, calls SmartForward method with the new TemplateID in request to forward the received email to User3.
            string forwardSubject = string.Format("FW: {0}", subject);
            string forwardMime = Common.CreatePlainTextMime(
                Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain),
                Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain),
                string.Empty,
                string.Empty,
                forwardSubject,
                Common.GenerateResourceName(Site, "forward: body"));

            SmartForwardRequest forwardRequest = TestSuiteHelper.CreateSmartForwardRequest(this.UserTwoInformation.InboxCollectionId, item.ServerId, forwardMime);
            forwardRequest.RequestData.TemplateID = templateID;
            SmartForwardResponse smartForwardResponse = this.ASRMAdapter.SmartForward(forwardRequest);
            Site.Assert.AreEqual<string>(string.Empty, smartForwardResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");
            #endregion

            #region The client logs on User3's account, calls Sync command to synchronize changes of Inbox folder in User3's mailbox, and gets the rights-managed e-mail message.
            this.SwitchUser(this.UserThreeInformation, true);
            TestSuiteBase.AddCreatedItemToCollection(this.UserThreeInformation, this.UserThreeInformation.InboxCollectionId, forwardSubject);
            item = this.SyncEmail(forwardSubject, this.UserThreeInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R361");

            // Verify MS-ASRM requirement: MS-ASRM_R361
            Site.CaptureRequirementIfAreEqual<string>(
                "MSASRM_EditExport_NotAllowedTemplate",
                item.Email.RightsManagementLicense.TemplateName,
                361,
                @"[In Handling SmartForward and SmartReply Requests] [When the client sends the server a SmartForward command request for a message with a rights policy template, the server MUST do the following to enforce the rights policy template: If the TemplateID element is included in the command request, the server does the following:] If the original message is protected and the specified TemplateID value is different than the TemplateID value on the original message, the server verifies that the new TemplateID value exists on the server.");

            #endregion
        }
        #endregion
    }
}