namespace Microsoft.Protocols.TestSuites.MS_ASRM
{
    using System.Globalization;
    using System.Xml;
    using Common.DataStructures;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the Settings, SendMail and Sync commands.
    /// </summary>
    [TestClass]
    public class S01_Settings_SendMail_Sync : TestSuiteBase
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

        #region MSASRM_S01_TC01_Sync_RightsManagedEmailMessages
        /// <summary>
        /// This test case is designed to call Sync command to synchronize a rights-managed e-mail message with different values of RightsManagementSupport element.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S01_TC01_Sync_RightsManagedEmailMessages()
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
            Site.Assert.IsNull(item.Email.Attachments, "The Attachments element in expected rights-managed e-mail message should be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R127");

            // Verify MS-ASRM requirement: MS-ASRM_R127
            // If the RightsManagementLicense element is not null, represents the message has IRM protection.
            Site.CaptureRequirementIfIsNotNull(
                item.Email.RightsManagementLicense,
                127,
                @"[In RightsManagementSupport] If the value of this element[RightsManagementSupport] is TRUE (1), the server will decompress rights-managed email messages before sending them to the client, as specified in section 3.2.4.3. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R347");

            // Verify MS-ASRM requirement: MS-ASRM_R347
            // If the RightsManagementLicense element is not null, represents the message has IRM protection.
            Site.CaptureRequirementIfIsNotNull(
                item.Email.RightsManagementLicense,
                347,
                @"[In RightsManagementSupport] If the value of this element[RightsManagementSupport] is TRUE (1), the server will decrypt rights-managed email messages before sending them to the client, as specified in section 3.2.4.3. ");

            XmlElement lastRawResponse = (XmlElement)this.ASRMAdapter.LastRawResponseXml;
            string contentExpiryDate = TestSuiteHelper.GetElementInnerText(lastRawResponse, "RightsManagementLicense", "ContentExpiryDate", subject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R23");

            // Verify MS-ASRM requirement: MS-ASRM_R23
            Site.CaptureRequirementIfAreEqual<string>(
                "9999-12-30T23:59:59.999Z",
                contentExpiryDate,
                23,
                @"[In ContentExpiryDate] The ContentExpiryDate element is set to ""9999-12-30T23:59:59.999Z"" if the rights management license has no expiration date set.");
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to false to synchronize changes of Inbox folder in User2's mailbox, and gets the neither decompressed nor decrypted rights-managed e-mail message.
            item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, false, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.Attachments, "The Attachments element in expected rights-managed e-mail message should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R335");

            // Verify MS-ASRM requirement: MS-ASRM_R335
            // If the RightsManagementLicense element is null, represents the message has no IRM protection.
            Site.CaptureRequirementIfIsNull(
                item.Email.RightsManagementLicense,
                335,
                @"[In RightsManagementSupport] If the value is FALSE (0), the server will not decompress rights-managed email messages before sending them to the client. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R348");

            // Verify MS-ASRM requirement: MS-ASRM_R348
            // If the RightsManagementLicense element is null, represents the message has no IRM protection.
            Site.CaptureRequirementIfIsNull(
                item.Email.RightsManagementLicense,
                348,
                @"[In RightsManagementSupport] If the value is FALSE (0), the server will not decrypt rights-managed email messages before sending them to the client. ");

            #endregion

            #region The client logs on User2's account, calls Sync command without the RightsManagementSupport element in a request message to synchronize changes of Inbox folder in User2's mailbox, and gets the neither decompressed nor decrypted rights-managed e-mail message.
            item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, null, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.Attachments, "The Attachments element in expected rights-managed e-mail message should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R128");

            // Verify MS-ASRM requirement: MS-ASRM_R128
            // If the response contains RightsManagementLicense element as null, and Attachments element not null, just as the same response when RightsManagementSupport is set to false, this requirement can be verified.
            Site.CaptureRequirementIfIsNull(
                item.Email.RightsManagementLicense,
                128,
                @"[In RightsManagementSupport] If the RightsManagementSupport element is not included in a request message, a default value of FALSE is assumed.");
            #endregion
        }
        #endregion

        #region MSASRM_S01_TC02_Sync_Owner_RightsManagedEmailMessages
        /// <summary>
        /// This test case is designed to test the Owner element.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S01_TC02_Sync_Owner_RightsManagedEmailMessages()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights denied except view rights.
            string templateID = this.GetTemplateID("MSASRM_View_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, switches to User2, and call FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, true, null);
            #endregion

            #region The client logs on User2's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of Inbox folder in User2's mailbox, and gets the decompressed and decrypted rights-managed e-mail message, checks the Owner element.
            Sync item = this.SyncEmail(subject, this.UserTwoInformation.InboxCollectionId, true, true);
            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R330");

            // Verify MS-ASRM requirement: MS-ASRM_R330
            Site.CaptureRequirementIfIsFalse(
                item.Email.RightsManagementLicense.Owner,
                330,
                @"[In Owner] if the value is FALSE (0), the user is not the owner of the e-mail message.");

            #endregion

            #region The client logs on User1's account, calls Sync command with RightsManagementSupport element set to true to synchronize changes of SentItems folder in User1's mailbox, and gets the decompressed and decrypted rights-managed e-mail message, checks the Owner element.
            this.SwitchUser(this.UserOneInformation, false);
            item = this.SyncEmail(subject, this.UserOneInformation.SentItemsCollectionId, true, true);

            Site.Assert.IsNotNull(item, "The returned item should not be null.");
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R72");

            // Verify MS-ASRM requirement: MS-ASRM_R72
            Site.CaptureRequirementIfIsTrue(
                item.Email.RightsManagementLicense.Owner,
                72,
                @"[In Owner] If the value is TRUE (1), the user is the owner of the e-mail message.");

            Site.Assert.AreEqual<string>(Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain).ToUpper(CultureInfo.CurrentCulture), item.Email.RightsManagementLicense.ContentOwner.ToUpper(CultureInfo.CurrentCulture), "The value of ContentOwner element should be equal to the User1's e-mail address.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R31");

            // Verify MS-ASRM requirement: MS-ASRM_R31
            Site.CaptureRequirementIfIsTrue(
                item.Email.RightsManagementLicense.Owner,
                31,
                @"[In ContentOwner] The Owner element is set to TRUE for the user specified by the ContentOwner element.");
            #endregion
        }
        #endregion

        #region MSASRM_S01_TC03_Settings_InvalidXMLBody_ActiveSyncVersionNot141
        /// <summary>
        /// This test case is designed to test that the server considers the XML body of the command request to be invalid when ActiveSync version is not equal to 14.1.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S01_TC03_Settings_InvalidXMLBody_ActiveSyncVersionNot141()
        {
            Site.Assume.AreNotEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Implementation does consider the XML body of the command request to be invalid, if the protocol version specified by in the command request is not 14.1.");
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Implementation does consider the XML body of the command request to be invalid, if the protocol version specified by in the command request is not 16.0.");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Implementation does consider the XML body of the command request to be invalid, if the protocol version specified by in the command request is not 16.1.");

            #region The client logs on User1's account, calls Settings command and checks the response of Settings command.

            if (Common.IsRequirementEnabled(418, this.Site))
            {
                SettingsResponse settingsResponse = this.ASRMAdapter.Settings();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R418");

                // Verify MS-ASRM requirement: MS-ASRM_R418
                // The value of Status element of Settings response could not be 1, which means the operation is unsuccessful.
                Site.CaptureRequirementIfAreNotEqual<string>(
                    "1",
                    settingsResponse.ResponseData.Status,
                    418,
                    @"[In Appendix B: Product Behavior]Implementation does consider the XML body of the command request to be invalid, if the protocol version that is specified by the command request does not support the XML elements that are defined for this protocol. (Exchange Server 2010 and above follow this behavior.)");
            }

            #endregion
        }
        #endregion
    }
}