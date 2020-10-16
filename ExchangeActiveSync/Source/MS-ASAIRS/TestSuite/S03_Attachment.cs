namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test the Attachments element and its sub elements in the AirSyncBase namespace, which is used by the Sync command, Search command and ItemOperations command to identify the data sent by and returned to client.
    /// </summary>
    [TestClass]
    public class S03_Attachment : TestSuiteBase
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

        #region MSASAIRS_S03_TC01_FileReference_ZeroLengthString
        /// <summary>
        /// This case is designed to test if the client includes a zero-length string for the value of the FileReference (Fetch) element in an ItemOperations command request, the server responds with a protocol status error of 15.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S03_TC01_FileReference_ZeroLengthString()
        {
            #region Send a mail with normal attachment.
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.NormalAttachment, subject, body);
            #endregion

            #region Send an ItemOperations request with the value of FileReference element as a zero-length string.
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, null);

            // Send an ItemOperations request with the value of FileReference element as a zero-length string.
            ItemOperationsRequest request = TestSuiteHelper.CreateItemOperationsRequest(this.User2Information.InboxCollectionId, syncItem.ServerId, string.Empty, null, null);

            DataStructures.ItemOperationsStore itemOperationsStore = this.ASAIRSAdapter.ItemOperations(request, DeliveryMethodForFetch.Inline);

            Site.Assert.AreEqual<int>(
                1,
                itemOperationsStore.Items.Count,
                "There should be 1 item in ItemOperations response.");
            #endregion

            #region Verify requirements
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R214");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R214
            Site.CaptureRequirementIfAreEqual<string>(
                "15",
                itemOperationsStore.Items[0].Status,
                214,
                @"[In FileReference (Fetch)] If the client includes a zero-length string for the value of this element [the FileReference (Fetch) element] in an ItemOperations command request, the server responds with a protocol status error of 15.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S03_TC02_NormalAttachment
        /// <summary>
        /// This case is designed to test the method value 1 which specifies the attachment is a normal attachment.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S03_TC02_NormalAttachment()
        {
            #region Send a mail with normal attachment
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.NormalAttachment, subject, body);
            #endregion

            #region Verify requirements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, null);
            this.VerifyMethodElementValue(syncItem.Email, 1);

            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, null, null, null);
            this.VerifyMethodElementValue(itemOperationsItem.Email, 1);

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, null, null, null);
                this.VerifyMethodElementValue(searchItem.Email, 1);
            }

            // According to above steps, requirement MS-ASAIRS_R225 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R225");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R225
            Site.CaptureRequirement(
                225,
                @"[In Method (Attachment)] [The value] 1 [of the Method element] meaning ""Normal attachment"" specifies that the attachment is a normal attachment.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R160");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R160
            Site.CaptureRequirementIfIsNotNull(
                ((Response.AttachmentsAttachment)syncItem.Email.Attachments.Items[0]).ContentId,
                160,
                @"[In ContentId (Attachment)] The ContentId element is an optional child element of the Attachment element (section 2.2.2.7) that contains the unique object ID for an attachment.");

            this.Site.CaptureRequirementIfIsNotNull(
                ((Response.AttachmentsAttachment)syncItem.Email.Attachments.Items[0]).ContentId,
                1343,
                @"[In ContentId (Attachment)] [The ContentId element is an optional child element of the Attachment element (section 2.2.2.7) that contains the unique identifier of the attachment, and] is used to reference the attachment within the item to which the attachment belongs.");

            #endregion
        }
        #endregion

        #region MSASAIRS_S03_TC03_EmbeddedMessageAttachment
        /// <summary>
        /// This case is designed to test the method value 5 which indicates that the attachment is an e-mail message, and that the attachment file has an .eml extension.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S03_TC03_EmbeddedMessageAttachment()
        {
            #region Send a mail with an e-mail messsage attachment
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.EmbeddedAttachment, subject, body);
            #endregion

            #region Verify requirements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, null);
            this.VerifyMethodElementValue(syncItem.Email, 5);

            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, null, null, null);
            this.VerifyMethodElementValue(itemOperationsItem.Email, 5);

            if (Common.IsRequirementEnabled(53, this.Site))
            { 
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, null, null);
                this.VerifyMethodElementValue(searchItem.Email, 5);
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R2299");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R2299
            Site.CaptureRequirementIfIsTrue(
                ((Response.AttachmentsAttachment)syncItem.Email.Attachments.Items[0]).DisplayName.EndsWith(".eml", System.StringComparison.CurrentCultureIgnoreCase),
                2299,
                @"[In Method (Attachment)] [The value] 5 [of the Method element] meaning ""Embedded message"" indicates that the attachment is an e-mail message, and that the attachment file has an .eml extension.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R100298");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R100298
            Site.CaptureRequirementIfIsFalse(
                ((Response.AttachmentsAttachment)syncItem.Email.Attachments.Items[0]).IsInlineSpecified,
                100298,
                @"[In IsInline (Attachment)] If the value[IsInline] is FALSE, then the attachment is not embedded in the message.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S03_TC04_OLEAttachment
        /// <summary>
        /// This case is designed to test the method value 6 which indicates that the attachment is an embedded Object Linking and Embedding (OLE) object.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S03_TC04_OLEAttachment()
        {
            #region Send a mail with an embedded OLE object
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.AttachOLE, subject, body);
            #endregion

            #region Verify requirements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, null);
            this.VerifyMethodElementValue(syncItem.Email, 6);

            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, null, null, null);
            this.VerifyMethodElementValue(itemOperationsItem.Email, 6);

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, null, null);
                this.VerifyMethodElementValue(searchItem.Email, 6);
            }

            // According to above steps, requirement MS-ASAIRS_R230 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R230");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R230
            Site.CaptureRequirement(
                 230,
                 @"[In Method (Attachment)] [The value] 6 [of the Method element] meaning ""Attach OLE"" indicates that the attachment is an embedded Object Linking and Embedding (OLE) object, such as an inline image.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R100299");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R100299
            Site.CaptureRequirementIfIsTrue(
                ((Response.AttachmentsAttachment)syncItem.Email.Attachments.Items[0]).IsInline,
                100299,
                @"[In IsInline (Attachment)] If the value[IsInline] is TRUE, then the attachment is embedded in the message.");
            #endregion
        }
        #endregion

        #region private methods
        /// <summary>
        /// This method is used to verify whether the Method element is the expected value.
        /// </summary>
        /// <param name="email">The email item got from server.</param>
        /// <param name="methodValue">The expected value of Method element .</param>
        private void VerifyMethodElementValue(DataStructures.Email email, byte methodValue)
        {
            Site.Assert.IsNotNull(
                email.Attachments,
                "The Attachments element in response should not be null.");

            Site.Assert.IsNotNull(
                email.Attachments.Items,
                "The Attachment element in response should not be null.");

            Site.Assert.AreEqual<int>(
                1,
                email.Attachments.Items.Length,
                "There should be only one Attachment element in response.");

            Site.Assert.IsNotNull(
                email.Attachments.Items[0],
                "The Attachment element in response should not be null.");

            Site.Assert.AreEqual<byte>(
                methodValue, 
                ((Response.AttachmentsAttachment)email.Attachments.Items[0]).Method, 
                "The value of Method element in response should be equal to the expected value.");
        }
        #endregion
    }
}