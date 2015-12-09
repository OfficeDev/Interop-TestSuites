namespace Microsoft.Protocols.TestSuites.MS_ASRM
{
    using System.Globalization;
    using Common.DataStructures;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to test the ItemOperations command.
    /// </summary>
    [TestClass]
    public class S02_ItemOperations : TestSuiteBase
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

        #region MSASRM_S02_TC01_ItemOperations_RightsManagedEmailMessages
        /// <summary>
        /// This test case is designed to call ItemOperations command to fetch a rights-managed e-mail message with or without RemoveRightsManagementProtection element.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S02_TC01_ItemOperations_RightsManagedEmailMessages()
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
            Site.Assert.IsNotNull(item.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            Site.Assert.IsTrue(item.Email.RightsManagementLicense.ExportAllowed, "The ExportAllowed element in expected rights-managed e-mail message should be true.");
            #endregion

            #region The client logs on User2's account, calls ItemOperations command which contains the Fetch element without RemoveRightsManagementProtection element.
            Request.ItemOperationsFetch fetch = new Request.ItemOperationsFetch
            {
                CollectionId = this.UserTwoInformation.InboxCollectionId,
                Store = SearchName.Mailbox.ToString(),
                ServerId = item.ServerId,
                Options = new Request.ItemOperationsFetchOptions
                {
                    Items = new object[] { true },
                    ItemsElementName = new Request.ItemsChoiceType5[] { Request.ItemsChoiceType5.RightsManagementSupport }
                }
            };
            ItemOperationsRequest itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { fetch });
            ItemOperationsStore itemOperationsStore = this.ASRMAdapter.ItemOperations(itemOperationsRequest);
            Site.Assert.AreEqual<int>(1, itemOperationsStore.Items.Count, "There should be only 1 item fetched in ItemOperations command response.");
            ItemOperations itemOperations = itemOperationsStore.Items[0];
            Site.Assert.IsNotNull(itemOperations, "The returned item should not be null.");
            Site.Assert.IsNotNull(itemOperations.Email, "The expected rights-managed e-mail message should not be null.");
            Site.Assert.IsNull(itemOperations.Email.Attachments, "The Attachments element in expected rights-managed e-mail message should be null.");
            Site.Assert.IsNotNull(itemOperations.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            #endregion

            #region The client logs on User2's account, calls ItemOperations command which contains the Fetch element with RemoveRightsManagementProtection element.
            fetch.RemoveRightsManagementProtection = string.Empty;
            itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { fetch });
            itemOperationsStore = this.ASRMAdapter.ItemOperations(itemOperationsRequest);
            Site.Assert.AreEqual<int>(1, itemOperationsStore.Items.Count, "There should be only 1 item fetched in ItemOperations command response.");
            itemOperations = itemOperationsStore.Items[0];
            Site.Assert.IsNotNull(itemOperations, "The returned item should not be null.");
            Site.Assert.IsNotNull(itemOperations.Email, "The expected rights-managed e-mail message should not be null.");
            Site.Assert.IsNull(itemOperations.Email.Attachments, "The Attachments element in expected rights-managed e-mail message should be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R91");

            // Verify MS-ASRM requirement: MS-ASRM_R91
            // If the response contains RightsManagementLicense element as null, represents the rights management protection is removed from the e-mail.
            Site.CaptureRequirementIfIsNull(
                itemOperations.Email.RightsManagementLicense,
                91,
                @"[In RemoveRightsManagementProtection] If present, it[RemoveRightsManagementProtection] indicates that the client is removing the IRM protection from the e-mail.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R94");

            // Verify MS-ASRM requirement: MS-ASRM_R94
            // If the response contains RightsManagementLicense element as null, represents the rights management protection is removed from the e-mail.
            Site.CaptureRequirementIfIsNull(
                itemOperations.Email.RightsManagementLicense,
                94,
                @"[In RemoveRightsManagementProtection] IRM protection can only be removed from e-mail messages when the ExportAllowed element is set to TRUE.");
            #endregion
        }
        #endregion
    }
}