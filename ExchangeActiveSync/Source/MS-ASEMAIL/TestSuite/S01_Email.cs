namespace Microsoft.Protocols.TestSuites.MS_ASEMAIL
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Threading;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to test normal e-mail events, including sending an e-mail to server, synchronizing e-mail with server, retrieving e-mail items that match the criteria specified by the client from server, retrieving data from the server for one or more specific e-mail items.
    /// </summary>
    [TestClass]
    public class S01_Email : TestSuiteBase
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

        #region MSASEMAIL_S01_TC01_SyncMail_OnlyMetadata
        /// <summary>
        /// This case is designed to test synchronizing only metadata of e-mail item in inbox folder with server.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC01_SyncMail_OnlyMetadata()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command wihtout including BodyPreference to synchronize the e-mail items with server.
            // Get the new added email item
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);
            #endregion

            #region Verify requirement.
            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R350");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R350
                // If the ConversationIndex element is included in Sync command response, then requirement MS-ASEAMIL_R350 can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    item.Email.ConversationIndex,
                    350,
                    @"[In ConversationIndex] The email2:ConversationIndex element is a required element in server responses.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R649");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R649
                // Since there is no bcc user, if the value of ReceivedAsBcc is null, then requirement MS-ASEAMIL_R649 can be captured.
                Site.CaptureRequirementIfIsNull(
                    item.Email.ReceivedAsBcc,
                    649,
                    @"[In ReceivedAsBcc] The email2:ReceivedAsBcc element is not included in the command response if the value is 0 (zero, meaning FALSE).");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC02_SyncMail_MetadataAndBody
        /// <summary>
        /// This case is designed to test synchronizing metadata and body of e-mail item in inbox folder with server.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC02_SyncMail_MetadataAndBody()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command including BodyPreference element to synchronize the e-mail items with server.
            // Call method Sync to synchronize(add or change) the e-mail items with the server.
            Request.BodyPreference bodyPreference = new Request.BodyPreference
            {
                AllOrNone = false,
                AllOrNoneSpecified = true,
                TruncationSize = 5120,
                TruncationSizeSpecified = true,
                Type = 1
            };

            // Get the new added email item
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, bodyPreference);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);
            #endregion

            #region Verify requirements.
            // If the Data element is included in Sync command response, then requirement MS-ASEAMIL_R1057 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1057");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1057
            Site.CaptureRequirementIfIsNotNull(
                item.Email.Body.Data,
                1057,
                @"[In Body (Airsyncbase Namespace)] [When[airsyncbase:Body] included in a Sync command response ([MS-ASCMD] section 2.2.1.21), a Search command response ([MS-ASCMD] section 2.2.1.16), or an ItemOperations command response ([MS-ASCMD] section 2.2.1.10), the airsyncbase:Body element can contain the following child element: airsyncbase:Data] This element [airsyncbase:Data] is only included if a nonzero airsyncbase:TruncationSize ([MS-ASAIRS] section 2.2.2.40.2) element value was included in the request and the airsyncbase:AllOrNone ([MS-ASAIRS] section 2.2.2.3.2) element value included in the request does not restrict content from being returned in the response.");

            if (Common.IsRequirementEnabled(439, this.Site))
            {
                // The client doesn't set the flag of the email item, so if the Flag element is included in Sync command response, then MS-ASEMAIL_R439 can be captured
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R439");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R439
                Site.CaptureRequirementIfIsTrue(
                    ((XmlElement)this.EMAILAdapter.LastRawResponseXml).InnerXml.Contains("<Flag xmlns=\"Email\" />"),
                    439,
                    @"[In Appendix B: Product Behavior] If no flags are present on the e-mail item, the Flag element is included in the command as an empty container element (that is, <Flag/>). (Exchange Server 2007 SP1 and above follow this behavior.)");
            }

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // If the server responds with a Sync command response and it contains an empty Categories, then MS-ASEMAIL_R318 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R318");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R318
                Site.CaptureRequirementIfIsTrue(
                    item.Email.CategoriesIsInclude && item.Email.Categories.Category == null,
                    318,
                    @"[In Categories] An empty Categories element is included as a child of the Add ([MS-ASCMD] section 2.2.3.7.2) element in a Sync ([MS-ASCMD] section 2.2.1.21) command if no child Category elements have been set on the message.");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC03_SyncChange_NotChangeCategories
        /// <summary>
        /// This case is designed to test if Sync command with Change element does not contain a Categories element, it indicates the categories on the e-mail item have not been changed.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC03_SyncChange_NotChangeCategories()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Categories element is not supported when the ActiveSyncProtocolVersion is 12.1.");

            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag and category for the e-mail item and synchronize it with server.
            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Get the result of adding flag and categories
            SyncStore getChangedResult = this.AddFlagAndCategories(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, item.ServerId);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The email item with subject '{0}' should be found.", emailSubject);
            #endregion

            #region Call Sync command with Change element to update the flag of the e-mail item and synchronize it with server.
            // Update the FlagType to "Flag for follow up"
            Request.Flag newFlag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                CompleteTime = DateTime.Now,
                CompleteTimeSpecified = true,
                DateCompleted = DateTime.Now,
                DateCompletedSpecified = true
            };

            // Update email flag on server
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, true, item.ServerId, newFlag, null);

            // Synchronize the update result with server
            SyncStore updateResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(updateResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The email item with subject '{0}' should be found.", emailSubject);
            #endregion

            #region Verify requirement.
            // If the server responds with a Sync command response and the Sync change command that does not contain a Categories element, then R320 can be captured.
            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R320");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R320
                Site.CaptureRequirementIfIsFalse(
                        item.Email.CategoriesIsInclude,
                        320,
                        @"[In Categories] A Sync command that contains a Change element that does not contain a Categories element indicates that the categories on the message have not changed.");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC04_SyncChange_RemoveCategories
        /// <summary>
        /// This case is designed to test Sync command with Change element containing an empty Categories element to indicate that all categories associated with the item have been removed.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC04_SyncChange_RemoveCategories()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Categories element is not supported when the ActiveSyncProtocolVersion is 12.1.");

            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add category for the e-mail item and synchronize it with server.
            string category = Common.GenerateResourceName(Site, "Category");
            Collection<string> categories = new Collection<string> { category };

            // Get the new added email item
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);

            // Add categories to the email item
            this.UpdateEmail(this.User2Information.InboxCollectionId, syncChangeResult.SyncKey, true, item.ServerId, null, categories);

            SyncStore updateResult = this.SyncChanges(syncChangeResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(updateResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Call Sync command with Change element to remove categories of the e-mail item and synchronize it with server.
            // Remove categories from email on server
            categories.Remove(category);
            this.UpdateEmail(this.User2Information.InboxCollectionId, updateResult.SyncKey, null, item.ServerId, null, categories);

            // Get changes from server
            SyncStore getChangedResult = this.SyncChanges(updateResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirement.
            // If the server responds with a Sync command response and it contains an empty Categories, then MS-ASEMAIL_R319 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R319");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R319
            Site.CaptureRequirementIfIsTrue(
                item.Email.CategoriesIsInclude && item.Email.Categories.Category == null,
                319,
                @"[In Categories] A Sync command that contains a Change element ([MS-ASCMD] section 2.2.3.24) that contains an empty Categories element indicates that all categories associated with the message have been removed.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC05_OnlyReadPropertyChanged
        /// <summary>
        /// This case is designed to test if only the Read element has been changed for the e-mail item, the server will include the Read element as the only child element of the airsync:ApplicationData within the airsync:Change element for an e-mail item in the Sync command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC05_OnlyReadPropertyChanged()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag and categories for the e-mail item and synchronize it with the server.
            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Get the result of adding flag and categories
            SyncStore getChangedResult = this.AddFlagAndCategories(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, item.ServerId);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The email item with subject '{0}' should be found.", emailSubject);
            #endregion

            #region Call Sync command with Change element to update the Read property of the e-mail and synchronize it with server.
            // Update the value of Read property to FALSE
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, false, item.ServerId, null, null);

            SyncStore updateResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(updateResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirements.
            // If the server will include the Read element as the only child element of the airSync:ApplicationData within the airSync:Change element for that e-mail item in the Sync command response, then MS-ASEMAIL_R75 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R75");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R75
            Site.CaptureRequirementIfIsTrue(
                TestSuiteHelper.IsOnlySpecifiedElement((XmlElement)this.EMAILAdapter.LastRawResponseXml, "ApplicationData", "Read"),
                75,
                @"[In Sending E-Mail Changes to the Client] If only the Read flag has changed for an e-mail item, the server MUST include the Read element as the only child element of the airsync:ApplicationData element ([MS-ASCMD] section 2.2.3.11)  within the airsync:Change element ([MS-ASCMD] section 2.2.3.24) for that e-mail item in the Sync command response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R641");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R641
            Site.CaptureRequirementIfIsFalse(
                Convert.ToBoolean(item.Email.Read),
                641,
                @"[In Read] a value of 0 (zero, meaning FALSE) indicates the e-mail message has not been viewed by the current recipient.");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R10604
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R10604");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R10604
            Site.CaptureRequirementIfIsTrue(
                item.Email.ReadIsInclude,
                10604,
                @"[In Sending E-Mail Changes to the Client]  If Read flag changed, Flag properties, Categories properties, Other E-Mail class properties and Non-E-Mail class properties not changed, Server action for e-mail item in Sync commend response is: Send Read flag only.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC06_OnlyFlagPropertiesChanged
        /// <summary>
        /// This case is designed to test if only the Flag properties has been changed for an e-mail item, the server will include the Flag element as the only child element of the airSync:ApplicationData within the airSync:Change element for that e-mail item in the Sync command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC06_OnlyFlagPropertiesChanged()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag and categories for the e-mail item and synchronize it with the server.
            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Get the result of adding flag and categories
            SyncStore getChangedResult = this.AddFlagAndCategories(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, item.ServerId);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Call Sync command with Change element to update the Flag property of the e-mail and synchronize it with server.
            Request.Flag newFlag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                CompleteTime = DateTime.Now,
                CompleteTimeSpecified = true,
                DateCompleted = DateTime.Now,
                DateCompletedSpecified = true
            };

            // Update email flag
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, null, item.ServerId, newFlag, null);

            // Get update result from server 
            this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            #endregion

            #region Verify requirements.
            // If the server includes the Flag element as the only child element of the airSync:ApplicationData within the airSync:Change element for that e-mail item in the Sync command response, 
            // then MS-ASEMAIL_R1019 will be captured.
            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R76");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R76
                Site.CaptureRequirementIfIsTrue(
                    TestSuiteHelper.IsOnlySpecifiedElement((XmlElement)this.EMAILAdapter.LastRawResponseXml, "ApplicationData", "Flag"),
                    76,
                    @"[In Sending E-Mail Changes to the Client] If only Flag properties have changed for an e-mail item, the server MUST include the Flag element as the only child element of the airsync:ApplicationData element within the airsync:Change element for that e-mail item in the Sync command response.");

                // If the server includes the Flag element as the only child element of the airSync:ApplicationData within the airSync:Change element for that e-mail item in the Sync command response, then MS-ASEMAIL_R96 will be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R100");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R100
                Site.CaptureRequirementIfIsTrue(
                    TestSuiteHelper.IsOnlySpecifiedElement((XmlElement)this.EMAILAdapter.LastRawResponseXml, "ApplicationData", "Flag"),
                    100,
                    @"[In Sending E-Mail Changes to the Client] If Flag properties changed, Read flag, Categories properties,  Non-E-Mail class properties and Other E-Mail class properties not changed, Server action for e-mail item in Sync commend response is: Send Flag block only.");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC07_BothReadAndFlagChanged
        /// <summary>
        /// This case is designed to test if both Read element and Flag element have changed for the e-mail item, the server will send only Read element and Flag element to the client for an e-mail item in Sync command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC07_BothReadAndFlagChanged()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag and categories for the e-mail item and synchronize it with the server.
            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Get the result of adding flag and categories
            SyncStore getChangedResult = this.AddFlagAndCategories(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, item.ServerId);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region all Sync command with Change element to update the flag and read properties of the e-mail and synchronize it with server.
            Request.Flag newFlag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                CompleteTime = DateTime.Now,
                CompleteTimeSpecified = true,
                DateCompleted = DateTime.Now,
                DateCompletedSpecified = true
            };

            // Update Flag and Read property
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, false, item.ServerId, newFlag, null);

            // Get server changes 
            SyncStore updateResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(updateResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirement.
            // If the server send Read flag and Flag block, then MS-ASEMAIL_R10612 can be captured.
            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R10612");
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R10612
                Site.CaptureRequirementIfIsTrue(
                    item.Email.FlagIsInclude && item.Email.Flag != null && item.Email.ReadIsInclude,
                    10612,
                    @"[In Sending E-Mail Changes to the Client] If Read flag and Flag properties changed, Categories properties, Other E-Mail class properties and Non-E-Mail class properties not changed, Server action for e-mail item in Sync commend response is: Send Read flag and Flag block");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC08_FlagReadAndOtherEmailClassPropertiesChanged
        /// <summary>
        /// This case is designed to test if Flag element, Read element, and other email class elements have been changed for the e-mail item, the server will send full item airsync:Change to the client for an e-mail item in Sync command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC08_FlagReadAndOtherEmailClassPropertiesChanged()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Categories element is not supported when the ActiveSyncProtocolVersion is 12.1.");

            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag and categories for the e-mail item and synchronize it with the server.
            // Get the email item
            SyncStore store = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(store, emailSubject);

            // Get the result of adding flag and categories
            SyncStore getChangedResult = this.AddFlagAndCategories(this.User2Information.InboxCollectionId, store.SyncKey, item.ServerId);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region all Sync command with Change element to update the flag , categories and read properties of the e-mail and synchronize it with server.
            Request.Flag newFlag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                CompleteTime = DateTime.Now,
                CompleteTimeSpecified = true,
                DateCompleted = DateTime.Now,
                DateCompletedSpecified = true
            };

            string newCategory = Common.GenerateResourceName(Site, "NewCategory");
            Collection<string> newCategories = new Collection<string> { newCategory };

            // Update Read, Flag and Categories properties
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, false, item.ServerId, newFlag, newCategories);

            // Get server changes 
            SyncStore updateResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(updateResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirements.
            // If the server Send Read flag, Flag block and categories block, then MS-ASEMAIL_R104 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R104");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R104
            Site.CaptureRequirementIfIsTrue(
                item.Email.FlagIsInclude && item.Email.Flag != null && item.Email.ReadIsInclude && item.Email.CategoriesIsInclude && item.Email.Categories != null,
                104,
                @"[In Sending E-Mail Changes to the Client]  If Read flag , Flag properties and Categories properties changed, Other E-Mail class properties and Non-E-Mail class properties not changed, Server action for e-mail item in Sync commend response is: Send Read flag, Flag block and categories block.");

            // If MS-ASEMAIL_R104 can be captured successfully, it means server partition email changes into the case "changes to Read flag, Flag properties and other email class properties, so MS-ASEMAIL_R1017 can also be captured."
            if (Common.IsRequirementEnabled(1017, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1017");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1017
                Site.CaptureRequirement(
                    1017,
                    @"[In Appendix B: Product Behavior] Implementation does partition email changes into one or more of the following categories:
Changes to the Read flag (section 2.2.2.47)
Changes to the Flag properties (section 2.2.2.27)
Changes to other E-Mail class properties, such as Subject (section 2.2.2.63)
Changes to non-E-Mail class properties (Exchange Server 2007 Sp1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC09_WithoutAnyEmailClassPropertiesChanged
        /// <summary>
        /// This case is designed to test if Flag element, Read element and other email class elements have not been changed for the e-mail item, the server will send nothing to the client for an e-mail item in Sync command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC09_WithoutAnyEmailClassPropertiesChanged()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag and categories for the e-mail item and synchronize it with the server.
            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Get the result of adding flag and categories
            SyncStore getChangedResult = this.AddFlagAndCategories(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, item.ServerId);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Call Sync command with Change element but without changing any property to synchronize the item.
            SyncRequest syncRequest = TestSuiteHelper.CreateSyncRequest(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            SyncStore syncChangedResult = this.EMAILAdapter.Sync(syncRequest);
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1018");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R81
            // According to the definition of SyncStore, if the count of ChangeElements in syncChangedResult is 0 which means the Collections element in Sync Response is empty,
            // then requirement MS-ASEAMIL_R81 can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                0,
                syncChangedResult.ChangeElements.Count,
                81,
                @"[In Sending E-Mail Changes to the Client] If Read flag, Flag properties, Other E-mail class properties and Non-E-mail class properties not changed, the Server action for e-mail item in Sync command response is: Send nothing to client.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC10_NonBooleanRead
        /// <summary>
        /// This case is designed to test if a non-Boolean value is used for Read element in a Sync command request, the server responds with Status element value of 6 in the Sync command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC10_NonBooleanRead()
        {
            #region Call SendMail command to send the mail to recipient
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to change the Read element with a boolean value
            // Calls Sync method to synchronize the email item on the server
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);

            // Call Sync command with Change element to change the value of Read property
            this.UpdateEmail(this.User2Information.InboxCollectionId, syncChangeResult.SyncKey, true, item.ServerId, null, null);
            #endregion

            #region Verify requirement
            syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);
            Site.Assert.IsNotNull(item.Email.Read, "The Read value should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R640");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R640
            Site.CaptureRequirementIfAreEqual<bool>(
                true,
                item.Email.Read.Value,
                640,
                @"[In Read] A value of 1 (TRUE) indicates the e-mail message has been viewed by the current recipient;");
            #endregion

            #region Calls Sync command with Change element to change the Read element with a non-boolean value
            SyncRequest invalidSyncRequest = TestSuiteHelper.CreateSyncChangeRequest(
                syncChangeResult.SyncKey,
                this.User2Information.InboxCollectionId,
                new Request.SyncCollectionChange { ServerId = item.ServerId });
            SendStringResponse syncStringResponse = this.EMAILAdapter.InvalidSync(
                invalidSyncRequest,
                "<ApplicationData><Read xmlns=\"Email\">NON_BOOL</Read></ApplicationData>",
                "</Change>");
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R643");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R643
            Site.CaptureRequirementIfAreEqual<string>(
                "6",
                TestSuiteHelper.GetStatusCode(syncStringResponse.ResponseDataXML),
                643,
                @"[In Read] If a non-boolean value is used in a Sync command request ([MS-ASCMD] section 2.2.1.21), the server responds with Status element ([MS-ASCMD] section 2.2.3.177.17) value of 6 in the Sync command response ([MS-ASCMD] section 2.2.1.21).");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC11_ChangeSenderValueInRequest
        /// <summary>
        /// This case is designed to test that the server will return a Status element value of 6 in Sync command response, if the client attempts to change the value of Sender element.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC11_ChangeSenderValueInRequest()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Sender element is not supported when the ActiveSyncProtocolVersion is 12.1.");

            #region The sender calls SendMail method to send the mail to recipient 'User2'
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to change the value of Sender element
            // Synchronize the email item on the server
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);

            SyncRequest invalidSyncRequest = TestSuiteHelper.CreateSyncChangeRequest(
                syncChangeResult.SyncKey,
                this.User2Information.InboxCollectionId,
                new Request.SyncCollectionChange { ServerId = item.ServerId });
            SendStringResponse syncStringResponse = this.EMAILAdapter.InvalidSync(
                invalidSyncRequest,
                string.Format("<ApplicationData><Sender xmlns=\"Email2\">{0}</Sender></ApplicationData>", this.User2Information.UserName),
                "</Change>");
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R723");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R723
            Site.CaptureRequirementIfAreEqual<string>(
                "6",
                TestSuiteHelper.GetStatusCode(syncStringResponse.ResponseDataXML),
                723,
                @"[In Sender] If the client attempts to change this value, the server returns a Status element ([MS-ASCMD] section 2.2.3.177.17) value of 6 in the Sync command response ([MS-ASCMD] section 2.2.1.21).");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC12_SenderAndFromHaveSameValue
        /// <summary>
        /// This case is designed to test that the Send element is not present in the response when the email2:Sender element and the From element have the same value.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC12_SenderAndFromHaveSameValue()
        {
            #region The sender calls SendMail command with Sender element and From element having same value to send the email to recipient.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            string fromAddress = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string senderAddress = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            this.SendPlaintextEmail(emailSubject, null, null, null, senderAddress, null, fromAddress);
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case.
            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region The recipient calls Sync command to synchronize the email item on server.
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R726");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R726
            // The Sender element is set by server and is read-only for client, so if the response doesn't include this element, that's to say, its value is null.
            Site.CaptureRequirementIfIsTrue(
                item.Email.Sender == null && string.Compare(fromAddress, senderAddress, StringComparison.CurrentCultureIgnoreCase) == 0,
                726,
                @"[In Sender] The email2:Sender element is not sent to the client when the email2:Sender element and the From element have the same value.");

            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC13_SenderValueIsNull
        /// <summary>
        /// This case is designed to test that the Sender element is not present in the response when the email2:Sender element value is NULL.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC13_SenderValueIsNull()
        {
            #region The sender calls SendMail method without including the Sender element to send the email to recipient.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, null, null, Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), null, null, null);
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case.
            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region The recipient calls Sync method to synchronize the email items on the server.
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1065");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1065
            // The Sender element is set by server and is read-only for client, so if the response doesn't include this element, that's to say, its value is null.
            Site.CaptureRequirementIfIsNull(
                item.Email.Sender,
                1065,
                @"[In Sender] The email2:Sender element is not sent to the client when the email2:Sender element value is NULL.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC14_ChangeReceivedAsBccInRequest
        /// <summary>
        /// This case is designed to test that the server will respond with Status element value of 6 in the Sync command response, if the client changes the value of ReceivedAsBcc element.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC14_ChangeReceivedAsBccInRequest()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The ReceivedAsBcc element is not supported when the ActiveSyncProtocolVersion is 12.1.");

            #region The sender calls SendMail method to send the email to recipient.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Concat(this.User3Information.UserName, '@', this.User3Information.UserDomain));
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case.
            this.SwitchUser(this.User3Information, true);
            this.RecordCaseRelativeItems(this.User3Information.UserName, this.User3Information.InboxCollectionId, emailSubject);
            #endregion

            #region The recipient calls Sync method to change the value of ReceivedAsBcc element.

            // Synchronize the email item
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User3Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);

            // Try to change the value of ReceivedAsBcc
            SyncRequest invalidSyncRequest = TestSuiteHelper.CreateSyncChangeRequest(
                syncChangeResult.SyncKey,
                this.User3Information.InboxCollectionId,
                new Request.SyncCollectionChange { ServerId = item.ServerId });
            SendStringResponse syncStringResponse = this.EMAILAdapter.InvalidSync(
                invalidSyncRequest,
                string.Format("<ApplicationData><ReceivedAsBcc xmlns=\"Email2\">{0}</ReceivedAsBcc></ApplicationData>", false),
                "</Change>");
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R648");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R648
            Site.CaptureRequirementIfAreEqual<string>(
                "6",
                TestSuiteHelper.GetStatusCode(syncStringResponse.ResponseDataXML),
                648,
                @"[In ReceivedAsBcc] If the client changes this element value, the server responds with Status element ([MS-ASCMD] section 2.2.3.177.17) value of 6 in the Sync command response ([MS-ASCMD] section 2.2.1.21).");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC15_SendAndReplyToMultipleAddress
        /// <summary>
        /// This case is designed to test sending an email to multiple addresses, replying an email to multiple addresses and ConversationIndex element: Additional timestamps are added when the message is replied to.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC15_SendAndReplyToMultipleAddress()
        {
            #region The sender calls SendMail command to send the email to multiple recipients.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            string to = string.Format(
                "{0}, {1}",
                Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain),
                Common.GetMailAddress(this.User3Information.UserName, this.User3Information.UserDomain));
            string cc = string.Format(
                "{0}, {1}, {2}",
                Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain),
                Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain),
                Common.GetMailAddress(this.User3Information.UserName, this.User3Information.UserDomain));
            string replyTo = string.Format(
                 "{0}, {1}",
                Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain),
                Common.GetMailAddress(this.User3Information.UserName, this.User3Information.UserDomain));

            this.SendPlaintextEmail(emailSubject, cc, null, to, null, replyTo, null);
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case.
            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.InboxCollectionId, emailSubject);

            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, emailSubject);

            this.SwitchUser(this.User3Information, true);
            this.RecordCaseRelativeItems(this.User3Information.UserName, this.User3Information.InboxCollectionId, emailSubject);
            #endregion

            #region The sender calls Sync method to synchronize the email items on the server.
            this.SwitchUser(this.User1Information, false);
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User1Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R328");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R328
            // If GetSyncAddItem can get a non-null SyncItem, then we can capture this requirement since the emailSubject must be unique.
            Site.CaptureRequirementIfIsNotNull(
                item,
                328,
                @"[In Cc] The message is directed at the primary recipient as specified by the To element (section 2.2.2.79), but the secondary recipients also receive a copy of the message.");
            #endregion

            #region The recipient calls Sync method to synchronize the email items on the server.
            this.SwitchUser(this.User2Information, false);

            syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);
            #endregion

            #region Verify requirements
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R400");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R400
            bool isVerifiedR400 = item.Email.DisplayTo.Contains(";");

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR400,
                400,
                @"[In DisplayTo] If there are multiple display names, they are separated by semi-colons.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R331");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R331
            bool isVerifiedR331 = item.Email.CC.Contains(",");

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR331,
                331,
                @"[In Cc] If there are multiple e-mail addresses, they are separated by commas.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R711");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R711
            bool isVerifiedR711 = item.Email.ReplyTo.Contains(";");

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR711,
                711,
                @"[In ReplyTo] If there are multiple e-mail addresses, they are separated by a semi-colon.");

            Site.CaptureRequirementIfIsTrue(
                item.Email.To.Contains(","),
                787,
                @"[In To] If there are multiple e-mail addresses, they are separated by commas.");
            #endregion

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                #region The recipient calls SmartReply method to reply the received email and send the reply to multiple recipients.
                string replyFromUser = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
                string replySubject = string.Format("Re: {0}", emailSubject);
                string replyContent = Common.GenerateResourceName(Site, "reply: body");

                string replyMime = TestSuiteHelper.CreatePlainTextMime(
                    replyFromUser,
                    replyTo,
                    string.Empty,
                    string.Empty,
                    replySubject,
                    replyContent);

                SmartReplyRequest replyRequest = TestSuiteHelper.CreateSmartReplyRequest(this.User2Information.InboxCollectionId, item.ServerId, replyMime);
                this.EMAILAdapter.SmartReply(replyRequest);
                #endregion

                #region Record user name, folder collectionId and item subject that are used in this case.
                this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.InboxCollectionId, replySubject);
                this.RecordCaseRelativeItems(this.User3Information.UserName, this.User3Information.InboxCollectionId, replySubject);
                #endregion

                #region The recipient calls Sync method to synchronize the email items on the server.
                syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
                item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);
                #endregion

                #region Verify requirement
                Site.Assert.IsNotNull(item.Email.LastVerbExecuted, "The last action on the message should not be null.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R510");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R510
                Site.CaptureRequirementIfAreEqual<int>(
                    1,
                    item.Email.LastVerbExecuted.Value,
                    510,
                    @"[In LastVerbExecuted] The email2:LastVerbExecuted element is an optional element that indicates the last action, such as reply or forward, that was taken on the message.");
                #endregion

                #region The reply email recipient calls Sync method to synchronize the email item on the server.
                this.SwitchUser(this.User3Information, false);
                syncChangeResult = this.GetSyncResult(replySubject, this.User3Information.InboxCollectionId, null);
                Sync replyItem = TestSuiteHelper.GetSyncAddItem(syncChangeResult, replySubject);
                #endregion

                #region Verify requirements
                byte[] firstEmailIndex = Convert.FromBase64String(item.Email.ConversationIndex);
                byte[] replyEmailIndex = Convert.FromBase64String(replyItem.Email.ConversationIndex);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R351");

                // The value of the first timestamp is derived from the date and time when the message was originally sent by the server, for more information, refer to [MS-ASCON], so it can't be decode to the time when the message was sent by the server, capture MS-ASEMAIL_R351 if this element is not null.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R351
                Site.CaptureRequirementIfIsNotNull(
                    item.Email.ConversationIndex,
                    351,
                    @"[In ConversationIndex] The first timestamp identifies the date and time when the message was originally sent by the server.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1054");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1054
                // The ConversationIndex of first email only have 5 bytes Conversation Index Header and the later reply or forward email will add the extra 5 bytes response level. So the second email index is 2 times of the first email.
                Site.CaptureRequirementIfAreEqual<int>(
                    2,
                    replyEmailIndex.Length / firstEmailIndex.Length,
                    1054,
                    @"[In ConversationIndex] Additional timestamps are added when the message is [forwarded or] replied to.");
                #endregion
            }
        }
        #endregion

        #region MSASEMAIL_S01_TC16_ForwardEmail
        /// <summary>
        /// This case is designed to test the email2:LastVerbExecuted, email2:LastVerbExecutedTime and ConversationIndex element: Additional timestamps are added when the message is forwarded.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC16_ForwardEmail()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The email2:LastVerbExecuted element, email2:LastVerbExecutedTime element and ConversationIndex element are not supported when the ActiveSyncProtocolVersion is 12.1.");

            #region The sender calls SendMail method to send the email to recipient.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty, Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), null, null, null);
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case.
            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region The recipient calls Sync command to synchronize the email items on the server.
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);
            #endregion

            #region The recipient calls SmartForward method to forward the received email to the second recipient.
            string forwardFromUser = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            string forwardToUser = Common.GetMailAddress(this.User3Information.UserName, this.User3Information.UserDomain);
            string forwardSubject = string.Format("FW: {0}", emailSubject);
            string forwardContent = Common.GenerateResourceName(Site, "forward: body");

            string forwardMime = TestSuiteHelper.CreatePlainTextMime(
                forwardFromUser,
                forwardToUser,
                string.Empty,
                string.Empty,
                forwardSubject,
                forwardContent);

            SmartForwardRequest forwardRequest = TestSuiteHelper.CreateSmartForwardRequest(this.User2Information.InboxCollectionId, item.ServerId, forwardMime);
            this.EMAILAdapter.SmartForward(forwardRequest);
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case.
            this.SwitchUser(this.User3Information, true);
            this.RecordCaseRelativeItems(this.User3Information.UserName, this.User3Information.InboxCollectionId, forwardSubject);
            #endregion

            #region The second recipient calls Sync method to synchronize the email item on the server.
            syncChangeResult = this.GetSyncResult(forwardSubject, this.User3Information.InboxCollectionId, null);
            Sync forwardItem = TestSuiteHelper.GetSyncAddItem(syncChangeResult, forwardSubject);
            #endregion

            #region Verify requirement
            byte[] firstEmailIndex = Convert.FromBase64String(item.Email.ConversationIndex);
            byte[] forwardEmailIndex = Convert.FromBase64String(forwardItem.Email.ConversationIndex);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R352");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R352
            // The ConversationIndex of first email only have 5 bytes Conversation Index Header and the later reply or forward email will add the extra 5 bytes response level. So the second email index is 2 times of the first email.
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                forwardEmailIndex.Length / firstEmailIndex.Length,
                352,
                @"[In ConversationIndex] Additional timestamps are added when the message is forwarded [or replied to].");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC17_IncludedConversationIdInRequest
        /// <summary>
        /// This case is designed to test that the server will return a Status element value of 6 in the Sync command response, if the email2:ConversationId element is included within a Change element in a Sync command request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC17_IncludedConversationIdInRequest()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The email2:ConversationId element is not supported when the ActiveSyncProtocolVersion is 12.1.");

            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element and include "ConversationId" element to update the email item.
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            Request.SyncCollectionChange changeData = new Request.SyncCollectionChange
            {
                ServerId = item.ServerId,
                ApplicationData = new Request.SyncCollectionChangeApplicationData()
            };

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType7> itemsElementName = new List<Request.ItemsChoiceType7>();

            items.Add(true);
            itemsElementName.Add(Request.ItemsChoiceType7.Read);

            changeData.ApplicationData.Items = items.ToArray();
            changeData.ApplicationData.ItemsElementName = itemsElementName.ToArray();

            SyncRequest syncRequest = TestSuiteHelper.CreateSyncChangeRequest(getEmailItem.SyncKey, this.User2Information.InboxCollectionId, changeData);
            SendStringResponse response = this.EMAILAdapter.InvalidSync(syncRequest, "<ConversationId xmlns=\"Email2\"><![CDATA[6E83QrDB2UWpEw+EHaJC3A==]]></ConversationId>", "<Read xmlns=\"Email\">");
            #endregion

            #region Verify requirement
            // If the status value in sync change response is 6, then requirement MS-ASEMAIL_R347 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R347");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R347
            Site.CaptureRequirementIfAreEqual<string>(
                "6",
                TestSuiteHelper.GetStatusCode(response.ResponseDataXML),
                347,
                @"[In ConversationId] The server returns a Status element value of 6 in the Sync command response ([MS-ASCMD] section 2.2.1.21) when the email2:ConversationId element is included within a Change element ([MS-ASCMD] section 2.2.3.24) in a Sync command request.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC18_SyncAdd_AdministrativeMessage
        /// <summary>
        /// This case is designed to test that a non-delivery report message generated by the server has a message class deriving from message classes.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC18_SyncAdd_AdministrativeMessage()
        {
            #region Call SendMail command to send an email message to a none exist email address
            string emailBody = Common.GenerateResourceName(Site, "content");
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            string emailMime = TestSuiteHelper.CreatePlainTextMime(
                Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain),
                Common.GetMailAddress("NoneExistUser", this.User1Information.UserDomain),
                null,
                null,
                emailSubject,
                emailBody);
            string clientId = TestSuiteHelper.GetClientId();

            SendMailRequest sendMailRequest = TestSuiteHelper.CreateSendMailRequest(clientId, false, emailMime);
            SendMailResponse response = this.EMAILAdapter.SendMail(sendMailRequest);

            Site.Assert.AreEqual<string>(
                 string.Empty,
                 response.ResponseDataXML,
                 "The server should return an empty xml response data to indicate SendMail command executes successfully.");
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case.
            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.InboxCollectionId, emailSubject);
            #endregion

            #region Call Sync command to get the message in the organizer's inbox
            // Sync mailbox changes
            // Get the new added email item
            SyncStore syncItemResult;
            Sync item = null;
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            do
            {
                Thread.Sleep(waitTime);

                // Get the new added email item
                SyncStore initSyncResult = this.InitializeSync(this.User1Information.InboxCollectionId);
                syncItemResult = this.SyncChanges(initSyncResult.SyncKey, this.User1Information.InboxCollectionId, null);
                if (syncItemResult != null && syncItemResult.CollectionStatus == 1)
                {
                    if (syncItemResult.AddElements != null)
                    {
                        foreach (Sync syncItem in syncItemResult.AddElements)
                        {
                            if (syncItem.Email.Subject.Contains(emailSubject))
                            {
                                item = syncItem;
                                break;
                            }
                        }
                    }
                }

                counter++;
            }
            while ((syncItemResult == null || item == null) && counter < retryCount);

            // Verify sync result
            Site.Assert.IsNotNull(
                item,
                "If the Sync command executes successfully, the item in response shouldn't be null. Retry count: {0}",
                counter);
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC19_ItemOperations_NotIncludedDataInResponse
        /// <summary>
        /// This case is designed to test if an airsync:Schema element is included in the ItemOperations command request, server response MUST be restricted to the elements that were included as child elements of the airsync:Schema element in the command request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC19_ItemOperations_OnlyIncludedSubjectInResponse()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call ItemOperations command without including BodyPreference element to fetch all the information about the e-mail
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);

            // Just including subject element in schema
            Request.Schema schema = new Request.Schema();
            List<object> elements = new List<object> { string.Empty };

            List<Request.ItemsChoiceType4> names = new List<Request.ItemsChoiceType4>
            {
                Request.ItemsChoiceType4.Subject1
            };

            schema.Items = elements.ToArray();
            schema.ItemsElementName = names.ToArray();

            ItemOperationsRequest itemOperationRequest = TestSuiteHelper.CreateItemOperationsFetchRequest(this.User2Information.InboxCollectionId, item.ServerId, null, null, schema);
            ItemOperationsStore itemOperationResult = this.EMAILAdapter.ItemOperations(itemOperationRequest);
            ItemOperations itemOperationsItem = TestSuiteHelper.GetItemOperationsItem(itemOperationResult, emailSubject);
            Site.Assert.IsNotNull(itemOperationsItem, "The email item with subject '{0}' should be found in the {1} folder of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirements
            // If the server responds with an ItemOperations response, then MS-ASEMAIL_R61 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R61");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R61
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                itemOperationResult.Status,
                61,
                @"[In ItemOperations Command Response] When a client uses an ItemOperations command request ([MS-ASCMD] section 2.2.1.10), as specified in section 3.1.5.2, to retrieve data from the server for one or more specific e-mail items, the server responds with an ItemOperations command response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R63");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R63
            Site.CaptureRequirementIfIsTrue(
                TestSuiteHelper.IsOnlySpecifiedElement((XmlElement)this.EMAILAdapter.LastRawResponseXml, "Properties", "Subject"),
                63,
                @"[In ItemOperations Command Response] If an airsync:Schema element ([MS-ASCMD] section 2.2.3.158) is included in the ItemOperations command request, then the elements returned in the ItemOperations command response MUST be restricted to the elements that were included as child elements of the airsync:Schema element in the command request.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC20_ItemOperations_IncludedDataInResponse
        /// <summary>
        /// This case is designed to test ItemOperations command with including the BodyPreference element to get the Data information.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC20_ItemOperations_IncludedDataInResponse()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call ItemOperations command including BodyPreference element and BodyPartPreference element to fetch all the information about email using ServerId
            // Call method Sync to synchronize(add or change) the e-mail items with the server.
            Request.BodyPreference bodyPreference = new Request.BodyPreference
            {
                AllOrNone = true,
                AllOrNoneSpecified = true,
                TruncationSize = 5120,
                TruncationSizeSpecified = true,
                Type = 1
            };

            Request.BodyPartPreference bodyPartPreference = null;
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.1"))
            {
                bodyPartPreference = new Request.BodyPartPreference
                {
                    AllOrNone = true,
                    AllOrNoneSpecified = true,
                    TruncationSize = 5120,
                    TruncationSizeSpecified = true,
                    Type = 2
                };
            }

            // Get the new added email item
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, bodyPreference);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);

            // Fetch email from server
            ItemOperationsRequest itemOperationRequest = TestSuiteHelper.CreateItemOperationsFetchRequest(this.User2Information.InboxCollectionId, item.ServerId, bodyPreference, bodyPartPreference, null);
            ItemOperationsStore itemOperationResult = this.EMAILAdapter.ItemOperations(itemOperationRequest);
            #endregion

            #region Verify requirement
            // If the server responds with an ItemOperations response, then MS-ASEMAIL_R61 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R61");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R61
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                itemOperationResult.Status,
                61,
                @"[In ItemOperations Command Response] When a client uses an ItemOperations command request ([MS-ASCMD] section 2.2.1.10), as specified in section 3.1.5.2, to retrieve data from the server for one or more specific e-mail items, the server responds with an ItemOperations command response.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC21_ItemOperations_IncludedAllOrNone_False
        /// <summary>
        /// This case is designed to test server returns the truncated data if the size (in bytes) of the available data exceeds the value of the TruncationSize element in ItemOperations command request and setting the airsyncbase: AllOrNone element value to 0 (false).
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC21_ItemOperations_IncludedAllOrNone_False()
        {
            Request.BodyPreference fetchBodyPreference = new Request.BodyPreference
            {
                TruncationSize = 10,
                TruncationSizeSpecified = true,
                AllOrNone = false,
                AllOrNoneSpecified = true
            };

            this.VerifyAllOrNone(fetchBodyPreference);
        }
        #endregion

        #region MSASEMAIL_S01_TC22_ItemOperations_IncludedAllOrNone_True
        /// <summary>
        /// This case is designed to test server doesn't return a truncated response if the size (in bytes) of the available data not exceeds the value of the TruncationSize element in ItemOperations command request and setting the airsyncbase: AllOrNone element value to 1 (true).
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC22_ItemOperations_IncludedAllOrNone_True()
        {
            Request.BodyPreference fetchBodyPreference = new Request.BodyPreference
            {
                TruncationSize = 100,
                TruncationSizeSpecified = true,
                AllOrNone = true
            };

            this.VerifyAllOrNone(fetchBodyPreference);
        }
        #endregion

        #region MSASEMAIL_S01_TC23_InvalidSearch
        /// <summary>
        /// This case is designed to test that if email class elements are included in the request, a search:Status element value of 2 is contained as a child element of the search:Store element in Search command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC23_InvalidSearch()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Search command including "To" element to Search email from server.
            // Ensure the email item has reached the recipient's inbox folder.
            this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);

            // Search email from server
            SearchRequest searchRequest = TestSuiteHelper.CreateSearchRequest(emailSubject, this.User2Information.InboxCollectionId);
            SendStringResponse result = this.EMAILAdapter.InvalidSearch(searchRequest, string.Format("<To xmlns=\"Email\">{0}</To>", Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain)));

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(result.ResponseDataXML);
            XmlNamespaceManager xnm = new XmlNamespaceManager(doc.NameTable);
            xnm.AddNamespace("e", "Search");

            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 1;
            XmlNode searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);

            while (counter < retryCount && searchStatus != null && searchStatus.InnerXml.Equals("10"))
            {
                Thread.Sleep(waitTime);
                result = this.EMAILAdapter.InvalidSearch(searchRequest, string.Format("<To xmlns=\"Email\">{0}</To>", Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain)));
                doc.LoadXml(result.ResponseDataXML);
                xnm = new XmlNamespaceManager(doc.NameTable);
                xnm.AddNamespace("e", "Search");
                searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);
                counter++;
            }

            XmlNode status = doc.SelectSingleNode("/e:Search/e:Response/e:Store/e:Status", xnm);

            Site.Assert.IsNotNull(status, "The status of Search command response shouldn't be null");
            #endregion

            #region Verify requirement
            // If the Search command response from the server contains a search:Status element value of 2 as a child element of the search:Store element, then R68 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R68");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R68
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                status.InnerText,
                68,
                @"[In Search Command Response] If E-mail class elements are included in the Search command request, the Search command response from the server contains a search:Status element ([MS-ASCMD] section 2.2.3.177.13) value of 2 as a child element of the search:Store element ([MS-ASCMD] section 2.2.3.178.3).");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC24_SearchEmail
        /// <summary>
        /// This case is designed to test using Search command to retrieve E-mail class items from the server that match the criteria specified by the client.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC24_SearchEmail()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Search command to Search email from server.
            // Search email from server
            SearchRequest searchRequest = TestSuiteHelper.CreateSearchRequest(emailSubject, this.User2Information.InboxCollectionId);
            SearchResponse searchResponse = this.EMAILAdapter.Search(searchRequest);

            SearchStore searchStore = Common.LoadSearchResponse(searchResponse, Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site));
            Search searchItem = null;
            if (searchStore.Results.Count != 0)
            {
                foreach (Search item in searchStore.Results)
                {
                    if (item.Email.Subject == emailSubject)
                    {
                        searchItem = item;
                        break;
                    }
                }
            }

            Site.Assert.IsNotNull(searchItem, "The email message with subject {0} should be found.", emailSubject);
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC25_OnlyEmailClassPropertiesChangedOtherThanFlagRead

        /// <summary>
        /// This case is designed to test if E-Mail Class properties of an e-mail item other than the Read flag and Flag properties have changed for an e-mail item, the server MUST include the changed e-mail properties as child elements of the airsync:ApplicationData element within the airsync:Change element for that e-mail item in the Sync command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC25_OnlyEmailClassPropertiesChangedOtherThanFlagRead()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Categories element is not supported when the ActiveSyncProtocolVersion is 12.1.");

            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag and categories for the e-mail item and synchronize it with the server.
            // Get the email item
            SyncStore store = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(store, emailSubject);

            // Get the result of adding flag and categories
            SyncStore getChangedResult = this.AddFlagAndCategories(this.User2Information.InboxCollectionId, store.SyncKey, item.ServerId);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Call Sync command with Change element to update the categories property of the e-mail and synchronize it with server.
            string newCategory = Common.GenerateResourceName(Site, "NewCategory");
            Collection<string> newCategories = new Collection<string> { newCategory };

            // Update Categories property
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, null, item.ServerId, null, newCategories);

            // Get server changes 
            SyncStore updateResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(updateResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirements.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R77");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R77
            Site.CaptureRequirementIfIsTrue(
                item.Email.CategoriesIsInclude && item.Email.Categories != null,
                77,
                @"[In Sending E-Mail Changes to the Client] If E-Mail class properties of an e-mail item other than the Read flag and Flag properties have changed for an e-mail item, the server MUST include the changed e-mail properties as child elements of the airsync:ApplicationData element within the airsync:Change element for that e-mail item in the Sync command response.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC26_CreateDraftEMail

        /// <summary>
        /// This case is designed to test the requirements related with draft E-Mail.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC26_CreateDraftEMail()
        {
            Site.Assume.AreEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Bcc element is supported when the ActiveSyncProtocolVersion is 16.0.");
            Site.Assume.AreEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Bcc element is supported when the ActiveSyncProtocolVersion is 16.1.");

            #region Add an email item with Sync command.
            // Call FolderSync command to synchronize the collection hierarchy.
            FolderSyncRequest folderSyncRequest = Common.CreateFolderSyncRequest("0");
            FolderSyncResponse folderSyncResponse = this.EMAILAdapter.FolderSync(folderSyncRequest);

            string draftCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Drafts, this.Site);

            string subject = Common.GenerateResourceName(Site, "subject");
            string to = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string bcc = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Request.SyncCollectionAdd syncAddCollection = new Request.SyncCollectionAdd();
            string clientId = TestSuiteHelper.GetClientId();
            syncAddCollection.ClientId = clientId;
            syncAddCollection.ApplicationData = new Request.SyncCollectionAddApplicationData();
            List<object> items = new List<object>();
            List<Request.ItemsChoiceType8> itemsElementName = new List<Request.ItemsChoiceType8>();
            itemsElementName.Add(Request.ItemsChoiceType8.Subject3);
            items.Add(subject);
            itemsElementName.Add(Request.ItemsChoiceType8.To);
            items.Add(to);
            itemsElementName.Add(Request.ItemsChoiceType8.Bcc);
            items.Add(bcc);
 
            syncAddCollection.ApplicationData.Items = items.ToArray();
            syncAddCollection.ApplicationData.ItemsElementName = itemsElementName.ToArray();
            syncAddCollection.Class = "Email";

            SyncStore initSyncResponse = this.InitializeSync(draftCollectionId);
            SyncRequest addEMailRequest = TestSuiteHelper.CreateSyncAddRequest(initSyncResponse.SyncKey, draftCollectionId, syncAddCollection);
            this.EMAILAdapter.Sync(addEMailRequest);
            #endregion

            #region Call Sync command wihtout including BodyPreference to synchronize the e-mail items with server.
            // Get the new added email item
            SyncStore syncChangeResult = this.GetSyncResult(subject, draftCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, subject);

            this.Site.CaptureRequirementIfAreEqual<bool>(
                true,
                item.Email.IsDraft.Value,
                1278,
                @"[In IsDraft] The value 1 (TRUE) indicates that the email is a draft.");

            this.Site.CaptureRequirementIfIsNotNull(
                item.Email.IsDraft,
                1273,
                @"[In IsDraft] This element [email2:IsDraft] is present in a Sync command response ([MS-ASCMD] section 2.2.2.20)[, a Search command response ([MS-ASCMD] section 2.2.2.15), or an ItemOperations command response ([MS-ASCMD] section 2.2.2.9)].");
            
            this.Site.CaptureRequirementIfIsTrue(
                !string.IsNullOrEmpty(item.Email.ConversationId),
                1184,
                @"[In ConversationId] In protocol version 16.0: When the client adds a new draft item, the server response will contain the email2:ConversationId element for that draft item.");

            this.Site.CaptureRequirementIfIsTrue(
                !string.IsNullOrEmpty(item.Email.ConversationIndex),
                1190,
                @"[In ConversationIndex] In protocol version 16.0: When the client adds a new draft item, the server response will include the email2:ConversationIndex element for that draft item.");
            #endregion

            #region Call Search command to Search email from server.
            // Search email from server
            SearchRequest searchRequest = TestSuiteHelper.CreateSearchRequest(subject, draftCollectionId);
            SearchResponse searchResponse = this.EMAILAdapter.Search(searchRequest);

            SearchStore searchStore = Common.LoadSearchResponse(searchResponse, Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site));
            Search searchItem = null;
            if (searchStore.Results.Count != 0)
            {
                foreach (Search resultItem in searchStore.Results)
                {
                    if (resultItem.Email.Subject == subject)
                    {
                        searchItem = resultItem;
                        break;
                    }
                }
            }

            Site.Assert.IsNotNull(searchItem, "The email message with subject {0} should be found.", subject);

            this.Site.CaptureRequirementIfIsNotNull(
                searchItem.Email.IsDraft,
                1274,
                @"[In IsDraft] This element [email2:IsDraft] is present in [a Sync command response ([MS-ASCMD] section 2.2.2.20),] a Search command response ([MS-ASCMD] section 2.2.2.15)[, or an ItemOperations command response ([MS-ASCMD] section 2.2.2.9)].");
            #endregion

            #region Call ItemOperations command without including BodyPreference element to fetch all the information about the e-mail
            ItemOperationsRequest itemOperationRequest = TestSuiteHelper.CreateItemOperationsFetchRequest(draftCollectionId, item.ServerId, null, null, null);
            ItemOperationsStore itemOperationResult = this.EMAILAdapter.ItemOperations(itemOperationRequest);
            ItemOperations itemOperationsItem = TestSuiteHelper.GetItemOperationsItem(itemOperationResult, subject);
     
            this.Site.CaptureRequirementIfIsNotNull(
                itemOperationsItem.Email.IsDraft,
                1275,
                @"[In IsDraft] This element [email2:IsDraft] is present in [a Sync command response ([MS-ASCMD] section 2.2.2.20), a Search command response ([MS-ASCMD] section 2.2.2.15), or] an ItemOperations command response ([MS-ASCMD] section 2.2.2.9).");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC27_CreateDraftEMailAndSend

        /// <summary>
        /// This case is designed to test to draft E-Mail and send it.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC27_CreateDraftEMailAndSend()
        {
            Site.Assume.AreEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Bcc element is supported when the ActiveSyncProtocolVersion is 16.0.");
            Site.Assume.AreEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Bcc element is supported when the ActiveSyncProtocolVersion is 16.1.");

            #region Add an email item with Sync command and send it.
            // Call FolderSync command to synchronize the collection hierarchy.
            FolderSyncRequest folderSyncRequest = Common.CreateFolderSyncRequest("0");
            FolderSyncResponse folderSyncResponse = this.EMAILAdapter.FolderSync(folderSyncRequest);

            string draftCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Drafts, this.Site);

            string subject = Common.GenerateResourceName(Site, "subject");
            string to = Common.GetMailAddress(this.User2Information.UserName, this.User1Information.UserDomain);
            Request.SyncCollectionAdd syncAddCollection = new Request.SyncCollectionAdd();
           
            string clientId = TestSuiteHelper.GetClientId();
            syncAddCollection.ClientId = clientId;
            syncAddCollection.ApplicationData = new Request.SyncCollectionAddApplicationData();
            List<object> items = new List<object>();
            List<Request.ItemsChoiceType8> itemsElementName = new List<Request.ItemsChoiceType8>();
            itemsElementName.Add(Request.ItemsChoiceType8.Subject3);
            items.Add(subject);
            itemsElementName.Add(Request.ItemsChoiceType8.To);
            items.Add(to);

            syncAddCollection.ApplicationData.Items = items.ToArray();
            syncAddCollection.ApplicationData.ItemsElementName = itemsElementName.ToArray();
            syncAddCollection.Class = "Email";
            syncAddCollection.Send = string.Empty;

            SyncStore initSyncResponse = this.InitializeSync(draftCollectionId);
            SyncRequest addEMailRequest = TestSuiteHelper.CreateSyncAddRequest(initSyncResponse.SyncKey, draftCollectionId, syncAddCollection);
            this.EMAILAdapter.Sync(addEMailRequest);
            #endregion

            #region Call Sync command wihtout including BodyPreference to synchronize the e-mail items with server.
            this.SwitchUser(this.User2Information, true);
            // Get the new added email item
            SyncStore syncChangeResult = this.GetSyncResult(subject, User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, subject);
            
            this.Site.CaptureRequirementIfAreEqual<bool>(
                false,
                item.Email.IsDraft.Value,
                1279,
                @"[In IsDraft] The value 0 (FALSE) indicates that the email is not a draft.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC28_BothReadAndFlagAndCategoriesChanged
        /// <summary>
        /// This case is designed to test if both Read element and Flag element have changed for the e-mail item, the server will send only Read element and Flag element to the client for an e-mail item in Sync command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC28_ReadAndFlagAndCategoriesChanged()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag and categories for the e-mail item and synchronize it with the server.
            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Get the result of adding flag and categories
            SyncStore getChangedResult = this.AddFlagAndCategories(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, item.ServerId);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region all Sync command with Change element to update the flag and read properties of the e-mail and synchronize it with server.
            Request.Flag newFlag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                CompleteTime = DateTime.Now,
                CompleteTimeSpecified = true,
                DateCompleted = DateTime.Now,
                DateCompletedSpecified = true
            };

            // Update Flag, Read and category property   
            string newCategory = Common.GenerateResourceName(Site, "NewCategory");
            Collection<string> newCategories = new Collection<string> { newCategory };

            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, false, item.ServerId, newFlag, newCategories);


            // Get server changes 
            SyncStore updateResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(updateResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirement.
            // If the server send Read flag and Flag block, then MS-ASEMAIL_R104 can be captured.
            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R104");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R104
                Site.CaptureRequirementIfIsTrue(
                    item.Email.FlagIsInclude && item.Email.Flag != null && item.Email.ReadIsInclude && item.Email.CategoriesIsInclude && item.Email.Categories != null,
                    104,
                    @"[In Sending E-Mail Changes to the Client]  If Read flag , Flag properties and Categories properties changed, Other E-Mail class properties and Non-E-Mail class properties not changed, Server action for e-mail item in Sync commend response is: Send Read flag, Flag block and categories block.");
            }
            #endregion
        }
        #endregion      

        #region MSASEMAIL_S01_TC29_OnlyCategoriesChanged
        /// <summary>
        /// This case is designed to test if both Read element and Flag element have changed for the e-mail item, the server will send only Read element and Flag element to the client for an e-mail item in Sync command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC29_OnlyCategoriesChanged()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag and categories for the e-mail item and synchronize it with the server.
            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Get the result of adding flag and categories
            SyncStore getChangedResult = this.AddFlagAndCategories(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, item.ServerId);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region all Sync command with Change element to update the flag and read properties of the e-mail and synchronize it with server.

            // Update category property   
            string newCategory = Common.GenerateResourceName(Site, "NewCategory");
            Collection<string> newCategories = new Collection<string> { newCategory };
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, null, item.ServerId, null, newCategories);


            // Get server changes 
            SyncStore updateResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(updateResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirement.
            // If the server send Read flag and Flag block, then MS-ASEMAIL_R104 can be captured.
            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R96");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R96
                Site.CaptureRequirementIfIsTrue(
                    item.Email.CategoriesIsInclude && item.Email.Categories != null,
                    96,
                    @"[In Sending E-Mail Changes to the Client] If Read flag, Other E-mail class properties, Flag properties  and Non-E-mail class properties not changed, Categories properties changed, Server action for e-mail item in Sync command response is: Send Categories block only");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC30_BothReadandCategoriesChanged
        /// <summary>
        /// This case is designed to test if both Read element and Flag element have changed for the e-mail item, the server will send only Read element and Flag element to the client for an e-mail item in Sync command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC30_BothReadandCategoriesChanged()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag and categories for the e-mail item and synchronize it with the server.
            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Get the result of adding flag and categories
            SyncStore getChangedResult = this.AddFlagAndCategories(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, item.ServerId);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region all Sync command with Change element to update the flag and read properties of the e-mail and synchronize it with server.

            // Update category property   
            string newCategory = Common.GenerateResourceName(Site, "NewCategory");
            Collection<string> newCategories = new Collection<string> { newCategory };
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, false, item.ServerId, null, newCategories);


            // Get server changes 
            SyncStore updateResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(updateResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirement.
            // If the server send Read flag and Flag block, then MS-ASEMAIL_R104 can be captured.
            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R10608");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R10608
                Site.CaptureRequirementIfIsTrue(
                    item.Email.ReadIsInclude && item.Email.CategoriesIsInclude && item.Email.Categories != null,
                    10608,
                    @"[In Sending E-Mail Changes to the Client]  If Read flag and Categories properties changed, Flag properties, Other E-Mail class properties and Non-E-Mail class properties not changed, Server action for e-mail item in Sync commend response is: Send Read flag and Categories block");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S01_TC31_BothFlagandCategoriesChanged
        /// <summary>
        /// This case is designed to test if both Read element and Flag element have changed for the e-mail item, the server will send only Read element and Flag element to the client for an e-mail item in Sync command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S01_TC31_BothFlagandCategoriesChanged()
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag and categories for the e-mail item and synchronize it with the server.
            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Get the result of adding flag and categories
            SyncStore getChangedResult = this.AddFlagAndCategories(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, item.ServerId);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region all Sync command with Change element to update the flag and read properties of the e-mail and synchronize it with server.

            // Update category property   
            Request.Flag newFlag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                CompleteTime = DateTime.Now,
                CompleteTimeSpecified = true,
                DateCompleted = DateTime.Now,
                DateCompletedSpecified = true
            };
            string newCategory = Common.GenerateResourceName(Site, "NewCategory");
            Collection<string> newCategories = new Collection<string> { newCategory };
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, null, item.ServerId, newFlag, newCategories);


            // Get server changes 
            SyncStore updateResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(updateResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirement.
            // If the server send Read flag and Flag block, then MS-ASEMAIL_R104 can be captured.
            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R10600");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R10600
                Site.CaptureRequirementIfIsTrue(
                    item.Email.FlagIsInclude && item.Email.Flag != null && item.Email.CategoriesIsInclude && item.Email.Categories != null,
                    10600,
                    @"[In Sending E-Mail Changes to the Client]  If Flag properties, Categories properties changed, Read flag, Other E-Mail class properties and Non-E-Mail class properties not changed, Server action for e-mail item in Sync commend response is: Send Flag block and Categories block");
            }
            #endregion
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Add flag and categories for an email item
        /// </summary>
        /// <param name="collectionId">The collectionId of the folder which contains the specified item.</param>
        /// <param name="syncKey">The synchronization key to get the new added email item.</param>
        /// <param name="serverId">The server id of the new added email item.</param>
        /// <returns>The result of adding email flag and categories.</returns>
        private SyncStore AddFlagAndCategories(string collectionId, string syncKey, string serverId)
        {
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "for Follow Up",
                CompleteTime = DateTime.Now,
                CompleteTimeSpecified = true,
                DateCompleted = DateTime.Now,
                DateCompletedSpecified = true
            };

            Collection<string> categories = new Collection<string>();

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                string category = Common.GenerateResourceName(Site, "Category");
                categories.Add(category);
            }
            else
            {
                categories = null;
            }

            // Add category and flag to the email item and set Read property to true
            this.UpdateEmail(collectionId, syncKey, true, serverId, flag, categories);

            // Get the result of adding flag and category
            return this.SyncChanges(syncKey, collectionId, null);
        }

        /// <summary>
        /// Verify the result of setting element AllOrNone to true or false.
        /// </summary>
        /// <param name="fetchBodyPreference">The preference information related to the type and size of information that is returned from fetching.</param>
        private void VerifyAllOrNone(Request.BodyPreference fetchBodyPreference)
        {
            #region Call method SendMail to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command including BodyPreference element to get the email item
            // Get the new added email item
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);
            #endregion

            #region Call ItemOperations command including BodyPreference element and setting the value of AllOrNone to false to fetch all the information about email using ServerId.
            // Call method ItemOperations to fetch all the information about e-mail using serverId
            fetchBodyPreference.Type = (byte)0x1;

            // Fetch email from server
            ItemOperationsRequest itemOperationRequest = TestSuiteHelper.CreateItemOperationsFetchRequest(this.User2Information.InboxCollectionId, item.ServerId, fetchBodyPreference, null, null);
            ItemOperationsStore itemOperationResult = this.EMAILAdapter.ItemOperations(itemOperationRequest);
            ItemOperations itemOperationsItem = TestSuiteHelper.GetItemOperationsItem(itemOperationResult, emailSubject);
            Site.Assert.IsNotNull(itemOperationsItem, "The email item with subject '{0}' should be found.", emailSubject);
            #endregion
        }
        #endregion
    }
}