namespace Microsoft.Protocols.TestSuites.MS_ASEMAIL
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to test flag events, include setting a flag on email, updating a flag on email, marking a flag on email as complete, clearing a flag from email.
    /// </summary>   
    [TestClass]
    public class S03_EmailFlag : TestSuiteBase
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

        #region MSASEMAIL_S03_TC01_AddFlag_AllElements
        /// <summary>
        /// This case is designed to test flagging an email with all elements.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC01_AddFlag_AllElements()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag with all flag child elements.
            DateTime startDate = DateTime.Now;
            DateTime dueDate = DateTime.Now.AddDays(20);

            // Create flag element with all properties
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "for Follow Up",
                Subject = Common.GenerateResourceName(Site, "FlagSubject"),
                StartDate = startDate,
                StartDateSpecified = true,
                UtcStartDate = startDate.ToUniversalTime(),
                UtcStartDateSpecified = true,
                DueDate = dueDate,
                DueDateSpecified = true,
                UtcDueDate = dueDate.ToUniversalTime(),
                UtcDueDateSpecified = true,
                OrdinalDate = DateTime.Now,
                OrdinalDateSpecified = true,
                ReminderSet = 1,
                ReminderSetSpecified = true,
                ReminderTime = DateTime.Now.AddDays(12),
                ReminderTimeSpecified = true,
                CompleteTime = DateTime.Now.AddDays(30),
                CompleteTimeSpecified = true,
                DateCompleted = DateTime.Now.AddDays(30),
                DateCompletedSpecified = true
            };

            flag.SubOrdinalDate = flag.OrdinalDate.ToString();

            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);

            // Add flag to an email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirements
            if (Common.IsRequirementEnabled(1007, this.Site))
            {
                // Because MS-ASEMAIL doesn't test sorting, if the Sync command succeeds and the response contains a SubOrdinalDate value, then MS-ASEMAIL_R1007 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1007");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1007
                Site.CaptureRequirementIfIsNotNull(
                    item.Email.Flag.SubOrdinalDate,
                    1007,
                    @"[In Appendix B: Product Behavior] It[tasks:SubOrdinalDate] can contain any string value and is used for sorting if there are duplicate tasks:OrdinalDate element (section 2.2.2.45) values. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R692");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R692
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                item.Email.Flag.ReminderSet,
                692,
                @"[In ReminderSet] The value of this element[tasks:ReminderSet] is set to 1 if a reminder has been set for the task.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC02_AddFlag_InvalidStatus
        /// <summary>
        /// This case is designed to test that the server will return a Status element value of 6 in the Sync command response when adding Flag with invalid Status.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC02_AddFlag_InvalidStatus()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add a flag with invalid Status value.
            Request.Flag flag = new Request.Flag
            {
                Status = "3",
                FlagType = "for Follow Up",
                CompleteTime = DateTime.Now,
                CompleteTimeSpecified = true,
                DateCompleted = DateTime.Now,
                DateCompletedSpecified = true
            };

            // Get the synchronization result.
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);

            // Get the new added email item
            Sync emailItem = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to the email item
            SyncStore updateResult = this.UpdateEmail(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, null, emailItem.ServerId, flag, null);

            Site.Assert.AreEqual<int>(
                6,
                int.Parse(updateResult.ChangeResponses[0].Status),
                "The server should return a Status 6 in the Sync command response if the flag Status element is set to a value other than 0, 1, or 2.");
            #endregion

            #region Verify requirement
            // If the server returns a Status element value of 6 in the Sync command response, then MS-ASEMAIL_R757 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R757");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R757
            Site.CaptureRequirement(
                757,
                @"[In Status] The server returns a Status element ([MS-ASCMD] section 2.2.3.177.17) value of 6 in the Sync command response ([MS-ASCMD] section 2.2.1.21) if the flag Status element is set to a value other than 0, 1, or 2.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC03_AddFlag_CompleteTimeAndDateCompleted
        /// <summary>
        /// This case is designed to test that the message will also include a value for the tasks:DateComplete element, if the message includes a value for the CompleteTime element.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC03_AddFlag_CompleteTimeAndDateCompleted()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag with DateComplete element and CompleteTime element.
            Request.Flag flag = new Request.Flag
            {
                Status = "1",
                FlagType = "for Follow Up",
                CompleteTime = DateTime.Now,
                CompleteTimeSpecified = true,
                DateCompleted = DateTime.Now,
                DateCompletedSpecified = true
            };

            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to an email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirement
            if (Common.IsRequirementEnabled(336, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R336");

                // Verify the prerequisite first
                Site.Assert.IsTrue(item.Email.Flag.CompleteTimeSpecified, "If Sync command updates email flag with CompleteTime element successfully, then the Flag.CompleteTimeSpecified field should be true");
                Site.Assert.IsNotNull(item.Email.Flag.CompleteTime, "If Sync command updates email flag with CompleteTime element successfully, then the Flag.CompleteTime field should not be null");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R336
                Site.CaptureRequirementIfIsTrue(
                    item.Email.Flag.DateCompletedSpecified,
                    336,
                    @"[In Appendix B: Product Behavior] If a message includes a value for the CompleteTime element, the message does also include a value for the tasks:DateCompleted element. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC04_AddFlag_InvalidReminder
        /// <summary>
        /// This case is designed to test that the server will return a Status element value of 6 in the command response, if tasks: ReminderSet element and tasks: ReminderTime element are invalid.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC04_AddFlag_InvalidReminder()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag with the ReminderSet element but missing the ReminderTime element.
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "for Follow Up",
                ReminderSet = 1,
                ReminderSetSpecified = true
            };

            // Get the synchronization result.
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);

            // Get the new added email item
            Sync emailItem = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to the email item
            SyncStore updateResult = this.UpdateEmail(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, null, emailItem.ServerId, flag, null);

            Site.Assert.AreEqual<int>(
                6,
                int.Parse(updateResult.ChangeResponses[0].Status),
                "The server should return a Status 6 in the Sync command response if the ReminderSet element value is set to 1 (TRUE) and the ReminderTime element is not included in the Sync command request.");
            #endregion

            #region Verify requirement
            // If the server returns a Status element value of 6 in the Sync command response, then MS-ASEMAIL_R702 can be captured. 
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R702");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R702
            Site.CaptureRequirement(
                702,
                @"[In ReminderTime] The server returns a Status element ([MS-ASCMD] section 2.2.3.177.17) value of 6 in the Sync command response ([MS-ASCMD] section 2.2.1.21) if the tasks:ReminderSet element value is set to 1 (TRUE) and the tasks:ReminderTime element is not included in the Sync command request ([MS-ASCMD] section 2.2.1.21).");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC05_AddFlag_DueDateBeforeStartDate
        /// <summary>
        /// This case is designed to test that the server will return a Status element value of 6 in the Sync command response, when DueDate occurs before StartDate.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC05_AddFlag_DueDateBeforeStartDate()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag with DueDate occurs before StartDate.
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "for Follow Up",
                StartDate = DateTime.Now.AddDays(20),
                StartDateSpecified = true,
                UtcStartDate = DateTime.Now.AddDays(20),
                UtcStartDateSpecified = true,
                DueDate = DateTime.Now.AddDays(10),
                DueDateSpecified = true,
                UtcDueDate = DateTime.Now.AddDays(30),
                UtcDueDateSpecified = true
            };

            // Get the synchronization result.
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);

            // Get the new added email item
            Sync emailItem = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to the email item
            SyncStore updateResult = this.UpdateEmail(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, null, emailItem.ServerId, flag, null);

            Site.Assert.AreEqual<int>(
                6,
                int.Parse(updateResult.ChangeResponses[0].Status),
                "The server should return a Status 6 in the Sync command response if the DueDate occurs before StartDate.");
            #endregion

            #region Verify requirements
            // If the server returns a Status element value of 6 in the Sync command response, then MS-ASEMAIL_R407, MS-ASEMAIL_R739 and MS-ASEMAIL_R741 can be captured. 
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R407");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R407
            Site.CaptureRequirement(
                407,
                @"[In DueDate] The server returns a Status element ([MS-ASCMD] section 2.2.3.177.17) value of 6 in the Sync command response ([MS-ASCMD] section 2.2.1.21) if this condition[When a flag is updated, the tasks:DueDate element value MUST NOT occur before the tasks:StartDate element (section 2.2.2.72) value.] is not met.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R739");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R739
            Site.CaptureRequirement(
                739,
                @"[In StartDate] The server returns a Status element ([MS-ASCMD] section 2.2.3.177.17) value of 6 in the Sync command response ([MS-ASCMD] section 2.2.1.21) if this condition[When a flag is updated, the tasks:StartDate element value MUST NOT occur after the tasks:DueDate element (section 2.2.2.31) value. ] is not met.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R741");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R741
            Site.CaptureRequirement(
                741,
                @"[In StartDate] The server returns a Status element ([MS-ASCMD] section 2.2.3.177.17) value of 6 in the Sync command response if this condition[the tasks:StartDate element value MUST NOT occur after the tasks:DueDate element (section 2.2.2.31) value] is not met.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC06_AddFlag_UtcDueDateBeforeUtcStartDate
        /// <summary>
        /// This case is designed to test that the server will return a Status element value of 6 in the Sync command response, when UtcDueDate occurs before UtcStartDate.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC06_AddFlag_UtcDueDateBeforeUtcStartDate()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag with UtcDueDate occurs before UtcStartDat.
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "for Follow Up",
                StartDate = DateTime.Now.AddDays(20),
                StartDateSpecified = true,
                UtcStartDate = DateTime.Now.AddDays(20),
                UtcStartDateSpecified = true,
                DueDate = DateTime.Now.AddDays(30),
                DueDateSpecified = true,
                UtcDueDate = DateTime.Now,
                UtcDueDateSpecified = true
            };

            // Get the synchronization result
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);

            // Get the new added email item
            Sync emailItem = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to the email item
            SyncStore updateResult = this.UpdateEmail(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, null, emailItem.ServerId, flag, null);

            Site.Assert.AreEqual<int>(
                6,
                int.Parse(updateResult.ChangeResponses[0].Status),
                "The server should return a Status 6 in the Sync command response if the UtcDueDate occurs before UtcStartDate.");
            #endregion

            #region Verify requirements
            // If the server returns a Status element value of 6 in the Sync command response, then MS-ASEMAIL_R862 and MS-ASEMAIL_R851 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R851");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R851
            Site.CaptureRequirement(
                851,
                @"[In UtcDueDate] The server returns a Status element ([MS-ASCMD] section 2.2.3.177.17) value of 6 in the Sync command response ([MS-ASCMD] section 2.2.1.21) if this condition[When a flag is updated, the tasks:UtcDueDate element value MUST NOT occur before the tasks:UtcStartDate element (section 2.2.2.87) value. ] is not met.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R862");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R862
            Site.CaptureRequirement(
                862,
                @"[In UtcStartDate] The server returns a Status element ([MS-ASCMD] section 2.2.3.177.17) value of 6 in the Sync command response ([MS-ASCMD] section 2.2.1.21) if this condition[When a flag is updated, the tasks:UtcStartDate element value MUST occur before the tasks:UtcDueDate element (section 2.2.2.86) value. ] is not met.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC07_AddFlag_InvalidDate
        /// <summary>
        /// This case is designed to test that the server will return a Status element value of 6 in the Sync command response, when client just sets the value of StartDate and DueDate.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC07_AddFlag_InvalidDate()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag and just set the value of StartDate and DueDate.
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "for Follow Up",
                StartDate = DateTime.Now,
                StartDateSpecified = true,
                DueDate = DateTime.Now.AddDays(20),
                DueDateSpecified = true
            };

            // Get the synchronization result
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);

            // Get the new added email item
            Sync emailItem = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to the email item
            SyncStore updateResult = this.UpdateEmail(this.User2Information.InboxCollectionId, getEmailItem.SyncKey, null, emailItem.ServerId, flag, null);

            Site.Assert.AreEqual<int>(
                6,
                int.Parse(updateResult.ChangeResponses[0].Status),
                "The server should return a Status 6 in the Sync command response if the the request just sets StartDate and DueDate.");
            #endregion

            #region Verify requirements
            // If the server returns a Status element value of 6 in the Sync command response, then MS-ASEMAIL_R409, MS-ASEMAIL_R864 and MS-ASEMAIL_R853 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R409");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R409
            Site.CaptureRequirement(
                409,
                @"[In DueDate] The server returns a Status element value of 6 in the Sync response if this condition[To set a flag, the tasks:StartDate element, tasks:DueDate element, tasks:UtcStartDate element (section 2.2.2.87), and tasks:UtcDueDate element (section 2.2.2.86) either all MUST be set, or all MUST be NULL.] is not met.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R864");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R864
            Site.CaptureRequirement(
                864,
                @"[In UtcStartDate] The server returns a Status element value of 6 in the Sync command response if this condition[To set a flag, the tasks:StartDate element, tasks:DueDate element (section 2.2.2.31), tasks:UtcStartDate element, and tasks:UtcDueDate element either all MUST be set, or all MUST be NULL. ] is not met.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R853");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R853
            Site.CaptureRequirement(
                853,
                @"[In UtcDueDate] The server returns a Status element value of 6 in the Sync command response if this condition[To set a flag, the tasks:StartDate element (section 2.2.2.72), tasks:DueDate element, tasks:UtcStartDate element, and tasks:UtcDueDate element either all MUST be set, or all MUST be NULL.] is not met.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC08_AddFlag_ValidDate
        /// <summary>
        /// This case is designed to test that the tasks:UtcDueDate element value must not occur before the tasks:UtcStartDate element and the DueDate element value must not occur before the StartDate element.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC08_AddFlag_ValidDate()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command to add flag with valid StartDate, DueDate, UtcStartDate and UtcDueStartDate.
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "for Follow Up",
                StartDate = DateTime.Now,
                StartDateSpecified = true,
                UtcStartDate = DateTime.Now,
                UtcStartDateSpecified = true,
                DueDate = DateTime.Now.AddDays(20),
                DueDateSpecified = true,
                UtcDueDate = DateTime.Now.AddDays(20),
                UtcDueDateSpecified = true
            };

            // Get the synchronization result
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);

            // Add flag to an email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);

            Site.Assert.AreEqual<byte>(
                1,
                getChangedResult.CollectionStatus,
                "The server returns a Status 1 in the Sync command response indicate set flag success.");
            #endregion

            #region Verify requirements
            // If server returned a Status 1, that is to say, the sync command is successful, then MS-ASEMAIL_R850, MS-ASEMAIL_R861, MS-ASEMAIL_R738 and MS-ASEMAIL_R406 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R850");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R850
            Site.CaptureRequirement(
                850,
                @"[In UtcDueDate] When a flag is updated, the tasks:UtcDueDate element value MUST NOT occur before the tasks:UtcStartDate element (section 2.2.2.87) value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R861");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R861
            Site.CaptureRequirement(
                861,
                @"[In UtcStartDate] When a flag is updated, the tasks:UtcStartDate element value MUST occur before the tasks:UtcDueDate element (section 2.2.2.86) value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R738");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R738
            Site.CaptureRequirement(
                738,
                @"[In StartDate] When a flag is updated, the tasks:StartDate element value MUST NOT occur after the tasks:DueDate element (section 2.2.2.31) value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R406");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R406
            Site.CaptureRequirement(
                406,
                @"[In DueDate] When a flag is updated, the tasks:DueDate element value MUST NOT occur before the tasks:StartDate element (section 2.2.2.72) value.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC09_AddFlag_OptionalFlagTypeInMeetingRequest
        /// <summary>
        /// This case is designed to test that the FlagType element is not required if the e-mail message is a meeting request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC09_AddFlag_OptionalFlagTypeInMeetingRequest()
        {
            #region Call SendMail command to send one meeting request email
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            string organizerEmailAddress = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateDefaultCalendar(emailSubject, organizerEmailAddress, attendeeEmailAddress, null, null, null, null);

            // Record the calendar item that created in calendar folder of user1
            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, emailSubject);

            // Send a meeting request email to user2
            this.SendMeetingRequest(emailSubject, calendar);
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case.
            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, emailSubject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, emailSubject);
            #endregion

            #region Sync changes in user2 mailbox .
            // Get the new added email item
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync emailItem = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);
            #endregion

            #region Verify requirement
            // If server returned flag element without containing flagType element, then MS-ASEMAIL_R470 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R470");

            // Verify MS-ASEMAIL requirement:MS-ASEMAIL_R470
            Site.CaptureRequirementIfIsNull(
                emailItem.Email.Flag.FlagType,
                470,
                @"[In FlagType] The FlagType element is not required if the e-mail message is a meeting request.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC10__AddFlag_OptionalFlagTypeInMeetingResponse
        /// <summary>
        /// This case is designed to test that the FlagType element is not required if the e-mail message is a meeting response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC10_AddFlag_OptionalFlagTypeInMeetingResponse()
        {
            #region Call SendMail command to send a meeting response message
            this.SwitchUser(this.User2Information, true);

            string emailSubject = Common.GenerateResourceName(Site, "subject");
            string organizerEmailAddress = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = TestSuiteHelper.CreateDefaultCalendar(emailSubject, organizerEmailAddress, attendeeEmailAddress);

            // Send meeting response reply email to organizer
            this.SendMeetingResponse(calendar);
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case.
            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.InboxCollectionId, emailSubject);
            #endregion

            #region Call Sync command to synchronize the meeting response message
            // Switch to user1 mailbox
            this.SwitchUser(this.User1Information, false);

            // Get the new added email item
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User1Information.InboxCollectionId, null);
            Sync emailItem = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1058");

            // If server returned a Flag element which doesn't contain FlagType element, that is to say, FlagType element is not required then MS-ASEMAIL_R1058 can be captured.
            // Verify MS-ASEMAIL requirement:MS-ASEMAIL_R1058
            Site.CaptureRequirementIfIsNull(
                emailItem.Email.Flag.FlagType,
                1058,
                @"[In FlagType] The FlagType element is not required if the e-mail message is a meeting response.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC11_FlagEmail_RequiredPropertiesOption1
        /// <summary>
        /// This case is designed to test when action is flagging an email, required properties from device are: Status = 2, FlagType = "Flag for follow up", tasks:StartDate, tasks:UtcStartDate, tasks:DueDate and tasks:UtcDueDate.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC11_FlagEmail_RequiredPropertiesOption1()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag with properties Status = 2, FlagType = "Flag for follow up", StartDate, UtcStartDate, DueDate and UtcDueDate.
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                StartDate = DateTime.Now,
                StartDateSpecified = true,
                UtcStartDate = DateTime.Now,
                UtcStartDateSpecified = true,
                DueDate = DateTime.Now.AddDays(20),
                DueDateSpecified = true,
                UtcDueDate = DateTime.Now.AddDays(20),
                UtcDueDateSpecified = true
            };

            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to an email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);

            Site.Assert.AreEqual<byte>(
                1,
                getChangedResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate set flag success.");
            #endregion

            #region Verify requirements
            // If server returned a Status 1, that is to say, the sync command is successful, then MS-ASEMAIL_R940 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R940");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R940
            Site.CaptureRequirement(
                940,
                @"[In Updating E-Mail Flags] When Action is Flag an email, Required Properties from Device are: Status = 2, FlagType (section 2.2.2.35) = ""Flag for follow up"",tasks:StartDate (section 2.2.2.72) and tasks:UtcStartDate (section 2.2.2.87) tasks:DueDate (section 2.2.2.31) and tasks:UtcDueDate (section 2.2.2.86)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R693");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R693
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                item.Email.Flag.ReminderSet,
                693,
                @"[In ReminderSet] The default value is 0 (zero).");

            // If MS-ASEMAIL_R693 is captured successfully, it means the reminder is not set for the task, then MS-ASEMAIL_R1072 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1072");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1072
            Site.CaptureRequirement(
                1072,
                @"[In ReminderSet] [The value of this element[tasks:ReminderSet] is set to 1 if a reminder has been set for the task;] otherwise, the value of this element[tasks:ReminderSet] is set to 0 (zero).");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC12_MarkEmailComplete
        /// <summary>
        /// This case is designed to test when action is marking an email complete, required properties from device are: Status = 1, CompleteTime and tasks: DateCompleted.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC12_MarkEmailComplete()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to mark an email as complete.
            Request.Flag flag = new Request.Flag
            {
                Status = "1",
                CompleteTime = DateTime.Now,
                CompleteTimeSpecified = true,
                DateCompleted = DateTime.Now,
                DateCompletedSpecified = true
            };

            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to an email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            Site.Assert.AreEqual<byte>(
                1,
                getChangedResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response to indicate sets flag success.");
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirements
            // If server returned a Status 1, that is to say, the sync command is successful, then MS-ASEMAIL_R86 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R86");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R86
            Site.CaptureRequirement(
                86,
                @"[In Updating E-Mail Flags] When Action is Mark an email flag as complete, Required Properties from Device are: Status = 1, CompleteTime (section 2.2.2.19), tasks:DateCompleted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R334");

            // If the CompleteTime is specified in response, then MS-ASEMAIL_R334 can be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R334
            Site.CaptureRequirementIfIsTrue(
                item.Email.Flag.CompleteTimeSpecified,
                334,
                @"[In CompleteTime] The CompleteTime element is required to mark a flagged item as complete.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC13_ClearEmailFlag_Status0
        /// <summary>
        /// This case is designed to test when action is clearing the flag on an email, required properties from device are: Status = 0.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC13_ClearEmailFlag_Status0()
        {
            Request.Flag newFlag = new Request.Flag { Status = "0" };

            // If server returned an empty Flag node, that is to say, the flag was cleared, then MS-ASEMAIL_R88 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R88");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R88
            Site.CaptureRequirementIfIsTrue(
                this.VerifyClearEmailFlag(newFlag),
                88,
                @"[In Updating E-Mail Flags] When Action is Clearing the flag on an email, Required Properties from Device are: Status=0.");
        }
        #endregion

        #region MSASEMAIL_S03_TC14_ClearEmailFlag_EmptyNode
        /// <summary>
        /// This case is designed to test when action is clearing the flag on an email, required properties from device are: Flag node empty.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC14_ClearEmailFlag_EmptyNode()
        {
            Request.Flag newFlag = new Request.Flag();

            // If server returned an empty Flag node, that is to say, the flag was cleared, then MS-ASEMAIL_R8800 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R8800");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R8800
            Site.CaptureRequirementIfIsTrue(
                this.VerifyClearEmailFlag(newFlag),
                8800,
                @"[In Updating E-Mail Flags] When Action is Clearing the flag on an email, Required Properties from Device are: Flag node empty.");
        }
        #endregion

        #region MSASEMAIL_S03_TC15_UpdateEmailFlagMetadata
        /// <summary>
        /// This case is designed to test when action is updating the flag metadata on an email, required properties from device are: All updated properties.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC15_UpdateEmailFlagMetadata()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add email flag.
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                StartDate = DateTime.Now,
                StartDateSpecified = true,
                UtcStartDate = DateTime.Now,
                UtcStartDateSpecified = true,
                UtcDueDate = DateTime.Now.AddDays(20),
                UtcDueDateSpecified = true,
                DueDate = DateTime.Now.AddDays(20),
                DueDateSpecified = true
            };

            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to an email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);

            Site.Assert.AreEqual<byte>(
                1,
                getChangedResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response to indicate sets flag success.");
            #endregion

            #region Call Sync command with Change element to update the email flag.
            Request.Flag newFlag = new Request.Flag { Status = "2", FlagType = "for Follow Up" };

            // Update flag
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, null, item.ServerId, newFlag, null);

            // Get update result from server 
            SyncStore getUpdateResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(getUpdateResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The email item with subject '{0}' should be found in the {1} folder of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            Site.Assert.AreEqual<byte>(
                1,
                getUpdateResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response to indicate clears flag success.");
            #endregion

            #region Verify requirements
            if (Common.IsRequirementEnabled(1021, this.Site))
            {
                // If the update flag operation executes successfully, then requirement MS-ASEMAIL_R1021 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1021");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1021
                Site.CaptureRequirement(
                    1021,
                    @"[In Appendix B: Product Behavior] Implementation does update the flag on the e-mail message by using the Change element ([MS-ASCMD] section 2.2.3.24) of the Sync command, for every flag update that the client sends to the server in a Sync command request ([MS-ASCMD] section 2.2.2.19). (Exchange Server 2007 SP1 and above follow this behavior.)");
            }

            // If the FlagType value returned from server is "for Follow Up", that is to say, FlagType has been updated successfully, then MS-ASEMAIL_R90 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R90");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R90
            Site.CaptureRequirementIfAreEqual<string>(
                newFlag.FlagType,
                item.Email.Flag.FlagType,
                90,
                @"[In Updating E-Mail Flags] When Action is Update the flag metadata on an email, Required Properties from Device are: All updated properties.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC16_UpdateEmailFlag_Error
        /// <summary>
        /// This case is designed to test the Sync command response includes an airsync: Status element value of 6, if any of the required elements are missing from the Sync command request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC16_UpdateEmailFlag_Error()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add a flag with the StartDate, DueDate, UtcStartDate and UtcDueDate.
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                StartDate = DateTime.Now,
                StartDateSpecified = true,
                UtcStartDate = DateTime.Now,
                UtcStartDateSpecified = true,
                UtcDueDate = DateTime.Now.AddDays(20),
                UtcDueDateSpecified = true,
                DueDate = DateTime.Now.AddDays(20),
                DueDateSpecified = true
            };

            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);

            // Get the new added email item
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);

            // Add flag to the email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            #endregion

            #region Call Sync command with Change element to update the email flag.
            Request.Flag newFlag = new Request.Flag { Status = "2" };

            SyncStore updateResult = this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, null, item.ServerId, newFlag, null);
            Site.Assert.AreEqual<int>(
                6,
                int.Parse(updateResult.ChangeResponses[0].Status),
                "The server should return a Status 6 in the Sync command response if any of the required elements:StartDate, DueDate, UtcStartDate and UtcDueDate are missing from the Sync command request.");
            #endregion

            #region Verify requirement
            // If server returned a Status 6, then MS-ASEMAIL_R92 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R92");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R92
            Site.CaptureRequirement(
                92,
                @"[In Updating E-Mail Flags] The Sync command response includes an airsync:Status element ([MS-ASCMD] section 2.2.3.177.17) value of 6 if any of the required elements listed in the table are missing from the Sync command request.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC17_IncludedTwoDueDateElement
        /// <summary>
        /// This case is designed to test that the server does not return a protocol status error in response if including two DueDate elements per Flag in Sync command request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC17_IncludedTwoDueDateElement()
        {
            #region Update email with additional DueDate element
            string insertTag = "<DueDate xmlns=\"Tasks\">";
            string dueDateElement = "<DueDate xmlns=\"Tasks\">" + DateTime.UtcNow.AddDays(20).ToString("yyyy-MM-ddThh:mm:ss.fffZ") + "</DueDate>";

            // Get status code from the update flag result
            string statusCode = this.FlagEmailWithAddtionalFlagElement(dueDateElement, insertTag);
            #endregion

            #region Verify requirement
            if (Common.IsRequirementEnabled(581, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R581");

                // If the Status returned from server response is null, it indicates the server doesn't return a protocol error then MS-ASEMAIL_R581 can be captured.
                // Verify MS-ASEMAIL requirement:MS-ASEMAIL_R581
                Site.CaptureRequirementIfIsNull(
                    statusCode,
                    581,
                    @"[In Appendix B: Product Behavior] Implementation does not return a protocol status error in response to such a command request [A maximum of one tasks:DueDate child element is allowed per Flag element. The result of including more than one tasks:DueDate child element per Flag element is undefined.]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC18_IncludedTwoOrdinalDateElement
        /// <summary>
        /// This case is designed to test the server does not return a protocol status error in response if including two OrdinalDate elements per Flag in Sync command request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC18_IncludedTwoOrdinalDateElement()
        {
            #region Update email with additional OridinalDate element
            DateTime secondOrdinalDate = DateTime.Now;
            string insertTag = "<OrdinalDate xmlns=\"Tasks\">";
            string ordinalDateElement = "<OrdinalDate xmlns=\"Tasks\">" + secondOrdinalDate.ToString("yyyy-MM-ddThh:mm:ss.fffZ") + "</OrdinalDate>";
            string statusCode = this.FlagEmailWithAddtionalFlagElement(ordinalDateElement, insertTag);
            #endregion

            #region Verify requirement
            if (Common.IsRequirementEnabled(634, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R634");

                // If the Status returned from server response is null, it indicates the server doesn't return a protocol error then MS-ASEMAIL_R634 can be captured
                // Verify MS-ASEMAIL requirement:MS-ASEMAIL_R634
                Site.CaptureRequirementIfIsNull(
                    statusCode,
                    634,
                    @"[In Appendix B: Product Behavior] Implementation does not return a protocol status error in response to such a command request [A maximum of one tasks:OrdinalDate child element is allowed per Flag element. The result of including more than one tasks:OrdinalDate child element per Flag element is undefined.]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }

        #endregion

        #region MSASEMAIL_S03_TC19_IncludedTwoReminderSetElement
        /// <summary>
        /// This case is designed to test that the server does not return a protocol status error in response if including two ReminderSet elements per Flag in Sync command request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC19_IncludedTwoReminderSetElement()
        {
            #region Update email with additional ReminderSet element
            byte secondReminderSet = 1;
            string insertTag = "<ReminderSet xmlns=\"Tasks\">";
            string reminderSetElement = "<ReminderSet xmlns=\"Tasks\">" + secondReminderSet.ToString() + "</ReminderSet>";

            // Get status code from the update flag operation result
            string statusCode = this.FlagEmailWithAddtionalFlagElement(reminderSetElement, insertTag);
            #endregion

            #region Verify requirement
            if (Common.IsRequirementEnabled(697, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R697");

                // If the Status returned from server response is null, it indicates the server doesn't return a protocol error then MS-ASEMAIL_R697 can be captured
                // Verify MS-ASEMAIL requirement:MS-ASEMAIL_R697
                Site.CaptureRequirementIfIsNull(
                    statusCode,
                    697,
                    @"[In Appendix B: Product Behavior] Implementation does not return a protocol status error in response to such a command request [A maximum of one tasks:ReminderSet child element is allowed per Flag element. The result of including more than one tasks:ReminderSet child element per Flag element is undefined.]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC20_IncludedTwoReminderTimeElement
        /// <summary>
        /// This case is designed to test that the server doesn't return a protocol status error in response if including two ReminderTime elements per Flag in Sync command request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC20_IncludedTwoReminderTimeElement()
        {
            #region Update email with additional ReminderTime element
            DateTime secondRemindTime = DateTime.Now.AddDays(12);
            string insterTag = "<ReminderTime xmlns=\"Tasks\">";
            string reminderTimeElement = "<ReminderTime xmlns=\"Tasks\">" + secondRemindTime.ToString("yyyy-MM-ddThh:mm:ss.fffZ") + "</ReminderTime>";

            // After Send mail then update mail
            string statusCode = this.FlagEmailWithAddtionalFlagElement(reminderTimeElement, insterTag);
            #endregion

            #region Verify requirement
            if (Common.IsRequirementEnabled(706, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R706");

                // If the Status returned from server response is null, it indicates the server doesn't return a protocol error then MS-ASEMAIL_R706 can be captured.
                // Verify MS-ASEMAIL requirement:MS-ASEMAIL_R706
                Site.CaptureRequirementIfIsNull(
                    statusCode,
                    706,
                    @"[In Appendix B: Product Behavior] Implementation does not return a protocol status error in response to such a command request [A maximum of one tasks:ReminderTime child element is allowed per Flag element. The result of including more than one tasks:ReminderTime child element per Flag element is undefined.]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC21_IncludedTwoStartDateElement
        /// <summary>
        /// This case is designed to test that the server does not return a protocol status error in response if including two StartDate elements per Flag in Sync command request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC21_IncludedTwoStartDateElement()
        {
            #region Update email with the additional StartDate element
            DateTime secondStartDate = DateTime.Now;

            // Additional startDate elements
            string insertTag = "<StartDate xmlns=\"Tasks\">";
            string reminderTimeElement = "<StartDate xmlns=\"Tasks\">" + secondStartDate.ToString("yyyy-MM-ddThh:mm:ss.fffZ") + "</StartDate>";

            // Send mail and update email flag
            string statusCode = this.FlagEmailWithAddtionalFlagElement(reminderTimeElement, insertTag);
            #endregion

            #region Verify requirement
            if (Common.IsRequirementEnabled(746, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R746");

                // If the Status returned from server response is null, it indicates the server doesn't return a protocol error then MS-ASEMAIL_R746 can be captured.
                // Verify MS-ASEMAIL requirement:MS-ASEMAIL_R746
                Site.CaptureRequirementIfIsNull(
                    statusCode,
                    746,
                    @"[In Appendix B: Product Behavior] Implementation does not return a protocol status error in response to such a command request [A maximum of one tasks:StartDate child element is allowed per Flag element. The result of including more than one tasks:StartDate child element per Flag element is undefined.]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC22_IncludedTwoStatusElement
        /// <summary>
        /// This case is designed to test that the server does not return a protocol status error in response if including two Status elements per Flag in Sync command request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC22_IncludedTwoStatusElement()
        {
            #region Update email with the additional Status element
            string secondStatus = "1";

            // Update email with additional status elements
            string insertTag = "<Status>";
            string sencondStatusElement = "<Status>" + secondStatus + "</Status>";

            // Send email and update email
            string statusCode = this.FlagEmailWithAddtionalFlagElement(sencondStatusElement, insertTag);
            #endregion

            #region Verify requirement
            if (Common.IsRequirementEnabled(761, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R761");

                // If the Status returned from server response is null, it indicates the server doesn't return a protocol error then MS-ASEMAIL_R761 can be captured.
                // Verify MS-ASEMAIL requirement:MS-ASEMAIL_R761
                Site.CaptureRequirementIfIsNull(
                    statusCode,
                    761,
                    @"[In Appendix B: Product Behavior] Implementation does not return a protocol status error in response to such a command request [A maximum of one Status element is allowed per Flag. The result of including more than one Status element per Flag is undefined.]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC23_IncludedTwoSubOrdinalDateElement
        /// <summary>
        /// This case is designed to test that the server does not return a protocol status error in response if including two SubOrdinalDate elements per Flag in Sync command request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC23_IncludedTwoSubOrdinalDateElement()
        {
            #region Update email flag with the additional SubOrdinalDate element
            string secondSubOrdinalDate = DateTime.Now.ToString();
            string insertTag = "<SubOrdinalDate xmlns=\"Tasks\">";
            string subOrdinalDateElement = "<SubOrdinalDate xmlns=\"Tasks\">" + secondSubOrdinalDate + "</SubOrdinalDate>";

            // Send email ,then update email flag
            string statusCode = this.FlagEmailWithAddtionalFlagElement(subOrdinalDateElement, insertTag);
            #endregion

            #region Verify requirement
            if (Common.IsRequirementEnabled(777, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R777");

                // If the Status returned from server response is null, it indicates the server doesn't return a protocol error then MS-ASEMAIL_R777 can be captured.
                // Verify MS-ASEMAIL requirement:MS-ASEMAIL_R777
                Site.CaptureRequirementIfIsNull(
                    statusCode,
                    777,
                    @"[In Appendix B: Product Behavior] Implementation does not return a protocol status error in response to such a command request [A maximum of one tasks:SubOrdinalDate child element is allowed per Flag element. The result of including more than one tasks:SubOrdinalDate child element per Flag element is undefined.]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC24_IncludedTwoUtcDueDateElement
        /// <summary>
        /// This case is designed to test the server does not return a protocol status error in response if including two UtcDueDate elements per Flag in Sync command request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC24_IncludedTwoUtcDueDateElement()
        {
            #region Update email with the additional UtcDueDate element
            DateTime secondUtcDueDate = DateTime.UtcNow.AddDays(20);
            string insertTag = "<UtcDueDate xmlns=\"Tasks\">";
            string secondUteStartDateElement = "<UtcDueDate xmlns=\"Tasks\">" + secondUtcDueDate.ToString("yyyy-MM-ddThh:mm:ss.fffZ") + "</UtcDueDate>";
            string statusCode = this.FlagEmailWithAddtionalFlagElement(secondUteStartDateElement, insertTag);
            #endregion

            #region Verify requirement
            if (Common.IsRequirementEnabled(858, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R858");

                // If the Status returned from server response is null, it indicates the server doesn't return a protocol error then MS-ASEMAIL_R858 can be captured.
                // Verify MS-ASEMAIL requirement:MS-ASEMAIL_R858
                Site.CaptureRequirementIfIsNull(
                    statusCode,
                    858,
                    @"[In Appendix B: Product Behavior] Implementation does not return a protocol status error in response to such a command request [A maximum of one tasks:UtcDueDate child element is allowed per Flag element. The result of including more than one tasks:UtcDueDate child elements per Flag element is undefined.]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC25_IncludedTwoUtcStartDateElement
        /// <summary>
        /// This case is designed to test that the server does not return a protocol status error in response if including two UtcStartDate elements per Flag in Sync command request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC25_IncludedTwoUtcStartDateElement()
        {
            #region Update email with the additional UtcStartDate element
            DateTime secondUteStartDate = DateTime.UtcNow;
            string insertTag = "<UtcStartDate xmlns=\"Tasks\">";
            string secondUteStartDateElement = "<UtcStartDate xmlns=\"Tasks\">" + secondUteStartDate.ToString("yyyy-MM-ddThh:mm:ss.fffZ") + "</UtcStartDate>";

            // Get status code from update result
            string statusCode = this.FlagEmailWithAddtionalFlagElement(secondUteStartDateElement, insertTag);
            #endregion

            #region Verify requirement
            if (Common.IsRequirementEnabled(869, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R869");

                // If the Status returned from server response is null, it indicates the server doesn't return a protocol error then MS-ASEMAIL_R869 can be captured.
                // Verify MS-ASEMAIL requirement:MS-ASEMAIL_R869
                Site.CaptureRequirementIfIsNull(
                    statusCode,
                    869,
                    @"[In Appendix B: Product Behavior] Implementation does not return a protocol status error in response to such a command request [A maximum of one tasks:UtcStartDate child element is allowed per Flag element. The result of including more than one tasks:UtcStartDate child elements per Flag element is undefined.]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC26_FlagEmail_RequiredPropertiesOption2
        /// <summary>
        /// This case is designed to test when action is flagging an email, required properties from device are: Status = 2, FlagType = "Flag for follow up", tasks:DateCompleted, CompleteTime.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC26_FlagEmail_RequiredPropertiesOption2()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag with properties Status = 2, FlagType = "Flag for follow up", tasks:DateCompleted, CompleteTime.
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                DateCompletedSpecified = true,
                DateCompleted = DateTime.Now.AddDays(1),
                CompleteTimeSpecified = true,
                CompleteTime = DateTime.Now.AddDays(1)
            };

            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to an email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);

            Site.Assert.AreEqual<byte>(
                1,
                getChangedResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate set flag success.");
            #endregion

            #region Verify requirements
            // If server returned a Status 1, that is to say, the sync command is successful, then MS-ASEMAIL_R941 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R941");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R941
            Site.CaptureRequirement(
                941,
                @"[In Updating E-Mail Flags] When Action is Flag an email, Required Properties from Device are:Status = 2, FlagType (section 2.2.2.35) = ""Flag for follow up"", tasks:DateCompleted (section 2.2.2.23), CompleteTime (section 2.2.2.19).");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC27_FlagTask_RequiredProperties
        /// <summary>
        /// This case is designed to test when action is Flag a task.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC27_FlagTask_RequiredProperties()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag with properties Status = 2,tasks:Subject,FlagType = "Flag for follow up",tasks:StartDate and tasks:UtcStartDate,tasks:DueDate and tasks:UtcDueDate,tasks:ReminderSet,tasks:ReminderTime
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                Subject = Common.GenerateResourceName(Site, "FlagSubject"),
                StartDate = DateTime.Now,
                StartDateSpecified = true,
                UtcStartDate = DateTime.Now,
                UtcStartDateSpecified = true,
                DueDate = DateTime.Now.AddDays(20),
                DueDateSpecified = true,
                UtcDueDate = DateTime.Now.AddDays(20),
                UtcDueDateSpecified = true,
                ReminderSet=1,
                ReminderSetSpecified=true,
                ReminderTime = DateTime.Now,
                ReminderTimeSpecified=true
            };

            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to an email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);

            Site.Assert.AreEqual<byte>(
                1,
                getChangedResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate set flag success.");
            #endregion

            #region Verify requirements
            // If server returned a Status 1, that is to say, the sync command is successful, then MS-ASEMAIL_R942 can be captured.
            // Add the debug information
            this.Site.CaptureRequirement(
                942,
                @"[In Updating E-Mail Flags] When Action is Flag a task, Required Properties from Device are: Status = 2, tasks:Subject (section 2.2.2.75.2) = user defined, FlagType = ""Flag for follow up"", tasks:StartDate and tasks:UtcStartDate, tasks:DueDate and tasks:UtcDueDate, tasks:ReminderSet (section 2.2.2.64), tasks:ReminderTime (section 2.2.2.65).");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC28_ClearTaskFlag_Status0
        /// <summary>
        /// This case is designed to test when action is clearing the flag on a task, required properties from device are: Status = 0.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC28_ClearTaskFlag_Status0()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag task.
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                Subject = Common.GenerateResourceName(Site, "FlagSubject"),
                StartDate = DateTime.Now,
                StartDateSpecified = true,
                UtcStartDate = DateTime.Now,
                UtcStartDateSpecified = true,
                DueDate = DateTime.Now.AddDays(20),
                DueDateSpecified = true,
                UtcDueDate = DateTime.Now.AddDays(20),
                UtcDueDateSpecified = true,
                ReminderSet = 1,
                ReminderSetSpecified = true,
                ReminderTime = DateTime.Now,
                ReminderTimeSpecified = true
            };

            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to an email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);

            Site.Assert.AreEqual<byte>(
                1,
                getChangedResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate set flag success.");
            #endregion

            #region Call method Sync to clear the flag added in the above step.
            Request.Flag newFlag = new Request.Flag { Status = "0" };

            // Clear flag
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, null, item.ServerId, newFlag, null);

            // Get update result from server 
            SyncStore getClearResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(getClearResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The email item with subject '{0}' should be found.", emailSubject);
            Site.Assert.AreEqual<byte>(
                1,
                getClearResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response to indicate clears flag success.");
            #endregion

            #region Verify requirements
            bool isVerifiedR89 =
                item.Email.Flag.Status == null &&
                item.Email.Flag.FlagType == null
                && string.IsNullOrEmpty(item.Email.Flag.Subject) &&
                !item.Email.Flag.StartDateSpecified &&
                !item.Email.Flag.DueDateSpecified &&
                !item.Email.Flag.UtcStartDateSpecified &&
                !item.Email.Flag.UtcDueDateSpecified &&
                !item.Email.Flag.ReminderSetSpecified &&
                !item.Email.Flag.ReminderTimeSpecified;
            
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR89,
                89,
                @"[In Updating E-Mail Flags] When Action is Clearing the flag on a task, Required Properties from Device are: Status = 0.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC29_ClearTaskFlag_EmptyNode
        /// <summary>
        /// This case is designed to test when action is clearing the flag on a task, required properties from device are: Flag node empty.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC29_ClearTaskFlag_EmptyNode()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag task.
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                Subject = Common.GenerateResourceName(Site, "FlagSubject"),
                StartDate = DateTime.Now,
                StartDateSpecified = true,
                UtcStartDate = DateTime.Now,
                UtcStartDateSpecified = true,
                DueDate = DateTime.Now.AddDays(20),
                DueDateSpecified = true,
                UtcDueDate = DateTime.Now.AddDays(20),
                UtcDueDateSpecified = true,
                ReminderSet = 1,
                ReminderSetSpecified = true,
                ReminderTime = DateTime.Now,
                ReminderTimeSpecified = true
            };

            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to an email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);

            Site.Assert.AreEqual<byte>(
                1,
                getChangedResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate set flag success.");
            #endregion

            #region Call method Sync to clear the flag added in the above step.
            Request.Flag newFlag = new Request.Flag();

            // Clear flag
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, null, item.ServerId, newFlag, null);

            // Get update result from server 
            SyncStore getClearResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(getClearResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The email item with subject '{0}' should be found.", emailSubject);
            Site.Assert.AreEqual<byte>(
                1,
                getClearResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response to indicate clears flag success.");
            #endregion

            #region Verify requirements
            bool isVerifiedR8900 =
                item.Email.Flag.Status == null &&
                item.Email.Flag.FlagType == null
                && string.IsNullOrEmpty(item.Email.Flag.Subject) &&
                !item.Email.Flag.StartDateSpecified &&
                !item.Email.Flag.DueDateSpecified &&
                !item.Email.Flag.UtcStartDateSpecified &&
                !item.Email.Flag.UtcDueDateSpecified &&
                !item.Email.Flag.ReminderSetSpecified &&
                !item.Email.Flag.ReminderTimeSpecified;
            
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR8900,
                8900,
                @"[In Updating E-Mail Flags] When Action is Clearing the flag on a task, Required Properties from Device are: Flag node empty.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC30_MarkTaskComplete
        /// <summary>
        /// This case is designed to test when action is marking a task complete, required properties from device are: Status = 1, CompleteTime and tasks: DateCompleted.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC30_MarkTaskComplete()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag task.
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                Subject = Common.GenerateResourceName(Site, "FlagSubject"),
                StartDate = DateTime.Now,
                StartDateSpecified = true,
                UtcStartDate = DateTime.Now,
                UtcStartDateSpecified = true,
                DueDate = DateTime.Now.AddDays(20),
                DueDateSpecified = true,
                UtcDueDate = DateTime.Now.AddDays(20),
                UtcDueDateSpecified = true,
                ReminderSet = 1,
                ReminderSetSpecified = true,
                ReminderTime = DateTime.Now,
                ReminderTimeSpecified = true
            };

            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to an email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);

            Site.Assert.AreEqual<byte>(
                1,
                getChangedResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate set flag success.");
            #endregion

            #region Call Sync command with Change element to mark a task as complete.
            flag = new Request.Flag
            {
                Status = "1",
                CompleteTime = DateTime.Now,
                CompleteTimeSpecified = true,
                DateCompleted = DateTime.Now,
                DateCompletedSpecified = true
            };

            // Get the email item
            getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to an email item
            getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            Site.Assert.AreEqual<byte>(
                1,
                getChangedResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response to indicate sets flag success.");
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            #endregion

            #region Verify requirements
            // If server returned a Status 1, that is to say, the sync command is successful, then MS-ASEMAIL_R87 can be captured.
            // Add the debug information
            this.Site.CaptureRequirement(
                87,
                @"[In Updating E-Mail Flags] When Action is Mark a task flag as complete, Required Properties from Device are:  Status = 1, CompleteTime, tasks:DateCompleted.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S03_TC31_UpdateTaskFlagMetadata
        /// <summary>
        /// This case is designed to test when action is updating the flag metadata on an task, required properties from device are: All updated properties.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S03_TC31_UpdateTaskFlagMetadata()
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call Sync command with Change element to add flag with properties Status = 2,tasks:Subject,FlagType = "Flag for follow up",tasks:StartDate and tasks:UtcStartDate,tasks:DueDate and tasks:UtcDueDate,tasks:ReminderSet,tasks:ReminderTime
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "Flag for follow up",
                Subject = Common.GenerateResourceName(Site, "FlagSubject"),
                StartDate = DateTime.Now,
                StartDateSpecified = true,
                UtcStartDate = DateTime.Now,
                UtcStartDateSpecified = true,
                DueDate = DateTime.Now.AddDays(20),
                DueDateSpecified = true,
                UtcDueDate = DateTime.Now.AddDays(20),
                UtcDueDateSpecified = true,
                ReminderSet = 1,
                ReminderSetSpecified = true,
                ReminderTime = DateTime.Now,
                ReminderTimeSpecified = true
            };

            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to an email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);

            Site.Assert.AreEqual<byte>(
                1,
                getChangedResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate set flag success.");
            #endregion

            #region Call Sync command with Change element to update the email flag.
            Request.Flag newFlag = new Request.Flag { Status = "2", FlagType = "for Follow Up" };

            // Update flag
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, null, item.ServerId, newFlag, null);

            // Get update result from server 
            SyncStore getUpdateResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(getUpdateResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The email item with subject '{0}' should be found in the {1} folder of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);
            Site.Assert.AreEqual<byte>(
                1,
                getUpdateResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response to indicate clears flag success.");
            #endregion

            #region Verify requirements
            // If the FlagType value returned from server is "for Follow Up", that is to say, FlagType has been updated successfully, then MS-ASEMAIL_R91 can be captured.
            this.Site.CaptureRequirementIfAreEqual<string>(
                newFlag.FlagType,
                item.Email.Flag.FlagType,
                91,
                @"[In Updating E-Mail Flags] When Action is Update flag metadata on a task, Required Properties from Device are: All updated properties.");
            #endregion
        }
        #endregion
        #region Private methods
        /// <summary>
        /// Send mail and update email with additional Flag element
        /// </summary>
        /// <param name="additionalFlagElement">Additional flag element</param>
        /// <param name="insertTag">Additional element will insert before this tag</param>
        /// <returns>Status code after update email flag</returns>
        private string FlagEmailWithAddtionalFlagElement(string additionalFlagElement, string insertTag)
        {
            string emailSubject = Common.GenerateResourceName(Site, "subject");

            // Send email to user2
            this.SendPlaintextEmail(emailSubject, null, null);

            // Get the new added email item
            SyncStore syncChangeResult = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync emailItem = TestSuiteHelper.GetSyncAddItem(syncChangeResult, emailSubject);

            DateTime startDate = DateTime.Now;
            DateTime dueDate = DateTime.Now.AddDays(20);

            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "for Follow Up",
                Subject = emailSubject,
                StartDate = startDate,
                StartDateSpecified = true,
                UtcStartDate = startDate.ToUniversalTime(),
                UtcStartDateSpecified = true,
                DueDate = dueDate,
                DueDateSpecified = true,
                UtcDueDate = dueDate.ToUniversalTime(),
                UtcDueDateSpecified = true,
                OrdinalDateSpecified = true,
                ReminderSet = 1,
                ReminderSetSpecified = true,
                ReminderTimeSpecified = true
            };

            flag.SubOrdinalDate = flag.OrdinalDate.ToString();
            flag.OrdinalDate = DateTime.Now;
            flag.ReminderTime = DateTime.Now.AddDays(12);

            // Update email with additional elements
            SendStringResponse result = this.UpdateEmailWithMoreData(this.User2Information.InboxCollectionId, syncChangeResult.SyncKey, true, emailItem.ServerId, flag, null, additionalFlagElement, insertTag);

            return TestSuiteHelper.GetStatusCode(result.ResponseDataXML);
        }

        /// <summary>
        /// Verify the operation of clearing email flag.
        /// </summary>
        /// <param name="newFlag">The new flag of the email.</param>
        /// <returns>True indicates the flag is cleared, false indicates the flag is not cleared.</returns>
        private bool VerifyClearEmailFlag(Request.Flag newFlag)
        {
            #region Call SendMail command to send an email.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlaintextEmail(emailSubject, string.Empty, string.Empty);
            #endregion

            #region Call method Sync to add a flag.
            Request.Flag flag = new Request.Flag
            {
                Status = "2",
                FlagType = "for Follow Up",
                StartDate = DateTime.Now,
                StartDateSpecified = true,
                UtcStartDate = DateTime.Now,
                UtcStartDateSpecified = true,
                DueDate = DateTime.Now.AddDays(20),
                DueDateSpecified = true,
                UtcDueDate = DateTime.Now.AddDays(20),
                UtcDueDateSpecified = true
            };

            // Get the email item
            SyncStore getEmailItem = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);

            // Add flag to an email item
            SyncStore getChangedResult = this.AddEmailFag(this.User2Information, getEmailItem, emailSubject, flag);
            item = TestSuiteHelper.GetSyncChangeItem(getChangedResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The message with subject {0} should be found in the folder {1} of user {2}.", emailSubject, FolderType.Inbox.ToString(), this.User2Information.UserName);

            Site.Assert.AreEqual<byte>(
                1,
                getChangedResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response to indicate sets flag success.");
            #endregion

            #region Call method Sync to clear the flag added in the above step.
            // Clear flag
            this.UpdateEmail(this.User2Information.InboxCollectionId, getChangedResult.SyncKey, null, item.ServerId, newFlag, null);

            // Get update result from server 
            SyncStore getClearResult = this.SyncChanges(getChangedResult.SyncKey, this.User2Information.InboxCollectionId, null);
            item = TestSuiteHelper.GetSyncChangeItem(getClearResult, item.ServerId);
            Site.Assert.IsNotNull(item, "The email item with subject '{0}' should be found.", emailSubject);
            Site.Assert.AreEqual<byte>(
                1,
                getClearResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response to indicate clears flag success.");
            #endregion

            #region Check whether the flag was cleared
            // If server returned an empty Flag node, that is to say, the flag was cleared.
            return item.Email.Flag.Status == null && item.Email.Flag.FlagType == null && !item.Email.Flag.StartDateSpecified && !item.Email.Flag.DueDateSpecified && !item.Email.Flag.UtcStartDateSpecified && !item.Email.Flag.UtcDueDateSpecified;
            #endregion
        }

        /// <summary>
        /// Add flag for an email item.
        /// </summary>
        /// <param name="userInformation">The information for user</param>
        /// <param name="getEmailItem">The specified email item which should be added the flag.</param>
        /// <param name="emailSubject">The subject of the specified email item.</param>
        /// <param name="flag">The flag which should be added to the specified email item.</param>
        /// <returns>The result store of adding flag.</returns>
        private SyncStore AddEmailFag(UserInformation userInformation, SyncStore getEmailItem, string emailSubject, Request.Flag flag)
        {
            // Get the new added email item
            Sync emailItem = TestSuiteHelper.GetSyncAddItem(getEmailItem, emailSubject);
            Site.Assert.IsNotNull(emailItem, "The email item with subject '{0}' should be found.", emailSubject);

            // Add flag to the email item
            this.UpdateEmail(userInformation.InboxCollectionId, getEmailItem.SyncKey, null, emailItem.ServerId, flag, null);

            // Get the server changes of adding flag and category
            return this.SyncChanges(getEmailItem.SyncKey, userInformation.InboxCollectionId, null);
        }
        #endregion
    }
}