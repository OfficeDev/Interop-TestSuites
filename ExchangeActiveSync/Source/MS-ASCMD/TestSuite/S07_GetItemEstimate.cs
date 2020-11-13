namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test the GetItemEstimate command.
    /// </summary>
    [TestClass]
    public class S07_GetItemEstimate : TestSuiteBase
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
        /// This test case is used to verify the requirements related to a successful GetItemEstimate command.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC01_GetItemEstimate_Success()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send a MIME-formatted email from User1 to User2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null);

            #region Call method GetItemEstimate to get an estimate of the number of items in the Inbox folder.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User2Information.InboxCollectionId), false);

            Request.Options option = new Request.Options
            {
                Items = new object[] { "Email" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            GetItemEstimateRequest getItemEstimateRequest = TestSuiteBase.CreateGetItemEstimateRequest(this.LastSyncKey, this.User2Information.InboxCollectionId, new Request.Options[] { option });
            GetItemEstimateResponse getItemEstimateResponse = this.CMDAdapter.GetItemEstimate(getItemEstimateRequest);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The Response element in the GetItemEstimate command response should not be null.");
            #endregion

            #region Call Sync command to get all items in the Inbox folder.
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User2Information.InboxCollectionId);
            this.Sync(syncRequest, false);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            SyncResponse syncResponse = this.Sync(syncRequest, false);

            Response.SyncCollectionsCollectionCommands commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            Site.Assert.IsNotNull(commands, "The Commands element in the Sync command response should not be null.");
            Site.Assert.IsNotNull(commands.Add, "The Add element of the Commands element in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4134");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4134
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                getItemEstimateResponse.ResponseData.Response[0].Status,
                4134,
                @"[In Status(GetItemEstimate)] [When the scope is Global], [the cause of the status value 1 is] Server successfully completed command.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R171");

            // Verify MS-ASCMD requirement: MS-ASCMD_R171
            Site.CaptureRequirementIfAreEqual<int>(
                commands.Add.Length,
                Convert.ToInt32(getItemEstimateResponse.ResponseData.Response[0].Collection.Estimate),
                171,
                @"[In GetItemEstimate] The GetItemEstimate command gets an estimate of the number of items in a collection or folder on the server that have to be synchronized.");

            // If R171 has been verified, then the GetItemEstimate command gets an estimate of the number of items in a collection or folder. 
            // So R5056 will be verified.
            this.Site.CaptureRequirement(
                5056,
                @"[In Synchronizing Inbox, Calendar, Contacts, and Tasks Folders] [Command sequence for folder synchronization, order [3]:] The server responds to indicate how many items will be added, changed, or deleted, for each collection.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetItemEstimate command, if Options element is not included in a request, server will enumerate all of the items within the collection, without any filter.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC02_GetItemEstimate_WithoutOptions()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call method GetItemEstiamte with filters within the airsync:Options container to Email class to get an estimate of the number of the items in the Inbox folder.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId));
            Request.Options option = new Request.Options
            {
                Items = new object[] { "Email" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            GetItemEstimateRequest getItemEstimateRequest = TestSuiteBase.CreateGetItemEstimateRequest(this.LastSyncKey, this.User1Information.InboxCollectionId, new Request.Options[] { option });
            GetItemEstimateResponse getItemEstimateResponse = CMDAdapter.GetItemEstimate(getItemEstimateRequest);
            #endregion

            #region Call method GetItemEstimate without Options element to get an estimate of the number of items in the Inbox folder on server.
            getItemEstimateRequest = TestSuiteBase.CreateGetItemEstimateRequest(this.LastSyncKey, this.User1Information.InboxCollectionId, null);
            GetItemEstimateResponse getItemEstimateResponseNoOption = CMDAdapter.GetItemEstimate(getItemEstimateRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3543");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3543
            Site.CaptureRequirementIfAreEqual<string>(
                getItemEstimateResponse.ResponseData.Response[0].Collection.Estimate,
                getItemEstimateResponseNoOption.ResponseData.Response[0].Collection.Estimate,
                3543,
                @"[In Options(GetItemEstimate)] If the airsync:Options element is not included in a request, then the GetItemEstimate command (section 2.2.2.7) will enumerate all of the items within the collection, without any filter (up to a maximum of 512 items).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetItemEstimate command for a contact collection, if FilterType element is included, no error is returned.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC03_GetItemEstimate_Contacts_FilterType()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            #region Call GetItemEstimate with setting FilterType to get an estimate number of items in Contacts folder.
            GetItemEstimateResponse getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.ContactsCollectionId, 0);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2967");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2967
            // If the returned value of the Estimate in previous step is not null, then R2967 should be covered.
            Site.CaptureRequirementIfIsNotNull(
                getItemEstimateResponse.ResponseData.Response[0].Collection.Estimate,
                2967,
                @"[In FilterType(GetItemEstimate)] If a filter type is specified, then the server sends an estimate of the items within the filter specifications.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2973");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2973
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                2973,
                @"[In FilterType(GetItemEstimate)] However, if the airsync:FilterType element is included in a GetItemEstimate request for a contact collection, no error is returned.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetItemEstimate command for e-mail, the status should be correspond to the value of FilterType.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC04_GetItemEstimate_Email_FilterType()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send an email from user1 to user2
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);            
            this.SwitchUser(this.User2Information);
            this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User2Information.InboxCollectionId));

            #region Call GetItemEstimate with FilterType setting to 0 to get estimate number of all items in Inbox folder.
            GetItemEstimateResponse getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User2Information.InboxCollectionId, 0);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");
            int original = Convert.ToInt32(getItemEstimateResponse.ResponseData.Response[0].Collection.Estimate);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2999");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2999
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                2999,
                @"[In FilterType(GetItemEstimate)] Yes. [Applies to email, if FilterType is 0, Status element value is 1.]");
            #endregion

            #region Call GetItemEstimate without FilterType
            getItemEstimateResponse = this.CMDAdapter.GetItemEstimate(TestSuiteBase.CreateGetItemEstimateRequest(this.LastSyncKey, this.User2Information.InboxCollectionId, null));
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");
            int current = Convert.ToInt32(getItemEstimateResponse.ResponseData.Response[0].Collection.Estimate);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2970");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2970
            Site.CaptureRequirementIfAreEqual<int>(
                original,
                current,
                2970,
                @"[In FilterType(GetItemEstimate)] If the airsync:FilterType element is omitted, then all objects are sent from the server.");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 1 to get estimate number of the items within 1 day in Inbox folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User2Information.InboxCollectionId, 1);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");
            original = Convert.ToInt32(getItemEstimateResponse.ResponseData.Response[0].Collection.Estimate);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3000");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3000
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                3000,
                @"[In FilterType(GetItemEstimate)] Yes. [Applies to email, if FilterType is 1, Status element value is 1.]");
            #endregion

            #region Send an email from user1 to user2
            this.SwitchUser(this.User1Information);
            string anotherEmailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, anotherEmailSubject, this.User1Information.UserName, this.User2Information.UserName, null);

            this.SwitchUser(this.User2Information);
            this.CheckEmail(this.User2Information.InboxCollectionId, anotherEmailSubject, null);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, anotherEmailSubject);

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User2Information.InboxCollectionId));
            #endregion

            #region Call GetItemEstimate with FilterType setting to 1 to get estimate number of the items within 1 day in Inbox folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User2Information.InboxCollectionId, 1);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");
            current = Convert.ToInt32(getItemEstimateResponse.ResponseData.Response[0].Collection.Estimate);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2969");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2969
            Site.CaptureRequirementIfAreEqual<int>(
                original + 1,
                current,
                2969,
                @"[In FilterType(GetItemEstimate)] New objects are added to the client when they are within the time window.");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 2 to get estimate number of the items within 3 days in Inbox folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User2Information.InboxCollectionId, 2);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3001");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3001
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                3001,
                @"[In FilterType(GetItemEstimate)] Yes. [Applies to email, if FilterType is 2, Status element value is 1.]");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 3 to get estimate number of the items within 1 week in Inbox folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User2Information.InboxCollectionId, 3);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3002");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3002
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                3002,
                @"[In FilterType(GetItemEstimate)] Yes. [Applies to email, if FilterType is 3, Status element value is 1.]");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 4 to get estimate number of the items within 2 weeks in Inbox folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User2Information.InboxCollectionId, 4);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3003");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3003
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                3003,
                @"[In FilterType(GetItemEstimate)] Yes. [Applies to email, if FilterType is 4, Status element value is 1.]");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 5 to get estimate number of the items within 1 month in Inbox folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User2Information.InboxCollectionId, 5);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3004");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3004
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                3004,
                @"[In FilterType(GetItemEstimate)] Yes. [Applies to email, if FilterType is 5, Status element value is 1.]");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 6 to get estimate number of items within 3 months in Inbox folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User2Information.InboxCollectionId, 6);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3005");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3005
            Site.CaptureRequirementIfAreEqual<int>(
                110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3005,
                @"[In FilterType(GetItemEstimate)] No,[Applies to email, if FilterType is 6,] Status element value 110");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 7 to get estimate number of items within 6 months in Inbox folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User2Information.InboxCollectionId, 7);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3006");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3006
            Site.CaptureRequirementIfAreEqual<int>(
                110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3006,
                @"[In FilterType(GetItemEstimate)] No,[Applies to email, if FilterType is 7,] Status element value 110");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 8 to get estimate number of incomplete tasks in Inbox folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User2Information.InboxCollectionId, 8);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3007");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3007
            Site.CaptureRequirementIfAreEqual<int>(
                110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3007,
                @"[In FilterType(GetItemEstimate)] No, [Applies to email, if FilterType is 8,] Status element value is 110.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetItemEstimate command for calendar, the value of status should be correspond to the value of FilterType. 
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC05_GetItemEstimate_Calendar_FilterType()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Recurrences cannot be added in protocol version 16.0");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Recurrences cannot be added in protocol version 16.1");

            #region Add a new calendar
            string calendarSubject = Common.GenerateResourceName(Site, "calendarSubject");
            DateTime startTime = DateTime.Now.AddDays(1.0);
            DateTime endTime = startTime.AddMinutes(10.0);

            Request.SyncCollectionAdd calendarData = new Request.SyncCollectionAdd
            {
                ClientId = TestSuiteBase.ClientId,
                ApplicationData = new Request.SyncCollectionAddApplicationData
                {
                    ItemsElementName =
                        new Request.ItemsChoiceType8[]
                        {
                            Request.ItemsChoiceType8.Subject, 
                            Request.ItemsChoiceType8.StartTime, 
                            Request.ItemsChoiceType8.EndTime                 
                        },
                    Items =
                        new object[]
                        {
                            calendarSubject, 
                            startTime.ToString("yyyyMMddTHHmmssZ"),
                            endTime.ToString("yyyyMMddTHHmmssZ")
                        }
                },
                Class = "Calendar"
            };

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId));

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, calendarData);
            SyncResponse syncResponse = this.Sync(syncRequest);

            Response.SyncCollectionsCollectionResponses responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The calendar should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, calendarSubject);
            #endregion

            #region Call GetItemEstimate with FilterType setting to 0 to get estimate number of all items in Calendar folder.
            GetItemEstimateResponse getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.CalendarCollectionId, (byte)0);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3008");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3008
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                3008,
                @"[In FilterType(GetItemEstimate)] Yes. [Applies to calendar, if FilterType is 0, Status element value is 1.]");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 1 to get estimate number of items within 1 day in Calendar folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.CalendarCollectionId, (byte)1);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3009");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3009
            Site.CaptureRequirementIfAreEqual<int>(
                110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3009,
                @"[In FilterType(GetItemEstimate)] No, [Applies to calendar, if FilterType is 1,] Status element (section 2.2.3.162.6) value is 110.");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 2 to get estimate number of items within 3 days in Calendar folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.CalendarCollectionId, (byte)2);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3010");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3010
            Site.CaptureRequirementIfAreEqual<int>(
                110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3010,
                @"[In FilterType(GetItemEstimate)] No, [Applies to calendar, if FilterType is 2,] Status element value is 110.");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 3 to get estimate number of items within 1 week in Calendar folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.CalendarCollectionId, (byte)3);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3011");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3011
            Site.CaptureRequirementIfAreEqual<int>(
                110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3011,
                @"[In FilterType(GetItemEstimate)] No, [Applies to calendar, if FilterType is 3,] Status element value is 110.");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 4 to get estimate number of items within 2 weeks in Calendar folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.CalendarCollectionId, (byte)4);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3012");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3012
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                3012,
                @"[In FilterType(GetItemEstimate)] Yes. [Applies to calendar, if FilterType is 4, Status element value 1]");
            #endregion

            #region Create a future calendar
            this.GetInitialSyncResponse(this.User1Information.CalendarCollectionId);
            calendarSubject = Common.GenerateResourceName(this.Site, "canlendarSubject");
            startTime = DateTime.Now.AddDays(15.0);
            endTime = startTime.AddMinutes(10.0);

            calendarData = new Request.SyncCollectionAdd
            {
                ClientId = TestSuiteBase.ClientId,
                ApplicationData = new Request.SyncCollectionAddApplicationData
                {
                    ItemsElementName = new Request.ItemsChoiceType8[]
                    {
                        Request.ItemsChoiceType8.Subject,
                        Request.ItemsChoiceType8.StartTime,
                        Request.ItemsChoiceType8.EndTime,
                    },
                    Items = new object[]
                    {
                        calendarSubject,
                        startTime.ToString("yyyyMMddTHHmmssZ"),
                        endTime.ToString("yyyyMMddTHHmmssZ"),
                    }
                },
                Class = "Calendar"
            };

            syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, calendarData);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The calendar should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, calendarSubject);

            syncResponse = this.SyncChanges(this.User1Information.CalendarCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            bool isVerifyR2971 = !string.IsNullOrEmpty(TestSuiteBase.FindServerId(syncResponse, "Subject", calendarSubject));

            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 4);
            this.Sync(syncRequest);

            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            isVerifyR2971 = isVerifyR2971 && !string.IsNullOrEmpty(TestSuiteBase.FindServerId(syncResponse, "Subject", calendarSubject));

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2971");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2971
            Site.CaptureRequirementIfIsTrue(
                isVerifyR2971,
                2971,
                @"[In FilterType(GetItemEstimate)] Calendar items that are in the future [or that have recurrence, but no end date], are sent to the client regardless of the airsync:FilterType element value.");
            #endregion

            #region Create a recurrence calendar without EndTime
            string recurrenceCalendarSubject = Common.GenerateResourceName(Site, "recurrenceCanlendarSubject");

            Request.SyncCollectionAdd recurrenceCalendarData = new Request.SyncCollectionAdd
            {
                ClientId = TestSuiteBase.ClientId,
                ApplicationData = new Request.SyncCollectionAddApplicationData
                {
                    ItemsElementName = new Request.ItemsChoiceType8[] { Request.ItemsChoiceType8.Subject, Request.ItemsChoiceType8.Recurrence },
                    Items = new object[]
                    {
                        recurrenceCalendarSubject,
                        new Request.Recurrence
                        {
                            Type = 1,
                            OccurrencesSpecified = false,
                            DayOfWeek = 2,
                            DayOfWeekSpecified = true,
                            IsLeapMonthSpecified = false
                        },
                    }
                },
                Class = "Calendar"
            };

            syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, recurrenceCalendarData);
            syncResponse = this.Sync(syncRequest);
            responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The calendar should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, recurrenceCalendarSubject);

            syncResponse = this.SyncChanges(this.User1Information.CalendarCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            bool isVerifyR5877 = !string.IsNullOrEmpty(TestSuiteBase.FindServerId(syncResponse, "Subject", recurrenceCalendarSubject));

            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 4);
            this.Sync(syncRequest);

            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            isVerifyR5877 = isVerifyR5877 && !string.IsNullOrEmpty(TestSuiteBase.FindServerId(syncResponse, "Subject", recurrenceCalendarSubject));

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5877");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5877
            Site.CaptureRequirementIfIsTrue(
                isVerifyR5877,
                5877,
                @"[In FilterType(GetItemEstimate)] Calendar items [that are in the future or] that have recurrence, but no end date, are sent to the client regardless of the airsync:FilterType element value.");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 5 to get estimate number of items within 1 month in Calendar folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.CalendarCollectionId, (byte)5);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3013");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3013
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                3013,
                @"[In FilterType(GetItemEstimate)] Yes. [Applies to calendar, if FilterType is 5, Status element value is 1.]");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 6 to get estimate number of items within 3 months in Calendar folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.CalendarCollectionId, (byte)6);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3014");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3014
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                3014,
                @"[In FilterType(GetItemEstimate)] Yes. [Applies to calendar, if FilterType is 6, Status element value 1]");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 7 to get estimate number of items within 6 months in Calendar folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.CalendarCollectionId, (byte)7);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3015");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3015
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                3015,
                @"[In FilterType(GetItemEstimate)] Yes. [Applies to calendar, if FilterType is 7, Status element value is 1.]");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 8 to get estimate number of incomplete tasks in Calendar folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.CalendarCollectionId, (byte)8);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3016");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3016
            Site.CaptureRequirementIfAreEqual<int>(
                110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3016,
                @"[In FilterType(GetItemEstimate)] No, [Applies to calendar, if FilterType is 8,] Status element value is 110.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetItemEstimate command for tasks, the value of status should be correspond to the value of FilterType.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC06_GetItemEstimate_Tasks_FilterType()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.TasksCollectionId));

            #region Call GetItemEstimate with FilterType setting to 0 to get estimate number of all items in Tasks folder.
            GetItemEstimateResponse getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.TasksCollectionId, (byte)0);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3017");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3017
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                3017,
                @"[In FilterType(GetItemEstimate)] Yes. [Applies to tasks, if FilterType is 0, Status element value is 1.]");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 1 to get estimate number of items within 1 day in Tasks folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.TasksCollectionId, (byte)1);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3018");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3018
            Site.CaptureRequirementIfAreEqual<int>(
                110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3018,
                @"[In FilterType(GetItemEstimate)] No, [Applies to tasks, if FilterType is 1, ] Status element value is 110.");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 2 to get estimate number of items within 3 days in Tasks folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.TasksCollectionId, (byte)2);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3019");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3019
            Site.CaptureRequirementIfAreEqual<int>(
                (byte)110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3019,
                @"[In FilterType(GetItemEstimate)] No, [In FilterType] No, [Applies to tasks, if FilterType is 2, ] Status element value is 110.");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 3 to get estimate number of items within 1 week in Tasks folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.TasksCollectionId, (byte)3);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3020");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3020
            Site.CaptureRequirementIfAreEqual<int>(
                (byte)110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3020,
                @"[In FilterType(GetItemEstimate)] No, [In FilterType] No, [Applies to tasks, if FilterType is 3, ] Status element value is 110.");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 4 to get estimate number of items within 2 weeks in Tasks folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.TasksCollectionId, (byte)4);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3021");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3021
            Site.CaptureRequirementIfAreEqual<int>(
                110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3021,
                @"[In FilterType(GetItemEstimate)] No, [Applies to tasks, if FilterType is 4, ] Status element value is 110.");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 5 to get estimate number of items within 1 month in Tasks folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.TasksCollectionId, (byte)5);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3022");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3022
            Site.CaptureRequirementIfAreEqual<int>(
                110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3022,
                @"[In FilterType(GetItemEstimate)] No, [Applies to tasks, if FilterType is 5, ]Status element value 110");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 6 to get estimate number of items within 3 months in Tasks folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.TasksCollectionId, (byte)6);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3023");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3023
            Site.CaptureRequirementIfAreEqual<int>(
                110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3023,
                @"[In FilterType(GetItemEstimate)] No, [Applies to tasks, if FilterType is 6, ] Status element value is 110.");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 7 to get estimate number of items within 6 months in Tasks folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.TasksCollectionId, (byte)7);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3024");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3024
            Site.CaptureRequirementIfAreEqual<int>(
                110,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                3024,
                @"[In FilterType(GetItemEstimate)] No, [Applies to tasks, if FilterType is 7, ] Status element value is 110.");
            #endregion

            #region Call GetItemEstimate with FilterType setting to 8 to get estimate number of incomplete tasks in Tasks folder.
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.TasksCollectionId, (byte)8);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3025");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3025
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                3025,
                @"[In FilterType(GetItemEstimate)] Yes. [Applies to tasks, if FilterType is 8, Status element value is 1.]");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetItemEstimate command, if FilterType element is invalid, the status should be equal to 103.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC07_GetItemEstimate_InvalidFilterType()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            #region Call method GetItemEstimate with invalid FilterType value to get an estimate of the number of items in Contacts folder on server.
            GetItemEstimateResponse getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.ContactsCollectionId, 9);
            int statusOfContacts = int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status);
            Assert.AreEqual<int>(103, statusOfContacts, "Specifying a airsync:FilterType of 9 or above for when the CollectionId element identifies contact collection should result in a Status element value of 103.");
            #endregion

            #region Call method GetItemEstimate with invalid FilterType value to get an estimate of the number of items in Inbox folder on server.
            this.GetInitialSyncResponse(this.User1Information.InboxCollectionId);
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.InboxCollectionId, 9);
            int statusOfInbox = int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status);
            Assert.AreEqual<int>(103, statusOfInbox, "Specifying a airsync:FilterType of 9 or above for when the CollectionId element identifies e-mail collection should result in a Status element value of 103.");
            #endregion

            #region Call method GetItemEstimate with invalid FilterType value to get an estimate of the number of items in Calendar folder on server.
            this.GetInitialSyncResponse(this.User1Information.CalendarCollectionId);
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.CalendarCollectionId, 9);
            int statusOfCalendar = int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status);
            Assert.AreEqual<int>(103, statusOfCalendar, "Specifying a airsync:FilterType of 9 or above for when the CollectionId element identifies calendar collection should result in a Status element value of 103.");
            #endregion

            #region Call method GetItemEstimate with invalid FilterType value to get an estimate of the number of items in Tasks folder on server.
            this.GetInitialSyncResponse(this.User1Information.TasksCollectionId);
            getItemEstimateResponse = this.GetItemEstimateWithFilterType(this.User1Information.TasksCollectionId, 9);
            int statusOfTasks = int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status);
            Assert.AreEqual<int>(103, statusOfTasks, "Specifying a airsync:FilterType of 9 or above for when the CollectionId element identifies tasks collection should result in a Status element value of 103.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2984");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2984
            // When all above four steps are passed, it means that a airsync:FilterType of 9 or above 
            // can cause a Status element of 103 when the CollectionId element (section 2.2.3.30.1) 
            // identifies any email, contact, calendar or task collection.
            Site.CaptureRequirement(
                2984,
                @"[In FilterType(GetItemEstimate)] Specifying a airsync:FilterType of 9 or above for when the CollectionId element (section 2.2.3.30.1) identifies any email, contact, calendar or task collection results in a Status element value of 103.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetItemEstimate command, if the request includes more than one MaxItem element, the server doesn't return error.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC08_GetItemEstimate_MoreThanOneMaxItems()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with 2 MaxItems elements on RI.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.RecipientInformationCacheCollectionId));

            Request.Options option = new Request.Options
            {
                Items = new object[] { "2", "3" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.MaxItems, Request.ItemsChoiceType1.MaxItems }
            };

            GetItemEstimateRequest getItemEstimateRequest = TestSuiteBase.CreateGetItemEstimateRequest(this.LastSyncKey, this.User1Information.RecipientInformationCacheCollectionId, new Request.Options[] { option });
            GetItemEstimateResponse getItemEstimateResponse = CMDAdapter.GetItemEstimate(getItemEstimateRequest);

            if (Common.IsRequirementEnabled(3272, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3272");

                // Verify MS-ASCMD requirement: MS-ASCMD_R3272
                Site.CaptureRequirementIfAreEqual<int>(
                    1,
                    int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                    3272,
                    @"[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one airsync:MaxItems element as the child of the airsync:Options element is undefined]. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetItemEstimate command, if the specified collection id is invalid, server should return a status code 2.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC09_GetItemEstimate_Status2()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            #region Call method GetItemEstimate with two collection ids, one is invalid CollectionId, to get an estimate of the number of items in Contacts folder on server.
            GetItemEstimateRequest getItemEstimateRequest = new GetItemEstimateRequest
            {
                RequestData = new Request.GetItemEstimate
                {
                    Collections = new Request.GetItemEstimateCollection[]
                    { 
                        new Request.GetItemEstimateCollection
                        {
                            ItemsElementName=new Request.ItemsChoiceType10[]
                            {
                                Request.ItemsChoiceType10.SyncKey,
                                Request.ItemsChoiceType10.CollectionId,
                                
                            },
                             Items = new object[]{
                                this.LastSyncKey,
                                this.User1Information.ContactsCollectionId,
                             
                               
                            },
                            Options = new Request.Options[]{
                                new Request.Options()
                                {
                                    Items =new object[]
                                    {
                                        "Contacts"
                                    },
                                    ItemsElementName =new Request.ItemsChoiceType1[]
                                    {
                                        Request.ItemsChoiceType1.Class
                                    }
                                }
                            }
                        },
                        new Request.GetItemEstimateCollection
                        {
                             ItemsElementName = new Request.ItemsChoiceType10[]{
                                 Request.ItemsChoiceType10.SyncKey,
                                 Request.ItemsChoiceType10.CollectionId,                              
                        },
                        Items = new object[]
                        {
                            this.LastSyncKey,
                            Common.GenerateResourceName(Site, "InvalidCollectionId"),    
                        },
                        Options = new Request.Options[]
                            {
                                new Request.Options
                                {
                                    Items = new object[]
                                    {
                                        "Email"
                                    },
                                    ItemsElementName = new Request.ItemsChoiceType1[]
                                    {
                                        Request.ItemsChoiceType1.Class
                                    }
                                }
                            }
                        }
                    }
                }
            };

            GetItemEstimateResponse getItemEstimateResponse = CMDAdapter.GetItemEstimate(getItemEstimateRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4135");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4135
            // When the Status value is 2, it means one of the specified CollectionIds is invalid.
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                int.Parse(getItemEstimateResponse.ResponseData.Response[1].Status),
                4135,
                @"[In Status(GetItemEstimate)] [When the scope is] Item, [the meaning of the status value] 2 [is] A collection was invalid or one of the specified collection IDs was invalid.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4136");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4136
            // When the Status value is 2, it means one or more of the specified folders does not exist or an incorrect folder is requested..
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                int.Parse(getItemEstimateResponse.ResponseData.Response[1].Status),
                4136,
                @"[In Status(GetItemEstimate)] [When the scope is Item], [the cause of the status value 2 is] One or more of the specified folders does not exist or an incorrect folder was requested.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4130");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4130
            bool isVerifyR4130 = int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status) == 1 && int.Parse(getItemEstimateResponse.ResponseData.Response[1].Status) != 1;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR4130,
                4130,
                @"[In Status(GetItemEstimate)] However, if the failure occurs at the Collection (section 2.2.3.29.1) level, a Status value is returned per Collection as a child of the Response element.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetItemEstimate command, if the synchronization state has not been primed, server should return a status code 3.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC10_GetItemEstimate_Status3()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call method GetItemEstimate without priming the synchronization state.
            Request.Options option = new Request.Options
            {
                Items = new object[] { "Email" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            GetItemEstimateRequest getItemEstimateRequest = TestSuiteBase.CreateGetItemEstimateRequest("0", this.User1Information.InboxCollectionId, new Request.Options[] { option });
            GetItemEstimateResponse getItemEstimateResponse = CMDAdapter.GetItemEstimate(getItemEstimateRequest);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4139");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4139
            // When the Status value is 3, it means the synchronization state has not been primed.
            Site.CaptureRequirementIfAreEqual<int>(
                3,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                4139,
                @"[In Status(GetItemEstimate)] [When the scope is] Item, [the meaning of the status value] 3 [is] The synchronization state has not been primed.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4140");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4140
            // When the Status value is 3, it means the issued GetItemEstimate command without first issuing a Sync command request with a SyncKey value of 0.
            Site.CaptureRequirementIfAreEqual<int>(
                3,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                4140,
                @"[In Status(GetItemEstimate)] [When the scope is Item], [the cause of the status value 3 is] The client has issued a GetItemEstimate command without first issuing a Sync command request (section 2.2.2.19) with an airsync:SyncKey element (section 2.2.3.156.4) value of zero (0).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetItemEstimate command, if the synchronization key is invalid, server should return a status code 4.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC11_GetItemEstimate_Status4_InvalidSyncKey()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId));

            #region Call method GetItemEstimate with an invalid SyncKey.
            string invalidSyncKey = Common.GenerateResourceName(Site, "InvalidSyncKey");

            Request.Options option = new Request.Options
            {
                Items = new object[] { "Email" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            GetItemEstimateRequest getItemEstimateRequest = TestSuiteBase.CreateGetItemEstimateRequest(invalidSyncKey, this.User1Information.InboxCollectionId, new Request.Options[] { option });
            GetItemEstimateResponse getItemEstimateResponse = CMDAdapter.GetItemEstimate(getItemEstimateRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4142");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4142
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                4142,
                @"[In Status(GetItemEstimate)] [When the scope is] Global, [the meaning of the status value] 4 [is] The specified synchronization key was invalid.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetItemEstimate command, if the SyncKey value provided in the GetItemEstimate request does not match those expected within the next Sync command request, server should return a status value of 4.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC12_GetItemEstimate_Status4_MismatchedSyncKey()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));
            string unmatchedSyncKey = this.LastSyncKey;

            #region Call Sync command on the Inbox folder.
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            this.Sync(syncRequest);
            #endregion

            #region Call method GetItemEstimate with a SyncKey that does not match the expected value.
            Request.Options option = new Request.Options
            {
                Items = new object[] { "Email" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            GetItemEstimateRequest getItemEstimateRequest = TestSuiteBase.CreateGetItemEstimateRequest(unmatchedSyncKey, this.User1Information.InboxCollectionId, new Request.Options[] { option });
            GetItemEstimateResponse getItemEstimateResponse = CMDAdapter.GetItemEstimate(getItemEstimateRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4594");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4594
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                4594,
                @"[In SyncKey(GetItemEstimate)] The server MUST provide a Status element (section 2.2.3.162.6) value of 4 if the airsync:SyncKey value provided in the GetItemEstimate request does not match those expected within the next Sync command request (section 2.2.2.19).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetItemEstimate command, if the ConversationMode element for collections that do not store e-mail results in an invalid XML error, server should return a status value of 103.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC13_GetItemEstimate_ConversationMode_Status103()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call GetItemEstimate with the ConversationMode element for collections that do not store e-mail.
            GetItemEstimateRequest getItemEstimateRequest = TestSuiteBase.CreateGetItemEstimateRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, null, true);
            //getItemEstimateRequest.RequestData.Collections[0].ConversationMode = true;
            //getItemEstimateRequest.RequestData.Collections[0].ConversationModeSpecified = true;
            GetItemEstimateResponse getItemEstimateResponse = this.CMDAdapter.GetItemEstimate(getItemEstimateRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2100");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2100
            Site.CaptureRequirementIfAreEqual<int>(
                103,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                2100,
                @"[In ConversationMode (GetItemEstimate)] Specifying the airsync:ConversationMode element for collections that do not store email results in an invalid XML error, Status element (section 2.2.3.162.6) value 103.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4129");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4129
            Site.CaptureRequirementIfAreEqual<int>(
                103,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                4129,
                @"[In Status(GetItemEstimate)] If the entire request command fails, the Status element is present as a child of the GetItemEstimate element and contains a value that indicates the type of global failure.");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify only SMS messages and email messages can be synchronized at the same time.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC14_GetItemEstimate_WithCombinationClasses()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId));

            #region Call GetItemEstimate to get the number of both Email and SMS messages.
            Request.Options option1 = new Request.Options
            {
                Items = new object[] { "Email" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            Request.Options option2 = new Request.Options
            {
                Items = new object[] { "SMS" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            GetItemEstimateRequest getItemEstimateRequest = TestSuiteBase.CreateGetItemEstimateRequest(this.LastSyncKey, this.User1Information.InboxCollectionId, new Request.Options[] { option1, option2 });
            GetItemEstimateResponse getItemEstimateResponse = CMDAdapter.GetItemEstimate(getItemEstimateRequest);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData.Response, "The response of GetItemEstimate command should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R907");

            // Verify MS-ASCMD requirement: MS-ASCMD_R907
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(getItemEstimateResponse.ResponseData.Response[0].Status),
                907,
                @"[In Class(GetItemEstimate)] Only SMS messages and email messages can be synchronized at the same time.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the limit value of Collection element of GetItemEstimate command.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC15_GetItemEstimate_Collection_LimitValue()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send an email from user1 to user2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);

            this.SwitchUser(this.User2Information);
            this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region Create 1001 objects of GetItemEstimateCollection type for GetItemEstimate command.
            Request.Options option = new Request.Options
            {
                Items = new object[] { "Email" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            List<Request.GetItemEstimateCollection> collections = new List<Request.GetItemEstimateCollection>();
            for (int i = 0; i <= 1000; i++)
            {
                Request.GetItemEstimateCollection collection = new Request.GetItemEstimateCollection
                {
                    ItemsElementName = new Request.ItemsChoiceType10[]
                    {
                        Request.ItemsChoiceType10.SyncKey,
                        Request.ItemsChoiceType10.CollectionId,                       
                        
                    },
                    Items = new object[]
                    {
                        this.LastSyncKey,
                        this.User2Information.InboxCollectionId,                     
                       
                    },
                    Options = new Request.Options[] { option } 
                };

                collections.Add(collection);
            }

            GetItemEstimateRequest getItemEstimateRequest = Common.CreateGetItemEstimateRequest(collections.ToArray());
            GetItemEstimateResponse getItemEstimateResponse = this.CMDAdapter.GetItemEstimate(getItemEstimateRequest);
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData, "The response data of GetItemEstimate command should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5647");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5647
            Site.CaptureRequirementIfAreEqual<int>(
                103,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                5647,
                @"[In Limiting Size of Command Requests] In GetItemEstimate (section 2.2.2.7) command request, when the limit value of Collection element is bigger than 1000 (minimum 30, maximum 5000), the error returned by server is Status element (section 2.2.3.162.6) value of 103.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the combination of classes other than SMS messages or email messages synchronization at the same time cause status value 4 returned.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S07_TC16_GetItemEstimate_WithStatusValue4Returned()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId));

            #region Call GetItemEstimate to get the number of both Email and SMS messages.
            Request.Options option1 = new Request.Options
            {
                Items = new object[] { "Tasks" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            Request.Options option2 = new Request.Options
            {
                Items = new object[] { "Contacts" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            GetItemEstimateRequest getItemEstimateRequest = TestSuiteBase.CreateGetItemEstimateRequest(this.LastSyncKey, this.User1Information.InboxCollectionId, new Request.Options[] { option1, option2 });
            GetItemEstimateResponse getItemEstimateResponse = CMDAdapter.GetItemEstimate(getItemEstimateRequest);
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R908");

            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                908,
                @"[In Class(GetItemEstimate)] A request for any other combination of classes will fail with a status value of 4.");
            #endregion
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Get estimate number of items in specified time window depends on FilterType in a collection.
        /// </summary>
        /// <param name="collectionId">Specifies collection id of the folder.</param>
        /// <param name="filterType">Specifies time window for the objects sent from the server to the client.</param>
        /// <returns>GetItemEstimate command response.</returns>
        private GetItemEstimateResponse GetItemEstimateWithFilterType(string collectionId, byte filterType)
        {
            Request.Options option = new Request.Options
            {
                Items = new object[] { filterType },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.FilterType }
            };

            return this.CMDAdapter.GetItemEstimate(TestSuiteBase.CreateGetItemEstimateRequest(this.LastSyncKey, collectionId, new Request.Options[] { option }));
        }
        #endregion
    }
}