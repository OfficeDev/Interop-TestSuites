namespace Microsoft.Protocols.TestSuites.MS_ASTASK
{
    using System;
    using System.Collections.Generic;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;
    using SyncItem = Microsoft.Protocols.TestSuites.Common.DataStructures.Sync;
    using SyncStore = Microsoft.Protocols.TestSuites.Common.DataStructures.SyncStore;

    /// <summary>
    /// This scenario is to test Task class element on the server by using Sync command.
    /// </summary>
    [TestClass]
    public class S01_SyncCommand : TestSuiteBase
    {
        #region Test Class initialize and clean up

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

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence whose Type element is 'recurs daily'.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC01_CreateTaskItemRecursDaily()
        {
            #region Call Sync command to create a task item with Recurrence whose Type element is recurs daily.
            SyncStore initializeSyncResponse = this.TASKAdapter.Sync(Common.CreateInitialSyncRequest(this.UserInformation.TasksCollectionId));

            string subject = Common.GenerateResourceName(Site, "subject");
            string clientId = System.Guid.NewGuid().ToString();
            DateTime startTime = DateTime.Now;
            DateTime utcStartTime = startTime.ToUniversalTime();
            DateTime until = startTime.AddDays(10);

            // Create a task Item with Type 0 and DayOfWeek
            string stringRequest = "<?xml version=\"1.0\" encoding=\"utf-8\"?><Sync xmlns=\"AirSync\"><Collections><Collection><SyncKey>" + initializeSyncResponse.SyncKey + "</SyncKey><CollectionId>" + this.UserInformation.TasksCollectionId + "</CollectionId><DeletesAsMoves>0</DeletesAsMoves><GetChanges>1</GetChanges><WindowSize>512</WindowSize><Options><BodyPreference xmlns=\"AirSyncBase\"><Type>2</Type></BodyPreference></Options><Commands><Add><ClientId>" + clientId + "</ClientId><ApplicationData><Body xmlns=\"AirSyncBase\"><Type>1</Type><Data>Content of the body.</Data></Body><UtcStartDate xmlns=\"Tasks\">" + utcStartTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</UtcStartDate><StartDate xmlns=\"Tasks\">" + startTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</StartDate><UtcDueDate xmlns=\"Tasks\">" + utcStartTime.AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</UtcDueDate><DueDate xmlns=\"Tasks\">" + startTime.AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</DueDate><ReminderSet xmlns=\"Tasks\">1</ReminderSet><ReminderTime xmlns=\"Tasks\">" + startTime.AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</ReminderTime><Subject xmlns=\"Tasks\">" + subject + "</Subject><Importance xmlns=\"Tasks\">0</Importance><Categories xmlns=\"Tasks\"><Category xmlns=\"Tasks\">Business</Category><Category xmlns=\"Tasks\">Waiting</Category></Categories><Recurrence xmlns=\"Tasks\"><Type>0</Type><Start>" + DateTime.Now.AddHours(1).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</Start><Until>" + until.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</Until><Interval>1</Interval><DayOfWeek>2</DayOfWeek></Recurrence></ApplicationData></Add></Commands></Collection></Collections></Sync>";
            SendStringResponse sendStringResponse = this.TASKAdapter.SendStringRequest(stringRequest, CommandName.Sync);

            SyncStore response;
            if (Common.IsRequirementEnabled(631, Site))
            {
                response = this.ExtractSyncStore(sendStringResponse);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R631");

                // Verify MS-ASTASK requirement: MS-ASTASK_R631
                Site.CaptureRequirementIfAreEqual<int>(
                    6,
                    int.Parse(response.AddResponses[0].Status),
                    631,
                    @"[In Appendix A: Product Behavior]  If the Type element value is 0 (zero), the DayOfWeek element is not a required child element of the Recurrence element. (<1> Section 2.2.2.10:  When the Type element value is 0, Exchange 2007 SP1 responds with a status 6 if DayOfWeek is set in the request.)");

                int dayOfWeekIndex = stringRequest.IndexOf("<DayOfWeek>");
                int dayOfWeekEndIndex = stringRequest.IndexOf("</DayOfWeek>") + 11;
                stringRequest = stringRequest.Remove(dayOfWeekIndex, dayOfWeekEndIndex - dayOfWeekIndex + 1);
                sendStringResponse = this.TASKAdapter.SendStringRequest(stringRequest, CommandName.Sync);
            }

            // Extract status code from string response
            response = this.ExtractSyncStore(sendStringResponse);

            Site.Assert.AreEqual<int>(1, int.Parse(response.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);
            #endregion

            #region Call Sync command to get the task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R158");

            // Verify MS-ASTASK requirement: MS-ASTASK_R158
            // The DeadOccure is not set in the request and the DeadOccur in the response value is 0, this requirement can be covered.
            Site.CaptureRequirementIfAreEqual<byte?>(
                0,
                syncedTaskItem.Task.Recurrence.DeadOccur,
                158,
                @"[In DeadOccur]The default value of the DeadOccur element is 0 (zero).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R264");

            // Verify MS-ASTASK requirement: MS-ASTASK_R264
            Site.CaptureRequirementIfAreEqual<byte?>(
                0,
                syncedTaskItem.Task.Sensitivity,
                264,
                @"[In Sensitivity] The default value of the Sensitivity element is 0 (zero) (normal).");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence whose Type element is 'recurs daily', but without Start element.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC02_CreateTaskItemRecursDailyWithoutStart()
        {
            #region Call Sync command to create task item without start time

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);

            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 0,
                OccurrencesSpecified = true,
                Occurrences = 2,
                DayOfWeekSpecified = true,
                DayOfWeek = 1
            };
            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);

            SyncStore syncResponse = this.SyncAddTask(taskItem);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R37011");

            // Verify MS-ASTASK requirement: MS-ASTASK_R37011
            // If the Start is not specified in the request, and the Status in the response is 6, this requirement can be covered.
            Site.CaptureRequirementIfAreEqual<int>(
                6,
                int.Parse(syncResponse.AddResponses[0].Status),
                37011,
                @"[In Start Element] If a client does not include the Start element, as specified in section 2.2.2.23, in a Sync command request ([MS-ASCMD] section 2.2.2.19) whenever a Recurrence element is present, then the server MUST respond with status error 6.");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence whose Type element is not 2 or 5 and contains a 'DayOfMonth' element.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC03_CreateTaskItemFailWithDayOfMonth()
        {
            #region Call Sync command to create task item with Type element set to 0 and DayOfMonth element set

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);
            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 0,
                Start = DateTime.Now,
                OccurrencesSpecified = true,
                Occurrences = 2,
                DayOfWeekSpecified = true,
                DayOfWeek = 1,
                DayOfMonthSpecified = true,
                DayOfMonth = 10
            };

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(6, int.Parse(syncResponse.AddResponses[0].Status), "Status code 6 should be returned when Type element set to 0 and DayOfMonth element set");

            #endregion

            #region Call Sync command to create task item with Type element set to 1 and DayOfMonth element set

            taskItem = new Dictionary<Request.ItemsChoiceType8, object> { { Request.ItemsChoiceType8.Subject2, subject } };
            recurrence = new Request.Recurrence1
            {
                Type = 1,
                Start = DateTime.Now,
                DayOfWeekSpecified = true,
                DayOfWeek = 1,
                DayOfMonthSpecified = true,
                DayOfMonth = 10
            };

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(6, int.Parse(syncResponse.AddResponses[0].Status), "Status code 6 should be returned when Type element set to 1 and DayOfMonth element set");

            #endregion

            #region Call Sync command to create task item with Type element set to 3 and DayOfMonth element set

            taskItem = new Dictionary<Request.ItemsChoiceType8, object> { { Request.ItemsChoiceType8.Subject2, subject } };
            recurrence = new Request.Recurrence1
            {
                Type = 3,
                Start = DateTime.Now,
                DayOfWeekSpecified = true,
                DayOfWeek = 1,
                WeekOfMonthSpecified = true,
                WeekOfMonth = 2,
                DayOfMonthSpecified = true,
                DayOfMonth = 10
            };

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(6, int.Parse(syncResponse.AddResponses[0].Status), "Status code 6 should be returned when Type element set to 3 and DayOfMonth element set");

            #endregion

            #region Call Sync command to create task item with Type element set to 6 and DayOfMonth element set

            taskItem = new Dictionary<Request.ItemsChoiceType8, object> { { Request.ItemsChoiceType8.Subject2, subject } };
            recurrence = new Request.Recurrence1
            {
                Type = 6,
                Start = DateTime.Now,
                DayOfWeekSpecified = true,
                DayOfWeek = 1,
                WeekOfMonthSpecified = true,
                WeekOfMonth = 2,
                MonthOfYearSpecified = true,
                MonthOfYear = 2,
                DayOfMonthSpecified = true,
                DayOfMonth = 10
            };

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(6, int.Parse(syncResponse.AddResponses[0].Status), "Status code 6 should be returned when Type element set to 6 and DayOfMonth element set");

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R132");

            // Verify MS-ASTASK requirement: MS-ASTASK_R132
            // If the Type element value is in {0,1,3,6}, and the server responds with a status 6 error (conversion error), this requirement can be covered.
            Site.CaptureRequirement(
                132,
                @"[In DayofMonth]When a request is issued with the DayOfMonth element in other instances[when the Type element value is not  2 or 5], the server responds with a status 6 error (conversion error).");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence of which the Type element is 'recurs monthly'.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC04_CreateTaskItemRecursMonthly()
        {
            #region Call Sync command to create a task with Recurrence whose Type element is recurs monthly.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);
            taskItem.Add(Request.ItemsChoiceType8.Importance1, (byte)1);
            Request.Categories3 categories = new Request.Categories3 { Category = "Business,Waiting".Split(',') };
            taskItem.Add(Request.ItemsChoiceType8.Categories3, categories);

            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 2,
                Start = DateTime.Now,
                UntilSpecified = true,
                DayOfMonthSpecified = true,
                DayOfMonth = 10
            };
            recurrence.Until = recurrence.Start.AddMonths(3);

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(1, int.Parse(syncResponse.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Sync command to get the task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R404");

            // Verify MS-ASTASK requirement: MS-ASTASK_R404
            // If DayOfMonthSpecified is true, DayOfMonth element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.DayOfMonthSpecified,
                404,
                @"[In DayOfMonth] A command [request or] response has a minimum of one DayOfMonth element per Recurrence element if the value of the Type element (section 2.2.2.27) is 2 [or 5].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R409");

            // Verify MS-ASTASK requirement: MS-ASTASK_R409
            // If DayOfMonthSpecified is true, DayOfMonth element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.DayOfMonthSpecified,
                409,
                @"[In DayOfMonth] The DayOfMonth element MUST only be included in [requests or] responses when the Type element value is 2 [or 5].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R366");

            // Verify MS-ASTASK requirement: MS-ASTASK_R366
            // If response is returned from server successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(syncResponse.AddResponses[0].Status),
                366,
                @"[In Sync Command Response] When a client uses the Sync command request ([MS-ASCMD] section 2.2.2.19) to synchronize its Task class items for a specified user with the tasks currently stored by the server, as specified in section 3.1.5.3, the server responds with a Sync command response ([MS-ASCMD] section 2.2.2.19).");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence of which the Type element is 'recurs yearly'.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC05_CreateTaskItemRecursYearly()
        {
            #region Call Sync command to create a task item with Recurrence whose Type element is recurs yearly.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);
            taskItem.Add(Request.ItemsChoiceType8.Importance1, (byte)2);
            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 5,
                Start = DateTime.Now,
                UntilSpecified = true,
                DayOfMonthSpecified = true,
                DayOfMonth = 10,
                MonthOfYearSpecified = true,
                MonthOfYear = 2
            };
            recurrence.Until = recurrence.Start.AddYears(3);

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(1, int.Parse(syncResponse.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Sync command to get the task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R406");

            // Verify MS-ASTASK requirement: MS-ASTASK_R406
            // If DayOfMonthSpecified is true, DayOfMonth element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.DayOfMonthSpecified,
                406,
                @"[In DayOfMonth] A command [request or] response has a minimum of one DayOfMonth element per Recurrence element if the value of the Type element (section 2.2.2.27) is [2 or] 5.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R410");

            // Verify MS-ASTASK requirement: MS-ASTASK_R410
            // If DayOfMonthSpecified is true, DayOfMonth element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.DayOfMonthSpecified,
                410,
                @"[In DayOfMonth] The DayOfMonth element MUST only be included in [requests or] responses when the Type element value is [2 or] 5.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R417");

            // Verify MS-ASTASK requirement: MS-ASTASK_R417
            // If MonthOfYearSpecified is true, MonthOfYear element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.MonthOfYearSpecified,
                417,
                @"[In MonthOfYear] A command [request or] response has a minimum of one MonthofYear child element per Recurrence element if the value of the Type element (section 2.2.2.27) is [either] 5 [or 6].");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence whose Type element is "recurs weekly".
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC06_CreateTaskItemRecursWeekly()
        {
            #region Call Sync command to create a task item with Recurrence whose Type element is recurs weekly.
            SyncStore initializeSyncResponse = this.TASKAdapter.Sync(Common.CreateInitialSyncRequest(this.UserInformation.TasksCollectionId));

            string subject = Common.GenerateResourceName(Site, "subject");
            string clientId = System.Guid.NewGuid().ToString();
            DateTime startTime = DateTime.Now;
            DateTime utcStartTime = startTime.ToUniversalTime();
            DateTime until = startTime.AddDays(21);

            // Create a task Item with Type 1
            string stringRequest = "<?xml version=\"1.0\" encoding=\"utf-8\"?><Sync xmlns=\"AirSync\"><Collections><Collection><SyncKey>" + initializeSyncResponse.SyncKey + "</SyncKey><CollectionId>" + this.UserInformation.TasksCollectionId + "</CollectionId><DeletesAsMoves>0</DeletesAsMoves><GetChanges>1</GetChanges><WindowSize>512</WindowSize><Options><BodyPreference xmlns=\"AirSyncBase\"><Type>2</Type></BodyPreference></Options><Commands><Add><ClientId>" + clientId + "</ClientId><ApplicationData><Body xmlns=\"AirSyncBase\"><Type>1</Type><Data>Content of the body.</Data></Body><UtcStartDate xmlns=\"Tasks\">" + utcStartTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</UtcStartDate><StartDate xmlns=\"Tasks\">" + startTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</StartDate><UtcDueDate xmlns=\"Tasks\">" + utcStartTime.AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</UtcDueDate><DueDate xmlns=\"Tasks\">" + startTime.AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</DueDate><ReminderSet xmlns=\"Tasks\">1</ReminderSet><ReminderTime xmlns=\"Tasks\">" + startTime.AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</ReminderTime><Subject xmlns=\"Tasks\">" + subject + "</Subject><Importance xmlns=\"Tasks\">0</Importance><Categories xmlns=\"Tasks\"><Category xmlns=\"Tasks\">Business</Category><Category xmlns=\"Tasks\">Waiting</Category></Categories><Recurrence xmlns=\"Tasks\"><Type>1</Type><Start>" + DateTime.Now.AddHours(1).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</Start><Until>" + until.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</Until><Interval>1</Interval><DayOfWeek>2</DayOfWeek></Recurrence></ApplicationData></Add></Commands></Collection></Collections></Sync>";
            SendStringResponse sendStringResponse = this.TASKAdapter.SendStringRequest(stringRequest, CommandName.Sync);

            // Extract status code from string response
            SyncStore response = this.ExtractSyncStore(sendStringResponse);
            Site.Assert.AreEqual<int>(1, int.Parse(response.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Sync command to get the task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R135");

            // Verify MS-ASTASK requirement: MS-ASTASK_R135
            // If DayOfWeekSpecified is true, DayOfWeek element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.DayOfWeekSpecified,
                135,
                @"[In DayOfWeek] A command [request or] response has a minimum of one DayOfWeek element per Recurrence element when the Type element value is 1[, 3, or 6].");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with a Recurrence whose Type element is not 0, 1, 3 and 6 and contains a 'DayOfWeek' element.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC07_CreateTaskItemFailWithDayOfWeek()
        {
            #region Call Sync command to create task item with Type element set to 2 and DayOfWeek element set

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);
            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 2,
                Start = DateTime.Now,
                OccurrencesSpecified = true,
                Occurrences = 2,
                DayOfWeekSpecified = true,
                DayOfWeek = 1,
                DayOfMonthSpecified = true,
                DayOfMonth = 10
            };
            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R532");

            Site.CaptureRequirementIfAreEqual<int>(
                6,
                int.Parse(syncResponse.AddResponses[0].Status),
                532,
                "[In DayOfWeek] If a request is issued with the DayOfWeek element when the Type element value is 2 [or 5], the server responds with a status 6 error (conversion error). ");

            #endregion

            #region Call Sync command to create task item with Type element set to 5 and DayOfWeek element set

            taskItem = new Dictionary<Request.ItemsChoiceType8, object> { { Request.ItemsChoiceType8.Subject2, subject } };
            recurrence = new Request.Recurrence1
            {
                Type = 5,
                Start = DateTime.Now,
                DayOfMonthSpecified = true,
                DayOfMonth = 10,
                MonthOfYearSpecified = true,
                MonthOfYear = 5,
                DayOfWeekSpecified = true,
                DayOfWeek = 1
            };
            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            syncResponse = this.SyncAddTask(taskItem);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R533");

            Site.CaptureRequirementIfAreEqual<int>(
                6,
                int.Parse(syncResponse.AddResponses[0].Status),
                533,
                "[In DayOfWeek] If a request is issued with the DayOfWeek element when the Type element value is [2 or] 5, the server responds with a status 6 error (conversion error). ");

            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence whose Type element is 'recurs weekly' and contains an invalid 'FirstDayOfWeek' value.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC08_CreateTaskItemRecursWeeklyWithInvalidFirstDayOfWeek()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The FirstDayOfWeek element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The FirstDayOfWeek element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command to create task item with Type element set to 1 and an invalid FirstDayOfWeek value.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);
            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 1,
                Start = DateTime.Now,
                OccurrencesSpecified = true,
                Occurrences = 2,
                DayOfWeekSpecified = true,
                DayOfWeek = 1,
                FirstDayOfWeekSpecified = true,
                FirstDayOfWeek = 7
            };

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(6, int.Parse(syncResponse.AddResponses[0].Status), "Status code 6 should be returned when an invalid FirstDayOfWeek element value is set.");

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R171");

            // Verify MS-ASTASK requirement: MS-ASTASK_R171
            // If the FirstDayOfWeek element value is not between 0 and 6, and the server responds with a status 6 error (conversion error),this requirement can be covered.
            Site.CaptureRequirement(
                171,
                @"[In FirstDayOfWeek] If the client uses the Sync command request ([MS-ASCMD] section 2.2.2.19) to transmit a value not included in this table[the value is between 0 to 6], then the server MUST return protocol status error 6.");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with an invalid 'Importance' value.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC09_CreateTaskItemWithInvalidImportanceElement()
        {
            #region Call Sync command to create task item with an invalid Importance value.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);

            // The valid value for importance is 0,1,2
            taskItem.Add(Request.ItemsChoiceType8.Importance1, (byte)3);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(1, int.Parse(syncResponse.AddResponses[0].Status), "Task item should be created successfully.");
            ItemsNeedToDelete.Add(subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R187");

            // Verify MS-ASTASK requirement: MS-ASTASK_R187
            // If the Importance in the request is not in {0,1,2} and the response Status is 1(Success),the requirement can be covered.
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(syncResponse.AddResponses[0].Status),
                187,
                @"[In Importance] If the Importance element is set to a value other than 0 (zero), 1, or 2 in a command request, the server will process the request successfully (that is, will not return an error code in the response) [and return the same value that is set in the request].");

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R550");

            // Verify MS-ASTASK requirement: MS-ASTASK_R550
            Site.CaptureRequirementIfAreEqual<byte?>(
                3,
                syncedTaskItem.Task.Importance,
                550,
                @"[In Importance] If the Importance element is set to a value other than 0 (zero), 1, or 2 in a command request, the server will [process the request successfully (that is, will not return an error code in the response) and] return the same value that is set in the request.");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence whose Type element is 'recurs yearly on the Nth day'.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC10_CreateTaskItemRecursYearlyOnTheNthDay()
        {
            #region Call Sync command to create task item which recurs yearly on the nth day.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);
            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 6,
                Start = DateTime.Now,
                OccurrencesSpecified = true,
                Occurrences = 4,
                DayOfWeekSpecified = true,
                DayOfWeek = 3,
                WeekOfMonthSpecified = true,
                WeekOfMonth = 2,
                MonthOfYearSpecified = true,
                MonthOfYear = 10
            };
            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(1, int.Parse(syncResponse.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Sync command to get the task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R531");

            // Verify MS-ASTASK requirement: MS-ASTASK_R531
            // If DayOfWeekSpecified is true, DayOfWeek element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.DayOfWeekSpecified,
                531,
                @"[In DayOfWeek] A command [request or] response has a minimum of one DayOfWeek element per Recurrence element when the Type element value is [1, 3, or] 6.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R418");

            // Verify MS-ASTASK requirement: MS-ASTASK_R418
            // If MonthOfYearSpecified is true, MonthOfYear element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.MonthOfYearSpecified,
                418,
                @"[In MonthOfYear] A command [request or] response has a minimum of one MonthofYear child element per Recurrence element if the value of the Type element (section 2.2.2.27) is [either 5 or] 6.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R431");

            // Verify MS-ASTASK requirement: MS-ASTASK_R431
            // If WeekOfMonthSpecified is true, WeekOfMonth element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.WeekOfMonthSpecified,
                431,
                @"[In WeekOfMonth] A command [request or] response has a minimum of one WeekOfMonth child element per Recurrence element when the value of the Type element (section 2.2.2.27) is [either 3 or] 6.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R433");

            // Verify MS-ASTASK requirement: MS-ASTASK_R433
            // If WeekOfMonthSpecified is true, WeekOfMonth element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.WeekOfMonthSpecified,
                433,
                @"[In WeekOfMonth] The WeekOfMonth element MUST only be included in [requests or] responses when the Type element value is [either 3 or] 6.");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with an invalid 'ReminderSet' value.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC11_CreateTaskItemWithInvalidReminderSetElement()
        {
            #region Call Sync command to create task item with an invalid ReminderSet value.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");
            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);

            // Set an invalid ReminderSet value.
            taskItem.Add(Request.ItemsChoiceType8.ReminderSet, (byte)2);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(6, int.Parse(syncResponse.AddResponses[0].Status), "A status code 6 should be returned when an invalid ReminderSet value is set.");

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R252");

            // Verify MS-ASTASK requirement: MS-ASTASK_R252
            // If the ReminderSet contains a value other than 0 (zero) or 1 in a command request, and the Status in the command response is 6,this requirement can be covered.
            Site.CaptureRequirementIfAreEqual<int>(
                6,
                int.Parse(syncResponse.AddResponses[0].Status),
                252,
                @"[In  ReminderSet]If the ReminderSet element contains a value other than 0 (zero) or 1 in a command request, the server responds with a status value of 6 in the command response.");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing task with Recurrence whose Type element is 'recurs monthly on the Nth day'.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC12_CreateTaskItemRecursMonthlyOnTheNthDay()
        {
            #region Call Sync command to create task item which recurs monthly on the nth day.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);

            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 3,
                Start = DateTime.Now,
                UntilSpecified = true,
                DayOfWeekSpecified = true,
                DayOfWeek = 1,
                WeekOfMonthSpecified = true,
                WeekOfMonth = 2
            };
            recurrence.Until = recurrence.Start.AddMonths(3);

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);

            SyncStore syncResponse = this.SyncAddTask(taskItem);
            Site.Assert.AreEqual<int>(1, int.Parse(syncResponse.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);
            #endregion

            #region Call Sync command to get the task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R530");

            // Verify MS-ASTASK requirement: MS-ASTASK_R530
            // If DayOfWeekSpecified is true, DayOfWeek element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.DayOfWeekSpecified,
                530,
                @"[In DayOfWeek] A command [request or] response has a minimum of one DayOfWeek element per Recurrence element when the Type element value is [1,] 3[, or 6].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R430");

            // Verify MS-ASTASK requirement: MS-ASTASK_R430
            // If WeekOfMonthSpecified is true, WeekOfMonth element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.WeekOfMonthSpecified,
                430,
                @"[In WeekOfMonth] A command [request or] response has a minimum of one WeekOfMonth child element per Recurrence element when the value of the Type element (section 2.2.2.27) is [either] 3 [or 6].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R319");

            // Verify MS-ASTASK requirement: MS-ASTASK_R319
            // If WeekOfMonthSpecified is true, WeekOfMonth element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.WeekOfMonthSpecified,
                319,
                @"[In WeekOfMonth] The WeekOfMonth element MUST only be included in [requests or] responses when the Type element value is [either] 3 [or 6].");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence whose Type value is not 3 or 6 and contains a 'WeekOfMonth' element.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC13_CreateTaskItemFailWithWeekOfMonth()
        {
            #region Call Sync command to create task item with Type element set to 0 and WeekOfMonth set.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);
            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 0,
                Start = DateTime.Now,
                OccurrencesSpecified = true,
                Occurrences = 2,
                DayOfWeekSpecified = true,
                DayOfWeek = 1,
                WeekOfMonthSpecified = true,
                WeekOfMonth = 2
            };

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(6, int.Parse(syncResponse.AddResponses[0].Status), "Status code 6 should be returned when Type element value is 0 and WeekOfMonth is set.");

            #endregion

            #region Call Sync command to create task item with Type element set to 1 and WeekOfMonth set.

            taskItem = new Dictionary<Request.ItemsChoiceType8, object> { { Request.ItemsChoiceType8.Subject2, subject } };
            recurrence = new Request.Recurrence1
            {
                Type = 1,
                Start = DateTime.Now,
                DayOfWeekSpecified = true,
                DayOfWeek = 1,
                WeekOfMonthSpecified = true,
                WeekOfMonth = 2
            };

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(6, int.Parse(syncResponse.AddResponses[0].Status), "Status code 6 should be returned when Type element value is 1 and WeekOfMonth is set.");

            #endregion

            #region Call Sync command to create task item with Type element set to 2 and WeekOfMonth set.

            taskItem = new Dictionary<Request.ItemsChoiceType8, object> { { Request.ItemsChoiceType8.Subject2, subject } };
            recurrence = new Request.Recurrence1
            {
                Type = 2,
                Start = DateTime.Now,
                DayOfMonthSpecified = true,
                DayOfMonth = 10,
                WeekOfMonthSpecified = true,
                WeekOfMonth = 2
            };

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(6, int.Parse(syncResponse.AddResponses[0].Status), "Status code 6 should be returned when Type element value is 2 and WeekOfMonth is set.");

            #endregion

            #region Call Sync command to create task item with Type element set to 5 and WeekOfMonth set.

            taskItem = new Dictionary<Request.ItemsChoiceType8, object> { { Request.ItemsChoiceType8.Subject2, subject } };
            recurrence = new Request.Recurrence1
            {
                Type = 5,
                Start = DateTime.Now,
                DayOfMonthSpecified = true,
                DayOfMonth = 10,
                MonthOfYearSpecified = true,
                MonthOfYear = 2,
                WeekOfMonthSpecified = true,
                WeekOfMonth = 2
            };

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(6, int.Parse(syncResponse.AddResponses[0].Status), "Status code 6 should be returned when Type element value is 5 and WeekOfMonth is set.");

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R320");

            // Verify MS-ASTASK requirement: MS-ASTASK_R320
            // If the Type element value is in {0,1,2,5}, and the server responds with a status 6 error (conversion error),this requirement can be covered.
            Site.CaptureRequirement(
                320,
                @"[In WeekOfMonth] When a client's request is issued with the WeekOfMonth element in other instances[when the Type element value is not either 3 or 6], the server responds with a status 6 error (conversion error).");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence whose type value is not 2, 3, 5 or 6 and contains a 'CalendarType' element.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC14_CreateTaskItemFailWithCalendarType()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The FirstDayOfWeek element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The FirstDayOfWeek element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command to create task item with Type element set to 0 and CalendarType set.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);
            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 0,
                Start = DateTime.Now,
                OccurrencesSpecified = true,
                Occurrences = 2,
                DayOfWeekSpecified = true,
                DayOfWeek = 1,
                CalendarType = 0,
                CalendarTypeSpecified = true
            };

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(6, int.Parse(syncResponse.AddResponses[0].Status), "Status code 6 should be returned when Type element value is 0 and CalendarType is set.");

            #endregion

            #region Call Sync command to create task item with Type element set to 1 and CalendarType set.

            taskItem = new Dictionary<Request.ItemsChoiceType8, object> { { Request.ItemsChoiceType8.Subject2, subject } };
            recurrence = new Request.Recurrence1
            {
                Type = 1,
                Start = DateTime.Now,
                DayOfWeekSpecified = true,
                DayOfWeek = 1,
                CalendarType = 0,
                CalendarTypeSpecified = true
            };

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(6, int.Parse(syncResponse.AddResponses[0].Status), "Status code 6 should be returned when Type element value is 1 and CalendarType is set.");

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R81");

            // Verify MS-ASTASK requirement: MS-ASTASK_R81
            // If the Type element value is not in {2,3,5,6} and the CalendarType element is included, the server responds with a status 6 error (conversion error), this requirement can be covered.
            Site.CaptureRequirement(
                81,
                @"[In CalendarType] If the CalendarType element is included in other instances[when the Type element is not set to a value of 2, 3, 5, or 6 ], the server responds with a status 6 error (conversion error).");
        }

        /// <summary>
        /// This test case is designed to verify the Recurrence element when Occurrence and Until elements are both set.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC15_CreateTaskItemOccurrencesAndUntilBothSet()
        {
            #region Call Sync command to create task item with both Occurrences and Until elements set
            SyncStore initializeSyncResponse = this.TASKAdapter.Sync(Common.CreateInitialSyncRequest(this.UserInformation.TasksCollectionId));

            string subject = Common.GenerateResourceName(Site, "subject");
            string clientId = System.Guid.NewGuid().ToString();
            DateTime startTime = DateTime.Now;
            DateTime utcStartTime = startTime.ToUniversalTime();
            DateTime until = startTime.AddDays(5);

            // Create a task Item with Type 0
            string stringRequest = "<?xml version=\"1.0\" encoding=\"utf-8\"?><Sync xmlns=\"AirSync\"><Collections><Collection><SyncKey>" + initializeSyncResponse.SyncKey + "</SyncKey><CollectionId>" + this.UserInformation.TasksCollectionId + "</CollectionId><DeletesAsMoves>0</DeletesAsMoves><GetChanges>1</GetChanges><WindowSize>512</WindowSize><Options><BodyPreference xmlns=\"AirSyncBase\"><Type>2</Type></BodyPreference></Options><Commands><Add><ClientId>" + clientId + "</ClientId><ApplicationData><Body xmlns=\"AirSyncBase\"><Type>1</Type><Data>Content of the body.</Data></Body><UtcStartDate xmlns=\"Tasks\">" + utcStartTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</UtcStartDate><StartDate xmlns=\"Tasks\">" + startTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</StartDate><UtcDueDate xmlns=\"Tasks\">" + utcStartTime.AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</UtcDueDate><DueDate xmlns=\"Tasks\">" + startTime.AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</DueDate><ReminderSet xmlns=\"Tasks\">1</ReminderSet><ReminderTime xmlns=\"Tasks\">" + startTime.AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</ReminderTime><Subject xmlns=\"Tasks\">" + subject + "</Subject><Categories xmlns=\"Tasks\"><Category xmlns=\"Tasks\">Business</Category><Category xmlns=\"Tasks\">Waiting</Category></Categories><Recurrence xmlns=\"Tasks\"><Type>0</Type><Start>" + DateTime.Now.AddHours(1).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</Start><Until>" + until.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</Until><Occurrences>2</Occurrences><Interval>1</Interval><DayOfWeek>2</DayOfWeek></Recurrence></ApplicationData></Add></Commands></Collection></Collections></Sync>";
            SendStringResponse sendStringResponse = this.TASKAdapter.SendStringRequest(stringRequest, CommandName.Sync);

            // Extract status code from string response
            SyncStore response = this.ExtractSyncStore(sendStringResponse);

            Site.Assert.AreEqual<int>(1, int.Parse(response.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Sync command to get the task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R37013");

            // Verify MS-ASTASK requirement: MS-ASTASK_R37013
            // If OccurrencesSpecified is true, UntilSpecified is false the server respect the value of Occurrences.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.OccurrencesSpecified == true && syncedTaskItem.Task.Recurrence.UntilSpecified == false,
                37013,
                @"[In Occurrences and Until Elements] If both the Occurrences element, as specified in section 2.2.2.16, and the Until element, as specified in section 2.2.2.28, are included in a Sync command request ([MS-ASCMD] section 2.2.2.19), the server MUST respect the value of the Occurrences element and ignore the Until element.");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence without Type element.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC16_CreateTaskItemRecurrenceWithoutType()
        {
            #region Call Sync command to create a task item without Type element contained in Recurrence.

            SyncStore initializeSyncResponse = this.TASKAdapter.Sync(Common.CreateInitialSyncRequest(this.UserInformation.TasksCollectionId));

            string subject = Common.GenerateResourceName(Site, "subject");
            string clientId = System.Guid.NewGuid().ToString();
            DateTime startTime = DateTime.Now.AddHours(1).AddDays(1);
            DateTime utcStartTime = startTime.ToUniversalTime();

            // Send a string creating task item request without Type element contained in Recurrence.
            string stringRequest = "<?xml version=\"1.0\" encoding=\"utf-8\"?><Sync xmlns=\"AirSync\"><Collections><Collection><SyncKey>" + initializeSyncResponse.SyncKey + "</SyncKey><CollectionId>" + this.UserInformation.TasksCollectionId + "</CollectionId><DeletesAsMoves>0</DeletesAsMoves><GetChanges>1</GetChanges><WindowSize>512</WindowSize><Options><BodyPreference xmlns=\"AirSyncBase\"><Type>2</Type></BodyPreference></Options><Commands><Add><ClientId>" + clientId + "</ClientId><ApplicationData><Body xmlns=\"AirSyncBase\"><Type>1</Type><Data>Content of the body.</Data></Body><UtcStartDate xmlns=\"Tasks\">" + utcStartTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</UtcStartDate><StartDate xmlns=\"Tasks\">" + startTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</StartDate><UtcDueDate xmlns=\"Tasks\">" + utcStartTime.AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</UtcDueDate><DueDate xmlns=\"Tasks\">" + startTime.AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</DueDate><ReminderSet xmlns=\"Tasks\">1</ReminderSet><ReminderTime xmlns=\"Tasks\">" + startTime.AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</ReminderTime><Subject xmlns=\"Tasks\">" + subject + "</Subject><Recurrence xmlns=\"Tasks\"><Start>" + DateTime.Now.AddHours(1).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</Start><DayOfWeek>1</DayOfWeek></Recurrence></ApplicationData></Add></Commands></Collection></Collections></Sync>";
            SendStringResponse sendStringResponse = this.TASKAdapter.SendStringRequest(stringRequest, CommandName.Sync);

            // Extract status code from string response
            SyncStore response = new SyncStore();
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(sendStringResponse.ResponseDataXML);
            XmlNodeList nodes = doc.DocumentElement.GetElementsByTagName("Collection");

            foreach (XmlNode node in nodes)
            {
                foreach (XmlNode item in node.ChildNodes)
                {
                    if (item.Name == "Responses")
                    {
                        foreach (XmlNode add in item)
                        {
                            if (add.Name == "Add")
                            {
                                Response.SyncCollectionsCollectionResponsesAdd responseData = new Response.SyncCollectionsCollectionResponsesAdd();

                                foreach (XmlNode addItem in add)
                                {
                                    if (addItem.Name == "Status")
                                    {
                                        responseData.Status = addItem.InnerText;
                                    }
                                }

                                response.AddResponses.Add(responseData);
                            }
                        }
                    }
                }
            }

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R289");

            // Verify MS-ASTASK requirement: MS-ASTASK_R289
            Site.CaptureRequirementIfAreEqual<int>(
                6,
                int.Parse(response.AddResponses[0].Status),
                289,
                @"[In Type] If a client does not include this element[Type] in a Sync command request ([MS-ASCMD] section 2.2.2.19) whenever a Recurrence element is present, then the server MUST respond with status error 6.");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with a 'DateCompleted' element.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC17_CreateTaskItemWithDateCompletedElement()
        {
            #region Call Sync command to create task item with DateCompletedElement.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);
            taskItem.Add(Request.ItemsChoiceType8.Complete, (byte)1);
            taskItem.Add(Request.ItemsChoiceType8.DateCompleted, DateTime.Now.AddHours(2));
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(1, int.Parse(syncResponse.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Sync command to get task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R124");

            // Verify MS-ASTASK requirement: MS-ASTASK_R124
            // If the Complete is '1' and the DateCompleted is not null, this requirement can be covered.
            Site.CaptureRequirementIfIsNotNull(
                syncedTaskItem.Task.DateCompleted,
                124,
                @"[In DateCompleted] The DateCompleted element MUST be included in the response if the Complete element (section 2.2.2.5) value is 1.");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about not including ReminderSet and importance element into airsync:Change.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC18_CreateAndChangeSubject()
        {
            #region Call Sync command to create a task item.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);
            taskItem.Add(Request.ItemsChoiceType8.ReminderSet, (byte)1);
            taskItem.Add(Request.ItemsChoiceType8.Importance1, (byte)2);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(1, int.Parse(syncResponse.AddResponses[0].Status), "Task item should be created successfully.");
            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Sync command to get the change in TaskFolder.

            SyncStore changeInTaskFolder = this.SyncChanges(this.UserInformation.TasksCollectionId);

            #endregion

            SyncItem syncedTaskItem = null;
            foreach (SyncItem item in changeInTaskFolder.AddElements)
            {
                if (item.Task.Subject.Equals(subject))
                {
                    syncedTaskItem = item;
                    break;
                }
            }

            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #region Call Sync command to update task item subject without changing Importance or ReminderSet.

            string newSubject = Common.GenerateResourceName(Site, "new subject");
            Dictionary<Request.ItemsChoiceType7, object> changedPropertyValue = new Dictionary<Request.ItemsChoiceType7, object>
            {
                {
                    Request.ItemsChoiceType7.Subject2,
                    newSubject
                }
            };
            SyncStore updateTaskResponse = this.SyncChangeTask(changeInTaskFolder.SyncKey, syncedTaskItem.ServerId, changedPropertyValue);

            Site.Assert.AreEqual<int>(1, updateTaskResponse.CollectionStatus, "Task item should be updated successfully.");
            ItemsNeedToDelete.Remove(subject);
            ItemsNeedToDelete.Add(newSubject);

            #endregion

            #region Call Sync command to get the task item.

            SyncItem updatedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, newSubject);
            Site.Assert.IsNotNull(updatedTaskItem.Task, "The task which subject is {0} should exist in server.", newSubject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R438");

            // Verify MS-ASTASK requirement: MS-ASTASK_R438
            // If response is successfully returned from server, the response message would be verified by schema validation and this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                syncedTaskItem.Task,
                438,
                @"[In Sync Command Response] Top-level Task class elements, as specified in section 2.2, are returned as child elements of the airsync:ApplicationData element ([MS-ASCMD] section 2.2.3.11) within an airsync:Change element ([MS-ASCMD] section 2.2.3.24) in the Sync command response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R368");

            // Verify MS-ASTASK requirement: MS-ASTASK_R368
            // If response is successfully returned from server, the response message would be verified by schema validation and this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                changeInTaskFolder.AddElements.Count > 0,
                368,
                @"[In Sync Command Response] Top-level Task class elements, as specified in section 2.2, are returned as child elements of the airsync:ApplicationData element ([MS-ASCMD] section 2.2.3.11) within an airsync:Add element ([MS-ASCMD] section 2.2.3.7.2) in the Sync command response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R369");

            // Verify MS-ASTASK requirement: MS-ASTASK_R369
            Site.CaptureRequirementIfAreEqual<byte?>(
                2,
                syncedTaskItem.Task.Importance,
                369,
                @"[In Importance Element] If the Importance element (section 2.2.2.12) is not included as a child element of the airsync:Change element in a Sync command request, the server MUST NOT delete the element from its message store, but rather keep its value unchanged.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R370");

            // Verify MS-ASTASK requirement: MS-ASTASK_R370
            Site.CaptureRequirementIfAreEqual<byte?>(
                1,
                syncedTaskItem.Task.ReminderSet,
                370,
                @"[In ReminderSet Element] If the ReminderSet element (section 2.2.2.20) was previously set on a task but is not included as a child element of the airsync:Change element in a Sync command request, the server MUST NOT delete the element from its message store but rather keep its value unchanged.");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks Categories element that doesn't contain child elements.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC19_CreateTaskItemCategoriesWithoutChildElements()
        {
            #region Call Sync command to create task item with categories not containing any child elements.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);

            // Set Categories element without any child elements.
            Request.Categories3 categories = new Request.Categories3();
            taskItem.Add(Request.ItemsChoiceType8.Categories3, categories);

            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(1, int.Parse(syncResponse.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Sync command to get the task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R436");

            // Verify MS-ASTASK requirement: MS-ASTASK_R436
            // If Categories element is not returned from server, this element has been removed.
            Site.CaptureRequirementIfIsNull(
                syncedTaskItem.Task.Categories,
                436,
                @"[In Categories] If a Categories element contains no Category child elements in a request [or response], then the categories for the specified task will be removed.");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence whose Type value is 2 and the CalendarType is returned in Sync response.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC20_CreateTaskItemRecursMonthlyWithCalendarTypeReturned()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The CalendarType element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command to create task item which recurs monthly and with CalendarType set.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);

            Request.Categories3 categories = new Request.Categories3 { Category = "Business,Waiting".Split(',') };
            taskItem.Add(Request.ItemsChoiceType8.Categories3, categories);

            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 2,
                Start = DateTime.Now,
                OccurrencesSpecified = true,
                Occurrences = 3,
                DayOfMonthSpecified = true,
                DayOfMonth = 10
            };
            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(1, int.Parse(syncResponse.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Sync command to get the task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R82");

            // Verify MS-ASTASK requirement: MS-ASTASK_R82
            // If CalendarTypeSpecified is true, CalendarType element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.CalendarTypeSpecified,
                82,
                @"[In CalendarType] A command response has a minimum of one CalendarType child element per Recurrence element when the Type element is set to a value of 2;");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence whose Type value is 3 and the CalendarType is returned in Sync response.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC21_CreateTaskItemRecursMonthlyOnTheNthDayWithCalendarTypeReturned()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The CalendarType element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command to create task item which recurs monthly on the nth day and with CalendarType set.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);

            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 3,
                Start = DateTime.Now,
                OccurrencesSpecified = true,
                Occurrences = 3,
                DayOfWeekSpecified = true,
                DayOfWeek = 1,
                WeekOfMonthSpecified = true,
                WeekOfMonth = 2
            };

            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);

            SyncStore syncResponse = this.SyncAddTask(taskItem);
            Site.Assert.AreEqual<int>(1, int.Parse(syncResponse.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);
            #endregion

            #region Call Sync command to get the task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R8211");

            // Verify MS-ASTASK requirement: MS-ASTASK_R8211
            // If CalendarTypeSpecified is true, CalendarType element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.CalendarTypeSpecified,
                8211,
                @"[In CalendarType] A command response has a minimum of one CalendarType child element per Recurrence element when the Type element is set to a value of 3;");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence whose Type value is 5 and the CalendarType is returned in Sync response.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC22_CreateTaskItemRecursYearlyWithCalendarTypeReturned()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The CalendarType element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command to create task item which recurs yearly and with CalendarType set.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);
            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 5,
                Start = DateTime.Now,
                OccurrencesSpecified = true,
                Occurrences = 3,
                DayOfMonthSpecified = true,
                DayOfMonth = 10,
                MonthOfYearSpecified = true,
                MonthOfYear = 2
            };
            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(1, int.Parse(syncResponse.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Sync command to get the task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R8212");

            // Verify MS-ASTASK requirement: MS-ASTASK_R8212
            // If CalendarTypeSpecified is true, CalendarType element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.CalendarTypeSpecified,
                8212,
                @"[In CalendarType] A command response has a minimum of one CalendarType child element per Recurrence element when the Type element is set to a value of 5;");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R204");

            // Verify MS-ASTASK requirement: MS-ASTASK_R204
            Site.CaptureRequirementIfAreEqual<byte?>(
                0,
                syncedTaskItem.Task.Recurrence.IsLeapMonth,
                204,
                @"[In IsLeapMonth] The default value of the IsLeapMonth element is 0.");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with Recurrence whose Type value is 6 and the CalendarType is returned in Sync response.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC23_CreateTaskItemRecursYearlyOnTheNthDayWithCalendarTypeReturned()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The CalendarType element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command to create task item which recurs yearly on the nth day and with CalendarType set.

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");

            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);
            Request.Recurrence1 recurrence = new Request.Recurrence1
            {
                Type = 6,
                Start = DateTime.Now,
                OccurrencesSpecified = true,
                Occurrences = 4,
                DayOfWeekSpecified = true,
                DayOfWeek = 3,
                WeekOfMonthSpecified = true,
                WeekOfMonth = 2,
                MonthOfYearSpecified = true,
                MonthOfYear = 10
            };
            taskItem.Add(Request.ItemsChoiceType8.Recurrence1, recurrence);
            SyncStore syncResponse = this.SyncAddTask(taskItem);

            Site.Assert.AreEqual<int>(1, int.Parse(syncResponse.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Sync command to get the task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R8213");

            // Verify MS-ASTASK requirement: MS-ASTASK_R8213
            // If CalendarTypeSpecified is true, CalendarType element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.CalendarTypeSpecified,
                8213,
                @"[In CalendarType] A command response has a minimum of one CalendarType child element per Recurrence element when the Type element is set to a value of 6;");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing tasks with FirstDayOfWeek returned in response.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S01_TC24_CreateTaskItemRecursWeeklyWithFirstDayOfWeekReturned()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The FirstDayOfWeek element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The FirstDayOfWeek element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command to create a task item which recurs weekly.
            SyncStore initializeSyncResponse = this.TASKAdapter.Sync(Common.CreateInitialSyncRequest(this.UserInformation.TasksCollectionId));

            string subject = Common.GenerateResourceName(Site, "subject");
            string clientId = System.Guid.NewGuid().ToString();
            DateTime startTime = DateTime.Now;
            DateTime utcStartTime = startTime.ToUniversalTime();
            DateTime until = startTime.AddDays(21);

            // Create a task Item with Type 0
            string stringRequest = "<?xml version=\"1.0\" encoding=\"utf-8\"?><Sync xmlns=\"AirSync\"><Collections><Collection><SyncKey>" + initializeSyncResponse.SyncKey + "</SyncKey><CollectionId>" + this.UserInformation.TasksCollectionId + "</CollectionId><DeletesAsMoves>0</DeletesAsMoves><GetChanges>1</GetChanges><WindowSize>512</WindowSize><Options><BodyPreference xmlns=\"AirSyncBase\"><Type>2</Type></BodyPreference></Options><Commands><Add><ClientId>" + clientId + "</ClientId><ApplicationData><Body xmlns=\"AirSyncBase\"><Type>1</Type><Data>Content of the body.</Data></Body><UtcStartDate xmlns=\"Tasks\">" + utcStartTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</UtcStartDate><StartDate xmlns=\"Tasks\">" + startTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</StartDate><UtcDueDate xmlns=\"Tasks\">" + utcStartTime.AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</UtcDueDate><DueDate xmlns=\"Tasks\">" + startTime.AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</DueDate><ReminderSet xmlns=\"Tasks\">1</ReminderSet><ReminderTime xmlns=\"Tasks\">" + startTime.AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</ReminderTime><Subject xmlns=\"Tasks\">" + subject + "</Subject><Importance xmlns=\"Tasks\">0</Importance><Categories xmlns=\"Tasks\"><Category xmlns=\"Tasks\">Business</Category><Category xmlns=\"Tasks\">Waiting</Category></Categories><Recurrence xmlns=\"Tasks\"><Type>1</Type><Start>" + DateTime.Now.AddHours(1).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</Start><Until>" + until.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "</Until><Interval>1</Interval><DayOfWeek>2</DayOfWeek></Recurrence></ApplicationData></Add></Commands></Collection></Collections></Sync>";
            SendStringResponse sendStringResponse = this.TASKAdapter.SendStringRequest(stringRequest, CommandName.Sync);

            // Extract status code from string response
            SyncStore response = this.ExtractSyncStore(sendStringResponse);

            Site.Assert.AreEqual<int>(1, int.Parse(response.AddResponses[0].Status), "Task item should be created successfully.");

            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Sync command to get the task item.

            SyncItem syncedTaskItem = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(syncedTaskItem.Task, "The task which subject is {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R167");

            // Verify MS-ASTASK requirement: MS-ASTASK_R167
            // If FirstDayOfWeekSpecified is true, FirstDayOfWeek element is returned from server.
            Site.CaptureRequirementIfIsTrue(
                syncedTaskItem.Task.Recurrence.FirstDayOfWeekSpecified,
                167,
                @"[In FirstDayOfWeek] The server MUST return a FirstDayOfWeek element when the value of the Type element (section 2.2.2.27) is 1.");
        }
    }
}