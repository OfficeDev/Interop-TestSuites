namespace Microsoft.Protocols.TestSuites.MS_OXWSTASK
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test create, get and delete operations.
    /// </summary>
    [TestClass]
    public class S01_CreateGetDeleteTaskItem : TestSuiteBase
    {
        #region Class initialize and clean up

        /// <summary>
        /// Initialize the test class.
        /// </summary>
        /// <param name="context">Context to initialize.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            TestClassBase.Initialize(context);
        }

        /// <summary>
        /// Clean up the test class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test cases

        /// <summary>
        /// This test case is used to verify the related requirements about DailyRegeneratingPatternType type
        ///  when exchanging the DailyRegeneratingPatternType task item between client and server.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S01_TC01_VerifyDailyRegeneratingPatternType()
        {
            #region Client calls CreateItem to create a task item that contains the recurrence element, which includes the DailyRegeneratingPatternType.

            // Configure the DailyRegeneratingPatternType.
            TaskRecurrenceType taskRecurrence = TestSuiteHelper.GenerateTaskRecurrence(TestSuiteHelper.GenerateDailyRegeneratingPattern, TestSuiteHelper.GenerateNumberedRecurrenceRange);

            // Define the task item.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            TaskType sentTaskItem = TestSuiteHelper.DefineTaskItem(subject, taskRecurrence);

            // Create a task on the server and save the item id.
            ItemIdType[] creatItemIds = this.CreateTasks(sentTaskItem);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = creatItemIds[0];
            #endregion

            #region Client calls GetItem to get the created task item.
            TaskType[] retrievedTaskItemDailys = this.GetTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemDaily = retrievedTaskItemDailys[0];
            #endregion

            #region Verify the related requirements about DailyRegeneratingPatternType type.
            bool isValidDailyRegeneration = retrievedTaskItemDaily.Recurrence.Item is DailyRegeneratingPatternType;
            Site.Log.Add(LogEntryKind.Debug, "The expected recurrence pattern type is DailyRegeneratingPatternType and the actual value is:" + retrievedTaskItemDaily.Recurrence.Item.GetType().Name);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R70");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R70
            Site.CaptureRequirementIfIsTrue(
                isValidDailyRegeneration,
                70,
                @"[In t:DailyRegeneratingPatternType Complex Type] The DailyRegeneratingPatternType complex type extends the RegeneratingPatternBaseType complex type, as specified in section 2.2.4.3.
                    <xs:complexType name=""DailyRegeneratingPatternType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""t:RegeneratingPatternBaseType""/>
                      </xs:complexContent>
                    </xs:complexType>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R241");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R241
            Site.CaptureRequirementIfIsTrue(
                isValidDailyRegeneration,
                241,
                @"[In TaskRecurrencePatternTypes Group] The type of DailyRegeneration is t:DailyRegeneratingPatternType (section 2.2.4.1).");

            NumberedRecurrenceRangeType sentNumRange = sentTaskItem.Recurrence.Item1 as NumberedRecurrenceRangeType;
            NumberedRecurrenceRangeType retrievedNumRange = retrievedTaskItemDaily.Recurrence.Item1 as NumberedRecurrenceRangeType;
            string sentStartDate = sentNumRange.StartDate.ToShortDateString();
            string retrievedStartDate = retrievedNumRange.StartDate.ToUniversalTime().ToShortDateString();

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1354: the expected start date:" + sentStartDate + " the actual value:" + retrievedStartDate);
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1354: the expected number:" + sentNumRange.NumberOfOccurrences + " the actual value:" + retrievedNumRange.NumberOfOccurrences);

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1354
            bool isVerifiedR1354 = retrievedStartDate.Equals(sentStartDate)
                && retrievedNumRange.NumberOfOccurrences == sentNumRange.NumberOfOccurrences;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1354,
                "MS-OXWSCDATA",
                1354,
                @"[In t:RecurrenceRangeTypes Group] The element ""NumberedRecurrence"" with type ""t:NumberedRecurrenceRangeType"" specifies the start date and the number of occurrences of a recurring item.");

            // Verify the TaskRecurrencePatternTypes.
            this.VerifyTaskRecurrencePatternTypes(isValidDailyRegeneration);

            // Verify the IntervalRecurrencePatternBaseType.
            this.VerifyIntervalRecurrencePatternBaseType(sentTaskItem, retrievedTaskItemDaily);

            // After the details of DailyRegeneration verified in above requirements, the following requirement
            // can be verified directly.
            this.Site.CaptureRequirement(
                127,
                @"[In TaskRecurrencePatternTypes Group] DailyRegeneration: Specifies how many days after the completion of the current task the next occurrence will happen.");

            #endregion

            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is used to validate the DailyRecurrencePatternType server behavior related requirements
        ///  when exchanging the DailyRecurrencePatternType task item between client and server.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S01_TC02_VerifyDailyRecurrencePatternType()
        {
            #region Client calls CreateItem to create a task item that contains the recurrence element, which includes the DailyRecurrencePatternType.

            // Configure the DailyRecurrencePatternType.
            TaskRecurrenceType taskRecurrence = TestSuiteHelper.GenerateTaskRecurrence(TestSuiteHelper.GenerateDailyRecurrencePattern, TestSuiteHelper.GenerateEndDateRecurrenceRange);

            // Define the task item.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            TaskType sentTaskItem = TestSuiteHelper.DefineTaskItem(subject, taskRecurrence);

            // Create a task item and save the item id.
            ItemIdType[] creatItemIds = this.CreateTasks(sentTaskItem);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = creatItemIds[0];
            #endregion

            #region Client calls GetItem to get the created task item.
            TaskType[] retrievedTaskItemDailys = GetTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemDaily = retrievedTaskItemDailys[0];
            #endregion

            #region Verify the related requirements about DailyRecurrencePatternType type.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R240");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R240
            bool isValidDailyRecurrence = retrievedTaskItemDaily.Recurrence.Item is DailyRecurrencePatternType;

            Site.CaptureRequirementIfIsTrue(
                isValidDailyRecurrence,
                240,
                @"[In TaskRecurrencePatternTypes Group] The type of DailyRecurrence is t:DailyRecurrencePatternType ([MS-OXWSCDATA] section 2.2.4.24).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1109");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1109 
            Site.CaptureRequirementIfIsTrue(
                isValidDailyRecurrence,
                "MS-OXWSCDATA",
                1109,
                @"[In t:DailyRecurrencePatternType Complex Type] The type [DailyRecurrencePatternType] is defined as follow:
                     <xs:complexType name=""DailyRecurrencePatternType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""t:IntervalRecurrencePatternBaseType""
                         />
                      </xs:complexContent>
                    </xs:complexType>");

            EndDateRecurrenceRangeType sentEndDateRange = sentTaskItem.Recurrence.Item1 as EndDateRecurrenceRangeType;
            EndDateRecurrenceRangeType retrievedEndDateRange = retrievedTaskItemDaily.Recurrence.Item1 as EndDateRecurrenceRangeType;
            string sentStartDate = sentEndDateRange.StartDate.ToShortDateString();
            string retrievedStartDate = retrievedEndDateRange.StartDate.ToUniversalTime().ToShortDateString();
            string sentEndDate = sentEndDateRange.EndDate.ToShortDateString();
            string retrievedEndDate = retrievedEndDateRange.EndDate.ToUniversalTime().ToShortDateString();

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1353: the expected start date:" + sentStartDate + " the actual value:" + retrievedStartDate);
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1353: the expected end date:" + sentEndDate + " the actual value:" + retrievedEndDate);

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1353
            bool isVerifiedR1353 = retrievedStartDate.Equals(sentStartDate)
                && retrievedEndDate.Equals(sentEndDate);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1353,
                "MS-OXWSCDATA",
                1353,
                @"[In t:RecurrenceRangeTypes Group] The element ""EndDateRecurrence"" with type ""t:EndDateRecurrenceRangeType"" specifies the start date and the end date of an item recurrence pattern.");

            // Verify TaskRecurrencePatternTypes.
            this.VerifyTaskRecurrencePatternTypes(isValidDailyRecurrence);

            // Verify the IntervalRecurrencePatternBaseType.
            this.VerifyIntervalRecurrencePatternBaseType(sentTaskItem, retrievedTaskItemDaily);

            // After the details of DailyRecurrence verified in above requirements, the following requirement
            // can be verified directly.            
            this.Site.CaptureRequirement(
                125,
                @"[In TaskRecurrencePatternTypes Group] DailyRecurrence: Specifies the interval, in days, at which a task recurs.");

            #endregion

            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is used to validate the MonthlyRegeneratingPatternType server behavior related requirements
        ///  when exchanging the MonthlyRegeneratingPatternType task item between client and server.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S01_TC03_VerifyMonthlyRegeneratingPatternType()
        {
            #region Client calls CreateItem to create a task item that contains the recurrence element, which includes the MonthlyRegeneratingPatternType.

            // Configure the MonthlyRegeneratingPatternType.
            TaskRecurrenceType taskRecurrence = TestSuiteHelper.GenerateTaskRecurrence(TestSuiteHelper.GenerateMonthlyRegeneratingPattern, TestSuiteHelper.GenerateNoEndRecurrenceRange);

            // Define the task item.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            TaskType sentTaskItem = TestSuiteHelper.DefineTaskItem(subject, taskRecurrence);

            // Create a task item and save the item id.
            ItemIdType[] createItemIds = this.CreateTasks(sentTaskItem);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion

            #region Client calls GetItem to get the created task item.
            TaskType[] retrievedTaskItemMonthlys = GetTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemMonthly = retrievedTaskItemMonthlys[0];
            #endregion

            #region Verify the related requirements about MonthlyRegeneratingPatternType type.
            bool isValidMonthlyRegeneratingPattern = retrievedTaskItemMonthly.Recurrence.Item is MonthlyRegeneratingPatternType;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R73");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R73
            Site.CaptureRequirementIfIsTrue(
                isValidMonthlyRegeneratingPattern,
                73,
                @"[In t:MonthlyRegeneratingPatternType Complex Type] The MonthlyRegeneratingPatternType complex type extends the RegeneratingPatternBaseType complex type, as specified in section 2.2.4.3.
                    <xs:complexType name=""MonthlyRegeneratingPatternType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""t:RegeneratingPatternBaseType""/>
                      </xs:complexContent>
                    </xs:complexType>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R243");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R243
            Site.CaptureRequirementIfIsTrue(
                isValidMonthlyRegeneratingPattern,
                243,
                @"[In TaskRecurrencePatternTypes Group] The type of MonthlyRegeneration is t:MonthlyRegeneratingPatternType (section 2.2.4.2).");

            NoEndRecurrenceRangeType sentNoEndDateRange = sentTaskItem.Recurrence.Item1 as NoEndRecurrenceRangeType;
            NoEndRecurrenceRangeType retrievedNoEndDateRange = retrievedTaskItemMonthly.Recurrence.Item1 as NoEndRecurrenceRangeType;
            string sentStartDate = sentNoEndDateRange.StartDate.ToShortDateString();
            string retrievedStartDate = retrievedNoEndDateRange.StartDate.ToUniversalTime().ToShortDateString();

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1352: the expected start date:" + sentStartDate + " the actual value:" + retrievedStartDate);

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1352
            this.Site.CaptureRequirementIfAreEqual<string>(
                sentStartDate,
                retrievedStartDate,
                "MS-OXWSCDATA",
                1352,
                @"[In t:RecurrenceRangeTypes Group] The element ""NoEndRecurrence"" with type ""t:NoEndRecurrenceRangeType"" specifies a recurrence pattern that does not have a defined end date.");

            // Verify TaskRecurrencePatternTypes.
            this.VerifyTaskRecurrencePatternTypes(isValidMonthlyRegeneratingPattern);

            // Verify the IntervalRecurrencePatternBaseType.
            this.VerifyIntervalRecurrencePatternBaseType(sentTaskItem, retrievedTaskItemMonthly);

            // After the details of MonthlyRegeneration verified in above requirements, the following requirement
            // can be verified directly. 
            this.Site.CaptureRequirement(
                131,
                @"[In TaskRecurrencePatternTypes Group] MonthlyRegeneration: Specifies how many months after the completion of the current task the next occurrence will happen.");

            #endregion

            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is used to validate the WeeklyRegeneratingPatternType server behavior related requirements
        ///  when exchanging the WeeklyRegeneratingPatternType task item between client and server.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S01_TC04_VerifyWeeklyRegeneratingPatternType()
        {
            #region Client calls CreateItem to create a task item that contains the recurrence element, which includes the WeeklyRegeneratingPatternType.

            // Configure the WeeklyRegeneratingPatternType.
            TaskRecurrenceType taskRecurrence = TestSuiteHelper.GenerateTaskRecurrence(TestSuiteHelper.GenerateWeeklyRegeneratingPattern, TestSuiteHelper.GenerateNumberedRecurrenceRange);

            // Define a task item.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            TaskType sentTaskItem = TestSuiteHelper.DefineTaskItem(subject, taskRecurrence);

            // Create a task and save the item id.
            ItemIdType[] createItemIds = this.CreateTasks(sentTaskItem);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion

            #region Client calls GetItem to get the created task item.
            TaskType[] retrievedTaskItemWeeklys = GetTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemWeekly = retrievedTaskItemWeeklys[0];
            #endregion

            #region Verify the related requirements about WeeklyRegeneratingPatternType type.
            bool isValidWeeklyRegeneratingPattern = retrievedTaskItemWeekly.Recurrence.Item is WeeklyRegeneratingPatternType;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R78");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R78
            Site.CaptureRequirementIfIsTrue(
                isValidWeeklyRegeneratingPattern,
                78,
                @"[In t:WeeklyRegeneratingPatternType Complex Type] The WeeklyRegeneratingPatternType complex type extends the RegeneratingPatternBaseType complex type, as specified in section 2.2.4.3.
                    <xs:complexType name=""WeeklyRegeneratingPatternType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""t:RegeneratingPatternBaseType""/>
                      </xs:complexContent>
                    </xs:complexType>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R242");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R242
            Site.CaptureRequirementIfIsTrue(
                isValidWeeklyRegeneratingPattern,
                242,
                @"[In TaskRecurrencePatternTypes Group] The type of WeeklyRegeneration is t:WeeklyRegeneratingPatternType (section 2.2.4.7).");

            // Verify TaskRecurrencePatternTypes.
            this.VerifyTaskRecurrencePatternTypes(isValidWeeklyRegeneratingPattern);

            // Verify the IntervalRecurrencePatternBaseType.
            this.VerifyIntervalRecurrencePatternBaseType(sentTaskItem, retrievedTaskItemWeekly);

            // After the details of WeeklyRegeneration verified in above requirements, the following requirement
            // can be verified directly.             
            this.Site.CaptureRequirement(
                129,
                @"[In TaskRecurrencePatternTypes Group] WeeklyRegeneration: Specifies how many weeks after the completion of the current task the next occurrence will happen.");

            #endregion

            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is used to validate the YearlyRegeneratingPatternType server behavior related requirements
        ///  when exchanging the YearlyRegeneratingPatternType task item between client and server. 
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S01_TC05_VerifyYearlyRegeneratingPatternType()
        {
            #region Client calls CreateItem to create a task item that contains the recurrence element, which includes the YearlyRegeneratingPatternType.

            // Configure the YearlyRegeneratingPatternType.
            TaskRecurrenceType taskRecurrence = TestSuiteHelper.GenerateTaskRecurrence(TestSuiteHelper.GenerateYearlyRegeneratingPattern, TestSuiteHelper.GenerateNumberedRecurrenceRange);

            // Define a task item.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            TaskType sentTaskItem = TestSuiteHelper.DefineTaskItem(subject, taskRecurrence);

            // Create a task and save the item id.
            ItemIdType[] createItemIds = this.CreateTasks(sentTaskItem);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion

            #region Client calls GetItem to get the created task item.
            TaskType[] retrievedTaskItemYearlys = GetTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemYearly = retrievedTaskItemYearlys[0];
            #endregion

            #region Verify the related requirements about YearlyRegeneratingPatternType type.
            bool isValidYearlyRegeneratingPattern = retrievedTaskItemYearly.Recurrence.Item is YearlyRegeneratingPatternType;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R81");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R81
            Site.CaptureRequirementIfIsTrue(
                isValidYearlyRegeneratingPattern,
                81,
                @"[In t:YearlyRegeneratingPatternType Complex Type] The YearlyRegeneratingPatternType complex type extends the RegeneratingPatternBaseType complex type, as specified in section 2.2.4.3.
                    <xs:complexType name=""YearlyRegeneratingPatternType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""t:RegeneratingPatternBaseType""/>
                      </xs:complexContent>
                    </xs:complexType>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R244");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R244
            Site.CaptureRequirementIfIsTrue(
                isValidYearlyRegeneratingPattern,
                244,
                @"[In TaskRecurrencePatternTypes Group] The type of YearlyRegeneration is t:YearlyRegeneratingPatternType (section 2.2.4.8).");

            // Verify the TaskRecurrencePatternTypes.
            this.VerifyTaskRecurrencePatternTypes(isValidYearlyRegeneratingPattern);

            // Verify the IntervalRecurrencePatternBaseType.
            this.VerifyIntervalRecurrencePatternBaseType(sentTaskItem, retrievedTaskItemYearly);

            // After the details of YearlyRegeneration verified in above requirements, the following requirement
            // can be verified directly.              
            this.Site.CaptureRequirement(
                133,
                @"[In TaskRecurrencePatternTypes Group] YearlyRegeneration: Specifies how many years after the completion of the current task the next occurrence will happen.");

            #endregion

            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is used to validate the YearlyRecurrencePatternType server behavior related requirements
        ///  when exchanging the YearlyRecurrencePatternType task item between client and server.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S01_TC06_VerifyRelativeYearlyRecurrencePatternType()
        {
            #region Client calls CreateItem to create a task item that contains the recurrence element, which includes the RelativeYearlyRecurrencePatternType.

            // Configure the RelativeYearlyRecurrencePatternType.
            TaskRecurrenceType taskRecurrence = TestSuiteHelper.GenerateTaskRecurrence(TestSuiteHelper.GenerateRelativeYearlyRecurrencePattern, TestSuiteHelper.GenerateNumberedRecurrenceRange);

            // Define a task item.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            TaskType sentTaskItem = TestSuiteHelper.DefineTaskItem(subject, taskRecurrence);

            // Create a task and save the item id.
            ItemIdType[] createItemIds = this.CreateTasks(sentTaskItem);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion

            #region Client calls GetItem to get the created task item.
            TaskType[] retrievedTaskItemYearlys = GetTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemYearly = retrievedTaskItemYearlys[0];
            #endregion

            #region Verify the related requirements about RelativeYearlyRecurrencePatternType type.
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R235");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R235
            bool isValidRelativeYearlyRecurrence = retrievedTaskItemYearly.Recurrence.Item is RelativeYearlyRecurrencePatternType;

            Site.CaptureRequirementIfIsTrue(
                isValidRelativeYearlyRecurrence,
                235,
                @"[In TaskRecurrencePatternTypes Group] The type of RelativeYearlyRecurrence is t:RelativeYearlyRecurrencePatternType ([MS-OXWSCDATA] section 2.2.4.63).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1259");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1259 
            Site.CaptureRequirementIfIsTrue(
                isValidRelativeYearlyRecurrence,
                "MS-OXWSCDATA",
                1259,
                @"[In t:RelativeYearlyRecurrencePatternType Complex Type] The type [RelativeYearlyRecurrencePatternType] is defined as follow:
                    <xs:complexType name=""RelativeYearlyRecurrencePatternType"">
                        <xs:complexContent>
                        <xs:extension
                            base=""t:RecurrencePatternBaseType""
                        >
                            <xs:sequence>
                            <xs:element name=""DaysOfWeek""
                                type=""t:DayOfWeekType""
                                />
                            <xs:element name=""DayOfWeekIndex""
                                type=""t:DayOfWeekIndexType""
                                />
                            <xs:element name=""Month""
                                type=""t:MonthNamesType""
                                />
                            </xs:sequence>
                        </xs:extension>
                        </xs:complexContent>
                    </xs:complexType>");

            RelativeYearlyRecurrencePatternType sentRecurPattern = sentTaskItem.Recurrence.Item as RelativeYearlyRecurrencePatternType;
            RelativeYearlyRecurrencePatternType retrievedRecurPattern = retrievedTaskItemYearly.Recurrence.Item as RelativeYearlyRecurrencePatternType;

            // Add the debug information.
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1261: the expected DaysOfWeek:" + sentRecurPattern.DaysOfWeek + " the actual value:" + retrievedRecurPattern.DaysOfWeek);

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1261 
            bool isVerifyR1261 = isValidRelativeYearlyRecurrence && retrievedRecurPattern.DaysOfWeek.Equals(sentRecurPattern.DaysOfWeek, StringComparison.CurrentCultureIgnoreCase);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1261,
                "MS-OXWSCDATA",
                1261,
                @"[In t:RelativeYearlyRecurrencePatternType Complex Type] This element [DaysOfWeek] MUST be present.");

            // Add the debug information.
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1263: the expected DayOfWeekIndex:" + sentRecurPattern.DayOfWeekIndex + " the actual value:" + retrievedRecurPattern.DayOfWeekIndex);

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1263 
            bool isVerifyR1263 = isValidRelativeYearlyRecurrence && retrievedRecurPattern.DayOfWeekIndex.Equals(sentRecurPattern.DayOfWeekIndex);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1263,
                "MS-OXWSCDATA",
                1263,
                @"[In t:RelativeYearlyRecurrencePatternType Complex Type] This element [DayOfWeekIndex] MUST be present.");

            // Add the debug information.
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1265: the expected Month:" + sentRecurPattern.Month + " the actual value:" + retrievedRecurPattern.Month);

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1265 
            bool isVerifyR1265 = isValidRelativeYearlyRecurrence && retrievedRecurPattern.Month.Equals(sentRecurPattern.Month);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1265,
                "MS-OXWSCDATA",
                1265,
                @"[In t:RelativeYearlyRecurrencePatternType Complex Type] This element [Month] MUST be present.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1264");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1264
            this.Site.CaptureRequirementIfAreEqual<MonthNamesType>(
                sentRecurPattern.Month,
                retrievedRecurPattern.Month,
                "MS-OXWSCDATA",
                1264,
                @"[In t:RelativeYearlyRecurrencePatternType Complex Type] The element ""Month"" with type ""t:MonthNamesType"" specifies the month when a yearly recurring item occurs.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1262");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1262                
            this.Site.CaptureRequirementIfAreEqual<DayOfWeekIndexType>(
                sentRecurPattern.DayOfWeekIndex,
                retrievedRecurPattern.DayOfWeekIndex,
                "MS-OXWSCDATA",
                1262,
                @"[In t:RelativeYearlyRecurrencePatternType Complex Type] The element ""DayOfWeekIndex"" with type ""t:DayOfWeekIndexType"" specifies the days of the week that are used in a relative yearly recurrence pattern.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1260");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1260                
            this.Site.CaptureRequirementIfAreEqual<String>(
                sentRecurPattern.DaysOfWeek,
                retrievedRecurPattern.DaysOfWeek,
                "MS-OXWSCDATA",
                1260,
                @"[In t:RelativeYearlyRecurrencePatternType Complex Type] The element ""DayOfWeekIndex"" with type ""t:DayOfWeekIndexType"" specifies the week that is used in a relative monthly recurrence pattern.");

            // Verify the TaskRecurrencePatternTypes.
            this.VerifyTaskRecurrencePatternTypes(isValidRelativeYearlyRecurrence);

            // After the details of RelativeYearlyRecurrence verified in above requirements, the following requirement
            // can be verified directly.              
            this.Site.CaptureRequirement(
                115,
                @"[In TaskRecurrencePatternTypes Group] RelativeYearlyRecurrence: Specifies a relative yearly recurrence pattern for a recurring task.");

            #endregion

            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is used to validate the AbsoluteYearlyRecurrencePatternType server behavior related requirements
        ///  when exchanging the AbsoluteYearlyRecurrencePatternType task item between client and server.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S01_TC07_VerifyAbsoluteYearlyRecurrencePatternType()
        {
            #region Client calls CreateItem to create a task item that contains the recurrence element, which includes the AbsoluteYearlyRecurrencePatternType.

            // Configure the AbsoluteYearlyRecurrencePatternType.
            TaskRecurrenceType taskRecurrence = TestSuiteHelper.GenerateTaskRecurrence(TestSuiteHelper.GenerateAbsoluteYearlyRecurrencePattern, TestSuiteHelper.GenerateNumberedRecurrenceRange);

            // Define a task item.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            TaskType sentTaskItem = TestSuiteHelper.DefineTaskItem(subject, taskRecurrence);

            // Create a task and save the item id.
            ItemIdType[] createItemIds = this.CreateTasks(sentTaskItem);     
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion

            #region Client calls GetItem to get the created task item.
            TaskType[] retrievedTaskItemYearlys = this.GetTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemYearly = retrievedTaskItemYearlys[0];
            #endregion

            #region Verify the related requirements about AbsoluteYearlyRecurrencePatternType type.
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R236");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R236
            bool isValidAbsoluteYearlyRecurrencePattern = retrievedTaskItemYearly.Recurrence.Item is AbsoluteYearlyRecurrencePatternType;

            Site.CaptureRequirementIfIsTrue(
                isValidAbsoluteYearlyRecurrencePattern,
                236,
                @"[In TaskRecurrencePatternTypes Group] The type of AbsoluteYearlyRecurrence is t:AbsoluteYearlyRecurrencePatternType ([MS-OXWSCDATA] section 2.2.4.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R999");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R999 
            Site.CaptureRequirementIfIsTrue(
                isValidAbsoluteYearlyRecurrencePattern,
                "MS-OXWSCDATA",
                999,
                @"[In t:AbsoluteYearlyRecurrencePatternType Complex Type] The type [AbsoluteYearlyRecurrencePatternType] is defined as follow:
                    <xs:complexType name=""AbsoluteYearlyRecurrencePatternType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""t:RecurrencePatternBaseType""
                        >
                          <xs:sequence>
                            <xs:element name=""DayOfMonth""
                              type=""xs:int""
                             />
                            <xs:element name=""Month""
                              type=""t:MonthNamesType""
                             />
                          </xs:sequence>
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");

            AbsoluteYearlyRecurrencePatternType sentRecurPattern = sentTaskItem.Recurrence.Item as AbsoluteYearlyRecurrencePatternType;
            AbsoluteYearlyRecurrencePatternType retrievedRecurPattern = retrievedTaskItemYearly.Recurrence.Item as AbsoluteYearlyRecurrencePatternType;

            // Add the debug information.
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1001: the expected DayOfMonth:" + sentRecurPattern.DayOfMonth + " the actual value:" + retrievedRecurPattern.DayOfMonth);

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1001 
            bool isVerifyR1001 = isValidAbsoluteYearlyRecurrencePattern && retrievedRecurPattern.DayOfMonth == sentRecurPattern.DayOfMonth;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1001,
                "MS-OXWSCDATA",
                1001,
                @"[In t:AbsoluteYearlyRecurrencePatternType Complex Type] This property [DayOfMonth] MUST be present.");

            // Add the debug information.
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1003: the expected Month:" + sentRecurPattern.Month + " the actual value:" + retrievedRecurPattern.Month);

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1003 
            bool isVerifyR1003 = isValidAbsoluteYearlyRecurrencePattern && retrievedRecurPattern.Month.Equals(sentRecurPattern.Month);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1003,
                "MS-OXWSCDATA",
                1003,
                @"[In t:AbsoluteYearlyRecurrencePatternType Complex Type] This property [Month] MUST be present.");

            // Verify TaskRecurrencePatternTypes.
            this.VerifyTaskRecurrencePatternTypes(isValidAbsoluteYearlyRecurrencePattern);

            // After the details of AbsoluteYearlyRecurrence verified in above requirements, the following requirement
            // can be verified directly.             
            this.Site.CaptureRequirement(
                117,
                @"[In TaskRecurrencePatternTypes Group] AbsoluteYearlyRecurrence: Specifies a yearly recurrence pattern for a recurring task.");

            #endregion

            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is used to validate the RelativeMonthlyRecurrencePatternType server behavior related requirements
        ///  when exchanging the RelativeMonthlyRecurrencePatternType task item between client and server.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S01_TC08_VerifyRelativeMonthlyRecurrencePatternType()
        {
            #region Client calls CreateItem to create a task item that contains the recurrence element, which includes the RelativeMonthlyRecurrencePatternType.

            // Configure the RelativeMonthlyRecurrencePatternType.
            TaskRecurrenceType taskRecurrence = TestSuiteHelper.GenerateTaskRecurrence(TestSuiteHelper.GenerateRelativeMonthlyRecurrencePattern, TestSuiteHelper.GenerateNumberedRecurrenceRange);

            // Define a task item.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            TaskType sentTaskItem = TestSuiteHelper.DefineTaskItem(subject, taskRecurrence);

            // Create a task and save the item id.
            ItemIdType[] createItemIds = this.CreateTasks(sentTaskItem);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion

            #region Client calls GetItem to get the created task item.
            TaskType[] retrievedTaskItemMonthlys = this.GetTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemMonthly = retrievedTaskItemMonthlys[0];
            #endregion

            #region Verify the related requirements about RelativeMonthlyRecurrencePatternType type.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R237");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R237
            bool isValidRelativeMonthlyRecurrence = retrievedTaskItemMonthly.Recurrence.Item is RelativeMonthlyRecurrencePatternType;

            Site.CaptureRequirementIfIsTrue(
                isValidRelativeMonthlyRecurrence,
                237,
                @"[In TaskRecurrencePatternTypes Group] The type of RelativeMonthlyRecurrence is t:RelativeMonthlyRecurrencePatternType ([MS-OXWSCDATA] section 2.2.4.62).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1255");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1255 
            Site.CaptureRequirementIfIsTrue(
                isValidRelativeMonthlyRecurrence,
                "MS-OXWSCDATA",
                1255,
                @"[In t:RelativeMonthlyRecurrencePatternType Complex Type] The type [RelativeMonthlyRecurrencePatternType] is defined as follow:
                    <xs:complexType name=""RelativeMonthlyRecurrencePatternType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""t:IntervalRecurrencePatternBaseType""
                        >
                          <xs:sequence>
                            <xs:element name=""DaysOfWeek""
                              type=""t:DayOfWeekType""
                             />
                            <xs:element name=""DayOfWeekIndex""
                              type=""t:DayOfWeekIndexType""
                             />
                          </xs:sequence>
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");

            RelativeMonthlyRecurrencePatternType sentRecurPattern = sentTaskItem.Recurrence.Item as RelativeMonthlyRecurrencePatternType;
            RelativeMonthlyRecurrencePatternType retrievedRecurPattern = retrievedTaskItemMonthly.Recurrence.Item as RelativeMonthlyRecurrencePatternType;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1256");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1256
            this.Site.CaptureRequirementIfAreEqual<DayOfWeekType>(
                sentRecurPattern.DaysOfWeek,
                retrievedRecurPattern.DaysOfWeek,
                "MS-OXWSCDATA",
                1256,
                @"[In t:RelativeMonthlyRecurrencePatternType Complex Type] The element ""DaysOfWeek"" with type ""t:DayOfWeekType"" specifies the days of the week that are used in a relative monthly recurrence pattern.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1257");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1257
            this.Site.CaptureRequirementIfAreEqual<DayOfWeekIndexType>(
                sentRecurPattern.DayOfWeekIndex,
                retrievedRecurPattern.DayOfWeekIndex,
                "MS-OXWSCDATA",
                1257,
                @"[In t:RelativeMonthlyRecurrencePatternType Complex Type] The element ""DayOfWeekIndex"" with type ""t:DayOfWeekIndexType"" specifies the week that is used in a relative monthly recurrence pattern.");

            // Verify TaskRecurrencePatternTypes.
            this.VerifyTaskRecurrencePatternTypes(isValidRelativeMonthlyRecurrence);

            // Verify the IntervalRecurrencePatternBaseType.
            this.VerifyIntervalRecurrencePatternBaseType(sentTaskItem, retrievedTaskItemMonthly);

            // After the details of RelativeMonthlyRecurrence verified in above requirements, the following requirement
            // can be verified directly.              
            this.Site.CaptureRequirement(
                119,
                @"[In TaskRecurrencePatternTypes Group] RelativeMonthlyRecurrence:  Specifies a relative monthly recurrence pattern for a recurring task.");

            #endregion

            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is used to validate the AbsoluteMonthlyRecurrencePatternType server behavior related requirements
        ///  when exchanging the AbsoluteMonthlyRecurrencePatternType task item between client and server.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S01_TC09_VerifyAbsoluteMonthlyRecurrencePatternType()
        {
            #region Client calls CreateItem to create a task item that contains the recurrence element, which includes the AbsoluteMonthlyRecurrencePatternType.

            // Configure the AbsoluteMonthlyRecurrencePatternType.
            TaskRecurrenceType taskRecurrence = TestSuiteHelper.GenerateTaskRecurrence(TestSuiteHelper.GenerateAbsoluteMonthlyRecurrencePattern, TestSuiteHelper.GenerateNumberedRecurrenceRange);

            // Define a task item.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            TaskType sentTaskItem = TestSuiteHelper.DefineTaskItem(subject, taskRecurrence);

            // Create a task and save the item id.
            ItemIdType[] createItemIds = this.CreateTasks(sentTaskItem);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion

            #region Client calls GetItem to get the created task item.
            TaskType[] retrievedTaskItemMonthlys = this.GetTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemMonthly = retrievedTaskItemMonthlys[0];
            #endregion

            #region Verify the related requirements about AbsoluteMonthlyRecurrencePatternType type.
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R238");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R238
            bool isValidAbsoluteMonthlyRecurrence = retrievedTaskItemMonthly.Recurrence.Item is AbsoluteMonthlyRecurrencePatternType;

            Site.CaptureRequirementIfIsTrue(
                isValidAbsoluteMonthlyRecurrence,
                238,
                @"[In TaskRecurrencePatternTypes Group] The type of AbsoluteMonthlyRecurrence is t:AbsoluteMonthlyRecurrencePatternType ([MS-OXWSCDATA] section 2.2.4.1).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R995");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R995 
            Site.CaptureRequirementIfIsTrue(
                isValidAbsoluteMonthlyRecurrence,
                "MS-OXWSCDATA",
                995,
                @"[In t:AbsoluteMonthlyRecurrencePatternType Complex Type] The type [AbsoluteMonthlyRecurrencePatternType] is defined as follow:
                    <xs:complexType name=""AbsoluteMonthlyRecurrencePatternType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""t:IntervalRecurrencePatternBaseType""
                        >
                          <xs:sequence>
                            <xs:element name=""DayOfMonth""
                              type=""xs:int""
                             />
                          </xs:sequence>
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");

            AbsoluteMonthlyRecurrencePatternType sentRecurPattern = sentTaskItem.Recurrence.Item as AbsoluteMonthlyRecurrencePatternType;
            AbsoluteMonthlyRecurrencePatternType retrievedRecurPattern = retrievedTaskItemMonthly.Recurrence.Item as AbsoluteMonthlyRecurrencePatternType;

            // Add the debug information.
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R997: the expected DayOfMonth:" + sentRecurPattern.DayOfMonth + " the actual value:" + retrievedRecurPattern.DayOfMonth);

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R997
            bool isVerifyR997 = isValidAbsoluteMonthlyRecurrence && retrievedRecurPattern.DayOfMonth == sentRecurPattern.DayOfMonth;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR997,
                "MS-OXWSCDATA",
                997,
                @"[In t:AbsoluteMonthlyRecurrencePatternType Complex Type] This property [DayOfMonth] MUST be present.");

            // Verify TaskRecurrencePatternTypes.
            this.VerifyTaskRecurrencePatternTypes(isValidAbsoluteMonthlyRecurrence);

            // Verify the IntervalRecurrencePatternBaseType.
            this.VerifyIntervalRecurrencePatternBaseType(sentTaskItem, retrievedTaskItemMonthly);

            // After the details of AbsoluteMonthlyRecurrence verified in above requirements, the following requirement
            // can be verified directly.           
            this.Site.CaptureRequirement(
                121,
                @"[In TaskRecurrencePatternTypes Group] AbsoluteMonthlyRecurrence: Specifies a monthly recurrence pattern for a recurring task.");

            #endregion

            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is used to validate the WeeklyRecurrencePatternType server behavior related requirements
        ///  when exchanging the WeeklyRecurrencePatternType task item between client and server.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S01_TC10_VerifyWeeklyRecurrencePatternType()
        {
            #region Client calls CreateItem to create a task item that contains the recurrence element, which includes the WeeklyRecurrencePatternType.

            #region Configure the WeeklyRecurrencePatternType.
            TaskRecurrenceType taskRecurrence = TestSuiteHelper.GenerateTaskRecurrence(TestSuiteHelper.GenerateWeeklyRecurrencePattern, TestSuiteHelper.GenerateNumberedRecurrenceRange);

            // According to the description of MS-OXWSCDATA, Exchange 2013 and Exchange 2010 do include the FirstDayOfWeek element.
            bool isR4005Implementated = Common.IsRequirementEnabled(4005, this.Site);
            if (isR4005Implementated)
            {
                // Set FirstDayOfWeek of WeeklyRecurrencePatternType to any FirstDayOfWeekType string value. Here is set to "Tuesday".
                (taskRecurrence.Item as WeeklyRecurrencePatternType).FirstDayOfWeek = "Tuesday";
            }

            // According to the description of MS-OXWSCDATA, Exchange 2007 do not include the FirstDayOfWeek element.
            bool isR1489Implementated = Common.IsRequirementEnabled(1489, this.Site);
            if (isR1489Implementated)
            {
                (taskRecurrence.Item as WeeklyRecurrencePatternType).FirstDayOfWeek = null;
            }

            #endregion

            // Define a task item.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            TaskType sentTaskItem = TestSuiteHelper.DefineTaskItem(subject, taskRecurrence);

            // Create a task and save the item id.
            ItemIdType[] createItemIds = this.CreateTasks(sentTaskItem);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion

            #region Client calls GetItem to get the created task item.
            TaskType[] retrievedTaskItemWeeklys = this.GetTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemWeekly = retrievedTaskItemWeeklys[0];
            #endregion

            #region Verify the related requirements about WeeklyRecurrencePatternType type.
            WeeklyRecurrencePatternType sentRecurPattern = sentTaskItem.Recurrence.Item as WeeklyRecurrencePatternType;
            WeeklyRecurrencePatternType retrievedRecurPattern = retrievedTaskItemWeekly.Recurrence.Item as WeeklyRecurrencePatternType;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R239");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R239
            bool isValidWeeklyRecurrence = retrievedTaskItemWeekly.Recurrence.Item is WeeklyRecurrencePatternType;

            Site.CaptureRequirementIfIsTrue(
                isValidWeeklyRecurrence,
                239,
                @"[In TaskRecurrencePatternTypes Group] The type of WeeklyRecurrence is t:WeeklyRecurrencePatternType ([MS-OXWSCDATA] section 2.2.4.77).");

            if (isR1489Implementated)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1489");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1489
                Site.CaptureRequirementIfIsNull(
                    retrievedRecurPattern.FirstDayOfWeek,
                    "MS-OXWSCDATA",
                    1489,
                    @"[In Appendix C: Product Behavior] Implementation does not include the FirstDayOfWeek element. (<253> Section 2.2.4.64: Exchange 2007 do not include the FirstDayOfWeek element.)");
            }

            if (isR4005Implementated)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R4005");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R4005
                Site.CaptureRequirementIfAreEqual<string>(
                    (taskRecurrence.Item as WeeklyRecurrencePatternType).FirstDayOfWeek,
                    retrievedRecurPattern.FirstDayOfWeek,
                    "MS-OXWSCDATA",
                    4005,
                    @"[In Appendix C: Product Behavior] Implementation does include the element ""FirstDayOfWeek"" with type ""t:DayOfWeekType (section 2.2.3.5)"" which specifies the first day of the week. (Exchange Server 2013 and above follow this behavior.)");
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1307");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1307 
            Site.CaptureRequirementIfIsTrue(
                isValidWeeklyRecurrence,
                "MS-OXWSCDATA",
                1307,
                @"[In t:WeeklyRecurrencePatternType Complex Type] The type [WeeklyRecurrencePatternType] is defined as follow:
                    <xs:complexType name=""WeeklyRecurrencePatternType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""t:IntervalRecurrencePatternBaseType""
                        >
                          <xs:sequence>
                            <xs:element name=""DaysOfWeek""
                              type=""t:DaysOfWeekType""
                             />
                            <xs:element name=""FirstDayOfWeek""
                              type=""t:DayOfWeekType""
                              minOccurs=""0""
                             />
                          </xs:sequence>
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1308");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1308
            this.Site.CaptureRequirementIfAreEqual<string>(
                sentRecurPattern.DaysOfWeek,
                retrievedRecurPattern.DaysOfWeek,
                "MS-OXWSCDATA",
                1308,
                @"[In t:WeeklyRecurrencePatternType Complex Type] The element ""DaysOfWeek"" with type ""t:DaysOfWeekType"" specifies the days of the week that are in the weekly recurrence pattern.");

            // Verify TaskRecurrencePatternTypes.
            this.VerifyTaskRecurrencePatternTypes(isValidWeeklyRecurrence);

            // Verify the IntervalRecurrencePatternBaseType.
            this.VerifyIntervalRecurrencePatternBaseType(sentTaskItem, retrievedTaskItemWeekly);

            // After the details of WeeklyRecurrence verified in above requirements, the following requirement
            // can be verified directly.            
            this.Site.CaptureRequirement(
                123,
                @"[In TaskRecurrencePatternTypes Group] WeeklyRecurrence: Specifies the weekly interval at which and the days on which a task recurs.");

            #endregion

            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is used to validate the TaskDelegateStateType server behavior related requirements.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S01_TC11_VerifyTaskDelegateStateType()
        {
            #region Client calls CreateItem to create a task item with the DelegationState element is not set to any value.
            // Create a task and save the item id.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            ItemIdType[] createItemIds = this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, null));
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion

            #region Client calls CreateItem to create a task with the DelegationState element is set to the Declined taskDelegateState.
            // Set the taskDelegateState of task item to TaskDelegateStateType.Declined. Except that the response code is ErrorInvalidPropertySet.
            TaskDelegateStateType taskDelegateState = TaskDelegateStateType.Declined;
            subject = Common.GenerateResourceName(this.Site, "This is a task");
            this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, taskDelegateState, true));
            Site.Assert.AreNotEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should not be success!", null);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, (ResponseCodeType)this.ResponseCode[0], "This create response status information should be ErrorInvalidPropertySet!", null);
            #endregion

            #region Client calls CreateItem to create a task with the DelegationState element is set to the Accepted taskDelegateState.

            // Set the taskDelegateState of task item to TaskDelegateStateType.Accepted. Except that the response code is ErrorInvalidPropertySet.
            taskDelegateState = TaskDelegateStateType.Accepted;
            subject = Common.GenerateResourceName(this.Site, "This is a task");
            this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, taskDelegateState, true));
            Site.Assert.AreNotEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should not be success!", null);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, (ResponseCodeType)this.ResponseCode[0], "This create response status information should be ErrorInvalidPropertySet!", null);
            #endregion

            #region Client calls CreateItem to create a task with the DelegationState element is set to the NoMatch taskDelegateState.
            // Set the taskDelegateState of task item to TaskDelegateStateType.NoMatch. Except that the response code is ErrorInvalidPropertySet.
            taskDelegateState = TaskDelegateStateType.NoMatch;
            subject = Common.GenerateResourceName(this.Site, "This is a task");
            this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, taskDelegateState, true));
            Site.Assert.AreNotEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should not be success!", null);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, (ResponseCodeType)this.ResponseCode[0], "This create response status information should be ErrorInvalidPropertySet!", null);
            #endregion

            #region Client calls CreateItem to create a task with the DelegationState element is set to the Owned taskDelegateState.
            // Set the taskDelegateState of task item to TaskDelegateStateType.Owned. Except that the response code is ErrorInvalidPropertySet.
            taskDelegateState = TaskDelegateStateType.Owned;
            subject = Common.GenerateResourceName(this.Site, "This is a task");
            this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, taskDelegateState, true));
            Site.Assert.AreNotEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should not be success!", null);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, (ResponseCodeType)this.ResponseCode[0], "This create response status information should be ErrorInvalidPropertySet!", null);
            #endregion

            #region Client calls CreateItem to create a task with the DelegationState element is set to the OwnNew taskDelegateState.
            // Set the taskDelegateState of task item to TaskDelegateStateType.OwnNew. Except that the response code is ErrorInvalidPropertySet.
            taskDelegateState = TaskDelegateStateType.OwnNew;
            subject = Common.GenerateResourceName(this.Site, "This is a task");
            this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, taskDelegateState, true));
            Site.Assert.AreNotEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should not be success!", null);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, (ResponseCodeType)this.ResponseCode[0], "This create response status information should be ErrorInvalidPropertySet!", null);
            #endregion

            #region Client calls CreateItem to create a task with the DelegationState element is set to the Max taskDelegateState.
            // Set the taskDelegateState of task item to TaskDelegateStateType.Max. Except that the response code is ErrorInvalidPropertySet.
            taskDelegateState = TaskDelegateStateType.Max;
            subject = Common.GenerateResourceName(this.Site, "This is a task");
            this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, taskDelegateState, true));
            Site.Assert.AreNotEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should not be success!", null);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, (ResponseCodeType)this.ResponseCode[0], "This create response status information should be ErrorInvalidPropertySet!", null);
            #endregion

            #region Verify TaskDelegateStateType simple type is never set.
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R85");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R85
            // Create task item without setting the TaskDelegateStateType value success, but failed when creating the task item setting any value of TaskDelegateStateType.
            // So this requirement can be verified that this enumeration is never set via directly method.
            Site.CaptureRequirement(
                85,
                @"[In Simple Types] This enumeration [TaskDelegateStateType] is never set.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R88");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R88
            Site.CaptureRequirement(
                88,
                @"[In t:TaskDelegateStateType Simple Type] The values [Accepted, Declined, Max, NoMatch, Owned and OwnNew] for this simple type [t:TaskDelegateStateType Simple Type] are never set.");

            #endregion

            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is used to validate the TaskStatusType server behavior related requirements.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S01_TC12_VerifyTaskStatusType()
        {
            #region Client calls CreateItem to create a task item which sets the task Status to NotStarted.

            TaskStatusType taskStatus = TaskStatusType.NotStarted;

            // Create a task and save the item id.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            ItemIdType[] createItemIdsFirst = this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, taskStatus));
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemIdFirst = createItemIdsFirst[0];
            #endregion

            #region Client calls GetItem to get the task item created in above step.
            TaskType[] taskItems = this.GetTasks(createItemIdFirst);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType taskItem = taskItems[0];
            #endregion

            #region Verify the related requirements about NotStarted TaskStatusType type.
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R105");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R105
            Site.CaptureRequirementIfAreEqual<TaskStatusType>(
                TaskStatusType.NotStarted,
                taskItem.Status,
                105,
                @"[In t:TaskStatusType Simple Type] NotStarted: Specifies that the task is not started.");

            #endregion

            #region Client calls CreateItem to create a task item which sets the task Status to Completed.

            taskStatus = TaskStatusType.Completed;

            // Create a task and save the item id.
            subject = Common.GenerateResourceName(this.Site, "This is a task");
            ItemIdType[] createItemIdsSecond = this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, taskStatus));
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemIdSecond = createItemIdsSecond[0];
            #endregion

            #region Client calls GetItem to get the task item created in above step.
            taskItems = this.GetTasks(createItemIdSecond);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            taskItem = taskItems[0];
            #endregion

            #region Verify the related requirements about Completed TaskStatusType type.
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R102");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R102
            Site.CaptureRequirementIfAreEqual<TaskStatusType>(
                TaskStatusType.Completed,
                taskItem.Status,
                102,
                @"[In t:TaskStatusType Simple Type] Completed: Specifies that the task is completed.");

            #endregion

            #region Client calls CreateItem to create a task item which sets the task Status to InProgress.

            taskStatus = TaskStatusType.InProgress;

            // Create a task and save the item id.
            subject = Common.GenerateResourceName(this.Site, "This is a task");
            ItemIdType[] createItemIdsThird = this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, taskStatus));
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemIdThird = createItemIdsThird[0];
            #endregion

            #region Client calls GetItem to get the task item created in above step.
            taskItems = this.GetTasks(createItemIdThird);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            taskItem = taskItems[0];
            #endregion

            #region Verify the related requirements about InProgress TaskStatusType type.
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R104");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R104
            Site.CaptureRequirementIfAreEqual<TaskStatusType>(
                TaskStatusType.InProgress,
                taskItem.Status,
                104,
                @"[In t:TaskStatusType Simple Type]  InProgress: Specifies that the task is in progress.");

            #endregion

            #region Client calls CreateItem to create a task item which sets the task Status to Deferred.

            taskStatus = TaskStatusType.Deferred;

            // Create a task and save the item id.
            subject = Common.GenerateResourceName(this.Site, "This is a task");
            ItemIdType[] createItemIdsForth = this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, taskStatus));
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemIdForth = createItemIdsForth[0];
            #endregion

            #region Client calls GetItem to get the task item created in above step.
            taskItems = this.GetTasks(createItemIdForth);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            taskItem = taskItems[0];
            #endregion

            #region Verify the related requirements about Deferred TaskStatusType type.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R103");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R103
            Site.CaptureRequirementIfAreEqual<TaskStatusType>(
                TaskStatusType.Deferred,
                taskItem.Status,
                103,
                @"[In t:TaskStatusType Simple Type] Deferred: Specifies that the task is deferred.");

            #endregion

            #region Client calls CreateItem to create a task item which sets the task Status to WaitingOnOthers.
            taskStatus = TaskStatusType.WaitingOnOthers;

            // Create a task and save the item id.
            subject = Common.GenerateResourceName(this.Site, "This is a task");
            ItemIdType[] createItemIdsFifth = this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, taskStatus));
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemIdFifth = createItemIdsFifth[0];

            #endregion

            #region Client calls GetItem to get the task item created in above step.
            taskItems = this.GetTasks(createItemIdFifth);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            taskItem = taskItems[0];
            #endregion

            #region Verify the related requirements about WaitingOnOthers TaskStatusType type.
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R106");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R106
            Site.CaptureRequirementIfAreEqual<TaskStatusType>(
                TaskStatusType.WaitingOnOthers,
                taskItem.Status,
                106,
                @"[In t:TaskStatusType Simple Type] WaitingOnOthers: Specifies that the task is waiting on other tasks.");

            #endregion

            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(createItemIdFirst, createItemIdSecond, createItemIdThird, createItemIdForth, createItemIdFifth);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S01_TC13_VerifyTaskPercentComplete()
        {
            #region Client calls CreateItem to create a task item which sets the task Status to NotStarted.

            // Create a task and save the item id.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            TaskType taskNew = TestSuiteHelper.DefineTaskItem(subject, null);
            taskNew.CompleteDate = DateTime.UtcNow.Date;
            taskNew.CompleteDateSpecified = true;
            taskNew.Status = TaskStatusType.NotStarted;
            taskNew.StatusSpecified = true;
            taskNew.PercentComplete = 100;
            taskNew.PercentCompleteSpecified = true;
            ItemIdType[] createItemIdsFirst = this.CreateTasks(taskNew);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemIdFirst = createItemIdsFirst[0];
            #endregion

            #region Client calls GetItem to get the task item created in above step.
            TaskType[] taskItems = this.GetTasks(createItemIdFirst);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType taskItem = taskItems[0];
            #endregion

            //Verify the CompleteDateSpecified==false and PercentComplete=0.0 for capture R67001 and R67002
            if ((taskItem.CompleteDateSpecified == false) && (taskItem.PercentComplete == 0))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R67001");

                // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R67001
                Site.CaptureRequirement(
                    67001,
                    @"[In t:TaskType Complex Type] Setting CompleteDate has the same effect as setting PercentComplete to 100 or Status to Completed.");
                
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R67002");

                // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R67002
                Site.CaptureRequirement(
                    67002,
                    @"[In t:TaskType Complex Type] In a request that sets at least two of these properties, the last processed property will determine the value that is set for these elements.");

            }
            
            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(createItemIdFirst);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }



        #endregion

        #region Private methods
        /// <summary>
        /// Verify TaskRecurrencePatternTypes group.
        /// </summary>
        /// <param name="isValidTaskRecurrencePatternTypes">Whether a task recurrence pattern is valid TaskRecurrencePatternTypes.</param>
        private void VerifyTaskRecurrencePatternTypes(bool isValidTaskRecurrencePatternTypes)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R112");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R112
            // The TaskRecurrencePatternTypes Group contains ten elements and xs:choice is displayed in the schema.
            // So R112 can be verified when each element of ten is verified one by one in the schema. 
            Site.CaptureRequirementIfIsTrue(
                isValidTaskRecurrencePatternTypes,
                112,
                @"[In TaskRecurrencePatternTypes Group] The TaskRecurrencePatternTypes group specifies recurrence information for recurring tasks.
                    <xs:group name=""TaskRecurrencePatternTypes"">
                      <xs:sequence>
                        <xs:choice>
                          <xs:element name=""RelativeYearlyRecurrence""
                            type=""t:RelativeYearlyRecurrencePatternType""/>
                          <xs:element name=""AbsoluteYearlyRecurrence""
                           type=""t:AbsoluteYearlyRecurrencePatternType""/>
                          <xs:element name=""RelativeMonthlyRecurrence""
                        type=""t:RelativeMonthlyRecurrencePatternType""/>
                          <xs:element name=""AbsoluteMonthlyRecurrence""
                      type=""t:AbsoluteMonthlyRecurrencePatternType""/>
                          <xs:element name=""WeeklyRecurrence""
                            type=""t:WeeklyRecurrencePatternType""/>
                          <xs:element name=""DailyRecurrence""
                            type=""t:DailyRecurrencePatternType""/>
                          <xs:element name=""DailyRegeneration""
                            type=""t:DailyRegeneratingPatternType""/>
                          <xs:element name=""WeeklyRegeneration""
                            type=""t:WeeklyRegeneratingPatternType""/>
                          <xs:element name=""MonthlyRegeneration""
                            type=""t:MonthlyRegeneratingPatternType""/>
                          <xs:element name=""YearlyRegeneration""
                            type=""t:YearlyRegeneratingPatternType""/>
                        </xs:choice>
                      </xs:sequence>
                    </xs:group>");
        }

        /// <summary>
        /// Verify the IntervalRecurrencePatternBaseType.
        /// </summary>
        /// <param name="sentTaskItem">The created task item.</param>
        /// <param name="retrievedTaskItemDaily">The retrieved task item.</param>
        private void VerifyIntervalRecurrencePatternBaseType(TaskType sentTaskItem, TaskType retrievedTaskItemDaily)
        {
            IntervalRecurrencePatternBaseType sentDailyRecur = sentTaskItem.Recurrence.Item as IntervalRecurrencePatternBaseType;
            IntervalRecurrencePatternBaseType retrievedDailyRecur = retrievedTaskItemDaily.Recurrence.Item as IntervalRecurrencePatternBaseType;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1179");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1179
            this.Site.CaptureRequirementIfAreEqual<int>(
                sentDailyRecur.Interval,
                retrievedDailyRecur.Interval,
                "MS-OXWSCDATA",
                1179,
                @"[In t:IntervalRecurrencePatternBaseType Complex Type] The element ""Interval"" with type ""xs:int"" specifies the interval between two consecutive recurring items.");
        }

        #endregion
    }
}