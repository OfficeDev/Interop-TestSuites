//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSTASK
{
    using System;
    using System.Linq;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The delegate used to generate RecurrencePatternBaseType.
    /// </summary>
    /// <returns>The generated RecurrencePatternBaseType.</returns>
    public delegate RecurrencePatternBaseType GenerateRecurPattern();

    /// <summary>
    /// The delegate used to generate RecurrenceRangeBaseType.
    /// </summary>
    /// <returns>The generated RecurrenceRangeBaseType.</returns>
    public delegate RecurrenceRangeBaseType GenerateRecurRange();

    /// <summary>
    /// A class contains all helper methods used in test cases.
    /// </summary>
    public static class TestSuiteHelper
    {
        /// <summary>
        /// Compare two string array to verify whether their value are equal.
        /// </summary>
        /// <param name="arrayFirst">The first string array</param>
        /// <param name="arraySecond">The second string array</param>
        /// <returns>Returning true means two string array value are equal. Otherwise, return false.</returns>
        public static bool CompareStringArray(string[] arrayFirst, string[] arraySecond)
        {
            bool isEqual = false;
            if (arrayFirst == null || arraySecond == null)
            {
                return isEqual;
            }
            else
            {
                isEqual = arrayFirst.SequenceEqual(arraySecond);
            }

            return isEqual;
        }

        /// <summary>
        /// Define the task item via setting different kind of taskRecurrence value.
        /// </summary>
        /// <param name="subject">The subject of the task.</param>
        /// <param name="taskRecurrence">Recurrence element value of the task.</param>
        /// <returns>The task object.</returns>
        public static TaskType DefineTaskItem(string subject, TaskRecurrenceType taskRecurrence)
        {
            #region Define Task item

            TaskType taskItem = new TaskType
            {
                // The subject of a task.
                Subject = subject,

                // The actual amount of time that is spent on a task. 
                ActualWork = 5,
                ActualWorkSpecified = true,

                // The total amount of time that is associated with a task.
                TotalWork = 10,
                TotalWorkSpecified = true,

                // The billing information for a task.
                BillingInformation = "Discount: 10 dollars",

                // The collection of companies that are associated with a task.
                Companies = new string[] { "CompanyFirst", "CompanySecond" },

                // The collection of contacts that are associated with a task.
                Contacts = new string[] { "Alice", "Bob" },

                // The start date of a task. The start date cannot occur after due date.
                StartDate = DateTime.Parse("2011-8-4"),
                StartDateSpecified = true,

                // The due date of a task. The due date cannot occur before start date.
                DueDate = DateTime.Parse("2011-8-10"),
                DueDateSpecified = true,

                // The mileage for a task.
                Mileage = "15 km.",

                // The completion percentage of a task.
                PercentComplete = 50,
                PercentCompleteSpecified = true
            };

            // The recurrence of a task.
            if (taskRecurrence != null)
            {
                taskItem.Recurrence = taskRecurrence;
            }
            #endregion

            return taskItem;
        }

        /// <summary>
        /// Define the task item via setting taskDelegateState value.
        /// </summary>
        /// <param name="subject">The subject of the task.</param>
        /// <param name="taskDelegateState">DelegationState element value of the task.</param>
        /// <param name="isSetting">The flag to indicate whether the taskDelegateState is setting or not, true means setting.</param>
        /// <returns>The task object.</returns>
        public static TaskType DefineTaskItem(string subject, TaskDelegateStateType taskDelegateState, bool isSetting)
        {
            TaskType taskNew = DefineTaskItem(subject, null);
            if (isSetting)
            {
                taskNew.DelegationStateSpecified = true;
                taskNew.DelegationState = taskDelegateState;
            }

            return taskNew;
        }

        /// <summary>
        /// Define the task item via setting taskStatus value.
        /// </summary>
        /// <param name="subject">The subject of the task.</param>
        /// <param name="taskStatus">Status element value of the task.</param>
        /// <returns>The task object.</returns>
        public static TaskType DefineTaskItem(string subject, TaskStatusType taskStatus)
        {
            TaskType taskNew = DefineTaskItem(subject, null);

            if (taskStatus == TaskStatusType.Completed)
            {
                taskNew.CompleteDate = DateTime.UtcNow.Date;
                taskNew.CompleteDateSpecified = true;
            }

            taskNew.Status = taskStatus;
            taskNew.StatusSpecified = true;
            return taskNew;
        }

        /// <summary>
        /// Define the task item with index.
        /// </summary>
        /// <param name="subject">The subject of the task</param>>
        /// <returns>The task object.</returns>
        public static TaskType DefineTaskItem(string subject)
        {
            TaskType taskNew = DefineTaskItem(subject, null);
            return taskNew;
        }

        /// <summary>
        /// Generate CreateItemRequest.
        /// </summary>
        /// <param name="taskItems">The task item will be created.</param>
        /// <returns>Generated CreateItemRequest.</returns>
        public static CreateItemType GenerateCreateItemRequest(params TaskType[] taskItems)
        {
            return new CreateItemType
            {
                Items = new NonEmptyArrayOfAllItemsType { Items = taskItems }
            };
        }

        /// <summary>
        /// Generate GetItemRequest.
        /// </summary>
        /// <param name="createItemIds">The created item id.</param>
        /// <returns>Generated GetItemRequest.</returns>
        public static GetItemType GenerateGetItemRequest(params ItemIdType[] createItemIds)
        {
            // Configure the ItemIds and ItemShape parameters for GetItem request.
            return new GetItemType
            {
                ItemIds = createItemIds,

                ItemShape = new ItemResponseShapeType()
                {
                    BaseShape = DefaultShapeNamesType.AllProperties
                }
            };
        }

        /// <summary>
        /// Generate UpdateItemRequest.
        /// </summary>
        /// <param name="createItemIds">The created item id.</param>
        /// <returns>Generated GetItemRequest.</returns>
        public static UpdateItemType GenerateUpdateItemRequest(params ItemIdType[] createItemIds)
        {
            // Specify needed to update value for the task item.
            TaskType taskUpdate = new TaskType
            {
                Companies = new string[] { "Company3", "Company4" }
            };

            // Define the ItemChangeType element for updating the task item's companies.
            PathToUnindexedFieldType pathTo = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.taskCompanies
            };

            SetItemFieldType setItemField = new SetItemFieldType()
            {
                Item = pathTo,
                Item1 = taskUpdate
            };

            ItemChangeType[] itemChanges = new ItemChangeType[createItemIds.Length];
            for (int i = 0; i < createItemIds.Length; i++)
            {
                ItemChangeType itemChange = new ItemChangeType()
                {
                    Item = createItemIds[i],
                    Updates = new ItemChangeDescriptionType[] { setItemField }
                };
                itemChanges[i] = itemChange;
            }

            // Return the UpdateItemType request to update the task item.
            return new UpdateItemType()
            {
                ItemChanges = itemChanges,
                ConflictResolution = ConflictResolutionType.AlwaysOverwrite
            };
        }

        /// <summary>
        /// Generate CopyItemRequest.
        /// </summary>
        /// <param name="createItemIds">The created item id.</param>
        /// <returns>Generated CopyItemRequest.</returns>
        public static CopyItemType GenerateCopyItemRequest(params ItemIdType[] createItemIds)
        {
            return new CopyItemType
            {
                // Configure ItemIds.
                ItemIds = createItemIds,

                // Configure folder id.
                ToFolderId = new TargetFolderIdType
                {
                    Item = new DistinguishedFolderIdType { Id = DistinguishedFolderIdNameType.tasks }
                }
            };
        }

        /// <summary>
        /// Generate MoveItemRequest.
        /// </summary>
        /// <param name="createItemIds">The created item id.</param>
        /// <returns>Generated CopyItemRequest.</returns>
        public static MoveItemType GenerateMoveItemRequest(params ItemIdType[] createItemIds)
        {
            return new MoveItemType
            {
                // Configure ItemIds.
                ItemIds = createItemIds,

                // Configure folder id.
                ToFolderId = new TargetFolderIdType
                {
                    Item = new DistinguishedFolderIdType { Id = DistinguishedFolderIdNameType.deleteditems }
                }
            };
        }

        /// <summary>
        /// Generate DeleteItemRequest.
        /// </summary>
        /// <param name="createItemIds">Created item id.</param>
        /// <returns>Generated DeleteItemRequest.</returns>
        public static DeleteItemType GenerateDeleteItemRequest(params ItemIdType[] createItemIds)
        {
            return new DeleteItemType
            {
                // Configure the DeleteItem request.
                ItemIds = createItemIds,

                // Configure the DeleteItem request.
                DeleteType = DisposalType.HardDelete,
                AffectedTaskOccurrencesSpecified = true,
                AffectedTaskOccurrences = AffectedTaskOccurrencesType.AllOccurrences
            };
        }

        /// <summary>
        /// Generate the task recurrence used by TaskType.
        /// </summary>
        /// <param name="generateRecurPatternMethod">The method to generate recurrence pattern type.</param>
        /// <param name="generateRecurRangeMethod">The method to generate recurrence range type.</param>
        /// <returns>The generated task recurrence.</returns>
        public static TaskRecurrenceType GenerateTaskRecurrence(GenerateRecurPattern generateRecurPatternMethod, GenerateRecurRange generateRecurRangeMethod)
        {
            return new TaskRecurrenceType
            {
                Item = generateRecurPatternMethod(),
                Item1 = generateRecurRangeMethod()
            };
        }

        /// <summary>
        /// Generate the daily regenerating pattern.
        /// </summary>
        /// <returns>The generated daily regenerating pattern.</returns>
        public static DailyRegeneratingPatternType GenerateDailyRegeneratingPattern()
        {
            return new DailyRegeneratingPatternType
            {
                // Set the interval of dailyRegeneratingPatternType as any integer value. Set it to 3 here.
                Interval = 3
            };
        }

        /// <summary>
        /// Generate the daily recurrence pattern.
        /// </summary>
        /// <returns>The generated daily recurrence pattern.</returns>
        public static DailyRecurrencePatternType GenerateDailyRecurrencePattern()
        {
            return new DailyRecurrencePatternType
            {
                // Set the interval of dailyRecurrencePatternType as any integer value. Set it to 3 here.
                Interval = 3
            };
        }

        /// <summary>
        /// Generate the weekly regenerating pattern.
        /// </summary>
        /// <returns>The generated weekly regenerating pattern.</returns>
        public static WeeklyRegeneratingPatternType GenerateWeeklyRegeneratingPattern()
        {
            return new WeeklyRegeneratingPatternType
            {
                // Set the interval of weeklyRegeneratingPatternType as any integer value. Set it to 3 here.
                Interval = 3
            };
        }

        /// <summary>
        /// Generate the weekly recurrence pattern.
        /// </summary>
        /// <returns>The generated weekly recurrence pattern.</returns>
        public static WeeklyRecurrencePatternType GenerateWeeklyRecurrencePattern()
        {
            return new WeeklyRecurrencePatternType
            {
                // Set DaysOfWeek of WeeklyRecurrencePatternType to any valid DaysOfWeekType string value. Set it to "Friday" here.
                DaysOfWeek = "Friday",
                Interval = 5
            };
        }

        /// <summary>
        /// Generate the monthly regenerating pattern.
        /// </summary>
        /// <returns>The generated monthly regenerating pattern.</returns>
        public static MonthlyRegeneratingPatternType GenerateMonthlyRegeneratingPattern()
        {
            return new MonthlyRegeneratingPatternType
            {
                // Set the interval of monthlyRegeneratingPatternType as any integer value. Set it to 3 here.
                Interval = 3
            };
        }

        /// <summary>
        /// Generate the relative monthly recurrence pattern.
        /// </summary>
        /// <returns>The generated relative monthly recurrence pattern.</returns>
        public static RelativeMonthlyRecurrencePatternType GenerateRelativeMonthlyRecurrencePattern()
        {
            return new RelativeMonthlyRecurrencePatternType
            {
                // Set the DayOfWeekIndex of RelativeMonthlyRecurrencePatternType to any DayOfWeekIndexType value. Set it to DayOfWeekIndexType.First here.
                DayOfWeekIndex = DayOfWeekIndexType.First,

                // Set the DaysOfWeek of RelativeMonthlyRecurrencePatternType to any DayOfWeekType type value. Set it to DayOfWeekType.Monday here.
                DaysOfWeek = DayOfWeekType.Monday,

                // Set the Interval of RelativeMonthlyRecurrencePatternType to any integer value. Set it to 5 here.
                Interval = 5
            };
        }

        /// <summary>
        /// Generate the absolute monthly recurrence pattern.
        /// </summary>
        /// <returns>The generated absolute monthly recurrence pattern.</returns>
        public static AbsoluteMonthlyRecurrencePatternType GenerateAbsoluteMonthlyRecurrencePattern()
        {
            return new AbsoluteMonthlyRecurrencePatternType
            {
                // Set DayOfMonth of AbsoluteMonthlyRecurrencePatternType to any integer value. Here set to 5.
                DayOfMonth = 5,

                // Set Interval of AbsoluteMonthlyRecurrencePatternType to any integer value. Here set to 4.
                Interval = 4
            };
        }

        /// <summary>
        /// Generate the yearly regenerating pattern.
        /// </summary>
        /// <returns>The generated yearly regenerating pattern.</returns>
        public static YearlyRegeneratingPatternType GenerateYearlyRegeneratingPattern()
        {
            return new YearlyRegeneratingPatternType
            {
                // Set the interval of yearlyRegeneratingPatternType as any integer value. Set it to 3 here.
                Interval = 3
            };
        }

        /// <summary>
        /// Generate the relative yearly recurrence pattern.
        /// </summary>
        /// <returns>The generated relative yearly recurrence pattern.</returns>
        public static RelativeYearlyRecurrencePatternType GenerateRelativeYearlyRecurrencePattern()
        {
            return new RelativeYearlyRecurrencePatternType
            {
                // Set the DayOfWeekIndex of RelativeYearlyRecurrencePatternType as any DayOfWeekIndexType type value. Here set to DayOfWeekIndexType.First.
                DayOfWeekIndex = DayOfWeekIndexType.First,

                // Set the DaysOfWeek of RelativeYearlyRecurrencePatternType as any valid string value. Here set to "Friday".
                DaysOfWeek = "Friday",

                // Set the Month of RelativeYearlyRecurrencePatternType as any valid MonthNamesType value. Here set to MonthNamesType.August.
                Month = MonthNamesType.August
            };
        }

        /// <summary>
        /// Generate the absolute yearly recurrence pattern.
        /// </summary>
        /// <returns>The generated absolute yearly recurrence pattern.</returns>
        public static AbsoluteYearlyRecurrencePatternType GenerateAbsoluteYearlyRecurrencePattern()
        {
            return new AbsoluteYearlyRecurrencePatternType
            {
                // Set the DayOfMonth of AbsoluteYearlyRecurrencePatternType as any integer value. Here set to 2. 
                DayOfMonth = 2,

                // Set the Month of AbsoluteYearlyRecurrencePatternType as any MonthNamesType value. Here set to MonthNamesType.August.
                Month = MonthNamesType.August
            };
        }

        /// <summary>
        /// Generate the numbered recurrence range.
        /// </summary>
        /// <returns>Generated the numbered recurrence range.</returns>
        public static NumberedRecurrenceRangeType GenerateNumberedRecurrenceRange()
        {
            return new NumberedRecurrenceRangeType
            {
                // Set the NumberOfOccurrences of NumberedRecurrenceRange as any integer value. Set it to 4 here.
                NumberOfOccurrences = 4,

                // Set the StartDate of NumberedRecurrenceRange as any valid date time. Set it to now here.
                StartDate = DateTime.Now
            };
        }

        /// <summary>
        /// Generate the no end recurrence range.
        /// </summary>
        /// <returns>Generated the no end recurrence range.</returns>
        public static NoEndRecurrenceRangeType GenerateNoEndRecurrenceRange()
        {
            return new NoEndRecurrenceRangeType
            {
                // Set the StartDate of NoEndRecurrenceRange as any valid date time. Set it to now here.
                StartDate = DateTime.Now
            };
        }

        /// <summary>
        /// Generate the end date recurrence range.
        /// </summary>
        /// <returns>Generated the end date recurrence range.</returns>
        public static EndDateRecurrenceRangeType GenerateEndDateRecurrenceRange()
        {
            return new EndDateRecurrenceRangeType
            {
                // Set the StartDate of EndDateRecurrenceRange as any valid date time. Set it to now here.
                StartDate = DateTime.Now,

                // Set the end date of EndDateRecurrenceRange as any time after start time. Set it to 2 days after here.
                EndDate = DateTime.Now.AddDays(2)
            };
        }
    }
}