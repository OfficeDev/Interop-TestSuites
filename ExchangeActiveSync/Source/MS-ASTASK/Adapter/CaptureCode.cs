namespace Microsoft.Protocols.TestSuites.MS_ASTASK
{
    using System.Collections.Generic;
    using Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using ItemOperationsStore = Microsoft.Protocols.TestSuites.Common.DataStructures.ItemOperationsStore;
    using SearchStore = Microsoft.Protocols.TestSuites.Common.DataStructures.SearchStore;
    using SyncStore = Microsoft.Protocols.TestSuites.Common.DataStructures.SyncStore;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-ASTASK.
    /// </summary>
    public partial class MS_ASTASKAdapter
    {
        /// <summary>
        /// This method is used to verify transport related requirement.
        /// </summary>
        private void VerifyTransport()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R2");

            // Verify MS-ASTASK requirement: MS-ASTASK_R2
            // ActiveSyncClient encodes XML request into WBXML and decodes WBXML to XML response, capture it directly if server responds successfully.
            Site.CaptureRequirement(
                2,
                @"[In Transport] The XML markup that constitutes the request body or the response body that is transmitted between the client and the server uses Wireless Application Protocol (WAP) Binary XML (WBXML), as specified in [MS-ASWBXML].");
        }

        /// <summary>
        /// This method is used to verify the message syntax related requirement.
        /// </summary>
        private void VerifyMessageSyntax()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R3");

            // Verify MS-ASTASK requirement: MS-ASTASK_R3
            // If the schema is validated, this requirement will be captured.
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                3,
                @"[In Message Syntax] The markup that is used by this protocol MUST be well-formed XML, as specified in [XML10/5].");
        }

        /// <summary>
        /// This method is used to verify the Sync response related requirements.
        /// </summary>
        /// <param name="syncResponse">Specified the SyncStore result returned from the server.</param>
        private void VerifySyncCommandResponse(SyncStore syncResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema should be validated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R349");

            // Verify MS-ASTASK requirement: MS-ASTASK_R349
            // If Sync response exists, this requirement will be captured.
            Site.CaptureRequirementIfIsNotNull(
                syncResponse,
                349,
                @"[In Synchronizing Task Data Between Client and Server][If the client sends a Sync command request to the server] The server responds with a Sync command response ([MS-ASCMD] section 2.2.2.19).");

            if (null != syncResponse.AddElements)
            {
                for (int i = syncResponse.AddElements.Count - 1; i >= 0; i--)
                {
                    Task task = syncResponse.AddElements[i].Task;

                    if (null != task)
                    {
                        if (null != task.Body)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R74");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                74,
                                @"[In Body (AirSyncBase Namespace)] The airsyncbase:Body element is a container ([MS-ASDTYPE] section 2.2) element that specifies details about the body of a task item.");

                            this.VerifyContainerDataType();
                        }

                        if (null != task.Categories)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R106");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                106,
                                @"[In Categories] The Categories element is a container ([MS-ASDTYPE] section 2.2) element that specifies a collection of user-managed labels assigned to the task.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R108");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                108,
                                @"[In Categories] A command response has a maximum of one Categories element per command.");

                            if (null != task.Categories.Category)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R114");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    114,
                                    @"[In Category] The value of this element[Category] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                                if (Common.IsRequirementEnabled(402, Site))
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R402, the actual number of category is: {0}", task.Categories.Category.Length);

                                    // Verify MS-ASTASK requirement: MS-ASTASK_R402
                                    // If the number of categories is not more than 300, this requirement will be captured.
                                    Site.CaptureRequirementIfIsTrue(
                                        task.Categories.Category.Length <= 300,
                                        402,
                                        @"[In Appendix B: Product Behavior] A Categories element does contain no more than 300 Category child elements. (Exchange 2007 SP1 and above follow this behavior.)");
                                }

                                this.VerifyStringDataType();
                            }

                            this.VerifyContainerDataType();
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R116");

                        // Since schema is validated, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            116,
                            @"[In Complete] The Complete element is a required element that specifies whether the task has been completed.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R118");

                        // Since schema is validated, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            118,
                            @"[In Complete] The value of this element[Complete] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R119, the actual value of Complete is: &", task.Complete);

                        // Verify MS-ASTASK requirement: MS-ASTASK_R119
                        Site.CaptureRequirementIfIsTrue(
                            task.Complete == 0 || task.Complete == 1,
                            119,
                            @"[In Complete] The value of the Complete element MUST be one of the following values[The value is 0 or1].");

                        this.VerifyUnsignedByteDataType(task.Complete);

                        if (null != task.DateCompleted)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R122");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                122,
                                @"[In DateCompleted] The DateCompleted element specifies the date and time at which the task was completed.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R125");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                125,
                                @"[In DateCompleted] The value of this element [DateCompleted] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

                            this.VerifyDateTimeDataType();
                        }

                        if (null != task.DueDate)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R161");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                161,
                                @"[In DueDate] The value of this element [DueDate] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

                            this.VerifyDateTimeDataType();
                        }

                        if (null != task.Importance)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R182");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                182,
                                @"[In Importance] The value of this element [Importance] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7.");

                            if (Common.IsRequirementEnabled(636, Site))
                            {
                                Site.Log.Add(LogEntryKind.Debug, "The returned value of Importance is: {0}.", task.Importance);
                                if (task.Importance >= 0 && task.Importance <= 2)
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R636");

                                    // Capture R636 when the value is valid
                                    Site.CaptureRequirement(
                                        636,
                                        @"[In Appendix B: Product Behavior] The value of the Importance element is one of the following:[the value is between 0 to 2] (Exchange 2007 SP1 and above follow this behavior.)");
                                }
                            }

                            this.VerifyUnsignedByteDataType(task.Importance);
                        }

                        if (null != task.Recurrence)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R221");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                221,
                                @"[In Recurrence] The Recurrence element is a container ([MS-ASDTYPE] section 2.2) element that specifies when and how often the task recurs.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R223");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                223,
                                @"[In Recurrence] A command [request or] response has a maximum of one Recurrence element per command.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R225");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                225,
                                @"[In Recurence][The Recurrence element can have the following child elements:] Type (section 2.2.2.25): This element is required.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R226");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                226,
                                @"[In Recurence][The Recurrence element can have the following child elements:] Start (section 2.2.2.22): This element is required.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R77");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                77,
                                @"[In CalendarType] The CalendarType element is a child element of the Recurrence element (section 2.2.2.20) that specifies the calendar system used by the task recurrence.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R79");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                79,
                                @"[In CalendarType] The value of this element[CalendarType] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.8.");

                            string[] expectedCalendarTypeValues = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "14", "15", "20" };

                            Common.VerifyActualValues("CalendarType", expectedCalendarTypeValues, task.Recurrence.CalendarType.ToString(), Site);

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R84, the actual value of CalendarType is: &", task.Recurrence.CalendarType);

                            // Since Common.VerifyActualValues runs successfully, this requirement can be captured.
                            Site.CaptureRequirement(
                                84,
                                @"[In CalendarType] The value of the CalendarType element MUST be one of the values listed in the following table[The value is among 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 15 and 20].");

                            this.VerifyUnsignedByteDataType(task.Recurrence.CalendarType);

                            if (task.Recurrence.DayOfMonthSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R128");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    128,
                                    @"[In DayOfMonth] The value of this element[DayOfMonth] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R130");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    130,
                                    @"[In DayOfMonth] A command [request or] response has a maximum of one DayOfMonth element per Recurrence element.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R525, the actual value of DayOfMonth is: &", task.Recurrence.DayOfMonth);

                                // Verify MS-ASTASK requirement: MS-ASTASK_R525
                                Site.CaptureRequirementIfIsTrue(
                                    task.Recurrence.DayOfMonth >= 1 && task.Recurrence.DayOfMonth <= 31,
                                    525,
                                    @"[In DayOfMonth] The value of the DayOfMonth element MUST be between 1 and 31.");

                                this.VerifyUnsignedByteDataType(task.Recurrence.DayOfMonth);
                            }

                            if (task.Recurrence.DayOfWeekSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R139");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    139,
                                    @"[In DayOfWeek] A command response has a maximum of one DayOfWeek element per Recurrence element.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R140");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    140,
                                    @"[In DayOfWeek] The value of this element [DayOfWeek] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R141, the actual value of DayOfWeek is: &", task.Recurrence.DayOfWeek);

                                // Verify MS-ASTASK requirement: MS-ASTASK_R141
                                // After calculation, the range of the numbers possible is between 1 to 254. 
                                Site.CaptureRequirementIfIsTrue(
                                    task.Recurrence.DayOfWeek >= 1 && task.Recurrence.DayOfWeek <= 254,
                                    141,
                                    @"[In DayOfWeek] The value of the DayOfWeek element MUST be either one of the following values[the value is 1,2,4,8,16,32,64 or 127], or the sum of more than one of the following values (in which case this task recurs on more than one day).");

                                this.VerifyUnsignedByteDataType(task.Recurrence.DayOfWeek);
                            }

                            if (task.Recurrence.DeadOccurSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R154");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    154,
                                    @"[In DeadOccur] The value of this element [DeadOccur] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R155");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    155,
                                    @"[In DeadOccur] A command [request or] response has a maximum of one DeadOccur child element per Recurrence element.");

                                this.VerifyUnsignedByteDataType(task.Recurrence.DeadOccur);
                            }

                            if (task.Recurrence.FirstDayOfWeekSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R162");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    162,
                                    @"[In FirstDayOfWeek] The FirstDayOfWeek element is a child element of the Recurrence element (section 2.2.2.18) that specifies which day is considered the first day of the calendar week for this recurrence.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R164");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    164,
                                    @"[In FirstDayOfWeek] The value of this element[FirstDayOfWeek] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R166");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    166,
                                    @"[In FirstDayOfWeek] A command [request or] response has a maximum of one FirstDayOfWeek child element per Recurrence element.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R170, the actual value of FirstDayOfWeek is: &", task.Recurrence.FirstDayOfWeek);

                                // Verify MS-ASTASK requirement: MS-ASTASK_R170
                                Site.CaptureRequirementIfIsTrue(
                                    task.Recurrence.FirstDayOfWeek <= 6,
                                    170,
                                    @"[In FirstDayOfWeek] The value of the FirstDayOfWeek element MUST be one of the values defined in the following table[the value is between 0 to 6].");

                                this.VerifyUnsignedByteDataType(task.Recurrence.FirstDayOfWeek);
                            }

                            if (task.Recurrence.IntervalSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R190");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    190,
                                    @"[In Interval] A command [request or] response has a maximum of one Interval element per Recurrence element.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R191");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    191,
                                    @"[In Interval] The value of the Interval element is an unsignedShort, as specified in [XMLSCHEMA2/2] section 3.3.23. ");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R192, the actual value of Interval is: &", task.Recurrence.Interval);

                                // Verify MS-ASTASK requirement: MS-ASTASK_R192
                                Site.CaptureRequirementIfIsTrue(
                                    task.Recurrence.Interval <= 999,
                                    192,
                                    @"[In Interval] The maximum value of this element[Interval] is 999.");
                            }

                            if (task.Recurrence.IsLeapMonthSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R195");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    195,
                                    @"[In IsLeapMonth] The value of this element[IsLeapMonth] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R197");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    197,
                                    @"[In IsLeapMonth] A command [request or] response has a maximum of one IsLeapMonth child element per Recurrence element.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R201, the actual value of IsLeapMonth is: &", task.Recurrence.IsLeapMonth);

                                // Verify MS-ASTASK requirement: MS-ASTASK_R201
                                Site.CaptureRequirementIfIsTrue(
                                    task.Recurrence.IsLeapMonth == 0 || task.Recurrence.IsLeapMonth == 1,
                                    201,
                                    @"[In IsLeapMonth] The value of the IsLeapMonth element MUST be one of the following values[the value is 0 or 1].");

                                this.VerifyUnsignedByteDataType(task.Recurrence.IsLeapMonth);
                            }

                            if (task.Recurrence.MonthOfYearSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R207");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    207,
                                    @"[In MonthOfYear] The value of this element[MonthOfYear] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R209");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    209,
                                    @"[In MonthOfYear] A command [request or] response has a maximum of one MonthofYear child element per Recurrence element.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R210, the actual value of MonthOfYear is: &", task.Recurrence.MonthOfYear);

                                // Verify MS-ASTASK requirement: MS-ASTASK_R210
                                Site.CaptureRequirementIfIsTrue(
                                    task.Recurrence.MonthOfYear >= 1 && task.Recurrence.MonthOfYear <= 12,
                                    210,
                                    @"[In MonthOfYear] The value of the MonthofYear element MUST be between 1 and 12.");

                                this.VerifyUnsignedByteDataType(task.Recurrence.MonthOfYear);
                            }

                            if (task.Recurrence.OccurrencesSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R213");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    213,
                                    @"[In Occurrences] The value of this element[Occurrences] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R214");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    214,
                                    @"[In Occurrences] A command [request or] response has a maximum of one Occurrences child element per Recurrence element.");

                                this.VerifyUnsignedByteDataType(task.Recurrence.Occurrences);
                            }

                            if (task.Recurrence.RegenerateSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R241");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    241,
                                    @"[In Regenerate] The value of this element [Regenerate] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R242");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    242,
                                    @"[In Regenerate] A command [request or] response has a maximum of one Regenerate child element per Recurrence element.");

                                this.VerifyUnsignedByteDataType(task.Recurrence.Regenerate);
                            }

                            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
                            {
                                if (Common.IsRequirementEnabled(633, Site))
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R633");

                                    // Verify MS-ASTASK requirement: MS-ASTASK_R633
                                    Site.CaptureRequirementIfIsFalse(
                                        task.Recurrence.Start != null,
                                        633,
                                        @"[In Appendix A: Product Behavior] <2> Section 2.2.2.25:  Microsoft Exchange Server 2010 Service Pack 1 (SP1), the initial release version of Exchange 2013, and Exchange 2016 Preview do not return the Start element when protocol version is 14.0.");
                                }
                            }
                            else
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R267");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    267,
                                    @"[In Start] A command [request or] response has a minimum of one Start child element per Recurrence element.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R268");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    268,
                                    @"[In Start] A command [request or] response has a maximum of one Start child element per Recurrence element.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R271");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    271,
                                    @"[In Start] The value of this element [Start] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");
                                this.VerifyDateTimeDataType();
                            }

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R285");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                285,
                                @"[In Type] The Type element is a required child element of the Recurrence element (section 2.2.2.18) that specifies the type of the recurrence item.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R287");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                287,
                                @"[In Type] A command [request or] response has a minimum of one Type child element per Recurrence element.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R288");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                288,
                                @"[In Type] A command [request or] response has a maximum of one Type child element per Recurrence element.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R291");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                291,
                                @"[In Type] The value of this element[Type] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7.");

                            string[] expectedTypeValues = { "0", "1", "2", "3", "5", "6" };

                            Common.VerifyActualValues("Type", expectedTypeValues, task.Recurrence.CalendarType.ToString(), Site);

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R292, the actual value of Type is: &", task.Recurrence.Type);

                            // Since Common.VerifyActualValues runs successfully, this requirement can be captured.
                            Site.CaptureRequirement(
                                292,
                                @"[In Type] The value of the Type element MUST be one of the following values[the value is between 0 to 6, except for 4].");

                            if (task.Recurrence.Type == 1)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R238");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    238,
                                    @"[In Recurence][The Recurrence element can have the following child elements:] FirstDayOfWeek (section 2.2.2.11): This element is required in server's responses for weekly recurrences.");
                            }

                            if (task.Recurrence.UntilSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R301");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    301,
                                    @"[In Until] A command [request or] response has a maximum of one Until child element per Recurrence element.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R305");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    305,
                                    @"[In Until] The value of this element[Until] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

                                this.VerifyDateTimeDataType();
                            }

                            if (task.Recurrence.WeekOfMonthSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R314");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    314,
                                    @"[In WeekOfMonth] The value of this element [WeekOfMonth] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R316");

                                // Since schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    316,
                                    @"[In WeekOfMonth] A command [request or] response has a maximum of one WeekOfMonth child element per Recurrence element.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R317, the actual value of WeekOfMonth is: &", task.Recurrence.WeekOfMonth);

                                // Verify MS-ASTASK requirement: MS-ASTASK_R317
                                Site.CaptureRequirementIfIsTrue(
                                    task.Recurrence.WeekOfMonth >= 1 && task.Recurrence.WeekOfMonth <= 5,
                                    317,
                                    @"[In WeekOfMonth] The value of the WeekOfMonth element MUST be between 1 and 5.");

                                this.VerifyUnsignedByteDataType(task.Recurrence.WeekOfMonth);
                            }

                            this.VerifyContainerDataType();
                        }

                        if (null != task.ReminderSet)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R248");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                248,
                                @"[In ReminderSet] The value of this element[ReminderSet] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R249, the actual value of ReminderSet is: &", task.ReminderSet);

                            // Verify MS-ASTASK requirement: MS-ASTASK_R249
                            Site.CaptureRequirementIfIsTrue(
                                task.ReminderSet == 0 || task.ReminderSet == 1,
                                249,
                                @"[In ReminderSet] The value of the ReminderSet element MUST be one of the following values[the value is 0 or 1].");

                            this.VerifyUnsignedByteDataType(task.ReminderSet);
                        }

                        if (null != task.ReminderTime)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R255");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                255,
                                @"[In ReminderTime] The value of this element [ReminderTime] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

                            this.VerifyDateTimeDataType();
                        }

                        if (null != task.Sensitivity)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R258");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                258,
                                @"[In Sensitivity] The value of this element[Sensitivity] is an unsignedbyte data type, as specified in [MS-ASDTYPE] section 2.7.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R259, the actual value of Sensitivity is: &", task.Sensitivity);

                            // Verify MS-ASTASK requirement: MS-ASTASK_R259
                            Site.CaptureRequirementIfIsTrue(
                                task.Sensitivity >= 0 && task.Sensitivity <= 3,
                                259,
                                @"[In Sensitivity] The value of the Sensitivity element MUST be one of the following values[the value is between 0 and 3].");

                            this.VerifyUnsignedByteDataType(task.Sensitivity);
                        }

                        if (null != task.StartDate)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R274");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                274,
                                @"[In StartDate] The value of this element[StartDate] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

                            this.VerifyDateTimeDataType();
                        }

                        if (null != task.Subject)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R277");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                277,
                                @"[In Subject] The value of this element [Subject] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                            this.VerifyStringDataType();
                        }

                        if (null != task.UtcDueDate)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R308");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                308,
                                @"[In UtcDueDate] The value of this element [UtcDueDate] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

                            this.VerifyDateTimeDataType();
                        }

                        if (null != task.UtcStartDate)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R311");

                            // Since schema is validated, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                311,
                                @"[In UtcStartDate] The value of this element [UtcStartDate] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

                            this.VerifyDateTimeDataType();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// This method is used to verify the ItemOperations response related requirement.
        /// </summary>
        /// <param name="itemOperationsResponse">Specified ItemOperationsStore result returned from the server.</param>
        private void VerifyItemOperationsResponse(ItemOperationsStore itemOperationsResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R353");

            // Verify MS-ASTASK requirement: MS-ASTASK_R353
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                itemOperationsResponse.Status,
                353,
                @"[In Requesting Details for Specific Tasks][If the client sends an ItemOperations command request to the server] The server responds with an ItemOperations command response ([MS-ASCMD] section 2.2.2.8).");
        }

        /// <summary>
        /// This method is used to verify the Search response related requirement.
        /// </summary>
        /// <param name="searchResponse">Specified SearchStore result returned from the server.</param>
        private void VerifySearchCommandResponse(SearchStore searchResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R351");

            // Verify MS-ASTASK requirement: MS-ASTASK_R351
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                searchResponse.Status,
                351,
                @"[In Searching for Task Data][If the client sends a search command request to the server] The server responds with a Search command response ([MS-ASCMD] section 2.2.2.14).");
        }

        /// <summary>
        /// This method is used to verify the container data type related requirements.
        /// </summary>
        private void VerifyContainerDataType()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R8");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R8
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                8,
                @"[In container Data Type] A container is an XML element that encloses other elements but has no value of its own.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R9");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R9
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                9,
                @"[In container Data Type] It [container] is a complex type with complex content, as specified in [XMLSCHEMA1/2] section 3.4.2.");
        }

        /// <summary>
        /// This method is used to verify the dateTime data type related requirements.
        /// </summary>
        private void VerifyDateTimeDataType()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R12");

            // Since schema is validated, this requirement can be captured directly.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                12,
                @"[In dateTime Data Type] It [dateTime] is declared as an element whose type attribute is set to ""dateTime"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R15");

            // Since schema is validated, this requirement can be captured directly.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                15,
                @"[In dateTime Data Type] All dates are given in Coordinated Universal Time (UTC) and are represented as a string in the following format.
YYYY-MM-DDTHH:MM:SS.MSSZ where
YYYY = Year (Gregorian calendar year)
MM = Month (01 - 12)
DD = Day (01 - 31)
HH = Number of complete hours since midnight (00 - 24)
MM = Number of complete minutes since start of hour (00 - 59)
SS = Number of seconds since start of minute (00 - 59)
MSS = Number of milliseconds");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R16");

            // Since schema is validated, this requirement can be captured directly.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                16,
                @"[In dateTime Data Type][in YYYY-MM-DDTHH:MM:SS.MSSZ ] The T serves as a separator, and the Z indicates that this time is in UTC.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R18");

            // Since schema is validated, this requirement can be captured directly.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                18,
                @"[In dateTime Data Type] Note: Dates and times in calendar items (as specified in [MS-ASCAL]) MUST NOT include punctuation separators.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R20");

            // ActiveSyncClient encoded dateTime data as inline strings, so if response is successfully returned this requirement can be verified.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R20
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                20,
                @"[In dateTime Data Type] Elements with a dateTime data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");
        }

        /// <summary>
        /// This method is used to verify the string data type related requirements.
        /// </summary>
        private void VerifyStringDataType()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R88");

            // Since schema is validated, this requirement can be captured directly.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                88,
                @"[In string Data Type] A string is a chunk of Unicode text.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R90");

            // Since schema is validated, this requirement can be captured directly.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                90,
                @"[In string Data Type] An element of this [string] type is declared as an element with a type attribute of ""string"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R91");

            // ActiveSyncClient encoded string data as inline strings, so if response is successfully returned this requirement can be covered.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R91
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                91,
                @"[In string Data Type] Elements with a string data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R94");

            // Since schema is validated, this requirement can be captured directly.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                94,
                @"[In string Data Type] Elements of these types [ActiveSync defines several conventions for strings that adhere to commonly used formats]are defined as string types in XML schemas.");
        }

        /// <summary>
        /// This method is used to verify the unsignedByte data type related requirements.
        /// </summary>
        /// <param name="byteValue">byte type value.</param>
        private void VerifyUnsignedByteDataType(byte? byteValue)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R123");

            // Since schema is validated, this requirement can be captured directly.
            Site.CaptureRequirementIfIsTrue(
                (byteValue >= 0) && (byteValue <= 255),
                "MS-ASDTYPE",
                123,
                @"[In unsignedByte Data Type] The unsignedByte data type is an integer value between 0 and 255, inclusive.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R125");

            // Since schema is validated, this requirement can be captured directly.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                125,
                @"[In unsignedByte Data Type] Elements of this type [unsignedByte type] are declared with an element whose type attribute is set to ""unsignedByte"".");
        }

        /// <summary>
        /// This method is used to verify MS-ASWBXML related requirements.
        /// </summary>
        private void VerifyWBXMLRequirements()
        {
            // Get WBXML decoded data.MS_ASWBXMLSyntheticImplementation.
            Dictionary<string, int> decodedData = this.activeSyncClient.GetMSASWBXMLImplementationInstance().DecodeDataCollection;

            if (null != decodedData)
            {
                // Find Code Page 9.
                foreach (KeyValuePair<string, int> decodeDataItem in decodedData)
                {
                    byte token;
                    string tagName = Common.GetTagName(decodeDataItem.Key, out token);
                    string codePageName = Common.GetCodePageName(decodeDataItem.Key);
                    int codePage = decodeDataItem.Value;
                    bool isValidCodePage = codePage >= 0 && codePage <= 24;
                    Site.Assert.IsTrue(isValidCodePage, "Code page value should between 0-24, the actual value is :{0}", codePage);

                    // Capture requirements.
                    if (9 == codePage)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R19");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R19
                        Site.CaptureRequirementIfAreEqual<string>(
                            "tasks",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            19,
                            @"[In Code Pages] [This algorithm supports] [Code page] 9 [that indicates] [XML namespace] Tasks");

                        switch (tagName)
                        {
                            case "Categories":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R282");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R282
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x08,
                                        token,
                                        "MS-ASWBXML",
                                        282,
                                        @"[In Code Page 9: Tasks] [Tag name] Categories [Token] 0x08 [supports protocol versions] All");

                                    break;
                                }

                            case "Category":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R283");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R283
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x09,
                                        token,
                                        "MS-ASWBXML",
                                        283,
                                        @"[In Code Page 9: Tasks] [Tag name] Category [Token] 0x09 [supports protocol versions] All");

                                    break;
                                }

                            case "Complete":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R284");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R284
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x0A,
                                        token,
                                        "MS-ASWBXML",
                                        284,
                                        @"[In Code Page 9: Tasks] [Tag name] Complete [Token] 0x0A [supports protocol versions] All");

                                    break;
                                }

                            case "DateCompleted":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R285");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R285
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x0B,
                                        token,
                                        "MS-ASWBXML",
                                        285,
                                        @"[In Code Page 9: Tasks] [Tag name] DateCompleted [Token] 0x0B [supports protocol versions] All");

                                    break;
                                }

                            case "DueDate":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R286");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R286
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x0C,
                                        token,
                                        "MS-ASWBXML",
                                        286,
                                        @"[In Code Page 9: Tasks] [Tag name] DueDate [Token] 0x0C [supports protocol versions] All");

                                    break;
                                }

                            case "UtcDueDate":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R287");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R287
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x0D,
                                        token,
                                        "MS-ASWBXML",
                                        287,
                                        @"[In Code Page 9: Tasks] [Tag name] UtcDueDate [Token] 0x0D [supports protocol versions] All");

                                    break;
                                }

                            case "Importance":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R288");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R288
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x0E,
                                        token,
                                        "MS-ASWBXML",
                                        288,
                                        @"[In Code Page 9: Tasks] [Tag name] Importance [Token] 0x0E [supports protocol versions] All");

                                    break;
                                }

                            case "Recurrence":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R289");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R289
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x0F,
                                        token,
                                        "MS-ASWBXML",
                                        289,
                                        @"[In Code Page 9: Tasks] [Tag name] Recurrence [Token] 0x0F [supports protocol versions] All");

                                    break;
                                }

                            case "Type":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R290");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R290
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x10,
                                        token,
                                        "MS-ASWBXML",
                                        290,
                                        @"[In Code Page 9: Tasks] [Tag name] Type [Token] 0x10 [supports protocol versions] All");

                                    break;
                                }

                            case "Start":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R291");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R291
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x11,
                                        token,
                                        "MS-ASWBXML",
                                        291,
                                        @"[In Code Page 9: Tasks] [Tag name] Start [Token] 0x11 [supports protocol versions] All");

                                    break;
                                }

                            case "Until":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R292");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R292
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x12,
                                        token,
                                        "MS-ASWBXML",
                                        292,
                                        @"[In Code Page 9: Tasks] [Tag name] Until [Token] 0x12 [supports protocol versions] All");

                                    break;
                                }

                            case "Occurrences":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R293");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R293
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x13,
                                        token,
                                        "MS-ASWBXML",
                                        293,
                                        @"[In Code Page 9: Tasks] [Tag name] Occurrences [Token] 0x13 [supports protocol versions] All");

                                    break;
                                }

                            case "Interval":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R294");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R294
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x14,
                                        token,
                                        "MS-ASWBXML",
                                        294,
                                        @"[In Code Page 9: Tasks] [Tag name] Interval [Token] 0x14 [supports protocol versions] All");

                                    break;
                                }

                            case "DayOfWeek":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R296");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R296
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x16,
                                        token,
                                        "MS-ASWBXML",
                                        296,
                                        @"[In Code Page 9: Tasks] [Tag name] DayOfWeek [Token] 0x16 [supports protocol versions] All");

                                    break;
                                }

                            case "DayOfMonth":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R295");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R295
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x15,
                                        token,
                                        "MS-ASWBXML",
                                        295,
                                        @"[In Code Page 9: Tasks] [Tag name] DayOfMonth [Token] 0x15 [supports protocol versions] All");

                                    break;
                                }

                            case "WeekOfMonth":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R297");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R297
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x17,
                                        token,
                                        "MS-ASWBXML",
                                        297,
                                        @"[In Code Page 9: Tasks] [Tag name] WeekOfMonth [Token] 0x17 [supports protocol versions] All");

                                    break;
                                }

                            case "MonthOfYear":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R298");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R298
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x18,
                                        token,
                                        "MS-ASWBXML",
                                        298,
                                        @"[In Code Page 9: Tasks] [Tag name] MonthOfYear [Token] 0x18 [supports protocol versions] All");

                                    break;
                                }

                            case "Regenerate":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R299");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R299
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x19,
                                        token,
                                        "MS-ASWBXML",
                                        299,
                                        @"[In Code Page 9: Tasks] [Tag name] Regenerate [Token] 0x19 [supports protocol versions] All");

                                    break;
                                }

                            case "DeadOccur":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R300");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R300
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x1A,
                                        token,
                                        "MS-ASWBXML",
                                        300,
                                        @"[In Code Page 9: Tasks] [Tag name] DeadOccur [Token] 0x1A [supports protocol versions] All");

                                    break;
                                }

                            case "ReminderSet":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R301");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R301
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x1B,
                                        token,
                                        "MS-ASWBXML",
                                        301,
                                        @"[In Code Page 9: Tasks] [Tag name] ReminderSet [Token] 0x1B [supports protocol versions] All");

                                    break;
                                }

                            case "ReminderTime":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R302");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R302
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x1C,
                                        token,
                                        "MS-ASWBXML",
                                        302,
                                        @"[In Code Page 9: Tasks] [Tag name] ReminderTime [Token] 0x1C [supports protocol versions] All");

                                    break;
                                }

                            case "Sensitivity":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R303");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R303
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x1D,
                                        token,
                                        "MS-ASWBXML",
                                        303,
                                        @"[In Code Page 9: Tasks] [Tag name] Sensitivity [Token] 0x1D [supports protocol versions] All");

                                    break;
                                }

                            case "StartDate":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R304");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R304
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x1E,
                                        token,
                                        "MS-ASWBXML",
                                        304,
                                        @"[In Code Page 9: Tasks] [Tag name] StartDate [Token] 0x1E [supports protocol versions] All");

                                    break;
                                }

                            case "UtcStartDate":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R305");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R305
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x1F,
                                        token,
                                        "MS-ASWBXML",
                                        305,
                                        @"[In Code Page 9: Tasks] [Tag name] UtcStartDate [Token] 0x1F [supports protocol versions] All");

                                    break;
                                }

                            case "Subject":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R306");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R306
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x20,
                                        token,
                                        "MS-ASWBXML",
                                        306,
                                        @"[In Code Page 9: Tasks] [Tag name] Subject [Token] 0x20 [supports protocol versions] All");

                                    break;
                                }

                            case "CalendarType":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R309");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R309
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x24,
                                        token,
                                        "MS-ASWBXML",
                                        309,
                                        @"[In Code Page 9: Tasks] [Tag name] CalendarType [Token] 0x24 [supports protocol versions] 14.0, 14.1, 16.0");

                                    break;
                                }

                            case "IsLeapMonth":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R310");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R310
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x25,
                                        token,
                                        "MS-ASWBXML",
                                        310,
                                        @"[In Code Page 9: Tasks] [Tag name] IsLeapMonth [Token] 0x25 [supports protocol versions] 14.0, 14.1, 16.0");

                                    break;
                                }

                            case "FirstDayOfWeek":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R311");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R311
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x26,
                                        token,
                                        "MS-ASWBXML",
                                        311,
                                        @"[In Code Page 9: Tasks] [Tag name] FirstDayOfWeek [Token] 0x26 [supports protocol versions] 14.1, 16.0");

                                    break;
                                }

                            default:
                                {
                                    Site.Assert.Fail("There exists unexpected Tag in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePage, tagName, token);
                                    break;
                                }
                        }
                    }
                }
            }
        }
    }
}