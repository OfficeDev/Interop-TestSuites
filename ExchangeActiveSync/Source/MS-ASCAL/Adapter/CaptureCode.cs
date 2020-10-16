namespace Microsoft.Protocols.TestSuites.MS_ASCAL
{
    using System;
    using System.Reflection;
    using System.Collections.Generic;
    using Common;
    using Microsoft.Protocols.TestSuites.Common.Response;
    using Microsoft.Protocols.TestTools;
    using ItemOperationsStore = Microsoft.Protocols.TestSuites.Common.DataStructures.ItemOperationsStore;
    using SearchItem = Microsoft.Protocols.TestSuites.Common.DataStructures.Search;
    using SearchStore = Microsoft.Protocols.TestSuites.Common.DataStructures.SearchStore;
    using SyncStore = Microsoft.Protocols.TestSuites.Common.DataStructures.SyncStore;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-ASCAL.
    /// </summary>
    public partial class MS_ASCALAdapter
    {
        /// <summary>
        /// This method is used to verify transport related requirement.
        /// </summary>
        private void VerifyTransport()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R13");

            // Verify MS-ASCAL requirement: MS-ASCAL_R13
            // ActiveSyncClient encodes XML request into WBXML and decodes WBXML to XML response, capture it directly if server responses succeed.
            Site.CaptureRequirement(
                13,
                @"[In Transport] The XML markup that constitutes the request body or the response body that is transmitted between the client and the server uses Wireless Application Protocol (WAP) Binary XML (WBXML), as specified in [MS-ASWBXML].");
        }

        /// <summary>
        /// This method is used to verify the message syntax related requirements.
        /// </summary>
        private void VerifyMessageSyntax()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R14");

            // Verify MS-ASCAL requirement: MS-ASCAL_R14
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                14,
                @"[In Message Syntax] The markup that is used by this protocol[MS-ASCAL] MUST be well-formed XML, as specified in [XML].");
        }

        /// <summary>
        /// This method is used to verify the Sync response related requirements
        /// </summary>
        /// <param name="syncResponse">Specified the SyncStore result returned from the server</param>
        private void VerifySyncCommandResponse(SyncStore syncResponse)
        {
            string activeSyncProtocolVersion = Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R521");

            // Verify MS-ASCAL requirement: MS-ASCAL_R521
            // If Sync response exists, this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                syncResponse,
                521,
                @"[In Synchronizing Calendar Data Between Client and Server] The server responds with a Sync command response ([MS-ASCMD] section 2.2.2.19).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R537");

            // Verify MS-ASCAL requirement: MS-ASCAL_R537
            // Since R521's been captured, this requirement can be captured too.
            Site.CaptureRequirement(
                537,
                @"[In Sync Command Response] When a client uses the Sync command request ([MS-ASCMD] section 2.2.2.19), as specified in section 3.1.5.3, to synchronize its Calendar class items for a specified user with the calendar items that are currently stored by the server, the server responds with a Sync command response ([MS-ASCMD] section 2.2.2.19).");

            if (null != syncResponse.AddElements)
            {
                for (int i = syncResponse.AddElements.Count - 1; i >= 0; i--)
                {
                    if (null != syncResponse.AddElements[i].Calendar)
                    {
                        Type type = typeof(DataStructures.Calendar);
                        PropertyInfo[] properties = type.GetProperties();
                        if (null != syncResponse.AddElements[i].Calendar.AllDayEvent)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R564");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R564
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                564,
                                @"[In AllDayEvent] The value of this element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R565. The actual value of AllDayEvent element is {0}.", syncResponse.AddElements[i].Calendar.AllDayEvent);

                            // Verify MS-ASCAL requirement: MS-ASCAL_R565
                            Site.CaptureRequirementIfIsTrue(
                                syncResponse.AddElements[i].Calendar.AllDayEvent == 0 || syncResponse.AddElements[i].Calendar.AllDayEvent == 1,
                                565,
                                @"[In AllDayEvent] The value of the AllDayEvent element MUST be one of the values[0,1] listed in the following table.");

                            this.VerifyUnsignedByteDataType(syncResponse.AddElements[i].Calendar.AllDayEvent);
                        }

                        if (!activeSyncProtocolVersion.Equals("12.1"))
                        {
                            if (null != syncResponse.AddElements[i].Calendar.AppointmentReplyTime)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R91");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R91
                                // The schema has been validated, so this requirement can be captured.
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    91,
                                    @"[In AppointmentReplyTime] The value of this element is a string data type, represented as a Compact DateTime ([MS-ASDTYPE] section 2.6.5).");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_88011");

                                // Since R91 is captured, this requirement can be captured too.
                                Site.CaptureRequirement(
                                    88011,
                                    @"[In AppointmentReplyTime] As a top-level element of the Calendar class, the AppointmentReplyTime<1> element specifies the date and time that the current user responded to the meeting request.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_566");

                                // Since R91 is captured, this requirement can be captured too.
                                Site.CaptureRequirement(
                                    566,
                                    @"[In AppointmentReplyTime] A command response has a maximum of one top-level AppointmentReplyTime element per response");

                                this.VerifyCompactDateTimeDataType();
                            }
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Attendees && null != syncResponse.AddElements[i].Calendar.Attendees.Attendee)
                        {
                            for (int j = syncResponse.AddElements[i].Calendar.Attendees.Attendee.Length - 1; j >= 0; j--)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R100");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R100
                                // If Email exists, this requirement can be captured.
                                Site.CaptureRequirementIfIsNotNull(
                                    syncResponse.AddElements[i].Calendar.Attendees.Attendee[j].Email,
                                    100,
                                    @"[In Attendee] [The Attendee element can have the following child elements:] Email (section 2.2.2.17): One instance of this element is required.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R101");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R101
                                // If Name exists, this requirement can be captured.
                                Site.CaptureRequirementIfIsNotNull(
                                    syncResponse.AddElements[i].Calendar.Attendees.Attendee[j].Name,
                                    101,
                                    @"[In Attendee][The Attendee element can have the following child elements:] Name (section 2.2.2.28): One instance of this element is required.");

                                this.VerifyContainerDataType();

                                if (syncResponse.AddElements[i].Calendar.Attendees.Attendee[j].AttendeeTypeSpecified)
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R122");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R122
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        122,
                                        @"[In AttendeeType] The value of this element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R121");

                                    // Since R122 is captured, this requirement can be captured too.
                                    Site.CaptureRequirement(
                                        121,
                                        @"[In AttendeeType] A command response has a maximum of one AttendeeType element per Attendee element.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R123, the value of AttendeeType is {0}", syncResponse.AddElements[i].Calendar.Attendees.Attendee[j].AttendeeType);

                                    string[] expectedValues = new string[] { "1", "2", "3" };
                                    Common.VerifyActualValues("AttendeeType", expectedValues, syncResponse.AddElements[i].Calendar.Attendees.Attendee[j].AttendeeType.ToString(), this.Site);

                                    // If the verification of actual values success, then requirement MS-ASCAL_R123 can be captured directly.
                                    // Verify MS-ASCAL requirement: MS-ASCAL_R123
                                    Site.CaptureRequirement(
                                        123,
                                        @"[In AttendeeType] The value of the AttendeeType element MUST be one of the values[1,2,3] specified in the following table.");

                                    this.VerifyUnsignedByteDataType(syncResponse.AddElements[i].Calendar.Attendees.Attendee[j].AttendeeType);
                                    
                                    foreach (PropertyInfo prop in properties)
                                    {
                                        var propName = prop.Name;
                                        var propValue = prop.GetValue(syncResponse.AddElements[i].Calendar);
                                        if (propName != "Attendees" && propValue != null)
                                        {
                                            // Add the debug information
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R434001");

                                            // Verify MS-ASCAL requirement: MS-ASCAL_R434001
                                            // If StartTime exists, this requirement can be captured.
                                            Site.CaptureRequirementIfIsNotNull(
                                                syncResponse.AddElements[i].Calendar.StartTime,
                                                434001,
                                                @"[In StartTime] [In protocol version 12.0, 12.1, 14.0, 14.1, 16.0, and 16.1, a Sync command response MUST contain one instance of the StartTime element if more than just] AttendeeType (section 2.2.2.6) has changed.");
                                            break;
                                        }
                                    }
                                }

                                if (syncResponse.AddElements[i].Calendar.Attendees.Attendee[j].AttendeeStatusSpecified)
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R117");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R117
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        117,
                                        @"[In AttendeeStatus] The value of this element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R116");

                                    // Since R117 is captured, this requirement can be captured too.
                                    Site.CaptureRequirement(
                                        116,
                                        @"[In AttendeeStatus] A command response has a maximum of one AttendeeStatus element per Attendee element.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R118, the value of AttendeeStatus is {0}", syncResponse.AddElements[i].Calendar.Attendees.Attendee[j].AttendeeStatus);

                                    string[] expectedValues = new string[] { "0", "2", "3", "4", "5" };
                                    Common.VerifyActualValues("AttendeeStatus", expectedValues, syncResponse.AddElements[i].Calendar.Attendees.Attendee[j].AttendeeStatus.ToString(), this.Site);

                                    // If the verification of actual values success, then requirement MS-ASCAL_R118 can be captured directly.
                                    // Verify MS-ASCAL requirement: MS-ASCAL_R118
                                    Site.CaptureRequirement(
                                        118,
                                        @"[In AttendeeStatus] The value of the AttendeeStatus element MUST be one of the values[0,2,3,4,5] listed in the following table.");

                                    this.VerifyUnsignedByteDataType(syncResponse.AddElements[i].Calendar.Attendees.Attendee[j].AttendeeStatus);
                                }
                            }

                            this.VerifyContainerDataType();
                        }

                        if (null != syncResponse.AddElements[i].Calendar.DtStamp)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R224");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R224
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                224,
                                @"[In DtStamp] The value of this element is a string data type, represented as a Compact DateTime ([MS-ASDTYPE] section 2.6.5).");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R22111");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R22111
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                22111,
                                @"[In DtStamp] [DtStamp] specifies the date and time that this exception was created.");

                            this.VerifyCompactDateTimeDataType();

                            foreach (PropertyInfo prop in properties)
                            {
                                var propName = prop.Name;
                                var propValue = prop.GetValue(syncResponse.AddElements[i].Calendar);
                                if (propName != "DtStamp" && propValue != null)
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R434");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R434
                                    // If StartTime exists, this requirement can be captured.
                                    Site.CaptureRequirementIfIsNotNull(
                                        syncResponse.AddElements[i].Calendar.StartTime,
                                        434,
                                        @"[In StartTime] In protocol version 12.0, 12.1, 14.0, 14.1, 16.0, and 16.1, a Sync command response MUST contain one instance of the StartTime element if more than just DtStamp (section 2.2.2.18) has changed.");
                                    break;
                                }
                            }
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Body)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R12411");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R12411
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                12411,
                                @"[In Body (AirSyncBase Namespace)] As a top-level element of the Calendar class, the airsyncbase:Body element specifies the body text of the calendar item.");

                            this.VerifyContainerDataType();
                        }

                        if (null != syncResponse.AddElements[i].Calendar.BusyStatus)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R137");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R137
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                137,
                                @"[In BusyStatus] The value of this element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R13111");

                            // Since R137 is captured, this requirement can be captured too.
                            Site.CaptureRequirement(
                                13111,
                                @"[In BusyStatus] As a top-level element of the Calendar class, the BusyStatus element specifies whether the recipient is busy at the time of the meeting.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R138, the value of BusyStatus is {0}", syncResponse.AddElements[i].Calendar.BusyStatus);

                            string[] expectedValues = new string[] { "0", "1", "2", "3", "4" };
                            Common.VerifyActualValues("BusyStatus", expectedValues, syncResponse.AddElements[i].Calendar.BusyStatus.ToString(), this.Site);

                            // If the verification of actual values success, then requirement MS-ASCAL_R138 can be captured directly.
                            // Verify MS-ASCAL requirement: MS-ASCAL_R138
                            Site.CaptureRequirement(
                                138,
                                @"[In BusyStatus] The value of the BusyStatus element MUST be one of the values[0,1,2,3,4] listed in the following table.");

                            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "12.1")
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2235");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R2235
                                Site.CaptureRequirementIfAreNotEqual<byte>(
                                    4,
                                    (byte)syncResponse.AddElements[i].Calendar.BusyStatus,
                                    2235,
                                    @"[In BusyStatus] The value 4 (working elsewhere) is not supported in protocol versions 2.5, 12.0, 12.1[, 14.0, and 14.1].");
                            }

                            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.0")
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2043");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R2043
                                Site.CaptureRequirementIfAreNotEqual<byte>(
                                    4,
                                    (byte)syncResponse.AddElements[i].Calendar.BusyStatus,
                                    2043,
                                    @"[In BusyStatus] The value 4 (working elsewhere) is not supported in protocol versions [2.5, 12.0, 12.1,] 14.0[, and 14.1].");
                            }

                            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.1")
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2044");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R2044
                                Site.CaptureRequirementIfAreNotEqual<byte>(
                                    4,
                                    (byte)syncResponse.AddElements[i].Calendar.BusyStatus,
                                    2044,
                                    @"[In BusyStatus] The value 4 (working elsewhere) is not supported in protocol versions [2.5, 12.0, 12.1, 14.0, and] 14.1.");
                            }

                            this.VerifyUnsignedByteDataType(syncResponse.AddElements[i].Calendar.BusyStatus);
                        }

                        if (!activeSyncProtocolVersion.Equals("12.1"))
                        {
                            if (null != syncResponse.AddElements[i].Calendar.Recurrence && syncResponse.AddElements[i].Calendar.Recurrence.CalendarTypeSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R145");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R145
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    145,
                                    @"[In CalendarType] The value of this element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R146");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R146
                                Site.CaptureRequirementIfIsTrue(
                                    syncResponse.AddElements[i].Calendar.Recurrence.CalendarType <= 23,
                                    146,
                                    @"[In CalendarType] The value of the CalendarType element MUST be one of the values[0~23] listed in the following table.");

                                this.VerifyUnsignedByteDataType(syncResponse.AddElements[i].Calendar.Recurrence.CalendarType);
                            }
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Categories && syncResponse.AddElements[i].Calendar.Categories.Category.Length >= 0 && syncResponse.AddElements[i].Calendar.Categories.Category.Length <= 300)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R182");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R182
                            // If schema is verified successfully, this requirement can be captured.
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                182,
                                @"[In Categories] [The Categories element can have the following child element:]Category (section 2.2.2.11)");

                            this.VerifyContainerDataType();

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R185");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R185
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                185,
                                @"[In Category] The value of this element is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                            this.VerifyStringDataType();
                        }

                        // DayOfMonth is required when Type is set to 2 or 5.
                        if (null != syncResponse.AddElements[i].Calendar.Recurrence && (syncResponse.AddElements[i].Calendar.Recurrence.Type == 2 || syncResponse.AddElements[i].Calendar.Recurrence.Type == 5))
                        {
                            if (syncResponse.AddElements[i].Calendar.Recurrence.DayOfMonthSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R192");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R192
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    192,
                                    @"[In DayOfMonth] The value of the DayOfMonth element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R193");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R193
                                Site.CaptureRequirementIfIsTrue(
                                    syncResponse.AddElements[i].Calendar.Recurrence.DayOfMonth >= 1 && syncResponse.AddElements[i].Calendar.Recurrence.DayOfMonth <= 31,
                                    193,
                                    @"[In DayOfMonth] The value of this element[DayOfMonth] MUST be between 1 and 31.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R188");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R188
                                // If schema is verified successfully, this requirement can be captured.
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    188,
                                    @"[In DayOfMonth] The DayOfMonth element is a child element of the Recurrence element (section 2.2.2.35).");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R18811");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R18811
                                // DayofMonth element means the day of month when Type is 2 or 5.
                                // If schema is verified successfully, this requirement can be captured
                                Site.CaptureRequirement(
                                    18811,
                                    @"[In DayOfMonth] The DayOfMonth element specifies the day of the month for the recurrence.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R19111");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R19111
                                // If schema is verified successfully, this requirement can be captured
                                Site.CaptureRequirement(
                                    19111,
                                    @"[In DayOfMonth] A command response has a maximum of one DayOfMonth child element per Recurrence element.");

                                this.VerifyUnsignedByteDataType(syncResponse.AddElements[i].Calendar.Recurrence.DayOfMonth);

                                if (syncResponse.AddElements[i].Calendar.Recurrence.Type == 2)
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R38935");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R38935
                                    Site.CaptureRequirementIfIsTrue(
                                        syncResponse.AddElements[i].Calendar.Recurrence.DayOfMonthSpecified,
                                        38935,
                                        @"[In Recurrence Patterns][When the Type element is set to 2, meaning a monthly occurrence, the following elements are supported] DayOfMonth (section 2.2.2.12). Required.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R19411");

                                    // When Type is 2 and DayOfMonth is included in response, this requirement can be captured.
                                    // If schema is verified successfully, this requirement can be captured
                                    Site.CaptureRequirement(
                                        19411,
                                        @"[In DayOfMonth] The DayOfMonth element MUST be included in responses when the Type element value is 2[or 5].");

                                    if (null != syncResponse.AddElements[i].Calendar.Recurrence.Until)
                                    {
                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_470");

                                        // If Until exists, this requirement can be captured.
                                        // If schema is verified successfully, this requirement can be captured
                                        Site.CaptureRequirement(
                                            470,
                                            @"[In Until] The value of the Until element is a string data type, represented as a Compact DateTime ([MS-ASDTYPE] section 2.6.5).");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_46411");

                                        // If Until exists, this requirement can be captured.
                                        // If schema is verified successfully, this requirement can be captured
                                        Site.CaptureRequirement(
                                            46411,
                                            @"[In Until] The Until element specifies the end date and time of the recurrence.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_ 46611");

                                        // If Until exists, this requirement can be captured.
                                        // If schema is verified successfully, this requirement can be captured
                                        Site.CaptureRequirement(
                                            46611,
                                            @"[In Until] A command response has a maximum of one Until child element per Recurrence element.");

                                        this.VerifyCompactDateTimeDataType();
                                    }
                                }

                                if (syncResponse.AddElements[i].Calendar.Recurrence.Type == 5)
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R38941");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R38941
                                    Site.CaptureRequirementIfIsTrue(
                                        syncResponse.AddElements[i].Calendar.Recurrence.MonthOfYearSpecified,
                                        38941,
                                        @"[In Recurrence Patterns][When the Type element is set to 5, meaning a yearly occurrence, the following elements are supported:] MonthOfYear (section 2.2.2.27). Required.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R38943");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R38943
                                    Site.CaptureRequirementIfIsTrue(
                                        syncResponse.AddElements[i].Calendar.Recurrence.DayOfMonthSpecified,
                                        38943,
                                        @"[In Recurrence Patterns][When the Type element is set to 5, meaning a yearly occurrence, the following elements are supported:] DayOfMonth. Required.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R19412");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R19412
                                    // When Type is 5 and DayOfMonth is included in response, this requirement can be captured.
                                    // If schema is verified successfully, this requirement can be captured
                                    Site.CaptureRequirement(
                                        19412,
                                        @"[In DayOfMonth] The DayOfMonth element MUST be included in responses when the Type element value is [2 or] 5.");
                                }
                            }
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Recurrence)
                        {
                            if (syncResponse.AddElements[i].Calendar.Recurrence.OccurrencesSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R38916");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R38916
                                // Since the schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    38916,
                                    @"[In Recurrence Patterns][For all values of the Type element, the following elements are optional] Occurrences (section 2.2.2.30) [or Until (section 2.2.2.45)]. Either the Occurrences [or Until element] is required to specify an end date. ");
                            }
                            else if (null != syncResponse.AddElements[i].Calendar.Recurrence.Until)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R389161");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R389161
                                // Since the schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    389161,
                                    @"[In Recurrence Patterns][For all values of the Type element, the following elements are optional] [Occurrences (section 2.2.2.30) or] Until (section 2.2.2.45).  [Either the Occurrences or] Until element is required to specify an end date. ");
                            }
                            else
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R389162");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R389162
                                // Since the schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    389162,
                                    @"[In Recurrence Patterns][For all values of the Type element, the following elements are optional] If neither value [Occurrences or Until element] is set, the event has no end date.  ");
                            }
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Recurrence && (syncResponse.AddElements[i].Calendar.Recurrence.Type == 0 || syncResponse.AddElements[i].Calendar.Recurrence.Type == 1 || syncResponse.AddElements[i].Calendar.Recurrence.Type == 3 || syncResponse.AddElements[i].Calendar.Recurrence.Type == 6))
                        {
                            if (syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeekSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R200, the value of DayOfWeek is {0}", syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeek);

                                string[] expectedValues = new string[] { "1", "2", "4", "8", "16", "32", "62", "64", "65", "127" };
                                Common.VerifyActualValues("DayOfWeek", expectedValues, syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeek.ToString(), this.Site);

                                // If the verification of actual values success, then requirement MS-ASCAL_R200 can be captured directly.
                                // Verify MS-ASCAL requirement: MS-ASCAL_R200
                                Site.CaptureRequirement(
                                    200,
                                    @"[In DayOfWeek] The value of the DayOfWeek element MUST be one of the values[1,2,4,8,16,32,62,64,65,127] listed in the following table [or the sum of more than one of the values listed in the following table (in which case this task recurs on more than one day)].");
                            }
                            else
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R3892311");

                                // Since the schema is validated, this requirement can be captured directly.
                                Site.CaptureRequirement(
                                    3892311,
                                    @"[In Recurrence Patterns][When the Type element is set to zero (0), meaning a weekly occurrence, the following elements are supported] If the DayOfWeek element is not set, the recurrence is a daily occurrence, occurring n days apart, where n is the value of the Interval element. ");
                            }

                            if (syncResponse.AddElements[i].Calendar.Recurrence.Type == 0)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R204111");

                                // If Type is 0 and DayOfWeek is included in response, this requirement can be captured.
                                // If schema is verified successfully, this requirement can be captured
                                Site.CaptureRequirement(
                                    204111,
                                    @"[In DayOfWeek] The DayOfWeek element MUST be included in responses when the Type element (section 2.2.2.43) value is 0 (zero).");
                            }

                            if (syncResponse.AddElements[i].Calendar.Recurrence.Type == 1)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R38929");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R38929
                                Site.CaptureRequirementIfIsTrue(
                                    syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeekSpecified,
                                    38929,
                                    @"[In Recurrence Patterns][When the Type element is set to 1, meaning a weekly occurrence, the following elements are supported] DayOfWeek. Required.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R204112");

                                // If Type is 1 and DayOfWeek is included in response, this requirement can be captured.
                                Site.CaptureRequirementIfIsTrue(
                                    syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeekSpecified,
                                    204112,
                                    @"[In DayOfWeek] The DayOfWeek element MUST be included in responses when the Type element (section 2.2.2.43) value is 1.");
                            }

                            if (syncResponse.AddElements[i].Calendar.Recurrence.Type == 3)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R3893713");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R3893713
                                Site.CaptureRequirementIfIsTrue(
                                    syncResponse.AddElements[i].Calendar.Recurrence.WeekOfMonthSpecified,
                                    3893713,
                                    @"[In Recurrence Patterns][When the Type element is set to 3, meaning a monthly occurrence on the nth day, the following elements are supported] WeekOfMonth (section 2.2.2.46). Required.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R3893717");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R3893717
                                Site.CaptureRequirementIfIsTrue(
                                    syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeekSpecified,
                                    3893717,
                                    @"[In Recurrence Patterns][When the Type element is set to 3, meaning a monthly occurrence on the nth day, the following elements are supported] DayOfWeek. Required. ");

                                if (syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeek == 127)
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R3893714");

                                    // Since R3893713 is captured, this requirement can be captured.
                                    Site.CaptureRequirement(
                                        3893714,
                                        @"[In Recurrence Patterns][When the Type element is set to 3, meaning a monthly occurrence on the nth day, the following elements are supported] If the DayOfWeek element is set to 127, the WeekOfMonth element indicates the day of the month that the event occurs. ");
                                }

                                if (syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeek == 62)
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R3893715");

                                    // Since R3893713 is captured, this requirement can be captured.
                                    Site.CaptureRequirement(
                                        3893715,
                                        @"[In Recurrence Patterns][When the Type element is set to 3, meaning a monthly occurrence on the nth day, the following elements are supported] If the DayOfWeek element is set to 62, to specify weekdays, the WeekOfMonth element indicates the nth weekday of the month, where n is the value of WeekOfMonth element. ");
                                }

                                if (syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeek == 65)
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R3893716");

                                    // Since R3893713 is captured, this requirement can be captured.
                                    Site.CaptureRequirement(
                                        3893716,
                                        @"[In Recurrence Patterns][When the Type element is set to 3, meaning a monthly occurrence on the nth day, the following elements are supported] If the DayOfWeek element is set to 65, to specify weekends, the WeekOfMonth element indicates the nth weekend day of the month, where n is the value of WeekOfMonth element.");
                                }

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R204113");

                                // If Type is 3 and DayOfWeek is included in response, this requirement can be captured.
                                Site.CaptureRequirementIfIsTrue(
                                    syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeekSpecified,
                                    204113,
                                    @"[In DayOfWeek] The DayOfWeek element MUST be included in responses when the Type element (section 2.2.2.43) value is 3.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R478113");

                                // If Type is 3 and WeekOfMonth is included in response, this requirement can be captured.
                                // If schema is verified successfully, this requirement can be captured
                                Site.CaptureRequirement(
                                    478113,
                                    @"[In WeekOfMonth] The WeekOfMonth element MUST be included in responses when the Type element (section 2.2.2.43) value is 3.");
                            }

                            if (syncResponse.AddElements[i].Calendar.Recurrence.Type == 6)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R38948");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R38948
                                Site.CaptureRequirementIfIsTrue(
                                    syncResponse.AddElements[i].Calendar.Recurrence.WeekOfMonthSpecified,
                                    38948,
                                    @"[In Recurrence Patterns][When the Type element is set to 6, meaning a yearly occurrence on the nth day, the following elements are supported] WeekOfMonth. Required.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R38950");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R38950
                                Site.CaptureRequirementIfIsTrue(
                                    syncResponse.AddElements[i].Calendar.Recurrence.MonthOfYearSpecified,
                                    38950,
                                    @"[In Recurrence Patterns][When the Type element is set to 6, meaning a yearly occurrence on the nth day, the following elements are supported] MonthOfYear. Required.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R204114");

                                // If Type is 6 and DayOfWeek is included in response, this requirement can be captured.
                                // If schema is verified successfully, this requirement can be captured
                                Site.CaptureRequirement(
                                    204114,
                                    @"[In DayOfWeek] The DayOfWeek element MUST be included in responses when the Type element (section 2.2.2.43) value is 6.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R478114");

                                // If Type is 6 and WeekOfMonth is included in response, this requirement can be captured.
                                // If MS-ASCAL_R38948 is verified successfully, this requirement can be captured
                                Site.CaptureRequirement(
                                    478114,
                                    @"[In WeekOfMonth] The WeekOfMonth element MUST be included in responses when the Type element (section 2.2.2.43) value is 6.");
                            }
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Recurrence && syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeekSpecified)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R196");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R196
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                196,
                                @"[In DayOfWeek] The DayOfWeek element is a child element of the Recurrence element (section 2.2.2.35) that specifies the day of the week for the recurrence.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R199");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R199
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                199,
                                @"[In DayOfWeek] The value of this element is an unsignedShort data type, as specified in [XMLSCHEMA2/2].");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R19811");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R19811
                            // Since R199 is captured, this requirement can be captured too.
                            Site.CaptureRequirement(
                                19811,
                                @"[In DayOfWeek] A command response has a maximum of one DayOfWeek child element per Recurrence element.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R20011");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R20011
                            Site.CaptureRequirementIfIsTrue(
                                syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeek >= 1 && syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeek <= 381,
                                20011,
                                @"[In DayOfWeek] The value of the DayOfWeek element MUST be [one of the values[1,2,4,8,16,32,62,64,65,127] listed in the following table, or] the sum of more than one of the values listed in the following table (in which case this task recurs on more than one day).");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R201");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R201
                            Site.CaptureRequirementIfIsTrue(
                                syncResponse.AddElements[i].Calendar.Recurrence.DayOfWeek <= 127,
                                201,
                                @"[In DayOfWeek] The value of the DayOfWeek element MUST NOT be greater than 127.");
                        }

                        if (!activeSyncProtocolVersion.Equals("12.1") && !activeSyncProtocolVersion.Equals("14.0"))
                        {
                            if (!syncResponse.AddElements[i].Calendar.Reminder.HasValue)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R3951111");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R3951111
                                Site.CaptureRequirement(
                                    3951111,
                                    @"[In Reminder] The value of this element is [an unsignedInt data type, as specified in [XMLSCHEMA2/2], or] an EmptyTag data type, which contains no value.<15>");
                            }
                        }
                        else
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2185");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R2185
                            Site.CaptureRequirementIfIsTrue(
                                syncResponse.AddElements[i].Calendar.Reminder == null || syncResponse.AddElements[i].Calendar.Reminder.ToString().Length != 0,
                                2185,
                                @"[In Reminder] When protocol version 2.5, 12.0, 12.1, or 14.0 is used, the value of the Reminder element cannot be an EmptyTag data type.");
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Reminder)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R39511");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R39511
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                39511,
                                @"[In Reminder] The value of this element is an unsignedInt data type, as specified in [XMLSCHEMA2/2], [or an EmptyTag data type, which contains no value].<15>");
                        }

                        if (!activeSyncProtocolVersion.Equals("12.1") && !activeSyncProtocolVersion.Equals("14.0"))
                        {
                            if (null != syncResponse.AddElements[i].Calendar.Recurrence && syncResponse.AddElements[i].Calendar.Recurrence.FirstDayOfWeekSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R269");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R269
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    269,
                                    @"[In FirstDayOfWeek] The FirstDayOfWeek<10> element is a child element of the Recurrence element (section 2.2.2.35) that specifies which day is considered the first day of the calendar week for the recurrence.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R276");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R276
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    276,
                                    @"[In FirstDayOfWeek] The value of this element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R277, the value of FirstDayOfWeek is {0}", syncResponse.AddElements[i].Calendar.Recurrence.FirstDayOfWeek);

                                string[] expectedValues = new string[] { "0", "1", "2", "3", "4", "5", "6" };
                                Common.VerifyActualValues("FirstDayOfWeek", expectedValues, syncResponse.AddElements[i].Calendar.Recurrence.FirstDayOfWeek.ToString(), this.Site);

                                // If the verification of actual values success, then requirement MS-ASCAL_R277 can be captured directly.
                                // Verify MS-ASCAL requirement: MS-ASCAL_R277
                                Site.CaptureRequirement(
                                    277,
                                    @"[In FirstDayOfWeek] The value of the FirstDayOfWeek element MUST be one of the values[0,1,2,3,4,5,6] listed in the following table.");

                                this.VerifyUnsignedByteDataType(syncResponse.AddElements[i].Calendar.Recurrence.FirstDayOfWeek);
                            }
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Recurrence)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R283");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R283
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                283,
                                @"[In Interval] The value of this element is an unsignedShort data type.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R28011");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R28011
                            // Since R283 is captured, this requirement can be captured too.
                            Site.CaptureRequirement(
                                28011,
                                @"[In Interval] The Interval element specifies the interval between recurrences.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R28211");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R28211
                            // Since R283 is captured, this requirement can be captured too.
                            Site.CaptureRequirement(
                                28211,
                                @"[In Interval] A command response has a maximum of one Interval child element per Recurrence element.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R28311");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R28311
                            Site.CaptureRequirementIfIsTrue(
                                syncResponse.AddElements[i].Calendar.Recurrence.Interval < 1000,
                                28311,
                                @"[In Interval] The value of this element is between a minimum value of 0 and a maximum value of 999.");
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Recurrence && syncResponse.AddElements[i].Calendar.Recurrence.OccurrencesSpecified)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R344");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R344
                            // If the command returns responses and the schema validation is successful, this requirement can be verified.
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                344,
                                @"[In Occurrences] The value of the Occurrences element is an unsignedShort, as specified in [XMLSCHEMA2/2].");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R33811");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R33811
                            // Since R344 is captured, this requirement can be captured too.
                            Site.CaptureRequirement(
                                33811,
                                @"[In Occurrences] The Occurrences element specifies the number of occurrences before the series ends.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R34011");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R34011
                            // Since R344 is captured, this requirement can be captured too.
                            Site.CaptureRequirement(
                                34011,
                                @"[In Occurrences] A command response has a maximum of one Occurrences child element per Recurrence element.");
                        }

                        if (!activeSyncProtocolVersion.Equals("12.1"))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R139");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R139
                            // If schema is verified successfully, this requirement can be captured.
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                139,
                                @"[In CalendarType] The CalendarType element<3> is a child element of the Recurrence element (section 2.2.2.35) that specifies the calendar system used by the recurrence.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R284");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R284
                            // IsLeapMonth is optional so this requirement can be verified no matter whether it appears or not.
                            // If schema is verified successfully, this requirement can be captured
                            Site.CaptureRequirement(
                                284,
                                @"[In IsLeapMonth] The IsLeapMonth element<11> is an optional child element of the Recurrence element (section 2.2.2.35).");

                            if (null != syncResponse.AddElements[i].Calendar.Recurrence && syncResponse.AddElements[i].Calendar.Recurrence.IsLeapMonthSpecified)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R292. The actual value of IsLeapMonth element is {0}.", syncResponse.AddElements[i].Calendar.Recurrence.IsLeapMonth);

                                // Verify MS-ASCAL requirement: MS-ASCAL_R292
                                Site.CaptureRequirementIfIsTrue(
                                    syncResponse.AddElements[i].Calendar.Recurrence.IsLeapMonth == 0 ||
                                    syncResponse.AddElements[i].Calendar.Recurrence.IsLeapMonth == 1,
                                    292,
                                    @"[In IsLeapMonth] The value of the IsLeapMonth element MUST be one of the values[0,1] listed in the following table.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R291");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R291
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    291,
                                    @"[In IsLeapMonth] The value of this element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R287");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R287
                                // Since R291 is captured, this requirement can be captured too.
                                Site.CaptureRequirement(
                                    287,
                                    @"[In IsLeapMonth] A command response has a maximum of one IsLeapMonth child element per Recurrence element.");

                                this.VerifyUnsignedByteDataType(syncResponse.AddElements[i].Calendar.Recurrence.IsLeapMonth);
                            }

                            if (null != syncResponse.AddElements[i].Calendar.ResponseType)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R405");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R405
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    405,
                                    @"[In ResponseType] A command response has a maximum of one top-level ResponseType element per response [and a maximum of one ResponseType child element per Exception element].");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R40611");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R40611
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    40611,
                                    @"[In ResponseType] The value of this element is an unsignedInt data type, as specified in [XMLSCHEMA2/2].");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R407, the value of ResponseType is {0}", syncResponse.AddElements[i].Calendar.ResponseType);

                                string[] expectedValues = new string[] { "0", "1", "2", "3", "4", "5" };
                                Common.VerifyActualValues("ResponseType", expectedValues, syncResponse.AddElements[i].Calendar.ResponseType.ToString(), this.Site);

                                // If the verification of actual values success, then requirement MS-ASCAL_R407 can be captured directly.
                                // Verify MS-ASCAL requirement: MS-ASCAL_R407
                                Site.CaptureRequirement(
                                    407,
                                    @"[In ResponseType] The value of the ResponseType element MUST be one of the values[0,1,2,3,4,5] listed in the following table.");
                            }
                        }

                        if (!activeSyncProtocolVersion.Equals("12.1") && !activeSyncProtocolVersion.Equals("14.0"))
                        {
                            if (null != syncResponse.AddElements[i].Calendar.MeetingStatus)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R309, the value of MeetingStatus is {0}", syncResponse.AddElements[i].Calendar.MeetingStatus);

                                string[] expectedValues = new string[] { "0", "1", "3", "5", "7", "9", "11", "13", "15" };
                                Common.VerifyActualValues("MeetingStatus", expectedValues, syncResponse.AddElements[i].Calendar.MeetingStatus.ToString(), this.Site);

                                // If the verification of actual values success, then requirement MS-ASCAL_R309 can be captured directly.
                                // Verify MS-ASCAL requirement: MS-ASCAL_R309
                                Site.CaptureRequirement(
                                    309,
                                    @"[In MeetingStatus] The value of the MeetingStatus element MUST be one of the values[0,1,3,5,7,9,11,13,15,] listed in the following table.");
                                
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R308");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R308
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    308,
                                    @"[In MeetingStatus] The value of this element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R30311");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R30311
                                // Since R308 is captured, this requirement can be captured too.
                                Site.CaptureRequirement(
                                    30311,
                                    @"[In MeetingStatus][As a top-level element of the Calendar class] the MeetingStatus element specifies whether the event is a meeting or an appointment, whether the event is canceled or active, and whether the user was the organizer.");

                                this.VerifyUnsignedByteDataType(syncResponse.AddElements[i].Calendar.MeetingStatus);
                            }
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Recurrence && syncResponse.AddElements[i].Calendar.Recurrence.MonthOfYearSpecified && (syncResponse.AddElements[i].Calendar.Recurrence.Type == 5 || syncResponse.AddElements[i].Calendar.Recurrence.Type == 6))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R321");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R321
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                321,
                                @"[In MonthOfYear] The MonthOfYear element is a child element of the Recurrence element (section 2.2.2.35) that specifies the month of the year for the recurrence.");

                            Site.Assert.IsTrue(syncResponse.AddElements[i].Calendar.Recurrence.MonthOfYearSpecified, "MonthOfYear is required when type is 5 or 6.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R325");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R325
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                325,
                                @"[In MonthOfYear] The value of this element[MonthOfYear] is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R32411");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R32411
                            // Since R325 is captured, this requirement can be captured too.
                            Site.CaptureRequirement(
                                32411,
                                @"[In MonthOfYear] A command response has a maximum of one MonthOfYear child element per Recurrence element.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R326");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R326
                            Site.CaptureRequirementIfIsTrue(
                                syncResponse.AddElements[i].Calendar.Recurrence.MonthOfYear >= 1 && syncResponse.AddElements[i].Calendar.Recurrence.MonthOfYear <= 12,
                                326,
                                @"[In MonthOfYear] The value of the MonthOfYear element MUST be between 1 and 12.");

                            if (syncResponse.AddElements[i].Calendar.Recurrence.Type == 5)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R32711");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R32711
                                // If Type is 5 and MonthOfYear is included in response, this requirement can be captured.
                                Site.CaptureRequirementIfIsTrue(
                                    syncResponse.AddElements[i].Calendar.Recurrence.MonthOfYearSpecified,
                                    32711,
                                    @"[In MonthOfYear] The MonthOfYear element MUST be included in responses when the Type element value is 5.");
                            }

                            if (syncResponse.AddElements[i].Calendar.Recurrence.Type == 6)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R32712");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R32712
                                // If Type is 5 and MonthOfYear is included in response, this requirement can be captured.
                                // If schema is verified successfully, this requirement can be captured
                                Site.CaptureRequirement(
                                    32712,
                                    @"[In MonthOfYear] The MonthOfYear element MUST be included in responses when the Type element value is 6.");
                            }

                            this.VerifyUnsignedByteDataType(syncResponse.AddElements[i].Calendar.Recurrence.MonthOfYear);
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Subject)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R43911");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R43911
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                43911,
                                @"[In Subject] As a top-level element of the Calendar class, the Subject element specifies the subject of the calendar item.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R444");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R444
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                444,
                                @"[In Subject] The value of this element is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                            this.VerifyStringDataType();
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Sensitivity)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R427");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R427
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                427,
                                @"[In Sensitivity] The value of this element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R428, the value of Sensitivity is {0}", syncResponse.AddElements[i].Calendar.Sensitivity);

                            string[] expectedValues = new string[] { "0", "1", "2", "3" };
                            Common.VerifyActualValues("Sensitivity", expectedValues, syncResponse.AddElements[i].Calendar.Sensitivity.ToString(), this.Site);

                            // If the verification of actual values success, then requirement MS-ASCAL_R428 can be captured directly.
                            // Verify MS-ASCAL requirement: MS-ASCAL_R428
                            Site.CaptureRequirement(
                                428,
                                @"[In Sensitivity] The value of the Sensitivity element MUST be one of the values[0,1,2,3] listed in the following table.");

                            this.VerifyUnsignedByteDataType(syncResponse.AddElements[i].Calendar.Sensitivity);
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R438");

                        // Verify MS-ASCAL requirement: MS-ASCAL_R438
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            438,
                            @"[In StartTime] The value of this element is a string data type, represented as a Compact DateTime ([MS-ASDTYPE] section 2.6.5).");

                        this.VerifyCompactDateTimeDataType();

                        if (null != syncResponse.AddElements[i].Calendar.Recurrence)
                        {
                            if (activeSyncProtocolVersion.Equals("2.5")
                                || activeSyncProtocolVersion.Equals("12.0")
                                || activeSyncProtocolVersion.Equals("12.1")
                                || activeSyncProtocolVersion.Equals("14.0")
                                || activeSyncProtocolVersion.Equals("14.1"))
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R374");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R374
                                // If schema is verified this requirement can be captured.
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    374,
                                    @"[In Recurrence] [The Recurrence element can have the following child elements:]Type (section 2.2.2.43): One instance of this element is required in protocol version 2.5, 12.0, 12.1, 14.0 and 14.1.");

                            }

                            this.VerifyContainerDataType();

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R448");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R448
                            // Since R374's been verified, and the schema is verified too, this requirement can be captured.
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                448,
                                @"[In Type] The Type element is a required child element of the Recurrence element (section 2.2.2.35) that specifies the type of the recurrence.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R45011");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R45011
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                45011,
                                @"[In Type] A command response has only one Type child element per Recurrence element.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R452");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R452
                            // The type is a required element which defined in schema.
                            // so if the schema is correct ,this element is never null.
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                452,
                                @"[In Type] The value of this element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R453, the value of Type is {0}", syncResponse.AddElements[i].Calendar.Recurrence.Type);

                            string[] expectedValues = new string[] { "0", "1", "2", "3", "5", "6" };
                            Common.VerifyActualValues("Type", expectedValues, syncResponse.AddElements[i].Calendar.Recurrence.Type.ToString(), this.Site);

                            // If the verification of actual values success, then requirement MS-ASCAL_R453 can be captured directly.
                            // Verify MS-ASCAL requirement: MS-ASCAL_R453
                            Site.CaptureRequirement(
                                453,
                                @"[In Type] The value of the Type element MUST be one of the values[0,1,2,3,5,6] listed in the following table.");

                            this.VerifyUnsignedByteDataType(syncResponse.AddElements[i].Calendar.Recurrence.Type);

                            if (syncResponse.AddElements[i].Calendar.Recurrence.Type == 3 || syncResponse.AddElements[i].Calendar.Recurrence.Type == 6)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R471");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R471
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    471,
                                    @"[In WeekOfMonth] The WeekOfMonth element is a child element of the Recurrence element (section 2.2.2.35) that specifies either the week of the month or the day of the month for the recurrence, depending on the value of the Type element.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R475");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R475
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    475,
                                    @"[In WeekOfMonth] The value of this element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R47411");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R47411
                                // Since R475 is verified, this requirement can be captured.
                                // If schema is verified successfully, this requirement can be captured
                                Site.CaptureRequirement(
                                    47411,
                                    @"[In WeekOfMonth] A command response has a maximum of one WeekOfMonth child element per Recurrence element.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R476. The actual value of WeekOfMonth element is {0}.", syncResponse.AddElements[i].Calendar.Recurrence.WeekOfMonth);

                                // Verify MS-ASCAL requirement: MS-ASCAL_R476
                                Site.CaptureRequirementIfIsTrue(
                                    syncResponse.AddElements[i].Calendar.Recurrence.WeekOfMonth >= 1 && syncResponse.AddElements[i].Calendar.Recurrence.WeekOfMonth <= 5,
                                    476,
                                    @"[In WeekOfMonth] The value of the WeekOfMonth element MUST be between 1 and 5.");

                                this.VerifyUnsignedByteDataType(syncResponse.AddElements[i].Calendar.Recurrence.WeekOfMonth);
                            }
                        }

                        if (null != syncResponse.AddElements[i].Calendar.UID)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R462");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R462
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                462,
                                @"[In UID] The value of the UID element is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R463. The actual length of UID element is {0}.", syncResponse.AddElements[i].Calendar.UID.ToCharArray().Length);

                            // Verify MS-ASCAL requirement: MS-ASCAL_R463
                            Site.CaptureRequirementIfIsTrue(
                                syncResponse.AddElements[i].Calendar.UID.ToCharArray().Length <= 300,
                                463,
                                @"[In UID] The maximum length of this element is 300 characters.");

                            this.VerifyStringDataType();
                        }

                        if (!activeSyncProtocolVersion.Equals("12.1"))
                        {
                            if (null != syncResponse.AddElements[i].Calendar.DisallowNewTimeProposal)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R216");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R216
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    216,
                                    @"[In DisallowNewTimeProposal] The value of the DisallowNewTimeProposal element is a boolean data type, as specified in [MS-ASDTYPE] section 2.1.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R215");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R215
                                // It can be verified by schema validation, if the validation is successful.
                                // This requirement can be verified.
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    215,
                                    @"[In DisallowNewTimeProposal] A command response contains one DisallowNewTimeProposal element per response.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R21111");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R21111
                                // Since R215 is captured, this requirement can be captured too.
                                Site.CaptureRequirement(
                                    21111,
                                    @"[In DisallowNewTimeProposal] The DisallowNewTimeProposal<4> element specifies whether a meeting request recipient can propose a new time for the scheduled meeting.");

                                this.VerifyBooleanDataType(syncResponse.AddElements[i].Calendar.DisallowNewTimeProposal);
                            }
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Exceptions)
                        {
                            if (null != syncResponse.AddElements[i].Calendar.Exceptions.Exception)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R242");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R242
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    242,
                                    @"[In Exception] It[ Exception] is a child element of the Exceptions element (section 2.2.2.20).");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "The Exceptions element actually has {0} Exception child elements.", syncResponse.AddElements[i].Calendar.Exceptions.Exception.Length);

                                // Verify MS-ASCAL requirement: MS-ASCAL_R24311
                                Site.CaptureRequirementIfIsTrue(
                                    syncResponse.AddElements[i].Calendar.Exceptions.Exception.Length >= 0 && syncResponse.AddElements[i].Calendar.Exceptions.Exception.Length <= 256,
                                    24311,
                                    @"[In Exception] A command response has between zero and 256 Exception child elements per Exceptions element.");

                                this.VerifyContainerDataType();

                                foreach (ExceptionsException exception in syncResponse.AddElements[i].Calendar.Exceptions.Exception)
                                {
                                    if (exception.DeletedSpecified)
                                    {
                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R209");

                                        // Verify MS-ASCAL requirement: MS-ASCAL_R209
                                        Site.CaptureRequirementIfIsTrue(
                                            this.activeSyncClient.ValidationResult,
                                            209,
                                            @"[In Deleted] The value of this element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");

                                        this.VerifyUnsignedByteDataType(exception.Deleted);
                                    }

                                    // Element ExceptionStartTime is supported when protocol version is 12.1/14.0/14.1
                                    if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "12.1"
                                        || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.0"
                                        || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.1")
                                    { 
                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R264");

                                        // Verify MS-ASCAL requirement: MS-ASCAL_R264
                                        // If ExceptionStartTime exists, this requirement can be captured.
                                        Site.CaptureRequirementIfIsNotNull(
                                            exception.ExceptionStartTime,
                                            264,
                                            @"[In ExceptionStartTime] The ExceptionStartTime element is a required child element of the Exception element (section 2.2.2.19) that specifies the original start time of the occurrence that the exception is replacing in the recurring series.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2106");

                                        // Verify MS-ASCAL requirement: MS-ASCAL_R2106
                                        Site.CaptureRequirementIfIsNotNull(
                                            exception.ExceptionStartTime,
                                            2106,
                                            @"[In Exception] The ExceptionStartTime element is a required child element of the the Exception element only when protocol version 2.5, 12.0, 12.1, 14.0, or 14.1 is used.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R246");

                                        // Verify MS-ASCAL requirement: MS-ASCAL_R246
                                        // Since R264 is captured, this requirement can be captured too.
                                        Site.CaptureRequirement(
                                            246,
                                            @"[In Exception][The Exception element can have the following child elements:] ExceptionStartTime (section 2.2.2.21): One instance of this element is required.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R268");

                                        // Verify MS-ASCAL requirement: MS-ASCAL_R268
                                        // Since R264 is captured, this requirement can be captured too.
                                        Site.CaptureRequirement(
                                            268,
                                            @"[In ExceptionStartTime] The value of the ExceptionStartTime element is a string data type, represented as a Compact DateTime ([MS-ASDTYPE] section 2.6.5).");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R26711");

                                        // Verify MS-ASCAL requirement: MS-ASCAL_R26711
                                        // Since R264 is captured, this requirement can be captured too.
                                        Site.CaptureRequirement(
                                            26711,
                                            @"[In ExceptionStartTime] A command response has only one ExceptionStartTime child element per Exception element.");

                                        this.VerifyCompactDateTimeDataType();                                    
                                    }

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R8111");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R8111
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        8111,
                                        @"[In AllDayEvent] A command response has a maximum of one AllDayEvent child element per Exception element.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R567");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R567
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        567,
                                        @"[In AppointmentReplyTime][A command response has] a maximum of one AppointmentReplyTime child element per Exception element.");

                                    if (null != exception.Attendees)
                                    {
                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R10711");

                                        // Verify MS-ASCAL requirement: MS-ASCAL_R10711
                                        Site.CaptureRequirementIfIsTrue(
                                            this.activeSyncClient.ValidationResult,
                                            10711,
                                            @"[In Attendees] A command response has a maximum of one Attendees child element per Exception element.");
                                    }

                                    if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "12.1")
                                    {
                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2023");

                                        // Verify MS-ASCAL requirement: MS-ASCAL_R2023
                                        Site.CaptureRequirementIfIsNull(
                                            exception.Attendees,
                                            2023,
                                            @"[In Attendees] When protocol version 2.5, 12.0, or 12.1 is used, the Attendees element is not supported as a child element of the Exception element.");
                                    }

                                    if (null != exception.Body)
                                    {
                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R12711");

                                        // Verify MS-ASCAL requirement: MS-ASCAL_R12711
                                        Site.CaptureRequirementIfIsTrue(
                                            this.activeSyncClient.ValidationResult,
                                            12711,
                                            @"[In Body (AirSyncBase Namespace)] A command response has a maximum of one airsyncbase:Body child element per Exception element.");
                                    }

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R22211");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R22211
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        22211,
                                        @"[In DtStamp] A command response has a maximum of one DtStamp child element per Exception element.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R23711");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R23711
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        23711,
                                        @"[In EndTime] A command response has a maximum of one EndTime child element per Exception element.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R29711");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R29711
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        29711,
                                        @"[In Location] A command response has a maximum of one Location child element per Exception element.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R39311");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R39311
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        39311,
                                        @"[In Reminder] A command response has a maximum of one Reminder child element per Exception element.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R40511");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R40511
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        40511,
                                        @"[In ResponseType] A command response has [a maximum of one top-level ResponseType element per response, and] a maximum of one ResponseType child element per Exception element.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R42411");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R42411
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        42411,
                                        @"[In Sensitivity] A command response has a maximum of one Sensitivity child element per Exception element.");

                                    if (null != exception.OnlineMeetingConfLink)
                                    {
                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R35111");

                                        // Verify MS-ASCAL requirement: MS-ASCAL_R35111
                                        Site.CaptureRequirementIfIsTrue(
                                            this.activeSyncClient.ValidationResult,
                                            35111,
                                            @"[In OnlineMeetingConfLink] The value of the OnlineMeetingConfLink element is either a GRUU as specified in [MS-SIPRE], [or an empty tag when included as a child of the Exception element].");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R34812");

                                        // Verify MS-ASCAL requirement: MS-ASCAL_R34812
                                        // This can be verified by schema validation and if MS-ASCAL_R35111 is verified, this can be captured.
                                        Site.CaptureRequirement(
                                            34812,
                                            @"[In OnlineMeetingConfLink] A command response has a maximum of one OnlineMeetingConfLink child element per Exception element.");

                                        if (0 == exception.OnlineMeetingConfLink.Length)
                                        {
                                            // Add the debug information
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R351");

                                            // Verify MS-ASCAL requirement: MS-ASCAL_R351
                                            Site.CaptureRequirementIfIsTrue(
                                                this.activeSyncClient.ValidationResult,
                                                351,
                                                @"[In OnlineMeetingConfLink] The value of the OnlineMeetingConfLink element is [either a GRUU as specified in [MS-SIPRE], or] an empty tag when included as a child of the Exception element.");
                                        }
                                    }

                                    if (!activeSyncProtocolVersion.Equals("12.1"))
                                    {
                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R402");

                                        // Verify MS-ASCAL requirement: MS-ASCAL_R402
                                        Site.CaptureRequirementIfIsTrue(
                                            this.activeSyncClient.ValidationResult,
                                            402,
                                            @"[In ResponseType] The ResponseType<18> element is an optional child element of the Exception element (section 2.2.2.19).");
                                    }
                                }
                            }

                            this.VerifyContainerDataType();
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Attendees)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R10411");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R10411
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                10411,
                                @"[In Attendees] As a top-level element of the Calendar class, the Attendees element specifies the collection of attendees for the calendar item.");

                            if (null != syncResponse.AddElements[i].Calendar.Attendees.Attendee)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R99");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R99
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    99,
                                    @"[In Attendee] It[Attendee] is a child element of the Attendees element (section 2.2.2.4).");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R98011");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R98011
                                // Since R99 is captured, this requirement can be captured too.
                                Site.CaptureRequirement(
                                    98011,
                                    @"[In Attendee] The Attendee element specifies an attendee who is invited to the event.");

                                foreach (AttendeesAttendee attendee in syncResponse.AddElements[i].Calendar.Attendees.Attendee)
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R225");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R225
                                    // If Email exists, this requirement can be captured.
                                    Site.CaptureRequirementIfIsNotNull(
                                        attendee.Email,
                                        225,
                                        @"[In Email] The Email element is a required child element of the Attendee element (section 2.2.2.3) that specifies the e-mail address of an attendee.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R227");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R227
                                    Site.CaptureRequirementIfIsTrue(
                                        null != attendee.Email && this.activeSyncClient.ValidationResult,
                                        227,
                                        @"[In Email] The value of this element is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R231");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R231
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        231,
                                        @"[In Email] It is recommended that the string format adhere to the format specified in [MS-ASDTYPE] section 2.6.2.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R22811");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R22811
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        22811,
                                        @"[In Email] A command response has only Email child element per Attendee element.");
                                    
                                    this.VerifyStringDataType();

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R329");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R329
                                    // If Name exists, this requirement can be captured.
                                    Site.CaptureRequirementIfIsNotNull(
                                        attendee.Name,
                                        329,
                                        @"[In Name] The Name element is a required child element of the Attendee element (section 2.2.2.3) that specifies the name of an attendee.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R331");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R331
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        331,
                                        @"[In Name] The value of this element[Name] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R33211");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R33211
                                    // Since R331 is captured, this requirement can be captured too.
                                    Site.CaptureRequirement(
                                        33211,
                                        @"[In Name] A command response has only one Name child element per Attendee element.");

                                    this.VerifyStringDataType();

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R336");

                                    // Verify MS-ASCAL requirement: MS-ASCAL_R336
                                    Site.CaptureRequirementIfIsTrue(
                                        this.activeSyncClient.ValidationResult,
                                        336,
                                        @"[In NativeBodyType] The value of this element is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.7.");
                                }
                            }
                        }

                        if (null != syncResponse.AddElements[i].Calendar.Location)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R300");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R300
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                300,
                                @"[In Location] The value of this element is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                            this.VerifyStringDataType();
                        }

                        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "16.0")
                        {
                            if (syncResponse.AddElements[i].Calendar.Location1.DisplayName != null)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2135");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R2135
                                Site.CaptureRequirementIfIsNull(
                                    syncResponse.AddElements[i].Calendar.Location,
                                    2135,
                                    @"[In Location] The airsyncbase:Location element ([MS-ASAIRS] section 2.2.2.27) is used instead of the calendar:Location element in protocol version 16.0 [and 16.1].");
                            }
                        }

                        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "16.1")
                        {
                            if (syncResponse.AddElements[i].Calendar.Location1.DisplayName != null)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2135001");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R2135001
                                Site.CaptureRequirementIfIsNull(
                                    syncResponse.AddElements[i].Calendar.Location,
                                    2135001,
                                    @"[In Location] The airsyncbase:Location element ([MS-ASAIRS] section 2.2.2.27) is used instead of the calendar:Location element in protocol version [16.0 and] 16.1.");
                            }
                        }
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R235");

                        // Verify MS-ASCAL requirement: MS-ASCAL_R235
                        // If EndTime is not null and the value of the AllyDayEvent is set as 1 successfully in request and EndTime can be got in the response
                        // so this requirement can be verified.
                        Site.CaptureRequirementIfIsNotNull(
                            syncResponse.AddElements[i].Calendar.EndTime,
                            235,
                            @"[In EndTime] The EndTime element MUST be present in the response as a top-level element, even if the value of the AllDayEvent element (section 2.2.2.1) is 1.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R240");

                        // Verify MS-ASCAL requirement: MS-ASCAL_R240
                        // This can be verified when EndTime is not null and schema validation is passed.
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            240,
                            @"[In EndTime] The value of this element is a string data type, represented as a Compact DateTime ([MS-ASDTYPE] section 2.6.5).");

                        this.VerifyCompactDateTimeDataType();

                        if (null != syncResponse.AddElements[i].Calendar.OrganizerEmail)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R364");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R364  
                            Site.CaptureRequirementIfIsTrue(
                                RFC822AddressParser.IsValidAddress(syncResponse.AddElements[i].Calendar.OrganizerEmail),
                                364,
                                @"[In OrganizerEmail] The value of the OrganizerEmail element is a string ([MS-ASDTYPE] section 2.6) in valid e-mail address format, as specified in [MS-ASDTYPE] section 2.6.2.");

                            this.VerifyStringDataType();
                            this.VerifyEmailAddressDataType(syncResponse.AddElements[i].Calendar.OrganizerEmail);
                        }

                        if (null != syncResponse.AddElements[i].Calendar.OrganizerName)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R369");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R369
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                369,
                                @"[In OrganizerName] The value of this element is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                            this.VerifyStringDataType();

                            if (activeSyncProtocolVersion.Equals("16.0"))
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2175");

                                //Verify MS-ASCAL requirment: MS-ASCAL_R2175
                                Site.CaptureRequirementIfAreEqual<string>(
                                    Common.GetConfigurationPropertyValue("OrganizerUserName", this.Site),
                                    syncResponse.AddElements[i].Calendar.OrganizerName,
                                    2175,
                                    @"[In OrganizerName] [When protocol version 16.0 is used, the client MUST NOT include the OrganizerName element in command requests and] the server will use the name of the current user."
                                    );
                            }

                            if (activeSyncProtocolVersion.Equals("16.1"))
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2175001");

                                //Verify MS-ASCAL requirment: MS-ASCAL_R2175001
                                Site.CaptureRequirementIfAreEqual<string>(
                                    Common.GetConfigurationPropertyValue("OrganizerUserName", this.Site),
                                    syncResponse.AddElements[i].Calendar.OrganizerName,
                                    2175001,
                                    @"[In OrganizerName] [When protocol version 16.1 is used, the client MUST NOT include the OrganizerName element in command requests and] the server will use the name of the current user."
                                    );
                            }

                        }

                        if (!activeSyncProtocolVersion.Equals("12.1"))
                        {
                            if (null != syncResponse.AddElements[i].Calendar.ResponseRequested)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R398");

                                // Verify MS-ASCAL requirement: MS-ASCAL_R398
                                Site.CaptureRequirementIfIsTrue(
                                    this.activeSyncClient.ValidationResult,
                                    398,
                                    @"[In ResponseRequested] The value of the ResponseRequested element is a boolean data type, as specified in [MS-ASDTYPE] section 2.1.");

                                this.VerifyBooleanDataType(syncResponse.AddElements[i].Calendar.ResponseRequested);
                            }

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R396");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R396
                            // When ResponseRequested is null, this requirement can be captured.
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                396,
                                @"[In ResponseRequested] The ResponseRequested<16> element is an optional element.");
                        }

                        if (syncResponse.AddElements[i].Calendar.Timezone != null)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R447");

                            // Verify MS-ASCAL requirement: MS-ASCAL_R447
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                447,
                                @"[In Timezone] The value of the Timezone element is a TimeZone data type, as specified in [MS-ASDTYPE] section 2.6.4.");

                            this.VerifyTimeZoneDataType();
                        }
                    }
                }
            }

            // Add the debug information        
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R538");

            // Verify MS-ASCAL requirement: MS-ASCAL_R538
            // If the schema validation is successful
            // it means that Sync command is operated successfully.
            // so this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                538,
                @"[In Sync Command Response] Top-level Calendar class elements, as specified in section 2.2.2, can be included in a Sync command response as child elements of the airsync:ApplicationData element ([MS-ASCMD] section 2.2.3.11) within either an airsync:Add element ([MS-ASCMD] section 2.2.3.7.2) or an airsync:Change element ([MS-ASCMD] section 2.2.3.24) in the Sync command response.");
        }

        /// <summary>
        /// This method is used to verify the ItemOperation response related requirements.
        /// </summary>
        /// <param name="itemOperationResponse">Specified ItemOperationsStore result returned from the server</param>
        private void VerifyItemOperationsResponse(ItemOperationsStore itemOperationResponse)
        {
            if (null != itemOperationResponse.Items)
            {
                for (int i = itemOperationResponse.Items.Count - 1; i >= 0; i--)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R529");

                    // Verify MS-ASCAL requirement: MS-ASCAL_R529
                    // If the Calendar exists, it means ItemOperations Command Response can have any of the elements which belong to Calendar class.
                    // so this requirement can be verified.
                    Site.CaptureRequirementIfIsNotNull(
                        itemOperationResponse.Items[i].Calendar,
                        529,
                        @"[In ItemOperations Command Response] Any of the elements that belong to the Calendar class, as specified in section 2.2.2, can be included in an ItemOperations command response.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R531");

                    // Verify MS-ASCAL requirement: MS-ASCAL_R531
                    // If Calendar exists, this requirement can be captured.
                    Site.CaptureRequirementIfIsNotNull(
                        itemOperationResponse.Items[i].Calendar,
                        531,
                        @"[In ItemOperations Command Response] Top-level Calendar class elements, as specified in section 2.2.2, MUST be returned as child elements of the itemoperations:Properties element ([MS-ASCMD] section 2.2.3.128) in the ItemOperations command response.");
                }
            }
        }

        /// <summary>
        /// This method is used to verify the Search response related requirements.
        /// </summary>
        /// <param name="searchResponse">Specified SearchStore result returned from the server</param>
        private void VerifySearchCommandResponse(SearchStore searchResponse)
        {
            Site.Assert.IsNotNull(searchResponse.Results, "Search results should not be null.");

            foreach (SearchItem itemInSearch in searchResponse.Results)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R534");

                // Verify MS-ASCAL requirement: MS-ASCAL_R534
                // If the calendar class exists, then any of its elements can be included, so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    itemInSearch.Calendar,
                    534,
                    @"[In Search Command Response] Any of the elements that belong to the Calendar class, as specified in section 2.2.2, can be included in a Search command response.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R535");

                // Since R534 is captured, then this requirement can be captured too.
                // If schema is verified successfully, this requirement can be captured
                Site.CaptureRequirement(
                    535,
                    @"[In Search Command Response] Top-level Calendar class elements MUST be returned as child elements of the search:Properties element ([MS-ASCMD] section 2.2.3.128) in the Search command response.");
            }
        }

        /// <summary>
        /// This method is used to verify the boolean related requirements.
        /// </summary>
        /// <param name="boolValue">A bool value</param>
        private void VerifyBooleanDataType(bool? boolValue)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R4");

            // If the schema validation is successful, then MS-ASDTYPE_R4 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R4
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                4,
                @"[In boolean Data Type] It [a boolean] is declared as an element with a type attribute of ""boolean"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R5");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R5
            Site.CaptureRequirementIfIsTrue(
                boolValue.Equals(true) || boolValue.Equals(false),
                "MS-ASDTYPE",
                5,
                @"[In boolean Data Type] The value of a boolean element is an integer whose only valid values are 1 (TRUE) or 0 (FALSE).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R7");

            // ActiveSyncClient encoded boolean data as inline strings, so if response is successfully returned MS-ASDTYPE_R7 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R7
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                7,
                @"[In boolean Data Type] Elements with a boolean data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");
        }

        /// <summary>
        /// This method is used to verify the dateTime related requirements.
        /// </summary>
        private void VerifyCompactDateTimeDataType()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R12211");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R12211
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                12211,
                @"[In Compact DateTime] A Compact DateTime value is a representation of a UTC date and time within an element of type xs:string, as specified in [XMLSCHEMA2/2] section 3.2.1. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R12213");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R12213
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                12213,
                @"[In Compact DateTime] [The format of a Compact DateTime value is specified by the following Augmented Backus-Naur Form (ABNF) notation. ]
                  date_string   = year month day ""T"" hour minute seconds ""Z""
                  year          = 4*DIGIT
                  month         = (""0"" DIGIT) / ""10"" / ""11"" / ""12""
                  day           = (""0"" DIGIT) / (""1"" DIGIT) / (""2"" DIGIT) / ""30"" / ""31""
                  hour          = (""0"" DIGIT) / (""1"" DIGIT) / ""20"" / ""21"" / ""22"" / ""23""
                  minute        = (""0"" DIGIT) / (""1"" DIGIT) / (""2"" DIGIT) / (""3"" DIGIT) / (""4"" DIGIT) / (""5"" DIGIT)
                  seconds       = (""0"" DIGIT) / (""1"" DIGIT) / (""2"" DIGIT) / (""3"" DIGIT) / (""4"" DIGIT) / (""5"" DIGIT)
                  ");
        }

        /// <summary>
        /// This method is used to verify the container related requirements.
        /// </summary>
        private void VerifyContainerDataType()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R8");

            // If the schema validation is successful, then MS-ASDTYPE_R8 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R9
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                8,
                @"[In container Data Type] A container is an XML element that encloses other elements but has no value of its own.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R9");

            // If the schema validation is successful, then MS-ASDTYPE_R9 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R9
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                9,
                @"[In container Data Type] It [container] is a complex type with complex content, as specified in [XMLSCHEMA1/2] section 3.4.2.");
        }

        /// <summary>
        /// This method is used to verify the E-mail address related requirements.
        /// </summary>
        /// <param name="emailAddress">The email address.</param>
        private void VerifyEmailAddressDataType(string emailAddress)
        {
            // If the validation is successful, then MS-ASDTYPE_R99 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R99");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R99
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                99,
                @"[In E-Mail Address] An e-mail address is an unconstrained value of an element of the string type (section 2.6).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R100");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R100
            Site.CaptureRequirementIfIsTrue(
                RFC822AddressParser.IsValidAddress(emailAddress),
                "MS-ASDTYPE",
                100,
                @"[In E-Mail Address] However, a valid individual e-mail address MUST have the following format: ""local-part@domain"".");
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
                @"[In string Data Type] An element of this[string] type is declared as an element with a type attribute of ""string"".");

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
        /// This method is used to verify the TimeZone related requirements.
        /// </summary>
        private void VerifyTimeZoneDataType()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R103");

            // If the schema validation is successful, then MS-ASDTYPE_R103 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R103
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                103,
                @"[In TimeZone] The TimeZone structure is a structure inside of an element of the string type (section 2.6).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R104");

            // If the schema validation is successful, then MS-ASDTYPE_R104 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R104
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                104,
                @"[In TimeZone] Bias (4 bytes): The value of this [Bias]field is a LONG, as specified in [MS-DTYP].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R107");

            // If the schema validation is successful, then MS-ASDTYPE_R107 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R107
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                107,
                @"[In TimeZone] StandardName (64 bytes): The value of this field is an array of 32 WCHARs, as specified in [MS-DTYP].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R108");

            // If the schema validation is successful, then MS-ASDTYPE_R108 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R108
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                108,
                @"[In TimeZone] It [TimeZone]contains an optional description for standard time.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R109");

            // If the schema validation is successful, then MS-ASDTYPE_R109 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R109
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                109,
                @"[In TimeZone] Any unused WCHARs in the array MUST be set to 0x0000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R110");

            // If the schema validation is successful, then MS-ASDTYPE_R110 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R110
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                110,
                @"[In TimeZone] StandardDate (16 bytes): The value of this field is a SYSTEMTIME structure, as specified in [MS-DTYP].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R111");

            // If the schema validation is successful, then MS-ASDTYPE_R111 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R111
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                111,
                @"[In TimeZone] It [TimeZone]contains the date and time when the transition from DST to standard time occurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R112");

            // If the schema validation is successful, then MS-ASDTYPE_R112 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R112
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                112,
                @"[In TimeZone] StandardBias (4 bytes): The value of this field is a LONG.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R113");

            // If the schema validation is successful, then MS-ASDTYPE_R113 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R113
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                113,
                @"[In TimeZone] It[TimeZone] contains the number of minutes to add to the value of the Bias field during standard time.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R114");

            // If the schema validation is successful, then MS-ASDTYPE_R114 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R114
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                114,
                @"[In TimeZone] DaylightName (64 bytes): The value of this field is an array of 32 WCHARs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R115");

            // If the schema validation is successful, then MS-ASDTYPE_R115 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R115
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                115,
                @"[In TimeZone] It [TimeZone] contains an optional description for DST.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R116");

            // If the schema validation is successful, then MS-ASDTYPE_R116 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R116
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                116,
                @"[In TimeZone] Any unused WCHARs in the array MUST be set to 0x0000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R117");

            // If the schema validation is successful, then MS-ASDTYPE_R117 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R117
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                117,
                @"[In TimeZone] DaylightDate (16 bytes): The value of this field is a SYSTEMTIME structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R118");

            // If the schema validation is successful, then MS-ASDTYPE_R118 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R118
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                118,
                @"[In TimeZone] It [TimeZone]contains the date and time when the transition from standard time to DST occurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R119");

            // If the schema validation is successful, then MS-ASDTYPE_R119 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R119
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                119,
                @"[In TimeZone] DaylightBias (4 bytes): The value of this field is a LONG.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R120");

            // If the schema validation is successful, then MS-ASDTYPE_R120 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R120
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                120,
                @"[In TimeZone] It [TimeZone]contains the number of minutes to add to the value of the Bias field during DST.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R121");

            // If the schema validation is successful, then MS-ASDTYPE_R121 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R121
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                121,
                @"[In TimeZone] The TimeZone structure is encoded using base64 encoding prior to being inserted in an XML element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R122");

            // If the schema validation is successful, then MS-ASDTYPE_R122 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R122
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                122,
                @"[In TimeZone] Elements with a TimeZone structure MUST be encoded and transmitted as [WBXML1.2] inline strings.");
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
                @"[In unsignedByte Data Type] Elements of this type[unsignedByte type] are declared with an element whose type attribute is set to ""unsignedByte"".");
        }

        /// <summary>
        /// Verify WBXML Capture for WBXML process.
        /// </summary>
        private void VerifyWBXMLRequirements()
        {
            // Get WBXML decoded data.MS_ASWBXMLSyntheticImplementation
            Dictionary<string, int> decodedData = this.activeSyncClient.GetMSASWBXMLImplementationInstance().DecodeDataCollection;

            if (null != decodedData)
            {
                // Find Code Page 4.
                foreach (KeyValuePair<string, int> decodeDataItem in decodedData)
                {
                    byte token;
                    string tagName = Common.GetTagName(decodeDataItem.Key, out token);
                    string codePageName = Common.GetCodePageName(decodeDataItem.Key);
                    int codePage = decodeDataItem.Value;
                    bool isValidCodePage = codePage >= 0 && codePage <= 24;
                    Site.Assert.IsTrue(isValidCodePage, "Code page value should between 0-24, the actual value is :{0}", codePage);

                    // Capture requirements.
                    if (4 == codePage)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R14");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R14
                        Site.CaptureRequirementIfAreEqual<string>(
                            "calendar",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            14,
                            @"[In Code Pages] [This algorithm supports] [Code page] 4 [that indicates] [XML namespace] Calendar");

                        switch (tagName)
                        {
                            case "Timezone":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R180");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R180
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x05,
                                        token,
                                        "MS-ASWBXML",
                                        180,
                                        @"[In Code Page 4: Calendar] [Tag name] Timezone [Token] 0x05 [supports protocol versions] All");

                                    break;
                                }

                            case "AllDayEvent":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R181");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R181
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x06,
                                        token,
                                        "MS-ASWBXML",
                                        181,
                                        @"[In Code Page 4: Calendar] [Tag name] AllDayEvent [Token] 0x06 [supports protocol versions] All");

                                    break;
                                }

                            case "Attendees":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R182");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R182
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x07,
                                        token,
                                        "MS-ASWBXML",
                                        182,
                                        @"[In Code Page 4: Calendar] [Tag name] Attendees [Token] 0x07 [supports protocol versions] All");

                                    break;
                                }

                            case "Attendee":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R183");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R183
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x08,
                                        token,
                                        "MS-ASWBXML",
                                        183,
                                        @"[In Code Page 4: Calendar] [Tag name] Attendee [Token] 0x08 [supports protocol versions] All");

                                    break;
                                }

                            case "Email":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R184");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R184
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x09,
                                        token,
                                        "MS-ASWBXML",
                                        184,
                                        @"[In Code Page 4: Calendar] [Tag name] Email [Token] 0x09 [supports protocol versions] All");

                                    break;
                                }

                            case "Name":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R185");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R185
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x0A,
                                        token,
                                        "MS-ASWBXML",
                                        185,
                                        @"[In Code Page 4: Calendar] [Tag name] Name [Token] 0x0A [supports protocol versions] All");

                                    break;
                                }

                            case "BusyStatus":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R186");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R186
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x0D,
                                        token,
                                        "MS-ASWBXML",
                                        186,
                                        @"[In Code Page 4: Calendar] [Tag name] BusyStatus [Token] 0x0D [supports protocol versions] All");

                                    break;
                                }

                            case "Categories":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R187");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R187
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x0E,
                                        token,
                                        "MS-ASWBXML",
                                        187,
                                        @"[In Code Page 4: Calendar] [Tag name] Categories [Token] 0x0E [supports protocol versions] All");

                                    break;
                                }

                            case "Category":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R188");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R188
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x0F,
                                        token,
                                        "MS-ASWBXML",
                                        188,
                                        @"[In Code Page 4: Calendar] [Tag name] Category [Token] 0x0F [supports protocol versions] All");

                                    break;
                                }

                            case "DtStamp":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R189");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R189
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x11,
                                        token,
                                        "MS-ASWBXML",
                                        189,
                                        @"[In Code Page 4: Calendar] [Tag name] DtStamp [Token] 0x11 [supports protocol versions] All");

                                    break;
                                }

                            case "EndTime":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R190");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R190
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x12,
                                        token,
                                        "MS-ASWBXML",
                                        190,
                                        @"[In Code Page 4: Calendar] [Tag name] EndTime [Token] 0x12 [supports protocol versions] All");

                                    break;
                                }

                            case "Exception":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R191");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R191
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x13,
                                        token,
                                        "MS-ASWBXML",
                                        191,
                                        @"[In Code Page 4: Calendar] [Tag name] Exception [Token] 0x13 [supports protocol versions] All");

                                    break;
                                }

                            case "Exceptions":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R192");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R192
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x14,
                                        token,
                                        "MS-ASWBXML",
                                        192,
                                        @"[In Code Page 4: Calendar] [Tag name] Exceptions [Token] 0x14 [supports protocol versions] All");

                                    break;
                                }

                            case "Deleted":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R193");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R193
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x15,
                                        token,
                                        "MS-ASWBXML",
                                        193,
                                        @"[In Code Page 4: Calendar] [Tag name] Deleted [Token] 0x15 [supports protocol versions] All");

                                    break;
                                }

                            case "ExceptionStartTime":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R194");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R194
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x16,
                                        token,
                                        "MS-ASWBXML",
                                        194,
                                        @"[In Code Page 4: Calendar] [Tag name] ExceptionStartTime [Token] 0x16 [supports protocol versions] 2.5, 12.0, 12.1, 14.0, 14.1");

                                    break;
                                }

                            case "Location":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R195");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R195
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x17,
                                        token,
                                        "MS-ASWBXML",
                                        195,
                                        @"[In Code Page 4: Calendar] [Tag name] Location - see note 2 following this table [Token] 0x17 [supports protocol versions] 2.5, 12.0, 12.1, 14.0, 14.1");

                                    break;
                                }

                            case "MeetingStatus":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R196");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R196
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x18,
                                        token,
                                        "MS-ASWBXML",
                                        196,
                                        @"[In Code Page 4: Calendar] [Tag name] MeetingStatus [Token] 0x18 [supports protocol versions] All");

                                    break;
                                }

                            case "OrganizerEmail":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R197");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R197
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x19,
                                        token,
                                        "MS-ASWBXML",
                                        197,
                                        @"[In Code Page 4: Calendar] [Tag name] OrganizerEmail [Token] 0x19 [supports protocol versions] All");

                                    break;
                                }

                            case "OrganizerName":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R198");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R198
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x1A,
                                        token,
                                        "MS-ASWBXML",
                                        198,
                                        @"[In Code Page 4: Calendar] [Tag name] OrganizerName [Token] 0x1A [supports protocol versions] All");

                                    break;
                                }

                            case "Recurrence":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R199");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R199
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x1B,
                                        token,
                                        "MS-ASWBXML",
                                        199,
                                        @"[In Code Page 4: Calendar] [Tag name] Recurrence [Token] 0x1B [supports protocol versions] All");

                                    break;
                                }

                            case "Type":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R200");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R200
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x1C,
                                        token,
                                        "MS-ASWBXML",
                                        200,
                                        @"[In Code Page 4: Calendar] [Tag name] Type [Token] 0x1C [supports protocol versions] All");

                                    break;
                                }

                            case "Until":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R201");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R201
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x1D,
                                        token,
                                        "MS-ASWBXML",
                                        201,
                                        @"[In Code Page 4: Calendar] [Tag name] Until [Token] 0x1D [supports protocol versions] All");

                                    break;
                                }

                            case "Occurrences":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R202");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R202
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x1E,
                                        token,
                                        "MS-ASWBXML",
                                        202,
                                        @"[In Code Page 4: Calendar] [Tag name] Occurrences [Token] 0x1E [supports protocol versions] All");

                                    break;
                                }

                            case "Interval":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R203");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R203
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x1F,
                                        token,
                                        "MS-ASWBXML",
                                        203,
                                        @"[In Code Page 4: Calendar] [Tag name] Interval [Token] 0x1F [supports protocol versions] All");

                                    break;
                                }

                            case "DayOfWeek":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R204");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R204
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x20,
                                        token,
                                        "MS-ASWBXML",
                                        204,
                                        @"[In Code Page 4: Calendar] [Tag name] DayOfWeek [Token] 0x20 [supports protocol versions] All");

                                    break;
                                }

                            case "DayOfMonth":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R205");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R205
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x21,
                                        token,
                                        "MS-ASWBXML",
                                        205,
                                        @"[In Code Page 4: Calendar] [Tag name] DayOfMonth [Token] 0x21 [supports protocol versions] All");

                                    break;
                                }

                            case "WeekOfMonth":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R206");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R206
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x22,
                                        token,
                                        "MS-ASWBXML",
                                        206,
                                        @"[In Code Page 4: Calendar] [Tag name] WeekOfMonth [Token] 0x22 [supports protocol versions] All");

                                    break;
                                }

                            case "MonthOfYear":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R207");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R207
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x23,
                                        token,
                                        "MS-ASWBXML",
                                        207,
                                        @"[In Code Page 4: Calendar] [Tag name] MonthOfYear [Token] 0x23 [supports protocol versions] All");

                                    break;
                                }

                            case "Reminder":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R208");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R208
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x24,
                                        token,
                                        "MS-ASWBXML",
                                        208,
                                        @"[In Code Page 4: Calendar] [Tag name] Reminder [Token] 0x24 [supports protocol versions] All");

                                    break;
                                }

                            case "Sensitivity":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R209");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R209
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x25,
                                        token,
                                        "MS-ASWBXML",
                                        209,
                                        @"[In Code Page 4: Calendar] [Tag name] Sensitivity [Token] 0x25 [supports protocol versions] All");

                                    break;
                                }

                            case "Subject":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R210");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R210
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x26,
                                        token,
                                        "MS-ASWBXML",
                                        210,
                                        @"[In Code Page 4: Calendar] [Tag name] Subject [Token]  0x26 [supports protocol versions] All");

                                    break;
                                }

                            case "StartTime":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R211");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R211
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x27,
                                        token,
                                        "MS-ASWBXML",
                                        211,
                                        @"[In Code Page 4: Calendar] [Tag name] StartTime [Token] 0x27 [supports protocol versions] All");

                                    break;
                                }

                            case "UID":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R212");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R212
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x28,
                                        token,
                                        "MS-ASWBXML",
                                        212,
                                        @"[In Code Page 4: Calendar] [Tag name] UID [Token] 0x28 [supports protocol versions] All");

                                    break;
                                }

                            case "AttendeeStatus":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R213");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R213
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x29,
                                        token,
                                        "MS-ASWBXML",
                                        213,
                                        @"[In Code Page 4: Calendar] [Tag name] AttendeeStatus [Token] 0x29 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                                    break;
                                }

                            case "AttendeeType":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R214");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R214
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x2A,
                                        token,
                                        "MS-ASWBXML",
                                        214,
                                        @"[In Code Page 4: Calendar] [Tag name] AttendeeType [Token] 0x2A [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                                    break;
                                }

                            case "DisallowNewTimeProposal":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R215");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R215
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x33,
                                        token,
                                        "MS-ASWBXML",
                                        215,
                                        @"[In Code Page 4: Calendar] [Tag name] DisallowNewTimeProposal [Token] 0x33 [supports protocol versions] 14.0, 14.1, 16.0");

                                    break;
                                }

                            case "ResponseRequested":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R217");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R217
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x34,
                                        token,
                                        "MS-ASWBXML",
                                        217,
                                        @"[In Code Page 4: Calendar] [Tag name] ResponseRequested [Token] 0x34 [supports protocol versions] 14.0, 14.1, 16.0");

                                    break;
                                }

                            case "AppointmentReplyTime":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R218");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R218
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x35,
                                        token,
                                        "MS-ASWBXML",
                                        218,
                                        @"[In Code Page 4: Calendar] [Tag name] AppointmentReplyTime [Token] 0x35 [supports protocol versions] 14.0, 14.1, 16.0");

                                    break;
                                }

                            case "ResponseType":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R219");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R219
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x36,
                                        token,
                                        "MS-ASWBXML",
                                        219,
                                        @"[In Code Page 4: Calendar] [Tag name] ResponseType [Token] 0x36 [supports protocol versions] 14.0, 14.1, 16.0");

                                    break;
                                }

                            case "CalendarType":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R220");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R220
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x37,
                                        token,
                                        "MS-ASWBXML",
                                        220,
                                        @"[In Code Page 4: Calendar] [Tag name] CalendarType [Token] 0x37 [supports protocol versions] 14.0, 14.1, 16.0");

                                    break;
                                }

                            case "IsLeapMonth":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R221");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R221
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x38,
                                        token,
                                        "MS-ASWBXML",
                                        221,
                                        @"[In Code Page 4: Calendar] [Tag name] IsLeapMonth [Token] 0x38 [supports protocol versions] 14.0, 14.1, 16.0");

                                    break;
                                }

                            case "FirstDayOfWeek":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R222");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R222
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x39,
                                        token,
                                        "MS-ASWBXML",
                                        222,
                                        @"[In Code Page 4: Calendar] [Tag name] FirstDayOfWeek [Token] 0x39 [supports protocol versions] 14.1, 16.0");

                                    break;
                                }

                            case "OnlineMeetingConfLink":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R223");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R223
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x3A,
                                        token,
                                        "MS-ASWBXML",
                                        223,
                                        @"[In Code Page 4: Calendar] [Tag name] OnlineMeetingConfLink [Token] 0x3A [supports protocol versions] 14.1, 16.0");

                                    break;
                                }

                            case "OnlineMeetingExternalLink":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R224");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R224
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x3B,
                                        token,
                                        "MS-ASWBXML",
                                        224,
                                        @"[In Code Page 4: Calendar] [Tag name] OnlineMeetingExternalLink [Token] 0x3B [supports protocol versions] 14.1, 16.0");

                                    break;
                                }
                            
                            default:
                                {
                                    Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePage, tagName, token);
                                    break;
                                }
                        }
                    }
                }
            }
        }
    }
}