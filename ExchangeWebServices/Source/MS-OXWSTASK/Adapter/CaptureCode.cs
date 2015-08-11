namespace Microsoft.Protocols.TestSuites.MS_OXWSTASK
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-OXWSTASK.
    /// </summary>
    public partial class MS_OXWSTASKAdapter
    {
        /// <summary>
        /// Verify the SOAP version.
        /// </summary>
        private void VerifySoapVersion()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R1");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R1
            // According to the implementation of adapter, the message is formatted as SOAP 1.1. If the operation is invoked successfully, then this requirement can be verified. 
            Site.CaptureRequirement(
                1,
                @"[In Transport] The SOAP version supported is SOAP 1.1. For details, see [SOAP1.1].");
        }

        /// <summary>
        /// Verify the transport related requirements.
        /// </summary>
        private void VerifyTransportType()
        {
            // Get the transport type
            TransportProtocol transport = (TransportProtocol)Enum.Parse(typeof(TransportProtocol), Common.GetConfigurationPropertyValue("TransportType", Site), true);
            if (transport == TransportProtocol.HTTPS && Common.IsRequirementEnabled(212, Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R212");

                // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R212
                // Because the adapter uses SOAP and HTTPS to communicate with server, if server returns data without exception, this requirement will be captured.
                Site.CaptureRequirement(
                    212,
                    @"[In Appendix C: Product Behavior] Implementation does use secure communications by means of HTTPS, as defined in [RFC2818]. (Exchange Server 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// Verify the CopyItemResponseType structure.
        /// </summary>
        /// <param name="copyItemResponse">The response got from server via CopyItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyCopyItemOperation(CopyItemResponseType copyItemResponse, bool isSchemaValidated)
        {
            // If the validation event return any error or warning, the schema validation is false, 
            // which indicates the schema is not matched with the expected result. 
            // So these requirement can be verified when the isSchemaValidation is true.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R148");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R148            
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                148,
                @"[In CopyItem Operation] The following is the WSDL port type specification for the CopyItem operation. <wsdl:operation name=""CopyItem"">
                        <wsdl:input message=""tns:CopyItemSoapIn"" />
                        <wsdl:output message=""tns:CopyItemSoapOut"" />
            </wsdl:operation>");

            // Proxy handles operation's soap in and soap out, if the server didn't respond the soap out message for the operation,  
            // proxy will fail. Now it didn't fail, that indicates server responds corresponding soap out message for the operation.
            // So the following requirement can be captured.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R150");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R150            
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                150,
                @"[In CopyItem Operation] The following is the WSDL binding specification for the CopyItem operation.<wsdl:operation name=""CopyItem"">
                       <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/CopyItem""/>
                       <wsdl:input>
                          <soap:header message=""tns:CopyItemSoapIn"" part=""Impersonation"" use=""literal""/>
                          <soap:header message=""tns:CopyItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
                          <soap:header message=""tns:CopyItemSoapIn"" part=""RequestVersion"" use=""literal""/>
                          <soap:body parts=""request"" use=""literal""/>
                       </wsdl:input>
                       <wsdl:output>
                          <soap:body parts=""CopyItemResult"" use=""literal""/>
                          <soap:header message=""tns:CopyItemSoapOut"" part=""ServerVersion"" use=""literal""/>
                       </wsdl:output>
                    </wsdl:operation>");

            if (copyItemResponse.ResponseMessages.Items != null && copyItemResponse.ResponseMessages.Items.Length > 0)
            {
                foreach (ItemInfoResponseMessageType itemInfo in copyItemResponse.ResponseMessages.Items)
                {
                    if (itemInfo.ResponseClass == ResponseClassType.Success)
                    {
                        this.VerifyTaskType(isSchemaValidated, itemInfo);
                    }
                }
            }
        }

        /// <summary>
        /// Verify the TaskType properties.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        /// <param name="itemInfo">The task item information.</param>
        private void VerifyTaskType(bool isSchemaValidated, ItemInfoResponseMessageType itemInfo)
        {
            foreach (ItemType item in itemInfo.Items.Items)
            {
                TaskType taskItem = item as TaskType;
                if (taskItem != null)
                {
                    // If the ActualWorkSpecified field of task item is true means ActualWork is specified, then verify the related requirements about ActualWork.
                    if (taskItem.ActualWorkSpecified)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R213");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R213 
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            213,
                            @"[In t:TaskType Complex Type] The type of ActualWork is xs:int [XMLSCHEMA2].");
                    }

                    if (taskItem.BillingInformation != null)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R215");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R215
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            215,
                            @"[In t:TaskType Complex Type] The type of BillingInformation is xs:string [XMLSCHEMA2].");
                    }

                    // If the ChangeCountSpecified field of task item is true means ChangeCount is specified, then verify the related requirements about ChangeCount.
                    if (taskItem.ChangeCountSpecified)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R216");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R216
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            216,
                            @"[In t:TaskType Complex Type] The type of ChangeCount is xs:int.");
                    }

                    if (taskItem.Companies != null)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R217");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R217
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            217,
                            @"[In t:TaskType Complex Type]The type of Companies is t:ArrayOfStringsType ([MS-OXWSCDATA] section 2.2.4.11).");

                        this.VerifyArrayOfStringsType(isSchemaValidated);
                    }

                    // If the CompleteDateSpecified field of task item is true means CompleteDate is specified, then verify the related requirements about CompleteDate.
                    if (taskItem.CompleteDateSpecified)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R218");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R218
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            218,
                            @"[In t:TaskType Complex Type] The type of CompleteDate is xs:dateTime.");
                    }

                    if (taskItem.Contacts != null)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R219");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R219
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            219,
                            @"[In t:TaskType Complex Type] The type of Contacts is t:ArrayOfStringsType.");

                        this.VerifyArrayOfStringsType(isSchemaValidated);
                    }

                    // If the DueDateSpecified field of task item is true means DueDate is specified, then verify the related requirements about DueDate.
                    if (taskItem.DueDateSpecified)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R222");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R222
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            222,
                            @"[In t:TaskType Complex Type] The type of DueDate is xs:dateTime.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R89");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R89
                        bool isVerifyR89 = isSchemaValidated && ((taskItem.DelegationState == TaskDelegateStateType.Accepted)
                                        || (taskItem.DelegationState == TaskDelegateStateType.Declined)
                                        || (taskItem.DelegationState == TaskDelegateStateType.Max)
                                        || (taskItem.DelegationState == TaskDelegateStateType.NoMatch)
                                        || (taskItem.DelegationState == TaskDelegateStateType.Owned)
                                        || (taskItem.DelegationState == TaskDelegateStateType.OwnNew));

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR89,
                            89,
                            @"[In t:TaskDelegateStateType Simple Type] [t:TaskDelegateStateType Simple Type] is defined as  
                                name=""TaskDelegateStateType"">
                                  <xs:restriction
                                    base=""xs:string"">
                                    <xs:enumeration
                                      value=""Accepted""/>
                                    <xs:enumeration
                                      value=""Declined""/>
                                    <xs:enumeration
                                      value=""Max""/>
                                    <xs:enumeration
                                      value=""NoMatch""/>
                                    <xs:enumeration
                                      value=""Owned""/>
                                    <xs:enumeration
                                      value=""OwnNew""/>
                                  </xs:restriction>
                                </xs:simpleType>");
                    }

                    // If the IsCompleteSpecified field of task item is true means IsComplete is specified, then verify the related requirements about IsComplete.
                    if (taskItem.IsCompleteSpecified)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R224");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R224
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            224,
                            @"[In t:TaskType Complex Type] The type of IsComplete is xs:boolean [XMLSCHEMA2].");
                    }

                    // If the IsRecurringSpecified field of task item is true means IsRecurring is specified, then verify the related requirements about IsRecurring.
                    if (taskItem.IsRecurringSpecified)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R225");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R225
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            225,
                            @"[In t:TaskType Complex Type] The type of IsRecurring is xs:boolean.");
                    }

                    if (taskItem.Mileage != null)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R227");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R227
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            227,
                            @"[In t:TaskType Complex Type] The type of Mileage is xs:string.");
                    }

                    // If the PercentCompleteSpecified field of task item is true means PercentComplete is specified, then verify the related requirements about PercentComplete.
                    if (taskItem.PercentCompleteSpecified)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R229");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R229
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            229,
                            @"[In t:TaskType Complex Type] The type of PercentComplete is xs:double [XMLSCHEMA2].");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R60");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R60
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated && taskItem.PercentComplete >= 0 && taskItem.PercentComplete <= 100,
                            60,
                            @"[In t:TaskType Complex Type] PercentComplete: Specifies a double value from 0 through 100 that describes the completion status of a task.");
                    }

                    #region Verify element Recurrence of TaskType

                    if (taskItem.Recurrence != null)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R230");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R230
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            230,
                            @"[In t:TaskType Complex Type] The type of Recurrence is t:TaskRecurrenceType (section 2.2.5.1).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R29");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R29
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            29,
                            @"[In t:TaskRecurrenceType Complex Type] The TaskRecurrenceType complex type specifies the recurrence pattern for tasks.
                                <xs:complexType name=""TaskRecurrenceType"">
                                  <xs:sequence>
                                    <xs:group
                                      ref=""t:TaskRecurrencePatternTypes""/>
                                    <xs:group
                                      ref=""t:RecurrenceRangeTypes""/>
                                  </xs:sequence>
                                </xs:complexType>");
                        if (taskItem.Recurrence.Item is RegeneratingPatternBaseType)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R209");

                            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R209
                            Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                209,
                                @"[In t:RegeneratingPatternBaseType Complex Type] The RegeneratingPatternBaseType complex type extends the IntervalRecurrencePatternBaseType complex type, as specified in [MS-OXWSCDATA] section 2.2.4.36.
                                    <xs:complexType name=""RegeneratingPatternBaseType""
                                      abstract=""true"">
                                      <xs:complexContent>
                                        <xs:extension
                                          base=""t:IntervalRecurrencePatternBaseType""/>
                                      </xs:complexContent>
                                    </xs:complexType>");
                        }

                        if (taskItem.Recurrence.Item1 != null)
                        {
                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1351");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1351        
                            this.Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1351,
                                @"[In t:RecurrenceRangeTypes Group] The group [t:RecurrenceRangeTypes] is defined as follow:
                                    <xs:group name=""t:RecurrenceRangeTypes"">
                                      <xs:sequence>
                                        <xs:choice>
                                          <xs:element name=""NoEndRecurrence""
                                            type=""t:NoEndRecurrenceRangeType""
                                           />
                                          <xs:element name=""EndDateRecurrence""
                                            type=""t:EndDateRecurrenceRangeType""
                                           />
                                          <xs:element name=""NumberedRecurrence""
                                            type=""t:NumberedRecurrenceRangeType""
                                           />
                                        </xs:choice>
                                      </xs:sequence>
                                    </xs:group>");
                        }

                        if (taskItem.Recurrence.Item1 is NumberedRecurrenceRangeType)
                        {
                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1662");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1662      
                            this.Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1662,
                                @"[In t:RecurrenceRangeTypes Group] The element ""NumberedRecurrence"" is""t:NumberedRecurrenceRangeType"" type (section 2.2.4.46).");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1236");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1236         
                            this.Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1236,
                                @"[In t:NumberedRecurrenceRangeType Complex Type] The type [NumberedRecurrenceRangeType] is defined as follow:
                                    <xs:complexType name=""NumberedRecurrenceRangeType"">
                                      <xs:complexContent>
                                        <xs:extension
                                          base=""t:RecurrenceRangeBaseType""
                                        >
                                          <xs:sequence>
                                            <xs:element name=""NumberOfOccurrences""
                                              type=""xs:int""
                                             />
                                          </xs:sequence>
                                        </xs:extension>
                                      </xs:complexContent>
                                    </xs:complexType>");
                        }

                        if (taskItem.Recurrence.Item1 is NoEndRecurrenceRangeType)
                        {
                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1660");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1660            
                            this.Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1660,
                                @"[In t:RecurrenceRangeTypes Group] The element ""NoEndRecurrence"" is ""t:NoEndRecurrenceRangeType"" type (section 2.2.4.41).");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1203");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1203      
                            this.Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1203,
                                @"[In t:NoEndRecurrenceRangeType Complex Type] The type [NoEndRecurrenceRangeType] is defined as follow:
                                    <xs:complexType name=""NoEndRecurrenceRangeType"">
                                      <xs:complexContent>
                                        <xs:extension
                                          base=""t:RecurrenceRangeBaseType""
                                         />
                                      </xs:complexContent>
                                    </xs:complexType>");
                        }

                        if (taskItem.Recurrence.Item1 is EndDateRecurrenceRangeType)
                        {
                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1661");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1661         
                            this.Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1661,
                                @"[In t:RecurrenceRangeTypes Group] The element ""EndDateRecurrence"" is ""t:EndDateRecurrenceRangeType"" type (section 2.2.4.28).");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1152");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1152            
                            this.Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1152,
                                @"[In t:EndDateRecurrenceRangeType Complex Type] The type [EndDateRecurrenceRangeType] is defined as follow:
                                     <xs:complexType name=""EndDateRecurrenceRangeType"">
                                      <xs:complexContent>
                                        <xs:extension
                                          base=""t:RecurrenceRangeBaseType""
                                        >
                                          <xs:sequence>
                                            <xs:element name=""EndDate""
                                              type=""xs:date""
                                             />
                                          </xs:sequence>
                                        </xs:extension>
                                      </xs:complexContent>
                                    </xs:complexType>");
                        }

                        if (taskItem.Recurrence.Item is IntervalRecurrencePatternBaseType)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1431");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1431            
                            Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1431,
                                @"[In t:IntervalRecurrencePatternBaseType Complex Type] The type [IntervalRecurrencePatternBaseType] is defined as follow:
                                    <xs:complexType name=""IntervalRecurrencePatternBaseType""
                                      abstract=""true""
                                    >
                                      <xs:complexContent>
                                        <xs:extension
                                          base=""t:RecurrencePatternBaseType""
                                        >
                                          <xs:sequence>
                                            <xs:element name=""Interval""
                                              type=""xs:int""
                                             />
                                          </xs:sequence>
                                        </xs:extension>
                                      </xs:complexContent>
                                    </xs:complexType>");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1632");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1632            
                            this.Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1632,
                                @"[In t:IntervalRecurrencePatternBaseType Complex Type] The element ""Interval"" is ""xs:int"" type ([XMLSCHEMA2]).");
                        }

                        if (taskItem.Recurrence.Item is WeeklyRecurrencePatternType)
                        {
                            WeeklyRecurrencePatternType recPattern = taskItem.Recurrence.Item as WeeklyRecurrencePatternType;

                            if (recPattern.DaysOfWeek != null)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R52");

                                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R52
                                Site.CaptureRequirementIfIsTrue(
                                    isSchemaValidated,
                                    "MS-OXWSCDATA",
                                    52,
                                    @"[In t:DaysOfWeekType Simple Type] The type [DaysOfWeekType] is defined as follow:<xs:simpleType name=""DaysOfWeekType"">
                                        <xs:list>
                                            <xs:itemType name=""t:DayOfWeekType""/>
                                        </xs:list>
                                    </xs:simpleType>");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R55");

                                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R55
                                Site.CaptureRequirementIfIsTrue(
                                    isSchemaValidated,
                                    "MS-OXWSCDATA",
                                    55,
                                    @"[In t:DaysOfWeekType Simple Type] The syntax [DaysOfWeekType] is defined as follow:
                                        <xs:simpleType name=""DaysOfWeekType"">
                                            <xs:list itemType=""t:DayOfWeekType""/>
                                        </xs:simpleType>");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R2110");

                                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R2110
                                Site.CaptureRequirementIfAreNotEqual<string>(
                                    DayOfWeekType.Day.ToString(),
                                    recPattern.DaysOfWeek,
                                    "MS-OXWSCDATA",
                                    2110,
                                    @"[In t:DayOfWeekType Simple Type] This value MUST NOT be used in the t:WeeklyRecurrencePatternType (section 2.2.4.64) complex types.");

                                // Add the debug information
                                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1653");

                                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1653
                                this.Site.CaptureRequirementIfIsTrue(
                                    isSchemaValidated,
                                    "MS-OXWSCDATA",
                                    1653,
                                    @"[In t:WeeklyRecurrencePatternType Complex Type] The element ""DaysOfWeek"" is ""t:DaysOfWeekType"" type (section 2.2.3.6).");
                            }
                        }

                        if (taskItem.Recurrence.Item is RelativeMonthlyRecurrencePatternType)
                        {
                            this.VerifyDayOfWeekType(isSchemaValidated);
                            this.VerifyDayOfWeekIndexType(isSchemaValidated);

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1638");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1638          
                            this.Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1638,
                                @"[In t:RelativeMonthlyRecurrencePatternType Complex Type] The element ""DaysOfWeek"" is ""t:DayOfWeekType"" type (section 2.2.3.5).");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1639");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1639
                            this.Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1639,
                                @"[In t:RelativeMonthlyRecurrencePatternType Complex Type] The element ""DayOfWeekIndex"" is ""t:DayOfWeekIndexType"" type(section 2.2.3.4).");
                        }

                        if (taskItem.Recurrence.Item is RelativeYearlyRecurrencePatternType)
                        {
                            RelativeYearlyRecurrencePatternType recPattern = taskItem.Recurrence.Item as RelativeYearlyRecurrencePatternType;

                            this.VerifyDayOfWeekType(isSchemaValidated);
                            this.VerifyDayOfWeekIndexType(isSchemaValidated);

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1491");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1491
                            Site.CaptureRequirementIfAreNotEqual<string>(
                                DayOfWeekType.Day.ToString(),
                                recPattern.DaysOfWeek,
                                "MS-OXWSCDATA",
                                1491,
                                @"[In t:DayOfWeekType Simple Type] This value MUST NOT be used in the t:RelativeYearlyRecurrencePatternType (section 2.2.4.53) complex types.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1492");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1492
                            Site.CaptureRequirementIfAreNotEqual<string>(
                                DayOfWeekType.Weekday.ToString(),
                                recPattern.DaysOfWeek,
                                "MS-OXWSCDATA",
                                1492,
                                @"[In t:DayOfWeekType Simple Type] This value MUST NOT be used in the t:RelativeYearlyRecurrencePatternType complex types.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1493");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1493
                            Site.CaptureRequirementIfAreNotEqual<string>(
                                DayOfWeekType.WeekendDay.ToString(),
                                recPattern.DaysOfWeek,
                                "MS-OXWSCDATA",
                                1493,
                                @"[In t:DayOfWeekType Simple Type] This value MUST NOT be used in the t:RelativeYearlyRecurrencePatternType complex types.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R174");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R174
                            Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                174,
                                @"[In t:MonthNamesType Simple Type] The type [MonthNamesType] is defined as follow: 
                                    <xs:simpleType name=""MonthNamesType"">
                                      <xs:restriction
                                        base=""xs:string""
                                      >
                                        <xs:enumeration
                                          value=""January""
                                         />
                                        <xs:enumeration
                                          value=""February""
                                         />
                                        <xs:enumeration
                                          value=""March""
                                         />
                                        <xs:enumeration
                                          value=""April""
                                         />
                                        <xs:enumeration
                                          value=""May""
                                         />
                                        <xs:enumeration
                                          value=""June""
                                         />
                                        <xs:enumeration
                                          value=""July""
                                         />
                                        <xs:enumeration
                                          value=""August""
                                         />
                                        <xs:enumeration
                                          value=""September""
                                         />
                                        <xs:enumeration
                                          value=""October""
                                         />
                                        <xs:enumeration
                                          value=""November""
                                         />
                                        <xs:enumeration
                                          value=""December""
                                         />
                                      </xs:restriction>
                                    </xs:simpleType>");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1640");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1640 
                            this.Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1640,
                                @"[In t:RelativeYearlyRecurrencePatternType Complex Type] The element ""DaysOfWeek"" is ""t:DayOfWeekType"" type.");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1641");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1641   
                            this.Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1641,
                                @"[In t:RelativeYearlyRecurrencePatternType Complex Type] The element ""DayOfWeekIndex"" is ""t:DayOfWeekIndexType"" type.");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1642");

                            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1642 
                            this.Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                "MS-OXWSCDATA",
                                1642,
                                @"[In t:RelativeYearlyRecurrencePatternType Complex Type] The element ""Month"" is ""t:MonthNamesType"" type.");
                        }
                    }

                    #endregion

                    // If the StartDateSpecified field of task item is true means StartDate is specified, then verify the related requirements about StartDate.
                    if (taskItem.StartDateSpecified)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R231");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R231
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            231,
                            @"[In t:TaskType Complex Type] The type of StartDate is xs:dateTime.");
                    }

                    #region Verify element Status of TaskType

                    // If the StatusSpecified field of task item is true means Status is specified, then verify the related requirements about Status.
                    if (taskItem.StatusSpecified)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R232");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R232
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            232,
                            @"[In t:TaskType Complex Type] The Type of Status is t:TaskStatusType (section 2.2.5.2).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R100");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R100
                        bool isVerifyR100 = isSchemaValidated && ((taskItem.Status == TaskStatusType.Completed)
                                                           || (taskItem.Status == TaskStatusType.Deferred)
                                                           || (taskItem.Status == TaskStatusType.InProgress)
                                                           || (taskItem.Status == TaskStatusType.NotStarted)
                                                           || (taskItem.Status == TaskStatusType.WaitingOnOthers));

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR100,
                            100,
                            @"[In t:TaskStatusType Simple Type] The TaskStatusType simple type specifies the status of a task item.
                                <xs:simpleType name=""TaskStatusType"">
                                  <xs:restriction
                                    base=""xs:string"">
                                    <xs:enumeration
                                      value=""Completed""/>
                                    <xs:enumeration
                                      value=""Deferred""/>
                                    <xs:enumeration
                                      value=""InProgress""/>
                                    <xs:enumeration
                                      value=""NotStarted""/>
                                    <xs:enumeration
                                      value=""WaitingOnOthers""/>
                                  </xs:restriction>
                                </xs:simpleType>");
                    }

                    #endregion

                    if (taskItem.StatusDescription != null)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R233");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R233
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            233,
                            @"[In t:TaskType Complex Type] The type of StatusDescription is xs:string.");
                    }

                    // If the TotalWorkSpecified field of task item is true means TotalWork is specified, then verify the related requirements about TotalWork.
                    if (taskItem.TotalWorkSpecified)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R234");

                        // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R234
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            234,
                            @"[In t:TaskType Complex Type] The type of TotalWork is xs:int.");
                    }

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R39");

                    // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R39
                    Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        39,
                        @"[In t:TaskType Complex Type] The TaskType complex type extends the ItemType complex type, as specified in [MS-OXWSCORE] section 2.2.4.8.
                            <xs:complexType name=""TaskType"">
                                <xs:complexContent>
                                <xs:extension
                                    base=""t:ItemType"">
                                    <xs:sequence>
                                    <xs:element name=""ActualWork""
                                        type=""xs:int""
                                        minOccurs=""0""/>
                                    <xs:element name=""AssignedTime""
                                        type=""xs:dateTime""
                                        minOccurs=""0""/>
                                    <xs:element name=""BillingInformation""
                                        type=""xs:string""
                                        minOccurs=""0""/>
                                    <xs:element name=""ChangeCount""
                                        type=""xs:int""
                                        minOccurs=""0""/>
                                    <xs:element name=""Companies""
                                        type=""t:ArrayOfStringsType""
                                        minOccurs=""0""/>
                                    <xs:element name=""CompleteDate""
                                        type=""xs:dateTime""
                                        minOccurs=""0""/>
                                    <xs:element name=""Contacts""
                                        type=""t:ArrayOfStringsType""
                                        minOccurs=""0""/>
                                    <xs:element name=""DelegationState""
                                        type=""t:TaskDelegateStateType""
                                        minOccurs=""0""/>
                                    <xs:element name=""Delegator""
                                        type=""xs:string""
                                        minOccurs=""0""/>
                                    <xs:element name=""DueDate""
                                        type=""xs:dateTime""
                                        minOccurs=""0""/>
                                    <xs:element name=""IsAssignmentEditable""
                                        type=""xs:int""
                                        minOccurs=""0""/>
                                    <xs:element name=""IsComplete""
                                        type=""xs:boolean""
                                        minOccurs=""0""/>
                                    <xs:element name=""IsRecurring""
                                        type=""xs:boolean""
                                        minOccurs=""0""/>
                                    <xs:element name=""IsTeamTask""
                                        type=""xs:boolean""
                                        minOccurs=""0""/>
                                    <xs:element name=""Mileage""
                                        type=""xs:string""
                                        minOccurs=""0""/>
                                    <xs:element name=""Owner""
                                        type=""xs:string""
                                        minOccurs=""0""/>
                                    <xs:element name=""PercentComplete""
                                        type=""xs:double""
                                        minOccurs=""0""/>
                                    <xs:element name=""Recurrence""
                                        type=""t:TaskRecurrenceType""
                                        minOccurs=""0""/>
                                    <xs:element name=""StartDate""
                                        type=""xs:dateTime""
                                        minOccurs=""0""/>
                                    <xs:element name=""Status""
                                        type=""t:TaskStatusType""
                                        minOccurs=""0""/>
                                    <xs:element name=""StatusDescription""
                                        type=""xs:string""
                                        minOccurs=""0""/>
                                    <xs:element name=""TotalWork""
                                        type=""xs:int""
                                        minOccurs=""0""/>
                                    </xs:sequence>
                                </xs:extension>
                                </xs:complexContent>
                            </xs:complexType>");
                }
                else
                {
                    Site.Assert.Fail("The verified item is not a task item.", null);
                }
            }
        }

        /// <summary>
        /// Verify the DayOfWeekIndexType.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyDayOfWeekIndexType(bool isSchemaValidated)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R25");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R25
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                25,
                @"[In t:DayOfWeekIndexType Simple Type] The type [DayOfWeekIndexType] is defined as follow:
                     <xs:simpleType name=""DayOfWeekIndexType"">
                        <xs:restriction base=""xs:string"">
                            <xs:enumeration value=""First""/>
                            <xs:enumeration value=""Fourth""/>
                            <xs:enumeration value=""Last""/>
                            <xs:enumeration value=""Second""/>
                            <xs:enumeration value=""Third""/>
                        </xs:restriction>
                    </xs:simpleType>");
        }

        /// <summary>
        /// Verify the DayOfWeekType.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyDayOfWeekType(bool isSchemaValidated)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R33");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R33
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                33,
                @"[In t:DayOfWeekType Simple Type] The type [DayOfWeekType] is defined as follow:
                    <xs:simpleType name=""DayOfWeekType"">
                        <xs:restriction base=""xs:string"">
                            <xs:enumeration value=""Day""/>
                            <xs:enumeration value=""Friday""/>
                            <xs:enumeration value=""Monday""/>
                            <xs:enumeration value=""Saturday""/>
                            <xs:enumeration value=""Sunday""/>
                            <xs:enumeration value=""Thursday""/>
                            <xs:enumeration value=""Tuesday""/>
                            <xs:enumeration value=""Wednesday""/>
                            <xs:enumeration value=""Weekday""/>
                            <xs:enumeration value=""WeekendDay""/>
                        </xs:restriction>
                    </xs:simpleType>");
        }

        /// <summary>
        /// Verify the CreateItemResponseType structure.
        /// </summary>
        /// <param name="createItemResponse">The response got from server via CreateItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyCreateItemOperation(CreateItemResponseType createItemResponse, bool isSchemaValidated)
        {
            // If the validation event return any error or warning, the schema validation is false, 
            // which indicates the schema is not matched with the expected result. 
            // So these requirement can be verified when the isSchemaValidation is true.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R155");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R155            
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                155,
                @"[In CreateItem Operation] The following is the WSDL port type specification for the CreateItem operation.<wsdl:operation name=""CreateItem"">
                     <wsdl:input message=""tns:CreateItemSoapIn"" />
                     <wsdl:output message=""tns:CreateItemSoapOut""/>
                </wsdl:operation>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R157");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R157
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                157,
                @"[In CreateItem Operation] The following is the WSDL binding specification for the CreateItem operation.<wsdl:operation name=""CreateItem"">
                   <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem""/>
                   <wsdl:input>
                      <soap:header message=""tns:CreateItemSoapIn"" part=""Impersonation"" use=""literal""/>
                      <soap:header message=""tns:CreateItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
                      <soap:header message=""tns:CreateItemSoapIn"" part=""RequestVersion"" use=""literal""/>
                      <soap:header message=""tns:CreateItemSoapIn"" part=""TimeZoneContext"" use=""literal""/>
                      <soap:body parts=""request"" use=""literal""/>
                   </wsdl:input>
                   <wsdl:output>
                      <soap:body parts=""CreateItemResult"" use=""literal""/>
                      <soap:header message=""tns:CreateItemSoapOut"" part=""ServerVersion"" use=""literal""/>
                   </wsdl:output>
                </wsdl:operation>");

            if (createItemResponse.ResponseMessages.Items != null && createItemResponse.ResponseMessages.Items.Length > 0)
            {
                foreach (ItemInfoResponseMessageType itemInfo in createItemResponse.ResponseMessages.Items)
                {
                    if (itemInfo.ResponseClass == ResponseClassType.Success)
                    {
                        this.VerifyTaskType(isSchemaValidated, itemInfo);
                    }
                }
            }
        }

        /// <summary>
        /// Verify the DeleteItemResponseType structure.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyDeleteItemOperation(bool isSchemaValidated)
        {
            // If the validation event return any error or warning, the schema validation is false, 
            // which indicates the schema is not matched with the expected result. 
            // So these requirement can be verified when the isSchemaValidation is true.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R162");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R162  
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                162,
                @"[In DeleteItem Operation] The following is the WSDL port type specification for the DeleteItem operation.<wsdl:operation name=""DeleteItem"">
                   <wsdl:input message=""tns:DeleteItemSoapIn"" />
                   <wsdl:output message=""tns:DeleteItemSoapOut"" />
                </wsdl:operation>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R164");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R164            
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                164,
                @"[In DeleteItem Operation] The following is the WSDL binding specification for the DeleteItem operation. <wsdl:operation name=""DeleteItem"">
                   <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/DeleteItem""/>
                   <wsdl:input>
                      <soap:header message=""tns:DeleteItemSoapIn"" part=""Impersonation"" use=""literal""/>
                      <soap:header message=""tns:DeleteItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
                      <soap:header message=""tns:DeleteItemSoapIn"" part=""RequestVersion"" use=""literal""/>
                      <soap:body parts=""request"" use=""literal""/>
                   </wsdl:input>
                   <wsdl:output>
                      <soap:body parts=""DeleteItemResult"" use=""literal""/>
                      <soap:header message=""tns:DeleteItemSoapOut"" part=""ServerVersion"" use=""literal""/>
                   </wsdl:output>
                </wsdl:operation>");
        }

        /// <summary>
        /// Verify the GetItemResponseType structure.
        /// </summary>
        /// <param name="getItemResponse">The response got from server via GetItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyGetItemOperation(GetItemResponseType getItemResponse, bool isSchemaValidated)
        {
            // If the validation event return any error or warning, the schema validation is false, 
            // which indicates the schema is not matched with the expected result. 
            // So these requirement can be verified when the isSchemaValidation is true.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R176");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R176
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                176,
                @"[In GetItem Operation] The following is the WSDL port type specification for the GetItem operation. <wsdl:operation name=""GetItem"">
                   <wsdl:input message=""tns:GetItemSoapIn"" />
                   <wsdl:output message=""tns:GetItemSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R178");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R178
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                178,
                @"[In GetItem Operation] The following is the WSDL binding specification for the GetItem operation. <wsdl:operation name=""GetItem"">
                   <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/GetItem""/>
                   <wsdl:input>
                      <soap:header message=""tns:GetItemSoapIn"" part=""Impersonation"" use=""literal""/>
                      <soap:header message=""tns:GetItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
                      <soap:header message=""tns:GetItemSoapIn"" part=""RequestVersion"" use=""literal""/>
                      <soap:header message=""tns:GetItemSoapIn"" part=""TimeZoneContext"" use=""literal""/>
                      <soap:body parts=""request"" use=""literal""/>
                   </wsdl:input>
                   <wsdl:output>
                      <soap:body parts=""GetItemResult"" use=""literal""/>
                      <soap:header message=""tns:GetItemSoapOut"" part=""ServerVersion"" use=""literal""/>
                   </wsdl:output>
                </wsdl:operation>");

            if (getItemResponse.ResponseMessages.Items != null && getItemResponse.ResponseMessages.Items.Length > 0)
            {
                foreach (ItemInfoResponseMessageType itemInfo in getItemResponse.ResponseMessages.Items)
                {
                    if (itemInfo.ResponseClass == ResponseClassType.Success)
                    {
                        this.VerifyTaskType(isSchemaValidated, itemInfo);
                    }
                }
            }
        }

        /// <summary>
        /// Verify the MoveItemResponseType structure.
        /// </summary>
        /// <param name="moveItemResponse">The response got from server via MoveItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyMoveItemOperation(MoveItemResponseType moveItemResponse, bool isSchemaValidated)
        {
            // If the validation event return any error or warning, the schema validation is false, 
            // which indicates the schema is not matched with the expected result. 
            // So these requirement can be verified when the isSchemaValidation is true.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R183");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R183
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                183,
                @"[In MoveItem Operation] The following is the WSDL port type specification for the MoveItem operation. <wsdl:operation name=""MoveItem"">
                   <wsdl:input message=""tns:MoveItemSoapIn"" />
                   <wsdl:output message=""tns:MoveItemSoapOut"" />
                </wsdl:operation>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R185");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R185
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                185,
                @"[In MoveItem Operation] The following is the WSDL binding specification for the MoveItem operation. <wsdl:operation name=""MoveItem"">
                   <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/MoveItem""/>
                   <wsdl:input>
                      <soap:header message=""tns:MoveItemSoapIn"" part=""Impersonation"" use=""literal""/>
                      <soap:header message=""tns:MoveItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
                      <soap:header message=""tns:MoveItemSoapIn"" part=""RequestVersion"" use=""literal""/>
                      <soap:body parts=""request"" use=""literal""/>
                   </wsdl:input>
                   <wsdl:output>
                      <soap:body parts=""MoveItemResult"" use=""literal""/>
                      <soap:header message=""tns:MoveItemSoapOut"" part=""ServerVersion"" use=""literal""/>
                   </wsdl:output>
                </wsdl:operation>");

            if (moveItemResponse.ResponseMessages.Items != null && moveItemResponse.ResponseMessages.Items.Length > 0)
            {
                foreach (ItemInfoResponseMessageType itemInfo in moveItemResponse.ResponseMessages.Items)
                {
                    if (itemInfo.ResponseClass == ResponseClassType.Success)
                    {
                        this.VerifyTaskType(isSchemaValidated, itemInfo);
                    }
                }
            }
        }

        /// <summary>
        /// Verify the UpdateItemResponseType structure.
        /// </summary>
        /// <param name="updateItemResponse">The response got from server via UpdateItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyUpdateItemOperation(UpdateItemResponseType updateItemResponse, bool isSchemaValidated)
        {
            // If the validation event return any error or warning, the schema validation is false, 
            // which indicates the schema is not matched with the expected result. 
            // So these requirement can be verified when the isSchemaValidation is true.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R191");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R191        
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                191,
                @"[In UpdateItem Operation] The following is the WSDL port type specification for the UpdateItem operation. <wsdl:operation name=""UpdateItem"">
                   <wsdl:input message=""tns:UpdateItemSoapIn"" />
                   <wsdl:output message=""tns:UpdateItemSoapOut""/>
                </wsdl:operation>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R193");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R193 
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                193,
                @"[In UpdateItem Operation] The following is the WSDL binding specification for the UpdateItem operation. <wsdl:operation name=""UpdateItem"">
                   <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/UpdateItem""/>
                   <wsdl:input>
                      <soap:header message=""tns:UpdateItemSoapIn"" part=""Impersonation"" use=""literal""/>
                      <soap:header message=""tns:UpdateItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
                      <soap:header message=""tns:UpdateItemSoapIn"" part=""RequestVersion"" use=""literal""/>
                      <soap:header message=""tns:UpdateItemSoapIn"" part=""TimeZoneContext"" use=""literal""/>
                      <soap:body parts=""request"" use=""literal""/>
                   </wsdl:input>
                   <wsdl:output>
                      <soap:body parts=""UpdateItemResult"" use=""literal""/>
                      <soap:header message=""tns:UpdateItemSoapOut"" part=""ServerVersion"" use=""literal""/>
                   </wsdl:output>
                </wsdl:operation>");

            if (updateItemResponse.ResponseMessages.Items != null && updateItemResponse.ResponseMessages.Items.Length > 0)
            {
                foreach (ItemInfoResponseMessageType itemInfo in updateItemResponse.ResponseMessages.Items)
                {
                    if (itemInfo.ResponseClass == ResponseClassType.Success)
                    {
                        this.VerifyTaskType(isSchemaValidated, itemInfo);
                    }
                }
            }
        }

        /// <summary>
        /// Verify the ArrayOfStringsType structure.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyArrayOfStringsType(bool isSchemaValidated)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1081");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1081            
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1081,
                @"[In t:ArrayOfStringsType Complex Type] The type [ArrayOfStringsType] is defined as follow:
                    <xs:complexType name=""ArrayOfStringsType"">
                      <xs:sequence>
                        <xs:element name=""String""
                          type=""xs:string""
                          minOccurs=""0""
                          maxOccurs=""unbounded""
                         />
                      </xs:sequence>
                    </xs:complexType>");
        }
    }
}