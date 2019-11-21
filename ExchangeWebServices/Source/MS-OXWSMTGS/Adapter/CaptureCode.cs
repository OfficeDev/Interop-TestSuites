namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-OXWSMTGS.
    /// </summary>
    public partial class MS_OXWSMTGSAdapter
    {
        #region Operation Verification
        /// <summary>
        /// Verify the WSDL port type specifications for the CopyItem operation and CopyItemResponseType structure. 
        /// </summary>
        /// <param name="response">The response message of CopyItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyCopyItemOperation(CopyItemResponseType response, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema validation should be true.");

            // Verify the WSDL port type specifications for the CopyItem operation
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R464");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R464
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             464,
             @"[In CopyItem operation] The following is the WSDL port type specification for the CopyItem operation.
<wsdl:operation name=""CopyItem"">
     <wsdl:input message=""tns:CopyItemSoapIn"" />
     <wsdl:output message=""tns:CopyItemSoapOut"" />
</wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R597");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                597,
                @"[In CopyItem Operation] The following is the WSDL binding specification for the CopyItem operation.
<wsdl:operation name=""CopyItem"">
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

            // Verify calendar related items
            foreach (ResponseMessageType responseMessage in response.ResponseMessages.Items)
            {
                if (responseMessage.ResponseClass == ResponseClassType.Success)
                {
                    ItemInfoResponseMessageType itemInfo = responseMessage as ItemInfoResponseMessageType;

                    ItemType item = itemInfo.Items.Items[0];
                    this.VerifyItemTypeItems(item, isSchemaValidated);
                }
            }
        }

        /// <summary>
        /// Verify the CreateItemResponseType structure. 
        /// </summary>
        /// <param name="response">The response message of CreateItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyCreateItemOperation(CreateItemResponseType response, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema validation should be true.");

            // Verify the WSDL port type specifications for the CreateItem operation
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R472");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R472
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             472,
             @"[In CreateItem operation] The following is the WSDL port type specification for the CreateItem operation.
<wsdl:operation name=""CreateItem"">
     <wsdl:input message=""tns:CreateItemSoapIn"" />
     <wsdl:output message=""tns:CreateItemSoapOut"" />
</wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R596");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                596,
                @"[In CreateItem Operation] The following is the WSDL binding specification for the CreateItem operation. 
<wsdl:operation name=""CreateItem"">
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

            // Verify calendar related items
            foreach (ResponseMessageType responseMessage in response.ResponseMessages.Items)
            {
                if (responseMessage.ResponseClass == ResponseClassType.Success)
                {
                    ItemInfoResponseMessageType itemInfo = responseMessage as ItemInfoResponseMessageType;
                    if (itemInfo.Items.Items != null)
                    {
                        ItemType item = itemInfo.Items.Items[0];
                        this.VerifyItemTypeItems(item, isSchemaValidated);
                    }
                }
            }
        }

        /// <summary>
        /// Verify the WSDL port type specifications for the DeleteItem operation. 
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyDeleteItemOperation(bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema validation should be true.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R444");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R444
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             444,
             @"[In DeleteItem operation] The following is the WSDL port type specification for the DeleteItem operation.
<wsdl:operation name=""DeleteItem"">
     <wsdl:input message=""tns:DeleteItemSoapIn"" />
     <wsdl:output message=""tns:DeleteItemSoapOut"" />
</wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R614");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                614,
                @"[In DeleteItem Operation] The following is the WSDL binding specification for the DeleteItem operation.
<wsdl:operation name=""DeleteItem"">
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
        /// Verify the WSDL port type specifications for the GetItem operation and GetItemResponseType structure
        /// </summary>
        /// <param name="response">The response message of GetItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyGetItemOperation(GetItemResponseType response, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema validation should be true.");

            // Verify the WSDL port type specifications for the GetItem operation.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R436");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R436
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             436,
             @"[In GetItem operation] The following is the WSDL port type specification for the GetItem operation. 
<wsdl:operation name=""GetItem"">
     <wsdl:input message=""tns:GetItemSoapIn"" />
     <wsdl:output message=""tns:GetItemSoapOut"" />
</wsdl:operation>");

            // Verify the WSDL binding specification for the GetItem operation
            if (Common.IsRequirementEnabled(694, Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R694");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R694           
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    694,
                    @"[In Appendix C: Product Behavior] Implementation does include the following WSDL binding specification for the GetItem operation.
    <wsdl:operation name=""GetItem"">
       <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/GetItem""/>
       <wsdl:input>
          <soap:header message=""tns:GetItemSoapIn"" part=""Impersonation"" use=""literal""/>
          <soap:header message=""tns:GetItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
          <soap:header message=""tns:GetItemSoapIn"" part=""RequestVersion"" use=""literal""/>
          <soap:header message=""tns:GetItemSoapIn"" part=""TimeZoneContext"" use=""literal""/>
          <soap:header message=""tns:GetItemSoapIn"" part=""DateTimePrecision"" use=""literal"" />
    <soap:body parts=""request"" use=""literal""/>
       </wsdl:input>
       <wsdl:output>
          <soap:body parts=""GetItemResult"" use=""literal""/>
          <soap:header message=""tns:GetItemSoapOut"" part=""ServerVersion"" use=""literal""/>
       </wsdl:output>
    </wsdl:operation> (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(695, Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R695");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R695           
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    695,
                    @"[In Appendix C: Product Behavior] Implementation does include the following WSDL binding specification for the GetItem operation.
    <wsdl:operation name=""GetItem"">
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
    </wsdl:operation> (Exchange 2007, Exchange 2010 and Exchange 2010 SP1 follow this behavior.)");
            }

            // Verify items
            foreach (ResponseMessageType responseMessage in response.ResponseMessages.Items)
            {
                if (responseMessage.ResponseClass == ResponseClassType.Success)
                {
                    ItemInfoResponseMessageType itemInfo = responseMessage as ItemInfoResponseMessageType;

                    // Each ItemInfoResponseMessageType contains one calendar related item for GetItem operation.
                    ItemType item = itemInfo.Items.Items[0];
                    this.VerifyItemTypeItems(item, isSchemaValidated);
                }
            }
        }

        /// <summary>
        /// Verify the ArrayOfRemindersType structure
        /// </summary>
        /// <param name="arrayOfReminders">Contains a list representing reminders for a meeting</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyArrayOfRemindersType(ReminderType[] arrayOfReminders, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2068");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2068
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             2068,
             @"[In t:ArrayOfRemindersType Complex Type]The ArrayOfRemindersType complex type specifies an array of reminders.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2069");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2069
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             2069,
             @"[In t:ArrayOfRemindersType Complex Type][The schema of ArrayOfRemindersType is defined as:]
	<xs:complexType name=""ArrayOfRemindersType"" >
      < xs:sequence >

        < xs:element name = ""Reminder"" type = ""t:ReminderType"" minOccurs = ""0"" maxOccurs = ""unbounded"" />

      </ xs:sequence >

    </ xs:complexType > ");



            foreach (ReminderType reminder in arrayOfReminders)
            {
                this.VerifyReminderType(reminder, isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify the ReminderType structure.
        /// </summary>
        /// <param name="reminder">Reminder for a calendar</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyReminderType(ReminderType reminder, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2071");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2071
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             2071,
             @"[In t:ArrayOfRemindersType Complex Type]The type of Reminder is t:ReminderType (section 3.1.4.5.3.4).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2074");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2074
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2074,
                @"[In t:ReminderType Complex Type][The schema of ReminderType is defined as:]
	<xs:complexType name=""ReminderType"" >
      < xs:sequence >

        < xs:element name = ""Subject"" type = ""xs:string"" minOccurs = ""1"" maxOccurs = ""1"" />
        xs:element name = ""Location"" type = ""xs:string"" minOccurs = ""0"" maxOccurs = ""1"" />

        < xs:element name = ""ReminderTime"" type = ""xs:dateTime"" minOccurs = ""1"" maxOccurs = ""1"" />

        < xs:element name = ""StartDate"" type = ""xs:dateTime"" minOccurs = ""1"" maxOccurs = ""1"" />

        < xs:element name = ""EndDate"" type = ""xs:dateTime"" minOccurs = ""1"" maxOccurs = ""1"" />

        < xs:element name = ""ItemId"" type = ""t:ItemIdType"" minOccurs = ""1"" maxOccurs = ""1"" />

        < xs:element name = ""RecurringMasterItemId"" type = ""t:ItemIdType"" minOccurs = ""0"" maxOccurs = ""1"" />

        < xs:element name = ""ReminderGroup"" type = ""t:ReminderGroupType"" minOccurs = ""0"" maxOccurs = ""1"" />

        < xs:element name = ""UID"" type = ""xs:string"" minOccurs = ""1"" maxOccurs = ""1"" />

      </ xs:sequence >

    </ xs:complexType > ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2076");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2076
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2076,
                @"[In t:ReminderType Complex Type]The type of Subject is xs:string ([XMLSCHEMA2]).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2078");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2078,
                @"[In t:ReminderType Complex Type]The type of Location is xs:string.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2080");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2080
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2080,
                @"[In t:ReminderType Complex Type]The type of ReminderTime is xs:dateTime ([XMLSCHEMA2]).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2082");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2082
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2082,
                @"[In t:ReminderType Complex Type]The type of StartDate is xs:string ([XMLSCHEMA2]).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2084");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2084
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2084,
                @"[In t:ReminderType Complex Type]The type of EndDate is xs:dateTime .");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2086");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2086
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2086,
                @"[In t:ReminderType Complex Type]The type of ItemId is t:ItemIdType ([MS-OXWSCORE] section 2.2.4.25).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2088");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2088
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2088,
                @"[In t:ReminderType Complex Type]The type of RecurringMasterItemId is t:ItemIdType.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2092");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2092
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2092,
                @"[In t:ReminderType Complex Type]The type of UID is xs:string .");


            if (reminder.ReminderGroupSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2090");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2090
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    2090,
                    @"[In t:ReminderType Complex Type]The type of ReminderGroup is t:ReminderGroupType (section 3.1.4.5.4.1).");

                this.VerifyReminderGroupType(reminder.ReminderGroup, isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify the ReminderGroupType structure.
        /// </summary>
        /// <param name="reminderGroup">Reminder group for a calendar</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyReminderGroupType(ReminderGroupType reminderGroup, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2097");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2097
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2097,
                @"[In ReminderGroupType Simple Type][The schema of ReminderGroupType is defined as:]
	<xs:simpleType name=""ReminderGroupType"" >
      < xs:restriction base = ""xs:string"" >

        < xs:enumeration value = ""Calendar"" />

        < xs:enumeration value = ""Task"" />

      </ xs:restriction >

    </ xs:simpleType > ");
        }

        /// <summary>
        /// Captures GetRemindersResponseMessageType related requirements.
        /// </summary>
        /// <param name="getReminderResponseMessage">An GetRemindersResponseMessageType instance.</param>
        /// <param name="isSchemaValidated">A Boolean value indicates the schema validation result.</param>
        private void VerifyGetReminderResponseMessageType(GetRemindersResponseMessageType getReminderResponseMessage, bool isSchemaValidated)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2046");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2046,
                @"[In Complex Types]GetRemindersResponseMessageType:Specifies a response to a GetReminders operation request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2050");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2050
            this.Site.CaptureRequirementIfIsNotNull(
                getReminderResponseMessage,
                2050,
                @"[In m:GetRemindersResponseMessageType Complex Type]The GetRemindersResponseMessageType complex type specifies a response to a request to return reminders. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2051");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2051
            this.Site.CaptureRequirementIfIsNotNull(
                getReminderResponseMessage,
                2051,
                @"[In m:GetRemindersResponseMessageType Complex Type]This type[GetRemindersResponseMessageType] extends the ResponseMessageType ([MS-OXWSCDATA] section 2.2.4.67).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R2052");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2052
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2052,
                @"[In m:GetRemindersResponseMessageType Complex Type][The schema of GetRemindersResponseMessageType is defined as:]	
<xs:complexType name=""GetRemindersResponseMessageType"" >
    < xs:complexContent >

    < xs:extension base = ""m:ResponseMessageType"" >

        < xs:sequence >

        < xs:element name = ""Reminders"" type = ""t:ArrayOfRemindersType"" minOccurs = ""1"" maxOccurs = ""1"" />

        </ xs:sequence >

    </ xs:extension >

    </ xs:complexContent >
");

            if(getReminderResponseMessage.Reminders != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R2054");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2054
                Site.CaptureRequirementIfIsInstanceOfType(
                    getReminderResponseMessage.Reminders,
                    typeof(ReminderType[]),
                    2054,
                    @"[In m:GetRemindersResponseMessageType Complex Type]the type of Reminders is t:ArrayOfRemindersType (section 3.1.4.5.3.3).");

                this.VerifyArrayOfRemindersType(getReminderResponseMessage.Reminders, isSchemaValidated);
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2012");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2012
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2012,
                @"[In GetReminders operation]The following is the WSDL port type specification for the GetReminders operation.
        <wsdl:operation name=""GetReminders"" >
                < wsdl:input message = ""tns:GetRemindersSoapIn"" />

                < wsdl:output message = ""tns:GetRemindersSoapOut"" />

        </ wsdl:operation >");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2013");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2013,
                @"[In GetReminders operation]The following is the WSDL binding specification for the GetReminders operation.	
                    <wsdl:operation name=""GetReminders"" >
            < soap:operation soapAction = ""http://schemas.microsoft.com/exchange/services/2006/messages/GetReminders"" />

            < wsdl:input >

                < soap:header message = ""tns:GetRemindersSoapIn"" part = ""RequestVersion"" use = ""literal"" />

                < soap:body parts = ""request"" use = ""literal"" />

            </ wsdl:input >

            < wsdl:output >

                < soap:body parts = ""GetRemindersResult"" use = ""literal"" />

                < soap:header message = ""tns:GetRemindersSoapOut"" part = ""ServerVersion"" use = ""literal"" />

            </ wsdl:output >

        </ wsdl:operation > ");


            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2025");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2025,
                @"[In tns:GetRemindersSoapOut Message][The schema of GetRemindersSoapOut is defined as:] 	<wsdl:message name=""GetRemindersSoapOut"" >
            < wsdl:part name = ""GetRemindersResult"" element = ""tns:GetRemindersResponse"" />

            < wsdl:part name = ""ServerVersion"" element = ""t:ServerVersionInfo"" />

        </ wsdl:message > ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2016");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2016,
                @"[In Messages] GetRemindersSoapOut:Specifies the SOAP message that is returned by the server in response to a GetRemindersSoapIn operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2019");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2019
            // According to the schema, getRemindersResponseMessage is the SOAP body of a response message returned by server, this requirement can be verified directly.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2019,
                @"[In tns:GetRemindersSoapOut Message] GetRemindersResult:Specifies the SOAP body of a response message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2028");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2028
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2028,
                @"[In tns:GetRemindersSoapOut Message]the element of GetRemindersResult is tns:GetRemindersResponse (section 3.1.4.5.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2030");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2030
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2030,
                @"[In tns:GetRemindersSoapOut Message]the element of ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.3.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2031");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2031
            // According to the schema, ServerVersion is the SOAP header that contains the server version information, this requirement can be verified directly.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2031,
                @"[In tns:GetRemindersSoapOut Message]ServerVersion:Specifies a SOAP header that identifies the server version for the response to a GetReminders operation request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2044");

            Site.CaptureRequirementIfIsNotNull(
                getReminderResponseMessage,
                2044,
                @"[In m:GetRemindersResponse Element][The schema of GetRemindersResponse is defined as:]	<xs:element name=""GetRemindersResponse"" type =""m: GetRemindersResponseMessageType"" />");
        }

        /// <summary>
        /// Verify the WSDL port type specifications for the MoveItem operation and MoveItemResponseType structure. 
        /// </summary>
        /// <param name="response">The response message of MoveItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyMoveItemOperation(MoveItemResponseType response, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema validation should be true.");

            // Verify the WSDL port type specifications for the MoveItem operation
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R457");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R457
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             457,
             @"[In MoveItem operation] The following is the WSDL port type specification for the MoveItem operation. 
<wsdl:operation name=""MoveItem"">
     <wsdl:input message=""tns:MoveItemSoapIn"" />
     <wsdl:output message=""tns:MoveItemSoapOut"" />
</wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R635");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                635,
                @"[In MoveItem Operation] The following is the WSDL binding specification for the MoveItem operation.
<wsdl:operation name=""MoveItem"">
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

            // Verify calendar related items
            foreach (ResponseMessageType respMsg in response.ResponseMessages.Items)
            {
                if (respMsg.ResponseClass == ResponseClassType.Success)
                {
                    ItemInfoResponseMessageType itemInfo = respMsg as ItemInfoResponseMessageType;

                    // Each ItemInfoResponseMessageType contains one calendar related item for MoveItem operation.
                    ItemType item = itemInfo.Items.Items[0];
                    this.VerifyItemTypeItems(item, isSchemaValidated);
                }
            }
        }

        /// <summary>
        /// Verify the WSDL port type specifications for the UpdateItem operation and UpdateItemResponseType structure. 
        /// </summary>
        /// <param name="response">The response message of UpdateItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyUpdateItemOperation(UpdateItemResponseType response, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema validation should be true.");

            // Verify the WSDL port type specifications for the UpdateItem operation
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R451");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R451
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             451,
             @"[In UpdateItem operation] The following is the WSDL port type specification for the UpdateItem operation. 
<wsdl:operation name=""UpdateItem"">
     <wsdl:input message=""tns:UpdateItemSoapIn"" />
     <wsdl:output message=""tns:UpdateItemSoapOut"" />
</wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R645");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                645,
                @"[In UpdateItem Operation] The following is the WSDL binding specification for the UpdateItem operation.
<wsdl:operation name=""UpdateItem"">
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

            // Verify calendar related items
            foreach (ResponseMessageType respMsg in response.ResponseMessages.Items)
            {
                if (respMsg.ResponseClass == ResponseClassType.Success)
                {
                    UpdateItemResponseMessageType updateItemInfo = respMsg as UpdateItemResponseMessageType;
                    if (updateItemInfo.Items.Items != null)
                    {
                        // Each ItemInfoResponseMessageType contains one calendar related item for UpdateItem operation.
                        ItemType item = updateItemInfo.Items.Items[0];
                        this.VerifyItemTypeItems(item, isSchemaValidated);
                    }
                }
            }
        }
        #endregion

        #region Element Types Verification
        /// <summary>
        /// Verify SOAP version.
        /// </summary>
        private void VerifySoapVersion()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1
            // According to the implementation of adapter, the message is formatted as SOAP 1.1. If the operation is invoked successfully, then this requirement can be verified.
            Site.CaptureRequirement(
                1,
                @"[In Transport] Messages are transported by using SOAP version 1.1, as specified in [SOAP1.1].");
        }

        /// <summary>
        /// Verify the transport related requirements.
        /// </summary>
        private void VerifyTransportType()
        {
            // Get the transport type
            TransportProtocol transport = (TransportProtocol)Enum.Parse(typeof(TransportProtocol), Common.GetConfigurationPropertyValue("TransportType", Site), true);
            if (transport == TransportProtocol.HTTPS)
            {
                if (Common.IsRequirementEnabled(504, Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R504");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R504
                    // Because Adapter uses SOAP and HTTPS to communicate with server, if server returned data without exception, this requirement has been captured.
                    Site.CaptureRequirement(
                        504,
                        @"[In Appendix C: Product Behavior] Implementation does support SOAP over HTTPS, as specified in [RFC2818]. (Exchange 2007 and above follow this behavior.)");
                }
            }
            else if (transport == TransportProtocol.HTTP)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R502");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R502
                // Because Adapter uses SOAP and HTTP to communicate with server, if server returned data without exception, this requirement has been captured.
                Site.CaptureRequirement(
                    502,
                    @"[In Transport] The protocol MUST support SOAP over HTTP, as specified in [RFC2616].");
            }
        }

        /// <summary>
        /// Verify the structure of calendar related item.
        /// </summary>
        /// <param name="item">One kind of calendar related item</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyItemTypeItems(ItemType item, bool isSchemaValidated)
        {
            if (item != null)
            {
                CalendarItemType calendarItem = item as CalendarItemType;
                if (calendarItem != null)
                {
                    this.VerifyCalendarItemType(calendarItem, isSchemaValidated);
                }

                MeetingRequestMessageType meetingrequest = item as MeetingRequestMessageType;
                if (meetingrequest != null)
                {
                    this.VerifyMeetingRequestMessageType(meetingrequest, isSchemaValidated);
                }

                MeetingCancellationMessageType meetingCancel = item as MeetingCancellationMessageType;
                if (meetingCancel != null)
                {
                    this.VerifyMeetingCancellationMessageType(meetingCancel, isSchemaValidated);
                }

                MeetingResponseMessageType meetingResponse = item as MeetingResponseMessageType;
                if (meetingResponse != null)
                {
                    this.VerifyMeetingResponseMessageType(meetingResponse, isSchemaValidated);
                }
            }
        }

        /// <summary>
        /// Verify the InboxReminderType structure.
        /// </summary>
        /// <param name="inboxReminder">InboxReminder for a calendar</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyInboxReminderType(InboxReminderType inboxReminder, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1278");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1278
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             1278,
             @"[In Appendix C: Product Behavior] Implementation does support InboxReminderType which specifies an inbox reminder. (Exchange 2013 and above follow this behavior.)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1033");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1033
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             1033,
             @"[In t:ArrayOfInboxReminderType] The type of InboxReminder is t:InboxReminderType (section 2.2.4.13).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1349");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1349
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                1349,
                @"[In t:InboxReminderType] [its schema is] <xs:complexType name=""InboxReminderType"" >
   < xs:sequence >
     < xs:element name = ""Id"" type = ""t:GuidType"" minOccurs = ""0"" maxOccurs = ""1"" />
     < xs:element name = ""ReminderOffset"" type = ""xs:int"" minOccurs = ""0"" maxOccurs = ""1"" />
  < xs:element name = ""Message"" type = ""xs:string"" minOccurs = ""0"" maxOccurs = ""1"" />
  < xs:element name = ""IsOrganizerReminder"" type = ""xs:boolean"" minOccurs = ""0"" maxOccurs = ""1"" />
  < xs:element name = ""OccurrenceChange""
         type = ""t:EmailReminderChangeType"" minOccurs = ""0"" maxOccurs = ""1"" />
  < xs:element name = ""IsImportedFromOLC"" type = ""xs:boolean"" minOccurs = ""0"" maxOccurs = ""1"" />
   < xs:element name = ""SendOption""
         type = ""t:EmailReminderSendOption"" minOccurs = ""0"" maxOccurs = ""1"" />
   </ xs:sequence >
 </ xs:complexType > ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1060");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1060
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                1060,
                @"[In t:InboxReminderType] The type of Id is t:GuidType ([XMLSCHEMA2] ).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1062");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                1062,
                @"[In t:InboxReminderType] The type of ReminderOffset is xs:int ([XMLSCHEMA2] ).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1064");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1064
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                1064,
                @"[In t:InboxReminderType] The type of Message is xs:string ([XMLSCHEMA2] ).");

            if (inboxReminder.IsOrganizerReminderSpecified)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1066");

                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    1066,
                    @"[In t:InboxReminderType] The type of IsOrganizerReminder is xs:boolean ([XMLSCHEMA2] ).");
            }

            if (Common.IsRequirementEnabled(1310, this.Site) && inboxReminder.OccurrenceChangeSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1069");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1069
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    1069,
                    @"[In t:InboxReminderType] OccurrenceChange: Specifies how this reminder has been modified for an occurrence");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1068");

                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    1068,
                    @"[In t:InboxReminderType] The type of OccurrenceChange t:EmailReminderChangeType (section 2.2.5.6).");

                this.VerifyEmailReminderChangeType(inboxReminder.OccurrenceChange, isSchemaValidated);
            }

            if (inboxReminder.IsImportedFromOLCSpecified)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2005");

                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    2005,
                    @"[In t:InboxReminderType] The type of IsImportedFromOLC is xs:boolean.");
            }
            if (Common.IsRequirementEnabled(1312, this.Site) && inboxReminder.SendOptionSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1341");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1341
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    1341,
                    @"[In t:InboxReminderType] SendOption: Specifies the send option.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1070");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1070
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    1070,
                    @"[In t:InboxReminderType] The type of SendOption is t:EmailReminderSendOption (section 2.2.5.7).");

                this.VerifyEmailReminderSendOption(inboxReminder.SendOption, isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify the AttendeeType structure.
        /// </summary>
        /// <param name="attendee">Attendee or resource for a meeting</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyAttendeeType(AttendeeType attendee, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R368");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R368
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             368,
             @"[In t:NonEmptyArrayOfAttendeesType Complex Type] The type of Attendee type is t:AttendeeType (section 2.2.4.4).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R135");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R135
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                135,
                @"[In t:AttendeeType Complex Type] [its schema is]  <xs:complexType name=""AttendeeType"">
                   <xs:sequence>
                     <xs:element name=""Mailbox""
                       type=""t:EmailAddressType""
                      />
                     <xs:element name=""ResponseType""
                       type=""t:ResponseTypeType""
                       minOccurs=""0""
                      />
                     <xs:element name=""LastResponseTime""
                       type=""xs:dateTime""
                       minOccurs=""0""                
                      />
                     <xs:element name=""ProposedStart""
                       type=""xs:dateTime"" 
                       minOccurs=""0""
                      />
                     <xs:element name=""ProposedEnd"" 
                       type=""xs:dateTime"" 
                       minOccurs=""0""
                      />
                   </xs:sequence>
                 </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R136");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R136
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                136,
                @"[In t:AttendeeType Complex Type] The type of Mailbox is t:EmailAddressType ([MS-OXWSCDATA] section 2.2.4.27).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R138");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                138,
                @"[In t:AttendeeType Complex Type] The type of  ResponseType is t:ResponseTypeType (section 2.2.5.10).");

            if (attendee.LastResponseTimeSpecified)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R140");

                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    140,
                    @"[In t:AttendeeType Complex Type] The type of LastResponseTime is xs:dateTime ([XMLSCHEMA2]).");
            }

            if (attendee.ProposedStartSpecified)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1042");

                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    1042,
                    @"[In t:AttendeeType Complex Type] The type of ProposedStart is xs:dateTime.");
            }

            if (attendee.ProposedEndSpecified)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1044");

                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    1044,
                    @"[In t:AttendeeType Complex Type] The type of ProposedEnd is xs:dateTime.");
            }

            // Verify child element EmailAddressType.
            this.VerifyEmailAddressType(attendee.Mailbox, isSchemaValidated);

            if (attendee.ResponseTypeSpecified == true)
            {
                this.VerifyResponseTypeType(isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify the CalendarItemType structure.
        /// </summary>
        /// <param name="calendarItem">Represents a server calendar item</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyCalendarItemType(CalendarItemType calendarItem, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R149");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R149
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                149,
                @"[In t:CalendarItemType Complex Type] [its schema is]
                <xs:complexType name=""CalendarItemType"">
                    <xs:complexContent>
                      <xs:extension base=""t:ItemType"">
                        <xs:sequence>
                          <!-- iCalendar properties -->
                          <xs:element name=""UID"" type=""xs:string"" minOccurs=""0""/>
                          <xs:element name=""RecurrenceId"" type=""xs:dateTime"" minOccurs=""0""/>
                          <xs:element name=""DateTimeStamp"" type=""xs:dateTime"" minOccurs=""0""/>
                          <!-- Single and Occurrence only -->
                          <xs:element name=""Start"" type=""xs:dateTime"" minOccurs=""0""/>
                          <xs:element name=""End"" type=""xs:dateTime"" minOccurs=""0""/>
                          <!-- Occurrence only -->
                          <xs:element name=""OriginalStart"" type=""xs:dateTime"" minOccurs=""0""/>
                          <xs:element name=""IsAllDayEvent"" type=""xs:boolean"" minOccurs=""0""/>
                          <xs:element name=""LegacyFreeBusyStatus"" type=""t:LegacyFreeBusyType"" minOccurs=""0""/>
                          <xs:element name=""Location"" type=""xs:string"" minOccurs=""0""/>
                          <xs:element name=""When"" type=""xs:string"" minOccurs=""0""/>
                          <xs:element name=""IsMeeting"" type=""xs:boolean"" minOccurs=""0""/>
                          <xs:element name=""IsCancelled"" type=""xs:boolean"" minOccurs=""0""/>
                          <xs:element name=""IsRecurring"" type=""xs:boolean"" minOccurs=""0""/>
                          <xs:element name=""MeetingRequestWasSent"" type=""xs:boolean"" minOccurs=""0""/>
                          <xs:element name=""IsResponseRequested"" type=""xs:boolean"" minOccurs=""0""/>
                          <xs:element name=""CalendarItemType"" type=""t:CalendarItemTypeType"" minOccurs=""0""/>
                          <xs:element name=""MyResponseType"" type=""t:ResponseTypeType"" minOccurs=""0""/>
                          <xs:element name=""Organizer"" type=""t:SingleRecipientType"" minOccurs=""0""/>
                          <xs:element name=""RequiredAttendees"" type=""t:NonEmptyArrayOfAttendeesType"" minOccurs=""0""/>
                          <xs:element name=""OptionalAttendees"" type=""t:NonEmptyArrayOfAttendeesType"" minOccurs=""0""/>
                           <xs:element name=""Resources"" type=""t:NonEmptyArrayOfAttendeesType"" minOccurs=""0""/>
                          <!-- Conflicting and adjacent meetings -->
                          <xs:element name=""ConflictingMeetingCount"" type=""xs:int"" minOccurs=""0""/>
                          <xs:element name=""AdjacentMeetingCount"" type=""xs:int"" minOccurs=""0""/>
                          <xs:element name=""ConflictingMeetings"" type=""t:NonEmptyArrayOfAllItemsType"" minOccurs=""0""/>
                          <xs:element name=""AdjacentMeetings"" type=""t:NonEmptyArrayOfAllItemsType"" minOccurs=""0""/>
                          <xs:element name=""Duration"" type=""xs:string"" minOccurs=""0""/>
                          <xs:element name=""TimeZone"" type=""xs:string"" minOccurs=""0""/>
                          <xs:element name=""AppointmentReplyTime"" type=""xs:dateTime"" minOccurs=""0""/>
                          <xs:element name=""AppointmentSequenceNumber"" type=""xs:int"" minOccurs=""0""/>
                          <xs:element name=""AppointmentState"" type=""xs:int"" minOccurs=""0""/>
                          <!-- Recurrence specific data, only valid if CalendarItemType is RecurringMaster -->
                          <xs:element name=""Recurrence"" type=""t:RecurrenceType"" minOccurs=""0""/>
                          <xs:element name=""FirstOccurrence"" type=""t:OccurrenceInfoType"" minOccurs=""0""/>
                          <xs:element name=""LastOccurrence"" type=""t:OccurrenceInfoType"" minOccurs=""0""/>
                          <xs:element name=""ModifiedOccurrences"" type=""t:NonEmptyArrayOfOccurrenceInfoType"" minOccurs=""0""/>
                          <xs:element name=""DeletedOccurrences"" type=""t:NonEmptyArrayOfDeletedOccurrencesType"" minOccurs=""0""/>
                          <xs:element name=""MeetingTimeZone"" type=""t:TimeZoneType"" minOccurs=""0""/>
                          <xs:element name=""StartTimeZone"" type=""t:TimeZoneDefinitionType"" minOccurs=""0""/>
                          <xs:element name=""EndTimeZone"" type=""t:TimeZoneDefinitionType"" minOccurs=""0""/>
                          <xs:element name=""ConferenceType"" type=""xs:int"" minOccurs=""0""/>
                          <xs:element name=""AllowNewTimeProposal"" type=""xs:boolean"" minOccurs=""0""/>
                          <xs:element name=""IsOnlineMeeting"" type=""xs:boolean"" minOccurs=""0""/>
                          <xs:element name=""MeetingWorkspaceUrl"" type=""xs:string"" minOccurs=""0""/>
                          <xs:element name=""NetShowUrl"" type=""xs:string"" minOccurs=""0""/>
                          <xs:element name=""EnhancedLocation"" type=""t:EnhancedLocationType"" minOccurs=""0""/>
                          <xs:element name=""StartWallClock"" type=""xs:dateTime"" minOccurs=""0"" maxOccurs=""1""/>
                          <xs:element name=""EndWallClock"" type=""xs:dateTime"" minOccurs=""0"" maxOccurs=""1""/>
                          <xs:element name=""StartTimeZoneId"" type=""xs:string"" minOccurs=""0"" maxOccurs=""1""/>
                          <xs:element name=""EndTimeZoneId"" type=""xs:string"" minOccurs=""0"" maxOccurs=""1""/>
                          <xs:element name=""IntendedFreeBusyStatus"" type=""t:LegacyFreeBusyType"" minOccurs=""0"" />
                          <xs:element name=""JoinOnlineMeetingUrl"" type=""xs:string"" minOccurs=""0"" maxOccurs=""1"" />
                          <xs:element name=""OnlineMeetingSettings"" type=""t:OnlineMeetingSettingsType"" minOccurs=""0"" maxOccurs=""1""/>
                          <xs:element name=""IsOrganizer"" type=""xs:boolean"" minOccurs=""0""/>
                           <xs:element name=""InboxReminders"" type=""t:ArrayOfInboxReminderType"" minOccurs=""0""/>
                        </xs:sequence>
                      </xs:extension>
                    </xs:complexContent>
                  </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R509");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R509
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                509,
                @"[In t:CalendarItemType Complex Type] This complex type extends the ItemType complex type, as specified in [MS-OXWSCORE] section 2.2.4.8.");

            if (calendarItem.UID != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R150");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R150
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    150,
                    @"[In t:CalendarItemType Complex Type] The type of UID is xs:string ([XMLSCHEMA2]).");
            }

            if (calendarItem.RecurrenceId != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R511");

                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    511,
                    @"[In t:CalendarItemType Complex Type] The type of RecurrenceId is xs:dateTime ([XMLSCHEMA2]).");
            }

            if (calendarItem.DateTimeStampSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R154");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R154
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    154,
                    @"[In t:CalendarItemType Complex Type] The type of DateTimeStamp is xs:dateTime.");
            }

            if (calendarItem.StartSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R156");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R156
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    156,
                    @"[In t:CalendarItemType Complex Type] The type of Start is xs:dateTime.");
            }

            if (calendarItem.EndSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R158");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R158
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    158,
                    @"[In t:CalendarItemType Complex Type] The type of End is xs:dateTime.");
            }

            if (calendarItem.OriginalStartSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R160");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R160
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    160,
                    @"[In t:CalendarItemType Complex Type] The type of OriginalStart is xs:dateTime.");
            }

            if (calendarItem.IsAllDayEventSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R162");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R162
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    162,
                    @"[In t:CalendarItemType Complex Type] The type of IsAllDayEvent is xs:boolean ([XMLSCHEMA2]).");
            }

            if (calendarItem.LegacyFreeBusyStatusSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R164");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R164
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    164,
                    @"[In t:CalendarItemType Complex Type] The type of LegacyFreeBusyStatus is t:LegacyFreeBusyType ([MS-OXWSCDATA] section 2.2.3.17)");

                this.VerifyLegacyFreeBusyType(isSchemaValidated);
            }

            if (calendarItem.Location != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R166");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R166
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    166,
                    @"[In t:CalendarItemType Complex Type] The type of Location is xs:string.");
            }

            if (calendarItem.When != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R168");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R168
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    168,
                    @"[In t:CalendarItemType Complex Type] The type of When is xs:string.");
            }

            if (calendarItem.IsMeetingSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R170");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R170
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    170,
                    @"[In t:CalendarItemType Complex Type] The type of IsMeeting is xs:boolean.");
            }

            if (calendarItem.IsCancelledSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R172");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R172
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    172,
                    @"[In t:CalendarItemType Complex Type] The type of IsCancelled is xs:boolean.");
            }

            if (calendarItem.IsRecurringSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R174");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R174
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    174,
                    @"[In t:CalendarItemType Complex Type] The type of IsRecurring is xs:boolean.");
            }

            if (calendarItem.MeetingRequestWasSentSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R176");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R176
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    176,
                    @"[In t:CalendarItemType Complex Type] The type of MeetingRequestWasSent is xs:boolean.");
            }

            if (calendarItem.IsResponseRequestedSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R178");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R178
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    178,
                    @"[In t:CalendarItemType Complex Type] The type of IsResponseRequested is xs:boolean.");
            }

            if (calendarItem.CalendarItemType1Specified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R180");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R180
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    180,
                    @"[In t:CalendarItemType Complex Type] The type of CalendarItemType is t:CalendarItemTypeType.");

                this.VerifyCalendarItemTypeType(isSchemaValidated);
            }

            if (calendarItem.MyResponseTypeSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R182");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R182
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    182,
                    @"[In t:CalendarItemType Complex Type] The type of MyResponseType is t:ResponseTypeType (section 2.2.5.12).");

                this.VerifyResponseTypeType(isSchemaValidated);
            }

            if (calendarItem.Organizer != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R184");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R184
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    184,
                    @"[In t:CalendarItemType Complex Type] The type of Organizer is t:SingleRecipientType ([MS-OXWSCDATA] section 2.2.4.60).");

                this.VerifySingleRecipientType(calendarItem.Organizer, isSchemaValidated);
            }

            if (calendarItem.RequiredAttendees != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R513");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R513
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    513,
                    @"[In t:CalendarItemType Complex Type] The type of RequiredAttendees is t:NonEmptyArrayOfAttendeesType (section 2.2.4.19)");

                this.VeirfyNonEmptyArrayOfAttendeesType(calendarItem.RequiredAttendees, isSchemaValidated);
            }

            if (calendarItem.OptionalAttendees != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R188");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R188
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    188,
                    @"[In t:CalendarItemType Complex Type] The type of OptionalAttendees is t:NonEmptyArrayOfAttendeesType.");

                this.VeirfyNonEmptyArrayOfAttendeesType(calendarItem.OptionalAttendees, isSchemaValidated);
            }

            if (calendarItem.Resources != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R190");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R190
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    190,
                    @"[In t:CalendarItemType Complex Type] The type of Resources is t:NonEmptyArrayOfAttendeesType.");

                this.VeirfyNonEmptyArrayOfAttendeesType(calendarItem.Resources, isSchemaValidated);
            }

            if (calendarItem.Duration != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R200");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R200
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    200,
                    @"[In t:CalendarItemType Complex Type] The type of Duration is xs:string.");
            }

            if (calendarItem.ConflictingMeetingCountSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R192");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R192
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    192,
                    @"[In t:CalendarItemType Complex Type] The type of ConflictingMeetingCount is xs:int ([XMLSCHEMA2]).");
            }

            if (calendarItem.ConflictingMeetings != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R196");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R196
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    196,
                    @"[In t:CalendarItemType Complex Type] The type of ConflictingMeetings is t:NonEmptyArrayOfAllItemsType ([MS-OXWSCDATA] section 2.2.4.48).");
            }

            if (calendarItem.AdjacentMeetings != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R198");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R198
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    198,
                    @"[In t:CalendarItemType Complex Type] The type of AdjacentMeetings is t:NonEmptyArrayOfAllItemsType.");
            }

            if (calendarItem.AdjacentMeetingCountSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R194");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R194
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    194,
                    @"[In t:CalendarItemType Complex Type] The type of AdjacentMeetingCount is xs:int.");
            }

            if (calendarItem.TimeZone != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R202");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R202
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    202,
                    @"[In t:CalendarItemType Complex Type] The type of TimeZone is xs:string.");
            }

            if (calendarItem.AppointmentReplyTimeSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R204");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R204
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    204,
                    @"[In t:CalendarItemType Complex Type] The type of AppointmentReplyTime is xs:dateTime.");
            }
            
            if (calendarItem.AppointmentSequenceNumberSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R206");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R206
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    206,
                    @"[In t:CalendarItemType Complex Type] The type of AppointmentSequenceNumber is xs:int.");
            }

            if (calendarItem.AppointmentStateSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R208");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R208
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    208,
                    @"[In t:CalendarItemType Complex Type] The type of AppointmentState is xs:int.");
            }

            if (calendarItem.Recurrence != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R210");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R210
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    210,
                    @"[In t:CalendarItemType Complex Type] The type of Recurrence is t:RecurrenceType (section 2.2.4.25).");

                this.VerifyRecurrenceType(calendarItem.Recurrence, isSchemaValidated);
            }

            if (calendarItem.FirstOccurrence != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R515");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R515
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    515,
                    @"[In t:CalendarItemType Complex Type] The type of FirstOccurrence is t:OccurrenceInfoType (section 2.2.4.22).");

                this.VerifyOccurrenceInfoType(isSchemaValidated);
            }

            if (calendarItem.LastOccurrence != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R214");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R214
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    214,
                    @"[In t:CalendarItemType Complex Type] The type of LastOccurrence is t:OccurrenceInfoType.");

                this.VerifyOccurrenceInfoType(isSchemaValidated);
            }

            if (calendarItem.ModifiedOccurrences != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R216");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R216
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    216,
                    @"[In t:CalendarItemType Complex Type] The type of ModifiedOccurrences is t:NonEmptyArrayOfOccurrenceInfoType (section 2.2.4.21).");

                this.VerifyNonEmptyArrayOfOccurrenceInfoType(calendarItem.ModifiedOccurrences, isSchemaValidated);
            }

            if (calendarItem.DeletedOccurrences != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R218");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R218
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    218,
                    @"[In t:CalendarItemType Complex Type] The type of DeletedOccurrences is t:NonEmptyArrayOfDeletedOccurrencesType (section 2.2.4.20).");

                this.VerifyNonEmptyArrayOfDeletedOccurrencesType(calendarItem.DeletedOccurrences, isSchemaValidated);
            }

            if(calendarItem.MeetingTimeZone != null && Common.IsRequirementEnabled(911, this.Site))
            {
                this.VerifyTimeZoneType(calendarItem.MeetingTimeZone, isSchemaValidated);
            }

            if (calendarItem.RecurrenceIdSpecified == true)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R511");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R511
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    511,
                    @"[In t:CalendarItemType Complex Type] The type of RecurrenceId is xs:dateTime ([XMLSCHEMA2]).");
            }

            if (calendarItem.ConferenceTypeSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R226");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R226
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    226,
                    @"[In t:CalendarItemType Complex Type] The type of ConferenceType is xs:int.");
            }

            if (calendarItem.AllowNewTimeProposalSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R228");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R228
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    228,
                    @"[In t:CalendarItemType Complex Type] The type of AllowNewTimeProposal is xs:boolean.");
            }

            if (calendarItem.MeetingWorkspaceUrl != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R232");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R232
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    232,
                    @"[In t:CalendarItemType Complex Type] The type of MeetingWorkspaceUrl is xs:string.");
            }

            if (calendarItem.NetShowUrl != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R234");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R234
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    234,
                    @"[In t:CalendarItemType Complex Type] The type of NetShowUrl is xs:string.");
            }

            if (Common.IsRequirementEnabled(1278, this.Site) && calendarItem.InboxReminders != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R2004");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R2004
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    2004,
                    @"[In t:CalendarItemType Complex Type] The type of complex type ""InboxReminders"" is ""t: ArrayOfInboxReminderType(section 2.2.4.3)"".");

                this.VerifyArrayOfInboxReminderType(calendarItem.InboxReminders, isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify the MeetingCancellationMessageType structure.
        /// </summary>
        /// <param name="meetingCancellationMessageType">Represents a meeting cancellation</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyMeetingCancellationMessageType(MeetingCancellationMessageType meetingCancellationMessageType, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R260");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R260
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             260,
                @"[In t:MeetingCancellationMessageType Complex Type] [Its schema is]
<xs:complexType name=""MeetingCancellationMessageType"">
  <xs:complexContent>
    <xs:extension base=""t:MeetingMessageType"">
      <xs:sequence>
        <xs:element name=""Start"" type=""xs:dateTime"" minOccurs=""0""/>
        <xs:element name=""End"" type=""xs:dateTime"" minOccurs=""0""/>
        <xs:element name=""Location"" type=""xs:string"" minOccurs=""0""/>
        <xs:element name=""Recurrence"" type=""t:RecurrenceType"" minOccurs=""0""/>
        <xs:element name=""CalendarItemType"" type=""xs:string"" minOccurs=""0""/>
        <xs:element name=""EnhancedLocation"" type=""t:EnhancedLocationType"" minOccurs=""0""/>
      </xs:sequence>
    </xs:extension>
  </xs:complexContent>
</xs:complexType>");

            // MeetingMessageType is the base type of MeetingCancellationMessageType.
            this.VerifyMeetingMessageType(meetingCancellationMessageType, isSchemaValidated);
        }

        /// <summary>
        /// Verify the MeetingMessageType structure.
        /// </summary>
        /// <param name="meetingMessageType">Represents a meeting in the messaging server store</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyMeetingMessageType(MeetingMessageType meetingMessageType, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R262");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R262
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             262,
             @"[In t:MeetingMessageType Complex Type] [Its schema is]
<xs:complexType name=""MeetingMessageType"">
  <xs:complexContent>
    <xs:extension
      base=""t:MessageType""
    >
      <xs:sequence>
        <xs:element name=""AssociatedCalendarItemId""
          type=""t:ItemIdType""
          minOccurs=""0""
         />
        <xs:element name=""IsDelegated""
          type=""xs:boolean""
          minOccurs=""0""
         />
        <xs:element name=""IsOutOfDate""
          type=""xs:boolean""
          minOccurs=""0""
         />
        <xs:element name=""HasBeenProcessed""
          type=""xs:boolean""
          minOccurs=""0""
         />
        <xs:element name=""ResponseType""
          type=""t:ResponseTypeType""
          minOccurs=""0""
         />
        <xs:element name=""UID""
          type=""xs:string""
          minOccurs=""0""
         />
        <xs:element name=""RecurrenceId""
          type=""xs:dateTime""
          minOccurs=""0""
         />
        <xs:element name=""DateTimeStamp""
          type=""xs:dateTime""
          minOccurs=""0""
         /> 
        <xs:element name=""IsOrganizer"" type=""xs:boolean"" minOccurs=""0""/>
      </xs:sequence>
    </xs:extension>
  </xs:complexContent>");

            #region Verify MeetingMessageType structure
            if (meetingMessageType.AssociatedCalendarItemId != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R263");

                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    263,
                    @"[In t:MeetingMessageType Complex Type] The type of AssociatedCalendarItemId is t:ItemIdType ([MS-OXWSCORE] section 2.2.4.19).");
            }

            if (meetingMessageType.IsDelegatedSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R265");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R265
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    265,
                    @"[In t:MeetingMessageType Complex Type] The type of IsDelegated is xs:boolean ([XMLSCHEMA2]).");
            }

            if (meetingMessageType.IsOutOfDateSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R267");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R267
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    267,
                    @"[In t:MeetingMessageType Complex Type] The type of IsOutOfDate is xs:boolean.");
            }

            if (meetingMessageType.HasBeenProcessedSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R269");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R269
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    269,
                    @"[In t:MeetingMessageType Complex Type] The type of HasBeenProcessed is xs:boolean.");
            }

            if (meetingMessageType.ResponseTypeSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R271");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R271
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    271,
                    @"[In t:MeetingMessageType Complex Type] The type of ResponseType is t:ResponseTypeType (section 2.2.5.12).");
            }

            if (!string.IsNullOrEmpty(meetingMessageType.UID))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R274");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R274
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    274,
                    @"[In t:MeetingMessageType Complex Type] The type of UID is xs:string ([XMLSCHEMA2]).");
            }

            if (meetingMessageType.RecurrenceIdSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R276");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R276
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    276,
                    @"[In t:MeetingMessageType Complex Type] The type of RecurrenceId is xs:dateTime ([XMLSCHEMA2]).");
            }

            if (meetingMessageType.DateTimeStamp != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R278");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R278
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    278,
                    @"[In t:MeetingMessageType Complex Type] The type of DateTimeStamp is xs:dateTime.");
            }

            if (meetingMessageType.IsOrganizerSpecified)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R80004");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R80004
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    80004,
                    @"[In t:MeetingMessageType Complex Type] The type of IsOrganizer is xs:boolean ([XMLSCHEMA2]).");
            }
            #endregion

            if (meetingMessageType.ResponseTypeSpecified)
            {
                this.VerifyResponseTypeType(isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify the MeetingRequestMessageType structure
        /// </summary>
        /// <param name="meetingRequestMessage">Represents a meeting request in the messaging server store</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyMeetingRequestMessageType(MeetingRequestMessageType meetingRequestMessage, bool isSchemaValidated)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R281");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                281,
                @"[In t:MeetingRequestMessageType Complex Type] [Its schema is]
<xs:complexContent>
  <xs:extension base=""t:MeetingMessageType"">
    <xs:sequence>
      <!--- MeetingRequest properties -->
      <xs:element name=""MeetingRequestType"" type=""t:MeetingRequestTypeType"" minOccurs=""0""/>
      <xs:element name=""IntendedFreeBusyStatus"" type=""t:LegacyFreeBusyType"" minOccurs=""0""/>
      <!-- Calendar Properties of the associated meeting request -->
      <!-- Single and Occurrence only -->
      <xs:element name=""Start"" type=""xs:dateTime"" minOccurs=""0""/>
      <xs:element name=""End"" type=""xs:dateTime"" minOccurs=""0""/>
      <!-- Occurrence only -->
      <xs:element name=""OriginalStart"" type=""xs:dateTime"" minOccurs=""0""/>
      <xs:element name=""IsAllDayEvent"" type=""xs:boolean"" minOccurs=""0""/>
      <xs:element name=""LegacyFreeBusyStatus"" type=""t:LegacyFreeBusyType"" minOccurs=""0""/>
      <xs:element name=""Location"" type=""xs:string"" minOccurs=""0""/>
      <xs:element name=""When"" type=""xs:string"" minOccurs=""0""/>
      <xs:element name=""IsMeeting"" type=""xs:boolean"" minOccurs=""0""/>
      <xs:element name=""IsCancelled"" type=""xs:boolean"" minOccurs=""0""/>
      <xs:element name=""IsRecurring"" type=""xs:boolean"" minOccurs=""0""/>
      <xs:element name=""MeetingRequestWasSent"" type=""xs:boolean"" minOccurs=""0""/>
      <xs:element name=""CalendarItemType"" type=""t:CalendarItemTypeType"" minOccurs=""0""/>
      <xs:element name=""MyResponseType"" type=""t:ResponseTypeType"" minOccurs=""0""/>
      <xs:element name=""Organizer"" type=""t:SingleRecipientType"" minOccurs=""0""/>
      <xs:element name=""RequiredAttendees"" type=""t:NonEmptyArrayOfAttendeesType"" minOccurs=""0""/>
      <xs:element name=""OptionalAttendees"" type=""t:NonEmptyArrayOfAttendeesType"" minOccurs=""0""/>
      <xs:element name=""Resources"" type=""t:NonEmptyArrayOfAttendeesType"" minOccurs=""0""/>
      <!-- Conflicting and adjacent meetings -->
      <xs:element name=""ConflictingMeetingCount"" type=""xs:int"" minOccurs=""0""/>
      <xs:element name=""AdjacentMeetingCount"" type=""xs:int"" minOccurs=""0""/>
      <xs:element name=""ConflictingMeetings"" type=""t:NonEmptyArrayOfAllItemsType"" minOccurs=""0""/>
      <xs:element name=""AdjacentMeetings"" type=""t:NonEmptyArrayOfAllItemsType"" minOccurs=""0""/>
      <xs:element name=""Duration"" type=""xs:string"" minOccurs=""0""/>
      <xs:element name=""TimeZone"" type=""xs:string"" minOccurs=""0""/>
      <xs:element name=""AppointmentReplyTime"" type=""xs:dateTime"" minOccurs=""0""/>
      <xs:element name=""AppointmentSequenceNumber"" type=""xs:int"" minOccurs=""0""/>
      <xs:element name=""AppointmentState"" type=""xs:int"" minOccurs=""0""/>
      <!-- Recurrence specific data, only valid if CalendarItemType is RecurringMaster -->
      <xs:element name=""Recurrence"" type=""t:RecurrenceType"" minOccurs=""0""/>
      <xs:element name=""FirstOccurrence"" type=""t:OccurrenceInfoType"" minOccurs=""0""/>
      <xs:element name=""LastOccurrence"" type=""t:OccurrenceInfoType"" minOccurs=""0""/>
      <xs:element name=""ModifiedOccurrences"" type=""t:NonEmptyArrayOfOccurrenceInfoType"" minOccurs=""0""/>
      <xs:element name=""DeletedOccurrences"" type=""t:NonEmptyArrayOfDeletedOccurrencesType"" minOccurs=""0""/>
      <xs:element name=""MeetingTimeZone"" type=""t:TimeZoneType"" minOccurs=""0""/>
      <xs:element name=""StartTimeZone"" type=""t:TimeZoneDefinitionType"" minOccurs=""0""/>
      <xs:element name=""EndTimeZone"" type=""t:TimeZoneDefinitionType"" minOccurs=""0""/>
      <xs:element name=""ConferenceType"" type=""xs:int"" minOccurs=""0""/>
      <xs:element name=""AllowNewTimeProposal"" type=""xs:boolean"" minOccurs=""0""/>
      <xs:element name=""IsOnlineMeeting"" type=""xs:boolean"" minOccurs=""0""/>
      <xs:element name=""MeetingWorkspaceUrl"" type=""xs:string"" minOccurs=""0""/>
      <xs:element name=""NetShowUrl"" type=""xs:string"" minOccurs=""0""/>
      <xs:element name=""EnhancedLocation"" type=""t:EnhancedLocationType"" minOccurs=""0""/>
      <xs:element name=""ChangeHighlights"" type=""t:ChangeHighlightsType"" minOccurs=""0""/>
      <xs:element name=""StartWallClock"" type=""xs:dateTime"" minOccurs=""0"" maxOccurs=""1""/>
      <xs:element name=""EndWallClock"" type=""xs:dateTime"" minOccurs=""0"" maxOccurs=""1""/>
      <xs:element name=""StartTimeZoneId"" type=""xs:string"" minOccurs=""0"" maxOccurs=""1""/>
      <xs:element name=""EndTimeZoneId"" type=""xs:string"" minOccurs=""0"" maxOccurs=""1""/>
    </xs:sequence>
  </xs:extension>
</xs:complexContent>
</xs:complexType>");

            #region Verify MeetingRequestMessageType structure

            if (meetingRequestMessage.Resources != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R318");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R318
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    318,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of Resources is t:NonEmptyArrayOfAttendeesType.");
            }

            if (meetingRequestMessage.MeetingRequestTypeSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R282");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R282
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    282,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of MeetingRequestType is t:MeetingRequestTypeType (section 2.2.5.9).");

                this.VerifyMeetingRequestTypeType(isSchemaValidated);
            }

            if (meetingRequestMessage.IntendedFreeBusyStatusSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R284");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R284
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    284,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of IntendedFreeBusyStatus is t:LegacyFreeBusyType ([MS-OXWSCDATA] section 2.2.3.17).");
            }

            if (meetingRequestMessage.Start != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R499");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R499
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    499,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of Start is  xs:dateTime ([XMLSCHEMA2])");
            }

            if (meetingRequestMessage.Start != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R286");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R286
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    286,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of End is xs:dateTime.");
            }

            if (!string.IsNullOrEmpty(meetingRequestMessage.Location))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R296");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R296
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    296,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of Location is xs:string ([XMLSCHEMA2]).");
            }

            if (meetingRequestMessage.IsMeetingSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R300");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R300
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    300,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of IsMeeting is xs:boolean.");
            }

            if (meetingRequestMessage.IsCancelledSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R302");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R302
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    302,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of IsCancelled is xs:boolean.");
            }

            if (meetingRequestMessage.IsRecurringSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R304");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R304
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    304,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of IsRecurring is xs:boolean.");
            }

            if (meetingRequestMessage.MeetingRequestWasSentSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R306");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R306
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    306,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of MeetingRequestWasSent is xs:boolean.");
            }

            if (meetingRequestMessage.CalendarItemTypeSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R308");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R308
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    308,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of CalendarItemType is t:CalendarItemTypeType.(section 2.2.4.6)");

                this.VerifyCalendarItemTypeType(isSchemaValidated);
            }

            if (meetingRequestMessage.Organizer != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R312");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R312
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    312,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of Organizer is t:SingleRecipientType ([MS-OXWSCDATA] section 2.2.4.60).");

                this.VerifySingleRecipientType(meetingRequestMessage.Organizer, isSchemaValidated);
            }

            if (meetingRequestMessage.RequiredAttendees != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R314");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R314
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    314,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of RequiredAttendees is t:NonEmptyArrayOfAttendeesType (section 2.2.4.19).");

                this.VeirfyNonEmptyArrayOfAttendeesType(meetingRequestMessage.RequiredAttendees, isSchemaValidated);
            }

            if (meetingRequestMessage.OptionalAttendees != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R316");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R316
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    316,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of OptionalAttendees is t:NonEmptyArrayOfAttendeesType.");

                this.VeirfyNonEmptyArrayOfAttendeesType(meetingRequestMessage.OptionalAttendees, isSchemaValidated);
            }

            if (meetingRequestMessage.ConflictingMeetingCountSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R320");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R320
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    320,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of ConflictingMeetingCount is xs:int ([XMLSCHEMA2]).");
            }

            if (meetingRequestMessage.AdjacentMeetingCountSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R322");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R322
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    322,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of AdjacentMeetingCount is xs:int.");
            }

            if (meetingRequestMessage.ConflictingMeetings != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R324");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R324
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    324,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of ConflictingMeetings is t:NonEmptyArrayOfAllItemsType ([MS-OXWSCDATA] section 2.2.4.42)");

                this.VerifyNonEmptyArrayOfAllItemsType(isSchemaValidated);
            }

            if (meetingRequestMessage.AdjacentMeetings != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R326");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R326
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    326,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of AdjacentMeetings is t:NonEmptyArrayOfAllItemsType.");

                this.VerifyNonEmptyArrayOfAllItemsType(isSchemaValidated);
            }

            if (!string.IsNullOrEmpty(meetingRequestMessage.Duration))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R328");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R328
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    328,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of Duration is xs:string.");
            }

            if (!string.IsNullOrEmpty(meetingRequestMessage.TimeZone))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R330");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R330
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    330,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of TimeZone is xs:string.");
            }

            if (meetingRequestMessage.AppointmentSequenceNumberSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R334");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R334
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    334,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of AppointmentSequenceNumber is xs:int.");
            }

            if (meetingRequestMessage.AppointmentStateSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R336");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R336
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    336,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of AppointmentState is xs:int.");
            }

            if (meetingRequestMessage.Recurrence != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R338");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R338
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    338,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of Recurrence is t:RecurrenceType (section 2.2.4.25).");

                this.VerifyRecurrenceType(meetingRequestMessage.Recurrence, isSchemaValidated);
            }

            if (meetingRequestMessage.FirstOccurrence != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R340");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R340
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    340,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of FirstOccurrence is t:OccurrenceInfoType (section 2.2.4.22).");

                this.VerifyOccurrenceInfoType(isSchemaValidated);
            }

            if (meetingRequestMessage.LastOccurrence != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R342");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R342
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    342,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of LastOccurrence is t:OccurrenceInfoType.");

                this.VerifyOccurrenceInfoType(isSchemaValidated);
            }

            if (meetingRequestMessage.ModifiedOccurrences != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R344");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R344
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    344,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of ModifiedOccurrences is t:NonEmptyArrayOfOccurrenceInfoType (section 2.2.4.21).");
            }

            if (meetingRequestMessage.DeletedOccurrences != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R346");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R346
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    346,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of DeletedOccurrences is t:NonEmptyArrayOfDeletedOccurrencesType (section 2.2.4.20).");
            }

            if (meetingRequestMessage.IsOnlineMeetingSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R358");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R358
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    358,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of IsOnlineMeeting is xs:boolean.");
            }

            if (!string.IsNullOrEmpty(meetingRequestMessage.MeetingWorkspaceUrl))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R360");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R360
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    360,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of MeetingWorkspaceUrl is xs:string.");
            }

            if (!string.IsNullOrEmpty(meetingRequestMessage.NetShowUrl))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R362");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R362
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    362,
                    @"[In t:MeetingRequestMessageType Complex Type] The type of NetShowUrl is xs:string.");
            }
            #endregion

            if (meetingRequestMessage != null)
            {
                this.VerifyMeetingMessageType(meetingRequestMessage, isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify the MeetingResponseMessageType structure.
        /// </summary>
        /// <param name="meetingResponseMessageType">Represents a meeting response in the messaging server store</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyMeetingResponseMessageType(MeetingResponseMessageType meetingResponseMessageType, bool isSchemaValidated)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R365");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                365,
                @"[In t:MeetingResponseMessageType Complex Type] [Its schema is]
<xs:complexType name=""MeetingResponseMessageType"">
  <xs:complexContent>
    <xs:extension base=""t:MeetingMessageType"">
      <xs:sequence>
        <xs:element name=""Start"" type=""xs:dateTime"" minOccurs=""0""/>
        <xs:element name=""End"" type=""xs:dateTime"" minOccurs=""0""/>
        <xs:element name=""Location"" type=""xs:string"" minOccurs=""0""/>
        <xs:element name=""Recurrence"" type=""t:RecurrenceType"" minOccurs=""0""/>
        <xs:element name=""CalendarItemType"" type=""xs:string"" minOccurs=""0""/>
        <xs:element name=""ProposedStart"" type=""xs:dateTime"" minOccurs=""0""/>
        <xs:element name=""ProposedEnd"" type=""xs:dateTime"" minOccurs=""0""/>
        <xs:element name=""EnhancedLocation"" type=""t:EnhancedLocationType"" minOccurs=""0""/>
      </xs:sequence>
    </xs:extension>
  </xs:complexContent>
</xs:complexType>");

            // MeetingMessageType is the base type of MeetingResponseMessageType
            this.VerifyMeetingMessageType(meetingResponseMessageType, isSchemaValidated);
        }

        /// <summary>
        /// Verify the NonEmptyArrayOfAttendeesType structure
        /// </summary>
        /// <param name="arrayOfAttendee">Contains a list representing attendees and resources for a meeting</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyArrayOfInboxReminderType(InboxReminderType[] arrayOfInboxReminder, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1345");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1345
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             1345,
             @"[In t:ArrayOfInboxReminderType] [its schema is] <xs:complexType name=""ArrayOfInboxReminderType"" >
  < xs:sequence >
       < xs:element name = ""InboxReminder""
            type = ""t:InboxReminderType"" minOccurs = ""0"" maxOccurs = ""unbounded"" />
  </ xs:sequence >
 </ xs:complexType > ");

            foreach (InboxReminderType inboxReminderType in arrayOfInboxReminder)
            {
                this.VerifyInboxReminderType(inboxReminderType, isSchemaValidated);
            }
        }


        /// <summary>
        /// Verify the NonEmptyArrayOfAttendeesType structure
        /// </summary>
        /// <param name="arrayOfAttendee">Contains a list representing attendees and resources for a meeting</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VeirfyNonEmptyArrayOfAttendeesType(AttendeeType[] arrayOfAttendee, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R367");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R367
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             367,
             @"[In t:NonEmptyArrayOfAttendeesType Complex Type] [its schema is] <xs:complexType name=""NonEmptyArrayOfAttendeesType"">
                  <xs:sequence>
                    <xs:element name=""Attendee""
                      type=""t:AttendeeType""
                      maxOccurs=""unbounded""
                     />
                  </xs:sequence>
                </xs:complexType>");

            foreach (AttendeeType attendeeType in arrayOfAttendee)
            {
                this.VerifyAttendeeType(attendeeType, isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify the NonEmptyArrayOfDeletedOccurrencesType structure.
        /// </summary>
        /// <param name="arrayOfDeletedOccurrences">Contains a list of deleted occurrences of a recurring calendar item</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyNonEmptyArrayOfDeletedOccurrencesType(DeletedOccurrenceInfoType[] arrayOfDeletedOccurrences, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R371");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R371
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             371,
             @"[In t:NonEmptyArrayOfDeletedOccurrencesType Complex Type] [its schema is] <xs:complexType name=""NonEmptyArrayOfDeletedOccurrencesType"">
                  <xs:sequence>
                    <xs:element name=""DeletedOccurrence""
                      type=""t:DeletedOccurrenceInfoType""
                      maxOccurs=""unbounded""
                     />
                  </xs:sequence>
                </xs:complexType>");

            if (arrayOfDeletedOccurrences != null && arrayOfDeletedOccurrences.Length > 0)
            {
                this.VerifyDeletedOccurrenceInfoType(isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify TimeZoneType structure.
        /// </summary>
        /// <param name="timeZoneType">Represents the time zone of the location where a meeting is hosted</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyTimeZoneType(TimeZoneType timeZoneType, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R911");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R911
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             911,
             @"[In Appendix C: Product Behavior] Implementation does support the TimeZoneType complex type which does not represent the time zone of the location where a meeting is hosted. 
<xs:complexType name=""TimeZoneType"">
  <xs:sequence
    minOccurs=""0""
  >
    <xs:element name=""BaseOffset""
      type=""xs:duration""
     />
    <xs:sequence
      minOccurs=""0""
    >
      <xs:element name=""Standard""
        type=""t:TimeChangeType""
       />
      <xs:element name=""Daylight""
        type=""t:TimeChangeType""
       />
    </xs:sequence>
  </xs:sequence>
  <xs:attribute name=""TimeZoneName""
    type=""xs:string""
    use=""optional""
  />
</xs:complexType> (<57> Section 2.2.4.29:  Only Exchange 2007 supports the TimeZoneType complex type.)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R406");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R406
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             406,
             @"[In t:TimeZoneType Complex Type] The type of BaseOffset is xs:duration ([XMLSCHEMA2]).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R408");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R408
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             408,
             @"[In t:TimeZoneType Complex Type] The type of Standard is t:TimeChangeType (section 2.2.4.28).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R410");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R410
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             410,
             @"[In t:TimeZoneType Complex Type] The type of Daylight is t:TimeChangeType.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R584");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R584
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             584,
             @"[In t:TimeZoneType Complex Type] The type of TimeZoneName is xs:string ([XMLSCHEMA2]).");

            if(timeZoneType != null)
            {
                this.VerifyTimeChangeType(isSchemaValidated);
            }
        }

       /// <summary>
        /// Verify TimeChangeType structure.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyTimeChangeType( bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R397");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R397
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             397,
             @"[In t:TimeChangeType Complex Type] [its schema is] <xs:complexType name=""TimeChangeType"">
  <xs:sequence>
    <xs:element name=""Offset""
      type=""xs:duration""
     />
    <xs:group
      minOccurs=""0""
      ref=""t:TimeChangePatternTypes""
     />
    <xs:element name=""Time""
      type=""xs:time""
     />
  </xs:sequence>
  <xs:attribute name=""TimeZoneName""
    type=""xs:string""
    use=""optional""
   />
</xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R398");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R398
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             398,
             @"[In t:TimeChangeType Complex Type] The type of Offset is xs:duration ([XMLSCHEMA2]).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R400");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R400
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             400,
             @"[In t:TimeChangeType Complex Type] The type of Time is xs:time ([XMLSCHEMA2]).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R402");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R402
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             402,
             @"[In t:TimeChangeType Complex Type] The type of TimeZoneName is xs:string ([XMLSCHEMA2]).");
        }

        /// <summary>
        /// Verify NonEmptyArrayOfOccurrenceInfoType structure.
        /// </summary>
        /// <param name="arrayOfOccurrenceInfoType">Contains a list of modified occurrences of a recurring calendar item</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyNonEmptyArrayOfOccurrenceInfoType(OccurrenceInfoType[] arrayOfOccurrenceInfoType, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R375");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R375
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             375,
             @"[In t:NonEmptyArrayOfOccurrenceInfoType Complex Type] [its schema is] <xs:complexType name=""NonEmptyArrayOfOccurrenceInfoType"">
                      <xs:sequence>
                        <xs:element name=""Occurrence""
                          type=""t:OccurrenceInfoType""
                          maxOccurs=""unbounded""
                         />
                      </xs:sequence>
                    </xs:complexType>");

            if (arrayOfOccurrenceInfoType != null)
            {
                this.VerifyOccurrenceInfoType(isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify occurrenceInfoType structure.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyOccurrenceInfoType(bool isSchemaValidated)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R376");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                376,
                @"[In t:NonEmptyArrayOfOccurrenceInfoType Complex Type] The type of Occurrence is t:OccurrenceInfoType (section 2.2.4.18).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R379");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R379
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                379,
                @"[In t:OccurrenceInfoType Complex Type] [its schema is] <xs:complexType name=""OccurrenceInfoType"">
                          <xs:sequence>
                            <xs:element name=""ItemId""
                              type=""t:ItemIdType""
                             />
                            <xs:element name=""Start""
                              type=""xs:dateTime""
                             />
                            <xs:element name=""End""
                              type=""xs:dateTime""
                             />
                            <xs:element name=""OriginalStart""
                              type=""xs:dateTime""
                             />
                          </xs:sequence>
                        </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R380");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R380
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                380,
                @"[In t:OccurrenceInfoType Complex Type] The type of ItemId is t:ItemIdType ([MS-OXWSCORE] section 2.2.4.19).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R382");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R382
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                382,
                @"[In t:OccurrenceInfoType Complex Type] The type of Start is xs:dateTime ([XMLSCHEMA2]).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R384");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R384
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                384,
                @"[In t:OccurrenceInfoType Complex Type] The type of End is xs:dateTime.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R386");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R386
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                386,
                @"[In t:OccurrenceInfoType Complex Type] The type of OriginalStart is xs:dateTime.");
        }

        /// <summary>
        /// Verify the RecurrenceType structure.
        /// </summary>
        /// <param name="recurrenceType">Specified the recurrence pattern and recurrence range for calendar items and meeting requests.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyRecurrenceType(RecurrenceType recurrenceType, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R389");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R389
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             389,
             @"[In t:RecurrenceType Complex Type] [its schema is] <xs:complexType name=""RecurrenceType"">
                      <xs:sequence>
                        <xs:group
                          ref=""t:RecurrencePatternTypes""
                         />
                        <xs:group
                          ref=""t:RecurrenceRangeTypes""
                         />
                      </xs:sequence>
                    </xs:complexType>");

            if (recurrenceType.Item != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1343");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1343
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    "MS-OXWSCDATA",
                    1343,
                    @"[In t:RecurrencePatternTypes Group] The group [t:RecurrencePatternTypes] is define as follow:
    <xs:group name=""t:RecurrencePatternTypes"">
        <xs:sequence>
        <xs:choice>
            <xs:element name=""t:RelativeYearlyRecurrence""
            type=""t:RelativeYearlyRecurrencePatternType""
            />
            <xs:element name=""AbsoluteYearlyRecurrence""
            type=""t:AbsoluteYearlyRecurrencePatternType""
            />
            <xs:element name=""RelativeMonthlyRecurrence""
            type=""t:RelativeMonthlyRecurrencePatternType""
            />
            <xs:element name=""AbsoluteMonthlyRecurrence""
            type=""t:AbsoluteMonthlyRecurrencePatternType""
            />
            <xs:element name=""WeeklyRecurrence""
            type=""t:WeeklyRecurrencePatternType""
            />
            <xs:element name=""DailyRecurrence""
            type=""t:DailyRecurrencePatternType""
            />
        </xs:choice>
        </xs:sequence>
    </xs:group>");

                this.VerifyRecurrencePatternTypes(recurrenceType.Item, isSchemaValidated);
            }

            if (recurrenceType.Item1 != null)
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

                this.VerifyRecurrenceRangeTypes(recurrenceType.Item1, isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify the CalendarItemTypeType schema.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyCalendarItemTypeType(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R51");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R51
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                51,
                @"[In t:CalendarItemTypeType Simple Type] [its schema is] <xs:simpleType name=""CalendarItemTypeType"">
                  <xs:restriction
                    base=""xs:string""
                  >
                    <xs:enumeration
                      value=""Single""
                     />
                    <xs:enumeration
                      value=""Occurrence""
                     />
                    <xs:enumeration
                      value=""Exception""
                     />
                    <xs:enumeration
                      value=""RecurringMaster""
                     />
                  </xs:restriction>
                </xs:simpleType>");
        }

        /// <summary>
        /// Verify MeetingRequestTypeType structure.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyMeetingRequestTypeType(bool isSchemaValidated)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R67");

            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                67,
                @"[In t:MeetingRequestTypeType Simple Type] [its schema is] <xs:simpleType name=""MeetingRequestTypeType"">
  <xs:restriction
    base=""xs:string""
  >
    <xs:enumeration
      value=""None""
     />
    <xs:enumeration
      value=""FullUpdate""
     />
    <xs:enumeration
      value=""InformationUpdate""
     />
    <xs:enumeration
      value=""NewMeetingRequest""
     />
    <xs:enumeration
      value=""Outdated""
     />
    <xs:enumeration
      value=""SilentUpdate""
     />
    <xs:enumeration
      value=""PrincipalWantsCopy""
     />
  </xs:restriction>
</xs:simpleType>");
        }

        /// <summary>
        /// Verify the ResponseTypeType schema.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyResponseTypeType(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R77");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R77
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             77,
             @"[In t:ResponseTypeType Simple Type] [its schema is] <xs:simpleType name=""ResponseTypeType"">
                  <xs:restriction
                    base=""xs:string""
                  >
                    <xs:enumeration
                      value=""Unknown""
                     />
                    <xs:enumeration
                      value=""Organizer""
                     />
                    <xs:enumeration
                      value=""Tentative""
                     />
                    <xs:enumeration
                      value=""Accept""
                     />
                    <xs:enumeration
                      value=""Decline""
                     />
                    <xs:enumeration
                      value=""NoResponseReceived""
                     />
                  </xs:restriction>
                </xs:simpleType>");
        }

        /// <summary>
        /// Verify LegacyFreeBusyType structure.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyLegacyFreeBusyType(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R155");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R155.
            // The LegacyFreeBusyType is the type of child elements LegacyFreeBusyStatus for CalendarItemType and MeetingRequestMessageType. If schema verification passes, it indicates the schemas of contained elements are also verified.
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             "MS-OXWSCDATA",
             155,
             @"[In t:LegacyFreeBusyType Simple Type] The type [LegacyFreeBusyType] is defined as follow:
<xs:simpleType name=""LegacyFreeBusyType"">
    <xs:restriction base=""xs:string"">
        <xs:enumeration value=""Busy""/>
        <xs:enumeration value=""Free""/>
        <xs:enumeration value=""NoData""/>
        <xs:enumeration value=""OOF""/>
        <xs:enumeration value=""Tentative""/>
        <xs:enumeration value=""WorkingElsewhere""/>
    </xs:restriction>
</xs:simpleType>");
        }

        /// <summary>
        /// Verify MailboxTypeType structure.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyMailboxTypeType(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R163");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R163.
            // The MailboxTypeType is the type of element MailboxType contained in CalendarItemType and MeetingRequestMessageType. If schema verification passes, it indicates the schemas of contained elements are also verified.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                163,
                @"[In t:MailboxTypeType Simple Type] The type [MailboxTypeType] is defined as follow:
                    <xs:simpleType name=""MailboxTypeType"">
                      <xs:restriction base=""xs:string"">
                        <xs:enumeration value=""Unknown""/>
                        <xs:enumeration value=""OneOff""/>
                        <xs:enumeration value=""Contact""/>
                        <xs:enumeration value=""Mailbox""/>
                        <xs:enumeration value=""PrivateDL""/>
                        <xs:enumeration value=""PublicDL""/>
                        <xs:enumeration value=""PublicFolder""/>
                      </xs:restriction>
                    </xs:simpleType>");
        }

        /// <summary>
        /// Verify NonEmptyStringType structure.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyNonEmptyStringType(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R189");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R189.
            // The NonEmptyStringType is the type of child elements Attendee for CalendarItemType. If isSchemaValidated is true, it indicates the schemas of contained elements are also verified.
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             "MS-OXWSCDATA",
             189,
             @"[In t:NonEmptyStringType Simple Type] The type [NonEmptyStringType] is defined as follow:
                    <xs:simpleType name=""NonEmptyStringType"">
                    <xs:restriction
                    base=""xs:string""
                    >
                    <xs:minLength
                        value=""1""
                        />
                    </xs:restriction>
                </xs:simpleType>");
        }

        /// <summary>
        /// Verify DeletedOccurrenceInfoType structure.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyDeletedOccurrenceInfoType(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R372");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R372
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             372,
             @"[In t:NonEmptyArrayOfDeletedOccurrencesType Complex Type] The type of DeletedOccurrence is t:DeletedOccurrenceInfoType ([MS-OXWSCDATA] section 2.2.4.22).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R111");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1111.
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             "MS-OXWSCDATA",
             1111,
             @"[In t:DeletedOccurrenceInfoType Complex Type] The type [DeletedOccurrenceInfoType] is defined as follow:
                    <xs:complexType name=""DeletedOccurrenceInfoType"">
                    <xs:sequence>
                    <xs:element name=""Start""
                        type=""xs:dateTime""
                        />
                    </xs:sequence>
                </xs:complexType>");
        }

        /// <summary>
        /// Verify EmailReminderChangeType structure.
        /// </summary>
        /// <param name="emailReminderchangeType">The EmaiReminderChangeType object.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyEmailReminderChangeType(EmailReminderChangeType emailReminderchangeType, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1310");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1310
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                1310,
                @"[In Appendix C: Product Behavior] Implementation does support the EmailReminderChangeType simple type, which specifies the type of the change. (Exchange 2016 and above follow this behavior.)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1361");

            // Verify MS-OXWSCDATA requirement: MS-OXWSMTGS_R1361
            // The EmailReminderchangeType is the type of OccurrenceChange contained in InboxReminderType. If schema verification passes, it indicates the schemas of contained elements are also verified.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                1361,
                @"[In t:EmailReminderChangeType] [its schema is] <xs:simpleType name=""EmailReminderChangeType"" >
   < xs:restriction base = ""xs:string"" >
    < xs:enumeration value = ""None"" />
    < xs:enumeration value = ""Added"" />
    < xs:enumeration value = ""Override"" />
    < xs:enumeration value = ""Deleted"" />
   </ xs:restriction >
  </ xs:simpleType > ");
        }

        /// <summary>
        /// Verify EmailReminderSendOption structure.
        /// </summary>
        /// <param name="emailReminderSendOption">The EmailReminderSendOption object.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyEmailReminderSendOption(EmailReminderSendOption emailReminderSendOption, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1312");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1312
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                1312,
                @"[In Appendix C: Product Behavior] Implementation does support the EmailReminderSendOption simple type, which specifies the send options for the reminder. (Exchange 2016 and above follow this behavior.)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1362");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1362
            // The EmailReminderchangeType is the type of OccurrenceChange contained in InboxReminderType. If schema verification passes, it indicates the schemas of contained elements are also verified.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                1362,
                @"[In t:EmailReminderSendOption] [its schema is] <xs:simpleType name=""EmailReminderSendOption"" >
   < xs:restriction base = ""xs:string"" >
     < xs:enumeration value = ""NotSet"" />
     < xs:enumeration value = ""User"" />
     < xs:enumeration value = ""AllAttendees"" />
     < xs:enumeration value = ""Staff"" />
     < xs:enumeration value = ""Customer"" />
   </ xs:restriction >
 </ xs:simpleType >
c");
        }

        /// <summary>
        /// Verify EmailAddressType structure.
        /// </summary>
        /// <param name="emailAddressType">The EmailAddressType object.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyEmailAddressType(EmailAddressType emailAddressType, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1147");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1147
            // The EmailAddressType is the type of Mailbox contained in CalendarItemType and MeetingRequestMessageType. If schema verification passes, it indicates the schemas of contained elements are also verified.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1147,
                @"[In t:EmailAddessType Complex Type] The type [EmailAddressType] is defined as follow:
<xs:complexType name=""EmailAddressType"">
  <xs:complexContent>
    <xs:extension
      base=""t:BaseEmailAddressType""
    >
      <xs:sequence>
        <xs:element name=""Name""
          type=""xs:string""
          minOccurs=""0""
         />
        <xs:element name=""EmailAddress""
          type=""t:NonEmptyStringType""
          minOccurs=""0""
         />
        <xs:element name=""RoutingType""
          type=""t:NonEmptyStringType""
          minOccurs=""0""
         />
        <xs:element name=""MailboxType""
          type=""t:MailboxTypeType""
          minOccurs=""0""
         />
        <xs:element name=""ItemId""
          type=""t:ItemIdType""
          minOccurs=""0""
         />
        <xs:element name=""OriginalDisplayName"" 
          type=""xs:string"" 
          minOccurs=""0""/>
      </xs:sequence>
    </xs:extension>
  </xs:complexContent>
</xs:complexType>");

            if (emailAddressType.MailboxTypeSpecified == true)
            {
                this.VerifyMailboxTypeType(isSchemaValidated);
            }

            if (emailAddressType.EmailAddress != null || emailAddressType.RoutingType != null)
            {
                this.VerifyNonEmptyStringType(isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify NonEmptyArrayOfAllItemsType structure.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyNonEmptyArrayOfAllItemsType(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1205");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1205
            Site.CaptureRequirementIfIsTrue(
             isSchemaValidated,
             "MS-OXWSCDATA",
             1205,
             @"[In t:NonEmptyArrayOfAllItemsType Complex Type] The type [NonEmptyArrayOfAllItemsType] is defined as follow:
                        <xs:complexType name=""NonEmptyArrayOfAllItemsType"">
                          <xs:sequence>
                            <xs:choice
                              minOccurs=""1""
                              maxOccurs=""unbounded""
                            >
                              <xs:element name=""Item""
                                type=""t:ItemType""
                               />
                              <xs:element name=""Message""
                                type=""t:MessageType""
                               />
                              <xs:element name=""CalendarItem""
                                type=""t:CalendarItemType""
                               />
                              <xs:element name=""Contact""
                                type=""t:ContactItemType""
                               />
                              <xs:element name=""DistributionList""
                                type=""t:DistributionListType""
                               />
                              <xs:element name=""MeetingMessage""
                                type=""t:MeetingMessageType""
                               />
                              <xs:element name=""MeetingRequest""
                                type=""t:MeetingRequestMessageType""
                               />
                              <xs:element name=""MeetingResponse""
                                type=""t:MeetingResponseMessageType""
                               />
                              <xs:element name=""MeetingCancellation""
                                type=""t:MeetingCancellationMessageType""
                               />
                              <xs:element name=""Task""
                                type=""t:TaskType""
                               />
                              <xs:element name=""PostItem""
                                type=""t:PostItemType""
                               />
                              <xs:element name=""ReplyToItem""
                                type=""t:ReplyToItemType""
                               />
                              <xs:element name=""ForwardItem""
                                type=""t:ForwardItemType""
                               />
                              <xs:element name=""ReplyAllToItem""
                                type=""t:ReplyAllToItemType""
                               />
                              <xs:element name=""AcceptItem""
                                type=""t:AcceptItemType""
                               />
                              <xs:element name=""TentativelyAcceptItem""
                                type=""t:TentativelyAcceptItemType""
                               />
                              <xs:element name=""DeclineItem""
                                type=""t:DeclineItemType""
                               />
                              <xs:element name=""CancelCalendarItem""
                                type=""t:CancelCalendarItemType""
                               />
                              <xs:element name=""RemoveItem""
                                type=""t:RemoveItemType""
                               />
                              <xs:element name=""SuppressReadReceipt""
                                type=""t:SuppressReadReceiptType""
                               />
                              <xs:element name=""PostReplyItem""
                                type=""t:PostReplyItemType""
                               />
                              <xs:element name=""AcceptSharingInvitation""
                                type=""t:AcceptSharingInvitationType""
                               />
                            </xs:choice>
                          </xs:sequence>
                        </xs:complexType>");
        }

        /// <summary>
        /// Verify SingleRecipientType structure.
        /// </summary>
        /// <param name="singleRecipientType">Specifies the e-mail address information for a single message recipient.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifySingleRecipientType(SingleRecipientType singleRecipientType, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1292");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1292
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1292,
                @"[In t:SingleRecipientType Complex Type] The type [SingleRecipientType] is defined as follow:
                    <xs:complexType name=""SingleRecipientType"">
                    <xs:choice>
                    <xs:element name=""Mailbox""
                        type=""t:EmailAddressType""
                        />
                    </xs:choice>
                </xs:complexType>");

            // The "Mailbox" is renamed to "Item" in the auto-generated ItemStack.
            this.VerifyEmailAddressType(singleRecipientType.Item, isSchemaValidated);
        }

        /// <summary>
        /// Verify RecurrencePatternBaseType structure.
        /// </summary>
        /// <param name="recurrencePatternBaseType">Specifies the recurrence pattern for calendar items and meeting requests.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyRecurrencePatternTypes(RecurrencePatternBaseType recurrencePatternBaseType, bool isSchemaValidated)
        {
            RelativeYearlyRecurrencePatternType relativeYearlyRecurrencePatternType = recurrencePatternBaseType as RelativeYearlyRecurrencePatternType;
            if (relativeYearlyRecurrencePatternType != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1259");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1259
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
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

                this.VerifyDayOfWeekIndexType(isSchemaValidated);
                this.VerifyMonthNamesType(isSchemaValidated);
            }

            AbsoluteYearlyRecurrencePatternType absoluteYearlyRecurrencePatternType = recurrencePatternBaseType as AbsoluteYearlyRecurrencePatternType;
            if (absoluteYearlyRecurrencePatternType != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R999");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R999
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
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

                this.VerifyMonthNamesType(isSchemaValidated);
            }

            RelativeMonthlyRecurrencePatternType relativeMonthlyRecurrencePatternType = recurrencePatternBaseType as RelativeMonthlyRecurrencePatternType;
            if (relativeMonthlyRecurrencePatternType != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1255");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1255
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
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

                this.VerifyDayOfWeekType(isSchemaValidated);
                this.VerifyDayOfWeekIndexType(isSchemaValidated);
            }

            AbsoluteMonthlyRecurrencePatternType absoluteMonthlyRecurrencePatternType = recurrencePatternBaseType as AbsoluteMonthlyRecurrencePatternType;
            if (absoluteMonthlyRecurrencePatternType != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R995");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R995
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
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
            }

            WeeklyRecurrencePatternType weeklyRecurrencePatternType = recurrencePatternBaseType as WeeklyRecurrencePatternType;
            if (weeklyRecurrencePatternType != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1307");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1307
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
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
            }

            DailyRecurrencePatternType dailyRecurrencePatternType = recurrencePatternBaseType as DailyRecurrencePatternType;
            if (dailyRecurrencePatternType != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1109");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1109
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
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
            }
        }

        /// <summary>
        /// Verify RecurrenceRangeBaseType structure.
        /// </summary>
        /// <param name="recurrenceRangeBaseType">Specifies the recurrence pattern with numbered recurrence, non-ending recurrence patterns, and recurrence patterns with a set start and end date.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyRecurrenceRangeTypes(RecurrenceRangeBaseType recurrenceRangeBaseType, bool isSchemaValidated)
        {
            NoEndRecurrenceRangeType recurrenceRangeWithoutEnd = recurrenceRangeBaseType as NoEndRecurrenceRangeType;
            if (recurrenceRangeWithoutEnd != null)
            {
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
            else
            {
                EndDateRecurrenceRangeType endDateRecurrenceRangeType = recurrenceRangeBaseType as EndDateRecurrenceRangeType;
                if (endDateRecurrenceRangeType != null)
                {
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
                else
                {
                    NumberedRecurrenceRangeType numberedRecurrenceRangeType = recurrenceRangeBaseType as NumberedRecurrenceRangeType;
                    if (numberedRecurrenceRangeType != null)
                    {
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
                }
            }
        }

        /// <summary>
        /// Verify DayOfWeekType structure.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyDayOfWeekType(bool isSchemaValidated)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R33");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R33
            this.Site.CaptureRequirementIfIsTrue(
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

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R52");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R52
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                52,
                @"[In t:DaysOfWeekType Simple Type] The type [DaysOfWeekType] is defined as follow:<xs:simpleType name=""DaysOfWeekType"">
    <xs:list>
        <xs:itemType name=""t:DayOfWeekType""/>
    </xs:list>
</xs:simpleType>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R55");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R55
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                55,
                @"[In t:DaysOfWeekType Simple Type] The syntax [DaysOfWeekType] is defined as follow:
<xs:simpleType name=""DaysOfWeekType"">
    <xs:list itemType=""t:DayOfWeekType""/>
</xs:simpleType>");
        }

        /// <summary>
        /// Verify DayOfWeekIndexType structure.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyDayOfWeekIndexType(bool isSchemaValidated)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R25");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R25
            this.Site.CaptureRequirementIfIsTrue(
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
        /// Verify MonthNamesType structure.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyMonthNamesType(bool isSchemaValidated)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R174");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R174
            this.Site.CaptureRequirementIfIsTrue(
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
        }
        #endregion
    }
}