namespace Microsoft.Protocols.TestSuites.MS_WSSREST
{
    using System;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter requirements capture code for MS-WSSREST server role. 
    /// </summary>
    public partial class MS_WSSRESTAdapter
    {
        /// <summary>
        /// Capture underlying transport protocol related requirements.
        /// </summary>
        private void CaptureTransportRelatedRequirements()
        {
            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            switch (transport)
            {
                case TransportProtocol.HTTP:

                    // As response successfully returned, the transport related requirements can be captured.
                    Site.CaptureRequirement(
                        110,
                        @"[In Transport] It [MS-WSSREST] transmits these messages using the HTTP protocol as described in [RFC2616].");
                    break;

                case TransportProtocol.HTTPS:

                    if (Common.IsRequirementEnabled(113, this.Site))
                    {
                        // Having received the response successfully have proved the HTTPS 
                        // transport is supported. If the HTTPS transport is not supported, the 
                        // response can't be received successfully.
                        Site.CaptureRequirement(
                        113,
                        @"[In Appendix C: Product Behavior] Implement does support transmitting these messages using the HTTPS protocol as described in [RFC2818]. (Microsoft SharePoint Foundation 2010 and above products follow this behavior.)");
                    }

                    break;

                default:
                    Site.Debug.Fail("Unknown transport type " + transport);
                    break;
            }
        }

        /// <summary>
        /// To catch Adapter requirements of RetrieveCSDLDocument operation.
        /// </summary>
        /// <param name="csdlDocument">the conceptual schema definition language (CSDL) document</param>
        private void ValidateRetrieveCSDLDocument(XmlDocument csdlDocument)
        {
            this.CaptureTransportRelatedRequirements();
            this.ValidateAndCaptureSchemaValidation();

            // If the ID property contained in the CSDL document which retrieved by Key, the requirement: MS-WSSREST_R63 can be verified.
            bool isVerifyR63 = false;
            XmlNodeList keys = csdlDocument.GetElementsByTagName("Key");

            foreach (XmlNode node in keys)
            {
                if (this.IsContainsIdProperty(node.ParentNode) && node.ChildNodes != null && node.ChildNodes.Count == 1 && !node.ChildNodes[0].Attributes["Name"].Value.Equals("ID", System.StringComparison.OrdinalIgnoreCase))
                {
                    isVerifyR63 = false;
                    break;
                }
                else
                {
                    isVerifyR63 = true;
                }
            }

            Site.Log.Add(LogEntryKind.Debug, "If the ID property contained in the CSDL document which retrieved by Key, the requirement: MS-WSSREST_R63 can be verified.");
            Site.CaptureRequirementIfIsTrue(
                isVerifyR63,
                63,
                @"[In Abstract Data Model] The Site and list data structure: ID Field maps to the Entity Data Model term: EntityKey.");

            // If the Attachments EntitySet exist in metadata, the requirements: MS-WSSREST_R75 and MS-WSSREST_R106 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the Attachments EntitySet exist in metadata, the requirements: MS-WSSREST_R75 and MS-WSSREST_R106 can be verified.");
            bool isAttachmentsExist = this.CheckEntitySet("attachments", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isAttachmentsExist,
                75,
                @"[In Attachment] To facilitate delete operation support for list item attachments, an additional EntitySet is created.");

            Site.CaptureRequirementIfIsTrue(
                isAttachmentsExist,
                106,
                @"[In Attachment] To facilitate create operation support for list item attachments, an additional EntitySet is created.");

            // If the retrieved CSDL document contains the properties "Owshiddenversion"and "Path", the requirement: MS-WSSREST_R14 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the retrieved CSDL document contains the properties 'Owshiddenversion' and 'Path', the requirement: MS-WSSREST_R14 can be verified.");
            bool isVerifyR14 = this.CheckProperty(csdlDocument, "Owshiddenversion") && this.CheckProperty(csdlDocument, "Path");
            Site.CaptureRequirementIfIsTrue(
                isVerifyR14,
                14,
                @"	[In List] The following table specifies minimum set of hidden fields [Owshiddenversion and Path], as specified in [MS-WSSTS] section 2.4.2, which MUST be provided by server.");

            // If the typeMapping contains the field type "AllDayEvent" and the property type is 'Boolean', the requirement: MS-WSSREST_R23 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'AllDayEvent', the requirement: MS-WSSREST_R23 can be verified.");
            bool isVerifyR23 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("AllDayEventFieldName", this.Site), "AllDayEvent") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("AllDayEventFieldName", this.Site), "Boolean", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR23,
                23,
                @"[In List Item] The Field type: AllDayEvent maps to the Entity Data Model property type: Primitive (Boolean).");

            // If the typeMapping contains the field type "Attachments" and the property type is 'Navigation' , the requirements: MS-WSSREST_R24 and MS-WSSREST_R62 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'Attachments', the requirement: MS-WSSREST_R24 and MS-WSSREST_R62 can be verified.");
            bool isAttachmentsVerified = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("AttachmentsFieldName", this.Site), "Attachments") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("AttachmentsFieldName", this.Site), "Navigation", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isAttachmentsVerified,
                24,
                @"[In List Item] The Field type: Attachments maps to the Entity Data Model property type: Navigation.");

            Site.CaptureRequirementIfIsTrue(
                isAttachmentsVerified,
                62,
                @"[In Abstract Data Model] The Site and list data structure: Attachments Field maps to the Entity Data Model term: NavigationProperty.");

            // If the typeMapping contains the field type "Boolean" and the property type is "Boolean", the requirement: MS-WSSREST_R25 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'Boolean', the requirement: MS-WSSREST_R25 can be verified.");
            bool isVerifyR25 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("BooleanFieldName", this.Site), "Boolean") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("BooleanFieldName", this.Site), "Boolean", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR25,
                25,
                @"[In List Item] The Field type: Boolean maps to the Entity Data Model property type: Primitive (Boolean).");

            // If the typeMapping contains the field type "Choice" and the property type is "Navigation", the requirement: MS-WSSREST_R26, MS-WSSREST_R61 and MS-WSSREST_R91 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'Choice', the requirement: MS-WSSREST_R25 can be verified.");
            bool isChoiceVerified = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("ChoiceFieldName", this.Site), "Choice") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("ChoiceFieldName", this.Site), "Navigation", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isChoiceVerified,
                26,
                @"[In List Item] The Field type: Choice maps to the Entity Data Model property type: Navigation.");

            Site.CaptureRequirementIfIsTrue(
                isChoiceVerified,
                61,
                @"[In Abstract Data Model] The Site and list data structure: Choice Field maps to the Entity Data Model term: NavigationProperty.");

            Site.CaptureRequirementIfIsTrue(
                isChoiceVerified,
                91,
                @"[In Choice or Multi-Choice Field] Choice field is mapped as navigation properties between the EntityType corresponding to the list item and the EntityType mentioned earlier [in section 2.2.2.2].");

            // If the typeMapping contains the field type "ContentTypeId" and the property type is "String", the requirement: MS-WSSREST_R27 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'ContentTypeId', the requirement: MS-WSSREST_R27 can be verified.");
            bool isVerifyR27 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("ContentTypeIdFieldName", this.Site), "ContentTypeId") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("ContentTypeIdFieldName", this.Site), "String", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR27,
                27,
                @"[In List Item] The Field type: ContentTypeId maps to the Entity Data Model property type: Primitive (String).");

            // If the typeMapping contains the field type "Counter" and the property type is "Int32", the requirement: MS-WSSREST_R28 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'Counter', the requirement: MS-WSSREST_R28 can be verified.");
            bool isVerifyR28 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("CounterFieldName", this.Site), "Counter") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("CounterFieldName", this.Site), "Int32", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR28,
                28,
                @"[In List Item] The Field type: Counter maps to the Entity Data Model property type: Primitive (Int32).");

            // If the typeMapping contains the field type "CrossProjectLink" and the property type is "Boolean", the requirement: MS-WSSREST_R112 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'CrossProjectLink', the requirement: MS-WSSREST_R112 can be verified.");

            if (Common.IsRequirementEnabled(112, this.Site))
            {
                // Verify MS-WSSREST requirement: MS-WSSREST_R112
                bool isVerifyR112 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("CrossProjectLinkFieldName", this.Site), "CrossProjectLink") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("CrossProjectLinkFieldName", this.Site), "Boolean", csdlDocument);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR112,
                    112,
                    @"[In Appendix C: Product Behavior] Implement does support the Field type: CrossProjectLink mapping to the Entity Data Model property type: Primitive (Boolean).(Microsoft SharePoint Foundation 2010 and Microsoft SharePointServer2010 follow this behavior )");
            }

            // If the typeMapping contains the field type "Currency" and the property type is "Double" , the requirement: MS-WSSREST_R30 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'Currency', the requirement: MS-WSSREST_R30 can be verified.");
            bool isVerifyR30 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("CurrencyFieldName", this.Site), "Currency") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("CurrencyFieldName", this.Site), "Double", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR30,
                30,
                @"[In List Item] The Field type: Currency maps to the Entity Data Model property type: Primitive (Double).");

            // If the typeMapping contains the field type "DateTime" and the property type is "DateTime", the requirement: MS-WSSREST_R31 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'DateTime', the requirement: MS-WSSREST_R31 can be verified.");
            bool isVerifyR31 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("DateTimeFieldName", this.Site), "DateTime") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("DateTimeFieldName", this.Site), "DateTime", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR31,
                31,
                @"[In List Item] The Field type: DateTime maps to the Entity Data Model property type: Primitive (DateTime).");

            // If the typeMapping contains the field type "File" and the property type is "String", the requirement: MS-WSSREST_R32 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'File', the requirement: MS-WSSREST_R31 can be verified.");
            bool isVerifyR32 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("FileFieldName", this.Site), "File") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("FileFieldName", this.Site), "String", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR32,
                32,
                @"[In List Item] The Field type: File maps to the Entity Data Model property type: Primitive (String).");

            // If the typeMapping contains the field type "GridChoice" and the property type is "String", the requirement: MS-WSSREST_R33 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'GridChoice', the requirement: MS-WSSREST_R33 can be verified.");
            bool isVerifyR33 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("GridChoiceFieldName", this.Site), "GridChoice") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("GridChoiceFieldName", this.Site), "String", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR33,
                33,
                @"[In List Item] The Field type: GridChoice maps to the Entity Data Model property type: Primitive (String).");

            // If the typeMapping contains the field type "Integer" and the property type is "Int32", the requirement: MS-WSSREST_R34 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'Integer', the requirement: MS-WSSREST_R34 can be verified.");
            bool isVerifyR34 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("IntegerFieldName", this.Site), "Integer") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("IntegerFieldName", this.Site), "Int32", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR34,
                34,
                @"[In List Item] The Field type: Integer maps to the Entity Data Model property type: Primitive (Int32).");

            // If the typeMapping contains the field type "Lookup" and the property type is "Navigation", the requirement: MS-WSSREST_R35 and MS-WSSREST_R60 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'Lookup', the requirement: MS-WSSREST_R34 can be verified.");
            bool isLookupVerified = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("LookupFieldName", this.Site), "Lookup") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("LookupFieldName", this.Site), "Navigation", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isLookupVerified,
                35,
                @"[In List Item] The Field type: Lookup maps to the Entity Data Model property type: Navigation.");

            Site.CaptureRequirementIfIsTrue(
                isLookupVerified,
                60,
                @"[In Abstract Data Model] The Site and list data structure: Lookup field maps to the Entity Data Model term: NavigationProperty.");

            // If the typeMapping contains the field type "ModStat" and the property type is "String", the requirement: MS-WSSREST_R37 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'ModStat', the requirement: MS-WSSREST_R37 can be verified.");
            bool isVerifyR37 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("ModStatFieldName", this.Site), "ModStat") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("ModStatFieldName", this.Site), "String", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR37,
                37,
                @"[In List Item] The Field type: ModStat maps to the Entity Data Model property type: Primitive (String).");

            // If the typeMapping contains the field type "MultiChoice" and the property type is "Navigation", the requirement: MS-WSSREST_R38 and MS-WSSREST_R92 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'MultiChoice', the requirement: MS-WSSREST_R38 and MS-WSSREST_R92 can be verified.");
            bool isMultiChoiceVerified = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("MultiChoiceFieldName", this.Site), "MultiChoice") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("MultiChoiceFieldName", this.Site), "Navigation", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isMultiChoiceVerified,
                38,
                @"[In List Item] The Field type: MultiChoice maps to the Entity Data Model property type: Navigation.");

            Site.CaptureRequirementIfIsTrue(
                isMultiChoiceVerified,
                92,
                @"[In Choice or Multi-Choice Field] Multi-choice fields (2) is mapped as navigation properties between the EntityType corresponding to the list item and the EntityType mentioned earlier [in section 2.2.2.2].");

            // If the typeMapping contains the field type "Note" and the property type is "String", the requirement: MS-WSSREST_R39 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'Note', the requirement: MS-WSSREST_R39 can be verified.");
            bool isVerifyR39 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("NoteFieldName", this.Site), "Note") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("NoteFieldName", this.Site), "String", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR39,
                39,
                @"[In List Item] The Field type: Note maps to the Entity Data Model property type: Primitive (String).");

            // If the typeMapping contains the field type "Number" and the property type is "Double", the requirement: MS-WSSREST_R40 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'Number', the requirement: MS-WSSREST_R40 can be verified.");
            bool isVerifyR40 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("NumberFieldName", this.Site), "Number") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("NumberFieldName", this.Site), "Double", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR40,
                40,
                @"[In List Item] The Field type: Number maps to the Entity Data Model property type: Primitive (Double).");

            // If the typeMapping contains the field type "PageSeparator" and the property type is "String", the requirement: MS-WSSREST_R41 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'PageSeparator', the requirement: MS-WSSREST_R41 can be verified.");
            bool isVerifyR41 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("PageSeparatorFieldName", this.Site), "PageSeparator") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("PageSeparatorFieldName", this.Site), "String", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR41,
                41,
                @"[In List Item] The Field type: PageSeparator maps to the Entity Data Model property type: Primitive (String).");

            // If the typeMapping contains the field type "Recurrence" and the property type is "Boolean", the requirement: MS-WSSREST_R42 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'Recurrence', the requirement: MS-WSSREST_R42 can be verified.");
            bool isVerifyR42 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("RecurrenceFieldName", this.Site), "Recurrence") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("RecurrenceFieldName", this.Site), "Boolean", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR42,
                42,
                @"[In List Item] The Field type: Recurrence maps to the Entity Data Model property type: Primitive (Boolean).");

            // If the typeMapping contains the field type "Text" and the property type is "String", the requirement: MS-WSSREST_R43 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'Text', the requirement: MS-WSSREST_R43 can be verified.");
            bool isVerifyR43 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("TextFieldName", this.Site), "Text") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("TextFieldName", this.Site), "String", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR43,
                43,
                @"[In List Item] The Field type: Text maps to the Entity Data Model property type: Primitive (String).");

            // If the typeMapping contains the field type "URL" and the property type is "String", the requirement: MS-WSSREST_R45 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'URL', the requirement: MS-WSSREST_R45 can be verified.");
            bool isVerifyR45 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("URLFieldName", this.Site), "URL") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("URLFieldName", this.Site), "String", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR45,
                45,
                @"[In List Item] The Field type: URL maps to the  Entity Data Model property type: Primitive (String).");

            // If the typeMapping contains the field type "User" and the property type is "Navigation", the requirement: MS-WSSREST_R46 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'User', the requirement: MS-WSSREST_R46 can be verified.");
            bool isVerifyR46 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("UserFieldName", this.Site), "User") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("UserFieldName", this.Site), "Navigation", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR46,
                46,
                @"[In List Item] The Field type: User maps to the Entity Data Model property type: Navigation.");

            // If the typeMapping contains the field type "WorkFlowEventType" and the property type is "String", the requirement: MS-WSSREST_R47 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'WorkFlowEventType', the requirement: MS-WSSREST_R47 can be verified.");
            bool isVerifyR47 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("WorkFlowEventTypeFieldName", this.Site), "WorkFlowEventType") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("WorkFlowEventTypeFieldName", this.Site), "String", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR47,
                47,
                @"[In List Item] The Field type: WorkFlowEventType maps to the Entity Data Model property type: Primitive (String).");

            // If the typeMapping contains the field type "WorkFlowStatus" and the property type is "Int32", the requirement: MS-WSSREST_R48 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the typeMapping contains the field type 'WorkFlowStatus', the requirement: MS-WSSREST_R48 can be verified.");
            bool isVerifyR48 = this.sutAdapter.CheckFieldType(Common.GetConfigurationPropertyValue("WorkFlowStatusFieldName", this.Site), "WorkFlowStatus") && this.CheckPropertyType(Common.GetConfigurationPropertyValue("WorkFlowStatusFieldName", this.Site), "Int32", csdlDocument);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR48,
                48,
                @"[In List Item] The Field type: WorkFlowStatus maps to the Entity Data Model property type: Primitive (Int32).");
        }

        /// <summary>
        /// Validate the requirement by schema validation.
        /// </summary>
        private void ValidateAndCaptureSchemaValidation()
        {
            // If the Schema validation is success, below requirements can be verified.
            // Verify MS-WSSREST requirement: MS-WSSREST_R8
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                SchemaValidation.ValidationResult,
                ValidationResult.Success,
                8,
                @"[In Elements] The List element is a container within a site (2) that stores list items.");

            // Verify MS-WSSREST requirement: MS-WSSREST_R9
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                SchemaValidation.ValidationResult,
                ValidationResult.Success,
                9,
                @"[In Elements] The List item element is an individual entry within a list (1).");

            // Verify MS-WSSREST requirement: MS-WSSREST_R10
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                SchemaValidation.ValidationResult,
                ValidationResult.Success,
                10,
                @"[In List] Each list (1) is represented as an EntitySet as specified in [MC-CSDL] section 2.1.17, which[EntitySet] contains Entities of a single EntityType as specified in [MC-CSDL] section 2.1.2.");

            // Verify MS-WSSREST requirement: MS-WSSREST_R11
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                SchemaValidation.ValidationResult,
                ValidationResult.Success,
                11,
                @"[In List] This EntityType [for list] contains properties for every non-hidden field (2) in the list (1) whose field type is supported as well as a subset of hidden fields (2).");

            // Verify MS-WSSREST requirement: MS-WSSREST_R19
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                SchemaValidation.ValidationResult,
                ValidationResult.Success,
                19,
                @"[In List Item] Every list item is represented as an Entity of a particular EntityType as specified in [MC-CSDL] section 2.1.2.");

            // Verify MS-WSSREST requirement: MS-WSSREST_R20
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                SchemaValidation.ValidationResult,
                ValidationResult.Success,
                20,
                @"[In List Item] EntityTypes are created based on the list (1) to which the list item belongs.");

            // Verify MS-WSSREST requirement: MS-WSSREST_R59
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                SchemaValidation.ValidationResult,
                ValidationResult.Success,
                59,
                @"[In Abstract Data Model] The Site and list data structure: List maps to the Entity Data Model term: EntitySet.");

            // Verify MS-WSSREST requirement: MS-WSSREST_R64
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                SchemaValidation.ValidationResult,
                ValidationResult.Success,
                64,
                @"[In Message Processing Events and Sequencing Rules] This protocol [MS-WSSREST] provides operations [create, retrieve, update, delete; operations on choice and multi-choice fields; inserting a new document into a document library] support as described in the following table.");

            // Verify MS-WSSREST requirement: MS-WSSREST_R77
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                SchemaValidation.ValidationResult,
                ValidationResult.Success,
                77,
                @"[In Attachment] Attachments are represented as a multi-valued navigation property on an Entity.");

            // Verify MS-WSSREST requirement: MS-WSSREST_R78
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                SchemaValidation.ValidationResult,
                ValidationResult.Success,
                78,
                @"[In Attachment] The EntityType for this EntitySet [support for list item attachments] contains three properties, ""EntitySet"", ""ItemId"" and ""Name"".");

            // Verify MS-WSSREST requirement: MS-WSSREST_R79
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                SchemaValidation.ValidationResult,
                ValidationResult.Success,
                79,
                @"[In Attachment] All the three properties [""EntitySet"", ""ItemId"" and ""Name""] together serve as its [EntitySet's] EntityKey as specified in [MC-CSDL] section 2.1.5.");
        }

        /// <summary>
        /// Check whether the special xml node contains an id property.
        /// </summary>
        /// <param name="node">The special xml node.</param>
        /// <returns>True if the specified xml node contains an id property, otherwise false.</returns>
        private bool IsContainsIdProperty(XmlNode node)
        {
            bool result = false;

            if (node != null && node.ChildNodes != null)
            {
                foreach (XmlNode item in node.ChildNodes)
                {
                    if (item.Name.Equals("Property", System.StringComparison.OrdinalIgnoreCase) && item.Attributes["Name"].Value.Equals("ID", System.StringComparison.OrdinalIgnoreCase))
                    {
                        result = true;
                        break;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Check whether the specified property contained in the response.
        /// </summary>
        /// <param name="doc">The response xml.</param>
        /// <param name="propertyName">The specified property name.</param>
        /// <returns>True if the specified property contained in the response xml, otherwise false.</returns>
        private bool CheckProperty(XmlDocument doc, string propertyName)
        {
            bool result = false;
            XmlNodeList entryTypes = doc.GetElementsByTagName("EntityType");

            if (null != entryTypes && entryTypes.Count > 0)
            {
                foreach (XmlNode node in entryTypes)
                {
                    string listName = node.Attributes["Name"].Value;
                    if (listName.Contains(Common.GetConfigurationPropertyValue("CalendarListName", this.Site))
                         || listName.Contains(Common.GetConfigurationPropertyValue("DiscussionBoardListName", this.Site))
                         || listName.Contains(Common.GetConfigurationPropertyValue("DoucmentLibraryListName", this.Site))
                         || listName.Contains(Common.GetConfigurationPropertyValue("GeneralListName", this.Site))
                         || listName.Contains(Common.GetConfigurationPropertyValue("SurveyListName", this.Site))
                         || listName.Contains(Common.GetConfigurationPropertyValue("TaskListName", this.Site))
                         || listName.Contains(Common.GetConfigurationPropertyValue("WorkflowHistoryListName", this.Site)))
                    {
                        foreach (XmlNode xn in node.ChildNodes)
                        {
                            if (xn.Name == "Property")
                            {
                                string fieldName = xn.Attributes["Name"].Value;

                                if (fieldName.Equals(propertyName, System.StringComparison.OrdinalIgnoreCase))
                                {
                                    result = true;
                                    break;
                                }
                            }
                        }

                        if (!result)
                        {
                            break;
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Check whether the type of the specified property equals to the specified property type.
        /// </summary>
        /// <param name="fieldName">The specified field name.</param>
        /// <param name="propertyType">The specified property type.</param>
        /// <param name="csdlMetadata">The conceptual schema definition language document of a data service.</param>
        /// <returns>True if the type of the specified property equals to the specified property type, otherwise false.</returns>
        private bool CheckPropertyType(string fieldName, string propertyType, XmlDocument csdlMetadata)
        {
            bool result = false;
            XmlNodeList xnl = csdlMetadata.GetElementsByTagName("EntityType");

            foreach (XmlNode node in xnl)
            {
                string listName = node.Attributes["Name"].Value;
                if (listName.Contains(Common.GetConfigurationPropertyValue("CalendarListName", this.Site))
                     || listName.Contains(Common.GetConfigurationPropertyValue("DiscussionBoardListName", this.Site))
                     || listName.Contains(Common.GetConfigurationPropertyValue("DoucmentLibraryListName", this.Site))
                     || listName.Contains(Common.GetConfigurationPropertyValue("GeneralListName", this.Site))
                     || listName.Contains(Common.GetConfigurationPropertyValue("SurveyListName", this.Site))
                     || listName.Contains(Common.GetConfigurationPropertyValue("TaskListName", this.Site))
                     || listName.Contains(Common.GetConfigurationPropertyValue("WorkflowHistoryListName", this.Site)))
                {
                    foreach (XmlNode xn in node.ChildNodes)
                    {
                        if (xn.Name == "Property" || xn.Name == "NavigationProperty")
                        {
                            string tempFieldName = xn.Attributes["Name"].Value;

                            if (tempFieldName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                            {
                                if (propertyType == "Navigation" && xn.Name == "NavigationProperty")
                                {
                                    result = true;
                                    break;
                                }

                                if (propertyType != "Navigation" && xn.Attributes["Type"].Value.Contains(propertyType))
                                {
                                    result = true;
                                    break;
                                }
                            }
                        }
                    }

                    if (result == true)
                    {
                        break;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Check whether the specified EntitySet exist in the metadata.
        /// </summary>
        /// <param name="entitySetName">The specified EntitySet name.</param>
        /// <param name="csdlMetadata">The conceptual schema definition language document of a data service.</param>
        /// <returns>True if the specified EntitySet exist in the metadata, otherwise false.</returns>
        private bool CheckEntitySet(string entitySetName, XmlDocument csdlMetadata)
        {
            bool result = false;
            XmlNodeList xnl = csdlMetadata.GetElementsByTagName("EntitySet");

            if (xnl != null && xnl.Count > 0)
            {
                foreach (XmlNode node in xnl)
                {
                    if (node.Attributes["Name"] != null && node.Attributes["Name"].Value.Equals(entitySetName, StringComparison.OrdinalIgnoreCase))
                    {
                        result = true;
                        break;
                    }
                }
            }

            return result;
        }
    }
}