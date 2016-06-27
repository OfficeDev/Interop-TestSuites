namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using System.Collections.Generic;
    /// <summary>
    /// MS_OXCMSGAdapter partial class
    /// </summary>
    public partial class MS_OXCMSGAdapter
    {
        /// <summary>
        /// The name of Open Specification MS-OXCDATA.
        /// </summary>
        private const string CdataShortName = "MS-OXCDATA";

        /// <summary>
        /// Verify MAPIHTTP transport.
        /// </summary>
        private void VerifyMAPITransport()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http" && Common.IsRequirementEnabled(10000, this.Site))
            {
                // If the transport sequence is MAPIHTTP and the code can reach here, it means that the implementation does support MAPIHTTP transport.
                this.Site.CaptureRequirement(
                    10000,
                    @"[In Appendix A: Product Behavior] Implementation does support this specification [MS-OXCMAPIHTTP]. (Exchange Server 2013 Service Pack 1 (SP1) follows this behavior.)");
            }
        }

        /// <summary>
        /// Verify properties PidTagAccess, PidTagAccessLevel, PidTagChangeKey, PidTagCreationTime, PidTagLastModificationTime, PidTagLastModifierName, PidTagObjectType, PidTagRecordKey and PidTagSearchKey exist on all Message objects.
        /// </summary>
        /// <param name="pidTagAccess">PidTagAccess property</param>
        /// <param name="pidTagAccessLevel">PidTagAccessLevel property</param>
        /// <param name="pidTagChangeKey">PidTagChangeKey property</param>
        /// <param name="pidTagCreationTime">PidTagCreationTime property</param>
        /// <param name="pidTagLastModificationTime">PidTagLastModificationTime property</param>
        /// <param name="pidTagLastModifierName">PidTagLastModifierName property</param>
        /// <param name="pidTagSearchKey">PidTagSearchKey property</param>
        private void VerifyPropertiesExistOnAllMessageObject(PropertyObj pidTagAccess, PropertyObj pidTagAccessLevel, PropertyObj pidTagChangeKey, PropertyObj pidTagCreationTime, PropertyObj pidTagLastModificationTime, PropertyObj pidTagLastModifierName, PropertyObj pidTagSearchKey)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1097");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1097
            // Property PidTagAccess is not null means it exists on the message object
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagAccess,
                1097,
                @"[In General Properties] The Property PidTagAccess exists on the Message objects.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1098");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1098
            // Property PidTagAccessLevel is not null means it exists on the message object
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagAccessLevel,
                1098,
                @"[In General Properties] The Property PidTagAccessLevel exists on the Message objects.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1099");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1099
            // Property PidTagChangeKey is not null means it exists on the message object
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagChangeKey,
                1099,
                @"[In General Properties] The Property PidTagChangeKey exists on the Message objects.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1100");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1100
            // Property PidTagCreationTime is not null means it exists on the message object
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagCreationTime,
                1100,
                @"[In General Properties] The Property PidTagCreationTime exists on the Message objects.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1101");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1101
            // Property PidTagLastModificationTime is not null means it exists on the message object
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagLastModificationTime,
                1101,
                @"[In General Properties] The Property PidTagLastModificationTime exists on the Message objects.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1102");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1102
            // Property PidTagLastModifierName is not null means it exists on the message object
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagLastModifierName,
                1102,
                @"[In General Properties] The Property PidTagLastModifierName exists on the Message objects.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1105");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1105
            // Property PidTagSearchKey is not null means it exists on the message object
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagSearchKey,
                1105,
                @"[In General Properties] The Property PidTagSearchKey exists on the Message objects.");
        }

        /// <summary>
        /// Verify requirements related with PidTagObjectType and PidTagRecordKey
        /// </summary>
        /// <param name="pidTagObjectType"></param>
        /// <param name="pidTagRecordKey"></param>
        private void VerifyPidTagObjectTypeAndPidTagRecordKey(PropertyObj pidTagObjectType, PropertyObj pidTagRecordKey)
        {
            if (Common.IsRequirementEnabled(3002, this.Site))
            {
                if(PropertyHelper.IsPropertyValid(pidTagObjectType))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3002");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R3002
                    this.Site.CaptureRequirement(
                        3002,
                        @"[In Appendix A: Product Behavior] Implementation does support the PidTagObjectType property. (Exchange 2007 follows this behavior.)");
                } 
            }

            if (Common.IsRequirementEnabled(3004, this.Site))
            {
                if (PropertyHelper.IsPropertyValid(pidTagRecordKey))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3004");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R3004
                    this.Site.CaptureRequirement(
                        3004,
                        @"[In Appendix A: Product Behavior] Implementation does support the PidTagRecordKey property. (Exchange 2007 follows this behavior.)");
                }
            }
        }
        /// <summary>
        /// Verify properties PidTagAccessLevel and PidTagRecordKey exist on any Attachment object.
        /// </summary>
        /// <param name="pidTagAccessLevel">PidTagAccessLevel property</param>
        /// <param name="pidTagRecordKey">PidTagRecordKey property</param>
        private void VerifyPropertiesExistOnAllAttachmentObject(PropertyObj pidTagAccessLevel, PropertyObj pidTagRecordKey)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1825");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1825
            // Property PidTagAccessLevel is not null means it exists on the attachment object
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagAccessLevel,
                1825,
                @"[In General Properties] Property PidTagAccessLevel exists on Attachment object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1827");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1827
            // Property PidTagRecordKey is not null means it exists on the attachment object
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagRecordKey,
                1827,
                @"[In General Properties] Property PidTagRecordKey exists on Attachment object.");
        }

        /// <summary>
        /// Verify the DataType of Messages.
        /// </summary>
        /// <param name="propertyObj"> The structure of message. </param>
        private void VerifyMessageSyntaxDataType(PropertyObj propertyObj)
        {
            switch (propertyObj.PropertyName)
            {
                case PropertyNames.PidTagHasAttachments:
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R497");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R497
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBoolean,
                        propertyObj.ValueType,
                        497,
                        @"[In PidTagHasAttachments Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");

                    // Verify the PtypBoolean type
                    bool isVerifyR2698 = (Convert.ToByte((bool)propertyObj.Value) >= byte.MinValue && Convert.ToByte((bool)propertyObj.Value) <= byte.MaxValue) && (propertyObj.ValueTypeCode == 0x000B);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2698:{0}", propertyObj.Value);

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2698
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR2698,
                        CdataShortName,
                        2698,
                        @"[In Property Data Types] PtypBoolean (PT_BOOLEAN. bool) is that 1 byte, restricted to 1 or 0 [MS-DTYP]: BOOLEAN with Property Type Value 0x000B, %x0B.00.");
                    break;

                case PropertyNames.PidTagMessageClass:
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R499");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R499
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        499,
                        @"[In PidTagMessageClass Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");

                    // Verify the PtypString type
                    bool isVerifyR2700 = propertyObj.ValueTypeCode == 0x001F;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2700:{0}", propertyObj.Value);

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2700
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR2700,
                        CdataShortName,
                        2700,
                        @"[In Property Data Types] PtypString (PT_UNICODE, string) is that Variable size; a string of Unicode characters in UTF-16LE format encoding with terminating null character (0x0000). with Property Type Value  0x001F,%x1F.00.");

                    string strMsg = (string)propertyObj.Value;

                    if (PropertyHelper.IsPropertyValid(propertyObj))
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R22. The value of PidTagMessageClass is {0}.", strMsg);

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R22
                        this.Site.CaptureRequirementIfIsTrue(
                            strMsg.IndexOf(".") > 0,
                            22,
                            @"[In PidTagMessageClass Property] The value of this property [PidTagMessageClass] is interpreted in groups of characters separated by periods (""."").");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R19:{0}", propertyObj.Value);

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R19
                        // Characters in PidTagMessageClass are from the ASCII characters 0x20 to 0x7F.
                        bool isVerifiedR19 = AdapterHelper.IsStringValueValid(strMsg, (char)0x20, (char)0x7F);

                        Site.CaptureRequirementIfIsTrue(
                            isVerifiedR19,
                            19,
                            @"[In PidTagMessageClass Property] All characters in this property [PidTagMessageClass] MUST be from the ASCII characters 0x20 through 0x7F.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R501:{0}", propertyObj.Value);

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R501
                        // ASCII character of a period is 0x2E 
                        bool isVerifyR501 = strMsg[strMsg.Length - 1] != 0x2E;

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR501,
                            501,
                            @"[In PidTagMessageClass Property] It [PtypString] MUST NOT end with a period (ASCII character 0x2E).");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R502:{0}", propertyObj.Value);

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R502
                        bool isVerifyR502 = AdapterHelper.IsStringLengthValid(strMsg, 0, 256);

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR502,
                            502,
                            @"[In PidTagMessageClass Property] Its length [PtypString] MUST be greater than zero and less than 256 characters.");

                        if (Common.IsRequirementEnabled(496, this.Site))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R496:{0}", propertyObj.Value);

                            // Verify MS-OXCMSG requirement: MS-OXCMSG_R496
                            bool isR496Satisfied = AdapterHelper.IsStringLengthValid(strMsg, 0, 128);

                            Site.CaptureRequirementIfIsTrue(
                                isR496Satisfied,
                                496,
                                @"[In Appendix A: Product Behavior] Furthermore, its [PtypString] length is fewer than 128 characters because some operations require extending the value of PidTagMessageClass property. (Exchange 2007 and above follow this behavior.)");
                        }
                    }

                    break;

                case PropertyNames.PidTagMessageFlags:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R507");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R507
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        507,
                        @"[In PidTagMessageFlags Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                    bool isVerifiedR1133 = false;
                    int valueOfPidTagMessageFlags = (int)propertyObj.Value;
                    if (valueOfPidTagMessageFlags == 0)
                    {
                        isVerifiedR1133 = true;
                    }
                    else
                    {
                        int value = 0x00000001 | 0x00000008 | 0x00000080 | 0x00000002 | 0x00000004 | 0x00000010 | 0x00000020 | 0x00000040 | 0x00000100 | 0x00000200 | 0x00000400 | 0x00002000 | 0x00008000;
                        if ((valueOfPidTagMessageFlags | value) == value)
                        {
                            isVerifiedR1133 = true;
                        }
                    }

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1133. The value of PidTagMessageFlags is {0}.", propertyObj.Value);

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1133
                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR1133,
                        1133,
                        @"[In PidTagMessageFlags Property] Set to zero or to a bitwise OR of one or more of the values [0x00000001, 0x00000008, 0x00000080, 0x00000002, 0x00000004, 0x00000010, 0x00000020, 0x00000040, 0x00000100, 0x00000200, 0x00000400, 0x00002000 and 0x00008000] from the following tables [two tables in section 2.2.1.6].");

                    // Verify the PtypInteger32 type
                    bool isVerifyR2691 = ((int)propertyObj.Value >= int.MinValue && (int)propertyObj.Value <= int.MaxValue) && (propertyObj.ValueTypeCode == 0x0003);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2691:{0}", propertyObj.Value);

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2691
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR2691,
                         CdataShortName,
                         2691,
                        @"[In Property Data Types] PtypInteger32 (PT_LONG, PT_I4, int, ui4) is that 4 bytes; a 32-bit integer [MS-DTYP]: INT32 with Property Type Value 0x0003,%x03.00.");
                    break;

                case PropertyNames.PidTagMessageCodepage:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R503");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R503
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        503,
                        @"[In PidTagMessageCodepage Property] Type: PtypInteger32, unsigned.");
                    break;

                case PropertyNames.PidTagMessageLocaleId:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R505");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R505
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        505,
                        @"[In PidTagMessageLocaleId Property] Type: PtypInteger32, unsigned.");
                    break;

                case PropertyNames.PidTagMessageSize:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R539");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R539
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        539,
                        @"[In PidTagMessageSize Property] Type: PtypInteger32, unsigned.");
                    break;

                case PropertyNames.PidTagMessageStatus:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R118");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R118
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        118,
                        @"[In PidTagMessageStatus Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                    bool isVerifyR120 = false;

                    int valueOfPidTagMessageStatus = (int)propertyObj.Value;

                    int possibleValues = 0x00000000 | 0x00001000 | 0x00000800 | 0x00002000;
                    if ((valueOfPidTagMessageStatus | possibleValues) == possibleValues)
                    {
                        isVerifyR120 = true;
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R120");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R120
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR120,
                        120,
                        @"[In PidTagMessageStatus Property] Contains a bitwise OR of zero or more of the following values [0x00001000, 0x00000800 and 0x00002000].");
                    break;

                case PropertyNames.PidTagSubjectPrefix:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R57");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R57
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        57,
                        @"[In PidTagSubjectPrefix Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagNormalizedSubject:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R63");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R63
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        63,
                        @"[In PidTagNormalizedSubject Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagImportance:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R69");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R69
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        69,
                        @"[In PidTagImportance Property]Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                    bool isVerifyR71 = false;

                    int propertyObject = (int)propertyObj.Value;
                    if (propertyObject == 0x00000000 || propertyObject == 0x00000001 || propertyObject == 0x00000002)
                    {
                        isVerifyR71 = true;
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R71");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R71
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR71,
                        71,
                        @"[In PidTagImportance Property] This property [PidTagImportance] MUST be set to one of the following values [0x00000000, 0x00000001 and 0x00000002].");
                    break;

                case PropertyNames.PidTagPriority:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R75");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R75
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        75,
                        @"[In PidTagPriority Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                    bool isVerifyR77 = false;

                    int valueOfPidTagPriority = (int)propertyObj.Value;
                    if (valueOfPidTagPriority == 0x00000001 || valueOfPidTagPriority == 0x00000000 || valueOfPidTagPriority == unchecked((int)0xffffffff))
                    {
                        isVerifyR77 = true;
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R77");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R77
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR77,
                        77,
                        @"[In PidTagPriority Property] This property [PidTagPriority] is set to one of the following values [0x00000001, 0x00000000, 0xFFFFFFFF].");
                    break;

                case PropertyNames.PidTagSensitivity:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R81");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R81
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        81,
                        @"[In PidTagSensitivity Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                    bool isVerifyR83 = false;

                    int valueOfPidTagSensitivity = (int)propertyObj.Value;
                    if (valueOfPidTagSensitivity == 0x00000000 || valueOfPidTagSensitivity == 0x00000001 || valueOfPidTagSensitivity == 0x00000002 || valueOfPidTagSensitivity == 0x00000003)
                    {
                        isVerifyR83 = true;
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R83");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R83
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR83,
                        83,
                        @"[In PidTagSensitivity Property] The value of this property [PidTagSensitivity] is one of the following [0x00000000, 0x00000001, 0x00000002 and 0x00000003].");
                    break;

                case PropertyNames.PidTagBody:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R123");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R123
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        2052,
                        @"[In PidTagBody Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagNativeBody:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R125");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R125
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        2055,
                        @"[In PidTagNativeBody Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                    if (Common.IsRequirementEnabled(1714, this.Site))
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R127");

                        bool isVerifyR127 = false;
                        int valueOfPidTagNativeBody = (int)propertyObj.Value;
                        if (valueOfPidTagNativeBody == 0x00 || valueOfPidTagNativeBody == 0x01 || valueOfPidTagNativeBody == 0x02 || valueOfPidTagNativeBody == 0x03 || valueOfPidTagNativeBody == 0x04)
                        {
                            isVerifyR127 = true;
                        }

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R127
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR127,
                            2057,
                            @"[In PidTagNativeBody Property] The value of this property [PidTagNativeBody] is limited to one of the property values [0x00, 0x01, 0x02, 0x03, 0x04] shown in the following table [table in section 2.2.1.49.2].");
                    }

                    break;
                case PropertyNames.PidTagInternetCodepage:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2069");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R2069
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        2069,
                        @"[In PidTagInternetCodepage Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                    if (PropertyHelper.IsPropertyValid(propertyObj))
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2070");

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R2070
                        // MS-OXCMSG_R140 can be indirectly verified when the value of property PidTagInternetCodepage is not null.
                        this.Site.CaptureRequirement(
                            2070,
                            @"[In PidTagInternetCodepage Property] The PidTagInternetCodepage property ([MS-OXPROPS] section 2.735) indicates the code page used for the PidTagBody property (section 2.2.1.56.1) or the PidTagBodyHtml property (section 2.2.1.56.3).");
                    }

                    break;

                case PropertyNames.PidTagArchiveTag:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R157");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R157
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        2109,
                        @"[In PidTagArchiveTag Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagRetentionPeriod:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R163");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R163
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        2116,
                        @"[In PidTagRetentionPeriod Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagStartDateEtc:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R170");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R170
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        2125,
                        @"[In PidTagStartDateEtc Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");

                    if (PropertyHelper.IsPropertyValid(propertyObj))
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R171");

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R171
                        // The length of PidTagStartDateEtc should be 12 bytes when the value of PidTagStartDateEtc exists, because Default retention period is 4 bytes and Start date is 8 bytes.
                        this.Site.CaptureRequirementIfAreEqual<int>(
                            12,
                            ((byte[])propertyObj.Value).Length - 2,
                            2126,
                            @"[In PidTagStartDateEtc Property] The PidTagStartDateEtc property ([MS-OXPROPS] section 2.1014) has the following structure [Default retention period, Start date].");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R172");

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R172
                        // The length of PidTagStartDateEtc should be 12 bytes when the value of PidTagStartDateEtc exists, because Default retention period is 4 bytes and Start date is 8 bytes.
                        this.Site.CaptureRequirementIfAreEqual<int>(
                            12,
                            ((byte[])propertyObj.Value).Length - 2,
                            2127,
                            @"[In PidTagStartDateEtc Property] The length of Default retention period is 4 bytes.");
                    }

                    break;

                case PropertyNames.PidTagRetentionDate:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R179");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R179
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypTime,
                        propertyObj.ValueType,
                        2134,
                        @"[In PidTagRetentionDate Property] Type: PtypTime ([MS-OXCDATA] section 2.11.1).");

                    // Verify the PtypTime type
                    bool isVerifyR2702 = ((System.DateTime)propertyObj.Value >= DateTime.MinValue && (System.DateTime)propertyObj.Value <= DateTime.MaxValue) && (propertyObj.ValueTypeCode == 0x0040);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2702:{0}", propertyObj.Value);

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2702
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR2702,
                        CdataShortName,
                        2702,
                        @"[In Property Data Types] PtypTime (PT_SYSTIME, time, datetime, datetime.tz, datetime.rfc1123, Date, time, time.tz) is that 8 bytes; a 64-bit integer representing the number of 100-nanosecond intervals since January 1, 1601 [MS-DTYP]: FILETIME with Property Type Value 0x0040,%x40.00.");

                    break;

                case PropertyNames.PidTagRetentionFlags:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R194");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R194
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        2147,
                        @"[In PidTagRetentionFlags Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                    bool isVerifyR2153 = false;

                    int valueOfPidTagRetentionFlags = (int)propertyObj.Value;
                    int possibleValueOfPidTagRetentionFlags = 0 | 0x00000001 | 0x00000002 | 0x00000004 | 0x00000008 | 0x000000010 | 0x000000020 | 0x000000040 | 0x00000080 | 0x00000100;
                    if ((valueOfPidTagRetentionFlags & possibleValueOfPidTagRetentionFlags) == valueOfPidTagRetentionFlags)
                    {
                        isVerifyR2153 = true;
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R197");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R197
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR2153,
                        2153,
                        @"[In PidTagRetentionFlags Property] The value of the PidTagRetentionFlags property [0x00000001, 0x00000002, 0x00000004, 0x00000008, 0x000000010, 0x000000020, 0x000000040, 0x00000080, 0x00000100] is a bitwise OR of zero or more of the values from the following table [table in section 2.2.1.51.6].");
                    break;

                case PropertyNames.PidTagArchivePeriod:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R205");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R205
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        2163,
                        @"[In PidTagArchivePeriod Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagLastModificationTime:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R222");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R222
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypTime,
                        propertyObj.ValueType,
                        222,
                        @"[In PidTagLastModificationTime Property] Type: PtypTime, in UTC ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagCreationTime:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R224");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R224
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypTime,
                        propertyObj.ValueType,
                        224,
                        @"[In PidTagCreationTime Property] Type: PtypTime, in UTC ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachSize:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R229");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R229
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        229,
                        @"[In PidTagAttachSize Property] Type: PtypInteger32, unsigned ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachNumber:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R232");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R232
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        232,
                        @"[In PidTagAttachNumber Property] Type: PtypInteger32, unsigned ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachDataBinary:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R235");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R235
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        235,
                        @"[In PidTagAttachDataBinary Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachMethod:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R239");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R239
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        239,
                        @"[In PidTagAttachMethod Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                    bool isVerifyR241 = false;
                    int valueOfPidTagAttachMethod = (int)propertyObj.Value;
                    if (valueOfPidTagAttachMethod == 0x00000000 || valueOfPidTagAttachMethod == 0x00000001 || valueOfPidTagAttachMethod == 0x00000002 || valueOfPidTagAttachMethod == 0x00000004 || valueOfPidTagAttachMethod == 0x00000005 || valueOfPidTagAttachMethod == 0x00000006)
                    {
                        isVerifyR241 = true;
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R241");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R241
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR241,
                        241,
                        @"[In PidTagAttachMethod Property] This property [PidTagAttachMethod] is set to one of the following values [0x00000000, 0x00000001, 0x00000002, 0x00000004, 0x00000005, 0x00000006].");
                    break;

                case PropertyNames.PidTagAttachLongFilename:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R599");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R599
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        599,
                        @"[In PidTagAttachLongFilename Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachFilename:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R601");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R601
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        601,
                        @"[In PidTagAttachFilename Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachExtension:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R603");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R603
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        603,
                        @"[In PidTagAttachExtension Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachLongPathname:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R605");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R605
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        605,
                        @"[In PidTagAttachLongPathname Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachTag:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R609");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R609
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        609,
                        @"[In PidTagAttachTag Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");

                    if (PropertyHelper.IsPropertyValid(propertyObj))
                    {
                        byte[] pidTagAttachTagTextFir = new byte[] { 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x14, 0x03, 0x0A, 0x01 };
                        byte[] pidTagAttachTagTextSec = new byte[] { 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x14, 0x03, 0x0A, 0x03, 0x02, 0x01 };
                        byte[] pidTagAttachTagTextThird = new byte[] { 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x14, 0x03, 0x0A, 0x04 };

                        byte[] valueOfPidTagAttachTag = new byte[((byte[])propertyObj.Value).Length - 2];
                        Buffer.BlockCopy((byte[])propertyObj.Value, 2, valueOfPidTagAttachTag, 0, ((byte[])propertyObj.Value).Length - 2);

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R612. The value of PidTagAttachTag is {0}.", propertyObj.Value.ToString());

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R612
                        bool isVerifiedR612 = propertyObj.Value == null
                            || Common.CompareByteArray(pidTagAttachTagTextFir, valueOfPidTagAttachTag)
                            || Common.CompareByteArray(pidTagAttachTagTextSec, valueOfPidTagAttachTag)
                            || Common.CompareByteArray(pidTagAttachTagTextThird, valueOfPidTagAttachTag);

                        this.Site.CaptureRequirementIfIsTrue(
                            isVerifiedR612,
                            612,
                            @"[In PidTagAttachTag Property] This property can be left unset; if set, it [PidTagAttachTag] MUST be one of the following [TNEF, afStorage, MIME].");
                    }

                    break;

                case PropertyNames.PidTagRenderingPosition:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R613");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R613
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        613,
                        @"[In PidTagRenderingPosition Property] Type: PtypInteger32, unsigned ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachRendering:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R616");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R616
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        616,
                        @"[In PidTagAttachRendering Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");

                    // Verify the PtypBinary type
                    bool isVerifyR2707 = propertyObj.ValueTypeCode == 0x0102;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2707:{0}", propertyObj.Value);

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2707
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR2707,
                        CdataShortName,
                        2707,
                        @"[In Property Data Types] PtypBinary (PT_BINARY) is that variable size; a COUNT field followed by that many bytes with Property Type Value 0x0102,%x02.01.");
                    break;

                case PropertyNames.PidTagAttachFlags:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R618");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R618
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        618,
                        @"[In PidTagAttachFlags Property] Type: PtypInteger32, as a bit field ([MS-OXCDATA] section 2.11.1).");

                    bool isVerifyR620 = false;

                    int valueOfPidTagAttachFlags = (int)propertyObj.Value;
                    int possibleValueOfPidTagAttachFlags = 0 | 0x00000001 | 0x00000002 | 0x00000004;

                    if ((valueOfPidTagAttachFlags | possibleValueOfPidTagAttachFlags) == possibleValueOfPidTagAttachFlags)
                    {
                        isVerifyR620 = true;
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R620");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R620
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR620,
                        620,
                        @"[In PidTagAttachFlags Property] This property [PidTagAttachFlags] contains a bitwise OR of zero or more of the following flags [0x00000001, 0x00000002 and 0x00000004].");
                    break;

                case PropertyNames.PidTagAttachTransportName:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R622");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R622
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        622,
                        @"[In PidTagAttachTransportName Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");

                    if (PropertyHelper.IsPropertyValid(propertyObj))
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R623");

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R623
                        // MS-OXCMSG_R623 can be indirectly verified if the value of PidTagAttachTransportName is not null.
                        this.Site.CaptureRequirementIfIsNotNull(
                            propertyObj.Value,
                            623,
                            @"[In PidTagAttachTransportName Property] The PidTagAttachTransportName property ([MS-OXPROPS] section 2.601) contains the name of an attachment file, modified so that it can be correlated with TNEF messages, as specified in [MS-OXTNEF].");
                    }

                    break;

                case PropertyNames.PidTagAttachEncoding:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R624");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R624
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        624,
                        @"[In PidTagAttachEncoding Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachAdditionalInformation:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R629");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R629
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        629,
                        @"[In PidTagAttachAdditionalInformation Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachmentLinkId:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R632");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R632
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        632,
                        @"[In PidTagAttachmentLinkId Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachmentFlags:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R635");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R635
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        635,
                        @"[In PidTagAttachmentFlags Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachmentHidden:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R638");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R638
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBoolean,
                        propertyObj.ValueType,
                        638,
                        @"[In PidTagAttachmentHidden Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachMimeTag:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R959");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R959
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        959,
                        @"[In MIME properties] The type of PidTagAttachMimeTag ([MS-OXPROPS] section 2.593) is PtypString.");

                    break;

                case PropertyNames.PidTagAttachContentId:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R960");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R960
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        960,
                        @"[In MIME Properties] The type of PidTagAttachContentId ([MS-OXPROPS] section 2.578) is PtypString.");
                    break;

                case PropertyNames.PidTagAttachContentLocation:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R961");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R961
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        961,
                        @"[In MIME Properties] The type of PidTagAttachContentLocation ([MS-OXPROPS] section 2.579) is PtypString.");
                    break;

                case PropertyNames.PidTagAttachContentBase:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R962");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R962
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        962,
                        @"[In MIME Properties] The type of PidTagAttachContentBase ([MS-OXPROPS] section 2.577) is PtypString.");
                    break;

                case PropertyNames.PidLidSmartNoAttach:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R89");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R89
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBoolean,
                        propertyObj.ValueType,
                        89,
                        @"[In PidLidSmartNoAttach Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidLidPrivate:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R92");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R92
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBoolean,
                        propertyObj.ValueType,
                        92,
                        @"[In PidLidPrivate Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidLidSideEffects:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R94");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R94
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        94,
                        @"[In PidLidSideEffects Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                    bool isVerifiedR1076 = false;
                    int valueOfPidLidSideEffects = (int)propertyObj.Value;
                    int possibleValueOfPidLidSideEffects = 0 | 0x00000001 | 0x00000008 | 0x00000010 | 0x00000020 | 0x00000040 | 0x00000100 | 0x00000400 | 0x00000800 | 0x00001000 | 0x00002000 | 0x00004000;
                    if ((valueOfPidLidSideEffects | possibleValueOfPidLidSideEffects) == possibleValueOfPidLidSideEffects)
                    {
                        isVerifiedR1076 = true;
                    }

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1076");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1076
                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR1076,
                        1076,
                        @"[In PidLidSideEffects Property] This property [PidLidSideEffects] is set to a bitwise OR of zero or more of the following flags [seOpenToDelete, seNoFrame, seCoerceToInbox, seOpenToCopy, seOpenToMove, seOpenForCtxMenu, seCannotUndoDelete, seCannotUndoCopy, seCannotUndoMove, seHasScript and seOpenToPermDelete].");
                    break;

                case PropertyNames.PidLidCommonStart:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R111");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R111
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypTime,
                        propertyObj.ValueType,
                        111,
                        @"[In PidLidCommonStart Property] Type: PtypTime ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidLidCategories:
                     // Add the debug information
                     this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1153");

                     // Verify MS-OXCMSG requirement: MS-OXCMSG_R1153
                     this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                         PropertyType.PtypMultipleString,
                         propertyObj.ValueType,
                         1153,
                         @"[In PidLidCategories Property] Type: PtypMultipleString ([MS-OXCDATA] section 2.11.1).");

                     break;

                case PropertyNames.PidLidCommonEnd:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R115");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R115
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypTime,
                        propertyObj.ValueType,
                        115,
                        @"[In PidLidCommonEnd Property] Type: PtypTime ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAutoForwarded:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1106");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1106
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBoolean,
                        propertyObj.ValueType,
                        1106,
                        @"[In PidTagAutoForwarded Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAutoForwardComment:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1152");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1152
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1152,
                        @"[In PidTagAutoForwardComment Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidLidClassified:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1159");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1159
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBoolean,
                        propertyObj.ValueType,
                        1159,
                        @"[In PidLidClassified Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagInternetReferences:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1161");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1161
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1161,
                        @"[In PidTagInternetReferences Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");

                    if (PropertyHelper.IsPropertyValid(propertyObj))
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1162");

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R1162
                        // Each message Id is 8 bytes and the first 2 bytes indicate the length of propertyObj.Value, so MS-OXCMSG_R1162 can be verified when (propertyObj.Value.ToString().Length - 2) % 8 == 0.
                        this.Site.CaptureRequirementIfAreEqual<int>(
                            0,
                            (propertyObj.Value.ToString().Length - 2) % 8,
                            1162,
                            @"[In PidTagInternetReferences Property] The PidTagInternetReferences property ([MS-OXPROPS] section 2.738) contains a list of message IDs (MIDs), as specified in [MS-OXCDATA] section 2.2.1.2, that specify the messages to which this reply is related.");
                    }

                    break;

                case PropertyNames.PidLidInfoPathFormName:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1165");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1165
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1165,
                        @"[In PidLidInfoPathFormName Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagMimeSkeleton:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1168");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1168
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        1168,
                        @"[In PidTagMimeSkeleton Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagTnefCorrelationKey:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1175");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1175
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        1175,
                        @"[In PidTagTnefCorrelationKey Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAddressBookDisplayNamePrintable:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2197");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R2197
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        2197,
                        @"[In PidTagAddressBookDisplayNamePrintable Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagCreatorEntryId:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1181");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1181
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        1181,
                        @"[In PidTagCreatorEntryId Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");

                    if (PropertyHelper.IsPropertyValid(propertyObj))
                    {
                        AddressBookEntryID creatorEntryId = new AddressBookEntryID();
                        creatorEntryId.Deserialize((byte[])propertyObj.Value, 0);

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2319");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2319
                        this.Site.CaptureRequirementIfAreEqual<uint>(
                            0x00000000,
                            creatorEntryId.Flags,
                            CdataShortName,
                            2319,
                            @"[In Address Book EntryID Structure] Flags (4 bytes): This value MUST be set to 0x00000000.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2753");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2753
                        this.Site.CaptureRequirementIfAreEqual<uint>(
                            0x00000000,
                            creatorEntryId.Flags,
                            CdataShortName,
                            2753,
                            @"[In Address Book EntryID Structure] However, in any EntryID stored in a property value, these 4 bytes [Flags] MUST be zero, indicating a long-term EntryID.");

                        byte[] expectedValue = new byte[] { 0xDC, 0xA7, 0x40, 0xC8, 0xC0, 0x42, 0x10, 0x1A, 0xB4, 0xB9, 0x08, 0x00, 0x2B, 0x2F, 0xE1, 0x82 };

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2320. The value of ProviderUID is {0}.", creatorEntryId.ProviderUID);

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2320
                        bool isVerifiedR2320 = Common.CompareByteArray(expectedValue, creatorEntryId.ProviderUID);

                        this.Site.CaptureRequirementIfIsTrue(
                            isVerifiedR2320,
                            CdataShortName,
                            2320,
                            @"[In Address Book EntryID Structure] This value [ProviderUID] is used to route EntryIDs to the correct provider and MUST be set to %xDC.A7.40.C8.C0.42.10.1A.B4.B9.08.00.2B.2F.E1.82.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2321");

                        uint version = BitConverter.ToUInt32((byte[])creatorEntryId.Version, 0);

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2321
                        this.Site.CaptureRequirementIfAreEqual<uint>(
                            0x00000001,
                            version,
                            CdataShortName,
                            2321,
                            @"[In Address Book EntryID Structure] Version (4 bytes): This value MUST be set to %x01.00.00.00.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2323. The value of Type is {0}.", creatorEntryId.Type);

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2323
                        bool isVerifiedR2323 = creatorEntryId.Type == 0x00000000
                            || creatorEntryId.Type == 0x00000001
                            || creatorEntryId.Type == 0x00000002
                            || creatorEntryId.Type == 0x00000003
                            || creatorEntryId.Type == 0x00000004
                            || creatorEntryId.Type == 0x00000005
                            || creatorEntryId.Type == 0x00000006
                            || creatorEntryId.Type == 0x00000100
                            || creatorEntryId.Type == 0x00000101
                            || creatorEntryId.Type == 0x00000102
                            || creatorEntryId.Type == 0x00000200;

                        this.Site.CaptureRequirementIfIsTrue(
                            isVerifiedR2323,
                            CdataShortName,
                            2323,
                            @"[In Address Book EntryID Structure,Type (4 bytes): ] It [Type] MUST be one of the values [0x00000000 %x00.00.00.00,0x00000001 %x01.00.00.00,0x00000002 %x02.00.00.00,0x00000003 %x03.00.00.00,0x00000004 %x04.00.00.00,0x00000005 %x05.00.00.00,0x00000006 %x06.00.00.00,0x00000100 %x00.01.00.00,0x00000101 %x01.01.00.00,0x00000102 %x02.01.00.00,0x00000200 %x00.02.00.00] from the following table.");
                    }

                    break;

                case PropertyNames.PidTagLastModifierEntryId:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1184");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1184
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        1184,
                        @"[In PidTagLastModifierEntryId Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidLidAgingDontAgeMe:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1187");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1187
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBoolean,
                        propertyObj.ValueType,
                        1187,
                        @"[In PidLidAgingDontAgeMe Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidLidCurrentVersion:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1190");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1190
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        1190,
                        @"[In PidLidCurrentVersion Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidLidCurrentVersionName:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1192");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1192
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1192,
                        @"[In PidLidCurrentVersionName Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAlternateRecipientAllowed:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1194");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1194
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBoolean,
                        propertyObj.ValueType,
                        1194,
                        @"[In PidTagAlternateRecipientAllowed Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagResponsibility:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1196");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1196
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBoolean,
                        propertyObj.ValueType,
                        1196,
                        @"[In PidTagResponsibility Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagHasNamedProperties:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1203");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1203
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBoolean,
                        propertyObj.ValueType,
                        1203,
                        @"[In PidTagHasNamedProperties Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagRecipientOrder:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1205");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1205
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        1205,
                        @"[In PidTagRecipientOrder Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidNameAcceptLanguage:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1207");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1207
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1207,
                        @"[In PidNameContentBase Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidNameContentBase:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1209");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1209
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1209,
                        @"[In PidNameAcceptLanguage Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagPurportedSenderDomain:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1213");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1213
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1213,
                        @"[In PidTagPurportedSenderDomain Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagStoreEntryId:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1215");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1215
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        1215,
                        @"[In PidTagStoreEntryId Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagTrustSender:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1218");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1218
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypInteger32,
                        propertyObj.ValueType,
                        1218,
                        @"[In PidTagTrustSender] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                    if (PropertyHelper.IsPropertyValid(propertyObj))
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1220. The value of PidTagTrustSender is {0}.", propertyObj.Value);

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R1220
                        bool isVerifiedR1220 = (int)propertyObj.Value == 0 || (int)propertyObj.Value == 1;

                        this.Site.CaptureRequirementIfIsTrue(
                            isVerifiedR1220,
                            1220,
                            @"[In PidTagTrustSender] This property [PidTagTrustSender] is a Boolean integer.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1221. The value of PidTagTrustSender is {0}.", propertyObj.Value);

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R1221
                        bool isVerifiedR1221 = (int)propertyObj.Value == 0 || (int)propertyObj.Value == 1;

                        this.Site.CaptureRequirementIfIsTrue(
                            isVerifiedR1221,
                            1221,
                            @"[In PidTagTrustSender] Valid values [0x00000000, 0x00000001] are given in the following table [table in section 2.2.1.45].");
                    }

                    break;

                case PropertyNames.PidTagSubject:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1237");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1237
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1237,
                        @"[In PidTagSubject Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidNameContentClass:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1248");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1248
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1248,
                        @"[In PidNameContentClass Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagLocalCommitTime:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1887");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1887
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypTime,
                        propertyObj.ValueType,
                        1887,
                        @"[In PidTagLocalCommitTime Property] Type: PtypTime ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagRtfCompressed:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2065");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R2065
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        2065,
                        @"[In PidTagRtfCompressed Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagRtfInSync:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2067");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R2067
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBoolean,
                        propertyObj.ValueType,
                        2067,
                        @"[In PidTagRtfInSync Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagBodyContentId:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2071");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R2071
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        2071,
                        @"[In PidTagBodyContentId Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");

                    if (PropertyHelper.IsPropertyValid(propertyObj))
                    {
                        Guid valueOfPidTagBodyContentId;

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2072. The value of PidTagBodyContentId is {0}.", propertyObj.Value);

                        // Verify MS-OXCMSG requirement: MS-OXCMSG_R2072
                        bool isVerifiedR2072 = Guid.TryParse(propertyObj.Value.ToString(), out valueOfPidTagBodyContentId);

                        this.Site.CaptureRequirementIfIsTrue(
                            isVerifiedR2072,
                            2072,
                            @"[In PidTagBodyContentId Property] The PidTagBodyContentId property ([MS-OXPROPS] section 2.610) contains a GUID corresponding to the current message body.");
                    }

                    break;

                case PropertyNames.PidTagBodyContentLocation:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2075");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R2075
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        2075,
                        @"[In PidTagBodyContentLocation Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagHtml:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2081");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R2081
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        2081,
                        @"[In PidTagHtml Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagPolicyTag:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R160");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R160
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBinary,
                        propertyObj.ValueType,
                        2111,
                        @"[In PidTagPolicyTag Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagArchiveDate:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R212");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R212
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypTime,
                        propertyObj.ValueType,
                        2172,
                        @"[In PidTagArchiveDate Property] Type: PtypTime ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagDisplayName:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R226");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R226
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        226,
                        @"[In PidTagDisplayName Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachPathname:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R607");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R607
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        607,
                        @"[In PidTagAttachPathname Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachDataObject:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R237");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R237
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypComObject,
                        propertyObj.ValueType,
                        237,
                        @"[In PidTagAttachDataObject Property] Type: PtypObject ([MS-OXCDATA] section 2.11.1).");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2722");

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2722
                    bool isVerifiedR2722 = propertyObj.ValueTypeCode == 0x000D;
                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR2722,
                        CdataShortName,
                        2722,
                        @"[In Property Value Types] PtypObject or PtypEmbeddedTable (PT_OBJECT) is  that  the property value is a COM object, as specified in section 2.11.1.5 with Property Type Value 0x000D,%x0D.00.");
                    break;

                case PropertyNames.PidTagTextAttachmentCharset:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1278");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1278
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1278,
                        @"[In PidTagTextAttachmentCharset Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagAttachPayloadClass:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1554");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1554
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1554,
                        @"[In MIME Properties] The type of PidTagAttachPayloadClass ([MS-OXPROPS] section 2.596) is PtypString.");
                    break;

                case PropertyNames.PidTagAttachPayloadProviderGuidString:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1553");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1553
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1553,
                        @"[In MIME Properties] The type of PidTagAttachPayloadProviderGuidString ([MS-OXPROPS] section 2.597) is PtypString.");
                    break;
                case PropertyNames.PidNameContentType:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2037");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R2037
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        2037,
                        @"[In PidNameContentType Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;
                case PropertyNames.PidLidClassification:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1155");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1155
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1155,
                        @"[In PidLidClassification] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidLidClassificationDescription:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1157");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R1157
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        1157,
                        @"[In PidLidClassificationDescription Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagCreatorName:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2040");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R2040
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        2040,
                        @"[In PidTagCreatorName Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;
                case PropertyNames.PidTagRead:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2044");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R2044
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypBoolean,
                        propertyObj.ValueType,
                        2044,
                        @"[In PidTagRead Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                    break;
                case PropertyNames.PidTagBodyHtml:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2063");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R2063
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        2063,
                        @"[In PidTagBodyHtml Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;

                case PropertyNames.PidTagRecipientDisplayName:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2046");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R2046
                    this.Site.CaptureRequirementIfAreEqual<PropertyType>(
                        PropertyType.PtypString,
                        propertyObj.ValueType,
                        2046,
                        @"[In PidTagRecipientDisplayName Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                    break;
                case PropertyNames.PidNameKeywords:
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R108");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R108
                    Site.CaptureRequirementIfAreEqual<PropertyType>(
                         PropertyType.PtypMultipleString,
                         propertyObj.ValueType,
                         108,
                         @"[In PidNameKeywords Property] Type: PtypMultipleString ([MS-OXCDATA] section 2.11.1).");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R110");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R110
                    Site.CaptureRequirementIfIsTrue(
                         ((string[])propertyObj.Value).Length<256,
                         110,
                         @"[In PidNameKeywords Property] The length of each string within the multivalue string is less than 256 characters.");

                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// Verify the value of PidTagDisplayName
        /// </summary>
        /// <param name="pidTagDisplayName">The structure of PidTagDisplayName</param>
        /// <param name="pidTagAttachLongFilename">The structure of PidTagAttachLongFilename</param>
        private void VerifyMessageSyntaxPidTagDisplayName(PropertyObj pidTagDisplayName, PropertyObj pidTagAttachLongFilename)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R228");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R228
            this.Site.CaptureRequirementIfAreEqual<string>(
                pidTagDisplayName.Value.ToString(),
                pidTagAttachLongFilename.Value.ToString(),
                228,
                @"[In PidTagDisplayName Property] This property is set to the same value as the PidTagAttachLongFilename property (section 2.2.2.13).");
        }

        /// <summary>
        /// Verify the length of PidTagSubjectPrefix and PidTagNormalizedSubject
        /// </summary>
        /// <param name="propertyObjPidTagSubjectPrefix"> The structure of PidTagSubjectPrefix </param>
        /// <param name="propertyObjPidTagNormalizedSubject"> The structure of PidTagNormalizedSubject </param>
        private void VerifyMessageSyntaxPidTagSubjectPrefixAndPidTagNormalizedSubject(PropertyObj propertyObjPidTagSubjectPrefix, PropertyObj propertyObjPidTagNormalizedSubject)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R60");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R60
            bool isVerifyR60 = propertyObjPidTagSubjectPrefix.Value.ToString().Length + propertyObjPidTagNormalizedSubject.Value.ToString().Length < 254;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR60,
                60,
                @"[In PidTagSubjectPrefix Property] The sum of the lengths of the PidTagNormalizedSubject property (section 2.2.1.10) and the PidTagSubjectPrefix property MUST be less than 254 characters.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R66");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R66
            bool isVerifyR66 = propertyObjPidTagSubjectPrefix.Value.ToString().Length + propertyObjPidTagNormalizedSubject.Value.ToString().Length < 254;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR66,
                66,
                @"[In PidTagNormalizedSubject Property] The sum of the lengths of the PidTagNormalizedSubject property and the PidTagSubjectPrefix property (section 2.2.1.9) MUST be less than 254 characters.");
        }

        /// <summary>
        /// Verify the value of PidLidCommonStart and PidLidCommonEnd
        /// </summary>
        /// <param name="propertyObjPidLidCommonStart">The structure of PidLidCommonStart.</param>
        /// <param name="propertyObjPidLidCommonEnd"> The structure of PidLidCommonEnd.</param>
        private void VerifyMessageSyntaxPidLidCommonStartAndPidLidCommonEnd(PropertyObj propertyObjPidLidCommonStart, PropertyObj propertyObjPidLidCommonEnd)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R113. The start time is {0}, the end time is {1}.", Convert.ToDateTime(propertyObjPidLidCommonStart.Value), Convert.ToDateTime(propertyObjPidLidCommonEnd.Value));

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R113
            bool isVerifiedR113 = Convert.ToDateTime(propertyObjPidLidCommonStart.Value) <= Convert.ToDateTime(propertyObjPidLidCommonEnd.Value);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR113,
                113,
                @"[In PidLidCommonStart Property] The value of this property [PidLidCommonStart] is less than or equal to the value of the PidLidCommonEnd property (section 2.2.1.19).");

            DateTime pidLidCommonStart = Convert.ToDateTime(propertyObjPidLidCommonStart.Value);
            Site.Assert.IsNotNull(pidLidCommonStart, "The pidLidCommonStart should not null.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R117The start time is {0}, the end time is {1}.", Convert.ToDateTime(propertyObjPidLidCommonStart.Value), Convert.ToDateTime(propertyObjPidLidCommonEnd.Value));

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R117
            bool isVerifiedR117 = Convert.ToDateTime(propertyObjPidLidCommonEnd.Value) >= Convert.ToDateTime(propertyObjPidLidCommonStart.Value);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR117,
                117,
                @"[In PidLidCommonEnd Property] The value of this property [PidLidCommonEnd Property] MUST be greater than or equal to the value of the PidLidCommonStart property (section 2.2.1.18).");
        }

        /// <summary>
        /// Verify the message of PidTagAttachmentLinkIdAndPidTagAttachmentFlags
        /// </summary>
        /// <param name="propertyObj"> The structure of PidTagAttachmentLinkIdAndPidTagAttachmentFlags </param>
        private void VerifyMessageSyntaxPidTagAttachmentLinkIdAndPidTagAttachmentFlags(PropertyObj propertyObj)
        {
            switch (propertyObj.PropertyName)
            {
                case PropertyNames.PidTagAttachmentLinkId:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R634");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R634
                    Site.CaptureRequirementIfAreEqual<int>(
                        0x00000000,
                        (int)propertyObj.Value,
                        634,
                        @"[In PidTagAttachmentLinkId Property] This property [PidTagAttachmentLinkId] MUST be set to 0x00000000 unless overridden by other protocols that extend the Message and Attachment Object Protocol as noted in section 1.4.");
                    break;

                case PropertyNames.PidTagAttachmentFlags:

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R637");

                    // Verify MS-OXCMSG requirement: MS-OXCMSG_R637
                    Site.CaptureRequirementIfAreEqual<int>(
                         0x00000000,
                        (int)propertyObj.Value,
                         637,
                         @"[In PidTagAttachmentFlags Property] This property [PidTagAttachmentFlags] MUST be set to 0x00000000 unless overridden by other protocols that extend the Message and Attachment Object Protocol as noted in section 1.4.");
                    break;
                 default:
                    break;
            }
        }

        /// <summary>
        /// Verify the property of HasNamedProperties 
        /// </summary>
        /// <param name="openMessageResponse"> The response of RopOpenMessage.</param>
        private void VerifyRopOpenMessageResponse(RopOpenMessageResponse openMessageResponse)
        {
            int hasNamedProperties = openMessageResponse.HasNamedProperties;

            bool isVerifyR672 = false;
            if (hasNamedProperties >= 0 && hasNamedProperties <= 0xFF)
            {
                isVerifyR672 = true;
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R672");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R672
            Site.CaptureRequirementIfIsTrue(
                isVerifyR672,
                672,
                @"[In RopOpenMessage ROP Response Buffer] HasNamedProperties: 1 byte.");

            bool isVerifyR684 = false;
            if (openMessageResponse.RowCount <= openMessageResponse.RecipientCount)
            {
                isVerifyR684 = true;
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R684");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R684
            Site.CaptureRequirementIfIsTrue(
                isVerifyR684,
                684,
                @"[In RopOpenMessage ROP Response Buffer] [RowCount] The value MUST be less than or equal to the RecipientCount field.");

            if (openMessageResponse.RecipientRows == null)
            {
                return;
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R685");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R685
            this.Site.CaptureRequirementIfAreEqual<int>(
                (int)openMessageResponse.RowCount,
                openMessageResponse.RecipientRows.Length,
                685,
                @"[In RopOpenMessage ROP Response Buffer] RecipientRows: An array of OpenRecipientRow structures whose number is equal to the value of the RowCount field.");

            // All the reserved fields are set to 0x0000. Take the first recipient row as a sample to validate the format and capture the relevant requirements. 
            OpenRecipientRow openRecipientRow = openMessageResponse.RecipientRows[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R686");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R686
            bool isVerifyR686 = ((openRecipientRow.RecipientType & 0x0F) == 0x00 ||
                   (openRecipientRow.RecipientType & 0x0F) == 0x01 ||
                   (openRecipientRow.RecipientType & 0x0F) == 0x02 ||
                   (openRecipientRow.RecipientType & 0x0F) == 0x03) &&
                   ((openRecipientRow.RecipientType & 0xF0) == 0x00
                   || ((openRecipientRow.RecipientType & 0xF0) | (0x10 | 0x80)) == (0x10 | 0x80));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR686,
                686,
                @"[In RopOpenMessage ROP Response Buffer] [RecipientRows] The value of the RecipientType field specified in [MS-OXCROPS] section 2.2.6.1.2.1 is a bitwise OR of zero or one value [0x01, 0x02, 0x03] from the Types table with zero or more values [0x10, 0x80] from the flags table.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R680");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R680
            Site.CaptureRequirementIfAreEqual<int>(
                openMessageResponse.RecipientColumns.Length,
                openMessageResponse.ColumnCount,
                680,
                @"[In RopOpenMessage ROP Response Buffer] ColumnCount: A 2-BYTE unsigned integer containing the number of elements in the RecipientColumns field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R681");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R681
            Site.CaptureRequirementIfAreEqual<int>(
                openMessageResponse.ColumnCount,
                openMessageResponse.RecipientColumns.Length,
                681,
                @"[In RopOpenMessage ROP Response Buffer] RecipientColumns: An array of PropertyTag structures ([MS-OXCDATA] section 2.9) with the number of elements specified in the ColumnCount field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R683");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R683
            this.Site.CaptureRequirementIfAreEqual<byte>(
                (byte)openMessageResponse.RecipientRows.Length,
                openMessageResponse.RowCount,
                683,
                @"[In RopOpenMessage ROP Response Buffer] RowCount: A 1-byte unsigned integer containing the number of rows in the RecipientRows field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R489. The value of SubjectPrefix is {0}.", openMessageResponse.SubjectPrefix);

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R489
            bool isVerifiedR489 = openMessageResponse.SubjectPrefix.StringType == 0x00
                || openMessageResponse.SubjectPrefix.StringType == 0x01
                || openMessageResponse.SubjectPrefix.StringType == 0x02
                || openMessageResponse.SubjectPrefix.StringType == 0x03
                || openMessageResponse.SubjectPrefix.StringType == 0x04;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR489,
                CdataShortName,
                489,
                @"[In TypedString Structure] StringType (1 byte): The value MUST be one of the following: [0x00, 0x01, 0x02, 0x03, 0x04].");

            if (openMessageResponse.SubjectPrefix.StringType == 0x00)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R490. The value of SubjectPrefix is {0}.", openMessageResponse.SubjectPrefix);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R490
                this.Site.CaptureRequirementIfIsNull(
                    openMessageResponse.SubjectPrefix.String,
                    CdataShortName,
                    490,
                    @"[In TypedString Structure] StringType (1 byte): 0x00: No string is present.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R499. The value of SubjectPrefix is {0}.", openMessageResponse.SubjectPrefix);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R499
                this.Site.CaptureRequirementIfIsNull(
                    openMessageResponse.SubjectPrefix.String,
                    CdataShortName,
                    499,
                    @"[In TypedString Structure] String (optional) (4 bytes): Otherwise [If the StringType field is not set to 0x02, 0x03, or 0x04] this field MUST NOT be present.");
            }
        }

        /// <summary>
        /// Verify the property of HasMessageId
        /// </summary>
        /// <param name="hasMessageId"> Value of HasMessageId </param>
        private void VerifyMessageSyntaxHasMessageId(int hasMessageId)
        {
            bool isVerifyR698 = false;
            if (hasMessageId >= 0 && hasMessageId <= 0xFF)
            {
                isVerifyR698 = true;
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R698");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R698
            Site.CaptureRequirementIfIsTrue(
                isVerifyR698,
                698,
                @"[In RopCreateMessage ROP Response Buffer] HasMessageId: The size of HasMessageId is 1 byte.");
        }

        /// <summary>
        /// Verify the property of SyntaxRowCount
        /// </summary>
        /// <param name="retuResponse"> The response of RopReadRecipients. </param>
        private void VerifyMessageSyntaxRowCount(RopReadRecipientsResponse retuResponse)
        {
            int rowCount = retuResponse.RowCount;

            if (retuResponse.RecipientRows == null)
            {
                return;
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R763");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R763
            bool isVerifyR763 = false;
            if (rowCount >= 0x00 && rowCount < 0xFF)
            {
                isVerifyR763 = true;
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR763,
                763,
                @"[In RopReadRecipients ROP Response Buffer] [RowCount] Is greater than or equal to 0x00 and less than 0xFF.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R762");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R762
            this.Site.CaptureRequirementIfAreEqual<int>(
                retuResponse.RecipientRows.Length,
                retuResponse.RowCount,
                762,
                @"[In RopReadRecipients ROP Response Buffer] RowCount: 1 byte containing the number of elements in the RecipientRows field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R765");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R765
            this.Site.CaptureRequirementIfAreEqual<int>(
                retuResponse.RowCount,
                retuResponse.RecipientRows.Length,
                765,
                @"[In RopReadRecipients ROP Response Buffer] RecipientRows: An array of ReadRecipientRow structures whose number of elements equals the RowCount field.");
        }

        /// <summary>
        /// Verify the property of MessageStatusFlags
        /// </summary>
        /// <param name="retuResponse"> The response of RopSetMessageStatus.</param>
        private void VerifyMessageSyntaxMessageStatusFlags(RopSetMessageStatusResponse retuResponse)
        {
            uint messageStatusFlags = retuResponse.MessageStatusFlags;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R807");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R807
            bool isVerifyR807 = false;
            uint value = 0 | 0x00001000 | 0x00000800 | 0x00002000;
            if ((messageStatusFlags | value) == value)
            {
                isVerifyR807 = true;
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR807,
                807,
                @"[In RopSetMessageStatus ROP Response Buffer] [MessageStatusFlags] This field MUST contain a bitwise OR of zero or more values from the table in section 2.2.1.8.");
        }

        /// <summary>
        /// Verify the property of ReadStatusChanged
        /// </summary>
        /// <param name="ropSetMessageReadFlagResponse"> A structure of RopSetMessageReadFlagResponse </param>
        /// <param name="ropSetMessageReadFlagRequest">Return RopSetMessageReadFlagRequest</param>
        private void VerifyMessageSyntaxReadStatusChanged(RopSetMessageReadFlagResponse ropSetMessageReadFlagResponse, RopSetMessageReadFlagRequest ropSetMessageReadFlagRequest)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R846");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R846
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                ropSetMessageReadFlagResponse.ReadStatusChanged.ToString().Length,
                846,
                @"[In RopSetMessageReadFlag ROP Response Buffer] ReadStatusChanged: 1 byte containing one of the following values [0x00, Nonzero].");

            if (ropSetMessageReadFlagResponse.ReadStatusChanged != 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R848");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R848
                Site.CaptureRequirementIfAreEqual<byte>(
                    ropSetMessageReadFlagRequest.LogonId,
                    (byte)ropSetMessageReadFlagResponse.LogonId,
                    848,
                    @"[In RopSetMessageReadFlag ROP Response Buffer] LogonId: 1 byte containing the LogonID from the request when the value in the ReadStatusChanged field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R851");

                bool isR851Verified = Common.CompareByteArray(ropSetMessageReadFlagRequest.ClientData, ropSetMessageReadFlagResponse.ClientData);

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R851
                this.Site.CaptureRequirementIfIsTrue(
                    isR851Verified,
                    851,
                    @"[In RopSetMessageReadFlag ROP Response Buffer] ClientData: 24 bytes containing the ClientData field from the request when the value in the ReadStatusChanged field is nonzero.");
            }
            else
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R849");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R849
                Site.CaptureRequirementIfIsNull(
                    ropSetMessageReadFlagResponse.LogonId,
                    849,
                    @"[In RopSetMessageReadFlag ROP Response Buffer] [LogonId] 0 bytes otherwise [when the value in ReadStatusChanged is zero].");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R852");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R852
                this.Site.CaptureRequirementIfIsNull(
                    ropSetMessageReadFlagResponse.ClientData,
                    852,
                    @"[In RopSetMessageReadFlag ROP Response Buffer] [ClientData] 0 bytes otherwise [when the value in ReadStatusChanged is zero].");
            }
        }

        /// <summary>
        /// Verify the message of RopSetReadFlags.
        /// </summary>
        /// <param name="setReadFlagsResponse"> The response of RopSetReadFlags.</param>
        private void VerifyRopSetReadFlagsResponse(RopSetReadFlagsResponse setReadFlagsResponse)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R834. The value of PartialCompletion is {0}.", setReadFlagsResponse.PartialCompletion);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R834
            bool isVerifiedR834 = (typeof(byte) == setReadFlagsResponse.PartialCompletion.GetType()) &&
                                 ((Convert.ToBoolean(setReadFlagsResponse.PartialCompletion) == true) ||
                                 (Convert.ToBoolean(setReadFlagsResponse.PartialCompletion) == false));

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR834,
                834,
                @"[In RopSetReadFlags ROP Response Buffer] PartialCompletion: 1 byte boolean flag.");
        }

        /// <summary>
        /// Verify the message of RopCreateAttachmentResponse
        /// </summary>
        /// <param name="ropCreateAttachmentResponse"> The message of RopCreateAttachmentResponse </param>
        /// <param name="nID"> ID for the Attachment object </param>
        private void VerifyDataStructureRopCreateAttachmentResponse(RopCreateAttachmentResponse ropCreateAttachmentResponse, int nID)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R867");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R867
            bool isVerifyR867 = ropCreateAttachmentResponse.AttachmentID.GetType() == typeof(uint) && ropCreateAttachmentResponse.AttachmentID == nID;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR867,
                867,
                @"[In RopCreateAttachment ROP Response Buffer] AttachmentID: 4 bytes containing the ID for the Attachment object that was created.");
        }

        /// <summary>
        /// Verify the message of RopOpenEmbeddedMessageResponse
        /// </summary>
        /// <param name="ropOpenEmbeddedMessageResponse"> The message of RopOpenEmbeddedMessageResponse</param>
        /// <param name="mID"> MID for the Message object </param>
        private void VerifyDataStructureRopOpenEmbeddedMessageResponse(RopOpenEmbeddedMessageResponse ropOpenEmbeddedMessageResponse, ulong mID)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R889");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R889
            bool isVerifyR889 = ropOpenEmbeddedMessageResponse.MessageId.GetType() == typeof(ulong) && ropOpenEmbeddedMessageResponse.MessageId.ToString().IndexOf(mID.ToString()) >= 0;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR889,
                889,
                @"[In RopOpenEmbeddedMessage ROP Response Buffer] MessageId: 8 bytes containing the MID ([MS-OXCDATA] section 2.2.1.2) for the Message object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2180");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2180
            // The message ID should be 8 bytes when the ropOpenEmbeddedMessageResponse.MessageId is not null or empty, because the ReplicaId is 2 bytes and GlobalCounter is 6 bytes.
            this.Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ulong),
                ropOpenEmbeddedMessageResponse.MessageId.GetType(),
                CdataShortName,
                2180,
                @"[In Message ID Structure] It [Message ID] is an 8-byte structure.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2181");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2181
            // The message ID should be 8 bytes when the ropOpenEmbeddedMessageResponse.MessageId is not null or empty, because the ReplicaId is 2 bytes and GlobalCounter is 6 bytes.
            this.Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ulong),
                ropOpenEmbeddedMessageResponse.MessageId.GetType(),
                CdataShortName,
                2181,
                @"[In Message ID Structure] ReplicaId (2 bytes): An unsigned integer identifying a Store object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2182");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2182
            // The message ID should be 8 bytes when the ropOpenEmbeddedMessageResponse.MessageId is not null or empty, because the ReplicaId is 2 bytes and GlobalCounter is 6 bytes.
            this.Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ulong),
                ropOpenEmbeddedMessageResponse.MessageId.GetType(),
                CdataShortName,
                2182,
                @"[In Message ID Structure] GlobalCounter (6 bytes): An unsigned integer identifying the message within its Store object.");
        }

        /// <summary>
        /// Verify the message of RopGetMessageStatus.
        /// </summary>
        /// <param name="retuResponse">The response of RopGetMessageStatus</param>
        private void VerifyGetMessageStatusResponse(RopSetMessageStatusResponse retuResponse)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R793");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R793
            this.Site.CaptureRequirementIfAreEqual<int>(
                4,
                retuResponse.MessageStatusFlags.ToString().Length,
                793,
                @"[In RopGetMessageStatus ROP Response Buffer] MessageStatusFlags: 4 bytes indicating the status of the Message object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R794. The value of MessageStatusFlags is {0}.", retuResponse.MessageStatusFlags);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R794
            bool isVerifiedR794 = false;
            uint messageStatusFlags = retuResponse.MessageStatusFlags;
            uint value = 0 | 0x00001000 | 0x00000800 | 0x00002000;
            if ((messageStatusFlags | value) == value)
            {
                isVerifiedR794 = true;
            }

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR794,
                794,
                @"[In RopGetMessageStatus ROP Response Buffer] [MessageStatusFlags] This field contains a bitwise OR of zero or more values from the table in section 2.2.1.8.");
        }

        /// <summary>
        /// Verify the message of PidTagArchiveDate
        /// </summary>
        /// <param name="pidTagArchiveDateObj"> Structure of PidTagArchiveDate </param>
        /// <param name="pidTagStartDateEtc"> Structure of PidTagStartDateEtc </param>
        private void VerifyMessageSyntaxPidTagArchiveDate(PropertyObj pidTagArchiveDateObj, DateTime pidTagStartDateEtc)
        {
            DateTime dateTimeObj = DateTime.Now;
            dateTimeObj = (DateTime)pidTagArchiveDateObj.Value;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2181");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2181
            Site.CaptureRequirementIfAreEqual<DateTime>(
                dateTimeObj,
                pidTagStartDateEtc,
                2181,
                @"[In PidTagArchiveDate Property] The start date is obtained from the last eight bytes of the PidTagStartDateEtc property (section 2.2.1.58.4).");
        }
    }
}