namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter capture code for MS_OXORULEAdapter.
    /// </summary>
    public partial class MS_OXORULEAdapter : ManagedAdapterBase, IMS_OXORULEAdapter
    {
        /// <summary>
        /// Verify MAPIHTTP transport.
        /// </summary>
        private void VerifyMAPITransport()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower(System.Globalization.CultureInfo.CurrentCulture) == "mapi_http" && Common.IsRequirementEnabled(933, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R933");

                // Verify requirement MS-OXORULE_R933
                // If the transport sequence is MAPIHTTP and the code can reach here, it means that the implementation does support MAPIHTTP transport.
                Site.CaptureRequirement(
                    933,
                    @"[In Appendix A: Product Behavior] Implementation does support this specification [MS-OXCMAPIHTTP]. (Exchange Server 2013 Service Pack 1 (SP1) follows this behavior.)");
            }
        }

        /// <summary>
        /// Verify the Properties in the DEM.
        /// </summary>
        /// <param name="ropGetPropertiesAllResponse">The response of the RopGetProperties call.</param>
        private void VerifyRopGetPropertiesAllForDEM(RopGetPropertiesAllResponse ropGetPropertiesAllResponse)
        {
            // Used to verify R811.
            ushort verifiedPropertiesCount = 0;

            // ropGetPropertiesAllResponse.PropertyValues contains a list of properties set on a Message.
            for (int i = 0; i < ropGetPropertiesAllResponse.PropertyValues.Length; i++)
            {
                // propertyId indicates the Id of a property.
                ushort propertyId = ropGetPropertiesAllResponse.PropertyValues[i].PropertyTag.PropertyId;

                // propertyType indicates the Type of a property.
                ushort propertyType = ropGetPropertiesAllResponse.PropertyValues[i].PropertyTag.PropertyType;

                // propertyValue indicates the value of a property.
                byte[] propertyValue = ropGetPropertiesAllResponse.PropertyValues[i].Value;

                switch (propertyId)
                {
                    // If propertyId is 0x001A, it means the property is PidTagMessageClass.
                    case 0x001A:
                        // Add the debug information. 
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R745");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R745.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypString,
                            propertyType,
                            745,
                            @"[In PidTagMessageClass Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");

                        // Add the debug information. 
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R746");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R746.
                        string messageClassValue = AdapterHelper.PropertyValueConvertToString(propertyValue);
                        Site.CaptureRequirementIfAreEqual<string>(
                            "IPC.Microsoft Exchange 4.0.Deferred Error",
                            messageClassValue,
                            746,
                            @"[In PidTagMessageClass Property] The PidTagMessageClass property ([MS-OXCMSG] section 2.2.1.3) MUST be set to ""IPC.Microsoft Exchange 4.0.Deferred Error"".");

                        if (propertyValue != null)
                        {
                            // This property value was set.
                            verifiedPropertiesCount++;
                        }

                        break;

                    // If propertyId is 0x6648, it means the property is PidTagRuleError.
                    case 0x6648:
                        {
                            // Add the debug information. 
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R747");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R747.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypInteger32,
                                propertyType,
                                747,
                                @"[In PidTagRuleError Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                            uint ruleErrorValue = AdapterHelper.PropertyValueConvertToUint(propertyValue);

                            // Add the debug information. 
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R748: the value of the PidTagRuleError property is {0}", ruleErrorValue);

                            // Verify MS-OXORULE requirement: MS-OXORULE_R748.
                            // The valid error numbers are from 1 to 14, so if the ruleErrorValue is less than 15 means the ruleErrorValue is among the valid error numbers.
                            bool isVerifyR748 = ruleErrorValue > 0 && ruleErrorValue < 15;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR748,
                                748,
                                @"[In PidTagRuleError Property] The PidTagRuleError property ([MS-OXPROPS] section 2.938) MUST be set to one of the following values [0x00000001, 0x00000002, 0x00000003, 0x00000004, 0x00000005, 0x00000006, 0x00000007, 0x00000008, 0x00000009, 0x0000000a, 0x0000000b, 0x0000000c, 0x0000000d, and 0x0000000e], indicating the cause of the error encountered during the execution of the rule (2).");

                            if (propertyValue != null)
                            {
                                // This property value was set.
                                verifiedPropertiesCount++;
                            }
                        }

                        break;

                    // If propertyId is 0x6649, it means the property is PidTagRuleActionType.
                    case 0x6649:
                        {
                            // Add the debug information. 
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R749");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R749.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypInteger32,
                                propertyType,
                                749,
                                @"[In PidTagRuleActionType Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                            if (propertyValue != null)
                            {
                                // This property value was set.
                                verifiedPropertiesCount++;
                            }
                        }

                        break;

                    // If propertyId is 0x6650, it means the property is PidTagRuleActionNumber.
                    case 0x6650:
                        {
                            // Add the debug information. 
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R706");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R706.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypInteger32,
                                propertyType,
                                706,
                                @"[In PidTagRuleActionNumber Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");

                            if (propertyValue != null)
                            {
                                // This property value was set.
                                verifiedPropertiesCount++;
                            }
                        }

                        break;

                    // If propertyId is 0x6681, it means the property is PidTagRuleProvider.
                    case 0x6681:
                        {
                            // Add the debug information. 
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R710");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R710.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypString,
                                propertyType,
                                710,
                                @"[In PidTagRuleProvider Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");

                            if (propertyValue != null)
                            {
                                // This property value was set.
                                verifiedPropertiesCount++;
                            }
                        }

                        break;

                    // If propertyId is 0x6674, it means the property is PidTagRuleId.
                    case 0x6674:
                        {
                            // Add the debug information. 
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R716");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R716.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypInteger64,
                                propertyType,
                                716,
                                @"[In PidTagRuleId Property] Type: PtypInteger64 ([MS-OXCDATA] section 2.11.1).");

                            if (propertyValue != null)
                            {
                                // This property value was set.
                                verifiedPropertiesCount++;
                            }
                        }

                        break;
                }
            }

            // Add the debug information. 
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R811");

            // Verify MS-OXORULE requirement: MS-OXORULE_R811.
            // There are 8 properties specified in section 2.2.7,
            // but the PidTagDamOriginalEntryId property and the PidTagRuleFolderEntryId property cannot get in RopGetProperteisAll ROP,
            // so currently there are 6 properties verified in this ROP, others will be verified in VerifyPropertiesInTable function.
            Site.CaptureRequirementIfAreEqual<ushort>(
                6,
                verifiedPropertiesCount,
                811,
                @"[In Handling Errors During Rule Processing] The server MUST generate the DEM in the following manner: 2. Set the property values on the DEM as specified in section 2.2.7.");
        }

        /// <summary>
        /// Verify the RopModifyRules operation.
        /// </summary>
        /// <param name="ropModifyRulesResponse">RopModifyRules response structure.</param>
        /// <param name="ropModifyRulesRequest">RopModifyRules request structure.</param>
        private void VerifyRopModifyRules(RopModifyRulesResponse ropModifyRulesResponse, RopModifyRulesRequest ropModifyRulesRequest)
        {
            // Add the debug information. 
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R678");

            // Verify MS-OXORULE requirement: MS-OXORULE_R678.
            Site.CaptureRequirementIfAreEqual<byte>(
                ropModifyRulesRequest.InputHandleIndex,
                ropModifyRulesResponse.InputHandleIndex,
                678,
                @"[In RopModifyRules ROP Response Buffer] InputHandleIndex (1 byte): The input handle in the response buffer MUST be the same as the index to the input handle in the request buffer for this operation [Processing RopModifyRules ROP Response].");
        }

        /// <summary>
        /// Verify the RopGetRulesTable operation.
        /// </summary>
        /// <param name="ropGetRulesTableResponse">RopGetRulesTable response buffer.</param>
        /// <param name="ropGetRulesTableRequest">RopGetRulesTable request buffer.</param>
        private void VerifyRopGetRulesTable(RopGetRulesTableResponse ropGetRulesTableResponse, RopGetRulesTableRequest ropGetRulesTableRequest)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R690");

            // Verify MS-OXORULE requirement: MS-OXORULE_R690.
            Site.CaptureRequirementIfAreEqual<byte>(
                ropGetRulesTableRequest.OutputHandleIndex,
                ropGetRulesTableResponse.OutputHandleIndex,
                690,
                @"[In RopGetRulesTable ROP Response Buffer] OutputHandleIndex (1 byte): MUST be set to the value of the OutputHandleIndex field specified in the request.");
        }

        /// <summary>
        /// Verify the RopGetPropertiesAll operation performed on extended rules.
        /// </summary>
        /// <param name="ropGetPropertiesAllResponse">RopGetPropertiesAll response buffer.</param>
        private void VerifyRopGetPropertiesAllForExtendedRules(RopGetPropertiesAllResponse ropGetPropertiesAllResponse)
        {
            // isFAIMessage indicates whether the current message is an FAI Message.
            bool isFAIMessage = false;

            // PidTagRuleMessageNameValue indicates the property value of PidTagRuleMessageName.
            string pidTagRuleMessageNameValue = null;

            // PidTagMessageClassValue indicates the property value of PidTagMessageClass.
            string pidTagMessageClassValue = null;

            // PidTagRuleMessageSequenceValue indicates the property value of PidTagRuleMessageSequence.
            byte[] pidTagRuleMessageSequenceValue = null;

            // PidTagRuleMessageStateValue indicates the property value of PidTagRuleMessageState.
            byte[] pidTagRuleMessageStateValue = null;

            // PidTagRuleMessageUserFlagsValue indicates the property value of PidTagRuleMessageUserFlags.
            byte[] pidTagRuleMessageUserFlagsValue = null;

            // PidTagRuleMessageLevelValue indicates the property value of PidTagRuleMessageLevel.
            byte[] pidTagRuleMessageLevelValue = null;

            // PidTagRuleMessageProviderValue indicates the property value of PidTagRuleMessageProvider.
            string pidTagRuleMessageProviderValue = null;

            // PidTagRuleMessageProviderDataValue indicates the property value of PidTagRuleMessageProviderData.
            byte[] pidTagRuleMessageProviderDataValue = null;

            // PidTagExtendedRuleMessageActionsValue indicates the property value of PidTagExtendedRuleMessageActions.
            ExtendedRuleActions pidTagExtendedRuleMessageActionsValue = null;

            // PidTagExtendedRuleMessageConditionValue indicates the property value of PidTagExtendedRuleMessageCondition.
            ExtendedRuleCondition pidTagExtendedRuleMessageConditionValue = null;

            // ropGetPropertiesAllResponse.PropertyValues contains a list of properties set on a Message.
            for (int i = 0; i < ropGetPropertiesAllResponse.PropertyValues.Length; i++)
            {
                // propertyId indicates the Id of a property.
                ushort propertyId = ropGetPropertiesAllResponse.PropertyValues[i].PropertyTag.PropertyId;

                // propertyType indicates the Type of a property.
                ushort propertyType = ropGetPropertiesAllResponse.PropertyValues[i].PropertyTag.PropertyType;

                // propertyValue indicates the value of a property.
                byte[] propertyValue = ropGetPropertiesAllResponse.PropertyValues[i].Value;

                switch (propertyId)
                {
                    // If propertyId is 0x001A, it means the property is PidTagMessageClass.
                    case 0x001A:
                        {
                            // Save the property value of PidTagMessageClass.
                            pidTagMessageClassValue = AdapterHelper.PropertyValueConvertToString(propertyValue);

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R167");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R167.
                            Site.CaptureRequirementIfAreEqual<string>(
                                "IPM.ExtendedRule.Message",
                                pidTagMessageClassValue,
                                167,
                                @"[In PidTagMessageClass Property] The PidTagMessageClass property ([MS-OXCMSG] section 2.2.1.3) MUST have a value of ""IPM.ExtendedRule.Message"".");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R165");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R165.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypString,
                                propertyType,
                                165,
                                @"[In PidTagMessageClass Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                        }

                        break;

                    // If propertyId is 0x65EC, it means the property is PidTagRuleMessageName.
                    case 0x65EC:
                        {
                            // Save the property value of PidTagRuleMessageName.
                            pidTagRuleMessageNameValue = AdapterHelper.PropertyValueConvertToString(propertyValue);

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R161");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R161.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypString,
                                propertyType,
                                161,
                                @"[In PidTagRuleMessageName Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2700.");

                            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2700.
                            bool isVerifyR2700 = Common.IsUtf16LEString(propertyValue) && propertyType == 0x001F;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR2700,
                                "MS-OXCDATA",
                                2700,
                                @"[In Property Data Types] PtypString (PT_UNICODE, string) is that Variable size; a string of Unicode characters in UTF-16LE format encoding with terminating null character (0x0000). with Property Type Value  0x001F,%x1F.00.");
                        }

                        break;

                    // If propertyId is 0x65F3, it means the property is PidTagRuleMessageSequence.
                    case 0x65F3:
                        {
                            // Save the property value of PidTagRuleMessageSequence.
                            pidTagRuleMessageSequenceValue = propertyValue;

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2691, the propertyValue.Length should be 4 and the actual value is {0}, the propertyType of this property should be 0x0003 and the actual value is {1}.", propertyValue.Length, propertyType);

                            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2691.
                            bool isVerifyR2691 = propertyValue.Length == 4 && propertyType == 0x0003;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR2691,
                                "MS-OXCDATA",
                                2691,
                                @"[In Property Data Types] PtypInteger32 (PT_LONG, PT_I4, int, ui4) is that 4 bytes; a 32-bit integer [MS-DTYP]: INT32 with Property Type Value 0x0003,%x03.00.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R168");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R168.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypInteger32,
                                propertyType,
                                168,
                                @"[In PidTagRuleMessageSequence Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");
                        }

                        break;

                    // If propertyId is 0x65E9, it means the property is PidTagRuleMessageState.
                    case 0x65E9:
                        {
                            // Save the property value of PidTagRuleMessageState.
                            pidTagRuleMessageStateValue = propertyValue;

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R171");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R171.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypInteger32,
                                propertyType,
                                171,
                                @"[In PidTagRuleMessageState Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");
                        }

                        break;

                    // If propertyId is 0x65ED, it means the property is PidTagRuleMessageLevel.
                    case 0x65ED:
                        {
                            // Save the property value of PidTagRuleMessageLevel.
                            pidTagRuleMessageLevelValue = propertyValue;

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R178");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R178.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypInteger32,
                                propertyType,
                                178,
                                @"[In PidTagRuleMessageLevel Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");
                        }

                        break;

                    // If propertyId is 0x65EB, it means the property is PidTagRuleMessageProvider.
                    case 0x65EB:
                        {
                            // Save the property value of PidTagRuleMessageProvider.
                            pidTagRuleMessageProviderValue = AdapterHelper.PropertyValueConvertToString(propertyValue);

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R182");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R182.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypString,
                                propertyType,
                                182,
                                @"[In PidTagRuleMessageProvider Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                        }

                        break;

                    // If propertyId is 0x0E99, it means the property is PidTagExtendedRuleMessageActions.
                    case 0x0E99:
                        {
                            // Save the property value of PidTagExtendedRuleMessageActions.
                            pidTagExtendedRuleMessageActionsValue = AdapterHelper.PropertyValueConvertToExtendedRuleActions(propertyValue);

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2707, the COUNT value({0}) of the propertyValue should equal to the propertyValue.Length({1}) - 2, and the propertyType({2}) of the property should be 0x0102.", BitConverter.ToInt16(propertyValue, 0), propertyValue.Length, propertyType);

                            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2707.
                            bool isVerifyR2707 = BitConverter.ToInt16(propertyValue, 0) == propertyValue.Length - 2 && propertyType == 0x0102;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR2707,
                                "MS-OXCDATA",
                                2707,
                                @"[In Property Data Types] PtypBinary (PT_BINARY) is that variable size; a COUNT field followed by that many bytes with Property Type Value 0x0102,%x02.01.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R189");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R189.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypBinary,
                                propertyType,
                                189,
                                @"[In PidTagExtendedRuleMessageActions Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");

                            // Verify the ExtendedRuleActions structure related requirements.
                            this.VerifyExtendedRuleMessageActions(pidTagExtendedRuleMessageActionsValue);
                        }

                        break;

                    // If propertyId is 0x0E9A, it means the property is PidTagExtendedRuleMessageCondition.
                    case 0x0E9A:
                        {
                            // Save the property value of PidTagExtendedRuleMessageCondition.
                            pidTagExtendedRuleMessageConditionValue = AdapterHelper.PropertyValueConvertToExtendedRuleCondition(propertyValue);

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R195");

                            // Verify MS-OXORULE requirement: MS-OXORULE_R195.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypBinary,
                                propertyType,
                                195,
                                @"[In PidTagExtendedRuleMessageCondition Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");

                            // Verify the ExtendedRuleCondition structure related requirements.
                            this.VerifyExtendedRuleMessageCondition(pidTagExtendedRuleMessageConditionValue);
                        }

                        break;

                    // If propertyId is 0x0E07, it means the property is PidTagMessageFlags.
                    case 0x0E07:
                        {
                            // If mfFAI flag is set on the property value of PidTagMessageFlags, it means this message is an FAI message.
                            if ((AdapterHelper.PropertyValueConvertToInt(propertyValue) & (uint)PidTagMessageFlag.mfFAI) == (uint)PidTagMessageFlag.mfFAI)
                            {
                                isFAIMessage = true;
                            }
                        }

                        break;
                }
            }

            // isFAIMessage indicates whether the current message is an FAI message.
            if (isFAIMessage)
            {
                if (Common.IsRequirementEnabled(894, this.Site))
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R894.");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R894.
                    // The current message is an FAI message. If the property value of PidTagRuleMessageName is not null, it means this property is set on FAI message.
                    Site.CaptureRequirementIfIsNotNull(
                        pidTagRuleMessageNameValue,
                        894,
                        @"[In PidTagRuleMessageName Property] The implementation does set the PidTagRuleMessageName property ([MS-OXPROPS] section 2.942)  on the FAI message. (Exchange 2003 and above follow this behavior.)");
                }

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R166.");

                // Verify MS-OXORULE requirement: MS-OXORULE_R166.
                // The current message is an FAI message. If the property value of PidTagMessageClass is not null, it means this property is set on FAI message.
                Site.CaptureRequirementIfIsNotNull(
                    pidTagMessageClassValue,
                    166,
                    @"[In PidTagMessageClass Property] The PidTagMessageClass property ([MS-OXCMSG] section 2.2.1.3) MUST be set on the FAI message.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R169");

                // Verify MS-OXORULE requirement: MS-OXORULE_R169.
                // The current message is an FAI message. If the property value of PidTagRuleMessageSequence is not null, it means this property is set on FAI message.
                Site.CaptureRequirementIfIsNotNull(
                    pidTagRuleMessageSequenceValue,
                    169,
                    @"[In PidTagRuleMessageSequence Property] The PidTagRuleMessageSequence property ([MS-OXPROPS] section 2.947) MUST be set on the FAI message.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R172.");

                // Verify MS-OXORULE requirement: MS-OXORULE_R172.
                // The current message is an FAI message. If the property value of PidTagRuleMessageState is not null, it means this property is set on FAI message.
                Site.CaptureRequirementIfIsNotNull(
                    pidTagRuleMessageStateValue,
                    172,
                    @"[In PidTagRuleMessageState Property] The PidTagRuleMessageState property ([MS-OXPROPS] section 2.948) MUST be set on the FAI message.");

                if (Common.IsRequirementEnabled(891, this.Site))
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R891, the property value of the PidTagRuleMessageUserFlags is {0}.", pidTagRuleMessageUserFlagsValue);

                    // Verify MS-OXORULE requirement: MS-OXORULE_R891.
                    Site.CaptureRequirementIfIsNull(
                        pidTagRuleMessageUserFlagsValue,
                        891,
                        @"[In PidTagRuleMessageUserFlags Property] The implementation does not set the PidTagRuleMessageUserFlags property ([MS-OXPROPS] section 2.947) on the FAI message. (Exchange 2003 and above follow this behavior.)");
                }

                if (Common.IsRequirementEnabled(892, this.Site))
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R892");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R892.
                    // The current message is an FAI message. If the property value of PidTagRuleMessageLevel is not null, it means this property is not set on FAI message.
                    Site.CaptureRequirementIfIsNotNull(
                        pidTagRuleMessageLevelValue,
                        892,
                        @"[In PidTagRuleMessageLevel Property] The implementation does set the PidTagRuleMessageLevel property ([MS-OXPROPS] section 2.941) on the FAI message. (Exchange 2003 and above follow this behavior.)");
                }

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R183.");

                // Verify MS-OXORULE requirement: MS-OXORULE_R183.
                // The current message is an FAI message. If the property value of PidTagRuleMessageProvider is not null, it means this property is set on FAI message.
                Site.CaptureRequirementIfIsNotNull(
                    pidTagRuleMessageProviderValue,
                    183,
                    @"[In PidTagRuleMessageProvider Property] The PidTagRuleMessageProvider property ([MS-OXPROPS] section 2.945) MUST be set on the FAI message.");

                if (Common.IsRequirementEnabled(893, this.Site))
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R893");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R893.
                    // The current message is an FAI message. If the property value of PidTagRuleMessageProviderData is null, it means this property is not set on FAI message.
                    Site.CaptureRequirementIfIsNull(
                        pidTagRuleMessageProviderDataValue,
                        893,
                        @"[In PidTagRuleMessageProviderData Property] The implementation does not set the PidTagRuleMessageProviderData property ([MS-OXPROPS] section 2.943) on the FAI message. (Exchange 2003 and above follow this behavior.)");
                }

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R190.");

                // Verify MS-OXORULE requirement: MS-OXORULE_R190.
                // The current message is an FAI message. If the property value of PidTagExtendedRuleMessageActions is not null, it means this property is set on FAI message.
                Site.CaptureRequirementIfIsNotNull(
                    pidTagExtendedRuleMessageActionsValue,
                    190,
                    @"[In PidTagExtendedRuleMessageActions Property] The PidTagExtendedRuleMessageActions property ([MS-OXPROPS] section 2.683) MUST be set on the FAI message.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R196");

                // Verify MS-OXORULE requirement: MS-OXORULE_R196.
                // The current message is an FAI message. If the property value of PidTagExtendedRuleMessageCondition is not null, it means this property is set on FAI message.
                Site.CaptureRequirementIfIsNotNull(
                    pidTagExtendedRuleMessageConditionValue,
                    196,
                    @"[In PidTagExtendedRuleMessageCondition Property] The PidTagExtendedRuleMessageCondition property ([MS-OXPROPS] section 2.684) MUST be set on the FAI message.");
            }
        }

        /// <summary>
        /// Verify ExtendedRuleActions buffer.
        /// </summary>
        /// <param name="extendeRuleActions">ExtendedRuleActions structure to be verified.</param>
        private void VerifyExtendedRuleMessageActions(ExtendedRuleActions extendeRuleActions)
        {
            // Get all of the Named Properties contained in the action buffer.
            PropertyName[] propertyNames = extendeRuleActions.NamedPropertyInformation.NamedProperty;

            // If the propertyNames exists, check whether the names' value contained in this structure are all in Unicode format.
            if (propertyNames != null)
            {
                // isContainNameValue is used to indicate whether exist nameValue in the PropertyName array.
                bool isContainNameValue = false;
                bool isVerifyR192 = true;

                // Check whether every name value contained in the PropertyName structure is in Unicode format.
                for (int i = 0; i < propertyNames.Length; i++)
                {
                    // Get the name value contained in the action buffer.
                    byte[] nameValue = propertyNames[i].Name;
                    if (nameValue != null)
                    {
                        isContainNameValue = true;

                        // If the nameValue cannot be converted to a Unicode string, it means the name value is not in Unicode format.
                        if (Encoding.Unicode.GetString(nameValue) == null)
                        {
                            isVerifyR192 = false;
                            break;
                        }
                    }
                }

                if (isContainNameValue)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R192");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R192.
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR192,
                        192,
                        @"[In PidTagExtendedRuleMessageActions Property] All string values contained in any part of the RuleAction structure MUST be in Unicode format.");
                }
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R208.");

            // Verify MS-OXORULE requirement: MS-OXORULE_R208.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000001,
                extendeRuleActions.RuleVersion,
                208,
                @"[In PidTagExtendedRuleMessageActions Property] RuleVersion (4 bytes): This document defines version 1, and thus this value MUST be set to 0x00000001.");

            // Verify the requirement related to the NamedPropertyInformation.
            this.VerifyNamedPropertyInformation(extendeRuleActions.NamedPropertyInformation);

            // Verify the requirement related to the RuleAction.
            this.VerifyExtendRuleAction(extendeRuleActions.RuleActionBuffer);

            // Add the debug information.
            // The format of the PidTagExtendedRuleMessageActions property: NamedPropertyInformation, RuleVersion, RuleActionsBuffer has been verified by above capture code, so R194 can be verified directly. 
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R194");

            // Verify MS-OXORULE requirement: MS-OXORULE_R194.
            Site.CaptureRequirement(
                194,
                @"[In PidTagExtendedRuleMessageActions Property] The format of the PidTagExtendedRuleMessageActions property is as follows: NamedPropertyInformation, RuleVersion, RuleActionsBuffer.");
        }

        /// <summary>
        /// Verify ExtendedRuleCondition buffer.
        /// </summary>
        /// <param name="extendedRuleCondition">ExtendedRuleCondition structure to be verified.</param>
        private void VerifyExtendedRuleMessageCondition(ExtendedRuleCondition extendedRuleCondition)
        {
            // Get all of the Named Properties contained in this condition property value.
            PropertyName[] propertyNames = extendedRuleCondition.NamedPropertyInformation.NamedProperty;

            if (extendedRuleCondition.RuleRestriction != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R215");

                // Verify MS-OXORULE requirement: MS-OXORULE_R215.
                // If the value of the restrictionType is less than 0x0c means the RuleRestriction is one of the restriction structure.
                bool isVerifyR215 = (uint)extendedRuleCondition.RuleRestriction.RestrictType < 0x0c;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR215,
                    215,
                    @"[In PidTagExtendedRuleMessageCondition Property] RuleRestriction (variable): The condition is expressed as a restriction, as specified in [MS-OXCDATA] section 2.12.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R214");

                // Verify MS-OXORULE requirement: MS-OXORULE_R214.
                // The condition is a ContentRestriction type, so if RuleRestriction in the extendRuleCondition is a ContentRestriction type R214 can be verified.
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ContentRestriction),
                    extendedRuleCondition.RuleRestriction.GetType(),
                    214,
                    @"[In PidTagExtendedRuleMessageCondition Property] RuleRestriction (variable): A structure containing the condition to be evaluated.");
            }

            // If the propertyNames exists, check whether the names' value contained in this structure are all in Unicode format.
            if (propertyNames != null)
            {
                // isContainNameValue is used to indicate whether exist nameValue in the PropertyName array.
                bool isContainNameValue = false;
                bool isVerifyR198 = true;

                // Check whether every name value contained in the PropertyName structure is in Unicode format.
                for (int i = 0; i < propertyNames.Length; i++)
                {
                    // Get the name value contained in the action buffer.
                    // Name is a string value.
                    byte[] nameValue = propertyNames[i].Name;
                    if (nameValue != null)
                    {
                        isContainNameValue = true;

                        // If the nameValue cannot be converted to a Unicode string, it means the name value is not in Unicode format.
                        if (Encoding.Unicode.GetString(nameValue) == null)
                        {
                            isVerifyR198 = false;
                            break;
                        }
                    }
                }

                if (isContainNameValue)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R198");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R198.
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR198,
                        198,
                        @"[In PidTagExtendedRuleMessageCondition Property] All string values contained in any part of this condition property value MUST be in Unicode format.");
                }
            }

            this.VerifyNamedPropertyInformation(extendedRuleCondition.NamedPropertyInformation);

            // Add the debug information.
            // The format of the PidTagExtendedRuleMessageCondition property: NamedPropertyInformation, RuleRestriction has been verified by above capture code, so R200 can be verified directly. 
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R200");

            // Verify MS-OXORULE requirement: MS-OXORULE_R200.
            Site.CaptureRequirement(
                200,
                @"[In PidTagExtendedRuleMessageCondition Property] The format of the PidTagExtendedRuleMessageCondition property is as follows: NamedPropertyInformation, RuleRestriction.");
        }

        /// <summary>
        /// Verify NamedPropertyInformation buffer.
        /// </summary>
        /// <param name="namedPropertyInformation">NamedPropertyInfo structure to be verified.</param>
        private void VerifyNamedPropertyInformation(NamedPropertyInfo namedPropertyInformation)
        {
            // propIds indicates a list of Property Ids of the named property used in the Named Property Information structure it precedes.
            uint[] propIds = namedPropertyInformation.PropId;
            uint namedPropertiesCount = 0;

            // namedProperties indicates a list of Property Names of the named property used in the Named Property Information structure it precedes.
            PropertyName[] namedProperties = namedPropertyInformation.NamedProperty;
            if (namedProperties != null)
            {
                foreach (PropertyName propertyName in namedProperties)
                {
                    this.VerifyPropertyName(propertyName);
                    namedPropertiesCount++;
                }

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R940");

                // Verify MS-OXORULE requirement: MS-OXORULE_R940.
                Site.CaptureRequirement(
                    940,
                    @"[In NamedPropertyInformation Structure] NamedProperties (variable): An array of PropertyName structures.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R221");

                // Verify MS-OXORULE requirement: MS-OXORULE_R221.
                Site.CaptureRequirementIfAreEqual<uint>(
                    namedPropertiesCount,
                    namedPropertyInformation.NoOfNamedProps,
                    221,
                    @"[In NamedPropertyInformation Structure] NoOfNamedProps (2 bytes): An integer that specifies the number of named property mappings that are packed in this structure [NamedPropertyInformation].");

                if (namedPropertyInformation.NoOfNamedProps > 0)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R225");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R225.
                    uint sizeOfNamedProperty = 0;
                    for (int i = 0; i < namedPropertiesCount; i++)
                    {
                        sizeOfNamedProperty += Convert.ToUInt32(namedProperties[i].Size());
                    }

                    Site.CaptureRequirementIfAreEqual<uint>(
                        sizeOfNamedProperty,
                        namedPropertyInformation.NamedPropertiesSize,
                        225,
                        @"[In NamedPropertyInformation Structure] NamedPropertiesSize (4 bytes): The total size, in bytes, of the following fields [NamedProperties].");
                }
            }

            // According to open specification, the Named Property Information structure must contain the PropId-NamedProperty pair.
            // So if some named properties are used, the corresponding property Ids of these named properties must exist,
            // or else, both of them are null.
            // If the named properties are actually used.
            if (propIds != null && namedProperties != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R217, the number of propIds is {0}, the number of namedProperties is {1}", propIds.Length, namedProperties.Length);

                // Verify MS-OXORULE requirement: MS-OXORULE_R217
                // If the number of PropId and NamedProperty are the same, it means the Named Property Information structure contains a list of PropId-NamedProperty pair,
                // and each pair represents a distinct named property used in the structure it precedes.
                Site.CaptureRequirementIfAreEqual<int>(
                    propIds.Length,
                    namedProperties.Length,
                    217,
                    @"[In NamedPropertyInformation Structure] For every distinct (unique) named property used in the structure it precedes, the NamedPropertyInformation structure contains one property ID - named property pair.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R938, the number of propIds is {0}, the number of namedProperties is {1}", propIds.Length, namedProperties.Length);

                // Verify MS-OXORULE requirement: MS-OXORULE_R938
                // If the number of PropId and NamedProperty are the same, it means there are one property ID in PropIds array for each PropertyName structure in the NamedProperties field.
                Site.CaptureRequirementIfAreEqual<int>(
                    propIds.Length,
                    namedProperties.Length,
                    938,
                    @"[In NamedPropertyInformation Structure] There MUST be one property ID in this array [PropIds] for each PropertyName structure in the NamedProperties field.");

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R218");

                // Verify MS-OXORULE requirement: MS-OXORULE_R218.
                bool isVerifyR218 = true;

                // To verify whether each property ID is less than 0x8000.
                foreach (uint propId in propIds)
                {
                    if (propId < 0x8000)
                    {
                        isVerifyR218 = false;
                    }
                }

                // If there doesn't contain the same value in the propIds array, IsContainSameValue(propIds) equals to false,
                // which means each property Id uniquely identifies the named property within an extended rule. 
                isVerifyR218 = isVerifyR218 && (AdapterHelper.IsContainSameValue(propIds) == false);
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR218,
                    218,
                    @"[In NamedPropertyInformation Structure] PropIds (variable): An array of property IDs, each of which is a value of 0x8000 or greater.");
            }

            if (namedPropertyInformation.NoOfNamedProps > 0)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R226");

                // Verify MS-OXORULE requirement: MS-OXORULE_R226.
                // NamedPropertiesSize is not zero means it is present.
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    0x00000000,
                    namedPropertyInformation.NamedPropertiesSize,
                    226,
                    @"[In NamedPropertyInformation Structure] NamedPropertiesSize (4 bytes): Only present if NoOfNamedProps is greater than zero.");
            }
        }

        /// <summary>
        /// Verify RuleAction buffer for stand rule.
        /// </summary>
        /// <param name="ruleAction">RuleAction structure to be verified.</param>
        private void VerifyRuleAction(RuleAction ruleAction)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R232, The value of NoOfActions is {0}", ruleAction.NoOfActions);

            // Verify MS-OXORULE requirement: MS-OXORULE_R232.
            // NoOfActions can be either unsigned-integer or unsigned-short.
            uint actionsNumber = (ruleAction.NoOfActions is uint) ? ((uint)ruleAction.NoOfActions) : ((ushort)ruleAction.NoOfActions);
            bool isVerifyR232 = actionsNumber > 0;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR232,
                232,
                @"[In RuleAction Structure] NoOfActions (2 bytes): This number MUST be greater than zero.");
            foreach (ActionBlock actionBlock in ruleAction.Actions)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R236");

                // Verify MS-OXORULE requirement: MS-OXORULE_R236.
                // There are 2 or 4 bytes of the ActionLength in the ActionBlocks,
                // so the cumulative length in bytes of the subsequent fields of the ActionLength must equal to the length of  ActionBlock.length - lengthOfActionLength.
                int lengthOfActionLength = actionBlock.ActionLength is ushort ? 2 : 4;
                bool isVerifyR236 = (actionBlock.ActionLength is ushort ? (ushort)actionBlock.ActionLength : (uint)actionBlock.ActionLength) == (actionBlock.Serialize().Length - lengthOfActionLength);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR236,
                    236,
                    @"[In ActionBlock Structure] ActionLength (2 bytes): An integer that specifies the cumulative length, in bytes, of the subsequent fields in this ActionBlock structure.");

                // Verify MS-OXORULE requirement: MS-OXORULE_R947.
                bool isVerifyR947 = (byte)actionBlock.ActionType == 0x01 ||
                                    (byte)actionBlock.ActionType == 0x02 ||
                                    (byte)actionBlock.ActionType == 0x03 ||
                                    (byte)actionBlock.ActionType == 0x04 ||
                                    (byte)actionBlock.ActionType == 0x05 ||
                                    (byte)actionBlock.ActionType == 0x06 ||
                                    (byte)actionBlock.ActionType == 0x07 ||
                                    (byte)actionBlock.ActionType == 0x08 ||
                                    (byte)actionBlock.ActionType == 0x09 ||
                                    (byte)actionBlock.ActionType == 0x0A ||
                                    (byte)actionBlock.ActionType == 0x0B;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR947,
                    947,
                    @"[In RuleAction Structure] ActionBlocks (variable):  An array of ActionBlock structures, each of which specifies an action (2) of the rule (2), as specified in section 2.2.5.1.");
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R231");

            // Verify MS-OXORULE requirement: MS-OXORULE_R231.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)ruleAction.Actions.Length,
                actionsNumber,
                231,
                @"[In RuleAction Structure] NoOfActions (2 bytes): Specifies the number of structures that are contained in the ActionBlocks field.");

            // isVerifyR274 is used to verify R274 and R272.
            bool isVerifyR274 = false;

            // isVerifyR287 is used to verify R287 and R272.
            bool isVerifyR287 = false;

            // isVerifyR879 is used to verify R879 and R272.
            bool isVerifyR879 = false;

            for (int i = 0; i < ruleAction.Actions.Length; i++)
            {
                ActionBlock actionBlock = ruleAction.Actions[i];

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R245", "The value of the actionBlock.ActionType{0} should be one of 0x01 to 0x0b", actionBlock.ActionType);

                // Verify MS-OXORULE requirement: MS-OXORULE_R245.
                bool isVerifyR245 = (byte)actionBlock.ActionType == 0x01 ||
                                    (byte)actionBlock.ActionType == 0x02 ||
                                    (byte)actionBlock.ActionType == 0x03 ||
                                    (byte)actionBlock.ActionType == 0x04 ||
                                    (byte)actionBlock.ActionType == 0x05 ||
                                    (byte)actionBlock.ActionType == 0x06 ||
                                    (byte)actionBlock.ActionType == 0x07 ||
                                    (byte)actionBlock.ActionType == 0x08 ||
                                    (byte)actionBlock.ActionType == 0x09 ||
                                    (byte)actionBlock.ActionType == 0x0A ||
                                    (byte)actionBlock.ActionType == 0x0B;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR245,
                    245,
                    @"[In ActionBlock Structure] The valid actions (2) [of ActionType] are listed in the following table: OP_MOVE, OP_COPY, OP_REPLY, OP_OOF_REPLY, OP_DEFER_ACTION, OP_BOUNCE, OP_FORWARD, OP_DELEGATE, OP_TAG, OP_DELETE, OP_MARK_AS_READ.");

                if ((actionBlock.ActionType == ActionType.OP_MOVE) || (actionBlock.ActionType == ActionType.OP_COPY))
                {
                    MoveCopyActionData moveCopyActionData = (MoveCopyActionData)actionBlock.ActionDataValue;
                    ServerEID serverEID = new ServerEID(moveCopyActionData.FolderEID);

                    // The storeObjectEntryID structure could only verified in ExchangeServer2007, R632 also only executed in ExchangeServer2007.
                    if (Common.IsRequirementEnabled(632, this.Site))
                    {
                        StoreObjectEntryID storeObjectEntryID = new StoreObjectEntryID();
                        storeObjectEntryID.Deserialize(moveCopyActionData.StoreEID);
                        this.VerifyStoreObjectEntryID(storeObjectEntryID);
                    }

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R982: the value of Ours is {0}", serverEID.Ours);

                    // Verify MS-OXORULE_R982.
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        0x01,
                        serverEID.Ours,
                        982,
                        @"[In ServerEid Structure] Ours (1 byte): This field MUST be set to 0x01.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R980");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R980
                    // Verify MS-OXORULE_R982 has verified the value of Ours is 0x01, MS-OXORULE_R980 verifies the instance's type.
                    this.Site.CaptureRequirementIfIsInstanceOfType(
                        serverEID,
                        typeof(ServerEID),
                        980,
                        @"[In ServerEid Structure] Ours (1 byte): The value 0x01 indicates that the remaining bytes conform to this structure [ServerEid].");

                    if (moveCopyActionData.StoreEIDSize != 0)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R297");

                        // Verify MS-OXORULE_R297.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)moveCopyActionData.StoreEID.Length,
                            moveCopyActionData.StoreEIDSize,
                            297,
                            @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Standard Rules] StoreEIDSize (2 bytes): An integer that specifies the size, in bytes, of the StoreEID field.");
                    }

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R300");

                    // Verify MS-OXORULE_R300.
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)moveCopyActionData.FolderEID.Length,
                        moveCopyActionData.FolderEIDSize,
                        300,
                        @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Standard Rules] FolderEIDSize (2 bytes): An integer that specifies the size, in bytes, of the FolderEID field.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R294");

                    // Verify MS-OXORULE_R294.
                    // The structure of the OP_MOVE Action Date Buffer is verified by the requirements MS-OXORULE_R297, MS-OXORULE_R298, MS-OXORULE_R300, and MS-OXORULE_R301.
                    Site.CaptureRequirement(
                        294,
                        @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Standard Rules] The OP_MOVE ActionData structure MUST be in the following format for a standard rule. [FolderInThisStore, StoreEIDSize, StoreEID, FolderEIDSize, and FolderEID].");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R880");

                    // Verify MS-OXORULE_R880.
                    // The structure of the OP_MOVE Action Date Buffer is verified by the requirements MS-OXORULE_R297, MS-OXORULE_R298, MS-OXORULE_R300, and MS-OXORULE_R301.
                    Site.CaptureRequirement(
                        880,
                        @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Standard Rules] The OP_COPY ActionData structure MUST be in the following format for a standard rule. [FolderInThisStore, StoreEIDSize, StoreEID, FolderEIDSize, and FolderEID].");
                }

                if (actionBlock.ActionType == ActionType.OP_COPY)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R249");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R249.
                    // If the ActionType of the actionBlock equals the value ActionType.OP_COPY means server could parse the ActionType 0x02.
                    Site.CaptureRequirement(
                        249,
                        @"[In ActionBlock Structure] The value of action type OP_COPY: 0x02.");
                }

                if (actionBlock.ActionType == ActionType.OP_MOVE)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R246");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R246.
                    // If the ActionType of the actionBlock equals the value ActionType.OP_MOVE means server could parse the ActionType 0x01.
                    Site.CaptureRequirement(
                        246,
                        @"[In ActionBlock Structure] The value of action type OP_MOVE: 0x01.");
                }

                if (actionBlock.ActionType == ActionType.OP_FORWARD)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R261");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R261.
                    // If the ActionType of the actionBlock equals the value ActionType.OP_FORWARD means server could parse the ActionType 0x07.
                    Site.CaptureRequirement(
                        261,
                        @"[In ActionBlock Structure] The value of action type OP_FORWARD: 0x07.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R274, the value of ActionFlavor is {0}, the ActionFlavorsForward.x is {1}.", actionBlock.ActionFlavor, ActionFlavorsForward.x);

                    // Verify MS-OXORULE requirement: MS-OXORULE_R274.
                    // In ActionFlavor, the not used Flags must not be set, the used Flags can be set or all the bits are not set, it means ActionFlavor is set according to the description in this requirement.
                    isVerifyR274 = (actionBlock.ActionFlavor & (uint)ActionFlavorsForward.x) == 0;

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR274,
                        274,
                        @"[In Action Flavors] If the value of the ActionType field is ""OP_FORWARD"", the ActionFlavor field contains a combination of the bitwise flags [XXXX (TM) (AT) (NC) (PR) XXXXXXXXXXXXXXXXXXXXXXXXXX] specified as follows.");
                }

                if ((actionBlock.ActionType == ActionType.OP_REPLY) || (actionBlock.ActionType == ActionType.OP_OOF_REPLY))
                {
                    ReplyActionData actionData = new ReplyActionData();
                    actionData.Deserialize(ruleAction.Actions[0].ActionDataValue.Serialize());

                    if (actionBlock.ActionType == ActionType.OP_REPLY)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R287, the value of ActionFlavor is {0}, it should be one of ActionFlavorsReply.NS({1}), ActionFlavorsReply.ST({2}), or 0x00000000", actionBlock.ActionFlavor, ActionFlavorsReply.NS, ActionFlavorsReply.ST);

                        // Verify MS-OXORULE requirement: MS-OXORULE_R287.
                        // According to open specification, in this condition, the allowed value of ActionFlavor are NS, ST and 0x00000000.
                        isVerifyR287 = (actionBlock.ActionFlavor == (uint)ActionFlavorsReply.NS) ||
                                            (actionBlock.ActionFlavor == (uint)ActionFlavorsReply.ST) ||
                                            (actionBlock.ActionFlavor == 0x00000000);

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR287,
                            287,
                            @"[In Action Flavors] If the ActionType field value is ""OP_REPLY"", the ActionFlavor field MUST have one of the values specified in the following table [XXXXXX (ST) (NS) XXXXXXXXXXXXXXXXXXXXXXXX] or zero (0x00000000).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R252");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R252.
                        // If the ActionType of the actionBlock equals the value ActionType.OP_REPLY means server could parse the ActionType 0x03.
                        Site.CaptureRequirement(
                            252,
                            @"[In ActionBlock Structure] The value of action type OP_REPLY: 0x03.");

                        bool isVerifyR992 = actionData.ReplyTemplateGUID.Length != 0 && actionData.ReplyTemplateFID != 0 && actionData.ReplyTemplateMID != 0;

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R992");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R992.
                        // If ReplyTemplateFID, ReplyTemplateMID, ReplyTemplateGUID in the ActionData is not null, R992 can be verified.
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR992,
                            992,
                            @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Standard Rules] The OP_REPLY ActionData structure MUST be in the following format for a standard rule. [ReplyTemplateFID, ReplyTemplateMID, ReplyTemplateGUID]");
                    }

                    ReplyActionData replyActionData = (ReplyActionData)actionBlock.ActionDataValue;

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R308");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R308.
                    Site.CaptureRequirementIfIsTrue(
                        Common.IsGUID(replyActionData.ReplyTemplateGUID),
                        308,
                        @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Standard Rules] ReplyTemplateGUID (16 bytes): A GUID that is generated by the client in the process of creating a reply template.");

                    if (actionBlock.ActionType == ActionType.OP_OOF_REPLY)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R879, the value of ActionFlavor is {0}", "it should be one of ActionFlavorsReply.NS({1}), ActionFlavorsReply.ST({2}), or 0x00000000", actionBlock.ActionFlavor, ActionFlavorsReply.NS, ActionFlavorsReply.ST);

                        // Verify MS-OXORULE requirement: MS-OXORULE_R879.
                        // According to open specification, in this condition, the allowed value of ActionFlavor are NS, ST and 0x00000000.
                        isVerifyR879 = (actionBlock.ActionFlavor == (uint)ActionFlavorsReply.NS) ||
                                            (actionBlock.ActionFlavor == (uint)ActionFlavorsReply.ST) ||
                                            (actionBlock.ActionFlavor == 0x00000000);

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR879,
                            879,
                            @"[In Action Flavors] If the ActionType field value is ""OP_OOF_REPLY"", the ActionFlavor field MUST have one of the values specified in the following table [XXXXXX (ST) (NS) XXXXXXXXXXXXXXXXXXXXXXXX] or zero (0x00000000). ");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R254");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R254.
                        // If the ActionType of the actionBlock equals the value ActionType. OP_OOF_REPLY means server could parse the ActionType and its value must be 0x04.
                        Site.CaptureRequirement(
                            254,
                            @"[In ActionBlock Structure] The value of action type OP_OOF_REPLY: 0x04.");

                        bool isVerifyR993 = actionData.ReplyTemplateGUID.Length != 0 && actionData.ReplyTemplateFID != 0 && actionData.ReplyTemplateMID != 0;

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R993");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R993.
                        // If ReplyTemplateFID, ReplyTemplateMID, ReplyTemplateGUID in the ActionData is not null, R993 can be verified.
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR993,
                            993,
                            @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Standard Rules] The OP_OOF_REPLY ActionData structure MUST be in the following format for a standard rule. [ReplyTemplateFID, ReplyTemplateMID, ReplyTemplateGUID]");
                    }

                    FolderID folderID = new FolderID();
                    folderID.Deserialize(replyActionData.ReplyTemplateFID);
                    if (folderID != null)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2175");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2175.
                        Site.CaptureRequirementIfAreEqual<int>(
                            8,
                            folderID.Size,
                            "MS-OXCDATA",
                            2175,
                            @"[In Folder ID Structure] It [Folder ID] is an 8-byte structure.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2176");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2176.
                        // "Identifying a Store object" is informative.
                        Site.CaptureRequirementIfAreEqual<int>(
                            2,
                            folderID.ReplicaId.Length,
                            "MS-OXCDATA",
                            2176,
                            @"[In Folder ID Structure] ReplicaId (2 bytes): An unsigned integer identifying a Store object.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2177");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2177.
                        // "Identifying the folder within its Store object" is informative.
                        Site.CaptureRequirementIfAreEqual<int>(
                            6,
                            folderID.GlobalCounter.Length,
                            "MS-OXCDATA",
                            2177,
                            @"[In Folder ID Structure] GlobalCounter (6 bytes): An unsigned integer identifying the folder within its Store object.");
                    }

                    MessageID messageID = new MessageID();
                    messageID.Deserialize(replyActionData.ReplyTemplateMID);
                    if (messageID != null)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2180");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2180.
                        Site.CaptureRequirementIfAreEqual<int>(
                            8,
                            messageID.Size,
                            "MS-OXCDATA",
                            2180,
                            @"[In Message ID Structure] It [Message ID] is an 8-byte structure.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2181");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2181.
                        // "Identifying a Store object" is informative.
                        Site.CaptureRequirementIfAreEqual<int>(
                            2,
                            messageID.ReplicaId.Length,
                            "MS-OXCDATA",
                            2181,
                            @"[In Message ID Structure] ReplicaId (2 bytes): An unsigned integer identifying a Store object.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2182");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2182.
                        // "Identifying the message within its Store object" is informative.
                        Site.CaptureRequirementIfAreEqual<int>(
                            6,
                            messageID.GlobalCounter.Length,
                            "MS-OXCDATA",
                            2182,
                            @"[In Message ID Structure] GlobalCounter (6 bytes): An unsigned integer identifying the message within its Store object.");
                    }
                }

                if (actionBlock.ActionType == ActionType.OP_DEFER_ACTION)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R256");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R256.
                    // If the ActionType of the actionBlock equals the value ActionType.OP_DEFER_ACTION means server could parse the ActionType 0x05.
                    Site.CaptureRequirement(
                        256,
                        @"[In ActionBlock Structure] The value of action type OP_DEFER_ACTION: 0x05.");
                }

                if (actionBlock.ActionType == ActionType.OP_TAG)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R265");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R265.
                    // If the ActionType of the actionBlock equals the value ActionType.OP_TAG means server could parse the ActionType 0x09.
                    Site.CaptureRequirement(
                        265,
                        @"[In ActionBlock Structure] The value of action type OP_TAG: 0x09.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R739");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R739.
                    // If the taggedPropertyValueOP_TAG is not null means the action data buffer can be parsed by the TaggedPropertyValue.
                    TaggedPropertyValue taggedPropertyValueOP_Tag = AdapterHelper.ReadTaggedProperty(actionBlock.ActionDataValue.Serialize());

                    Site.CaptureRequirementIfIsNotNull(
                        taggedPropertyValueOP_Tag,
                        739,
                        @"[In OP_TAG ActionData Structure] An OP_TAG ActionData structure is a TaggedPropertyValue structure, packaged as specified in [MS-OXCDATA] section 2.11.4.");
                }

                if ((actionBlock.ActionType == ActionType.OP_FORWARD) ||
                    (actionBlock.ActionType == ActionType.OP_REPLY) ||
                    (actionBlock.ActionType == ActionType.OP_OOF_REPLY))
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R272", "The value of the ActionFlavor is {0}, and the ActionType is {1}", actionBlock.ActionFlavor, actionBlock.ActionType);

                    // Verify MS-OXORULE requirement: MS-OXORULE_R272.
                    // According to open specification, the only action types that support the Action Flavor are OP_REPLY, OP_OOF_REPLY and OP_FORWARD.
                    // The setting of ActionFlavor is verified in R274 and R287, so the condition to verify R272 is "isVerifyR274 || isVerifyR287".
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR274 || isVerifyR287 || isVerifyR879,
                        272,
                        @"[In Action Flavors] The only ActionType field values that currently support an Action Flavor are ""OP_REPLY"", ""OP_OOF_REPLY"" and ""OP_FORWARD"".");
                }
                else
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R273");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R273.
                    Site.CaptureRequirementIfAreEqual<uint>(
                        0,
                        actionBlock.ActionFlavor,
                        273,
                        @"[In Action Flavors] The value of the ActionFlavor field MUST be 0x00000000 if the value of the ActionType field is not one of these values [OP_REPLY, OP_OOF_REPLY, and OP_FORWARD].");
                }

                // When the AT flag in ActionFlavor is set.
                if ((actionBlock.ActionFlavor & (uint)ActionFlavorsForward.AT) == (uint)ActionFlavorsForward.AT)
                {
                    // Verify MS-OXORULE requirement: MS-OXORULE_R280.
                    uint otherActionFlavorFlags = (uint)ActionFlavorsForward.PR |
                                                  (uint)ActionFlavorsForward.NC |
                                                  (uint)ActionFlavorsForward.TM |
                                                  (uint)ActionFlavorsForward.x;

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R280, the value of ActionFlavor is {0}, this value should not be combined with otherActionFlavorFlags({1}).", actionBlock.ActionFlavor, otherActionFlavorFlags);

                    // To verify whether the other ActionFlavor Flags are 0 when the AT flag in ActionFlavor is set.
                    bool isVerifyR280 = (actionBlock.ActionFlavor & otherActionFlavorFlags) == 0;

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR280,
                        280,
                        @"[In Action Flavors] AT (Bitmask 0x00000004): This value MUST NOT be combined with other ActionFlavor flags.");
                }

                // When the TM flag in ActionFlavor is set.
                if ((actionBlock.ActionFlavor & (uint)ActionFlavorsForward.TM) == (uint)ActionFlavorsForward.TM)
                {
                    // Verify MS-OXORULE requirement: MS-OXORULE_R283.
                    uint otherActionFlavorFlags = (uint)ActionFlavorsForward.PR |
                                                  (uint)ActionFlavorsForward.NC |
                                                  (uint)ActionFlavorsForward.AT |
                                                  (uint)ActionFlavorsForward.x;

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R283, the value of ActionFlavor is {0}, this value should not be combined with otherActionFlavorFlags({1}).", actionBlock.ActionFlavor, otherActionFlavorFlags);

                    // To verify whether the other ActionFlavor Flags are 0 when the TM flag in ActionFlavor is set.
                    bool isVerifyR283 = (actionBlock.ActionFlavor & otherActionFlavorFlags) == 0;

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR283,
                        283,
                        @"[In Action Flavors] TM (Bitmask 0x00000008): This value MUST NOT be combined with other ActionFlavor flags.");
                }

                if (actionBlock.ActionType == ActionType.OP_FORWARD && Common.IsRequirementEnabled(636, this.Site))
                {
                    ForwardDelegateActionData forwardOrDelegateData = (ForwardDelegateActionData)actionBlock.ActionDataValue;

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R636");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R636.
                    bool isVerifyR636 = true;
                    foreach (RecipientBlock recipientBlock in forwardOrDelegateData.RecipientsData)
                    {
                        // Whether the PidTagEntryId exists.
                        bool hasPidTagEntryId = false;
                        foreach (TaggedPropertyValue property in recipientBlock.PropertiesData)
                        {
                            // propertyID of PidTagEntryId is 0x0FFF.
                            if (property.PropertyTag.PropertyId == 0x0FFF)
                            {
                                hasPidTagEntryId = true;
                            }
                        }

                        if (!hasPidTagEntryId)
                        {
                            isVerifyR636 = false;
                            break;
                        }
                    }

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR636,
                        636,
                        @"[In Appendix A: Product Behavior] Implementation does require the PidTagEntryId property for action ""OP_FORWARD"". [<8> Section 2.2.5.1.2.4.1: Exchange 2003 and Exchange 2007 also require the PidTagEntryId property for action ""OP_FORWARD"".]");
                }

                if ((actionBlock.ActionType == ActionType.OP_FORWARD) || (actionBlock.ActionType == ActionType.OP_DELEGATE))
                {
                    if (actionBlock.ActionType == ActionType.OP_DELEGATE)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R263");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R263.
                        // If the ActionType of the actionBlock equals the value ActionType.DELEGATE means server could parse the ActionType 0x08.
                        Site.CaptureRequirement(
                            263,
                            @"[In ActionBlock Structure] The value of action type OP_DELEGATE: 0x08.");
                    }

                    ForwardDelegateActionData forwardOrDelegateData = (ForwardDelegateActionData)actionBlock.ActionDataValue;

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R323, the recipient count is {0}.", forwardOrDelegateData.RecipientCount);

                    // Verify MS-OXORULE requirement: MS-OXORULE_R323.
                    bool isVerifyR323 = (forwardOrDelegateData.RecipientCount is ushort ? (ushort)forwardOrDelegateData.RecipientCount : (uint)forwardOrDelegateData.RecipientCount) == forwardOrDelegateData.RecipientsData.Length;

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR323,
                        323,
                        @"[In OP_FORWARD and OP_DELEGATE ActionData Structure] RecipientCount (4 bytes): An integer that specifies the number of RecipientBlockData structures, as specified in section 2.2.5.1.2.4.1, contained in the RecipientBlocks field. ");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R324", "The value of the RecipientCount is {0}, it should greater than 0.", forwardOrDelegateData.RecipientCount);

                    // Verify MS-OXORULE requirement: MS-OXORULE_R324.
                    // The RecipientCount can be either unsigned-integer or unsigned-short.
                    uint recipientCount = (forwardOrDelegateData.RecipientCount is uint) ? ((uint)forwardOrDelegateData.RecipientCount) : ((ushort)forwardOrDelegateData.RecipientCount);
                    bool isVerifyR324 = recipientCount > 0;

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR324,
                        324,
                        @"[In OP_FORWARD and OP_DELEGATE ActionData Structure] RecipientCount (4 bytes): This number MUST be greater than zero.");

                    bool isVerifyR633 = true;
                    bool isVerifyR632 = true;
                    bool isVerifyR331 = true;
                    foreach (RecipientBlock recipientBlock in forwardOrDelegateData.RecipientsData)
                    {
                        if (Common.IsRequirementEnabled(633, this.Site))
                        {
                            if (recipientBlock.Reserved != 0x00)
                            {
                                isVerifyR633 = false;
                                break;
                            }
                        }

                        if (Common.IsRequirementEnabled(632, this.Site))
                        {
                            if (recipientBlock.Reserved != 0x01)
                            {
                                isVerifyR632 = false;
                                break;
                            }
                        }

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R330, the NoOfProperties of the RecipientBlock is {0}.", recipientBlock.NoOfProperties);

                        // Verify MS-OXORULE requirement: MS-OXORULE_R330.
                        bool isVerifyR330 = (recipientBlock.NoOfProperties is ushort ? (ushort)recipientBlock.NoOfProperties : (uint)recipientBlock.NoOfProperties) == (ulong)recipientBlock.PropertiesData.Length;

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR330,
                            330,
                            @"[In RecipientBlockData Structure] NoOfProperties (4 bytes): An integer that specifies the number of structures present in the PropertyValues field.");

                        TaggedPropertyValue[] taggedPropertyValue = recipientBlock.PropertiesData;
                        bool isVerifyR468 = taggedPropertyValue[0].PropertyTag.PropertyId != 0 && taggedPropertyValue[0].PropertyTag.PropertyType != 0;

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R468");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R468.
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR468,
                            "MS-OXCDATA",
                            468,
                            @"[In TaggedPropertyValue Structure] PropertyTag (4 bytes): A PropertyTag structure, as specified in section 2.9, giving the values of the PropertyId and PropertyType fields for the property.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R469");
                        PropertyValue propertyValue = (PropertyValue)taggedPropertyValue[0];

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R469.
                        Site.CaptureRequirementIfIsInstanceOfType(
                            propertyValue,
                            typeof(PropertyValue),
                            "MS-OXCDATA",
                            469,
                            @"[In TaggedPropertyValue Structure] PropertyValue (variable): A PropertyValue structure, as specified in section 2.11.2.1.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R1008");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R1008.
                        // The format of the TaggedPropertyValue structure has been verified by MS-OXCDATA_R468, and MS-OXCDATA_R469.
                        Site.CaptureRequirement(
                            1008,
                            @"[In RecipientBlockData Structure] The format of the TaggedPropertyValue structure is specified in [MS-OXCDATA] section 2.11.4.");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R331.
                        uint properties = (recipientBlock.NoOfProperties is uint) ? ((uint)recipientBlock.NoOfProperties) : ((ushort)recipientBlock.NoOfProperties);
                        if (properties <= 0)
                        {
                            isVerifyR331 = false;
                            break;
                        }
                    }

                    if (Common.IsRequirementEnabled(633, this.Site))
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R633");
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR633,
                            633,
                            @"[[In Appendix A: Product Behavior] Implementation does set this value [Reserved (1 byte)] to 0x00. [<9> Section 2.2.5.1.2.4.1: Exchange 2010, Exchange 2013, and Exchange 2016 set this value to 0x00.]");
                    }

                    if (Common.IsRequirementEnabled(632, this.Site))
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R632");

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR632,
                            632,
                            @"[In Appendix A: Product Behavior] Implementation does set this value [Reserved (1 byte)] to 0x01. [<9> Section 2.2.5.1.2.4.1: Exchange 2003 and Exchange 2007 set this value to 0x01.]");
                    }

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R331");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R331.
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR331,
                        331,
                        @"[In RecipientBlockData Structure] NoOfProperties (4 bytes): This number MUST be greater than zero.");

                    if (Common.IsRequirementEnabled(895, this.Site))
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R895");

                        // Verify requirement: MS-OXORULE_R895.
                        bool isVerifyR895 = true;
                        foreach (RecipientBlock recipientBlock in forwardOrDelegateData.RecipientsData)
                        {
                            // Whether the PidTagDisplayName exists.
                            bool hasPidTagDisplayName = false;

                            // Whether the PidTagEmailAddress exists.
                            bool hasPidTagEmailAddress = false;

                            // Whether the PidTagRecipientType exists.
                            bool hasPidTagRecipientType = false;
                            foreach (TaggedPropertyValue property in recipientBlock.PropertiesData)
                            {
                                // propertyID of hasPidTagDisplayName is 0x3001.
                                if (property.PropertyTag.PropertyId == 0x3001)
                                {
                                    hasPidTagDisplayName = true;
                                }

                                // propertyID of PidTagEmailAddress is 0x3003.
                                if (property.PropertyTag.PropertyId == 0x3003)
                                {
                                    hasPidTagEmailAddress = true;
                                }

                                // propertyID of PidTagRecipientType is 0x0C15.
                                if (property.PropertyTag.PropertyId == 0x0C15)
                                {
                                    hasPidTagRecipientType = true;
                                }
                            }

                            if (!(hasPidTagDisplayName && hasPidTagEmailAddress && hasPidTagRecipientType))
                            {
                                isVerifyR895 = false;
                                break;
                            }
                        }

                        // Verify MS-OXORULE requirement: MS-OXORULE_R895.
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR895,
                            895,
                            @"[In RecipientBlock Data Buffer Packet Structure] No rules (2) does require more [specify values for more properties besides the PidTagDisplayName, PidTagEmailAddress, and PidTagRecipientType in the forward/delegate ActionData buffer] on the implementation. (Exchange 2010 and above follow this behavior.)");
                    }
                }

                if (actionBlock.ActionType == ActionType.OP_BOUNCE)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R259");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R259.
                    // If the ActionType of the actionBlock equals the value ActionType.OP_COPY means server could parse the ActionType 0x06.
                    Site.CaptureRequirement(
                        259,
                        @"[In ActionBlock Structure] The value of action type OP_BOUNCE: 0x06.");

                    BounceActionData bounceActionData = (BounceActionData)actionBlock.ActionDataValue;

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R340", "The value of the BounceCode field is {0}.", bounceActionData.Bounce);

                    // Verify MS-OXORULE requirement: MS-OXORULE_R340.
                    bool isVerifyR340 =
                          (uint)bounceActionData.Bounce == 0x0000000D
                        || (uint)bounceActionData.Bounce == 0x0000001F
                        || (uint)bounceActionData.Bounce == 0x00000026;
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR340,
                        340,
                        @"[In OP_BOUNCE ActionData Structure] The bounce code MUST be one of the following values. [0x0000000D, 0x0000001F, and 0x00000026].");
                }

                if ((actionBlock.ActionType == ActionType.OP_DELETE) || (actionBlock.ActionType == ActionType.OP_MARK_AS_READ))
                {
                    if (actionBlock.ActionType == ActionType.OP_DELETE)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R267");

                        // If the ActionType of the actionBlock equals the value ActionType.OP_MOVE means server could parse the ActionType 0x0A.
                        Site.CaptureRequirement(
                                      267,
                                       @"[In ActionBlock Structure] The value of action type OP_DELETE: 0x0A.");
                    }

                    if (actionBlock.ActionType == ActionType.OP_MARK_AS_READ)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R269");

                        // If the ActionType of the actionBlock equals the value ActionType.OP_COPY means server could parse the ActionType 0x0B.
                        Site.CaptureRequirement(
                                      269,
                                       @"[In ActionBlock Structure] The value of action type OP_MARK_AS_READ: 0x0B.");
                    }

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R346");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R346.
                    // If the Size of ActionDataValue is 0, it means these actions have no ActionData buffer.
                    Site.CaptureRequirementIfAreEqual<int>(
                        0,
                        actionBlock.ActionDataValue.Size(),
                        346,
                        @"[In OP_DELETE or OP_MARK_AS_READ ActionData Structure] These actions [OP_DELETE, OP_MARK_AS_READ] (3) have no ActionData structure.");
                }
            }
        }

        /// <summary>
        /// Verify RuleAction buffer for extend rule.
        /// </summary>
        /// <param name="ruleAction">RuleAction structure to be verified.</param>
        private void VerifyExtendRuleAction(RuleAction ruleAction)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R945: the size of the NoOfActions field is {0} bytes", (uint)ruleAction.NoOfActions);

            // Verify MS-OXORULE requirement: MS-OXORULE_R945.
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ruleAction.NoOfActions.GetType(),
                945,
                @"[In RuleAction Structure] For extended rules, the size of the NoOfActions field is 4 bytes instead of 2 bytes.");

            for (int i = 0; i < ruleAction.Actions.Length; i++)
            {
                ActionBlock actionBlock = ruleAction.Actions[i];

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R949: the size of the ActionLength field is {0} bytes", (uint)actionBlock.ActionLength);

                // Verify MS-OXORULE requirement: MS-OXORULE_R949.
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(uint),
                    actionBlock.ActionLength.GetType(),
                    949,
                    @"[In ActionBlock Structure] ActionLength (2 bytes): For extended rules, the size of the ActionLength field is 4 bytes instead of 2 bytes.");

                if ((actionBlock.ActionType == ActionType.OP_MOVE) || (actionBlock.ActionType == ActionType.OP_COPY))
                {
                    MoveCopyActionDataOfExtendedRule actionData = (MoveCopyActionDataOfExtendedRule)actionBlock.ActionDataValue;

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R970.");

                    // Verify MS-OXORULE_R970.
                    Site.CaptureRequirementIfAreEqual<uint>(
                        (uint)actionData.StoreEID.Length,
                        actionData.StoreEIDSize,
                        970,
                        @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Extended Rules] StoreEIDSize (4 bytes): An integer that specifies the size, in bytes, of the StoreEID field.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R975.");

                    // Verify MS-OXORULE_R975.
                    Site.CaptureRequirementIfAreEqual<uint>(
                        (uint)actionData.FolderEID.Length,
                        actionData.FolderEIDSize,
                        975,
                        @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Extended Rules] FolderEIDSize (4 bytes): An integer that specifies the size, in bytes, of the FolderEID field.");

                    if (actionBlock.ActionType == ActionType.OP_COPY)
                    {
                        bool isVerifyR967 = actionData.FolderEID.Length != 0 && actionData.StoreEID.Length != 0;

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R967");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R967.
                        // If StoreEIDSize, StoreEID, FolderEIDSize, and FolderEID in the ActionData is not null, R967 can be verified.
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR967,
                            967,
                            @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Extended Rules] The OP_COPY ActionData structure MUST be in the following format for an extended rule.  [StoreEIDSize, StoreEID, FolderEIDSize, and FolderEID].");
                    }

                    if (actionBlock.ActionType == ActionType.OP_MOVE)
                    {
                        bool isVerifyR966 = actionData.FolderEID.Length != 0 && actionData.StoreEID.Length != 0;

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R966, the length of FolderEID is {0}, and the length of StoreEID is {1}", actionData.FolderEID.Length, actionData.StoreEID.Length);

                        // Verify MS-OXORULE requirement: MS-OXORULE_R966.
                        // If StoreEIDSize, StoreEID, FolderEIDSize, and FolderEID in the ActionData is not null, R966 can be verified.
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR966,
                            966,
                            @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Extended Rules] The OP_MOVE ActionData structure MUST be in the following format for an extended rule.  [StoreEIDSize, StoreEID, FolderEIDSize, and FolderEID].");
                    }
                }

                if ((actionBlock.ActionType == ActionType.OP_REPLY) || (actionBlock.ActionType == ActionType.OP_OOF_REPLY))
                {
                    ReplyActionDataOfExtendedRule actionData = (ReplyActionDataOfExtendedRule)actionBlock.ActionDataValue;

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R997.");

                    // Verify MS-OXORULE_R997.
                    Site.CaptureRequirementIfAreEqual<uint>(
                        (uint)actionData.ReplyTemplateMessageEID.Length,
                        actionData.MessageEIDSize,
                        997,
                        @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Extended Rules] MessageEIDSize (4 bytes): An integer that specifies the size, in bytes, of the ReplyTemplateMessageEID field.");

                    if (actionBlock.ActionType == ActionType.OP_REPLY)
                    {
                        bool isVerifyR995 = actionData.ReplyTemplateMessageEID.Length != 0 && actionData.ReplyTemplateGUID.Length != 0;

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R995, the length of ReplyTemplateMessageEID is {0}, and the actionData.ReplyTemplateGUID is {1}", (uint)actionData.ReplyTemplateMessageEID.Length, actionData.ReplyTemplateGUID.Length);

                        // Verify MS-OXORULE requirement: MS-OXORULE_R995.
                        // If ReplyTemplateFID, ReplyTemplateMID, ReplyTemplateGUID in the ActionData is not null, R995 can be verified.
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR995,
                            995,
                            @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Extended Rules] The OP_REPLY ActionData structure MUST be in the following format for an extended rule. [MessageEIDSize, ReplyTemplateMessageEID, ReplyTemplateGUID]");
                    }

                    if (actionBlock.ActionType == ActionType.OP_OOF_REPLY)
                    {
                        bool isVerifyR996 = actionData.ReplyTemplateMessageEID.Length != 0 && actionData.ReplyTemplateGUID.Length != 0;

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R996, the length of ReplyTemplateMessageEID is {0}, and the actionData.ReplyTemplateGUID is {1}", (uint)actionData.ReplyTemplateMessageEID.Length, actionData.ReplyTemplateGUID.Length);

                        // Verify MS-OXORULE requirement: MS-OXORULE_R996.
                        // If ReplyTemplateFID, ReplyTemplateMID, ReplyTemplateGUID in the ActionData is not null, R996 can be verified.
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR996,
                            996,
                            @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Extended Rules] The OP_OOF_REPLY ActionData structure MUST be in the following format for an extended rule. [MessageEIDSize, ReplyTemplateMessageEID, ReplyTemplateGUID]");
                    }
                }
            }
        }

        /// <summary>
        /// Verify the specific property.
        /// </summary>
        /// <param name="propertyTags">PropertyTag array.</param>
        private void VerifyPropertiesSpecific(PropertyTag[] propertyTags)
        {
            foreach (PropertyTag propertyTag in propertyTags)
            {
                // propertyID of PidTagHasRules is 0x663A.
                if (propertyTag.PropertyId == 0x663A)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R718");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R718.
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)PropertyType.PtypBoolean,
                        propertyTag.PropertyType,
                        718,
                        @"[In PidTagHasRules Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                }

                // propertyID of PidTagReplyTemplateId is 0x65C2.
                if (propertyTag.PropertyId == 0x65C2)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R726");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R726.
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)PropertyType.PtypBinary,
                        propertyTag.PropertyType,
                        726,
                        @"[In PidTagReplyTemplateId Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                }
            }
        }

        /// <summary>
        /// Verify the RopGetProperties in RuleData Structure.
        /// </summary>
        /// <param name="ruleData">RuleData Structure.</param>
        private void VerifyPropertiesInRuleData(RuleData ruleData)
        {
            // RuleData contains a list of properties.
            for (int i = 0; i < ruleData.PropertyValues.Length; i++)
            {
                // propertyId indicates the Id of a property.
                ushort propertyId = (ruleData.PropertyValues[i] as TaggedPropertyValue).PropertyTag.PropertyId;

                // propertyType indicates the Type of a property.
                ushort propertyType = (ruleData.PropertyValues[i] as TaggedPropertyValue).PropertyTag.PropertyType;

                // propertyValue indicates the value of a property.
                byte[] propertyValue = (ruleData.PropertyValues[i] as TaggedPropertyValue).Value;
                switch (propertyId)
                {
                    // If propertyId is 0x6679, it means the property is PidTagRuleCondition.
                    case 0x6679:
                        {
                            RuleCondition ruleCondition = new RuleCondition();
                            ruleCondition.Deserialize(propertyValue);

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2705, the value of the propertyType is {0}.", propertyType);

                            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2705.
                            // Stack De-serialize this structure as this requirement, if ruleCondition is not null,
                            // "a byte array representing one or more Restriction structures" can be covered.
                            bool isVerifyR2705 = ruleCondition != null && propertyType == 0x00FD;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR2705,
                                "MS-OXCDATA",
                                2705,
                                @"[In Property Value Types] PtypRestriction (PT_SRESTRICT) is that variable size; a byte array representing one or more Restriction structures as specified in section 2.12. with Property Type Value 0x00FD,%xFD.00.");
                        }

                        break;

                    // If propertyId is 0x6680, it means the property is PidTagRuleActions.
                    case 0x6680:
                        {
                            RuleAction ruleAction = new RuleAction();
                            ruleAction.Deserialize(propertyValue);

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2706, the value of the propertyType is {0}.", propertyType);

                            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2706.
                            // Stack De-serialize this structure as this requirement, if ruleAction is not null,
                            // "a 16-bit COUNT of actions (not bytes) followed by that many Rule Action structures" can be covered.
                            bool isVerifyR2706 = ruleAction != null && propertyType == 0x00FE;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR2706,
                                "MS-OXCDATA",
                                2706,
                                @"[In Property Value Types] PtypRuleAction (PT_ACTIONS) is that variable size; a 16-bit COUNT field followed by that many rule (2) action (2) structures, as specified in [MS-OXORULE] section 2.2.5. with Property Type Value 0x00FE,%xFE.00.");
                        }

                        break;

                    // If propertyId is 0x6674, it means the property is PidTagRuleId.
                    case 0x6674:
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2699", "The length({0}) of the propertyValue should be 8 and the value of the propertyType({1}) should be 0x0014.", propertyValue.Length, propertyType);

                            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2699.
                            bool isVerifyR2699 = propertyValue.Length == 8 && propertyType == 0x0014;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR2699,
                                "MS-OXCDATA",
                                2699,
                                @"[In Property Data Types] PtypInteger64 (PT_LONGLONG, PT_I8, i8, ui8) is that 8 bytes; a 64-bit integer [MS-DTYP]: LONGLONG with Property Type Value 0x0014,%x14.00.");
                        }

                        break;
                }
            }
        }

        /// <summary>
        /// Verify the RopGetProperties in Rules Table.
        /// </summary>
        /// <param name="propertyTags">PropertyTag array.</param>
        /// <param name="quertyRowsResponse">Response of quertyRows.</param>
        private void VerifyPropertiesInTable(PropertyTag[] propertyTags, RopQueryRowsResponse quertyRowsResponse)
        {
            // propertyTags contains a list of PropertyTag.
            for (int i = 0; i < propertyTags.Length; i++)
            {
                // propertyId indicates the Id of a property.
                ushort propertyId = propertyTags[i].PropertyId;

                // propertyType indicates the Type of a property.
                ushort propertyType = propertyTags[i].PropertyType;

                // propertyValue indicates the value of a property.
                byte[] propertyValue = quertyRowsResponse.RowData.PropertyRows[quertyRowsResponse.RowData.PropertyRows.Count - 1].PropertyValues[i].Value;

                // If propertyId is 0x6674, it means the property is PidTagRuleId.
                if (propertyId == 0x6674)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2699", "The length({0}) of the propertyValue should be 8 and the value of the propertyType({1}) should be 0x0014.", propertyValue.Length, propertyType);

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2699.
                    bool isVerifyR2699 = propertyValue.Length == 8 && propertyType == 0x0014;

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR2699,
                        "MS-OXCDATA",
                        2699,
                        @"[In Property Data Types] PtypInteger64 (PT_LONGLONG, PT_I8, i8, ui8) is that 8 bytes; a 64-bit integer [MS-DTYP]: LONGLONG with Property Type Value 0x0014,%x14.00.");
                }

                // If propertyId is 0x001A, it means the property is PidTagMessageClass.
                if (propertyId == 0x001A)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R351");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R351.
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)PropertyType.PtypString,
                        propertyType,
                        351,
                        @"[In PidTagMessageClass] Type: PtypString ([MS-OXCDATA] section 2.11.1).");

                    if (this.targetOfRop == TargetOfRop.ForDAM)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R352");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R352.
                        Site.CaptureRequirementIfAreEqual<string>(
                            "IPC.Microsoft Exchange 4.0.Deferred Action",
                            AdapterHelper.PropertyValueConvertToString(propertyValue),
                            352,
                            @"[In PidTagMessageClass] The PidTagMessageClass property ([MS-OXCMSG] section 2.2.1.3) MUST be set to ""IPC.Microsoft Exchange 4.0.Deferred Action"".");
                    }
                }

                // If propertyId is 0x6647, it means the property is PidTagDamBackPatched.
                if (propertyId == 0x6647)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2698", "The length({0}) of the propertyValue should be 1 and the value of the propertyType({1}) should be 0x000b.", propertyValue.Length, propertyType);

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2698.
                    bool isVerifyR2698 = propertyValue.Length == 1 && propertyType == 0x000B;

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR2698,
                        "MS-OXCDATA",
                        2698,
                        @"[In Property Data Types] PtypBoolean (PT_BOOLEAN. bool) is that 1 byte, restricted to 1 or 0 [MS-DTYP]: BOOLEAN with Property Type Value 0x000B, %x0B.00.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R353");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R353.
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)PropertyType.PtypBoolean,
                        propertyType,
                        353,
                        @"[In PidTagDamBackPatched property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                }

                // If propertyId is 0x6646, it means the property is PidTagDamOriginalEntryId.
                if (propertyId == 0x6646)
                {
                    if (this.targetOfRop == TargetOfRop.ForDAM)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R356");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R356.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypBinary,
                            propertyType,
                            356,
                            @"[In PidTagDamOriginalEntryId] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    }
                    else if (this.targetOfRop == TargetOfRop.ForDEM)
                    {
                        // Add the debug information. 
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R712");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R712.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypBinary,
                            propertyType,
                            712,
                            @"[In PidTagDamOriginalEntryId Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    }
                }

                // If propertyId is 0x6681, it means the property is PidTagRuleProvider.
                if (propertyId == 0x6681)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R358");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R358.
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)PropertyType.PtypString,
                        propertyType,
                        358,
                        @"[In PidTagRuleProvider] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
                }

                // If propertyId is 0x6651, it means the property is PidTagRuleFolderEntryId.
                if (propertyId == 0x6651)
                {
                    if (this.targetOfRop == TargetOfRop.ForDAM)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R360");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R360.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypBinary,
                            propertyType,
                            360,
                            @"[In PidTagRuleFolderEntryId] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    }
                    else if (this.targetOfRop == TargetOfRop.ForDEM)
                    {
                        // Add the debug information. 
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R714");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R714.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypBinary,
                            propertyType,
                            714,
                            @"[In PidTagRuleFolderEntryId Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                    }
                }

                // If propertyId is 0x6645, it means the property is PidTagClientActions.
                if (propertyId == 0x6645)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R362");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R362.
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)PropertyType.PtypBinary,
                        propertyType,
                        362,
                        @"[In PidTagClientActions] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");

                    if (this.targetOfRop == TargetOfRop.ForDAM)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R364");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R364.
                        // Stack has constructed the rule action buffer according to section 2.2.5.
                        // If ruleAction is not null, it means the propertyValue of PidTagClientActions can be parsed as the RuleAction format successfully,
                        // which indicates this buffer has been packed according to this requirement.
                        RuleAction ruleAction = new RuleAction();
                        ruleAction = AdapterHelper.PropertyValueConvertToRuleAction(propertyValue);

                        Site.CaptureRequirementIfIsNotNull(
                            ruleAction,
                            364,
                            @"[In PidTagClientActions] The buffer MUST be packed according to the RuleAction structure specified in section 2.2.5.");
                    }
                }

                // If propertyId is 0x6675, it means the property is PidTagRuleIds.
                if (propertyId == 0x6675)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R367");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R367.
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)PropertyType.PtypBinary,
                        propertyType,
                        367,
                        @"[In PidTagRuleIds] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");

                    if (this.targetOfRop == TargetOfRop.ForDAM)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R369");

                        // Verify MS-OXORULE requirement: MS-OXORULE_R369.
                        bool isVerifyR369 = AdapterHelper.PropertyValueConvertToBinary(propertyValue).Length % 8 == 0;

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR369,
                            369,
                            @"[In PidTagRuleIds] The length of this binary property[PidTagRuleIds] MUST be a multiple of 8 bytes.");
                    }
                }

                // If propertyId is 0x6741, it means the property is PidTagDeferredActionMessageOriginalEntryId.
                if (propertyId == 0x6741)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2704");

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2704.
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        0x00FB,
                        propertyType,
                        "MS-OXCDATA",
                        2704,
                        @"[In Property Value Types] PtypServerId (PT_SVREID) is that variable size; a 16-bit COUNT field followed by a structure specified in  section 2.11.1.4. with Property Type Value 0x00FB,%xFB.00.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R370");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R370.
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)PropertyType.PtypServerId,
                        propertyType,
                        370,
                        @"[In PidTagDeferredActionMessageOriginalEntryId Property] Type: PtypServerId ([MS-OXCDATA] section 2.11.1).");
                }

                // propertyID of PidTagHasDeferredActionMessages is 0x3FEA.
                if (propertyId == 0x3FEA)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R722, the propertyType of the PidTagHasDeferredActionMessages is {0}", propertyType);

                    // Verify MS-OXORULE requirement: MS-OXORULE_R722.
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)PropertyType.PtypBoolean,
                        propertyType,
                        722,
                        @"[In PidTagHasDeferredActionMessages Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1).");
                }

                // propertyID of PidTagRwRulesStream is 0x6802.
                if (propertyId == 0x6802)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R1010, the propertyType of the PidTagRwRulesStream is {0}", propertyType);

                    // Verify MS-OXORULE requirement: MS-OXORULE_R1010.
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)PropertyType.PtypBoolean,
                        propertyType,
                        1010,
                        @"[PidTagRwRulesStream Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
                }
            }
        }

        /// <summary>
        /// Verify the StoreObjectEntryID structure.
        /// </summary>
        /// <param name="storeEntryId">The StoreObjectEntryID to be verified.</param>
        private void VerifyStoreObjectEntryID(StoreObjectEntryID storeEntryId)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2249");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2249.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                storeEntryId.Flags,
                "MS-OXCDATA",
                2249,
                @"[In Store Object EntryIDs Structure] Flags (4 bytes): This value MUST be set to 0x00000000.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2250.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2250.
            bool isVerifyR2250 = Common.CompareByteArray(
                                     storeEntryId.ProviderUID,
                                     new byte[] { 0x38, 0xA1, 0xBB, 0x10, 0x05, 0xE5, 0x10, 0x1A, 0xA1, 0xBB, 0x08, 0x00, 0x2B, 0x2A, 0x56, 0xC2 });

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2250,
                "MS-OXCDATA",
                2250,
                @"[In Store Object EntryID Structure] This value [ProviderUID] is used to route EntryIDs to the correct provider and MUST be set to %x38.A1.BB.10.05.E5.10.1A.A1.BB.08.00.2B.2A.56.C2.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2251");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2251.
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                storeEntryId.Version,
                "MS-OXCDATA",
                2251,
                @"[In Store Object EntryIDs Structure] Version (1 byte): This value MUST be set to zero.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2252");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2252.
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                storeEntryId.Flag,
                "MS-OXCDATA",
                2252,
                @"[In Store Object EntryIDs Structure] Flag (1 byte): This value MUST be set to zero.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2253.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2253.
            bool isVerifyR2253 = Common.CompareByteArray(
                                     storeEntryId.DLLFileName,
                                     new byte[] { 0x45, 0x4D, 0x53, 0x4D, 0x44, 0x42, 0x2E, 0x44, 0x4C, 0x4C, 0x00, 0x00, 0x00, 0x00 });

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2253,
                "MS-OXCDATA",
                2253,
                @"[In Store Object EntryIDs Structure] DLLFileName (14 bytes): This field MUST be set to the following value, which represents ""emsmdb.dll"": %x45.4D.53.4D.44.42.2E.44.4C.4C.00.00.00.00.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2254");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2254.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                storeEntryId.WrappedFlags,
                "MS-OXCDATA",
                2254,
                @"[In Store Object EntryIDs Structure] WrappedFlags (4 bytes): This value MUST be set to 0x00000000.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2255.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2255.
            bool isVerifyR2255 = Common.CompareByteArray(storeEntryId.WrappedProviderUID, new byte[] { 0x1B, 0x55, 0xFA, 0x20, 0xAA, 0x66, 0x11, 0xCD, 0x9B, 0xC8, 0x00, 0xAA, 0x00, 0x2F, 0xC4, 0x5A }) ||
                                 Common.CompareByteArray(storeEntryId.WrappedProviderUID, new byte[] { 0x1C, 0x83, 0x02, 0x10, 0xAA, 0x66, 0x11, 0xCD, 0x9B, 0xC8, 0x00, 0xAA, 0x00, 0x2F, 0xC4, 0x5A });

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2255,
                "MS-OXCDATA",
                2255,
                @"[In Store Object EntryIDs Structure] WrappedProvider UID (16 bytes): This field MUST be set to one of the values [%x1B.55.FA.20.AA.66.11.CD.9B.C8.00.AA.00.2F.C4.5A,%x1C.83.02.10.AA.66.11.CD.9B.C8.00.AA.00.2F.C4.5A] in the following table.");

            if (storeEntryId.ObjectType == StoreObjectType.Mailbox)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2256.");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R2256.
                bool isVerifyR2256 = Common.CompareByteArray(
                                         storeEntryId.WrappedProviderUID,
                                         new byte[] { 0x1B, 0x55, 0xFA, 0x20, 0xAA, 0x66, 0x11, 0xCD, 0x9B, 0xC8, 0x00, 0xAA, 0x00, 0x2F, 0xC4, 0x5A });

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2256,
                    "MS-OXCDATA",
                    2256,
                    @"[In Store Object EntryIDs Structure] WrappedProvider UID (16 bytes): The store object type is Mailbox Store object, providerUID value is %x1B.55.FA.20.AA.66.11.CD.9B.C8.00.AA.00.2F.C4.5A.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2260");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R2260.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x0000000C,
                    storeEntryId.WrappedType,
                    "MS-OXCDATA",
                    2260,
                    @"[In Store Object EntryID Structure] For a mailbox store this value [WrappedType] MUST be set to %x0C.00.00.00.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2265");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R2265.
                // If MailboxDN is not null, it means this field is present.
                Site.CaptureRequirementIfIsNotNull(
                    storeEntryId.MailBoxDN,
                    "MS-OXCDATA",
                    2265,
                    @"[In Store Object EntryIDs Structure] MailboxDN (variable): This field is present only for mailbox databases.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2264");

                // Remove the "\0" from string.
                StringBuilder mailBoxDN = new StringBuilder();
                for (int i = 0; i < storeEntryId.MailBoxDN.Length - 1; i++)
                {
                    mailBoxDN.Append(storeEntryId.MailBoxDN.ToCharArray()[i]);
                }

                bool isSingleZeroByteTerminated = storeEntryId.MailBoxDN.ToCharArray()[storeEntryId.MailBoxDN.Length - 1] == '\0';
                bool isVerifiedR2264 = Common.IsDNMatchABNF(mailBoxDN.ToString(), DNFormat.X500Dn) && isSingleZeroByteTerminated;

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R2264.
                Site.CaptureRequirementIfIsTrue(
                    isVerifiedR2264,
                    "MS-OXCDATA",
                    2264,
                    @"[In Store Object EntryIDs Structure] MailboxDN (variable):  A string of single-byte characters terminated by a single zero byte and representing the X500 DN of the mailbox, as specified in [MS-OXOAB].");
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2263");
            bool isVerifiedR2263 = storeEntryId.ServerShortname.ToCharArray()[storeEntryId.ServerShortname.Length - 1] == '\0';

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2263.
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR2263,
                "MS-OXCDATA",
                2263,
                @"[In Store Object EntryIDs Structure] ServerShortname (variable): A string of single-byte characters terminated by a single zero byte, indicating the short name or NetBIOS name of the server.");
        }

        /// <summary>
        /// Verify PropertyName structure.
        /// </summary>
        /// <param name="propertyName">The property name.</param>
        private void VerifyPropertyName(PropertyName propertyName)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R17, The actual value of Kind is {0}", propertyName.Kind);

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R17.
            bool isVerifyR17 = propertyName.Kind == 0x00 ||
                               propertyName.Kind == 0x01 ||
                               propertyName.Kind == 0xFF;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR17,
                "MS-OXCDATA",
                17,
                @"[In PropertyName Structure] Kind (1 byte): The possible values [0x00, 0x01, 0xFF] for the Kind field are in the following table.");

            if (propertyName.Kind == 0x00)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R25");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R25.
                // LID is not 0 means it is present.
                Site.CaptureRequirementIfIsNotNull(
                    propertyName.LID,
                    "MS-OXCDATA",
                    25,
                    @"[In PropertyName Structure] LID (optional) (4 bytes): This field is present only if the value of the Kind field is equal to 0x00.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R26");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R26.
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(uint),
                    propertyName.LID.GetType(),
                    "MS-OXCDATA",
                    26,
                    @"[In PropertyName Structure] LID (optional) (4 bytes):  An unsigned integer that identifies the named property within its property set.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R27");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R27.
                // NameSize!=null means NameSize present.
                Site.CaptureRequirementIfIsNull(
                    propertyName.NameSize,
                    "MS-OXCDATA",
                    27,
                    @"[In PropertyName Structure] NameSize (optional) (1 byte):  This field is present only if the value of the Kind field is equal to 0x01.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R29");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R29.
                // That Name does not equal null means Name presents.
                Site.CaptureRequirementIfIsNull(
                    propertyName.Name,
                    "MS-OXCDATA",
                    29,
                    @"[In PropertyName Structure] Name (optional) (variable): This field is present only if Kind is equal to 0x01.");
            }
            else
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R25");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R25.
                // LID is not 0 means it is present.
                Site.CaptureRequirementIfIsNull(
                    propertyName.LID,
                    "MS-OXCDATA",
                    25,
                    @"[In PropertyName Structure] LID (optional) (4 bytes): This field is present only if the value of the Kind field is equal to 0x00.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R27");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R27.
                // NameSize!=null means NameSize present.
                Site.CaptureRequirementIfIsNotNull(
                    propertyName.NameSize,
                    "MS-OXCDATA",
                    27,
                    @"[In PropertyName Structure] NameSize (optional) (1 byte):  This field is present only if the value of the Kind field is equal to 0x01.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R29");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R29.
                // That Name does not equal null means Name presents.
                Site.CaptureRequirementIfIsNotNull(
                    propertyName.Name,
                    "MS-OXCDATA",
                    29,
                    @"[In PropertyName Structure] Name (optional) (variable): This field is present only if Kind is equal to 0x01.");
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R941");

            // Verify MS-OXORULE requirement: MS-OXORULE_R941.
            // The format of the PropertyName structure has been verified by the above capture code, so R941 can be verified directly.
            Site.CaptureRequirement(
                941,
                @"[In NamedPropertyInformation Structure] The format of the PropertyName structure is specified in [MS-OXCDATA] section 2.6.1.");
        }

        /// <summary>
        /// Verify the FolderEntryID structure.
        /// </summary>
        /// <param name="folderEntryID">The FolderEntryID to be verified.</param>
        /// <param name="storeObjectType">Store Object Type.</param>
        private void VerifyFolderEntryID(FolderEntryID folderEntryID, StoreObjectType storeObjectType)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2225");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2225.
            Site.CaptureRequirementIfAreEqual<int>(
                0,
                BitConverter.ToInt32(folderEntryID.Flag, 0),
                "MS-OXCDATA",
                2225,
                @"[In Folder EntryID Structure] Flags (4 bytes): This value MUST be set to 0x00000000.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2229");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2229.
            Site.CaptureRequirementIfAreEqual<int>(
                16,
                folderEntryID.DataBaseGUID.Length,
                "MS-OXCDATA",
                2229,
                @"[In Folder EntryID Structure] DatabaseGuid (16 bytes): A GUID associated with the Store object and corresponding to the ReplicaId field of the FID structure.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2230");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2230.
            Site.CaptureRequirementIfAreEqual<int>(
                6,
                folderEntryID.GlobalCounter.Length,
                "MS-OXCDATA",
                2230,
                @"[In Folder EntryID Structure] GlobalCounter (6 bytes): An unsigned integer identifying the folder.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2231");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2231.
            Site.CaptureRequirementIfAreEqual<short>(
                0,
               BitConverter.ToInt16(folderEntryID.Pad, 0),
               "MS-OXCDATA",
                2231,
                @"[In Folder EntryID Structure] Pad (2 bytes): This value MUST be set to zero.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R301");

            // The structure of FolderEId is verified by the requirements MS-OXCDATA_R2225, MS-OXCDATA_R2229, MS-OXCDATA_R2230, and MS-OXCDATA_R2231.
            // Verify MS-OXORULE_R301.
            Site.CaptureRequirement(
                301,
                @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Standard Rules] FolderEID (variable): A structure that identifies the destination folder.");

            if (storeObjectType == StoreObjectType.Mailbox)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2226");

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R2226.
                Site.CaptureRequirementIfAreEqual<byte[]>(
                    this.mailboxGUID,
                    folderEntryID.ProviderUID,
                    "MS-OXCDATA",
                    2226,
                    @"[In Folder EntryID Structure] Provider UID (16 bytes): For a folder in a private mailbox, this value MUST be set to value of the MailboxGuid field from the RopLogon ROP response buffer ([MS-OXCROPS] section 2.2.3.1).");
            }
        }

        /// <summary>
        /// Verify the MessageEntryID structure.
        /// </summary>
        /// <param name="messageEntryID">The MessageEntryID to be verified.</param>
        private void VerifyMessageEntryID(MessageEntryID messageEntryID)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2236");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2236.
            Site.CaptureRequirementIfAreEqual<int>(
                0,
                BitConverter.ToInt32(messageEntryID.Flag, 0),
                "MS-OXCDATA",
                2236,
                @"[In Message EntryID Structure] Flags (4 bytes): This value MUST be set to 0x00000000.");

            // Add the debug information.
            int messageType = BitConverter.ToInt16(messageEntryID.MessageType, 0);
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2239, The actual value of MessageType is {0}", messageType);

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2239.
            bool isVerifyR2239 = messageType == 0x0001 ||
                                 messageType == 0x0003 ||
                                 messageType == 0x0005 ||
                                 messageType == 0x0007 ||
                                 messageType == 0x0009 ||
                                 messageType == 0x000B ||
                                 messageType == 0x000C;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2239,
                "MS-OXCDATA",
                2239,
                @"[In Message EntryID Structure] MessageType (2 bytes): One of several Store object types specified in the table in section 2.2.4.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2241");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2241.
            Site.CaptureRequirementIfIsTrue(
                Common.IsGUID(messageEntryID.FolderDataBaseGUID),
                "MS-OXCDATA",
                2241,
                @"[In Message EntryID Structure] FolderDatabaseGuid (16 bytes): A GUID associated with the Store object of the folder in which the message resides and corresponding to the DatabaseReplicationId field of the folder ID structure.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2242");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2242.
            Site.CaptureRequirementIfAreEqual<int>(
                6,
                messageEntryID.FolderGlobalCounter.Length,
                "MS-OXCDATA",
                2242,
                @"[In Message EntryID Structure] FolderGlobalCounter (6 bytes): An unsigned integer identifying the folder in which the message resides.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2243");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2243.
            Site.CaptureRequirementIfAreEqual<short>(
                0,
               BitConverter.ToInt16(messageEntryID.Pad, 0),
               "MS-OXCDATA",
                2243,
                @"[In Message EntryID Structure] Pad (2 bytes): This value MUST be set to zero.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2244");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2244.
            Site.CaptureRequirementIfIsTrue(
                Common.IsGUID(messageEntryID.FolderDataBaseGUID),
                "MS-OXCDATA",
                2244,
                @"[In Message EntryID Structure] MessageDatabaseGuid (16 bytes): A GUID associated with the Store object of the message and corresponding to the DatabaseReplicationId field of the message ID structure.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2245");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2245.
            Site.CaptureRequirementIfAreEqual<int>(
                6,
                messageEntryID.FolderGlobalCounter.Length,
                "MS-OXCDATA",
                2245,
                @"[In Message EntryID Structure] MessageGlobalCounter (6 bytes): An unsigned integer identifying the message.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2246");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2246.
            Site.CaptureRequirementIfAreEqual<short>(
                0,
               BitConverter.ToInt16(messageEntryID.Pad, 0),
               "MS-OXCDATA",
                2246,
                @"[In Message EntryID Structure] Pad (2 bytes): This value MUST be set to zero.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2237");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2237.
            Site.CaptureRequirementIfAreEqual<byte[]>(
                this.mailboxGUID,
                messageEntryID.ProviderUID,
                "MS-OXCDATA",
                2237,
                @"[In Message EntryID Structure,ProviderUID (16 bytes):] For a folder in a private mailbox, this value MUST be set to the value of the MailboxGuid field from the RopLogon ROP response buffer ([MS-OXCROPS] section 2.2.3.1).");
        }
    }
}