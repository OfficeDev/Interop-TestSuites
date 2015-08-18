namespace Microsoft.Protocols.TestSuites.MS_OXCTABL
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Provide AdapterHelper Methods.
    /// </summary>
    public static class AdapterHelper
    {
        /// <summary>
        /// The site which is used to print log information.
        /// </summary>
        private static ITestSite site;

        /// <summary>
        /// Gets or sets the instance of site.
        /// </summary>
        public static ITestSite Site
        {
            get { return AdapterHelper.site; }
            set { AdapterHelper.site = value; }
        }

        /// <summary>
        /// Generate a sample valid RestrictData structure.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        /// <param name="tableType">Identify the table type</param>
        /// <returns>A valid RestrictData structure</returns>
        public static byte[] GenerateValidRestrictData(ITestSite site, TableType tableType)
        {
            byte[] tempData = null;
            byte[] restrictData = null;
            int index = 0;

            switch (tableType)
            {
                case TableType.CONTENT_TABLE:
                case TableType.ATTACHMENTS_TABLE:
                case TableType.PERMISSIONS_TABLE:
                case TableType.INVALID_TABLE:
                    string senderName = Common.GetConfigurationPropertyValue("Sender1Name", site) + "\0";
                    tempData = Encoding.Unicode.GetBytes(senderName);
                    restrictData = new byte[tempData.Length + 13];
                    index = 0;
                    restrictData[index++] = 0x03;
                    restrictData[index++] = 0x02;
                    restrictData[index++] = 0x00;
                    restrictData[index++] = 0x04;
                    restrictData[index++] = 0x00;
                    restrictData[index++] = 0x1f;
                    restrictData[index++] = 0x00;
                    restrictData[index++] = 0x1a;
                    restrictData[index++] = 0x0c;
                    restrictData[index++] = 0x1f;
                    restrictData[index++] = 0x00;
                    restrictData[index++] = 0x1a;
                    restrictData[index++] = 0x0c;
                    Array.Copy(tempData, 0, restrictData, index, tempData.Length);
                    break;
                case TableType.HIERARCHY_TABLE:
                    string displayName = Constants.TestGetHierarchyTableFolderName1 + "\0";
                    tempData = Encoding.Unicode.GetBytes(displayName);
                    restrictData = new byte[tempData.Length + 13];
                    index = 0;
                    restrictData[index++] = 0x03;
                    restrictData[index++] = 0x02;
                    restrictData[index++] = 0x00;
                    restrictData[index++] = 0x04;
                    restrictData[index++] = 0x00;
                    restrictData[index++] = 0x1f;
                    restrictData[index++] = 0x00;
                    restrictData[index++] = 0x01;
                    restrictData[index++] = 0x30;
                    restrictData[index++] = 0x1f;
                    restrictData[index++] = 0x00;
                    restrictData[index++] = 0x01;
                    restrictData[index++] = 0x30;
                    Array.Copy(tempData, 0, restrictData, index, tempData.Length);
                    break;
                case TableType.RULES_TABLE:
                    string ruleName = Constants.RuleNameMarkAsRead + "\0"; // The RuleProvider property is set to a default value.
                    tempData = Encoding.Unicode.GetBytes(ruleName);
                    restrictData = new byte[tempData.Length + 13];
                    restrictData[index++] = 0x03;
                    restrictData[index++] = 0x02;
                    restrictData[index++] = 0x00;
                    restrictData[index++] = 0x04;
                    restrictData[index++] = 0x00;
                    restrictData[index++] = 0x1f;
                    restrictData[index++] = 0x00;
                    restrictData[index++] = 0x82;
                    restrictData[index++] = 0x66;
                    restrictData[index++] = 0x1f;
                    restrictData[index++] = 0x00;
                    restrictData[index++] = 0x82;
                    restrictData[index++] = 0x66;
                    Array.Copy(tempData, 0, restrictData, index, tempData.Length);
                    break;
                default:
                    Site.Assert.Fail("The value of the tableType field is {0}, this is not a valid table type", tableType);
                    break;
            }

            return restrictData;
        }

        /// <summary>
        /// Generate rule properties value.
        /// </summary>
        /// <param name="name">Rule name.</param>
        /// <returns>Rule properties</returns>
        public static RuleProperties GenerateRuleProperties(string name)
        {
            RuleProperties ruleProperties = new RuleProperties
            {
                UserFlag = Constants.PidTagRuleUserFlags1,
                Provider = Constants.PidTagRuleProvider,
                ProviderData = Constants.PidTagRuleProviderData,
                ConditionSubjectName = Constants.RuleConditionSubjectContainString,
                Name = name
            };
            return ruleProperties;
        }

        /// <summary>
        /// Generate a sample valid RuleData structure.
        /// </summary>
        /// <param name="actionType">ActionType in a rule</param>
        /// <param name="ruleDataType">Rule data value.</param>
        /// <param name="ruleSequence">Rule sequence in PidTagRuleSequence of the RuleData</param>
        /// <param name="ruleState">Specifies value of PidTagRuleState of the RuleData</param>
        /// <param name="actionData">The actionData needed to construct the RuleData</param>
        /// <param name="ruleProperties">Rule properties structure.</param>
        /// <returns>A valid RuleData structure</returns>
        public static RuleData GenerateValidRuleData(ActionTypes actionType, TestRuleDataType ruleDataType, int ruleSequence, RuleState ruleState, IActionData actionData, RuleProperties ruleProperties)
        {
            RuleData ruleData = new RuleData
            {
                RuleDataFlags = (byte)ruleDataType
            };

            List<TaggedPropertyValue> propertyList = new List<TaggedPropertyValue>();
            TaggedPropertyValue pidTagRuleSequence = new TaggedPropertyValue();
            PropertyTag pidTagRuleSequencePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)TestPropertyName.PidTagRuleSequence,
                PropertyType = (ushort)PropertyTypeName.PtypInteger32
            };
            pidTagRuleSequence.PropertyTag = pidTagRuleSequencePropertyTag;
            pidTagRuleSequence.Value = BitConverter.GetBytes(ruleSequence);
            propertyList.Add(pidTagRuleSequence);

            TaggedPropertyValue pidTagRuleLevel = new TaggedPropertyValue();
            PropertyTag pidTagRuleLevelPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)TestPropertyName.PidTagRuleLevel,
                PropertyType = (ushort)PropertyTypeName.PtypInteger32
            };
            pidTagRuleLevel.PropertyTag = pidTagRuleLevelPropertyTag;
            pidTagRuleLevel.Value = BitConverter.GetBytes(0);
            propertyList.Add(pidTagRuleLevel);

            TaggedPropertyValue pidTagRuleState = new TaggedPropertyValue();
            PropertyTag pidTagRuleStatePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)TestPropertyName.PidTagRuleState,
                PropertyType = (ushort)PropertyTypeName.PtypInteger32
            };
            pidTagRuleState.PropertyTag = pidTagRuleStatePropertyTag;
            pidTagRuleState.Value = BitConverter.GetBytes((uint)ruleState);
            propertyList.Add(pidTagRuleState);

            if (ruleProperties.Name.Length > 0)
            {
                TaggedPropertyValue pidTagRuleName = new TaggedPropertyValue();
                PropertyTag pidTagRuleNamePropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)TestPropertyName.PidTagRuleName,
                    PropertyType = (ushort)PropertyTypeName.PtypString
                };
                pidTagRuleName.PropertyTag = pidTagRuleNamePropertyTag;
                pidTagRuleName.Value = Encoding.Unicode.GetBytes((string)ruleProperties.Name + "\0");
                propertyList.Add(pidTagRuleName);
            }

            if (ruleProperties.Provider.Length > 0)
            {
                TaggedPropertyValue pidTagRuleProvider = new TaggedPropertyValue();
                PropertyTag pidTagRuleProviderPropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)TestPropertyName.PidTagRuleProvider,
                    PropertyType = (ushort)PropertyTypeName.PtypString
                };
                pidTagRuleProvider.PropertyTag = pidTagRuleProviderPropertyTag;
                pidTagRuleProvider.Value = Encoding.Unicode.GetBytes((string)ruleProperties.Provider + "\0");
                propertyList.Add(pidTagRuleProvider);
            }

            if (ruleProperties.UserFlag.Length > 0)
            {
                TaggedPropertyValue pidTagRuleUserFlags = new TaggedPropertyValue();
                PropertyTag pidTagRuleUserFlagsPropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)TestPropertyName.PidTagRuleUserFlags,
                    PropertyType = (ushort)PropertyTypeName.PtypInteger32
                };
                pidTagRuleUserFlags.PropertyTag = pidTagRuleUserFlagsPropertyTag;
                pidTagRuleUserFlags.Value = BitConverter.GetBytes(int.Parse(ruleProperties.UserFlag));
                propertyList.Add(pidTagRuleUserFlags);
            }

            if (ruleProperties.ProviderData.Length > 0)
            {
                TaggedPropertyValue pidTagRuleProviderData = new TaggedPropertyValue();
                PropertyTag pidTagRuleProviderDataPropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)TestPropertyName.PidTagRuleProviderData,
                    PropertyType = (ushort)PropertyTypeName.PtypBinary
                };
                pidTagRuleProviderData.PropertyTag = pidTagRuleProviderDataPropertyTag;
                pidTagRuleProviderData.Value = Common.GetBytesFromBinaryHexString(ruleProperties.ProviderData);
                pidTagRuleProviderData.Value = Common.AddInt16LengthBeforeBinaryArray(pidTagRuleProviderData.Value);
                propertyList.Add(pidTagRuleProviderData);
            }

            if (ruleProperties.ConditionSubjectName.Length > 0)
            {
                AddRuleCondition(propertyList, ruleProperties.ConditionSubjectName);
                AddRuleAction(propertyList, actionType, actionData, Constants.CommonActionFlavor, Constants.RuleActionFlags);
            }

            ruleData.PropertyValues = propertyList.ToArray();
            ruleData.PropertyValueCount = (ushort)propertyList.Count;
            return ruleData;
        }

        /// <summary>
        /// Build a RestrictionData as single ContentRestriction
        /// </summary>
        /// <param name="taggedPValue">A TaggedPropertyValue of this restriction</param>
        /// <returns>Byte array of Restriction Data</returns>
        public static byte[] GenerateContentRestriction(TaggedPropertyValue taggedPValue)
        {
            ContentRestrictions result = new ContentRestrictions
            {
                FuzzyLevelLow = ContentRestrictions.FuzzyLevelLowValue.FL_SUBSTRING,
                FuzzyLevelHigh = ContentRestrictions.FuzzyLevelHighValue.FL_IGNORECASE,
                PropertyTag = taggedPValue.PropertyTag,
                TaggedValue = taggedPValue
            };
            return result.Serialize();
        }

        /// <summary>
        /// Read a TaggedPropertyValue structure from buffer.
        /// </summary>
        /// <param name="buffer">Buffer contain TaggedPropertyValue instance</param>
        /// <returns>A TaggedPropertyvalue structure</returns>
        public static TaggedPropertyValue ReadTaggedProperty(byte[] buffer)
        {
            TaggedPropertyValue tagValue = new TaggedPropertyValue();
            BufferReader bufferReader = new BufferReader(buffer);

            PropertyTag newPropertyTag = new PropertyTag
            {
                PropertyType = bufferReader.ReadUInt16(),
                PropertyId = bufferReader.ReadUInt16()
            };
            tagValue.PropertyTag = newPropertyTag;
            tagValue.Value = ReadValueByType(tagValue.PropertyTag.PropertyType, bufferReader.ReadToEnd());

            return tagValue;
        }

        /// <summary>
        /// Read a value from buffer for special Type.
        /// </summary>
        /// <param name="type">The Type of value</param>
        /// <param name="buffer">Buffer contains value</param>
        /// <returns>Byte array of the value</returns>
        public static byte[] ReadValueByType(ushort type, byte[] buffer)
        {
            byte[] value = null;
            BufferReader bufferReader = new BufferReader(buffer);

            uint length = 0;
            byte[] tmpArray = null;
            uint startIndex = 0;
            uint endIndex = 0;
            switch (type)
            {
                // 2-byte
                // PtypInteger16
                case 0x0002:
                    value = bufferReader.ReadBytes(2);
                    break;

                // 4-byte
                // PtypInteger32
                case 0x0003:

                // PtypFloating32
                case 0x0004:

                // PtypErrorCode 
                case 0x000A:
                    value = bufferReader.ReadBytes(4);
                    break;

                // 8-byte
                // PtypFloating64
                case 0x0005:

                // PtypCurrency 
                case 0x0006:

                // PtypFloatingTime
                case 0x0007:

                // PtypInteger64
                case 0x0014:

                // PtypTime 
                case 0x0040:
                    value = bufferReader.ReadBytes(8);
                    break;

                // 1 byte
                // PtypBoolean 
                case 0x000B:
                    value = new byte[1] { bufferReader.ReadByte() };
                    break;

                // PtypString PT_UNICODE
                case 0x001F:
                    value = Encoding.Unicode.GetBytes(bufferReader.ReadUnicodeString());
                    break;

                // PtypString8
                case 0x001E:
                    value = Encoding.ASCII.GetBytes(bufferReader.ReadASCIIString());
                    break;

                // 16-byte
                // PtypGuid  16bytes
                case 0x0048:

                // PtypServerId 
                case 0x00FB:
                    value = bufferReader.ReadBytes(16);
                    break;

                // PtypRestriction  
                case 0x00FD:
                    tmpArray = bufferReader.ReadToEnd();
                    IRestrictions restriction = null;
                    restriction = new ContentRestrictions();

                    length += restriction.Deserialize(tmpArray);
                    value = bufferReader.ReadBytes(0, length);
                    break;

                // PtypRuleAction  
                case 0x00FE:
                    tmpArray = bufferReader.ReadToEnd();
                    RuleAction ruleAction = new RuleAction();
                    length = ruleAction.Deserialize(tmpArray);
                    bufferReader = new BufferReader(tmpArray);
                    value = bufferReader.ReadBytes(length);
                    break;

                // PtypBinary  
                case 0x0102:
                    length = (uint)(buffer[bufferReader.Position] + (buffer[bufferReader.Position + 1] << 8) + 2);
                    value = bufferReader.ReadBytes(length);
                    break;

                // PtypMultipleInteger16 
                case 0x1002:
                    length = (uint)(buffer[bufferReader.Position] + (buffer[bufferReader.Position + 1] << 8));
                    length = (length * 2) + 2;
                    value = bufferReader.ReadBytes(length);
                    break;

                // PtypMultipleInteger32  
                case 0x1003:

                // PtypMultipleFloating32  
                case 0x1004:
                    length = (uint)(buffer[bufferReader.Position] + (buffer[bufferReader.Position + 1] << 8));
                    length = (length * 4) + 2;
                    value = bufferReader.ReadBytes(length);
                    break;

                // PtypMultipleFloating64 
                case 0x1005:

                // PtypMultipleCurrency
                case 0x1006:

                // PtypMultipleFloatingTime
                case 0x1007:

                // PtypMultipleInteger64
                case 0x1014:

                // PtypMultipleTime 
                case 0x1040:
                    length = (uint)(buffer[bufferReader.Position] + (buffer[bufferReader.Position + 1] << 8));
                    length = (length * 8) + 2;
                    value = bufferReader.ReadBytes(length);
                    break;

                // PtypMultipleString
                case 0x101F:
                    startIndex = bufferReader.Position;
                    ushort strCount = bufferReader.ReadUInt16();
                    for (int istr = 0; istr < strCount; istr++)
                    {
                        bufferReader.ReadUnicodeString();
                    }

                    endIndex = bufferReader.Position;
                    length = endIndex - startIndex;
                    value = bufferReader.ReadBytes(startIndex, length);
                    break;

                // PtypMultipleString8
                case 0x101E:
                    startIndex = bufferReader.Position;
                    ushort str8Count = bufferReader.ReadUInt16();
                    for (int istr8 = 0; istr8 < str8Count; istr8++)
                    {
                        bufferReader.ReadASCIIString();
                    }

                    endIndex = bufferReader.Position;
                    length = endIndex - startIndex;
                    value = bufferReader.ReadBytes(startIndex, length);
                    break;

                // PtypMultipleGuid
                case 0x1048:
                    length = (uint)(buffer[bufferReader.Position] + (buffer[bufferReader.Position + 1] << 8));
                    length = (length * 16) + 2;
                    value = bufferReader.ReadBytes(length);
                    break;

                // PtypMultipleBinary 
                case 0x1102:
                    startIndex = bufferReader.Position;
                    ushort binCount = bufferReader.ReadUInt16();
                    for (int ibin = 0; ibin < binCount; ibin++)
                    {
                        uint binLength = bufferReader.ReadUInt16();
                        bufferReader.ReadBytes(binLength);
                    }

                    endIndex = bufferReader.Position;
                    length = endIndex - startIndex;
                    value = bufferReader.ReadBytes(startIndex, length);
                    break;

                // PtypUnspecified 
                case 0x0000:
                    throw new NotImplementedException();

                // PtypNull
                case 0x0001:
                    value = null;
                    break;

                // PtypObject or PtypEmbeddedTable 
                case 0x000D:
                    throw new NotImplementedException();

                default:
                    throw new NotImplementedException();
            }

            return value;
        }

        /// <summary>
        /// Add ruleCondition to the properties list
        /// </summary>
        /// <param name="list">The properties list of a rule</param>
        /// <param name="mailSubjectContent">The mail subject of the content restriction</param>
        private static void AddRuleCondition(IList<TaggedPropertyValue> list, string mailSubjectContent)
        {
            TaggedPropertyValue pidTagRuleCondition = new TaggedPropertyValue();
            PropertyTag pidTagRuleConditionPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)TestPropertyName.PidTagRuleCondition,
                PropertyType = (ushort)PropertyTypeName.PtypRestriction
            };
            pidTagRuleCondition.PropertyTag = pidTagRuleConditionPropertyTag;

            TaggedPropertyValue taggedProperty = new TaggedPropertyValue();
            PropertyTag taggedPropertyPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)TestPropertyName.PidTagSubject,
                PropertyType = (ushort)PropertyTypeName.PtypString
            };
            taggedProperty.PropertyTag = taggedPropertyPropertyTag;
            taggedProperty.Value = Encoding.Unicode.GetBytes(mailSubjectContent + "\0");

            pidTagRuleCondition.Value = GenerateContentRestriction(taggedProperty);
            list.Add(pidTagRuleCondition);
        }

        /// <summary>
        /// Add rule action to the properties list
        /// </summary>
        /// <param name="list">The properties list for a rule action</param>
        /// <param name="actionType">The action Type</param>
        /// <param name="actionData">The actionData</param>
        /// <param name="actionFlavor">The actionFlavor of the action</param>
        /// <param name="actionFlag">The action flag value.</param>
        private static void AddRuleAction(IList<TaggedPropertyValue> list, ActionTypes actionType, IActionData actionData, uint actionFlavor, uint actionFlag)
        {
            TaggedPropertyValue pidTagRuleActions = new TaggedPropertyValue();
            PropertyTag pidTagRuleActionsPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)TestPropertyName.PidTagRuleActions,
                PropertyType = (ushort)PropertyTypeName.PtypRuleAction
            };
            pidTagRuleActions.PropertyTag = pidTagRuleActionsPropertyTag;
            switch (actionType)
            {
                case ActionTypes.OP_MARK_AS_READ:
                case ActionTypes.OP_DELETE:
                    pidTagRuleActions.Value = GenerateRuleAction(actionType, Count.TwoBytesCount, new DeleteMarkReadActionData(), actionFlavor, actionFlag).Serialize();
                    break;
                case ActionTypes.OP_MOVE:
                case ActionTypes.OP_COPY:
                case ActionTypes.OP_FORWARD:
                case ActionTypes.OP_DEFER_ACTION:
                case ActionTypes.OP_TAG:
                case ActionTypes.OP_DELEGATE:
                case ActionTypes.OP_BOUNCE:
                case ActionTypes.OP_OOF_REPLY:
                case ActionTypes.OP_REPLY:
                    pidTagRuleActions.Value = GenerateRuleAction(actionType, Count.TwoBytesCount, actionData, actionFlavor, actionFlag).Serialize();
                    break;
                default:
                    Site.Assert.Fail("The actionType is {0}, this is not a valid action type", actionType);
                    break;
            }

            list.Add(pidTagRuleActions);
        }

        /// <summary>
        /// Generate a sample rule action.
        /// </summary>
        /// <param name="actionType">The action Type</param>
        /// <param name="bytesCount">The length of the bytes count</param>
        /// <param name="actionDataBufferValue">The actionData buffer</param>
        /// <param name="actionFlavor">Action flavor value.</param>
        /// <param name="actionFlag">Action flag value.</param>
        /// <returns>An instance of the RuleAction</returns>
        private static RuleAction GenerateRuleAction(ActionTypes actionType, Count bytesCount, IActionData actionDataBufferValue, uint actionFlavor, uint actionFlag)
        {
            ActionBlock actionBlock = new ActionBlock(bytesCount)
            {
                ActionType = actionType,
                ActionFlags = actionFlag,
                ActionDataValue = actionDataBufferValue,
                ActionFlavor = actionFlavor
            };

            actionBlock.ActionLength = (bytesCount == Count.TwoBytesCount) ? (actionBlock.Size() - 2) : (actionBlock.Size() - 4);
            RuleAction ruleAction = new RuleAction(bytesCount)
            {
                NoOfActions = 0x01,
                Actions = new ActionBlock[1]
                {
                    actionBlock
                }
            };

            // Only one rule action is generated.
            return ruleAction;
        }
    }
}