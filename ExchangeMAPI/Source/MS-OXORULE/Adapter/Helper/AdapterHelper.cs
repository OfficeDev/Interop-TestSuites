namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using System.Runtime.InteropServices;
    using System.Net;
    using System.Collections;
    
    /// <summary>
    /// Help method collection.
    /// </summary>
    public class AdapterHelper
    {
        /// <summary>
        /// The current session context cookies for the request.
        /// </summary>
        private static CookieCollection sessionContextCookies;

        /// <summary>
        /// The transport used by the test suite.
        /// </summary>
        private static string transport;


        /// <summary>
        /// Gets or sets the current session context cookies for the request.
        /// </summary>
        public static CookieCollection SessionContextCookies
        {
            get
            {
                if (sessionContextCookies == null)
                {
                    sessionContextCookies = new CookieCollection();
                }

                return sessionContextCookies;
            }

            set
            {
                sessionContextCookies = value;
            }
        }

        /// <summary>
        /// Gets or sets the transport used by the test suite.
        /// </summary>
        public static string Transport
        {
            get { return AdapterHelper.transport; }
            set { AdapterHelper.transport = value; }
        }

        /// <summary>
        /// Get a TaggedPropertyValue structure from buffer.
        /// </summary>
        /// <param name="buffer">Buffer contain TaggedPropertyValue instance.</param>
        /// <returns>A TaggedPropertyvalue structure.</returns>
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
        /// <param name="type">The Type of value.</param>
        /// <param name="buffer">Buffer contains value.</param>
        /// <returns>Byte array of the value.</returns>
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
                    IRestriction restriction = null;
                    RestrictionType restrictionType = (RestrictionType)tmpArray[0];
                    switch (restrictionType)
                    {
                        case RestrictionType.AndRestriction:
                            restriction = new AndRestriction();
                            break;
                        case RestrictionType.BitMaskRestriction:
                            restriction = new BitMaskRestriction();
                            break;
                        case RestrictionType.CommentRestriction:
                            restriction = new CommentRestriction();
                            break;
                        case RestrictionType.ComparePropertiesRestriction:
                            restriction = new ComparePropertiesRestriction();
                            break;
                        case RestrictionType.ContentRestriction:
                            restriction = new ContentRestriction();
                            break;
                        case RestrictionType.CountRestriction:
                            restriction = new CountRestriction();
                            break;
                        case RestrictionType.ExistRestriction:
                            restriction = new ExistRestriction();
                            break;
                        case RestrictionType.NotRestriction:
                            restriction = new NotRestriction();
                            break;
                        case RestrictionType.OrRestriction:
                            restriction = new OrRestriction();
                            break;
                        case RestrictionType.PropertyRestriction:
                            restriction = new PropertyRestriction();
                            break;
                        case RestrictionType.SizeRestriction:
                            restriction = new SizeRestriction();
                            break;
                        case RestrictionType.SubObjectRestriction:
                            restriction = new SubObjectRestriction();
                            break;
                    }

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
            }

            return value;
        }

        /// <summary>
        /// This method converts a string to a byte array, the string is consist of hex16 values which are separated by ",".
        /// </summary>
        /// <param name="str">A string to be converted.</param>
        /// <returns>Return the converted byte array.</returns>
        public static byte[] ConvertStringToBytes(string str)
        {
            char[] charSeparators = { ',' };
            string[] strArray = str.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);
            byte[] bytes = new byte[strArray.Length];
            for (int i = 0; i < strArray.Length; i++)
            {
                bytes[i] = byte.Parse(strArray[i].Trim().Substring(2), System.Globalization.NumberStyles.HexNumber);
            }

            return bytes;
        }

        /// <summary>
        /// Generate test data for creating extended rule.
        /// </summary>
        /// <param name="rulename">The rule name.</param>
        /// <param name="ruleSequence">The rule sequence.</param>
        /// <param name="ruleState">The rule state.</param>
        /// <param name="provider">The rule provider.</param>
        /// <param name="actionType">The rule action Type.</param>
        /// <param name="actionData">The rule action data.</param>
        /// <param name="contentRestrictSubjectName">The subject name of the rule content restriction.</param>
        /// <param name="namedPropertyInfo">The namedPropertyInfo that needed for construct the rule data.</param>
        /// <returns>An array of TaggedPropertyValue of an extended rule.</returns>
        public static TaggedPropertyValue[] GenerateExtendedRuleTestData(string rulename, int ruleSequence, uint ruleState, string provider, ActionType actionType, IActionData actionData, string contentRestrictSubjectName, NamedPropertyInfo namedPropertyInfo)
        {
            List<TaggedPropertyValue> propList = new List<TaggedPropertyValue>();
            TaggedPropertyValue pidTagRuleMessageName = new TaggedPropertyValue();
            PropertyTag pidTagRuleMessageNameTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleMessageName,
                PropertyType = (ushort)PropertyType.PtypString
            };
            pidTagRuleMessageName.PropertyTag = pidTagRuleMessageNameTag;
            pidTagRuleMessageName.Value = Encoding.Unicode.GetBytes(rulename + "\0");
            propList.Add(pidTagRuleMessageName);

            TaggedPropertyValue pidTagMessageClass = new TaggedPropertyValue();
            PropertyTag pidTagMessageClassTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            pidTagMessageClass.PropertyTag = pidTagMessageClassTag;
            pidTagMessageClass.Value = Encoding.Unicode.GetBytes(Constants.ExtendedRuleMessageClass + "\0");
            propList.Add(pidTagMessageClass);

            TaggedPropertyValue pidTagRuleMessageSequence = new TaggedPropertyValue();
            PropertyTag pidTagRuleMessageSequencePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleMessageSequence,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            pidTagRuleMessageSequence.PropertyTag = pidTagRuleMessageSequencePropertyTag;
            pidTagRuleMessageSequence.Value = BitConverter.GetBytes(ruleSequence);
            propList.Add(pidTagRuleMessageSequence);

            TaggedPropertyValue pidTagRuleMessageState = new TaggedPropertyValue();
            PropertyTag pidTagRuleMessageStatePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleMessageState,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            pidTagRuleMessageState.PropertyTag = pidTagRuleMessageStatePropertyTag;
            pidTagRuleMessageState.Value = BitConverter.GetBytes(ruleState);
            propList.Add(pidTagRuleMessageState);

            TaggedPropertyValue pidTagRuleMessageLevel = new TaggedPropertyValue();
            PropertyTag pidTagRuleMessageLevelPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleMessageLevel,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            pidTagRuleMessageLevel.PropertyTag = pidTagRuleMessageLevelPropertyTag;
            pidTagRuleMessageLevel.Value = BitConverter.GetBytes(Constants.ExtendedRuleMessageLevel);
            propList.Add(pidTagRuleMessageLevel);

            TaggedPropertyValue pidTagRuleMessageProvider = new TaggedPropertyValue();
            PropertyTag pidTagRuleMessageProviderPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleMessageProvider,
                PropertyType = (ushort)PropertyType.PtypString
            };
            pidTagRuleMessageProvider.PropertyTag = pidTagRuleMessageProviderPropertyTag;
            pidTagRuleMessageProvider.Value = Encoding.Unicode.GetBytes(provider + "\0");
            propList.Add(pidTagRuleMessageProvider);

            ExtendedRuleActions extendedRuleActions = new ExtendedRuleActions
            {
                NamedPropertyInformation = namedPropertyInfo
            };

            extendedRuleActions.RuleVersion = Constants.ExtendedRuleVersion;
            extendedRuleActions.RuleActionBuffer = GetRuleAction(actionType, CountByte.FourBytesCount, actionData, Constants.CommonActionFlavor, Constants.RuleActionFlags);

            TaggedPropertyValue pidTagExtendedRuleMessageActions = new TaggedPropertyValue();
            PropertyTag pidTagExtendedRuleMessageActionsPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagExtendedRuleMessageActions,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            pidTagExtendedRuleMessageActions.PropertyTag = pidTagExtendedRuleMessageActionsPropertyTag;
            pidTagExtendedRuleMessageActions.Value = Common.AddInt16LengthBeforeBinaryArray(extendedRuleActions.Serialize());
            propList.Add(pidTagExtendedRuleMessageActions);

            TaggedPropertyValue pidTagExtendedRuleMessageCondition = new TaggedPropertyValue();
            PropertyTag pidTagExtendedRuleMessageConditionPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagExtendedRuleMessageCondition,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            pidTagExtendedRuleMessageCondition.PropertyTag = pidTagExtendedRuleMessageConditionPropertyTag;

            TaggedPropertyValue taggedProperty = new TaggedPropertyValue();
            PropertyTag taggedPropertyPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagSubject,
                PropertyType = (ushort)PropertyType.PtypString
            };
            taggedProperty.PropertyTag = taggedPropertyPropertyTag;
            taggedProperty.Value = Encoding.Unicode.GetBytes(contentRestrictSubjectName + "\0");
            ContentRestriction contentRestriction = new ContentRestriction
            {
                FuzzyLevelLow = ContentRestriction.FuzzyLevelLowValue.FL_SUBSTRING,
                FuzzyLevelHigh = ContentRestriction.FuzzyLevelHighValue.FL_IGNORECASE,
                PropertyTag = taggedProperty.PropertyTag,
                TaggedValue = taggedProperty
            };

            ExtendedRuleCondition extendedRuleCondition = new ExtendedRuleCondition
            {
                NamedPropertyInformation = namedPropertyInfo,
                RuleRestriction = contentRestriction
            };
            pidTagExtendedRuleMessageCondition.Value = Common.AddInt16LengthBeforeBinaryArray(extendedRuleCondition.Serialize());
            propList.Add(pidTagExtendedRuleMessageCondition);
            return propList.ToArray();
        }

        /// <summary>
        /// Generate RuleData from ptf config base on different action Type.
        /// </summary>
        /// <param name="actionType">ActionType in a rule.</param>
        /// <param name="ruleDataType">The type of the rule data.</param>
        /// <param name="ruleSequence">Rule sequence in PidTagRuleSequence of the RuleData.</param>
        /// <param name="ruleState">Specifies value of PidTagRuleState of the RuleData.</param>
        /// <param name="actionData">The actionData needed to construct the RuleData.</param>
        /// <param name="otherInfo">
        /// An array of Other needed information of Actions.
        /// Index may be from 0-4
        /// 0. Rule name
        /// 1. Rule provider
        /// 2. Rule user Flags
        /// 3. Rule provider data
        /// 4. Subject name of content restriction
        /// </param>
        /// <param name="ruleId">A 64-bit unsigned integer represents the rule.</param>
        /// <returns>A valid RuleData structure.</returns>
        public static RuleData GenerateValidRuleData(ActionType actionType, TestRuleDataType ruleDataType, int ruleSequence, RuleState ruleState, IActionData actionData, RuleProperties otherInfo, ulong? ruleId)
        {
            RuleData ruleData = new RuleData
            {
                RuleDataFlags = (byte)ruleDataType
            };

            List<TaggedPropertyValue> propertyList = new List<TaggedPropertyValue>();
            if (ruleDataType == TestRuleDataType.ForModify || ruleDataType == TestRuleDataType.ForRemove)
            {
                TaggedPropertyValue pidTagRuleId = new TaggedPropertyValue();
                PropertyTag pidTagRuleIdPropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleId,
                    PropertyType = (ushort)PropertyType.PtypInteger64
                };
                pidTagRuleId.PropertyTag = pidTagRuleIdPropertyTag;
                pidTagRuleId.Value = BitConverter.GetBytes((ulong)ruleId);
                propertyList.Add(pidTagRuleId);
            }

            if (ruleDataType != TestRuleDataType.ForRemove)
            {
                TaggedPropertyValue pidTagRuleSequence = new TaggedPropertyValue();
                PropertyTag pidTagRuleSequencePropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleSequence,
                    PropertyType = (ushort)PropertyType.PtypInteger32
                };
                pidTagRuleSequence.PropertyTag = pidTagRuleSequencePropertyTag;
                pidTagRuleSequence.Value = BitConverter.GetBytes(ruleSequence);
                propertyList.Add(pidTagRuleSequence);

                TaggedPropertyValue pidTagRuleLevel = new TaggedPropertyValue();
                PropertyTag pidTagRuleLevelPropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleLevel,
                    PropertyType = (ushort)PropertyType.PtypInteger32
                };
                pidTagRuleLevel.PropertyTag = pidTagRuleLevelPropertyTag;
                pidTagRuleLevel.Value = BitConverter.GetBytes(0);
                propertyList.Add(pidTagRuleLevel);

                TaggedPropertyValue pidTagRuleState = new TaggedPropertyValue();
                PropertyTag pidTagRuleStatePropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleState,
                    PropertyType = (ushort)PropertyType.PtypInteger32
                };
                pidTagRuleState.PropertyTag = pidTagRuleStatePropertyTag;
                pidTagRuleState.Value = BitConverter.GetBytes((uint)ruleState);
                propertyList.Add(pidTagRuleState);

                if (otherInfo.Name.Length > 0)
                {
                    TaggedPropertyValue pidTagRuleName = new TaggedPropertyValue();
                    PropertyTag pidTagRuleNamePropertyTag = new PropertyTag
                    {
                        PropertyId = (ushort)PropertyId.PidTagRuleName,
                        PropertyType = (ushort)PropertyType.PtypString
                    };
                    pidTagRuleName.PropertyTag = pidTagRuleNamePropertyTag;
                    pidTagRuleName.Value = Encoding.Unicode.GetBytes((string)otherInfo.Name + "\0");
                    propertyList.Add(pidTagRuleName);
                }

                if (otherInfo.Provider.Length > 0)
                {
                    TaggedPropertyValue pidTagRuleProvider = new TaggedPropertyValue();
                    PropertyTag pidTagRuleProviderPropertyTag = new PropertyTag
                    {
                        PropertyId = (ushort)PropertyId.PidTagRuleProvider,
                        PropertyType = (ushort)PropertyType.PtypString
                    };
                    pidTagRuleProvider.PropertyTag = pidTagRuleProviderPropertyTag;
                    pidTagRuleProvider.Value = Encoding.Unicode.GetBytes((string)otherInfo.Provider + "\0");
                    propertyList.Add(pidTagRuleProvider);
                }

                if (otherInfo.UserFlag.Length > 0)
                {
                    TaggedPropertyValue pidTagRuleUserFlags = new TaggedPropertyValue();
                    PropertyTag pidTagRuleUserFlagsPropertyTag = new PropertyTag
                    {
                        PropertyId = (ushort)PropertyId.PidTagRuleUserFlags,
                        PropertyType = (ushort)PropertyType.PtypInteger32
                    };
                    pidTagRuleUserFlags.PropertyTag = pidTagRuleUserFlagsPropertyTag;
                    pidTagRuleUserFlags.Value = BitConverter.GetBytes(int.Parse(otherInfo.UserFlag));
                    propertyList.Add(pidTagRuleUserFlags);
                }

                if (otherInfo.ProviderData.Length > 0)
                {
                    TaggedPropertyValue pidTagRuleProviderData = new TaggedPropertyValue();
                    PropertyTag pidTagRuleProviderDataPropertyTag = new PropertyTag
                    {
                        PropertyId = (ushort)PropertyId.PidTagRuleProviderData,
                        PropertyType = (ushort)PropertyType.PtypBinary
                    };
                    pidTagRuleProviderData.PropertyTag = pidTagRuleProviderDataPropertyTag;
                    pidTagRuleProviderData.Value = Common.GetBytesFromBinaryHexString(otherInfo.ProviderData);
                    pidTagRuleProviderData.Value = Common.AddInt16LengthBeforeBinaryArray(pidTagRuleProviderData.Value);
                    propertyList.Add(pidTagRuleProviderData);
                }

                if (otherInfo.ConditionSubjectName.Length > 0)
                {
                    AddRuleCondition(propertyList, otherInfo.ConditionSubjectName);
                    AddRuleAction(propertyList, actionType, actionData, Constants.CommonActionFlavor, Constants.RuleActionFlags);
                }
            }

            ruleData.PropertyValues = propertyList.ToArray();
            ruleData.PropertyValueCount = (ushort)propertyList.Count;
            return ruleData;
        }

        /// <summary>
        /// Generate RuleData from ptf config base on different action Type.
        /// </summary>
        /// <param name="actionType">ActionType in a rule.</param>
        /// <param name="ruleSequence">Rule sequence in PidTagRuleSequence of the RuleData.</param>
        /// <param name="ruleState">Specifies value of PidTagRuleState of the RuleData.</param>
        /// <param name="actionData">The actionData needed to construct the RuleData.</param>
        /// <param name="actionFlavor">The actionFlavor of the rule action.</param>
        /// <param name="otherInfo">
        /// An array of Other needed information of Actions.
        /// Index may be from 0-4
        /// 0. Rule name
        /// 1. Rule provider
        /// 2. Rule user Flags
        /// 3. Rule provider data
        /// 4. Subject name of content restriction
        /// </param>
        /// <returns>A valid RuleData structure.</returns>
        public static RuleData GenerateValidRuleDataWithFlavor(ActionType actionType, int ruleSequence, RuleState ruleState, IActionData actionData, uint actionFlavor, RuleProperties otherInfo)
        {
            RuleData ruleData = new RuleData
            {
                RuleDataFlags = (byte)TestRuleDataType.ForAdd
            };

            List<TaggedPropertyValue> propertyList = new List<TaggedPropertyValue>();
            TaggedPropertyValue pidTagRuleSequence = new TaggedPropertyValue();
            PropertyTag pidTagRuleSequencePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleSequence,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            pidTagRuleSequence.PropertyTag = pidTagRuleSequencePropertyTag;
            pidTagRuleSequence.Value = BitConverter.GetBytes(ruleSequence);
            propertyList.Add(pidTagRuleSequence);

            TaggedPropertyValue pidTagRuleLevel = new TaggedPropertyValue();
            PropertyTag pidTagRuleLevelPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleLevel,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            pidTagRuleLevel.PropertyTag = pidTagRuleLevelPropertyTag;
            pidTagRuleLevel.Value = BitConverter.GetBytes(Constants.RuleLevel);
            propertyList.Add(pidTagRuleLevel);

            TaggedPropertyValue pidTagRuleState = new TaggedPropertyValue();
            PropertyTag pidTagRuleStatePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleState,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            pidTagRuleState.PropertyTag = pidTagRuleStatePropertyTag;
            pidTagRuleState.Value = BitConverter.GetBytes((uint)ruleState);
            propertyList.Add(pidTagRuleState);

            if (otherInfo.Name.Length > 0)
            {
                TaggedPropertyValue pidTagRuleName = new TaggedPropertyValue();
                PropertyTag pidTagRuleNamePropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleName,
                    PropertyType = (ushort)PropertyType.PtypString
                };
                pidTagRuleName.PropertyTag = pidTagRuleNamePropertyTag;
                pidTagRuleName.Value = Encoding.Unicode.GetBytes((string)otherInfo.Name + "\0");
                propertyList.Add(pidTagRuleName);
            }

            if (otherInfo.Provider.Length > 0)
            {
                TaggedPropertyValue pidTagRuleProvider = new TaggedPropertyValue();
                PropertyTag pidTagRuleProviderPropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleProvider,
                    PropertyType = (ushort)PropertyType.PtypString
                };
                pidTagRuleProvider.PropertyTag = pidTagRuleProviderPropertyTag;
                pidTagRuleProvider.Value = Encoding.Unicode.GetBytes((string)otherInfo.Provider + "\0");
                propertyList.Add(pidTagRuleProvider);
            }

            if (otherInfo.UserFlag.Length > 0)
            {
                TaggedPropertyValue pidTagRuleUserFlags = new TaggedPropertyValue();
                PropertyTag pidTagRuleUserFlagsPropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleUserFlags,
                    PropertyType = (ushort)PropertyType.PtypInteger32
                };
                pidTagRuleUserFlags.PropertyTag = pidTagRuleUserFlagsPropertyTag;
                pidTagRuleUserFlags.Value = BitConverter.GetBytes(int.Parse(otherInfo.UserFlag));
                propertyList.Add(pidTagRuleUserFlags);
            }

            if (otherInfo.ProviderData.Length > 0)
            {
                TaggedPropertyValue pidTagRuleProviderData = new TaggedPropertyValue();
                PropertyTag pidTagRuleProviderDataPropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleProviderData,
                    PropertyType = (ushort)PropertyType.PtypBinary
                };
                pidTagRuleProviderData.PropertyTag = pidTagRuleProviderDataPropertyTag;
                pidTagRuleProviderData.Value = Common.GetBytesFromBinaryHexString(otherInfo.ProviderData);
                pidTagRuleProviderData.Value = Common.AddInt16LengthBeforeBinaryArray(pidTagRuleProviderData.Value);
                propertyList.Add(pidTagRuleProviderData);
            }

            if (otherInfo.ConditionSubjectName.Length > 0)
            {
                AddRuleCondition(propertyList, otherInfo.ConditionSubjectName);
                AddRuleAction(propertyList, actionType, actionData, actionFlavor, Constants.RuleActionFlags);
            }

            ruleData.PropertyValues = propertyList.ToArray();
            ruleData.PropertyValueCount = (ushort)propertyList.Count;
            return ruleData;
        }

        /// <summary>
        /// Generate RuleData from ptf config base on different action Type.
        /// </summary>
        /// <param name="actionType">ActionType of each blocks, the sequence must be the same as the actionData array.</param>
        /// <param name="ruleSequence">Rule sequence in PidTagRuleSequence of the RuleData.</param>
        /// <param name="ruleState">Specifies value of PidTagRuleState of the RuleData.</param>
        /// <param name="actionData">The actionData array needed to construct the ActionBlock in RuleActions.</param>
        /// <param name="actionFlavor">The actionFlavor array of the rule ActionBlocks.</param>
        /// <param name="otherInfo">
        /// An array of Other needed information of Actions.
        /// Index may be from 0-4
        /// 0. Rule name
        /// 1. Rule provider
        /// 2. Rule user Flags
        /// 3. Rule provider data
        /// 4. Subject name of content restriction
        /// </param>
        /// <returns>A valid RuleData structure.</returns>
        public static RuleData GenerateValidRuleDataWithFlavor(ActionType[] actionType, int ruleSequence, RuleState ruleState, IActionData[] actionData, uint[] actionFlavor, RuleProperties otherInfo)
        {
            RuleData ruleData = new RuleData
            {
                RuleDataFlags = (byte)TestRuleDataType.ForAdd
            };

            List<TaggedPropertyValue> propertyList = new List<TaggedPropertyValue>();
            TaggedPropertyValue pidTagRuleSequence = new TaggedPropertyValue();
            PropertyTag pidTagRuleSequencePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleSequence,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            pidTagRuleSequence.PropertyTag = pidTagRuleSequencePropertyTag;
            pidTagRuleSequence.Value = BitConverter.GetBytes(ruleSequence);
            propertyList.Add(pidTagRuleSequence);

            TaggedPropertyValue pidTagRuleLevel = new TaggedPropertyValue();
            PropertyTag pidTagRuleLevelPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleLevel,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            pidTagRuleLevel.PropertyTag = pidTagRuleLevelPropertyTag;
            pidTagRuleLevel.Value = BitConverter.GetBytes(Constants.RuleLevel);
            propertyList.Add(pidTagRuleLevel);

            TaggedPropertyValue pidTagRuleState = new TaggedPropertyValue();
            PropertyTag pidTagRuleStatePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleState,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            pidTagRuleState.PropertyTag = pidTagRuleStatePropertyTag;
            pidTagRuleState.Value = BitConverter.GetBytes((uint)ruleState);
            propertyList.Add(pidTagRuleState);

            if (otherInfo.Name.Length > 0)
            {
                TaggedPropertyValue pidTagRuleName = new TaggedPropertyValue();
                PropertyTag pidTagRuleNamePropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleName,
                    PropertyType = (ushort)PropertyType.PtypString
                };
                pidTagRuleName.PropertyTag = pidTagRuleNamePropertyTag;
                pidTagRuleName.Value = Encoding.Unicode.GetBytes((string)otherInfo.Name + "\0");
                propertyList.Add(pidTagRuleName);
            }

            if (otherInfo.Provider.Length > 0)
            {
                TaggedPropertyValue pidTagRuleProvider = new TaggedPropertyValue();
                PropertyTag pidTagRuleProviderPropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleProvider,
                    PropertyType = (ushort)PropertyType.PtypString
                };
                pidTagRuleProvider.PropertyTag = pidTagRuleProviderPropertyTag;
                pidTagRuleProvider.Value = Encoding.Unicode.GetBytes((string)otherInfo.Provider + "\0");
                propertyList.Add(pidTagRuleProvider);
            }

            if (otherInfo.UserFlag.Length > 0)
            {
                TaggedPropertyValue pidTagRuleUserFlags = new TaggedPropertyValue();
                PropertyTag pidTagRuleUserFlagsPropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleUserFlags,
                    PropertyType = (ushort)PropertyType.PtypInteger32
                };
                pidTagRuleUserFlags.PropertyTag = pidTagRuleUserFlagsPropertyTag;
                pidTagRuleUserFlags.Value = BitConverter.GetBytes(int.Parse(otherInfo.UserFlag));
                propertyList.Add(pidTagRuleUserFlags);
            }

            if (otherInfo.ProviderData.Length > 0)
            {
                TaggedPropertyValue pidTagRuleProviderData = new TaggedPropertyValue();
                PropertyTag pidTagRuleProviderDataPropertyTag = new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleProviderData,
                    PropertyType = (ushort)PropertyType.PtypBinary
                };
                pidTagRuleProviderData.PropertyTag = pidTagRuleProviderDataPropertyTag;
                pidTagRuleProviderData.Value = Common.GetBytesFromBinaryHexString(otherInfo.ProviderData);
                pidTagRuleProviderData.Value = Common.AddInt16LengthBeforeBinaryArray(pidTagRuleProviderData.Value);
                propertyList.Add(pidTagRuleProviderData);
            }

            if (otherInfo.ConditionSubjectName.Length > 0)
            {
                AddRuleCondition(propertyList, otherInfo.ConditionSubjectName);

                uint[] defaultActionFlags = new uint[actionData.Length];
                for (int i = 0; i < actionData.Length; i++)
                {
                    defaultActionFlags[i] = Constants.RuleActionFlags;
                }

                AddRuleAction(propertyList, actionType, actionData, actionFlavor, defaultActionFlags);
            }

            ruleData.PropertyValues = propertyList.ToArray();
            ruleData.PropertyValueCount = (ushort)propertyList.Count;
            return ruleData;
        }

        /// <summary>
        /// Build a RestrictionData as single ContentRestriction.
        /// </summary>
        /// <param name="taggedPValue">A TaggedPropertyValue of this restriction.</param>
        /// <returns>Restriction Data.</returns>
        public static byte[] GenerateContentRestriction(TaggedPropertyValue taggedPValue)
        {
            ContentRestriction result = new ContentRestriction
            {
                FuzzyLevelLow = ContentRestriction.FuzzyLevelLowValue.FL_SUBSTRING,
                FuzzyLevelHigh = ContentRestriction.FuzzyLevelHighValue.FL_IGNORECASE,
                PropertyTag = taggedPValue.PropertyTag,
                TaggedValue = taggedPValue
            };

            return result.Serialize();
        }

        /// <summary>
        /// This method help to convert the property value, which is of variable bytes, to RuleAction structure.
        /// </summary>
        /// <param name="byteArray">The byte array to be converted.</param>
        /// <returns>Return the RuleAction structure.</returns>
        public static RuleAction PropertyValueConvertToRuleAction(byte[] byteArray)
        {
            // The first 2 bytes of byteArray only indicates the total number of subsequent bytes,
            // byteArrayTobeConverted is the actual bytes used to convert to the RuleAction structure,
            // which should not include the first 2 bytes of byteArray.
            byte[] byteArrayToBeConverted = new byte[byteArray.Length - 2];
            Array.Copy(byteArray, 2, byteArrayToBeConverted, 0, byteArray.Length - 2);

            // Deserialize the byte array to the RuleAction structure.
            RuleAction ruleAction = new RuleAction();
            ruleAction.Deserialize(byteArrayToBeConverted);

            return ruleAction;
        }

        /// <summary>
        /// This method help to convert the property value, which is of variable bytes, to string value.
        /// </summary>
        /// <param name="byteArray">The byte array to be converted.</param>
        /// <returns>Return the string value.</returns>
        public static string PropertyValueConvertToString(byte[] byteArray)
        {
            // Convert the byte array to the Unicode string value.
            string value = Encoding.Unicode.GetString(byteArray);

            // The actual string value should not contain the last character '\0'.
            // So here get the substring which doesn't include the last character '\0'.
            value = value.Substring(0, value.Length - 1);
            return value;
        }

        /// <summary>
        /// This method help to convert the property value, which is of 1 byte, to Boolean value.
        /// </summary>
        /// <param name="byteArray">The byte array to be converted.</param>
        /// <returns>Return the Boolean value.</returns>
        public static bool PropertyValueConvertToBool(byte[] byteArray)
        {
            bool value = BitConverter.ToBoolean(byteArray, 0);
            return value;
        }

        /// <summary>
        /// This method helps to convert the property value, which is of 4 bytes, to unsigned integer value.
        /// </summary>
        /// <param name="byteArray">The byte array to be converted.</param>
        /// <returns>Return the unsigned integer value.</returns>
        public static uint PropertyValueConvertToUint(byte[] byteArray)
        {
            uint value = (uint)(byteArray[3] << 24) +
                         (uint)(byteArray[2] << 16) +
                         (uint)(byteArray[1] << 8) +
                         (uint)byteArray[0];
            return value;
        }

        /// <summary>
        /// This method helps to convert the property value, which is of 8 bytes, to unsigned long value.
        /// </summary>
        /// <param name="byteArray">The byte array to be converted.</param>
        /// <returns>Return the unsigned long value.</returns>
        public static ulong PropertyValueConvertToUint64(byte[] byteArray)
        {
            ulong value = 0;
            for (int i = byteArray.Length - 1; i >= 0; i--)
            {
                ulong temp = (ulong)byteArray[i];
                value += temp << (i * 8);
            }

            return value;
        }

        /// <summary>
        /// This method help to convert the property value, which is of variable bytes, to binary value.
        /// </summary>
        /// <param name="byteArray">The byte array to be converted.</param>
        /// <returns>Return the byte array value.</returns>
        public static byte[] PropertyValueConvertToBinary(byte[] byteArray)
        {
            // The first 2 bytes of byteArray only indicates the total length of subsequent bytes,
            // value is used to indicate the actual binary value in the Property Value returned,
            // which should not include the first 2 bytes of byteArray.
            byte[] value = new byte[byteArray.Length - 2];
            Array.Copy(byteArray, 2, value, 0, byteArray.Length - 2);

            return value;
        }

        /// <summary>
        /// This method help to convert the property value, which is of variable bytes, to ExtendedRuleActions structure.
        /// </summary>
        /// <param name="byteArray">The byte array to be converted.</param>
        /// <returns>Return the ExtendedRuleActions structure.</returns>
        public static ExtendedRuleActions PropertyValueConvertToExtendedRuleActions(byte[] byteArray)
        {
            // The first 2 bytes of byteArray only indicates the total number of subsequent bytes,
            // byteArrayTobeConvert is the actual bytes used to convert to the ExtendedRuleActions structure,
            // which should not include the first 2 bytes of byteArray.
            byte[] byteArrayTobeConvert = new byte[byteArray.Length - 2];
            Array.Copy(byteArray, 2, byteArrayTobeConvert, 0, byteArray.Length - 2);

            // De-serialize the byte array into the ExtendedRuleActions structure.
            ExtendedRuleActions extendedRuleActions = new ExtendedRuleActions();
            extendedRuleActions.Deserialize(byteArrayTobeConvert);

            return extendedRuleActions;
        }

        /// <summary>
        /// Generate PropertyTag arrays for rule properties of DAM.
        /// </summary>
        /// <returns>PropertyTag arrays for rule properties of DAM</returns>
        public static PropertyTag[] GenerateRuleInfoPropertiesOfDAM()
        {
            PropertyTag[] propertyTagOfDAM = new PropertyTag[0x09];

            // PidTagMessageClass
            propertyTagOfDAM[0].PropertyId = (ushort)PropertyId.PidTagMessageClass;
            propertyTagOfDAM[0].PropertyType = (ushort)PropertyType.PtypString;

            // PidTagDamBackPatched
            propertyTagOfDAM[1].PropertyId = (ushort)PropertyId.PidTagDamBackPatched;
            propertyTagOfDAM[1].PropertyType = (ushort)PropertyType.PtypBoolean;

            // PidTagDamOriginalEntryId
            propertyTagOfDAM[2].PropertyId = (ushort)PropertyId.PidTagDamOriginalEntryId;
            propertyTagOfDAM[2].PropertyType = (ushort)PropertyType.PtypBinary;

            // PidTagRuleProvider
            propertyTagOfDAM[3].PropertyId = (ushort)PropertyId.PidTagRuleProvider;
            propertyTagOfDAM[3].PropertyType = (ushort)PropertyType.PtypString;

            // PidTagRuleFolderEntryId
            propertyTagOfDAM[4].PropertyId = (ushort)PropertyId.PidTagRuleFolderEntryId;
            propertyTagOfDAM[4].PropertyType = (ushort)PropertyType.PtypBinary;

            // PidTagClientActions
            propertyTagOfDAM[5].PropertyId = (ushort)PropertyId.PidTagClientActions;
            propertyTagOfDAM[5].PropertyType = (ushort)PropertyType.PtypBinary;

            // PidTagRuleIds
            propertyTagOfDAM[6].PropertyId = (ushort)PropertyId.PidTagRuleIds;
            propertyTagOfDAM[6].PropertyType = (ushort)PropertyType.PtypBinary;

            // PidTagDeferredActionMessageOriginalEntryId
            propertyTagOfDAM[7].PropertyId = (ushort)PropertyId.PidTagDeferredActionMessageOriginalEntryId;
            propertyTagOfDAM[7].PropertyType = (ushort)PropertyType.PtypServerId;

            // PidTagMID
            propertyTagOfDAM[8].PropertyId = (ushort)PropertyId.PidTagMid;
            propertyTagOfDAM[8].PropertyType = (ushort)PropertyType.PtypInteger64;

            return propertyTagOfDAM;
        }

        /// <summary>
        /// Generate property value on recipient.
        /// </summary>
        /// <param name="userName">Recipient user name.</param>
        /// <param name="userDN">Recipient user dn.</param>
        /// <returns>Property array value.</returns>
        public static TaggedPropertyValue[] GenerateRecipientPropertiesBlock(string userName, string userDN)
        {
            TaggedPropertyValue[] recipientProperties = new TaggedPropertyValue[4];

            // Add PidTagDisplayName
            recipientProperties[0] = new TaggedPropertyValue();

            PropertyTag pidTagDisplayNamePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagDisplayName,
                PropertyType = (ushort)PropertyType.PtypString
            };
            recipientProperties[0].PropertyTag = pidTagDisplayNamePropertyTag;
            recipientProperties[0].Value = Encoding.Unicode.GetBytes(userName + "\0");

            // Add PidTagEmailAddress
            recipientProperties[1] = new TaggedPropertyValue();
            PropertyTag pidTagEmailAddressPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagEmailAddress,
                PropertyType = (ushort)PropertyType.PtypString
            };
            recipientProperties[1].PropertyTag = pidTagEmailAddressPropertyTag;
            recipientProperties[1].Value = Encoding.Unicode.GetBytes(userDN + "\0");

            // Add PidTagRecipientType
            recipientProperties[2] = new TaggedPropertyValue();
            PropertyTag pidTagRecipientTypePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRecipientType,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            recipientProperties[2].PropertyTag = pidTagRecipientTypePropertyTag;
            recipientProperties[2].Value = BitConverter.GetBytes(0x00000001);

            AddressBookEntryID addressBookEntryID = new AddressBookEntryID(userDN);

            // Add PidTagEntryID
            recipientProperties[3] = new TaggedPropertyValue();
            PropertyTag pidTagEntryIDPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagEntryId,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            recipientProperties[3].PropertyTag = pidTagEntryIDPropertyTag;
            recipientProperties[3].Value = Common.AddInt16LengthBeforeBinaryArray(addressBookEntryID.Serialize());

            return recipientProperties;
        }

        /// <summary>
        /// Serialize recipient properties.
        /// </summary>
        /// <returns>Serialized properties value.</returns>
        public static uint[] SerializeRecipientProperties()
        {
            uint[] aulPropTag = new uint[5];

            // PidTagEntryId
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagEntryId,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            aulPropTag[0] = Common.ConvertByteArrayToUint(propertyTag.Serialize());

            // PidTagAddressType
            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagAddressType,
                PropertyType = (ushort)PropertyType.PtypString
            };
            aulPropTag[1] = Common.ConvertByteArrayToUint(propertyTag.Serialize());

            // PidTagEmailAddress
            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagEmailAddress,
                PropertyType = (ushort)PropertyType.PtypString
            };
            aulPropTag[2] = Common.ConvertByteArrayToUint(propertyTag.Serialize());

            // PidTagDisplayName
            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagDisplayName,
                PropertyType = (ushort)PropertyType.PtypString
            };
            aulPropTag[3] = Common.ConvertByteArrayToUint(propertyTag.Serialize());

            // PidTagSearchKey
            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagSearchKey,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            aulPropTag[4] = Common.ConvertByteArrayToUint(propertyTag.Serialize());
            return aulPropTag;
        }

        /// <summary>
        /// Generate rule properties value.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        /// <param name="name">Rule name.</param>
        /// <returns>Rule properties.</returns>
        public static RuleProperties GenerateRuleProperties(ITestSite site, string name)
        {
            RuleProperties ruleProperties = new RuleProperties();
            switch (name)
            {
                case Constants.RuleNameCopy:
                case Constants.RuleNameDelete:
                case Constants.RuleNameMarkAsRead:
                case Constants.RuleNameMoveOne:
                case Constants.RuleNameMoveTwo:
                case Constants.RuleNameForward:
                case Constants.RuleNameDeferredAction:
                case Constants.RuleNameBounce:
                case Constants.RuleNameTag:
                case Constants.RuleNameDelegate:
                case Constants.RuleNameReply:
                case Constants.RuleNameForwardAT:
                case Constants.RuleNameForwardTM:
                case Constants.RuleNameOOFReply:
                    ruleProperties.UserFlag = Constants.PidTagRuleUserFlags1;
                    ruleProperties.Provider = Constants.PidTagRuleProvider;
                    ruleProperties.ProviderData = Constants.PidTagRuleProviderData;
                    ruleProperties.ConditionSubjectName = Constants.RuleConditionSubjectContainString + DateTime.Now.ToString("HHmmss");
                    break;

                case Constants.DAMPidTagRuleNameOne:
                case Constants.DAMPidTagRuleNameTwo:
                    ruleProperties.Provider = Constants.DAMPidTagRuleProviderOne;
                    ruleProperties.ConditionSubjectName = Constants.DAMPidTagRuleConditionSubjectContainString + DateTime.Now.ToString("HHmmss");
                    ruleProperties.UserFlag = Constants.DAMPidTagRuleUserFlags;
                    ruleProperties.ProviderData = Constants.DAMPidTagRuleProviderData;
                    break;

                case Constants.DEMRule:
                    ruleProperties.ConditionSubjectName = Constants.DEMPidTagRuleConditionSubjectContainString + DateTime.Now.ToString("HHmmss");
                    ruleProperties.UserFlag = Constants.DEMPidTagRuleUserFlags;
                    ruleProperties.ProviderData = Constants.DEMPidTagRuleProviderData;
                    ruleProperties.Provider = Constants.DEMPidTagRuleProvider;
                    break;

                default:
                    break;
            }

            ruleProperties.Name = Common.GenerateResourceName(site, name);
            return ruleProperties;
        }

        #region CaptureCode Helper method
        /// <summary>
        /// This method helps to convert the property value, which is of 4 bytes, to int value.
        /// </summary>
        /// <param name="byteArray">The byte array to be converted.</param>
        /// <returns>Return the integer value.</returns>
        public static int PropertyValueConvertToInt(byte[] byteArray)
        {
            int value = (byteArray[3] << 24) +
                        (byteArray[2] << 16) +
                        (byteArray[1] << 8) +
                        byteArray[0];
            return value;
        }

        /// <summary>
        /// This method help to convert the property value, which is of variable bytes, to ExtendedRuleCondition structure.
        /// </summary>
        /// <param name="byteArray">The byte array to be converted.</param>
        /// <returns>Return the ExtendedRuleCondition structure.</returns>
        public static ExtendedRuleCondition PropertyValueConvertToExtendedRuleCondition(byte[] byteArray)
        {
            // The first 2 bytes of byteArray only indicates the total number of subsequent bytes,
            // byteArrayTobeConvert is the actual bytes used to convert to the ExtendedRuleCondition structure,
            // which should not include the first 2 bytes of byteArray.
            byte[] byteArrayTobeConvert = new byte[byteArray.Length - 2];
            Array.Copy(byteArray, 2, byteArrayTobeConvert, 0, byteArray.Length - 2);

            // Deserialize the byte array value into the ExtendedRuleCondition structure.
            ExtendedRuleCondition extendedRuleCondition = new ExtendedRuleCondition();
            extendedRuleCondition.Deserialize(byteArrayTobeConvert);

            return extendedRuleCondition;
        }

        /// <summary>
        /// This method is used to determine whether there are two elements of the same value in a uint array.
        /// </summary>
        /// <param name="uintArray">The unsigned integer array to be evaluated.</param>
        /// <returns>The return value is true means there has the same value in this array, false means each element in this 
        /// array is not the same.</returns>
        public static bool IsContainSameValue(uint[] uintArray)
        {
            for (int i = 0; i < uintArray.Length; i++)
            {
                for (int j = i + 1; j < uintArray.Length; j++)
                {
                    if (uintArray[i] == uintArray[j])
                    {
                        return true;
                    }
                }
            }

            return false;
        }
        #endregion

        /// <summary>
        /// Add rule action to the properties list.
        /// </summary>
        /// <param name="list">The properties list for a rule action.</param>
        /// <param name="actionType">The action Type.</param>
        /// <param name="actionData">The actionData.</param>
        /// <param name="actionFlavor">The actionFlavor of the action.</param>
        /// <param name="actionFlags">The action flag value.</param>
        private static void AddRuleAction(IList<TaggedPropertyValue> list, ActionType actionType, IActionData actionData, uint actionFlavor, uint actionFlags)
        {
            TaggedPropertyValue pidTagRuleActions = new TaggedPropertyValue();
            PropertyTag pidTagRuleActionsPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleActions,
                PropertyType = (ushort)PropertyType.PtypRuleAction
            };
            pidTagRuleActions.PropertyTag = pidTagRuleActionsPropertyTag;
            pidTagRuleActions.Value = GetRuleAction(actionType, CountByte.TwoBytesCount, actionData, actionFlavor, actionFlags).Serialize();
            list.Add(pidTagRuleActions);
        }

        /// <summary>
        /// Add rule action to the properties list.
        /// </summary>
        /// <param name="list">The properties list for a rule action.</param>
        /// <param name="actionTypes">The action Type.</param>
        /// <param name="actionData">The actionData.</param>
        /// <param name="actionFlavor">The actionFlavor of the action.</param>
        /// <param name="actionFlags">The action flag value.</param>
        private static void AddRuleAction(IList<TaggedPropertyValue> list, ActionType[] actionTypes, IActionData[] actionData, uint[] actionFlavor, uint[] actionFlags)
        {
            TaggedPropertyValue pidTagRuleActions = new TaggedPropertyValue();
            PropertyTag pidTagRuleActionsPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleActions,
                PropertyType = (ushort)PropertyType.PtypRuleAction
            };
            pidTagRuleActions.PropertyTag = pidTagRuleActionsPropertyTag;
            pidTagRuleActions.Value = GetRuleAction(actionTypes, CountByte.TwoBytesCount, actionData, actionFlavor, actionFlags).Serialize();
            list.Add(pidTagRuleActions);
        }

        /// <summary>
        /// Generate a rule action.
        /// </summary>
        /// <param name="actionType">The action Type.</param>
        /// <param name="countBytes">The length of the bytes count.</param>
        /// <param name="actionDataBufferValue">The actionData buffer.</param>
        /// <param name="actionFlavor">Action flavor value.</param>
        /// <param name="actionFlags">Action flag value.</param>
        /// <returns>An instance of the RuleAction.</returns>
        private static RuleAction GetRuleAction(ActionType actionType, CountByte countBytes, IActionData actionDataBufferValue, uint actionFlavor, uint actionFlags)
        {
            ActionBlock actionBlock = new ActionBlock(countBytes)
            {
                ActionType = actionType,
                ActionFlags = actionFlags,
                ActionDataValue = actionDataBufferValue,
                ActionFlavor = actionFlavor
            };

            // Get actionBlock size
            int lengthOfActionLength = 0;
            if (actionBlock.CountType == CountByte.TwoBytesCount)
            {
                lengthOfActionLength += 2;
            }
            else if (actionBlock.CountType == CountByte.FourBytesCount)
            {
                lengthOfActionLength += 4;
            }

            // Length of ActionType is 1
            // Length of ActionFlavor is 4
            // Length of ActionFlags is 4
            int size = lengthOfActionLength + 1 + 4 + 4 + actionBlock.ActionDataValue.Size();
            actionBlock.ActionLength = (countBytes == CountByte.TwoBytesCount) ? (size - 2) : (size - 4);
            RuleAction ruleAction = new RuleAction(countBytes)
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

        /// <summary>
        /// Generate a rule action.
        /// </summary>
        /// <param name="actionTypes">The action Type.</param>
        /// <param name="countBytes">The length of the bytes count.</param>
        /// <param name="actionDataBufferValue">The actionData buffer.</param>
        /// <param name="actionFlavor">Action flavor value.</param>
        /// <param name="actionFlags">Action flag value.</param>
        /// <returns>An instance of the RuleAction.</returns>
        private static RuleAction GetRuleAction(ActionType[] actionTypes, CountByte countBytes, IActionData[] actionDataBufferValue, uint[] actionFlavor, uint[] actionFlags)
        {
            ActionBlock[] actionBlocks = new ActionBlock[actionTypes.Length];

            for (int i = 0; i < actionDataBufferValue.Length; i++)
            {
                ActionBlock actionBlock = new ActionBlock(countBytes)
                {
                    ActionType = actionTypes[i],
                    ActionFlags = actionFlags[i],
                    ActionDataValue = actionDataBufferValue[i],
                    ActionFlavor = actionFlavor[i]
                };

                // Get actionBlock size
                int lengthOfActionLength = 0;
                if (actionBlock.CountType == CountByte.TwoBytesCount)
                {
                    lengthOfActionLength += 2;
                }
                else if (actionBlock.CountType == CountByte.FourBytesCount)
                {
                    lengthOfActionLength += 4;
                }

                // Length of ActionType is 1
                // Length of ActionFlavor is 4
                // Length of ActionFlags is 4
                int size = lengthOfActionLength + 1 + 4 + 4 + actionBlock.ActionDataValue.Size();
                actionBlock.ActionLength = (countBytes == CountByte.TwoBytesCount) ? (size - 2) : (size - 4);
                actionBlocks[i] = actionBlock;
            }

            RuleAction ruleAction = new RuleAction(countBytes)
            {
                NoOfActions = actionBlocks.Length,
                Actions = actionBlocks
            };
            return ruleAction;
        }

        /// <summary>
        /// Add ruleCondition to the properties list.
        /// </summary>
        /// <param name="list">The properties list of a rule.</param>
        /// <param name="mailSubjectContent">The mail subject of the content restriction.</param>
        private static void AddRuleCondition(IList<TaggedPropertyValue> list, string mailSubjectContent)
        {
            TaggedPropertyValue pidTagRuleCondition = new TaggedPropertyValue();
            PropertyTag pidTagRuleConditionPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleCondition,
                PropertyType = (ushort)PropertyType.PtypRestriction
            };
            pidTagRuleCondition.PropertyTag = pidTagRuleConditionPropertyTag;

            TaggedPropertyValue taggedProperty = new TaggedPropertyValue();
            PropertyTag taggedPropertyPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagSubject,
                PropertyType = (ushort)PropertyType.PtypString
            };
            taggedProperty.PropertyTag = taggedPropertyPropertyTag;
            taggedProperty.Value = Encoding.Unicode.GetBytes(mailSubjectContent + "\0");

            pidTagRuleCondition.Value = GenerateContentRestriction(taggedProperty);
            list.Add(pidTagRuleCondition);
        }

        /// <summary>
        /// Allocate memory for stat instance.
        /// </summary>
        /// <param name="stat">The stat instance.</param>
        /// <returns>The pointer points to the allocated memory.</returns>
        public static IntPtr AllocStat(STAT stat)
        {
            IntPtr ptr = Marshal.AllocHGlobal(Marshal.SizeOf(stat));
            int offset = 0;
            Marshal.WriteInt32(ptr, offset, (int)stat.SortType);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.ContainerID);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.CurrentRec);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, stat.Delta);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.NumPos);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.TotalRecs);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.CodePage);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.TemplateLocale);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.SortLocale);

            return ptr;
        }

        /// <summary>
        /// Allocate memory for the specific property values.
        /// </summary>
        /// <param name="pta_r">PropertyTagArray_r instance.</param>
        /// <returns>A pointer points to the allocated memory.</returns>
        public static IntPtr AllocPropertyTagArray_r(PropertyTagArray_r pta_r)
        {
            int offset = 0;
            int cb = (int)(sizeof(uint) + (pta_r.Values * sizeof(uint)));

            IntPtr ptr = Marshal.AllocHGlobal(cb);

            Marshal.WriteInt32(ptr, offset, (int)pta_r.Values);
            offset += sizeof(uint);

            for (int i = 0; i < pta_r.Values; i++)
            {
                Marshal.WriteInt32(ptr, offset, (int)pta_r.AulPropTag[i]);
                offset += sizeof(uint);
            }

            return ptr;
        }

        /// <summary>
        /// Parse EphemeralEntryID structure from byte array.
        /// </summary>
        /// <param name="bytes">The byte array to be parsed.</param>
        /// <returns>An EphemeralEntryID structure instance.</returns>
        public static EphemeralEntryID ParseEphemeralEntryIDFromBytes(byte[] bytes)
        {
            int index = 0;

            EphemeralEntryID entryID = new EphemeralEntryID
            {
                IDType = bytes[index++],
                R1 = bytes[index++],
                R2 = bytes[index++],
                R3 = bytes[index++],
                ProviderUID = new FlatUID_r
                {
                    Ab = new byte[Constants.FlatUIDByteSize]
                }
            };
            for (int i = 0; i < Constants.FlatUIDByteSize; i++)
            {
                entryID.ProviderUID.Ab[i] = bytes[index++];
            }

            // R4: 4 bytes
            entryID.R4 = (uint)BitConverter.ToInt32(bytes, index);
            index += 4;

            // DisplayType: 4 bytes
            entryID.DisplayType = (DisplayTypeValue)BitConverter.ToInt32(bytes, index);
            index += 4;

            // Mid: 4 bytes
            entryID.Mid = (uint)BitConverter.ToInt32(bytes, index);
            index += 4;

            return entryID;
        }

        /// <summary>
        /// Parse PermanentEntryID structure from byte array.
        /// </summary>
        /// <param name="bytes">The byte array to be parsed.</param>
        /// <returns>An PermanentEntryID structure instance.</returns>
        public static PermanentEntryID ParsePermanentEntryIDFromBytes(byte[] bytes)
        {
            int index = 0;

            PermanentEntryID entryID = new PermanentEntryID
            {
                IDType = bytes[index++],
                R1 = bytes[index++],
                R2 = bytes[index++],
                R3 = bytes[index++],
                ProviderUID = new FlatUID_r
                {
                    Ab = new byte[Constants.FlatUIDByteSize]
                }
            };
            for (int i = 0; i < Constants.FlatUIDByteSize; i++)
            {
                entryID.ProviderUID.Ab[i] = bytes[index++];
            }

            // R4: 4 bytes
            entryID.R4 = (uint)BitConverter.ToInt32(bytes, index);
            index += 4;

            // DisplayType: 4 bytes
            entryID.DisplayTypeString = (DisplayTypeValue)BitConverter.ToInt32(bytes, index);
            index += 4;

            // DistinguishedName: variable 
            entryID.DistinguishedName = System.Text.Encoding.Default.GetString(bytes, index, bytes.Length - index - 1);
            return entryID;
        }

        /// <summary>
        /// Parse a STAT structure instance from pointer.
        /// </summary>
        /// <param name="ptr">Pointer points to the memory.</param>
        /// <returns>A STAT structure.</returns>
        public static STAT ParseStat(IntPtr ptr)
        {
            STAT stat = new STAT();
            int offset = 0;

            stat.SortType = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.ContainerID = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.CurrentRec = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.Delta = Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.NumPos = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.TotalRecs = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.CodePage = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.TemplateLocale = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.SortLocale = (uint)Marshal.ReadInt32(ptr, offset);

            return stat;
        }

        /// <summary>
        /// Parse Binary_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of Binary_r structure.</returns>
        public static Binary_r ParseBinary_r(IntPtr ptr)
        {
            Binary_r b_r = new Binary_r
            {
                Cb = (uint)Marshal.ReadInt32(ptr)
            };

            if (b_r.Cb == 0)
            {
                b_r.Lpb = null;
            }
            else
            {
                b_r.Lpb = new byte[b_r.Cb];
                if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                {
                    IntPtr baddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                    for (uint i = 0; i < b_r.Cb; i++)
                    {
                        b_r.Lpb[i] = Marshal.ReadByte(baddr, (int)i);
                    }
                }
                else
                {
                    for (uint i = 0; i < b_r.Cb; i++)
                    {
                        b_r.Lpb[i] = Marshal.ReadByte(ptr, (int)i + sizeof(uint));
                    }
                }
            }

            return b_r;
        }

        /// <summary>
        /// Parse GUIDs.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of GUIDs.</returns>
        public static FlatUID_r ParseFlatUID_r(IntPtr ptr)
        {
            FlatUID_r fuid_r = new FlatUID_r
            {
                Ab = new byte[Constants.FlatUIDByteSize]
            };
            for (int i = 0; i < Constants.FlatUIDByteSize; i++)
            {
                fuid_r.Ab[i] = Marshal.ReadByte(ptr, i);
            }

            return fuid_r;
        }

        /// <summary>
        /// Parse ShortArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of ShortArray_r structure.</returns>
        public static ShortArray_r ParseShortArray_r(IntPtr ptr)
        {
            ShortArray_r shortArray = new ShortArray_r
            {
                CValues = (uint)Marshal.ReadInt32(ptr)
            };

            if (shortArray.CValues == 0)
            {
                shortArray.Lpi = null;
            }
            else
            {
                if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                {
                    IntPtr saaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                    ptr = saaddr;
                }

                shortArray.Lpi = new short[shortArray.CValues];
                int offset = 0;
                for (uint i = 0; i < shortArray.CValues; i++, offset += sizeof(short))
                {
                    shortArray.Lpi[i] = Marshal.ReadInt16(ptr, offset);
                }
            }

            return shortArray;
        }

        /// <summary>
        /// Parse LongArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of LongArray_r structure.</returns>
        public static LongArray_r ParseLongArray_r(IntPtr ptr)
        {
            LongArray_r longArray = new LongArray_r
            {
                CValues = (uint)Marshal.ReadInt32(ptr)
            };

            if (longArray.CValues == 0)
            {
                longArray.Lpl = null;
            }
            else
            {
                if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                {
                    IntPtr laaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                    ptr = laaddr;
                }

                longArray.Lpl = new int[longArray.CValues];
                int offset = 0;
                for (uint i = 0; i < longArray.CValues; i++, offset += sizeof(int))
                {
                    longArray.Lpl[i] = Marshal.ReadInt32(ptr, offset);
                }
            }

            return longArray;
        }

        /// <summary>
        /// Parse String_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of String_r structure.</returns>
        public static byte[] ParseString_r(IntPtr ptr)
        {
            List<byte> stringArray = new List<byte>();
            if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
            {
                IntPtr szaPtr = new IntPtr(Marshal.ReadInt32(ptr));

                if (szaPtr == IntPtr.Zero)
                {
                    return null;
                }

                ptr = szaPtr;
            }

            int offsetOfszA = 0;
            byte curValueOfszA = 0;
            ArrayList listOfszA = new ArrayList();
            while (Marshal.ReadByte(ptr, offsetOfszA) != '\0')
            {
                curValueOfszA = Marshal.ReadByte(ptr, offsetOfszA);
                offsetOfszA++;
                listOfszA.Add(curValueOfszA);
            }

            for (int i = 0; i < listOfszA.Count; i++)
            {
                stringArray.Add(byte.Parse(listOfszA[i].ToString()));
            }

            return stringArray.ToArray();
        }

        /// <summary>
        /// Parse WString_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of WString_r structure.</returns>
        public static byte[] ParseWString_r(IntPtr ptr)
        {
            List<byte> stringArray = new List<byte>();

            if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
            {
                IntPtr szwPtr = new IntPtr(Marshal.ReadInt32(ptr));

                if (szwPtr == IntPtr.Zero)
                {
                    return null;
                }

                ptr = szwPtr;
            }

            int offsetOfszW = 0;
            byte curValueOfszW = 0;
            short shortValueOfszW = Marshal.ReadInt16(ptr, offsetOfszW);
            ArrayList listOfszW = new ArrayList();

            while (shortValueOfszW != '\0')
            {
                curValueOfszW = Marshal.ReadByte(ptr, offsetOfszW);
                offsetOfszW++;
                listOfszW.Add(curValueOfszW);

                if (offsetOfszW % 2 == 0)
                {
                    shortValueOfszW = Marshal.ReadInt16(ptr, offsetOfszW);
                }
            }

            for (int i = 0; i < listOfszW.Count; i++)
            {
                stringArray.Add(byte.Parse(listOfszW[i].ToString()));
            }

            return stringArray.ToArray();
        }

        /// <summary>
        /// Parse StringArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of StringArray_r structure.</returns>
        public static StringArray_r ParseStringArray_r(IntPtr ptr)
        {
            StringArray_r stringArray = new StringArray_r
            {
                CValues = (uint)Marshal.ReadInt32(ptr)
            };

            if (stringArray.CValues == 0)
            {
                stringArray.LppszA = null;
            }
            else
            {
                if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                {
                    IntPtr szaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                    stringArray.LppszA = new string[stringArray.CValues];
                    int offset = 0;
                    for (uint i = 0; i < stringArray.CValues; i++)
                    {
                        stringArray.LppszA[i] = Marshal.PtrToStringAnsi(new IntPtr(Marshal.ReadInt32(szaddr, offset)));
                        offset += 4;
                    }
                }
                else
                {
                    stringArray.LppszA = new string[stringArray.CValues];
                    for (uint i = 0; i < stringArray.CValues; i++)
                    {
                        stringArray.LppszA[i] = BitConverter.ToString(ParseString_r(ptr));
                    }
                }
            }

            return stringArray;
        }

        /// <summary>
        /// Allocate memory for GUIDs.
        /// </summary>
        /// <param name="fuid_r">Instance of GUIDs.</param>
        /// <returns>A pointer points to the allocated memory.</returns>
        public static IntPtr AllocFlatUID_r(FlatUID_r fuid_r)
        {
            IntPtr ptr = Marshal.AllocHGlobal(Constants.FlatUIDByteSize);

            for (int i = 0; i < Constants.FlatUIDByteSize; i++)
            {
                Marshal.WriteByte(ptr, i, fuid_r.Ab[i]);
            }

            return ptr;
        }

        /// <summary>
        /// Parse BinaryArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of BinaryArray_r structure.</returns>
        public static BinaryArray_r ParseBinaryArray_r(IntPtr ptr)
        {
            BinaryArray_r binaryArray = new BinaryArray_r
            {
                CValues = (uint)Marshal.ReadInt32(ptr)
            };

            if (binaryArray.CValues == 0)
            {
                binaryArray.Lpbin = null;
            }
            else
            {
                if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                {
                    IntPtr baaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                    binaryArray.Lpbin = new Binary_r[binaryArray.CValues];
                    for (uint i = 0; i < binaryArray.CValues; i++)
                    {
                        binaryArray.Lpbin[i] = ParseBinary_r(baaddr);
                        baaddr = new IntPtr(baaddr.ToInt32() + 8);
                    }
                }
                else
                {
                    binaryArray.Lpbin = new Binary_r[binaryArray.CValues];
                    for (uint i = 0; i < binaryArray.CValues; i++)
                    {
                        binaryArray.Lpbin[i] = ParseBinary_r(ptr);
                    }
                }
            }

            return binaryArray;
        }

        /// <summary>
        /// Parse WStringArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of WStringArray_r structure.</returns>
        public static WStringArray_r ParseWStringArray_r(IntPtr ptr)
        {
            WStringArray_r wsa_r = new WStringArray_r
            {
                CValues = (uint)Marshal.ReadInt32(ptr)
            };

            if (wsa_r.CValues == 0)
            {
                wsa_r.LppszW = null;
            }
            else
            {
                if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                {
                    IntPtr szwaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                    wsa_r.LppszW = new string[wsa_r.CValues];
                    for (uint i = 0; i < wsa_r.CValues; i++)
                    {
                        wsa_r.LppszW[i] = Marshal.PtrToStringUni(new IntPtr(Marshal.ReadInt32(szwaddr)));
                        szwaddr = new IntPtr(szwaddr.ToInt32() + 4);
                    }
                }
                else
                {
                    wsa_r.LppszW = new string[wsa_r.CValues];
                    for (uint i = 0; i < wsa_r.CValues; i++)
                    {
                        wsa_r.LppszW[i] = BitConverter.ToString(ParseWString_r(ptr));
                    }
                }
            }

            return wsa_r;
        }

        /// <summary>
        /// Parse FlatUIDArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of FlatUIDArray_r structure.</returns>
        public static FlatUIDArray_r ParseFlatUIDArray_r(IntPtr ptr)
        {
            FlatUIDArray_r fuida_r = new FlatUIDArray_r
            {
                CValues = (uint)Marshal.ReadInt32(ptr)
            };

            if (fuida_r.CValues == 0)
            {
                fuida_r.Lpguid = null;
            }
            else
            {
                fuida_r.Lpguid = new FlatUID_r[fuida_r.CValues];
                IntPtr fuidaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                for (uint i = 0; i < fuida_r.CValues; i++)
                {
                    fuida_r.Lpguid[i] = ParseFlatUID_r(new IntPtr(Marshal.ReadInt32(fuidaddr)));
                    fuidaddr = new IntPtr(fuidaddr.ToInt32() + 4);
                }
            }

            return fuida_r;
        }


        /// <summary>
        /// Parse PROP_VAL_UNION structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <param name="proptype">Property Types.</param>
        /// <returns>Instance of PROP_VAL_UNION structure.</returns>
        public static PROP_VAL_UNION ParsePROP_VAL_UNION(IntPtr ptr, PropertyTypeValue proptype)
        {
            PROP_VAL_UNION pvu = new PROP_VAL_UNION();

            switch (proptype)
            {
                case PropertyTypeValue.PtypInteger16:
                    pvu.I = Marshal.ReadInt16(ptr);
                    break;

                case PropertyTypeValue.PtypInteger32:
                    pvu.L = Marshal.ReadInt32(ptr);
                    break;

                case PropertyTypeValue.PtypBoolean:
                    if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                    {
                        pvu.B = (ushort)Marshal.ReadInt16(ptr);
                    }
                    else
                    {
                        pvu.B = (byte)Marshal.ReadByte(ptr);
                    }

                    break;

                case PropertyTypeValue.PtypString8:
                    pvu.LpszA = ParseString_r(ptr);
                    break;

                case PropertyTypeValue.PtypBinary:
                    pvu.Bin = ParseBinary_r(ptr);
                    break;

                case PropertyTypeValue.PtypString:
                    pvu.LpszW = ParseWString_r(ptr);
                    break;

                case PropertyTypeValue.PtypGuid:
                    if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                    {
                        IntPtr uidaddr = new IntPtr(Marshal.ReadInt32(ptr));

                        if (uidaddr == IntPtr.Zero)
                        {
                            pvu.Lpguid = null;
                        }
                        else
                        {
                            pvu.Lpguid = new FlatUID_r[1];
                            pvu.Lpguid[0] = ParseFlatUID_r(uidaddr);
                        }
                    }
                    else
                    {
                        if (ptr == IntPtr.Zero)
                        {
                            pvu.Lpguid = null;
                        }
                        else
                        {
                            pvu.Lpguid = new FlatUID_r[1];
                            pvu.Lpguid[0] = ParseFlatUID_r(ptr);
                        }
                    }

                    break;

                case PropertyTypeValue.PtypTime:
                    pvu.Ft.LowDateTime = (uint)Marshal.ReadInt32(ptr);
                    pvu.Ft.HighDateTime = (uint)Marshal.ReadInt32(ptr, sizeof(uint));
                    break;

                case PropertyTypeValue.PtypErrorCode:
                    pvu.Err = Marshal.ReadInt32(ptr);
                    break;

                case PropertyTypeValue.PtypMultipleInteger16:
                    pvu.MVi = ParseShortArray_r(ptr);
                    break;

                case PropertyTypeValue.PtypMultipleInteger32:
                    pvu.MVl = ParseLongArray_r(ptr);
                    break;

                case PropertyTypeValue.PtypMultipleString8:
                    uint isFound = (uint)Marshal.ReadInt32(ptr);

                    if (isFound == (uint)ErrorCodeValue.NotFound)
                    {
                        pvu.MVszA.CValues = 0;
                        pvu.MVszA.LppszA = null;
                    }
                    else
                    {
                        pvu.MVszA = ParseStringArray_r(ptr);
                    }

                    break;

                case PropertyTypeValue.PtypMultipleBinary:
                    pvu.MVbin = ParseBinaryArray_r(ptr);
                    break;

                case PropertyTypeValue.PtypMultipleString:
                    pvu.MVszW = ParseWStringArray_r(ptr);
                    break;

                case PropertyTypeValue.PtypMultipleGuid:
                    pvu.MVguid = ParseFlatUIDArray_r(ptr);
                    break;

                case PropertyTypeValue.PtypNull:
                case PropertyTypeValue.PtypEmbeddedTable:
                    pvu.Reserved = Marshal.ReadInt32(ptr);
                    break;

                default:
                    throw new ParseException("Parsing PROP_VAL_UNION failed!");
            }

            return pvu;
        }

        /// <summary>
        /// Parse PROP_VAL_UNION structure.
        /// </summary>
        /// <param name="propertyValue">The property value used for parsing.</param>
        /// <param name="proptype">The Property Types used for parsing.</param>
        /// <returns>Instance of PROP_VAL_UNION structure.</returns>
        public static PROP_VAL_UNION ParsePROP_VAL_UNION(AddressBookPropertyValue propertyValue, PropertyTypeValue proptype)
        {
            PROP_VAL_UNION pvu = new PROP_VAL_UNION();
            IntPtr valuePtr = IntPtr.Zero;
            valuePtr = Marshal.AllocHGlobal(propertyValue.Value.Length);
            Marshal.Copy(propertyValue.Value, 0, valuePtr, propertyValue.Value.Length);
            pvu = AdapterHelper.ParsePROP_VAL_UNION(valuePtr, proptype);
            Marshal.FreeHGlobal(valuePtr);
            return pvu;
        }

        /// <summary>
        /// Parse PropertyRowSet_r structure.
        /// </summary>
        /// <param name="columns">The columns which contain property tags.</param>
        /// <param name="rowCount">The row count of the PropertyRowSet_r.</param>
        /// <param name="rowData">The row data which contain the property values.</param>
        /// <returns>Instance of PropertyRowSet_r structure.</returns>
        public static PropertyRowSet_r ParsePropertyRowSet_r(LargePropTagArray columns, uint rowCount, AddressBookPropertyRow[] rowData)
        {
            PropertyRowSet_r propertyRowSet_r = new PropertyRowSet_r();

            propertyRowSet_r.Rows = rowCount;
            propertyRowSet_r.PropertyRowSet = new PropertyRow_r[rowCount];

            for (int i = 0; i < rowCount; i++)
            {
                propertyRowSet_r.PropertyRowSet[i].Reserved = 0;
                propertyRowSet_r.PropertyRowSet[i].Values = columns.PropertyTagCount;
                propertyRowSet_r.PropertyRowSet[i].Props = new PropertyValue_r[columns.PropertyTagCount];
                for (int j = 0; j < columns.PropertyTagCount; j++)
                {
                    propertyRowSet_r.PropertyRowSet[i].Props[j].PropTag = BitConverter.ToUInt32(columns.PropertyTags[j].Serialize(), 0);
                    propertyRowSet_r.PropertyRowSet[i].Props[j].Reserved = 0;
                    propertyRowSet_r.PropertyRowSet[i].Props[j].Value = AdapterHelper.ParsePROP_VAL_UNION(rowData[i].ValueArray.ToArray()[j], (PropertyTypeValue)columns.PropertyTags[j].PropertyType);
                }
            }

            return propertyRowSet_r;
        }

        /// <summary>
        /// Parse PropertyRowSet_r structure.
        /// </summary>
        /// <param name="rowsCount">The row count of the PropertyRowSet_r.</param>
        /// <param name="rows">The rows which contains property tags and property values.</param>
        /// <returns>Instance of PropertyRowSet_r structure.</returns>
        public static PropertyRowSet_r ParsePropertyRowSet_r(uint rowsCount, AddressBookPropValueList[] rows)
        {
            PropertyRowSet_r propertyRowSet_r = new PropertyRowSet_r();

            propertyRowSet_r.Rows = rowsCount;
            if (rowsCount == 0)
            {
                propertyRowSet_r.PropertyRowSet = null;
            }
            else
            {
                propertyRowSet_r.PropertyRowSet = new PropertyRow_r[rowsCount];

                for (int i = 0; i < rowsCount; i++)
                {
                    propertyRowSet_r.PropertyRowSet[i].Reserved = 0;
                    propertyRowSet_r.PropertyRowSet[i].Values = rows[i].PropertyValueCount;
                    propertyRowSet_r.PropertyRowSet[i].Props = new PropertyValue_r[rows[i].PropertyValueCount];
                    for (int j = 0; j < rows[i].PropertyValueCount; j++)
                    {
                        propertyRowSet_r.PropertyRowSet[i].Props[j].PropTag = BitConverter.ToUInt32(rows[i].PropertyValues[j].PropertyTag.Serialize(), 0);
                        propertyRowSet_r.PropertyRowSet[i].Props[j].Reserved = 0;
                        propertyRowSet_r.PropertyRowSet[i].Props[j].Value = AdapterHelper.ParsePROP_VAL_UNION((AddressBookPropertyValue)rows[i].PropertyValues[j], (PropertyTypeValue)rows[i].PropertyValues[j].PropertyTag.PropertyType);
                    }
                }
            }

            return propertyRowSet_r;
        }

        /// <summary>
        /// Parse the PropertyRow_r structure.
        /// </summary>
        /// <param name="row">The row which contains property tags and property values</param>
        /// <returns>Instance of the PropertyRow_r.</returns>
        public static PropertyRow_r ParsePropertyRow_r(AddressBookPropValueList row)
        {
            PropertyRow_r propertyRow_r = new PropertyRow_r();
            propertyRow_r.Reserved = 0;
            propertyRow_r.Values = row.PropertyValueCount;
            propertyRow_r.Props = new PropertyValue_r[row.PropertyValueCount];
            for (int i = 0; i < row.PropertyValueCount; i++)
            {
                propertyRow_r.Props[i].PropTag = BitConverter.ToUInt32(row.PropertyValues[i].PropertyTag.Serialize(), 0);
                propertyRow_r.Props[i].Reserved = 0;
                propertyRow_r.Props[i].Value = AdapterHelper.ParsePROP_VAL_UNION((AddressBookPropertyValue)row.PropertyValues[i], (PropertyTypeValue)row.PropertyValues[i].PropertyTag.PropertyType);
            }

            return propertyRow_r;
        }

        /// <summary>
        /// Parse PropertyValue_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of PropertyValue_r structure.</returns>
        public static PropertyValue_r ParsePropertyValue_r(IntPtr ptr)
        {
            PropertyValue_r protertyValue = new PropertyValue_r();

            int offset = 0;

            protertyValue.PropTag = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(uint);

            protertyValue.Reserved = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(uint);

            IntPtr newPtr = new IntPtr(ptr.ToInt32() + offset);
            protertyValue.Value = ParsePROP_VAL_UNION(newPtr, (PropertyTypeValue)(protertyValue.PropTag & 0x0000FFFF));

            return protertyValue;
        }

        /// <summary>
        /// Parse PropertyRow_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of PropertyRow_r structure.</returns>
        public static PropertyRow_r ParsePropertyRow_r(IntPtr ptr)
        {
            PropertyRow_r protertyRow = new PropertyRow_r();
            int offset = 0;

            protertyRow.Reserved = (uint)Marshal.ReadInt32(ptr);
            offset += sizeof(uint);

            protertyRow.Values = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(uint);

            if (protertyRow.Values == 0)
            {
                protertyRow.Props = null;
            }
            else
            {
                protertyRow.Props = new PropertyValue_r[protertyRow.Values];
                IntPtr pvaddr = new IntPtr(Marshal.ReadInt32(ptr, offset));

                const int PropertyValueLengthInBytes = 16;

                for (uint i = 0; i < protertyRow.Values; i++)
                {
                    protertyRow.Props[i] = ParsePropertyValue_r(pvaddr);
                    pvaddr = new IntPtr(pvaddr.ToInt32() + PropertyValueLengthInBytes);
                }
            }

            return protertyRow;
        }

        /// <summary>
        /// Parse PropertyRowSet_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of PropertyRowSet_r structure.</returns>
        public static PropertyRowSet_r ParsePropertyRowSet_r(IntPtr ptr)
        {
            PropertyRowSet_r prs_r = new PropertyRowSet_r();
            int offset = 0;

            prs_r.Rows = (uint)Marshal.ReadInt32(ptr);
            offset += sizeof(uint);

            const int PropertyRowLengthInBytes = 12;

            if (prs_r.Rows == 0)
            {
                prs_r.PropertyRowSet = null;
            }
            else
            {
                ptr = new IntPtr(ptr.ToInt32() + offset);
                prs_r.PropertyRowSet = new PropertyRow_r[prs_r.Rows];

                for (uint i = 0; i < prs_r.Rows; i++)
                {
                    prs_r.PropertyRowSet[i] = ParsePropertyRow_r(ptr);
                    ptr = new IntPtr(ptr.ToInt32() + PropertyRowLengthInBytes);
                }
            }

            return prs_r;
        }
    }
}