namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A driver for property initialization.
    /// </summary>
    public static class PropertyHelper
    {
        /// <summary>
        /// The dictionary for Property name and PropertyTag.
        /// </summary>
        private static Dictionary<PropertyNames, PropertyTag> propertyTagDic = new Dictionary<PropertyNames, PropertyTag>();

        /// <summary>
        /// The error code of response: access denied
        /// </summary>
        private static byte[] resultDeny = new byte[] { 0x05, 0x00, 0x07, 0x80 };

        /// <summary>
        /// The error code of response: not found
        /// </summary>
        private static byte[] resultNotFound = new byte[] { 0x0f, 0x01, 0x04, 0x80 };

        /// <summary>
        /// The error code of response: no value
        /// </summary>
        private static byte[] resultNullReference = new byte[] { 0xB9, 0x04, 0x00, 0x00 };

        /// <summary>
        /// Transfer ITestSite into adapter, make adapter can use ITestSite's function.
        /// </summary>
        private static ITestSite site;

        /// <summary>
        /// Gets the dictionary for Property name and PropertyTag.
        /// </summary>
        public static Dictionary<PropertyNames, PropertyTag> PropertyTagDic
        {
            get { return propertyTagDic; }
        }

        /// <summary>
        /// Initialize the AdapterHelper class.
        /// </summary>
        /// <param name="testSite">The instance of ITestSite.</param>
        public static void Initialize(ITestSite testSite)
        {
            site = testSite;
        }

        /// <summary>
        /// Check whether the return value in response is error.
        /// </summary>
        /// <param name="responseValue">The response value</param>
        /// <returns>If the response is error code, return true</returns>
        public static bool IsErrorCode(byte[] responseValue)
        {
            return Common.CompareByteArray(responseValue, resultDeny)
                || Common.CompareByteArray(responseValue, resultNotFound)
                || Common.CompareByteArray(responseValue, resultNullReference);
        }

        /// <summary>
        /// Check whether the return value in response is error.
        /// </summary>
        /// <param name="responseValue">The response value</param>
        /// <returns>If the response is error code, return true</returns>
        public static bool IsErrorCode(object responseValue)
        {
            return (responseValue is byte[]) && IsErrorCode((byte[])responseValue);
        }

        /// <summary>
        /// Initialize the dictionary for Property name and PropertyTag.
        /// </summary>
        public static void InitializePropertyTagDic()
        {
            AddPropertyTagToDic(PropertyNames.PidTagHasAttachments, new PropertyTag(0x0E1B, (ushort)PropertyType.PtypBoolean));
            AddPropertyTagToDic(PropertyNames.PidTagMessageClass, new PropertyTag(0x001A, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagMessageCodepage, new PropertyTag(0x3FFD, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagMessageLocaleId, new PropertyTag(0x3FF1, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagMessageSize, new PropertyTag(0x0E08, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagMessageStatus, new PropertyTag(0x0E17, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagSubjectPrefix, new PropertyTag(0x003D, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagNormalizedSubject, new PropertyTag(0x0E1D, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagImportance, new PropertyTag(0x0017, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagPriority, new PropertyTag(0x0026, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagSensitivity, new PropertyTag(0x0036, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidNameKeywords, new PropertyTag(0x0000, (ushort)PropertyType.PtypMultipleString));
            AddPropertyTagToDic(PropertyNames.PidTagAutoForwardComment, new PropertyTag(0x0004, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagBody, new PropertyTag(0x1000, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagNativeBody, new PropertyTag(0x1016, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagBodyHtml, new PropertyTag(0x1013, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagRtfCompressed, new PropertyTag(0x1009, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagRtfInSync, new PropertyTag(0x0E1F, (ushort)PropertyType.PtypBoolean));
            AddPropertyTagToDic(PropertyNames.PidTagInternetCodepage, new PropertyTag(0x3FDE, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagArchiveTag, new PropertyTag(0x3018, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagPolicyTag, new PropertyTag(0x3019, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagRetentionPeriod, new PropertyTag(0x301A, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagStartDateEtc, new PropertyTag(0x301B, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagRetentionDate, new PropertyTag(0x301C, (ushort)PropertyType.PtypTime));
            AddPropertyTagToDic(PropertyNames.PidTagRetentionFlags, new PropertyTag(0x301D, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagArchivePeriod, new PropertyTag(0x301E, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagArchiveDate, new PropertyTag(0x301F, (ushort)PropertyType.PtypTime));
            AddPropertyTagToDic(PropertyNames.PidTagLastModificationTime, new PropertyTag(0x3008, (ushort)PropertyType.PtypTime));
            AddPropertyTagToDic(PropertyNames.PidTagCreationTime, new PropertyTag(0x3007, (ushort)PropertyType.PtypTime));
            AddPropertyTagToDic(PropertyNames.PidTagDisplayName, new PropertyTag(0x3001, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagAttachSize, new PropertyTag(0x0E20, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagAttachNumber, new PropertyTag(0x0E21, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagAttachPathname, new PropertyTag(0x3708, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagAttachDataBinary, new PropertyTag(0x3701, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagAttachMethod, new PropertyTag(0x3705, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagAttachFilename, new PropertyTag(0x3704, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagAttachLongFilename, new PropertyTag(0x3707, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagAttachExtension, new PropertyTag(0x3703, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagAttachLongPathname, new PropertyTag(0x370D, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagAttachTag, new PropertyTag(0x370A, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagRenderingPosition, new PropertyTag(0X370B, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagAttachRendering, new PropertyTag(0x3709, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagAttachFlags, new PropertyTag(0x3714, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagAttachTransportName, new PropertyTag(0x370C, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagAttachEncoding, new PropertyTag(0x3702, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagAttachAdditionalInformation, new PropertyTag(0x370F, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagAttachmentLinkId, new PropertyTag(0x7FFA, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagAttachmentFlags, new PropertyTag(0x7FFD, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagAttachmentHidden, new PropertyTag(0x7FFE, (ushort)PropertyType.PtypBoolean));
            AddPropertyTagToDic(PropertyNames.PidTagAttachMimeTag, new PropertyTag(0x370E, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagAttachContentId, new PropertyTag(0x3712, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagAttachContentLocation, new PropertyTag(0x3713, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagAttachContentBase, new PropertyTag(0x3711, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagAttachDataObject, new PropertyTag(0x3701, (ushort)PropertyType.PtypComObject));
            AddPropertyTagToDic(PropertyNames.PidTagMessageFlags, new PropertyTag(0x0E07, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagAccessLevel, new PropertyTag(0x0FF7, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagDisplayBcc, new PropertyTag(0x0E02, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagDisplayCc, new PropertyTag(0x0E03, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagDisplayTo, new PropertyTag(0x0E04, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagSecurityDescriptor, new PropertyTag(0x0E27, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagUrlCompNameSet, new PropertyTag(0x0E62, (ushort)PropertyType.PtypBoolean));
            AddPropertyTagToDic(PropertyNames.PidTagTrustSender, new PropertyTag(0x0E79, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagUrlCompName, new PropertyTag(0x10F3, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagSearchKey, new PropertyTag(0x300B, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagAccess, new PropertyTag(0x0FF4, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagCreatorName, new PropertyTag(0x3FF8, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagCreatorEntryId, new PropertyTag(0x3FF9, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagLastModifierName, new PropertyTag(0x3FFA, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagLastModifierEntryId, new PropertyTag(0x3FFB, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagHasNamedProperties, new PropertyTag(0x664A, (ushort)PropertyType.PtypBoolean));
            AddPropertyTagToDic(PropertyNames.PidTagLocaleId, new PropertyTag(0x66A1, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagChangeKey, new PropertyTag(0x65E2, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagObjectType, new PropertyTag(0x0FFE, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagRecordKey, new PropertyTag(0x0FF9, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagLocalCommitTime, new PropertyTag(0x6709, (ushort)PropertyType.PtypTime));
            AddPropertyTagToDic(PropertyNames.PidTagAutoForwarded, new PropertyTag(0x0005, (ushort)PropertyType.PtypBoolean));
            AddPropertyTagToDic(PropertyNames.PidTagAddressBookDisplayNamePrintable, new PropertyTag(0x39FF, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagRowid, new PropertyTag(0x3000, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagRecipientOrder, new PropertyTag(0x5FDF, (ushort)PropertyType.PtypInteger32));
            AddPropertyTagToDic(PropertyNames.PidTagSubject, new PropertyTag(0x0037, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagInternetReferences, new PropertyTag(0x1039, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagMimeSkeleton, new PropertyTag(0x64F0, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagTnefCorrelationKey, new PropertyTag(0x007F, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagAlternateRecipientAllowed, new PropertyTag(0x0002, (ushort)PropertyType.PtypBoolean));
            AddPropertyTagToDic(PropertyNames.PidTagResponsibility, new PropertyTag(0x0E0F, (ushort)PropertyType.PtypBoolean));
            AddPropertyTagToDic(PropertyNames.PidTagPurportedSenderDomain, new PropertyTag(0x4083, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagStoreEntryId, new PropertyTag(0x0FFB, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagMessageRecipients, new PropertyTag(0x0E12, (ushort)PropertyType.PtypComObject));
            AddPropertyTagToDic(PropertyNames.PidTagBodyContentId, new PropertyTag(0x1015, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagBodyContentLocation, new PropertyTag(0x1014, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagHtml, new PropertyTag(0x1013, (ushort)PropertyType.PtypBinary));
            AddPropertyTagToDic(PropertyNames.PidTagTextAttachmentCharset, new PropertyTag(0x371B, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagAttachPayloadClass, new PropertyTag(0x371A, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagAttachPayloadProviderGuidString, new PropertyTag(0x3719, (ushort)PropertyType.PtypString));
            AddPropertyTagToDic(PropertyNames.PidTagRead, new PropertyTag(0x0E69, (ushort)PropertyType.PtypBoolean));
            AddPropertyTagToDic(PropertyNames.PidTagRecipientDisplayName, new PropertyTag(0x5FF6, (ushort)PropertyType.PtypString));
        }

        /// <summary>
        /// Find a property name from dictionary by property ID.
        /// </summary>
        /// <param name="propertyID">A unsigned integer value</param>
        /// <param name="propertyType">Property type code</param>
        /// <returns>Return Property Name value</returns>
        public static PropertyNames GetPropertyNameByID(uint propertyID, uint propertyType)
        {
            if (PropertyTagDic == null)
            {
                InitializePropertyTagDic();
            }

            PropertyNames keyTarget = PropertyNames.UserSpecified;

            foreach (PropertyNames propertyName in PropertyTagDic.Keys)
            {
                if ((uint)PropertyTagDic[propertyName].PropertyId == propertyID && (uint)PropertyTagDic[propertyName].PropertyType == propertyType)
                {
                    keyTarget = propertyName;
                    break;
                }
            }

            return keyTarget;
        }

        /// <summary>
        /// Override GetPropertyObjFromBuffer with PropertyTag list type.
        /// </summary>
        /// <param name="propertyTags">List of PropertyTag</param>
        /// <param name="response">RopGetPropertiesSpecificResponse packet</param>
        /// <returns>List of PropertyObj</returns>
        public static List<PropertyObj> GetPropertyObjFromBuffer(PropertyTag[] propertyTags, RopGetPropertiesSpecificResponse response)
        {
            List<PropertyObj> propertyList = new List<PropertyObj>();

            for (int i = 0; i < propertyTags.Length; i++)
            {
                PropertyNames propertyName = GetPropertyNameByID(propertyTags[i].PropertyId, propertyTags[i].PropertyType);
                propertyList.Add(GetPropertyObjFromBuffer(new PropertyObj(propertyName), response.RowData.PropertyValues[i].Value));
            }

            return propertyList;
        }

        /// <summary>
        /// Override GetPropertyObjFromBuffer without PropertyTag list type. Instead of that the properties will be collected from response.
        /// </summary>
        /// <param name="response">RopGetPropertiesAllResponse packet</param>
        /// <returns>List of PropertyObj</returns>
        public static List<PropertyObj> GetPropertyObjFromBuffer(RopGetPropertiesAllResponse response)
        {
            List<PropertyObj> propertyList = new List<PropertyObj>();

            foreach (TaggedPropertyValue taggedPropertyValue in response.PropertyValues)
            {
                if (GetPropertyNameByID(taggedPropertyValue.PropertyTag.PropertyId, taggedPropertyValue.PropertyTag.PropertyType) == PropertyNames.UserSpecified)
                {
                    continue;
                }

                PropertyObj property = new PropertyObj(taggedPropertyValue.PropertyTag.PropertyId, taggedPropertyValue.PropertyTag.PropertyType);

                propertyList.Add(GetPropertyObjFromBuffer(property, taggedPropertyValue.Value));
            }

            return propertyList;
        }

        /// <summary>
        /// The extended method for List of PropertyObj to find the exact one by property name.
        /// </summary>
        /// <param name="propertyList">List of PropertyObj</param>
        /// <param name="propertyName">A property name value</param>
        /// <returns>Return propertyObj</returns>
        public static PropertyObj GetPropertyByName(List<PropertyObj> propertyList, PropertyNames propertyName)
        {
            PropertyObj property = null;
            foreach (PropertyObj pt in propertyList)
            {
                if (pt.PropertyName == propertyName)
                {
                    property = pt;
                    break;
                }
            }

            return property;
        }

        /// <summary>
        /// Determine whether the property value is valid.
        /// </summary>
        /// <param name="property">The input PropertyObj.</param>
        /// <returns>Returns Boolean value that indicate whether property value is valid.</returns>
        public static bool IsPropertyValid(PropertyObj property)
        {
            if (property == null)
            {
                return false;
            }

            return !IsErrorCode(property.Value);
        }

        /// <summary>
        /// Returns a byte array which append the length of the Input as the first two bytes of input Byte array.
        /// </summary>
        /// <param name="bytes">The input Byte array</param>
        /// <returns>The Byte array corresponding to the Input</returns>
        public static byte[] GetBinaryFromGeneral(byte[] bytes)
        {
            List<byte> lstBytes = new List<byte>();
            byte[] bytesSize = BitConverter.GetBytes(bytes.Length);
            lstBytes.Add(bytesSize[0]);
            lstBytes.Add(bytesSize[1]);
            lstBytes.AddRange(bytes);

            return lstBytes.ToArray();
        }

        /// <summary>
        ///  Get a property's value from bytes.
        /// </summary>
        /// <param name="property">Send the property</param>
        /// <param name="bytes">The response buffer binary bytes</param>
        /// <returns>Return PropertyObj</returns>
        public static PropertyObj GetPropertyObjFromBuffer(PropertyObj property, byte[] bytes)
        {
            if (IsErrorCode(bytes))
            {
                property.Value = bytes;
                return property;
            }

            switch (property.ValueType)
            {
                case PropertyType.PtypInteger32:
                    property.Value = BitConverter.ToInt32(bytes, 0);
                    break;
                case PropertyType.PtypString:
                    string strResult = string.Empty;
                    foreach (byte b in bytes)
                    {
                        if (b == 0x00)
                        {
                            continue;
                        }

                        strResult += (char)b;
                    }

                    property.Value = strResult;
                    break;
                case PropertyType.PtypBoolean:
                    site.Assert.AreEqual<int>(1, bytes.Length, "PtypBoolean should be 1 byte");
                    site.Assert.IsTrue(bytes[0] == 0x00 || bytes[0] == 0x01, "PtypBoolean should be restricted to 1 or 0.");
                    property.Value = BitConverter.ToBoolean(bytes, 0);
                    break;
                case PropertyType.PtypMultipleString:
                    List<byte> lstBytes = new List<byte>();
                    List<string> lstStrs = new List<string>();

                    for (int i = 2; i < bytes.Length; i += 2)
                    {
                        if (bytes[i] == 0x00)
                        {
                            // Find the 0x0000 end flag to convert each string.
                            string strSingleResult = string.Empty;
                            foreach (byte b in lstBytes)
                            {
                                strSingleResult += (char)b;
                            }

                            lstStrs.Add(strSingleResult);
                            lstBytes.Clear();
                        }
                        else
                        {
                            lstBytes.Add(bytes[i]);
                        }
                    }

                    property.Value = lstStrs.ToArray();

                    break;

                case PropertyType.PtypTime:
                    property.Value = DateTime.FromFileTimeUtc(BitConverter.ToInt64(bytes, 0));
                    break;

                case PropertyType.PtypBinary: 
                case PropertyType.PtypComObject:
                    property.Value = bytes;
                    break;

                default:
                    break;
            }

            return property;
        }

        /// <summary>
        /// Add a PropertyTag object to PropertyTag dictionary.
        /// </summary>
        /// <param name="propertyName">The name of property.</param>
        /// <param name="propertyTag">The PropertyTag object.</param>
        private static void AddPropertyTagToDic(PropertyNames propertyName, PropertyTag propertyTag)
        {
            if (PropertyTagDic.ContainsKey(propertyName) == false)
            {
                PropertyTagDic.Add(propertyName, propertyTag);
            }
        }
    }
}