namespace Microsoft.Protocols.TestSuites.MS_OXCPRPT
{
    using System;
    using System.Runtime.InteropServices;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter capture code for MS_OXCPRPTAdapter
    /// </summary>
    public partial class MS_OXCPRPTAdapter : IMS_OXCPRPTAdapter
    {
        #region Verify requirements in MS-OXCPRPT

        #region Common Data
        /// <summary>
        /// The GUID of a Property Name. If GUID is PS_MAPI, the associated value is as below shown.
        /// </summary>
        private byte[] valuePSMAPI = { 0x28, 0x03, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 };

        #endregion Common Data

        #region MAPIHTTP transport

        /// <summary>
        /// Verify MAPIHTTP transport.
        /// </summary>
        private void VerifyMAPITransport()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http" && Common.IsRequirementEnabled(5000, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R5000");

                // Verify requirement MS-MS-OXCPRPT_R5000
                // If the transport sequence is MAPIHTTP and the code can reach here, it means that the implementation does support MAPIHTTP transport.
                Site.CaptureRequirement(
                        5000,
                        @"[In Appendix A: Product Behavior] Implementation does support this specification [MS-OXCMAPIHTTP]. (Exchange Server 2013 Service Pack 1 (SP1) and above follow this behavior.)");
            }
        }
        #endregion

        #region Verify Common Object Properties
        /// <summary>
        /// Verify Common Object properties related requirements.
        /// </summary>
        /// <param name="propertyTag">A property tag both identifies a property and gives the data type its value.</param>
        /// <param name="ropGetPropertiesSecificResponse">The RopGetPropertiesSpecific response buffer structure.</param>
        /// <param name="objectType">Server object type.</param>
        private void VerifyCommonProperties(PropertyTag propertyTag, RopGetPropertiesSpecificResponse ropGetPropertiesSecificResponse, ServerObjectType objectType)
        {
            // Identifies the property
            ushort propertyId = propertyTag.PropertyId;

            foreach (PropertyValue propertyVal in ropGetPropertiesSecificResponse.RowData.PropertyValues)
            {
                // The value of the property. Its syntax is specified by the PropertyType field of the tag, and its semantics by the PropertyId field of the tag, defined in MS-OXCDATA section 2.12.4 TaggedPropertyValue
                uint propertyValue = Common.ConvertByteArrayToUint(propertyVal.Value);

                switch (propertyId)
                {
                    // According to MS-OXPROPS section 2.575 PidTagAccess, if PropertyId is 0x0FF4, it means the property is PidTagAccess.
                    case 0x0FF4:
                        {
                            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R311
                            // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to Indicates that all property values are present and without error.
                            Site.CaptureRequirementIfAreEqual<uint>(
                                0x00,
                                ropGetPropertiesSecificResponse.RowData.Flag,
                                311,
                                @"[In PidTagAccess Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");

                            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R14
                            int notUsedbits = ~(0x00000001 | 0x00000002 | 0x00000004 | 0x00000008 | 0x00000010 | 0x00000020);
                            bool isVerifyR14 = (propertyValue & notUsedbits) == 0;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR14,
                                14,
                                @"[In PidTagAccess Property] The value is a bitwise-OR of zero or more values[0x00000001,0x00000002,0x00000004,0x00000008,0x00000010,0x00000020] from the following table.");
                        }

                        break;

                    // According to MS-OXPROPS 2.577 PidTagAccessLevel, if PropertyId is 0x0FF7, it means the property is PidTagAccessLevel.
                    case 0x0FF7:
                        {
                            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R312
                            // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to Indicates that all property values are present and without error.
                            Site.CaptureRequirementIfAreEqual<uint>(
                                0x00,
                                ropGetPropertiesSecificResponse.RowData.Flag,
                                312,
                                @"[In PidTagAccessLevel Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");

                            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R26
                            bool isVerifyR26 = (propertyValue == 0x00000000) ||
                                               (propertyValue == 0x00000001);

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR26,
                                26,
                                @"[In PidTagAccessLevel Property] This value of this property MUST be one of the values in the following table. [0x00000000, 0x00000001].");

                            Site.CaptureRequirementIfAreNotEqual<ServerObjectType>(
                              ServerObjectType.Folder,
                              objectType,
                              2401,
                              @"[In PidTagAccessLevel Property] This property does not apply to Folder objects.");

                            Site.CaptureRequirementIfAreNotEqual<ServerObjectType>(
                                ServerObjectType.Logon,
                                objectType,
                                2402,
                                @"[In PidTagAccessLevel Property] This property does not apply to Logon objects.");
                        }

                        break;

                    // According to MS-OXPROPS 2.701 PidTagChangeKey, if PropertyId is 0x65E2, it means the property is PidTagChangeKey
                    case 0x65E2:
                        {
                            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R313
                            // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to Indicates that all property values are present and without error.
                            Site.CaptureRequirementIfAreEqual<uint>(
                                0x00,
                                ropGetPropertiesSecificResponse.RowData.Flag,
                                313,
                                @"[In PidTagChangeKey Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1)");
                        }

                        break;

                    // According to MS-OXPROPS 2.724 PidTagCreationTime, if PropertyId is 0x3007, it means the property is PidTagCreationTime
                    case 0x3007:
                        {
                            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R314
                            // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to indicate that all property values are present and without error.
                            Site.CaptureRequirementIfAreEqual<uint>(
                                0x00,
                                ropGetPropertiesSecificResponse.RowData.Flag,
                                314,
                                @"[In PidTagCreationTime Property] Type: PtypTime ([MS-OXCDATA] section 2.11.1)");
                        }

                        break;

                    // According to MS-OXPROPS 2.883 PidTagLastModifierName, if PropertyId is 0x3FFA, it means the property is PidTagLastModifierName.
                    case 0x3FFA:
                        {
                            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R315.
                            // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to indicate that all property values are present and without error.
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                0x00,
                                ropGetPropertiesSecificResponse.RowData.Flag,
                                315,
                                @"[In PidTagLastModifierName Property] Type: PtypString ([MS-OXCDATA] section 2.11.1)");

                            // The property value should be the user who creates the object.
                            string result = Encoding.Unicode.GetString(ropGetPropertiesSecificResponse.RowData.PropertyValues[0].Value).TrimEnd('\0').ToLower();
                            Site.CaptureRequirementIfAreEqual<string>(
                                Common.GetConfigurationPropertyValue("UserName", this.Site).ToLower(),
                                result,
                                34,
                                @"[In PidTagLastModifierName Property] The PidTagLastModifierName property ([MS-OXPROPS] section 2.757) contains the name of the last mail user to modify the object.");

                            Site.CaptureRequirementIfAreNotEqual<ServerObjectType>(
                                 ServerObjectType.Folder,
                                 objectType,
                                 3401,
                                 @"[In PidTagLastModifierName Property] This property does not apply to Folder objects.");

                            Site.CaptureRequirementIfAreNotEqual<ServerObjectType>(
                                ServerObjectType.Logon,
                                objectType,
                                3402,
                                @"[In PidTagLastModifierName Property] This property does not apply to Logon objects.");
                        }

                        break;

                    // According to MS-OXPROPS 2.831 PidTagLastModificationTime, if PropertyId is 0x3008, it means the property is PidTagLastModificationTime.
                    case 0x3008:
                        {
                            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R316
                            // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to indicate that all property values are present and without error.
                            Site.CaptureRequirementIfAreEqual<uint>(
                                0x00,
                                ropGetPropertiesSecificResponse.RowData.Flag,
                                316,
                                @"[In PidTagLastModificationTimeProperty] Type: PtypTime ([MS-OXCDATA] section 2.11.1)");
                        }

                        break;

                    // According to MS-OXPROPS 2.1059 PidTagSearchKey, if PropertyId is 0x300B, it means the property is PidTagSearchKey
                    case 0x300B:
                        {
                            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R319
                            // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to indicate that all property values are present and without error.
                            Site.CaptureRequirementIfAreEqual<uint>(
                                0x00,
                                ropGetPropertiesSecificResponse.RowData.Flag,
                                319,
                                @"[In PidTagSearchKey Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");

                            Site.CaptureRequirementIfAreNotEqual<ServerObjectType>(
                                ServerObjectType.Folder,
                                objectType,
                                5501,
                                @"[In PidTagSearchKey Property] This property does not apply to Folder objects.");

                            Site.CaptureRequirementIfAreNotEqual<ServerObjectType>(
                                ServerObjectType.Logon,
                               objectType,
                               5502,
                               @"[In PidTagSearchKey Property] This property does not apply to Logon objects. ");
                        }

                        break;

                    default:
                        break;
                }
            }
        }
        #endregion

        #region RopGetPropertiesSpecific
        /// <summary>
        /// Verify the RopGetPropertiesSpecific operation related requirements.
        /// </summary>
        /// <param name="ropGetPropertiesSpecificResponse">The RopGetPropertiesSpecific response buffer structure.</param>
        /// <param name="objectType">Indicates which object type the RopGetPropertiesSpecific operation is acting on.</param>
        /// <param name="isOrdered">A boolean value which indicates whether the order of elements in request/response buffer is right.</param>
        /// <param name="isPropertySizeLimitZero">Indicates if PropertySizeLimit is set to 0.</param>
        /// <param name="isPropertyLagerThanLimit">Indicates if returned value size is larger than the size limit.</param>
        /// <param name="isValueContainsNotEnoughMemory">Indicates if the value is larger than limit or not.</param>
        private void VerifyRopGetPropertiesSpecific(
            RopGetPropertiesSpecificResponse ropGetPropertiesSpecificResponse,
            ServerObjectType objectType,
            bool isOrdered,
            bool isPropertySizeLimitZero,
            bool isPropertyLagerThanLimit,
            bool isValueContainsNotEnoughMemory)
        {
            // Since the RopGetPropertiesSpecific ROP response was parsed successfully, MS-OXCPRPT_R45602 can be captured directly.
            Site.CaptureRequirement(
                45602,
                @"[In Processing RopGetPropertiesSpecific] The server responds with a RopGetPropertiesSpecific ROP response buffer. ");

            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R56
            // Since the return value is in order, it verifies the return value is returned in order as the properties specified in the PropertyTags field. So MS-OXCPRPT_R56 can be captured.
            Site.CaptureRequirementIfIsTrue(
                isOrdered,
                56,
                @"[In RopGetPropertiesSpecific ROP] The RopGetPropertiesSpecific ROP ([MS-OXCROPS] section 2.2.8.3) queries for and returns the values of properties specified in the PropertyTags field.");

            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R45702
            // Since the response for RopGetPropertiesSpecific doesn't contain PropertyTags information, if the returned values for PropertyTags fields in RopGetPropertiesSpecific are same with the setting values in RopSetProperties. 
            // It means the properties to be returned are specified in the PropertyTags field of the request. 
            Site.CaptureRequirementIfIsTrue(
                isOrdered,
                45702,
                @"[In Processing RopGetPropertiesSpecific] The properties to be returned are specified in the PropertyTags field of the ROP request buffer.");

            // The order has been verified when de-serializing the related structures and represented by the variable isOrdered.
            Site.CaptureRequirementIfIsTrue(
               isOrdered,
               "MS-OXCDATA",
               64,
               @"[In PropertyRow Structures] For the RopGetPropertiesSpecific ROP, property values are returned in the order that the properties were requested.");

            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R45701
            // The order has been verified when de-serializing the related structures and represented by the variable isOrdered.
            Site.CaptureRequirementIfIsTrue(
                isOrdered,
                45701,
                @"[In Processing RopGetPropertiesSpecific] The server MUST order properties in the PropertyValues field of the ROP response buffer in the same order in which properties are specified in the PropertyTags field.");

            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R76
            // The order has been verified when de-serializing the related structures and represented by the variable isOrdered.
            Site.CaptureRequirementIfIsTrue(
                isOrdered,
                76,
                @"[In RopGetPropertiesSpecific ROP Response Buffer] RowData (variable): A PropertyRow structure ([MS-OXCDATA] section 2.8.1) that contains the values of the properties specified in the ROP request buffer.");

            if (isPropertySizeLimitZero)
            {
                bool isPropertyLimitedByBufferSize = false;
                uint len = 0;
                foreach (PropertyValue value in ropGetPropertiesSpecificResponse.RowData.PropertyValues)
                {
                    len = len + (uint)value.Size();
                }

                isPropertyLimitedByBufferSize = len < uint.MaxValue;

                Site.CaptureRequirementIfIsTrue(
                    isPropertyLimitedByBufferSize,
                    62,
                    @"[In RopGetPropertiesSpecific ROP Request Buffer] PropertySizeLimit: If this value is zero, the property values are limited only by the size of the ROP response buffer. ");
            }

            if (ropGetPropertiesSpecificResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                switch (objectType)
                {
                    case ServerObjectType.Message:
                        
                         // If the object type this operation acting on is Message object and this operation is performed successfully, then the following requirement can be captured.
                        Site.CaptureRequirement(
                            5801,
                            @"[In RopGetPropertiesSpecific ROP] Message objects are supported for this operation [RopGetPropertiesSpecific ROP].");
                        break;
                    case ServerObjectType.Attachment:
                        
                        // If the object type this operation acting on is Attachment object and this operation is performed successfully, then the following requirement can be captured.
                        Site.CaptureRequirement(
                            5803,
                            @"[In RopGetPropertiesSpecific ROP] Attachment objects are supported for this operation [RopGetPropertiesSpecific ROP].");
                        break;
                    case ServerObjectType.Folder:
                        
                        // If the object type this operation acting on is Folder object and this operation is performed successfully, then the following requirement can be captured.
                        Site.CaptureRequirement(
                            5802,
                            @"[In RopGetPropertiesSpecific ROP] Folder objects are supported for this operation [RopGetPropertiesSpecific ROP].");
                        break;
                    case ServerObjectType.Logon:
                        
                        // If the object type this operation acting on is Logon object and this operation is performed successfully, then the following requirement can be captured.
                        Site.CaptureRequirement(
                            5804,
                            @"[In RopGetPropertiesSpecific ROP] Logon objects are supported for this operation [RopGetPropertiesSpecific ROP].");
                        break;
                    default:
                        Site.Assert.Fail("The returned server object type is {0}, which is invalid.", objectType);
                        break;
                }
            }

            if (Common.IsRequirementEnabled(86703, this.Site))
            {
                if (this.cprptCurrentType == ServerObjectType.Message || this.cprptCurrentType == ServerObjectType.Attachment)
                {
                    if (isPropertyLagerThanLimit)
                    {
                        // If MS-OXCPRPT_R86703 is enabled, when the property is a PtypBinary type, a PtypObject type, or a string property, 
                        // the NotEnoughMemory (0x8007000E) will be returned, if the value is larger than either the available space in the ROP response buffer or the size specified in the PropertySizeLimit field of the ROP request buffer.
                        // MS-OXCPRPT_R86703 is captured in model. This requirement is captured here since the requirements from other protocols cannot be handled in model.
                        Site.CaptureRequirementIfIsTrue(
                             isValueContainsNotEnoughMemory,
                             "MS-OXCDATA",
                             2052,
                             @"[In Additional Error Codes] NotEnoughMemory (E_NOMEMORY, MAPI_E_NOT_ENOUGH_MEMORY) will be returned, On get, indicates that the property or column value is too large to be retrieved by the request, and the property value needs to be accessed with the RopOpenStream ROP ([MS-OXCROPS] section 2.2.9.1).");
                    }
                }
            }
        }
        #endregion

        #region VerifyRopGetPropertiesSpecificForWantUnicode
        /// <summary>
        /// Verify the RopGetPropertiesSpecific operation related requirements for WantUnicode. 
        /// </summary>
        /// <param name="ropGetPropertiesSpecificResponse">The RopGetPropertiesSpecific response buffer structure.</param>
        /// <param name="isUnicode">Indicates if string properties are to be encoded in the Unicode format in the ROP response buffer.</param>
        private void VerifyRopGetPropertiesSpecificForWantUnicode(RopGetPropertiesSpecificResponse ropGetPropertiesSpecificResponse, bool isUnicode)
        {
            TypedPropertyValue value = (TypedPropertyValue)ropGetPropertiesSpecificResponse.RowData.PropertyValues[0];

            if (isUnicode)
            {
                Site.CaptureRequirementIfAreEqual<ushort>(
                    (ushort)0x001f,
                    value.PropertyType,
                    6401,
                    @"[In RopGetPropertiesSpecific ROP Request Buffer] WantUnicode: A Boolean value that is nonzero if the string properties that are requested with PtypUnspecified ([MS-OXCDATA] section 2.11.1) as the property type are to be encoded in the Unicode format in the ROP response buffer. ");

                Site.CaptureRequirementIfAreEqual<ushort>(
                    (ushort)0x001f,
                    value.PropertyType,
                    458,
                    @"[In Processing RopGetPropertiesSpecific] If the WantUnicode field is set to a nonzero value, the server MUST return string properties that are requested without a specified type (PtypUnspecified) in Unicode format. ");
            }
            else
            {
                Site.CaptureRequirementIfAreEqual<ushort>(
                   (ushort)0x001e,
                   value.PropertyType,
                   65,
                   @"[In RopGetPropertiesSpecific ROP Request Buffer] WantUnicode: If WantUnicode is set to zero, the string properties that are requested with PtypUnspecified as the property type are to be encoded in multibyte character set (MBCS) format.");

                Site.CaptureRequirementIfAreEqual<ushort>(
                   (ushort)0x001e,
                   value.PropertyType,
                   459,
                   @"[In Processing RopGetPropertiesSpecific] If the WantUnicode field is set to zero, the server MUST return string properties that are requested without a specified type (PtypUnspecified) in MBCS format.");
            }

            Site.CaptureRequirementIfAreNotEqual<ushort>(
                (ushort)0x0000,
                value.PropertyType,
                "MS-OXCDATA",
                2720,
                @"[In Property Value Types] PtypUnspecified(PT_UNSPECIFIED) is 
                Any: this property type value matches any type; a server MUST return the actual type in its response. Servers MUST NOT return this type in response to a client request other than NspiGetIDsFromNames or the RopGetPropertyIdsFromNames ROP request ([MS-OXCROPS] section 2.2.8.1). with Property Type Value 0x0000,%x00.00.");
        }
        #endregion

        #region VerifyRopGetPropertiesSpecificForTaggedProperties
        /// <summary>
        /// Verify the RopGetPropertiesSpecific operation related requirements for tagged properties. 
        /// </summary>
        /// <param name="ropGetPropertiesSpecificResponse">The RopGetPropertiesSpecific response buffer structure.</param>
        private void VerifyRopGetPropertiesSpecificForTaggedProperties(RopGetPropertiesSpecificResponse ropGetPropertiesSpecificResponse)
        {
            Site.Assert.AreEqual<int>(3, ropGetPropertiesSpecificResponse.RowData.PropertyValues.Count, "The RopGetPropertiesSpecific operation should return 3 properties.");

            FlaggedPropertyValue firstProperty = ropGetPropertiesSpecificResponse.RowData.PropertyValues[0] as FlaggedPropertyValue;
            Site.Assert.IsNotNull(firstProperty, "The first property returned by the RopGetPropertiesSpecific operation should not be of FlaggedPropertyValue type.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R457001");

            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R457001
            // This requirement can be verified since the value (0x0) of flag field in FlaggedPropertyValue means the property value returned from server successfully.
            // The value (0x1) means the PropertyValue field is not present.
            // The value (0xa) means the value is a property error code.
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                firstProperty.Flag,
                457001,
                @"[In Processing RopGetPropertiesSpecific] The server MUST return the values for the explicitly requested properties on the object, including those set by any client.");

            // The PidTagRecordKey property does not apply to Folder objects and Logon objects.
            if (this.cprptCurrentType == ServerObjectType.Attachment)
            {
                FlaggedPropertyValue secondProperty = ropGetPropertiesSpecificResponse.RowData.PropertyValues[1] as FlaggedPropertyValue;
                Site.Assert.IsNotNull(secondProperty, "The second property returned by the RopGetPropertiesSpecific operation should not be of FlaggedPropertyValue type.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R457002");

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R457002
                // This requirement can be verified since the value (0x0) of flag field in FlaggedPropertyValue means the property value returned from server successfully.
                // The value (0x1) means the PropertyValue field is not present.
                // The value (0xa) means the value is a property error code.
                Site.CaptureRequirementIfAreEqual<byte>(
                    0,
                    secondProperty.Flag,
                    457002,
                    @"[In Processing RopGetPropertiesSpecific] The server MUST return the values for the explicitly requested properties on the object, including those set by any server.");
            }

            // The computation of the value of the PidTagConversationId property is not supported by Exchange 2007.
            if (Common.IsRequirementEnabled(457003, this.Site) && this.cprptCurrentType == ServerObjectType.Message)
            {
                FlaggedPropertyValue thirdProperty = ropGetPropertiesSpecificResponse.RowData.PropertyValues[2] as FlaggedPropertyValue;
                Site.Assert.IsNotNull(thirdProperty, "The third property returned by the RopGetPropertiesSpecific operation should not be of FlaggedPropertyValue type.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R457003");

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R457003
                // This requirement can be verified since the value (0x0) of flag field in FlaggedPropertyValue means the property value returned from server successfully.
                // The value (0x1) means the PropertyValue field is not present.
                // The value (0xa) means the value is a property error code.
                Site.CaptureRequirementIfAreEqual<byte>(
                    0,
                    thirdProperty.Flag,
                    457003,
                    @"[In Processing RopGetPropertiesSpecific] The server MUST return the values for the explicitly requested properties on the object, including those computed properties. ");
            }
        }
        #endregion

        #region VerifyPropertyValueAfterPropertyDeleted
        /// <summary>
        /// Verify whether the property value is NotFound when querying this property after deleted.
        /// </summary>
        /// <param name="propertyValue">The returned value of the property.</param>
        /// <param name="isRopDeletePropertiesMethod">Indicates if the ROP method is RopDeleteProperties.</param>
        private void VerifyPropertyValueAfterPropertyDeleted(byte[] propertyValue, bool isRopDeletePropertiesMethod)
        {
            byte[] notFound = new byte[] { 0x0F, 0x01, 0x04, 0x80 };
            bool isR875Satisfied = Common.CompareByteArray(propertyValue, notFound);

            if (Common.IsRequirementEnabled(875, this.Site))
            {                
                if (isRopDeletePropertiesMethod)
                {
                    // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R875
                    Site.CaptureRequirementIfIsTrue(
                        isR875Satisfied,
                        875,
                        @"[In Processing RopDeleteProperties] Implementation does return the NotFound error(0x8004010F) ([MS-OXCDATA] section 2.4.2) in place of a value, if a client request the value of this property[which is deleted].(Microsoft Exchange Server 2007 and above follow this behavior.)");
                }
            }

            if (Common.IsRequirementEnabled(83603, this.Site))
            {
                if (!isRopDeletePropertiesMethod)
                {
                    // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R83603
                    Site.CaptureRequirementIfIsTrue(
                       isR875Satisfied,
                       83603,
                       @"[In RopDeletePropertiesNoReplicate] Implementation does return the NotFound error(0x8004010F) ([MS-OXCDATA] section 2.4.2) in place of a value, if a client request the value of this property.(Microsoft Exchange Server 2007 and above follow this behavior.).");
                }
            }
        }
        #endregion

        #region RopGetPropertiesAll
        /// <summary>
        /// Verify the RopGetPropertiesAll operation related requirements.
        /// </summary>
        /// <param name="ropGetPropertiesAllResponse">The RopGetPropertiesAll response buffer structure.</param>
        /// <param name="objectType">Indicates which object type the RopGetPropertiesAll operation is acting on.</param>
        /// <param name="isPropertySizeLimitZero">Indicates if the PropertySizeLimit is set to 0.</param>
        /// <param name="isPropertyLagerThanLimit">Indicates if the property size is larger than the limit.</param>
        /// <param name="isUnicode">Indicates if string properties are to be encoded in the Unicode format in the ROP response buffer.</param>
        /// <param name="isValueContainsNotEnoughMemory">Indicates if returned property is larger than the limit or not.</param>
        private void VerifyRopGetPropertiesAll(RopGetPropertiesAllResponse ropGetPropertiesAllResponse, ServerObjectType objectType, bool isPropertySizeLimitZero, bool isPropertyLagerThanLimit, bool isUnicode, bool isValueContainsNotEnoughMemory)
        {
            // Since the RopGetPropertiesAll ROP response was parsed successfully, MS-OXCPRPT_R90703 can be captured directly.
            Site.CaptureRequirement(
                 90703,
                 @"[In Processing RopGetPropertiesAll] The server responds with a RopGetPropertiesAll ROP response buffer");

            if (ropGetPropertiesAllResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                bool containString8 = false;
                bool containString = false;
                bool hasString = false;
                foreach (TaggedPropertyValue value in ropGetPropertiesAllResponse.PropertyValues)
                {
                    if (value.PropertyTag.PropertyType == (uint)0x001e)
                    {
                        containString8 = true;
                        hasString = true;
                    }
                    else if (value.PropertyTag.PropertyType == (uint)0x001f)
                    {
                        containString = true;
                        hasString = true;
                    }
                }

                if (hasString)
                {
                    if (isUnicode)
                    {
                        Site.CaptureRequirementIfIsTrue(
                           !containString8 && containString,
                           8502,
                           @"[In RopGetPropertiesAll ROP Request Buffer] WantUnicode: A Boolean value that is nonzero if string properties that are requested with PtypUnspecified ([MS-OXCDATA] section 2.11.1) as the property type are to be encoded in the Unicode format in the ROP response buffer. ");  
                    }
                    else
                    {
                        Site.CaptureRequirementIfIsFalse(
                           containString,
                           8702,
                           @"[In RopGetPropertiesAll ROP Request Buffer] WantUnicode: If WantUnicode is set to zero, the string properties that are requested with PtypUnspecified as the property type are to be encoded in MBCS format.");
                    }
                }

                bool isReturnPropertyTypeNotPtypUnspecified = true;
                foreach (TaggedPropertyValue value in ropGetPropertiesAllResponse.PropertyValues)
                {
                    if (value.PropertyTag.PropertyType == 0x0000)
                    {
                        isReturnPropertyTypeNotPtypUnspecified = false;
                        break;
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isReturnPropertyTypeNotPtypUnspecified,
                    "MS-OXCDATA",
                    2720,
                    @"[In Property Value Types] PtypUnspecified(PT_UNSPECIFIED) is 
                Any: this property type value matches any type; a server MUST return the actual type in its response. Servers MUST NOT return this type in response to a client request other than NspiGetIDsFromNames or the RopGetPropertyIdsFromNames ROP request ([MS-OXCROPS] section 2.2.8.1). with Property Type Value 0x0000,%x00.00.");

                if (isPropertySizeLimitZero)
                {
                    bool isPropertyLimitedByBufferSize = true;
                    foreach (TaggedPropertyValue value in ropGetPropertiesAllResponse.PropertyValues)
                    {
                        if (value.Value != null && value.Size() > 0xBABE)
                        {
                            isPropertyLimitedByBufferSize = false;
                            break;
                        }
                    }

                    Site.CaptureRequirementIfIsTrue(
                        isPropertyLimitedByBufferSize,
                        84,
                        @"[In RopGetPropertiesAll ROP Request Buffer] PropertySizeLimit: If this value[PropertySizeLimit] is zero, the property values are only limited by the size of the ROP[RopGetPropertiesAll] response buffer.");
                }

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R9201
                // numberOfPropertyValues indicates the number of elements in PropertyValues.
                ushort numberOfPropertyValues = 0;
                if (ropGetPropertiesAllResponse.PropertyValues != null)
                {
                    numberOfPropertyValues = (ushort)ropGetPropertiesAllResponse.PropertyValues.Length;
                }

                Site.CaptureRequirementIfAreEqual<ushort>(
                    numberOfPropertyValues,
                    ropGetPropertiesAllResponse.PropertyValueCount,
                    9201,
                    @"[In RopGetPropertiesAll ROP Response Buffer] PropertyValueCount: An integer that specifies the number of elements in the PropertyValues field.");

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                        92,
                        @"[In RopGetPropertiesAll ROP Response Buffer] PropertyValueCount: 2 bytes integer.");

                if (objectType == ServerObjectType.Message)
                {
                    // If the object type this operation acting on is Message object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        8001,
                        @"[In RopGetPropertiesAll ROP] This operation [RopGetPropertiesAll ROP] is valid on Message objects.");
                }

                if (objectType == ServerObjectType.Folder)
                {
                    // If the object type this operation acting on is Folder object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        8002,
                        @"[In RopGetPropertiesAll ROP] This operation [RopGetPropertiesAll ROP] is valid on Folder objects.");
                }

                if (objectType == ServerObjectType.Attachment)
                {
                    // If the object type this operation acting on is Attachment object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        8003,
                        @"[In RopGetPropertiesAll ROP] This operation [RopGetPropertiesAll ROP] is valid on Attachment objects.");
                }

                if (objectType == ServerObjectType.Logon)
                {
                    // If the object type this operation acting on is Logon object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        8004,
                        @"[In RopGetPropertiesAll ROP] This operation [RopGetPropertiesAll ROP] is valid on Logon objects.");
                }
            }

            if (Common.IsRequirementEnabled(86703, this.Site))
            {
                if (this.cprptCurrentType == ServerObjectType.Message || this.cprptCurrentType == ServerObjectType.Attachment)
                {
                    if (isPropertyLagerThanLimit)
                    {
                        // If MS-OXCPRPT_R86703 is enabled, when the property is a PtypBinary type, a PtypObject type, or a string property, 
                        // the NotEnoughMemory (0x8007000E) will be returned, if the value is larger than either the available space in the ROP response buffer or the size specified in the PropertySizeLimit field of the ROP request buffer.
                        // MS-OXCPRPT_R86703 is captured in model. This requirement is captured here since the requirements from other protocols cannot be handled in model.
                        Site.CaptureRequirementIfIsTrue(
                            isValueContainsNotEnoughMemory,
                            "MS-OXCDATA",
                            2052,
                            @"[In Additional Error Codes] NotEnoughMemory (E_NOMEMORY, MAPI_E_NOT_ENOUGH_MEMORY) will be returned, On get, indicates that the property or column value is too large to be retrieved by the request, and the property value needs to be accessed with the RopOpenStream ROP ([MS-OXCROPS] section 2.2.9.1).");
                    }
                }
            }

            if (ropGetPropertiesAllResponse.PropertyValueCount != 0x00)
            {
                foreach (TaggedPropertyValue propertyValue in ropGetPropertiesAllResponse.PropertyValues)
                {
                    this.VerifyPropertyTagInCDATA(propertyValue.PropertyTag);
                }
            }
        }
        #endregion RopGetPropertiesAll

        #region RopGetPropertiesList
        /// <summary>
        /// Verify the RopGetPropertiesList operation related requirements.
        /// </summary>
        /// <param name="ropGetPropertiesListResponse">The RopGetPropertiesList response buffer structure.</param>
        /// <param name="objectType">Indicates which object type the RopGetPropertiesList operation is acting on.</param>
        private void VerifyRopGetPropertiesList(RopGetPropertiesListResponse ropGetPropertiesListResponse, ServerObjectType objectType)
        {
            // Since the RopGetPropertiesList ROP response was parsed successfully, MS-OXCPRPT_R47102 can be captured directly.
            Site.CaptureRequirement(
              47102,
              @"[In Processing RopGetPropertiesList] The server responds with a RopGetPropertiesList ROP response buffer.");

            if (ropGetPropertiesListResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                if (objectType == ServerObjectType.Message)
                {
                    // If the object type this operation acting on is Message object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        9701,
                        @"[In RopGetPropertiesList ROP] This operation[RopGetPropertiesList ROP] is valid on Message objects.");
                }

                if (objectType == ServerObjectType.Folder)
                {
                    // If the object type this operation acting on is Folder object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        9702,
                        @"[In RopGetPropertiesList ROP] This operation[RopGetPropertiesList ROP] is valid on Folder objects.");
                }

                if (objectType == ServerObjectType.Attachment)
                {
                    // If the object type this operation acting on is Attachment object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        9703,
                        @"[In RopGetPropertiesList ROP] This operation[RopGetPropertiesList ROP] is valid on Attachment objects.");
                }

                if (objectType == ServerObjectType.Logon)
                {
                    // If the object type this operation acting on is Logon object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        9704,
                        @"[In RopGetPropertiesList ROP] This operation[RopGetPropertiesList ROP] is valid on Logon objects.");
                }
            }

            if (ropGetPropertiesListResponse.PropertyTagCount != 0x00)
            {
                foreach (PropertyTag propertyTag in ropGetPropertiesListResponse.PropertyTags)
                {
                    this.VerifyPropertyTagInCDATA(propertyTag);
                }
            }

            ushort numberOfPropertyTag = 0;
            if (ropGetPropertiesListResponse.PropertyTagCount > 0)
            {
                numberOfPropertyTag = (ushort)ropGetPropertiesListResponse.PropertyTags.Length;

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                   100,
                   @"[In RopGetPropertiesList ROP Response Buffer] PropertyTags (variable): An array of PropertyTag structures ([MS-OXCDATA] section 2.9). ");
            }

            Site.CaptureRequirementIfAreEqual<ushort>(
                numberOfPropertyTag,
                ropGetPropertiesListResponse.PropertyTagCount,
                9903,
                @"[In RopGetPropertiesList ROP Response Buffer] PropertyTagCount: An integer that specifies the number of property tags contained in the PropertyTags field.");

            // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
            Site.CaptureRequirement(
                99,
                @"[In RopGetPropertiesList ROP Response Buffer] PropertyTagCount: 2 bytes integer.");
        }
        #endregion

        #region RopSetProperties
        /// <summary>
        /// Verify the RopSetProperties operation related requirements.
        /// </summary>
        /// <param name="ropSetPropertiesResponse">The RopSetProperties response buffer structure.</param>
        /// <param name="objectType">Indicates which object type the RopSetProperties operation is acting on.</param>
        private void VerifyRopSetProperties(RopSetPropertiesResponse ropSetPropertiesResponse, ServerObjectType objectType)
        {
            // Since the RopSetProperties ROP response was parsed successfully, MS-OXCPRPT_R47502 can be captured directly.
            Site.CaptureRequirement(
                 47502,
                 @"[In Processing RopSetProperties] The server responds with a RopSetProperties ROP response buffer.");

            if (ropSetPropertiesResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                if (objectType == ServerObjectType.Message)
                {
                    // If the object type this operation acting on is Message object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        10201,
                        @"[In RopSetProperties ROP] This operation [RopSetProperties ROP] is valid on Message objects.");
                }

                if (objectType == ServerObjectType.Folder)
                {
                    // If the object type this operation acting on is Folder object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        10202,
                        @"[In RopSetProperties ROP] This operation [RopSetProperties ROP] is valid on Folder objects.");
                }

                if (objectType == ServerObjectType.Attachment)
                {
                    // If the object type this operation acting on is Attachment object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        10203,
                        @"[In RopSetProperties ROP] This operation [RopSetProperties ROP] is valid on Attachment objects.");
                }

                if (objectType == ServerObjectType.Logon)
                {
                    // If the object type this operation acting on is Logon object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        10204,
                        @"[In RopSetProperties ROP] This operation [RopSetProperties ROP] is valid on Logon objects.");
                }

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                    107,
                   @"[In RopSetProperties ROP Response Buffer] propertyProblemCount: 2 bytes integer.");

                ushort numberOfPropertyProblem = 0;
                if (ropSetPropertiesResponse.PropertyProblemCount != 0)
                {
                    // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                    Site.CaptureRequirement(
                        108,
                        @"[In RopSetProperties ROP Response Buffer] propertyProblems (variable): An array of PropertyProblem structures ([MS-OXCDATA] section 2.7).");

                    for (int counter = 0; counter < ropSetPropertiesResponse.PropertyProblems.Length; counter++)
                    {
                        this.VerifyPropertyProblemSturctureInCDATA();
                    }

                    numberOfPropertyProblem = (ushort)ropSetPropertiesResponse.PropertyProblems.Length;
                }

                Site.CaptureRequirementIfAreEqual<ushort>(
                    numberOfPropertyProblem,
                    ropSetPropertiesResponse.PropertyProblemCount,
                    10701,
                    @"[In RopSetProperties ROP Response Buffer] propertyProblemCount: An integer that specifies the number of elements contained in the PropertyProblems field.");
            }
        }
        #endregion

        #region RopSetPropertiesNoReplicate
        /// <summary>
        /// Verify the RopSetPropertiesNoReplicate operation related requirements.
        /// </summary>
        /// <param name="ropSetPropertiesNoReplicateResponse">The RopSetPropertiesNoReplicate response buffer structure.</param>
        /// <param name="objectType">Indicates which object type the RopSetPropertiesNoReplicate operation is acting on.</param>
        private void VerifyRopSetPropertiesNoReplicate(RopSetPropertiesNoReplicateResponse ropSetPropertiesNoReplicateResponse, ServerObjectType objectType)
        {
            if (ropSetPropertiesNoReplicateResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                if (objectType == ServerObjectType.Message)
                {
                    // If the object type this operation acting on is Message object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        826,
                        @"[In RopSetPropertiesNoReplicate ROP] This ROP is supported for Message objects.");
                }

                if (objectType == ServerObjectType.Folder)
                {
                    // If the object type this operation acting on is Folder object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        828,
                        @"[In RopSetPropertiesNoReplicate ROP] This ROP is supported for Folder objects.");
                }

                if (objectType == ServerObjectType.Attachment)
                {
                    // If the object type this operation acting on is Attachment object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        830,
                        @"[In RopSetPropertiesNoReplicate ROP] This ROP is supported for Attachment objects.");
                }

                if (objectType == ServerObjectType.Logon)
                {
                    // If the object type this operation acting on is Logon object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        832,
                        @"[In RopSetPropertiesNoReplicate ROP] This ROP is supported for Logon objects.");
                }

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                  83307,
                  @"[In RopSetPropertiesNoReplicate ROP Response Buffer] propertyProblemCount: 2 bytes integer.");
            }

            ushort numberOfPropertyProblem = 0;
            if (ropSetPropertiesNoReplicateResponse.PropertyProblemCount != 0)
            {
                for (int counter = 0; counter < ropSetPropertiesNoReplicateResponse.PropertyProblems.Length; counter++)
                {
                    this.VerifyPropertyProblemSturctureInCDATA();
                }

                numberOfPropertyProblem = (ushort)ropSetPropertiesNoReplicateResponse.PropertyProblems.Length;
            }

            Site.CaptureRequirementIfAreEqual<ushort>(
             numberOfPropertyProblem,
             ropSetPropertiesNoReplicateResponse.PropertyProblemCount,
             83308,
             @"[In RopSetPropertiesNoReplicate ROP Response Buffer] propertyProblemCount: An integer that specifies the number of elements contained in the PropertyProblems field.");
        }
        #endregion

        #region RopDeleteProperties
        /// <summary>
        /// Verify the RopDeleteProperties operation related requirements.
        /// </summary>
        /// <param name="ropDeletePropertiesResponse">The RopDeleteProperties response buffer structure.</param>
        /// <param name="objectType">Indicates which object type the RopDeleteProperties operation is acting on </param>
        private void VerifyRopDeleteProperties(RopDeletePropertiesResponse ropDeletePropertiesResponse, ServerObjectType objectType)
        {
            // Since the RopDeleteProperties ROP response was parsed successfully, MS-OXCPRPT_R48402 can be captured directly.
            Site.CaptureRequirement(
                 48402,
                 @"[In Processing RopDeleteProperties] The server responds with a RopDeleteProperties ROP response buffer. ");

            if (ropDeletePropertiesResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                if (objectType == ServerObjectType.Message)
                {
                    // If the object type this operation acting on is Message object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        11401,
                        @"[In RopDeleteProperties ROP] This operation [RopDeleteProperties ROP] is valid on Message objects.");
                }

                if (objectType == ServerObjectType.Folder)
                {
                    // If the object type this operation acting on is Folder object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        11402,
                        @"[In RopDeleteProperties ROP] This operation [RopDeleteProperties ROP] is valid on Folder objects.");
                }

                if (objectType == ServerObjectType.Attachment)
                {
                    // If the object type this operation acting on is Attachment object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        11403,
                        @"[In RopDeleteProperties ROP] This operation [RopDeleteProperties ROP] is valid on Attachment objects.");
                }

                if (objectType == ServerObjectType.Logon)
                {
                    // If the object type this operation acting on is Logon object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        11404,
                        @"[In RopDeleteProperties ROP] This operation [RopDeleteProperties ROP] is valid on Logon objects.");
                }

                if (ropDeletePropertiesResponse.PropertyProblemCount != 0)
                {
                    // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                    Site.CaptureRequirement(
                       120,
                       @"[In RopDeleteProperties ROP Response Buffer] PropertyProblems (variable): An array of PropertyProblem structures ([MS-OXCDATA] section 2.7).");
                }

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                       119,
                       @"[In RopDeleteProperties ROP Response Buffer] PropertyProblemCount: 2 bytes integer.");
            }

            if (ropDeletePropertiesResponse.ReturnValue == 0x8004010F)
            {
                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2055,
                        @"[In Additional Error Codes] The numeric value (hex) for error code NotFound is 0x8004010F, %x0F.01.04.80.");
            }

            ushort numberOfPropertyProblem = 0;
            if (ropDeletePropertiesResponse.PropertyProblemCount != 0x00)
            {
                numberOfPropertyProblem = (ushort)ropDeletePropertiesResponse.PropertyProblems.Length;
            }

            Site.CaptureRequirementIfAreEqual<ushort>(
                numberOfPropertyProblem,
                ropDeletePropertiesResponse.PropertyProblemCount,
                11901,
                @"[In RopDeleteProperties ROP Response Buffer] PropertyProblemCount: An integer that specifies the number of elements contained in the PropertyProblems field.");

            if (ropDeletePropertiesResponse.PropertyProblemCount != 0x00)
            {
                for (int counter = 0; counter < ropDeletePropertiesResponse.PropertyProblems.Length; counter++)
                {
                    this.VerifyPropertyProblemSturctureInCDATA();
                }
            }
        }
        #endregion

        #region RopDeletePropertiesNoReplicate
        /// <summary>
        /// Verify the RopDeletePropertiesNoReplicate operation related requirements.
        /// </summary>
        /// <param name="ropDeletePropertiesNoReplicateResponse">The RopDeletePropertiesNoReplicate response buffer structure.</param>
        /// <param name="objectType">Indicates which object type the RopDeletePropertiesNoReplicate operation is acting on.</param>
        private void VerifyRopDeletePropertiesNoReplicate(RopDeletePropertiesNoReplicateResponse ropDeletePropertiesNoReplicateResponse, ServerObjectType objectType)
        {
            if (ropDeletePropertiesNoReplicateResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                if (objectType == ServerObjectType.Message)
                {
                    // If the object type this operation acting on is Message object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        837,
                        @"[In RopDeletePropertiesNoReplicate] This ROP is supported for Message objects.");
                }

                if (objectType == ServerObjectType.Folder)
                {
                    // If the object type this operation acting on is Folder object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        839,
                        @"[In RopDeletePropertiesNoReplicate] This ROP is supported for Folder objects.");
                }

                if (objectType == ServerObjectType.Attachment)
                {
                    // If the object type this operation acting on is Attachment object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        841,
                        @"[In RopDeletePropertiesNoReplicate] This ROP is supported for Attachment objects.");
                }

                if (objectType == ServerObjectType.Logon)
                {
                    // If the object type this operation acting on is Logon object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        843,
                        @"[In RopDeletePropertiesNoReplicate] This ROP is supported for Logon objects.");
                }

                if (ropDeletePropertiesNoReplicateResponse.PropertyProblemCount != 0)
                {
                    // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                    Site.CaptureRequirement(
                        84407,
                         @"[In RopDeletePropertiesNoReplicate ROP Response Buffer] PropertyProblems (variable): An array of PropertyProblem structures ([MS-OXCDATA] section 2.7).");
                }

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                      84405,
                      @"[In RopDeletePropertiesNoReplicate ROP Response Buffer] PropertyProblemCount: 2 bytes integer.");
            }

            if (ropDeletePropertiesNoReplicateResponse.ReturnValue == 0x8004010F)
            {
                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2055,
                        @"[In Additional Error Codes] The numeric value (hex) for error code NotFound is 0x8004010F, %x0F.01.04.80.");
            }

            ushort numberOfPropertyProblem = 0;

            if (ropDeletePropertiesNoReplicateResponse.PropertyProblemCount > 0)
            {
                for (int counter = 0; counter < ropDeletePropertiesNoReplicateResponse.PropertyProblems.Length; counter++)
                {
                    this.VerifyPropertyProblemSturctureInCDATA();
                }

                numberOfPropertyProblem = (ushort)ropDeletePropertiesNoReplicateResponse.PropertyProblems.Length;
            }

            Site.CaptureRequirementIfAreEqual<ushort>(
                numberOfPropertyProblem,
                ropDeletePropertiesNoReplicateResponse.PropertyProblemCount,
                84406,
                @"[In RopDeletePropertiesNoReplicate ROP Response Buffer] PropertyProblemCount: An integer that specifies the number of elements contained in the PropertyProblems field.");
        }
        #endregion

        #region RopQueryNamedProperties
        /// <summary>
        /// Verify the RopQueryNamedProperties operation related requirements.
        /// </summary>
        /// <param name="queryFlags">The QueryFlags specified in request buffer.</param>
        /// <param name="propertyGuid">The PropertyGuid specified in request buffer.</param>       
        /// <param name="ropQueryNamedPropertiesResopnse">The RopQueryNamedProperties response buffer structure.</param>
        /// <param name="objectType">Indicates which object type the RopQueryNamedProperties operation is acting on.</param>
        /// <param name="isOrdered">A boolean value which indicates whether the order of elements in request/response buffer is right.</param>
        private void VerifyRopQueryNamedProperties(
            QueryFlags queryFlags,
            byte[] propertyGuid,
            RopQueryNamedPropertiesResponse ropQueryNamedPropertiesResopnse,
            ServerObjectType objectType,
            bool isOrdered)
        {
            // Since the RopQueryNamedProperties ROP response was parsed successfully, MS-OXCPRPT_R49502 can be captured directly.
            Site.CaptureRequirement(
                 49502,
                 @"[In Processing RopQueryNamedProperties] The server responds with a RopQueryNamedProperties ROP response buffer.");

            if (ropQueryNamedPropertiesResopnse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                if (objectType == ServerObjectType.Message)
                {
                    // If the object type this operation acting on is Message object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        12901,
                        @"[In RopQueryNamedProperties ROP] This operation is valid on Message objects.");
                }

                if (objectType == ServerObjectType.Folder)
                {
                    // If the object type this operation acting on is Folder object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        12902,
                        @"[In RopQueryNamedProperties ROP] This operation is valid on Folder objects.");
                }

                if (objectType == ServerObjectType.Attachment)
                {
                    // If the object type this operation acting on is Attachment object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        12903,
                        @"[In RopQueryNamedProperties ROP] This operation is valid on Attachment objects.");
                }

                if (objectType == ServerObjectType.Logon)
                {
                    if (Common.IsRequirementEnabled(12904, this.Site))
                    {
                        // If the object type this operation acting on is Logon object and this operation is performed successfully, then the following requirement can be captured.
                        Site.CaptureRequirement(
                            12904,
                            @"[In RopQueryNamedProperties ROP] Implementation is valid on Logon objects. (Microsoft Exchange Server 2010 and above follow this behavior)");
                    }
                }

                if (ropQueryNamedPropertiesResopnse.IdCount != 0)
                {
                    // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                    Site.CaptureRequirement(
                        14301,
                        @"[In RopQueryNamedProperties ROP Response Buffer] propertyNames (variable): An array of PropertyName structures ([MS-OXCDATA] section 2.6.1). ");
                }

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                    14201,
                     @"[In RopQueryNamedProperties ROP Response Buffer] propertyIds (variable): An array of 16-bit integers, each of which is a property ID.");

                if (ropQueryNamedPropertiesResopnse.PropertyNames != null)
                {
                    Site.CaptureRequirementIfAreEqual<int>(
                       ropQueryNamedPropertiesResopnse.PropertyIds.Length,
                       ropQueryNamedPropertiesResopnse.PropertyNames.Length,
                       14203,
                       @"[In RopQueryNamedProperties ROP Response Buffer] PropertyIds: The array MUST contain one property ID for each of the named properties specified in the PropertyNames field.");

                    if (queryFlags == QueryFlags.NoStrings)
                    {
                        // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R132
                        bool isVerifyR132 = true;

                        // Examine every PropertyName in the ropQueryNamedPropertiesResponse.
                        for (int i = 0; i < ropQueryNamedPropertiesResopnse.IdCount; i++)
                        {
                            // If any Named properties with a Kind of 0x1 actually presents in the response, isVerifyR132 should be set to false.
                            if (ropQueryNamedPropertiesResopnse.PropertyNames[i].Kind == 0x01)
                            {
                                // Add the debug information
                                Site.Log.Add(
                                    LogEntryKind.Debug,
                                    "Verify MS-OXCPRPT_R132, QueryFlags is {0}, PropertyName.Kind in the response is {1}",
                                    queryFlags,
                                    ropQueryNamedPropertiesResopnse.PropertyNames[i].Kind);

                                isVerifyR132 = false;
                                break;
                            }
                        }

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR132,
                            132,
                            @"[In RopQueryNamedProperties ROP Request Buffer] QueryFlags: NoStrings:Named properties that have a property name identifier MUST NOT be included in the response.");
                    }

                    if (queryFlags == QueryFlags.NoIds)
                    {
                        // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R134
                        bool isVerifyR134 = true;

                        // Examine every PropertyName in the ropQueryNamedPropertiesResponse.
                        for (int i = 0; i < ropQueryNamedPropertiesResopnse.IdCount; i++)
                        {
                            // If any Named properties with a Kind of 0x00 actually presents in the response, isVerifyR134 should be set to false.
                            if (ropQueryNamedPropertiesResopnse.PropertyNames[i].Kind == 0x00)
                            {
                                // Add the debug information
                                Site.Log.Add(
                                    LogEntryKind.Debug,
                                    "Verify MS-OXCPRPT_R134, QueryFlags is {0}, PropertyName.Kind in the response is {1}",
                                    queryFlags,
                                    ropQueryNamedPropertiesResopnse.PropertyNames[i].Kind);

                                isVerifyR134 = false;
                                break;
                            }
                        }

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR134,
                            134,
                            @"[In RopQueryNamedProperties ROP Request Buffer] QueryFlags: NoIds:Named properties that have a long ID (LID) MUST NOT be included in the response.");
                    }

                    // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R141
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)ropQueryNamedPropertiesResopnse.PropertyNames.Length,
                        ropQueryNamedPropertiesResopnse.IdCount,
                        141,
                        @"[In IdCount] IdCount: 2 byte integer.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R14101");

                    // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R14101
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)ropQueryNamedPropertiesResopnse.PropertyIds.Length,
                        ropQueryNamedPropertiesResopnse.IdCount,
                        14101,
                        @"[In RopQueryNamedProperties ROP Response Buffer] IdCount: An integer that specifies the number of elements contained in the PropertyIds field.");

                    Site.CaptureRequirementIfAreEqual<ushort>(
                        (ushort)ropQueryNamedPropertiesResopnse.PropertyIds.Length,
                        ropQueryNamedPropertiesResopnse.IdCount,
                        14202,
                        @"[In RopQueryNamedProperties ROP Response Buffer] propertyIds: The number of integers contained in the array MUST equal the value specified in the IdCount field. ");

                    bool isVerifyR238 = true;
                    foreach (PropertyId id in ropQueryNamedPropertiesResopnse.PropertyIds)
                    {
                        if ((id.ID & 0x8000) != 0x8000)
                        {
                            isVerifyR238 = false;
                            break;
                        }
                    }

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR238,
                        238,
                        @"[In RopGetNamesFromPropertyIds ROP] Property IDs for named properties are identified by having their most significant bit set (0x8000). ");

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR238,
                        41405,
                        @"[In Getting Property IDs for Named Properties] The property ID of a named property has the most significant bit set (0x8000).");
                }

                if (queryFlags == QueryFlags.OtherValue)
                {
                    Site.CaptureRequirementIfAreEqual<uint>(
                        0x00,
                        ropQueryNamedPropertiesResopnse.ReturnValue,
                        87802,
                        @"[In Processing RopQueryNamedProperties] No matter what the invalid bits that are set in the QueryFlags field of the ROP request buffer, the server return the same.");
                }

                // propertyGuid is not null means that the client specified this field in the request buffer, this field is present.
                if (propertyGuid != null)
                {
                    // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R138
                    bool isVerifyR138 = true;

                    // Examine every PropertyName's Guid return by server.
                    for (int i = 0; i < ropQueryNamedPropertiesResopnse.PropertyNames.Length; i++)
                    {
                        // Check whether the Guid returned in response buffer matches the PropertyGuid specified in request buffer.
                        if (!Common.CompareByteArray(ropQueryNamedPropertiesResopnse.PropertyNames[i].Guid, propertyGuid))
                        {
                            // Add the debug information
                            Site.Log.Add(
                                LogEntryKind.Debug,
                                "Verify MS-OXCPRPT_R138,PropertyGuid in request buffer is {0}, Guid in response buffer is {1}",
                                propertyGuid.ToString(),
                                ropQueryNamedPropertiesResopnse.PropertyNames[i].Guid.ToString());

                            isVerifyR138 = false;
                            break;
                        }
                    }

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR138,
                        138,
                        @"[In RopQueryNamedProperties ROP Request Buffer] PropertyGuid: A GUID that specifies the property set of properties to be returned. If this field is present, only named properties with a property set matching the GUID are returned in a successful response.");
                }

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R145
                // The order has been verified when de-serializing the related structures and represented by the variable isOrdered.
                // Whether the named properties match the PropertyGUID specified in the request has been verified in MS-OXCPRPT_R138 above.
                Site.CaptureRequirementIfIsTrue(
                    isOrdered,
                    145,
                    @"[In RopQueryNamedProperties ROP Response Buffer] PropertyNames:The entries in this list MUST match the order of the entries in the PropertyIds field, and the number of entries MUST be equal.");
            }
        }
        #endregion

        #region RopCopyProperties
        /// <summary>
        /// Verify the RopCopyProperties operation related requirements.
        /// </summary>
        /// <param name="ropCopyPropertiesResponse">The RopCopyProperties response buffer structure.</param>
        /// <param name="objectType">Indicates which object type the RopCopyProperties operation is acting on.</param>
        /// <param name="copyFlags">CopyFlags parameter in request of RopCopyProperties.</param>
        /// <param name="isPropertyFound">Indicates the property is found or not.</param>
        private void VerifyRopCopyProperties(
            RopCopyPropertiesResponse ropCopyPropertiesResponse,
            ServerObjectType objectType,
            CopyFlags copyFlags,
            bool isPropertyFound)
        {
            // Since the RopCopyProperties ROP response was parsed successfully, MS-OXCPRPT_R50003 can be captured directly.
            Site.CaptureRequirement(
                50003,
                @"[In Processing RopCopyProperties] The server responds with a RopCopyProperties ROP response buffer.");

            if (ropCopyPropertiesResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                  16801,
                  @"[In RopCopyProperties ROP Response Buffer] DestHandleIndex: 4 bytes integer.");

                if (objectType == ServerObjectType.Message)
                {
                    // If the object type this operation acting on is Message object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        14901,
                        @"[In RopCopyProperties ROP] This operation [RopCopyProperties ROP] is valid on Message objects.");
                }

                if (objectType == ServerObjectType.Folder)
                {
                    // If the object type this operation acting on is Folder object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        14902,
                        @"[In RopCopyProperties ROP] This operation [RopCopyProperties ROP] is valid on Folder objects.");
                }

                if (objectType == ServerObjectType.Attachment)
                {
                    // If the object type this operation acting on is Attachment object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        14903,
                        @"[In RopCopyProperties ROP] This operation [RopCopyProperties ROP] is valid on Attachment objects.");
                }

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                165,
                @"[In RopCopyProperties ROP Response Buffer] propertyProblemCount: 2 bytes integer.");

                Site.CaptureRequirementIfAreEqual<ushort>(
                      ropCopyPropertiesResponse.PropertyProblemCount,
                      (ushort)ropCopyPropertiesResponse.PropertyProblems.Length,
                      16502,
                      @"[In RopCopyProperties ROP Response Buffer] propertyProblemCount: An integer that specifies the number of elements contained in the PropertyProblems field. ");

                if (ropCopyPropertiesResponse.PropertyProblemCount != 0)
                {
                    // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                    Site.CaptureRequirement(
                        167,
                        @"[In RopCopyProperties ROP Response Buffer] PropertyProblems (variable): An array of PropertyProblem structures ([MS-OXCDATA] section 2.7).");

                    for (int counter = 0; counter < ropCopyPropertiesResponse.PropertyProblems.Length; counter++)
                    {
                        this.VerifyPropertyProblemSturctureInCDATA();
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                   isPropertyFound,
                   500,
                   @"[In Processing RopCopyProperties] The server MUST copy or move the properties specified from the source object to the destination object.");
            }

            if (ropCopyPropertiesResponse.ReturnValue == 0x80070057)
            {
                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                    "MS-OXCDATA",
                      904,
                     @"[In Error Codes] The numeric value (hex) for error code InvalidParameter is 0x80070057, %x57.00.07.80.");
            }

            if (ropCopyPropertiesResponse.ReturnValue == 0x80040102)
            {
                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                    "MS-OXCDATA",
                      928,
                     @"[In Error Codes] The numeric value (hex) for error code NotSupported is 0x80040102, %x02.01.04.80.");
            }

            if (ropCopyPropertiesResponse.ReturnValue.Equals((uint)CPRPTErrorCode.NullDestinationObject))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R168");

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R168
                // The value of DestHandleIndex is not 0 means this field is present.
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    0,
                    ropCopyPropertiesResponse.DestHandleIndex,
                    168,
                    @"[In RopCopyProperties ROP Response Buffer] DestHandleIndex: The DestHandleIndex field MUST be set to the value of the DestHandleIndex field of the ROP request buffer. ");

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R737
                // The value of PropertyProblemCount is 0 means this field is not present.
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0,
                    ropCopyPropertiesResponse.PropertyProblemCount,
                    737,
                    @"[In RopCopyProperties ROP Response Buffer] PropertyProblemCount: This field MUST NOT be present if the ReturnValue field is set to NullDestinationObject (0x00000503).");

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R738
                // PropertyProblems is null means this field is not present.
                Site.CaptureRequirementIfIsNull(
                    ropCopyPropertiesResponse.PropertyProblems,
                    738,
                    @"[In RopCopyProperties ROP Response Buffer] propertyProblems: This field MUST NOT be present if the ReturnValue field is set to NullDestinationObject (0x00000503).");
            }
            else
            {
                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R170
                // The value of DestHandleIndex is 0 means this field is not present.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    ropCopyPropertiesResponse.DestHandleIndex,
                    170,
                    @"[In RopCopyProperties ROP Response Buffer] DestHandleIndex: The DestHandleIndex field MUST NOT be present if the ReturnValue field is set to any value other than NullDestinationObject (0x00000503).");
            }

            if ((objectType == ServerObjectType.Folder) && (copyFlags == CopyFlags.Move))
            {
                if (Common.IsRequirementEnabled(882, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<uint>(
                        (uint)CPRPTErrorCode.NotSupported,
                        ropCopyPropertiesResponse.ReturnValue,
                        882,
                        @"[In Processing RopCopyProperties] Implementation does return NotSupported error, if the original object is a Folder object and the CopyFlags field has the Move flag set.(Microsoft Exchange Server 2007 and above follow this behavior)");
                }
            }
        }
        #endregion

        #region VerifyRopCopyToExcludeElement
        /// <summary>
        /// Verify if RopCopyTo method copies the excluded value.
        /// </summary>
        /// <param name="isExcludedPropertyCopied">Indicates if the excluded value has been copied.</param>
        private void VerifyRopCopyToExcludeElement(bool isExcludedPropertyCopied)
        {
            Site.CaptureRequirementIfIsFalse(
                isExcludedPropertyCopied,
                171,
                @"[In RopCopyTo ROP] The RopCopyTo ROP ([MS-OXCROPS] section 2.2.8.12) is used to copy or move all but a specified few properties from a source object to a destination object.");
 
            Site.CaptureRequirementIfIsFalse(
                isExcludedPropertyCopied,
                191,
                @"[In RopCopyTo ROP Request Buffer] ExcludedTags: Each structure contains the property tag of a property that MUST NOT be copied or moved as part of this operation.");
        }
        #endregion

        #region RopCopyTo
        /// <summary>
        /// Verify the RopCopyTo operation related requirements.
        /// </summary>
        /// <param name="ropCopyToResponse">The RopCopyTo response buffer structure.</param>
        /// <param name="copyToCondition">The condition to generate the corresponding error codes.</param>
        /// <param name="objectType">Indicates which object type the RopCopyProperties operation is acting on.</param>
        /// <param name="copyFlags">Indicates the copy flags.</param>
        private void VerifyRopCopyTo(RopCopyToResponse ropCopyToResponse, CopyToCondition copyToCondition, ServerObjectType objectType, CopyFlags copyFlags)
        {
            // Since the RopCopyTo ROP response was parsed successfully, MS-OXCPRPT_R50702 can be captured directly.
            Site.CaptureRequirement(
                50702,
                @"[In Processing RopCopyTo] The server responds with a RopCopyTo ROP response buffer.");

            if (ropCopyToResponse.ReturnValue == (uint)CPRPTErrorCode.NotSupported)
            {
                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                    "MS-OXCDATA",
                     928,
                     @"[In Error Codes] The numeric value (hex) for error code NotSupported is 0x80040102, %x02.01.04.80.");
            }

            if (ropCopyToResponse.ReturnValue == (uint)CPRPTErrorCode.None)
            {
                if (objectType == ServerObjectType.Message)
                {
                    // If the object type this operation acting on is Message object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        17201,
                        @"[In RopCopyTo ROP] This operation [RopCopyTo ROP] is valid on Message objects.");
                }

                if (objectType == ServerObjectType.Folder)
                {
                    if (copyFlags == CopyFlags.Move)
                    {
                        if (Common.IsRequirementEnabled(5070511, this.Site))
                        {
                            // If the implementation return a successful response, it means it doesn't return NotSupported error. So it can be capture directly.
                            Site.CaptureRequirement(
                                5070511,
                                @"[In Processing RopCopyTo] Implementation does not return a NotSupported error (0x80040102), if the original object is a Folder object and the CopyFlags field has the Move flag set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                        }
                    }

                    // If the object type this operation acting on is Folder object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        17203,
                        @"[In RopCopyTo ROP] This operation [RopCopyTo ROP] is valid on Folder objects.");
                }

                if (objectType == ServerObjectType.Attachment)
                {
                    // If the object type this operation acting on is Attachment object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        17202,
                        @"[In RopCopyTo ROP] This operation [RopCopyTo ROP] is valid on Attachment objects.");
                }

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                   192,
                   @"[In RopCopyTo ROP Response Buffer] propertyProblemCount: 2 bytes integer.");

                if (ropCopyToResponse.PropertyProblemCount != 0)
                {
                    for (int counter = 0; counter < ropCopyToResponse.PropertyProblems.Length; counter++) 
                    {
                        this.VerifyPropertyProblemSturctureInCDATA();
                    }

                    Site.CaptureRequirementIfAreEqual<ushort>(
                        ropCopyToResponse.PropertyProblemCount,
                        (ushort)ropCopyToResponse.PropertyProblems.Length,
                        19201,
                        @"[In RopCopyTo ROP Response Buffer] propertyProblemCount: An integer that specifies the number of elements contained in the PropertyProblems field. ");

                    // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                    Site.CaptureRequirement(
                        193,
                        @"[In RopCopyTo ROP Response Buffer] propertyProblems (variable): An array of PropertyProblem structures ([MS-OXCDATA] section 2.7).");
                }

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                    19601,
                    @"[In RopCopyTo ROP Response Buffer] DestHandleIndex (4 bytes):  An integer.");
            }

            if (ropCopyToResponse.ReturnValue.Equals((uint)CPRPTErrorCode.NullDestinationObject))
            {
                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R195
                // The value of DestHandleIndex is not 0 means this field is present.
                Site.CaptureRequirementIfAreEqual<uint>(
                    1,
                    ropCopyToResponse.DestHandleIndex,
                    196,
                    @"[In RopCopyTo ROP Response Buffer] DestHandleIndex: The DestHandleIndex field MUST be set to the value of the DestHandleIndex field of the ROP request buffer");

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R739
                // The value of PropertyProblemCount is 0 means this field is not present.
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0,
                    ropCopyToResponse.PropertyProblemCount,
                    739,
                    @"[In RopCopyTo ROP Response Buffer] PropertyProblemCount: This field MUST NOT be present if the ReturnValue field is set to NullDestinationObject (0x00000503).");

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R740
                // PropertyProblems is null means this field is not present.
                Site.CaptureRequirementIfIsNull(
                    ropCopyToResponse.PropertyProblems,
                    740,
                    @"[In RopCopyTo ROP Response Buffer] PropertyProblems: This field MUST NOT be present if the ReturnValue field is set to NullDestinationObject (0x00000503).");
            }
            else
            {
                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R197
                // The value of DestHandleIndex is 0 means this field is not present.
                Site.CaptureRequirementIfAreEqual<uint>(
                     0,
                     ropCopyToResponse.DestHandleIndex,
                    197,
                    @"[In RopCopyTo ROP Response Buffer] DestHandleIndex: The DestHandleIndex field MUST NOT be present if the ReturnValue field is set to any value other than NullDestinationObject (0x00000503).");
            }

            if (copyToCondition == CopyToCondition.SourceContainsDest)
            {
                if (objectType == ServerObjectType.Folder)
                {
                    if (Common.IsRequirementEnabled(89603, this.Site))
                    {
                        Site.CaptureRequirementIfAreEqual<uint>(
                            (uint)CPRPTErrorCode.FolderCycle,
                            ropCopyToResponse.ReturnValue,
                            89603,
                            @"[In Processing RopCopyTo] Implementation does return error code ""0x8004060B"" with name ""FolderCycle"" when The source folder contains the destination folder.(Microsoft Exchange Server 2007 and above follow this behavior)");
                    }
                }
            }
            else if (copyToCondition == CopyToCondition.SourceDestNotCompatible)
            {
                // CopyToCondition.SourceDestNotCompatible means source object and destination object are not compatible with each other for the copy operation.
                bool isR890Satisfied = ropCopyToResponse.ReturnValue == (uint)CPRPTErrorCode.NotSupported;

                if (Common.IsRequirementEnabled(890, this.Site))
                {
                    // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R890
                    Site.CaptureRequirementIfIsTrue(
                        isR890Satisfied,
                        890,
                        @"[In Processing RopCopyTo] Implementation does return error code ""0x80040102"" with name ""NotSupported"",  when ""The source object and destination object are not compatible with each other for the copy operation."" (Microsoft Exchange Server 2007 and above follow this behavior)");

                    Site.CaptureRequirementIfIsTrue(
                        isR890Satisfied,
                        150,
                        @"[In RopCopyProperties ROP] Also, the source and destination object MUST be of the same type.");
                }
            }
            else if (copyToCondition == CopyToCondition.SourceDestHasSubObjWithSameDisplayName)
            {
                // CopyToCondition.SourceDestHasSubObjWithSameDisplayName means a sub-object cannot be copied because there is already a sub-object existing 
                // in the destination object with the same display name (PidTagDisplayName) as the sub-object to be copied.
                if (Common.IsRequirementEnabled(899, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R899, the return error code value is: {0}", ropCopyToResponse.ReturnValue);

                    // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R899
                    Site.CaptureRequirementIfAreEqual<uint>(
                        (uint)CPRPTErrorCode.CollidingNames,
                        ropCopyToResponse.ReturnValue,
                        899,
                        @"[In Processing RopCopyTo] Implementation does return error code ""0x80040604"" with name ""CollidingNames"", when a subobject cannot be copied because there is already a subobject existing in the destination object with the same display name, which is specified in the PidTagDisplayName property ([MS-OXCFOLD] section 2.2.2.2.2.4), as the subobject to be copied.(Microsoft Exchange Server 2007 and above follow this behavior)");
                }
            }
            else if (copyToCondition == CopyToCondition.SourceMessageContainsDestMessage && Common.IsRequirementEnabled(89601, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R89601, the return error code value is: {0}", ropCopyToResponse.ReturnValue);

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R89601
                Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)CPRPTErrorCode.MessageCycle,
                    ropCopyToResponse.ReturnValue,
                    89601,
                    @"[In Appendix A: Product Behavior] Implementation does return error code ""0x00000504"" with name ""MessageCycle"" when the source message directly contains the destination message. (Exchange 2007 and above follow this behavior.)");
            }
            else if (copyToCondition == CopyToCondition.SourceMessageIndirectlyContainsDestMessage && Common.IsRequirementEnabled(89604, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R89604, the return error code value is: {0}", ropCopyToResponse.ReturnValue);

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R89604
                Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)CPRPTErrorCode.MessageCycle,
                    ropCopyToResponse.ReturnValue,
                    89604,
                    @"[In Appendix A: Product Behavior] Implementation does return error code ""0x00000504"" with name ""MessageCycle"" when the source message indirectly contains the destination message. (Exchange 2007 and above follow this behavior.)");            
            }
        }

        /// <summary>
        /// Verify the RopCopyTo operation about public folder related requirements.
        /// </summary>
        /// <param name="ropCopyToResponse">The RopCopyTo response buffer structure.</param>
        private void VerifyRopCopyToForPublicFolder(RopCopyToResponse ropCopyToResponse)
        {
            if (Common.IsRequirementEnabled(5070513, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R5070513, the return error code value is: {0}", ropCopyToResponse.ReturnValue);

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R5070513
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    (uint)CPRPTErrorCode.NotSupported,
                    ropCopyToResponse.ReturnValue,
                    5070513,
                    @"[In Appendix A: Product Behavior] Implementation does not return a NotSupported error (0x80040102) if the RopCopyTo ROP request specifies a public folder as either the source object or the destination object. (<10> Section 3.2.5.8: Exchange 2007 and Exchange 2010 follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(5070514, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R5070514, the return error code value is: {0}", ropCopyToResponse.ReturnValue);

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R5070514
                Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)CPRPTErrorCode.NotSupported,
                    ropCopyToResponse.ReturnValue,
                    5070514,
                    @"[In Appendix A: Product Behavior] Implementation does return a NotSupported error (0x80040102) if the RopCopyTo ROP request specifies a public folder as either the source object or the destination object. (<10> Section 3.2.5.8: Exchange 2013 follows this behavior.)");
            }
        }
        #endregion

        #region RopGetPropertyIdsFromNames
        /// <summary>
        /// Verify the RopCopyTo operation related requirements.
        /// </summary>
        /// <param name="propertyNameCount">The PropertyNameCount parameter in the request buffer.</param>
        /// <param name="propertyNames">A list of PropertyName in the request buffer.</param>
        /// <param name="flags">The Flags set by client in the request buffer.</param>
        /// <param name="ropGetPropertyIdsFromNamesResponse">The RopCopyTo response buffer structure.</param>
        /// <param name="isOrdered">A boolean value which indicates whether the order of elements in request/response buffer is right.</param>
        /// <param name="objectType">Indicates which object type the RopGetPropertyIdsFromNames operation is acting on.</param>
        /// <param name="isPropertyNameExisting">A boolean value indicates whether the queried PropertyNames exists on server.</param>
        /// <param name="isNamesAndIdsMapping">A boolean value indicates whether the Names and the returned Ids are mapping.</param>
        private void VerifyRopGetPropertyIdsFromNames(
            ushort propertyNameCount,
            PropertyName[] propertyNames,
            byte flags,
            RopGetPropertyIdsFromNamesResponse ropGetPropertyIdsFromNamesResponse,
            bool isOrdered,
            ServerObjectType objectType,
            bool isPropertyNameExisting,
            bool isNamesAndIdsMapping)
        {
            // Since the RopGetPropertyIdsFromNames ROP response was parsed successfully, MS-OXCPRPT_R52502 can be captured directly.
            Site.CaptureRequirement(
                52502,
                @"[In Processing RopGetPropertyIdsFromNames] The server responds with a RopGetPropertyIdsFromNames ROP response buffer.");

            // Since the RopGetPropertyIdsFromNames ROP response was parsed successfully, MS-OXCPRPT_R41404 can be captured directly.
            Site.CaptureRequirement(
               41404,
               @"[In Getting Property IDs for Named Properties] The client provides a property set and an identifier in either integer or string form, and the server returns the property ID, which the client uses for any operations performed on that property.");

            if (propertyNames[0].Kind == 0x00)
            {
                Site.CaptureRequirementIfAreNotEqual<ushort>(
                    0x0000,
                    ropGetPropertyIdsFromNamesResponse.PropertyIds[0].ID,
                    225,
                    @"[In RopGetPropertyIdsFromNames ROP Response Buffer] PropertyIds: Reasons a name couldn't be mapped include: Use of the PS_MAPI namespace and not specifying 0x00 for the Kind field of the PropertyName structure ([MS-OXCDATA] section 2.6.1).");
            }

            if (flags != 0x02 && isPropertyNameExisting)
            {
                Site.CaptureRequirementIfAreNotEqual<ushort>(
                   0x0000,
                   ropGetPropertyIdsFromNamesResponse.PropertyIds[0].ID,
                   226,
                   @"[In RopGetPropertyIdsFromNames ROP Response Buffer] PropertyIds: Reasons a name couldn't be mapped include:The name wasn't found in the mapping table and the Flags field of the ROP request buffer was not set to 0x02.");
            }

            if (ropGetPropertyIdsFromNamesResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                if (objectType == ServerObjectType.Message)
                {
                    // If the object type this operation acting on is Message object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        20701,
                        @"[In RopGetPropertyIdsFromNames ROP] This operation [RopGetPropertyIdsFromNames ROP] is valid on Message objects.");
                }

                if (objectType == ServerObjectType.Attachment)
                {
                    // If the object type this operation acting on is Folder object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        20702,
                        @"[In RopGetPropertyIdsFromNames ROP] This operation [RopGetPropertyIdsFromNames ROP] is valid on Attachment objects.");
                }

                if (objectType == ServerObjectType.Folder)
                {
                    // If the object type this operation acting on is Attachment object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        20703,
                        @"[In RopGetPropertyIdsFromNames ROP] This operation [RopGetPropertyIdsFromNames ROP] is valid on Folder objects.");
                }

                if (objectType == ServerObjectType.Logon)
                {
                    // If the object type this operation acting on is Logon object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        20704,
                        @"[In RopGetPropertyIdsFromNames ROP] This operation [RopGetPropertyIdsFromNames ROP] is valid on Logon objects.");
                }

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                    215,
                    @"[In RopGetPropertyIdsFromNames ROP Response Buffer] PropertyIdCount: 2 bytes integer.");

                Site.CaptureRequirementIfAreEqual<ushort>(
                      (ushort)ropGetPropertyIdsFromNamesResponse.PropertyIds.Length,
                      ropGetPropertyIdsFromNamesResponse.PropertyIdCount,
                      21502,
                      @"[In RopGetPropertyIdsFromNames ROP Response Buffer] PropertyIdCount: An integer that specifies the number of property IDs contained in the PropertyIds field. ");

                if (propertyNameCount != 0)
                {
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        propertyNameCount,
                        ropGetPropertyIdsFromNamesResponse.PropertyIdCount,
                        216,
                        @"[In RopGetPropertyIdsFromNames ROP Response Buffer] PropertyIdCount: The value of this field MUST be equal to the value of the PropertyNameCount field of the ROP request buffer unless the value of the PropertyNameCount field is zero.");
                }
            }

            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R218
            bool isVerifyR218 = true;
            for (int i = 0; i < ropGetPropertyIdsFromNamesResponse.PropertyIds.Length; i++)
            {
                // Check that every element's size in PropertyIds is 16-bit 
                if (Marshal.SizeOf(ropGetPropertyIdsFromNamesResponse.PropertyIds[i]) != 2)
                {
                    // Add the debug information
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "Verify MS-OXCPRPT_R218, The actual length of PropertyId is {0}",
                        Marshal.SizeOf(ropGetPropertyIdsFromNamesResponse.PropertyIds[i]));
                    isVerifyR218 = false;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR218,
                218,
                @"[In RopGetPropertyIdsFromNames ROP Response Buffer] propertyIds (variable): An array of 16-bit integers. ");

            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R222
            // The order has been verified when de-serializing the related structures and represented by the variable isOrdered
            Site.CaptureRequirementIfIsTrue(
                isOrdered,
                222,
                @"[In RopGetPropertyIdsFromNames ROP Response Buffer] PropertyIds: The order of property IDs in this array MUST match the order of the named properties specified in the PropertyNames field of the ROP request buffer.");

            Site.CaptureRequirementIfIsTrue(
               isOrdered,
               21801,
               @"[In RopGetPropertyIdsFromNames ROP Response Buffer] PropertyIds: Each integer is a property ID that is mapped from a named property that is specified in the PropertyNames field of the ROP request buffer");

            if (!((propertyNameCount == 0) && (objectType == ServerObjectType.Logon)))
            {
                // !isPropertyNameExisting means that the propertyNames doesn't exist on server.
                // The property flag doesn't equal 0x02 means the Create bit in the Flags parameter is not set.
                if (!isPropertyNameExisting && (flags != 0x02))
                {
                    for (int i = 0; i < ropGetPropertyIdsFromNamesResponse.PropertyIds.Length; i++)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R532");

                        // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R532
                        // The PropertyNames don't exist on server means the PropertyNames requested are the unfound rows.
                        // The logon user is Administrator and the test suite will not reach the server-imposed limit on property ID 
                        // mappings, so the user always has permission to create new entries and the server-imposed limit is not reached.
                        // Since the Flags parameter doesn't have the Create flag bit set, so the "unless" condition will be still false.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0000,
                            ropGetPropertyIdsFromNamesResponse.PropertyIds[i].ID,
                            532,
                            @"[In Processing RopGetPropertyIdsFromNames] [If the PropertyNameCount parameter is zero, and the RopGetPropertyIdsFromNames is acting on a Logon object, the server must enumerate all PropertyNames associated with property IDs] Otherwise, the server MUST, for each entry in the PropertyNames field of the ROP request buffer, follow this procedure: 3. For unfound rows, the returned property ID MUST be 0x0000 unless all of the following conditions are true:
                            1.The Flags field of the ROP request buffer is set to 0x02.
                            2. The user has permission to create new entries.
                            3. The server-imposed limit on property ID mappings specified later in this section hasn't yet been reached.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R220");

                        // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R220
                        // The PropertyNames don't exist on server and the Flags parameter doesn't have the Create flag bit set means 
                        // the names could not be mapped.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0000,
                            ropGetPropertyIdsFromNamesResponse.PropertyIds[i].ID,
                            220,
                            @"[In RopGetPropertyIdsFromNames ROP Response Buffer] PropertyIds: If a named property cannot be mapped, the associated entry in the PropertyIds field MUST be 0x0000.");
                    }
                }

                // !isPropertyNameExisting means the propertyNames don't exist on server
                // flags equals to 0x02 means the Flags parameter has the Create flag bit set.
                // The logon user is Administrator and the test suite will not reach the server-imposed limit on property ID 
                // mappings, so the user always has permission to create new entries and the server-imposed limit is not reached.
                // If all of these above condition are met, it means the returned id is newly assigned.
                if (!isPropertyNameExisting && (flags == 0x02))
                {
                    for (int i = 0; i < ropGetPropertyIdsFromNamesResponse.PropertyIds.Length; i++)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R534");

                        // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R534
                        Site.CaptureRequirementIfAreNotEqual<ushort>(
                            0xFFFF,
                            ropGetPropertyIdsFromNamesResponse.PropertyIds[i].ID,
                            534,
                            @"[In Processing RopGetPropertyIdsFromNames] [If the PropertyNameCount parameter is zero, and the RopGetPropertyIdsFromNames is acting on a Logon object, the server must enumerate all PropertyNames associated with property IDs] Otherwise, the server MUST, for each entry in the PropertyNames field of the ROP request buffer, follow this procedure: 4. [If the above conditions in step three are all met, a new property ID is registered for the named property.] The newly assigned property ID MUST NOT be equal to 0xFFFF.");

                        // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R535
                        Site.CaptureRequirementIfIsTrue(
                            ropGetPropertyIdsFromNamesResponse.PropertyIds[i].ID > 0x8000,
                            535,
                            @"[In Processing RopGetPropertyIdsFromNames] [If the PropertyNameCount parameter is zero, and the RopGetPropertyIdsFromNames is acting on a Logon object, the server must enumerate all PropertyNames associated with property IDs] Otherwise, the server MUST, for each entry in the PropertyNames field of the ROP request buffer, follow this procedure: 4. [If the above conditions in step three are all met, a new property ID is registered for the named property.]  The newly assigned property ID MUST be greater than 0x8000.");
                    }
                }

                // When the PropertyNames requested exist on the server, then the server can find the property ID associated with the PropertyName.
                if (isPropertyNameExisting)
                {
                    // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R531
                    // Whether the PropertyNames queried and the returned Ids are mapping has been verified and represented as the Boolean value isNameAndIdMapping.
                    Site.CaptureRequirementIfIsTrue(
                        isNamesAndIdsMapping,
                        531,
                        @"[In Processing RopGetPropertyIdsFromNames] [If the PropertyNameCount parameter is zero, and the RopGetPropertyIdsFromNames is acting on a Logon object, the server must enumerate all PropertyNames associated with property IDs] Otherwise, the server MUST, for each entry in the PropertyNames field of the ROP request buffer, follow this procedure: 2. Find the property ID registered for the named property that matches the PropertyName structure as follows: If the Kind field is set to 0x00, the named property has a LID that matches the value of the LID field of the PropertyName structure.");

                    for (int i = 0; i < ropGetPropertyIdsFromNamesResponse.PropertyIds.Length; i++)
                    {
                        if (Common.CompareByteArray(propertyNames[i].Guid, this.valuePSMAPI))
                        {
                            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R642
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)propertyNames[i].LID,
                                ropGetPropertyIdsFromNamesResponse.PropertyIds[i].ID,
                                642,
                                @"[In Processing RopGetPropertyIdsFromNames] [If the PropertyNameCount parameter is zero, and the RopGetPropertyIdsFromNames is acting on a Logon object, the server must enumerate all PropertyNames associated with property IDs] Otherwise, the server MUST, for each entry in the PropertyNames field of the ROP request buffer, follow this procedure: 1.  If the GUID field of the PropertyName structure ([MS-OXCDATA] section 2.6.1) in the ROP request buffer specifies the PS_MAPI property set, the returned property ID is obtained from the LID field.");
                        }
                    }
                }
            }

            // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
            Site.CaptureRequirement(
                198,
                @"[In RopGetPropertyIdsFromNames ROP] The RopGetPropertyIdsFromNames ROP ([MS-OXCROPS] section 2.2.8.1) maps abstract, client-defined named properties to concrete 16-bit property IDs (of which 15 bits are significant).");
        }
        #endregion

        #region RopGetNamesFromPropertyIds
        /// <summary>
        /// Verify the RopGetNamesFromPropertyIds operation related requirements.
        /// </summary>
        /// <param name="propertyIdCount">The PropertyIdCount parameter in the request buffer which indicates the number of PropertyId.</param>
        /// <param name="propertyIds">A list of PropertyId in the request buffer.</param>
        /// <param name="ropGetNamesFromPropertyIdsResponse">The RopGetNamesFromPropertyIds response buffer structure.</param>
        /// <param name="objectType">Indicates which object type the RopGetNamesFromPropertyIds operation is acting on.</param>
        private void VerifyRopGetNamesFromPropertyIds(
            ushort propertyIdCount,
            PropertyId[] propertyIds,
            RopGetNamesFromPropertyIdsResponse ropGetNamesFromPropertyIdsResponse,
            ServerObjectType objectType)
        {
            // Since the RopGetNamesFromPropertyIds ROP response was parsed successfully, MS-OXCPRPT_R51802 can be captured directly.
            Site.CaptureRequirement(
                51802,
                @"[In Processing RopGetNamesFromPropertyIds] The server responds with a RopGetNamesFromPropertyIds ROP response buffer.");

            if (ropGetNamesFromPropertyIdsResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                if (objectType == ServerObjectType.Message)
                {
                    // If the object type this operation acting on is Message object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        23701,
                        @"[In RopGetNamesFromPropertyIds ROP] This operation [RopGetNamesFromPropertyIds ROP] is valid on Message objects.");
                }

                if (objectType == ServerObjectType.Folder)
                {
                    // If the object type this operation acting on is Folder object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        23703,
                        @"[In RopGetNamesFromPropertyIds ROP] This operation [RopGetNamesFromPropertyIds ROP] is valid on Folder objects.");
                }

                if (objectType == ServerObjectType.Attachment)
                {
                    // If the object type this operation acting on is Attachment object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        23702,
                        @"[In RopGetNamesFromPropertyIds ROP] This operation [RopGetNamesFromPropertyIds ROP] is valid on Attachment objects.");
                }

                if (objectType == ServerObjectType.Logon)
                {
                    // If the object type this operation acting on is Logon object and this operation is performed successfully, then the following requirement can be captured.
                    Site.CaptureRequirement(
                        23704,
                        @"[In RopGetNamesFromPropertyIds ROP] This operation [RopGetNamesFromPropertyIds ROP] is valid on Logon objects.");
                }

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                  248,
                  @"[In RopGetNamesFromPropertyIds ROP Response Buffer] PropertyNameCount: 2 bytes integer.");

                if (ropGetNamesFromPropertyIdsResponse.PropertyNameCount != 0)
                {
                    // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                    Site.CaptureRequirement(
                          250,
                          @"[In RopGetNamesFromPropertyIds ROP Response Buffer] PropertyNames (variable): An array of PropertyName structures ([MS-OXCDATA] section 2.6).");
                }

                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXCPRPT_R658, PropertyNameCount is {0}, ProertyNames is {1}",
                    ropGetNamesFromPropertyIdsResponse.PropertyNameCount,
                    ropGetNamesFromPropertyIdsResponse.PropertyNames);
            }

            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R249
            Site.CaptureRequirementIfAreEqual<ushort>(
                propertyIdCount,
                ropGetNamesFromPropertyIdsResponse.PropertyNameCount,
                249,
                @"[In RopGetNamesFromPropertyIds ROP Response Buffer] PropertyNameCount: This value MUST be equal to the value of the PropertyIdCount field of the ROP request buffer. ");

            Site.CaptureRequirementIfAreEqual<ushort>(
               propertyIdCount,
               (ushort)ropGetNamesFromPropertyIdsResponse.PropertyNames.Length,
               24802,
               @"[In RopGetNamesFromPropertyIds ROP Response Buffer] propertyNameCount: An integer that specifies the number of structures contained in the PropertyNames field. ");

            // Examine each property ID in the request message and each associated PropertyName in the response message.The order of the ID and Name matches.
            for (int i = 0; i < propertyIds.Length; i++)
            {
                if (propertyIds[i].ID < 0x8000)
                {
                    // Add the debug information
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "Verify MS-OXCPRPT_R518,The PropertyId requested is {0}, The GUID response in PropertyName is {1}, The expected value of GUID is {2}",
                        propertyIds[i].ID,
                        ropGetNamesFromPropertyIdsResponse.PropertyNames[i].Guid.ToString(),
                        this.valuePSMAPI.ToString());

                    // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R518
                    // Check whether the PropertyName's GUID is PS-MAPI
                    bool isVerifyR518 = Common.CompareByteArray(ropGetNamesFromPropertyIdsResponse.PropertyNames[i].Guid, this.valuePSMAPI);

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR518,
                        518,
                        @"[In Processing RopGetNamesFromPropertyIds] For each property ID in the PropertyIds field of the ROP request buffer, the server MUST perform the following: If the property ID is less than 0x8000, the associated PropertyName structure ([MS-OXCDATA] section 2.6.1) contained in the PropertyNames field of the ROP response buffer MUST be composed as follows: The structure's GUID field is set to the PS_MAPI property set ([MS-OXPROPS] section 1.3.2).");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R519");

                    Site.CaptureRequirementIfAreEqual<byte>(
                        0x00,
                        ropGetNamesFromPropertyIdsResponse.PropertyNames[i].Kind,
                        519,
                        @"[In Processing RopGetNamesFromPropertyIds] For each property ID in the PropertyIds field of the ROP request buffer, the server MUST perform the following: If the property ID is less than 0x8000, the associated PropertyName structure ([MS-OXCDATA] section 2.6.1) contained in the PropertyNames field of the ROP response buffer MUST be composed as follows:: The structure's Kind field is set to 0x00.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPRPT_R520");

                    Site.CaptureRequirementIfAreEqual<uint>(
                        (uint)propertyIds[i].ID,
                        (uint)ropGetNamesFromPropertyIdsResponse.PropertyNames[i].LID,
                        520,
                        @"[In Processing RopGetNamesFromPropertyIds] For each property ID in the PropertyIds field of the ROP request buffer, the server MUST perform the following: If the property ID is less than 0x8000, the associated PropertyName structure ([MS-OXCDATA] section 2.6.1) contained in the PropertyNames field of the ROP response buffer MUST be composed as follows:: The structure's LID field is set to the property ID.");
                }
            }

            bool isCDATAR25Valid = true;
            bool isCDATAR27Valid = true;
            foreach (PropertyName propertyName in ropGetNamesFromPropertyIdsResponse.PropertyNames)
            {
                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                    "MS-OXCDATA",
                    30,
                    @"[In PropertyName Structure] Name (optional) (variable):  The value is a Unicode (UTF-16 format) string, followed by two zero bytes as terminating null characters, that identifies the property within its property set.");

                Site.CaptureRequirementIfIsNotNull(
                    propertyName.Guid,
                    "MS-OXCDATA",
                    21,
                    @"[In PropertyName Structure] GUID (16 bytes):  The GUID that identifies the property set for the named property.");

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                    "MS-OXCDATA",
                    26,
                    @"[In PropertyName Structure] LID (optional) (4 bytes):  An unsigned integer that identifies the named property within its property set.");

                if ((propertyName.LID != 0x00 && propertyName.Kind != 0x00) || ((propertyName.Kind == 0x01 || propertyName.Kind == 0x0ff) && propertyName.LID != 0x00))
                {
                    isCDATAR25Valid = false;
                }

                if ((propertyName.NameSize != null && propertyName.Kind != 0x01) || ((propertyName.Kind == 0x00 || propertyName.Kind == 0x0ff) && propertyName.NameSize != null))
                {
                    isCDATAR27Valid = false;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isCDATAR25Valid,
                "MS-OXCDATA",
                25,
                @"[In PropertyName Structure] LID (optional) (4 bytes): This field is present only if the value of the Kind field is equal to 0x00.");

            Site.CaptureRequirementIfIsTrue(
               isCDATAR27Valid,
               "MS-OXCDATA",
               27,
               @"[In PropertyName Structure] NameSize (optional) (1 byte):  This field is present only if the value of the Kind field is equal to 0x01.");

            Site.CaptureRequirementIfIsTrue(
               isCDATAR27Valid,
               "MS-OXCDATA",
               29,
               @"[In PropertyName Structure] Name (optional) (variable): This field is present only if Kind is equal to 0x01.");

            // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R521
            // The property ID values has been verified that have an associated PropertyName.
            // Server will set the PropertyName value according to the PropertyId. So if the returned PropertyNames is not null, this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                ropGetNamesFromPropertyIdsResponse.PropertyNames,
                521,
                @"[In Processing RopGetNamesFromPropertyIds] For each property ID in the PropertyIds field of the ROP request buffer, the server MUST perform the following: For property IDs that have an associated PropertyName structure, the server MUST return the PropertyName structure associated with the property ID.");

            // Since the RopGetNamesFromPropertyIds ROP response was parsed successfully, MS-OXCPRPT_R230 can be captured directly.
            Site.CaptureRequirement(
                230,
                @"[In RopGetNamesFromPropertyIds ROP] The RopGetNamesFromPropertyIds ROP ([MS-OXCROPS] section 2.2.8.2) maps concrete property IDs to abstract, client-defined named properties");
        }
        #endregion

        #region RopOpenStream
        /// <summary>
        /// Verify the RopOpenStream operation related requirements.
        /// </summary>
        /// <param name="ropOpenStreamResponse">The RopOpenStream response buffer structure.</param>
        /// <param name="objectType">Indicates which object type the RopOpenStream operation is acting on.</param>
        /// <param name="isPropertyTagExist">A boolean value indicates if the property tag exist for the object or not.</param>
        /// <param name="openFlag">Specifies OpenModeFlags for RopOpenStream.</param>
        private void VerifyRopOpenStream(RopOpenStreamResponse ropOpenStreamResponse, ServerObjectType objectType, bool isPropertyTagExist, OpenModeFlags openFlag)
        {
            if (!isPropertyTagExist && openFlag != OpenModeFlags.Create)
            {
                // Error code NotFound value is 0x8004010F
                bool isR891Satisfied = ropOpenStreamResponse.ReturnValue == 0x8004010F;
                if (isR891Satisfied)
                {
                    // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                    Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2055,
                        @"[In Additional Error Codes] The numeric value (hex) for error code NotFound is 0x8004010F, %x0F.01.04.80.");
                }

                if (Common.IsRequirementEnabled(891, this.Site))
                {
                    // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R891
                    Site.CaptureRequirementIfIsTrue(
                        isR891Satisfied,
                        891,
                        @"[In Processing RopOpenStream] Implementation does return error ""0x8004010F"" with name ""NotFound"", when The property tag does not exist for the object, and it cannot be created because the Create bit was not specified in OpenModeFlags field.(Microsoft Exchange Server 2007 and above follow this behavior)");
                }
            }
            else
            {
                // Since the RopOpenStream ROP response was parsed successfully, MS-OXCPRPT_R537 can be captured directly.
                Site.CaptureRequirement(
                    537,
                    @"[In Processing RopOpenStream] The server responds with a RopOpenStream ROP response buffer.");

                // If this operation is performed successfully, that means this operation opens a property as a Stream object.
                Site.CaptureRequirement(
                    251,
                    @"[In RopOpenStream ROP] The RopOpenStream ROP ([MS-OXCROPS] section 2.2.9.1) opens a property as a Stream object.");
                if (ropOpenStreamResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
                {
                    if (objectType == ServerObjectType.Message)
                    {
                        // If the object type this operation acting on is Message object and this operation is performed successfully, then the following requirement can be captured.
                        Site.CaptureRequirement(
                            25301,
                            @"[In RopOpenStream ROP] This operation [RopOpenStream ROP] is valid on Message objects.");
                    }

                    if (objectType == ServerObjectType.Folder)
                    {
                        // If the object type this operation acting on is Folder object and this operation is performed successfully, then the following requirement can be captured.
                        Site.CaptureRequirement(
                            25302,
                            @"[In RopOpenStream ROP] This operation [RopOpenStream ROP] is valid on Folder objects.");
                    }

                    if (objectType == ServerObjectType.Attachment)
                    {
                        // If the object type this operation acting on is Attachment object and this operation is performed successfully, then the following requirement can be captured.
                        Site.CaptureRequirement(
                            25303,
                            @"[In RopOpenStream ROP] This operation [RopOpenStream ROP] is valid on Attachment objects.");
                    }

                    // If the OpenStream operation succeeds, it means the StreamSize filed has been set by server internally with the current number of BYTES in the stream.
                    // So this requirement can be verified directly. 
                    Site.CaptureRequirement(
                        276,
                        @"[In RopOpenStream ROP Response Buffer]StreamSize (4 bytes).");

                    // If the OpenStream operation succeeds, it means the StreamSize filed has been set by server internally with the current number of BYTES in the stream.
                    // So this requirement can be verified directly. 
                    Site.CaptureRequirement(
                       27601,
                       @"[In RopOpenStream ROP Response Buffer]StreamSize: An integer that specifies the number of bytes in the opened stream. ");
                }
            }
        }
        #endregion

        #region RopReadStream
        /// <summary>
        /// Verify the RopReadStream operation related requirements.
        /// </summary>
        /// <param name="ropReadStreamResponse">The RopReadStream response buffer structure.</param>
        /// <param name="isReadingFailed">Indicates whether the RopReadStream failed.</param>
        private void VerifyRopReadStream(RopReadStreamResponse ropReadStreamResponse, bool isReadingFailed)
        {
            if (!isReadingFailed)
            {
                // Since the RopReadStream ROP response was parsed successfully, MS-OXCPRPT_R55202 can be captured directly.
                Site.CaptureRequirement(
                    55202,
                    @"[In Processing RopReadStream] The server responds with a RopReadStream ROP response buffer.");

                if (ropReadStreamResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
                {
                    // Since the RopReadStream ROP response was parsed successfully, MS-OXCPRPT_R277 can be captured directly.
                    Site.CaptureRequirement(
                        277,
                        @"[In RopReadStream ROP] The RopReadStream ROP ([MS-OXCROPS] section 2.2.9.2) reads the stream of bytes from a Stream object.");

                    // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                    Site.CaptureRequirement(
                        286,
                        @"[In RopReadStream ROP Response Buffer] DataSize (2 bytes): An integer.");

                    Site.CaptureRequirementIfAreEqual<ushort>(
                        ropReadStreamResponse.DataSize,
                        (ushort)ropReadStreamResponse.Data.Length,
                        28601,
                        @"[In RopReadStream ROP Response Buffer] DataSize: An integer that specifies the number of bytes in the Data field.");

                    Site.CaptureRequirementIfAreEqual<ushort>(
                        ropReadStreamResponse.DataSize,
                        (ushort)ropReadStreamResponse.Data.Length,
                        296,
                        @"[In RopReadStream ROP Response Buffer] Data: This field MUST contain exactly the number of bytes specified in the DataSize field.");

                    // If this operation is performed successfully, that means this operation is valid on Stream objects.
                    Site.CaptureRequirement(
                        27800,
                        @"[In RopReadStream ROP] This operation is valid on Stream objects.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled(88601, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        0,
                        ropReadStreamResponse.DataSize,
                        88601,
                        @"[In Processing RopReadStream] Implementation does set the DataSize field to zero in the case of a failure.(Microsoft Exchange Server 2007 and above follow this behavior)");
                }
            }
        }
        #endregion

        #region RopReadStream
        /// <summary>
        /// Verify the RopReadStream operation related requirements.
        /// </summary>
        /// <param name="readStreamResponse">The RopReadStream response buffer structure.</param>
        /// <param name="byteCount">Indicates the size to be read.</param>
        /// <param name="maxByteCount">If byteCount is 0xBABE, use MaximumByteCount to determine the size to be read.</param>
        private void VerifyRopReadStreamWithLimitedSize(RopReadStreamResponse readStreamResponse, ushort byteCount, uint maxByteCount)
        {
            if (byteCount != 0xBABE)
            {
                Site.CaptureRequirementIfIsTrue(
                    readStreamResponse.DataSize <= byteCount,
                    2830201,
                    @"[In RopReadStream ROP Request Buffer] ByteCount: An integer that specifies the maximum number of bytes to be read unless the value is 0xBABE");

                Site.CaptureRequirementIfIsTrue(
                    readStreamResponse.DataSize <= byteCount,
                    55206,
                    @"[In Processing RopReadStream]otherwise[If the ByteCount field of the ROP request buffer isn't set to 0xBABE], the number of bytes read MUST be less than or equal to the value of the ByteCount field.");
            }
            else
            {
                Site.CaptureRequirementIfIsTrue(
                   readStreamResponse.DataSize <= maxByteCount,
                   2830401,
                   @"[In RopReadStream ROP Request Buffer] MaximumByteCount: An integer that specifies the maximum number of bytes to be read if the ByteCount field is set to 0xBABE. ");

                Site.CaptureRequirementIfIsTrue(
                  readStreamResponse.DataSize <= maxByteCount,
                  55205,
                  @"[In Processing RopReadStream] If the ByteCount field of the ROP request buffer is set to 0xBABE, the number of bytes read MUST be less than or equal to the value of the MaximumByteCount field of the ROP request buffer;");
            }
        }
        #endregion

        #region RopWriteStream
        /// <summary>
        /// Verify the RopWriteStream operation related requirements.
        /// </summary>
        /// <param name="ropWriteStreamResponse">The RopWriteStream response buffer structure.</param>
        /// <param name="openModeFlags">Specifies OpenModeFlags in RopOpenStream.</param>
        /// <param name="writeData">The data to be written to the stream.</param>
        /// <param name="canBeRetrieval">Indicates if the property is immediately available for retrieval by a ROP that uses the same object handle.</param>
        /// <param name="isChangedInDB">Indicates if the value is changed in database or not.</param>
        /// <param name="forErrorCode">Indicates if expecting to fail this operation</param>
        /// <param name="isWriteSizeElementRight">Indicates if the written size is right or not.</param>
        private void VerifyRopWriteStream(RopWriteStreamResponse ropWriteStreamResponse, OpenModeFlags openModeFlags, string writeData, bool canBeRetrieval, bool isChangedInDB, bool forErrorCode, bool isWriteSizeElementRight)
        {
            if (!forErrorCode)
            {
                // Since the RopWriteStream ROP response was parsed successfully, MS-OXCPRPT_R55703 can be captured directly.
                Site.CaptureRequirement(
                55703,
                @"[In Processing RopWriteStream] The server responds with a RopWriteStream ROP response buffer.");

                Site.CaptureRequirementIfIsTrue(
                    canBeRetrieval,
                    56401,
                    @"[In Processing RopWriteStream]After a RopWriteStream ROP request buffer is processed, the new value of the property MUST be immediately available for retrieval by a ROP that uses the same object handle. ");

                if (Common.IsRequirementEnabled(56402, this.Site))
                {
                    Site.CaptureRequirementIfIsFalse(
                        isChangedInDB,
                        56402,
                        @"[In Processing RopWriteStream]However, the new value of the property is not persisted to the database. ");
                }

                Site.CaptureRequirementIfAreEqual<int>(
                    writeData.Length,
                    (int)ropWriteStreamResponse.WrittenSize,
                    297,
                    @"[In RopWriteStream ROP] The RopWriteStream ROP ([MS-OXCROPS] section 2.2.9.3) writes the stream of bytes into a Stream object.");

                Site.CaptureRequirementIfAreEqual<int>(
                   writeData.Length,
                   (int)ropWriteStreamResponse.WrittenSize,
                   303,
                   @"[In RopWriteStream ROP Request Buffer] Data (variable): An array of bytes that constitute the data to be written to the stream.");

                // CPRPTErrorCode.None indicates the operation is performed successfully.
                if (ropWriteStreamResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
                {
                    // If this operation is performed successfully, that means this operation is valid on Stream objects.
                    Site.CaptureRequirement(
                        29801,
                        @"[In RopWriteStream ROP] This operation is valid on Stream objects. ");

                    // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                    Site.CaptureRequirement(
                       304,
                       @"[In RopWriteStream ROP Response Buffer] WrittenSize (2 bytes): An integer.");

                    Site.CaptureRequirementIfIsTrue(
                        isWriteSizeElementRight,
                        30401,
                        @"[In RopWriteStream ROP Response Buffer] WrittenSize: An integer that specifies the number of bytes actually written to the stream.");
                }
            }
            else
            {
                // This mode is to test error code returned from server
                if (openModeFlags == OpenModeFlags.ReadOnly)
                {
                    if (Common.IsRequirementEnabled(901, this.Site))
                    {
                        Site.CaptureRequirementIfAreEqual<uint>(
                            (uint)CPRPTErrorCode.STG_E_ACCESSDENIED,
                            ropWriteStreamResponse.ReturnValue,
                            901,
                            @"[In Processing RopWriteStream] Implementation does return error code ""0x80030005"" with name ""StreamAccessDenied"", when Write access is denied.(Microsoft Exchange Server 2007 and above follow this behavior)");
                    }
                }
            }
        }
        #endregion

        #region RopWriteStreamWithCreatePermission
        /// <summary>
        /// Verify the client has read/write access when open property as stream with Create OpenModeFlags.
        /// </summary>
        /// <param name="writeStreamResponse">The RopWriteStream response buffer structure.</param>
        private void VerifyRopWriteStreamWithCreatePermission(RopWriteStreamResponse writeStreamResponse)
        {
            // CPRPTErrorCode.None indicates the operation is performed successfully.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)CPRPTErrorCode.None,
                writeStreamResponse.ReturnValue,
                271,
                @"[In RopOpenStream ROP Request Buffer] OpenModeFlags: Create: Open a new stream. This mode will delete the current property value and open the stream for read/write access. ");

            Site.CaptureRequirementIfAreEqual<uint>(
                 (uint)CPRPTErrorCode.None,
                 writeStreamResponse.ReturnValue,
                 272,
                 @"[In RopOpenStream ROP Request Buffer] OpenModeFlags: Create: This mode is required for a property that has not been set.");
        }
        #endregion

        #region RopCommitStream
        /// <summary>
        /// Verify the RopCommitStream operation related requirements.
        /// </summary>
        /// <param name="ropCommitStreamResponse">The RopCommitStream response buffer structure.</param>
        private void VerifyRopCommitStream(RopCommitStreamResponse ropCommitStreamResponse)
        {
            // Since the RopCommitStream ROP response was parsed successfully, MS-OXCPRPT_R56902 can be captured directly.
            Site.CaptureRequirement(
                56902,
                @"[In Processing RopCommitStream] The server responds with a RopCommitStream ROP response buffer.");

            // CPRPTErrorCode.None indicates the operation is performed successfully.
            if (ropCommitStreamResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                // If this operation is performed successfully, that means this operation is valid on Stream objects.
                Site.CaptureRequirement(
                    30601,
                    @"[In RopCommitStream ROP] This operation is valid on Stream objects.");
            }
        }
        #endregion

        #region RopGetStreamSize
        /// <summary>
        /// Verify the RopGetStreamSize operation related requirements.
        /// </summary>
        /// <param name="ropGetStreamSizeResponse">The RopGetStreamSize response buffer structure.</param>
        private void VerifyRopGetStreamSize(RopGetStreamSizeResponse ropGetStreamSizeResponse)
        {
            // Since the RopGetStreamSize ROP response was parsed successfully, MS-OXCPRPT_R57502 can be captured directly.
            Site.CaptureRequirement(
                57502,
                @"[In Processing RopGetStreamSize] The server responds with a RopGetStreamSize ROP response buffer.");

            // CPRPTErrorCode.None indicates the operation is performed successfully.
            if (ropGetStreamSizeResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                // Since the RopGetStreamSize ROP response is performed successfully, MS-OXCPRPT_R57502 can be captured directly.
                Site.CaptureRequirement(
                    308,
                    @"[In RopGetStreamSize ROP] The RopGetStreamSize ROP ([MS-OXCROPS] section 2.2.9.5) retrieves the size of the stream.");

                // If this operation is performed successfully, that means this operation is valid on Stream objects.
                Site.CaptureRequirement(
                    30901,
                    @"[In RopGetStreamSize ROP] This operation is valid on Stream objects.");

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                    320,
                    @"[In RopGetStreamSize ROP Response Buffer] StreamSize (4 bytes): An integer .");
            }
        }
        #endregion

        #region RopSetStreamSize
        /// <summary>
        /// Verify the RopSetStreamSize operation related requirements.
        /// </summary>
        /// <param name="ropSetStreamSizeResponse">The RopSetStreamSize response buffer structure.</param>
        /// <param name="isStreamSizeEqualSetSize">Indicates if the stream size is equal to the pre-set value.</param>
        private void VerifyRopSetStreamSize(RopSetStreamSizeResponse ropSetStreamSizeResponse, bool isStreamSizeEqualSetSize)
        {
            // Since the RopSetStreamSize ROP response was parsed successfully, MS-OXCPRPT_R57702 can be captured directly.
            Site.CaptureRequirement(
                57702,
                @"[In Processing RopSetStreamSize] The server responds with a RopSetStreamSize ROP response buffer.");

            // CPRPTErrorCode.None indicates the operation is performed successfully.
            if (ropSetStreamSizeResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                // If this operation is performed successfully, that means this operation is valid on Stream objects.
                Site.CaptureRequirement(
                    32301,
                    @"[In RopSetStreamSize ROP] This operation is valid on Stream objects. ");

                Site.CaptureRequirementIfIsTrue(
                    isStreamSizeEqualSetSize,
                    32003,
                    @"[In RopGetStreamSize ROP Response Buffer] StreamSize: An integer that specifies the number of bytes in the stream.");

                Site.CaptureRequirementIfIsTrue(
                    isStreamSizeEqualSetSize,
                    575,
                    @"[In Processing RopGetStreamSize] The server MUST return the current size of the Stream object. ");
            }
        }
        #endregion

        #region RopSeekStream
        /// <summary>
        /// Verify the RopSeekStream operation related requirements.
        /// </summary>
        /// <param name="ropSeekStreamResponse">The RopSeekStream response buffer structure.</param>
        /// <param name="seekStereamCondition">Specifies the particular scenario when performing RopSeekStream.</param>
        private void VerifyRopSeekStream(RopSeekStreamResponse ropSeekStreamResponse, SeekStreamCondition seekStereamCondition)
        {
            // Since the RopSeekStream ROP response was parsed successfully, MS-OXCPRPT_R57902 can be captured directly.
            Site.CaptureRequirement(
                57902,
                @"[In Processing RopSeekStream] The server responds with a RopSeekStream ROP response buffer.");

            // CPRPTErrorCode.None indicates the operation is performed successfully.
            if (ropSeekStreamResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                // If this operation is performed successfully, that means this operation is valid on Stream objects.
                Site.CaptureRequirement(
                    33101,
                    @"[In RopSeekStream ROP] This ROP is valid on Stream objects.");

                // Since the RopSeekStream ROP response is performed successfully, MS-OXCPRPT_R581 can be captured directly.
                Site.CaptureRequirement(
                  581,
                  @"[In Processing RopSeekStream] The server modifies the location of the seek point associated with the Stream object according to the ROP request buffer. ");

                // Since the RopSeekStream ROP response is performed successfully, MS-OXCPRPT_R330 can be captured directly.
                Site.CaptureRequirement(
                    330,
                    @"[In RopSeekStream ROP] RopSeekStream sets the seek pointer to a new location.");

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                   341,
                   @"[In  RopSeekStream ROP Response Buffer] NewPosition (8 bytes): An integer.");
            }

            if (seekStereamCondition == SeekStreamCondition.OriginInvalid)
            {
                if (Common.IsRequirementEnabled(903, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<uint>(
                        (uint)CPRPTErrorCode.StreamInvalidParam,
                        ropSeekStreamResponse.ReturnValue,
                        903,
                        @"[In Processing RopSeekStream] Implementation does return error code ""0x80030057"" with name ""StreamInvalidParam"", when The value of the Origin field is invalid.(Microsoft Exchange Server 2007 and above follow this behavior)");
                }
            }

            if (seekStereamCondition == SeekStreamCondition.MovedBeyondMaxStreamSize)
            {
                if (Common.IsRequirementEnabled(894, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<uint>(
                        (uint)CPRPTErrorCode.StreamSeekError,
                        ropSeekStreamResponse.ReturnValue,
                        894,
                        @"[In Appendix A: Product Behavior] Implementation does return error code ""0x80030019"" with name ""StreamSeekError"",  When Tried to seek to offset before the start or beyond the max stream size of 2^31.(Microsoft Exchange Server 2007 and above follow this behavior)");
                }
            }
        }
        #endregion

        #region RopCopyToStream
        /// <summary>
        /// Verify the RopCopyToStream operation related requirements.
        /// </summary>
        /// <param name="ropCopyToStreamResponse">The RopCopyToStream response buffer structure.</param>
        /// <param name="isDestinationExist">A boolean value indicates if the destination object exists or not</param>
        /// <param name="isCopyRightDataCount">Indicates if the number of properties to be copied is right or not.</param>
        /// <param name="destHandleIsRight">Indicates if the DestHandleIndex is right or not.</param>
        private void VerifyRopCopyToStream(RopCopyToStreamResponse ropCopyToStreamResponse, bool isDestinationExist, bool isCopyRightDataCount, bool destHandleIsRight)
        {
            // Since the RopCopyToStream ROP response was parsed successfully, MS-OXCPRPT_R76002 can be captured directly.
            Site.CaptureRequirement(
                76002,
                @"[In Processing RopCopyToStream] The server responds with a RopCopyToStream ROP response buffer.");

            if (isDestinationExist)
            {
                Site.CaptureRequirementIfIsTrue(
                    isCopyRightDataCount,
                    34902,
                    @"[In RopCopyToStream ROP Response Buffer] WrittenByteCount: An integer that specifies the number of bytes written to the destination stream. ");
            }

            if (ropCopyToStreamResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                // If this operation is performed successfully, that means this operation is valid on Stream objects.
                Site.CaptureRequirement(
                    34301,
                    @"[In RopCopyToStream ROP] This operation is valid on Stream objects.");

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                    35101,
                    @"[In RopCopyToStream ROP Response Buffer] DestHandleIndex (4 bytes): An integer.");

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                  34701,
                  @"[In RopCopyToStream ROP Response Buffer] ReadByteCount: 8 bytes integer.");

                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirement(
                 34901,
                 @"[In RopCopyToStream ROP Response Buffer] WrittenByteCount: 8 bytes integer.");

                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R591
                // ReadByteCount specifies the number of BYTES read from the source object while WrittenByteCount specifies the number of BYTES written to the destination object.
                // When this operation is performed successfully, if ReadByteCount equals WrittenByteCount, it means the position of destination streams is moved forward the same bytes as the source streams.
                Site.CaptureRequirementIfAreEqual<ulong>(
                    ropCopyToStreamResponse.WrittenByteCount,
                    ropCopyToStreamResponse.ReadByteCount,
                    591,
                    @"[In Processing RopCopyToStream] The server MUST move the seek pointer of both the source and destination streams forward the same number of bytes as were copied.");
            }

            if (!ropCopyToStreamResponse.ReturnValue.Equals((uint)CPRPTErrorCode.NullDestinationObject))
            {
                // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R352
                // The value of DestHandleIndex is 0 means this field is not present.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    ropCopyToStreamResponse.DestHandleIndex,
                    352,
                    @"[In RopCopyToStream ROP Response Buffer] DestHandleIndex: The DestHandleIndex field MUST NOT be present if the ReturnValue field is set to any value other than NullDestinationObject (0x00000503).");
            }
            else
            {
                // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                Site.CaptureRequirementIfIsTrue(
                    destHandleIsRight,
                    351,
                    @"[In RopCopyToStream ROP Response Buffer] DestHandleIndex: The DestHandleIndex field MUST be set to the value of the DestHandleIndex field of the ROP request buffer.");
            }

            if (!isDestinationExist)
            {
                if (Common.IsRequirementEnabled(895, this.Site))
                {
                    // Verify MS-OXCPRPT requirement: MS-OXCPRPT_R895
                    Site.CaptureRequirementIfAreEqual<uint>(
                        (uint)CPRPTErrorCode.NullDestinationObject,
                        ropCopyToStreamResponse.ReturnValue,
                        895,
                        @"[In Processing RopCopyToStream] Implementation does return error code ""0x00000503"" with name ""NullDestinationObject"", when Destination object does not exist.(Microsoft Exchange Server 2007 and above follow this behavior)");
                }
            }
        }
        #endregion

        #region RopLockRegionStream
        /// <summary>
        /// Verify the RopLockRegionStream operation related requirements.
        /// </summary>
        /// <param name="ropLockRegionStreamResponse">The RopLockRegionStream response buffer structure.</param>
        private void VerifyRopLockRegionStream(RopLockRegionStreamResponse ropLockRegionStreamResponse)
        {
            // Since the RopLockRegionStream ROP response was parsed successfully, MS-OXCPRPT_R60703 can be captured directly.
            Site.CaptureRequirement(
                60703,
                @"[In Processing RopLockRegionStream] The server responds with a RopLockRegionStream ROP response buffer.");

            if (Common.IsRequirementEnabled(750, this.Site))
            {
                if (ropLockRegionStreamResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
                {
                    // If this operation is performed successfully, that means this operation is valid on Stream objects.
                    Site.CaptureRequirement(
                       750,
                       @"[In Appendix A: Product Behavior] Implementation does implement the RopLockRegionStream ROP. (<15> Section 3.2.5.20: Exchange 2003 and Exchange 2007 implement the RopLockRegionStream ROP.)");

                    // If this operation is performed successfully, that means this operation is valid on Stream objects.
                    Site.CaptureRequirement(
                        37401,
                        @"[In RopLockRegionStream ROP] This operation is valid on Stream objects.");
                }
            }
        }
        #endregion

        #region RopUnlockRegionStream

        /// <summary>
        /// Verify the RopUnlockRegionStream operation related requirements.
        /// </summary>
        /// <param name="ropUnlockRegionStreamResponse">The RopUnlockRegionStream response buffer structure.</param>
        private void VerifyRopUnlockRegionStream(RopUnlockRegionStreamResponse ropUnlockRegionStreamResponse)
        {
            // Since the RopUnlockRegionStream ROP response was parsed successfully, MS-OXCPRPT_R62003 can be captured directly.
            Site.CaptureRequirement(
                62003,
                @"[In Processing RopUnlockRegionStream] The server responds with a RopUnlockRegionStream ROP response buffer.");

            if (Common.IsRequirementEnabled(751, this.Site))
            {
                if (ropUnlockRegionStreamResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
                {
                    // If this operation is performed successfully, that means this operation is valid on Stream objects.
                    Site.CaptureRequirement(
                       751,
                       @"[In Appendix A: Product Behavior] Implementation does implement the RopUnlockRegionStream ROP. (<16> Section 3.2.5.21: Exchange 2003 and Exchange 2007 implement the RopUnlockRegionStream ROP ([MS-OXCROPS] section 2.2.9.10).)");

                    // If this operation is performed successfully, that means this operation is valid on Stream objects.
                    Site.CaptureRequirement(
                        38103,
                        @"[In RopUnlockRegionStream ROP] This ROP is valid on Stream objects.");
                }
            }
        }
        #endregion

        #region RopWriteAndCommitStream
        /// <summary>
        /// Verify the RopWriteAndCommitStream operation related requirements.
        /// </summary>
        /// <param name="ropWriteAndCommitStreamResponse">The RopWriteAndCommitStream response buffer structure.</param>
        /// <param name="isWriteCorrectDataToStream">Indicates if the data is written to stream properly.</param>
        private void VerifyRopWriteAndCommitStream(RopWriteStreamResponse ropWriteAndCommitStreamResponse, bool isWriteCorrectDataToStream)
        {
            // Since the RopWriteAndCommitStream ROP response was parsed successfully, MS-OXCPRPT_R80503 can be captured directly.
            Site.CaptureRequirement(
                80503,
                @"[In Processing RopWriteAndCommitStream] The server responds with a RopWriteAndCommitStream ROP response buffer.");

            if (Common.IsRequirementEnabled(752, this.Site))
            {
                // CPRPTErrorCode.None indicates the operation is performed successfully.
                if (ropWriteAndCommitStreamResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
                {
                    // Since the RopWriteAndCommitStream ROP response is performed successfully, MS-OXCPRPT_R752 can be captured directly.
                    Site.CaptureRequirement(
                        752,
                        @"[In Appendix A: Product Behavior] Implementation does implement the RopWriteAndCommitStream ROP. (<17> Section 3.2.5.22: Exchange 2003 and Exchange 2007 implement the RopWriteAndCommitStream ROP ([MS-OXCROPS] section 2.2.9.11.))");

                    // Since the RopWriteAndCommitStream ROP response is performed successfully, MS-OXCPRPT_R805 can be captured directly.
                    Site.CaptureRequirement(
                       805,
                       @"[In Processing RopWriteAndCommitStream] If the server implements this ROP, it MUST first process the ROP request as specified in section 3.2.5.13 ");

                    // Since the RopWriteAndCommitStream ROP response is performed successfully, MS-OXCPRPT_R812 can be captured directly.
                    Site.CaptureRequirement(
                      812,
                      @"[In Processing RopWriteAndCommitStream] If the server implements this ROP, it MUST then process the ROP process the ROP request as specified in section 3.2.5.14.");

                    // If this operation is performed successfully, that means this operation is valid on Stream objects.
                    Site.CaptureRequirement(
                        63201,
                        @"[In RopWriteAndCommitStream ROP] This operation is valid on Stream objects.");

                    // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                    Site.CaptureRequirement(
                     393,
                     @"[In RopWriteAndCommitStream ROP Response Buffer] WrittenSize (2 bytes): An integer.");
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isWriteCorrectDataToStream,
                39301,
                @"[In RopWriteAndCommitStream ROP Response Buffer] WrittenSize: An integer that specifies the number of bytes actually written and committed to the stream.");
        }
        #endregion

        #region RopCloneStream
        /// <summary>
        /// Verify the RopCloneStream operation related requirements.
        /// </summary>
        /// <param name="ropCloneStreamResponse">The RopCloneStream response buffer structure.</param>
        private void VerifyRopCloneStream(RopCloneStreamResponse ropCloneStreamResponse)
        {
            // Since the RopCloneStream ROP response was parsed successfully, MS-OXCPRPT_R75603 can be captured directly.
            Site.CaptureRequirement(
                75603,
                @"[In Processing RopCloneStream] The server responds with a RopCloneStream ROP response buffer.");

            if (Common.IsRequirementEnabled(753, this.Site))
            {
                if (ropCloneStreamResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
                {
                    // If this operation is performed successfully, the following requirement can be captured.
                    Site.CaptureRequirement(
                       753,
                       @"[In Appendix A: Product Behavior] Implementation does implement the RopCloneStream ROP. (<18> Section 3.2.5.23: Exchange 2003 and Exchange 2007 implement the RopCloneStream ROP ([MS-OXCROPS] section 2.2.9.12).)");

                    // If this operation is performed successfully, that means this operation is valid on Stream objects.
                    Site.CaptureRequirement(
                        39801,
                        @"[In RopCloneStream ROP] This operation is valid on Stream objects.");
                }
            }
        }
        #endregion
        #endregion

        #region Verify requirements in MS-OXCDATA
        /// <summary>
        /// Verify the Common Properties related requirements in MS-OXCDATA.
        /// </summary>
        /// <param name="propertyTag">A property tag both identifies a property and gives the data type its value.</param>
        /// <param name="ropGetPropertiesSpecificResponse">The RopGetPropertiesSpecific response buffer structure.</param>
        private void VerifyCommonPropertiesInOXCDATA(PropertyTag propertyTag, RopGetPropertiesSpecificResponse ropGetPropertiesSpecificResponse)
        {
            // propertyId indicates the Id of a property.
            ushort propertyId = propertyTag.PropertyId;

            foreach (PropertyValue propertyVal in ropGetPropertiesSpecificResponse.RowData.PropertyValues)
            {
                // PropertyValue (Variable): A PropertyValue structure specifying the value of the property. Its syntax is specified by the PropertyType field of the tag, and its semantics by the PropertyId field of the tag, defined in MS-OXCDATA section 2.12.4 TaggedPropertyValue
                switch (propertyId)
                {
                    // According to MS-OXPROPS 2.700 PidTagChangeKey, if PropertyId is 0x65E2, it means the property is PidTagChangeKey, the Data type is PtypBinary
                    case 0x65E2:
                        {
                            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2707
                            Site.CaptureRequirementIfAreEqual<int>(
                                propertyVal.Value.Length - 2,
                                BitConverter.ToInt16(propertyVal.Value, 0),
                                "MS-OXCDATA",
                                2707,
                                 @"[In Property Data Types] PtypBinary (PT_BINARY) is that variable size; a COUNT field followed by that many bytes with Property Type Value 0x0102,%x02.01.");
                        }

                        break;

                    // According to MS-OXPROPS 2.589 PidTagAccess, if PropertyId is 0x0FF4, it means the property is PidTagAccess, the Data type is PtypInteger32
                    case 0x0FF4:
                        {
                            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2691
                            Site.CaptureRequirementIfAreEqual<int>(
                                4,
                                propertyVal.Value.Length,
                                "MS-OXCDATA",
                                2691,
                                @"[In Property Data Types] PtypInteger32 (PT_LONG, PT_I4, int, ui4) is that 4 bytes; a 32-bit integer [MS-DTYP]: INT32 with Property Type Value 0x0003,%x03.00.");
                        }

                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Verify the propertyTag structure with MS-OXCDATA.
        /// </summary>
        /// <param name="propertyTag">The propertyTag value</param>
        private void VerifyPropertyTagInCDATA(PropertyTag propertyTag)
        {
            switch (propertyTag.PropertyType)
            {
                case 0x001E:
                    {
                        // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                        Site.CaptureRequirement(
                             "MS-OXCDATA",
                             2701,
                             @"[In Property Data Types] PtypString8 is that variable size, a string of multi-byte characters in externally specified encoding with terminating null character (single 0 byte).PT_STRING8 with Property Type Value 0x001E, %z1E.00.");
                    }

                    break;
                case 0x0040:
                    {
                        // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                        Site.CaptureRequirement(
                            "MS-OXCDATA",
                            2702,
                            @"[In Property Data Types] PtypTime (PT_SYSTIME, time, datetime, datetime.tz, datetime.rfc1123, Date, time, time.tz) is that 8 bytes; a 64-bit integer representing the number of 100-nanosecond intervals since January 1, 1601 [MS-DTYP]: FILETIME with Property Type Value 0x0040,%x40.00.");
                    }

                    break;
                case 0x001F:
                    {
                        // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                        Site.CaptureRequirement(
                            "MS-OXCDATA",
                            2700,
                            @"[In Property Data Types] PtypString (PT_UNICODE, string) is that Variable size; a string of Unicode characters in UTF-16LE format encoding with terminating null character (0x0000). with Property Type Value  0x001F,%x1F.00.");
                    }

                    break;
                case 0x000A:
                    {
                        // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                        Site.CaptureRequirement(
                             "MS-OXCDATA",
                             2697,
                             @"[In Property Value Types] PtypErrorCode(PT_ERROR) is that 
                                     4 bytes; a 32-bit integer encoding error information as specified in section 2.4.1.
                                     with Property Type Value  0x000A,%x0A.00.");
                    }

                    break;
                case 0x0102:
                    {
                        // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
                        Site.CaptureRequirement(
                            "MS-OXCDATA",
                            2707,
                            @"[In Property Data Types] PtypBinary (PT_BINARY) is that variable size; a COUNT field followed by that many bytes with Property Type Value 0x0102,%x02.01.");
                    }

                    break;

                default:
                    break;
            }

            // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
            Site.CaptureRequirement(
              "MS-OXCDATA",
              181,
              @"[In PropertyTag Structure] PropertyType (2 bytes): An unsigned integer that identifies the data type of the property value, as specified by the table in section 2.11.1.");

            // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
            Site.CaptureRequirement(
               "MS-OXCDATA",
               182,
               @"[In PropertyTag Structure] PropertyId (2 bytes): An unsigned integer that identifies the property.");

            // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
            Site.CaptureRequirement(
                         "MS-OXCDATA",
                         468,
                         @"[In TaggedPropertyValue Structure] PropertyTag (4 bytes): A PropertyTag structure, as specified in section 2.9, giving the values of the PropertyId and PropertyType fields for the property.");
        }

        /// <summary>
        /// Verify the PropertyProblem structure defined in MS-OXCDATA.
        /// </summary>
        private void VerifyPropertyProblemSturctureInCDATA()
        {
            // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
            Site.CaptureRequirement(
                "MS-OXCDATA",
                46,
                @"[In PropertyProblem Structure] Index (2 bytes): An unsigned integer.");

            // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
            Site.CaptureRequirement(
                "MS-OXCDATA",
                48,
                @"[In PropertyProblem Structure] PropertyTag (4 bytes): A PropertyTag structure, as specified in section 2.9.");

            // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
            Site.CaptureRequirement(
                "MS-OXCDATA",
                50,
                @"[In PropertyProblem Structure] ErrorCode (4 bytes): An unsigned integer.");

            // The parser has ensured the field satisfied the format, otherwise the response cannot be received.
            Site.CaptureRequirement(
                "MS-OXCDATA",
                52,
                @"[In PropertyProblem Structure] An array of PropertyProblem structures is returned from the following ROPs:
                        RopDeleteProperties ([MS-OXCROPS] section 2.2.8.8), 
                        RopDeletePropertiesNoReplicate ([MS-OXCROPS] section 2.2.8.9), 
                        RopSetProperties ([MS-OXCROPS] section 2.2.8.6), 	RopSetPropertiesNoReplicate ([MS-OXCROPS] section 2.2.8.7), 
                        RopCopyProperties ([MS-OXCROPS] section 2.2.8.11), 
                        RopCopyTo ([MS-OXCROPS] section 2.2.8.12).");
        }
        #endregion

        #region Verify requirements in MS-OXPROPS
        /// <summary>
        /// Verify the type of Common Object Properties in MS-OXPROPS.
        /// </summary>
        /// <param name="propertyTag">propertyTag indicates the property currently set on the object.</param>
        /// <param name="ropGetPropertiesSpecificResponse">The RopGetPropertiesSpecific response buffer structure.</param>
        private void VerifyTypeOfCommonPropertiesInOXPROPS(PropertyTag propertyTag, RopGetPropertiesSpecificResponse ropGetPropertiesSpecificResponse)
        {
            switch (propertyTag.PropertyId)
            {
                // According to MS-OXPROPS 2.589 PidTagAccess, if PropertyId is 0x0FF4, it means the property is PidTagAccess
                case 0x0FF4:
                    {
                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R4927
                        // If the property ID is 0x0FF4, it means the associated property is PidTagAccess
                        Site.CaptureRequirement(
                            "MS-OXPROPS",
                            4927,
                            @"[In PidTagAccess] Property ID: 0x0FF4.");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R4928
                        // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to indicate that all property values are present and without error.
                        Site.CaptureRequirementIfAreEqual<uint>(
                            0x00,
                            ropGetPropertiesSpecificResponse.RowData.Flag,
                            "MS-OXPROPS",
                            4928,
                            @"[In PidTagAccess] Data type: PtypInteger32, 0x0003.");
                    }

                    break;

                // According to MS-OXPROPS 2.590 PidTagAccessLevel, if PropertyId is 0x0FF7, it means the property is PidTagAccessLevel
                case 0x0FF7:
                    {
                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R4934
                        // If the property ID is 0x0FF7, it means the associated property is PidTagAccessLevel
                        Site.CaptureRequirement(
                            "MS-OXPROPS",
                            4934,
                            @"[In PidTagAccessLevel] Property ID: 0x0FF7.");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R4935
                        // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to indicate that all property values are present and without error.
                        Site.CaptureRequirementIfAreEqual<uint>(
                            0x00,
                            ropGetPropertiesSpecificResponse.RowData.Flag,
                            "MS-OXPROPS",
                            4935,
                            @"[In PidTagAccessLevel] Data type: PtypInteger32, 0x0003.");
                    }

                    break;

                // According to MS-OXPROPS 2.700 PidTagChangeKey, if PropertyId is 0x65E2, it means the property is PidTagChangeKey
                case 0x65E2:
                    {
                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5686
                        // If the property ID is 0x65E2, it means the associated property is PidTagChangeKey
                        Site.CaptureRequirement(
                            "MS-OXPROPS",
                            5686,
                            @"[In PidTagChangeKey] Property ID: 0x65E2.");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5687
                        // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to indicate that all property values are present and without error.
                        Site.CaptureRequirementIfAreEqual<uint>(
                            0x00,
                            ropGetPropertiesSpecificResponse.RowData.Flag,
                            "MS-OXPROPS",
                            5687,
                            @"[In PidTagChangeKey] Data type: PtypBinary, 0x0102.");
                    }

                    break;

                // According to MS-OXPROPS 2.729 PidTagCreationTime, if PropertyId is 0x3007, it means the property is PidTagCreationTime
                case 0x3007:
                    {
                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5880
                        // If the property ID is 0x3007, it means the associated property is PidTagCreationTime
                        Site.CaptureRequirement(
                            "MS-OXPROPS",
                            5880,
                            @"[In PidTagCreationTime] Property ID: 0x3007.");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5881
                        // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to indicate that all property values are present and without error.
                        Site.CaptureRequirementIfAreEqual<uint>(
                            0x00,
                            ropGetPropertiesSpecificResponse.RowData.Flag,
                            "MS-OXPROPS",
                            5881,
                            @"[In PidTagCreationTime] Data type: PtypTime, 0x0040.");
                    }

                    break;

                // According to MS-OXPROPS 2.868 PidTagLastModifierName, if PropertyId is 0x3FFA, it means the property is PidTagLastModifierName
                case 0x3FFA:
                    {
                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6796
                        // If the property ID is 0x3FFA, it means the associated property is PidTagLastModifierName
                        Site.CaptureRequirement(
                            "MS-OXPROPS",
                            6796,
                            @"[In PidTagLastModifierName] Property ID: 0x3FFA.");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6797
                        // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to indicate that all property values are present and without error.
                        Site.CaptureRequirementIfAreEqual<uint>(
                            0x00,
                            ropGetPropertiesSpecificResponse.RowData.Flag,
                            "MS-OXPROPS",
                            6797,
                            @"[In PidTagLastModifierName] Data type: PtypString, 0x001F.");
                    }

                    break;

                // According to MS-OXPROPS 2.866 PidTagLastModificationTime, if PropertyId is 0x3008, it means the property is PidTagLastModificationTime
                case 0x3008:
                    {
                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6783
                        // If the property ID is 0x3008, it means the associated property is PidTagLastModificationTime
                        Site.CaptureRequirement(
                            "MS-OXPROPS",
                            6783,
                            @"[In PidTagLastModificationTime] Property ID: 0x3008.");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6784
                        // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to indicate that all property values are present and without error.
                        Site.CaptureRequirementIfAreEqual<uint>(
                            0x00,
                            ropGetPropertiesSpecificResponse.RowData.Flag,
                            "MS-OXPROPS",
                            6784,
                            @"[In PidTagLastModificationTime] Data type: PtypTime, 0x0040.");
                    }

                    break;

                // According to MS-OXPROPS 2.1110 PidTagSearchKey, if PropertyId is 0x300B, it means the property is PidTagSearchKey
                case 0x300B:
                    {
                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R8385
                        // If the property ID is 0x300B, it means the associated property is PidTagSearchKey
                        Site.CaptureRequirement(
                            "MS-OXPROPS",
                            8385,
                            @"[In PidTagSearchKey] Property ID: 0x300B");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R8386
                        // According to MS-OXCDATA section 2.8.1.1 StandardPropertyRow, Flag is equal to 0x00 to indicate that all property values are present and without error.
                        Site.CaptureRequirementIfAreEqual<uint>(
                            0x00,
                            ropGetPropertiesSpecificResponse.RowData.Flag,
                            "MS-OXPROPS",
                            8386,
                            @"[In PidTagSearchKey] Data type: PtypBinary, 0x0102.");
                    }

                    break;
                default:
                    break;
            }
        }
        #endregion
    }
}