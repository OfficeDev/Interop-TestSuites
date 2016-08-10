[assembly: Microsoft.Xrt.Runtime.NativeType("System.Diagnostics.Tracing.*")]

namespace Microsoft.Protocols.TestSuites.MS_OXCPRPT
{
   using Microsoft.Modeling;

    /// <summary>
    /// Model program for MS-OXCPRPT.
    /// </summary>
    public static class Model
    {
        #region State Variables
        /// <summary>
        /// Indicates whether test environment has been initialized
        /// </summary>
        private static bool isInitialized;
        
        /// <summary>
        /// Indicates whether client user successfully set properties
        /// </summary>
        private static bool isSetPropertiesSuccess;

        /// <summary>
        /// Indicates whether client user successfully delete properties
        /// </summary>
        private static bool isDeletePropertiesSuccess;

        /// <summary>
        /// Indicates whether client user commit stream properties successfully
        /// </summary>
        private static bool isCommitStreamSuccess;

        /// <summary>
        /// Indicates whether client user successfully copy properties
        /// </summary>
        private static bool isCopyPropertiesSuccess;

        /// <summary>
        /// Indicates whether client user successfully copyTo
        /// </summary>
        private static bool isCopyToSuccess;

        /// <summary>
        /// Indicates whether a stream is successfully opened by [RopOpenStream]
        /// </summary>
        private static bool isStreamOpenedSuccess;

        /// <summary>
        /// Indicates whether a stream is successfully written by [RopWriteStream]
        /// </summary>
        private static bool isStreamWriteSuccess;

        /// <summary>
        /// Indicates whether stream is locked
        /// </summary>
        private static bool isStreamLocked;

        /// <summary>
        /// Indicates stream open mode
        /// </summary>
        private static OpenModeFlags streamOpenFlag = OpenModeFlags.ReadOnly;

        /// <summary>
        /// Indicates whether the server is running under asynchronous mode
        /// </summary>
        private static bool isWorkSynchronously;

        /// <summary>
        /// Specifies the Object type that the ROPs act on
        /// </summary>
        private static ServerObjectType globalObj = ServerObjectType.Logon;

        /// <summary>
        /// Indicates whether the first object handle has been got
        /// </summary>
        private static bool isFirstObjectGot;

        /// <summary>
        /// Indicates whether the second object handle has been got
        /// </summary>
        private static bool isSecondObjectGot;

        /// <summary>
        /// The list to save requirement status
        /// </summary>
        private static MapContainer<int, bool> requirementContainer = new MapContainer<int, bool>();
        #endregion

        #region Client state variables for asynchronous operations

        /// <summary>
        /// Specifies the CopyFlags in RopCopyTo and RopCopyProperties request
        /// </summary>
        private static CopyFlags clientCopyFlag = CopyFlags.Move;

        /// <summary>
        /// Specifies whether the WantAsynchronous is zero or non-zero in RopCopyTo and RopCopyProperties request
        /// if WantAsynchronous is zero, isClientWantAsynchronous = false, otherwise, isClientWantAsynchronous = true.
        /// </summary>
        private static bool isClientWantAsynchronous;

        /// <summary>
        /// Specifies whether the WantSubObjects is zero or non-zero in RopCopyTo and RopCopyProperties request
        /// if WantSubObjects is zero, isClientWantSubObjects = false, otherwise, isClientWantSubObjects = true.
        /// </summary>
        private static bool isClientWantSubObjects;

        /// <summary>
        /// Specifies whether the destination specified by DestHandleIndex in RopCopyTo and RopCopyProperties request is exist or not
        /// </summary>
        private static bool isDestinationInRequestExist;

        /// <summary>
        /// Specifies whether another ROP operation is sent by client after receiving the response of RopProgress 
        /// </summary>
        private static bool isClientSendOtherRop;

        /// <summary>
        /// Specifies whether client want abort RopProgress
        /// if WantCancel is zero, isClientWantCancel = false, means the client wants the current operation to continue. 
        /// If WantCancel non-zero, isClientWantCancel = true, the client is requesting that the server attempt to cancel the operation.
        /// </summary>
        private static bool isClientWantCancel;

        /// <summary>
        /// Specifies whether RopSaveChangesAttachment is called successfully
        /// </summary>
        private static bool isSaveChangesAttachmentSuccess;

        #endregion

        #region Initialization

        /// <summary>
        /// Action Initialization which is used to initialize test environment
        /// </summary>
        [Rule(Action = "InitializeMailBox()")]
        public static void InitializeMailBox()
        {
            isInitialized = true;
            isSetPropertiesSuccess = false;
            isDeletePropertiesSuccess = false;
            isCopyPropertiesSuccess = false;
            isCopyToSuccess = false;
            isCommitStreamSuccess = false;
            isSaveChangesAttachmentSuccess = false;
            isStreamLocked = false;
            isStreamOpenedSuccess = false;
            isStreamWriteSuccess = false;
        }

        /// <summary>
        /// Action Initialization which is used to initialize test environment
        /// </summary>
        [Rule(Action = "InitializePublicFolder")]
        public static void InitializePublicFolder()
        {
            isInitialized = true;
        }

        #endregion

        /// <summary>
        /// This action is used to get different Object 
        /// </summary>
        /// <param name="objType">Specifies ServerObjectType</param>
        /// <param name="objToOperate">Specifies which object will be got</param>
        [Rule(Action = "GetObject(objType, objToOperate)")]
        public static void GetObject1(ServerObjectType objType, ObjectToOperate objToOperate)
        {
            Condition.IsTrue(objToOperate == ObjectToOperate.FirstObject || objToOperate == ObjectToOperate.FifthObject);

            isFirstObjectGot = true;
            globalObj = objType;
        }

        /// <summary>
        /// This action is used to get second object which has the same type with the first
        /// </summary>
        /// <param name="objType">Specifies ServerObjectType</param>
        /// <param name="objToOperate">Specifies which object will be got</param>
        [Rule(Action = "GetObject(objType,objToOperate)")]
        public static void GetObject2(ServerObjectType objType, ObjectToOperate objToOperate)
        {
            Condition.IsTrue(objType == globalObj && isFirstObjectGot);
            Condition.IsTrue(objToOperate == ObjectToOperate.SecondObject);

            isSecondObjectGot = true;
        }

        /// <summary>
        /// Action for [RopQueryNamedProperties] operation
        /// </summary>
        /// <param name="queryFlags">Specifies QueryFlags parameter in request</param>
        /// <param name="hasGuid">Indicates whether PropertyGUID is present</param>
        /// <param name="isKind0x01Return">Indicates whether kind with 0x01 is returned</param>
        /// <param name="isKind0x00Return">Indicates whether kind with 0x00 is returned</param>
        /// <param name="isNamedPropertyGuidReturn">Indicates whether name properties is returned</param>
        [Rule(Action = "RopQueryNamedPropertiesMethod(queryFlags, hasGuid, out isKind0x01Return, out isKind0x00Return, out isNamedPropertyGuidReturn)")]
        public static void RopQueryNamedPropertiesMethod(
            QueryFlags queryFlags,
            bool hasGuid,
            out bool isKind0x01Return,
            out bool isKind0x00Return,
            out bool isNamedPropertyGuidReturn)
        {
            Condition.IsTrue(isInitialized);
            Condition.IsTrue(
                (globalObj == ServerObjectType.Logon && requirementContainer[12904]) ||
                globalObj == ServerObjectType.Folder ||
                globalObj == ServerObjectType.Message ||
                globalObj == ServerObjectType.Attachment);

            isKind0x01Return = false;
            isKind0x00Return = false;
            isNamedPropertyGuidReturn = false;

            if (queryFlags == QueryFlags.NoIds)
            {
                if (hasGuid)
                {
                    isKind0x01Return = true;
                    ModelHelper.CaptureRequirement(
                        877,
                        @"[In Processing RopQueryNamedProperties] Starting with the full list of all  named properties:
                        If the NoIds bit is set in the QueryFlags field, named properties with the Kind field set to 0x0 MUST NOT be returned.");

                    isNamedPropertyGuidReturn = true;
                    ModelHelper.CaptureRequirement(
                        878,
                        @"[In Processing RopQueryNamedProperties] Starting with the full list of all  named properties: 
                        If the PropertyGuid field of the ROP request buffer is present, named properties with a GUID field ([MS-OXCDATA] section 2.6.1) value 
                        that does not match the value of the PropertyGuid field MUST NOT be returned.");
                }
            }
            else if (!hasGuid && queryFlags == QueryFlags.NoStrings)
            {
                isKind0x00Return = true;
                ModelHelper.CaptureRequirement(
                    876,
                    @"[In Processing RopQueryNamedProperties] Starting with the full list of all  named properties: 
                    If the NoStrings bit is set in the QueryFlags field of the ROP request buffer, named properties with the Kind field ([MS-OXCDATA] section 2.6.1) 
                    set to 0x1 MUST NOT be returned.");
            }
        }

        /// <summary>
        /// Action for [RopGetPropertiesAll] operation
        /// </summary>
        /// <param name="isPropertySizeLimitZero">Indicates whether PropertySizeLimit parameter is zero</param>
        /// <param name="isPropertyLargerThanLimit">Indicates whether request properties are larger than the limit
        /// When PropertySizeLimit is non-zero, it indicates whether request properties larger than PropertySizeLimit
        /// When PropertySizeLimit is zero, it indicates whether request properties larger than size of response</param>
        /// <param name="isUnicode">Indicates whether the requested property is encoded in Unicode format in response buffer</param>
        /// <param name="isValueContainsNotEnoughMemory">Indicates whether returned value contains NotEnoughMemory error when request properties too large</param>
        [Rule(Action = "RopGetPropertiesAllMethod(isPropertySizeLimitZero,isPropertyLargerThanLimit,isUnicode,out isValueContainsNotEnoughMemory)")]
        public static void RopGetPropertiesAllMethod(bool isPropertySizeLimitZero, bool isPropertyLargerThanLimit, bool isUnicode, out bool isValueContainsNotEnoughMemory)
        {
            Condition.IsTrue(isInitialized);
            Condition.IfThen(isPropertyLargerThanLimit, globalObj != ServerObjectType.Logon);

            // isPropertyLargerThanLimit, isValueContainsNotEnoughMemory
            isValueContainsNotEnoughMemory = false;
            if (requirementContainer[86703] && isPropertyLargerThanLimit && globalObj != ServerObjectType.Folder)
            {
                isValueContainsNotEnoughMemory = true;
                ModelHelper.CaptureRequirement(
                    644,
                    @"[In RopGetPropertiesAll ROP Response Buffer] PropertyValues: If the property value is larger than the size specified in the PropertySizeLimit field of the ROP request buffer, the type MUST be PtypErrorCode ([MS-OXCDATA] section 2.11.1) with a value of NotEnoughMemory ([MS-OXCDATA] section 2.4.2).");

                ModelHelper.CaptureRequirement(
                    85,
                    @"[In RopGetPropertiesAll ROP Request Buffer] PropertySizeLimit: If this value is nonzero, the property values are limited both by the size of the ROP response buffer and by the value of the PropertySizeLimit field. ");

                ModelHelper.CaptureRequirement(
                    86703,
                    @"Implementation does not ignore the PropertySizeLimit field. (<4> Section 3.2.5.1: Exchange 2003 and Exchange 2007 do not ignore the PropertySizeLimit field. When the property is a PtypBinary type, a PtypObject type, or a string property, Exchange 2003 and Exchange 2007 return the PtypErrorCode type with a value of NotEnoughMemory (0x8007000E) in place of the property value if the value is larger than either the available space in the ROP response buffer or the size specified in the PropertySizeLimit field of the ROP request buffer.)");
            }

            if (requirementContainer[90707] && isPropertyLargerThanLimit)
            {
                ModelHelper.CaptureRequirement(
                    90707,
                    @"[In Processing RopGetPropertiesAll] Implementation does return no error, no matter whether the property size exceed the value of PropertySizeLimit or not. (Microsoft Exchange Server 2010 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// Action for [RopGetPropertiesList] operations
        /// </summary>
        [Rule(Action = "RopGetPropertiesListMethod()")]
        public static void RopGetPropertiesListMethod()
        {
            Condition.IsTrue(isInitialized);
        }

        /// <summary>
        ///  Action for [RopGetPropertiesSpecificForWantUnicode] operation
        /// </summary>
        /// <param name="isUnicode">Indicates whether the requested property is encoded in Unicode format in response buffer</param>
        [Rule(Action = "RopGetPropertiesSpecificForWantUnicode(isUnicode)")]
        public static void RopGetPropertiesSpecificForWantUnicode(bool isUnicode)
        {
            Condition.IsTrue(isInitialized);
        }

        /// <summary>
        ///  Action for [RopGetPropertiesSpecific] operation
        /// </summary>
        /// <param name="isTestOrder">Indicates whether to test returned PropertyNames order</param>
        /// <param name="isPropertySizeLimitZero">Indicates whether PropertySizeLimit parameter is zero</param>
        /// <param name="isPropertyLargerThanLimit">Indicates whether request properties are larger than the limit
        /// When PropertySizeLimit is non-zero, it indicates whether request properties are larger than PropertySizeLimit
        /// When PropertySizeLimit is zero, it indicates whether request properties are larger than size of response</param>
        /// <param name="isValueContainsNotEnoughMemory">Indicates whether returned value contains NotEnoughMemory error when request properties are too large</param>
        [Rule(Action = "RopGetPropertiesSpecificMethod(isTestOrder, isPropertySizeLimitZero,isPropertyLargerThanLimit,out isValueContainsNotEnoughMemory)")]
        public static void RopGetPropertiesSpecificMethod(
            bool isTestOrder,
            bool isPropertySizeLimitZero,
            bool isPropertyLargerThanLimit,
            out bool isValueContainsNotEnoughMemory)
        {
            Condition.IsTrue(isInitialized);
            Condition.IfThen(isPropertyLargerThanLimit || isTestOrder, globalObj != ServerObjectType.Logon);

            isValueContainsNotEnoughMemory = false;
            if (requirementContainer[86703] && isPropertyLargerThanLimit && globalObj != ServerObjectType.Folder)
            {
                isValueContainsNotEnoughMemory = true;
                ModelHelper.CaptureRequirement(
                    63,
                    @"[In RopGetPropertiesSpecific ROP Request Buffer] PropertySizeLimit: If this value is nonzero, the property values are limited both by the size of the ROP response buffer and by the value of the PropertySizeLimit field. ");

                ModelHelper.CaptureRequirement(
                    86703,
                    @"Implementation does not ignore the PropertySizeLimit field. (<4> Section 3.2.5.1: Exchange 2003 and Exchange 2007 do not ignore the PropertySizeLimit field. When the property is a PtypBinary type, a PtypObject type, or a string property, Exchange 2003 and Exchange 2007 return the PtypErrorCode type with a value of NotEnoughMemory (0x8007000E) in place of the property value if the value is larger than either the available space in the ROP response buffer or the size specified in the PropertySizeLimit field of the ROP request buffer.)");

                ModelHelper.CaptureRequirement(
                    644,
                    @"[In RopGetPropertiesAll ROP Response Buffer] PropertyValues: If the property value is larger than the size specified in the PropertySizeLimit field of the ROP request buffer, the type MUST be PtypErrorCode ([MS-OXCDATA] section 2.11.1) with a value of NotEnoughMemory ([MS-OXCDATA] section 2.4.2).");
            }

            if (requirementContainer[9070102] && isPropertyLargerThanLimit)
            {
                ModelHelper.CaptureRequirement(
                    9070102,
                    @"[In Processing RopGetPropertiesSpecific] Implementation does return no error, no matter whether the property size exceed the value of PropertySizeLimit or not. (Microsoft Exchange Server 2010 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// Action for [RopGetPropertiesSpecific] operation
        /// </summary>
        [Rule(Action = "RopGetPropertiesSpecificForTaggedProperties()")]
        public static void RopGetPropertiesSpecificForTaggedProperties()
        {
            Condition.IsTrue(isInitialized);
            Condition.IsTrue(globalObj != ServerObjectType.Logon);
        }

        /// <summary>
        /// Action for [RopGetPropertyIdsFromNames] operation
        /// </summary>
        /// <param name="isTestOrder">Indicates whether to test returned PropertyNames order</param>
        /// <param name="isCreateFlagSet">Indicates whether the "Create" Flags in request parameter is set</param>
        /// <param name="isPropertyNameExisting">Indicates whether PropertyName is existing in object mapping</param>
        /// <param name="specialPropertyName">Specifies PropertyName of request parameter</param>
        /// <param name="isCreatedEntryReturned">If Create Flags is set:
        /// indicates that the server MUST create new entries for any name parameters that are not found in the existing mapping set, 
        /// and return existing entries for any name parameters that are found in the existing mapping set.
        /// </param>
        /// <param name="error">Specifies the ErrorCode when server reached limit</param>
        [Rule(Action = @"RopGetPropertyIdsFromNamesMethod(isTestOrder, isCreateFlagSet, isPropertyNameExisting, specialPropertyName, out isCreatedEntryReturned, out error)")]
        public static void RopGetPropertyIdsFromNamesMethod(
            bool isTestOrder,
            bool isCreateFlagSet,
            bool isPropertyNameExisting,
            SpecificPropertyName specialPropertyName,
            out bool isCreatedEntryReturned,
            out CPRPTErrorCode error)
        {
            Condition.IsTrue(isInitialized);

            isCreatedEntryReturned = false;
            error = CPRPTErrorCode.None;

            if (isCreateFlagSet && !isPropertyNameExisting)
            {
                isCreatedEntryReturned = true;
                ModelHelper.CaptureRequirement(
                    628,
                    @"[In RopGetPropertyIdsFromNames ROP Request Buffer] Flags: This field is set to 0x02 to request that a new entry be created for 
                    each named property that is not found in the existing mapping table; ");
            }

            if (!isCreateFlagSet)
            {
                isCreatedEntryReturned = false;
                ModelHelper.CaptureRequirement(
                    62801,
                    @"[In RopGetPropertyIdsFromNames ROP Request Buffer] Flags: This field is set to 0x00 otherwise[If not request a new entry to be created for each named property that is not found in the existing mapping table].");
            }
        }

        /// <summary>
        /// Action for [RopGetNamesFromPropertyIds] operation
        /// </summary>
        /// <param name="propertyIdType">Specifies different PropertyId type</param>
        [Rule(Action = "RopGetNamesFromPropertyIdsMethod(propertyIdType)")]
        public static void RopGetNamesFromPropertyIdsMethod(PropertyIdType propertyIdType)
        {
            Condition.IsTrue(isInitialized);
        }

        /// <summary>
        /// Action for [RopSetProperties] operation
        /// </summary>
        /// <param name="isModifiedValueReturned">
        /// Indicates whether the modified value of a property can be returned use a same handle
        /// </param>
        /// <param name="isChangedInDB">
        /// Indicates whether the modified value is submitted to DB.
        /// For Message and Attachment object, it requires another ROP for submit DB.
        /// For Logon and Folder object, it DOES NOT need any other ROPs for submit.
        /// </param>
        [Rule(Action = "RopSetPropertiesMethod(out isModifiedValueReturned,out isChangedInDB)")]
        public static void RopSetPropertiesMethod(out bool isModifiedValueReturned, out bool isChangedInDB)
        {
            Condition.IsTrue(isInitialized);

            isChangedInDB = false;
            isModifiedValueReturned = true;

            ModelHelper.CaptureRequirement(
                475,
                @"[In Processing RopSetProperties] The server MUST modify the value of each property specified in the PropertyValues field of the ROP request buffer.");

            ModelHelper.CaptureRequirement(
                101,
                @"[In RopSetProperties ROP] The RopSetProperties ROP ([MS-OXCROPS] section 2.2.8.6) updates the specified properties on an object. ");

            ModelHelper.CaptureRequirement(
                477,
                @"[In Processing RopSetProperties] For example, if the client uses the same object handle in a RopGetPropertiesAll ROP request ([MS-OXCROPS] section 2.2.8.4) 
                to read those same properties, the modified value MUST be returned. ");

            if (globalObj == ServerObjectType.Message)
            {
                isChangedInDB = false;

                ModelHelper.CaptureRequirement(
                    476,
                    @"[In Processing RopSetProperties] For Message objects, the new value of the properties 
                    MUST be made available immediately for retrieval by a ROP that uses  the same Message object handle. ");
            }

            if (globalObj == ServerObjectType.Attachment)
            {
                isChangedInDB = false;

                ModelHelper.CaptureRequirement(
                    479,
                    @"[In Processing RopSetProperties] For Attachment objects, the new value of the properties 
                    MUST be made available immediately for retrieval by a ROP that uses the same Attachment object handle. ");
            }

            // To verify whether the modified value persisted in Data base
            if (globalObj == ServerObjectType.Folder)
            {
                isChangedInDB = true;

                ModelHelper.CaptureRequirement(
                    846,
                    @"[In Processing RopSetProperties] For Folder objects,
                the new value of the properties MUST be persisted immediately without requiring another ROP to commit it.");
            }

            if (globalObj == ServerObjectType.Logon)
            {
                isChangedInDB = true;

                ModelHelper.CaptureRequirement(
                    847,
                    @"[In Processing RopSetProperties] For Logon objects, the new value of the properties MUST be persisted immediately without requiring another ROP to commit it.");
            }

            isSetPropertiesSuccess = true;
        }

        /// <summary>
        /// Action for [RopDeleteProperties] operation
        /// </summary>
        /// <param name="isPropertiesDeleted">
        /// If the server returns success, Properties will be deleted.
        /// </param>
        /// <param name="isChangedInDB">
        /// Indicates whether the modified value is submitted to DB.
        /// For Message and Attachment object, it requires another ROP for submit DB.
        /// For Logon and Folder object, it DOES NOT need any other ROPs for submit.
        /// </param>
        [Rule(Action = "RopDeletePropertiesMethod(out isPropertiesDeleted,out isChangedInDB)")]
        public static void RopDeletePropertiesMethod(out bool isPropertiesDeleted, out bool isChangedInDB)
        {
            Condition.IsTrue(isInitialized);

            isPropertiesDeleted = true;
            isChangedInDB = false;

            ModelHelper.CaptureRequirement(
                485,
                @"[In Processing RopDeleteProperties] If the server returns success, 
                it MUST NOT have a valid value to return to a client that asks for the value of this property. ");

            ModelHelper.CaptureRequirement(
                486,
                "[In Processing RopDeleteProperties] [In Processing RopDeleteProperties] The server MUST delete the property from the object.");

            // For Message object, verify MS-OXCPRPT_R487, MS-OXCPRPT_R488
            if (globalObj == ServerObjectType.Message)
            {
                isChangedInDB = false;

                ModelHelper.CaptureRequirement(
                    487,
                    @"[In Processing RopDeleteProperties] For Message objects, the properties MUST be removed immediately when using the same handle. ");

                ModelHelper.CaptureRequirement(
                    488,
                    @"[In Processing RopDeleteProperties] In other words, if the client uses the same handle to read those same properties using 
                    the RopGetPropertiesSpecific ROP ([MS-OXCROPS] section 2.2.8.3) or the RopGetPropertiesAll ROP 
                    ([MS-OXCROPS] section 2.2.8.4), the properties MUST be deleted. ");
            }

            // For attachment object, verify MS-OXCPRPT_R490
            if (globalObj == ServerObjectType.Attachment)
            {
                isChangedInDB = false;

                ModelHelper.CaptureRequirement(
                    490,
                    @"[In Processing RopDeleteProperties] For Attachment objects, the properties MUST be removed immediately when using the same handle.");
            }

            // For Folder objects and Logon objects, verify MS-OXCPRPT_R492
            if (globalObj == ServerObjectType.Folder)
            {
                isChangedInDB = true;
                ModelHelper.CaptureRequirement(
                    492,
                    @"[In Processing RopDeleteProperties] For Folder objects, the properties 
                    MUST be removed immediately without requiring another ROP to commit the change.");
            }

            if (globalObj == ServerObjectType.Logon)
            {
                if (!requirementContainer[49201])
                {
                    isChangedInDB = false;
                    isPropertiesDeleted = false;
                }
                else
                {
                    isChangedInDB = true;
                    ModelHelper.CaptureRequirement(
                        49201,
                        @"[In Processing RopDeleteProperties] For Logon objects, the properties MUST be removed immediately without requiring another ROP to commit the change.");
                }
            }

            isDeletePropertiesSuccess = true;

            ModelHelper.CaptureRequirement(
                    113,
                    @"[In RopDeleteProperties ROP] The RopDeleteProperties ROP ([MS-OXCROPS] section 2.2.8.8) removes the specified properties from an object. ");
        }

        /// <summary>
        /// Action for [RopSaveChangesMessage] operation
        /// </summary>
        /// <param name="isChangedInDB">Indicates whether changes of Message object submit to database 
        /// when [RopSetProperties] or [RopDeleteProperties]</param>
        [Rule(Action = "RopSaveChangesMessageMethod(out isChangedInDB)")]
        public static void RopSaveChangesMessageMethod(out bool isChangedInDB)
        {
            Condition.IsTrue(isInitialized);
            Condition.IsTrue((globalObj == ServerObjectType.Message || globalObj == ServerObjectType.Attachment) && (isCommitStreamSuccess || isSetPropertiesSuccess || isDeletePropertiesSuccess || isCopyPropertiesSuccess || isCopyToSuccess));
            isChangedInDB = true;
            if (globalObj == ServerObjectType.Message)
            {
                if (isSetPropertiesSuccess)
                {
                    ModelHelper.CaptureRequirement(
                        478,
                        @"[In Processing RopSetProperties] However, the modified value MUST NOT be persisted to the database until a successful RopSaveChangesMessage ROP is issued.");

                    ModelHelper.CaptureRequirement(
                       82401,
                       @"[In RopSetPropertiesNoReplicate ROP] However, the modified value MUST NOT be persisted to the database until a successful RopSaveChangesMessage ROP ([MS-OXCROPS] section 2.2.6.3) is issued.");
                }

                if (isDeletePropertiesSuccess)
                {
                    ModelHelper.CaptureRequirement(
                        489,
                        @"[In Processing RopDeleteProperties] [for message object] However, the deleted properties MUST NOT be persisted to the database until a successful RopSaveChangesMessage ROP ([MS-OXCROPS] section 2.2.6.3) is issued.");

                    ModelHelper.CaptureRequirement(
                        83606,
                        @"[In RopDeletePropertiesNoReplicate] However, the deleted properties MUST NOT be persisted to the database until a successful RopSaveChangesMessage ROP ([MS-OXCROPS] section 2.2.6.3) is issued.");
                }

                if (isCopyPropertiesSuccess)
                {
                    ModelHelper.CaptureRequirement(
                      849,
                      @"[In Processing RopCopyProperties] In the case of Message objects, the changes on either source or destination MUST NOT be persisted until the RopSaveChangesMessage ROP ([MS-OXCROPS] section 2.2.6.3) is successfully issued.");
                }

                if (isCopyToSuccess)
                {
                    ModelHelper.CaptureRequirement(
                       5070507,
                       @"[In Processing RopCopyTo] In the case of Message objects, the changes on either source or destination MUST NOT be persisted until the RopSaveChangesMessage ROP ([MS-OXCROPS] section 2.2.6.3) is successfully issued.");
                }
            }

            if (globalObj == ServerObjectType.Attachment)
            {
                Condition.IsTrue(isSaveChangesAttachmentSuccess);
                if (isSetPropertiesSuccess)
                {
                    ModelHelper.CaptureRequirement(
                        480,
                        @"[In Processing RopSetProperties] However, the modified value MUST NOT be persisted to the database until a successful RopSaveChangesAttachment ROP followed by a successful RopSaveChangesMessage ROP is issued.");
                    ModelHelper.CaptureRequirement(
                        82501,
                        @"[In RopSetPropertiesNoReplicate ROP] However, the modified value MUST NOT be persisted to the database until a successful RopSaveChangesAttachment ROP ([MS-OXCROPS] section 2.2.6.15) followed by a successful RopSaveChangesMessage ROP is issued.");
                }

                if (isCopyPropertiesSuccess)
                {
                    ModelHelper.CaptureRequirement(
                        848,
                        @"[In Processing RopCopyProperties] In the case of Attachment objects, the changes on either source or destination MUST NOT be persisted until the RopSaveChangesAttachment ROP ([MS-OXCROPS] section 2.2.6.15) and the RopSaveChangesMessage ROP are successfully issued, in that order.");
                }

                if (isDeletePropertiesSuccess)
                {
                    ModelHelper.CaptureRequirement(
                        491,
                        @"[In Processing RopDeleteProperties] However, the deleted properties MUST NOT be persisted to the database until a successful RopSaveChangesAttachment ROP ([MS-OXCROPS] section 2.2.6.15) and a successful RopSaveChangesMessage ROP are issued, in that order.");

                    ModelHelper.CaptureRequirement(
                       83608,
                       @"[In RopDeletePropertiesNoReplicate] However, the deleted properties MUST NOT be persisted to the database until a successful RopSaveChangesAttachment ROP ([MS-OXCROPS] section 2.2.6.15) and a successful RopSaveChangesMessage ROP are issued, in that order.");
                }

                // Reset isSaveChangesAttachmentSuccess
                isSaveChangesAttachmentSuccess = false;
            }
        }

        /// <summary>
        /// Action for [RopSaveChangesAttachment] operation
        /// </summary>
        /// <param name="isChangedInDB">Indicates whether changes of Attachment object submit to database 
        /// when [RopSetProperties] or [RopDeleteProperties]</param>
        [Rule(Action = "RopSaveChangesAttachmentMethod(out isChangedInDB)")]
        public static void RopSaveChangesAttachmentMethod(out bool isChangedInDB)
        {
            Condition.IsTrue(isInitialized);
            Condition.IsTrue(globalObj == ServerObjectType.Attachment && (isSetPropertiesSuccess || isDeletePropertiesSuccess || isCopyPropertiesSuccess));

            isChangedInDB = false;
            isSaveChangesAttachmentSuccess = true;
        }

        /// <summary>
        /// Action for [RopSetPropertiesNoReplicate] operation
        /// </summary>
        /// <param name="isModifiedValueReturned">
        /// Indicates whether the modified value of a property can be returned use a same handle
        /// </param>
        [Rule(Action = "RopSetPropertiesNoReplicateMethod(out isModifiedValueReturned)")]
        public static void RopSetPropertiesNoReplicateMethod(out bool isModifiedValueReturned)
        {
            Condition.IsTrue(isInitialized);
            isModifiedValueReturned = true;

            ModelHelper.CaptureRequirement(
                822,
                "[In RopSetPropertiesNoReplicate ROP] The server MUST modify the value of the properties of the object.");

            ModelHelper.CaptureRequirement(
                824,
                @"[In RopSetPropertiesNoReplicate ROP] If the client uses the same object handle in a RopGetPropertiesAll ROP request ([MS-OXCROPS] section 2.2.8.4) to read those same properties, the modified value MUST be returned.");

            if (globalObj == ServerObjectType.Message)
            {
                ModelHelper.CaptureRequirement(
                    823,
                    @"[In RopSetPropertiesNoReplicate ROP] For Message objects, the new value of the properties MUST be made available immediately when using the same handle.");
            }

            if (globalObj == ServerObjectType.Attachment)
            {
                ModelHelper.CaptureRequirement(
                    825,
                    @"[In RopSetPropertiesNoReplicate ROP] For Attachment objects, the new value of the properties MUST be made available immediately when using the same handle.");
            }

            if (globalObj == ServerObjectType.Logon)
            {
                ModelHelper.CaptureRequirement(
                    82502,
                    @"[In RopSetPropertiesNoReplicate ROP] For Logon objects, the new value of the properties MUST be persisted immediately without requiring another ROP to commit it.");
            }

            ModelHelper.CaptureRequirement(
                112,
                @"[In RopSetPropertiesNoReplicate ROP] On all other objects [Message objects, Attachment objects and Logon objects], the RopSetPropertiesNoReplicate ROP works the same way as the RopSetProperties ROP.");

            isSetPropertiesSuccess = true;
        }

        /// <summary>
        /// Action for [RopDeletePropertiesNoReplicate] operation
        /// </summary>
        /// <param name="isPropertiesDeleted">Indicates whether result is same as RopDeleteProperties </param>
        /// <param name="isChangedInDB">
        /// Indicates whether the modified value is submitted to DB.
        /// For Message and Attachment object, it requires another ROP for submit DB.
        /// For Logon and Folder object, it DOES NOT need any other ROPs for submit.
        /// </param>
        [Rule(Action = "RopDeletePropertiesNoReplicateMethod(out isPropertiesDeleted,out isChangedInDB)")]
        public static void RopDeletePropertiesNoReplicateMethod(out bool isPropertiesDeleted, out bool isChangedInDB)
        {
            Condition.IsTrue(isInitialized);

            isPropertiesDeleted = true;
            isChangedInDB = false;

            ModelHelper.CaptureRequirement(
                835,
                @"[In RopDeletePropertiesNoReplicate] If the server returns success 
                it MUST NOT have a valid value to return to a client that asks for the value of this property.");

            ModelHelper.CaptureRequirement(
                836, "[In RopDeletePropertiesNoReplicate] The server MUST remove the value of the property from the object.");

            ModelHelper.CaptureRequirement(
                124,
                @"[In RopDeletePropertiesNoReplicate ROP] On all other objects [Message objects, Attachment objects and Logon objects], the RopDeletePropertiesNoReplicate ROP works the same way as RopDeleteProperties.");

            if (globalObj == ServerObjectType.Message)
            {
                isChangedInDB = false;

                ModelHelper.CaptureRequirement(
                    83604,
                    @"[In RopDeletePropertiesNoReplicate] For Message objects, the properties MUST be removed immediately when using the same handle.");

                ModelHelper.CaptureRequirement(
                    83605,
                    @"[In RopDeletePropertiesNoReplicate] In other words, if the client uses the same handle to read those same properties using the RopGetPropertiesSpecific ROP ([MS-OXCROPS] section 2.2.8.3), the properties MUST be deleted. ");

                ModelHelper.CaptureRequirement(
                    8360501,
                    @"[In RopDeletePropertiesNoReplicate] In other words, if the client uses the same handle to read those same properties using the RopGetPropertiesAll ROP ([MS-OXCROPS] section 2.2.8.4), the properties MUST be deleted. ");
            }

            if (globalObj == ServerObjectType.Attachment)
            {
                isChangedInDB = false;

                ModelHelper.CaptureRequirement(
                    83607,
                    @"[In RopDeletePropertiesNoReplicate] For Attachment objects, the properties MUST be removed immediately when using the same handle.");
            }

            if (globalObj == ServerObjectType.Folder)
            {
                isChangedInDB = true;
            }

            if (globalObj == ServerObjectType.Logon)
            {
                if (!requirementContainer[83609])
                {
                    isChangedInDB = false;
                    isPropertiesDeleted = false;
                }
                else
                {
                    isChangedInDB = true;
                    ModelHelper.CaptureRequirement(
                        83609,
                        @"[In RopDeletePropertiesNoReplicate] For Logon objects, the properties MUST be removed immediately without requiring another ROP to commit the change.");
                }
            }

            isDeletePropertiesSuccess = true;
        }

        /// <summary>
        /// Action for [RopOpenStream] operation
        /// </summary>
        /// <param name="objectToOperate">Specifies which object will be operated on</param>
        /// <param name="openFlag">Specifies OpenModeFlags for [RopOpenStream]</param>
        /// <param name="isPropertyTagExist">Indicates whether request property exist</param>
        /// <param name="isStreamSizeEqualToStream">Indicates whether StreamSize in response is the same with 
        /// the current number of BYTES in the stream</param>
        /// <param name="error">
        /// If the property tag does not exist for the object and "Create" is not specified in OpenModeFlags
        /// NotFound error should be returned
        /// </param>
        [Rule(Action = "RopOpenStreamMethod(objectToOperate, openFlag, isPropertyTagExist,out isStreamSizeEqualToStream,out error)")]
        public static void RopOpenStreamMethod(
            ObjectToOperate objectToOperate,
            OpenModeFlags openFlag,
            bool isPropertyTagExist,
            out bool isStreamSizeEqualToStream,
            out CPRPTErrorCode error)
        {
            Condition.IsTrue(isInitialized);
            Condition.IsTrue(globalObj != ServerObjectType.Logon);

            // openFlag and error is designed for negative test case.
            error = CPRPTErrorCode.None;
            isStreamSizeEqualToStream = false;
            isStreamWriteSuccess = false;
            isStreamOpenedSuccess = true;

            streamOpenFlag = openFlag;
            ModelHelper.CaptureRequirement(
                885,
                @"[In Processing RopOpenStream] The server MUST open the stream in the mode indicated by the OpenModeFlags field as specified by the table in section 2.2.14.1.");

            if (isPropertyTagExist)
            {
                isStreamSizeEqualToStream = true;
            }
            else if (openFlag == OpenModeFlags.ReadOnly && isPropertyTagExist == false)
            {
                error = CPRPTErrorCode.NotFound;
                isStreamOpenedSuccess = false;
            }
        }

        /// <summary>
        /// Action for [RopOpenStream] operation.
        /// </summary>
        /// <param name="objectToOperate">Specifies which object will be operated on.</param>
        /// <param name="propertyType">Specifies which type of property will be operated on.</param>
        /// <param name="error">Returned error code.</param>
        [Rule(Action = "RopOpenStreamWithDifferentPropertyType(objectToOperate, propertyType, out error)")]
        public static void RopOpenStreamWithDifferentPropertyType(
            ObjectToOperate objectToOperate,
            PropertyTypeName propertyType,
            out CPRPTErrorCode error)
        {
            Condition.IsTrue(isInitialized);
            Condition.IsTrue(globalObj != ServerObjectType.Logon);
            Condition.IfThen(globalObj == ServerObjectType.Message || globalObj == ServerObjectType.Attachment, propertyType == PropertyTypeName.PtypBinary || propertyType == PropertyTypeName.PtypString || propertyType == PropertyTypeName.PtypObject);

            if (objectToOperate == ObjectToOperate.FirstObject)
            {
                Condition.IfThen(globalObj == ServerObjectType.Folder, propertyType == PropertyTypeName.PtypBinary);
            }
            else if (objectToOperate == ObjectToOperate.FifthObject)
            {
                Condition.IfThen(globalObj == ServerObjectType.Folder, propertyType == PropertyTypeName.PtypBinary || propertyType == PropertyTypeName.PtypString);
            }

            // openFlag and error is designed for negative test case.
            error = CPRPTErrorCode.None;            
            if (globalObj == ServerObjectType.Attachment && propertyType == PropertyTypeName.PtypBinary)
            {
                ModelHelper.CaptureRequirement(
                    25502,
                    @"[In RopOpenStream ROP] Single-valued PtypBinary type properties ([MS-OXCDATA] section 2.11.1) is supported for Attachment objects.");
            }

            if (globalObj == ServerObjectType.Attachment && propertyType == PropertyTypeName.PtypObject)
            {
                ModelHelper.CaptureRequirement(
                    25503,
                    @"[In RopOpenStream ROP] Single-valued PtypObject type properties ([MS-OXCDATA] section 2.11.1) is supported for Attachment objects.");
            }

            if (globalObj == ServerObjectType.Message && propertyType == PropertyTypeName.PtypBinary)
            {
                ModelHelper.CaptureRequirement(
                   25506,
                   @"[In RopOpenStream ROP] Single-valued PtypBinary type properties ([MS-OXCDATA] section 2.11.1) is supported for Message objects.");
            }

            if (globalObj == ServerObjectType.Message && propertyType == PropertyTypeName.PtypObject)
            {
                ModelHelper.CaptureRequirement(
                   25507,
                   @"[In RopOpenStream ROP] Single-valued PtypObject type properties ([MS-OXCDATA] section 2.11.1) is supported for Message objects.");
            }

            if (globalObj == ServerObjectType.Message && propertyType == PropertyTypeName.PtypString)
            {
                ModelHelper.CaptureRequirement(
                   25509,
                   @"[In RopOpenStream ROP] Single-valued PtypString type properties ([MS-OXCDATA] section 2.11.1) is supported for Message objects.");
            }

            if (globalObj == ServerObjectType.Attachment && propertyType == PropertyTypeName.PtypString)
            {
                if (requirementContainer[25505])
                {
                    ModelHelper.CaptureRequirement(
                        25505,
                        @"[In RopOpenStream ROP] Single-valued PtypString type properties ([MS-OXCDATA] section 2.11.1) is supported for Attachment objects.");
                }
                else
                {
                    error = CPRPTErrorCode.NotSupported;
                }            
            }
        }

        /// <summary>
        /// Action for [RopReadStream] operation
        /// </summary>
        /// <param name="isReadingFailed">Indicates whether reading stream get failure. E.g. object handle is not stream</param>
        [Rule(Action = "RopReadStreamMethod(isReadingFailed)")]
        public static void RopReadStreamMethod(bool isReadingFailed)
        {
            Condition.IsTrue(isInitialized && isStreamOpenedSuccess);
        }
        
        /// <summary>
        /// Action for [RopReadStream] operation
        /// </summary>
        /// <param name="byteCount">Indicates the size to be read.</param>
        /// <param name="maxByteCount">If byteCount is 0xBABE, use MaximumByteCount to determine the size to be read.</param>
        /// Disable CA1801, because the parameter 'isReadingFailed' is used for interface implementation.
        [Rule(Action = "RopReadStreamWithLimitedSize(byteCount,maxByteCount)")]
        public static void RopReadStreamWithLimitedSize(ushort byteCount, uint maxByteCount)
        {
            Condition.IsTrue(isInitialized && isStreamOpenedSuccess);
        }

        /// <summary>
        /// Action for [RopSeekStream]
        /// </summary>
        /// <param name="condition">Specifies particular scenario of RopSeekStream</param>
        /// <param name="isStreamExtended">Indicates whether a stream object is extended and zero filled to the new seek location</param>
        /// <param name="error">Returned error code</param>
        [Rule(Action = "RopSeekStreamMethod(condition,out isStreamExtended, out error)")]
        public static void RopSeekStreamMethod(SeekStreamCondition condition, out bool isStreamExtended, out CPRPTErrorCode error)
        {
            Condition.IsTrue(isInitialized && isStreamOpenedSuccess);

            if (condition == SeekStreamCondition.MovedBeyondMaxStreamSize)
            {
                error = CPRPTErrorCode.StreamSeekError;
                isStreamExtended = false;

                ModelHelper.CaptureRequirement(
                    582,
                    "[In Processing RopSeekStream] If the client requests the seek pointer be moved beyond 2^31 bytes, the server MUST return the StreamSeekError error code in the ReturnValue field of the ROP response buffer.");
            }
            else if (condition == SeekStreamCondition.MovedBeyondEndOfStream)
            {
                error = CPRPTErrorCode.None;
                isStreamExtended = true;
                ModelHelper.CaptureRequirement(
                    583,
                    @"[In Processing RopSeekStream] If the client requests the seek pointer be moved beyond the end of the stream, the stream is extended, and zeros filled to the new seek location.");
            }
            else if (condition == SeekStreamCondition.OriginInvalid)
            {
                error = CPRPTErrorCode.StreamInvalidParam;
                isStreamExtended = false;
            }
            else
            {
                error = CPRPTErrorCode.None;
                isStreamExtended = false;
            }
        }

        /// <summary>
        /// Action for [RopLockRegionStream] operation
        /// </summary>
        /// <param name="preState">Specifies the pre-state before call [RopLockRegionStream]</param>
        /// <param name="error">Returned error code</param>
        [Rule(Action = "RopLockRegionStreamMethod(preState, out error)")]
        public static void RopLockRegionStreamMethod(PreStateBeforeLock preState, out CPRPTErrorCode error)
        {
            Condition.IsTrue(isInitialized && isStreamOpenedSuccess);

            // RopUnlockRegionStream is only implemented on Exchange server 2007.
            Condition.IsTrue(requirementContainer[750]);
            error = CPRPTErrorCode.None;
            if (preState == PreStateBeforeLock.WithExpiredLock || preState == PreStateBeforeLock.Normal)
            {
                error = CPRPTErrorCode.None;
                ModelHelper.CaptureRequirement(
                     610,
                     "[In Processing RopLockRegionStream] If the server implements this ROP, if all previous locks are expired, or if there are no previous locks, the server MUST grant the requesting client a new lock.");
            }

            isStreamLocked = true;
        }

        /// <summary>
        /// Action for [RopUnlockRegionStream] operation
        /// </summary>
        /// <param name="isPreviousLockExists">Indicates whether a previous lock exists and not owned by this session</param>
        [Rule(Action = "RopUnlockRegionStreamMethod(isPreviousLockExists)")]
        public static void RopUnlockRegionStreamMethod(bool isPreviousLockExists)
        {
            Condition.IsTrue(isInitialized && isStreamOpenedSuccess && isStreamLocked);

            // RopUnlockRegionStream is only implemented on Exchange server 2007.
            Condition.IsTrue(requirementContainer[751]);
        }

        /// <summary>
        /// Action for [RopWriteStream] operation
        /// </summary>
        /// <param name="openFlag">Specifies the OpenModeFlags of the stream</param>
        /// <param name="isExceedMax">Indicates whether the write will exceed the maximum stream size.</param>
        /// <param name="error"> Specifies the ErrorCode when WriteStream failed</param>
        [Rule(Action = "RopWriteStreamMethod(openFlag, isExceedMax, out error)")]
        public static void RopWriteStreamMethod(OpenModeFlags openFlag, bool isExceedMax, out CPRPTErrorCode error)
        {
            Condition.IsTrue(isInitialized && isStreamOpenedSuccess);
            Condition.IsTrue(openFlag == streamOpenFlag);
            Condition.IsTrue(globalObj != ServerObjectType.Logon);
            Condition.IfThen(isExceedMax, openFlag == OpenModeFlags.ReadWrite);

            error = CPRPTErrorCode.None;
            isStreamWriteSuccess = true;
            if (streamOpenFlag == OpenModeFlags.ReadWrite)
            {
                ModelHelper.CaptureRequirement(
                       269,
                       @"[In RopOpenStream ROP Request Buffer] OpenModeFlags: ReadWrite: Open the stream for read/write access.");
            }

            if (streamOpenFlag == OpenModeFlags.ReadOnly)
            {
                error = CPRPTErrorCode.STG_E_ACCESSDENIED;
                isStreamWriteSuccess = false;
                ModelHelper.CaptureRequirement(
                    267,
                    @"[In RopOpenStream ROP Request Buffer] OpenModeFlags: ReadOnly: Open the stream for read-only access.");
            }

            if (isExceedMax)
            {
                // For ExchangeServer 2007, StreamSizeError error code returned.
                if (requirementContainer[86706])
                {
                    error = CPRPTErrorCode.StreamSizeError;
                    ModelHelper.CaptureRequirement(
                        86706,
                        @"[In Appendix A: Product Behavior] Implementation does return the StreamSizeError error code. (<12> Section 3.2.5.13: Exchange 2003 and Exchange 2007 return the StreamSizeError error code if they write less than the amount requested.)");
                }

                // For Exchange 2010, toobig error code returned.
                if (requirementContainer[55707])
                {
                    error = CPRPTErrorCode.ecTooBig;
                    ModelHelper.CaptureRequirement(
                        55707,
                        @"[In Processing RopWriteStream] Implementation does  return the TooBig error code if it writes less than the amount requested.(Microsoft Exchange Server 2010 and above follow this behavior)");
                }

                if (requirementContainer[90102])
                {
                    error = CPRPTErrorCode.ecTooBig;
                    ModelHelper.CaptureRequirement(
                        90102,
                        @"[In Processing RopWriteStream] Implementation does return error code ""0x80040305"" with name ""TooBig"", when the write will exceed the maximum stream size.(Microsoft Exchange Server 2007 and above follow this behavior)");
                }

                isStreamWriteSuccess = false;
            }
        }

        /// <summary>
        /// Action for [RopCommitStream] operations
        /// </summary>
        /// <param name="openFlag">Indicates the OpenModeFlags when stream is opened</param>
        /// <param name="isPropertyValueChanged">Indicates whether property value is changed
        /// </param>
        [Rule(Action = "RopCommitStreamMethod(openFlag, out isPropertyValueChanged)")]
        public static void RopCommitStreamMethod(OpenModeFlags openFlag, out bool isPropertyValueChanged)
        {
            Condition.IsTrue(isInitialized && isStreamOpenedSuccess);
            Condition.IsTrue(globalObj != ServerObjectType.Logon);
            Condition.IsTrue(streamOpenFlag == openFlag);
            if (isStreamWriteSuccess)
            {
                isPropertyValueChanged = true;
                ModelHelper.CaptureRequirement(
                    305,
                    @"[In RopCommitStream ROP] The RopCommitStream ROP ([MS-OXCROPS] section 2.2.9.4) ensures that any changes made to a Stream object are persisted in storage.");

                if (globalObj == ServerObjectType.Folder)
                {
                    ModelHelper.CaptureRequirement(
                       56405,
                       @"[In Processing RopWriteStream] For a Folder object, the value is persisted when the RopCommitStream ROP ([MS-OXCROPS] section 2.2.9.4) is issued on the Stream object or the Stream object is closed with a RopRelease ROP ([MS-OXCROPS] section 2.2.15.3).");
                }
            }
            else
            {
                isPropertyValueChanged = false;
            }

            isCommitStreamSuccess = true;
        }

        /// <summary>
        /// Action for RopRelease operation
        /// The client uses RopRelease ([MS-OXCROPS] section 2.2.14.3) after it is done with the Stream object.
        /// </summary>
        /// <param name="obj">Specifies which object will be operated on.</param>
        /// <param name="isPropertyValueChanged">
        /// For Folder Object, this ROP should not change the value in stream after RopWriteStream.
        /// For non-Folder Object, this ROP should change the value.
        /// </param>
        [Rule(Action = "RopReleaseMethod(obj, out isPropertyValueChanged)")]
        public static void RopReleaseMethod(ObjectToOperate obj, out bool isPropertyValueChanged)
        {
            Condition.IsTrue(isInitialized && isStreamOpenedSuccess);
            isPropertyValueChanged = false;

            if (isStreamWriteSuccess)
            {
                isPropertyValueChanged = true;
                if (globalObj == ServerObjectType.Attachment)
                {
                    ModelHelper.CaptureRequirement(
                        56404,
                        @"[In Processing RopWriteStream] For an Attachment object, the new value is persisted to the database when a successful RopSaveChangesAttachment ROP ([MS-OXCROPS] section 2.2.6.15) followed by a successful RopSaveChangesMessage ROP is issued.");
                }

                if (globalObj == ServerObjectType.Message)
                {
                    ModelHelper.CaptureRequirement(
                          56403,
                          @"[In Processing RopWriteStream] For a Message object, the new value is persisted to the database when a successful RopSaveChangesMessage ROP ([MS-OXCROPS] section 2.2.6.3) is issued.");
                }
            }

            // Check if the previous RopWriteStream is success
            if (isStreamWriteSuccess)
            {
                isPropertyValueChanged = true;
            }
        }

        /// <summary>
        /// Action for RopRelease operation
        /// The client uses RopRelease ([MS-OXCROPS] section 2.2.14.3) after it is done with the Stream object.
        /// </summary>
        /// <param name="obj">Specifies which object will be operated on</param>
        [Rule(Action = "RopReleaseMethodNoVerify(obj)")]
        public static void RopReleaseMethodNoVerify(ObjectToOperate obj)
        {
            Condition.IsTrue(isInitialized && isStreamOpenedSuccess);
        }

        /// <summary>
        /// Action for [RopCopyToStream] operation
        /// </summary>
        /// <param name="isDestinationExist">Specified the whether the destination existed.</param>
        /// <param name="isReadWriteSuccess"> When call success: The server MUST read the number of BYTES requested from the source Stream object, 
        /// and write those bytes into the destination Stream object</param>
        /// <param name="error"> If Destination object does not exist, expect DestinationNullObject error </param>
        [Rule(Action = "RopCopyToStreamMethod(isDestinationExist,out isReadWriteSuccess,out error)")]
        public static void RopCopyToStreamMethod(bool isDestinationExist, out bool isReadWriteSuccess, out CPRPTErrorCode error)
        {
            Condition.IsTrue(isInitialized && isStreamOpenedSuccess);
            Condition.IsTrue(isFirstObjectGot && isSecondObjectGot);

            if (isDestinationExist)
            {
                isReadWriteSuccess = true;
                error = CPRPTErrorCode.None;

                if (requirementContainer[867091])
                {
                    ModelHelper.CaptureRequirement(
                        867091,
                        @"[In Appendix A: Product Behavior] Implementation does implement the RopCopyToStream ROP. (<13> Section 3.2.5.18: Exchange 2007 and above products (except the initial release version of Exchange 2010) follow this behavior.)");
                }
                
                ModelHelper.CaptureRequirement(
                    592,
                    @"[In Processing RopCopyToStream] The server MUST read the number of bytes requested from the source Stream object and write those bytes into the destination Stream object.");
            }
            else
            {
                isReadWriteSuccess = false;
                error = CPRPTErrorCode.NullDestinationObject;
            }

            isStreamWriteSuccess = false;
        }

        /// <summary>
        /// Action for [RopWriteAndCommitStream] operation
        /// </summary>
        /// <param name="error">
        /// This ROP MUST NOT be used on Stream objects opened on properties on Folder objects
        /// which means it should be failed against Folder object
        /// </param>
        [Rule(Action = "RopWriteAndCommitStreamMethod(out error)")]
        public static void RopWriteAndCommitStreamMethod(out CPRPTErrorCode error)
        {
            // Exchange 2010 does not implement this ROP.
            Condition.IsTrue(requirementContainer[752]);
            isStreamWriteSuccess = true;

            error = CPRPTErrorCode.None;
        }

        /// <summary>
        /// Action for [RopCloneStream] operation
        /// </summary>
        [Rule(Action = "RopCloneStreamMethod")]
        public static void RopCloneStreamMethod()
        {
            Condition.IsTrue(isInitialized && isSecondObjectGot && isFirstObjectGot && isStreamOpenedSuccess);

            // RopCloneStream is not implemented on Exchange server 2010.
            Condition.IsTrue(requirementContainer[753]);
        }

        /// <summary>
        /// Action for [RopSetStreamSize] operation
        /// </summary>
        /// <param name="isSizeIncreased">
        /// Indicates the new size is increased or decreased 
        /// </param>
        /// <param name="isExtendedValueZero">
        /// If the size of the stream is increased, then value of the extended stream MUST be zero,
        /// </param>
        /// <param name="isLost">
        /// If the size of the stream is decreased, the information that extends past the end of the new size lost
        /// </param>
        /// <param name="isIncrease">If the size of the stream is increased, set this value to true</param>
        [Rule(Action = "RopSetStreamSizeMethod(isSizeIncreased,out isExtendedValueZero,out isLost,out isIncrease)")]
        public static void RopSetStreamSizeMethod(bool isSizeIncreased, out bool isExtendedValueZero, out bool isLost, out bool isIncrease)
        {
            // This operation requires an opened stream
            Condition.IsTrue(isInitialized && isStreamOpenedSuccess);
            Condition.IsTrue(globalObj != ServerObjectType.Folder);

            isExtendedValueZero = false;
            isLost = false;
            isIncrease = false;
            if (isSizeIncreased)
            {
                isIncrease = true;
                ModelHelper.CaptureRequirement(
                   322,
                   "[In RopSetStreamSize ROP] The RopSetStreamSize ROP ([MS-OXCROPS] section 2.2.9.6) sets the size of a stream. ");

                ModelHelper.CaptureRequirement(
                 32801,
                 "[In RopSetStreamSize ROP Request Buffer] StreamSize: An integer that specifies the size, in bytes, of the stream. ");

                ModelHelper.CaptureRequirement(
                57705,
                "[In Processing RopSetStreamSize] The server sets the current size of the Stream object according to the value specified in the StreamSize field of the ROP request buffer. ");

                isExtendedValueZero = true;
                ModelHelper.CaptureRequirement(
                    57706,
                    "[In Processing RopSetStreamSize] If the size of the stream is increased, the server MUST set the values in the extended stream to 0x00.");
            }
            else
            {
                isLost = true;
                ModelHelper.CaptureRequirement(
                    57707,
                    "[In Processing RopSetStreamSize] If the size of the stream is decreased, the server discards the values that are beyond the end of the new size.");

                ModelHelper.CaptureRequirement(
                   322,
                   "[In RopSetStreamSize ROP] The RopSetStreamSize ROP ([MS-OXCROPS] section 2.2.9.6) sets the size of a stream. ");

                ModelHelper.CaptureRequirement(
                32801,
                "[In RopSetStreamSize ROP Request Buffer] StreamSize: An integer that specifies the size, in bytes, of the stream. ");
                ModelHelper.CaptureRequirement(
                57705,
                "[In Processing RopSetStreamSize] The server sets the current size of the Stream object according to the value specified in the StreamSize field of the ROP request buffer. ");
            }
        }

        /// <summary>
        /// Action for [RopGetStreamSize] operation
        /// </summary>
        [Rule(Action = "RopGetStreamSizeMethod()")]
        public static void RopGetStreamSizeMethod()
        {
            // This operation requires an opened stream
            Condition.IsTrue(isInitialized && isStreamOpenedSuccess);
        }

        /// <summary>
        /// First part for Action [RopCopyProperies] operation
        /// </summary>
        /// <param name="copyFlag">Specifies the CopyFlags in the call request</param>
        /// <param name="isWantAsynchronousZero">Indicates whether WantAsynchronous parameter in call request is zero</param>
        /// <param name="isDestinationExist">Indicates whether destination object is exist for [RopCopyProperties]</param>
        [Rule(Action = @"call RopCopyPropertiesMethod(copyFlag,isWantAsynchronousZero,isDestinationExist, out _, out _, out _, out _, out _)")]
        public static void RopCopyPropertiesMethodCall(CopyFlags copyFlag, bool isWantAsynchronousZero, bool isDestinationExist)
        {
            Condition.IsTrue(
                isInitialized && isFirstObjectGot && isSecondObjectGot &&
                (globalObj == ServerObjectType.Message ||
                globalObj == ServerObjectType.Attachment ||
                globalObj == ServerObjectType.Folder));

            clientCopyFlag = copyFlag;
            isClientWantAsynchronous = !isWantAsynchronousZero;
            isDestinationInRequestExist = isDestinationExist;
        }

        /// <summary>
        ///  The second part of action RopCopyPropertiesMethod
        /// </summary>
        /// <param name="isPropertiesDeleted">If CopyFlags is set to Move, Source object will be deleted after copy.</param>
        /// <param name="isChangedInDB">Indicates whether the change is submitted to DB</param>
        /// <param name="isOverwriteDestination"> If CopyFlags is set to NoOverWrite, Destination should not be overwritten. </param>
        /// <param name="isReturnedRopProgress">
        /// If this ROP is performed asynchronously,
        /// RopProgress response returned instead of RopCopyProperties response
        /// </param>
        /// <param name="error"> If destination object is not exist, NullDestinationObject error will be returned </param>
        [Rule(Action = @"return RopCopyPropertiesMethod(_,_,_,
            out isPropertiesDeleted, out isChangedInDB,out isOverwriteDestination, out isReturnedRopProgress, out error)")]
        public static void RopCopyPropertiesMethodReturn(
            bool isPropertiesDeleted,
            bool isChangedInDB,
            bool isOverwriteDestination,
            bool isReturnedRopProgress,
            CPRPTErrorCode error)
        {
            Condition.IsTrue(
                error == CPRPTErrorCode.None ||
                error == CPRPTErrorCode.NullDestinationObject ||
                error == CPRPTErrorCode.NotSupported ||
                error == CPRPTErrorCode.InvalidParameter);

            if (error == CPRPTErrorCode.None || error == CPRPTErrorCode.NotSupported || error == CPRPTErrorCode.InvalidParameter)
            {
                Condition.IsTrue(isDestinationInRequestExist);
            }

            if (error == CPRPTErrorCode.NullDestinationObject)
            {
                Condition.IsTrue(!isDestinationInRequestExist && !isPropertiesDeleted && !isChangedInDB && !isOverwriteDestination);
            }

            if (error == CPRPTErrorCode.InvalidParameter || error == CPRPTErrorCode.NotSupported)
            {
                Condition.IsTrue(!isPropertiesDeleted && !isChangedInDB && !isOverwriteDestination);
            }

            // Test synchronous and asynchronous situation through "isClientWantAsynchronous", "isReturnedRopProgress" and "isWorkSynchronously"
            if (!isClientWantAsynchronous)
            {
                Condition.IsTrue(!isReturnedRopProgress);
                isWorkSynchronously = true;
                ModelHelper.CaptureRequirement(
                    15204,
                    "[In RopCopyProperties ROP Request Buffer] WantAsynchronous: If this field is set to zero, then the server performs the ROP[RopCopyProperties] synchronously.");

                // To test the CopyFlags
                // Distinguish behavior on folder object, then common behaviors on message and attachment objects
                if (isDestinationInRequestExist)
                {
                    if (clientCopyFlag == CopyFlags.None)
                    {
                        isCopyPropertiesSuccess = true;
                        Condition.IsTrue(error == CPRPTErrorCode.None && isOverwriteDestination && !isPropertiesDeleted);

                        if (globalObj == ServerObjectType.Folder)
                        {
                            Condition.IsTrue(isChangedInDB);

                            ModelHelper.CaptureRequirement(
                                503,
                                @"[In Processing RopCopyProperties] In the case of Folder objects, the changes on the source
                        and destination MUST be immediately persisted.");
                        }
                        else
                        {
                            Condition.IsFalse(isChangedInDB);
                        }
                    }
                    else if (clientCopyFlag == CopyFlags.NoOverWrite)
                    {
                        isCopyPropertiesSuccess = true;
                        Condition.IsTrue(error == CPRPTErrorCode.None && !isOverwriteDestination && !isChangedInDB && !isPropertiesDeleted);

                        ModelHelper.CaptureRequirement(
                            146,
                            @"[In RopCopyProperties ROP] The RopCopyProperties ROP ([MS-OXCROPS] section 2.2.8.11) copies or moves one or more properties from one object to another. ");

                        ModelHelper.CaptureRequirement(
                            264,
                            @"[In RopCopyProperties ROP Request Buffer] CopyFlags: NoOverwrite: If this bit[bit 0x02] is set, properties that already have a value on the destination object will not be overwritten;");

                        ModelHelper.CaptureRequirement(
                            879,
                            @"[In Processing RopCopyProperties] If the NoOverwrite flag is set in the CopyFlags field,
                        the server MUST NOT overwrite any properties that already have a value on the destination object.");

                        ModelHelper.CaptureRequirement(
                                15901,
                                @"[In RopCopyProperties ROP Request Buffer] CopyFlags: otherwise [If this bit is not set to 0x01], properties are copied.");
                    }
                    else if (clientCopyFlag == CopyFlags.MoveAndNoOverWrite)
                    {
                        if (requirementContainer[86701])
                        {
                            if (globalObj == ServerObjectType.Folder)
                            {
                                Condition.IsTrue(error == CPRPTErrorCode.NotSupported);
                            }
                            else
                            {
                                isCopyPropertiesSuccess = true;
                                Condition.IsTrue(!isPropertiesDeleted && !isOverwriteDestination && error == CPRPTErrorCode.None && !isChangedInDB);
                                ModelHelper.CaptureRequirement(
                                    86701,
                                    @"Implementation does support combination of the Move bit and the NoOverwrite bit. <2> Section 2.2.10.1: Exchange 2003 and Exchange 2007 support the combination of the Move bit and the NoOverwrite bit in the CopyFlags field.");
                            }
                        }

                        if (requirementContainer[86502])
                        {
                            Condition.IsTrue(error == CPRPTErrorCode.InvalidParameter);
                            ModelHelper.CaptureRequirement(
                                86502,
                                @"[In RopCopyProperties ROP Request Buffer] CopyFlags: Implementation does not support the combination of these bit (bit 0x01 and bit 0x02) in CopyFlags.(Microsoft Exchange Server 2010 and above follow this behavior)");
                        }
                    }
                    else if (clientCopyFlag == CopyFlags.Move)
                    {
                        if (globalObj != ServerObjectType.Folder)
                        {
                            Condition.IsTrue(error == CPRPTErrorCode.None);
                            Condition.IsTrue(isOverwriteDestination);
                            ModelHelper.CaptureRequirement(
                                265,
                                @"[In RopCopyProperties ROP Request Buffer] CopyFlags: NoOverwrite:  otherwise(If this bit is not set to 0x02), they [properties that already have a value on the destination object] are overwritten.");

                            if (requirementContainer[86704])
                            {
                                Condition.IsTrue(isPropertiesDeleted);
                                ModelHelper.CaptureRequirement(
                                    86704,
                                    @"Implementation does remove the properties from the source object, [if the Move flag is set in the CopyFlags field of the ROP request buffer.](<6> Section 3.2.5.7: Exchange 2003 and Exchange 2007 remove the properties from the source object.)");

                                ModelHelper.CaptureRequirement(
                                    159,
                                    @"[In RopCopyProperties ROP Request Buffer] CopyFlags: If this bit[bit 0x01] is set, properties are moved");
                            }

                            if (requirementContainer[50101])
                            {
                                Condition.IsTrue(!isPropertiesDeleted);
                                ModelHelper.CaptureRequirement(
                                    50101,
                                    @"[In Processing RopCopyProperties] Implementation doesn't delete the copied properties from the source object, if the move flag is set in the CopyFlags field of the ROP request buffer.(Microsoft Exchange Server 2010 and above follow this behavior.)");
                            }
                        }
                    }
                    else if (clientCopyFlag == CopyFlags.Other)
                    {
                        if (requirementContainer[86705])
                        {
                            if (globalObj != ServerObjectType.Folder)
                            {
                                Condition.IsTrue(error == CPRPTErrorCode.None);

                                ModelHelper.CaptureRequirement(
                                    86705,
                                    @"Implementation does ignore invalid bits and doesn't return the InvalidParameter error code. <7> Section 3.2.5.7: Exchange 2003 and Exchange 2007 ignore invalid bits and do not return the InvalidParameter error code (0x80070057).");
                            }
                            else
                            {
                                Condition.IsTrue(error == CPRPTErrorCode.NotSupported);
                            }
                        }

                        if (requirementContainer[88001])
                        {
                            Condition.IsTrue(error == CPRPTErrorCode.InvalidParameter);
                            ModelHelper.CaptureRequirement(
                                88001,
                                @"[In Processing RopCopyProperties] Implementation does return an InvalidParameter error (0x80070057) ([MS-OXCDATA] section 2.4), if any other bits are set in the CopyFlags field.(Microsoft Exchange Server 2010 and above follow this behavior)");
                        }
                    }
                }
            }
            else
            {
                if (!isReturnedRopProgress)
                {
                    isWorkSynchronously = true;
                    ModelHelper.CaptureRequirement(
                        153,
                        @"[In RopCopyProperties ROP Request Buffer] WantAsynchronous: If this field is set to nonzero, the ROP is processed either synchronously or asynchronously.");

                    ModelHelper.CaptureRequirement(
                        504,
                        @"[In Processing RopCopyProperties] If the client requests asynchronous execution, then the server can execute this ROP asynchronously.");

                    ModelHelper.CaptureRequirement(
                        50402,
                        @"[In Processing RopCopyProperties] During asynchronous processing, the server can indicate that the operation is still being processed by returning a RopProgress ROP response ([MS-OXCROPS] section 2.2.8.13), or it can indicate that the operation has already completed by returning a RopCopyProperties ROP response.");
                }
            }
        }

        /// <summary>
        /// First part for action [RopCopyTo] operation
        /// </summary>
        /// <param name="copyFlag">Specifies the CopyFlags in the call request</param>
        /// <param name="isWantAsynchronousZero">Indicates whether WantAsynchronous parameter in call request is zero</param>
        /// <param name="isWantSubObjectsZero">Indicates whether WantSubObjects parameter in call request is zero</param>
        /// <param name="isDestinationExist">Indicates whether destination object is exist for [RopCopyTo]</param>
        [Rule(Action = @"call RopCopyToMethod(copyFlag, isWantAsynchronousZero, isWantSubObjectsZero, isDestinationExist, out _,out _,out _, out _, out _, out _)")]
        public static void RopCopyToMethodCall(CopyFlags copyFlag, bool isWantAsynchronousZero, bool isWantSubObjectsZero, bool isDestinationExist)
        {
            Condition.IsTrue(isInitialized && isFirstObjectGot && isSecondObjectGot &&
                (globalObj == ServerObjectType.Message ||
                globalObj == ServerObjectType.Attachment ||
                globalObj == ServerObjectType.Folder));

            ModelHelper.CaptureRequirement(
                821,
                @"[In Processing RopCopyTo] The source object and destination object need to be of the same type, and MUST be a Message object, Folder object, or Attachment object.");

            clientCopyFlag = copyFlag;
            isClientWantAsynchronous = !isWantAsynchronousZero;
            isClientWantSubObjects = !isWantSubObjectsZero;
            isDestinationInRequestExist = isDestinationExist;
        }

        /// <summary>
        /// Second part of action RopCopyToMethod
        /// </summary>
        /// <param name="isPropertiesDeleted">If CopyFlags is set to Move, source object will be deleted after copy</param>
        /// <param name="isSubObjectCopied">Indicates whether sub-object properties is also be copied</param>
        /// <param name="isOverwriteDestination">If CopyFlags is set to NoOverWrite, destination should not be overwritten.</param>
        /// <param name="isReturnedRopProgress">If this ROP is performed asynchronously, RopProgress response returned instead of RopCopyProperties response</param>
        /// <param name="isChangedInDB">Indicates whether destination is changed in database.</param>
        /// <param name="error">
        /// If destination object is not exist,
        /// NullDestinationObject error will be returned
        /// </param>
        [Rule(Action = @"return RopCopyToMethod(_, _, _, _,
            out isPropertiesDeleted, out isSubObjectCopied, out isOverwriteDestination, out isReturnedRopProgress, out isChangedInDB, out error)")]
        public static void RopCopyToMethodReturn(
            bool isPropertiesDeleted,
            bool isSubObjectCopied,
            bool isOverwriteDestination,
            bool isReturnedRopProgress,
            bool isChangedInDB,
            CPRPTErrorCode error)
        {
            Condition.IfThen(clientCopyFlag == CopyFlags.Move, error == CPRPTErrorCode.None);

            Condition.IsTrue(error == CPRPTErrorCode.None
                || error == CPRPTErrorCode.NullDestinationObject
                || error == CPRPTErrorCode.InvalidParameter);

            if (error == CPRPTErrorCode.None)
            {
                isCopyToSuccess = true;
            }

            // Test synchronous and asynchronous situation through "isClientWantAsynchronous", "isReturnedRopProgress" and "isWorkSynchronously"
            if (!isClientWantAsynchronous)
            {
                Condition.IsTrue(!isReturnedRopProgress);
                isWorkSynchronously = true;
                ModelHelper.CaptureRequirement(
                    177,
                    "[In RopCopyTo ROP Request Buffer] WantAsynchronous: If this field is set to zero, this ROP [RopCopyTo ROP] is processed synchronously. ");

                if (isDestinationInRequestExist)
                {
                    // To test the CopyFlags
                    if (clientCopyFlag == CopyFlags.NoOverWrite)
                    {
                        Condition.IsTrue(!isOverwriteDestination && !isPropertiesDeleted);
                        ModelHelper.CaptureRequirement(
                              5070504,
                              @"[In Processing RopCopyTo] If the NoOverwrite flag is set in the CopyFlags field, the server MUST NOT overwrite any properties that already have a value on the destination object. ");

                        ModelHelper.CaptureRequirement(
                              18501,
                              @"[In RopCopyTo ROP Request Buffer] CopyFlags: Move: otherwise(If this bit is not set to 0x01), properties are copied.");

                        ModelHelper.CaptureRequirement(
                            624,
                            @"[In RopCopyTo ROP Request Buffer] CopyFlags: NoOverwrite: If this bit[bit 0x02] is set, properties that already have a value on the destination object will not be overwritten;");

                        if (globalObj == ServerObjectType.Folder)
                        {
                            Condition.IsTrue(isChangedInDB);
                            ModelHelper.CaptureRequirement(
                                5070509,
                                @"[In Processing RopCopyTo] In the case of Folder objects, the changes on the source and destination MUST be immediately persisted.");
                        }

                        if (globalObj == ServerObjectType.Message)
                        {
                            Condition.IsFalse(isChangedInDB);

                            if (isClientWantSubObjects)
                            {
                                Condition.IsTrue(isSubObjectCopied);
                                ModelHelper.CaptureRequirement(
                                     181,
                                     "[In RopCopyTo ROP Request Buffer] WantSubObjects: If WantSubObjects is nonzero then sub-objects MUST also be copied.");
                            }
                            else
                            {
                                Condition.IsTrue(!isSubObjectCopied);
                                ModelHelper.CaptureRequirement(
                                    182,
                                    "[In RopCopyTo ROP Request Buffer] WantSubObjects: Otherwise[If WantSubObjects is zero] they[sub-objects] are  not[copied].");
                            }
                        }
                    }
                    else if (clientCopyFlag == CopyFlags.MoveAndNoOverWrite)
                    {
                        if (requirementContainer[86702])
                        {
                            Condition.IsTrue(error == CPRPTErrorCode.None);
                            ModelHelper.CaptureRequirement(
                               86702,
                               @"Implementation does support combination. (<3> Section 2.2.11.1: Exchange 2003 and Exchange 2007 support the combination of the Move bit and the NoOverwrite bit in the CopyFlags field.)");
                        }

                        if (requirementContainer[18402])
                        {
                            Condition.IsTrue(error == CPRPTErrorCode.InvalidParameter);
                            ModelHelper.CaptureRequirement(
                                18402,
                                @"[In RopCopyTo ROP Request Buffer] CopyFlags: Implementation doesn't support the combination of these bit(bit0x01 and bit 0x02) in CopyFlags.(Microsoft Exchange Server 2010 and above follow this behavior)");
                        }
                    }
                    else if (clientCopyFlag == CopyFlags.Move)
                    {
                        Condition.IsTrue(error == CPRPTErrorCode.None);
                        Condition.IsTrue(isOverwriteDestination);

                        ModelHelper.CaptureRequirement(
                            625,
                            @"[In RopCopyTo ROP Request Buffer] CopyFlags: otherwise[If this bit is not set to 0x02], they [properties that already have a value on the destination object] are overwritten.");

                        if (requirementContainer[86707])
                        {
                            ModelHelper.CaptureRequirement(
                                86707,
                                @"Implementation does delete the property. <8> Section 3.2.5.8: Exchange 2003, Exchange 2007, and Exchange 2010 delete the properties from the source object.)");
                        }
                    }
                    else if (clientCopyFlag == CopyFlags.Other)
                    {
                        if (requirementContainer[86708])
                        {
                            Condition.IsTrue(error == CPRPTErrorCode.None);
                            ModelHelper.CaptureRequirement(
                                86708,
                                @"Implementation does not return the InvalidParameter error code (0x80070057). <9> Section 3.2.5.8: Exchange 2003 and Exchange 2007 ignore invalid bits and do not return the InvalidParameter error code (0x80070057).");
                        }

                        if (requirementContainer[5070506])
                        {
                            Condition.IsTrue(error == CPRPTErrorCode.InvalidParameter);

                            ModelHelper.CaptureRequirement(
                                5070506,
                                @"[In Processing RopCopyTo] Implementation does return an InvalidParameter error (0x80070057). if any other bits are set in the CopyFlags field.(Microsoft Exchange Server 2010 and above follow this behavior.)");
                        }
                    }
                }
            }
            else
            {
                if (!isReturnedRopProgress)
                {
                    isWorkSynchronously = true;
                    ModelHelper.CaptureRequirement(
                        178,
                        @"[In RopCopyTo ROP Request Buffer] WantAsynchronou:] If this field is set to nonzero, this ROP [RopCopyTo ROP] is processed either synchronously or asynchronously.");

                    ModelHelper.CaptureRequirement(
                        50801,
                        @"[In Processing RopCopyTo] If the client requests asynchronous processing, the server can process this ROP asynchronously.");

                    ModelHelper.CaptureRequirement(
                       50802,
                       @"[In Processing RopCopyTo] During asynchronous processing, the server can indicate that the operation is still being processed by returning a RopProgress ROP response ([MS-OXCROPS] section 2.2.8.13), or it can indicate that the operation has already completed by returning a RopCopyTo ROP response. ");
                }
            }
        }

        /// <summary>
        /// Action for generating error codes of [RopCopyTo] operation,
        /// all the requirements for error code table need to get value from PTFConfig, 
        /// so this action is only to generate corresponding test cases, and leave the requirements to adapter to verify
        /// </summary>
        /// <param name="condition">The condition to generate the corresponding error codes</param>
        [Rule(Action = "RopCopyToMethodForErrorCodeTable(condition)")]
        public static void RopCopyToMethodForErrorCodeTable(CopyToCondition condition)
        {
            Condition.IsTrue(
                isInitialized &&
                isFirstObjectGot &&
                isSecondObjectGot &&
                (globalObj == ServerObjectType.Message ||
                globalObj == ServerObjectType.Attachment ||
                globalObj == ServerObjectType.Folder));
        }

        /// <summary>
        /// Action for the Folder-To-Folder mode of [RopCopyTo] operation.
        /// </summary>
        [Rule(Action = "RopCopyToForPublicFolder()")]
        public static void RopCopyToForPublicFolder()
        {
            Condition.IsTrue(isInitialized);
        }

        /// <summary>
        /// First part of action RopProgressMethod
        /// </summary>
        /// <param name="isOtherRopSent">Indicates whether has other ROPs been sent</param>
        /// <param name="isWantCancel">Indicates whether client want cancel</param>
        [Rule(Action = "call RopProgressMethod(isOtherRopSent, isWantCancel, out _, out _)")]
        public static void RopProgressMethodCall(bool isOtherRopSent, bool isWantCancel)
        {
            Condition.IsTrue(isInitialized && !isWorkSynchronously);

            isClientSendOtherRop = isOtherRopSent;
            isClientWantCancel = isWantCancel;
        }

        /// <summary>
        /// Second part of action RopProgressMethod
        /// </summary>
        /// <param name="isOriginalOpsResponse">
        /// If original asynchronous ROPs are done or canceled,
        /// Response should be original ROPs response.
        /// Otherwise, it should be RopProgress response.
        /// </param>
        /// <param name="isOtherRopResponse"> Indicates the other ROP's response is returned
        ///  If the client sends a ROP other than RopProgress to the server with the same logon 
        ///  before the asynchronous operation is complete the server 
        ///  MUST abort the asynchronous operation and respond to the new ROP.
        /// </param>
        [Rule(Action = "return RopProgressMethod(_, _, out isOriginalOpsResponse, out isOtherRopResponse)")]
        public static void RopProgressMethodReturn(bool isOriginalOpsResponse, bool isOtherRopResponse)
        {
            if (!isClientSendOtherRop)
            {
                Condition.IsTrue(!isOtherRopResponse);
                if (isOriginalOpsResponse == true && isClientWantCancel == true)
                {
                    isWorkSynchronously = true;
                }

                if (isOriginalOpsResponse == true && isClientWantCancel == false)
                {
                    isWorkSynchronously = true;
                }
            }
            else if (isClientSendOtherRop)
            {
                Condition.IsTrue(isOtherRopResponse && !isOriginalOpsResponse);
            }
        }

        /// <summary>
        ///   Get common object properties in order to test their type
        /// </summary>
        /// <param name="commonProperty">The common object properties defined in section 2.2.1</param>
        [Rule(Action = "GetCommonObjectProperties(commonProperty)")]
        public static void GetCommonObjectProperties(CommonObjectProperty commonProperty)
        {
            Condition.IsTrue(isInitialized);
        }

        /// <summary>
        ///  Set common object properties in order to test whether each of them is read-only
        /// </summary>
        /// <param name="commonProperty">The  common object properties defined in section 2.2.1</param>
        /// <param name="error">When a property is specified as "read-only for the client", 
        /// the server MUST return an error and ignore any request to change the value of that property.
        /// </param>
        [Rule(Action = "SetCommonObjectProperties(commonProperty, out error)")]
        public static void SetCommonObjectProperties(CommonObjectProperty commonProperty, out CPRPTErrorCode error)
        {
            error = CPRPTErrorCode.None;
        }

        /// <summary>
        /// Get requirement enabled
        /// </summary>
        /// <param name="rsid">Requirement ID</param>
        /// <param name="enabled">True represents the requirement is enabled; false represents the requirement is disabled.</param>
        [Rule(Action = "CheckRequirementEnabled(rsid, out enabled)")]
        public static void CheckRequirementEnabled(int rsid, out bool enabled)
        {
            enabled = Choice.Some<bool>();
            requirementContainer.Add(rsid, enabled);

            if (requirementContainer.ContainsKey(86702) && requirementContainer.ContainsKey(18402))
            {
                Condition.IsTrue(requirementContainer[86702] != requirementContainer[18402]);
            }

            if (requirementContainer.ContainsKey(86708) && requirementContainer.ContainsKey(5070506))
            {
                Condition.IsTrue(requirementContainer[86708] != requirementContainer[5070506]);
            }

            if (requirementContainer.ContainsKey(86701) && requirementContainer.ContainsKey(86502))
            {
                Condition.IsTrue(requirementContainer[86701] != requirementContainer[86502]);
            }

            if (requirementContainer.ContainsKey(86704) && requirementContainer.ContainsKey(50101))
            {
                Condition.IsTrue(requirementContainer[86704] != requirementContainer[50101]);
            }

            if (requirementContainer.ContainsKey(86705) && requirementContainer.ContainsKey(88001))
            {
                Condition.IsTrue(requirementContainer[86705] != requirementContainer[88001]);
            }
        }

        /// <summary>
        /// This method is used to check whether MAPIHTTP transport is supported by SUT.
        /// </summary>
        /// <param name="isSupported">The transport is supported or not.</param>
        [Rule(Action = "CheckMAPIHTTPTransportSupported(out isSupported)")]
        public static void CheckMAPIHTTPTransportSupported(out bool isSupported)
        {
            isSupported = Choice.Some<bool>();
        }
    }
}