namespace Microsoft.Protocols.TestSuites.MS_OXCPRPT
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-OXCPRPTAdapter class.
    /// </summary>
    public interface IMS_OXCPRPTAdapter : IAdapter
    {
        /// <summary>
        /// This method is used to initialize the test environment for private mailbox.
        /// </summary>
        void InitializeMailBox();

        /// <summary>
        /// This method is used to initialize the test environment for public folders.
        /// </summary>
        void InitializePublicFolder();

        /// <summary>
        /// This method is used to get object for different object types.
        /// </summary>
        /// <param name="objType">Specifies the object type.</param>
        /// <param name="objToOperate">Specifies which objects to operate.</param>
        void GetObject(ServerObjectType objType, ObjectToOperate objToOperate);

        /// <summary>
        /// The method is used to query an object for all the named properties. 
        /// </summary>
        /// <param name="queryFlags">Specifies QueryFlags parameter in request.</param>
        /// <param name="hasGuid">Indicates whether HasGuid is zero, 
        /// If the HasGUID field is non-zero then the PropertyGUID field MUST be included in the request. 
        /// If no PropertyGUID is specified, then properties from any GUID MUST be returned in the results.</param>
        /// <param name="isKind0x01Returned">True if the named properties of the response with the Kind field 
        /// ([MS-OXCDATA] section 2.6.1) set to 0x1 was returned.</param>
        /// <param name="isKind0x00Returned">True if the named properties of the response with the Kind field 
        /// ([MS-OXCDATA] section 2.6.1) set to 0x0 was returned.</param>
        /// <param name="isNamedPropertyGuidReturned">True if the named properties  with a GUID field ([MS-OXCDATA] 
        /// section 2.6.1) value that does not match the value of the PropertyGUID field was returned.</param>
        void RopQueryNamedPropertiesMethod(
            QueryFlags queryFlags,
            bool hasGuid,
            out bool isKind0x01Returned,
            out bool isKind0x00Returned,
            out bool isNamedPropertyGuidReturned);

        /// <summary>
        /// This method is used to query for and return all of the property tags and values of properties that have been set. 
        /// </summary>
        /// <param name="isPropertySizeLimitZero">Indicates whether PropertySizeLimit parameter is zero.</param>
        /// <param name="isPropertyLargerThanLimit">Indicates whether request properties larger than limit:
        /// When PropertySizeLimit is non-zero, it indicates whether request properties larger than PropertySizeLimit,
        /// When PropertySizeLimit is zero, it indicates whether request properties larger than size of response.</param>
        /// <param name="isUnicode">Indicates whether the requested property is encoded in Unicode format in response buffer.</param>
        /// <param name="isValueContainsNotEnoughMemory">Indicates whether returned value contains NotEnoughMemory error when the request properties are too large.</param>
        void RopGetPropertiesAllMethod(bool isPropertySizeLimitZero, bool isPropertyLargerThanLimit, bool isUnicode, out bool isValueContainsNotEnoughMemory);

        /// <summary>
        /// This method is used to query for and return all of the property tags for properties that have been set on an object. 
        /// </summary>
        void RopGetPropertiesListMethod();

        /// <summary>
        /// This method is used to query for and return the values of properties specified in the PropertyTags field. 
        /// </summary>
        /// <param name="isTestOrder">Indicates whether to test returned PropertyNames order.</param>
        /// <param name="isPropertySizeLimitZero">Indicates whether PropertySizeLimit parameter is zero.</param>
        /// <param name="isPropertyLargerThanLimit">Indicates whether request properties larger than limit
        /// When PropertySizeLimit is non-zero, it indicates whether request properties larger than PropertySizeLimit
        /// When PropertySizeLimit is zero, it indicates whether request properties larger than size of response.</param>
        /// <param name="isValueContainsNotEnoughMemory">Indicates whether returned value contains NotEnoughMemory error when the request properties are too large.</param>
        void RopGetPropertiesSpecificMethod(bool isTestOrder, bool isPropertySizeLimitZero, bool isPropertyLargerThanLimit, out bool isValueContainsNotEnoughMemory);

        /// <summary>
        ///  This method is used to query for and return the values of properties specified in the PropertyTags field, which is related with unicode format. 
        /// </summary>
        /// <param name="isUnicode">Indicates whether the requested property is encoded in Unicode format in response buffer</param>
        void RopGetPropertiesSpecificForWantUnicode(bool isUnicode);

        /// <summary>
        /// This method is used to query for and return the values of properties specified in the PropertyTags field, which is related with tagged properties. 
        /// </summary>
        void RopGetPropertiesSpecificForTaggedProperties();

        /// <summary>
        /// This method is used to map abstract, client-defined named properties to concrete 16-bit property IDs. 
        /// </summary>
        /// <param name="isTestOrder">Indicates whether to test returned PropertyNames order.</param>
        /// <param name="isCreateFlagSet">Indicates whether the "Create" Flags in request parameter is set.</param>
        /// <param name="isPropertyNameExisting">Indicates whether PropertyName is existing in object mapping.</param>
        /// <param name="specialPropertyName">Specifies PropertyName of request parameter.</param>
        /// <param name="isCreatedEntryReturned">If Create Flags is set: If set, indicates that the server MUST create new
        /// entries for any name parameters that are not found in the existing mapping set, and return existing entries for any
        /// name parameters that are found in the existing mapping set.</param>
        /// <param name="error">Specifies the ErrorCode when server reached limit.</param>
        void RopGetPropertyIdsFromNamesMethod(
            bool isTestOrder,
            bool isCreateFlagSet,
            bool isPropertyNameExisting,
            SpecificPropertyName specialPropertyName,
            out bool isCreatedEntryReturned,
            out CPRPTErrorCode error);

        /// <summary>
        /// This method is used to map concrete property IDs to abstract, client-defined named properties. 
        /// </summary>
        /// <param name="propertyIdType">Specifies different PropertyId type</param>
        void RopGetNamesFromPropertyIdsMethod(PropertyIdType propertyIdType);

        /// <summary>
        /// This method is used to set property values for an object without invoking replication. 
        /// </summary>
        /// <param name="isSameWithSetProperties">
        /// Indicates whether result is same as RopSetProperties.
        /// </param>
        void RopSetPropertiesNoReplicateMethod(out bool isSameWithSetProperties);

        /// <summary>
        /// This method is used to delete property values from an object without invoking replication. 
        /// </summary>
        /// <param name="isSameWithDeleteProperties">Indicates whether result is same as RopDeleteProperties.</param>
        /// <param name="isChangedInDB">Indicates the database is changed or not</param>
        void RopDeletePropertiesNoReplicateMethod(out bool isSameWithDeleteProperties, out bool isChangedInDB);

        /// <summary>
        /// This method is used to update the specified properties on an object. 
        /// </summary>
        /// <param name="isModifiedValueReturned">Indicates whether the modified value of a property can be returned use a same handle.</param>
        /// <param name="isChangedInDB">Indicates whether the modified value is submit to DB.
        /// For Message and Attachment object, it require another ROP for submit DB.
        /// For Logon and Folder object, it DO NOT need any other ROPs for submit.</param>
        void RopSetPropertiesMethod(out bool isModifiedValueReturned, out bool isChangedInDB);

        /// <summary>
        /// This method is used to remove the specified properties from an object. 
        /// </summary>
        /// <param name="isNoValidValueReturnedForDeletedProperties">
        /// If the server returns success, it MUST NOT have a valid value to return to a client that asks for the value of this property.
        /// </param>
        /// <param name="isChangedInDB">
        /// Indicates whether the modified value is submit to DB For Message and Attachment object, it require another 
        /// ROP for submit DB. For Logon and Folder object, it DO NOT need any other ROPs for submit.</param>
        void RopDeletePropertiesMethod(out bool isNoValidValueReturnedForDeletedProperties, out bool isChangedInDB);

        /// <summary>
        /// This method is used to commit the changes made to a message. 
        /// </summary>
        /// <param name="isChangedInDB">Indicates whether changes of Message object submit to database 
        /// when [RopSetProperties] or [RopDeleteProperties].</param>
        void RopSaveChangesMessageMethod(out bool isChangedInDB);

        /// <summary>
        /// This method is used to commit the changes made to an attachment. 
        /// </summary>
        /// <param name="isChangedInDB">Indicates whether changes of Message object submit to database 
        /// when [RopSetProperties] or [RopDeleteProperties].</param>
        void RopSaveChangesAttachmentMethod(out bool isChangedInDB);

        /// <summary>
        /// This method is used to open a property as a Stream object, enabling the client to perform various streaming operations on the property. 
        /// </summary>
        /// <param name="obj">Specifies which object will be operated.</param>
        /// <param name="openFlag">Specifies OpenModeFlags for [RopOpenStream].</param>
        /// <param name="isPropertyTagExist">Indicates whether request property exist.</param>
        /// <param name="isStreamSizeEqualToStream">Indicates whether StreamSize in response is 
        /// the same with the current number of BYTES in the stream.</param>
        /// <param name="error">If the property tag does not exist for the object and "Create" 
        /// is not specified in OpenModeFlags, NotFound error should be returned.</param>
        void RopOpenStreamMethod(ObjectToOperate obj, OpenModeFlags openFlag, bool isPropertyTagExist, out bool isStreamSizeEqualToStream, out CPRPTErrorCode error);

        /// <summary>
        /// This method is used to open a different type of properties as a Stream object, enabling the client to perform various streaming operations on the property. 
        /// </summary>
        /// <param name="obj">Specifies which object will be operated.</param>
        /// <param name="propertyType">Specifies which type of property will be operated.</param>
        /// <param name="error">Returned error code.</param>
        void RopOpenStreamWithDifferentPropertyType(ObjectToOperate obj, PropertyTypeName propertyType, out CPRPTErrorCode error);

        /// <summary>
        /// This method is used to release all resources associated with a Server object. 
        /// The client uses RopRelease ([MS-OXCROPS] section 2.2.14.3) after it is done with the Stream object.
        /// </summary>
        /// <param name="obj">Specifies which object will be operated.</param>
        void RopReleaseMethodNoVerify(ObjectToOperate obj);

        /// <summary>
        /// This method is used to read the stream of bytes from a Stream object. 
        /// </summary>
        /// <param name="isReadingFailed">Indicates whether reading stream get failure. E.g. object handle is not stream.</param>
        void RopReadStreamMethod(bool isReadingFailed);

        /// <summary>
        /// This method is used to read the stream of limited size bytes from a Stream object.
        /// </summary>
        /// <param name="byteCount">Indicates the size to be read.</param>
        /// <param name="maxByteCount">If byteCount is 0xBABE, use MaximumByteCount to determine the size to be read.</param>
        void RopReadStreamWithLimitedSize(ushort byteCount, uint maxByteCount);

        /// <summary>
        /// This method is used to set the seek pointer to a new location, which is relative to the beginning of the stream, the end of the stream, or the location of the current seek pointer. 
        /// </summary>
        /// <param name="condition">Specifies particular scenario of RopSeekStream.</param>
        /// <param name="isStreamExtended">Indicates whether a stream object is extended and zero filled to the new seek location.</param>
        /// <param name="error">Returned error code.</param>
        void RopSeekStreamMethod(SeekStreamCondition condition, out bool isStreamExtended, out CPRPTErrorCode error);

        /// <summary>
        /// This method is used to lock a specified range of bytes in a Stream object. 
        /// </summary>
        /// <param name="preState">Specifies the pre-state before call [RopLockRegionStream]</param>
        /// <param name="error">Return error
        /// 1. If there are previous locks that are not expired, the server MUST return an AccessDenied error.
        /// 2. If a session with an expired lock calls any ROP for this Stream object that would encounter the locked region, 
        /// the server MUST return a NetworkError.</param>
        void RopLockRegionStreamMethod(PreStateBeforeLock preState, out CPRPTErrorCode error);

        /// <summary>
        /// This method is used to unlock a specified range of bytes in a Stream object. 
        /// </summary>
        /// <param name="isPreviousLockExists">Indicates whether a previous lock exists and not owned by this session.
        /// If there are previous locks that are not owned by the session calling the ROP, the server MUST leave them unmodified.</param>
        void RopUnlockRegionStreamMethod(bool isPreviousLockExists);

        /// <summary>
        /// This method is used to write the stream of bytes into a Stream object. 
        /// </summary>
        /// <param name="openFlag">Specifies the OpenModeFlags of the stream.</param>
        /// <param name="isExceedMax">Indicates whether the write will exceed the maximum stream size.</param>
        /// <param name="error"> Specifies the ErrorCode when WriteStream failed:
        /// STG_E_ACCESSDENIED 0x80030005 Write access is denied.
        /// When stream is opened with ReadOnly flag.</param>
        void RopWriteStreamMethod(OpenModeFlags openFlag, bool isExceedMax, out CPRPTErrorCode error);

        /// <summary>
        /// This method is used to write the stream of bytes into a Stream object. 
        /// </summary>
        /// <param name="openFlag">Specifies the OpenModeFlags of the stream.</param>
        /// <param name="isExceedMax">Indicates whether the write will exceed the maximum stream size.</param>
        /// <param name="error"> Specifies the ErrorCode when WriteStreamExtended failed:STG_E_ACCESSDENIED
        /// 0x80030005 Write access is denied.When stream is opened with ReadOnly flag.</param>
        void RopWriteStreamExtendedMethod(OpenModeFlags openFlag, bool isExceedMax, out CPRPTErrorCode error);

        /// <summary>
        /// This method is used to ensure that any changes made to a Stream object are persisted in storage for a Folder object. 
        /// </summary>
        /// <param name="openFlag">Indicates the OpenModeFlags when stream is opened.</param>
        /// <param name="isPropertyValueChanged">Indicates whether property value is changed.</param>
        void RopCommitStreamMethod(OpenModeFlags openFlag, out bool isPropertyValueChanged);

        /// <summary>
        /// Method for ROP lease operation.
        /// The client uses RopRelease ([MS-OXCROPS] section 2.2.14.3) after it is done with the Stream object. 
        /// </summary>
        /// <param name="obj">Specifies which object will be operated.</param>
        /// <param name="isPropertyValueChanged">
        /// For Folder Object, this ROP should not change the value in stream after RopWriteStream.
        /// For non-Folder Object, this ROP should change the value.
        /// </param>
        void RopReleaseMethod(ObjectToOperate obj, out bool isPropertyValueChanged);

        /// <summary>
        /// This method is used to copy a specified number of bytes from the current seek pointer in the source stream to the current seek pointer in the destination stream. 
        /// </summary>
        /// <param name="isDestinationExist">Specified the whether the destination existed.</param>
        /// <param name="isReadWriteSuccess">When call success:The server MUST read the number of BYTES
        /// requested from the source Stream object, and write those bytes into the destination Stream object.</param>
        /// <param name="error">If Destination object does not exist, expect DestinationNullObject error.</param>
        void RopCopyToStreamMethod(bool isDestinationExist, out bool isReadWriteSuccess, out CPRPTErrorCode error);

        /// <summary>
        /// This method is used to write bytes to a stream and commits the stream. 
        /// </summary>
        /// <param name="error">This ROP MUST NOT be used on Stream objects opened on 
        /// properties on Folder objects which means it should be failed against Folder object.</param>
        void RopWriteAndCommitStreamMethod(out CPRPTErrorCode error);

        /// <summary>
        /// This method is used to create a new Stream object that is a clone of another Stream object. 
        /// </summary>
        void RopCloneStreamMethod();

        /// <summary>
        /// This method is used to retrieve the size of the stream. 
        /// </summary>
        void RopGetStreamSizeMethod();

        /// <summary>
        /// This method is used to set the size of a stream. 
        /// </summary>
        /// <param name="isSizeIncreased"> Indicates the new size is increased or decreased.</param>
        /// <param name="isExtendedValueZero">
        /// If the size of the stream is increased, then value of the extended stream MUST be zero.
        /// </param>
        /// <param name="isLost">
        /// If the size of the stream is decreased, the information that extends past the end of the new size is lost.
        /// </param>
        /// <param name="isIncrease">If the size of the stream is increased, set this value to true</param>
        void RopSetStreamSizeMethod(bool isSizeIncreased, out bool isExtendedValueZero, out bool isLost, out bool isIncrease);

        /// <summary>
        /// This method is used to copy or move one or more properties from one object to another. 
        /// </summary>
        /// <param name="copyFlag">Specifies the CopyFlags in the call request.</param>
        /// <param name="isWantAsynchronousZero">Indicates whether WantAsynchronous parameter in call request is zero.</param>
        /// <param name="isDestinationExist">Indicates whether destination object is exist for [RopCopyProperties].</param>
        /// <param name="isPropertiesDeleted">If CopyFlags is set to Move,Source object will be deleted after copy to.</param>
        /// <param name="isChangedInDB">Indicates whether the change is submit to DB.</param>
        /// <param name="isOverwriteDestination">If CopyFlags is set to NoOverWrite,Destination should not be overwritten.</param>
        /// <param name="isReturnedRopProgress">If this ROP is performed Asynchronously,RopProgress response returned.</param>
        /// <param name="error">If destination object is not exist,NullDestinationObject error will be returned.</param>
        void RopCopyPropertiesMethod(
            CopyFlags copyFlag,
            bool isWantAsynchronousZero,
            bool isDestinationExist,
            out bool isPropertiesDeleted,
            out bool isChangedInDB,
            out bool isOverwriteDestination,
            out bool isReturnedRopProgress,
            out CPRPTErrorCode error);

        /// <summary>
        /// This method is used to copy or move all but a specified few properties from a source object to a destination object. 
        /// </summary>
        /// <param name="copyFlag">Specifies the CopyFlags in the call request.</param>
        /// <param name="isWantAsynchronousZero">Indicates whether WantAsynchronous parameter in call request is zero.</param>
        /// <param name="isWantSubObjectsZero">Indicates whether WantSubObjects parameter in call request is zero.</param>
        /// <param name="isDestinationExist">Indicates whether destination object is exist for [RopCopyTo]</param>
        /// <param name="isPropertiesDeleted">If CopyFlags is set to Move,Source object will be deleted after copy to.</param>
        /// <param name="isSubObjectCopied">Indicates whether sub-object properties is also be copied.</param>
        /// <param name="isOverwriteDestination">If CopyFlags is set to NoOverWrite, destination should not be overwritten.</param>
        /// <param name="isReturnedRopProgress">If this ROP is performed Asynchronously, RopProgress response will be returned.</param>
        /// <param name="isChangedInDB">Indicates whether destination is changed in database.</param>
        /// <param name="error">If destination object does not exist, NullDestinationObject error will be returned.</param>
        void RopCopyToMethod(
            CopyFlags copyFlag,
            bool isWantAsynchronousZero,
            bool isWantSubObjectsZero,
            bool isDestinationExist,
            out bool isPropertiesDeleted,
            out bool isSubObjectCopied,
            out bool isOverwriteDestination,
            out bool isReturnedRopProgress,
            out bool isChangedInDB,
            out CPRPTErrorCode error);

        /// <summary>
        /// This method is used to copy or move properties from a source object to a destination object with error code returned. 
        /// </summary>
        /// <param name="condition">Specifies a special scenario of RopCopyTo.</param>
        void RopCopyToMethodForErrorCodeTable(CopyToCondition condition);

        /// <summary>
        /// This method is used to copy or move properties from a source object to a destination object on public folder.
        /// </summary>
        void RopCopyToForPublicFolder();

        /// <summary>
        /// This method is used to report the progress status of an asynchronous operation. 
        /// </summary>
        /// <param name="isOtherRopSent">Indicates whether other ROP is sent.</param>
        /// <param name="isWantCancel">Indicates whether WantCancel parameter is set to non-zero,any 
        /// non-zero value means client want cancel the original operation.</param>
        /// <param name="isOriginalOpsResponse">If original asynchronous ROPs are done or canceled,
        /// response should be original ROPs response. Otherwise, it should be RopProgress response.</param>
        /// <param name="isOtherRopResponse">Indicates the other ROP's response is returned. If the client sends a 
        /// ROP other than RopProgress to the server with the same logon before the asynchronous operation is 
        /// complete the server MUST abort the asynchronous operation and respond to the new ROP.</param>
        void RopProgressMethod(bool isOtherRopSent, bool isWantCancel, out bool isOriginalOpsResponse, out bool isOtherRopResponse);

        /// <summary>
        ///   Get common object properties in order to test their type.
        /// </summary>
        /// <param name="commonProperty">The nine Common Object Properties defined in section 2.2.1.</param>
        void GetCommonObjectProperties(CommonObjectProperty commonProperty);

        /// <summary>
        ///  Set common object properties in order to test whether each of them is read-only.
        /// </summary>
        /// <param name="commonProperty">The nine Common Object Properties defined in section 2.2.1.</param>
        /// <param name="error">When a property is specified as "read-only for the client", the server MUST
        /// return an error and ignore any request to change the value of that property.</param>
        void SetCommonObjectProperties(CommonObjectProperty commonProperty, out CPRPTErrorCode error);

        /// <summary>
        /// Checks if the requirement is enabled in the SHOULDMAY ptfconfig files.
        /// </summary>
        /// <param name="rsid">Requirement ID</param>
        /// <param name="enabled">True represents the requirement is enabled; false represents the requirement is disabled.</param>
        void CheckRequirementEnabled(int rsid, out bool enabled);
        
        /// <summary>
        /// This method is used to check whether MAPIHTTP transport is supported by SUT.
        /// </summary>
        /// <param name="isSupported">The transport is supported or not.</param>
        void CheckMAPIHTTPTransportSupported(out bool isSupported);
    }
}