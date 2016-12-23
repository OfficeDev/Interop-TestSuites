namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    /// <summary>
    /// The enumeration of types of error codes in a sub response. 
    /// </summary>
    public enum ErrorCodeType
    {
        #region GenericErrorCodeTypes

        /// <summary>
        /// Indicating that the cell storage service sub request succeeded for the given URL for the file.
        /// </summary>
        Success,

        /// <summary>
        /// Indicating an error when any an incompatible version number is specified as part of the RequestVersion element of the cell storage service. 
        /// </summary>
        IncompatibleVersion,

        /// <summary>
        /// Indicating an error when the URL for the file being specified as part of the Request element is invalid. 
        /// </summary>
        InvalidUrl,

        /// <summary>
        /// Indicating an error when the targeted URL for the file specified as part of the Request element does not exists or file creation failed on the protocol server. 
        /// </summary>
        FileNotExistsOrCannotBeCreated,

        /// <summary>
        /// Indicating an error when the targeted URL for the file specified as part of the Request element does not have correct authorization.
        /// </summary>
        FileUnauthorizedAccess,

        /// <summary>
        /// Indicating an error when the file path is not found.
        /// </summary>
        PathNotFound,

        /// <summary>
        /// Indicating an error when one or more SubRequest elements for a targeted URL for the file was unable to be parsed. 
        /// </summary>
        InvalidSubRequest,

        /// <summary>
        /// Indicating an unknown error when processing any SubRequest element for a targeted URL for the file.
        /// </summary>
        SubRequestFail,

        /// <summary>
        /// Indicating an error when the targeted URL to the file’s file type is blocked on the protocol server.
        /// </summary>
        BlockedFileType,

        /// <summary>
        /// Indicating an error when the targeted URL for the file is not yet checked out by the current client before sending a lock request on the file. 
        /// </summary>
        DocumentCheckoutRequired,

        /// <summary>
        /// Indicating an error when any of the cell storage service sub requests for the targeted URL for the file contains invalid input parameters. 
        /// </summary>
        InvalidArgument,

        /// <summary>
        /// Indicating an error when the targeted cell storage service sub request is a valid sub request, but the server does not support that sub request. 
        /// </summary>
        RequestNotSupported,

        /// <summary>
        /// Indicating an error when the associated protocol server site URL is not found.
        /// </summary>
        InvalidWebUrl,

        /// <summary>
        /// Indicating an error when the web service is turned off during processing of the cell storage service request. 
        /// </summary>
        WebServiceTurnedOff,

        /// <summary>
        /// Indicating an error when the file that is correctly stored on server is modified by another user before the current user finished writing to the underlying file provider. 
        /// </summary>
        ColdStoreConcurrencyViolation,

        /// <summary>
        /// Indicating any undefined error that occurs during processing of the cell storage service request. 
        /// </summary>
        Unknown,

        /// <summary>
        /// Indicating the version number does not match a version of the file.
        /// </summary>
        VersionNotFound,
        #endregion

        #region CellRequestErrorCodeTypes

        /// <summary>
        /// Indicating an error when processing a cell sub request for the given URL for the file. 
        /// </summary>
        CellRequestFail,

        /// <summary>
        /// Indicating an error when a matching Etag is not found. 
        /// </summary>
        CellRequestEtagNotMatching,

        /// <summary>
        /// Indicating an error when the requested file is in an Information Rights Management (IRM) protected document, that is only supported through WebDav (Web Distributed Authoring and Versioning Protocol).
        /// </summary>
        IRMDocLibarysOnlySupportWebDAV,

        #endregion

        #region DependencyCheckRelatedErrorCodeTypes

        /// <summary>
        /// Indicating an error when the sub request on which this specific sub request is dependent on has not been executed yet.
        /// </summary>
        DependentRequestNotExecuted,

        /// <summary>
        /// Indicating an error when the sub request on which this specific sub request is dependent on has failed and this sub request MUST only execute on success of the sub request on which this specific sub request depends. 
        /// </summary>
        DependentOnlyOnSuccessRequestFailed,

        /// <summary>
        /// Indicating an error when the sub request on which this specific sub request is dependent on has succeeded and this sub request MUST only execute on failure of the original sub request on which this specific sub request depends.
        /// </summary>
        DependentOnlyOnFailRequestSucceeded,

        /// <summary>
        /// Indicating an error when the sub request on which this specific sub request is dependent on is supported and this sub request MUST only execute if the original sub request is not supported.
        /// </summary>
        DependentOnlyOnNotSupportedRequestGetSupported,

        /// <summary>
        /// Indicating an error when an invalid sub request dependency type is specified.
        /// </summary>
        InvalidRequestDependencyType,

        #endregion

        #region LockAndCoauthRelatedErrorCodeTypes

        /// <summary>
        /// Indicating key/value pairs exceed the quota.
        /// </summary>
        EditorMetadataQuotaReached,

        /// <summary>
        /// Indicating the key exceeds the length limit.
        /// </summary>
        EditorMetadataStringExceedsLengthLimit,

        /// <summary>
        /// Indicating an editor client id is not found.
        /// </summary>
        EditorClientIdNotFound,

        /// <summary>
        /// Indicating an undefined error that occurs during processing of lock operations requested as part of a cell storage service sub request.
        /// </summary>
        LockRequestFail,

        /// <summary>
        /// Indicating an error when there is an already existing exclusive lock on the targeted URL for the file or a schema lock on the file with a different schema lock identifier. 
        /// </summary>
        FileAlreadyLockedOnServer,

        /// <summary>
        /// Indicating an error when no exclusive lock or shared lock exists on a file and a release of the lock or a conversion of the lock is requested as part of a cell storage service request. 
        /// </summary>
        FileNotLockedOnServer,

        /// <summary>
        /// Indicating an error when no shared lock exists on a file, because co-authoring of file is disabled on the server. 
        /// </summary>
        FileNotLockedOnServerAsCoauthDisabled,

        /// <summary>
        /// Indicating an error when a protocol server fails to process a lock conversion request sent as part of a cell storage service request, because co-authoring of file is disabled on the server. 
        /// </summary>
        LockNotConvertedAsCoauthDisabled,

        /// <summary>
        /// Indicating an error when the file is checked out by another client that is preventing the file from being locked by the current client. 
        /// </summary>
        FileAlreadyCheckedOutOnServer,

        /// <summary>
        /// Indicating an error when convert to shared lock fails because the file is checked out by the current client. 
        /// </summary>
        ConvertToSchemaFailedFileCheckedOutByCurrentUser,

        /// <summary>
        /// Indicating an error when a save of the file co-authoring tracker, maintained by the protocol server fails after some other client edited the file co-authoring tracker before the save is done by the current client. 
        /// </summary>
        CoauthRefblobConcurrencyViolation,

        /// <summary>
        /// Indicating an error when all of the following conditions are true: 
        /// A co-authoring sub request of type, "Convert to exclusive lock" or schema lock sub request of type, "Convert to exclusive lock" is requested on a file 
        /// There is more than one client in the current co-authoring session for that file
        /// ReleaseLockOnConversionToExclusiveFailure attribute specified as part of the sub request is set to false.
        /// </summary>
        MultipleClientsInCoauthSession,

        /// <summary>
        /// Indicating an error when one of the following conditions is true when a co-authoring sub request or schema lock sub request is sent:
        /// No co-authoring session exists for the file.
        /// The current client does not exist in the co-authoring session for the file
        /// The current client exists in the co-authoring session, but protocol server is unable to remove it from the co-authoring session for the file. 
        /// </summary>
        InvalidCoauthSession,

        /// <summary>
        /// Indicating an error when the number of users that co-author a file has reached the threshold limit. The threshold limit specifies the maximum number of users allowed to co-author a file at any instant of time. 
        /// </summary>
        NumberOfCoauthorsReachedMax,

        /// <summary>
        /// Indicating an error when a co-authoring sub request or schema lock sub request of type, "Convert to exclusive lock" is sent by the client with the ReleaseLockOnConversionToExclusiveFailure attribute set to true and there is more than one client editing the file.
        /// </summary>
        ExitCoauthSessionAsConvertToExclusiveFailed

        #endregion
    }
}