namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class provides common functions to the TestSuite project.
    /// </summary>
    public class SharedTestSuiteHelper
    {
        /// <summary>
        /// A string indicates reserved SchemaLockId.
        /// </summary>
        public const string ReservedSchemaLockID = "29358EC1-E813-4793-8E70-ED0344E7B73C";

        /// <summary>
        /// A int indicates default timeout.
        /// </summary>
        public const int DefaultTimeOut = 3600;

        /// <summary>
        /// A string indicates default ExclusiveLockID.
        /// </summary>
        public const string DefaultExclusiveLockID = "e8adfa03-1af6-4272-b657-90ca322fbc7f";

        /// <summary>
        /// A string indicates default clientId.
        /// </summary>
        public const string DefaultClientID = "7936e3b0-d116-4d67-ad18-89a4fcfcfbe8";

        /// <summary>
        /// Prevents a default instance of the SharedTestSuiteHelper class from being created
        /// </summary>
        private SharedTestSuiteHelper()
        { 
        }

        #region MS-FSSHTTPB helper function
        /// <summary>
        /// A method used to extract FSSHTTPB subResponse from CellSubResponse element.
        /// </summary>
        /// <param name="cellSubResponse">A parameter represents SubResponse element returned by the protocol server.</param>
        /// <param name="site">A parameter represents the instance of ITestSite.</param>
        /// <returns>A return value represents FSSHTTPB subResponse.</returns>
        public static FsshttpbResponse ExtractFsshttpbResponse(CellSubResponseType cellSubResponse, ITestSite site)
        {
            site.Assert.AreEqual<int>(
                1,
                cellSubResponse.SubResponseData.Text.Length,
                "CellSubResponse should contain MS-FSSHTTPB embedded information.");

            string subResponseBase64 = cellSubResponse.SubResponseData.Text[0];
            byte[] subResponseBinary = Convert.FromBase64String(subResponseBase64);
            return FsshttpbResponse.DeserializeResponseFromByteArray(subResponseBinary, 0);
        }

        /// <summary>
        /// A method used to create a CellRequest object and initialize it.
        /// </summary>
        /// <returns>A return value represents the CellRequest object.</returns>
        public static FsshttpbCellRequest CreateFsshttpbCellRequest()
        {
            FsshttpbCellRequest cellRequest = new FsshttpbCellRequest();

            // MUST be great or equal to OxFA12994 
            cellRequest.Version = 0xFA12994;

            // MUST be 12 
            cellRequest.ProtocolVersion = 12;

            // MUST be 11 
            cellRequest.MinimumVersion = 11;

            // MUST be 0x9B069439F329CF9C 
            cellRequest.Signature = 0x9B069439F329CF9C;

            // Set the user agent GUID. 
            cellRequest.GUID = FsshttpbCellRequest.UserAgentGuid;

            // Set the value which MUST be 1. 
            cellRequest.RequestHashingSchema = new Compact64bitInt(1u);
            return cellRequest;
        }

        /// <summary>
        /// A method used to create a CellSubRequest object for QueryChanges and initialize it.
        /// </summary>
        /// <param name="subRequestId">A parameter represents the subRequest identifier.</param>
        /// <returns>A return value represents the CellRequest object for QueryChanges.</returns>
        public static CellSubRequestType CreateCellSubRequestEmbeddedQueryChanges(ulong subRequestId)
        {
            FsshttpbCellRequest cellRequest = CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = BuildFsshttpbQueryChangesSubRequest(subRequestId);
            cellRequest.AddSubRequest(queryChange, null);

            CellSubRequestType cellSubRequest = CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            return cellSubRequest;
        }

        /// <summary>
        /// A method used to create a AllocateExtendedGUIDRange subRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestToken">A parameter represents the subRequest token.</param>
        /// <param name="subRequestId">>A parameter represents the subRequest id.</param>
        /// <param name="requestIdCount">A parameter represents the number of extended GUIDs to allocate.</param>
        /// <returns>A return value represents the Allocate extended GUID range subRequest object.</returns>
        public static CellSubRequestType CreateCellSubRequestEmbeddedAllocateExtendedGuidRange(int subRequestToken, ulong subRequestId, Compact64bitInt requestIdCount)
        {
            FsshttpbCellRequest cellRequest = CreateFsshttpbCellRequest();
            AllocateExtendedGuidRangeCellSubRequest allocateExtendedGuidRange = new AllocateExtendedGuidRangeCellSubRequest(requestIdCount, subRequestId);
            cellRequest.AddSubRequest(allocateExtendedGuidRange, null);
            CellSubRequestType cellSubRequest = CreateCellSubRequest((ulong)subRequestToken, cellRequest.ToBase64());
            return cellSubRequest;
        }

        /// <summary>
        /// A method used to create PutChanges subRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestId">A parameter represents SubRequest id </param>
        /// <param name="fileContent">A parameter represents the local changes needed to be submitted to the protocol server.</param>
        /// <returns>A return value represents the CellSubRequest object for PutChanges.</returns>
        public static CellSubRequestType CreateCellSubRequestEmbeddedPutChanges(ulong subRequestId, byte[] fileContent)
        {
            FsshttpbCellRequest cellRequest = CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(fileContent, out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(subRequestId, storageIndexExGuid);
            cellRequest.AddSubRequest(putChange, dataElements);

            CellSubRequestType cellSubRequest = CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            return cellSubRequest;
        }

        /// <summary>
        /// A method used to create a CellSubRequest object for QueryAccess and initialize it.
        /// </summary>
        /// <param name="subRequestId">A parameter represents the subRequest identifier.</param>
        /// <returns>A return value represents the CellSubRequest object for QueryAccess.</returns>
        public static CellSubRequestType CreateCellSubRequestEmbeddedQueryAccess(ulong subRequestId)
        {
            FsshttpbCellRequest cellRequest = CreateFsshttpbCellRequest();
            QueryAccessCellSubRequest queryAccess = new QueryAccessCellSubRequest(subRequestId);
            cellRequest.AddSubRequest(queryAccess, null);

            CellSubRequestType cellSubRequest = CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            return cellSubRequest;
        }

        /// <summary>
        /// A method used to create a CellSubRequest object for QueryEditorsTable and initialize it.
        /// </summary>
        /// <param name="subRequestId">A parameter represents the subRequest identifier.</param>
        /// <returns>A return value represents the CellSubRequest object for QueryEditorsTable.</returns>
        public static CellSubRequestType CreateCellSubRequestEmbeddedQueryEditorsTable(ulong subRequestId)
        {
            FsshttpbCellRequest cellRequest = CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = BuildFsshttpbQueryChangesSubRequest(subRequestId);
            cellRequest.AddSubRequest(queryChange, null);

            CellSubRequestType cellSubRequest = CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            // For query editor table, the partition id MUST be 7808F4DD-2385-49d6-B7CE-37ACA5E43602.
            cellSubRequest.SubRequestData.PartitionID = "7808F4DD-2385-49d6-B7CE-37ACA5E43602";
            return cellSubRequest;
        }

        /// <summary>
        /// A method used to create a CellSubRequest object and initialize it.
        /// </summary>
        /// <param name="requestToken">A parameter represents Request token.</param>
        /// <param name="base64Content">A parameter represents serialized subRequest.</param>
        /// <returns>A return value represents CellSubRequest object.</returns>
        public static CellSubRequestType CreateCellSubRequest(ulong requestToken, string base64Content)
        {
            return CreateCellSubRequest(requestToken, base64Content, Convert.FromBase64String(base64Content).Length);
        }

        /// <summary>
        /// A method used to create a CellSubRequest object and initialize it.
        /// </summary>
        /// <param name="requestToken">A parameter represents Request token.</param>
        /// <param name="base64Content">A parameter represents serialized subRequest.</param>
        /// <param name="binaryDataSize">A parameter represents the number of bytes of data in the SubRequestData element of a cell sub-request.</param>
        /// <returns>A return value represents CellSubRequest object.</returns>
        public static CellSubRequestType CreateCellSubRequest(ulong requestToken, string base64Content, long binaryDataSize)
        {
            CellSubRequestType cellRequestType = new CellSubRequestType();
            cellRequestType.SubRequestToken = requestToken.ToString();
            CellSubRequestDataType subRequestData = new CellSubRequestDataType();
            subRequestData.BinaryDataSize = binaryDataSize;
            subRequestData.Text = new string[1];
            subRequestData.Text[0] = base64Content;

            cellRequestType.SubRequestData = subRequestData;

            return cellRequestType;
        }

        /// <summary>
        /// A method used to create a PutChanges CellSubRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestId">A parameter represents the subRequest identifier.</param>
        /// <param name="storageIndexExGuid">A parameter represents the Storage Index EXGUID.</param>
        /// <param name="expectStorageIndexExGuid">A parameter represents the Expected Storage Index EXGUID.</param>
        /// <param name="isImplyNullExpectedIfNoMapping">A parameter represents the behavior of checking the current storage index value prior to update.</param>
        /// <param name="isPartial">A parameter represents that this is a partial Put Changes, and not the full changes.</param>
        /// <param name="isPartialLast">A parameter represents if this is the last Put Changes in a partial set of changes.</param>
        /// <param name="isFavorCoherencyFailureOverNotFound">A parameter represents if to force a coherency check on the server.</param>
        /// <param name="isAbortRemainingPutChangesOnFailure">A parameter represents if to abort remaining Put Changes on failure.</param>
        /// <param name="isMultiRequestPutHint">A parameter represents to reduce the number of auto coalesces during multi-request put scenarios. If only one request for a Put Changes, this bit is zero.</param>
        /// <param name="isReturnCompleteKnowledgeIf">A parameter represents if to return the complete knowledge from the server.</param>
        /// <param name="isLastWriterWinsOnNextChange">A parameter represents if to allow the Put Changes to be subsequently overwritten on the next Put Changes, even if a client is not coherent with this change.</param>
        /// <returns>A return value represents PutChanges CellSubRequest object.</returns>
        public static PutChangesCellSubRequest BuildFsshttpbPutChangesSubRequestRequest(
                                ulong subRequestId,
                                ExGuid storageIndexExGuid,
                                ExGuid expectStorageIndexExGuid = null,
                                bool isImplyNullExpectedIfNoMapping = false,
                                bool isPartial = false,
                                bool isPartialLast = false,
                                bool isFavorCoherencyFailureOverNotFound = true,
                                bool isAbortRemainingPutChangesOnFailure = false,
                                //bool isMultiRequestPutHint = false,
                                bool isReturnCompleteKnowledgeIf = true,
                                bool isLastWriterWinsOnNextChange = false)
        {
            PutChangesCellSubRequest putChanges = new PutChangesCellSubRequest(subRequestId, storageIndexExGuid);
            putChanges.ExpectedStorageIndexExtendedGUID = expectStorageIndexExGuid;

            putChanges.ImplyNullExpectedIfNoMapping = Convert.ToInt32(isImplyNullExpectedIfNoMapping);
            putChanges.Partial = Convert.ToInt32(isPartial);
            putChanges.PartialLast = Convert.ToInt32(isPartialLast);
            putChanges.FavorCoherencyFailureOverNotFound = Convert.ToInt32(isFavorCoherencyFailureOverNotFound);
            putChanges.AbortRemainingPutChangesOnFailure = Convert.ToInt32(isAbortRemainingPutChangesOnFailure);
            //putChanges.MultiRequestPutHint = Convert.ToInt32(isMultiRequestPutHint);
            putChanges.ReturnCompleteKnowledgeIfPossible = Convert.ToInt32(isReturnCompleteKnowledgeIf);
            putChanges.LastWriterWinsOnNextChange = Convert.ToInt32(isLastWriterWinsOnNextChange);

            return putChanges;
        }

        /// <summary>
        /// A method used to create a QueryChanges CellSubRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestId">A parameter represents subRequest id.</param>
        /// <param name="reserved">A parameter that must be set to zero.</param>
        /// <param name="isAllowFragments">A parameter represents that if to allow fragments.</param>
        /// <param name="isExcludeObjectData">A parameter represents if to exclude object data.</param>
        /// <param name="isIncludeFilteredOutDataElementsInKnowledge">A parameter represents if to include the serial numbers of filtered out data elements in the response knowledge.</param>
        /// <param name="reserved1">A parameter represents a 4-bit reserved field that must be set to zero.</param>
        /// <param name="isStorageManifestIncluded">A parameter represents if to include the storage manifest.</param>
        /// <param name="isCellChangesIncluded">A parameter represents if to include the cell changes.</param>
        /// <param name="reserved2">A parameter represents a 6-bit reserved field that must be set to zero.</param>
        /// <param name="cellId">A parameter represents if the Query Changes are scoped to a specific cell. If the Cell ID is 0x0000, no scoping restriction is specified.</param>
        /// <param name="maxDataElements">A parameter represents the maximum data elements to return.</param>
        /// <param name="queryChangesFilterList">A parameter represents how the results of the query will be filtered before it is returned to the client.</param>
        /// <param name="knowledge">A parameter represents what the client knows about a state of a file.</param>
        /// <returns>A return value represents QueryChanges CellSubRequest object.</returns>
        public static QueryChangesCellSubRequest BuildFsshttpbQueryChangesSubRequest(
                                ulong subRequestId, 
                                int reserved = 0, 
                                bool isAllowFragments = false, 
                                bool isExcludeObjectData = false, 
                                bool isIncludeFilteredOutDataElementsInKnowledge = true, 
                                int reserved1 = 0, 
                                bool isStorageManifestIncluded = true, 
                                bool isCellChangesIncluded = true, 
                                int reserved2 = 0, 
                                CellID cellId = null, 
                                ulong? maxDataElements = null, 
                                List<Filter> queryChangesFilterList = null, 
                                Knowledge knowledge = null)
        {
            QueryChangesCellSubRequest queryChange = new QueryChangesCellSubRequest(subRequestId);

            queryChange.Reserved = reserved;
            queryChange.AllowFragments = Convert.ToInt32(isAllowFragments);
            queryChange.ExcludeObjectData = Convert.ToInt32(isExcludeObjectData);
            queryChange.IncludeFilteredOutDataElementsInKnowledge = Convert.ToInt32(isIncludeFilteredOutDataElementsInKnowledge);
            queryChange.Reserved1 = reserved1;

            queryChange.IncludeStorageManifest = Convert.ToInt32(isStorageManifestIncluded);
            queryChange.IncludeCellChanges = Convert.ToInt32(isCellChangesIncluded);
            queryChange.Reserved2 = reserved2;

            if (cellId == null)
            {
                cellId = new CellID(new ExGuid(0, Guid.Empty), new ExGuid(0, Guid.Empty));
            }

            queryChange.CellId = cellId;

            if (maxDataElements != null)
            {
                queryChange.MaxDataElements = new Compact64bitInt(maxDataElements.Value);
            }

            queryChange.QueryChangeFilters = queryChangesFilterList;
            queryChange.Knowledge = knowledge;

            return queryChange;
        }

        #endregion

        #region MS-FSSHTTP helper function

        /// <summary>
        /// A method used to create a WhoAmISubRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestToken">A parameter represents the subRequest token.</param>
        /// <returns>A return value represents the WhoAmISubRequest object.</returns>
        public static WhoAmISubRequestType CreateWhoAmISubRequest(uint subRequestToken)
        {
            WhoAmISubRequestType whoAmIRequestType = new WhoAmISubRequestType();
            whoAmIRequestType.SubRequestToken = subRequestToken.ToString();

            return whoAmIRequestType;
        }

        /// <summary>
        /// A method used to create a ServerTimeSubRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestToken">A parameter represents the subRequest token.</param>
        /// <returns>A return value represents the ServerTimeSubRequest object.</returns>
        public static ServerTimeSubRequestType CreateServerTimeSubRequest(uint subRequestToken)
        {
            ServerTimeSubRequestType serverTimeSubRequest = new ServerTimeSubRequestType();
            serverTimeSubRequest.SubRequestToken = subRequestToken.ToString();

            return serverTimeSubRequest;
        }

        /// <summary>
        /// A method used to create a GetDocMetaInfoSubRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestToken">A parameter represents the subRequest token.</param>
        /// <returns>A return value represents the GetDocMetaInfoSubRequest object.</returns>
        public static GetDocMetaInfoSubRequestType CreateGetDocMetaInfoSubRequest(uint subRequestToken)
        {
            GetDocMetaInfoSubRequestType getDocMetaInfoSubRequest = new GetDocMetaInfoSubRequestType();
            getDocMetaInfoSubRequest.SubRequestToken = subRequestToken.ToString();

            return getDocMetaInfoSubRequest;
        }

        /// <summary>
        /// A method used to create a GetVersionsSubRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestToken">A parameter represents the subRequest token.</param>
        /// <returns>A return value represents the GetVersionsSubRequest object.</returns>
        public static GetVersionsSubRequestType CreateGetVersionsSubRequest(uint subRequestToken)
        {
            GetVersionsSubRequestType getVersionsSubRequest = new GetVersionsSubRequestType();
            getVersionsSubRequest.SubRequestToken = subRequestToken.ToString();
            
            return getVersionsSubRequest;
        }

        /// <summary>
        /// A method used to create a AmIAloneSubRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestToken">A parameter represents the subRequest token.</param>
        /// <returns>A return value represents the AmIAloneSubRequest object.</returns>
        public static AmIAloneSubRequestType CreateAmIAloneSubRequest()
        {
            AmIAloneSubRequestType amIAloneSubRequest = new AmIAloneSubRequestType();
            amIAloneSubRequest.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
            amIAloneSubRequest.SubRequestData = new AmIAloneSubRequestDataType();
 
            return amIAloneSubRequest;
        }

        /// <summary>
        /// A method used to create a LockStatusSubRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestToken">A parameter represents the subRequest token.</param>
        /// <returns>A return value represents the LockStatusSubRequest object.</returns>
        public static LockStatusSubRequestType CreateLockStatusSubRequest()
        {
            LockStatusSubRequestType lockStatusSubRequest = new LockStatusSubRequestType();
            lockStatusSubRequest.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();

            return lockStatusSubRequest;
        }

        /// <summary>
        /// A method used to create a PropertiesSubRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestToken">A parameter represents the subRequest token.</param>
        /// <returns>A return value represents the PropertiesSubRequest object.</returns>
        public static PropertiesSubRequestType CreatePropertiesSubRequest(uint subRequestToken, PropertiesRequestTypes propertiesRequestTypes, PropertyIdType[] propertyIds, ITestSite site)
        {
            PropertiesSubRequestType propertiesSubRequest = new PropertiesSubRequestType();
            propertiesSubRequest.SubRequestToken = subRequestToken.ToString();
            propertiesSubRequest.SubRequestData = new PropertiesSubRequestDataType();
            propertiesSubRequest.SubRequestData.PropertiesSpecified = true;
            propertiesSubRequest.SubRequestData.Properties = propertiesRequestTypes;
            if (propertiesRequestTypes == PropertiesRequestTypes.PropertyGet)
            {
                site.Assert.IsTrue(propertyIds != null, "PropertyIds MUST be specified when the properties subrequest has a PropertiesSubRequestType attribute set to PropertyGet.");
                propertiesSubRequest.SubRequestData.PropertyIds = propertyIds;
            }
            return propertiesSubRequest;
        }

        /// <summary>
        /// A method used to create a VersioningSubRequestType object and initialize it.
        /// </summary>
        /// <param name="subRequestToken">A parameter represents the subRequest token.</param>
        /// <param name="versioningRequestType">Versioning request types</param>
        /// <param name="versionNumber">A FileVersionNumberType that serves to uniquely identify a version of a file on the server</param>
        /// <param name="site">A parameter represents the instance of ITestSite.</param>
        /// <returns>A return value represents the VersioningSubRequest object.</returns>
        public static VersioningSubRequestType CreateVersioningSubRequest(uint subRequestToken, VersioningRequestTypes versioningRequestType, string versionNumber, ITestSite site)
        {
            VersioningSubRequestType versioningSubRequest = new VersioningSubRequestType();
            versioningSubRequest.SubRequestToken = subRequestToken.ToString();
            versioningSubRequest.SubRequestData = new VersioningSubRequestDataType();
            versioningSubRequest.SubRequestData.VersioningRequestType = versioningRequestType;
            versioningSubRequest.SubRequestData.VersioningRequestTypeSpecified = true;

            if (versioningRequestType == VersioningRequestTypes.RestoreVersion)
            {
                site.Assert.IsTrue(!string.IsNullOrEmpty(versionNumber), "VersionNumber MUST be specified when the versioning subrequest has a VersioningSubRequestType attribute set to RestoreVersion.");
                versioningSubRequest.SubRequestData.Version = versionNumber;
            }

            return versioningSubRequest;
        }

        /// <summary>
        /// A method used to create a FileOperationSubRequestType object and initialize it.
        /// </summary>
        /// <param name="fileOperationRequestType">FileOperation request types</param>
        /// <param name="newName">A string that specifies a new name for the file on the server.</param>
        /// <param name="exclusiveLock">A string that serves as a unique identifier for the exclusive lock on the file at the time the file operation request is executed</param>
        /// <param name="site">A parameter represents the instance of ITestSite.</param>
        /// <returns>A return value represents the VersioningSubRequest object.</returns>
        public static FileOperationSubRequestType CreateFileOperationSubRequest(FileOperationRequestTypes fileOperationRequestType, string newName, string exclusiveLock, ITestSite site)
        {
            FileOperationSubRequestType fileOperationSubRequest = new FileOperationSubRequestType();
            fileOperationSubRequest.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
            fileOperationSubRequest.SubRequestData = new FileOperationSubRequestDataType();
            fileOperationSubRequest.SubRequestData.FileOperation = fileOperationRequestType;
            fileOperationSubRequest.SubRequestData.FileOperationSpecified = true;
            fileOperationSubRequest.SubRequestData.ExclusiveLockID = exclusiveLock;

            if (fileOperationRequestType == FileOperationRequestTypes.Rename)
            {
                site.Assert.IsTrue(!string.IsNullOrEmpty(newName), "FileOperation MUST be specified when the fileOperation subrequest has a FileOperationSubRequestType attribute set to Rename.");
                fileOperationSubRequest.SubRequestData.NewFileName = newName;
            }

            return fileOperationSubRequest;
        }

        #region Editors table helper function
        /// <summary>
        /// A method used to create a EditorsTable Sub-request object for JoinEditingSession and initialize it.
        /// </summary>
        /// <param name="clientId">A parameter represents the client identifier.</param>
        /// <param name="timeout">A parameter represents the time, in seconds, after which the editors table entry for that particular file will expire for that specific protocol client.</param>
        /// <returns>A return value represents the EditorsTable Sub-request object for JoinEditingSession.</returns>
        public static EditorsTableSubRequestType CreateEditorsTableSubRequestForJoinSession(string clientId, int timeout)
        {
            EditorsTableSubRequestType join = new EditorsTableSubRequestType();
            join.SubRequestData = new EditorsTableSubRequestDataType();
            join.SubRequestData.AsEditorSpecified = true;
            join.SubRequestData.AsEditor = true;
            join.SubRequestData.ClientID = clientId;
            join.SubRequestData.Timeout = timeout.ToString();
            join.SubRequestData.EditorsTableRequestTypeSpecified = true;
            join.SubRequestData.EditorsTableRequestType = EditorsTableRequestTypes.JoinEditingSession;
            join.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
            join.Type = SubRequestAttributeType.EditorsTable;

            return join;
        }

        /// <summary>
        /// A method used to create a EditorsTableSubRequest object for LeaveEditingSession and initialize it.
        /// </summary>
        /// <param name="clientId">A parameter represents the client identifier.</param>
        /// <returns>A return value represents the EditorsTable Sub-request object for LeaveEditingSession.</returns>
        public static EditorsTableSubRequestType CreateEditorsTableSubRequestForLeaveSession(string clientId)
        {
            EditorsTableSubRequestType leave = new EditorsTableSubRequestType();
            leave.SubRequestData = new EditorsTableSubRequestDataType();
            leave.SubRequestData.ClientID = clientId;
            leave.SubRequestData.EditorsTableRequestTypeSpecified = true;
            leave.SubRequestData.EditorsTableRequestType = EditorsTableRequestTypes.LeaveEditingSession;
            leave.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
            leave.Type = SubRequestAttributeType.EditorsTable;

            return leave;
        }

        /// <summary>
        /// A method used to create a EditorsTable Sub-request object for RefreshEditingSession and initialize it.
        /// </summary>
        /// <param name="clientId">A parameter represents the client identifier.</param>
        /// <param name="timeout">A parameter represents the time, in seconds, after which the editors table entry for that particular file will expire for that specific protocol client.</param>
        /// <returns>A return value represents the EditorsTable Sub-request object for RefreshEditingSession.</returns>
        public static EditorsTableSubRequestType CreateEditorsTableSubRequestForRefreshSession(string clientId, int timeout)
        {
            EditorsTableSubRequestType refresh = new EditorsTableSubRequestType();
            refresh.SubRequestData = new EditorsTableSubRequestDataType();
            refresh.SubRequestData.AsEditorSpecified = true;
            refresh.SubRequestData.AsEditor = true;
            refresh.SubRequestData.ClientID = clientId;
            refresh.SubRequestData.Timeout = timeout.ToString();
            refresh.SubRequestData.EditorsTableRequestTypeSpecified = true;
            refresh.SubRequestData.EditorsTableRequestType = EditorsTableRequestTypes.RefreshEditingSession;
            refresh.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
            refresh.Type = SubRequestAttributeType.EditorsTable;

            return refresh;
        }

        /// <summary>
        /// A method used to create a EditorsTable Sub-request object for RemoveEditorMetadata and initialize it.
        /// </summary>
        /// <param name="clientId">A parameter represents the client identifier.</param>
        /// <param name="key">A parameter represents a unique key in an arbitrary key/value pair of the client’s choice.</param>
        /// <returns>A return value represents the EditorsTable Sub-request object for RemoveEditorMetadata.</returns>
        public static EditorsTableSubRequestType CreateEditorsTableSubRequestForRemoveSessionMetadata(string clientId, string key)
        {
            EditorsTableSubRequestType remove = new EditorsTableSubRequestType();
            remove.SubRequestData = new EditorsTableSubRequestDataType();
            remove.SubRequestData.ClientID = clientId;
            remove.SubRequestData.Key = key;
            remove.SubRequestData.EditorsTableRequestTypeSpecified = true;
            remove.SubRequestData.EditorsTableRequestType = EditorsTableRequestTypes.RemoveEditorMetadata;
            remove.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
            remove.Type = SubRequestAttributeType.EditorsTable;

            return remove;
        }

        /// <summary>
        /// A method used to create a EditorsTable Sub-request object for UpdateEditorMetadata and initialize it.
        /// </summary>
        /// <param name="clientId">A parameter represents the client identifier.</param>
        /// <param name="key">A parameter represents the unique key in an arbitrary key/value pair of the client’s choice.</param>>
        /// <param name="content">A parameter represents the binary value that is associated with a key in an arbitrary key/value pair of the client’s choice.</param>
        /// <returns>A return value represents the EditorsTable Sub-request object for UpdateEditorMetadata.</returns>
        public static EditorsTableSubRequestType CreateEditorsTableSubRequestForUpdateSessionMetadata(string clientId, string key, byte[] content)
        {
            EditorsTableSubRequestType update = new EditorsTableSubRequestType();
            update.SubRequestData = new EditorsTableSubRequestDataType();
            update.SubRequestData.ClientID = clientId;
            update.SubRequestData.Key = key;
            update.SubRequestData.Text = new string[1];
            update.SubRequestData.Text[0] = Convert.ToBase64String(content);
            update.SubRequestData.EditorsTableRequestTypeSpecified = true;
            update.SubRequestData.EditorsTableRequestType = EditorsTableRequestTypes.UpdateEditorMetadata;
            update.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
            update.Type = SubRequestAttributeType.EditorsTable;

            return update;
        }
        #endregion

        /// <summary>
        /// A method used to create a SchemaLockSubRequest object and initialize it.
        /// </summary>
        /// <param name="allowFallbackToExclusive">A parameter represents whether a schema lock sub-request is allowed to fall back to an exclusive lock sub-request provided that shared locking on the file is not supported.</param>
        /// <param name="exclusiveLockId">A parameter represents a unique identifier for the exclusive lock on the file.</param>
        /// <returns>A return value represents a SchemaLock Sub-Request object.</returns>
        public static SchemaLockSubRequestType CreateSchemaLockSubRequestForGetLock(bool? allowFallbackToExclusive, string exclusiveLockId = DefaultExclusiveLockID)
        {
            SchemaLockSubRequestType schemaLockRequest = new SchemaLockSubRequestType();
            schemaLockRequest.Type = SubRequestAttributeType.SchemaLock;
            schemaLockRequest.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();

            schemaLockRequest.SubRequestData = new SchemaLockSubRequestDataType();
            schemaLockRequest.SubRequestData.ClientID = DefaultClientID;
            schemaLockRequest.SubRequestData.SchemaLockID = ReservedSchemaLockID;
            schemaLockRequest.SubRequestData.SchemaLockRequestType = SchemaLockRequestTypes.GetLock;
            schemaLockRequest.SubRequestData.SchemaLockRequestTypeSpecified = true;
            schemaLockRequest.SubRequestData.Timeout = DefaultTimeOut.ToString();

            if (allowFallbackToExclusive != null)
            {
                schemaLockRequest.SubRequestData.AllowFallbackToExclusive = allowFallbackToExclusive.Value;
                schemaLockRequest.SubRequestData.AllowFallbackToExclusiveSpecified = true;
                schemaLockRequest.SubRequestData.ExclusiveLockID = exclusiveLockId;
            }
            else
            {
                schemaLockRequest.SubRequestData.AllowFallbackToExclusiveSpecified = false;
            }

            return schemaLockRequest;
        }

        #region Coauthoring sub request helper function
        /// <summary>
        /// A method used to create the CoauthSubRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestToken">A parameter represents the sub-Request token.</param>
        /// <param name="coauthRequestType">A parameter represents the Coauthoring Request object.</param>
        /// <param name="clientId">A parameter represents the client identifier.</param>
        /// <param name="schemaLockId">A parameter represents a unique identifier for the schema lock on the file.</param>
        /// <param name="allowFallBack">A parameter represents whether a coauthoring sub-request is allowed to fall back to an exclusive lock sub-request provided shared locking on the file is not supported.</param>
        /// <param name="releaseLockOnConversionToExclusiveFailure">A parameter represents whether the server is allowed to remove the ClientId entry associated with the current client in the File coauthoring tracker.</param>
        /// <param name="exclusiveLockId">A parameter represents a unique identifier for the exclusive lock on the file when a coauthoring request of type "Convert to exclusive lock" is requested.</param>
        /// <param name="timeout">A parameter represents the time, in seconds, after which the shared lock for that particular file will expire for that specific protocol client.</param>
        /// <returns>A return value represents the Coauthoring Sub-Request object.</returns>
        public static CoauthSubRequestType CreateCoauthSubRequest(uint subRequestToken, CoauthRequestTypes coauthRequestType, string clientId, string schemaLockId, bool? allowFallBack, bool? releaseLockOnConversionToExclusiveFailure, string exclusiveLockId, int timeout)
        {
            CoauthSubRequestType coauthSubRequest = new CoauthSubRequestType();
            coauthSubRequest.SubRequestToken = subRequestToken.ToString();
            coauthSubRequest.SubRequestData = new CoauthSubRequestDataType();
            coauthSubRequest.SubRequestData.CoauthRequestType = coauthRequestType;
            coauthSubRequest.SubRequestData.CoauthRequestTypeSpecified = true;
            coauthSubRequest.SubRequestData.ClientID = clientId;
            coauthSubRequest.SubRequestData.SchemaLockID = schemaLockId;

            if (coauthRequestType == CoauthRequestTypes.JoinCoauthoring || coauthRequestType == CoauthRequestTypes.RefreshCoauthoring || coauthRequestType == CoauthRequestTypes.ConvertToExclusive)
            {
                coauthSubRequest.SubRequestData.Timeout = timeout.ToString();
            }

            if (coauthRequestType == CoauthRequestTypes.JoinCoauthoring || coauthRequestType == CoauthRequestTypes.RefreshCoauthoring)
            {
                if (allowFallBack != null)
                {
                    coauthSubRequest.SubRequestData.ExclusiveLockID = exclusiveLockId;
                    coauthSubRequest.SubRequestData.AllowFallbackToExclusive = allowFallBack.Value;
                    coauthSubRequest.SubRequestData.AllowFallbackToExclusiveSpecified = true;
                }
            }

            if (coauthRequestType == CoauthRequestTypes.ConvertToExclusive)
            {
                coauthSubRequest.SubRequestData.ReleaseLockOnConversionToExclusiveFailureSpecified = true;
                if (releaseLockOnConversionToExclusiveFailure != null)
                {
                    coauthSubRequest.SubRequestData.ReleaseLockOnConversionToExclusiveFailure = releaseLockOnConversionToExclusiveFailure.Value;
                }
                else
                {
                    coauthSubRequest.SubRequestData.ReleaseLockOnConversionToExclusiveFailure = false;
                }

                coauthSubRequest.SubRequestData.ExclusiveLockID = exclusiveLockId;
            }

            return coauthSubRequest;
        }

        /// <summary>
        /// A method used to create the CoauthSubRequest object for GetCoauthoringSession and initialize it.
        /// </summary>
        /// <param name="clientId">A parameter represents the client identifier.</param>
        /// <param name="schemaLockId">A parameter represents a unique identifier for the schema lock on the file.</param>
        /// <returns>A return value represents the Coauth SubRequest object for GetCoauthoringSession.</returns>
        public static CoauthSubRequestType CreateCoauthSubRequestForGetCoauthSessionStatus(string clientId, string schemaLockId)
        {
            return CreateCoauthSubRequest(SequenceNumberGenerator.GetCurrentToken(), CoauthRequestTypes.GetCoauthoringStatus, clientId, schemaLockId, null, null, null, DefaultTimeOut);
        }

        /// <summary>
        /// A method used to create the CoauthSubRequest object for JoinCoauthoringSession and initialize it.
        /// </summary>
        /// <param name="clientId">A parameter represents the client identifier.</param>
        /// <param name="schemaLockId">A parameter represents a unique identifier for the schema lock on the file.</param>
        /// <param name="allowFallBack">A parameter represents whether allow fall back to exclusive lock when join the coauth failed.</param>
        /// <param name="exclusiveLockId">A parameter represents a unique identifier for the exclusive lock on the file when a coauthoring request of type "Convert to exclusive lock" is requested.</param>
        /// <param name="timeout">A parameter represents the time, in seconds, after which the shared lock for that particular file will expire for that specific protocol client.</param>
        /// <returns>A return value represents the Coauth SubRequest object for JoinCoauthoringSession.</returns>
        public static CoauthSubRequestType CreateCoauthSubRequestForJoinCoauthSession(string clientId, string schemaLockId, bool? allowFallBack = null, string exclusiveLockId = DefaultExclusiveLockID, int timeout = DefaultTimeOut)
        {
            return CreateCoauthSubRequest(SequenceNumberGenerator.GetCurrentToken(), CoauthRequestTypes.JoinCoauthoring, clientId, schemaLockId, allowFallBack, null, exclusiveLockId, timeout);
        }

        /// <summary>
        /// A method used to create the CoauthSubRequest object for ConvertToExclusiveLock and initialize it.
        /// </summary>
        /// <param name="clientId">A parameter represents the client identifier.</param>
        /// <param name="schemaLockId">A parameter represents a unique identifier for the schema lock on the file.</param>
        /// <param name="exclusiveLockId">A parameter represents a unique identifier for the exclusive lock on the file when a coauthoring request of type "Convert to exclusive lock" is requested.</param>
        /// <param name="releaseLockOnConversionToExclusiveFailure">A parameter represents whether the server is allowed to remove the ClientId entry associated with the current client in the File coauthoring tracker.</param>
        /// <param name="timeout">A parameter represents the time, in seconds, after which the shared lock for that particular file will expire for that specific protocol client.</param>
        /// <returns>A return value represents the Coauth SubRequest object for ConvertToExclusiveLock.</returns>
        public static CoauthSubRequestType CreateCoauthSubRequestForConvertToExclusiveLock(string clientId, string schemaLockId, string exclusiveLockId, bool? releaseLockOnConversionToExclusiveFailure = null, int timeout = DefaultTimeOut)
        {
            return CreateCoauthSubRequest(SequenceNumberGenerator.GetCurrentToken(), CoauthRequestTypes.ConvertToExclusive, clientId, schemaLockId, null, releaseLockOnConversionToExclusiveFailure, exclusiveLockId, timeout);
        }

        /// <summary>
        /// A method used to create a CoauthSubRequest object for CheckLockAvailability and initialize it.
        /// </summary>
        /// <param name="clientId">A parameter represents a unique identifier for the client ID.</param>
        /// <param name="schemaLockId">A parameter represents a unique identifier for the schema lock on the file.</param>
        /// <returns>A return value represents the Coauth CellSubRequest object for CheckLockAvailability.</returns>
        public static CoauthSubRequestType CreateCoauthSubRequestForCheckLockAvailability(string clientId, string schemaLockId)
        {
            return CreateCoauthSubRequest(SequenceNumberGenerator.GetCurrentToken(), CoauthRequestTypes.CheckLockAvailability, clientId, schemaLockId, null, null, null, 0);
        }

        /// <summary>
        /// A method used to create a CoauthSubRequest object for ExitCoauthoringSession and initialize it.
        /// </summary>
        /// <param name="clientId">A parameter represents client identifier.</param>
        /// <param name="schemaLockId">A parameter represents a unique identifier for the schema lock on the file.</param>
        /// <returns>A return value represents the Coauth CellSubRequest object for ExitCoauthoringSession.</returns>
        public static CoauthSubRequestType CreateCoauthSubRequestForExitCoauthoringSession(string clientId, string schemaLockId)
        {
            return CreateCoauthSubRequest(SequenceNumberGenerator.GetCurrentToken(), CoauthRequestTypes.ExitCoauthoring, clientId, schemaLockId, null, null, null, 0);
        }

        /// <summary>
        /// A method used to create a CoauthSubRequest object for RefreshCoauthoringSession and initialize it.
        /// </summary>
        /// <param name="clientId">A parameter represents client identifier.</param>
        /// <param name="schemaLockId">A parameter represents a unique identifier for the schema lock on the file.</param>
        /// <param name="timeout">A parameter represents the time, in seconds, after which the shared lock for that particular file will expire for that specific protocol client.</param>
        /// <returns>A return value represents the Coauth CellSubRequest object for RefreshCoauthoringSession.</returns>
        public static CoauthSubRequestType CreateCoauthSubRequestForRefreshCoauthoringSession(string clientId, string schemaLockId, int timeout = DefaultTimeOut)
        {
            return CreateCoauthSubRequest(SequenceNumberGenerator.GetCurrentToken(), CoauthRequestTypes.RefreshCoauthoring, clientId, schemaLockId, null, null, null, timeout);
        }

        /// <summary>
        /// A method used to create a CoauthSubRequest object for MarkTransitionToComplete and initialize it.
        /// </summary>
        /// <param name="clientId">A parameter represents client identifier.</param>
        /// <param name="schemaLockId">A parameter represents a unique identifier for the schema lock on the file.</param>
        /// <returns>A return value represents the Coauth CellSubRequest object for MarkTransitionComplete.</returns>
        public static CoauthSubRequestType CreateCoauthSubRequestForMarkTransitionComplete(string clientId, string schemaLockId)
        {
            return CreateCoauthSubRequest(SequenceNumberGenerator.GetCurrentToken(), CoauthRequestTypes.MarkTransitionComplete, clientId, schemaLockId, null, null, null, 0);
        }
        #endregion

        /// <summary>
        /// A method used to make the current thread to sleep for the specified seconds.
        /// </summary>
        /// <param name="numSeconds">A parameter represents the number of seconds to sleep.</param>
        public static void Sleep(int numSeconds)
        {
            System.Threading.Thread.Sleep(1000 * numSeconds);
        }

        /// <summary>
        /// A method used to create an ExclusiveLockSubRequest object with specified exclusive lock operation types and initialize it.
        /// </summary>
        /// <param name="type">A parameter represents the exclusive lock operation types.<see cref="ExclusiveLockRequestTypes"/> </param>
        /// <returns>A return value represents ExclusiveLockSubRequest object.</returns>
        public static ExclusiveLockSubRequestType CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes type)
        {
            ExclusiveLockSubRequestType exclusiveLockSubRequest = new ExclusiveLockSubRequestType();
            exclusiveLockSubRequest.Type = SubRequestAttributeType.ExclusiveLock;
            exclusiveLockSubRequest.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();

            exclusiveLockSubRequest.SubRequestData = new ExclusiveLockSubRequestDataType();
            exclusiveLockSubRequest.SubRequestData.ExclusiveLockID = DefaultExclusiveLockID;
            exclusiveLockSubRequest.SubRequestData.ExclusiveLockRequestType = type;
            exclusiveLockSubRequest.SubRequestData.ExclusiveLockRequestTypeSpecified = true;

            if (type != ExclusiveLockRequestTypes.ReleaseLock && type != ExclusiveLockRequestTypes.CheckLockAvailability)
            {
                exclusiveLockSubRequest.SubRequestData.Timeout = DefaultTimeOut.ToString();
            }

            if (type == ExclusiveLockRequestTypes.ConvertToSchema || type == ExclusiveLockRequestTypes.ConvertToSchemaJoinCoauth)
            {
                exclusiveLockSubRequest.SubRequestData.SchemaLockID = ReservedSchemaLockID;
                exclusiveLockSubRequest.SubRequestData.ClientID = DefaultClientID;
            }

            return exclusiveLockSubRequest;
        }

        /// <summary>
        /// A method used to create a SchemaLockSubRequest object and initialize it.
        /// </summary>
        /// <param name="schemaLockRequesttype">A parameter represents the type of schema lock sub-request.</param>
        /// <param name="allowFallBack">A parameter represents whether a schema lock sub-request is allowed to fall back to an exclusive lock sub-request provided that shared locking on the file is not supported.</param>
        /// <param name="releaseLockOnConversionToExclusiveFailureAttr">A parameter represents whether the server is allowed to remove the ClientId entry associated with the current client in the File coauthoring tracker.</param>
        /// <returns>A return value represents the SchemaLockSubRequest object.</returns>
        public static SchemaLockSubRequestType CreateSchemaLockSubRequest(SchemaLockRequestTypes schemaLockRequesttype, bool? allowFallBack, bool? releaseLockOnConversionToExclusiveFailureAttr)
        {
            SchemaLockSubRequestType schemaLockRequest = new SchemaLockSubRequestType();
            schemaLockRequest.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
            schemaLockRequest.SubRequestData = new SchemaLockSubRequestDataType();
            schemaLockRequest.SubRequestData.SchemaLockRequestType = schemaLockRequesttype;
            schemaLockRequest.SubRequestData.SchemaLockRequestTypeSpecified = true;
            schemaLockRequest.SubRequestData.SchemaLockID = ReservedSchemaLockID;
            schemaLockRequest.SubRequestData.ClientID = DefaultClientID;
            schemaLockRequest.SubRequestData.Timeout = DefaultTimeOut.ToString();

            if (schemaLockRequesttype == SchemaLockRequestTypes.ConvertToExclusive)
            {
                schemaLockRequest.SubRequestData.ReleaseLockOnConversionToExclusiveFailureSpecified = true;
                if (releaseLockOnConversionToExclusiveFailureAttr != null)
                {
                    schemaLockRequest.SubRequestData.ReleaseLockOnConversionToExclusiveFailure = releaseLockOnConversionToExclusiveFailureAttr.Value;
                }
                else
                {
                    schemaLockRequest.SubRequestData.ReleaseLockOnConversionToExclusiveFailure = false;
                }

                schemaLockRequest.SubRequestData.ExclusiveLockID = DefaultExclusiveLockID;
            }

            if (schemaLockRequesttype == SchemaLockRequestTypes.GetLock && allowFallBack == true)
            {
                schemaLockRequest.SubRequestData.AllowFallbackToExclusiveSpecified = true;
                schemaLockRequest.SubRequestData.AllowFallbackToExclusive = true;
                schemaLockRequest.SubRequestData.ExclusiveLockID = DefaultExclusiveLockID;
            }

            return schemaLockRequest;
        }

        /// <summary>
        /// A method used to create a LockStatusSubRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestToken">A parameter represents the subRequest token.</param>
        /// <returns>A return value represents the LockStatusSubRequest object.</returns>
        public static LockStatusSubRequestType CreateLockStatusSubRequest(uint subRequestToken)
        {
            LockStatusSubRequestType lockStatusSubRequest = new LockStatusSubRequestType();
            lockStatusSubRequest.SubRequestToken = subRequestToken.ToString();

            return lockStatusSubRequest;
        }
        /// <summary>
        /// A method used to check if the specified subResponse element exists in the Response element.
        /// </summary>
        /// <param name="response">A parameter represents the response information.</param>
        /// <param name="site">A parameter represents the instance of ITestSite.</param>
        /// <param name="subResponseIndex">A parameter represents the index of subResponse element.</param>
        public static void CheckSubResponse(Response response, ITestSite site, int subResponseIndex = 0)
        {
            if (response == null)
            {
                site.Assert.Fail("In SharedTestSuiteHelper::CheckSubResponse, the parameter response is null.");
            }

            if (response.SubResponse.Length <= subResponseIndex)
            {
                site.Assert.Fail("Out of the range with the index value {0}", subResponseIndex);
            }
        }

        /// <summary>
        /// The method is used to expect the successful MS-FSSHTTPB response. 
        /// </summary>
        /// <param name="response">Specify the MS-FSSHTTPB response instance.</param>
        /// <param name="site">A parameter represents the instance of ITestSite.</param>
        /// <param name="subResponseIndex">A parameter represents the index of subResponse element.</param>
        public static void ExpectMsfsshttpbSubResponseSucceed(FsshttpbResponse response, ITestSite site, int subResponseIndex = 0)
        {
            if (response == null)
            {
                site.Assert.Fail("In SharedTestSuiteHelper::ExpectMsfsshttpbSubResponseSucceed, the parameter response is null.");
            }

            if (response.Status == true)
            {
                site.Assert.Fail("Expect the MS-FSSHTTPB response succeeds, actually it fails with the reason {0}", response.ResponseError.ErrorData.ErrorDetail);
            }

            if (response.CellSubResponses == null)
            {
                site.Assert.Fail("In SharedTestSuiteHelper::ExpectMsfsshttpbSubResponseSucceed, the parameter response.CellSubResponses is null.");
            }

            if (response.CellSubResponses.Count <= subResponseIndex)
            {
                site.Assert.Fail("In SharedTestSuiteHelper::ExpectMsfsshttpbSubResponseSucceed, out of the parameter response.CellSubResponses index with the value {0}.", subResponseIndex);
            }

            if (response.CellSubResponses[subResponseIndex].Status == true)
            {
                site.Assert.Fail("Expect the MS-FSSHTTPB sub response succeeds in the index {1}, actually it fails with the reason {0}", response.CellSubResponses[subResponseIndex].ResponseError.ErrorData.ErrorDetail, subResponseIndex);
            }
        }

        /// <summary>
        /// A method used to check if the specified Response element exists in the CellStorageResponse.
        /// </summary>
        /// <param name="response">A parameter represents the response information.</param>
        /// <param name="site">A parameter represents the instance of ITestSite.</param>
        /// <param name="responseIndex">A parameter represents the index of Response element.</param>
        public static void CheckCellStorageResponse(CellStorageResponse response, ITestSite site, int responseIndex = 0)
        {
            if (response == null)
            {
                site.Assert.Fail("In SharedTestSuiteHelper::CheckCellStorageResponse, the parameter response is null.");
            }

            if (response.ResponseCollection == null)
            {
                site.Assert.Fail("In SharedTestSuiteHelper::CheckCellStorageResponse, the parameter response.ResponseCollection is null.");
            }

            if (response.ResponseCollection.Response == null)
            {
                site.Assert.Fail("In SharedTestSuiteHelper::CheckCellStorageResponse, the parameter response.ResponseCollection.Response is null.");
            }

            if (response.ResponseCollection.Response.Length <= responseIndex)
            {
                site.Assert.Fail("In SharedTestSuiteHelper::CheckCellStorageResponse, out of the range of  parameter response.ResponseCollection.Response with the index value {0}.", responseIndex);
            }
        }

        /// <summary>
        /// A method used to extract subResponse elements from the CellStorageResponse element.
        /// </summary>
        /// <typeparam name="T">The type of subResponse.</typeparam>
        /// <param name="response">A parameter represents the response information.</param>
        /// <param name="responseIndex">A parameter represents the index of Response element.</param>
        /// <param name="subResponseIndex">A parameter represents the index of subResponse element.</param>
        /// <param name="site">A parameter represents an instance of ITestSite.</param>
        /// <returns>A return value represents the subResponse object.</returns>
        public static T ExtractSubResponse<T>(CellStorageResponse response, int responseIndex, int subResponseIndex, ITestSite site)
            where T : SubResponseType, new()
        {
            CheckCellStorageResponse(response, site, responseIndex);
            CheckSubResponse(response.ResponseCollection.Response[responseIndex], site, subResponseIndex);
            return FsshttpConverter.ConvertToSpecialSubResponse<T>(response.ResponseCollection.Response[responseIndex].SubResponse[subResponseIndex]);
        }

        /// <summary>
        /// A method used to convert a errorCode string to a ErrorCodeType value.
        /// </summary>
        /// <param name="errorCode">A parameter represents the error codes.</param>
        /// <param name="site">A parameter represents the instance of ITestSite.</param>
        /// <returns>A return value represents the errorCode value of type ErrorCodeType.</returns>
        public static ErrorCodeType ConvertToErrorCodeType(string errorCode, ITestSite site)
        {
            ErrorCodeType retValue;

            if (!Enum.TryParse<ErrorCodeType>(errorCode, true, out retValue))
            {
                site.Assert.Fail("Cannot convert the error code string {0} to the Enum type ErrorCodeType, {0} is not defined.", errorCode);
            }

            return retValue;
        }
        #endregion

        /// <summary>
        /// A method used to generate a file URL.
        /// </summary>
        /// <param name="site">A parameter represents the instance of ITestSite.</param>
        /// <returns>A return value represents the generated file URL.</returns>
        public static string GenerateNonExistFileUrl(ITestSite site)
        {
            string urlTemplate = Common.GetConfigurationPropertyValue("NormalFile", site);
            int index = urlTemplate.LastIndexOf('/');

            return urlTemplate.Remove(index) + "/" + Common.GenerateResourceName(site, "fileName") + ".txt";
        }

        /// <summary>
        /// This method is used to generate the random file content.
        /// </summary>
        /// <param name="site">Specify the ITestSite instance.</param>
        /// <returns>Return the file content in bytes.</returns>
        public static byte[] GenerateRandomFileContent(ITestSite site)
        {
            return System.Text.Encoding.Unicode.GetBytes(Common.GenerateResourceName(site, "FileContent"));
        }

        /// <summary>
        /// A method used to generate a random string in the range A-Z with the specified string size.
        /// </summary>
        /// <param name="size">A parameter represents the generated string size.</param>
        /// <returns>A return value represents the random generated string.</returns>
        public static string GenerateRandomString(int size)
        {
            Random random = new Random((int)DateTime.Now.Ticks);
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < size; i++)
            {
                int intIndex = Convert.ToInt32(Math.Floor((26 * random.NextDouble()) + 65));
                char ch = Convert.ToChar(intIndex);
                builder.Append(ch);
            }

            return builder.ToString();
        }

        /// <summary>
        /// This method is used to generate ETag value.
        /// </summary>
        /// <returns>Return the ETag value.</returns>
        public static string GenerateRandomETag()
        {
            return string.Format("{{{0},{1}}}", System.Guid.NewGuid().ToString("N"), SequenceNumberGenerator.GetCurrentSerialNumber());
        }

        /// <summary>
        /// This method is used to test two FsshttpbResponse which contains only PutChangesResponse are roughly equaled.
        /// </summary>
        /// <param name="fsshttpbResponse1">Specify the first FsshttpbResponse instance.</param>
        /// <param name="fsshttpbResponse2">Specify the second FsshttpbResponse instance.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        /// <returns>Return true if equals, otherwise it returns false.</returns>
        public static bool CompareSucceedFsshttpbPutChangesResponse(FsshttpbResponse fsshttpbResponse1, FsshttpbResponse fsshttpbResponse2, ITestSite site)
        {
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse1, site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse2, site);

            if (fsshttpbResponse1.Status != fsshttpbResponse2.Status)
            {
                return false;
            }

            if (fsshttpbResponse1.CellSubResponses != null && fsshttpbResponse2.CellSubResponses != null)
            {
                if (fsshttpbResponse1.CellSubResponses.Count != fsshttpbResponse2.CellSubResponses.Count)
                {
                    return false;
                }

                for (int index = 0; index < fsshttpbResponse1.CellSubResponses.Count; index++)
                {
                    if (fsshttpbResponse1.CellSubResponses[index].Status != fsshttpbResponse2.CellSubResponses[index].Status)
                    {
                        return false;
                    }
                }
            }
            else
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// This method is used to test two FsshttpbResponse which contains only AllocateExtendedGuidRangeResponse are roughly equaled.
        /// </summary>
        /// <param name="fsshttpbResponse1">Specify the first FsshttpbResponse instance.</param>
        /// <param name="fsshttpbResponse2">Specify the second FsshttpbResponse instance.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        /// <returns>Return true if equals, otherwise it returns false.</returns>
        public static bool ComapreSucceedFsshttpAllocateExtendedGuidRangeResposne(FsshttpbResponse fsshttpbResponse1, FsshttpbResponse fsshttpbResponse2, ITestSite site)
        {
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse1, site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse2, site);

            for (int index = 0; index < fsshttpbResponse1.CellSubResponses.Count; index++)
            {
                if (fsshttpbResponse1.CellSubResponses[index].Status != fsshttpbResponse2.CellSubResponses[index].Status)
                {
                    return false;
                }

                AllocateExtendedGuidRangeSubResponseData allocateResponse1 = fsshttpbResponse1.CellSubResponses[index].GetSubResponseData<AllocateExtendedGuidRangeSubResponseData>();
                AllocateExtendedGuidRangeSubResponseData allocateResponse2 = fsshttpbResponse2.CellSubResponses[index].GetSubResponseData<AllocateExtendedGuidRangeSubResponseData>();

                if (allocateResponse1.IntegerRangeMax.DecodedValue != allocateResponse2.IntegerRangeMax.DecodedValue)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// This method is used to test GUID Combined with the From sequence number forms the starting serial number of the range.
        /// </summary>
        /// <param name="fsshttpbResponse">Specify the FsshttpbResponse instance</param>
        /// <param name="cellSpecializedKnowledgeData">Specify the cellSpecializedKnowledgeData instance</param>
        /// <returns>Return true if From sequence number forms the starting serial number of the range, otherwise return false.</returns>
        public static bool CheckFromSequenceNumber(FsshttpbResponse fsshttpbResponse, CellKnowledge cellSpecializedKnowledgeData)
        {
            int istrue = 0;
            bool isVerifyR2126 = false;
            foreach (DataElement dataElement in fsshttpbResponse.DataElementPackage.DataElements)
            {
                foreach (CellKnowledgeRange cellKnowledgeRange in cellSpecializedKnowledgeData.CellKnowledgeRangeList)
                {
                    if (dataElement.SerialNumber.GUID == cellKnowledgeRange.CellKnowledgeRangeGUID)
                    {
                        if (dataElement.SerialNumber.Value >= cellKnowledgeRange.From.DecodedValue)
                        {
                            istrue++;
                            break;
                        }
                        else
                        {
                            istrue = -1;
                            break;
                        }
                    }
                }

                if (istrue == -1)
                {
                    break;
                }
            }

            if (istrue == 0 || istrue == -1)
            {
                isVerifyR2126 = false;
            }

            if (istrue > 0)
            {
                isVerifyR2126 = true;
            }

            return isVerifyR2126;
        }

        /// <summary>
        /// This method is used to test GUID Combined with the To sequence number forms the ending serial number of the range.
        /// </summary>
        /// <param name="fsshttpbResponse">Specify the FsshttpbResponse instance</param>
        /// <param name="cellSpecializedKnowledgeData">Specify the cellSpecializedKnowledgeData instance</param>
        /// <returns>Return true if To sequence number forms the ending serial number of the range, otherwise false. </returns>
        public static bool CheckToSequenceNumber(FsshttpbResponse fsshttpbResponse, CellKnowledge cellSpecializedKnowledgeData)
        {
            int istrue = 0;
            bool isVerifyR2127 = false;
            foreach (DataElement dataElement in fsshttpbResponse.DataElementPackage.DataElements)
            {
                foreach (CellKnowledgeRange cellKnowledgeRange in cellSpecializedKnowledgeData.CellKnowledgeRangeList)
                {
                    if (dataElement.SerialNumber.GUID == cellKnowledgeRange.CellKnowledgeRangeGUID)
                    {
                        if (dataElement.SerialNumber.Value <= cellKnowledgeRange.To.DecodedValue)
                        {
                            istrue++;
                            break;
                        }
                        else
                        {
                            istrue = -1;
                            break;
                        }
                    }
                }

                if (istrue == -1)
                {
                    break;
                }
            }

            if (istrue == 0 || istrue == -1)
            {
                isVerifyR2127 = false;
            }

            if (istrue > 0)
            {
                isVerifyR2127 = true;
            }

            return isVerifyR2127;
        }
    }
}