namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    /// <summary>
    /// The enumeration of request type.
    /// </summary>
    public enum RequestTypes
    {
        /// <summary>
        /// Query access.
        /// </summary>
        QueryAccess = 1,

        /// <summary>
        /// Query changes.
        /// </summary>
        QueryChanges = 2,

        /// <summary>
        /// Query knowledge.
        /// </summary>
        QueryKnowledge = 3,

        /// <summary>
        /// Put changes.
        /// </summary>
        PutChanges = 5,

        /// <summary>
        /// Query raw storage.
        /// </summary>
        QueryRawStorage = 6,

        /// <summary>
        /// Put raw storage.
        /// </summary>
        PutRawStorage = 7,

        /// <summary>
        /// Query diagnostic store info.
        /// </summary>
        QueryDiagnosticStoreInfo = 8,

        /// <summary>
        /// Allocate extended Guid range .
        /// </summary>
        AllocateExtendedGuidRange = 11
    }

    /// <summary>
    /// The enumeration of the filter type
    /// </summary>
    public enum FilterType
    {
        /// <summary>
        /// All filter
        /// </summary>
        AllFilter = 1,

        /// <summary>
        /// Data element type filter
        /// </summary>
        DataElementTypeFilter,

        /// <summary>
        /// Storage index referenced data elements filter
        /// </summary>
        StorageIndexReferencedDataElementsFilter,

        /// <summary>
        /// Cell ID filter
        /// </summary>
        CellIDFilter,

        /// <summary>
        /// Custom filter
        /// </summary>
        CustomFilter,

        /// <summary>
        /// Data element IDs filter
        /// </summary>
        DataElementIDsFilter,

        /// <summary>
        /// Hierarchy filter
        /// </summary>
        HierarchyFilter,
    }

    /// <summary>
    /// The enumeration of the data element type
    /// </summary>
    public enum DataElementType
    {
        /// <summary>
        /// None data element type
        /// </summary>
        None = 0,
        
        /// <summary>
        /// Storage Index Data Element
        /// </summary>
        StorageIndexDataElementData = 1,
        
        /// <summary>
        /// Storage Manifest Data Element
        /// </summary>
        StorageManifestDataElementData = 2,
        
        /// <summary>
        /// Cell Manifest Data Element
        /// </summary>
        CellManifestDataElementData = 3,
        
        /// <summary>
        /// Revision Manifest Data Element
        /// </summary>
        RevisionManifestDataElementData = 4,
        
        /// <summary>
        /// Object Group Data Element
        /// </summary>
        ObjectGroupDataElementData = 5,
        
        /// <summary>
        /// Fragment Data Element
        /// </summary>
        FragmentDataElementData = 6,
        
        /// <summary>
        /// Object Data BLOB Data Element
        /// </summary>
        ObjectDataBLOBDataElementData = 10
    }

    /// <summary>
    /// The enumeration of the hierarchy filter depth
    /// </summary>
    public enum HierarchyFilterDepth
    {
        /// <summary>
        /// Index values corresponding to the specified keys only.
        /// </summary>
        IndexOnly = 0,
        
        /// <summary>
        /// First data elements referenced by the storage index values corresponding to the specified keys only.
        /// </summary>
        FirstDataElement = 1,
        
        /// <summary>
        /// Single level. All data elements under the sub-graphs rooted by the specified keys stopping at any storage index entries.
        /// </summary>
        SingleLevel = 2,
        
        /// <summary>
        /// Deep. All data elements and storage index entries under the sub-graphs rooted by the specified keys.
        /// </summary>
        Deep = 3
    }

    /// <summary>
    /// The enumeration of the knowledge type
    /// </summary>
    public enum KnowledgeType
    {
        /// <summary>
        /// Cell knowledge
        /// </summary>
        Cellknowledge = 0,

        /// <summary>
        /// Waterline knowledge
        /// </summary>
        Waterlineknowledge,

        /// <summary>
        /// Fragment knowledge
        /// </summary>
        Fragmentknowledge,

        /// <summary>
        /// Content tag knowledge
        /// </summary>
        Contenttagknowledge
    }

    /// <summary>
    /// The enumeration of the stream object type header start
    /// </summary>
    public enum StreamObjectTypeHeaderStart
    {
        /// <summary>
        /// ErrorStringSupplementalInfo type in the ResponseError
        /// </summary>
        ErrorStringSupplementalInfo = 0x4E,

        /// <summary>
        /// Data Element
        /// </summary>
        DataElement = 0x01,

        /// <summary>
        /// Object Data BLOB
        /// </summary>
        ObjectDataBLOB = 0x02,

        /// <summary>
        /// Waterline Knowledge Entry
        /// </summary>
        WaterlineKnowledgeEntry = 0x04,

        /// <summary>
        /// Object Group Object BLOB Data Declaration
        /// </summary>
        ObjectGroupObjectBLOBDataDeclaration = 0x05,

        /// <summary>
        /// Storage Manifest Root Declare
        /// </summary>
        StorageManifestRootDeclare = 0x07,

        /// <summary>
        /// Revision Manifest Root Declare
        /// </summary>
        RevisionManifestRootDeclare = 0x0A,

        /// <summary>
        /// Cell Manifest Current Revision
        /// </summary>
        CellManifestCurrentRevision = 0x0B,

        /// <summary>
        /// Storage Manifest Schema GUID
        /// </summary>
        StorageManifestSchemaGUID = 0x0C,

        /// <summary>
        /// Storage Index Revision Mapping
        /// </summary>
        StorageIndexRevisionMapping = 0x0D,

        /// <summary>
        /// Storage Index Cell Mapping
        /// </summary>
        StorageIndexCellMapping = 0x0E,

        /// <summary>
        /// Cell Knowledge Range
        /// </summary>
        CellKnowledgeRange = 0x0F,

        /// <summary>
        /// The Knowledge
        /// </summary>
        Knowledge = 0x10,

        /// <summary>
        /// Storage Index Manifest Mapping
        /// </summary>
        StorageIndexManifestMapping = 0x11,

        /// <summary>
        /// Cell Knowledge
        /// </summary>
        CellKnowledge = 0x14,

        /// <summary>
        /// Data Element Package
        /// </summary>
        DataElementPackage = 0x15,

        /// <summary>
        /// Object Group Object Data
        /// </summary>
        ObjectGroupObjectData = 0x16,

        /// <summary>
        /// Cell Knowledge Entry
        /// </summary>
        CellKnowledgeEntry = 0x17,

        /// <summary>
        /// Object Group Object Declare
        /// </summary>
        ObjectGroupObjectDeclare = 0x18,

        /// <summary>
        /// Revision Manifest Object Group References
        /// </summary>
        RevisionManifestObjectGroupReferences = 0x19,

        /// <summary>
        /// Revision Manifest
        /// </summary>
        RevisionManifest = 0x1A,

        /// <summary>
        /// Object Group Object Data BLOB Reference
        /// </summary>
        ObjectGroupObjectDataBLOBReference = 0x1C,

        /// <summary>
        /// Object Group Declarations
        /// </summary>
        ObjectGroupDeclarations = 0x1D,

        /// <summary>
        /// Object Group Data
        /// </summary>
        ObjectGroupData = 0x1E,

        /// <summary>
        /// Intermediate Node Object
        /// </summary>
        LeafNodeObject = 0x1F, // Defined in MS-FSSHTTPD

        /// <summary>
        /// Root Node Object
        /// </summary>
        IntermediateNodeObject = 0x20, // Defined in MS-FSSHTTPD

        /// <summary>
        /// Signature Object
        /// </summary>
        SignatureObject = 0x21, // Defined in MS-FSSHTTPD

        /// <summary>
        /// Data Size Object
        /// </summary>
        DataSizeObject = 0x22, // Defined in MS-FSSHTTPD

        /// <summary>
        /// Waterline Knowledge
        /// </summary>
        WaterlineKnowledge = 0x29,

        /// <summary>
        /// Content Tag Knowledge
        /// </summary>
        ContentTagKnowledge = 0x2D,

        /// <summary>
        /// Content Tag Knowledge Entry
        /// </summary>
        ContentTagKnowledgeEntry = 0x2E,

        /// <summary>
        /// The Request
        /// </summary>
        Request = 0x040,

        /// <summary>
        /// FSSHTTPB Sub Response
        /// </summary>
        FsshttpbSubResponse = 0x041,

        /// <summary>
        /// Sub Request
        /// </summary>
        SubRequest = 0x042,

        /// <summary>
        /// Read Access Response
        /// </summary>
        ReadAccessResponse = 0x043,

        /// <summary>
        /// Specialized Knowledge
        /// </summary>
        SpecializedKnowledge = 0x044,

        /// <summary>
        /// PutChanges Response SerialNumber ReassignAll
        /// </summary>
        PutChangesResponseSerialNumberReassignAll = 0x045,

        /// <summary>
        /// Write Access Response
        /// </summary>
        WriteAccessResponse = 0x046,

        /// <summary>
        /// Query Changes Filter
        /// </summary>
        QueryChangesFilter = 0x047,

        /// <summary>
        /// Win32 Error
        /// </summary>
        Win32Error = 0x049,

        /// <summary>
        /// Protocol Error
        /// </summary>
        ProtocolError = 0x04B,

        /// <summary>
        /// Response Error
        /// </summary>
        ResponseError = 0x04D,

        /// <summary>
        /// User Agent version
        /// </summary>
        UserAgentversion = 0x04F,

        /// <summary>
        /// QueryChanges Filter Schema Specific
        /// </summary>
        QueryChangesFilterSchemaSpecific = 0x050,

        /// <summary>
        /// QueryChanges Request
        /// </summary>
        QueryChangesRequest = 0x051,

        /// <summary>
        /// HRESULT Error
        /// </summary>
        HRESULTError = 0x052,

        /// <summary>
        /// PutChanges Response SerialNumberReassign
        /// </summary>
        PutChangesResponseSerialNumberReassign = 0x053,

        /// <summary>
        /// QueryChanges Filter DataElement IDs
        /// </summary>
        QueryChangesFilterDataElementIDs = 0x054,

        /// <summary>
        /// User Agent GUID
        /// </summary>
        UserAgentGUID = 0x055,

        /// <summary>
        /// QueryChanges Filter Data Element Type
        /// </summary>
        QueryChangesFilterDataElementType = 0x057,

        /// <summary>
        /// Query Changes Data Constraint
        /// </summary>
        QueryChangesDataConstraint = 0x059,

        /// <summary>
        /// PutChanges Request
        /// </summary>
        PutChangesRequest = 0x05A,

        /// <summary>
        /// Query Changes Request Arguments
        /// </summary>
        QueryChangesRequestArguments = 0x05B,

        /// <summary>
        /// Query Changes Filter Cell ID
        /// </summary>
        QueryChangesFilterCellID = 0x05C,

        /// <summary>
        /// User Agent
        /// </summary>
        UserAgent = 0x05D,

        /// <summary>
        /// Query Changes Response
        /// </summary>
        QueryChangesResponse = 0x05F,

        /// <summary>
        /// Query Changes Filter Hierarchy
        /// </summary>
        QueryChangesFilterHierarchy = 0x060,

        /// <summary>
        /// The Response
        /// </summary>
        FsshttpbResponse = 0x062,

        /// <summary>
        /// Query Data Element Request
        /// </summary>
        QueryDataElementRequest = 0x065,

        /// <summary>
        /// Cell Error
        /// </summary>
        CellError = 0x066,

        /// <summary>
        /// Query Changes Filter Flags
        /// </summary>
        QueryChangesFilterFlags = 0x068,

        /// <summary>
        /// Data Element Fragment
        /// </summary>
        DataElementFragment = 0x06A,

        /// <summary>
        /// Fragment Knowledge
        /// </summary>
        FragmentKnowledge = 0x06B,

        /// <summary>
        /// Fragment Knowledge Entry
        /// </summary>
        FragmentKnowledgeEntry = 0x06C,

        /// <summary>
        /// Object Group Metadata Declarations
        /// </summary>
        ObjectGroupMetadataDeclarations = 0x79,
        
        /// <summary>
        /// Object Group Metadata
        /// </summary>
        ObjectGroupMetadata = 0x78,

        /// <summary>
        /// Allocate Extended GUID Range Request
        /// </summary>
        AllocateExtendedGUIDRangeRequest = 0x080,

        /// <summary>
        /// Allocate Extended GUID Range Response
        /// </summary>
        AllocateExtendedGUIDRangeResponse = 0x081,

        /// <summary>
        /// Request Hash Options
        /// </summary>
        RequestHashOptions = 0x088,

        /// <summary>
        /// Target Partition Id
        /// </summary>
        TargetPartitionId = 0x083,

        /// <summary>
        /// Put Changes Response
        /// </summary>
        PutChangesResponse = 0x087,

        /// <summary>
        /// Diagnostic Request Option Output
        /// </summary>
        DiagnosticRequestOptionOutput = 0x089,

        /// <summary>
        /// Additional Flags
        /// </summary>
        AdditionalFlags = 0x86,

        /// <summary>
        /// Put changes lock id
        /// </summary>
        PutChangesLockId = 0x85,

        /// <summary>
        /// This value is wrong
        /// </summary>
        DataElementHash = 0x100
    }

    /// <summary>
    /// The enumeration of the stream object type header end
    /// </summary>
    public enum StreamObjectTypeHeaderEnd
    {
        /// <summary>
        /// Data Element
        /// </summary>
        DataElement = 0x01,
        
        /// <summary>
        /// The Knowledge
        /// </summary>
        Knowledge = 0x10,
        
        /// <summary>
        /// Cell Knowledge
        /// </summary>
        CellKnowledge = 0x14,
        
        /// <summary>
        /// Data Element Package
        /// </summary>
        DataElementPackage = 0x15,
        
        /// <summary>
        /// Object Group Declarations
        /// </summary>
        ObjectGroupDeclarations = 0x1D,
        
        /// <summary>
        /// Object Group Data
        /// </summary>
        ObjectGroupData = 0x1E,
        
        /// <summary>
        /// Intermediate Node End
        /// </summary>
        IntermediateNodeEnd = 0x1F, // Defined in MS-FSSHTTPD
        
        /// <summary>
        /// Root Node End
        /// </summary>
        RootNodeEnd = 0x20, // Defined in MS-FSSHTTPD
        
        /// <summary>
        /// Waterline Knowledge
        /// </summary>
        WaterlineKnowledge = 0x29,
        
        /// <summary>
        /// Content Tag Knowledge
        /// </summary>
        ContentTagKnowledge = 0x2D,
        
        /// <summary>
        /// The Request
        /// </summary>
        Request = 0x040,
        
        /// <summary>
        /// Sub Response
        /// </summary>
        SubResponse = 0x041,
        
        /// <summary>
        /// Sub Request
        /// </summary>
        SubRequest = 0x042,
        
        /// <summary>
        /// Read Access Response
        /// </summary>
        ReadAccessResponse = 0x043,
        
        /// <summary>
        /// Specialized Knowledge
        /// </summary>
        SpecializedKnowledge = 0x044,
        
        /// <summary>
        /// Write Access Response
        /// </summary>
        WriteAccessResponse = 0x046,
        
        /// <summary>
        /// Query Changes Filter
        /// </summary>
        QueryChangesFilter = 0x047,
        
        /// <summary>
        /// The Error type
        /// </summary>
        Error = 0x04D,
        
        /// <summary>
        /// Query Changes Request
        /// </summary>
        QueryChangesRequest = 0x051,
        
        /// <summary>
        /// User Agent
        /// </summary>
        UserAgent = 0x05D,
        
        /// <summary>
        /// The Response
        /// </summary>
        Response = 0x062,
        
        /// <summary>
        /// Fragment Knowledge
        /// </summary>
        FragmentKnowledge = 0x06B,
        
        /// <summary>
        /// Object Group Metadata Declarations, new added in MOSS2013.
        /// </summary>
        ObjectGroupMetadataDeclarations = 0x79,
        
        /// <summary>
        /// Target PartitionId, new added in MOSS2013.
        /// </summary>
        TargetPartitionId = 0x083
    }

    /// <summary>
    /// The enumeration of the cell error code
    /// </summary>
    public enum CellErrorCode
    {
        /// <summary>
        /// Unknown error
        /// </summary>
        Unknownerror = 1,

        /// <summary>
        /// Invalid object
        /// </summary>
        InvalidObject = 2,

        /// <summary>
        /// Invalid partition
        /// </summary>
        Invalidpartition = 3,

        /// <summary>
        /// Request not supported
        /// </summary>
        Requestnotsupported = 4,

        /// <summary>
        /// Storage read-only
        /// </summary>
        Storagereadonly = 5,

        /// <summary>
        /// Revision ID not found
        /// </summary>
        RevisionIDnotfound = 6,

        /// <summary>
        /// The Bad token
        /// </summary>
        Badtoken = 7,

        /// <summary>
        /// Request not finished
        /// </summary>
        Requestnotfinished = 8,

        /// <summary>
        /// Incompatible token
        /// </summary>
        Incompatibletoken = 9,

        /// <summary>
        /// Scoped cell storage
        /// </summary>
        Scopedcellstorage = 11,

        /// <summary>
        /// Coherency failure
        /// </summary>
        Coherencyfailure = 12,

        /// <summary>
        /// Cell storage state deserialization failure
        /// </summary>
        Cellstoragestatedeserializationfailure = 13,
        
        /// <summary>
        /// Incompatible protocol version
        /// </summary>
        Incompatibleprotocolversion = 15,

        /// <summary>
        /// Referenced data element not found
        /// </summary>
        Referenceddataelementnotfound = 16,

        /// <summary>
        /// Request stream schema error
        /// </summary>
        Requeststreamschemaerror = 18,

        /// <summary>
        /// Response stream schema error
        /// </summary>
        Responsestreamschemaerror = 19,

        /// <summary>
        /// Unknown request
        /// </summary>
        Unknownrequest = 20,

        /// <summary>
        /// Storage failure
        /// </summary>
        Storagefailure = 21,

        /// <summary>
        /// Storage write only
        /// </summary>
        Storagewriteonly = 22,

        /// <summary>
        /// Invalid serialization
        /// </summary>
        Invalidserialization = 23,

        /// <summary>
        /// Data element not found
        /// </summary>
        Dataelementnotfound = 24,

        /// <summary>
        /// Invalid implementation
        /// </summary>
        Invalidimplementation = 25,

        /// <summary>
        /// Incompatible old storage
        /// </summary>
        Incompatibleoldstorage = 26,

        /// <summary>
        /// Incompatible new storage
        /// </summary>
        Incompatiblenewstorage = 27,
        
        /// <summary>
        /// Incorrect context for data element ID
        /// </summary>
        IncorrectcontextfordataelementID = 28,

        /// <summary>
        /// Object group duplicate objects
        /// </summary>
        Objectgroupduplicateobjects = 29,

        /// <summary>
        /// Object reference not founding revision
        /// </summary>
        Objectreferencenotfoundinrevision = 31,

        /// <summary>
        /// Merge cell storage state conflict
        /// </summary>
        Mergecellstoragestateconflict = 32,

        /// <summary>
        /// Unknown query changes filter
        /// </summary>
        Unknownquerychangesfilter = 33,

        /// <summary>
        /// Unsupported query changes filter
        /// </summary>
        Unsupportedquerychangesfilter = 34,

        /// <summary>
        /// Unable to provide knowledge
        /// </summary>
        Unabletoprovideknowledge = 35,

        /// <summary>
        /// Data element missing ID
        /// </summary>
        DataelementmissingID = 36,

        /// <summary>
        /// Data element missing serial number
        /// </summary>
        Dataelementmissingserialnumber = 37,

        /// <summary>
        /// Request argument invalid
        /// </summary>
        Requestargumentinvalid = 38,

        /// <summary>
        /// Partial changes not supported
        /// </summary>
        Partialchangesnotsupported = 39,

        /// <summary>
        /// Store busy retry later
        /// </summary>
        Storebusyretrylater = 40,

        /// <summary>
        /// GUIDID table not supported
        /// </summary>
        GUIDIDtablenotsupported = 41,

        /// <summary>
        /// Data element cycle
        /// </summary>
        Dataelementcycle = 42,

        /// <summary>
        /// Fragment knowledge error
        /// </summary>
        Fragmentknowledgeerror = 43,

        /// <summary>
        /// Fragment size mismatch
        /// </summary>
        Fragmentsizemismatch = 44,

        /// <summary>
        /// Fragments incomplete
        /// </summary>
        Fragmentsincomplete = 45,

        /// <summary>
        /// Fragment invalid
        /// </summary>
        Fragmentinvalid = 46,

        /// <summary>
        /// Aborted after failed put changes
        /// </summary>
        Abortedafterfailedputchanges = 47,

        /// <summary>
        /// Upgrade failed because there are no upgradeable contents.
        /// </summary>
        FailedNoUpgradeableContents = 79,

        /// <summary>
        /// Unable to allocate additional extended GUIDs.
        /// </summary>
        UnableAllocateAdditionalExtendedGuids = 106,

        /// <summary>
        /// Site is in read-only mode.
        /// </summary>
        SiteReadonlyMode = 108,
    
        /// <summary>
        /// Multi-Request partition reached quota.
        /// </summary>
        MultiRequestPartitionReachQutoa = 111,

        /// <summary>
        /// Extended GUID collision.
        /// </summary>
        ExtendedGuidCollision = 112,

        /// <summary>
        /// Upgrade failed because of insufficient permissions.
        /// </summary>
        InsufficientPermisssions = 113,

        /// <summary>
        /// Upgrade failed because of server throttling.
        /// </summary>
        ServerThrottling = 114,

        /// <summary>
        /// Upgrade failed because the upgraded file is too large.
        /// </summary>
        FileTooLarge = 115
    }

    /// <summary>
    /// The enumeration of the protocol error code
    /// </summary>
    public enum ProtocolErrorCode
    {
        /// <summary>
        /// Unknown error
        /// </summary>
        Unknownerror = 1,
        
        /// <summary>
        /// End of Stream
        /// </summary>
        EndofStream = 50,
        
        /// <summary>
        /// Unknown internal error
        /// </summary>
        Unknowninternalerror = 61,
        
        /// <summary>
        /// Input stream schema invalid
        /// </summary>
        Inputstreamschemainvalid = 108,
        
        /// <summary>
        /// Stream object invalid
        /// </summary>
        Streamobjectinvalid = 142,
        
        /// <summary>
        /// Stream object unexpected
        /// </summary>
        Streamobjectunexpected = 143,
        
        /// <summary>
        /// Server URL not found
        /// </summary>
        ServerURLnotfound = 144,
        
        /// <summary>
        /// Stream object serialization error
        /// </summary>
        Streamobjectserializationerror = 145,
    }

    /// <summary>
    /// The enumeration of the chunking methods.
    /// </summary>
    public enum ChunkingMethod
    {
        /// <summary>
        /// File data is passed to the Zip algorithm chunking method.
        /// </summary>
        ZipAlgorithm,

        /// <summary>
        /// File data is passed to the RDC Analysis chunking method.
        /// </summary>
        RDCAnalysis,

        /// <summary>
        /// File data is passed to the Simple algorithm chunking method.
        /// </summary>
        SimpleAlgorithm
    }
}