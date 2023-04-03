namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// This class specifies query change sub-request.
    /// </summary>
    public class QueryChangesCellSubRequest : FsshttpbCellSubRequest
    {
        /// <summary>
        /// Initializes a new instance of the QueryChangesCellSubRequest class
        /// </summary>
        /// <param name="subRequestID">Specify the sub request id</param>
        public QueryChangesCellSubRequest(ulong subRequestID)
        {
            this.RequestID = subRequestID;
            this.RequestType = Convert.ToUInt64(RequestTypes.QueryChanges);
            this.QueryChangesRequest = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.QueryChangesRequest, 2);
            this.Reserved = 0;
            this.AllowFragments = 0;
            this.Reserved1 = 0;
            this.IncludeStorageManifest = 0;
            this.IncludeCellChanges = 0;
            this.Reserved2 = 0;
            this.AllowFragments2 = 0;
            this.RoundKnowledgeToWholeCellChanges = 0;

            // CellId
            ExGuid extendGuid1 = new ExGuid(0x0, Guid.Empty);
            ExGuid extendGuid2 = new ExGuid(0x0, Guid.Empty);
            this.CellId = new CellID(extendGuid1, extendGuid2);
            this.QueryChangeFilters = new List<Filter>();
        }

        /// <summary>
        /// Gets or sets Query Changes Request (4 bytes): A 32-bit stream object header that specifies a query changes request.
        /// </summary>
        public StreamObjectHeaderStart32bit QueryChangesRequest { get; set; }

        /// <summary>
        /// Gets or sets R Reserved (1 bit): A reserved bit that MUST be set to zero and MUST be ignored.
        /// </summary>
        public int Reserved { get; set; }

        /// <summary>
        /// Gets or sets Allow Fragments (1 bit): A bit that specifies if set to allow fragments, otherwise it does not allow fragments.
        /// </summary>
        public int AllowFragments { get; set; }

        /// <summary>
        /// Gets or sets Allow Fragments (1 bit): A bit that specifies if set to allow fragments, otherwise it does not allow fragments.
        /// </summary>
        public int AllowFragments2 { get; set; }

        /// <summary>
        /// Gets or sets Round Knowledge to Whole Cell Changes (1 bit): F ?Round Knowledge to Whole Cell Changes (1 bit): If set, a bit that specifies that the knowledge specified in the request MUST be modified, prior to change enumeration, such that any changes under a cell node, as implied by the knowledge, cause the knowledge to be modified such that all changes in that cell are returned.
        /// </summary>
        public int RoundKnowledgeToWholeCellChanges { get; set; }

        /// <summary>
        /// Gets or sets exclude object data (1 bit): A bit that specifies to exclude object data; otherwise, object data is included.
        /// </summary>
        public int ExcludeObjectData { get; set; }

        /// <summary>
        /// Gets or sets Include Filtered Out Data Elements In Knowledge (1 bit): A bit that specifies to include the serial numbers of filtered out data elements in the response knowledge; otherwise, the serial numbers of filtered out data elements are not included in the response knowledge.
        /// </summary>
        public int IncludeFilteredOutDataElementsInKnowledge { get; set; }

        /// <summary>
        /// Gets or sets Reserved1 (2 bits): A 6-bit reserved field that MUST be set to zero and MUST be ignored.
        /// </summary>
        public int Reserved1 { get; set; }

        /// <summary>
        /// User Content Equivalent Version Ok (1 bit): This attribute MAY be set. If set, a bit that specifies that if the version of the file cannot be found, it is acceptable to return a substitute version that is equivalent from a user authored content perspective, if such a version exists.
        /// </summary>
        public int UserContentEquivalentVersionOk { get; set; }

        /// <summary>
        /// Reserved (7 bits): A 7-bit reserved field that MUST be set to zero and MUST be ignored.
        /// </summary>
        public int ReservedMustBeZero { get; set; }
        /// <summary>
        /// Gets or sets Query Changes Request Arguments (4 bytes): A 32-bit stream object header that specifies a query changes request arguments.
        /// </summary>
        public StreamObjectHeaderStart32bit QueryChangesRequestArguments { get; set; }

        /// <summary>
        /// Gets or sets S - Include Storage Manifest (1 bit): A bit that specifies if set to include the storage manifest, otherwise the storage manifest is not included.
        /// </summary>
        public int IncludeStorageManifest { get; set; }

        /// <summary>
        /// Gets or sets C - Include Cell Changes (1 bit): A bit that specifies if set to include cell changes, otherwise cell changes are not included.
        /// </summary>
        public int IncludeCellChanges { get; set; }

        /// <summary>
        /// Gets or sets Reserved2 (6 bits): A 6-bit reserved field that MUST be set to zero and MUST be ignored.
        /// </summary>
        public int Reserved2 { get; set; }

        /// <summary>
        /// Gets or sets Cell ID (variable): A cell ID that specifies if the query changes are scoped to a specific cell. If the cell ID is 0x0000, no scoping restriction is specified.
        /// </summary>
        public CellID CellId { get; set; }

        /// <summary>
        /// Gets or sets Query Changes Data Constraints (4 bytes): A 32-bit stream object header that specifies a query changes Data Constraints, this is optional if no maximum data elements constraint is required.
        /// </summary>
        public StreamObjectHeaderStart32bit QueryChangesDataConstraints { get; set; }

        /// <summary>
        /// Gets or sets Max Data Elements 
        /// </summary>
        public Compact64bitInt MaxDataElements { get; set; }

        /// <summary>
        /// Gets or sets Query Changes Filter (variable): An optional ordered array of filters that specifies how the results of the query will be filtered before it is returned to the client.
        /// </summary>
        public List<Filter> QueryChangeFilters { get; set; }
        
        /// <summary>
        /// Gets or sets Knowledge (variable): An optional knowledge that specify what the client knows about a state of a file.
        /// </summary>
        public Knowledge Knowledge { get; set; }

        /// <summary>
        /// This method is used to convert the element into a byte List 
        /// </summary>
        /// <returns>Return the Byte List</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(base.SerializeToByteList());
            
            // Query Changes Request 
            byteList.AddRange(this.QueryChangesRequest.SerializeToByteList());
            
            // Reserved1
            BitWriter bitWriter = new BitWriter(1);
            bitWriter.AppendInit32(this.Reserved, 1);
            bitWriter.AppendInit32(this.AllowFragments, 1);
            bitWriter.AppendInit32(this.ExcludeObjectData, 1);
            bitWriter.AppendInit32(this.IncludeFilteredOutDataElementsInKnowledge, 1);
            bitWriter.AppendInit32(this.AllowFragments2, 1);
            bitWriter.AppendInit32(this.RoundKnowledgeToWholeCellChanges, 1);
            bitWriter.AppendInit32(this.Reserved1, 2);
            byteList.AddRange(bitWriter.Bytes);

            // User Content Equivalent Version Ok and Reserved
            bitWriter = new BitWriter(1);
            bitWriter.AppendInit32(this.UserContentEquivalentVersionOk, 1);
            bitWriter.AppendInit32(this.ReservedMustBeZero, 7);
            byteList.AddRange(bitWriter.Bytes);

            // Cell ID bytes
            List<byte> cellIDBytes = this.CellId.SerializeToByteList();

            // Query Changes Request Arguments 
            this.QueryChangesRequestArguments = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.QueryChangesRequestArguments, 1 + cellIDBytes.Count);
            byteList.AddRange(this.QueryChangesRequestArguments.SerializeToByteList());

            // Reserved2
            bitWriter = new BitWriter(1);
            bitWriter.AppendInit32(this.IncludeStorageManifest, 1);
            bitWriter.AppendInit32(this.IncludeCellChanges, 1);
            bitWriter.AppendInit32(this.Reserved2, 6);
            byteList.AddRange(bitWriter.Bytes);

            // Cell ID
            byteList.AddRange(cellIDBytes);

            // optional
            if (this.MaxDataElements != null)
            {
                if (this.MaxDataElements.DecodedValue > 0)
                {
                    // Max Data Elements bytes
                    List<byte> maxDataElementsBytes = (new Compact64bitInt(this.MaxDataElements.DecodedValue)).SerializeToByteList();
                    
                    // Query Changes Data Constraints 
                    this.QueryChangesDataConstraints = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.QueryChangesDataConstraint, maxDataElementsBytes.Count);
                    byteList.AddRange(this.QueryChangesDataConstraints.SerializeToByteList());
                    
                    // Max Data Elements
                    byteList.AddRange(maxDataElementsBytes);
                }
            }

            if (this.QueryChangeFilters != null)
            {
                foreach (Filter filter in this.QueryChangeFilters)
                {
                    byteList.AddRange(filter.SerializeToByteList());
                }
            }

            if (this.Knowledge != null)
            {
                byteList.AddRange(this.Knowledge.SerializeToByteList());
            }

            byteList.AddRange(this.ToBytesEnd());
            return byteList;
        }
    }

    /// <summary>
    /// A filter that specifies the query changes request
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public abstract class Filter : IFSSHTTPBSerializable
    {
        /// <summary>
        /// Initializes a new instance of the Filter class
        /// </summary>
        /// <param name="filterType">Filter Category</param>
        protected Filter(FilterType filterType)
        {
            this.QueryChangesFilterEnd = new StreamObjectHeaderEnd16bit((int)StreamObjectTypeHeaderStart.QueryChangesFilter);
            this.QueryChangesFilterStart = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.QueryChangesFilter, 2);
            this.QueryChangesFilterFlags = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.QueryChangesFilterFlags, 1);
            this.FilterType = filterType;
        }

        /// <summary>
        /// Gets or sets Query Changes Filter Start (4 bytes): A 32-bit stream object header that specifies a query changes filter start.
        /// </summary>
        public StreamObjectHeaderStart32bit QueryChangesFilterStart { get; set; }

        /// <summary>
        /// Gets or sets Filter Type (1 byte): A flag that specifies filter type to apply to the category, either 0 to exclude that filter category, or 1 to include it.
        /// </summary>
        public FilterType FilterType { get; set; }

        /// <summary>
        /// Gets or sets Filter Operation (1 byte): A flag that specifies how the filter is applied to the data elements before they are added to the response data element package. This field MUST be set to zero or one.
        /// </summary>
        public int FilterOperation { get; set; }

        /// <summary>
        /// Gets or sets Query Changes Filter Data (variable): A structure that specifies additional data based on the filter category.
        /// </summary>
        public Filter QueryChangesFilterData { get; set; }

        /// <summary>
        /// Gets or sets Query Changes Filter End (2 bytes): A 16-bit stream object header that specifies a query changes filter end.
        /// </summary>
        public StreamObjectHeaderEnd16bit QueryChangesFilterEnd { get; set; }

        /// <summary>
        /// Gets or sets Query Changes Filter Flags (4 bytes): An optional 32-bit stream object header that specifies a query changes filter flags.
        /// </summary>
        public StreamObjectHeaderStart32bit QueryChangesFilterFlags { get; set; }

        /// <summary>
        /// Gets or sets F - Fail if Unsupported (1 bit): A bit that specifies if set to 1 to allow failure if a filter is not supported, otherwise unsupported filters are ignored. This bit is only sent if query changes filter flags are specified.
        /// </summary>
        public bool? FailifUnsupported { get; set; }

        /// <summary>
        /// Gets or sets Reserved (7 bits): A 7-bit reserved and MUST be set to 0 and MUST be ignored if query changes filter flags are specified
        /// </summary>
        public int Reserved { get; set; }

        #region IFSSHTTPBSerialiable Members

        /// <summary>
        /// This method is used to serialize the element into a byte List.
        /// </summary>
        /// <returns>Return the Byte List.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();

            byteList.AddRange(this.QueryChangesFilterStart.SerializeToByteList());
            byteList.Add((byte)this.FilterType);
            byteList.Add((byte)this.FilterOperation);

            this.SerializeFilterData(byteList);
            byteList.AddRange(this.QueryChangesFilterEnd.SerializeToByteList());

            // This QueryChangesFilterFlags seems to be not supported in MOSS2013 and not sure
            // in the MOSS2010. In the current stage, just comment this out.
            // byteList.AddRange(this.QueryChangesFilterFlags.SerializeToByteList());
            // if (FailifUnsupported != null)
            // {
            //    byteList.AddRange(this.QueryChangesFilterFlags.SerializeToByteList());
            //    byteList.Add(Convert.ToByte(FailifUnsupported.Value));
            //// }

            return byteList;
        }

        /// <summary>
        /// This method is used to serialize the filter data to the given byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which will contain the filter data.</param>
        protected abstract void SerializeFilterData(List<byte> byteList);

        #endregion
    }

    /// <summary>
    /// All filter
    /// </summary>
    public class AllFilter : Filter
    {
        /// <summary>
        /// Initializes a new instance of the AllFilter class
        /// </summary>
        public AllFilter()
            : base(FilterType.AllFilter)
        {
        }

        /// <summary>
        /// This method is used to serialize the filter data to the given byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which will contain the filter data.</param>
        protected override void SerializeFilterData(List<byte> byteList)
        {
            // Do nothing according the current open specification document.
        }
    }

    /// <summary>
    /// Specifies the data element type to query.
    /// </summary>
    public class DataElementTypeFilter : Filter
    {
        /// <summary>
        /// Initializes a new instance of the DataElementTypeFilter class
        /// </summary>
        /// <param name="dataElementType">Specify the data element type.</param>
       public DataElementTypeFilter(DataElementType dataElementType)
            : base(FilterType.DataElementTypeFilter)
        {
            this.DataElementType = dataElementType;
            this.QueryChangesFilterDataElementType = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.QueryChangesFilterDataElementType, 2);
        }
 
        /// <summary>
        /// Gets or sets Query Changes Filter Data Element Type (4 bytes): A 32-bit stream object header that specifies a query changes filter data element type.
        /// </summary>
        public StreamObjectHeaderStart32bit QueryChangesFilterDataElementType { get; set; }
        
        /// <summary>
        /// Gets or sets Data Element Type (Variable): A compact unsigned 64-bit integer that specifies the data element type as:
        /// </summary>
        public DataElementType DataElementType { get; set; }
       
        /// <summary>
        /// This method is used to convert the element into a byte List 
        /// </summary>
        /// <param name="byteList">Specify the byte list which will contain the filter data.</param>
        protected override void SerializeFilterData(List<byte> byteList)
        {
            byteList.AddRange(this.QueryChangesFilterDataElementType.SerializeToByteList());
            byteList.AddRange(new Compact64bitInt((ulong)this.DataElementType).SerializeToByteList());
        }
    }

    /// <summary>
    /// Specifies the storage index referenced data element type to query.
    /// </summary>
    public class StorageIndexReferencedDataElementsFilter : Filter
    {
        /// <summary>
        /// Initializes a new instance of the StorageIndexReferencedDataElementsFilter class
        /// </summary>
        public StorageIndexReferencedDataElementsFilter()
            : base(FilterType.StorageIndexReferencedDataElementsFilter)
        {
        }

        /// <summary>
        /// This method is used to convert the element into a byte List 
        /// </summary>
        /// <param name="byteList">Specify the byte list which will contain the filter data.</param>
        protected override void SerializeFilterData(List<byte> byteList)
        {
            // Do nothing according the current open specification document.
        }
    }

    /// <summary>
    /// Specifies a particular cell to query.
    /// </summary>
    public class CellIDFilter : Filter
    {
        /// <summary>
        /// Initializes a new instance of the CellIDFilter class
        /// </summary>
        /// <param name="cellId">Specify the CellID.</param>
        public CellIDFilter(CellID cellId)
            : base(FilterType.CellIDFilter)
        {
            this.CellID = new CellID(cellId);
            this.QueryChangesFilterCellID = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.QueryChangesFilterCellID, this.CellID.SerializeToByteList().Count);
        }

        /// <summary>
        /// Gets or sets Query Changes Filter Cell ID (4 bytes): A 32-bit stream object header that specifies a query changes filter cell ID.
        /// </summary>
        public StreamObjectHeaderStart32bit QueryChangesFilterCellID { get; set; }

        /// <summary>
        /// Gets or sets Cell ID (variable): A cell ID that specifies the cell the query changes is scoped to.
        /// </summary>
        public CellID CellID { get; set; }

        /// <summary>
        /// This method is used to convert the element into a byte List 
        /// </summary>
        /// <param name="byteList">Specify the byte list which will contain the filter data.</param>
        protected override void SerializeFilterData(List<byte> byteList)
        {
            byteList.AddRange(this.QueryChangesFilterCellID.SerializeToByteList());
            byteList.AddRange(this.CellID.SerializeToByteList());
        }
    }

    /// <summary>
    /// Specifies a custom filter to apply.
    /// </summary>
    public class CustomFilter : Filter
    {
        /// <summary>
        /// Initializes a new instance of the CustomFilter class. Schema Filter Data (variable): A byte stream that specifies the schema filters data opaque to this protocol.
        /// </summary>
        /// <param name="schemaIdentifier">Specify the schema guid.</param>
        /// <param name="schemaFilterData">Specify the schema filter data.</param>
        public CustomFilter(Guid schemaIdentifier, byte[] schemaFilterData)
            : base(FilterType.CustomFilter)
        {
            this.SchemaGUID = schemaIdentifier;
            this.SchemaFilterData = schemaFilterData;
            this.QueryChangesFilterSchema = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.QueryChangesFilterSchemaSpecific, 16 + this.SchemaFilterData.Length);
        }

        /// <summary>
        /// Gets or sets Query Changes Filter Schema Specific (4 bytes): A 32-bit stream object header that specifies a query changes filter schema specific.
        /// </summary>
        public StreamObjectHeaderStart32bit QueryChangesFilterSchema { get; set; }

        /// <summary>
        /// Gets or sets Schema GUID (16 bytes): A GUID that specifies the schema specific filter opaque to this protocol.
        /// </summary>
        public Guid SchemaGUID { get; set; }

        /// <summary>
        /// Gets or sets Schema Filter Data (variable): A byte stream that specifies the schema filters data opaque to this protocol.
        /// </summary>
        public byte[] SchemaFilterData { get; set; }

        /// <summary>
        /// This method is used to convert the element into a byte List 
        /// </summary>
        /// <param name="byteList">Specify the byte list which will contain the filter data.</param>
        protected override void SerializeFilterData(List<byte> byteList)
        {
            byteList.AddRange(this.QueryChangesFilterSchema.SerializeToByteList());
            byteList.AddRange(this.SchemaGUID.ToByteArray());
            byteList.AddRange(this.SchemaFilterData);
        }
    }

    /// <summary>
    /// Specifies the data element identifiers to query.
    /// </summary>
    public class DataElementIDsFilter : Filter
    {
        /// <summary>
        /// Initializes a new instance of the DataElementIDsFilter class
        /// </summary>
        /// <param name="dataElementIDs">Specify the data element ID</param>
        public DataElementIDsFilter(ExGUIDArray dataElementIDs)
            : base(FilterType.DataElementIDsFilter)
        {
           this.DataElementIDs = new ExGUIDArray(dataElementIDs);
           this.QueryChangesFilterDataElementIDs = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.QueryChangesFilterDataElementIDs, this.DataElementIDs.SerializeToByteList().Count);
        }

        /// <summary>
        /// Gets or sets Query Changes Filter Data Element IDs (4 bytes): A 32-bit stream object header that specifies a query changes filter data element IDs.
        /// </summary>
        public StreamObjectHeaderStart32bit QueryChangesFilterDataElementIDs { get; set; }

        /// <summary>
        /// Gets or sets Data Element IDs (Variable): An extended GUID Array that specifies the data element identifiers.
        /// </summary>
        public ExGUIDArray DataElementIDs { get; set; }

        /// <summary>
        /// This method is used to convert the element into a byte List 
        /// </summary>
        /// <param name="byteList">Specify the byte list which will contain the filter data.</param>
        protected override void SerializeFilterData(List<byte> byteList)
        {
            byteList.AddRange(this.QueryChangesFilterDataElementIDs.SerializeToByteList());
            byteList.AddRange(this.DataElementIDs.SerializeToByteList());
        }
    }

    /// <summary>
    /// Specifies the hierarchy storage index keys to query as well as the depth to query.
    /// </summary>
    public class HierarchyFilter : Filter
    {
        /// <summary>
        /// Initializes a new instance of the HierarchyFilter class
        /// </summary>
        /// <param name="content">Specify the hierarchy filter contents.</param>
        public HierarchyFilter(byte[] content)
            : base(FilterType.HierarchyFilter)
        {
            this.QueryChangesFilterHierarchy = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.QueryChangesFilterHierarchy, 1 + content.Length);
            this.Contents = content;
        }

        /// <summary>
        /// Gets or sets Query Changes Filter Hierarchy (4 bytes): A 32-bit stream object header that specifies a query changes filter hierarchy.
        /// </summary>
        public StreamObjectHeaderStart32bit QueryChangesFilterHierarchy { get; set; }

        /// <summary>
        /// Gets or sets Depth (1 byte): An unsigned integer that specifies the depth and MUST be one of the following values:
        /// 0 Index values corresponding to the specified keys only.
        /// 1 First data elements referenced by the storage index values corresponding to the specified keys only.
        /// 2 Single level. All data elements under the sub-graphs rooted by the specified keys stopping at any storage index entries.
        /// 3 Deep. All data elements and storage index entries under the sub-graphs rooted by the specified keys.
        /// </summary>
        public HierarchyFilterDepth Depth { get; set; }

        /// <summary>
        /// Gets or sets the contents.
        /// </summary>
        public byte[] Contents { get; set; }

        /// <summary>
        /// This method is used to convert the element into a byte List 
        /// </summary>
        /// <param name="byteList">Specify the byte list which will contain the filter data.</param>
        protected override void SerializeFilterData(List<byte> byteList)
        {
            byteList.AddRange(this.QueryChangesFilterHierarchy.SerializeToByteList());
            byteList.Add((byte)this.Depth);
            byteList.AddRange(this.Contents);
        }
    }
}