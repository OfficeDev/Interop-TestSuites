namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class specifies a base class for all the cell sub request.
    /// </summary>
    public abstract class FsshttpbCellSubRequest : IFSSHTTPBSerializable
    {
        /// <summary>
        /// Initializes a new instance of the FsshttpbCellSubRequest class
        /// </summary>
        protected FsshttpbCellSubRequest()
        {    
        }

        /// <summary>
        /// Gets or sets Request ID (variable): A compact unsigned 64-bit integer specifying the request number for each sub-request.
        /// </summary>
        public ulong RequestID { get; set; }
        
        /// <summary>
        /// Gets or sets Request Type (variable): A compact unsigned 64-bit integer that specifies the request type.
        /// </summary>
        public ulong RequestType { get; set; }
        
        /// <summary>
        /// Gets or sets Priority (variable): A compact unsigned 64-bit integer that specify the priority of the sub-request. (variable): A compact unsigned 64-bit integer that specify the priority of the sub-request.
        /// </summary>
        public ulong Priority { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the partition id guid is used.
        /// </summary>
        public bool IsPartitionIDGUIDUsed { get; set; }

        /// <summary>
        /// Gets or sets a GUID that specifies the partition.
        /// If the IsPartitionIDGUIDUsed is false, this property is ignored.
        /// </summary>
        public Guid PartitionIdGUID { get; set; }

        /// <summary>
        /// Gets or sets Sub-request End (2 bytes): A 16-bit stream object header that specifies a sub-request end.
        /// </summary>
        internal StreamObjectHeaderEnd16bit SubRequestEnd { get; set; }

        /// <summary>
        /// Gets or sets Target Partition Id Start (variable): A 32-bit stream object header that specifies the beginning of a target partition id.
        /// </summary>
        internal StreamObjectHeaderStart PartitionIdGUIDStart { get; set; }

        /// <summary>
        /// Gets or sets Sub-request Start (4 bytes): A 32-bit stream object header that specifies a sub-request start.
        /// </summary>
        internal StreamObjectHeaderStart32bit SubRequestStart { get; set; }

        /// <summary>
        /// This method is used to convert the element into a byte List 
        /// </summary>
        /// <returns>Return the Byte List</returns>
        public virtual List<byte> SerializeToByteList()
        {
            // Request ID bytes
            List<byte> requestIDBytes = (new Compact64bitInt(this.RequestID)).SerializeToByteList();
            
            // Request Type bytes
            List<byte> requestTypeBytes = (new Compact64bitInt(this.RequestType)).SerializeToByteList();
            
            // Priority bytes
            List<byte> priorityBytes = (new Compact64bitInt(this.Priority)).SerializeToByteList();

            // Sub-request Start bytes
            int subRequstStartLength = requestIDBytes.Count + requestTypeBytes.Count + priorityBytes.Count;
            this.SubRequestStart = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.SubRequest, subRequstStartLength);

            List<byte> byteList = new List<byte>();
            
            // Sub-request Start
            byteList.AddRange(this.SubRequestStart.SerializeToByteList());
            
            // Request ID
            byteList.AddRange(requestIDBytes);
            
            // Request Type
            byteList.AddRange(requestTypeBytes);
            
            // Priority
            byteList.AddRange(priorityBytes);

            // Target partition ID
            if (this.IsPartitionIDGUIDUsed)
            {
                this.PartitionIdGUIDStart = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.TargetPartitionId, 16);
                byteList.AddRange(this.PartitionIdGUIDStart.SerializeToByteList());
                byteList.AddRange(this.PartitionIdGUID.ToByteArray());
            
            }

            return byteList;
        }

        /// <summary>
        /// This method is used to convert the element end  into a byte List 
        /// </summary>
        /// <returns>Return the byte List</returns>
        public List<byte> ToBytesEnd()
        {
            this.SubRequestEnd = new StreamObjectHeaderEnd16bit(StreamObjectTypeHeaderEnd.SubRequest); 
            return this.SubRequestEnd.SerializeToByteList();
        }
    }
}