namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using Microsoft.Protocols.TestSuites.Common;
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class specifies Allocate extended Guid range sub-request 
    /// </summary>
    public class AllocateExtendedGuidRangeCellSubRequest : FsshttpbCellSubRequest
    {
        /// <summary>
        /// Initializes a new instance of the AllocateExtendedGuidRangeCellSubRequest class
        /// </summary>
        /// <param name="requestIDCount">A compact unsigned 64-bit integer that specifies the number of extended Guids to allocate.</param>
        /// <param name="subRequestID">Specify the sub-request id</param>
        public AllocateExtendedGuidRangeCellSubRequest(Compact64bitInt requestIDCount, ulong subRequestID)
        {
            this.RequestID = subRequestID;
            this.RequestType = Convert.ToUInt64(RequestTypes.AllocateExtendedGuidRange);
            this.AllocateExtendedGuidRangeRequest = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.AllocateExtendedGUIDRangeRequest, requestIDCount.SerializeToByteList().Count + 1);
            this.RequestIDCount = requestIDCount;
        }

        /// <summary>
        /// Gets or sets Put Raw Storage Request (4 bytes): A stream object header that specifies a Allocate extended Guid range request.
        /// </summary>
        public StreamObjectHeaderStart32bit AllocateExtendedGuidRangeRequest { get; set; }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the number of extended Guids to allocate.
        /// </summary>
        public Compact64bitInt RequestIDCount { get; set; }

        /// <summary>
        /// Gets or sets an 8-bit reserved field that MUST be set to zero and MUST be ignored.
        /// </summary>
        public int Reserved { get; set; }

        /// <summary>
        /// This method is used to convert the element into a byte List 
        /// </summary>
        /// <returns>Return the Byte List</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(base.SerializeToByteList());
            byteList.AddRange(this.AllocateExtendedGuidRangeRequest.SerializeToByteList());
            byteList.AddRange(this.RequestIDCount.SerializeToByteList());
            BitWriter bitWriter = new BitWriter(1);
            bitWriter.AppendInit32(this.Reserved, 8);
            List<byte> reservedBytes = new List<byte>(bitWriter.Bytes);
            byteList.AddRange(reservedBytes);
            byteList.AddRange(this.ToBytesEnd());

            return byteList;
        }
    }
}