namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class specifies query access sub-request.
    /// </summary>
    public class QueryAccessCellSubRequest : FsshttpbCellSubRequest 
    {
        /// <summary>
        /// Initializes a new instance of the QueryAccessCellSubRequest class
        /// </summary>
        /// <param name="subRequestID">Specify the sub request id</param>
        public QueryAccessCellSubRequest(ulong subRequestID)
        {
            this.RequestID = subRequestID;
            this.RequestType = Convert.ToUInt64(RequestTypes.QueryAccess);
        }

        /// <summary>
        /// This method is used to convert the element into a byte List 
        /// </summary>
        /// <returns>Return the Byte List</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(base.SerializeToByteList());
            byteList.AddRange(this.ToBytesEnd());

            return byteList;
        }
    }
}