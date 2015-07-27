//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

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