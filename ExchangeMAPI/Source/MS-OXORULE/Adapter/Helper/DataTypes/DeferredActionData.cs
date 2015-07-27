//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    /// <summary>
    /// Action Data buffer format for ActionType: OP_DEFER_ACTION
    /// </summary>
    public class DeferredActionData : IActionData
    {
        /// <summary>
        /// Client defined Data, will be treated as an opaque BLOB(binary large object) by the server
        /// </summary>
        private byte[] data;

        /// <summary>
        /// Gets or sets the data
        /// </summary>
        public byte[] Data
        {
            get { return this.data; }
            set { this.data = value; }
        }

        /// <summary>
        /// The total Size of this ActionData buffer
        /// </summary>
        /// <returns>Number of bytes in this ActionData buffer.</returns>
        public int Size()
        {
            return this.Serialize().Length;
        }

        /// <summary>
        /// Get serialized byte array for this ActionData
        /// </summary>
        /// <returns>Serialized byte array.</returns>
        public byte[] Serialize()
        {
            return this.Data;
        }

        /// <summary>
        /// Deserialized byte array to a DeferredActionData instance
        /// </summary>
        /// <param name="buffer">Byte array contains data of an ActionData instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public uint Deserialize(byte[] buffer)
        {
            this.Data = buffer;
            return (uint)buffer.Length;
        }
    }
}