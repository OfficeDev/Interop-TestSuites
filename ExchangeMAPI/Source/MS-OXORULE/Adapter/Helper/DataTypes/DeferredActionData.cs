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