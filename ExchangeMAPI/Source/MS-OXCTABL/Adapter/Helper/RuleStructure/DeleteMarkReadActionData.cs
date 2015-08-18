namespace Microsoft.Protocols.TestSuites.MS_OXCTABL
{
    /// <summary>
    /// Action Data buffer format for ActionType: OP_DELETE or OP_MARK_AS_READ
    /// The incoming messages are deleted or marked as read according to the ActionType itself. These actions have no ActionData buffer.
    /// </summary>
    public class DeleteMarkReadActionData : IActionData
    {
        /// <summary>
        /// The total Size of this ActionData buffer
        /// </summary>
        /// <returns>Number of bytes in this ActionData buffer</returns>
        public int Size()
        {
            // Length of BounceCode
            return 0;
        }

        /// <summary>
        /// Get serialized byte array for this ActionData
        /// </summary>
        /// <returns>Serialized byte array</returns>
        public byte[] Serialize()
        {
            return null;
        }

        /// <summary>
        /// Deserialized byte array to a BounceActionData instance
        /// </summary>
        /// <param name="buffer">Byte array contains data of an ActionData instance</param>
        /// <returns>Bytes count that deserialized in buffer</returns>
        public uint Deserialize(byte[] buffer)
        {
            return 0;
        }
    }
}