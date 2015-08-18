namespace Microsoft.Protocols.TestSuites.MS_OXCTABL
{
    /// <summary>
    /// Interface for ActionData, when using ActionData, must use a derived class base on different Action Type
    /// </summary>
    public interface IActionData
    {
        /// <summary>
        /// Get the total Size of ActionData
        /// </summary>
        /// <returns>The Size of ActionData buffer</returns>
        int Size();

        /// <summary>
        /// Get serialized byte array for this structure
        /// </summary>
        /// <returns>Serialized byte array</returns>
        byte[] Serialize();

        /// <summary>
        /// Deserialized byte array to an ActionData instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of an ActionData instance</param>
        /// <returns>Bytes count that deserialized in buffer</returns>
        uint Deserialize(byte[] buffer);
    }
}