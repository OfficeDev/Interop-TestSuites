namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// This interface define the methods that is needed to deserialize a bytes array into an ROP object
    /// </summary>
    public interface IDeserializable
    {
        /// <summary>
        /// Deserialize input bytes ropBytes into a ROP object
        /// </summary>
        /// <param name="ropBytes">The bytes array to deserialize</param>
        /// <param name="startIndex">The start index of the byte array</param>
        /// <returns>The bytes deserialized</returns>
        int Deserialize(byte[] ropBytes, int startIndex);
    }
}