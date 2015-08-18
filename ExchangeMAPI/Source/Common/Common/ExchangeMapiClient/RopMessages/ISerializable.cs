namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// This interface define the methods that is needed to serialize an ROP object
    /// </summary>
    public interface ISerializable
    {
        /// <summary>
        /// Serialize into a bytes array.
        /// </summary>
        /// <returns>The bytes array serialized</returns>
        byte[] Serialize();

        /// <summary>
        /// Return the size in bytes of the object serialized
        /// </summary>
        /// <returns>The size in bytes of the object serialized</returns>
        int Size();
    }
}