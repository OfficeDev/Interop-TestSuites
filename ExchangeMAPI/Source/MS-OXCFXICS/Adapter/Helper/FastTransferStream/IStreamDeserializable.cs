namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// An interface deserializable object must implement.
    /// </summary>
    public interface IStreamDeserializable
    {
        /// <summary>
        /// Deserialize object from memoryStream,
        /// after deserialization stream's read position += serialized object size;
        /// </summary>
        /// <param name="stream">Stream contains the serialized object</param>
        void Deserialize(FastTransferStream stream);
    }
}