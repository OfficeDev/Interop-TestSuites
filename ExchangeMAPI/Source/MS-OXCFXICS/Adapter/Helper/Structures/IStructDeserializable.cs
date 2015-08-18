namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System.IO;

    /// <summary>
    /// An interface that a deserializable object must implement.
    /// </summary>
    public interface IStructDeserializable
    {
        /// <summary>
        /// Deserialize an object from a stream.
        /// </summary>
        /// <param name="stream">A stream contains object fields.</param>
        /// <param name="size">Max length can used by this deserialization
        /// if -1 no limitation except stream length.
        /// </param>
        /// <returns>The number of bytes read from the stream.</returns>
        int Deserialize(Stream stream, int size);
    }
}