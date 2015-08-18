namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System.IO;

    /// <summary>
    /// An interface that every serializable object must implement.
    /// </summary>
    public interface IStructSerializable
    {
        /// <summary>
        /// Serialize current instance to a stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>The number of bytes written to the stream.</returns>
        int Serialize(Stream stream);
    }
}