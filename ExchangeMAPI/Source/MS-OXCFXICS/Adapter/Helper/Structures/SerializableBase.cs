namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.IO;

    /// <summary>
    /// Base serializable class.
    /// </summary>
    public abstract class SerializableBase : IStructSerializable, IStructDeserializable
    {
        /// <summary>
        /// The size of a GUID structure in bytes.
        /// </summary>
        public readonly int GuidSize = Guid.Empty.ToByteArray().Length;

        /// <summary>
        /// Serialize current instance to a stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>The number of bytes written to the stream.</returns>
        public virtual int Serialize(Stream stream)
        {
            return Serializer.Serialize(this, stream);
        }

        /// <summary>
        /// Deserialize an object from a stream.
        /// </summary>
        /// <param name="stream">A stream contains object fields.</param>
        /// <param name="size">Max length can used by this deserialization
        /// if -1 no limitation except stream length.
        /// </param>
        /// <returns>The number of bytes read from the stream.</returns>
        public abstract int Deserialize(Stream stream, int size);
    }
}