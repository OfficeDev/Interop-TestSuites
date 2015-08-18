namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// The messageChange element represents a change to a Message object.
    /// messageChange        = messageChangeFull / MessageChangePartial
    /// </summary>
    public abstract class MessageChange : SyntacticalBase
    {
        /// <summary>
        /// Initializes a new instance of the MessageChange class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        protected MessageChange(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets a value indicating whether has PidTagMid.
        /// </summary>
        public abstract bool HasPidTagMid { get; }

        /// <summary>
        /// Gets a value indicating whether has PidTagMessageSize.
        /// </summary>
        public abstract bool HasPidTagMessageSize { get; }

        /// <summary>
        /// Gets a value indicating whether has PidTagChangeNumber.
        /// </summary>
        public abstract bool HasPidTagChangeNumber { get; }

        /// <summary>
        /// Gets a value indicating the sourceKey property.
        /// </summary>
        public abstract byte[] SourceKey { get; }

        /// <summary>
        /// Gets a value indicating the PidTagChangekey property.
        /// </summary>
        public abstract byte[] PidTagChangeKey { get; }

        /// <summary>
        /// Gets a value indicating the PidTagChangeNumber property.
        /// </summary>
        public abstract long PidTagChangeNumber { get; }

        /// <summary>
        /// Gets a value indicating the PidTagMid property.
        /// </summary>
        public abstract long PidTagMid { get; }

        /// <summary>
        /// Gets LastModificationTime.
        /// </summary>
        public abstract DateTime LastModificationTime { get; }

        /// <summary>
        /// Verify that a stream's current position contains a serialized messageChange.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized messageChange, return true, else false</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return MessageChangeFull.Verify(stream)
                || MessageChangePartial.Verify(stream);
        }

        /// <summary>
        /// Deserialize a messageChange from a stream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>A messageChange object.</returns>
        public static SyntacticalBase DeserializeFrom(FastTransferStream stream)
        {
            if (MessageChangeFull.Verify(stream))
            {
                return new MessageChangeFull(stream);
            }
            else if (MessageChangePartial.Verify(stream))
            {
                return new MessageChangePartial(stream);
            }
            else
            {
                AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
                return null;
            }
        }
    }
}