namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Base class of Restriction
    /// </summary>
    public abstract class Restriction
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value indicates the Type of restriction.
        /// </summary>
        private Restrictions restrictType;

        /// <summary>
        /// Gets or sets unsigned 8-bit integer. This value indicates the Type of restriction.
        /// </summary>
        public Restrictions RestrictType
        {
            get { return this.restrictType; }
            protected set { this.restrictType = value; }
        }

        /// <summary>
        /// Get the total Size of Restriction
        /// </summary>
        /// <returns>The Size of Restriction buffer</returns>
        public abstract int Size();

        /// <summary>
        /// Get serialized byte array for this struct
        /// </summary>
        /// <returns>Serialized byte array</returns>
        public abstract byte[] Serialize();

        /// <summary>
        /// Deserialized byte array to an Restriction instance
        /// </summary>
        /// <param name="restrictionData">Byte array contain data of an Restriction instance</param>
        public abstract void Deserialize(byte[] restrictionData);
    }
}