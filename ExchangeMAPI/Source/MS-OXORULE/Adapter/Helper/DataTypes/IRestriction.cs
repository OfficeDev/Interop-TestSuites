namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    /// <summary>
    /// Interface of Restrictions
    /// </summary>
    public interface IRestriction
    {
        /// <summary>
        /// Gets unsigned 8-bit integer. This value indicates the Type of restriction.
        /// </summary>
        RestrictionType RestrictType
        {
            get;
        }

        /// <summary>
        /// Get the total Size of Restriction
        /// </summary>
        /// <returns>The Size of Restriction buffer.</returns>
        int Size();

        /// <summary>
        /// Get serialized byte array for this struct
        /// </summary>
        /// <returns>Serialized byte array.</returns>
        byte[] Serialize();

        /// <summary>
        /// Deserialized byte array to a Restriction instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of a Restriction instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        uint Deserialize(byte[] buffer);
    }
}