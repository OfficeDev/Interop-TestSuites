namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Interface of Restriction
    /// </summary>
    public interface IRestriction
    {
        /// <summary>
        /// Gets unsigned 8-bit integer. This value indicates the Type of restriction.
        /// </summary>
        Restrictions RestrictType
        {
            get;
        }

        /// <summary>
        /// Get the total Size of Restriction
        /// </summary>
        /// <returns>The Size of Restriction buffer</returns>
        int Size();

        /// <summary>
        /// Get serialized byte array for this struct
        /// </summary>
        /// <returns>Serialized byte array</returns>
        byte[] Serialize();

        /// <summary>
        /// Deserialized byte array to an Restriction instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of an Restriction instance</param>
        void Deserialize(byte[] buffer);
    }
}