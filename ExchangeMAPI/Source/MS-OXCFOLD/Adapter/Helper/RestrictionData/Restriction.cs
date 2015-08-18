namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    /// <summary>
    /// The common restricts type structure.
    /// </summary>
    public abstract class Restriction
    {
        /// <summary>
        /// Gets or sets value indicates the type of restriction (2) and MUST be set to 0x02.
        /// </summary>
        public RestrictType RestrictType
        {
            get;
            set;
        }

        /// <summary>
        /// Deserialize the Restriction data.
        /// </summary>
        /// <param name="restrictionData">The restriction data.</param>
        public abstract void Deserialize(byte[] restrictionData);

        /// <summary>
        /// Serialize the Restriction data.
        /// </summary>
        /// <returns>Format the type of Restriction data to byte array.</returns>
        public abstract byte[] Serialize();

        /// <summary>
        /// Get the size of the restriction data.
        /// </summary>
        /// <returns>Returns the size of restriction data.</returns>
        public abstract int Size();
    }
}