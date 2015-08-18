namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// Definition of ICS data structure.
    /// </summary>
    public struct ICSStateData
    {
        /// <summary>
        /// PidTagIdsetGiven data.
        /// </summary>
        public byte[] PidTagIdsetGiven;

        /// <summary>
        /// PidTagCnsetSeen data.
        /// </summary>
        public byte[] PidTagCnsetSeen;

        /// <summary>
        /// PidTagCnsetSeenFAI data.
        /// </summary>
        public byte[] PidTagCnsetSeenFAI;

        /// <summary>
        /// PidTagCnsetRead data.
        /// </summary>
        public byte[] PidTagCnsetRead;
    }
}