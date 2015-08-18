namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// The progressTotal element contains data that describes 
    /// the approximate size of all the messageChange elements 
    /// that will follow in this stream. 
    /// progressTotal        = IncrSyncProgressMode propList
    /// </summary>
    public class ProgressTotal : SyntacticalBase
    {
        /// <summary>
        /// A propList value.
        /// </summary>
        private PropList propList;

        /// <summary>
        /// Initializes a new instance of the ProgressTotal class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public ProgressTotal(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the propList.
        /// </summary>
        public PropList PropList
        {
            get { return this.propList; }
            set { this.propList = value; }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized progressTotal
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized progressTotal, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return stream.VerifyMarker(Markers.PidTagIncrSyncProgressMode);
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            this.Deserialize<PropList>(
                stream, 
                Markers.PidTagIncrSyncProgressMode,
                out this.propList);
        }
    }
}