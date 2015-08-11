namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    /// <summary>
    /// The result of custom processing of a legal hold.
    /// </summary>
    public class CustomProcessingResult
    {
        /// <summary>
        /// This property indicates the result of custom processing of a legal hold.
        /// </summary>
        private HoldProcessingResult? holdsProcessingResult;

        /// <summary>
        /// Gets or sets the result of custom processing of a legal hold.
        /// </summary>
        public HoldProcessingResult? HoldProcessingResult
        {
            get
            {
                return this.holdsProcessingResult;
            }

            set
            {
                this.holdsProcessingResult = value;
            }
        }
    }
}