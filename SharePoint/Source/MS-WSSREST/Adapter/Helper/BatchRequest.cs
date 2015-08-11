namespace Microsoft.Protocols.TestSuites.MS_WSSREST
{
    /// <summary>
    /// The batch request.
    /// </summary>
    public class BatchRequest : Request
    {
        /// <summary>
        /// The type of this request.
        /// </summary>
        private OperationType operationType;

        /// <summary>
        /// Gets or sets the type of this request.
        /// </summary>
        public OperationType OperationType
        {
            get
            {
                return this.operationType;
            }

            set
            {
                this.operationType = value;
            }
        }
    }
}