namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    /// <summary>
    /// A class used to store the response data of "CopyIntoItems" operation.
    /// </summary>
    public class CopyIntoItemsResponse
    {
        /// <summary>
        /// Represents the result status of "CopyIntoItems" operation.
        /// </summary>
        private uint copyIntoItemsResultValue;

        /// <summary>
        /// Represents the results collection of copy actions.
        /// </summary>
        private CopyResult[] resultsCollection;

        /// <summary>
        /// Initializes a new instance of the CopyIntoItemsResponse class.
        /// </summary>
        /// <param name="copyIntoItemsResult">A parameter represents the result status of "CopyIntoItems" operation.</param>
        /// <param name="results">A parameter represents the copy results collection of the copy actions performed in "CopyIntoItems" operation.</param>
        public CopyIntoItemsResponse(uint copyIntoItemsResult, CopyResult[] results)
        {
            this.copyIntoItemsResultValue = copyIntoItemsResult;
            this.resultsCollection = results;
        }

        /// <summary>
        /// Gets the result status of the "CopyIntoItems" operation.
        /// </summary>
        public uint CopyIntoItemsResult
        {
            get
            {
                return this.copyIntoItemsResultValue;
            }
        }

        /// <summary>
        /// Gets the copy results collection.
        /// </summary>
        public CopyResult[] Results
        {
            get
            {
                return this.resultsCollection;
            }
        }
    }
}