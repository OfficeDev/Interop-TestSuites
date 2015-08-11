namespace Microsoft.Protocols.TestSuites.MS_COPYS
{   
    /// <summary>
    /// A class used to store the response data of "CopyIntoItemsLocal" operation.
    /// </summary>
    public class CopyIntoItemsLocalResponse
    {   
        /// <summary>
        /// Represents the result status of "CopyIntoItemsLocal" operation.
        /// </summary>
        private uint copyIntoItemsLocalResultValue;

        /// <summary>
        /// Represents the results collection of copy actions.
        /// </summary>
        private CopyResult[] resultsCollection;

        /// <summary>
        /// Initializes a new instance of the CopyIntoItemsLocalResponse class.
        /// </summary>
        /// <param name="copyIntoItemsLocalResult">A parameter represents the result status of "CopyIntoItemsLocal" operation.</param>
        /// <param name="results">A parameter represents the copy results collection of the copy actions performed in "CopyIntoItemsLocal" operation.</param>
        public CopyIntoItemsLocalResponse(uint copyIntoItemsLocalResult, CopyResult[] results)
        {
            this.copyIntoItemsLocalResultValue = copyIntoItemsLocalResult;
            this.resultsCollection = results;
        }

        /// <summary>
        /// Gets the result status of the "CopyIntoItemsLocal" operation.
        /// </summary>
        public uint CopyIntoItemsLocalResult
        {
            get
            {
                return this.copyIntoItemsLocalResultValue;
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