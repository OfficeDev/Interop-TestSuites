namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This class specifies the base class for file chunking.
    /// </summary>
    public abstract class AbstractChunking
    {
        /// <summary>
        /// The file content that contains the data.
        /// </summary>
        private byte[] fileContent;

        /// <summary>
        /// Initializes a new instance of the AbstractChunking class.
        /// </summary>
        /// <param name="fileContent">The content of the file.</param>
        protected AbstractChunking(byte[] fileContent)
        {
            this.fileContent = fileContent;
        }

        /// <summary>
        /// Gets or sets the file content.
        /// </summary>
        protected byte[] FileContent
        {
            get
            {
                return this.fileContent;
            }

            set
            {
                this.fileContent = value;
            }
        }

        /// <summary>
        /// This method is used to chunk the file data.
        /// </summary>
        /// <returns>A list of LeafNodeObjectData.</returns>
        public abstract List<LeafNodeObjectData> Chunking();

        /// <summary>
        /// This method is used to analyze the chunk.
        /// </summary>
        /// <param name="rootNode">Specify the root node object which is needed to be analyzed.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public abstract void AnalyzeChunking(IntermediateNodeObject rootNode, ITestSite site);
    }
}