namespace Microsoft.Protocols.TestSuites.Common
{
    using System.Collections.Generic;

    /// <summary>
    /// Range structure.
    /// </summary>
    public struct PartMetaData
    {
        /// <summary>
        /// The Start index of a part
        /// </summary>
        private int start;

        /// <summary>
        /// The Count in byte of a part
        /// </summary>
        private int count;

        /// <summary>
        /// Gets or sets Start index of a part.
        /// </summary>
        public int Start
        {
            get { return this.start; }

            set { this.start = value; }
        }

        /// <summary>
        /// Gets or sets Count in byte of a part.
        /// </summary>
        public int Count
        {
            get { return this.count; }

            set { this.count = value; }
        }
    }

    /// <summary>
    /// Represents a Metadata for a multipart response.
    /// </summary>
    public class MultipartMetadata
    {
        /// <summary>
        /// The count of parts
        /// </summary>
        private int partsCount;

        /// <summary>
        /// The ranges of parts
        /// </summary>
        private PartMetaData[] partsMetaData;

        /// <summary>
        /// Initializes a new instance of the MultipartMetadata class.
        /// </summary>
        /// <param name="metadata">An integer array contains the metadata information</param>
        public MultipartMetadata(int[] metadata)
        {
            this.partsCount = metadata[0];

            List<PartMetaData> temp = new List<PartMetaData>();
            for (int i = 1; i < this.partsCount * 2; i = i + 2)
            {
                temp.Add(new PartMetaData() { Start = metadata[i], Count = metadata[i + 1] });
            }

            this.partsMetaData = temp.ToArray();
        }

        /// <summary>
        /// Gets the count of parts.
        /// </summary>
        public int PartsCount
        {
            get
            {
                return this.partsCount;
            }
        }

        /// <summary>
        /// Gets the ranges of the multipart.
        /// </summary>
        public PartMetaData[] PartsMetaData
        {
            get
            {
                return this.partsMetaData;
            }
        }
    }
}