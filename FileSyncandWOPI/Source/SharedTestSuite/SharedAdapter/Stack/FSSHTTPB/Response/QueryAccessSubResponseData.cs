namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    /// <summary>
    /// Specifies the access permissions requested for read/write to the file
    /// </summary>
    public class QueryAccessSubResponseData : SubResponseData
    {
        /// <summary>
        /// Initializes a new instance of the QueryAccessSubResponseData class. 
        /// </summary>
        public QueryAccessSubResponseData()
        {
        }

        /// <summary>
        /// Gets or sets Read Access Response Start (4 bytes): A 32-bit stream object header that specifies a read access response start.
        /// </summary>
        public ReadAccessResponse ReadAccessResponse { get; set; }

        /// <summary>
        /// Gets or sets Write Access Response Start (4 bytes): A 32-bit stream object header that specifies a write access response start.
        /// </summary>
        public WriteAccessResponse WriteAccessResponse { get; set; }

        /// <summary>
        /// Deserialize sub response data from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains sub response data.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        protected override void DeserializeSubResponseDataFromByteArray(byte[] byteArray, ref int currentIndex)
        {
            int index = currentIndex;

            this.ReadAccessResponse = StreamObject.GetCurrent<ReadAccessResponse>(byteArray, ref index);
            this.WriteAccessResponse = StreamObject.GetCurrent<WriteAccessResponse>(byteArray, ref index);

            currentIndex = index;
        }
    }

    /// <summary>
    /// Specifies the read access response.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class ReadAccessResponse : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the ReadAccessResponse class. 
        /// </summary>
        public ReadAccessResponse()
            : base(StreamObjectTypeHeaderStart.ReadAccessResponse)
        {
        }

        /// <summary>
        /// Gets or sets Response Error (variable): A response error that specifies read access permission.
        /// </summary> 
        public ResponseError ReadResponseError { get; set; }

        /// <summary>
        /// Deserialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains response message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 0)
            {
                throw new ResponseParseErrorException(currentIndex, "ReadAccessResponse over-parse error", null);
            }

            int index = currentIndex;
            this.ReadResponseError = StreamObject.GetCurrent<ResponseError>(byteArray, ref index);
            currentIndex = index;
        }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected override int SerializeItemsToByteList(System.Collections.Generic.List<byte> byteList)
        {
            throw new System.NotImplementedException();
        }
    }

    /// <summary>
    /// Specifies the write access response.
    /// </summary>
    public class WriteAccessResponse : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the WriteAccessResponse class. 
        /// </summary>
        public WriteAccessResponse() : base(StreamObjectTypeHeaderStart.WriteAccessResponse)
        {
        }

        /// <summary>
        /// Gets or sets Response Error (variable): A response error that specifies write access permission.
        /// </summary>
        public ResponseError WriteResponseError { get; set; }

        /// <summary>
        /// Deserialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains response message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 0)
            {
                throw new ResponseParseErrorException(currentIndex, "WriteAccessResponse over-parse error", null);
            }

            int index = currentIndex;
            this.WriteResponseError = StreamObject.GetCurrent<ResponseError>(byteArray, ref index);
            currentIndex = index;
        }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected override int SerializeItemsToByteList(System.Collections.Generic.List<byte> byteList)
        {
            throw new System.NotImplementedException();
        }
    }
}