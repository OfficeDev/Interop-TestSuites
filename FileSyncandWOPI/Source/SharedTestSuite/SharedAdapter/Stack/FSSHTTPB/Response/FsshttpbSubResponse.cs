namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Specifies the sub-responses corresponding to each sub-request
    /// </summary>
    public class FsshttpbSubResponse : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the FsshttpbSubResponse class.
        /// </summary>
        public FsshttpbSubResponse()
            : base(StreamObjectTypeHeaderStart.FsshttpbSubResponse)
        {
            this.RequestID = new Compact64bitInt();
            this.RequestType = new Compact64bitInt();
        }

        /// <summary>
        /// Gets or sets Request ID (variable): A compact unsigned 64-bit integer specifying the request number this sub-response is for.
        /// </summary>
        public Compact64bitInt RequestID { get; set; }

        /// <summary>
        /// Gets or sets Request Type (variable): A compact unsigned 64-bit integer specifying the response type matching the request.
        /// </summary>
        public Compact64bitInt RequestType { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the request has failed and a response error MUST follow.
        /// </summary>
        public bool Status { get; set; }

        /// <summary>
        /// Gets or sets Reserved (7 bits): A 7 bit reserved and MUST be set to 0 and MUST be ignored.
        /// </summary>
        public byte Reserved { get; set; }

        /// <summary>
        /// Gets or sets the response error.
        /// </summary>
        public ResponseError ResponseError { get; set; }

        /// <summary>
        /// Gets or sets the sub response data.
        /// </summary>
        public SubResponseData SubResponseData { get; set; }

        /// <summary>
        /// Get the sub response data with the specified type.
        /// </summary>
        /// <typeparam name="T">Specify the sub response data type.</typeparam>
        /// <returns>Return the sub response data if the sub response data is the type T.</returns>
        /// <exception cref="InvalidOperationException">If the sub response data is not type T, throw InvalidOperationException.</exception>
        public T GetSubResponseData<T>()
            where T : SubResponseData
        {
            if (this.SubResponseData is T)
            {
                return this.SubResponseData as T;
            }

            throw new InvalidOperationException(string.Format("The sub response data is not the type {0}, its type is {1}", typeof(T).Name, this.SubResponseData.GetType().Name));
        }

        /// <summary>
        /// Deserialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains response message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.RequestID = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.RequestType = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.Status = (byteArray[index] & 0x1) == 0x1 ? true : false;
            this.Reserved = (byte)(byteArray[index] >> 1);
            index += 1;

            if (index - currentIndex != lengthOfItems)
            {
                throw new ResponseParseErrorException(currentIndex, "CellSubResponse object over-parse error", null); 
            }

            if (this.Status)
            {
                this.ResponseError = StreamObject.GetCurrent<ResponseError>(byteArray, ref index);
            }
            else
            {
                this.SubResponseData = SubResponseData.GetCurrentSubResponseData((int)this.RequestType.DecodedValue, byteArray, ref index);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // Capture the requirements related response data.
                    new MsfsshttpbAdapterCapture().InvokeCaptureMethod(this.SubResponseData.GetType(), this.SubResponseData, SharedContext.Current.Site);
                }
            }

            currentIndex = index;
        }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected override int SerializeItemsToByteList(System.Collections.Generic.List<byte> byteList)
        {
            throw new NotImplementedException("The FsshttpbSubResponse::SerializeItemsToByteList does not implement in the current state.");
        }
    }
}