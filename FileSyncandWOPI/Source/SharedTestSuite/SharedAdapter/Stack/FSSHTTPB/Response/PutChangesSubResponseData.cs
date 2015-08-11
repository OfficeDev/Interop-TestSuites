namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Specifies an optional response to remap serial numbers
    /// </summary>
    public class PutChangesSubResponseData : SubResponseData
    {
        /// <summary>
        /// Initializes a new instance of the PutChangesSubResponseData class. 
        /// </summary>
        public PutChangesSubResponseData()
        {
        }

        /// <summary>
        /// Gets or sets Put Changes Response Serial Number Reassign All (4 bytes), an optional 32-bit stream object header that specifies a put changes response serial number reassign all.
        /// </summary>
        public PutChangesResponseSerialNumberReassignAll PutChangesResponseSerialNumberReassignAll { get; set; }

        /// <summary>
        /// Gets or sets the PutChangesResponse.
        /// </summary>
        public PutChangesResponse PutChangesResponse { get; set; }

        /// <summary>
        /// Gets or sets Knowledge (variable): A knowledge that specifies the current state of the file on the server after the changes is merged.
        /// </summary>
        public Knowledge Knowledge { get; set; }

        /// <summary>
        /// Gets or sets the DiagnosticRequestOptionOutput.
        /// </summary>
        public DiagnosticRequestOptionOutput DiagnosticRequestOptionOutput { get; set; }

        /// <summary>
        /// Deserialize sub response data from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains sub response data.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        protected override void DeserializeSubResponseDataFromByteArray(byte[] byteArray, ref int currentIndex)
        {
            PutChangesResponseSerialNumberReassignAll outValue;
            int index = currentIndex;

            if (StreamObject.TryGetCurrent<PutChangesResponseSerialNumberReassignAll>(byteArray, ref index, out outValue))
            {
                this.PutChangesResponseSerialNumberReassignAll = outValue;
            }

            PutChangesResponse putChangesResponse;
            if (StreamObject.TryGetCurrent<PutChangesResponse>(byteArray, ref index, out putChangesResponse))
            {
                this.PutChangesResponse = putChangesResponse;
            }

            this.Knowledge = StreamObject.GetCurrent<Knowledge>(byteArray, ref index);

            DiagnosticRequestOptionOutput diagnosticRequestOptionOutput;
            if (StreamObject.TryGetCurrent<DiagnosticRequestOptionOutput>(byteArray, ref index, out diagnosticRequestOptionOutput))
            {
                this.DiagnosticRequestOptionOutput = diagnosticRequestOptionOutput;
            }

            currentIndex = index;
        }
    }

    /// <summary>
    /// specifies a put changes response serial number reassign to reassign serial number.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class PutChangesResponseSerialNumberReassign : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the PutChangesResponseSerialNumberReassign class. 
        /// </summary>
        public PutChangesResponseSerialNumberReassign()
            : base(StreamObjectTypeHeaderStart.PutChangesResponseSerialNumberReassign)
        {
        }

        /// <summary>
        /// Gets or sets Data Element Extended GUID (variable): A data element extended GUID that specifies to remap if the put changes response serial number reassign is specified.
        /// </summary>
        public ExGuid DataElementExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets Serial Number (variable): A serial number that specifies to use for the data element if the put changes response serial number reassign is specified.
        /// </summary>
        public SerialNumber SerialNumber { get; set; }

        /// <summary>
        /// Deserialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains response message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;

            this.DataElementExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.SerialNumber = BasicObject.Parse<SerialNumber>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "PutChangesResponseSerialNumberReassign", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            throw new System.NotImplementedException();
        }
    }

    /// <summary>
    /// Specifies a put changes response serial number reassign all.
    /// </summary>
    public class PutChangesResponseSerialNumberReassignAll : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the PutChangesResponseSerialNumberReassignAll class. 
        /// </summary>
        public PutChangesResponseSerialNumberReassignAll()
            : base(StreamObjectTypeHeaderStart.PutChangesResponseSerialNumberReassignAll)
        {
            this.PutChangesResponseSNReassignList = new List<PutChangesResponseSerialNumberReassign>();
        }

        /// <summary>
        /// Gets or sets Reassigned Serial Number (variable): A serial number that specifies to map all data elements to the same number if the put changes response serial number reassign all is specified.
        /// </summary>
        public SerialNumber ReassignedSerialNumber { get; set; }

        /// <summary>
        /// Gets or sets Put Changes Response Serial Number Reassign (4 bytes): Zero or more 32-bit stream object header that specifies a put changes response serial number reassign to reassign serial number to (overriding reassign all if it exists).
        /// </summary>
        public List<PutChangesResponseSerialNumberReassign> PutChangesResponseSNReassignList { get; set; }

        /// <summary>
        /// Deserialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains response message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;

            this.ReassignedSerialNumber = BasicObject.Parse<SerialNumber>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "PutChangesResponseSerialNumberReassignAll", "Stream object over-parse error", null);
            }

            this.PutChangesResponseSNReassignList = new List<PutChangesResponseSerialNumberReassign>();
            PutChangesResponseSerialNumberReassign outValue;

            while (StreamObject.TryGetCurrent<PutChangesResponseSerialNumberReassign>(byteArray, ref index, out outValue))
            {
                this.PutChangesResponseSNReassignList.Add(outValue);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            throw new NotImplementedException("The operation PutChangesResponseSerialNumberReassignAll::SerializeItemsToByteList is not implemented.");
        }
    }

    /// <summary>
    /// This class is used to represent the PutChangesResponse.
    /// </summary>
    public class PutChangesResponse : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the PutChangesResponse class. 
        /// </summary>
        public PutChangesResponse()
            : base(StreamObjectTypeHeaderStart.PutChangesResponse)
        {
            this.DataElementAdded = new ExGUIDArray();
        }

        /// <summary>
        /// Gets or sets the applied storage index guid.
        /// </summary>
        public ExGuid AppliedStorageIndexID { get; set; }

        /// <summary>
        /// Gets or sets the list of data elements.
        /// </summary>
        public ExGUIDArray DataElementAdded { get; set; }

        /// <summary>
        /// Deserialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains response message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.AppliedStorageIndexID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.DataElementAdded = BasicObject.Parse<ExGUIDArray>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "PutChangesResponse", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            throw new NotImplementedException("The operation PutChangesResponse::SerializeItemsToByteList is not implemented.");
        }
    }

    /// <summary>
    /// This class is used to represent the DiagnosticRequestOptionOutput.
    /// </summary>
    public class DiagnosticRequestOptionOutput : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the DiagnosticRequestOptionOutput class. 
        /// </summary>
        public DiagnosticRequestOptionOutput()
            : base(StreamObjectTypeHeaderStart.DiagnosticRequestOptionOutput)
        {
        }

        /// <summary>
        /// Gets or sets a value indicating whether IsDiagnosticRequestOptionOutput.
        /// </summary>
        public bool IsDiagnosticRequestOptionOutput { get; set; }

        /// <summary>
        /// Gets or sets the reserved value (7 bit).
        /// </summary>
        public int Reserved { get; set; }

        /// <summary>
        /// Deserialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains response message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            this.IsDiagnosticRequestOptionOutput = Convert.ToBoolean(byteArray[currentIndex] & 0x01);
            this.Reserved = byteArray[currentIndex] & 0x7f;

            if (1 != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "DiagnosticRequestOptionOutput", "Stream object over-parse error", null);
            }

            currentIndex += 1;
        }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            throw new NotImplementedException("The operation PutChangesResponse::SerializeItemsToByteList is not implemented.");
        }
    }
}