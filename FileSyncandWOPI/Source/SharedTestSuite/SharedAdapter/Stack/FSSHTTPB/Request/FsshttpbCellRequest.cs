namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// This class specifies a kind of request.
    /// </summary>
    public class FsshttpbCellRequest : IFSSHTTPBSerializable
    {
        /// <summary>
        /// This user agent guid is used by the test suite.
        /// </summary>
        public static readonly Guid UserAgentGuid = new Guid("E731B87E-DD45-44AA-AB80-0C75FBD1530E");

        /// <summary>
        /// Initializes a new instance of the FsshttpbCellRequest class
        /// </summary>
        public FsshttpbCellRequest()
        {
            this.IsRequestHashingOptionsUsed = false;
        }

        /// <summary>
        /// Gets or sets Protocol Version (2bytes): An unsigned integer that specifies the protocol schema version number used in this request. This value MUST be 12.
        /// </summary>
        public ushort ProtocolVersion { get; set; }

        /// <summary>
        /// Gets or sets Minimum Version (2 bytes): An unsigned integer that specifies the oldest version of the protocol schema that this schema is compatible with. This value MUST be 11.
        /// </summary>
        public ushort MinimumVersion { get; set; }

        /// <summary>
        /// Gets or sets Signature (8 bytes): An unsigned integer that specifies a constant signature, to identify this as a request. This MUST be set to 0x9B069439F329CF9C.
        /// </summary>
        public ulong Signature { get; set; }

        /// <summary>
        /// Gets or sets GUID (16 bytes): A GUID that specifies the user agent.
        /// </summary>
        public System.Guid GUID { get; set; }

        /// <summary>
        /// Gets or sets Version (4 bytes): An unsigned integer that specifies the version of the client.
        /// </summary>
        public uint Version { get; set; }

        /// <summary>
        /// Gets or sets Request Hashing Schema: A compact unsigned 64-bit integer that specifies the Hashing Schema being requested that must be set to 1 indicating Content Information Data Structure Version 1.0 as specified in [MS-PCCRC].
        /// If the IsRequestHashingOptionsUsed is false, this property will be ignored.
        /// </summary>
        public Compact64bitInt RequestHashingSchema { get; set; }

        /// <summary>
        /// Gets or sets Reserved (1 bit): A reserved bit that MUST be set to zero and MUST be ignored.
        /// If the IsRequestHashingOptionsUsed is false, this property will be ignored.
        /// </summary>
        public int Reserve1 { get; set; }

        /// <summary>
        /// Gets or sets Reserved (1 bit): A reserved bit that MUST be set to zero and MUST be ignored.
        /// If the IsRequestHashingOptionsUsed is false, this property will be ignored.
        /// </summary>
        public int Reserve2 { get; set; }

        /// <summary>
        /// Gets or sets Request Data Element Hashes Instead of Data (1 bit): If set, a bit that specifies to exclude object data and instead return data element hashes; otherwise, object data is included.
        /// If the IsRequestHashingOptionsUsed is false, this property will be ignored.
        /// </summary>
        public int RequestDataElementHashesInsteadofData { get; set; }

        /// <summary>
        /// Gets or sets Request Data Element Hashes (1 bit): If set, a bit that specifies to include data element hashes (if available) when returning data elements; otherwise data element hashes should not be returned. If data element hashes are returned they MUST be encoded in the schema specified by Request Hashing Schema.
        /// If the IsRequestHashingOptionsUsed is false, this property will be ignored.
        /// </summary>
        public int RequestDataElementHashes { get; set; }

        /// <summary>
        /// Gets or sets Reserved (4 bits): A reserved bit that MUST be set to zero and MUST be ignored.
        /// If the IsRequestHashingOptionsUsed is false, this property will be ignored.
        /// </summary>
        public int Reserve3 { get; set; }

        /// <summary>
        /// Gets or sets Sub-requests
        /// </summary>
        public List<FsshttpbCellSubRequest> SubRequests { get; set; }

        /// <summary>
        /// Gets or sets Data Element Package
        /// </summary>
        public DataElementPackage DataElementPackage { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the request hash options declaration is specified.
        /// </summary>
        public bool IsRequestHashingOptionsUsed { get; set; }

        /// <summary>
        /// Gets or sets Cell Request End 
        /// </summary>
        internal StreamObjectHeaderEnd16bit CellRequestEnd { get; set; }

        /// <summary>
        /// Gets or sets Request Start (4 bytes): A 32-bit stream object header that specifies a request start.
        /// </summary>
        internal StreamObjectHeaderStart32bit RequestStart { get; set; }

        /// <summary>
        /// Gets or sets User Agent Start (4 bytes): A 32-bit stream object header that specifies a user agent start.
        /// </summary>
        internal StreamObjectHeaderStart32bit UserAgentStart { get; set; }

        /// <summary>
        /// Gets or sets User Agent GUID (4 bytes): A 32-bit stream object header that specifies a user agent GUID.
        /// </summary>
        internal StreamObjectHeaderStart32bit UserAgentGUID { get; set; }

        /// <summary>
        /// Gets or sets User Agent Version (4 bytes): A 32-bit stream object header that specifies a user agent version.
        /// </summary>
        internal StreamObjectHeaderStart32bit UserAgentVersion { get; set; }

        /// <summary>
        /// Gets or sets User Agent End (2 bytes): A 16-bit stream object header that specifies a user agent end.
        /// </summary>
        internal StreamObjectHeaderEnd16bit UserAgentEnd { get; set; }

        /// <summary>
        /// Gets or sets Request Hashing Options Declaration: A 32-bit stream object header that specifies a request hashing options declaration.
        /// </summary>
        internal StreamObjectHeaderStart RequestHashingOptionsDeclaration { get; set; }

        /// <summary>
        /// This method is used to convert the element into a byte List 
        /// </summary>
        /// <returns>Return the Byte List</returns>
        public List<byte> SerializeToByteList()
        {
            this.RequestStart = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.Request, 0);
            this.UserAgentStart = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.UserAgent, 0);
            this.UserAgentGUID = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.UserAgentGUID, 16);
            this.UserAgentVersion = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.UserAgentversion, 4);
            this.UserAgentEnd = new StreamObjectHeaderEnd16bit((int)StreamObjectTypeHeaderEnd.UserAgent);
            this.CellRequestEnd = new StreamObjectHeaderEnd16bit((int)StreamObjectTypeHeaderEnd.Request);
            
            List<byte> byteList = new List<byte>();

            // Protocol Version
            byteList.AddRange(LittleEndianBitConverter.GetBytes(this.ProtocolVersion));

            // Minimum Version
            byteList.AddRange(LittleEndianBitConverter.GetBytes(this.MinimumVersion));

            // Signature
            byteList.AddRange(LittleEndianBitConverter.GetBytes(this.Signature));

            // Request Start
            byteList.AddRange(this.RequestStart.SerializeToByteList());

            // User Agent Start
            byteList.AddRange(this.UserAgentStart.SerializeToByteList());

            // User Agent GUID
            byteList.AddRange(this.UserAgentGUID.SerializeToByteList());

            // GUID
            byteList.AddRange(this.GUID.ToByteArray());

            // User Agent Version
            byteList.AddRange(this.UserAgentVersion.SerializeToByteList());

            // Version
            byteList.AddRange(LittleEndianBitConverter.GetBytes(this.Version));

            // User Agent End
            byteList.AddRange(this.UserAgentEnd.SerializeToByteList());

            if (this.IsRequestHashingOptionsUsed)
            {
                List<byte> hashSchemaList = this.RequestHashingSchema.SerializeToByteList();
                this.RequestHashingOptionsDeclaration = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.RequestHashOptions, hashSchemaList.Count + 1);

                // Request Hashing Options Declaration
                byteList.AddRange(this.RequestHashingOptionsDeclaration.SerializeToByteList());

                // Request Hashing Schema
                byteList.AddRange(hashSchemaList);

                // Reserve
                BitWriter bw = new BitWriter(1);
                bw.AppendInit32(this.Reserve1, 1);
                bw.AppendInit32(this.Reserve2, 1);
                bw.AppendInit32(this.RequestDataElementHashesInsteadofData, 1);
                bw.AppendInit32(this.RequestDataElementHashes, 1);
                bw.AppendInit32(this.Reserve3, 4);
                byteList.AddRange(bw.Bytes);
            }

            // Sub-requests
            if (this.SubRequests != null && this.SubRequests.Count != 0)
            {
                foreach (FsshttpbCellSubRequest subRequest in this.SubRequests)
                {
                    byteList.AddRange(subRequest.SerializeToByteList());
                }
            }
            else
            {
                throw new InvalidOperationException("MUST contain sub request in request structure which is defined in the MS-FSSHTTPB.");
            }

            // Data Element Package 
            if (this.DataElementPackage != null)
            {
                byteList.AddRange(this.DataElementPackage.SerializeToByteList());
            }

            // Cell Request End 
            byteList.AddRange(this.CellRequestEnd.SerializeToByteList());

            return byteList;
        }

        /// <summary>
        /// This method is used to retrieve the Base64 encoding string.
        /// </summary>
        /// <returns>Return the Base64 string</returns>
        public string ToBase64()
        {
            return System.Convert.ToBase64String(this.SerializeToByteList().ToArray());
        }

        /// <summary>
        /// Used to add the sub request
        /// </summary>
        /// <param name="subRequest">Sub request</param>
        /// <param name="dataElement">Date elements list</param>
        public void AddSubRequest(FsshttpbCellSubRequest subRequest, List<DataElement> dataElement)
        {
            if (this.SubRequests == null)
            {
                this.SubRequests = new List<FsshttpbCellSubRequest>();
            }

            this.SubRequests.Add(subRequest);
            
            // Add the sub-request mapping for further validation usage.
            MsfsshttpbSubRequestMapping.Add((int)subRequest.RequestID, subRequest.GetType(), SharedContext.Current.Site);

            if (dataElement != null)
            {
                if (this.DataElementPackage == null)
                {
                    this.DataElementPackage = new DataElementPackage();
                }

                this.DataElementPackage.DataElements.AddRange(dataElement);
            }
        }
    }
}