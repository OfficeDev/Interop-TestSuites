namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The FolderReplicaInfo structure contains 
    /// information about server replicas of a public folder.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    [SerializableObjectAttribute(true, true)]
    public struct FolderReplicaInfo : IStructDeserializable, IStructSerializable
    {
        /// <summary>
        /// MUST be set to "0x00000000".
        /// </summary>
        [MarshalAs(UnmanagedType.U4)]
        public uint Flags;

        /// <summary>
        /// MUST be set to "0x00000000".
        /// </summary>
        [MarshalAs(UnmanagedType.U4)]
        public uint Depth;

        /// <summary>
        /// A LongTermID structure. Contains the LongTermID of a folder, 
        /// for which server replica information is being described.
        /// </summary>
        public LongTermId FolderLongTermId;

        /// <summary>
        /// An unsigned 32-bit integer value that determines 
        /// how many elements exist in ServerDNArray. 
        /// MUST NOT be zero (0x00000000).
        /// </summary>
        public uint ServerDNCount; // (4 bytes):  

        /// <summary>
        /// (4 bytes):  An unsigned 32-bit integer value that determines
        /// how many of the leading elements in ServerDNArray have the same,
        /// lowest, network access cost. CheapServerDNCount MUST be less than 
        /// or equal to ServerDNCount.
        /// </summary>
        public uint CheapServerDNCount;

        /// <summary>
        ///  An array of ASCII-encoded NULL-terminated strings. 
        ///  MUST contain ServerDNCount strings. Contains an 
        ///  enterprise/site/server distinguished name (ESSDN) 
        ///  of servers that have a replica of the folder identifier 
        ///  by FolderLongTermId.
        /// </summary>
        public string[] ServerDNArray;

        /// <summary>
        /// Initializes a new instance of the FolderReplicaInfo structure.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public FolderReplicaInfo(FastTransferStream stream)
        {
            this.Flags = stream.ReadUInt32();
            this.Depth = stream.ReadUInt32();
            this.FolderLongTermId = new LongTermId
            {
                DatabaseGuid = stream.ReadGuid().ToByteArray(),
                GlobalCounter = new byte[6]
            };
            stream.Read(
                this.FolderLongTermId.GlobalCounter,
                0,
                this.FolderLongTermId.GlobalCounter.Length);
            stream.Read(new byte[2], 0, 2);
            this.ServerDNCount = stream.ReadUInt32();
            this.CheapServerDNCount = stream.ReadUInt32();
            this.ServerDNArray = new string[this.ServerDNCount];

            for (int i = 0; i < this.ServerDNCount; i++)
            {
                this.ServerDNArray[i] = stream.ReadString8();
            }
        }

        /// <summary>
        /// Deserialize from a stream.
        /// </summary>
        /// <param name="stream">A stream contains serialize.</param>
        /// <param name="size">Must be -1.</param>
        /// <returns>The number of bytes read from the stream.</returns>
        public int Deserialize(System.IO.Stream stream, int size)
        {
            AdapterHelper.Site.Assert.AreEqual(-1, size, "The size value should be -1, but the actual value is {0}.", size);

            int bytesRead = 0;
            this.Flags = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;

            this.Depth = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;

            this.FolderLongTermId = StreamHelper.ReadLongTermId(stream);
            bytesRead += 0x10 + 6 + 2;

            this.ServerDNCount = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;

            this.CheapServerDNCount = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;

            this.ServerDNArray = new string[this.ServerDNCount];
            for (int i = 0; i < this.ServerDNCount; i++)
            {
                this.ServerDNArray[i] = StreamHelper.ReadString8(stream);
            }

            AdapterHelper.Site.Assert.AreEqual(this.ServerDNArray.Length, (int)this.ServerDNCount, "The deserialized serverDN count is not equal to the original server DN count. The expected value of the deserialized server DN is {0}, but the actual value is {1}.", this.ServerDNCount, this.ServerDNArray.Length);

            bytesRead += Common.GetBytesFromMutiUnicodeString(this.ServerDNArray).Length;
            return bytesRead;
        }

        /// <summary>
        /// Serialize this instance to a stream.
        /// </summary>
        /// <param name="stream">A data stream contains serialized object.</param>
        /// <returns>The number of bytes written to the stream.</returns>
        public int Serialize(System.IO.Stream stream)
        {
            int bytesWriten = 0;
            string serverDns = string.Empty;
            for (int i = 0; i < this.ServerDNArray.Length; i++)
            {
                serverDns += this.ServerDNArray[i];
            }

            byte[] serverDnBytes = System.Text.Encoding.ASCII.GetBytes(serverDns);

            bytesWriten += StreamHelper.WriteUInt32(stream, this.Flags);
            bytesWriten += StreamHelper.WriteUInt32(stream, this.Depth);
            bytesWriten += StreamHelper.WriteLongTermId(stream, this.FolderLongTermId);
            bytesWriten += StreamHelper.WriteUInt32(stream, this.ServerDNCount);
            bytesWriten += StreamHelper.WriteUInt32(stream, this.CheapServerDNCount);
            bytesWriten += StreamHelper.WriteBuffer(stream, serverDnBytes);
            return bytesWriten;
        }
    }
}