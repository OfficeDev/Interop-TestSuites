namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopGetSearchCriteria request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopGetSearchCriteriaRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x31.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table 
        /// where the handle for the input Server Object is stored.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This value specifies whether the client wants the restriction data 
        /// (returned in the RestrictionData field of the response) to be specified with Unicode strings or ASCII strings.
        /// </summary>
        public byte UseUnicode;     

        /// <summary>
        /// This value specifies whether the server includes the restriction information in the response.
        /// </summary>
        public byte IncludeRestriction;

        /// <summary>
        /// This value specifies whether the server includes the folders list in the response.
        /// </summary>
        public byte IncludeFolders;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            byte[] serializeBuffer = new byte[Marshal.SizeOf(this)];
            IntPtr requestBuffer = new IntPtr();
            requestBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.StructureToPtr(this, requestBuffer, true);
                Marshal.Copy(requestBuffer, serializeBuffer, 0, Marshal.SizeOf(this));
                return serializeBuffer;
            }
            finally
            {
                Marshal.FreeHGlobal(requestBuffer);
            }
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            return Marshal.SizeOf(this);
        }
    }

    /// <summary>
    /// RopGetSearchCriteria response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopGetSearchCriteriaResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x31.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation. 
        /// For successful response, this field is set to 0x00000000.
        /// For failure response, this field is set to a value other than 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// This value specifies the length of the RestrictionData field.
        /// </summary>
        public ushort RestrictionDataSize;

        /// <summary>
        /// This field contains a restriction packet, as specified in [MS-OXCDATA] section 2.13. 
        /// The restriction specifies the filter for this search folder.
        /// </summary>
        public byte[] RestrictionData;

        /// <summary>
        /// This field MUST be set to the value of the LogonId field in the request.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// This value specifies the number of IDs in the FolderIds field.
        /// </summary>
        public ushort FolderIdCount;

        /// <summary>
        /// This field contains identifiers that specify which folders are searched.
        /// </summary>
        public ulong[] FolderIds;

        /// <summary>
        /// These flags control the search for a search folder.
        /// </summary>
        public uint SearchFlags;

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.RopId = ropBytes[index++];
            this.InputHandleIndex = ropBytes[index++];
            this.ReturnValue = (uint)BitConverter.ToInt32(ropBytes, index);
            index += sizeof(uint);
                                                                                                                                                                             
            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.RestrictionDataSize = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += sizeof(ushort);
                if (this.RestrictionDataSize > 0)
                {
                    this.RestrictionData = new byte[this.RestrictionDataSize];
                    Array.Copy(ropBytes, index, this.RestrictionData, 0, this.RestrictionDataSize);
                    index += this.RestrictionDataSize;
                }

                this.LogonId = ropBytes[index++];
                this.FolderIdCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += sizeof(ushort);

                if (this.FolderIdCount > 0)
                {
                    this.FolderIds = new ulong[this.FolderIdCount];
                    for (int i = 0; i < this.FolderIdCount; i++)
                    {
                        this.FolderIds[i] = (ulong)BitConverter.ToInt64(ropBytes, index);
                        index += sizeof(ulong);
                    }
                }

                this.SearchFlags = (uint)BitConverter.ToInt32(ropBytes, index);
                index += sizeof(uint);
            }
                                                                                                                                                                                                
            return index - startIndex;
        }
    }
}