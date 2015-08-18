namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopCopyTo request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopCopyToRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x39.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// Unsigned 8-bit integer. This index specifies the location in the Server Object Handle Table 
        /// where the handle for the source Server Object is stored.
        /// </summary>
        public byte SourceHandleIndex;

        /// <summary>
        /// Unsigned 8-bit integer. This index specifies the location in the Server Object Handle Table 
        /// where the handle for the destination Server Object is stored.
        /// </summary>
        public byte DestHandleIndex;

        /// <summary>
        /// 8-bit Boolean. This value specifies whether the operation is to be executed asynchronously 
        /// with status reported via RopProgress.
        /// </summary>
        public byte WantAsynchronous;

        /// <summary>
        /// 8-bit Boolean. This value specifies whether to copy sub-objects.
        /// </summary>
        public byte WantSubObjects;

        /// <summary>
        /// 8-bit flags structure. The possible values are specified in [MS-OXCPRPT]. These flags control the operation behavior.
        /// </summary>
        public byte CopyFlags;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies how many tags are present in ExcludedTags.
        /// </summary>
        public ushort ExcludedTagCount;

        /// <summary>
        /// Array of PropertyTag structures. The number of structures contained in this field is specified 
        /// by the ExcludedTagCount field. The format of the PropertyTag structure is specified in [MS-OXCDATA]. 
        /// This field specifies the properties to exclude from the copy.
        /// </summary>
        public PropertyTag[] ExcludedTags;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            int index = 0;
            byte[] serializeBuffer = new byte[this.Size()];

            serializeBuffer[index++] = this.RopId;
            serializeBuffer[index++] = this.LogonId;
            serializeBuffer[index++] = this.SourceHandleIndex;
            serializeBuffer[index++] = this.DestHandleIndex;
            serializeBuffer[index++] = this.WantAsynchronous;
            serializeBuffer[index++] = this.WantSubObjects;
            serializeBuffer[index++] = this.CopyFlags;

            Array.Copy(BitConverter.GetBytes((ushort)this.ExcludedTagCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            for (int i = 0; i < this.ExcludedTagCount; i++)
            {
                Array.Copy(this.ExcludedTags[i].Serialize(), 0, serializeBuffer, index, this.ExcludedTags[i].Size());
                index += this.ExcludedTags[i].Size();
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 9 indicates sizeof(byte) * 7 + sizeof(UInt16)
            int size = sizeof(byte) * 9;
            for (int i = 0; i < this.ExcludedTagCount; i++)
            {
                size += this.ExcludedTags[i].Size();
            }

            return size;
        }
    }

    /// <summary>
    /// RopCopyTo response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopCopyToResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x39.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index MUST be set to the SourceHandleIndex specified in the request.
        /// </summary>
        public byte SourceHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation. 
        /// For this response, this field is set to 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the number of PropertyProblem structures in the PropertyProblems field.
        /// </summary>
        public ushort PropertyProblemCount;

        /// <summary>
        /// Array of PropertyProblem structures. The number of structures contained in this field is specified 
        /// by the PropertyProblemCount field. The format of the PropertyProblem structure is specified in [MS-OXCDATA].
        /// </summary>
        public PropertyProblem[] PropertyProblems;

        /// <summary>
        /// Unsigned 32-bit integer. Available for failure response. This index MUST be set to the DestHandleIndex 
        /// specified in the request. More information about Server Objects can be found in section 3.
        /// </summary>
        public uint DestHandleIndex;

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
            this.SourceHandleIndex = ropBytes[index++];
            this.ReturnValue = (uint)BitConverter.ToInt32(ropBytes, index);
            index += sizeof(uint);

            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.PropertyProblemCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += sizeof(ushort);
                if (this.PropertyProblemCount >= 0)
                {
                    this.PropertyProblems = new PropertyProblem[this.PropertyProblemCount];
                    for (int i = 0; i < this.PropertyProblemCount; i++)
                    {
                        index += this.PropertyProblems[i].Deserialize(ropBytes, index);
                    }
                }
            }
            else if (this.ReturnValue == 0x00000503)
            {
                this.DestHandleIndex = (uint)BitConverter.ToInt32(ropBytes, index);
                index += sizeof(uint);
            }

            return index - startIndex;
        }
    }
}