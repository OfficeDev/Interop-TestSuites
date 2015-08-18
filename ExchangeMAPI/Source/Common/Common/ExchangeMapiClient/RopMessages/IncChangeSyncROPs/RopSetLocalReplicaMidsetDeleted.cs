namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// This structure specifies the ranges of message identifiers.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct LongTermIdRange : ISerializable
    {
        /// <summary>
        /// LongTermId structure. The format of the LongTermId structure is specified in [MS-OXCDATA] section 2.2.1.3.1. 
        /// This identifier specifies the beginning of a range.
        /// </summary>
        public byte[] MinLongTermId;

        /// <summary>
        /// LongTermId structure. The format of the LongTermId structure is specified in [MS-OXCDATA] section 2.2.1.3.1. 
        /// This identifier specifies the end of a range.
        /// </summary>
        public byte[] MaxLongTermId;

        /// <summary>
        /// Serialize the LongTermIdRange structure.
        /// </summary>
        /// <returns>The LongTermIdRange structure serialized.</returns>
        public byte[] Serialize()
        {
            // 0 indicates start index
            int index = 0;
            byte[] serializedBuffer = new byte[this.Size()];

            // 24 indicates the total length of the LongTermID structure is 24 bytes as defined in [MS-OXCDATA]
            Array.Copy(this.MinLongTermId, 0, serializedBuffer, index, 24);
            index += 24;
            Array.Copy(this.MaxLongTermId, 0, serializedBuffer, index, 24);
            index += 24;

            return serializedBuffer;
        }

        /// <summary>
        /// Return the size of LongTermIdRange structure.
        /// </summary>
        /// <returns>The size of LongTermIdRange structure.</returns>
        public int Size()
        {
            // 48 indicates sizeof (byte) * 24 * 2
            int size = sizeof(byte) * 48;
            return size;
        }
    }

    /// <summary>
    /// RopSetLocalReplicaMidsetDeleted request buffer structure.
    /// </summary>
    public struct RopSetLocalReplicaMidsetDeletedRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x93.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table where the handle for the input Server Object is stored.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This value specifies the size of both the LongTermIdRangeCount and LongTermIdRanges fields.
        /// </summary>
        public ushort DataSize;

        /// <summary>
        /// This value specifies the number of structures in the LongTermIdRanges field.
        /// </summary>
        public uint LongTermIdRangeCount;

        /// <summary>
        /// These structures specify the ranges of message identifiers that have been deleted.
        /// </summary>
        public LongTermIdRange[] LongTermIdRanges;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            // 0 indicates start index
            int index = 0;
            byte[] serializedBuffer = new byte[this.Size()];

            serializedBuffer[index++] = this.RopId;
            serializedBuffer[index++] = this.LogonId;
            serializedBuffer[index++] = this.InputHandleIndex;

            Array.Copy(BitConverter.GetBytes((ushort)this.DataSize), 0, serializedBuffer, index, sizeof(ushort));
            index += 2;
            Array.Copy(BitConverter.GetBytes((uint)this.LongTermIdRangeCount), 0, serializedBuffer, index, sizeof(uint));
            index += 4;

            for (int i = 0; i < this.LongTermIdRangeCount; i++)
            {
                Array.Copy(this.LongTermIdRanges[i].Serialize(), 0, serializedBuffer, index, this.LongTermIdRanges[i].Size());
                index += this.LongTermIdRanges[i].Size();
            }

            return serializedBuffer;
        }

        /// <summary>
        /// Return the size of LongTermIdRange structure.
        /// </summary>
        /// <returns>The size of LongTermIdRange structure.</returns>
        public int Size()
        {
            // 9 indicates sizeof (byte) * 3 + sizeof (UInt16) + sizeof(UInt32)
            int size = sizeof(byte) * 9;

            for (int i = 0; i < this.LongTermIdRangeCount; i++)
            {
                size += this.LongTermIdRanges[i].Size();
            }

            return size;
        }
    }

    /// <summary>
    /// RopSetLocalReplicaMidsetDeleted response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSetLocalReplicaMidsetDeletedResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x93.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            IntPtr responseBuffer = new IntPtr();
            responseBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.Copy(ropBytes, startIndex, responseBuffer, Marshal.SizeOf(this));
                this = (RopSetLocalReplicaMidsetDeletedResponse)Marshal.PtrToStructure(responseBuffer, typeof(RopSetLocalReplicaMidsetDeletedResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}