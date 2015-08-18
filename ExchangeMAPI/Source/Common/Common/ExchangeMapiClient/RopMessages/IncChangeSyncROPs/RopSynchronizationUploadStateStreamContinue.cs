namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopSynchronizationUploadStateStreamContinue request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSynchronizationUploadStateStreamContinueRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x76.
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
        /// This value specifies the size of the StreamData field.
        /// </summary>
        public uint StreamDataSize;

        /// <summary>
        /// This array contains the state stream data to be uploaded.
        /// </summary>
        public byte[] StreamData;

        /// <summary>
        /// Serialize the RopSynchronizationUploadStateStreamContinue request buffer.
        /// </summary>
        /// <returns>The RopSynchronizationUploadStateStreamContinue request buffer structure serialized.</returns>
        public byte[] Serialize()
        {
            // 0 indicates start index
            int index = 0;
            byte[] serializedBuffer = new byte[this.Size()];

            serializedBuffer[index++] = this.RopId;
            serializedBuffer[index++] = this.LogonId;
            serializedBuffer[index++] = this.InputHandleIndex;

            Array.Copy(BitConverter.GetBytes((uint)this.StreamDataSize), 0, serializedBuffer, index, sizeof(uint));
            index += sizeof(uint);
            if (this.StreamDataSize > 0)
            {
                Array.Copy(this.StreamData, 0, serializedBuffer, index, this.StreamDataSize);
                index += (int)this.StreamDataSize;
            }

            return serializedBuffer;
        }

        /// <summary>
        /// Return the size of this RopSynchronizationUploadStateStreamContinue request buffer structure.
        /// </summary>
        /// <returns>The size of RopSynchronizationUploadStateStreamContinue request buffer structure.</returns>
        public int Size()
        {
            // 7 indicates sizeof (byte) * 3 + sizeof (UInt32)
            int size = sizeof(byte) * 7;
            if (this.StreamDataSize > 0)
            {
                size += (int)this.StreamDataSize;
            }

            return size;
        }
    }

    /// <summary>
    /// RopSynchronizationUploadStreamContinue response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSynchronizationUploadStateStreamContinueResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x76.
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
                this = (RopSynchronizationUploadStateStreamContinueResponse)Marshal.PtrToStructure(
                    responseBuffer, 
                    typeof(RopSynchronizationUploadStateStreamContinueResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}