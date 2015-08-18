namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopHardDeleteMessages request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopHardDeleteMessagesRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x91.
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
        /// This value specifies whether the operation is to be executed asynchronously with status reported via RopProgress.
        /// </summary>
        public byte WantAsynchronous;

        /// <summary>
        /// This value specifies whether the server sends a non-read receipt to the message-sender when a message is deleted.
        /// </summary>
        public byte NotifyNonRead;

        /// <summary>
        /// This value specifies the number of identifiers in the MessageIds field.
        /// </summary>
        public ushort MessageIdCount;

        /// <summary>
        /// These IDs specify the messages to be hard deleted.
        /// </summary>
        public ulong[] MessageIds;

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
            serializeBuffer[index++] = this.InputHandleIndex;
            serializeBuffer[index++] = this.WantAsynchronous;
            serializeBuffer[index++] = this.NotifyNonRead;

            Array.Copy(BitConverter.GetBytes((ushort)this.MessageIdCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            for (int i = 0; i < this.MessageIdCount; i++)
            {
                Array.Copy(BitConverter.GetBytes((ulong)this.MessageIds[i]), 0, serializeBuffer, index, sizeof(ulong));
                index += sizeof(ulong);
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 7 indicates sizeof(byte) * 5 + sizeof(UInt16)
            int size = sizeof(byte) * 7;
            for (int i = 0; i < this.MessageIdCount; i++)
            {
                size += sizeof(ulong);
            }

            return size;
        }
    }

    /// <summary>
    /// RopHardDeleteMessages response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopHardDeleteMessagesResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x91.
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
        /// This value indicates whether the operation was only partially completed.
        /// </summary>
        public byte PartialCompletion;

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
                this = (RopHardDeleteMessagesResponse)Marshal.PtrToStructure(
                    responseBuffer, 
                    typeof(RopHardDeleteMessagesResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}