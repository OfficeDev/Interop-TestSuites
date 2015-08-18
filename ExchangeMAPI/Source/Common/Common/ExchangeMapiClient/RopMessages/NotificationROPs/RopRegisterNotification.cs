namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopRegisterNotification request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopRegisterNotificationRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x29.
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
        /// This index specifies the location in the Server Object Handle Table 
        /// where the handle for the input Server Object is stored. 
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// The possible values are specified in [MS-OXCNOTIF]. These flags specify the types of events to register for.
        /// </summary>
        public byte NotificationTypes;

        /// <summary>
        /// The field is reserved. The field value MUST be zero. The behavior is undefined if the value is not zero.
        /// </summary>
        public byte Reserved;

        /// <summary>
        /// This value specifies whether the notification is scoped to the mailbox store instead of a specific folder or message.
        /// </summary>
        public byte WantWholeStore;

        /// <summary>
        /// This field is present when the WantWholeStore field is zero and is not present when it is non-zero. 
        /// This value specifies the folder to register notifications for.
        /// </summary>
        public ulong FolderId;

        /// <summary>
        /// This field is present when the WantWholeStore field is zero and is not present when it is non-zero.
        /// This value specifies the message to register notifications for.
        /// </summary>
        public ulong MessageId;

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
            serializeBuffer[index++] = this.OutputHandleIndex;
            serializeBuffer[index++] = this.NotificationTypes;
            serializeBuffer[index++] = this.Reserved;
            serializeBuffer[index++] = this.WantWholeStore;

            if (this.WantWholeStore == 0)
            {
                Array.Copy(BitConverter.GetBytes((ulong)this.FolderId), 0, serializeBuffer, index, sizeof(ulong));
                index += sizeof(ulong);
                Array.Copy(BitConverter.GetBytes((ulong)this.MessageId), 0, serializeBuffer, index, sizeof(ulong));
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
            int size = sizeof(byte) * 7;
            if (this.WantWholeStore == 0)
            {
                // 16 indicates sizeof(UInt64) * 2
                size += 16;
            }

            return size;
        }
    }

    /// <summary>
    /// RopRegisterNotification response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopRegisterNotificationResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x29.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the OutputHandleIndex specified in the request. 
        /// </summary>
        public byte OutputHandleIndex;

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
                this = (RopRegisterNotificationResponse)Marshal.PtrToStructure(
                    responseBuffer, 
                    typeof(RopRegisterNotificationResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}