namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopLogon request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopLogonRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0xFE.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This value specifies the ID that the client wants associated with the created logon. 
        /// Any value is allowed and the client does not have to use values in a certain numeric order. 
        /// If the client specifies an active logon ID, the current logon is released and replaced with the new one.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table where the handle for 
        /// the output Server Object will be stored.
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// The possible values are specified in [MS-OXCSTOR]. This structure contains flags that control the behavior of the logon.
        /// </summary>
        public byte LogonFlags;

        /// <summary>
        /// The possible values are specified in [MS-OXCSTOR]. This structure contains more flags that control the behavior of the logon.
        /// </summary>
        public uint OpenFlags;

        /// <summary>
        /// The possible values are specified in [MS-OXCSTOR]. This structure specifies ongoing action on the mailbox or public folder.
        /// </summary>
        public uint StoreState;

        /// <summary>
        /// This value specifies the size of the Essdn field.
        /// </summary>
        public ushort EssdnSize;

        /// <summary>
        /// Null terminated ASCII string. The number of characters (including the null) contained in this field is specified by the EssdnSize field. 
        /// This string specifies which mailbox to log on to.
        /// </summary>
        public byte[] Essdn;

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
            serializeBuffer[index++] = this.OutputHandleIndex;
            serializeBuffer[index++] = this.LogonFlags;
            Array.Copy(BitConverter.GetBytes((int)this.OpenFlags), 0, serializeBuffer, index, sizeof(uint));
            index += sizeof(uint);
            Array.Copy(BitConverter.GetBytes((int)this.StoreState), 0, serializeBuffer, index, sizeof(uint));
            index += sizeof(uint);
            Array.Copy(BitConverter.GetBytes((short)this.EssdnSize), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            if (this.EssdnSize > 0)
            {
                Array.Copy(this.Essdn, 0, serializeBuffer, index, this.EssdnSize);
                index += this.EssdnSize;
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            int size = (sizeof(byte) * 4) + (2 * sizeof(uint)) + sizeof(ushort);
            if (this.EssdnSize > 0)
            {
                size += this.EssdnSize;
            }

            return size;
        }
    }

    /// <summary>
    /// This structure specifies the time at which the logon occurred.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct LogonTime : IDeserializable
    {
        /// <summary>
        /// This value specifies the current second.
        /// </summary>
        public byte Seconds;

        /// <summary>
        /// This value specifies the current minute.
        /// </summary>
        public byte Minutes;

        /// <summary>
        /// This value specifies the current hour.
        /// </summary>
        public byte Hour;

        /// <summary>
        /// This value specifies the current day of the week (Sunday = 0, Monday = 1, and so on).
        /// </summary>
        public byte DayOfWeek;

        /// <summary>
        /// This value specifies the current day of the month.
        /// </summary>
        public byte Day;

        /// <summary>
        /// This value specifies the current month (January = 1, February = 2, and so on).
        /// </summary>
        public byte Month;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the current year.
        /// </summary>
        public ushort Year;

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
                this = (LogonTime)Marshal.PtrToStructure(responseBuffer, typeof(LogonTime));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }

    /// <summary>
    /// RopLogon response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopLogonResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0xFE.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table 
        /// where the handle for the output Server Object will be stored.
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation. For this response, this field is set to 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// The possible values are specified in [MS-OXCSTOR]. These flags specify the type of logon.
        /// </summary>
        public byte LogonFlags;

        /// <summary>
        /// 13 64-bit identifiers. These IDs specify a set of special folders for a mailbox.
        /// </summary>
        public ulong[] FolderIds;

        /// <summary>
        /// 16-bit identifier. This field specifies a replica ID for the logon.
        /// </summary>
        public byte[] ReplId;

        /// <summary>
        /// This field is not used and is ignored by the client. The server SHOULD set this field to an empty GUID (all zeroes).
        /// </summary>
        public byte[] PerUserGuid;

        /// <summary>
        /// 8-bit flags structure. The possible values are specified in [MS-OXCSTOR]. 
        /// These flags provide details about the state of the mailbox.
        /// </summary>
        public byte ResponseFlags;

        /// <summary>
        /// GUID. This value identifies the mailbox on which the logon was performed.
        /// </summary>
        public byte[] MailboxGuid;

        /// <summary>
        /// GUID. This field specifies the replica GUID that is associated with the replica ID, 
        /// which is specified in the ReplId field.
        /// </summary>
        public byte[] ReplGuid;

        /// <summary>
        /// LogonTime structure. The format of this structure is specified in section 2.2.2.1.2.1. 
        /// This structure specifies the time at which the logon occurred.
        /// </summary>
        public LogonTime LogonTime;

        /// <summary>
        /// Unsigned 64-bit integer. This value represents the number of 100-nanosecond intervals since January 1, 1601. 
        /// This time specifies when the Gateway Address Routing Table last changed.
        /// </summary>
        public ulong GwartTime;

        /// <summary>
        /// 32-bit flags structure. The possible values are specified in [MS-OXCSTOR]. 
        /// These flags specify ongoing action on the mailbox or public folder.
        /// </summary>
        public uint StoreState;

        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the length of the ServerName field.
        /// </summary>
        public byte ServerNameSize;

        /// <summary>
        /// Null terminated ASCII string. The number of characters (including the null) contained in this field 
        /// is specified by the ServerNameSize field. This string specifies a different server for the client to connect to.
        /// </summary>
        public byte[] ServerName;

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
            this.OutputHandleIndex = ropBytes[index++];
            this.ReturnValue = (uint)BitConverter.ToInt32(ropBytes, index);
            index += sizeof(uint);

            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.LogonFlags = ropBytes[index++];
                if ((this.LogonFlags & (byte)TestSuites.Common.LogonFlags.Private) != 0x00)
                {
                    // The length of FolderIds is 13.
                    this.FolderIds = new ulong[13];
                    for (int i = 0; i < 13; i++)
                    {
                        this.FolderIds[i] = (ulong)BitConverter.ToInt64(ropBytes, index);
                        index += sizeof(ulong);
                    }

                    this.ResponseFlags = ropBytes[index++];

                    // Guid holds 16 bytes.
                    this.MailboxGuid = new byte[16];
                    Array.Copy(ropBytes, index, this.MailboxGuid, 0, 16);
                    index += 16;

                    // ReplId holds 16 bits.
                    this.ReplId = new byte[2];
                    Array.Copy(ropBytes, index, this.ReplId, 0, 2);
                    index += 2;

                    // ReplGuid holds 16 bytes.
                    this.ReplGuid = new byte[16];
                    Array.Copy(ropBytes, index, this.ReplGuid, 0, 16);
                    index += 16;
                    index += this.LogonTime.Deserialize(ropBytes, index);
                    this.GwartTime = (ulong)BitConverter.ToInt64(ropBytes, index);
                    index += sizeof(ulong);
                    this.StoreState = (uint)BitConverter.ToInt32(ropBytes, index);
                    index += sizeof(uint);
                }
                else
                {
                    // The length of FolderIds is 13.
                    this.FolderIds = new ulong[13];
                    for (int i = 0; i < 13; i++)
                    {
                        this.FolderIds[i] = (ulong)BitConverter.ToInt64(ropBytes, index);
                        index += sizeof(ulong);
                    }

                    // ReplId holds 16 bits.
                    this.ReplId = new byte[2];
                    Array.Copy(ropBytes, index, this.ReplId, 0, 2);
                    index += 2;

                    // ReplGuid holds 16 bytes.
                    this.ReplGuid = new byte[16];
                    Array.Copy(ropBytes, index, this.ReplGuid, 0, 16);
                    index += 16;

                    // PerUserGuid holds 16 bytes.
                    this.PerUserGuid = new byte[16];
                    Array.Copy(ropBytes, index, this.PerUserGuid, 0, 16);
                    index += 16;
                }
            }
            else if (this.ReturnValue == 0x00000478)
            {
                // Redirect response 0x00000478.
                this.LogonFlags = ropBytes[index++];
                this.ServerNameSize = ropBytes[index++];
                if (this.ServerNameSize > 0)
                {
                    this.ServerName = new byte[this.ServerNameSize];
                    Array.Copy(ropBytes, index, this.ServerName, 0, this.ServerNameSize);
                    index += this.ServerNameSize;
                }
            }

            return index - startIndex;
        }
    }    
}