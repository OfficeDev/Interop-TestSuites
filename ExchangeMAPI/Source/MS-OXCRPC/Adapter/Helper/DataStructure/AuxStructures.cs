namespace Microsoft.Protocols.TestSuites.MS_OXCRPC
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// AUX_PERF_SESSIONINFO structure
    /// </summary>
    public struct AUX_PERF_SESSIONINFO 
    {
        /// <summary>
        /// SessionID (2 bytes):  Session identification number.
        /// </summary>
        public short SessionID;

        /// <summary>
        /// Reserved (2 bytes), padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream. 
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public short Reserved;

        /// <summary>
        /// SessionGuid (16 bytes), GUID representing the client session to associate with the session identification number
        /// in field SessionID.
        /// </summary>
        public byte[] SessionGuid;

        /// <summary>
        /// Serializes AUX_PERF_SESSIONINFO to a byte array
        /// </summary>
        /// <returns>Returns the byte array of serialized AUX_PERF_SESSIONINFO</returns>
        public byte[] Serialize()
        {
            if (this.SessionGuid == null)
            {
                // According to Open Specification, this field should be a 16 byte array
                this.SessionGuid = new byte[ConstValues.GuidByteSize];
            }

            // Refer to 2.2.2.4 AUX_PERF_SESSIONINFO
            int size = (sizeof(short) * 2) + this.SessionGuid.Length;
            byte[] resultBytes = new byte[size];
            int index = 0;
            Array.Copy(BitConverter.GetBytes(this.SessionID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.Reserved), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(this.SessionGuid, 0, resultBytes, index, this.SessionGuid.Length);
            return resultBytes;
        }
    }

    /// <summary>
    /// AUX_PERF_SESSIONINFO_V2 structure
    /// </summary>
    public struct AUX_PERF_SESSIONINFO_v2
    {
        /// <summary>
        /// SessionID (2 bytes):  Session identification number.
        /// </summary>
        public short SessionID;

        /// <summary>
        /// Reserved (2 bytes):  Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream.
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public short Reserved;

        /// <summary>
        /// SessionGuid (16 bytes):  GUID representing the client session to associate with the session identification number
        /// in field SessionID.
        /// </summary>
        public byte[] SessionGuid;

        /// <summary>
        /// ConnectionID (4 bytes):  Connection identification number.
        /// </summary>
        public int ConnectionID;

        /// <summary>
        /// Serializes AUX_PERF_SESSIONINFO_V2 to a byte array
        /// </summary>
        /// <returns>Returns the byte array of serialized AUX_PERF_SESSIONINFO_V2</returns>
        public byte[] Serialize()
        {
            if (this.SessionGuid == null)
            {
                // According to Open Specification, this field should be a 16 byte array
                this.SessionGuid = new byte[ConstValues.GuidByteSize];
            }

            // Refer to 2.2.2.5 AUX_PERF_SESSIONINFO_V2
            int size = (sizeof(short) * 2) + sizeof(int) + this.SessionGuid.Length;
            byte[] resultBytes = new byte[size];
            int index = 0;
            Array.Copy(BitConverter.GetBytes(this.SessionID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.Reserved), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(this.SessionGuid, 0, resultBytes, index, this.SessionGuid.Length);
            index += this.SessionGuid.Length;
            Array.Copy(BitConverter.GetBytes(this.ConnectionID), 0, resultBytes, index, sizeof(int));
            return resultBytes;
        }
    }

    /// <summary>
    /// AUX_PERF_CLIENTINFO structure
    /// </summary>
    public struct AUX_PERF_CLIENTINFO 
    {
        /// <summary>
        /// AdapterSpeed (4 bytes):  Speed of client computer's network adapter (kbits/s).
        /// </summary>
        public int AdapterSpeed;

        /// <summary>
        /// ClientID (2 bytes):  Client-assigned identification number.
        /// </summary>
        public short ClientID;

        /// <summary>
        /// MachineNameOffset (2 bytes): The offset from the beginning of the AUX_HEADER structure to the MachineName field.
        /// A value of zero indicates that the MachineName field is null or empty.
        /// </summary>
        public short MachineNameOffset;

        /// <summary>
        /// UserNameOffset (2 bytes): The offset from the beginning of the AUX_HEADER structure to the UserName field.
        /// A value of zero indicates that the UserName field is null or empty.
        /// </summary>
        public short UserNameOffset;

        /// <summary>
        /// ClientIPSize (2 bytes): Size of the client IP address referenced by the ClientIPOffset field.
        /// The client IP address is located in the ClientIP field.
        /// </summary>
        public short ClientIPSize;

        /// <summary>
        /// ClientIPOffset (2 bytes): The offset from the beginning of the AUX_HEADER structure to the ClientIP field.
        /// A value of zero indicates that the ClientIP field is null or empty.
        /// </summary>
        public short ClientIPOffset;

        /// <summary>
        /// ClientIPMaskSize (2 bytes): Size of the client IP subnet mask referenced by the ClientIPMaskOffset field.
        /// The client IP mask is located in the ClientIPMask field.
        /// </summary>
        public short ClientIPMaskSize;

        /// <summary>
        /// ClientIPMaskOffset (2 bytes): The offset from the beginning of the AUX_HEADER structure to the ClientIPMask field.
        /// The size of the IP subnet mask is found in the ClientIPMaskSize field.
        /// A value of zero indicates that the ClientIPMask field is null or empty.
        /// </summary>
        public short ClientIPMaskOffset;

        /// <summary>
        /// AdapterNameOffset (2 bytes): The offset from the beginning of the AUX_HEADER structure to the AdapterName field.
        /// A value of zero indicates that the AdapterName field is null or empty.
        /// </summary>
        public short AdapterNameOffset;

        /// <summary>
        /// MacAddressSize (2 bytes): Size of the network adapter MAC address referenced by the MacAddressOffset field.
        /// The network adapter MAC address is located in the MacAddress field.
        /// </summary>
        public short MacAddressSize;

        /// <summary>
        /// MacAddressOffset (2 bytes): The offset from the beginning of the AUX_HEADER structure to the MacAddress field.
        /// A value of zero indicates that the MacAddress field is null or empty.
        /// </summary>
        public short MacAddressOffset;

        /// <summary>
        /// ClientMode (2 bytes): Determines the mode in which the client is running.
        /// </summary>
        public short ClientMode;

        /// <summary>
        /// Reserved (2 bytes): Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream.
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public short Reserved;

        /// <summary>
        /// MachineName (variable): A null-terminated Unicode string that contains the client computer name.
        /// This variable field is offset from the beginning of the AUX_HEADER structure by the MachineNameOffset value.
        /// </summary>
        public byte[] MachineName;

        /// <summary>
        /// UserName (variable): A null-terminated Unicode string that contains the user's account name.
        /// This variable field is offset from the beginning of the AUX_HEADER structure by the UserNameOffset value.
        /// </summary>
        public byte[] UserName;

        /// <summary>
        /// ClientIP (variable): The client's IP address.
        /// This variable field is offset from the beginning of the AUX_HEADER structure by the ClientIPOffset value.
        /// The size of the client IP address data is found in the ClientIPSize field.
        /// </summary>
        public byte[] ClientIP;

        /// <summary>
        /// ClientIPMask (variable): The client's IP subnet mask.
        /// This variable field is offset from the beginning of the AUX_HEADER structure by the ClientIPMaskOffset value.
        /// The size of the client IP mask data is found in the ClientIPMaskSize field.
        /// </summary>
        public byte[] ClientIPMask;

        /// <summary>
        /// AdapterName (variable): A null-terminated Unicode string that contains the client network adapter name.
        /// This variable field is offset from the beginning of the AUX_HEADER structure by the AdapterNameOffset value.
        /// </summary>
        public byte[] AdapterName;

        /// <summary>
        /// MacAddress (variable):The client's network adapter MAC address.
        /// This variable field is offset from the beginning of the AUX_HEADER structure by the MacAddressOffset value.
        /// The size of the network adapter MAC address data is found in the MacAddressSize field.
        /// </summary>
        public byte[] MacAddress;

        /// <summary>
        /// Serializes AUX_PERF_CLIENTINFO to a byte array
        /// </summary>
        /// <returns>Returns the byte array of serialized AUX_PERF_CLIENTINFO</returns>
        public byte[] Serialize()
        {
            // Refer to 2.2.2.6 AUX_PERF_CLIENTINFO
            int size = sizeof(int) + (sizeof(short) * 12);
            if (this.MachineName != null)
            {
                size += this.MachineName.Length;
            }

            if (this.UserName != null)
            {
                size += this.UserName.Length;
            }

            if (this.ClientIP != null)
            {
                size += this.ClientIP.Length;
            }

            if (this.ClientIPMask != null)
            {
                size += this.ClientIPMask.Length;
            }

            if (this.AdapterName != null)
            {
                size += this.AdapterName.Length;
            }

            if (this.MacAddress != null)
            {
                size += this.MacAddress.Length;
            }

            byte[] resultBytes = new byte[size];
            int index = 0;
            Array.Copy(BitConverter.GetBytes(this.AdapterSpeed), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            Array.Copy(BitConverter.GetBytes(this.ClientID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.MachineNameOffset), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.UserNameOffset), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ClientIPSize), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ClientIPOffset), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ClientIPMaskSize), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.AdapterNameOffset), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.MacAddressSize), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ClientMode), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.Reserved), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ClientIPMaskOffset), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            if (this.MachineName != null)
            {
                Array.Copy(this.MachineName, 0, resultBytes, index, this.MachineName.Length);
                index += this.MachineName.Length;
            }

            if (this.UserName != null)
            {
                Array.Copy(this.UserName, 0, resultBytes, index, this.UserName.Length);
                index += this.UserName.Length;
            }

            if (this.ClientIP != null)
            {
                Array.Copy(this.ClientIP, 0, resultBytes, index, this.ClientIP.Length);
                index += this.ClientIP.Length;
            }

            if (this.ClientIPMask != null)
            {
                Array.Copy(this.ClientIPMask, 0, resultBytes, index, this.ClientIPMask.Length);
                index += this.ClientIPMask.Length;
            }

            if (this.AdapterName != null)
            {
                Array.Copy(this.AdapterName, 0, resultBytes, index, this.AdapterName.Length);
                index += this.AdapterName.Length;
            }

            if (this.MacAddress != null)
            {
                Array.Copy(this.MacAddress, 0, resultBytes, index, this.MacAddress.Length);
                index += this.MacAddress.Length;
            }

            return resultBytes;
        }
    }

    /// <summary>
    /// 2.2.2.8   AUX_PERF_PROCESSINFO
    /// </summary>
    public struct AUX_PERF_PROCESSINFO 
    {
        /// <summary>
        /// ProcessID (2 bytes): Client-assigned process identification number.
        /// </summary>
        public short ProcessID;

        /// <summary>
        /// Reserved_1 (2 bytes): Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream.
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public short Reserved1;

        /// <summary>
        /// ProcessGuid (16 bytes): GUID representing the client process to associate with the process identification number
        /// in field ProcessID.
        /// </summary>
        public byte[] ProcessGuid;

        /// <summary>
        /// ProcessNameOffset (2 bytes): The offset from the beginning of the AUX_HEADER structure to the ProcessName field.
        /// A value of zero indicates that the ProcessName field is null or empty.
        /// </summary>
        public short ProcessNameOffset;

        /// <summary>
        /// Reserved_2 (2 bytes): Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream.
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public short Reserved2;

        /// <summary>
        /// ProcessName (variable): A null-terminated Unicode string that contains the client process name.
        /// This variable field is offset from the beginning of the AUX_HEADER structure by the ProcessNameOffset value.
        /// </summary>
        public byte[] ProcessName;

        /// <summary>
        /// Serializes AUX_PERF_PROCESSINFO to a byte array
        /// </summary>
        /// <returns>Returns the byte array of serialized AUX_PERF_PROCESSINFO</returns>
        public byte[] Serialize()
        {
            if (this.ProcessGuid == null)
            {
                // According to Open Specification, this field should be a 16 byte array
                this.ProcessGuid = new byte[ConstValues.GuidByteSize];
            }
                                                                                                                                                                             
            // Refer to 2.2.2.8 AUX_PERF_PROCESSINFO
            int size = (sizeof(short) * 4) + this.ProcessGuid.Length;
            if (this.ProcessName != null)
            {
                size += this.ProcessName.Length;
            }
                                                                                                                                                                                                
            byte[] resultBytes = new byte[size];
            int index = 0;
            Array.Copy(BitConverter.GetBytes(this.ProcessID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.Reserved1), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(this.ProcessGuid, 0, resultBytes, index, this.ProcessGuid.Length);
            index += this.ProcessGuid.Length;
            Array.Copy(BitConverter.GetBytes(this.ProcessNameOffset), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.Reserved2), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            if (this.ProcessName != null)
            {
                Array.Copy(this.ProcessName, 0, resultBytes, index, this.ProcessName.Length);
                index += this.ProcessName.Length;
            }
                                                                                                                                                                                                
            return resultBytes;
        }
    }

    /// <summary>
    /// 2.2.2.9 AUX_PERF_DEFMDB_SUCCESS
    /// </summary>
    public struct AUX_PERF_DEFMDB_SUCCESS 
    {
        /// <summary>
        /// TimeSinceRequest (4 bytes): Number of milliseconds since successful request occurred.
        /// </summary>
        public int TimeSinceRequest;

        /// <summary>
        /// TimeToCompleteRequest (4 bytes): Number of milliseconds the successful request took to complete.
        /// </summary>
        public int TimeToCompleteRequest;

        /// <summary>
        /// RequestID (2 bytes): Request identification number.
        /// </summary>
        public short RequestID;

        /// <summary>
        /// Reserved (2 bytes): Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream.
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public short Reserved;

        /// <summary>
        /// Serializes AUX_PERF_DEFMDB_SUCCESS to a byte array
        /// </summary>
        /// <returns>Returns the byte array of serialized AUX_PERF_DEFMDB_SUCCESS</returns>
        public byte[] Serialize()
        {
            // 2.2.2.9 AUX_PERF_DEFMDB_SUCCESS
            int size = (sizeof(short) * 2) + (sizeof(int) * 2);
            byte[] resultBytes = new byte[size];
            int index = 0;
            Array.Copy(BitConverter.GetBytes(this.TimeSinceRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            Array.Copy(BitConverter.GetBytes(this.TimeToCompleteRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            Array.Copy(BitConverter.GetBytes(this.RequestID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.Reserved), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            return resultBytes;
        }
    }

    /// <summary>
    /// 2.2.2.10 AUX_PERF_DEFGC_SUCCESS
    /// </summary>
    public struct AUX_PERF_DEFGC_SUCCESS 
    {
        /// <summary>
        /// ServerID (2 bytes): Server identification number.
        /// </summary>
        public short ServerID;

        /// <summary>
        /// SessionID (2 bytes): Session identification number.
        /// </summary>
        public short SessionID;

        /// <summary>
        /// TimeSinceRequest (4 bytes): Number of milliseconds since successful request occurred.
        /// </summary>
        public int TimeSinceRequest;

        /// <summary>
        /// TimeToCompleteRequest (4 bytes): Number of milliseconds the successful request took to complete.
        /// </summary>
        public int TimeToCompleteRequest;

        /// <summary>
        /// RequestOperation (1 byte): Client-defined operation that was successful.
        /// </summary>
        public byte RequestOperation;

        /// <summary>
        /// Reserved (3 bytes): Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream.
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public int Reserved;

        /// <summary>
        /// Serializes AUX_PERF_DEFGC_SUCCESS to a byte array
        /// </summary>
        /// <returns>Returns the byte array of serialized AUX_PERF_DEFGC_SUCCESS</returns>
        public byte[] Serialize()
        {
            // 2.2.2.10 AUX_PERF_DEFGC_SUCCESS
            int size = 18;
            byte[] resultBytes = new byte[size];
            int index = 0;
            Array.Copy(BitConverter.GetBytes(this.ServerID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.SessionID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.TimeSinceRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            Array.Copy(BitConverter.GetBytes(this.TimeToCompleteRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            resultBytes[index] = this.RequestOperation;
            index += sizeof(byte);

            // The length of processName is three bytes, so the first byte will be abandoned.
            Array.Copy(BitConverter.GetBytes(this.Reserved), 1, resultBytes, index, 3);
            return resultBytes;
        }
    }

    /// <summary>
    /// 2.2.2.12 AUX_PERF_MDB_SUCCESS_V2
    /// </summary>
    public struct AUX_PERF_MDB_SUCCESS_V2 
    {
        /// <summary>
        /// ProcessID (2 bytes): Process identification number.
        /// </summary>
        public short ProcessID;

        /// <summary>
        /// ClientID (2 bytes): Client identification number.
        /// </summary>
        public short ClientID;

        /// <summary>
        /// ServerID (2 bytes): Server identification number.
        /// </summary>
        public short ServerID;

        /// <summary>
        /// SessionID (2 bytes): Session identification number.
        /// </summary>
        public short SessionID;

        /// <summary>
        /// RequestID (2 bytes): Request identification number.
        /// </summary>
        public short RequestID;

        /// <summary>
        /// Reserved (2 bytes): Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream.
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public short Reserved;

        /// <summary>
        /// TimeSinceRequest (4 bytes): Number of milliseconds since successful request occurred.
        /// </summary>
        public int TimeSinceRequest;

        /// <summary>
        /// TimeToCompleteRequest (4 bytes): Number of milliseconds the successful request took to complete.
        /// </summary>
        public int TimeToCompleteRequest;

        /// <summary>
        /// Serializes AUX_PERF_MDB_SUCCESS_V2 to a byte array
        /// </summary>
        /// <returns>Returns the byte array of serialized AUX_PERF_MDB_SUCCESS_V2</returns>
        public byte[] Serialize()
        {
            // 2.2.2.12 AUX_PERF_MDB_SUCCESS_V2
            int size = (sizeof(short) * 6) + (sizeof(int) * 2);
            byte[] resultBytes = new byte[size];
            int index = 0;
            Array.Copy(BitConverter.GetBytes(this.ProcessID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ClientID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ServerID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.SessionID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.RequestID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.Reserved), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.TimeSinceRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            Array.Copy(BitConverter.GetBytes(this.TimeToCompleteRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            return resultBytes;
        }
    }

    /// <summary>
    /// 2.2.2.13 AUX_PERF_GC_SUCCESS
    /// </summary>
    public struct AUX_PERF_GC_SUCCESS
    {
        /// <summary>
        /// ClientID (2 bytes): Client identification number.
        /// </summary>
        public short ClientID;

        /// <summary>
        /// ServerID (2 bytes): Server identification number.
        /// </summary>
        public short ServerID;

        /// <summary>
        /// SessionID (2 bytes): Session identification number.
        /// </summary>
        public short SessionID;

        /// <summary>
        /// Reserved_1 (2 bytes): Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream. 
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public short Reserved1;

        /// <summary>
        /// TimeSinceRequest (4 bytes): Number of milliseconds since successful request occurred.
        /// </summary>
        public int TimeSinceRequest;

        /// <summary>
        /// TimeToCompleteRequest (4 bytes): Number of milliseconds the successful request took to complete.
        /// </summary>
        public int TimeToCompleteRequest;

        /// <summary>
        /// RequestOperation (1 byte): Client-defined operation that was successful.
        /// </summary>
        public byte RequestOperation;

        /// <summary>
        /// Reserved_2 (3 bytes): Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream.
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public int Reserved2;

        /// <summary>
        /// Serializes AUX_PERF_GC_SUCCESS to a byte array
        /// </summary>
        /// <returns>Returns the byte array of serialized AUX_PERF_GC_SUCCESS</returns>
        public byte[] Serialize()
        {
            // 2.2.2.13 AUX_PERF_GC_SUCCESS
            int size = (sizeof(short) * 4) + (sizeof(int) * 3) + sizeof(byte);
            byte[] resultBytes = new byte[size];
            int index = 0;
            Array.Copy(BitConverter.GetBytes(this.ClientID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ServerID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.SessionID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.Reserved1), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.TimeSinceRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            Array.Copy(BitConverter.GetBytes(this.TimeToCompleteRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            resultBytes[index] = this.RequestOperation;
            index += sizeof(byte);

            // The length of processName is three bytes, so the first byte will be abandoned.
            Array.Copy(BitConverter.GetBytes(this.Reserved2), 1, resultBytes, index, 3);
            return resultBytes;
        }
    }

    /// <summary>
    /// 2.2.2.14 AUX_PERF_GC_SUCCESS_V2
    /// </summary>
    public struct AUX_PERF_GC_SUCCESS_V2 
    {
        /// <summary>
        /// ProcessID (2 bytes): Process identification number.
        /// </summary>
        public short ProcessID;

        /// <summary>
        /// ClientID (2 bytes): Client identification number.
        /// </summary>
        public short ClientID;

        /// <summary>
        /// ServerID (2 bytes): Server identification number.
        /// </summary>
        public short ServerID;

        /// <summary>
        /// SessionID (2 bytes): Session identification number.
        /// </summary>
        public short SessionID;

        /// <summary>
        /// TimeSinceRequest (4 bytes): Number of milliseconds since successful request occurred.
        /// </summary>
        public int TimeSinceRequest;

        /// <summary>
        /// TimeToCompleteRequest (4 bytes): Number of milliseconds the successful request took to complete.
        /// </summary>
        public int TimeToCompleteRequest;

        /// <summary>
        /// RequestOperation (1 byte): Client-defined operation that was successful.
        /// </summary>
        public byte RequestOperation;

        /// <summary>
        /// Reserved (3 bytes): Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream.
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public int Reserved;

        /// <summary>
        /// Serializes AUX_PERF_GC_SUCCESS_V2 to a byte array
        /// </summary>
        /// <returns>Returns the byte array of serialized AUX_PERF_GC_SUCCESS_V2</returns>
        public byte[] Serialize()
        {
            // 2.2.2.14 AUX_PERF_GC_SUCCESS_V2
            int size = (sizeof(short) * 4) + (sizeof(int) * 3) + sizeof(byte);
            byte[] resultBytes = new byte[size];
            int index = 0;
            Array.Copy(BitConverter.GetBytes(this.ProcessID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ClientID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ServerID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.SessionID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.TimeSinceRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            Array.Copy(BitConverter.GetBytes(this.TimeToCompleteRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            resultBytes[index] = this.RequestOperation;
            index += sizeof(byte);

            // The length of processName is three bytes, so the first byte will be abandoned.
            Array.Copy(BitConverter.GetBytes(this.Reserved), 1, resultBytes, index, 3);
            return resultBytes;
        }
    }

    /// <summary>
    /// 2.2.2.15 AUX_PERF_FAILURE
    /// </summary>
    public struct AUX_PERF_FAILURE 
    {
        /// <summary>
        /// ClientID (2 bytes): Client identification number.
        /// </summary>
        public short ClientID;

        /// <summary>
        /// ServerID (2 bytes): Server identification number.
        /// </summary>
        public short ServerID;

        /// <summary>
        /// SessionID (2 bytes): Session identification number.
        /// </summary>
        public short SessionID;

        /// <summary>
        /// RequestID (2 bytes): Request identification number.
        /// </summary>
        public short RequestID;

        /// <summary>
        /// TimeSinceRequest (4 bytes): Number of milliseconds since failure request occurred.
        /// </summary>
        public int TimeSinceRequest;

        /// <summary>
        /// TimeToFailRequest (4 bytes): Number of milliseconds the failure request took to complete.
        /// </summary>
        public int TimeToFailRequest;

        /// <summary>
        /// ResultCode (4 bytes): Error code return of failed request. Returned error codes are implementation specific.
        /// </summary>
        public int ResultCode;

        /// <summary>
        /// RequestOperation (1 byte): Client-defined operation that failed.
        /// </summary>
        public byte RequestOperation;

        /// <summary>
        /// Reserved (3 bytes): Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream.
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public int Reserved;

        /// <summary>
        /// Serializes AUX_PERF_FAILURE to a byte array
        /// </summary>
        /// <returns>Returns the byte array of serialized AUX_PERF_FAILURE</returns>
        public byte[] Serialize()
        {
            // 2.2.2.15 AUX_PERF_FAILURE
            int size = (sizeof(short) * 4) + (sizeof(int) * 4) + sizeof(byte);
            byte[] resultBytes = new byte[size];
            int index = 0;
            Array.Copy(BitConverter.GetBytes(this.ClientID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ServerID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.SessionID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.RequestID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.TimeSinceRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            Array.Copy(BitConverter.GetBytes(this.TimeToFailRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            Array.Copy(BitConverter.GetBytes(this.ResultCode), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            resultBytes[index] = this.RequestOperation;
            index += sizeof(byte);

            // The length of processName is three bytes, so the first byte will be abandoned.
            Array.Copy(BitConverter.GetBytes(this.Reserved), 1, resultBytes, index, 3);
            return resultBytes;
        }
    }

    /// <summary>
    /// 2.2.2.16 AUX_PERF_FAILURE_V2
    /// </summary>
    public struct AUX_PERF_FAILURE_V2 
    {
        /// <summary>
        /// ProcessID (2 bytes): Process identification number.
        /// </summary>
        public short ProcessID;

        /// <summary>
        /// ClientID (2 bytes): Client identification number.
        /// </summary>
        public short ClientID;

        /// <summary>
        /// ServerID (2 bytes): Server identification number.
        /// </summary>
        public short ServerID;

        /// <summary>
        /// SessionID (2 bytes): Session identification number.
        /// </summary>
        public short SessionID;

        /// <summary>
        /// RequestID (2 bytes): Request identification number.
        /// </summary>
        public short RequestID;

        /// <summary>
        /// Reserved_1 (2 bytes): Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream.
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public short Reserved1;

        /// <summary>
        /// TimeSinceRequest (4 bytes): Number of milliseconds since failure request occurred.
        /// </summary>
        public int TimeSinceRequest;

        /// <summary>
        /// TimeToFailRequest (4 bytes): Number of milliseconds the failure request took to complete.
        /// </summary>
        public int TimeToFailRequest;

        /// <summary>
        /// ResultCode (4 bytes): Error code return of failed request. Returned error codes are implementation specific.
        /// </summary>
        public int ResultCode;

        /// <summary>
        /// RequestOperation (1 byte): Client-defined operation that failed.
        /// </summary>
        public byte RequestOperation;

        /// <summary>
        /// Reserved_2 (3 bytes): Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream.
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public int Reserved2;

        /// <summary>
        /// Serializes AUX_PERF_FAILURE_V2 to a byte array
        /// </summary>
        /// <returns>Returns the byte array of serialized AUX_PERF_FAILURE_V2</returns>
        public byte[] Serialize()
        {
            // 2.2.2.16 AUX_PERF_FAILURE_V2
            int size = (sizeof(short) * 6) + (sizeof(int) * 4) + sizeof(byte);
            byte[] resultBytes = new byte[size];
            int index = 0;
            Array.Copy(BitConverter.GetBytes(this.ProcessID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ClientID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ServerID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.SessionID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.RequestID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.Reserved1), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.TimeSinceRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            Array.Copy(BitConverter.GetBytes(this.TimeToFailRequest), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            Array.Copy(BitConverter.GetBytes(this.ResultCode), 0, resultBytes, index, sizeof(int));
            index += sizeof(int);
            resultBytes[index] = this.RequestOperation;
            index += sizeof(byte);

            // The length of processName is three bytes, so the first byte will be abandoned.
            Array.Copy(BitConverter.GetBytes(this.Reserved2), 1, resultBytes, index, 3);
            return resultBytes;
        }
    }

    /// <summary>
    /// 2.2.2.20 AUX_PERF_ACCOUNTINFO
    /// </summary>
    public struct AUX_PERF_ACCOUNTINFO
    {
        /// <summary>
        /// ClientID (2 bytes): Client assigned identification number.
        /// </summary>
        public short ClientID;

        /// <summary>
        /// Reserved (2 bytes): Padding to enforce alignment of the data on a 4-byte field.
        /// The client can fill this field with any value when writing the stream.
        /// The server MUST ignore the value of this field when reading the stream.
        /// </summary>
        public short Reserved;

        /// <summary>
        /// Account (16 bytes): A GUID representing the client account information that relates to the client
        /// identification number in the ClientID field.
        /// </summary>
        public byte[] Account;

        /// <summary>
        /// Serializes AUX_PERF_ACCOUNTINFO to a byte array
        /// </summary>
        /// <returns>Returns the byte array of serialized AUX_PERF_ACCOUNTINFO</returns>
        public byte[] Serialize()
        {
            if (this.Account == null)
            {
                // According to Open Specification, this field should be a 16 byte array
                this.Account = new byte[ConstValues.GuidByteSize];
            }

            // 2.2.2.20 AUX_PERF_ACCOUNTINFO
            int size = (sizeof(short) * 2) + this.Account.Length;
            byte[] resultBytes = new byte[size];
            int index = 0;
            Array.Copy(BitConverter.GetBytes(this.ClientID), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.Reserved), 0, resultBytes, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(this.Account, 0, resultBytes, index, this.Account.Length);
            return resultBytes;
        }
    }

    /// <summary>
    /// The data structure of AUX_CLIENT_CONNECTION_INFO
    /// </summary>
    public struct AUX_CLIENT_CONNECTION_INFO 
    {
        /// <summary>
        /// The GUID of the connection to the server.
        /// </summary>
        public Guid ConnectionGUID;

        /// <summary>
        /// The offset from the beginning of the AUX_HEADER structure to the ConnectionContextInfo field.
        /// </summary>
        public short OffsetConnectionContextInfo;

        /// <summary>
        /// Padding to enforce alignment of the data on a 4-byte field.
        /// </summary>
        public short Reserved;

        /// <summary>
        /// The number of connection attempts.
        /// </summary>
        public short ConnectionAttempts;

        /// <summary>
        /// A value of 0x0001 for this field means that the client is running in cached mode.
        /// A value of 0x0000 means that the client is not designating a mode of operation. 
        /// </summary>
        public int ConnectionFlags;

        /// <summary>
        /// A null-terminated Unicode string that contains opaque connection context information to be logged by the server. 
        /// </summary>
        public string ConnectionContextInfo;

        /// <summary>
        /// Serializes AUX_CLIENT_CONNECTION_INFO to a byte array
        /// </summary>
        /// <returns>Returns the byte array of serialized AUX_CLIENT_CONNECTION_INFO.</returns>
        public byte[] Serialize()
        {
            byte[] connectionContextInfo = new byte[0];
            if (!string.IsNullOrEmpty(this.ConnectionContextInfo))
            {
                connectionContextInfo = System.Text.Encoding.Unicode.GetBytes(this.ConnectionContextInfo);
            }

            int size = 28 + connectionContextInfo.Length;
            byte[] resultBytes = new byte[size];

            int index = 0;
            Array.Copy(this.ConnectionGUID.ToByteArray(), 0, resultBytes, index, 16);
            index = index + 16;
            Array.Copy(BitConverter.GetBytes(this.OffsetConnectionContextInfo), 0, resultBytes, index, sizeof(short));
            index = index + sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.Reserved), 0, resultBytes, index, sizeof(short));
            index = index + sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ConnectionAttempts), 0, resultBytes, index, sizeof(short));
            index = index + sizeof(short);
            Array.Copy(BitConverter.GetBytes(this.ConnectionFlags), 0, resultBytes, index, sizeof(int));
            index = index + sizeof(int);
            Array.Copy(connectionContextInfo, 0, resultBytes, index, connectionContextInfo.Length);

            return resultBytes;
        }
    }

    /// <summary>
    /// The structure of rgbAuxOut that contains the AUX_CLIENT_CONTROL, AUX_OSVERSIONINFO or AUX_EXORGINFO.
    /// </summary>
    public struct AUX_SERVER_TOPOLOGY_STRUCTURE
    {
        /// <summary>
        /// The header before the payload in rgbAuxOut.
        /// </summary>
        public AUX_HEADER Header;

        /// <summary>
        /// Payload that contains the AUX_CLIENT_CONTROL, AUX_OSVERSIONINFO or AUX_EXORGINFO
        /// </summary>
        public byte[] Payload;
    }
    /// <summary>
    /// The structure of AUX_SERVER_CAPABILITIES.
    /// </summary>
    public struct AUX_SERVER_CAPABILITIES
    {
        /// <summary>
        /// A flag that indicates that the server supports specific capabilities.
        /// </summary>
        public uint ServerCapabilityFlags;
    }
}