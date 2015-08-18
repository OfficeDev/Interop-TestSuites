namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;
    using System.Text;

    /// <summary>
    /// RopOpenFolder request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopOpenFolderRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x02.
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
        /// where the handle for the output Server Object will be stored.
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// This identifier specifies the folder to be opened.
        /// </summary>
        public ulong FolderId;

        /// <summary>
        /// These flags are used to control how the folder is opened. The possible flags values are specified in [MS-OXCFOLD].
        /// </summary>
        public byte OpenModeFlags;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            byte[] serializeBuffer = new byte[Marshal.SizeOf(this)];

            // Get the request buffer.
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
    /// RopOpenFolder response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopOpenFolderResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x02.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the OutputHandleIndex specified in the request.
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation. For successful response, this field is set to 0x00000000.
        /// For failure response, this field is set to a value other than 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// This value indicates whether there are rules associated with the folder.
        /// </summary>
        public byte HasRules;

        /// <summary>
        /// This value specifies whether the folder is a ghosted folder.
        /// </summary>
        public byte IsGhosted;

        /// <summary>
        /// This field is present if IsGhosted is non-zero and is not present if IsGhosted is zero. 
        /// This value specifies the number of strings in the Servers field.
        /// </summary>
        public ushort? ServerCount;

        /// <summary>
        /// This field is present if IsGhosted is non-zero and is not present if IsGhosted is zero. 
        /// This value specifies the number of values in Servers that refers to lowest cost servers.
        /// </summary>
        public ushort? CheapServerCount;

        /// <summary>
        /// This field is present if IsGhosted is non-zero and is not present if IsGhosted is zero. 
        /// The number of strings contained in this field is specified by the ServerCount field. These strings specify which servers have replicas of this folder.
        /// </summary>
        public string[] Servers;

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            // Set null as default value.
            this.HasRules = 0;
            this.IsGhosted = 0;
            this.ServerCount = null;
            this.CheapServerCount = null;

            int index = startIndex;
            this.RopId = ropBytes[index++];
            this.OutputHandleIndex = ropBytes[index++];
            this.ReturnValue = (uint)BitConverter.ToInt32(ropBytes, index);
            index += 4;
                                                                                                                                                                             
            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.HasRules = ropBytes[index++];
                this.IsGhosted = ropBytes[index++];
                if (this.IsGhosted != 0)
                {
                    this.ServerCount = BitConverter.ToUInt16(ropBytes, index);
                    index += 2;
                    this.CheapServerCount = BitConverter.ToUInt16(ropBytes, index);
                    index += 2;
                    if (this.ServerCount > 0)
                    {
                        this.Servers = new string[(ushort)this.ServerCount];
                        for (int i = 0; i < this.ServerCount; i++)
                        {
                            int bytesLen = 0;
                                                                                                                                                                             
                            // Find the string with '\0' end
                            for (int j = index; j < ropBytes.Length; j++)
                            {
                                bytesLen++;
                                if (ropBytes[j] == 0)
                                {
                                    break;
                                }
                            }

                            this.Servers[i] = Encoding.ASCII.GetString(ropBytes, index, bytesLen);
                            index += bytesLen;
                        }
                    }
                }
            }
                                                                                                                                                                                                
            return index - startIndex;
        }
    }
}