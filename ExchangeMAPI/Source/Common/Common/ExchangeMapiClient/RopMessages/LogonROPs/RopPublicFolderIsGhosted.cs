//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;
    using System.Text;

    /// <summary>
    /// RopPublicFolderIsGhosted request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopPublicFolderIsGhostedRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x45.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// Unsigned 8-bit integer. This index specifies the location in the Server Object Handle Table 
        /// where the handle for the input Server Object is stored.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// 64-bit identifier. This value identifies the folder to check.
        /// </summary>
        public ulong FolderId;

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
    /// RopPublicFolderIsGhosted response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopPublicFolderIsGhostedResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x45.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation. 
        /// For this response, this field is set to 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// 8-bit Boolean. This value specifies whether the folder is a ghosted folder.
        /// </summary>
        public byte IsGhosted;

        /// <summary>
        /// Unsigned 16-bit integer. This field is present if IsGhosted is non-zero and is not present if IsGhosted is zero. 
        /// This value specifies the number of strings in the Servers field.
        /// </summary>
        public ushort? ServersCount;

        /// <summary>
        /// Unsigned 16-bit integer. This field is present if IsGhosted is non-zero and is not present if IsGhosted is zero. 
        /// This value specifies the number of strings in Servers that refer to lowest cost servers.
        /// </summary>
        public ushort? CheapServersCount;

        /// <summary>
        /// List of null-terminated ASCII strings. This field is present if IsGhosted is non-zero and is 
        /// not present if IsGhosted is zero. The number of strings contained in this field is specified by the ServersCount field. 
        /// These strings specify which servers have replicas of this folder.
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
            int index = startIndex;
            this.RopId = ropBytes[index++];
            this.InputHandleIndex = ropBytes[index++];
            this.ReturnValue = (uint)BitConverter.ToInt32(ropBytes, index);
            index += sizeof(uint);

            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.IsGhosted = ropBytes[index++];
                if (this.IsGhosted != 0)
                {
                    this.ServersCount = BitConverter.ToUInt16(ropBytes, index);
                    index += sizeof(ushort);
                    this.CheapServersCount = BitConverter.ToUInt16(ropBytes, index);
                    index += sizeof(ushort);
                    if (this.ServersCount > 0)
                    {
                        this.Servers = new string[(ushort)this.ServersCount];
                        for (int i = 0; i < this.ServersCount; i++)
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