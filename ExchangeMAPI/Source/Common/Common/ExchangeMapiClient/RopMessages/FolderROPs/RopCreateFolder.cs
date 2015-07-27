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
    /// RopCreateFolder request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopCreateFolderRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x1C.
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
        /// The possible values for this enumeration are specified in [MS-OXCFOLD]. 
        /// This value specifies what type of folder to create.
        /// </summary>
        public byte FolderType;

        /// <summary>
        /// This value specifies whether the DisplayName and Comment are specified in Unicode or ASCII.
        /// </summary>
        public byte UseUnicodeStrings;

        /// <summary>
        /// This value specifies whether this operation opens or fails when a folder already exists.
        /// </summary>
        public byte OpenExisting;

        /// <summary>
        /// Reserved. This field MUST be set to 0x00.
        /// </summary>
        public byte Reserved;

        /// <summary>
        /// Null-terminated string. This value specifies the name of the created folder. 
        /// If UseUnicodeStrings is non-zero, the string is composed of Unicode characters. 
        /// If UseUnicodeStrings is zero, the string is composed of ASCII characters.
        /// </summary>
        public byte[] DisplayName;

        /// <summary>
        /// Null-terminated string. This value specifies the folder comment that is associated with the created folder. 
        /// If UseUnicodeStrings is non-zero, the string is composed of Unicode characters. 
        /// If UseUnicodeStrings is zero, the string is composed of ASCII characters.
        /// </summary>
        public byte[] Comment;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            int index = 0;
            int bufSize = sizeof(byte) * 8;
            if (this.DisplayName != null)
            {
                bufSize += this.DisplayName.Length;
            }

            if (this.Comment != null)
            {
                bufSize += this.Comment.Length;
            }

            byte[] serializeBuffer = new byte[bufSize];
            serializeBuffer[index++] = this.RopId;
            serializeBuffer[index++] = this.LogonId;
            serializeBuffer[index++] = this.InputHandleIndex;
            serializeBuffer[index++] = this.OutputHandleIndex;
            serializeBuffer[index++] = this.FolderType;
            serializeBuffer[index++] = this.UseUnicodeStrings;
            serializeBuffer[index++] = this.OpenExisting;
            serializeBuffer[index++] = this.Reserved;
            if (this.DisplayName != null)
            {
                Array.Copy(this.DisplayName, 0, serializeBuffer, index, this.DisplayName.Length);
                index += this.DisplayName.Length;
            }

            if (this.Comment != null)
            {
                Array.Copy(this.Comment, 0, serializeBuffer, index, this.Comment.Length);
                index += this.Comment.Length;
            }
                                                                                                                                                                                                
            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            int size = sizeof(byte) * 8;
            if (this.DisplayName != null)
            {
                size += this.DisplayName.Length;
            }

            if (this.Comment != null)
            {
                size += this.Comment.Length;
            }
                                                                                                                                                                                                
            return size;
        }
    }

    /// <summary>
    /// RopCreateFolder response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopCreateFolderResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x1C.
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
        /// This value identifies the folder created or opened.
        /// </summary>
        public ulong FolderId;

        /// <summary>
        /// This value indicates whether an existing folder was opened or a new folder was created.
        /// </summary>
        public byte IsExistingFolder;

        /// <summary>
        /// This field is present if the IsExistingFolder field is non-zero and is not present otherwise. 
        /// This value indicates whether there are rules associated with the folder.
        /// </summary>
        public byte? HasRules;

        /// <summary>
        /// This field is present if the IsExistingFolder field is non-zero and is not present otherwise. 
        /// This value indicates whether the server is an active replica of this folder.
        /// </summary>
        public byte? IsGhosted;

        /// <summary>
        /// This field is present if both IsExistingFolder and IsGhosted are non-zero and is not present otherwise. 
        /// This value specifies the number of strings in the Servers field.
        /// </summary>
        public ushort? ServerCount;

        /// <summary>
        /// This field is present if both IsExistingFolder and IsGhosted are non-zero and is not present otherwise. 
        /// This value specifies the number of values in Servers that refers to lowest cost servers.
        /// </summary>
        public ushort? CheapServerCount;

        /// <summary>
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
            // Because HasRules is byte type, has default value 0, can't check it whether present
            // Set HasRules to null to decide it present, if equal 0xFF means not present else present
            this.HasRules = null;
            int index = startIndex;
            this.RopId = ropBytes[index++];
            this.OutputHandleIndex = ropBytes[index++];
            this.ReturnValue = (uint)BitConverter.ToInt32(ropBytes, index);
            index += sizeof(uint);
                                                                                                                                                                             
            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.FolderId = (ulong)BitConverter.ToInt64(ropBytes, index);
                index += sizeof(ulong);
                this.IsExistingFolder = ropBytes[index++];
                if (this.IsExistingFolder != 0)
                {
                    this.HasRules = ropBytes[index++];
                    this.IsGhosted = ropBytes[index++];
                    if (this.IsGhosted != 0)
                    {
                        this.ServerCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                        index += sizeof(ushort);
                        this.CheapServerCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                        index += sizeof(ushort);
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
            }
                                                                                                                                                                                                
            return index - startIndex;
        }
    }
}