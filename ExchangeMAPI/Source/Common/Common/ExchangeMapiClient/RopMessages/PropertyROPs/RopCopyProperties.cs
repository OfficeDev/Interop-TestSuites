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

    /// <summary>
    /// RopCopyProperties request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopCopyPropertiesRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x67.
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
        /// 8-bit flags structure. The possible values are specified in [MS-OXCPRPT]. These flags control the operation behavior.
        /// </summary>
        public byte CopyFlags;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies how many tags are present in the PropertyTags field.
        /// </summary>
        public ushort PropertyTagCount;

        /// <summary>
        /// Array of PropertyTag structures. The number of structures contained in this field is specified 
        /// by the PropertyTagCount field. The format of the PropertyTag structure is specified in [MS-OXCDATA]. 
        /// This field specifies the properties to copy.
        /// </summary>
        public PropertyTag[] PropertyTags;

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
            serializeBuffer[index++] = this.CopyFlags;

            Array.Copy(BitConverter.GetBytes((ushort)this.PropertyTagCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            for (int i = 0; i < this.PropertyTagCount; i++)
            {
                Array.Copy(this.PropertyTags[i].Serialize(), 0, serializeBuffer, index, this.PropertyTags[i].Size());
                index += this.PropertyTags[i].Size();
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 8 indicates sizeof(byte) * 6 + sizeof(UInt16)
            int size = sizeof(byte) * 8;

            for (int i = 0; i < this.PropertyTagCount; i++)
            {
                size += this.PropertyTags[i].Size();
            }

            return size;
        }
    }

    /// <summary>
    /// RopCopyProperties response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopCopyPropertiesResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x67.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index MUST be set to the SourceHandleIndex value specified in the request.
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
        /// Unsigned 32-bit integer. Only available for failure response. This index MUST be set to the DestHandleIndex specified in the request.
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