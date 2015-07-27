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
    /// RuleData structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RuleData : ISerializable
    {
        /// <summary>
        /// 8-bit flag structure. This field is used to specify the type of operation.
        /// </summary>
        public byte RuleDataFlags;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the number of value structures in the PropertyValues field.
        /// </summary>
        public ushort PropertyValueCount;

        /// <summary>
        /// Array of PropertyValue structures. The number of structures contained in this field is specified by the PropertyValueCount field. The format of the PropertyValue structure is specified in [MS-OXCDATA]. These values are used to specify and describe the rule operations.
        /// </summary>
        public PropertyValue[] PropertyValues;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            int index = 0;
            byte[] serializeBuffer = new byte[this.Size()];
            serializeBuffer[index++] = this.RuleDataFlags;
            Array.Copy(BitConverter.GetBytes((ushort)this.PropertyValueCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            for (int i = 0; i < this.PropertyValueCount; i++)
            {
                if (this.PropertyValues[i].Value != null)
                {
                    Array.Copy(this.PropertyValues[i].Serialize(), 0, serializeBuffer, index, this.PropertyValues[i].Size());
                    index += this.PropertyValues[i].Size();
                }
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 3 indicates sizeof(byte) + sizeof(UInt16)
            int size = sizeof(byte) * 3;
            for (int i = 0; i < this.PropertyValueCount; i++)
            {
                if (this.PropertyValues[i].Value != null)
                {
                    size += this.PropertyValues[i].Size();
                }
            }

            return size;
        }
    }

    /// <summary>
    /// RopModifyRules request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopModifyRulesRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x41.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the logon on which the operation is performed.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// Unsigned 8-bit integer. This index refers to the handle in the Server Object Handle 
        /// Table used as input for this operation.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// 8-bit flags structure. The possible values are specified in [MS-OXORULE]. 
        /// These flags specify behavior of this operation.
        /// </summary>
        public byte ModifyRulesFlags;

        /// <summary>
        /// Unsigned 16-bit integer. This count specifies the number of structures serialized in the RuleData field.
        /// </summary>
        public ushort RulesCount;

        /// <summary>
        /// A list of RuleData structures. The number of structures contained in this field is specified by the RulesCount field.
        /// </summary>
        public RuleData[] RulesData;

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
            serializeBuffer[index++] = this.ModifyRulesFlags;

            Array.Copy(BitConverter.GetBytes((ushort)this.RulesCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            for (int i = 0; i < this.RulesCount; i++)
            {
                Array.Copy(this.RulesData[i].Serialize(), 0, serializeBuffer, index, this.RulesData[i].Size());
                index += this.RulesData[i].Size();
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 6 indicates sizeof(byte) * 4 + sizeof(UInt16)
            int size = sizeof(byte) * 6;

            for (int i = 0; i < this.RulesCount; i++)
            {
                size += this.RulesData[i].Size();
            }

            return size;
        }
    }

    /// <summary>
    /// RopModifyRules response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopModifyRulesResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x41.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index refers to the handle in the Server Object Handle 
        /// Table used as input for this operation.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation.
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
                this = (RopModifyRulesResponse)Marshal.PtrToStructure(responseBuffer, typeof(RopModifyRulesResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}