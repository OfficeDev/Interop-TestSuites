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
    /// A property tag both identifies a property and gives the data type its value. 
    /// </summary>
    public struct PropertyTag : IDeserializable, ISerializable
    {
        /// <summary>
        /// 16-bit unsigned integer that identifies the data type of the property value.
        /// </summary>
        public ushort PropertyType;

        /// <summary>
        /// A 16-bit unsigned integer that identifies the property.
        /// </summary>
        public ushort PropertyId;

        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyTag" /> structure.
        /// </summary>
        /// <param name="propertyId">Property id</param>
        /// <param name="propertyType">Property type</param>
        public PropertyTag(ushort propertyId, ushort propertyType)
        {
            this.PropertyId = propertyId;
            this.PropertyType = propertyType;
        }

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.PropertyType = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += sizeof(ushort);
            this.PropertyId = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += sizeof(ushort);
            return index - startIndex;
        }

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            int index = 0;
            byte[] serializeBuffer = new byte[this.Size()];
            Array.Copy(BitConverter.GetBytes((ushort)this.PropertyType), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            Array.Copy(BitConverter.GetBytes((ushort)this.PropertyId), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 4 indicates sizeof(UInt16) * 2
            int size = sizeof(byte) * 4;
            return size;
        }
    }

    /// <summary>
    /// RopSetColumns request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSetColumnsRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x12.
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
        /// 8-bit flags structure. The possible values are specified in [MS-OXCTABL]. These flags control this operation.
        /// </summary>
        public byte SetColumnsFlags;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies how many tags are present in PropertyTags.
        /// </summary>
        public ushort PropertyTagCount;

        /// <summary>
        /// Array of PropertyTag structures. This field specifies the property values that are visible in table rows. 
        /// The number of structures contained in this field is specified by the PropertyTagCount field. 
        /// The format of the PropertyTag structure is specified in [MS-OXCDATA].
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
            serializeBuffer[index++] = this.InputHandleIndex;
            serializeBuffer[index++] = this.SetColumnsFlags;
            Array.Copy(BitConverter.GetBytes((short)this.PropertyTagCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            if (this.PropertyTagCount > 0)
            {
                IntPtr requestBuffer = new IntPtr();
                requestBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(PropertyTag)));
                try
                {
                    Context.Instance.Init();
                    foreach (PropertyTag propTag in this.PropertyTags)
                    {
                        Marshal.StructureToPtr(propTag, requestBuffer, true);
                        Marshal.Copy(requestBuffer, serializeBuffer, index, Marshal.SizeOf(typeof(PropertyTag)));
                        index += Marshal.SizeOf(typeof(PropertyTag));
   
                        // Insert properties into Context
                        Context.Instance.Properties.Add(new Property((PropertyType)propTag.PropertyType));
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(requestBuffer);
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
            int size = (sizeof(byte) * 4) + sizeof(ushort);
            if (this.PropertyTagCount > 0)
            {
                size += this.PropertyTags.Length * Marshal.SizeOf(typeof(PropertyTag));
            }

            return size;
        }
    }

    /// <summary>
    /// RopSetColumns response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSetColumnsResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x12.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation. 
        /// For success response, this field is set to 0x00000000.
        /// For failure response, this field is set to a value other than 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// 8-bit enumeration. The possible values for this enumeration are specified in [MS-OXCTABL]. 
        /// This value specifies the status of the table.
        /// </summary>
        public byte TableStatus;

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
                this.TableStatus = ropBytes[index++];
            }
            else
            {
                this.TableStatus = 0;
            }

            return index - startIndex;
        }
    }
}