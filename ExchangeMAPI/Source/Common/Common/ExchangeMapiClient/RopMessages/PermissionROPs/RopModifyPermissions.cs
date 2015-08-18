namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// PermissionData request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct PermissionData : ISerializable
    {
        /// <summary>
        /// This field is used to specify the type of operation.
        /// </summary>
        public byte PermissionDataFlags;

        /// <summary>
        /// This value specifies the number of value structures in the PropertyValues field.
        /// </summary>
        public ushort PropertyValueCount;

        /// <summary>
        /// The number of structures contained in this field is specified by the PropertyValueCount field. 
        /// These values are used to specify and describe the modify operations.
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
            serializeBuffer[index++] = this.PermissionDataFlags;
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
    /// RopModifyPermissionsRequest request buffer structure.
    /// </summary>
    public struct RopModifyPermissionsRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x40.
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
        /// The possible values are specified in [MS-OXCPERM]. These flags control behavior of this operation.
        /// </summary>
        public byte ModifyFlags;

        /// <summary>
        /// This count specifies the number of structures serialized in the PermissionData array.
        /// </summary>
        public ushort ModifyCount;

        /// <summary>
        /// A list of PermissionData structures. 
        /// The number of structures contained in this field is specified by the ModifyCount field. 
        /// </summary>
        public PermissionData[] PermissionsData;

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
            serializeBuffer[index++] = this.ModifyFlags;

            Array.Copy(BitConverter.GetBytes((ushort)this.ModifyCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            for (int i = 0; i < this.ModifyCount; i++)
            {
                Array.Copy(this.PermissionsData[i].Serialize(), 0, serializeBuffer, index, this.PermissionsData[i].Size());
                index += this.PermissionsData[i].Size();
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

            for (int i = 0; i < this.ModifyCount; i++)
            {
                size += this.PermissionsData[i].Size();
            }

            return size;
        }
    }

    /// <summary>
    /// RopModifyPermissionsRequest response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopModifyPermissionsResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x40.
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
                this = (RopModifyPermissionsResponse)Marshal.PtrToStructure(responseBuffer, typeof(RopModifyPermissionsResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}