//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System;
    using System.Runtime.InteropServices;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Address book EntryIDs can represent several types of Address Book objects including individual users, 
    /// distribution lists, containers, and templates as specified in table 4.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct AddressBookEntryID : ISerializable
    {
        /// <summary>
        /// A 4 bytes indicates a long-term EntryID, this value must be 0x00000000.
        /// </summary>
        public uint Flags;

        /// <summary>
        /// The value of ProviderUID
        /// </summary>
        public byte[] ProviderUID;

        /// <summary>
        /// The version for address book entry id,MUST be set to %x01.00.00.00. 
        /// </summary>
        public byte[] Version;

        /// <summary>
        /// A 32-bit integer representing the type of the object.
        /// </summary>
        public uint Type;

        /// <summary>
        /// The X500 DN of the Address Book object. X500DN is a null-terminated string of 8-bit characters. 
        /// </summary>
        public string X500DN;

        /// <summary>
        /// Follows X500DN
        /// </summary>
        public string Domain;

        /// <summary>
        /// The size of ProviderUID field.
        /// </summary>
        private const int ProviderUIDSize = 0x0010;

        /// <summary>
        /// The size of VersionSize field.
        /// </summary>
        private const int VersionSize = 0x0004;

        /// <summary>
        /// The size of TypeSize field.
        /// </summary>
        private const int TypeSize = 0x0004;

        /// <summary>
        /// The serialized AddressBookEntryID structure.
        /// </summary>
        private byte[] result;

        /// <summary>
        /// Size of AddressBookEntryID structure.
        /// </summary>
        private int size;

        /// <summary>
        /// Deserialize the response buffer.
        /// </summary>
        /// <param name="ropBytes">Bytes in response.</param>
        /// <param name="startIndex">The start index of this structure.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex + 2;
            this.Flags = BitConverter.ToUInt32(ropBytes, index);
            index = index + TypeSize;

            this.ProviderUID = new byte[ProviderUIDSize];
            Array.Copy(ropBytes, index, this.ProviderUID, 0, ProviderUIDSize);
            index = index + ProviderUIDSize;

            this.Version = new byte[VersionSize];
            Array.Copy(ropBytes, index, this.Version, 0, VersionSize);
            index = index + VersionSize;

            this.Type = BitConverter.ToUInt32(ropBytes, index);
            index += TypeSize;

            if (index >= ropBytes.Length)
            {
                throw new ParseException("The X500DN is not found from response buffer. The index was outside the bounds of the response buffer.");
            }

            this.X500DN = this.GetAddressValue(ropBytes, ref index);
            return 0;
        }

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The serialized ROP request buffer.</returns>
        public byte[] Serialize()
        {
            this.result = new byte[this.Size()];
            int index = 0;

            Array.Copy(new byte[] { 0x00, 0x00, 0x00, 0x00 }, 0, this.result, index, 4);
            index += 4;
            Array.Copy(this.ProviderUID, 0, this.result, index, 16);
            index += 16;
            Array.Copy(this.Version, 0, this.result, index, 4);
            index += 4;
            Array.Copy(BitConverter.GetBytes(this.Type), 0, this.result, index, sizeof(uint));
            index += 4;
            if (!string.IsNullOrEmpty(this.X500DN))
            {
                byte[] values = System.Text.Encoding.Default.GetBytes(this.X500DN);
                Array.Copy(values, 0, this.result, index, values.Length);
            }

            return this.result;
        }

        /// <summary>
        /// Size of address book EntryId structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // For address book EntryId Flags
            this.size = 0;

            // For address book EntryId ProviderUID
            this.size += ProviderUIDSize;

            // For address book EntryId Version
            this.size += VersionSize;

            // For address book EntryId Type
            this.size += TypeSize;
            if (null != this.X500DN)
            {
                this.size += this.X500DN.Length;
            }

            return this.size;
        }

        /// <summary>
        /// Indicates whether this instance and a specific object are equals
        /// </summary>
        /// <param name="obj">The object that compare with this instance.</param>
        /// <returns>A Boolean value indicates whether this instance and a specific object are equals.</returns>
        public override bool Equals(object obj)
        {
            if (obj.GetType() != typeof(AddressBookEntryID))
            {
                return false;
            }

            AddressBookEntryID addrObj = (AddressBookEntryID)obj;

            if (addrObj.Type != this.Type)
            {
                return false;
            }

            if (addrObj.X500DN != this.X500DN)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Return the hash code of this instance.
        /// </summary>
        /// <returns>The hash code of this instance </returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        /// <summary>
        /// Get Address Value in the byte array values
        /// </summary>
        /// <param name="values">the byte array which contains the value</param>
        /// <param name="index">The start index of values</param>
        /// <returns>The value of Address</returns>
        private string GetAddressValue(byte[] values, ref int index)
        {
            string res = string.Empty;
            int i = Array.IndexOf<byte>(values, 0x00, index);

            if (i == -1)
            {
                throw new ParseException("The X500DN is not found from response buffer. The index was outside the bounds of the response buffer.");
            }

            res = Encoding.ASCII.GetString(values, index, i + 1 - index);
            index = i + 1;

            return res;
        }
    }
}