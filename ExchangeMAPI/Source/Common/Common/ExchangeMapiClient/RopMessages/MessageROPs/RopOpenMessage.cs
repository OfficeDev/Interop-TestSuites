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
    /// TypedString structure, is used in certain ROPs in order to compact the string representation on the wire as much as possible.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct TypedString : ISerializable, IDeserializable
    {
        /// <summary>
        /// 8-bit enumeration.The possible values are specified in [MS-OXCDATA].
        /// </summary>
        public byte StringType;

        /// <summary>
        /// If the StringType field is set to 0x02, 0x03, or 0x04, then this field MUST be present and in the format specified by the Type field. 
        /// Otherwise, this field MUST NOT be present.
        /// </summary>
        public byte[] String;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            byte[] serializedBuffer = new byte[this.Size()];
            serializedBuffer[0] = this.StringType;
            Array.Copy(this.String, 0, serializedBuffer, 1, this.Size() - sizeof(byte));
            return serializedBuffer;
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
            int length = 0;
            this.StringType = ropBytes[index++];
            switch (this.StringType)
            {
                // There is no string represent
                case (byte)TestSuites.Common.StringType.None:

                // The string is empty
                case (byte)TestSuites.Common.StringType.Empty:
                    break;

                // Null-terminated 8-bit character string
                case (byte)TestSuites.Common.StringType.CharacterString:
                // Null-terminated Reduced Unicode character string
                case (byte)TestSuites.Common.StringType.ReducedUnicodeCharacterString:
                    while (true)
                    {
                        // Terminator found
                        if (ropBytes[index + length] == '\0')
                        {
                            ++length;
                            break;
                        }
                        else
                        {
                            ++length;
                        }
                    }

                    this.String = new byte[length];
                    Array.Copy(ropBytes, index, this.String, 0, length);
                    index += length;
                    break;

                // Null-terminated Unicode character string
                case (byte)TestSuites.Common.StringType.UnicodeCharacterString:
                    while (true)
                    {
                        // Terminator found
                        if ((ropBytes[index + length] == '\0') &&
                            (ropBytes[index + length + 1] == '\0'))
                        {
                            length += 2;
                            break;
                        }
                        else
                        {
                            length += 2;
                        }
                    }

                    this.String = new byte[length];
                    Array.Copy(ropBytes, index, this.String, 0, length);
                    index += length;
                    break;

                // Undefined flag
                default:
                    break;
            }

            return index - startIndex;
        }

        /// <summary>
        /// Return the size of TypedString  structure.
        /// </summary>
        /// <returns>The size of TypedString  structure.</returns>
        public int Size()
        {
            int size = 0;
            int length = 0;
            size += sizeof(byte);
            switch (this.StringType)
            {
                // There is no string represent
                case (byte)TestSuites.Common.StringType.None:
                // The string is empty
                case (byte)TestSuites.Common.StringType.Empty:
                    break;

                // Null-terminated 8-bit character string
                case (byte)TestSuites.Common.StringType.CharacterString:
                // Null-terminated Reduced Unicode character string
                case (byte)TestSuites.Common.StringType.ReducedUnicodeCharacterString:
                    while (true)
                    {
                        // Terminator found
                        if (this.String[length++] == '\0')
                        {
                            break;
                        }
                    }

                    size += length;
                    break;

                // Null-terminated Unicode character string
                case (byte)TestSuites.Common.StringType.UnicodeCharacterString:
                    while (true)
                    {
                        // Terminator found
                        if ((this.String[length++] == '\0') &&
                            (this.String[length++] == '\0'))
                        {
                            break;
                        }
                    }

                    size += length;
                    break;

                // Undefined flag
                default:
                    break;
            }

            return size;
        }
    }

    /// <summary>
    /// RopOpenMessage request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopOpenMessageRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x03.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table where the handle for the input Server Object is stored. 
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table where the handle for the output Server Object will be stored. 
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// This value specifies which code page will be used for string values associated with the message.
        /// </summary>
        public ushort CodePageId;

        /// <summary>
        /// . This value identifies the parent folder of the message to be opened.
        /// </summary>
        public ulong FolderId;

        /// <summary>
        /// These flags control the access to the message. The possible values are specified in [MS-OXCMSG].
        /// </summary>
        public byte OpenModeFlags;

        /// <summary>
        /// This value identifies the message to be opened.
        /// </summary>
        public ulong MessageId;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            IntPtr requestBuffer = new IntPtr();
            byte[] serializedBuffer = new byte[this.Size()];
            requestBuffer = Marshal.AllocHGlobal(this.Size());
            try
            {
                Marshal.StructureToPtr(this, requestBuffer, true);
                Marshal.Copy(requestBuffer, serializedBuffer, 0, this.Size());
                return serializedBuffer;
            }
            finally
            {
                Marshal.FreeHGlobal(requestBuffer);
            }
        }

        /// <summary>
        /// Return the size of RopOpenMessage request buffer structure.
        /// </summary>
        /// <returns>The size of RopOpenMessage request buffer structure.</returns>
        public int Size()
        {
            return Marshal.SizeOf(this);
        }
    }

    /// <summary>
    /// OpenRecipientRow structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct OpenRecipientRow : IDeserializable
    {
        /// <summary>
        /// 8-bit enumeration. The possible values for this enumeration are specified in [MS-OXCMSG]. 
        /// This enumeration specifies the type of recipient.
        /// </summary>
        public byte RecipientType;

        /// <summary>
        /// This value specifies the code page for the recipient.
        /// </summary>
        public ushort CodePageId;

        /// <summary>
        /// The server MUST set this field to 0x0000.
        /// </summary>
        public ushort Reserved;

        /// <summary>
        /// This value specifies the size of the RecipientRow field.
        /// </summary>
        public ushort RecipientRowSize;

        /// <summary>
        /// RecipientRow structure. The format of this structure is specified in [MS-OXCDATA]. 
        /// The size of this field, in bytes, is specified by the RecipientRowSize field.
        /// </summary>
        public RecipientRow RecipientRow;

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.RecipientType = ropBytes[index++];

            this.CodePageId = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += sizeof(ushort);
            this.Reserved = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += sizeof(ushort);
            this.RecipientRowSize = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += sizeof(ushort);
            this.RecipientRow = new RecipientRow();
            index += this.RecipientRow.Deserialize(ropBytes, index);   
            return index - startIndex;
        }
    }

    /// <summary>
    /// RecipientRow structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RecipientRow : ISerializable, IDeserializable
    {
        /// <summary>
        /// RecipientFlags structure. The format of this structure is defined in section 2.9.3.1. 
        /// This value specifies the type of recipient and which standard properties are included.
        /// </summary>
        public ushort RecipientFlags;

        /// <summary>
        /// This field MUST be present when the Type field of the RecipientFlags field is set to X500DN (0x1) and MUST NOT be present otherwise. 
        /// This value specifies the amount of the Address Prefix is used for this X500 DN.
        /// </summary>
        public byte AddressPrfixUsed;

        /// <summary>
        /// This field MUST be present when the Type field of the RecipientFlags field is set to X500DN (0x1) and MUST NOT be present otherwise.
        /// This value specifies the display type of this address.
        /// </summary>
        public byte DisplayType;

        /// <summary>
        /// Null-terminated ASCII string. This field MUST be present when the Type field of the RecipientFlags field is set to X500DN (0x1) and MUST NOT be present otherwise. 
        /// This value specifies the X500 DN of this recipient.
        /// </summary>
        public byte[] X500DN;

        /// <summary>
        /// This field MUST be present when the Type field of the RecipientFlags field is set to PersonalDistributionList1 (0x6) or PersonalDistributionList2 (0x7). 
        /// This field MUST NOT be present otherwise. This value specifies the size of the EntryID field.
        /// </summary>
        public ushort EntryIdSize;

        /// <summary>
        /// This field MUST be present when the Type field of the RecipientFlags field is set to PersonalDistributionList1 (0x6) or PersonalDistributionList2 (0x7). 
        /// This field MUST NOT be present otherwise. The number of bytes in this field MUST be the same as specified in the EntryIdSize field.
        /// This array specifies the address book EntryID of the distribution list.
        /// </summary>
        public byte[] EntryId;

        /// <summary>
        /// This field MUST be present when the Type field of the RecipientFlags field is set to PersonalDistributionList1 (0x6) or PersonalDistributionList2 (0x7).
        /// This field MUST NOT be present otherwise. This value specifies the size of the SearchKey field.
        /// </summary>
        public ushort SearchKeySize;

        /// <summary>
        /// This field is used when the Type field of the RecipientFlags field is set to PersonalDistributionList1 (0x6) or PersonalDistributionList2 (0x7). 
        /// This field MUST NOT be present otherwise. The number of bytes in this field MUST be the same as specified in the SearchKeySize field and can be 0. 
        /// This array specifies the Search Key of the distribution list.
        /// </summary>
        public byte[] SearchKey;
        
        /// <summary>
        /// Null-terminated ASCII string. This field MUST be present when the Type field of the RecipientsFlags field is set to NoType (0x0) and the O flag of the RecipientsFlags field is set. 
        /// This field MUST NOT be present otherwise. This string specifies the address type of the recipient.
        /// </summary>
        public byte[] AddressType;
        
        /// <summary>
        /// Null-terminated string. This field MUST be present when the E flag of the RecipientsFlags field is set and MUST NOT be present otherwise. 
        /// This field MUST be specified in Unicode characters if the U flag of the RecipientsFlags field is set and 8-bit character set otherwise. 
        /// This string specifies the Email Address of the recipient.
        /// </summary>
        public byte[] EmailAddress;
        
        /// <summary>
        /// Null-terminated string. This field MUST be present when the D flag of the RecipientsFlags field is set and MUST NOT be present otherwise. 
        /// This field MUST be specified in Unicode characters if the U flag of the RecipientsFlags field is set and 8-bit character set otherwise. 
        /// This string specifies the Email Address of the recipient.
        /// </summary>
        public byte[] DisplayName;

        /// <summary>
        /// Null-terminated string. This field MUST be present when the I flag of the RecipientsFlags field is set and MUST NOT be present otherwise. 
        /// This field MUST be specified in Unicode characters if the U flag of the RecipientsFlags field is set and 8-bit character set otherwise. 
        /// This string specifies the Email Address of the recipient.
        /// </summary>
        public byte[] SimpleDisplayName;

        /// <summary>
        /// Null-terminated string. This field MUST be present when the T flag of the RecipientsFlags field is set and MUST NOT be present otherwise. 
        /// This field MUST be specified in Unicode characters if the U flag of the RecipientsFlags field is set and 8-bit character set otherwise.
        /// This string specifies the Email Address of the recipient.
        /// </summary>
        public byte[] TransmittableDisplayName;

        /// <summary>
        /// This value specifies the number of columns from the RecipientColumns field that are included in RecipientProperties.
        /// </summary>
        public ushort RecipientColumnCount;

        /// <summary>
        /// PropertyRow structures.
        /// </summary>
        public PropertyRow RecipientProperties;

        /// <summary>
        /// Size of RecipientRow structure.
        /// </summary>
        private int size;

        /// <summary>
        /// The serialized request buffer of RecipientRow structure.
        /// </summary>
        private byte[] serializedBuffer;

        /// <summary>
        /// Size of RecipientRow structure.
        /// </summary>
        /// <returns>An integer represents the size of RecipientRow structure</returns>
        public int Size()
        {
            // For RecipientFlags
            this.size = sizeof(ushort);

            // For RecipientColumnCount
            this.size += sizeof(ushort);
            if ((ushort)TestSuites.Common.RecipientFlags.X500DN == (this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList2))
            {
                // For AddressPrefixUsed
                this.size++;

                // For DisplayType
                this.size++; 
            }

            if (((ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList1 ==
                (this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList2))
                || ((ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList2 ==
                (this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList2)))
            {
                // For EntryIdSize
                this.size += sizeof(ushort);

                // For SearchKeySize
                this.size += sizeof(ushort);
            }

            if (null != this.X500DN)
            {
                this.size += this.X500DN.Length;
            }

            if (null != this.EntryId)
            {
                this.size += this.EntryId.Length;
            }

            if (null != this.SearchKey)
            {
                this.size += this.SearchKey.Length;
            }

            if (null != this.AddressType)
            {
                this.size += this.AddressType.Length;
            }

            if (null != this.EmailAddress)
            {
                this.size += this.EmailAddress.Length;
            }

            if (null != this.DisplayName)
            {
                this.size += this.DisplayName.Length;
            }

            if (null != this.SimpleDisplayName)
            {
                this.size += this.SimpleDisplayName.Length;
            }

            if (null != this.TransmittableDisplayName)
            {
                this.size += this.TransmittableDisplayName.Length;
            }

            if (null != this.RecipientProperties)
            {
                // Flag of PropertyRow, one byte
                this.size++;

                if (null != this.RecipientProperties.PropertyValues)
                {
                    foreach (PropertyValue propertyValue in this.RecipientProperties.PropertyValues)
                    {
                        if (null != propertyValue.Value)
                        {
                            this.size += propertyValue.Value.Length;
                        }
                    }
                }
            }

            return this.size;
        }

        /// <summary>
        /// Serialize the request buffer.
        /// </summary>
        /// <returns>The serialized request buffer of RecipientRow structure.</returns>
        public byte[] Serialize()
        {
            this.serializedBuffer = new byte[this.Size()];
            int index = 0;

            Array.Copy(BitConverter.GetBytes((ushort)this.RecipientFlags), 0, this.serializedBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            if ((ushort)TestSuites.Common.RecipientFlags.X500DN == (this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList2))
            {
                this.serializedBuffer[index++] = this.AddressPrfixUsed;
                this.serializedBuffer[index++] = this.DisplayType;
            }

            if (null != this.X500DN)
            {
                Array.Copy(this.X500DN, 0, this.serializedBuffer, index, this.X500DN.Length);
                index += this.X500DN.Length;
            }

            if (((ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList1 ==
                (this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList2))
                || ((ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList2 ==
                (this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList2)))
            {
                Array.Copy(BitConverter.GetBytes((ushort)this.EntryIdSize), 0, this.serializedBuffer, index, sizeof(ushort));
                index += sizeof(ushort);
            }

            if (null != this.EntryId)
            {
                Array.Copy(this.EntryId, 0, this.serializedBuffer, index, this.EntryId.Length);
                index += this.EntryId.Length;
            }

            if (((ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList1 ==
                (this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList2))
                || ((ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList2 ==
                (this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList2)))
            {
                Array.Copy(BitConverter.GetBytes((ushort)this.SearchKeySize), 0, this.serializedBuffer, index, sizeof(ushort));
                index += sizeof(ushort);
            }

            if (null != this.SearchKey)
            {
                Array.Copy(this.SearchKey, 0, this.serializedBuffer, index, this.SearchKey.Length);
                index += this.SearchKey.Length;
            }

            if (null != this.AddressType)
            {
                Array.Copy(this.AddressType, 0, this.serializedBuffer, index, this.AddressType.Length);
                index += this.AddressType.Length;
            }

            if (null != this.EmailAddress)
            {
                Array.Copy(this.EmailAddress, 0, this.serializedBuffer, index, this.EmailAddress.Length);
                index += this.EmailAddress.Length;
            }

            if (null != this.DisplayName)
            {
                Array.Copy(this.DisplayName, 0, this.serializedBuffer, index, this.DisplayName.Length);
                index += this.DisplayName.Length;
            }

            if (null != this.SimpleDisplayName)
            {
                Array.Copy(this.SimpleDisplayName, 0, this.serializedBuffer, index, this.SimpleDisplayName.Length);
                index += this.SimpleDisplayName.Length;
            }

            if (null != this.TransmittableDisplayName)
            {
                Array.Copy(this.TransmittableDisplayName, 0, this.serializedBuffer, index, this.TransmittableDisplayName.Length);
                index += this.TransmittableDisplayName.Length;
            }

            Array.Copy(BitConverter.GetBytes((ushort)this.RecipientColumnCount), 0, this.serializedBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            if (null != this.RecipientProperties)
            {
                this.serializedBuffer[index++] = 0x00;
                foreach (PropertyValue propertyValue in this.RecipientProperties.PropertyValues)
                {
                    Array.Copy(propertyValue.Value, 0, this.serializedBuffer, index, propertyValue.Value.Length);
                    index += propertyValue.Value.Length;
                }
            }

            return this.serializedBuffer;
        }

        /// <summary>
        /// Deserialize the response buffer.
        /// </summary>
        /// <param name="ropBytes">Bytes in response.</param>
        /// <param name="startIndex">The start index of this structure.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.size = 0;

            this.RecipientFlags = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += sizeof(ushort);

            bool tflag, dflag, eflag, oflag, iflag, uflag;
            int recipientTye = 0;
            tflag = dflag = eflag = oflag = iflag = uflag = false;

            if ((this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.T) > 0)
            {
                tflag = true;
            }

            if ((this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.D) > 0)
            {
                dflag = true;
            }

            if ((this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.E) > 0)
            {
                eflag = true;
            }

            recipientTye = this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList2;
            if ((this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.O) > 0)
            {
                oflag = true;
            }

            if ((this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.I) > 0)
            {
                iflag = true;
            }

            if ((this.RecipientFlags & (ushort)TestSuites.Common.RecipientFlags.U) > 0)
            {
                uflag = true;
            }
          
            switch (recipientTye)
            {
                case (ushort)TestSuites.Common.RecipientFlags.X500DN:
                    this.AddressPrfixUsed = ropBytes[index++];
                    this.DisplayType = ropBytes[index++];

                    int bytesLen = 0;
                    bool isFound = false;

                    // Find the string with '\0' end
                    for (int i = index; i < ropBytes.Length; i++)
                    {
                        bytesLen++;
                        if (ropBytes[i] == 0)
                        {
                            isFound = true;
                            break;
                        }  
                    }

                    if (!isFound)
                    {
                        throw new ParseException("String too long or not found");
                    }
                    else
                    {
                        X500DN = new byte[bytesLen];
                        Array.Copy(ropBytes, index, X500DN, 0, bytesLen - 1);
                        index += bytesLen;
                    }

                    break;
                case (ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList1:
                case (ushort)TestSuites.Common.RecipientFlags.PersonalDistributionList2:
                    this.EntryIdSize = (ushort)BitConverter.ToInt16(ropBytes, index);
                    index += 2;
                    this.EntryId = new byte[this.EntryIdSize];
                    Array.Copy(ropBytes, index, this.EntryId, 0, this.EntryIdSize);
                    index += this.EntryIdSize;

                    this.SearchKeySize = (ushort)BitConverter.ToInt16(ropBytes, index);
                    index += 2;
                    this.SearchKey = new byte[this.SearchKeySize];
                    Array.Copy(ropBytes, index, this.SearchKey, 0, this.SearchKeySize);
                    index += this.SearchKeySize;

                    break;
                case (ushort)TestSuites.Common.RecipientFlags.None:
                    if (oflag == true)
                    {
                        bytesLen = 0;
                        isFound = false;

                        // Find the string with '\0' end
                        for (int i = index; i < ropBytes.Length; i++)
                        {
                            bytesLen++;
                            if (ropBytes[i] == 0)
                            {
                                isFound = true;
                                break;
                            }
                        }

                        if (!isFound)
                        {
                            throw new ParseException("String too long or not found");
                        }
                        else
                        {
                            this.AddressType = new byte[bytesLen];
                            Array.Copy(ropBytes, index, this.AddressType, 0, bytesLen - 1);
                            index += bytesLen;
                        } 
                    }

                    break;
            }

            if (eflag == true)
            {
                int bytesLen = 0;
                bool isFound = false;
                if (uflag == true)
                {
                    for (int i = index; i < ropBytes.Length; i += 2)
                    {
                        bytesLen += 2;
                        if ((ropBytes[i] == 0) && (ropBytes[i + 1] == 0))
                        {
                            isFound = true;
                            break;
                        }
                    }
                }
                else
                {
                    // Find the string with '\0' end
                    for (int i = index; i < ropBytes.Length; i++)
                    {
                        bytesLen++;
                        if (ropBytes[i] == 0)
                        {
                            isFound = true;
                            break;
                        }
                    }
                }

                if (!isFound)
                {
                    throw new ParseException("String too long or not found");
                }
                else
                {
                    this.EmailAddress = new byte[bytesLen];
                    Array.Copy(ropBytes, index, this.EmailAddress, 0, bytesLen - 1);
                    index += bytesLen;
                }  
            }

            if (dflag == true)
            {
                int bytesLen = 0;
                bool isFound = false;
                if (uflag == true)
                {
                    for (int i = index; i < ropBytes.Length; i += 2)
                    {
                        bytesLen += 2;
                        if ((ropBytes[i] == 0) && (ropBytes[i + 1] == 0))
                        {
                            isFound = true;
                            break;
                        }
                    }
                }
                else
                {
                    // Find the string with '\0' end
                    for (int i = index; i < ropBytes.Length; i++)
                    {
                        bytesLen++;
                        if (ropBytes[i] == 0)
                        {
                            isFound = true;
                            break;
                        }
                    }
                }

                if (!isFound)
                {
                    throw new ParseException("String too long or not found");
                }
                else
                {
                    if (uflag == true)
                    {
                        this.DisplayName = new byte[bytesLen];
                        Array.Copy(ropBytes, index, this.DisplayName, 0, bytesLen);
                        index += bytesLen;
                    }
                    else
                    {
                        this.DisplayName = new byte[bytesLen];
                        Array.Copy(ropBytes, index, this.DisplayName, 0, bytesLen);
                        index += bytesLen;
                    }
                }  
            }

            if (iflag == true)
            {
                int bytesLen = 0;
                bool isFound = false;
                if (uflag == true)
                {
                    for (int i = index; i < ropBytes.Length; i += 2)
                    {
                        bytesLen += 2;
                        if ((ropBytes[i] == 0) && (ropBytes[i + 1] == 0))
                        {
                            isFound = true;
                            break;
                        }
                    }
                }
                else
                {
                    // Find the string with '\0' end
                    for (int i = index; i < ropBytes.Length; i++)
                    {
                        bytesLen++;
                        if (ropBytes[i] == 0)
                        {
                            isFound = true;
                            break;
                        }
                    }
                }

                if (!isFound)
                {
                    throw new ParseException("String too long or not found");
                }
                else
                {
                    if (uflag == true)
                    {
                        this.SimpleDisplayName = new byte[bytesLen];
                        Array.Copy(ropBytes, index, this.SimpleDisplayName, 0, bytesLen);
                        index += bytesLen;
                    }
                    else
                    {
                        this.SimpleDisplayName = new byte[bytesLen];
                        Array.Copy(ropBytes, index, this.SimpleDisplayName, 0, bytesLen);
                        index += bytesLen;
                    }
                }   
            }

            if (tflag == true)
            {
                int bytesLen = 0;
                bool isFound = false;
                if (uflag == true)
                {
                    for (int i = index; i < ropBytes.Length; i += 2)
                    {
                        bytesLen += 2;
                        if ((ropBytes[i] == 0) && (ropBytes[i + 1] == 0))
                        {
                            isFound = true;
                            break;
                        }
                    }
                }
                else
                {
                    // Find the string with '\0' end
                    for (int i = index; i < ropBytes.Length; i++)
                    {
                        bytesLen++;
                        if (ropBytes[i] == 0)
                        {
                            isFound = true;
                            break;
                        }
                    }
                }

                if (!isFound)
                {
                    throw new ParseException("String too long or not found");
                }
                else
                {
                    if (uflag == true)
                    {
                        this.TransmittableDisplayName = new byte[bytesLen];
                        Array.Copy(ropBytes, index, this.TransmittableDisplayName, 0, bytesLen);
                        index += bytesLen;
                    }
                    else
                    {
                        this.TransmittableDisplayName = new byte[bytesLen];
                        Array.Copy(ropBytes, index, this.TransmittableDisplayName, 0, bytesLen);
                        index += bytesLen;
                    }
                } 
            }

            this.RecipientColumnCount = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += sizeof(ushort);
            Context.Instance.PropertyBytes = ropBytes;
            Context.Instance.CurIndex = index;
            this.RecipientProperties = new PropertyRow();
            this.RecipientProperties.Parse(Context.Instance);
            
            index = Context.Instance.CurIndex;
            this.size = index - startIndex;
            this.serializedBuffer = new byte[this.size];
            Array.Copy(ropBytes, startIndex, this.serializedBuffer, 0, this.size);

            return index - startIndex;
        }
    }

    /// <summary>
    /// RopOpenMessage response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopOpenMessageResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x03.
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
        /// 8-bit Boolean. This value specifies whether the message has named properties.
        /// </summary>
        public byte HasNamedProperties;

        /// <summary>
        /// TypedString structure. The format of the TypedString structure is specified in [MS-OXCDATA]. 
        /// This structure specifies the subject prefix of the message.
        /// </summary>
        public TypedString SubjectPrefix;

        /// <summary>
        /// TypedString structure. The format of the TypedString structure is specified in [MS-OXCDATA]. 
        /// This structure specifies the normalized subject of the message.
        /// </summary>
        public TypedString NormalizedSubject;

        /// <summary>
        /// This value specifies the number of recipients on the message.
        /// </summary>
        public ushort RecipientCount;

        /// <summary>
        /// This value specifies the number of structures in the RecipientColumns field.
        /// </summary>
        public ushort ColumnCount;

        /// <summary>
        /// Array of PropertyTag structures. The number of structures contained in this field is specified by the ColumnCount field. 
        /// The format of the PropertyTag structure is specified in [MS-OXCDATA]. 
        /// This field specifies the property values that can be included in each row that is specified in the RecipientRows field.
        /// </summary>
        public PropertyTag[] RecipientColumns;

        /// <summary>
        /// This value specifies the number of structures in the RecipientRows field.
        /// </summary>
        public byte RowCount;

        /// <summary>
        /// List of OpenRecipientRow structures. The number of structures contained in this field is specified by the RowCount field. 
        /// </summary>
        public OpenRecipientRow[] RecipientRows;

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
                this.HasNamedProperties = ropBytes[index++];
                index += this.SubjectPrefix.Deserialize(ropBytes, index);
                index += this.NormalizedSubject.Deserialize(ropBytes, index);
                this.RecipientCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += sizeof(ushort);
                this.ColumnCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += sizeof(ushort);

                // RecipientColumns
                if (this.ColumnCount > 0)
                {
                    this.RecipientColumns = new PropertyTag[this.ColumnCount];
                    Context.Instance.Init();
                    for (int i = 0; i < this.ColumnCount; i++)
                    {
                        index += this.RecipientColumns[i].Deserialize(ropBytes, index);
                        Context.Instance.Properties.Add(new Property((PropertyType)this.RecipientColumns[i].PropertyType));
                    }       
                }

                this.RowCount = ropBytes[index++];
                if (this.RowCount > 0)
                {
                    this.RecipientRows = new OpenRecipientRow[this.RowCount];
                    for (int i = 0; i < this.RowCount; i++)
                    {
                        index += this.RecipientRows[i].Deserialize(ropBytes, index);
                    }
                }
            }

            return index - startIndex;
        }
    }
}