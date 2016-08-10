namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// This define the base class for a property value node
    /// </summary>
    public class AddressBookPropertyValue : Node
    {
        /// <summary>
        /// Property's value
        /// </summary>
        private byte[] value;

        /// <summary>
        /// Indicates whether the property's value is an unfixed size or not.
        /// </summary>
        private bool isVariableSize = false;

        /// <summary>
        /// Gets or sets property's value
        /// </summary>
        public byte[] Value
        {
            get { return this.value; }
            set { this.value = value; }
        }

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public virtual byte[] Serialize()
        {
            int length = this.Size();
            byte[] resultBytes = new byte[length];

            if (this.isVariableSize)
            {
                // Fill 2 bytes with length
                resultBytes[0] = (byte)((ushort)this.value.Length & 0x00FF);
                resultBytes[1] = (byte)(((ushort)this.value.Length & 0xFF00) >> 8);
            }

            Array.Copy(this.value, 0, resultBytes, this.isVariableSize == false ? 0 : 2, this.value.Length);

            return resultBytes;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public virtual int Size()
        {
            return this.isVariableSize == false ? this.value.Length : this.value.Length + 2;
        }

        /// <summary>
        /// Parse bytes in context into a PropertyValueNode
        /// </summary>
        /// <param name="context">The value of Context</param>
        public override void Parse(Context context)
        {
            // Current processing property's Type
            PropertyType type = context.CurProperty.Type;
            int strBytesLen = 0;
            bool isFound = false;

            switch (type)
            {
                // 1 Byte
                case PropertyType.PtypBoolean:
                    if (context.AvailBytes() < sizeof(byte))
                    {
                        throw new ParseException("Not well formed PtypIBoolean");
                    }
                    else
                    {
                        this.value = new byte[sizeof(byte)];
                        Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, sizeof(byte));
                        context.CurIndex += sizeof(byte);
                    }

                    break;
                case PropertyType.PtypInteger16:
                    if (context.AvailBytes() < sizeof(short))
                    {
                        throw new ParseException("Not well formed PtypInteger16");
                    }
                    else
                    {
                        this.value = new byte[sizeof(short)];
                        Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, sizeof(short));
                        context.CurIndex += sizeof(short);
                    }

                    break;
                case PropertyType.PtypInteger32:
                case PropertyType.PtypFloating32:
                case PropertyType.PtypErrorCode:
                    if (context.AvailBytes() < sizeof(int))
                    {
                        throw new ParseException("Not well formed PtypInteger32");
                    }
                    else
                    {
                        // Assign value of int to property's value
                        this.value = new byte[sizeof(int)];
                        Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, sizeof(int));
                        context.CurIndex += sizeof(int);
                    }

                    break;
                case PropertyType.PtypFloating64:
                case PropertyType.PtypCurrency:
                case PropertyType.PtypFloatingTime:
                case PropertyType.PtypInteger64:
                case PropertyType.PtypTime:
                    if (context.AvailBytes() < sizeof(long))
                    {
                        throw new ParseException("Not well formed PtypInteger64");
                    }
                    else
                    {
                        // Assign value of Int64 to property's value
                        this.value = new byte[sizeof(long)];
                        Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, sizeof(long));
                        context.CurIndex += sizeof(long);
                    }

                    break;
                case PropertyType.PtypGuid:
                    if (context.AvailBytes() < sizeof(byte) * 16)
                    {
                        throw new ParseException("Not well formed 16 PtypGuid");
                    }
                    else
                    {
                        // Assign value of Int64 to property's value
                        this.value = new byte[sizeof(byte) * 16];
                        Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, sizeof(byte) * 16);
                        context.CurIndex += sizeof(byte) * 16;
                    }

                    break;
                case PropertyType.PtypBinary:
                    if (context.AvailBytes() < sizeof(uint))
                    {
                        throw new ParseException("Not well formed PtypBinary");
                    }
                    else
                    {
                        // Property start with "FF"
                        if (context.PropertyBytes[context.CurIndex] == (byte)0xFF)
                        {
                            context.CurIndex++;
                        }

                        // First parse the count of the binary bytes
                        int bytesCount = BitConverter.ToInt32(context.PropertyBytes, context.CurIndex);
                        this.value = new byte[sizeof(int) + bytesCount];
                        Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, sizeof(int));
                        context.CurIndex += sizeof(int);

                        // Then parse the binary bytes.
                        if (bytesCount == 0)
                        {
                            this.value = null;
                        }
                        else
                        {
                            Array.Copy(context.PropertyBytes, context.CurIndex, this.value, sizeof(int), bytesCount);
                            context.CurIndex += bytesCount;
                        }
                    }

                    break;
                case PropertyType.PtypMultipleInteger16:
                    if (context.AvailBytes() < sizeof(short))
                    {
                        throw new ParseException("Not well formed PtypMultipleInterger");
                    }
                    else
                    {
                        short bytesCount = BitConverter.ToInt16(context.PropertyBytes, context.CurIndex);
                        this.value = new byte[sizeof(short) + bytesCount * sizeof(short)];
                        Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, sizeof(short));
                        context.CurIndex += sizeof(short);
                        if (bytesCount == 0)
                        {
                            this.value = null;
                        }
                        else
                        {
                            Array.Copy(context.PropertyBytes, context.CurIndex, this.value, sizeof(short), bytesCount * sizeof(short));
                            context.CurIndex += bytesCount * sizeof(short);
                        }
                    }

                    break;
                case PropertyType.PtypMultipleInteger32:
                case PropertyType.PtypMultipleFloating32:
                    if (context.AvailBytes() < sizeof(short))
                    {
                        throw new ParseException("Not well formed PtypMultipleInterger");
                    }
                    else
                    {
                        short bytesCount = BitConverter.ToInt16(context.PropertyBytes, context.CurIndex);
                        this.value = new byte[sizeof(short) + bytesCount * sizeof(int)];
                        Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, sizeof(short));
                        context.CurIndex += sizeof(short);
                        if (bytesCount == 0)
                        {
                            this.value = null;
                        }
                        else
                        {
                            Array.Copy(context.PropertyBytes, context.CurIndex, this.value, sizeof(short), bytesCount * sizeof(int));
                            context.CurIndex += bytesCount * sizeof(int);
                        }
                    }

                    break;
                case PropertyType.PtypMultipleFloating64:
                case PropertyType.PtypMultipleCurrency:
                case PropertyType.PtypMultipleFloatingTime:
                case PropertyType.PtypMultipleInteger64:
                case PropertyType.PtypMultipleTime:
                    if (context.AvailBytes() < sizeof(short))
                    {
                        throw new ParseException("Not well formed PtypMultipleInterger");
                    }
                    else
                    {
                        short bytesCount = BitConverter.ToInt16(context.PropertyBytes, context.CurIndex);
                        this.value = new byte[sizeof(short) + bytesCount * sizeof(long)];
                        Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, sizeof(short));
                        context.CurIndex += sizeof(short);
                        if (bytesCount == 0)
                        {
                            this.value = null;
                        }
                        else
                        {
                            Array.Copy(context.PropertyBytes, context.CurIndex, this.value, sizeof(short), bytesCount * sizeof(long));
                            context.CurIndex += bytesCount * sizeof(long);
                        }
                    }

                    break;
                case PropertyType.PtypMultipleGuid:
                    if (context.AvailBytes() < sizeof(short))
                    {
                        throw new ParseException("Not well formed PtypMultipleInterger");
                    }
                    else
                    {
                        short bytesCount = BitConverter.ToInt16(context.PropertyBytes, context.CurIndex);
                        this.value = new byte[sizeof(short) + bytesCount * 16];
                        Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, sizeof(short));
                        context.CurIndex += sizeof(short);
                        if (bytesCount == 0)
                        {
                            this.value = null;
                        }
                        else
                        {
                            Array.Copy(context.PropertyBytes, context.CurIndex, this.value, sizeof(short), bytesCount * 16);
                            context.CurIndex += bytesCount * 16;
                        }
                    }

                    break;
                case PropertyType.PtypString8:

                    // The length in bytes of the unicode string to parse
                    strBytesLen = 0;
                    isFound = false;

                    // Property start with "FF"
                    if (context.PropertyBytes[context.CurIndex] == (byte)0xFF)
                    {
                        context.CurIndex++;
                    }

                    // Find the string with '\0' end
                    for (int i = context.CurIndex; i < context.PropertyBytes.Length; i++)
                    {
                        strBytesLen++;
                        if (context.PropertyBytes[i] == 0)
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
                        this.value = new byte[strBytesLen];
                        Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, strBytesLen);
                        context.CurIndex += strBytesLen;
                    }

                    break;
                case PropertyType.PtypString:
                    // The length in bytes of the unicode string to parse
                    strBytesLen = 0;
                    isFound = false;

                    // Property start with "FF"
                    if (context.PropertyBytes[context.CurIndex] == (byte)0xFF)
                    {
                        context.CurIndex++;
                    }
                    else
                    {
                        if (context.PropertyBytes[context.CurIndex] == (byte)0x00)
                        {
                            context.CurIndex++;
                            value = new byte[]
                            { 
                                0x00, 0x00 
                            };
                            break; 
                        } 
                    }

                    // Find the string with '\0''\0' end
                    for (int i = context.CurIndex; i < context.PropertyBytes.Length; i += 2)
                    {
                        strBytesLen += 2;
                        if ((context.PropertyBytes[i] == 0) && (context.PropertyBytes[i + 1] == 0))
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
                        this.value = new byte[strBytesLen];
                        Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, strBytesLen);
                        context.CurIndex += strBytesLen;
                    }

                    break;
                case PropertyType.PtypMultipleString:
                    if (context.AvailBytes() < sizeof(short))
                    {
                        throw new FormatException("Not well formed PtypMultipleString");
                    }
                    else
                    {
                        strBytesLen = 0;
                        isFound = false;
                        short stringCount = BitConverter.ToInt16(context.PropertyBytes, context.CurIndex);
                        context.CurIndex += sizeof(short);
                        if (stringCount == 0)
                        {
                            value = null;
                            break;
                        }

                        for (int i = context.CurIndex; i < context.PropertyBytes.Length; i += 2)
                        {
                            strBytesLen += 2;
                            if ((context.PropertyBytes[i] == 0) && (context.PropertyBytes[i + 1] == 0))
                            {
                                stringCount--;
                            }

                            if (stringCount == 0)
                            {
                                isFound = true;
                                break;
                            }
                        }

                        if (!isFound)
                        {
                            throw new FormatException("String too long or not found");
                        }
                        else
                        {
                            value = new byte[strBytesLen];
                            Array.Copy(context.PropertyBytes, context.CurIndex, value, 0, strBytesLen);
                            context.CurIndex += strBytesLen;
                        }
                    }

                    break;
                case PropertyType.PtypMultipleString8:
                    if (context.AvailBytes() < sizeof(int))
                    {
                        throw new FormatException("Not well formed PtypMultipleString8");
                    }
                    else
                    {
                        List<byte> listOfBytes = new List<byte>();

                        // Property start with "FF"
                        if (context.PropertyBytes[context.CurIndex] == (byte)0xFF)
                        {
                            context.CurIndex++;
                        }

                        if (context.PropertyBytes[context.CurIndex] == (byte)0x00)
                        {
                            this.value = null;
                        }

                        int stringCount = BitConverter.ToInt32(context.PropertyBytes, context.CurIndex);
                        byte[] countOfArray = new byte[sizeof(int)];
                        Array.Copy(context.PropertyBytes, context.CurIndex, countOfArray, 0, sizeof(int));
                        listOfBytes.AddRange(countOfArray);
                        context.CurIndex += sizeof(int);
                        if (stringCount == 0)
                        {
                            value = null;
                            break;
                        }

                        for (int i = 0; i < stringCount; i++)
                        {
                            // Property start with "FF"
                            if (context.PropertyBytes[context.CurIndex] == (byte)0xFF)
                            {
                                context.CurIndex++;
                            }

                            int countOfString8 = 0;
                            for (int j = context.CurIndex; j < context.PropertyBytes.Length; j++)
                            {
                                countOfString8++;
                                if (context.PropertyBytes[j] == 0)
                                {
                                    break;
                                }
                            }

                            byte[] bytesOfString8 = new byte[countOfString8];
                            Array.Copy(context.PropertyBytes, context.CurIndex, bytesOfString8, 0, countOfString8);
                            listOfBytes.AddRange(bytesOfString8);
                            context.CurIndex += countOfString8;
                        }

                        this.value = listOfBytes.ToArray();
                    }

                    break;
                case PropertyType.PtypRuleAction:
                    // Length of the property
                    int felength = 0;

                    // Length of the action blocks
                    short actionBolcksLength = BitConverter.ToInt16(context.PropertyBytes, context.CurIndex);
                    felength += 2;
                    short actionBlockLength = 0;
                    for (int i = 0; i < actionBolcksLength; i++)
                    {
                        actionBlockLength = BitConverter.ToInt16(context.PropertyBytes, context.CurIndex + felength);
                        felength += 2 + actionBlockLength;
                    }

                    this.value = new byte[felength];
                    Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, felength);
                    context.CurIndex += felength;
                    break;
                case PropertyType.PtypServerId:
                    if (context.AvailBytes() < sizeof(short))
                    {
                        throw new ParseException("Not well formed PtypServerId");
                    }
                    else
                    {
                        this.value = new byte[sizeof(short) + 21 * sizeof(byte)];
                        Array.Copy(context.PropertyBytes, context.CurIndex, this.value, 0, sizeof(short) + 21);

                        context.CurIndex += 21 + sizeof(short);
                    }

                    break;
                case PropertyType.PtypMultipleBinary:
                    if (context.AvailBytes() < sizeof(short))
                    {
                        throw new ParseException("Not well formed PtypMultipleBinary");
                    }
                    else
                    {
                        List<byte> listOfBytes = new List<byte>();

                        // Property start with "FF"
                        if (context.PropertyBytes[context.CurIndex] == (byte)0xFF)
                        {
                            context.CurIndex++;
                        }

                        int bytesCount = BitConverter.ToInt32(context.PropertyBytes, context.CurIndex);
                        byte[] countOfArray = new byte[sizeof(int)];
                        Array.Copy(context.PropertyBytes, context.CurIndex, countOfArray, 0, sizeof(int));
                        listOfBytes.AddRange(countOfArray);
                        context.CurIndex += sizeof(int);
                        for (int ibin = 0; ibin < bytesCount; ibin++)
                        {
                            // Property start with "FF"
                            if (context.PropertyBytes[context.CurIndex] == (byte)0xFF)
                            {
                                context.CurIndex++;
                            }

                            int binLength = BitConverter.ToInt32(context.PropertyBytes, context.CurIndex);
                            Array.Copy(context.PropertyBytes, context.CurIndex, countOfArray, 0, sizeof(int));
                            listOfBytes.AddRange(countOfArray);
                            context.CurIndex += sizeof(int);
                            if (binLength > 0)
                            {
                                byte[] bytesArray = new byte[binLength];
                                Array.Copy(context.PropertyBytes, context.CurIndex, bytesArray, 0, binLength);
                                listOfBytes.AddRange(bytesArray);
                                context.CurIndex += sizeof(byte) * binLength;
                            }
                        }

                        this.value = listOfBytes.ToArray();
                    }

                    break;

                default:
                    throw new FormatException("Type " + type.ToString() + " not found or not support.");
            }
        }
    }
}