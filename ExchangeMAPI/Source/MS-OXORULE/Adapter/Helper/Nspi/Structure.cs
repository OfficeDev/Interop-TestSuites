namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using System.Text;
    using Microsoft.Protocols.TestTools.Messages.Marshaling;

    #region Property Values
    /// <summary>
    /// The FlatUID_r structure is an encoding of the FlatUID data structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct FlatUID_r
    {
        /// <summary>
        /// Encodes the ordered bytes of the FlatUID data structure.
        /// </summary>
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = Constants.FlatUIDByteSize), Inline()]
        public byte[] Ab;

        /// <summary>
        /// Compare two FlatUID_r structures whether equal or not.
        /// </summary>
        /// <param name="flatUid">The second FlatUID_r structure.</param>
        /// <returns>True: the two FlatUID_r structure equal; Otherwise, false.</returns>
        public bool Compare(FlatUID_r flatUid)
        {
            for (int i = 0; i < Constants.FlatUIDByteSize; i++)
            {
                if (this.Ab[i] != flatUid.Ab[i])
                {
                    return false;
                }
            }

            return true;
        }
    }

    /// <summary>
    /// The PropertyTagArray_r structure is an encoding of the PropTagArray data structure.
    /// </summary>
    public struct PropertyTagArray_r
    {
        /// <summary>
        /// Encodes the Count field in the PropTagArray structure.
        /// </summary>
        public uint Values;

        /// <summary>
        /// Encodes the PropertyTags field of PropTagArray.
        /// </summary>
        [Size("CValues"), Inline()]
        public uint[] AulPropTag;

        /// <summary>
        /// This method is used for serializing the PropertyTagArray_r.
        /// </summary>
        /// <returns>The serialized bytes of PropertyTagArray_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.Values));
            for (int i = 0; i < this.AulPropTag.Length; i++)
            {
                buffer.AddRange(BitConverter.GetBytes(this.AulPropTag[i]));
            }

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The Binary_r structure encodes an array of un-interpreted bytes.
    /// </summary>
    public struct Binary_r
    {
        /// <summary>
        /// The number of un-interpreted bytes represented in this structure. This value MUST NOT exceed 2,097,152.
        /// </summary>
        public uint Cb;

        /// <summary>
        /// The un-interpreted bytes.
        /// </summary>
        [Size("Cb")]
        public byte[] Lpb;

        /// <summary>
        /// This method is used for serializing the Binary_r.
        /// </summary>
        /// <returns>The serialized bytes of Binary_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.Cb));
            buffer.AddRange(this.Lpb);
            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The ShortArray_r structure encodes an array of 16-bit integers.
    /// </summary>
    public struct ShortArray_r
    {
        /// <summary>
        /// The number of 16-bit integer values represented in the ShortArray_r structure.
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint CValues;

        /// <summary>
        /// The 16-bit integer values.
        /// </summary>
        [Size("CValues")]
        public short[] Lpi;

        /// <summary>
        /// This method is used for serializing the ShortArray_r.
        /// </summary>
        /// <returns>The serialized bytes of ShortArray_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.CValues));
            for (int i = 0; i < this.Lpi.Length; i++)
            {
                buffer.AddRange(BitConverter.GetBytes(this.Lpi[i]));
            }

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The LongArray_r structure encodes an array of 32-bit integers.
    /// </summary>
    public struct LongArray_r
    {
        /// <summary>
        /// The number of 32-bit integers represented in this structure. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint CValues;

        /// <summary>
        /// The 32-bit integer values.
        /// </summary>
        [Size("CValues")]
        public int[] Lpl;

        /// <summary>
        /// This method is used for serializing the LongArray_r.
        /// </summary>
        /// <returns>The serialized bytes of LongArray_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.CValues));
            for (int i = 0; i < this.Lpl.Length; i++)
            {
                buffer.AddRange(BitConverter.GetBytes(this.Lpl[i]));
            }

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The StringArray_r structure encodes an array of references to 8-bit character strings.
    /// </summary>
    public struct StringArray_r
    {
        /// <summary>
        /// The number of 8-bit character string references represented in the StringArray_r structure. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint CValues;

        /// <summary>
        /// The 8-bit character string references. The strings referred to are NULL-terminated.
        /// </summary>
        [Size("CValues"), String(StringEncoding.ASCII)]
        public string[] LppszA;

        /// <summary>
        /// This method is used for serializing the StringArray_r.
        /// </summary>
        /// <returns>The serialized bytes of StringArray_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.CValues));
            for (int i = 0; i < this.LppszA.Length; i++)
            {
                buffer.AddRange(Encoding.ASCII.GetBytes(this.LppszA[i]));
            }

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The BinaryArray_r structure is an array of Binary_r data structures.
    /// </summary>
    public struct BinaryArray_r
    {
        /// <summary>
        /// The number of Binary_r data structures represented in the BinaryArray_r structure. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint CValues;

        /// <summary>
        /// The Binary_r data structures.
        /// </summary>
        [Size("CValues")]
        public Binary_r[] Lpbin;

        /// <summary>
        /// This method is used for serializing the BinaryArray_r.
        /// </summary>
        /// <returns>The serialized bytes of BinaryArray_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.CValues));
            for (int i = 0; i < this.Lpbin.Length; i++)
            {
                buffer.AddRange(this.Lpbin[i].Serialize());
            }

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The FlatUIDArray_r structure encodes an array of FlatUID_r data structures.
    /// </summary>
    public struct FlatUIDArray_r
    {
        /// <summary>
        /// The number of FlatUID_r structures represented in the FlatUIDArray_r structure. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint CValues;

        /// <summary>
        /// The FlatUID_r data structures.
        /// </summary>
        [Size("CValues")]
        public FlatUID_r[] Lpguid;

        /// <summary>
        /// This method is used for serializing the FlatUIDArray_r.
        /// </summary>
        /// <returns>The serialized bytes of FlatUIDArray_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.CValues));
            for (int i = 0; i < this.Lpguid.Length; i++)
            {
                buffer.AddRange(this.Lpguid[i].Ab);
            }

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The WStringArray_r structure encodes an array of references to Unicode strings.
    /// </summary>
    public struct WStringArray_r
    {
        /// <summary>
        /// The number of Unicode character string references represented in the WStringArray_r structure. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint CValues;

        /// <summary>
        /// The Unicode character string references. The strings referred to are NULL-terminated.
        /// </summary>
        [Size("CValues"), String(StringEncoding.Unicode)]
        public string[] LppszW;

        /// <summary>
        /// This method is used for serializing the WStringArray_r.
        /// </summary>
        /// <returns>The serialized bytes of WStringArray_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.CValues));
            for (int i = 0; i < this.LppszW.Length; i++)
            {
                buffer.AddRange(Encoding.Unicode.GetBytes(this.LppszW[i]));
            }

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The DateTimeArray_r structure encodes an array of FILETIME structures.
    /// </summary>
    public struct DateTimeArray_r
    {
        /// <summary>
        /// The number of FILETIME data structures represented in the DateTimeArray_r structure. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint CValues;

        /// <summary>
        /// The FILETIME data structures.
        /// </summary>
        public FILETIME[] Lpft;

        /// <summary>
        /// This method is used for serializing the DateTimeArray_r.
        /// </summary>
        /// <returns>The serialized bytes of DateTimeArray_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.CValues));
            for (int i = 0; i < this.Lpft.Length; i++)
            {
                buffer.AddRange(this.Lpft[i].Serialize());
            }

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The FILETIME data structure.
    /// </summary>
    public struct FILETIME
    {
        /// <summary>
        /// The low-order part of the file time.
        /// </summary>
        public uint LowDateTime;

        /// <summary>
        /// The high-order part of the file time.
        /// </summary>
        public uint HighDateTime;

        /// <summary>
        /// This method is used for serializing the FILETIME.
        /// </summary>
        /// <returns>The serialized bytes of FILETIME.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.LowDateTime));
            buffer.AddRange(BitConverter.GetBytes(this.HighDateTime));
            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The PROP_VAL_UNION structure encodes a single instance of any type of property value. 
    /// </summary>
    [Union("System.Int32")]
    public struct PROP_VAL_UNION
    {
        /// <summary>
        /// A single 16-bit integer value.
        /// </summary>
        [Case("0x00000002")]
        public short I;

        /// <summary>
        /// A single 32-bit integer value.
        /// </summary>
        [Case("0x00000003")]
        public int L;

        /// <summary>
        /// A single Boolean value. 
        /// The client and server MUST NOT set this to values other than 1 or 0.
        /// </summary>
        [Case("0x0000000B")]
        public uint B;

        /// <summary>
        /// A single 8-bit character string value. This value is NULL-terminated.
        /// </summary>
        [Case("0x0000001E")]
        public byte[] LpszA;

        /// <summary>
        /// A single binary data value. 
        /// The number of bytes that can be encoded in this structure is 2,097,152.
        /// </summary>
        [Case("0x00000102")]
        public Binary_r Bin;

        /// <summary>
        /// A single Unicode string value. This value is NULL-terminated.
        /// </summary>
        [Case("0x0000001F")]
        public byte[] LpszW;

        /// <summary>
        /// A single GUID value. The value is encoded as a FlatUID_r data structure.
        /// </summary>
        [Case("0x00000048")]
        public FlatUID_r[] Lpguid;

        /// <summary>
        /// A single 64-bit integer value. 
        /// The value is encoded as a FILETIME structure. 
        /// </summary>
        [Case("0x00000040")]
        public FILETIME Ft;

        /// <summary>
        /// A single PtypErrorCode value.
        /// </summary>
        [Case("0x0000000A")]
        public int Err;

        /// <summary>
        /// Multiple 16-bit integer values. 
        /// The number of values that can be encoded in this structure is 100,000.
        /// </summary>
        [Case("0x00001002")]
        public ShortArray_r MVi;

        /// <summary>
        /// Multiple 32-bit integer values. 
        /// The number of values that can be encoded in this structure is 100,000.
        /// </summary>
        [Case("0x00001003")]
        public LongArray_r MVl;

        /// <summary>
        /// Multiple 8-bit character string values. These string values are NULL-terminated. 
        /// The number of values that can be encoded in this structure is 100,000.
        /// </summary>
        [Case("0x0000101E")]
        public StringArray_r MVszA;

        /// <summary>
        /// Multiple binary data values. The number of bytes that can be encoded in each value of this structure is 2,097,152. 
        /// The number of values that can be encoded in this structure is 100,000.
        /// </summary>
        [Case("0x00001102")]
        public BinaryArray_r MVbin;

        /// <summary>
        /// Multiple GUID values. The values are encoded as FlatUID_r data structures. 
        /// The number of values that can be encoded in this structure is 100,000.
        /// </summary>
        [Case("0x00001048")]
        public FlatUIDArray_r MVguid;

        /// <summary>
        /// Multiple Unicode string values. These string values are NULL-terminated.
        /// The number of values that can be encoded in this structure is 100,000.
        /// </summary>
        [Case("0x0000101F")]
        public WStringArray_r MVszW;

        /// <summary>
        /// Multiple 64-bit integer values. The values are encoded as FILETIME structures.
        /// The number of values that can be encoded in this structure is 100,000.
        /// </summary>
        [Case("0x00001040")]
        public DateTimeArray_r MVft;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000000.
        /// </summary>
        [Case("0x00000001, 0x0000000D")]
        public int Reserved;
    }

    /// <summary>
    /// The PropertyValue_r structure is an encoding of the PropertyValue data structure
    /// </summary>
    public struct PropertyValue_r
    {
        /// <summary>
        /// Encodes the property tag of the property whose value is represented by the PropertyValue_r data structure.
        /// </summary>
        public uint PropTag;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000000.
        /// </summary>
        public uint Reserved;

        /// <summary>
        /// Encodes the actual value of the property represented by the PropertyValue_r data structure.
        /// </summary>
        [Switch("PropTag & 0x0000FFFF")]
        public PROP_VAL_UNION Value;

        /// <summary>
        /// This method is used for serializing the PropertyValue_r.
        /// </summary>
        /// <returns>The serialized bytes of PropertyValue_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.PropTag));
            buffer.AddRange(BitConverter.GetBytes(this.Reserved));
            switch (this.PropTag & 0x0000FFFF)
            {
                case 0x00000002:
                    buffer.AddRange(BitConverter.GetBytes(this.Value.I));
                    break;

                case 0x00000003:
                    buffer.AddRange(BitConverter.GetBytes(this.Value.L));
                    break;

                case 0x0000000b:
                    buffer.AddRange(BitConverter.GetBytes(this.Value.B));
                    break;

                case 0x0000001e:
                    if (this.Value.LpszA.Length > 0)
                    {
                        buffer.Add(0xff);
                        buffer.AddRange(this.Value.LpszA);
                    }

                    if (this.Value.LpszA.Length == 0)
                    {
                        buffer.Add(0x00);
                    }

                    break;

                case 0x00000102:
                    if (this.Value.Bin.Serialize().Length > 0)
                    {
                        buffer.Add(0xff);
                        buffer.AddRange(this.Value.Bin.Serialize());
                    }

                    if (this.Value.Bin.Serialize().Length == 0)
                    {
                        buffer.Add(0x00);
                    }

                    break;

                case 0x0000001f:
                    if (this.Value.LpszW.Length > 0)
                    {
                        buffer.Add(0xff);
                        buffer.AddRange(this.Value.LpszW);
                    }

                    if (this.Value.LpszW.Length == 0)
                    {
                        buffer.Add(0x00);
                    }

                    break;

                case 0x00000048:
                    for (int i = 0; i < this.Value.Lpguid.Length; i++)
                    {
                        buffer.AddRange(this.Value.Lpguid[i].Ab);
                    }

                    break;

                case 0x00000040:
                    buffer.AddRange(this.Value.Ft.Serialize());
                    break;

                case 0x0000000a:
                    buffer.AddRange(BitConverter.GetBytes(this.Value.Err));
                    break;

                case 0x00001002:
                    buffer.AddRange(this.Value.MVi.Serialize());
                    break;

                case 0x00001003:
                    buffer.AddRange(this.Value.MVl.Serialize());
                    break;

                case 0x0000101e:
                    if (this.Value.MVszA.Serialize().Length > 0)
                    {
                        buffer.Add(0xff);
                        buffer.AddRange(BitConverter.GetBytes(this.Value.MVszA.CValues));
                        for (int i = 0; i < this.Value.MVszA.CValues; i++)
                        {
                            StringBuilder stringForAdd = new StringBuilder(this.Value.MVszA.LppszA[i]);
                            buffer.Add(0xff);
                            buffer.AddRange(System.Text.Encoding.UTF8.GetBytes(stringForAdd.ToString() + "\0"));
                        }
                    }

                    if (this.Value.MVszA.Serialize().Length == 0)
                    {
                        buffer.Add(0x00);
                    }

                    break;

                case 0x00001048:
                    buffer.AddRange(this.Value.MVguid.Serialize());
                    break;

                case 0x0000101f:
                    if (this.Value.MVszW.Serialize().Length > 0)
                    {
                        buffer.Add(0xff);
                        buffer.AddRange(BitConverter.GetBytes(this.Value.MVszW.CValues));
                        for (int i = 0; i < this.Value.MVszW.CValues; i++)
                        {
                            StringBuilder stringForAdd = new StringBuilder(this.Value.MVszW.LppszW[i]);
                            buffer.Add(0xff);
                            buffer.AddRange(System.Text.Encoding.Unicode.GetBytes(stringForAdd.ToString() + "\0"));
                        }
                    }

                    if (this.Value.MVszW.Serialize().Length == 0)
                    {
                        buffer.Add(0x00);
                    }

                    break;

                case 0x00001040:
                    buffer.AddRange(this.Value.MVft.Serialize());
                    break;

                case 0x00001102:
                    if (this.Value.MVbin.Serialize().Length > 0)
                    {
                        buffer.Add(0xff);
                        buffer.AddRange(BitConverter.GetBytes(this.Value.MVbin.CValues));
                        for (int i = 0; i < this.Value.MVbin.CValues; i++)
                        {
                            buffer.Add(0xff);
                            buffer.AddRange(BitConverter.GetBytes(this.Value.MVbin.Lpbin[i].Cb));
                            buffer.AddRange(this.Value.MVbin.Lpbin[i].Lpb);
                        }
                    }

                    if (this.Value.MVbin.Serialize().Length == 0)
                    {
                        buffer.Add(0x00);
                    }

                    break;

                case 0x00000001:
                case 0x0000000d:
                    buffer.AddRange(BitConverter.GetBytes(this.Value.Reserved));
                    break;
            }

            return buffer.ToArray();
        }
    }

    #endregion

    /// <summary>
    /// The PropertyRow_r structure is an encoding of the StandardPropertyRow data structure.
    /// </summary>
    public struct PropertyRow_r
    {
        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000000.
        /// </summary>
        public uint Reserved;

        /// <summary>
        /// The number of PropertyValue_r structures represented in the PropertyRow_r structure. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint Values;

        /// <summary>
        /// Encode the ValueArray field of the StandardPropertyRow data structure.
        /// </summary>
        [Size("CValues")]
        public PropertyValue_r[] Props;
    }

    /// <summary>
    /// The PropertyRowSet_r structure is an encoding of the PropertyRowSet data structure.
    /// </summary>
    public struct PropertyRowSet_r
    {
        /// <summary>
        /// Encode the RowCount field of the PropertyRowSet data structures. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint Rows;

        /// <summary>
        /// Encode the Rows field of the PropertyRowSet data structure.
        /// </summary>
        [Size("Rows")]
        public PropertyRow_r[] PropertyRowSet;
    }

    #region Restrictions
    /// <summary>
    /// The AndRestriction_r is a single RPC encoding.
    /// </summary>
    public struct AndRestriction_r
    {
        /// <summary>
        /// Encodes the RestrictCount field of the AndRestriction and OrRestriction data structures. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint CRes;

        /// <summary>
        /// Encodes the Restricts field of the AndRestriction and OrRestriction data structures. 
        /// </summary>
        [Size("CRes")]
        public Restriction_r[] LpRes;

        /// <summary>
        /// This method is used for serializing the AndRestriction_r.
        /// </summary>
        /// <returns>The serialized bytes of AndRestriction_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.CRes));
            for (int i = 0; i < this.LpRes.Length; i++)
            {
                buffer.AddRange(this.LpRes[i].Serialize());
            }

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The OrRestriction_r is a single RPC encoding.
    /// </summary>
    public struct OrRestriction_r
    {
        /// <summary>
        /// Encodes the RestrictCount field of the AndRestriction and OrRestriction data structures. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint CRes;

        /// <summary>
        /// Encodes the Restricts field of the AndRestriction and OrRestriction data structures. 
        /// </summary>
        [Size("CRes")]
        public Restriction_r[] LpRes;

        /// <summary>
        /// This method is used for serializing the OrRestriction_r.
        /// </summary>
        /// <returns>The serialized bytes of OrRestriction_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.CRes));
            for (int i = 0; i < this.LpRes.Length; i++)
            {
                buffer.AddRange(this.LpRes[i].Serialize());
            }

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The NotRestriction_r structure is an encoding of the NotRestriction data structure.
    /// </summary>
    public struct NotRestriction_r
    {
        /// <summary>
        /// Encodes the Restriction field of the NotRestriction data structure.
        /// </summary>
        [StaticSize(1, StaticSizeMode.Elements)]
        public Restriction_r[] LpRes;

        /// <summary>
        /// This method is used for serializing the NotRestriction_r.
        /// </summary>
        /// <returns>The serialized bytes of NotRestriction_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            for (int i = 0; i < this.LpRes.Length; i++)
            {
                buffer.AddRange(this.LpRes[i].Serialize());
            }

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The ContentRestriction_r structure is an encoding of the ContentRestriction data structure.
    /// </summary>
    public struct ContentRestriction_r
    {
        /// <summary>
        /// Encodes the FuzzyLevelLow and  FuzzyLevelHigh fields of the ContentRestriction data structure.
        /// </summary>
        public uint FuzzyLevel;

        /// <summary>
        /// Encodes the PropertyTag field of the ContentRestriction data structure.
        /// </summary>
        public uint PropTag;

        /// <summary>
        /// Encodes the TaggedValue field of the ContentRestriction data structure.
        /// </summary>
        [StaticSize(1, StaticSizeMode.Elements)]
        public PropertyValue_r[] Prop;

        /// <summary>
        /// This method is used for serializing the ContentRestriction_r.
        /// </summary>
        /// <returns>The serialized bytes of ContentRestriction_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.FuzzyLevel));
            buffer.AddRange(BitConverter.GetBytes(this.PropTag));
            for (int i = 0; i < this.Prop.Length; i++)
            {
                buffer.AddRange(this.Prop[i].Serialize());
            }

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The PropertyRestriction_r structure is an encoding of the PropertyRestriction data structure.
    /// </summary>
    public struct Propertyrestriction_r
    {
        /// <summary>
        /// Encodes the RelOp field of the PropertyRestriction data structure.
        /// </summary>
        public uint Relop;

        /// <summary>
        /// Encodes the PropTag field of the PropertyRestriction data structure.
        /// </summary>
        public uint PropTag;

        /// <summary>
        /// Encodes the TaggedValue field of the PropertyRestriction data structure.c
        /// </summary>
        [StaticSize(1, StaticSizeMode.Elements)]
        public PropertyValue_r[] Prop;

        /// <summary>
        /// This method is used for serializing the Propertyrestriction_r.
        /// </summary>
        /// <returns>The serialized bytes of Propertyrestriction_r</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.Relop));
            buffer.AddRange(BitConverter.GetBytes(this.PropTag));
            for (int i = 0; i < this.Prop.Length; i++)
            {
                buffer.AddRange(this.Prop[i].Serialize());
            }

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// Encodes the TaggedValue field of the PropertyRestriction data structure.
    /// </summary>
    public struct ExistRestriction_r
    {
        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000000.
        /// </summary>
        public uint Reserved1;

        /// <summary>
        /// Encodes the PropTag field of the ExistRestriction data structure.
        /// </summary>
        public uint PropTag;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000000.
        /// </summary>
        public uint Reserved2;

        /// <summary>
        /// This method is used for serializing the ExistRestriction_r.
        /// </summary>
        /// <returns>The serialized bytes of ExistRestriction_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.Reserved1));
            buffer.AddRange(BitConverter.GetBytes(this.PropTag));
            buffer.AddRange(BitConverter.GetBytes(this.Reserved2));
            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The RestrictionUnion_r structure encodes a single instance of any type of restriction. 
    /// </summary>
    [Union("System.Int32")]
    public struct RestrictionUnion_r
    {
        /// <summary>
        /// RestrictionUnion_r contains an encoding of an AndRestriction.
        /// </summary>
        [Case("0x00000000")]
        public AndRestriction_r ResAnd;

        /// <summary>
        /// RestrictionUnion_r contains an encoding of an OrRestriction.
        /// </summary>
        [Case("0x00000001")]
        public OrRestriction_r ResOr;

        /// <summary>
        /// RestrictionUnion_r contains an encoding of a NotRestriction.
        /// </summary>
        [Case("0x00000002")]
        public NotRestriction_r ResNot;

        /// <summary>
        /// RestrictionUnion_r contains an encoding of a ContentRestriction.
        /// </summary>
        [Case("0x00000003")]
        public ContentRestriction_r ResContent;

        /// <summary>
        /// RestrictionUnion_r contains an encoding of a PropertyRestriction.
        /// </summary>
        [Case("0x00000004")]
        public Propertyrestriction_r ResProperty;

        /// <summary>
        /// RestrictionUnion_r contains an encoding of an ExistRestriction.
        /// </summary>
        [Case("0x00000008")]
        public ExistRestriction_r ResExist;
    }

    /// <summary>
    /// The Restriction_r structure is an encoding of the Restriction filters.
    /// </summary>
    public struct Restriction_r
    {
        /// <summary>
        /// Encodes the RestrictType field common to all restriction structures.
        /// </summary>
        public uint Rt;

        /// <summary>
        /// Encodes the actual restriction specified by the type in the rt field.
        /// </summary>
        [Switch("Rt")]
        public RestrictionUnion_r Res;

        /// <summary>
        /// This method is used for serializing the Restriction_r.
        /// </summary>
        /// <returns>The serialized bytes of Restriction_r.</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.Rt));
            switch (this.Rt)
            {
                case 0x00000000:
                    buffer.AddRange(this.Res.ResAnd.Serialize());
                    break;

                case 0x00000001:
                    buffer.AddRange(this.Res.ResOr.Serialize());
                    break;

                case 0x00000002:
                    buffer.AddRange(this.Res.ResNot.Serialize());
                    break;

                case 0x00000003:
                    buffer.AddRange(this.Res.ResContent.Serialize());
                    break;

                case 0x00000004:
                    buffer.AddRange(this.Res.ResProperty.Serialize());
                    break;

                case 0x00000008:
                    buffer.AddRange(this.Res.ResExist.Serialize());
                    break;
            }

            return buffer.ToArray();
        }
    }
    #endregion

    /// <summary>
    /// The PropertyName_r structure is an encoding of the PropertyName data structure.
    /// </summary>
    public struct PropertyName_r
    {
        /// <summary>
        /// Encode the GUID field of the PropertyName data structure. This field is encoded as a FlatUID_r data structure.
        /// </summary>
        [StaticSize(1, StaticSizeMode.Elements)]
        public FlatUID_r[] Guid;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000000.
        /// </summary>
        public uint Reserved;

        /// <summary>
        /// Encode the lID field of the PropertyName data structure. 
        /// </summary>
        public int ID;
    }

    /// <summary>
    /// The StringsArray_r structure is used to aggregate a number of character type strings into a single data structure.
    /// </summary>
    public struct StringsArray_r
    {
        /// <summary>
        /// The number of character string structures in this aggregation. 
        /// The value MUST NOT exceed 100,000.
        /// </summary>
        public uint CValues;

        /// <summary>
        /// The list of character type strings in this aggregation. 
        /// The strings in this list are NULL-terminated.
        /// </summary>
        [Size("Count"), String(StringEncoding.ASCII)]
        public string[] LppszA;
    }

    /// <summary>
    /// The WStringsArray_r structure is used to aggregate a number of wchar_t type strings into a single data structure.
    /// </summary>
    public struct WStringsArray_r
    {
        /// <summary>
        /// The number of character strings structures in this aggregation. The value MUST NOT exceed 100,000.
        /// </summary>
        public uint CValues;

        /// <summary>
        /// The list of wchar_t type strings in this aggregation. The strings in this list are NULL-terminated.
        /// </summary>
        [Size("Count"), String(StringEncoding.Unicode)]
        public string[] LppszW;
    }

    /// <summary>
    /// The STAT structure is used to specify the state of a table and location information that applies to that table. 
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct STAT
    {
        /// <summary>
        /// This field contains a DWORD value that represents a sort order. 
        /// The client sets this field to specify the sort type of this table. 
        /// </summary>
        public uint SortType;

        /// <summary>
        /// This field contains a Minimal Entry ID. 
        /// The client sets this field to specify the Minimal Entry ID of the address book container that this STAT structure represents. 
        /// </summary>
        public uint ContainerID;

        /// <summary>
        /// This field contains a Minimal Entry ID. 
        /// The client sets this field to specify a beginning position in the table for the start of an NSPI method. 
        /// The server sets this field to report the end position in the table after processing an NSPI method.
        /// </summary>
        public uint CurrentRec;

        /// <summary>
        /// This field contains a long value. 
        /// The client sets this field to specify an offset from the beginning position in the table for the start of an NSPI method. 
        /// </summary>
        public int Delta;

        /// <summary>
        /// This field contains a DWORD value that specifies a position in the table.
        /// The client sets this field to specify a fractional position for the beginning position in the table for the start of an NSPI method.
        /// The server sets this field to specify the approximate fractional position at the end of an NSPI method. 
        /// </summary>
        public uint NumPos;

        /// <summary>
        /// This field contains a DWORD value that specifies the number of rows in the table. 
        /// The client sets this field to specify a fractional position for the beginning position in the table for the start of an NSPI method.
        /// The server sets this field to specify the total number of rows in the table. 
        /// </summary>
        public uint TotalRecs;

        /// <summary>
        /// This field contains a DWORD value that represents a code page. 
        /// The client sets this field to specify the code page the client uses for non-Unicode strings. 
        /// The server MUST use this value during string handling and MUST NOT modify this field.
        /// </summary>
        public uint CodePage;

        /// <summary>
        /// This field contains a DWORD value that represents a language code identifier (LCID).
        /// The client sets this field to specify the LCID associated with the template the client wants the server to return. 
        /// The server MUST NOT modify this field.
        /// </summary>
        public uint TemplateLocale;

        /// <summary>
        /// This field contains a DWORD value that represents an LCID. 
        /// The client sets this field to specify the LCID that it wants the server to use when sorting any strings. 
        /// The server MUST use this value during sorting and MUST NOT modify this field.
        /// </summary>
        public uint SortLocale;

        /// <summary>
        /// Parse the STAT from the response data.
        /// </summary>
        /// <param name="rawData">The response data.</param>
        /// <param name="startIndex">The start index.</param>
        /// <returns>The result of STAT.</returns>
        public static STAT Parse(byte[] rawData, ref int startIndex)
        {
            STAT state = new STAT();
            state.SortType = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.ContainerID = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.CurrentRec = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.Delta = BitConverter.ToInt32(rawData, startIndex);
            startIndex += sizeof(int);
            state.NumPos = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.TotalRecs = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.CodePage = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.TemplateLocale = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.SortLocale = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);

            return state;
        }

        /// <summary>
        /// Initiate the stat to initial values.
        /// </summary>
        public void InitiateStat()
        {
            this.CodePage = (uint)RequiredCodePage.CP_TELETEX;
            this.CurrentRec = (uint)MinimalEntryID.MID_BEGINNING_OF_TABLE;
            this.ContainerID = 0;
            this.Delta = 0;
            this.NumPos = 0;
            this.SortLocale = (uint)DefaultLCID.NSPI_DEFAULT_LOCALE;
            this.SortType = (uint)TableSortOrder.SortTypeDisplayName;
            this.TemplateLocale = (uint)DefaultLCID.NSPI_DEFAULT_LOCALE;
            this.TotalRecs = 0;
        }

        /// <summary>
        /// This method is used for serializing the STAT
        /// </summary>
        /// <returns>The serialized bytes of STAT</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.SortType));
            buffer.AddRange(BitConverter.GetBytes(this.ContainerID));
            buffer.AddRange(BitConverter.GetBytes(this.CurrentRec));
            buffer.AddRange(BitConverter.GetBytes(this.Delta));
            buffer.AddRange(BitConverter.GetBytes(this.NumPos));
            buffer.AddRange(BitConverter.GetBytes(this.TotalRecs));
            buffer.AddRange(BitConverter.GetBytes(this.CodePage));
            buffer.AddRange(BitConverter.GetBytes(this.TemplateLocale));
            buffer.AddRange(BitConverter.GetBytes(this.SortLocale));

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The EphemeralEntryID structure identifies a specific object in the address book. 
    /// </summary>
    public struct EphemeralEntryID
    {
        /// <summary>
        /// The type of this ID. The value is the constant 0x87. 
        /// </summary>
        public byte IDType;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00.
        /// </summary>
        public byte R1;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00.
        /// </summary>
        public byte R2;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00.
        /// </summary>
        public byte R3;

        /// <summary>
        /// Contain the GUID of the server that issues this Ephemeral Entry ID.
        /// </summary>
        public FlatUID_r ProviderUID;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000001.
        /// </summary>
        public uint R4;

        /// <summary>
        /// The display type of the object specified by this Ephemeral Entry ID. 
        /// </summary>
        public DisplayTypeValue DisplayType;

        /// <summary>
        /// The Minimal Entry ID of this object, as specified in section 2.3.8.1. 
        /// This value is expressed in little-endian format.
        /// </summary>
        public uint Mid;

        /// <summary>
        /// Compare whether the two EphemeralEntryID structures are equal or not.
        /// </summary>
        /// <param name="id1">The EphemeralEntryID structure to be compared.</param>
        /// <returns>If the two EphemeralEntryID structures are equal, return true, else false.</returns>
        public bool Compare(EphemeralEntryID id1)
        {
            if (this.DisplayType == id1.DisplayType && this.Mid == id1.Mid)
            {
                if (this.ProviderUID.Compare(id1.ProviderUID))
                {
                    return true;
                }
            }

            return false;
        }
    }

    /// <summary>
    /// The PermanentEntryID structure identifies a specific object in the address book. 
    /// </summary>
    public struct PermanentEntryID
    {
        /// <summary>
        /// The type of this ID. The value is the constant 0x00. 
        /// </summary>
        public byte IDType;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00.
        /// </summary>
        public byte R1;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00.
        /// </summary>
        public byte R2;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00.
        /// </summary>
        public byte R3;

        /// <summary>
        /// A FlatUID_r value that contains the constant GUID specified in Permanent Entry ID GUID.
        /// </summary>
        public FlatUID_r ProviderUID;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000001.
        /// </summary>
        public uint R4;

        /// <summary>
        /// The display type of the object specified by this Permanent Entry ID. 
        /// </summary>
        public DisplayTypeValue DisplayTypeString;

        /// <summary>
        /// The DN of the object specified by this Permanent Entry ID. The value is expressed as a DN.
        /// </summary>
        public string DistinguishedName;
    }
}