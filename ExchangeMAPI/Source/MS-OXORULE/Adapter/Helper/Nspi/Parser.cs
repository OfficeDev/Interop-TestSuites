namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Parse structures.
    /// </summary>
    public static class Parser
    {
        /// <summary>
        /// Parse Binary_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        /// <returns>Instance of Binary_r structure.</returns>
        public static Binary_r ParseBinary_r(IntPtr ptr)
        {
            Binary_r b_r = new Binary_r
            {
                Cb = (uint)Marshal.ReadInt32(ptr)
            };

            if (b_r.Cb == 0)
            {
                b_r.Lpb = null;
            }
            else
            {
                b_r.Lpb = new byte[b_r.Cb];
                IntPtr baddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                for (uint i = 0; i < b_r.Cb; i++)
                {
                    b_r.Lpb[i] = Marshal.ReadByte(baddr, (int)i);
                }
            }

            return b_r;
        }

        /// <summary>
        /// Parse GUIDs.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        /// <returns>Instance of GUIDs.</returns>
        public static FlatUID_r ParseFlatUID_r(IntPtr ptr)
        {
            FlatUID_r fuid_r = new FlatUID_r();

            const int UIDLengthInBytes = 16;
            fuid_r.Ab = new byte[UIDLengthInBytes];
            for (int i = 0; i < UIDLengthInBytes; i++)
            {
                fuid_r.Ab[i] = Marshal.ReadByte(ptr, i);
            }

            return fuid_r;
        }

        /// <summary>
        /// Parse ShortArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        /// <returns>Instance of ShortArray_r structure.</returns>
        public static ShortArray_r ParseShortArray_r(IntPtr ptr)
        {
            ShortArray_r shortArray = new ShortArray_r
            {
                Values = (uint)Marshal.ReadInt32(ptr)
            };

            if (shortArray.Values == 0)
            {
                shortArray.Lpi = null;
            }
            else
            {
                IntPtr saaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                shortArray.Lpi = new short[shortArray.Values];
                int offset = 0;
                for (uint i = 0; i < shortArray.Values; i++, offset += sizeof(short))
                {
                    shortArray.Lpi[i] = Marshal.ReadInt16(saaddr, offset);
                }
            }

            return shortArray;
        }

        /// <summary>
        /// Parse LongArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        /// <returns>Instance of LongArray_r structure.</returns>
        public static LongArray_r ParseLongArray_r(IntPtr ptr)
        {
            LongArray_r longArray = new LongArray_r
            {
                Values = (uint)Marshal.ReadInt32(ptr)
            };

            if (longArray.Values == 0)
            {
                longArray.Lpl = null;
            }
            else
            {
                IntPtr laaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                longArray.Lpl = new int[longArray.Values];
                int offset = 0;
                for (uint i = 0; i < longArray.Values; i++, offset += sizeof(int))
                {
                    longArray.Lpl[i] = Marshal.ReadInt32(laaddr, offset);
                }
            }

            return longArray;
        }

        /// <summary>
        /// Parse StringArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        /// <returns>Instance of StringArray_r structure.</returns>
        public static StringArray_r ParseStringArray_r(IntPtr ptr)
        {
            StringArray_r stringArray = new StringArray_r
            {
                Values = (uint)Marshal.ReadInt32(ptr)
            };

            if (stringArray.Values == 0)
            {
                stringArray.LppszA = null;
            }
            else
            {
                stringArray.LppszA = new string[stringArray.Values];
                IntPtr szaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                int offset = 0;
                for (uint i = 0; i < stringArray.Values; i++)
                {
                    stringArray.LppszA[i] = Marshal.PtrToStringAnsi(new IntPtr(Marshal.ReadInt32(szaddr, offset)));
                    offset += 4;
                }
            }

            return stringArray;
        }

        /// <summary>
        /// Parse BinaryArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        /// <returns>Instance of BinaryArray_r structure.</returns>
        public static BinaryArray_r ParseBinaryArray_r(IntPtr ptr)
        {
            BinaryArray_r binaryArray = new BinaryArray_r
            {
                Values = (uint)Marshal.ReadInt32(ptr)
            };

            if (binaryArray.Values == 0)
            {
                binaryArray.Lpbin = null;
            }
            else
            {
                binaryArray.Lpbin = new Binary_r[binaryArray.Values];
                IntPtr baaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                for (uint i = 0; i < binaryArray.Values; i++)
                {
                    binaryArray.Lpbin[i] = ParseBinary_r(baaddr);
                    baaddr = new IntPtr(baaddr.ToInt32() + 8);
                }
            }

            return binaryArray;
        }

        /// <summary>
        /// Parse WStringArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        /// <returns>Instance of WStringArray_r structure.</returns>
        public static WStringArray_r ParseWStringArray_r(IntPtr ptr)
        {
            WStringArray_r wsa_r = new WStringArray_r
            {
                Values = (uint)Marshal.ReadInt32(ptr)
            };

            if (wsa_r.Values == 0)
            {
                wsa_r.LppszW = null;
            }
            else
            {
                wsa_r.LppszW = new string[wsa_r.Values];
                IntPtr szwaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                for (uint i = 0; i < wsa_r.Values; i++)
                {
                    wsa_r.LppszW[i] = Marshal.PtrToStringUni(new IntPtr(Marshal.ReadInt32(szwaddr)));
                    szwaddr = new IntPtr(szwaddr.ToInt32() + 4);
                }
            }

            return wsa_r;
        }

        /// <summary>
        /// Parse FlatUIDArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        /// <returns>Instance of FlatUIDArray_r structure.</returns>
        public static FlatUIDArray_r ParseFlatUIDArray_r(IntPtr ptr)
        {
            FlatUIDArray_r fuida_r = new FlatUIDArray_r
            {
                Values = (uint)Marshal.ReadInt32(ptr)
            };

            if (fuida_r.Values == 0)
            {
                fuida_r.Guid = null;
            }
            else
            {
                fuida_r.Guid = new FlatUID_r[fuida_r.Values];
                IntPtr fuidaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                for (uint i = 0; i < fuida_r.Values; i++)
                {
                    fuida_r.Guid[i] = ParseFlatUID_r(new IntPtr(Marshal.ReadInt32(fuidaddr)));
                    fuidaddr = new IntPtr(fuidaddr.ToInt32() + 4);
                }
            }

            return fuida_r;
        }

        /// <summary>
        /// Parse PROP_VAL_UNION structure.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        /// <param name="proptype">Property Types.</param>
        /// <returns>Instance of PROP_VAL_UNION structure.</returns>
        public static PROP_VAL_UNION ParsePROP_VAL_UNION(IntPtr ptr, PropertyType proptype)
        {
            PROP_VAL_UNION pvu = new PROP_VAL_UNION();

            switch (proptype)
            {
                case PropertyType.PtypInteger16:
                    pvu.I = Marshal.ReadInt16(ptr);
                    break;

                case PropertyType.PtypInteger32:
                    pvu.L = Marshal.ReadInt32(ptr);
                    break;

                case PropertyType.PtypBoolean:
                    pvu.B = (ushort)Marshal.ReadInt16(ptr);
                    break;

                case PropertyType.PtypString8:
                    pvu.LpszA = Marshal.PtrToStringAnsi(new IntPtr(Marshal.ReadInt32(ptr)));
                    break;

                case PropertyType.PtypBinary:
                    pvu.Bin = ParseBinary_r(ptr);
                    break;

                case PropertyType.PtypString:
                    pvu.LpszW = Marshal.PtrToStringUni(new IntPtr(Marshal.ReadInt32(ptr)));
                    break;

                case PropertyType.PtypGuid:
                    IntPtr uidaddr = new IntPtr(Marshal.ReadInt32(ptr));
                    if (uidaddr == IntPtr.Zero)
                    {
                        pvu.Guid = null;
                    }
                    else
                    {
                        pvu.Guid = new FlatUID_r[1];
                        pvu.Guid[0] = ParseFlatUID_r(uidaddr);
                    }

                    break;

                case PropertyType.PtypTime:
                    pvu.FileTime.LowDateTime = (uint)Marshal.ReadInt32(ptr);
                    pvu.FileTime.HighDateTime = (uint)Marshal.ReadInt32(ptr, sizeof(uint));
                    break;

                case PropertyType.PtypErrorCode:
                    pvu.ErrorCode = Marshal.ReadInt32(ptr);
                    break;

                case PropertyType.PtypMultipleInteger16:
                    pvu.MVi = ParseShortArray_r(ptr);
                    break;

                case PropertyType.PtypMultipleInteger32:
                    pvu.MVl = ParseLongArray_r(ptr);
                    break;

                case PropertyType.PtypMultipleString8:
                    pvu.MVszA = ParseStringArray_r(ptr);
                    break;

                case PropertyType.PtypMultipleBinary:
                    pvu.MVbin = ParseBinaryArray_r(ptr);
                    break;

                case PropertyType.PtypMultipleString:
                    pvu.MVszW = ParseWStringArray_r(ptr);
                    break;

                case PropertyType.PtypMultipleGuid:
                    pvu.MVguid = ParseFlatUIDArray_r(ptr);
                    break;

                case PropertyType.PtypNull:
                case PropertyType.PtypEmbeddedTable:
                    pvu.Reserved = Marshal.ReadInt32(ptr);
                    break;

                default:
                    throw new ParseException("Parsing PROP_VAL_UNION failed!");
            }

            return pvu;
        }

        /// <summary>
        /// Parse PropertyValue_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        /// <returns>Instance of PropertyValue_r structure.</returns>
        public static PropertyValue_r ParsePropertyValue_r(IntPtr ptr)
        {
            PropertyValue_r protertyValue = new PropertyValue_r();

            int offset = 0;

            protertyValue.PropTag = (uint)Marshal.ReadInt32(ptr, offset);
            offset += 4;

            protertyValue.Reserved = (uint)Marshal.ReadInt32(ptr, offset);
            offset += 4;

            protertyValue.Value = ParsePROP_VAL_UNION(new IntPtr(ptr.ToInt32() + offset), (PropertyType)(protertyValue.PropTag & 0x0000FFFF));

            return protertyValue;
        }

        /// <summary>
        /// Parse PropertyRow_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        /// <returns>Instance of PropertyRow_r structure.</returns>
        public static PropertyRow_r ParsePropertyRow_r(IntPtr ptr)
        {
            PropertyRow_r protertyRow = new PropertyRow_r();
            int offset = 0;

            protertyRow.Reserved = (uint)Marshal.ReadInt32(ptr);
            offset += sizeof(uint);

            protertyRow.Values = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(uint);

            if (protertyRow.Values == 0)
            {
                protertyRow.Props = null;
            }
            else
            {
                protertyRow.Props = new PropertyValue_r[protertyRow.Values];
                IntPtr pvaddr = new IntPtr(Marshal.ReadInt32(ptr, offset));
                const int PropertyValueLengthInBytes = 16;
                for (uint i = 0; i < protertyRow.Values; i++)
                {
                    protertyRow.Props[i] = ParsePropertyValue_r(pvaddr);
                    pvaddr = new IntPtr(pvaddr.ToInt32() + PropertyValueLengthInBytes);
                }
            }

            return protertyRow;
        }

        /// <summary>
        /// Parse PropertyRowSet_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        /// <returns>Instance of PropertyRowSet_r structure.</returns>
        public static PropertyRowSet_r ParsePropertyRowSet_r(IntPtr ptr)
        {
            PropertyRowSet_r prs_r = new PropertyRowSet_r
            {
                Rows = (uint)Marshal.ReadInt32(ptr)
            };

            const int PropertyRowLengthInBytes = 12;

            if (prs_r.Rows == 0)
            {
                prs_r.PropertyRowSet = null;
            }
            else
            {
                ptr = new IntPtr(ptr.ToInt32() + sizeof(uint));
                prs_r.PropertyRowSet = new PropertyRow_r[prs_r.Rows];
                for (uint i = 0; i < prs_r.Rows; i++)
                {
                    prs_r.PropertyRowSet[i] = ParsePropertyRow_r(ptr);
                    ptr = new IntPtr(ptr.ToInt32() + PropertyRowLengthInBytes);
                }
            }

            return prs_r;
        }

        /// <summary>
        /// Parse TRowSet structure.
        /// </summary>
        /// <param name="data">Original data.</param>
        /// <param name="codePage">CodePage number.</param>
        /// <returns>TROW array.</returns>
        public static TRowSet ParseTRowSet(byte[] data, uint codePage)
        {
            TRowSet temp = new TRowSet();
            int index = 0;
            try
            {
                temp.Type = BitConverter.ToUInt32(data, index);
                index += sizeof(uint);
                temp.RowCount = BitConverter.ToUInt32(data, index);
                index += sizeof(uint);
                temp.Rows = new TROW[temp.RowCount];
                for (int i = 0; i < temp.RowCount; i++)
                {
                    temp.Rows[i].XPos = BitConverter.ToUInt32(data, index);
                    index += sizeof(uint);
                    temp.Rows[i].DeltaX = BitConverter.ToUInt32(data, index);
                    index += sizeof(uint);
                    temp.Rows[i].YPos = BitConverter.ToUInt32(data, index);
                    index += sizeof(uint);
                    temp.Rows[i].DeltaY = BitConverter.ToUInt32(data, index);
                    index += sizeof(uint);
                    temp.Rows[i].ControlType = BitConverter.ToUInt32(data, index);
                    index += sizeof(uint);
                    temp.Rows[i].ControlFlags = BitConverter.ToUInt32(data, index);
                    index += sizeof(uint);

                    temp.Rows[i].ControlStructure = new CNTRL
                    {
                        Type = BitConverter.ToUInt32(data, index)
                    };
                    index += sizeof(uint);
                    temp.Rows[i].ControlStructure.Size = BitConverter.ToUInt32(data, index);
                    index += sizeof(uint);
                    temp.Rows[i].ControlStructure.String = BitConverter.ToUInt32(data, index);
                    index += sizeof(uint);
                    byte[] stringValue = GetulStringValue(data, temp.Rows[i].ControlStructure.String);
                    if (null == stringValue)
                    {
                        temp.Rows[i].ControlStructure.StringValue = string.Empty;
                    }
                    else
                    {
                        temp.Rows[i].ControlStructure.StringValue = Encoding.GetEncoding((int)codePage).GetString(stringValue);
                    }
                }
            }
            catch (IndexOutOfRangeException)
            {
                return new TRowSet();
            }
            catch (NullReferenceException)
            {
                return new TRowSet();
            }

            return temp;
        }

        /// <summary>
        /// Parse Script.
        /// </summary>
        /// <param name="data">Original data.</param>
        /// <returns>Instance of ScriptInstruction.</returns>
        public static ScriptInstruction ParseScript(byte[] data)
        {
            ScriptInstruction temp = new ScriptInstruction();
            int index = 0;
            temp.Size = BitConverter.ToUInt32(data, index);

            // A uint32 field consumes 4 bytes.
            index += sizeof(uint);
            temp.ScriptData = new uint[temp.Size];
            for (int i = 0; i < temp.Size; i++)
            {
                temp.ScriptData[i] = BitConverter.ToUInt32(data, index);

                // A uint32 field consumes 4 bytes.
                index += sizeof(uint);
            }

            return temp;
        }

        /// <summary>
        /// Parse Script for Exchange server 2010.
        /// </summary>
        /// <param name="data">Original data.</param>
        /// <returns>Instance of ScriptInstruction.</returns>
        public static ScriptInstruction ParseScriptForEx2010(byte[] data)
        {
            ScriptInstruction temp = new ScriptInstruction();
            int index = 0;

            // Specifies the number of DWORDs of script data that follow. 
            temp.Size = (uint)(data.Length / sizeof(uint));
            temp.ScriptData = new uint[temp.Size];
            for (int i = 0; i < temp.Size; i++)
            {
                temp.ScriptData[i] = BitConverter.ToUInt32(data, index);

                // A uint32 field consumes 4 bytes.
                index += sizeof(uint);
            }

            return temp;
        }

        /// <summary>
        /// Parse Script Data.
        /// </summary>
        /// <param name="data">Original script data.</param>
        /// <returns>Operation set.</returns>
        public static List<Operation> ParseScriptData(uint[] data)
        {
            int index = 0;

            // Instruction set.
            List<Operation> result = new List<Operation>();
            Operation operation = new Operation();

            // Ten instructions value from open specification 2.2.2.2 Script Format.
            List<uint> set = new List<uint>
            {
                (uint)InstructionTypeValues.Emit_Property_Value,
                (uint)InstructionTypeValues.Emit_String,
                (uint)InstructionTypeValues.Emit_Upper_Property,
                (uint)InstructionTypeValues.Emit_Upper_String,
                (uint)InstructionTypeValues.Error,
                (uint)InstructionTypeValues.Halt,
                (uint)InstructionTypeValues.Jump,
                (uint)InstructionTypeValues.Jump_If_Equal_Properties,
                (uint)InstructionTypeValues.Jump_If_Equal_Values,
                (uint)InstructionTypeValues.Jump_If_Not_Exists
            };
            foreach (uint temp in data)
            {
                if (temp != (uint)InstructionTypeValues.Halt && set.Contains(temp))
                {
                    switch (temp)
                    {
                        // Jump If Not Exists Instruction.
                        case (uint)InstructionTypeValues.Jump_If_Not_Exists:
                            operation.OperationInstruction = data[index++];
                            operation.PropTag1 = data[index++];
                            operation.PropTag2 = null;
                            operation.Offset = data[index++];
                            result.Add(operation);
                            break;

                        // Emit Property Value Instruction.
                        case (uint)InstructionTypeValues.Emit_Property_Value:
                            operation.OperationInstruction = data[index++];
                            operation.PropTag1 = data[index++];
                            operation.PropTag2 = null;
                            operation.Offset = null;
                            result.Add(operation);
                            break;

                        // Error Instruction.
                        case (uint)InstructionTypeValues.Error:
                            operation.OperationInstruction = data[index++];
                            operation.PropTag1 = null;
                            operation.PropTag2 = null;
                            operation.Offset = null;
                            result.Add(operation);
                            break;

                        // Emit String Instruction.
                        case (uint)InstructionTypeValues.Emit_String:
                            operation.OperationInstruction = data[index++];
                            operation.PropTag1 = data[index++];
                            operation.PropTag2 = null;
                            operation.Offset = null;
                            result.Add(operation);
                            break;

                        // Jump Instruction.
                        case (uint)InstructionTypeValues.Jump:
                            operation.OperationInstruction = data[index++];
                            operation.PropTag1 = null;
                            operation.PropTag2 = null;
                            operation.Offset = data[index++];
                            result.Add(operation);
                            break;

                        // Jump If Equal Properties Instruction.
                        case (uint)InstructionTypeValues.Jump_If_Equal_Properties:
                            operation.OperationInstruction = data[index++];
                            operation.PropTag1 = data[index++];
                            operation.PropTag2 = data[index++];
                            operation.Offset = data[index++];
                            result.Add(operation);
                            break;

                        // Jump If Equal Values Instruction.
                        case (uint)InstructionTypeValues.Jump_If_Equal_Values:
                            operation.OperationInstruction = data[index++];
                            operation.PropTag1 = data[index++];
                            operation.PropTag2 = data[index++];
                            operation.Offset = data[index++];
                            result.Add(operation);
                            break;

                        // Emit Upper String Instruction.
                        case (uint)InstructionTypeValues.Emit_Upper_String:
                            operation.OperationInstruction = data[index++];
                            operation.PropTag1 = data[++index];
                            operation.PropTag2 = null;
                            operation.Offset = null;
                            result.Add(operation);
                            break;

                        // Emit Upper Property Instruction.
                        case (uint)InstructionTypeValues.Emit_Upper_Property:
                            operation.OperationInstruction = data[index++];
                            operation.PropTag1 = data[index++];
                            operation.PropTag2 = null;
                            operation.Offset = null;
                            result.Add(operation);
                            break;
                        default:

                            // Do nothing.
                            break;
                    }
                }
                else if (temp == (uint)InstructionTypeValues.Halt)
                {
                    // Halt Instruction.
                    // When this instruction is encountered, the script has finished and was successful.
                    operation.OperationInstruction = data[index++];
                    operation.PropTag1 = null;
                    operation.PropTag2 = null;
                    operation.Offset = null;
                    result.Add(operation);
                    break;
                }
                else
                {
                    // Do nothing.
                }
            }

            return result;
        }

        /// <summary>
        /// Get ulString Value.
        /// </summary>
        /// <param name="data">Original data.</param>
        /// <param name="offSet">Offset value.</param>
        /// <returns>ulString Value.</returns>
        private static byte[] GetulStringValue(byte[] data, uint offSet)
        {
            int length = 0;
            uint index = offSet;

            // A null-terminated non-Unicode string. 
            while ((data[index] != 0x00) && (index < data.Length))
            {
                length++;
                index++;
            }

            if (0 == length)
            {
                return null;
            }
            else
            {
                byte[] temp = new byte[length];
                Array.Copy(data, offSet, temp, 0, length);
                return temp;
            }
        }
    }
}