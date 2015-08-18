namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;
    using System.IO;

    /// <summary>
    /// Contain a GLOBCNT range.
    /// </summary>
    public struct GLOBCNTRange
    {
        /// <summary>
        /// The start GLOBCNT.
        /// </summary>
        private GLOBCNT startGLOBCNT;

        /// <summary>
        /// The end GLOBCNT.
        /// </summary>
        private GLOBCNT endGLOBCNT;

        /// <summary>
        /// Initializes a new instance of the GLOBCNTRange structure.
        /// </summary>
        /// <param name="start">The starting GLOBCNT of the range.</param>
        /// <param name="end">The end GLOBCNT of the range.</param>
        public GLOBCNTRange(GLOBCNT start, GLOBCNT end)
        {
            if (start > end)
            {
                AdapterHelper.Site.Assert.Fail("The start GLOBCNT should not large than the end GLOBCNT.");
            }

            this.startGLOBCNT = start;
            this.endGLOBCNT = end;
        }

        /// <summary>
        /// Gets or sets the start GLOBCNT.
        /// </summary>
        public GLOBCNT StartGLOBCNT
        {
            get { return this.startGLOBCNT; }
            set { this.startGLOBCNT = value; }
        }

        /// <summary>
        /// Gets or sets the end GLOBCNT.
        /// </summary>
        public GLOBCNT EndGLOBCNT
        {
            get { return this.endGLOBCNT; }
            set { this.endGLOBCNT = value; }
        }

        /// <summary>
        /// Gets a value indicating whether this range a singleton.
        /// </summary>
        public bool IsSingleton
        {
            get
            {
                return this.startGLOBCNT == this.endGLOBCNT;
            }
        }

        /// <summary>
        /// Indicates whether has the GLOBCNT.
        /// </summary>
        /// <param name="cnt">The GLOBCNT.</param>
        /// <returns>If the GLOBCNT is in this range, return true, else false.</returns>
        public bool Contains(GLOBCNT cnt)
        {
            return this.StartGLOBCNT <= cnt && this.EndGLOBCNT >= cnt;
        }

        /// <summary>
        /// Indicates whether has the GLOBCNTRange.
        /// </summary>
        /// <param name="range">A GLOBCNTRange.</param>
        /// <returns>If the GLOBCNTRange is in this range, return true, else false.</returns>
        public bool Contains(GLOBCNTRange range)
        {
            return range.StartGLOBCNT >= this.StartGLOBCNT
                && range.EndGLOBCNT <= this.EndGLOBCNT;
        }

        /// <summary>
        /// Get the same high order bytes of this range.
        /// </summary>
        /// <returns>The same high order bytes.</returns>
        public byte[] GetSameHighOrderValues()
        {
            int i = 0;
            byte[] tmp1 = StructureSerializer.Serialize(this.StartGLOBCNT);
            byte[] tmp2 = StructureSerializer.Serialize(this.EndGLOBCNT);
            while (i < tmp1.Length && tmp1[i] == tmp2[i])
            {
                i++;
            }

            byte[] r = new byte[i];
            Array.Copy(tmp1, r, i);
            return r;
        }
    }

    /// <summary>
    /// A GLOBSET is a set of GLOBCNT values 
    /// that are typically reduced to one or more GLOBCNT ranges. 
    /// </summary>
    [SerializableObjectAttribute(true, true)]
    public class GLOBSET : SerializableBase
    {
        /// <summary>
        /// Contains common bytes.
        /// </summary>
        private CommonByteStack stack;

        /// <summary>
        /// A stream instance which is used to deserialize object.
        /// </summary>
        private Stream stream;

        /// <summary>
        /// A list of GLOBCNTs.
        /// </summary>
        private List<GLOBCNT> globcntList;

        /// <summary>
        /// Contains GLOBCNTRanges in the stream.
        /// </summary>
        private List<GLOBCNTRange> globcntRangeList;

        /// <summary>
        /// A list of command generated while deserializing.
        /// </summary>
        private List<Command> deserializedcommandList;

        /// <summary>
        /// Indicates whether all GLOBCNTs in GLOBSET when serializing.
        /// </summary>
        private bool isAllGLOBCNTInGLOBSET;

        /// <summary>
        /// Gets a value indicating whether all Duplicate GLOBCNTs removed when serializing.
        /// </summary>
        private bool hasAllDuplicateGLOBCNTRemoved;

        /// <summary>
        /// Indicates whether all GLOBCNTs are arranged from lowest to highest.
        /// </summary>
        private bool isAllGLOBCNTRanged;

        /// <summary>
        /// Indicates whether GLOBCNT values are grouped into consecutive ranges 
        /// with a low GLOBCNT value and a high GLOBCNT value.
        /// </summary>
        private bool hasGLOBCNTGroupedIntoRanges;

        /// <summary>
        /// Indicates whether GLOBCNT value which is disjoint is made into a singleton range 
        /// with the low and high GLOBCNT values being the same.
        /// </summary>
        private bool isDisjointGLOBCNTMadeIntoSingleton;

        /// <summary>
        /// Command types.
        /// </summary>
        private enum Operation : byte
        {
            /// <summary>
            /// Represent the push1 command.
            /// </summary>
            Push1 = 0x01,

            /// <summary>
            /// Represent the push2 command.
            /// </summary>
            Push2 = 0x02,

            /// <summary>
            /// Represent the push3 command.
            /// </summary>
            Push3 = 0x03,

            /// <summary>
            /// Represent the push4 command.
            /// </summary>
            Push4 = 0x04,
            
            /// <summary>
            /// Represent the push5 command.
            /// </summary>
            Push5 = 0x05,

            /// <summary>
            /// Represent the push6 command.
            /// </summary>
            Push6 = 0x06,

            /// <summary>
            /// Represent the pop command.
            /// </summary>
            Pop = 0x50,

            /// <summary>
            /// Represent the bitmask command.
            /// </summary>
            Bitmask = 0x42,

            /// <summary>
            /// Represent the range command.
            /// </summary>
            Range = 0x52,

            /// <summary>
            /// Represent the end command.
            /// </summary>
            End = 0x00
        }

        /// <summary>
        /// Gets a value indicating whether GLOBCNT value which is disjoint is made into a singleton range 
        /// with the low and high GLOBCNT values being the same.
        /// </summary>
        public bool IsDisjointGLOBCNTMadeIntoSingleton
        {
            get
            {
                return this.isDisjointGLOBCNTMadeIntoSingleton;
            }
        }

        /// <summary>
        /// Gets or sets the GLOBCNTList.
        /// </summary>
        public List<GLOBCNT> GLOBCNTList
        {
            get
            {
                // Too many GLOBCNT from a range.
                if (this.globcntRangeList != null)
                {
                    this.globcntList = GetGLOBCNTList(this.globcntRangeList);
                }

                return this.globcntList;
            }

            set 
            {
                this.globcntList = value;
            }
        }

        /// <summary>
        /// Gets or sets the GLOBCNTRangeList which contains GLOBCNTRanges in the stream.
        /// </summary>
        public List<GLOBCNTRange> GLOBCNTRangeList
        {
            get { return this.globcntRangeList; }
            set { this.globcntRangeList = value; }
        }

        /// <summary>
        /// Gets a list of command generated while deserializing.
        /// </summary>
        public List<Command> DeserializedCommandList
        {
            get { return this.deserializedcommandList; }
        }

        /// <summary>
        /// Gets a value indicating whether all GLOBCNT in GLOBSET when serializing.
        /// </summary>
        public bool IsAllGLOBCNTInGLOBSET
        {
            get
            {
                return this.isAllGLOBCNTInGLOBSET;
            }
        }

        /// <summary>
        /// Gets a value indicating whether all Duplicate GLOBCNT removed when serializing.
        /// </summary>
        public bool HasAllDuplicateGLOBCNTRemoved
        {
            get
            {
                return this.hasAllDuplicateGLOBCNTRemoved;
            }
        }

        /// <summary>
        /// Gets a value indicating whether all GLOBCNTs are arranged from lowest to highest.
        /// </summary>
        public bool IsAllGLOBCNTRanged
        {
            get
            {
                return this.isAllGLOBCNTRanged;
            }
        }

        /// <summary>
        /// Gets a value indicating whether GLOBCNT values are grouped into consecutive ranges 
        /// with a low GLOBCNT value and a high GLOBCNT value
        /// </summary>
        public bool HasGLOBCNTGroupedIntoRanges
        {
            get
            {
                return this.hasGLOBCNTGroupedIntoRanges;
            }
        }

        /// <summary>
        /// Get GLOBCNTs from a GLOBCNTRange list.
        /// </summary>
        /// <param name="rangeList">A GLOBCNTRange list.</param>
        /// <returns>A GLOBCNT list corresponding to the GLOBCNTRange list.</returns>
        public static List<GLOBCNT> GetGLOBCNTList(List<GLOBCNTRange> rangeList)
        {
            List<GLOBCNT> cnts = new List<GLOBCNT>();
            foreach (GLOBCNTRange range in rangeList)
            {
                GLOBCNT tmp = range.StartGLOBCNT;
                cnts.Add(tmp);
                tmp = GLOBCNT.Inc(tmp);
                while (tmp <= range.EndGLOBCNT)
                {
                    cnts.Add(tmp);
                    tmp = GLOBCNT.Inc(tmp);
                }
            }

            return cnts;
        }

        /// <summary>
        /// Get GLOBCNTRanges from a GLOBCNT list.
        /// </summary>
        /// <param name="globcntList">A GLOBCNT list.</param>
        /// <returns>A GLOBCNTRange list corresponding to the GLOBCNT list.</returns>
        public static List<GLOBCNTRange> GetGLOBCNTRange(List<GLOBCNT> globcntList)
        {
            // _REPLID = id;
            int i, j;
            List<GLOBCNT> list = new List<GLOBCNT>();
            List<GLOBCNTRange> globSETRangeList = new List<GLOBCNTRange>();

            // Do a copy.
            for (i = 0; i < globcntList.Count; i++)
            {
                list.Add(globcntList[i]);
            }

            // Remove all the duplicate GLOBCNT values.
            for (i = 0; i < list.Count - 1; i++)
            {
                j = i + 1;
                while (j < list.Count)
                {
                    if (list[i] == list[j])
                    {
                        list.RemoveAt(j);
                        continue;
                    }
                    else
                    {
                        j++;
                    }
                }
            }

            // Sort GLOBCNT.
            list.Sort(new Comparison<GLOBCNT>(delegate(GLOBCNT c1, GLOBCNT c2)
            {
                if (c1 < c2)
                {
                    return -1;
                }

                if (c1 > c2)
                {
                    return 1;
                }

                return 0;
            }));

            // Make a GLOBCNTRange list.
            i = 0;
            while (i < list.Count)
            {
                GLOBCNT start = list[i];
                GLOBCNT end = start;
                GLOBCNT next = end;
                j = i + 1;
                while (j < list.Count)
                {
                    end = next;
                    next = GLOBCNT.Inc(end);
                    if (list[j] == next)
                    {
                        list.RemoveAt(j);
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }

                globSETRangeList.Add(new GLOBCNTRange(start, end));
                i++;
            }

            return globSETRangeList;
        }

        /// <summary>
        /// Serialize fields to a stream.
        /// </summary>
        /// <param name="stream">The stream where serialized instance will be wrote.</param>
        /// <returns>Bytes written to the stream.</returns>
        public override int Serialize(Stream stream)
        {
            int bytesWriren = 0;
            this.stream = stream;
            this.deserializedcommandList = null;
            this.stack = new CommonByteStack();
            this.globcntRangeList = GetGLOBCNTRange(this.GLOBCNTList);
            this.isAllGLOBCNTInGLOBSET = true;
            this.isAllGLOBCNTRanged = true;
            this.isDisjointGLOBCNTMadeIntoSingleton = true;
            this.hasAllDuplicateGLOBCNTRemoved = true;
            this.hasGLOBCNTGroupedIntoRanges = true;
            bytesWriren += this.Compress(0, this.globcntRangeList.Count - 1);
            bytesWriren += this.End(stream);
            return bytesWriren;
        }

        /// <summary>
        /// Deserialize fields in this class from a stream.
        /// </summary>
        /// <param name="stream">Stream contains a serialized instance of this class.</param>
        /// <param name="size">How many bytes can read if -1, no limitation.MUST be -1.</param>
        /// <returns>Bytes have been read from the stream.</returns>
        public override int Deserialize(Stream stream, int size)
        {
            AdapterHelper.Site.Assert.AreEqual(-1, size, "The size value MUST be -1, but the actual value is {0}.", size);

            int bytesRead = 0;
            this.stream = stream;
            this.globcntList = new List<GLOBCNT>();
            this.globcntRangeList = new List<GLOBCNTRange>();
            this.stack = new CommonByteStack();
            this.deserializedcommandList = new List<Command>();
            Operation op = this.ReadOperation();
            bytesRead += 1;
            while (op != Operation.End)
            {
                switch (op)
                {
                    case Operation.Bitmask:
                        if (this.stack.Bytes != 5)
                        {
                            AdapterHelper.Site.Assert.Fail("The deserialization operation should be successful.");
                        }
                        else
                        {
                            byte[] commonBytes = stack.GetCommonBytes();
                            byte startValue, bitmask;
                            bytesRead += ReadBitmaskValue(out startValue, out bitmask);
                            List<GLOBCNTRange> tmp = FromBitmask(commonBytes, startValue, bitmask);
                            BitmaskCommand bmCmd =
                                new BitmaskCommand((byte)op, startValue, bitmask)
                                {
                                    CorrespondingGLOBCNTRangeList = tmp
                                };
                            deserializedcommandList.Add(bmCmd);
                            for (int i = 0; i < tmp.Count; i++)
                            {
                                globcntRangeList.Add(tmp[i]);
                            }

                            tmp = null;
                        }

                        break;
                    case Operation.End:
                        this.deserializedcommandList.Add(new EndCommand((byte)op));
                        return bytesRead;
                    case Operation.Pop:
                        this.deserializedcommandList.Add(new PopCommand((byte)op));
                        this.stack.Pop();
                        break;
                    case Operation.Range:
                        {
                            byte[] lowValue, highValue;
                            bytesRead += this.ReadRangeValue(out lowValue, out highValue);
                            GLOBCNTRange range = this.FromRange(
                                this.stack.GetCommonBytes(),
                                lowValue,
                                highValue);
                            List<GLOBCNTRange> rangeList = new List<GLOBCNTRange>
                            {
                                range
                            };
                            RangeCommand rngCmd =
                                new RangeCommand((byte)op, lowValue, highValue)
                                {
                                    CorrespondingGLOBCNTRangeList = rangeList
                                };

                            this.deserializedcommandList.Add(rngCmd);
                            this.globcntRangeList.Add(range);
                        }

                        break;
                    case Operation.Push1:
                    case Operation.Push2:
                    case Operation.Push3:
                    case Operation.Push4:
                    case Operation.Push5:
                    case Operation.Push6:
                        int pushByteCount = (int)op;
                        byte[] pushBytes;
                        bytesRead += this.ReadPushedValue(pushByteCount, out pushBytes);
                        PushCommand pshCmd = new PushCommand((byte)op, pushBytes);
                        if (6 == pushByteCount + this.stack.Bytes)
                        {
                            List<GLOBCNTRange> rangeList = new List<GLOBCNTRange>();
                            GLOBCNTRange range = this.FromPush(this.stack.GetCommonBytes(), pushBytes);
                            rangeList.Add(range);
                            this.globcntRangeList.Add(range);
                            pshCmd.CorrespondingGLOBCNTRangeList = rangeList;
                            this.deserializedcommandList.Add(pshCmd);
                            break;
                        }

                        this.deserializedcommandList.Add(pshCmd);
                        this.stack.Push(pushBytes);
                        break;
                    default:
                        AdapterHelper.Site.Assert.Fail("The operation get from the stream is invalid, its value is {0}.", op);
                        break;
                }

                op = this.ReadOperation();
                bytesRead += 1;
            }

            if (op == Operation.End)
            {
                this.deserializedcommandList.Add(new EndCommand((byte)op));
            }

            this.isAllGLOBCNTInGLOBSET = true;
            this.isAllGLOBCNTRanged = true;
            this.isDisjointGLOBCNTMadeIntoSingleton = true;
            this.hasAllDuplicateGLOBCNTRemoved = true;
            this.hasGLOBCNTGroupedIntoRanges = true;
            return bytesRead;
        }

        /// <summary>
        /// Verifies a condition.
        /// </summary>
        /// <param name="condition">A boolean value.</param>
        private void Verify(bool condition)
        {
            AdapterHelper.Site.Assert.IsTrue(condition, "The condition should be true.");
        }

        /// <summary>
        /// Writes a push command.
        /// </summary>
        /// <param name="stream">A Stream object.</param>
        /// <param name="values">A byte array contains values of the push command.</param>
        /// <returns>The count of bytes have been wrote to the stream.</returns>
        private int Push(Stream stream, byte[] values)
        {
            int size = 0;
            if (values == null
                || values.Length == 0
                || values.Length > 6)
            {
                AdapterHelper.Site.Assert.Fail("The value of the parameter of \"values\" in Push method is invalid.");
            }

            stream.WriteByte((byte)values.Length);
            size += 1;
            stream.Write(values, 0, values.Length);
            size += values.Length;
            return size;
        }

        /// <summary>
        /// Writes a pop command to a stream.
        /// </summary>
        /// <param name="stream">A stream object.</param>
        /// <returns>The count of bytes have been wrote to the stream.</returns>
        private int Pop(Stream stream)
        {
            stream.WriteByte((byte)Operation.Pop);
            return 1;
        }

        /// <summary>
        /// Writes an end command to a stream.
        /// </summary>
        /// <param name="stream">A stream object.</param>
        /// <returns>The count of bytes have been wrote to the stream.</returns>
        private int End(Stream stream)
        {
            stream.WriteByte((byte)Operation.End);
            return 1;
        }

        /// <summary>
        /// Writes a bitmask command to a stream.
        /// </summary>
        /// <param name="stream">A Stream object.</param>
        /// <param name="startValue">The start value of the push command.</param>
        /// <param name="bitmask">The bitmask of the push command.</param>
        /// <returns>The count of bytes have been wrote to the stream.</returns>
        private int Bitmask(Stream stream, byte startValue, byte bitmask)
        {
            stream.WriteByte((byte)Operation.Bitmask);
            stream.WriteByte(startValue);
            stream.WriteByte(bitmask);
            return 3;
        }

        /// <summary>
        /// Writes a range command to a stream.
        /// </summary>
        /// <param name="stream">A Stream object.</param>
        /// <param name="lowValue">The lowValue of a range command.</param>
        /// <param name="highValue">The highValue of a range command.</param>
        /// <returns>The count of bytes have been wrote to the stream.</returns>
        private int Range(Stream stream, byte[] lowValue, byte[] highValue)
        {
            if (lowValue == null
                || highValue == null
                || lowValue.Length != highValue.Length
                || lowValue.Length == 0
                || lowValue.Length > 6)
            {
                AdapterHelper.Site.Assert.Fail("The values of highValue or lowValue arguments in Range method are invalid.");
            }

            stream.WriteByte((byte)Operation.Range);
            stream.Write(lowValue, 0, lowValue.Length);
            stream.Write(highValue, 0, highValue.Length);
            return 1 + (2 * lowValue.Length);
        }

        /// <summary>
        /// Reads a byte from the stream.
        /// </summary>
        /// <returns>A byte value read from the stream.</returns>
        private byte ReadByte()
        {
            return StreamHelper.ReadUInt8(this.stream);
        }

        /// <summary>
        /// Reads a bitmask command from the stream.
        /// </summary>
        /// <param name="startingValue">The start value of the bitmask command.</param>
        /// <param name="bitmask">The bitmask of the bitmask command.</param>
        /// <returns>The count of bytes have been read.</returns>
        private int ReadBitmaskValue(out byte startingValue, out byte bitmask)
        {
            startingValue = this.ReadByte();
            bitmask = this.ReadByte();
            return 2;
        }

        /// <summary>
        /// Reads a range command from the stream.
        /// </summary>
        /// <param name="lowValue">The low value of the range command.</param>
        /// <param name="highValue">The high value of the range command.</param>
        /// <returns>The count of bytes have been read from the stream.</returns>
        private int ReadRangeValue(out byte[] lowValue, out byte[] highValue)
        {
            int size = 6 - this.stack.Bytes;
            lowValue = new byte[size];
            highValue = new byte[size];
            this.stream.Read(lowValue, 0, size);
            this.stream.Read(highValue, 0, size);
            return size * 2;
        }

        /// <summary>
        /// Reads a push command from the stream.
        /// </summary>
        /// <param name="size">The size of the push command.</param>
        /// <param name="commonBytes">The common bytes in the push command.</param>
        /// <returns>The count of bytes have been read from the stream.</returns>
        private int ReadPushedValue(int size, out byte[] commonBytes)
        {
            commonBytes = new byte[size];
            this.stream.Read(commonBytes, 0, size);
            return size;
        }

        /// <summary>
        /// Reads an operation byte from a stream.
        /// </summary>
        /// <returns>An operation enumeration.</returns>
        private Operation ReadOperation()
        {
            return (Operation)this.ReadByte();
        }

        /// <summary>
        /// Gets high order command bytes in the GLOBCNTRange list.
        /// </summary>
        /// <param name="startIndex">The start index of the GLOBCNTRange list to get high order common bytes.</param>
        /// <param name="endIndex">The end index of the GLOBCNTRange list to get high order common bytes.</param>
        /// <param name="byteIndex">Specifies the index of GLOBCNTRange's common bytes to compare from.</param>
        /// <param name="firstDiffIndex">The index of the first GLOBCNTRange which have different byte with previous ones.</param>
        /// <returns>A byte array contain common bytes.</returns>
        private byte[] HighOrderCommonBytes(
            int startIndex,
            int endIndex,
            int byteIndex,
            out int firstDiffIndex)
        {
            this.Verify(startIndex < endIndex);
            int len = 0;
            int i, j = endIndex;
            bool hasDiff = false;
            byte[] firstRangeCommonBytes = this.globcntRangeList[startIndex]
                .GetSameHighOrderValues();
            for (i = byteIndex;
                i < firstRangeCommonBytes.Length && !hasDiff;
                i++)
            {
                byte common = firstRangeCommonBytes[i];
                for (j = startIndex + 1; j <= endIndex; j++)
                {
                    byte[] bytes = this.globcntRangeList[j]
                        .GetSameHighOrderValues();
                    if (bytes[i] != common)
                    {
                        hasDiff = true;
                        break;
                    }
                }

                if (hasDiff)
                {
                    break;
                }
            }

            len = hasDiff ? i - byteIndex : firstRangeCommonBytes.Length - byteIndex;
            firstDiffIndex = hasDiff ? j : endIndex + 1;
            byte[] r = new byte[len];
            Array.Copy(firstRangeCommonBytes, byteIndex, r, 0, len);
            return r;
        }

        /// <summary>
        /// Compresses GLOBCNTRanges in the GLOBCNTRange list to as a bitmask command.
        /// </summary>
        /// <param name="startIndex">The start index of the GLOBCNTRange in the GLOBCNTRange list to compress.</param>
        /// <param name="endIndex">The end index of the GLOBCNTRange in the GLOBCNTRange list to compress.</param>
        /// <returns>The count of bytes have been wrote to the stream.</returns>
        private int CompressBitmask(int startIndex, int endIndex)
        {
            this.Verify(startIndex <= endIndex);
            List<GLOBCNT> list = GetGLOBCNTList(
                this.globcntRangeList.GetRange(startIndex, endIndex - startIndex + 1));
            byte bitmask = 0;
            GLOBCNT tmp = list[0];
            byte startValue = tmp.Byte6;
            tmp = GLOBCNT.Inc(tmp);
            this.Verify(list.Count < 10);
            for (int i = 0; i < 9; i++)
            {
                if (list.Contains(tmp))
                {
                    bitmask |= checked((byte)(1 << i));
                }

                tmp = GLOBCNT.Inc(tmp);
            }

            return this.Bitmask(this.stream, startValue, bitmask);
        }

        /// <summary>
        /// Compresses a GLOBCNTRange in the GLOBCNTRange list to as a bitmask command.
        /// </summary>
        /// <param name="index">The index of the GLOBCNTRange.</param>
        /// <returns>The count of bytes have been wrote to the stream.</returns>
        private int CompressRange(int index)
        {
            int bytes = 6 - this.stack.Bytes;
            GLOBCNT cnt1 = this.globcntRangeList[index].StartGLOBCNT;
            GLOBCNT cnt2 = this.globcntRangeList[index].EndGLOBCNT;
            byte[] tmp1 = new byte[bytes];
            Array.Copy(
                StructureSerializer.Serialize(cnt1),
                6 - bytes,
                tmp1,
                0,
                bytes);
            byte[] tmp2 = new byte[bytes];
            Array.Copy(
                StructureSerializer.Serialize(
                cnt2),
                6 - bytes,
                tmp2,
                0,
                bytes);

            return this.Range(this.stream, tmp1, tmp2);
        }

        /// <summary>
        /// Compresses a singleton GLOBCNTRange.
        /// </summary>
        /// <param name="index">The index of the GLOBCNTRange.</param>
        /// <returns>The count of bytes have been wrote to the stream.</returns>
        private int CompressSingleton(int index)
        {
            this.Verify(this.globcntRangeList[index].StartGLOBCNT
                == this.globcntRangeList[index].EndGLOBCNT);
            int byteCount = 6 - this.stack.Bytes;
            GLOBCNT cnt = this.globcntRangeList[index].StartGLOBCNT;
            byte[] tmp = new byte[byteCount];
            Array.Copy(
                StructureSerializer.Serialize(cnt),
                6 - byteCount,
                tmp,
                0,
                byteCount);
            return this.Push(this.stream, tmp);
        }

        /// <summary>
        /// Compress the last byte(6th) of  GLOBCNTRanges.
        /// </summary>
        /// <param name="startIndex">The start index of GLOBCNTRanges to compress.</param>
        /// <param name="endIndex">The end index of GLOBCNTRanges to compress.</param>
        /// <returns>The count of bytes have been wrote to the stream.</returns>
        private int CompressLastByte(int startIndex, int endIndex)
        {
            this.Verify(this.stack.Bytes == 5);
            int bytesWriten = 0;
            int nextCompressIndex = startIndex;
            while (startIndex <= endIndex)
            {
                byte tmp1 = this.globcntRangeList[startIndex].StartGLOBCNT.Byte6;
                byte tmp2 = this.globcntRangeList[nextCompressIndex].EndGLOBCNT.Byte6;
                if (tmp2 - tmp1 >= 9)
                {
                    bytesWriten += this.CompressRange(nextCompressIndex);
                    nextCompressIndex++;
                }
                else
                {
                    nextCompressIndex = startIndex;
                    while (nextCompressIndex < endIndex)
                    {
                        tmp2 = this.globcntRangeList[nextCompressIndex + 1]
                            .EndGLOBCNT.Byte6;
                        if (tmp2 - tmp1 < 9)
                        {
                            nextCompressIndex++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    if (nextCompressIndex == startIndex)
                    {
                        if (this.globcntRangeList[startIndex].IsSingleton)
                        {
                            bytesWriten += this.CompressSingleton(startIndex);
                        }
                        else
                        {
                            bytesWriten += this.CompressRange(startIndex);
                        }
                    }
                    else
                    {
                        bytesWriten += this.CompressBitmask(
                            startIndex,
                            nextCompressIndex);
                    }

                    startIndex = nextCompressIndex + 1;
                }
            }

            return bytesWriten;
        }

        /// <summary>
        /// Compresses a GLOBCNTRange to the stream.
        /// </summary>
        /// <param name="index">The index of the GLOBCNTRange.</param>
        /// <returns>The count of bytes have been wrote to the stream.</returns>
        private int Compress(int index)
        {
            int bytesWriten = 0;
            if (this.globcntRangeList[index].IsSingleton)
            {
                bytesWriten += this.CompressSingleton(index);
            }
            else if (this.globcntRangeList[index].StartGLOBCNT
                < this.globcntRangeList[index].EndGLOBCNT)
            {
                bytesWriten += this.CompressRange(index);
            }

            return bytesWriten;
        }

        /// <summary>
        /// Compresses GLOBCNTRanges to the stream.
        /// </summary>
        /// <param name="startIndex">The start index.</param>
        /// <param name="endIndex">The end index.</param>
        /// <returns>The count of bytes have been wrote to the stream.</returns>
        private int Compress(int startIndex, int endIndex)
        {
            int bytesWriten = 0;
            bool pushed = false;
            this.Verify(this.stack.Bytes < 6);
            if (startIndex > this.globcntRangeList.Count)
            {
                this.Verify(false);
            }

            if (startIndex == endIndex)
            {
                this.Compress(startIndex);
            }
            else if (startIndex > endIndex)
            {
                AdapterHelper.Site.Assert.Fail(string.Format("The value of parameter 'startIndex'({0}) is bigger than the value of parameter 'endIndex'({1}).", startIndex, endIndex));
            }
            else if (startIndex < endIndex)
            {
                int firstDiffIndex;
                byte[] commonBytes = this.HighOrderCommonBytes(
                    startIndex,
                    endIndex,
                    this.stack.Bytes,
                    out firstDiffIndex);
                if (commonBytes.Length == 6 - this.stack.Bytes)
                {
                    // Two or more ranges have same bytes.
                    this.Verify(false);
                }
                else if (commonBytes.Length == 0)
                {
                    if (this.stack.Bytes == 5)
                    {
                        // Stack already has 5 bytes, ranges have different last byte.
                        bytesWriten += this.CompressLastByte(startIndex, endIndex);
                    }
                    else if ((endIndex + 1) != firstDiffIndex)
                    {
                        // Divide current ranges.
                        // Ranges' subset have same bytes.
                        bytesWriten += this.Compress(startIndex, firstDiffIndex - 1);
                        if (firstDiffIndex <= endIndex)
                        {
                            bytesWriten += this.Compress(firstDiffIndex, endIndex);
                        }
                    }
                    else
                    {
                        // The first range's same bytes are different with others.
                        bytesWriten += this.CompressSingleton(startIndex);
                        startIndex++;
                        if (startIndex <= endIndex)
                        {
                            bytesWriten += this.Compress(startIndex, endIndex);
                        }
                    }
                }
                else
                {
                    // FirstDiffIndex < endIndex.
                    this.stack.Push(commonBytes);
                    bytesWriten += this.Push(this.stream, commonBytes);
                    pushed = true;
                    if (this.stack.Bytes == 5)
                    {
                        bytesWriten += this.CompressLastByte(startIndex, endIndex);
                        this.stack.Pop();
                        bytesWriten += this.Pop(this.stream);
                        return bytesWriten;
                    }

                    bytesWriten += this.Compress(startIndex, firstDiffIndex - 1);
                    if (firstDiffIndex <= endIndex)
                    {
                        bytesWriten += this.Compress(firstDiffIndex, endIndex);
                    }
                }
            }

            if (pushed)
            {
                this.stack.Pop();
                bytesWriten += this.Pop(this.stream);
            }

            return bytesWriten;
        }

        /// <summary>
        /// Deserializes a GLOBCNTRange from a range command.
        /// </summary>
        /// <param name="comonBytes">The common bytes in the common byte stack.</param>
        /// <param name="lowBytes">The lowValue of the range command.</param>
        /// <param name="highBytes">The highValue of the range command.</param>
        /// <returns>A GLOBCNTRange.</returns>
        private GLOBCNTRange FromRange(
            byte[] comonBytes,
            byte[] lowBytes,
            byte[] highBytes)
        {
            this.Verify(comonBytes.Length + lowBytes.Length == 6
                    && lowBytes.Length == highBytes.Length);

            byte[] lowBuffer = new byte[6];
            byte[] highBuffer = new byte[6];

            Array.Copy(comonBytes, lowBuffer, comonBytes.Length);
            Array.Copy(lowBytes, 0, lowBuffer, comonBytes.Length, lowBytes.Length);

            Array.Copy(comonBytes, highBuffer, comonBytes.Length);
            Array.Copy(highBytes, 0, highBuffer, comonBytes.Length, highBytes.Length);

            return new GLOBCNTRange(
                StructureSerializer.Deserialize<GLOBCNT>(lowBuffer),
                StructureSerializer.Deserialize<GLOBCNT>(highBuffer));
        }

        /// <summary>
        /// Deserializes a list of GLOBCNTRanges from a bitmask command.
        /// </summary>
        /// <param name="commonBytes">The common bytes in the common byte stack.</param>
        /// <param name="startValue">The startValue of the bitmask command.</param>
        /// <param name="bitmask">The bitmaskValue of the bitmask command.</param>
        /// <returns>A list of GLOBCNTRanges.</returns>
        private List<GLOBCNTRange> FromBitmask(
            byte[] commonBytes,
            byte startValue,
            byte bitmask)
        {
            int bitIndex = 0;
            List<GLOBCNTRange> ranges = new List<GLOBCNTRange>();
            this.Verify(commonBytes.Length == 5);
            byte[] buffer = new byte[6];
            Array.Copy(commonBytes, buffer, 5);
            buffer[5] = startValue;
            byte start = startValue;
            GLOBCNT cnt1 = StructureSerializer.Deserialize<GLOBCNT>(buffer);
            GLOBCNT cnt2;
            do
            {
                if (((1 << bitIndex) & bitmask) == 0)
                {
                    // Use previous bitmask.
                    // ==bitIndex + 1 -1;
                    start = (byte)bitIndex;
                    start = checked((byte)(start + startValue));
                    buffer[5] = start;
                    cnt2 = StructureSerializer.Deserialize<GLOBCNT>(buffer);
                    ranges.Add(new GLOBCNTRange(cnt1, cnt2));
                    bitIndex++;
                    while (bitIndex < 8
                        && ((1 << bitIndex) & bitmask) == 0)
                    {
                        bitIndex++;
                    }

                    if (bitIndex == 8)
                    {
                        break;
                    }
                    else
                    {
                        start = (byte)(bitIndex + 1);
                        start = checked((byte)(start + startValue));
                        buffer[5] = start;
                        cnt1 = StructureSerializer.Deserialize<GLOBCNT>(buffer);
                        bitIndex++;
                    }
                }
                else
                {
                    bitIndex++;
                }
            } 
            while (bitIndex < 8);
            if (((1 << 7) & bitmask) != 0)
            {
                start = (byte)8;
                start = checked((byte)(8 + startValue));
                buffer[5] = start;
                cnt2 = StructureSerializer.Deserialize<GLOBCNT>(buffer);
                ranges.Add(new GLOBCNTRange(cnt1, cnt2));
            }

            return ranges;
        }

        /// <summary>
        /// Deserializes a GLOBCNTRange from a push command.
        /// </summary>
        /// <param name="comonBytes">The common bytes in the common byte stack.</param>
        /// <param name="pushedBytes">The bytes of the push command pushed.</param>
        /// <returns>A GLOBCNTRange.</returns>
        private GLOBCNTRange FromPush(byte[] comonBytes, byte[] pushedBytes)
        {
            this.Verify(pushedBytes.Length + comonBytes.Length == 6);
            byte[] pushedBuffer = new byte[6];
            Array.Copy(comonBytes, pushedBuffer, comonBytes.Length);
            Array.Copy(pushedBytes, 0, pushedBuffer, comonBytes.Length, pushedBytes.Length);
            return new GLOBCNTRange(
                StructureSerializer.Deserialize<GLOBCNT>(pushedBuffer),
                StructureSerializer.Deserialize<GLOBCNT>(pushedBuffer));
        }
    }
}