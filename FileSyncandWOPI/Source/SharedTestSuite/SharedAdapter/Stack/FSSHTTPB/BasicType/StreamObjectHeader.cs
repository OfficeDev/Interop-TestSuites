namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class specifies the base class for 16-bit or 32-bit stream object header start.
    /// </summary>
    public abstract class StreamObjectHeaderStart : BasicObject
    {
        /// <summary>
        /// Specify for 16-bit stream object header start.
        /// </summary>
        public const int StreamObjectHeaderStart16bit = 0;

        /// <summary>
        /// Specify for 32-bit stream object header start.
        /// </summary>
        public const int StreamObjectHeaderStart32bit = 2;

        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderStart class.
        /// </summary>
        protected StreamObjectHeaderStart()
        {
        }

        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderStart class with specified header type.
        /// </summary>
        /// <param name="streamObjectTypeHeaderStart">Specify the value of the StreamObjectHeaderStart Type.</param>
        protected StreamObjectHeaderStart(StreamObjectTypeHeaderStart streamObjectTypeHeaderStart)
        {
            this.Type = streamObjectTypeHeaderStart;
        }

        /// <summary>
        /// Gets or sets the type of the stream object.
        /// value 0 for 16-bit stream object header start, 
        /// value 2 for 32-bit stream object header start.
        /// </summary>
        public int HeaderType { get; set; }

        /// <summary>
        /// Gets or sets a value that specifies if set a compound parse type is needed and 
        /// MUST be ended with either an 8-bit stream object header end or a 16-bit stream object header end. 
        /// If the bit is zero, it specifies a single object. Otherwise it specifies a compound object.
        /// </summary>
        public int Compound { get; set; }

        /// <summary>
        /// Gets or sets a value that specifies the stream object type.
        /// </summary>
        public StreamObjectTypeHeaderStart Type { get; set; }

        /// <summary>
        /// Gets or sets a 15-bit unsigned integer that specifies the length in bytes for additional data (if any). 
        /// </summary>
        public int Length { get; set; }

        /// <summary>
        /// This method is used to parse the actual 16bit or 32bit stream header.
        /// </summary>
        /// <param name="byteArray">Specify the Byte array.</param>
        /// <param name="startIndex">Specify the start position.</param>
        /// <param name="streamObjectHeader">Specify the out value for the parse result.</param>
        /// <returns>Return true if success, otherwise returns false. </returns>
        public static int TryParse(byte[] byteArray, int startIndex, out StreamObjectHeaderStart streamObjectHeader)
        {
            uint headerType = (uint)(byteArray[startIndex] & 0x3);
            if (headerType == StreamObjectHeaderStart.StreamObjectHeaderStart16bit)
            {
                streamObjectHeader = new StreamObjectHeaderStart16bit();
            }
            else
            {
                if (headerType == StreamObjectHeaderStart.StreamObjectHeaderStart32bit)
                {
                    streamObjectHeader = new StreamObjectHeaderStart32bit();
                }
                else
                {
                    streamObjectHeader = null;
                    return 0;
                }
            }

            try
            {
                return streamObjectHeader.DeserializeFromByteArray(byteArray, startIndex);
            }
            catch (InvalidOperationException)
            {
                streamObjectHeader = null;
                return 0;
            }
        }
    }

    /// <summary>
    /// An 16-bit header for a compound object would indicate the start of a stream object 
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class StreamObjectHeaderStart16bit : StreamObjectHeaderStart
    {
        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderStart16bit class with specified type and length.
        /// </summary>
        /// <param name="type">Specify the type of the StreamObjectHeaderStart16bit.</param>
        /// <param name="length">Specify the length of the StreamObjectHeaderStart16bit.</param>
        public StreamObjectHeaderStart16bit(StreamObjectTypeHeaderStart type, int length)
        {
            if (this.Length > 127)
            {
                throw new ArgumentOutOfRangeException("Length", "16-bit Stream Object Header Start, Length (7-bits): A 7-bit unsigned integer that specifies the length in bytes for additional data (if any). If the length is more than 127 bytes, a 32-bit stream object header start MUST be used.");
            }

            this.HeaderType = 0x0;
            this.Type = type;
            this.Compound = StreamObject.CompoundTypes.Contains(this.Type) ? 1 : 0;
            this.Length = length;
        }

        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderStart16bit class with specified type.
        /// </summary>
        /// <param name="type">Specify the type of the StreamObjectHeaderStart16bit.</param>
        public StreamObjectHeaderStart16bit(StreamObjectTypeHeaderStart type)
            : this(type, 0)
        {
        }

        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderStart16bit class, this is the default constructor.
        /// </summary>
        public StreamObjectHeaderStart16bit()
        {
        }

        /// <summary>
        /// This method is used to convert the element of StreamObjectHeaderStart16bit basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of StreamObjectHeaderStart16bit.</returns>
        public override List<byte> SerializeToByteList()
        {
            BitWriter bitField = new BitWriter(2);
            bitField.AppendInit32(this.HeaderType, 2);
            bitField.AppendInit32(this.Compound, 1);
            bitField.AppendUInit32(Convert.ToUInt32(this.Type), 6);
            bitField.AppendInit32(this.Length, 7);
            return new List<byte>(bitField.Bytes);
        }

        /// <summary>
        /// This method is used to get the Uint16 value of the 16bit stream object header.
        /// </summary>
        /// <returns>Return the ushort value.</returns>
        public ushort ToUint16()
        {
            List<byte> bytes = this.SerializeToByteList();
            return LittleEndianBitConverter.ToUInt16(bytes.ToArray(), 0);
        }

        /// <summary>
        /// This method is used to deserialize the StreamObjectHeaderStart16bit basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the StreamObjectHeaderStart16bit basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            using (BitReader bitReader = new BitReader(byteArray, startIndex))
            {
                this.HeaderType = bitReader.ReadInt32(2);
                if (this.HeaderType != StreamObjectHeaderStart.StreamObjectHeaderStart16bit)
                {
                    throw new InvalidOperationException(string.Format("Failed to get the StreamObjectHeaderStart16bit header type value, expect value {0}, but actual value is {1}", StreamObjectHeaderStart16bit, this.HeaderType));
                }

                this.Compound = bitReader.ReadInt32(1);
                int typeValue = bitReader.ReadInt32(6);
                if (!Enum.IsDefined(typeof(StreamObjectTypeHeaderStart), typeValue))
                {
                    throw new InvalidOperationException(string.Format("Failed to get the StreamObjectHeaderStart16bit type value, the value {0} is not defined", typeValue));
                }

                this.Type = (StreamObjectTypeHeaderStart)typeValue;
                if (StreamObject.CompoundTypes.Contains(this.Type) && this.Compound != 1)
                {
                    throw new InvalidOperationException(string.Format("Failed to parse the StreamObjectHeaderStart16bit header. If the type value is {0} then the compound value should 1, but actual value is 0", typeValue));
                }

                this.Length = bitReader.ReadInt32(7);
                if (this.Length > 127)
                {
                    throw new InvalidOperationException("16-bit Stream Object Header Start, Length (7-bits): A 7-bit unsigned integer that specifies the length in bytes for additional data (if any). If the length is more than 127 bytes, a 32-bit stream object header start MUST be used.");
                }

                return 2;
            }
        }
    }

    /// <summary>
    /// An 32-bit header for a compound object would indicate the start of a stream object 
    /// </summary>
    public class StreamObjectHeaderStart32bit : StreamObjectHeaderStart
    {
        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderStart32bit class with specified type and length.
        /// </summary>
        /// <param name="type">Specify the type of the StreamObjectHeaderStart32bit.</param>
        /// <param name="length">Specify the length of the StreamObjectHeaderStart32bit.</param>
        public StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart type, int length)
        {
            this.HeaderType = StreamObjectHeaderStart.StreamObjectHeaderStart32bit;
            this.Type = type;
            this.Compound = StreamObject.CompoundTypes.Contains(this.Type) ? 1 : 0;

            if (length >= 32767)
            {
                this.Length = 32767;
                this.LargeLength = new Compact64bitInt((ulong)length);
            }
            else
            {
                this.Length = length;
                this.LargeLength = null;
            }
        }

        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderStart32bit class, this is the default constructor.
        /// </summary>
        public StreamObjectHeaderStart32bit()
        {
        }

        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderStart32bit class with specified type.
        /// </summary>
        /// <param name="streamObjectTypeHeaderStart">Specify the type of the StreamObjectHeaderStart32bit.</param>
        public StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart streamObjectTypeHeaderStart)
        {
            this.Type = streamObjectTypeHeaderStart;
        }

        /// <summary>
        /// Gets or sets an optional compact uint64 that specifies the length in bytes for additional data (if any). 
        /// This field MUST be specified if the Length field contains 32767 and MUST NOT be specified if the Length field 
        /// contains any other value than 32767.
        /// </summary>
        public Compact64bitInt LargeLength { get; set; }

        /// <summary>
        /// This method is used to convert the element of StreamObjectHeaderStart32bit basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of StreamObjectHeaderStart32bit.</returns>
        public override List<byte> SerializeToByteList()
        {
            BitWriter bitFieldWriter = new BitWriter(4);
            bitFieldWriter.AppendInit32(this.HeaderType, 2);
            bitFieldWriter.AppendInit32(this.Compound, 1);
            bitFieldWriter.AppendUInit32(Convert.ToUInt32(this.Type), 14);
            bitFieldWriter.AppendInit32(this.Length, 15);

            List<byte> listByte = new List<byte>(bitFieldWriter.Bytes);

            if (this.LargeLength != null)
            {
                listByte.AddRange(this.LargeLength.SerializeToByteList().ToArray());
            }

            return listByte;
        }

        /// <summary>
        /// This method is used to deserialize the StreamObjectHeaderStart32bit basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the StreamObjectHeaderStart32bit basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            using (BitReader bitReader = new BitReader(byteArray, startIndex))
            {
                this.HeaderType = bitReader.ReadInt32(2);
                if (this.HeaderType != StreamObjectHeaderStart.StreamObjectHeaderStart32bit)
                {
                    throw new InvalidOperationException(string.Format("Failed to get the StreamObjectHeaderStart32bit header type value, expect value {0}, but actual value is {1}", StreamObjectHeaderStart.StreamObjectHeaderStart32bit, this.HeaderType));
                }

                this.Compound = bitReader.ReadInt32(1);
                int typeValue = bitReader.ReadInt32(14);
                if (!Enum.IsDefined(typeof(StreamObjectTypeHeaderStart), typeValue))
                {
                    throw new InvalidOperationException(string.Format("Failed to get the StreamObjectHeaderStart32bit type value, the value {0} is not defined", typeValue));
                }

                this.Type = (StreamObjectTypeHeaderStart)typeValue;
                if (StreamObject.CompoundTypes.Contains(this.Type) && this.Compound != 1)
                {
                    throw new InvalidOperationException(string.Format("Failed to parse the StreamObjectHeaderStart32bit header. If the type value is {0} then the compound value should 1, but actual value is 0", typeValue));
                }

                this.Length = bitReader.ReadInt32(15);

                int index = startIndex;
                index += 4;

                if (this.Length == 32767)
                {
                    this.LargeLength = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
                }

                return index - startIndex;
            }
        }
    }

    /// <summary>
    /// This class specifies the base class for 8-bit or 16-bit stream object header end.
    /// </summary>
    public abstract class StreamObjectHeaderEnd : BasicObject
    {
        /// <summary>
        /// Gets or sets the type of the stream object.
        /// value 1 for 8-bit stream object header start, 
        /// value 3 for 16-bit stream object header start.
        /// </summary>
        public StreamObjectTypeHeaderEnd Type { get; set; }
    }

    /// <summary>
    /// An 8-bit header for a compound object would indicate the end of a stream object 
    /// </summary>
    public class StreamObjectHeaderEnd8bit : StreamObjectHeaderEnd
    {
        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderEnd8bit class with the specified type value.
        /// </summary>
        /// <param name="type">Specify the integer value of the type.</param>
        public StreamObjectHeaderEnd8bit(int type)
        {
            if (!Enum.IsDefined(typeof(StreamObjectTypeHeaderEnd), type))
            {
                throw new InvalidOperationException(string.Format("The type value {0} is not defined for the stream object end 8 bit header", type));
            }

            this.Type = (StreamObjectTypeHeaderEnd)type;
        }

        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderEnd8bit class, this is the default constructor.
        /// </summary>
        public StreamObjectHeaderEnd8bit()
        {
        }

        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderEnd8bit class with the specified type value.
        /// </summary>
        /// <param name="type">Specify the value of the type.</param>
        public StreamObjectHeaderEnd8bit(StreamObjectTypeHeaderEnd type)
            : this((int)type)
        {
        }

        /// <summary>
        /// This method is used to convert the element of StreamObjectHeaderEnd8bit basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of StreamObjectHeaderEnd8bit.</returns>
        public override List<byte> SerializeToByteList()
        {
            BitWriter bitFieldWriter = new BitWriter(1);
            bitFieldWriter.AppendInit32(0x1, 2);
            bitFieldWriter.AppendUInit32(Convert.ToUInt32(this.Type), 6);
            return new List<byte>(bitFieldWriter.Bytes);
        }

        /// <summary>
        /// This method is used to get the byte value of the 8bit stream object header End.
        /// </summary>
        /// <returns>Return StreamObjectHeaderEnd8bit value represented by byte.</returns>
        public byte ToByte()
        {
            List<byte> bytes = this.SerializeToByteList();

            if (bytes.Count != 1)
            {
                throw new InvalidOperationException("The unexpected StreamObjectHeaderEnd8bit length");
            }

            return bytes[0];
        }

        /// <summary>
        /// This method is used to deserialize the StreamObjectHeaderEnd8bit basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the StreamObjectHeaderEnd8bit basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            using (BitReader reader = new BitReader(byteArray, startIndex))
            {
                int headerType = reader.ReadInt32(2);

                if (headerType != 0x1)
                {
                    throw new InvalidOperationException(string.Format("Failed to get the StreamObjectHeaderEnd8bit header type value, expect value {0}, but actual value is {1}", 0x1, headerType));
                }

                uint typeValue = reader.ReadUInt32(6);
                if (!Enum.IsDefined(typeof(StreamObjectTypeHeaderEnd), (int)typeValue))
                {
                    throw new InvalidOperationException(string.Format("Failed to get the StreamObjectHeaderEnd8bit type value, the value {0} is not defined", typeValue));
                }

                this.Type = (StreamObjectTypeHeaderEnd)typeValue;
                return 1;
            }
        }
    }

    /// <summary>
    /// An 16-bit header for a compound object would indicate the end of a stream object 
    /// </summary>
    public class StreamObjectHeaderEnd16bit : StreamObjectHeaderEnd
    {
        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderEnd16bit class with the specified type value.
        /// </summary>
        /// <param name="type">Specify the integer value of the type.</param>
        public StreamObjectHeaderEnd16bit(int type)
        {
            if (!Enum.IsDefined(typeof(StreamObjectTypeHeaderEnd), type))
            {
                throw new InvalidOperationException(string.Format("The type value {0} is not defined for the stream object end 16-bit header", type));
            }

            this.Type = (StreamObjectTypeHeaderEnd)type;
        }

        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderEnd16bit class with the specified type value.
        /// </summary>
        /// <param name="headerType">Specify the value of the type.</param>
        public StreamObjectHeaderEnd16bit(StreamObjectTypeHeaderEnd headerType)
            : this((int)headerType)
        {
        }

        /// <summary>
        /// Initializes a new instance of the StreamObjectHeaderEnd16bit class, this is the default constructor.
        /// </summary>
        public StreamObjectHeaderEnd16bit()
        {
        }

        /// <summary>
        /// This method is used to convert the element of StreamObjectHeaderEnd16bit basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of StreamObjectHeaderEnd16bit.</returns>
        public override List<byte> SerializeToByteList()
        {
            BitWriter bitFieldWriter = new BitWriter(2);
            bitFieldWriter.AppendInit32(0x3, 2);
            bitFieldWriter.AppendUInit32(Convert.ToUInt32(this.Type), 14);
            return new List<byte>(bitFieldWriter.Bytes);
        }

        /// <summary>
        /// This method is used to get the byte value of the 16-bit stream object header End.
        /// </summary>
        /// <returns>Return StreamObjectHeaderEnd8bit value represented by unsigned short integer.</returns>
        public ushort ToUint16()
        {
            List<byte> bytes = this.SerializeToByteList();
            return LittleEndianBitConverter.ToUInt16(bytes.ToArray(), 0);
        }

        /// <summary>
        /// This method is used to deserialize the StreamObjectHeaderEnd16bit basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the StreamObjectHeaderEnd16bit basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            using (BitReader reader = new BitReader(byteArray, startIndex))
            {
                int headerType = reader.ReadInt32(2);

                if (headerType != 0x3)
                {
                    throw new InvalidOperationException(string.Format("Failed to get the StreamObjectHeaderEnd16bit header type value, expect value {0}, but actual value is {1}", 0x3, headerType));
                }

                uint typeValue = reader.ReadUInt32(14);
                if (!Enum.IsDefined(typeof(StreamObjectTypeHeaderEnd), (int)typeValue))
                {
                    throw new InvalidOperationException(string.Format("Failed to get the StreamObjectHeaderEnd16bit type value, the value {0} is not defined", typeValue));
                }

                this.Type = (StreamObjectTypeHeaderEnd)typeValue;
                return 2;
            }
        }
    }
}