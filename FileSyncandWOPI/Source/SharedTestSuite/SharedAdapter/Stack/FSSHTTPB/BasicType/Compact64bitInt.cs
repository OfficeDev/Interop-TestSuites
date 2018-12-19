namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using Microsoft.Protocols.TestSuites.Common;
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// A 9-byte encoding of values in the range 0x0002000000000000 through 0xFFFFFFFFFFFFFFFF
    /// </summary>
    public class Compact64bitInt : BasicObject
    {
        /// <summary>
        /// Specify the type value for compact uint zero type value.
        /// </summary>
        public const int CompactUintNullType = 0;
        
        /// <summary>
        /// Specify the type value for compact uint 7 bits type value.
        /// </summary>
        public const int CompactUint7bitType = 1;
        
        /// <summary>
        /// Specify the type value for compact uint 14 bits type value.
        /// </summary>
        public const int CompactUint14bitType = 2;
        
        /// <summary>
        /// Specify the type value for compact uint 21 bits type value.
        /// </summary>
        public const int CompactUint21bitType = 4;
        
        /// <summary>
        /// Specify the type value for compact uint 28 bits type value.
        /// </summary>
        public const int CompactUint28bitType = 8;
        
        /// <summary>
        /// Specify the type value for compact uint 35 bits type value.
        /// </summary>
        public const int CompactUint35bitType = 16;
        
        /// <summary>
        /// Specify the type value for compact uint 42 bits type value.
        /// </summary>
        public const int CompactUint42bitType = 32;
        
        /// <summary>
        /// Specify the type value for compact uint 49 bits type value.
        /// </summary>
        public const int CompactUint49bitType = 64;
        
        /// <summary>
        /// Specify the type value for compact uint 64 bits type value.
        /// </summary>
        public const int CompactUint64bitType = 128;

        /// <summary>
        /// Initializes a new instance of the Compact64bitInt class with specified value.
        /// </summary>
        /// <param name="decodedValue">Decoded value</param>
        public Compact64bitInt(ulong decodedValue)
        {
            this.DecodedValue = decodedValue;
        }

        /// <summary>
        /// Initializes a new instance of the Compact64bitInt class, this is the default constructor.
        /// </summary>
        public Compact64bitInt()
        {
            this.DecodedValue = 0;
        }

        /// <summary>
        /// Gets or sets the Type value.
        /// </summary>
        public uint Type { get; set; }

        /// <summary>
        /// Gets or sets the value represented by the compact uint value.
        /// </summary>
        public ulong DecodedValue { get; set; }

        /// <summary>
        /// This method is used to convert the element of Compact64bitInt basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of Compact64bitInt.</returns>
        public override List<byte> SerializeToByteList()
        {
            BitWriter bitWriter = new BitWriter(9);

            if (this.DecodedValue == 0)
            {
                bitWriter.AppendUInt64(0, 8);
            }
            else if (this.DecodedValue >= 0x01 && this.DecodedValue <= 0x7F)
            {
                bitWriter.AppendUInt64(CompactUint7bitType, 1);
                bitWriter.AppendUInt64(this.DecodedValue, 7);
            }
            else if (this.DecodedValue >= 0x0080 && this.DecodedValue <= 0x3FFF)
            {
                bitWriter.AppendUInt64(CompactUint14bitType, 2);
                bitWriter.AppendUInt64(this.DecodedValue, 14);
            }
            else if (this.DecodedValue >= 0x004000 && this.DecodedValue <= 0x1FFFFF)
            {
                bitWriter.AppendUInt64(CompactUint21bitType, 3);
                bitWriter.AppendUInt64(this.DecodedValue, 21);
            }
            else if (this.DecodedValue >= 0x0200000 && this.DecodedValue <= 0xFFFFFFF)
            {
                bitWriter.AppendUInt64(CompactUint28bitType, 4);
                bitWriter.AppendUInt64(this.DecodedValue, 28);
            }
            else if (this.DecodedValue >= 0x010000000 && this.DecodedValue <= 0x7FFFFFFFF)
            {
                bitWriter.AppendUInt64(CompactUint35bitType, 5);
                bitWriter.AppendUInt64(this.DecodedValue, 35);
            }
            else if (this.DecodedValue >= 0x00800000000 && this.DecodedValue <= 0x3FFFFFFFFFF)
            {
                bitWriter.AppendUInt64(CompactUint42bitType, 6);
                bitWriter.AppendUInt64(this.DecodedValue, 42);
            }
            else if (this.DecodedValue >= 0x0040000000000 && this.DecodedValue <= 0x1FFFFFFFFFFFF)
            {
                bitWriter.AppendUInt64(CompactUint49bitType, 7);
                bitWriter.AppendUInt64(this.DecodedValue, 49);
            }
            else if (this.DecodedValue >= 0x0002000000000000 && this.DecodedValue <= 0xFFFFFFFFFFFFFFFF)
            {
                bitWriter.AppendUInt64(CompactUint64bitType, 8);
                bitWriter.AppendUInt64(this.DecodedValue, 64);
            }

            return new List<byte>(bitWriter.Bytes);
        }

        /// <summary>
        /// This method is used to deserialize the Compact64bitInt basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the Compact64bitInt basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex) // return the length consumed
        {
            using (BitReader bitReader = new BitReader(byteArray, startIndex))
            {
                int numberOfContinousZeroBit = 0;
                while (numberOfContinousZeroBit < 8 && bitReader.MoveNext())
                {
                    if (bitReader.Current == false)
                    {
                        numberOfContinousZeroBit++;
                    }
                    else
                    {
                        break;
                    }
                }

                switch (numberOfContinousZeroBit)
                {
                    case 0:
                        this.DecodedValue = bitReader.ReadUInt64(7);
                        this.Type = CompactUint7bitType;
                        return 1;

                    case 1:
                        this.DecodedValue = bitReader.ReadUInt64(14);
                        this.Type = CompactUint14bitType;
                        return 2;

                    case 2:
                        this.DecodedValue = bitReader.ReadUInt64(21);
                        this.Type = CompactUint21bitType;
                        return 3;

                    case 3:
                        this.DecodedValue = bitReader.ReadUInt64(28);
                        this.Type = CompactUint28bitType;
                        return 4;

                    case 4:
                        this.DecodedValue = bitReader.ReadUInt64(35);
                        this.Type = CompactUint35bitType;
                        return 5;

                    case 5:
                        this.DecodedValue = bitReader.ReadUInt64(42);
                        this.Type = CompactUint42bitType;
                        return 6;

                    case 6:
                        this.DecodedValue = bitReader.ReadUInt64(49);
                        this.Type = CompactUint49bitType;
                        return 7;

                    case 7:
                        this.DecodedValue = bitReader.ReadUInt64(64);
                        this.Type = CompactUint64bitType;
                        return 9;

                    case 8:
                        this.DecodedValue = 0;
                        this.Type = CompactUintNullType;
                        return 1;

                    default:
                        throw new InvalidOperationException("Failed to parse the Compact64bitInt, the type value is unexpected");
                }
            }
        }
    }
}