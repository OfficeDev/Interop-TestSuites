namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections;
    using System.Collections.Generic;

    /// <summary>
    /// The class is used to read/set bit value for a byte array.
    /// </summary>
    public static class Bit
    {
        /// <summary>
        /// Read a bit value from a byte array with the specified bit position.
        /// </summary>
        /// <param name="array">Specify the byte array.</param>
        /// <param name="bit">Specify the bit position.</param>
        /// <returns>Return the bit value in the specified bit position.</returns>
        public static bool IsBitSet(byte[] array, long bit)
        {
            return (array[bit / 8] & (1 << (int)(bit % 8))) != 0;
        }

        /// <summary>
        /// Set a bit value to "On" in the specified byte array with the specified bit position.
        /// </summary>
        /// <param name="array">Specify the byte array.</param>
        /// <param name="bit">Specify the bit position.</param>
        public static void SetBit(byte[] array, long bit)
        {
            array[bit / 8] |= unchecked((byte)(1 << (int)(bit % 8)));
        }

        /// <summary>
        /// Set a bit value to "Off" in the specified byte array with the specified bit position.
        /// </summary>
        /// <param name="array">Specify the byte array.</param>
        /// <param name="bit">Specify the bit position.</param>
        public static void ClearBit(byte[] array, long bit)
        {
            array[bit / 8] &= unchecked((byte)(~(1 << (int)(bit % 8))));
        }
    }

    /// <summary>
    /// A class is used to extract values across byte boundaries with arbitrary bit positions.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public sealed class BitReader : IEnumerator<bool>
    {
        /// <summary>
        /// A byte array which contains the bytes need to be read.
        /// </summary>
        private byte[] byteArray;

        /// <summary>
        /// A start position which will be not changed in the process of reading.
        /// This value will be used for recording the start position and will be used by the function reset.
        /// </summary>
        private long startPosition;

        /// <summary>
        /// An offset which is used to keep trace for the current read position in bit.
        /// </summary>
        private long offset;

        /// <summary>
        /// The length of the byte Array which contains the byte need to be read.
        /// </summary>
        private long length;

        /// <summary>
        /// Initializes a new instance of the BitReader class with specified bytes buffer and start position in byte.
        /// </summary>
        /// <param name="array">Specify the byte array which contains the bytes need to be read.</param>
        /// <param name="index">Specify the start position in byte.</param>
        public BitReader(byte[] array, int index)
        {
            this.byteArray = array;
            this.offset = ((long)index * 8) - 1;
            this.startPosition = this.offset;
            this.length = (long)array.Length * 8;
        }

        /// <summary>
        /// Gets a value indicating whether the bit is true or false in the current bit position.
        /// </summary>
        public bool Current
        {
            get
            {
                return Bit.IsBitSet(this.byteArray, this.offset);
            }
        }

        /// <summary>
        /// Gets the bit in the byte array at the current position of the enumerator.
        /// </summary>
        object IEnumerator.Current
        {
            get { return this.Current; }
        }

        /// <summary>
        /// Read specified bit length content as an UInt64 type and increase the bit offset. 
        /// </summary>
        /// <param name="readingLength">Specify the reading bit length.</param>
        /// <returns>Return the UInt64 type value.</returns>
        public ulong ReadUInt64(int readingLength)
        {
            byte[] uint64Bytes = this.GetBytes(readingLength, 8);
            return LittleEndianBitConverter.ToUInt64(uint64Bytes, 0);
        }

        /// <summary>
        /// Read specified bit length content as an UInt32 type and increase the bit offset with the specified length. 
        /// </summary>
        /// <param name="readingLength">Specify the reading bit length.</param>
        /// <returns>Return the UInt32 type value.</returns>
        public uint ReadUInt32(int readingLength)
        {
            byte[] uint32Bytes = this.GetBytes(readingLength, 4);
            return LittleEndianBitConverter.ToUInt32(uint32Bytes, 0);
        }

        /// <summary>
        /// Reading the bytes specified by the byte length.
        /// </summary>
        /// <param name="readingLength">Specify the reading byte length.</param>
        /// <returns>Return the read bytes array.</returns>
        public byte[] ReadBytes(int readingLength)
        {
            byte[] readingByteArray = this.GetBytes(readingLength * 8, readingLength);
            return readingByteArray;
        }

        /// <summary>
        /// Read specified bit length content as an byte type and increase the bit offset with the specified length. 
        /// </summary>
        /// <param name="readingBitLength">Specify the reading bit length.</param>
        /// <returns>Return the byte value.</returns>
        public byte ReadByte(int readingBitLength)
        {
            byte[] readingByteArray = this.GetBytes(readingBitLength, 1);
            return readingByteArray[0];
        }

        /// <summary>
        /// Read specified bit length content as an UInt16 type and increase the bit offset with the specified length. 
        /// </summary>
        /// <param name="readingLength">Specify the reading bit length.</param>
        /// <returns>Return the UInt16 value.</returns>
        public short ReadInt16(int readingLength)
        {
            byte[] uint16Bytes = this.GetBytes(readingLength, 2);
            return LittleEndianBitConverter.ToInt16(uint16Bytes, 0);
        }

        /// <summary>
        /// Read specified bit length content as an Int32 type and increase the bit offset with the specified length. 
        /// </summary>
        /// <param name="readingLength">Specify the reading bit length.</param>
        /// <returns>Return the Int32 type value.</returns>
        public int ReadInt32(int readingLength)
        {
            byte[] uint32Bytes = this.GetBytes(readingLength, 4);
            return LittleEndianBitConverter.ToInt32(uint32Bytes, 0);
        }

        /// <summary>
        /// Read as a GUID from the current offset position and increate the bit offset with 128 bit.
        /// </summary>
        /// <returns>Return the GUID value.</returns>
        public Guid ReadGuid()
        {
            return new Guid(this.GetBytes(128, 16));
        }

        /// <summary>
        /// Advances the enumerator to the next bit of the byte array.
        /// </summary>
        /// <returns>true if the enumerator was successfully advanced to the next bit; false if the enumerator has passed the end of the byte array.</returns>
        public bool MoveNext()
        {
            return ++this.offset < this.length;
        }

        /// <summary>
        /// Assign the internal read buffer to null.
        /// </summary>
        public void Dispose()
        {
            this.byteArray = null;
        }

        /// <summary>
        /// Sets the enumerator to its initial position, which is before the first bit in the byte array.
        /// </summary>
        public void Reset()
        {
            this.offset = this.startPosition;
        }

        /// <summary>
        /// Construct a byte array with specified bit length and the specified the byte array size.
        /// </summary>
        /// <param name="needReadlength">Specify the need read bit length.</param>
        /// <param name="size">Specify the byte array size.</param>
        /// <returns>Returns the constructed byte array.</returns>
        private byte[] GetBytes(int needReadlength, int size)
        {
            byte[] retBytes = new byte[size];
            int i = 0;
            while (i < needReadlength)
            {
                if (!this.MoveNext())
                {
                    throw new InvalidOperationException("Unexpected to meet the byte array end.");
                }

                if (this.Current)
                {
                    Bit.SetBit(retBytes, i);
                }
                else
                {
                    Bit.ClearBit(retBytes, i);
                }

                i++;
            }

            return retBytes;
        }
    }

    /// <summary>
    /// A class is used to write various primitive number into a byte array with arbitrary bit positions.
    /// </summary>
    public sealed class BitWriter
    {
        /// <summary>
        /// A byte buffer will contain all the written byte.
        /// </summary>
        private byte[] bytes;

        /// <summary>
        /// An offset which is used to keep trace for the current write position in bit, staring with 0.
        /// </summary>
        private int bitOffset;

        /// <summary>
        /// Initializes a new instance of the BitWriter class with specified buffer size in byte.
        /// </summary>
        /// <param name="bufferSize">Specify the buffer byte size.</param>
        public BitWriter(int bufferSize)
        {
            this.bytes = new byte[bufferSize];
            this.bitOffset = 0;
        }

        /// <summary>
        /// Gets a copy byte array which contains the current written byte.
        /// </summary>
        public byte[] Bytes
        {
            get
            {
                if (this.bitOffset % 8 != 0)
                {
                    throw new InvalidOperationException("BitWriter:Bytes, Cannot get the current bytes because the last byte is not written completely.");
                }

                int retByteLength = this.bitOffset / 8;
                byte[] retByteArray = new byte[retByteLength];
                System.Array.Copy(this.bytes, 0, retByteArray, 0, retByteLength);
                return retByteArray;
            }
        }

        /// <summary>
        /// Append a specified Unit64 type value into the buffer with the specified bit length.
        /// </summary>
        /// <param name="value">Specify the value which needs to be appended.</param>
        /// <param name="length">Specify the bit length which the value will occupy in the buffer.</param>
        public void AppendUInt64(ulong value, int length)
        {
            byte[] convertedBytes = LittleEndianBitConverter.GetBytes(value);
            this.SetBytes(convertedBytes, length);
        }

        /// <summary>
        /// Append a specified Unit32 type value into the buffer with the specified bit length.
        /// </summary>
        /// <param name="value">Specify the value which needs to be appended.</param>
        /// <param name="length">Specify the bit length which the value will occupy in the buffer.</param>
        public void AppendUInit32(uint value, int length)
        {
            byte[] convertedBytes = LittleEndianBitConverter.GetBytes(value);
            this.SetBytes(convertedBytes, length);
        }

        /// <summary>
        /// Append a specified Init32 type value into the buffer with the specified bit length.
        /// </summary>
        /// <param name="value">Specify the value which needs to be appended.</param>
        /// <param name="length">Specify the bit length which the value will occupy in the buffer.</param>
        public void AppendInit32(int value, int length)
        {
            byte[] convertedBytes = LittleEndianBitConverter.GetBytes(value);
            this.SetBytes(convertedBytes, length);
        }

        /// <summary>
        /// Append a specified GUID value into the buffer.
        /// </summary>
        /// <param name="value">Specify the GUID value.</param>
        public void AppendGUID(Guid value)
        {
            this.SetBytes(value.ToByteArray(), 128);
        }

        /// <summary>
        /// Write the specified byte array into the buffer from the current position with the specified bit length.
        /// </summary>
        /// <param name="needWrittenBytes">Specify the needed written byte array.</param>
        /// <param name="length">Specify the bit length which the byte array will occupy in the buffer.</param>
        private void SetBytes(byte[] needWrittenBytes, int length)
        {
            for (uint i = 0; i < length; i++)
            {
                if (Bit.IsBitSet(needWrittenBytes, i))
                {
                    Bit.SetBit(this.bytes, this.bitOffset++);
                }
                else
                {
                    Bit.ClearBit(this.bytes, this.bitOffset++);
                }
            }
        }
    }
}