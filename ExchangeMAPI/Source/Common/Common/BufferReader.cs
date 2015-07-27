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
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Used to read buffer.
    /// </summary>
    public class BufferReader
    {
        /// <summary>
        /// Length of buffer.
        /// </summary>
        private uint length;

        /// <summary>
        /// Point to current position of buffer reader.
        /// </summary>
        private uint position;

        /// <summary>
        /// The buffer being read.
        /// </summary>
        private byte[] buffer;

        /// <summary>
        /// Initializes a new instance of the BufferReader class.
        /// </summary>
        /// <param name="buffer">The buffer being read.</param>
        public BufferReader(byte[] buffer)
        {
            if (buffer == null)
            {
                throw new Exception("buffer can't be null");
            }

            this.buffer = buffer;
            this.length = (uint)buffer.Length;
            this.position = 0;
        }

        /// <summary>
        /// Gets length of buffer.
        /// </summary>
        public uint Length
        {
            get { return this.length; }
        }

        /// <summary>
        /// Gets current position to be read.
        /// </summary>
        public uint Position
        {
            get { return this.position; }
        }

        /// <summary>
        /// Reads a 2-byte unsigned integer from the current position and advances the current position of the buffer by 2 bytes.
        /// </summary>
        /// <returns>A 2-byte unsigned integer read from the current buffer.</returns>
        public ushort ReadUInt16()
        {
            if (this.position + 1 >= this.length)
            {
                throw new Exception("bufferReader out of bound.");
            }

            short low = (short)this.buffer[this.position++];
            short high = (short)(this.buffer[this.position++] << 8);
            return (ushort)(low + high);
        }

        /// <summary>
        /// Reads a 2-byte signed integer from the current position and advances the current position of the buffer by 2 bytes.
        /// </summary>
        /// <returns>A 2-byte signed integer read from the current buffer.</returns>
        public short ReadInt16()
        {
            if (this.position + 1 >= this.length)
            {
                throw new Exception("bufferReader out of bound.");
            }

            short low = (short)this.buffer[this.position++];
            short high = (short)(this.buffer[this.position++] << 8);
            return (short)(low + high);
        }

        /// <summary>
        /// Reads a 4-byte unsigned integer from the current position and advances the current position of the buffer by 4 bytes.
        /// </summary>
        /// <returns>A 4-byte unsigned integer read from the current buffer.</returns>
        public uint ReadUInt32()
        {
            if (this.position + 3 >= this.length)
            {
                throw new Exception("bufferReader out of bound.");
            }

            int part1 = (int)this.buffer[this.position++];
            int part2 = (int)(this.buffer[this.position++] << 8);
            int part3 = (int)(this.buffer[this.position++] << 16);
            int part4 = (int)(this.buffer[this.position++] << 24);
            return (uint)(part1 + part2 + part3 + part4);
        }

        /// <summary>
        /// Reads a 4-byte signed integer from the current position and advances the current position of the buffer by 4 bytes.
        /// </summary>
        /// <returns>A 4-byte signed integer read from the current buffer.</returns>
        public int ReadInt32()
        {
            if (this.position + 3 >= this.length)
            {
                throw new Exception("bufferReader out of bound.");
            }

            int part1 = (int)this.buffer[this.position++];
            int part2 = (int)(this.buffer[this.position++] << 8);
            int part3 = (int)(this.buffer[this.position++] << 16);
            int part4 = (int)(this.buffer[this.position++] << 24);
            return (int)(part1 + part2 + part3 + part4);
        }

        /// <summary>
        /// Reads an 8-byte unsigned integer from the current position and advances the current position of the buffer by 8 bytes.
        /// </summary>
        /// <returns>An 8-byte unsigned integer read from the current buffer.</returns>
        public ulong ReadUInt64()
        {
            long low = (uint)this.ReadInt32();
            long high = (uint)this.ReadInt32();

            // 0x100000000 is 2 raised to the 32th power plus 1
            return (ulong)((high * 0x100000000) + low);
        }

        /// <summary>
        /// Reads an 8-byte signed integer from the current position and advances the current position of the buffer by 8 bytes.
        /// </summary>
        /// <returns>An 8-byte signed integer read from the current buffer.</returns>
        public long ReadInt64()
        {
            long low = this.ReadInt32();
            long high = this.ReadInt32();

            // 0x100000000 is 2 raised to the 32th power plus 1
            return (long)((high * 0x100000000) + low);
        }

        /// <summary>
        /// Reads a string into ASCII format from the current position and advances the current position of the buffer by string length bytes.
        /// </summary>
        /// <returns>An ASCII string from the current buffer.</returns>
        public string ReadASCIIString()
        {
            List<byte> byteList = new List<byte>();
            for (; this.buffer[this.position] != 0x00; this.position++)
            {
                if (this.position >= this.length)
                {
                    throw new Exception("bufferReader out of bound.");
                }

                byteList.Add(this.buffer[this.position]);
            }

            byteList.Add(0);
            this.position++;
            return Encoding.ASCII.GetString(byteList.ToArray());
        }

        /// <summary>
        /// Reads a string into Unicode format from the current position and advances the current position of the buffer by 2 * string length bytes.
        /// </summary>
        /// <returns>A Unicode string from the current buffer.</returns>
        public string ReadUnicodeString()
        {
            List<byte> byteList = new List<byte>();
            for (; this.buffer[this.position] != 0x00 || this.buffer[this.position + 1] != 0x00; this.position += 2)
            {
                if (this.position + 1 >= this.length)
                {
                    throw new Exception("bufferReader out of bound.");
                }

                byteList.Add(this.buffer[this.position]);
                byteList.Add(this.buffer[this.position + 1]);
            }

            byteList.Add(0);
            byteList.Add(0);
            this.position += 2;
            return Encoding.Unicode.GetString(byteList.ToArray());
        }

        /// <summary>
        /// Reads the specified number of bytes from the current position into a byte array and advances the current position by that number of bytes.
        /// </summary>
        /// <param name="length">The number of bytes to read.</param>
        /// <returns>A byte array containing data read from buffer. .</returns>
        public byte[] ReadBytes(uint length)
        {
            if (length <= 0)
            {
                return null;
            }

            if (this.position + length > this.length)
            {
                throw new Exception("bufferReader out of bound.");
            }

            byte[] bytes = new byte[length];
            for (int i = 0; i < length; i++)
            {
                bytes[i] = this.buffer[this.position++];
            }

            return bytes;
        }

        /// <summary>
        /// Reads the specified number of bytes from the position of index into a byte array and advances the current position by that number of bytes.
        /// </summary>
        /// <param name="index">The position of bytes to read.</param>
        /// <param name="length">The number of bytes to read.</param>
        /// <returns>A byte array containing data read from buffer. .</returns>
        public byte[] ReadBytes(uint index, uint length)
        {
            if (index + length > this.length)
            {
                throw new Exception("bufferReader out of bound.");
            }

            this.position = index;
            return this.ReadBytes(length);
        }

        /// <summary>
        /// Reads a byte from the current position and advances the current position of the buffer by one byte.
        /// </summary>
        /// <returns>The next byte read from the buffer.</returns>
        public byte ReadByte()
        {
            if (this.position >= this.length)
            {
                throw new Exception("bufferReader out of bound.");
            }

            return this.buffer[this.position++];
        }

        /// <summary>
        /// Read bytes from current position to the end of buffer.
        /// </summary>
        /// <returns>A byte array containing data read from buffer.</returns>
        public byte[] ReadToEnd()
        {
            byte[] byteArray = this.ReadBytes(this.length - this.position);
            return byteArray;
        }
    }
}