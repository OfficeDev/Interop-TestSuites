//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Used for serializing or deserializing a fast transfer stream.
    /// </summary>
    public class FastTransferStream : MemoryStream
    {
        /// <summary>
        /// The length of a GUID structure.
        /// </summary>
        public static readonly int GuidLength = Guid.Empty.ToByteArray().Length;

        /// <summary>
        /// The length of a PidTag property.
        /// </summary>
        private const int PidLength = MarkersHelper.PidTagLength;

        /// <summary>
        /// Initializes a new instance of the FastTransferStream class.
        /// </summary>
        /// <param name="buffer">A bytes array.</param>
        /// <param name="writable">Whether the stream supports writing.</param>
        public FastTransferStream(byte[] buffer, bool writable)
            : base(buffer, 0, buffer.Length, writable, true)
        {
        }

        /// <summary>
        /// Gets a value indicating whether the stream position is at the end of this stream
        /// </summary>
        public bool IsEndOfStream
        {
            get
            {
                return this.Position == this.Length;
            }
        }

        /// <summary>
        /// Read a Markers from stream, compare it to the desired Marker.
        /// </summary>
        /// <param name="desiredMarker">The desired Markers.</param>
        /// <returns>Whether read Markers equals the desired Markers.</returns>
        public bool ReadMarker(Markers desiredMarker)
        {
            byte[] buffer = new byte[PidLength];
            int len = this.Read(buffer, 0, PidLength);
            if (len == PidLength)
            {
                uint marker = BitConverter.ToUInt32(buffer, 0);
                return (Markers)marker == desiredMarker;
            }

            return false;
        }

        /// <summary>
        /// Read a MetaProperties value from stream,
        /// and advance the position within the stream by 4,
        /// compare it to the desired MetaProperties
        /// </summary>
        /// <param name="meta">The desired Markers</param>
        /// <returns>If read value equals the desired Markers return true, else false.</returns>
        public bool ReadMetaProperty(MetaProperties meta)
        {
            return this.ReadUInt32() == (uint)meta;
        }

        /// <summary>
        /// Read a Markers value from stream,
        /// and advance the position within the stream by 4
        /// </summary>
        /// <returns>The Markers value</returns>
        public Markers ReadMarker()
        {
            byte[] buffer = new byte[PidLength];
            int len = this.Read(buffer, 0, PidLength);
            uint marker;
            if (len == PidLength)
            {
                marker = BitConverter.ToUInt32(buffer, 0);
                return (Markers)marker;
            }

            AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");

            ////return a invalid value.
            return (Markers)0;
        }

        /// <summary>
        /// Read a uint value from stream,
        /// and advance the position within the stream by 4
        /// </summary>
        /// <returns>The uint value.</returns>
        public uint ReadUInt32()
        {
            byte[] buffer = new byte[4];
            this.Read(buffer, 0, buffer.Length);
            return BitConverter.ToUInt32(buffer, 0);
        }

        /// <summary>
        /// Read an int value from stream,
        /// and advance the position within the stream by 4
        /// </summary>
        /// <returns>The int value.</returns>
        public int ReadInt32()
        {
            byte[] buffer = new byte[4];
            this.Read(buffer, 0, buffer.Length);
            return BitConverter.ToInt32(buffer, 0);
        }

        /// <summary>
        /// Read a unsigned short integer value from stream,
        /// and advance the position within the stream by 2
        /// </summary>
        /// <returns>The unsigned short integer value</returns>
        public ushort ReadUInt16()
        {
            byte[] buffer = new byte[2];
            this.Read(buffer, 0, buffer.Length);
            return BitConverter.ToUInt16(buffer, 0);
        }

        /// <summary>
        /// Read a short value from stream,
        /// and advance the position within the stream by 2
        /// </summary>
        /// <returns>The short value</returns>
        public short ReadInt16()
        {
            byte[] buffer = new byte[2];
            this.Read(buffer, 0, buffer.Length);
            return BitConverter.ToInt16(buffer, 0);
        }

        /// <summary>
        /// Read a long value from stream,
        /// and advance the position within the stream by 8
        /// </summary>
        /// <returns>The long value</returns>
        public long ReadInt64()
        {
            byte[] buffer = new byte[8];
            this.Read(buffer, 0, buffer.Length);
            return BitConverter.ToInt64(buffer, 0);
        }

        /// <summary>
        /// Read the unsigned long integer value from stream,
        /// and advance the position within the stream by 8
        /// </summary>
        /// <returns>The unsigned long integer value</returns>
        public ulong ReadUInt64()
        {
            byte[] buffer = new byte[8];
            this.Read(buffer, 0, buffer.Length);
            return BitConverter.ToUInt64(buffer, 0);
        }

        /// <summary>
        /// Read a float value from stream,
        /// and advance the position within the stream by 4
        /// </summary>
        /// <returns>The float value</returns>
        public float ReadFloating32()
        {
            byte[] buffer = new byte[4];
            this.Read(buffer, 0, PidLength);
            return BitConverter.ToSingle(buffer, 0);
        }

        /// <summary>
        /// Read a double value from stream,
        /// and advance the position within the stream by 8
        /// </summary>
        /// <returns>The double value</returns>
        public double ReadFloating64()
        {
            byte[] buffer = new byte[8];
            this.Read(buffer, 0, buffer.Length);
            return BitConverter.ToDouble(buffer, 0);
        }

        /// <summary>
        /// Read a currency value from stream,
        /// and advance the position within the stream by 8
        /// </summary>
        /// <returns>The long value represents a currency value</returns>
        public long ReadCurrency()
        {
            return this.ReadInt64();
        }

        /// <summary>
        /// Read a FloatingTime value from stream,
        /// and advance the position within the stream by 8
        /// </summary>
        /// <returns>The double value represents a FloatingTime value</returns>
        public double ReadFloatingTime()
        {
            return this.ReadFloating64();
        }

        /// <summary>
        /// Read a Boolean value from stream,
        /// and advance the position within the stream by 2
        /// </summary>
        /// <returns>The unsigned short integer value represents a Boolean value</returns>
        public ushort ReadBoolean()
        {
            return this.ReadUInt16();
        }

        /// <summary>
        /// Read a Time value from stream,
        /// and advance the position within the stream by 2
        /// </summary>
        /// <returns>The unsigned long integer value represents a Time value</returns>
        public ulong ReadTime()
        {
            return this.ReadUInt64();
        }

        /// <summary>
        /// Read a string value from stream,
        /// and advance the position.
        /// </summary>
        /// <returns>The string value</returns>
        public string ReadString()
        {
            char tmp;
            byte[] buffer = new byte[2];
            StringBuilder b = new StringBuilder();
            do
            {
                this.Read(buffer, 0, 2);
                tmp = BitConverter.ToChar(buffer, 0);
                b.Append(tmp);
            }
            while (tmp != 0);
            return b.ToString();
        }

        /// <summary>
        /// Read a GUID value from stream,
        /// and advance the position.
        /// </summary>
        /// <returns>The GUID value</returns>
        public Guid ReadGuid()
        {
            byte[] buffer = new byte[Guid.Empty.ToByteArray().Length];
            this.Read(buffer, 0, buffer.Length);
            return new Guid(buffer);
        }

        /// <summary>
        /// Read a ASCII string value from stream,
        /// and advance the position.
        /// </summary>
        /// <returns>The ASCII string value</returns>
        public string ReadString8()
        {
            byte tmp;
            byte[] buffer = new byte[1];
            StringBuilder b = new StringBuilder();
            do
            {
                this.Read(buffer, 0, 1);
                tmp = buffer[0];
                b.Append(Convert.ToChar(tmp));
            }
            while (tmp != 0);
            return b.ToString();
        }

        /// <summary>
        /// Read a byte value from stream,
        /// and advance the position.
        /// </summary>
        /// <returns>The byte value</returns>
        public byte ReadUInt8()
        {
            byte result = 0;
            int tmp = ReadByte();
            if (tmp == -1)
            {
                AdapterHelper.Site.Assert.Fail("The end of the stream has been reached.");
            }
            else if (tmp >= 0)
            {
                result = (byte)tmp;
            }
            else
            {
                AdapterHelper.Site.Assert.Fail(string.Format("Expecting to get an unsigned integer from the stream, but the actual value is {0}.", tmp));
            }

            return result;
        }

        /// <summary>
        /// Read  bytes from stream,
        /// and advance the position.
        /// </summary>
        /// <param name="size">The size of bytes</param>
        /// <returns>The bytes array</returns>
        public byte[] ReadBlock(int size)
        {
            byte[] buffer = new byte[size];
            this.Read(buffer, 0, size);
            return buffer;
        }

        /// <summary>
        /// Read a list of blocks
        /// and advance the position.
        /// </summary>
        /// <param name="totalSize">The total number of bytes to read</param>
        /// <param name="blockSize">The size of each block</param>
        /// <returns>A list of blocks</returns>
        public List<byte[]> ReadBlocks(int totalSize, int blockSize)
        {
            int i;
            List<byte[]> l = new List<byte[]>();
            for (i = 0; i < totalSize; i += blockSize)
            {
                l.Add(this.ReadBlock(blockSize));
            }

            Debug.Assert(i == totalSize, "Assure have read all the data.");
            return l;
        }

        /// <summary>
        /// Read an int value and block
        /// and advance the position.
        /// </summary>
        /// <returns>A bytes array and a int value specifies the length of the bytes array</returns>
        public Tuple<int, byte[]> ReadLengthBlock()
        {
            int tmp = this.ReadInt32();
            byte[] buffer = this.ReadBlock(tmp);
            return new Tuple<int, byte[]>(tmp, buffer);
        }

        /// <summary>
        /// Read a list of an int value and a block
        /// and advance the position.
        /// </summary>
        /// <param name="totalLength">The number of bytes to read</param>
        /// <returns>A list of an int value and a block</returns>
        public List<Tuple<int, byte[]>> ReadLengthBlocks(int totalLength)
        {
            int i = 0;
            List<Tuple<int, byte[]>> list = new List<Tuple<int, byte[]>>();
            while (i < totalLength)
            {
                Tuple<int, byte[]> tmp = this.ReadLengthBlock();
                i++;
                list.Add(tmp);
            }

            Debug.Assert(i == totalLength, "Assure have read all the data.");
            return list;
        }

        /// <summary>
        /// Read a ServerID
        /// and advance the position.
        /// </summary>
        /// <returns>A ServerID</returns>
        public ServerID ReadServerID()
        {
            ServerID s;
            s.Ours = this.ReadUInt8();
            s.FID = this.ReadUInt64();
            s.MID = this.ReadUInt64();
            s.Instance = this.ReadUInt32();
            return s;
        }

        /// <summary>
        /// Get a uint value
        /// and do not advance the position.
        /// </summary>
        /// <returns>A uint value </returns>
        public uint VerifyUInt32()
        {
            return BitConverter.ToUInt32(
                this.GetBuffer(),
                (int)this.Position);
        }

        /// <summary>
        /// Get a uint value for current position plus an offset
        /// and do not advance the position.
        /// </summary>
        /// <param name="offset">An int value</param>
        /// <returns>A uint value</returns>
        public uint VerifyUInt32(int offset)
        {
            return BitConverter.ToUInt32(
                this.GetBuffer(),
                (int)this.Position + offset);
        }

        /// <summary>
        /// Get an unsigned short integer value for current position plus an offset
        /// and does not advance the position.
        /// </summary>
        /// <returns>An unsigned short integer value</returns>
        public ushort VerifyUInt16()
        {
            return BitConverter.ToUInt16(
                this.GetBuffer(),
                (int)this.Position);
        }

        /// <summary>
        /// Get an unsigned short integer value for current position plus an offset
        /// and do not advance the position.
        /// </summary>
        /// <param name="offset">An int value</param>
        /// <returns>An unsigned short integer value</returns>
        public ushort VerifyUInt16(int offset)
        {
            return BitConverter.ToUInt16(
                this.GetBuffer(),
                (int)this.Position + offset);
        }

        /// <summary>
        /// Backtracks the position
        /// </summary>
        /// <param name="count">An int value</param>
        public void BacktracePosition(int count)
        {
            this.Position -= count;
        }

        /// <summary>
        /// Backtracks the position with PidTagLength size.
        /// </summary>
        public void BacktraceMarker()
        {
            this.Position -= MarkersHelper.PidTagLength;
        }

        /// <summary>
        /// Indicate the Markers at the position equals a specified Markers.
        /// </summary>
        /// <param name="marker">A Markers value</param>
        /// <returns>True if the Markers at the position equals to the specified Markers,
        /// else false.
        /// </returns>
        public bool VerifyMarker(Markers marker)
        {
            return this.Verify((uint)marker);
        }

        /// <summary>
        /// Get a Markers value from current position plus an offset
        /// and do not advance the position.
        /// </summary>
        /// <returns>A Markers value</returns>
        public Markers VerifyMarker()
        {
            return (Markers)this.VerifyUInt32();
        }

        /// <summary>
        /// Get a Markers value from current position plus an offset
        /// and do not advance the position.
        /// </summary>
        /// <param name="offset">An int value</param>
        /// <returns>A Markers value</returns>
        public Markers VerifyMarker(int offset)
        {
            return (Markers)this.VerifyUInt32(offset);
        }

        /// <summary>
        /// Indicate the Markers at the current position plus an offset
        /// equals a specified Markers
        /// </summary>
        /// <param name="marker">A Markers to be verified</param>
        /// <param name="offset">An int value</param>
        /// <returns>
        /// True if the Markers at the current position plus an offset
        /// equals a specified Markers, else false.</returns>
        public bool VerifyMarker(Markers marker, int offset)
        {
            return this.Verify((uint)marker, offset);
        }

        /// <summary>
        /// Indicate the MetaProperties at the position equals a specified MetaProperties
        /// </summary>
        /// <param name="meta">A MetaProperties value</param>
        /// <returns>True if the MetaProperties at the position 
        /// equals the specified MetaProperties, 
        /// else false.
        /// </returns>
        public bool VerifyMetaProperty(MetaProperties meta)
        {
            return !this.IsEndOfStream && this.Verify((uint)meta, 0);
        }

        /// <summary>
        /// Indicate the MetaProperties at the position equals a specified PidTagFXErrorInfo
        /// </summary>
        /// <param name="offset">An int value</param>
        /// <returns>True if the uint at the position 
        /// equals the PidTagFXErrorInfo marker,
        /// else false.
        /// </returns>
        public bool VerifyErrorInfo(int offset)
        {
            return this.VerifyMarker(Markers.PidTagFXErrorInfo, offset);
        }

        /// <summary>
        /// Indicate the unsigned short integer value at the position equals a specified unsigned short integer value.
        /// </summary>
        /// <param name="val">An unsigned short integer value</param>
        /// <returns>True if the unsigned short integer at the position 
        /// equals to the specified unsigned short integer,
        /// else false.
        /// </returns>
        public bool Verify(ushort val)
        {
            return !this.IsEndOfStream && BitConverter.ToUInt16(
                this.GetBuffer(),
                (int)this.Position) == val;
        }

        /// <summary>
        /// Indicate the uint value at the position equals a specified uint value.
        /// </summary>
        /// <param name="val">A uint value.</param>
        /// <returns>True if the uint at the position 
        /// equals the specified uint.
        /// else false.
        /// </returns>
        public bool Verify(uint val)
        {
            return !this.IsEndOfStream && BitConverter.ToUInt32(
                this.GetBuffer(),
                (int)this.Position) == val;
        }

        /// <summary>
        /// Indicate the uint value at the position plus an offset
        /// equals a specified uint value.
        /// </summary>
        /// <param name="val">A uint value</param>
        /// <param name="offset">An int value</param>
        /// <returns>
        /// True if the uint at the position plus an offset
        /// equals the specified uint,
        /// else false.</returns>
        public bool Verify(uint val, int offset)
        {
            return !this.IsEndOfStream && BitConverter.ToUInt32(
                this.GetBuffer(),
                (int)this.Position + offset) == val;
        }

        /// <summary>
        /// Indicate the byte value at the position plus an offset
        /// equals a specified byte
        /// </summary>
        /// <param name="val">A uint value</param>
        /// <param name="offset">An int value</param>
        /// <returns>
        /// True if the byte at the position plus an offset
        /// equals the specified byte, 
        /// else false.</returns>
        public bool Verify(byte val, int offset)
        {
            byte[] tmp = this.GetBuffer();
            return !this.IsEndOfStream && tmp[(int)this.Position + offset] == val;
        }
    }
}