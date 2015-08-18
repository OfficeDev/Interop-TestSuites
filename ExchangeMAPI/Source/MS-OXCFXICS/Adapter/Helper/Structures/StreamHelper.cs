namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.IO;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A helper class for stream operations
    /// </summary>
    public static class StreamHelper
    {
        /// <summary>
        /// The size of a GUID structure.
        /// </summary>
        public static readonly int GuidLength = Guid.Empty.ToByteArray().Length;

        /// <summary>
        /// Read a GUID value from a stream
        /// and advances the position within the stream
        /// by 16.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>A GUID value.</returns>
        public static Guid ReadGuid(Stream stream)
        {
            byte[] buffer = new byte[GuidLength];
            if (buffer.Length == stream.Read(buffer, 0, buffer.Length))
            {
                return new Guid(buffer);
            }
            else
            {
                AdapterHelper.Site.Assert.Fail("The stream is parsed to end.");
                return Guid.Empty;
            }
        }

        /// <summary>
        /// Read a byte from a stream
        /// and advances the position within the stream by 1.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>The byte read from the stream.</returns>
        public static byte ReadUInt8(Stream stream)
        {
            byte[] buffer = new byte[1];
            if (buffer.Length == stream.Read(buffer, 0, 1))
            {
                return buffer[0];
            }
            else
            {
                AdapterHelper.Site.Assert.Fail("The stream is parsed to end.");
                return 0;
            }
        }

        /// <summary>
        /// Read a ulong value from a stream
        /// and advances the position within the stream by 8.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>A ulong value.</returns>
        public static ulong ReadUInt64(Stream stream)
        {
            byte[] buffer = new byte[8];
            if (buffer.Length == stream.Read(buffer, 0, 8))
            {
                return BitConverter.ToUInt64(buffer, 0);
            }
            else
            {
                AdapterHelper.Site.Assert.Fail("The stream is parsed to end.");
                return 0;
            }
        }

        /// <summary>
        /// Read a uint value from a stream
        /// and advances the position within the stream by 4.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>A uint value.</returns>
        public static uint ReadUInt32(Stream stream)
        {
            byte[] buffer = new byte[4];
            if (buffer.Length == stream.Read(buffer, 0, buffer.Length))
            {
                return BitConverter.ToUInt32(buffer, 0);
            }
            else
            {
                AdapterHelper.Site.Assert.Fail("The stream is parsed to end.");
                return 0;
            }
        }

        /// <summary>
        /// Read a ushort value from a stream
        /// and advances the position within the stream by 2.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>A ushort value.</returns>
        public static ushort ReadUInt16(Stream stream)
        {
            byte[] buffer = new byte[2];
            if (buffer.Length == stream.Read(buffer, 0, buffer.Length))
            {
                return BitConverter.ToUInt16(buffer, 0);
            }
            else
            {
                AdapterHelper.Site.Assert.Fail("The stream is parsed to end.");
                return 0;
            }
        }

        /// <summary>
        /// Read an ASCII string value from a stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>The ASCII string value.</returns>
        public static string ReadString8(Stream stream)
        {
            byte tmp;
            byte[] buffer = new byte[1];
            StringBuilder b = new StringBuilder();
            do
            {
                stream.Read(buffer, 0, 1);
                tmp = buffer[0];
                b.Append(Convert.ToChar(tmp));
            }
            while (tmp != 0);
            return b.ToString();
        }

        /// <summary>
        /// Write a ushort value to a stream
        /// and advances the position within the stream by 2.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="value">A ushort value.</param>
        /// <returns>The number of bytes written to the stream.</returns>
        public static int WriteUInt16(Stream stream, ushort value)
        {
            byte[] buffer = BitConverter.GetBytes(value);
            stream.Write(buffer, 0, buffer.Length);
            return buffer.Length;
        }

        /// <summary>
        /// Write a uint value to a stream
        /// and advances the position within the stream by 4.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="value">A uint value.</param>
        /// <returns>The number of bytes written to the stream.</returns>
        public static int WriteUInt32(Stream stream, uint value)
        {
            byte[] buffer = BitConverter.GetBytes(value);
            stream.Write(buffer, 0, buffer.Length);
            return buffer.Length;
        }

        /// <summary>
        /// Write a ulong value to a stream
        /// and advances the position within the stream by 8.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="value">A ulong value.</param>
        /// <returns>The number of bytes written to the stream.</returns>
        public static int WriteUInt64(Stream stream, ulong value)
        {
            byte[] buffer = BitConverter.GetBytes(value);
            stream.Write(buffer, 0, buffer.Length);
            return buffer.Length;
        }

        /// <summary>
        /// Write a GUID value to a stream
        /// and advances the position within the stream by 16
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="guid">A GUID value.</param>
        /// <returns>The number of bytes written to the stream.</returns>
        public static int WriteGuid(Stream stream, Guid guid)
        {
            byte[] buffer = guid.ToByteArray();
            return WriteBuffer(stream, buffer);
        }

        /// <summary>
        /// Write bytes to a stream
        /// and advances the position within the stream by 16.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="buffer">A byte array.</param>
        /// <returns>The number of bytes written to the stream.</returns>
        public static int WriteBuffer(Stream stream, byte[] buffer)
        {
            stream.Write(buffer, 0, buffer.Length);
            return buffer.Length;
        }

        /// <summary>
        /// Write  a LongTermId value to a stream
        /// and advances the position within the stream by 24.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="id">A LongTermId value.</param>
        /// <returns>The number of bytes written to the stream.</returns>
        public static int WriteLongTermId(Stream stream, LongTermId id)
        { 
            int size = 0;
            size += StreamHelper.WriteGuid(stream, new Guid(id.DatabaseGuid));
            size += StreamHelper.WriteBuffer(stream, id.GlobalCounter);
            size += StreamHelper.WriteBuffer(stream, new byte[2] { 0x00, 0x00 });
            return size;
        }

        /// <summary>
        /// Read a LongTermId from a stream
        /// and advances the position within the stream by 24.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>The number of bytes read from the stream.</returns>
        public static LongTermId ReadLongTermId(Stream stream)
        {
            LongTermId id = new LongTermId();
            id.DatabaseGuid = StreamHelper.ReadGuid(stream).ToByteArray();

            id.GlobalCounter = new byte[6];
            stream.Read(id.GlobalCounter, 0, 6);
            stream.Read(new byte[2], 0, 2);
            return id;
        }
    }
}