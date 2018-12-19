namespace Microsoft.Protocols.TestSuites.Common
{
    using System;

    /// <summary>
    /// Implement a converter which converts to/from little-endian byte arrays.
    /// </summary>
    public class LittleEndianBitConverter
    {
        /// <summary>
        /// Prevents a default instance of the LittleEndianBitConverter class from being created
        /// </summary>
        private LittleEndianBitConverter()
        { 
        }

        /// <summary>
        ///  Returns a 16-bit unsigned integer converted from two bytes at a specified position in a byte array.
        /// </summary>
        /// <param name="array">Specify an array of bytes.</param>
        /// <param name="index">Specify the starting position.</param>
        /// <returns>Returns a 16-bit unsigned integer formed by two bytes beginning at startIndex.</returns>
        public static ushort ToUInt16(byte[] array, int index)
        {
            CheckByteArgument(array, index, 2);
            return unchecked((ushort)ConvertFromBytes(array, index, 2));
        }

        /// <summary>
        ///  Returns a 32-bit unsigned integer converted from two bytes at a specified position in a byte array.
        /// </summary>
        /// <param name="array">Specify an array of bytes.</param>
        /// <param name="index">Specify the starting position.</param>
        /// <returns>Returns a 32-bit unsigned integer formed by two bytes beginning at startIndex.</returns>
        public static uint ToUInt32(byte[] array, int index)
        {
            CheckByteArgument(array, index, 4);
            return unchecked((uint)ConvertFromBytes(array, index, 4));
        }

        /// <summary>
        ///  Returns a 32-bit signed integer converted from two bytes at a specified position in a byte array.
        /// </summary>
        /// <param name="array">Specify an array of bytes.</param>
        /// <param name="index">Specify the starting position.</param>
        /// <returns>Returns a 32-bit signed integer formed by two bytes beginning at startIndex.</returns>
        public static int ToInt32(byte[] array, int index)
        {
            CheckByteArgument(array, index, 4);
            return unchecked((int)ConvertFromBytes(array, index, 4));
        }

        /// <summary>
        ///  Returns a 16-bit signed integer converted from two bytes at a specified position in a byte array.
        /// </summary>
        /// <param name="array">Specify an array of bytes.</param>
        /// <param name="index">Specify the starting position.</param>
        /// <returns>Returns a 16-bit signed integer formed by two bytes beginning at startIndex.</returns>
        public static short ToInt16(byte[] array, int index)
        {
            CheckByteArgument(array, index, 4);
            return unchecked((short)ConvertFromBytes(array, index, 2));
        }

        /// <summary>
        ///  Returns a 64-bit unsigned integer converted from two bytes at a specified position in a byte array.
        /// </summary>
        /// <param name="array">Specify an array of bytes.</param>
        /// <param name="index">Specify the starting position.</param>
        /// <returns>Returns a 64-bit unsigned integer formed by two bytes beginning at startIndex.</returns>
        public static ulong ToUInt64(byte[] array, int index)
        {
            CheckByteArgument(array, index, 8);
            return unchecked((ulong)ConvertFromBytes(array, index, 8));
        }

        /// <summary>
        /// Returns the specified 64-bit unsigned integer value as an array of bytes.
        /// </summary>
        /// <param name="value">Specify the number to convert.</param>
        /// <returns>Returns an array of bytes with length 8.</returns>
        public static byte[] GetBytes(ulong value)
        {
            byte[] buffer = new byte[8];
            ConvertToBytes(value, buffer);
            return buffer;
        }

        /// <summary>
        ///  Returns the specified 32-bit unsigned integer value as an array of bytes.
        /// </summary>
        /// <param name="value">Specify the number to convert.</param>
        /// <returns>Returns an array of bytes with length 4.</returns>
        public static byte[] GetBytes(uint value)
        {
            byte[] buffer = new byte[4];
            ConvertToBytes(value, buffer);
            return buffer;
        }

        /// <summary>
        /// Returns the specified 32-bit signed integer value as an array of bytes.
        /// </summary>
        /// <param name="value">Specify the number to convert.</param>
        /// <returns>Returns an array of bytes with length 4.</returns>
        public static byte[] GetBytes(int value)
        {
            byte[] buffer = new byte[4];
            ConvertToBytes((ulong)value, buffer);
            return buffer;
        }

        /// <summary>
        /// Returns the specified 16-bit unsigned integer value as an array of bytes.
        /// </summary>
        /// <param name="value">Specify the number to convert.</param>
        /// <returns>Returns an array of bytes with length 2.</returns>
        public static byte[] GetBytes(ushort value)
        {
            byte[] buffer = new byte[2];
            ConvertToBytes(value, buffer);
            return buffer;
        }

        /// <summary>
        /// Returns a value built from the specified number of bytes from the given buffer,
        /// starting at index.
        /// </summary>
        /// <param name="buffer">Specify the data in byte array format</param>
        /// <param name="startIndex">Specify the first index to use</param>
        /// <param name="bytesToConvert">Specify the number of bytes to use</param>
        /// <returns>Return the value built from the given bytes</returns>
        private static ulong ConvertFromBytes(byte[] buffer, int startIndex, int bytesToConvert)
        {
            ulong ret = 0;
            int bitCount = 0;
            for (int i = 0; i < bytesToConvert; i++)
            {
                ret |= (ulong)buffer[startIndex + i] << bitCount;

                bitCount += 8;
            }

            return ret;
        }

        /// <summary>
        /// This method is used to convert the specified value to the buffer.
        /// </summary>
        /// <param name="value">Specify the value to convert.</param>
        /// <param name="buffer">Specify the buffer which copies the bytes into.</param>
        private static void ConvertToBytes(ulong value, byte[] buffer)
        {
            for (int i = 0; i < buffer.Length; i++)
            {
                buffer[i] = unchecked((byte)(value & 0xff));
                value = value >> 8;
            }
        }

        /// <summary>
        /// This method is used to check the given argument for validity.
        /// </summary>
        /// <param name="value">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index.</param>
        /// <param name="bytesRequired">Specify the number of bytes.</param>
        /// <exception cref="ArgumentNullException">The byte array is a null reference.</exception>
        /// <exception cref="ArgumentOutOfRangeException">
        /// StartIndex is greater than the length of value minus bytesRequired.
        /// </exception>
        /// <exception cref="ArgumentException">
        /// StartIndex is less than zero.
        /// </exception>
        private static void CheckByteArgument(byte[] value, int startIndex, int bytesRequired)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            if (startIndex < 0)
            {
                throw new ArgumentException("The index cannot be less than 0.");
            }

            if (startIndex > value.Length - bytesRequired)
            {
                throw new ArgumentOutOfRangeException("startIndex");
            }
        }
    }
}