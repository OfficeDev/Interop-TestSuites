namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;
    using System.IO;

    /// <summary>
    /// Contains common bytes while serializing a GLOBSET.
    /// </summary>
    public class CommonByteStack : Stack<byte[]>
    {
        /// <summary>
        /// Gets byte count.
        /// </summary>
        public int Bytes
        {
            get
            {
                int count = 0;
                if (this.Count > 0)
                {
                    foreach (byte[] arr in this)
                    {
                        count += arr.Length;
                    }

                    return count;
                }
                else
                {
                    return 0;
                }
            }
        }

        /// <summary>
        /// Reverse elements in a byte[][] array.
        /// </summary>
        /// <param name="arrs">A byte[][] array.</param>
        public void ReverseArray(byte[][] arrs)
        {
            for (int i = 0; i < arrs.Length / 2; i++)
            {
                byte[] tmp = arrs[i];
                arrs[i] = arrs[arrs.Length - 1 - i];
                arrs[arrs.Length - 1 - i] = tmp;
            }
        }

        /// <summary>
        /// Get all bytes.
        /// </summary>
        /// <param name="reverse">To reverse byte[]s order in the stack or not.</param>
        /// <returns>All bytes contained by this object.</returns>
        public byte[] ToBytes(bool reverse)
        {
            byte[][] arrs = this.ToArray();
            byte[] r = new byte[this.Bytes];
            int index = 0;
            if (reverse)
            {
                this.ReverseArray(arrs);
            }

            for (int i = 0; i < arrs.Length; i++)
            {
                byte[] arr = arrs[i];
                Array.Copy(arr, 0, r, index, arr.Length);
                index += arr.Length;
            }

            return r;
        }

        /// <summary>
        /// Get all bytes in reserved byte[] order.
        /// </summary>
        /// <returns>All bytes contained by this object.</returns>
        public byte[] GetCommonBytes()
        {
            return this.ToBytes(true);
        }

        /// <summary>
        /// Write all bytes in reserved byte[] order to a stream.
        /// </summary>
        /// <param name="stream">Bytes write to.</param>
        /// <returns>Byte count has been written.</returns>
        public int WriteToStream(Stream stream)
        {
            byte[][] arrs = this.ToArray();
            this.ReverseArray(arrs);
            foreach (byte[] arr in arrs)
            {
                stream.Write(arr, 0, arr.Length);
            }

            return this.Bytes;
        }
    }
}