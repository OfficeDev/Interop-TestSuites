namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Security.Cryptography;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// Computes the CRC32 hash for the input data.
    /// </summary>
    public class CRC32 : HashAlgorithm
    {
        private UInt32[] crc32Table = new UInt32[256];
        private UInt32 crc32Result;

        /// <summary>
        /// Initializes a new instance of CRC32 class.
        /// </summary>
        public CRC32():this(0xEDB88320)
        {
            
        }

        public CRC32(uint polynomial) : base()
        {
            this.HashSizeValue = 32;
            for (uint i = 0; i < 256; i++)
            {
                UInt32 crc32 = i;
                for (int j = 8; j > 0; j--)
                {
                    if ((crc32 & 1) == 1)
                    {
                        crc32 = (crc32 >> 1) ^ polynomial;
                    }
                    else
                    {
                        crc32 >>= 1;
                    }
                }
                crc32Table[i] = crc32;
            }

            Initialize();
        }
        /// <summary>
        /// Gets the value of the computed CRC-32 hash code as unsigned 32-bit integer.
        /// </summary>
        /// <value>The current value of the computed CRC-32 hash code.</value>
        public UInt32 CRC32Hash { get; protected set; }

        public override void Initialize()
        {
            this.crc32Result = 0xFFFFFFFF;
        }
        /// <summary>
        /// Routes data written to the object into the CRC32 hash algorithm for computing the hash.
        /// </summary>
        /// <param name="array">The input to compute the hash code for.</param>
        /// <param name="ibStart">The offset into the byte array from which to begin using data.</param>
        /// <param name="cbSize">The number of bytes in the byte array to use as data.</param>
        protected override void HashCore(byte[] array, int ibStart, int cbSize)
        {
            int end = ibStart + cbSize;
            for (int i = ibStart; i < end; i++)
            {
                this.crc32Result = (this.crc32Result >> 8) ^ this.crc32Table[array[i] ^ (this.crc32Result & 0x000000FF)];
            }
        }
        /// <summary>
        /// Finalizes the hash computation after the last data is processed by the cryptographic stream object.
        /// </summary>
        /// <returns>The computed hash code.</returns>
        protected override byte[] HashFinal()
        {
            this.crc32Result = ~this.crc32Result;
            this.CRC32Hash = this.crc32Result;
            this.HashValue = BitConverter.GetBytes(this.crc32Result);
            return this.HashValue;
        }
    }
}
