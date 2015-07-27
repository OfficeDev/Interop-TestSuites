//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Security.Cryptography;

    /// <summary>
    /// A class implements MD4 algorithm to generate RDC signature.
    /// </summary>
    public class RDCSignatureGenerator : HashAlgorithm
    {
        /// <summary>
        /// A 4-word buffer is used to compute the message digest
        /// </summary>
        private uint[] state;

        /// <summary>
        /// A 16-word buffer is used to store current 16-word block.
        /// </summary>
        private uint[] x;

        /// <summary>
        /// The current index of byte which was processed.
        /// </summary>
        private int bytesProcessed;

        /// <summary>
        /// Initializes a new instance of the RDCSignatureGenerator class
        /// </summary>
        public RDCSignatureGenerator()
        {
            this.Initialize();
        }

        /// <summary>
        /// Initializes an implementation of the System.Security.Cryptography.HashAlgorithm class.
        /// </summary>
        public override void Initialize()
        {
            this.x = new uint[16];
            this.bytesProcessed = 0;

            this.state = new uint[4];
            this.state[0] = 0x67452301;
            this.state[1] = 0xefcdab89;
            this.state[2] = 0x98badcfe;
            this.state[3] = 0x10325476;
        }

        /// <summary>
        /// Routes data written to the object into the hash algorithm for computing the hash.
        /// </summary>
        /// <param name="array">The input to compute the hash code for.</param>
        /// <param name="start">The offset into the byte array from which to begin using data.</param>
        /// <param name="size">The number of bytes in the byte array to use as data.</param>
        protected override void HashCore(byte[] array, int start, int size)
        {
            this.ProcessMessage(array);
        }

        /// <summary>
        /// Finalizes the hash computation after the last data is processed by the cryptographic stream object.
        /// </summary>
        /// <returns>The computed hash code.</returns>
        protected override byte[] HashFinal()
        {
            try
            {
                this.ProcessMessage(this.Padding());

                BitWriter writer = new BitWriter(16);
                writer.AppendUInit32(this.state[0], 32);
                writer.AppendUInit32(this.state[1], 32);
                writer.AppendUInit32(this.state[2], 32);
                writer.AppendUInit32(this.state[3], 32);

                return writer.Bytes;
            }
            finally
            {
                this.Initialize();
            }
        }

        /// <summary>
        /// Computes the hash value for the specified region of the input byte.
        /// </summary>
        /// <param name="array">The specified region of the input byte</param>
        private void ProcessMessage(byte[] array)
        {
            for (int n = 0; n < array.Length; n++)
            {
                int c = this.bytesProcessed & 63;
                int i = c >> 2;
                int s = (c & 3) << 3;

                this.x[i] = (this.x[i] & ~((uint)255 << s)) | ((uint)array[n] << s);

                if (c == 63)
                {
                    this.MD4Transform();
                }

                this.bytesProcessed++;
            }
        }

        /// <summary>
        /// Get the padding byte array.
        /// </summary>
        /// <returns>The padding byte array.</returns>
        private byte[] Padding()
        {
            int padLenght = ((this.bytesProcessed % 64) < 56) ? (56 - (this.bytesProcessed % 64)) : (56 + 64 - (this.bytesProcessed % 64));
            byte[] padding = new byte[padLenght + 8];
            padding[0] = 0x80;
            BitConverter.GetBytes(this.bytesProcessed * 8).CopyTo(padding, padding.Length - 8);

            return padding;
        }

        /// <summary>
        /// MD4 basic transformation. Transforms state based on block.
        /// </summary>
        private void MD4Transform()
        {
            uint a = this.state[0];
            uint b = this.state[1];
            uint c = this.state[2];
            uint d = this.state[3];

            foreach (int i in new[] { 0, 4, 8, 12 })
            {
                a = this.Round1(a, b, c, d, this.x[i], 3);
                d = this.Round1(d, a, b, c, this.x[i + 1], 7);
                c = this.Round1(c, d, a, b, this.x[i + 2], 11);
                b = this.Round1(b, c, d, a, this.x[i + 3], 19);
            }

            foreach (int i in new[] { 0, 1, 2, 3 })
            {
                a = this.Round2(a, b, c, d, this.x[i], 3);
                d = this.Round2(d, a, b, c, this.x[i + 4], 5);
                c = this.Round2(c, d, a, b, this.x[i + 8], 9);
                b = this.Round2(b, c, d, a, this.x[i + 12], 13);
            }

            foreach (int i in new[] { 0, 2, 1, 3 })
            {
                a = this.Round3(a, b, c, d, this.x[i], 3);
                d = this.Round3(d, a, b, c, this.x[i + 8], 9);
                c = this.Round3(c, d, a, b, this.x[i + 4], 11);
                b = this.Round3(b, c, d, a, this.x[i + 12], 15);
            }

            this.state[0] += a;
            this.state[1] += b;
            this.state[2] += c;
            this.state[3] += d;
        }

        /// <summary>
        /// Rotate specify data left n bits .
        /// </summary>
        /// <param name="rotatedData">The data that be rotated.</param>
        /// <param name="n">The number of bits that rotate left.</param>
        /// <returns>The rotated data.</returns>
        private uint RotateLeft(uint rotatedData, int n)
        {
            return (rotatedData << n) | (rotatedData >> (32 - n));
        }

        /// <summary>
        /// The operation for round1.
        /// </summary>
        /// <param name="a">The state data a.</param>
        /// <param name="b">The state data b.</param>
        /// <param name="c">The state data c.</param>
        /// <param name="d">The state data d.</param>
        /// <param name="inputData">The specify input data.</param>
        /// <param name="s">The number of bits that rotate left.</param>
        /// <returns>The result after compute.</returns>
        private uint Round1(uint a, uint b, uint c, uint d, uint inputData, int s)
        {
            return this.RotateLeft(a + ((b & c) | (~b & d)) + inputData, s);
        }

        /// <summary>
        /// The operation for round2.
        /// </summary>
        /// <param name="a">The state data a.</param>
        /// <param name="b">The state data b.</param>
        /// <param name="c">The state data c.</param>
        /// <param name="d">The state data d.</param>
        /// <param name="inputData">The specify input data.</param>
        /// <param name="s">The number of bits that rotate left.</param>
        /// <returns>The result after compute.</returns>
        private uint Round2(uint a, uint b, uint c, uint d, uint inputData, int s)
        {
            return this.RotateLeft(a + ((b & c) | (b & d) | (c & d)) + inputData + 0x5a827999, s);
        }

        /// <summary>
        /// The operation for round3.
        /// </summary>
        /// <param name="a">The state data a.</param>
        /// <param name="b">The state data b.</param>
        /// <param name="c">The state data c.</param>
        /// <param name="d">The state data d.</param>
        /// <param name="inputData">The specify input data.</param>
        /// <param name="s">The number of bits that rotate left.</param>
        /// <returns>The result after compute.</returns>
        private uint Round3(uint a, uint b, uint c, uint d, uint inputData, int s)
        {
            return this.RotateLeft(a + (b ^ c ^ d) + inputData + 0x6ed9eba1, s);
        }
    }
}