namespace Microsoft.Protocols.TestSuites.MS_ASWBXML
{
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Implements a byte queue for MS-ASWBXML
    /// </summary>
    internal class ByteQueue : Queue<byte>
    {
        /// <summary>
        /// Initializes a new instance of the ByteQueue class.
        /// </summary>
        /// <param name="bytes">The encoded bytes</param>
        public ByteQueue(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Dequeues a multibyte int
        /// </summary>
        /// <returns>The multibyte int</returns>
        public int DequeueMultibyteInt()
        {
            int returnValue = 0;
            byte singleByte;

            do
            {
                returnValue <<= 7;

                singleByte = this.Dequeue();
                returnValue += (int)(singleByte & 0x7F);
            }
            while (CheckContinuationBit(singleByte));

            return returnValue;
        }

        /// <summary>
        /// Dequeues a string
        /// </summary>
        /// <returns>The string.</returns>
        public string DequeueString()
        {
            StringBuilder builder = new StringBuilder();
            byte currentByte;
            do
            {
                currentByte = this.Dequeue();
                if (currentByte != 0x00)
                {
                    builder.Append((char)currentByte);
                }
            }
            while (currentByte != 0x00);

            return builder.ToString();
        }

        /// <summary>
        /// Dequeues a string in specified length
        /// </summary>
        /// <param name="length">The length of the string to dequeue.</param>
        /// <returns>The dequeued string.</returns>
        public string DequeueString(int length)
        {
            StringBuilder builder = new StringBuilder();

            for (int i = 0; i < length; i++)
            {
                byte currentByte = this.Dequeue();
                builder.Append((char)currentByte);
            }

            return builder.ToString();
        }

        /// <summary>
        /// Dequeues a base64 string in specified length
        /// </summary>
        /// <param name="length">The length of the string to dequeue.</param>
        /// <returns>The base64 string.</returns>
        public string DequeueBase64String(int length)
        {
            byte[] byteArray = new byte[length];
            for (int i = 0; i < length; i++)
            {
                byteArray[i] = this.Dequeue();
            }

            return System.Convert.ToBase64String(byteArray, 0, byteArray.Length);
        }

        /// <summary>
        /// Checks if the continuation bit is set
        /// </summary>
        /// <param name="byteval">the byte need to check.</param>
        /// <returns>True, if the continuation bit is set. Otherwise, False.</returns>
        private static bool CheckContinuationBit(byte byteval)
        {
            return (0x80 & byteval) != 0;
        }
    }
}