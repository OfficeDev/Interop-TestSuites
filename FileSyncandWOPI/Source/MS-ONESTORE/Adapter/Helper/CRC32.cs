namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    /// <summary>
    /// Computes the CRC32 hash for the input data.
    /// </summary>
    public class CRC32
    {
        /// <summary>
        /// 
        /// </summary>
        private static uint[] Crc32Table;
        /// <summary>
        /// Generate CRC32 code table
        /// </summary>
        private static void GetCrc32Table()
        {
            uint Crc;
            Crc32Table = new uint[256];
            int i, j;
            for (i = 0; i < 256; i++)
            {
                Crc = (uint)i;
                for (j = 8; j > 0; j--)
                {
                    if ((Crc & 1) == 1)
                        Crc = (Crc >> 1) ^ 0xEDB88320;
                    else
                        Crc >>= 1;
                }
                Crc32Table[i] = Crc;
            }
        }
        /// <summary>
        /// Compute the CRC32 value.
        /// </summary>
        /// <param name="buffer">The byte array of the input data</param>
        /// <returns>Return the CRC value.</returns>
        public static uint GetCrc32(byte[] buffer)
        {
            GetCrc32Table();
            uint value = 0xffffffff;
            for (int i = 0; i < buffer.Length; i++)
            {
                value = ((value >> 8) & 0xffffffff) ^ Crc32Table[(value ^ buffer[i]) & 0xFF];
            }
            return value ^ 0xffffffff;
        }
    }
}
