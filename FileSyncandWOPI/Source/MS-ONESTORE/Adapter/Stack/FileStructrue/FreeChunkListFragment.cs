namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the FreeChunkListFragment structure.
    /// </summary>
    public class FreeChunkListFragment
    {
        private ulong size = 0;

        public FreeChunkListFragment(ulong size)
        {
            this.size = size;
        }
        /// <summary>
        /// Gets or sets the value of crc field.
        /// </summary>
        public uint crc { get; set; }

        /// <summary>
        /// Gets or sets the value of fcrNextChunk field.
        /// </summary>
        public FileChunkReference64x32 fcrNextChunk { get; set; }
        /// <summary>
        /// Gets or sets the value of fcrFreeChunk field.
        /// </summary>
        public FileChunkReference64[] fcrFreeChunk { get; set; }

        /// <summary>
        /// This method is used to convert the element of FreeChunkListFragment object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of FreeChunkListFragment</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(BitConverter.GetBytes(this.crc));
            byteList.AddRange(this.fcrNextChunk.SerializeToByteList());

            foreach(FileChunkReference64 f in this.fcrFreeChunk)
            {
                byteList.AddRange(f.SerializeToByteList());
            }

            return byteList;
        }

        /// <summary>
        /// This method is used to deserialize the FreeChunkListFragment object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the FreeChunkListFragment object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.crc = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.fcrNextChunk = new FileChunkReference64x32();
            int len = this.fcrNextChunk.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            ulong count = (this.size - 16) / 16;
            this.fcrFreeChunk = new FileChunkReference64[count];
            for (ulong i = 0; i < count; i++)
            {
                this.fcrFreeChunk[i] = new FileChunkReference64();
                len = this.fcrFreeChunk[i].DoDeserializeFromByteArray(byteArray, index);
                index += len;
            }

            return index - startIndex;
        }
    }
}
