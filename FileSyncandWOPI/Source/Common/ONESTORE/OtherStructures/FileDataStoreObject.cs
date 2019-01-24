namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// This class is used to represent the FileDataStoreObject structure.
    /// </summary>
    public class FileDataStoreObject
    {
        /// <summary>
        /// The format of cb field.
        /// </summary>
        private uint size;

        public FileDataStoreObject(uint size)
        {
            this.size = size;
        }

        /// <summary>
        /// Gets or sets the value of guidHeader.
        /// </summary>
        public Guid guidHeader { get; set; }

        /// <summary>
        /// Gets or sets the value of cbLength.
        /// </summary>
        public ulong cbLength { get; set; }
        /// <summary>
        /// Gets or sets the value of unused.
        /// </summary>
        public uint unused { get; set; }

        /// <summary>
        ///   Gets or sets the value of reserved.
        /// </summary>
        public ulong reserved { get; set; }

        /// <summary>
        /// Gets or sets the value of FileData.
        /// </summary>
        public byte[] FileData { get; set; }

        /// <summary>
        /// Gets or sets the value of guidFooter.
        /// </summary>
        public Guid guidFooter { get; set; }

        /// <summary>
        /// This method is used to deserialize the FileDataStoreObject object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the FileDataStoreObject object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.guidHeader= this.ReadGuid(byteArray, index);
            index += 16;
            this.cbLength = BitConverter.ToUInt64(byteArray, index);
            index += 8;
            this.unused = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.reserved = BitConverter.ToUInt64(byteArray, index);
            index += 8;
            this.FileData = new byte[this.size - 52];
            Array.Copy(byteArray, index, this.FileData, 0, this.FileData.Length);
            index += this.FileData.Length;
            this.guidFooter= this.ReadGuid(byteArray, index);
            index += 16;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of FileDataStoreObject object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of FileDataStoreObject.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.guidHeader.ToByteArray());
            byteList.AddRange(BitConverter.GetBytes(this.cbLength));
            byteList.AddRange(BitConverter.GetBytes(this.unused));
            byteList.AddRange(BitConverter.GetBytes(this.reserved));
            byteList.AddRange(this.FileData);
            byteList.AddRange(this.guidFooter.ToByteArray());

            return byteList;
        }

        /// <summary>
        /// This method is used to read the Guid for byte array.
        /// </summary>
        /// <param name="byteArray">The byte array.</param>
        /// <param name="startIndex">The offset of the Guid value.</param>
        /// <returns>Return the value of Guid.</returns>
        private Guid ReadGuid(byte[] byteArray, int startIndex)
        {
            byte[] guidBuffer = new byte[16];
            Array.Copy(byteArray, startIndex, guidBuffer, 0, 16);

            return new Guid(guidBuffer);
        }
    }
}
