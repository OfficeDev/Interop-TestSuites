namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the FileNodeListHeader structure.
    /// </summary>
    public class FileNodeListHeader
    {
        /// <summary>
        /// Gets or sets the value of uintMagic field.
        /// </summary>
        public UInt64 uintMagic { get; set; }

        /// <summary>
        /// Gets or sets the value of FileNodeListID field.
        /// </summary>
        public uint FileNodeListID { get; set; }

        /// <summary>
        /// Gets or sets the value of nFragmentSequence field.
        /// </summary>
        public uint nFragmentSequence { get; set; }

        /// <summary>
        /// This method is used to deserialize the FileNodeListHeader object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the FileNodeListHeader object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.uintMagic = BitConverter.ToUInt64(byteArray, index);
            index += 8;
            this.FileNodeListID = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.nFragmentSequence = BitConverter.ToUInt32(byteArray, index);
            index += 4;

            return index - startIndex;
        }
        /// <summary>
        /// This method is used to convert the element of FileNodeListHeader object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of FileNodeListHeader.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(BitConverter.GetBytes(this.uintMagic));
            byteList.AddRange(BitConverter.GetBytes(this.FileNodeListID));
            byteList.AddRange(BitConverter.GetBytes(this.nFragmentSequence));

            return byteList;
        }
    }
}
