namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the GlobalIdTableEntry3FNDX structure.
    /// </summary>
    public class GlobalIdTableEntry3FNDX : FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of iIndexCopyFromStart field.
        /// </summary>
        public uint iIndexCopyFromStart { get; set; }

        /// <summary>
        /// Gets or sets the value of cEntriesToCopy field.
        /// </summary>
        public uint cEntriesToCopy { get; set; }

        /// <summary>
        /// Gets or sets the value of iIndexCopyToStart field.
        /// </summary>
        public uint iIndexCopyToStart { get; set; }

        /// <summary>
        /// This method is used to deserialize the GlobalIdTableEntry3FNDX object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the GlobalIdTableEntry3FNDX object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.iIndexCopyFromStart = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.cEntriesToCopy = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.iIndexCopyToStart = BitConverter.ToUInt32(byteArray, index);
            index += 4;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of GlobalIdTableEntry3FNDX object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of GlobalIdTableEntry3FNDX.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(BitConverter.GetBytes(this.iIndexCopyFromStart));
            byteList.AddRange(BitConverter.GetBytes(this.cEntriesToCopy));
            byteList.AddRange(BitConverter.GetBytes(this.iIndexCopyToStart));

            return byteList;
        }
    }
}
