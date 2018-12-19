namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the GlobalIdTableEntry2FNDX structure.
    /// </summary>
    public class GlobalIdTableEntry2FNDX : FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of iIndexMapFrom field.
        /// </summary>
        public uint iIndexMapFrom { get; set; }
        /// <summary>
        /// Gets or sets the value of iIndexMapTo field.
        /// </summary>
        public uint iIndexMapTo { get; set; }

        /// <summary>
        /// This method is used to deserialize the GlobalIdTableEntry2FNDX object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the GlobalIdTableEntry2FNDX object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.iIndexMapFrom = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.iIndexMapTo = BitConverter.ToUInt32(byteArray, index);
            index += 4;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of GlobalIdTableEntry2FNDX object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of GlobalIdTableEntry2FNDX.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(BitConverter.GetBytes(this.iIndexMapFrom));
            byteList.AddRange(BitConverter.GetBytes(this.iIndexMapTo));

            return byteList;
        }
    }
}
