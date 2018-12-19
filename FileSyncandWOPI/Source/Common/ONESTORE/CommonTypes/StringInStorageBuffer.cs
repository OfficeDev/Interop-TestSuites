namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the StringInStorageBuffer structure.
    /// </summary>
    public class StringInStorageBuffer
    {
        /// <summary>
        /// Gets or sets the value of cch field.
        /// </summary>
        public uint Cch { get; set; }
        /// <summary>
        /// Gets or sets the value of StringData field.
        /// </summary>
        public string StringData { get; set; }

        /// <summary>
        /// This method is used to convert the element of ExtendedGUID object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ExtendedGUID</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(BitConverter.GetBytes(this.Cch));
            byteList.AddRange(System.Text.Encoding.Unicode.GetBytes(this.StringData));

            return byteList;
        }

        /// <summary>
        /// This method is used to deserialize the ExtendedGUID object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ExtendedGUID object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.Cch = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.StringData = System.Text.Encoding.Unicode.GetString(byteArray, index, (int)this.Cch * 2);
            index += (int)this.Cch * 2;

            return index - startIndex;
        }

    }
}
