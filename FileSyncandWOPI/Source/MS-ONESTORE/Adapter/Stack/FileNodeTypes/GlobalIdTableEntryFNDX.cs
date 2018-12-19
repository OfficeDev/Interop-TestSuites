namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// This class is used to represent the GlobalIdTableEntryFNDX structure.
    /// </summary>
    public class GlobalIdTableEntryFNDX : FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of index field.
        /// </summary>
        public uint index { get; set; }
        /// <summary>
        /// Gets or sets the value of guid field.
        /// </summary>
        public Guid guid { get; set; }

        /// <summary>
        /// This method is used to deserialize the GlobalIdTableEntryFNDX object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the GlobalIdTableEntryFNDX object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int idx = startIndex;
            this.index = BitConverter.ToUInt32(byteArray, idx);
            idx += 4;
            this.guid = AdapterHelper.ReadGuid(byteArray, idx);
            idx += 16;

            return idx - startIndex;
        }
        /// <summary>
        /// This method is used to convert the element of GlobalIdTableEntryFNDX object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of GlobalIdTableEntryFNDX.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(BitConverter.GetBytes(this.index));
            byteList.AddRange(this.guid.ToByteArray());

            return byteList;
        }
    }
}
