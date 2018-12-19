namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Common;
    using System.Collections.Generic;
    /// <summary>
    /// This class is used to represent a ObjectSpaceObjectStreamHeader.
    /// </summary>
    public class ObjectSpaceObjectStreamHeader
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the number of CompactID structures.
        /// </summary>
        public uint Count { get; set; }
        /// <summary>
        /// Gets or sets the Reserved field.
        /// </summary>
        public int Reserved { get; set; }
        /// <summary>
        /// Gets or sets the ExtendedStreamsPresent field.
        /// </summary>
        public int ExtendedStreamsPresent { get; set; }
        /// <summary>
        /// Gets or sets the OsidStreamNotPresent field.
        /// </summary>
        public int OsidStreamNotPresent { get; set; }
        /// <summary>
        /// This method is used to convert the element of ObjectSpaceObjectStreamHeader into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectSpaceObjectStreamHeader</returns>
        public List<byte> SerializeToByteList()
        {
            BitWriter bitWriter = new BitWriter(4);
            bitWriter.AppendUInit32(this.Count, 24);
            bitWriter.AppendInit32(this.Reserved, 6);
            bitWriter.AppendInit32(this.ExtendedStreamsPresent, 1);
            bitWriter.AppendInit32(this.OsidStreamNotPresent, 1);

            return new List<byte>(bitWriter.Bytes);
        }
        /// <summary>
        /// This method is used to deserialize the ObjectSpaceObjectStreamHeader object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectSpaceObjectStreamHeader object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            using (BitReader bitReader = new BitReader(byteArray, startIndex))
            {
                this.Count = bitReader.ReadUInt32(24);
                this.Reserved = bitReader.ReadInt32(6);
                this.ExtendedStreamsPresent = bitReader.ReadInt32(1);
                this.OsidStreamNotPresent = bitReader.ReadInt32(1);
                return 4;
            }
        }
    }
}
