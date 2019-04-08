namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Common;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ObjectDeclarationWithRefCountBody structure.
    /// </summary>
    public class ObjectDeclarationWithRefCountBody
    {
        /// <summary>
        /// Gets or sets the value of oid field.
        /// </summary>
        public CompactID oid { get; set; }

        /// <summary>
        /// Gets or sets the value of jci field.
        /// </summary>
        public uint jci { get; set; }

        /// <summary>
        /// Gets or sets the value of odcs field.
        /// </summary>
        public uint odcs { get; set; }

        /// <summary>
        /// Gets or sets the value of fReserved1 field.
        /// </summary>
        public uint fReserved1 { get; set; }

        /// <summary>
        /// Gets or sets the value of fHasOidReferences field.
        /// </summary>
        public uint fHasOidReferences { get; set; }

        /// <summary>
        /// Gets or sets the value of fHasOsidReferences field.
        /// </summary>
        public uint fHasOsidReferences { get; set; }

        /// <summary>
        /// Gets or sets the fReserved2 field.
        /// </summary>
        public uint fReserved2 { get; set; }

        /// <summary>
        /// This method is used to deserialize the ObjectDeclarationWithRefCountBody object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectDeclarationWithRefCountBody object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.oid = new CompactID();
            int len = this.oid.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            using (BitReader bitReader = new BitReader(byteArray, index))
            {
                this.jci = bitReader.ReadUInt32(10);
                this.odcs = bitReader.ReadUInt32(4);
                this.fReserved1 = bitReader.ReadUInt32(2);
                this.fHasOidReferences = bitReader.ReadUInt32(1);
                this.fHasOsidReferences = bitReader.ReadUInt32(1);
                this.fReserved2 = bitReader.ReadUInt32(30);
            }
            index += 6;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of ObjectDeclarationWithRefCountBody object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectDeclarationWithRefCountBody.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.oid.SerializeToByteList());
            BitWriter bitWriter = new BitWriter(6);
            bitWriter.AppendUInit32(this.jci, 10);
            bitWriter.AppendUInit32(this.odcs, 4);
            bitWriter.AppendUInit32(this.fReserved1, 2);
            bitWriter.AppendUInit32(this.fHasOidReferences, 1);
            bitWriter.AppendUInit32(this.fHasOsidReferences, 1);
            bitWriter.AppendUInit32(this.fReserved2, 30);

            byteList.AddRange(bitWriter.Bytes);

            return byteList;
        }
    }
}
