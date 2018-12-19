namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Common;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ObjectDeclaration2Body structure.
    /// </summary>
    public class ObjectDeclaration2Body
    {
        /// <summary>
        /// Gets or sets the value of oid field.
        /// </summary>
        public CompactID oid { get; set; }

        /// <summary>
        /// Gets or sets the value of jcid field.
        /// </summary>
        public JCID jcid { get; set; }

        /// <summary>
        /// Gets or sets the value of fHasOidReferences field.
        /// </summary>
        public uint fHasOidReferences { get; set; }

        /// <summary>
        /// Gets or sets the value of fHasOsidReferences field.
        /// </summary>
        public uint fHasOsidReferences { get; set; }

        /// <summary>
        /// Gets or sets the value of fReserved2 field.
        /// </summary>
        public uint fReserved2 { get; set; }

        /// <summary>
        /// This method is used to deserialize the ObjectDeclaration2Body object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectDeclaration2Body object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.oid = new CompactID();
            int len = this.oid.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.jcid = new JCID();
            len = this.jcid.DoDeserializeFromByteArray(byteArray, index);
            index += len;

            using (BitReader bitReader = new BitReader(byteArray, index))
            {
                this.fHasOidReferences = bitReader.ReadUInt32(1);
                this.fHasOsidReferences = bitReader.ReadUInt32(1);
                this.fReserved2 = bitReader.ReadUInt32(6);
            }
            index += 1;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of ObjectDeclaration2Body object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectDeclaration2Body.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.oid.SerializeToByteList());
            byteList.AddRange(this.jcid.SerializeToByteList());
            BitWriter bitWriter = new BitWriter(6);
            bitWriter.AppendUInit32(this.fHasOidReferences, 1);
            bitWriter.AppendUInit32(this.fHasOsidReferences, 1);
            bitWriter.AppendUInit32(this.fReserved2, 6);

            byteList.AddRange(bitWriter.Bytes);

            return byteList;
        }
    }
}
