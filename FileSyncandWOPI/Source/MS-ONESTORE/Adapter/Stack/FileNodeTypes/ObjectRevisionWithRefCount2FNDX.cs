namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Common;
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ObjectRevisionWithRefCount2FNDX structure.
    /// </summary>
    public class ObjectRevisionWithRefCount2FNDX : FileNodeBase
    {
        /// <summary>
        /// The format of stp field.
        /// </summary>
        private uint stpFormat;
        /// <summary>
        /// The format of cb field.
        /// </summary>
        private uint cbFormat;

        public ObjectRevisionWithRefCount2FNDX(uint stpFormat, uint cbFormat)
        {
            this.stpFormat = stpFormat;
            this.cbFormat = cbFormat;
        }

        /// <summary>
        /// Gets or sets the value of ref field.
        /// </summary>
        public FileNodeChunkReference Ref { get; set; }

        /// <summary>
        /// Gets or sets the value of oid field.
        /// </summary>
        public CompactID oid { get; set; }

        /// <summary>
        /// Gets or sets the value of fHasOidReferences field.
        /// </summary>
        public int fHasOidReferences { get; set; }
        /// <summary>
        /// Gets or sets the value of fHasOsidReferences field.
        /// </summary>
        public int fHasOsidReferences { get; set; }
        /// <summary>
        /// Gets or sets the value of Reserved field.
        /// </summary>
        public int Reserved { get; set; }

        /// <summary>
        /// Gets or sets the value of cRef field.
        /// </summary>
        public uint cRef { get; set; }

        /// <summary>
        /// Gets or sets the value of ObjectSpaceObjectPropSet.
        /// </summary>
        public ObjectSpaceObjectPropSet PropertySet { get; set; }
        /// <summary>
        /// This method is used to deserialize the ObjectRevisionWithRefCount2FNDX object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectRevisionWithRefCount2FNDX object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.Ref = new FileNodeChunkReference(this.stpFormat, this.cbFormat);
            int len = this.Ref.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.PropertySet = new ObjectSpaceObjectPropSet();
            this.PropertySet.DoDeserializeFromByteArray(byteArray, (int)this.Ref.StpValue);
            this.oid = new CompactID();
            len = this.oid.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            using (BitReader bitReader = new BitReader(byteArray, index))
            {
                this.fHasOidReferences = bitReader.ReadInt32(1);
                this.fHasOsidReferences = bitReader.ReadInt32(1);
                this.Reserved = bitReader.ReadInt32(30);
            }
            index += 4;
            this.cRef = BitConverter.ToUInt32(byteArray, index);
            index += 4;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of ObjectRevisionWithRefCount2FNDX object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectRevisionWithRefCount2FNDX.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.Ref.SerializeToByteList());
            byteList.AddRange(this.oid.SerializeToByteList());
            BitWriter bitWriter = new BitWriter(4);
            bitWriter.AppendInit32(this.fHasOidReferences, 1);
            bitWriter.AppendInit32(this.fHasOsidReferences, 1);
            bitWriter.AppendInit32(this.Reserved, 30);
            byteList.AddRange(bitWriter.Bytes);
            byteList.AddRange(BitConverter.GetBytes(this.cRef));

            return byteList;
        }
    }
}
