namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ObjectDeclaration2LargeRefCountFND structure.
    /// </summary>
    public class ObjectDeclaration2LargeRefCountFND : FileNodeBase
    {
        /// <summary>
        /// The format of stp field.
        /// </summary>
        private uint stpFormat;
        /// <summary>
        /// The format of cb field.
        /// </summary>
        private uint cbFormat;

        public ObjectDeclaration2LargeRefCountFND(uint stpFormat, uint cbFormat)
        {
            this.stpFormat = stpFormat;
            this.cbFormat = cbFormat;
        }

        /// <summary>
        /// Gets or sets the value of BlobRef field.
        /// </summary>
        public FileNodeChunkReference BlobRef { get; set; }

        /// <summary>
        /// Gets or sets the value of body field.
        /// </summary>
        public ObjectDeclaration2Body body { get; set; }

        /// <summary>
        /// Gets or sets the value of cRef field.
        /// </summary>
        public uint cRef { get; set; }
        /// <summary>
        /// Gets or sets the value of ObjectSpaceObjectPropSet.
        /// </summary>
        public ObjectSpaceObjectPropSet PropertySet { get; set; }
        /// <summary>
        /// This method is used to deserialize the ObjectDeclaration2LargeRefCountFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectDeclaration2LargeRefCountFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.BlobRef = new FileNodeChunkReference(this.stpFormat, this.cbFormat);
            int len = this.BlobRef.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.PropertySet = new ObjectSpaceObjectPropSet();
            this.PropertySet.DoDeserializeFromByteArray(byteArray, (int)this.BlobRef.StpValue);
            this.body = new ObjectDeclaration2Body();
            len = this.body.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.cRef = BitConverter.ToUInt32(byteArray, index);
            index += 4;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of ObjectDeclaration2LargeRefCountFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectDeclaration2LargeRefCountFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.BlobRef.SerializeToByteList());
            byteList.AddRange(this.body.SerializeToByteList());
            byteList.AddRange(BitConverter.GetBytes(this.cRef));

            return byteList;
        }
    }
}
