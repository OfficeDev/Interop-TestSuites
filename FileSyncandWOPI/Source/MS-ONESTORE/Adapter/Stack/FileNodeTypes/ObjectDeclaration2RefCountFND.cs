namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ObjectDeclaration2RefCountFND structure.
    /// </summary>
    public class ObjectDeclaration2RefCountFND : FileNodeBase
    {
        /// <summary>
        /// The format of stp field.
        /// </summary>
        private uint stpFormat;
        /// <summary>
        /// The format of cb field.
        /// </summary>
        private uint cbFormat;

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="stpFormat">The value of stpFormat.</param>
        /// <param name="cbFormat">The value of cbFormat.</param>
        public ObjectDeclaration2RefCountFND(uint stpFormat, uint cbFormat)
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
        public byte cRef { get; set; }

        public ObjectSpaceObjectPropSet PropertySet { get; set; }

        /// <summary>
        /// This method is used to deserialize the ObjectDeclaration2RefCountFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectDeclaration2RefCountFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.BlobRef = new FileNodeChunkReference(this.stpFormat, this.cbFormat);
            int len = this.BlobRef.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.body = new ObjectDeclaration2Body();
            len = this.body.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.cRef = byteArray[index];
            index += 1;

            if (OneNoteRevisionStoreFile.IsEncryption == false)
            {
                this.PropertySet = new ObjectSpaceObjectPropSet();
                this.PropertySet.DoDeserializeFromByteArray(byteArray, (int)this.BlobRef.StpValue);
            }
            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of ObjectDeclaration2RefCountFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectDeclaration2RefCountFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.BlobRef.SerializeToByteList());
            byteList.AddRange(this.body.SerializeToByteList());
            byteList.Add(this.cRef);

            return byteList;
        }
    }
}
