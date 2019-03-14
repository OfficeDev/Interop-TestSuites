namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ObjectSpaceManifestListReferenceFND structrue.
    /// </summary>
    public class ObjectSpaceManifestListReferenceFND : FileNodeBase
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
        public ObjectSpaceManifestListReferenceFND(uint stpFormat, uint cbFormat)
        {
            this.stpFormat = stpFormat;
            this.cbFormat = cbFormat;
        }
        /// <summary>
        /// Gets or sets the value of ref field.
        /// </summary>
        public FileNodeChunkReference refField { get; set; }
        /// <summary>
        /// Gets or sets the value of gosid field.
        /// </summary>
        public ExtendedGUID gosid { get; set; }
        /// <summary>
        /// This method is used to deserialize the ObjectSpaceManifestListReferenceFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectSpaceManifestListReferenceFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.refField = new FileNodeChunkReference(this.stpFormat, this.cbFormat);
            int len = this.refField.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.gosid = new ExtendedGUID();
            len = this.gosid.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            return index - startIndex;
        }
        /// <summary>
        /// This method is used to convert the element of ObjectSpaceManifestListReferenceFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectSpaceManifestListReferenceFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.refField.SerializeToByteList());
            byteList.AddRange(this.gosid.SerializeToByteList());

            return byteList;
        }
    }
}
