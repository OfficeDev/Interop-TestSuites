namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the RevisionManifestListReferenceFND structure.
    /// </summary>
    public class RevisionManifestListReferenceFND : FileNodeBase
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
        public RevisionManifestListReferenceFND(uint stpFormat, uint cbFormat)
        {
            this.stpFormat = stpFormat;
            this.cbFormat = cbFormat;
        }

        /// <summary>
        /// Gets or sets the value of ref field.
        /// </summary>
        public FileNodeChunkReference refField { get; set; }
        /// <summary>
        /// This method is used to deserialize the RevisionManifestListReferenceFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the RevisionManifestListReferenceFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            this.refField = new FileNodeChunkReference(this.stpFormat, this.cbFormat);

            return this.refField.DoDeserializeFromByteArray(byteArray, startIndex);
        }
        /// <summary>
        /// This method is used to convert the element of RevisionManifestListReferenceFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of RevisionManifestListReferenceFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            return this.refField.SerializeToByteList();
        }
    }
}
