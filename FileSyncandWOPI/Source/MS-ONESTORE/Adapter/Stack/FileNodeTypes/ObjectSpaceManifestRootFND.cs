namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ObjectSpaceManifestRootFND structrue.
    /// </summary>
    public class ObjectSpaceManifestRootFND : FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of gosidRoot field.
        /// </summary>
        public ExtendedGUID gosidRoot { get; set; }
        /// <summary>
        /// This method is used to deserialize the ObjectSpaceManifestRootFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectSpaceManifestRootFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.gosidRoot = new ExtendedGUID();
            int len = this.gosidRoot.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            return index - startIndex;
        }
        /// <summary>
        /// This method is used to convert the element of ObjectSpaceManifestRootFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectSpaceManifestRootFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            this.gosidRoot = new ExtendedGUID();
            return this.gosidRoot.SerializeToByteList();
        }
    }
}
