namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    /// <summary>
    /// This class is used to represent the ObjectSpaceManifestListStartFND structrue.
    /// </summary>
    public class ObjectSpaceManifestListStartFND : FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of gosid field.
        /// </summary>
        public ExtendedGUID gosid { get; set; }
        /// <summary>
        /// This method is used to deserialize the ObjectSpaceManifestListStartFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectSpaceManifestListStartFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            this.gosid = new ExtendedGUID();
            return this.gosid.DoDeserializeFromByteArray(byteArray, startIndex);
        }
        /// <summary>
        /// This method is used to convert the element of ObjectSpaceManifestListStartFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectSpaceManifestListStartFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            return this.gosid.SerializeToByteList();
        }
    }
}
