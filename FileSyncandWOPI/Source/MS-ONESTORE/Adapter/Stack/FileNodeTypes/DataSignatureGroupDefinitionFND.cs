namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the DataSignatureGroupDefinitionFND structure.
    /// </summary>
    public class DataSignatureGroupDefinitionFND:FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of DataSignatureGroup field.
        /// </summary>
        public ExtendedGUID DataSignatureGroup { get; set; }

        /// <summary>
        /// This method is used to deserialize the DataSignatureGroupDefinitionFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the DataSignatureGroupDefinitionFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            this.DataSignatureGroup = new ExtendedGUID();
            return this.DataSignatureGroup.DoDeserializeFromByteArray(byteArray, startIndex);
        }
        /// <summary>
        /// This method is used to convert the element of DataSignatureGroupDefinitionFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of DataSignatureGroupDefinitionFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            return this.DataSignatureGroup.SerializeToByteList();
        }
    }
}
