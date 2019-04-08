namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ObjectGroupStartFND structure.
    /// </summary>
    public class ObjectGroupStartFND:FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of oid field.
        /// </summary>
        public ExtendedGUID oid { get; set; }

        /// <summary>
        /// This method is used to deserialize the ObjectGroupStartFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectGroupStartFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            this.oid = new ExtendedGUID();
            return this.oid.DoDeserializeFromByteArray(byteArray, startIndex);
        }
        /// <summary>
        /// This method is used to convert the element of ObjectGroupStartFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectGroupStartFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            return this.oid.SerializeToByteList();
        }
    }
}
