namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;

    /// <summary>
    /// The class is used to represent the GlobalIdTableStartFNDX structure.
    /// </summary>
    public class GlobalIdTableStartFNDX : FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of Reserved field.
        /// </summary>
        public byte Reserved { get; set; }

        /// <summary>
        /// This method is used to deserialize the GlobalIdTableStartFNDX object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the GlobalIdTableStartFNDX object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            this.Reserved = byteArray[startIndex];
            return 1;
        }
        /// <summary>
        /// This method is used to convert the element of GlobalIdTableStartFNDX object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of GlobalIdTableStartFNDX.</returns>
        public override List<byte> SerializeToByteList()
        {
            return new List<byte>(this.Reserved);
        }
    }
}
