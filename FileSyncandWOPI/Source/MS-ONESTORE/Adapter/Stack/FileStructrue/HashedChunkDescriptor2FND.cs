namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// The class is used to represent HashedChunkDescriptor2FND structure.
    /// </summary>
    public class HashedChunkDescriptor2FND:FileNodeBase
    {
        /// <summary>
        /// The format of stp field.
        /// </summary>
        private uint stpFormat;
        /// <summary>
        /// The format of cb field.
        /// </summary>
        private uint cbFormat;
        public HashedChunkDescriptor2FND(uint stpFormat, uint cbFormat)
        {
            this.stpFormat = stpFormat;
            this.cbFormat = cbFormat;
        }

        /// <summary>
        /// Gets or sets the value of BlobRef field.
        /// </summary>
        public FileNodeChunkReference BlobRef { get; set; }
        /// <summary>
        /// Gets or sets the value of guidHash field.
        /// </summary>
        public Guid guidHash { get; set; }

        /// <summary>
        /// This method is used to convert the element of HashedChunkDescriptor2FND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of HashedChunkDescriptor2FND</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.BlobRef.SerializeToByteList());
            byteList.AddRange(this.guidHash.ToByteArray());

            return byteList;
        }

        /// <summary>
        /// This method is used to deserialize the HashedChunkDescriptor2FND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the HashedChunkDescriptor2FND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.BlobRef = new FileNodeChunkReference(this.stpFormat, this.cbFormat);
            int len = this.BlobRef.DoDeserializeFromByteArray(byteArray, index);
            index += len;

            this.guidHash = AdapterHelper.ReadGuid(byteArray, index);
            index += 16;

            return index - startIndex;
        }
    }
}
