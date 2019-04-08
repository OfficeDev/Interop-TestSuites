namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ReadOnlyObjectDeclaration2LargeRefCountFND structure.
    /// </summary>
    public class ReadOnlyObjectDeclaration2LargeRefCountFND:FileNodeBase
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
        public ReadOnlyObjectDeclaration2LargeRefCountFND(uint stpFormat, uint cbFormat)
        {
            this.stpFormat = stpFormat;
            this.cbFormat = cbFormat;
        }

        /// <summary>
        /// Gets or sets the value of base field.
        /// </summary>
        public ObjectDeclaration2LargeRefCountFND Base { get; set; }

        /// <summary>
        /// Gets or sets the value of md5Hash field.
        /// </summary>
        public byte[] md5Hash { get; set; }

        /// <summary>
        /// This method is used to deserialize the ReadOnlyObjectDeclaration2LargeRefCountFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ReadOnlyObjectDeclaration2LargeRefCountFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.Base = new ObjectDeclaration2LargeRefCountFND(this.stpFormat, this.cbFormat);
            int len = this.Base.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.md5Hash = new byte[16];
            Array.Copy(byteArray, index, this.md5Hash, 0, 16);
            index += 16;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of ReadOnlyObjectDeclaration2LargeRefCountFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ReadOnlyObjectDeclaration2LargeRefCountFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.Base.SerializeToByteList());
            byteList.AddRange(this.md5Hash);

            return byteList;
        }
    }
}
