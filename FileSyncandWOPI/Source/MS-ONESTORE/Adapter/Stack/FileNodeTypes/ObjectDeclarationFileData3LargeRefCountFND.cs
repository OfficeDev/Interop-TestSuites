namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ObjectDeclarationFileData3LargeRefCountFND structure.
    /// </summary>
    public class ObjectDeclarationFileData3LargeRefCountFND : FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of oid field.
        /// </summary>
        public CompactID oid { get; set; }

        /// <summary>
        /// Gets or sets the value of jcid field.
        /// </summary>
        public JCID jcid { get; set; }

        /// <summary>
        /// Gets or sets the value of cRef field.
        /// </summary>
        public uint cRef { get; set; }

        /// <summary>
        /// Gets or sets the value of FileDataReference field.
        /// </summary>
        public StringInStorageBuffer FileDataReference { get; set; }

        /// <summary>
        /// Gets or sets the value of Extension field.
        /// </summary>
        public StringInStorageBuffer Extension { get; set; }

        /// <summary>
        /// This method is used to deserialize the ObjectDeclarationFileData3LargeRefCountFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectDeclarationFileData3LargeRefCountFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.oid = new CompactID();
            int len = this.oid.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.jcid = new JCID();
            len = this.jcid.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.cRef = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.FileDataReference = new StringInStorageBuffer();
            len = this.FileDataReference.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.Extension = new StringInStorageBuffer();
            len = this.Extension.DoDeserializeFromByteArray(byteArray, index);
            index += len;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of ObjectDeclarationFileData3LargeRefCountFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectDeclarationFileData3LargeRefCountFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.oid.SerializeToByteList());
            byteList.AddRange(this.jcid.SerializeToByteList());
            byteList.AddRange(BitConverter.GetBytes(this.cRef));
            byteList.AddRange(this.FileDataReference.SerializeToByteList());
            byteList.AddRange(this.Extension.SerializeToByteList());

            return byteList;
        }
    }
}
