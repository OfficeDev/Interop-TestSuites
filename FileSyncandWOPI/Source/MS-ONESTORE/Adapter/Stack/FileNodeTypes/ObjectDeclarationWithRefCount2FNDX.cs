namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ObjectDeclarationWithRefCount2FNDX structure.
    /// </summary>
    public class ObjectDeclarationWithRefCount2FNDX : FileNodeBase
    {
        /// <summary>
        /// The format of stp field.
        /// </summary>
        private uint stpFormat;
        /// <summary>
        /// The format of cb field.
        /// </summary>
        private uint cbFormat;

        public ObjectDeclarationWithRefCount2FNDX(uint stpFormat, uint cbFormat)
        {
            this.stpFormat = stpFormat;
            this.cbFormat = cbFormat;
        }

        /// <summary>
        /// Gets or sets the value of ObjectRef field.
        /// </summary>
        public FileNodeChunkReference ObjectRef { get; set; }

        /// <summary>
        /// Gets or sets the value of body field. 
        /// </summary>
        public ObjectDeclarationWithRefCountBody body { get; set; }

        /// <summary>
        /// Gets or sets the value of cRef field.
        /// </summary>
        public uint cRef { get; set; }

        /// <summary>
        /// This method is used to deserialize the ObjectDeclarationWithRefCount2FNDX object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectDeclarationWithRefCount2FNDX object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.ObjectRef = new FileNodeChunkReference(this.stpFormat, this.cbFormat);
            int len = this.ObjectRef.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.body = new ObjectDeclarationWithRefCountBody();
            len = this.body.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.cRef = BitConverter.ToUInt32(byteArray, index);
            index += 4;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of ObjectDeclarationWithRefCount2FNDX object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectDeclarationWithRefCount2FNDX.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.ObjectRef.SerializeToByteList());
            byteList.AddRange(this.body.SerializeToByteList());
            byteList.AddRange(BitConverter.GetBytes(this.cRef));

            return byteList;
        }
    }
}
