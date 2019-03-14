namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ObjectDataEncryptionKeyV2FNDX structure.
    /// </summary>
    public class ObjectDataEncryptionKeyV2FNDX : FileNodeBase
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
        public ObjectDataEncryptionKeyV2FNDX(uint stpFormat, uint cbFormat)
        {
            this.stpFormat = stpFormat;
            this.cbFormat = cbFormat;
        }
        /// <summary>
        /// Gets or sets the value of ref field.
        /// </summary>
        public FileNodeChunkReference Ref { get; set; }
        /// <summary>
        /// The Header field.
        /// </summary>
        public ulong Header { get; set; }

        /// <summary>
        /// The Encryption Data field.
        /// </summary>
        public byte[] EncryptionData { get; set; }

        /// <summary>
        /// The Footer field.
        /// </summary>
        public ulong Footer { get; set; }

        /// <summary>
        /// This method is used to deserialize the ObjectDataEncryptionKeyV2FNDX object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectDataEncryptionKeyV2FNDX object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            this.Ref = new FileNodeChunkReference(this.stpFormat, this.cbFormat);
            int len = this.Ref.DoDeserializeFromByteArray(byteArray, startIndex);
            int index = (int)this.Ref.StpValue;
            this.Header = BitConverter.ToUInt64(byteArray, index);
            index += 8;
            this.EncryptionData = new byte[this.Ref.CbValue - 16];
            Array.Copy(byteArray, index, this.EncryptionData, 0, this.EncryptionData.Length);
            index += this.EncryptionData.Length - 1;
            this.Footer = BitConverter.ToUInt64(byteArray, index);
      
            return len;
        }

        /// <summary>
        /// This method is used to convert the element of ObjectDataEncryptionKeyV2FNDX object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectDataEncryptionKeyV2FNDX.</returns>
        public override List<byte> SerializeToByteList()
        {
            return this.Ref.SerializeToByteList();
        }
    }
}
