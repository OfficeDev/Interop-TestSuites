namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the FileDataStoreObjectReferenceFND structure.
    /// </summary>
    public class FileDataStoreObjectReferenceFND:FileNodeBase
    {
        /// <summary>
        /// The format of stp field.
        /// </summary>
        private uint stpFormat;
        /// <summary>
        /// The format of cb field.
        /// </summary>
        private uint cbFormat;

        public FileDataStoreObjectReferenceFND(uint stpFormat, uint cbFormat)
        {
            this.stpFormat = stpFormat;
            this.cbFormat = cbFormat;
        }

        /// <summary>
        /// Gets or sets the value of ref field.
        /// </summary>
        public FileNodeChunkReference Ref { get; set; }

        /// <summary>
        /// Gets or sets the value of guidReference field.
        /// </summary>
        public Guid guidReference { get; set; }

        public FileDataStoreObject fileDataStoreObject { get; set; }

        /// <summary>
        /// This method is used to deserialize the FileDataStoreObjectReferenceFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the FileDataStoreObjectReferenceFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.Ref = new FileNodeChunkReference(this.stpFormat, this.cbFormat);
            int len = this.Ref.DoDeserializeFromByteArray(byteArray, index);
            index += len;

            this.guidReference = AdapterHelper.ReadGuid(byteArray, index);
            index += 16;

            this.fileDataStoreObject = new FileDataStoreObject((uint)this.Ref.CbValue);
            this.fileDataStoreObject.DoDeserializeFromByteArray(byteArray, (int)this.Ref.StpValue);

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of FileDataStoreObjectReferenceFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of FileDataStoreObjectReferenceFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.Ref.SerializeToByteList());
            byteList.AddRange(this.guidReference.ToByteArray());

            return byteList;
        }
    }
}
