namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ObjectInfoDependencyOverridesFND structure.
    /// </summary>
    public class ObjectInfoDependencyOverridesFND : FileNodeBase
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
        public ObjectInfoDependencyOverridesFND(uint stpFormat, uint cbFormat)
        {
            this.stpFormat = stpFormat;
            this.cbFormat = cbFormat;
        }

        /// <summary>
        /// Gets or sets the value of ref field.
        /// </summary>
        public FileNodeChunkReference Ref { get; set; }

        /// <summary>
        /// Gets or sets the value of data field.
        /// </summary>
        public ObjectInfoDependencyOverrideData data { get; set; }

        /// <summary>
        /// This method is used to deserialize the ObjectInfoDependencyOverridesFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectInfoDependencyOverridesFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.Ref = new FileNodeChunkReference(this.stpFormat, this.cbFormat);
            int len = this.Ref.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.data = new ObjectInfoDependencyOverrideData();
            len = this.data.DoDeserializeFromByteArray(byteArray, index);
            index += len;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of ObjectInfoDependencyOverridesFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectInfoDependencyOverridesFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.Ref.SerializeToByteList());
            byteList.AddRange(this.data.SerializeToByteList());

            return byteList;
        }
    }
}
