namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// This class is used to represnet the RevisionManifestStart6FND structure.
    /// </summary>
    public class RevisionManifestStart6FND : FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of the rid field.
        /// </summary>
        public ExtendedGUID rid { get; set; }
        /// <summary>
        /// Gets or sets the value of ridDependent field.
        /// </summary>
        public ExtendedGUID ridDependent { get; set; }
        /// <summary>
        /// Gets or sets the value of the RevisionRole field.
        /// </summary>
        public int RevisionRole { get; set; }
        /// <summary>
        /// Gets or sets the value of the odcsDefault field.
        /// </summary>
        public ushort odcsDefault { get; set; }

        /// <summary>
        /// This method is used to deserialize the RevisionManifestStart6FND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the RevisionManifestStart6FND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.rid = new ExtendedGUID();
            int len = this.rid.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.ridDependent = new ExtendedGUID();
            len = this.ridDependent.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.RevisionRole = BitConverter.ToInt32(byteArray, index);
            index += 4;
            this.odcsDefault = BitConverter.ToUInt16(byteArray, index);
            index += 2;

            return index - startIndex;
        }
        /// <summary>
        /// This method is used to convert the element of RevisionManifestStart6FND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of RevisionManifestStart6FND.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.rid.SerializeToByteList());
            byteList.AddRange(this.ridDependent.SerializeToByteList());
            byteList.AddRange(BitConverter.GetBytes(this.RevisionRole));
            byteList.AddRange(BitConverter.GetBytes(this.odcsDefault));

            return byteList;
        }
    }
}
