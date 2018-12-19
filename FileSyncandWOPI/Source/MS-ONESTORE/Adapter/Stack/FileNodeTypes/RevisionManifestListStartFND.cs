namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    public class RevisionManifestListStartFND : FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of gosid field.
        /// </summary>
        public ExtendedGUID gosid { get; set; }
        /// <summary>
        /// Gets or sets the value of nInstance field.
        /// </summary>
        public byte[] nInstance { get; set; }
        /// <summary>
        /// This method is used to deserialize the RevisionManifestListStartFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the RevisionManifestListStartFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.gosid = new ExtendedGUID();
            int len = this.gosid.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.nInstance = new byte[4];
            Array.Copy(byteArray, index, this.nInstance, 0, 4);
            index += 4;

            return index - startIndex;
        }
        /// <summary>
        /// This method is used to convert the element of RevisionManifestListStartFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of RevisionManifestListStartFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.gosid.SerializeToByteList());
            byteList.AddRange(this.nInstance);

            return byteList;
        }
    }
}
