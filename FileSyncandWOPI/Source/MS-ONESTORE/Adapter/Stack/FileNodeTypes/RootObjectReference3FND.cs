namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the RootObjectReference3FND structure.
    /// </summary>
    public class RootObjectReference3FND : FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of oidRoot field.
        /// </summary>
        public ExtendedGUID oidRoot { get; set; }

        /// <summary>
        /// Gets or sets the value of RootRole field.
        /// </summary>
        public uint RootRole { get; set; }

        /// <summary>
        /// This method is used to deserialize the RootObjectReference3FND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the RootObjectReference3FND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.oidRoot = new ExtendedGUID();
            int len = this.oidRoot.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.RootRole = BitConverter.ToUInt32(byteArray, index);
            index += 4;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of RootObjectReference3FND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of RootObjectReference3FND.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.oidRoot.SerializeToByteList());
            byteList.AddRange(BitConverter.GetBytes(this.RootRole));

            return byteList;
        }
    }
}
