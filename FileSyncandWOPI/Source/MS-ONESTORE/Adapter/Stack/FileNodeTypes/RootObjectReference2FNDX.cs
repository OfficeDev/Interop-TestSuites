namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent RootObjectReference2FNDX structure.
    /// </summary>
    public class RootObjectReference2FNDX : FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of oidRoot field.
        /// </summary>
        public CompactID oidRoot { get; set; }

        /// <summary>
        /// Gets or sets the value of RootRole field.
        /// </summary>
        public uint RootRole { get; set; }

        /// <summary>
        /// This method is used to deserialize the RootObjectReference2FNDX object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the RootObjectReference2FNDX object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.oidRoot = new CompactID();
            int len = this.oidRoot.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.RootRole = BitConverter.ToUInt32(byteArray, index);
            index += 4;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of RootObjectReference2FNDX object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of RootObjectReference2FNDX.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.oidRoot.SerializeToByteList());
            byteList.AddRange(BitConverter.GetBytes(this.RootRole));

            return byteList;
        }
    }
}
