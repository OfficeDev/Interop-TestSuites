
namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the FileChunkReference32 structure.
    /// </summary>
    public class FileChunkReference32 : FileChunkReference
    {
        /// <summary>
        /// Gets or sets the value of stp field.
        /// </summary>
        public uint Stp { get; set; }
        /// <summary>
        /// Gets or sets the value of cb field.
        /// </summary>
        public uint Cb { get; set; }
        /// <summary>
        /// This method is used to deserialize the FileChunkReference32 object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the FileChunkReference32 object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.Stp = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.Cb = BitConverter.ToUInt32(byteArray, index);
            index += 4;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of FileChunkReference32 object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of FileChunkReference32</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(BitConverter.GetBytes(this.Stp));
            byteList.AddRange(BitConverter.GetBytes(this.Cb));

            return byteList;
        }
        /// <summary>
        /// This method is used the check the instance whether is fcrNil.
        /// </summary>
        /// <returns>return the whether the instance is fcrNil.</returns>
        public override bool IsfcrNil()
        {
            if(this.Stp==uint.MaxValue && this.Cb==uint.MinValue)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// This method is used the check the instance whether is fcrZero.
        /// </summary>
        /// <returns>return the whether the instance is fcrZero.</returns>
        public override bool IsfcrZero()
        {
            if (this.Stp == uint.MinValue && this.Cb == uint.MinValue)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
