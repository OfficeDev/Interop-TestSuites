namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the FileNodeChunkReference structure.
    /// </summary>
    public class FileNodeChunkReference:FileChunkReference
    {
        /// <summary>
        /// The format of stp field.
        /// </summary>
        private uint stpFormat;
        /// <summary>
        /// The format of cb field.
        /// </summary>
        private uint cbFormat;
        public FileNodeChunkReference(uint stpFormat,uint cbFormat)
        {
            this.stpFormat = stpFormat;
            this.cbFormat = cbFormat;
        }

        /// <summary>
        /// Gets or sets the value of stp field.
        /// </summary>
        public byte[] Stp { get; set; }
        /// <summary>
        /// Gets or sets the value of cb field.
        /// </summary>
        public byte[] Cb { get; set; }

        /// <summary>
        /// This method is used to deserialize the FileChunkReference32 object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the FileChunkReference32 object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            int stpLen = 0;
            switch(this.stpFormat)
            {
                case 0:
                    stpLen = 8;
                    break;
                case 1:
                case 3:
                    stpLen = 4;
                    break;
                case 2:
                    stpLen = 2;
                    break;
            }
            this.Stp = new byte[stpLen];
            Array.Copy(byteArray, index, this.Stp, 0, stpLen);
            index += stpLen;
            int cbLen = 0;
            switch(this.cbFormat)
            {
                case 0:
                    cbLen = 4;
                    break;
                case 1:
                    cbLen = 8;
                    break;
                case 2:
                    cbLen = 1;
                    break;
                case 3:
                    cbLen = 2;
                    break;
            }
            this.Cb = new byte[cbLen];
            Array.Copy(byteArray, index, this.Cb, 0, cbLen);
            index += cbLen;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of FileChunkReference32 object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of FileChunkReference32</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.Stp);
            byteList.AddRange(this.Cb);

            return byteList;
        }
        /// <summary>
        /// This method is used the check the instance whether is fcrNil.
        /// </summary>
        /// <returns>return the whether the instance is fcrNil.</returns>
        public override bool IsfcrNil()
        {
            foreach(byte b in this.Stp)
            {
                if(b!=byte.MaxValue)
                {
                    return false;
                }
            }
            foreach(byte b in this.Cb)
            {
                if (b != 0)
                {
                    return false;
                }
            }

            return true;
        }
        /// <summary>
        /// This method is used the check the instance whether is fcrZero.
        /// </summary>
        /// <returns>return the whether the instance is fcrZero.</returns>
        public override bool IsfcrZero()
        {
            foreach (byte b in this.Stp)
            {
                if (b != byte.MinValue)
                {
                    return false;
                }
            }
            foreach (byte b in this.Cb)
            {
                if (b != byte.MinValue)
                {
                    return false;
                }
            }

            return true;
        }
    }
}
