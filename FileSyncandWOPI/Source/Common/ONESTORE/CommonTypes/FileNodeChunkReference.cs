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
        /// Gets or sets the uncompressed value of the CB.
        /// </summary>
        public ulong CbValue { get; private set; }
        /// <summary>
        /// Gets or sets the uncompressed value of stp field.
        /// </summary>
        public ulong StpValue { get; private set; }

        /// <summary>
        /// This method is used to deserialize the FileNodeChunkReference object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the FileNodeChunkReference object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            int stpLen = 0;
            switch(this.stpFormat)
            {
                case 0:
                    stpLen = 8;
                    this.Stp = new byte[stpLen];
                    Array.Copy(byteArray, index, this.Stp, 0, stpLen);
                    this.StpValue = BitConverter.ToUInt64(this.Stp, 0);
                    break;
                case 1:
                    stpLen = 4;
                    this.Stp = new byte[stpLen];
                    Array.Copy(byteArray, index, this.Stp, 0, stpLen);
                    this.StpValue = (ulong)BitConverter.ToUInt32(this.Stp, 0);
                    break;
                case 3:
                    stpLen = 4;
                    this.Stp = new byte[stpLen];
                    Array.Copy(byteArray, index, this.Stp, 0, stpLen);
                    this.StpValue = (ulong)(BitConverter.ToUInt32(this.Stp, 0) * 8);
                    break;
                case 2:
                    stpLen = 2;
                    this.Stp = new byte[stpLen];
                    Array.Copy(byteArray, index, this.Stp, 0, stpLen);
                    this.StpValue = (ulong)(BitConverter.ToUInt16(this.Stp, 0) * 8);
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
                    this.Cb = new byte[4];
                    Array.Copy(byteArray, index, this.Cb, 0, 4);
                    this.CbValue = (ulong)BitConverter.ToUInt32(this.Cb, 0);
                    break;
                case 1:
                    cbLen = 8;
                    this.Cb = new byte[8];
                    Array.Copy(byteArray, index, this.Cb, 0, 8);
                    this.CbValue = (ulong)BitConverter.ToUInt64(this.Cb, 0);
                    break;
                case 2:
                    cbLen = 1;
                    this.Cb = new byte[1];
                    Array.Copy(byteArray, index, this.Cb, 0, 1);
                    this.CbValue = (ulong)(this.Cb[0] * 8);
                    break;
                case 3:
                    cbLen = 2;
                    this.Cb = new byte[2];
                    Array.Copy(byteArray, index, this.Cb, 0, 2);
                    this.CbValue = (ulong)(BitConverter.ToUInt16(this.Cb, 0) * 8);
                    break;
            }

            index += cbLen;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of FileNodeChunkReference object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of FileNodeChunkReference</returns>
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
