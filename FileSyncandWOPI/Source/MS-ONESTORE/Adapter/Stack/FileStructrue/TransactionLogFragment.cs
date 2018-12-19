namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;


    /// <summary>
    /// This class is used to represent the TransactionLogFragment structure.
    /// </summary>
    public class TransactionLogFragment
    {
        private uint size = 0;

        public TransactionLogFragment(uint size)
        {
            this.size = size;
        }
        /// <summary>
        /// Gets or sets the value of sizeTable field.
        /// </summary>
        public TransactionEntry[] sizeTable { get; set; }
        /// <summary>
        /// Gets or sets the value of nextFragment field.
        /// </summary>
        public FileChunkReference64x32 nextFragment { get; set; }

        /// <summary>
        /// This method is used to convert the element of TransactionLogFragment object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of TransactionLogFragment</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();

            foreach(TransactionEntry entry in this.sizeTable)
            {
                byteList.AddRange(entry.SerializeToByteList());
            }

            byteList.AddRange(this.nextFragment.SerializeToByteList());

            return byteList;
        }

        /// <summary>
        /// This method is used to deserialize the TransactionLogFragment object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the TransactionLogFragment object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            uint count = ((this.size - 12) / 8);
            this.sizeTable = new TransactionEntry[count];
            int len = 0;
            for (int i = 0; i < count; i++)
            {
                this.sizeTable[i] = new TransactionEntry();
                len = this.sizeTable[i].DoDeserializeFromByteArray(byteArray, index);
                index += len;
            }

            this.nextFragment = new FileChunkReference64x32();
            len = this.nextFragment.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            return index - startIndex;
        }
    }
}
