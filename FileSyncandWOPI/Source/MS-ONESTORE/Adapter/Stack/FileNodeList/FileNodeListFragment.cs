namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// This class is used to represent the FileNodeListFragment structure.
    /// </summary>
    public class FileNodeListFragment
    {
        /// <summary>
        /// The size of the FileNodeListFragment structure.
        /// </summary>
        private ulong size;
        public FileNodeListFragment(ulong size)
        {
            this.size = size;
        }

        /// <summary>
        /// Gets or sets the value of header field.
        /// </summary>
        public FileNodeListHeader Header { get; set; }

        /// <summary>
        /// Gets or sets the value of rgFileNodes field.
        /// </summary>
        public List<FileNode> rgFileNodes { get; set; }

        /// <summary>
        /// Gets or sets the value of padding field.
        /// </summary>
        public byte[] padding { get; set; }

        /// <summary>
        /// Gets or sets the value of nextFragment field.
        /// </summary>
        public FileChunkReference64x32 nextFragment { get; set; }

        /// <summary>
        /// Gets or sets the value of footer field.
        /// </summary>
        public UInt64 footer { get; set; }

        /// <summary>
        /// This method is used to deserialize the FileNodeListHeader object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the FileNodeListHeader object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.Header = new FileNodeListHeader();
            int len = this.Header.DoDeserializeFromByteArray(byteArray, index);
            index += len;

            FileNode fileNode = null;
            this.rgFileNodes = new List<FileNode>();
            int fileNodeSize = 0;
            uint fileNodeCount = uint.MaxValue;
            if (OneNoteRevisionStoreFile.FileNodeCountMapping.ContainsKey(this.Header.FileNodeListID))
            {
                fileNodeCount = OneNoteRevisionStoreFile.FileNodeCountMapping[this.Header.FileNodeListID];
            }

            do
            {
                fileNode = new FileNode();
                len = fileNode.DoDeserializeFromByteArray(byteArray, index);
                index += len;
                fileNodeSize += len;
                if (fileNode.FileNodeID != 0)
                {
                    this.rgFileNodes.Add(fileNode);
                    if (fileNode.FileNodeID != FileNodeIDValues.ChunkTerminatorFND)
                    {
                        fileNodeCount--;
                    }
                }
            }
            while ((int)this.size - 36 - fileNodeSize > 4 && fileNodeCount > 0);

            if (OneNoteRevisionStoreFile.FileNodeCountMapping.ContainsKey(this.Header.FileNodeListID))
            {
                OneNoteRevisionStoreFile.FileNodeCountMapping[this.Header.FileNodeListID] = fileNodeCount;
            }

            int paddinglength = (int)this.size - 36 - fileNodeSize;
            this.padding = new byte[paddinglength];
            Array.Copy(byteArray, index, this.padding, 0, paddinglength);
            index += paddinglength;
            this.nextFragment = new FileChunkReference64x32();
            len = this.nextFragment.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.footer = BitConverter.ToUInt64(byteArray, index);
            index += 8;

            return index - startIndex;
        }
        /// <summary>
        /// This method is used to convert the element of FileNodeListHeader object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of FileNodeListHeader.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.Header.SerializeToByteList());
            foreach (FileNode fileNode in this.rgFileNodes)
            {
                byteList.AddRange(fileNode.SerializeToByteList());
            }
            byteList.AddRange(this.padding);
            byteList.AddRange(this.nextFragment.SerializeToByteList());
            byteList.AddRange(BitConverter.GetBytes(this.footer));

            return byteList;
        }
    }
}
