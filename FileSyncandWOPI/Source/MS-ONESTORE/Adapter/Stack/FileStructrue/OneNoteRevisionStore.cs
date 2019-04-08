namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// This class is used to represent the OneNote Revision Store file.
    /// </summary>
    public class OneNoteRevisionStoreFile
    {
        public static Dictionary<uint, uint> FileNodeCountMapping = new Dictionary<uint, uint>();

        public static bool IsEncryption = false;
        /// <summary>
        /// Gets or sets the value of header field.
        /// </summary>
        public Header Header { get; set; }

        /// <summary>
        /// Gets or sets the value of FreeChunkList field.
        /// </summary>
        public List<FreeChunkListFragment> FreeChunkList { get; set; }
        /// <summary>
        /// Gets or sets the value of TransactionLog field.
        /// </summary>
        public List<TransactionLogFragment> TransactionLog { get; set; }

        /// <summary>
        /// Gets or sets the value of HashedChunkList field.
        /// </summary>
        public List<FileNodeListFragment> HashedChunkList { get; set; }

        /// <summary>
        /// Gets or sets the value of root file node list
        /// </summary>
        public RootFileNodeList RootFileNodeList { get; set; }

        /// <summary>
        /// Gets or sets the value of file node list
        /// </summary>
        public List<FileNodeListFragment> FileNodeList { get; set; }

        public OneNoteRevisionStoreFile()
        {
            OneNoteRevisionStoreFile.FileNodeCountMapping.Clear();

            this.FreeChunkList = new List<FreeChunkListFragment>();
            this.TransactionLog = new List<TransactionLogFragment>();
            this.HashedChunkList = new List<FileNodeListFragment>();
            OneNoteRevisionStoreFile.IsEncryption = false;
            OneNoteRevisionStoreFile.FileNodeCountMapping = new Dictionary<uint, uint>();
        }
        public void DoDeserializeFromByteArray(byte[] byteArray)
        {
            int index = 0;
            this.Header = new Header();
            int len = this.Header.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.FreeChunkList = new List<FreeChunkListFragment>();
            FileChunkReference64x32 freeChunkRef = this.Header.fcrFreeChunkList;
            if (freeChunkRef.IsfcrNil() == false && freeChunkRef.IsfcrZero() == false)
            {
                do
                {
                    FreeChunkListFragment freeChunkListFragment = new FreeChunkListFragment(freeChunkRef.Cb);
                    len = freeChunkListFragment.DoDeserializeFromByteArray(byteArray, (int)freeChunkRef.Stp);
                    index += len;
                    this.FreeChunkList.Add(freeChunkListFragment);
                    freeChunkRef = freeChunkListFragment.fcrNextChunk;
                }
                while (freeChunkRef.IsfcrNil() == false && freeChunkRef.IsfcrZero() == false);
            }
            this.TransactionLog = new List<TransactionLogFragment>();
            FileChunkReference64x32 transLogRef = this.Header.fcrTransactionLog;
            do
            {
                TransactionLogFragment transLogFragment = new TransactionLogFragment(transLogRef.Cb);
                len = transLogFragment.DoDeserializeFromByteArray(byteArray, (int)transLogRef.Stp);
                transLogRef = transLogFragment.nextFragment;
                this.TransactionLog.Add(transLogFragment);
                foreach (TransactionEntry entry in transLogFragment.sizeTable.Where(t => t.srcID != 0x00000001).ToArray())
                {
                    if (FileNodeCountMapping.ContainsKey(entry.srcID))
                    {
                        if (FileNodeCountMapping[entry.srcID] < entry.TransactionEntrySwitch)
                        {
                            FileNodeCountMapping[entry.srcID] = entry.TransactionEntrySwitch;
                        }
                    }
                    else
                    {
                        FileNodeCountMapping.Add(entry.srcID, entry.TransactionEntrySwitch);
                    }
                }
            }
            while (transLogRef.IsfcrNil() == false && transLogRef.IsfcrZero() == false);

            this.HashedChunkList = new List<FileNodeListFragment>();
            FileChunkReference64x32 hashChunkRef = this.Header.fcrHashedChunkList;
            if (hashChunkRef.IsfcrNil() == false && hashChunkRef.IsfcrZero() == false)
            {
                do
                {
                    FileNodeListFragment hashChunkList = new FileNodeListFragment(hashChunkRef.Cb);
                    hashChunkList.DoDeserializeFromByteArray(byteArray, (int)hashChunkRef.Stp);
                    this.HashedChunkList.Add(hashChunkList);
                    hashChunkRef = hashChunkList.nextFragment;
                }
                while (hashChunkRef.IsfcrNil() == false && hashChunkRef.IsfcrZero() == false);
            }

            this.FileNodeList = new List<FileNodeListFragment>();
            if (Header.fcrFileNodeListRoot.IsfcrNil() == false && Header.fcrFileNodeListRoot.IsfcrZero() == false)
            {
                this.RootFileNodeList = new RootFileNodeList();
                this.RootFileNodeList.DoDeserializeFromByteArray(byteArray, Header.fcrFileNodeListRoot);
            }
        }
    }
}
