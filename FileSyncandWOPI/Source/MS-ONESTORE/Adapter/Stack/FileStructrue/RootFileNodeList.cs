namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// The class is used to represent the Root File Node List.
    /// </summary>
    public class RootFileNodeList
    {
        public RootFileNodeList()
        {
            this.FileNodeListFragments = new List<FileNodeListFragment>();
            this.ObjectSpaceManifestList = new List<MS_ONESTORE.ObjectSpaceManifestList>();
            this.FileDataStoreListReference = new List<FileNode>();
            this.FileNodeSequence = new List<FileNode>();
        }
        /// <summary>
        /// Gets or sets the FileNodeListFragment list.
        /// </summary>
        public List<FileNodeListFragment> FileNodeListFragments { get; set; }
        /// <summary>
        /// Gets or sets the ObjectSpaceManifestRoot FileNode.
        /// </summary>
        public FileNode ObjectSpaceManifestRoot { get; set; }

        /// <summary>
        /// Gets or sets the FileNode list.
        /// </summary>
        public List<FileNode> FileNodeSequence { get; set; }
        /// <summary>
        /// Gets or sets the list of Object SpaceManifest List
        /// </summary>
        public List<ObjectSpaceManifestList> ObjectSpaceManifestList { get; set; }
        /// <summary>
        /// Gets or sets the list of FileDataStoreListReference FileNode.
        /// </summary>
        public List<FileNode> FileDataStoreListReference { get; set; }
        /// <summary>
        /// This method is used to deserialize the Root File Node List from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="reference">Specify the reference of Root File Node List.</param>
        public void DoDeserializeFromByteArray(byte[] byteArray, FileChunkReference64x32 reference)
        {
            FileNodeListFragment fragment = new FileNodeListFragment(reference.Cb);
            fragment.DoDeserializeFromByteArray(byteArray, (int)reference.Stp);
            this.FileNodeListFragments.Add(fragment);
            this.FileNodeSequence.AddRange(fragment.rgFileNodes.Where(f => f.FileNodeID != FileNodeIDValues.ChunkTerminatorFND).ToArray());
            FileChunkReference64x32 nextFragmentRef = fragment.nextFragment;
            while (nextFragmentRef.IsfcrNil()==false && nextFragmentRef.IsfcrZero() == false)
            {
                FileNodeListFragment nextFragment = new FileNodeListFragment(nextFragmentRef.Cb);
                nextFragment.DoDeserializeFromByteArray(byteArray, (int)nextFragmentRef.Stp);
                nextFragmentRef = nextFragment.nextFragment;
                this.FileNodeListFragments.Add(nextFragment);
                this.FileNodeSequence.AddRange(nextFragment.rgFileNodes.Where(f => f.FileNodeID != FileNodeIDValues.ChunkTerminatorFND).ToArray());
            }

            FileNode[] objectSpaceManifestListReferences = this.FileNodeSequence.Where(obj => obj.FileNodeID == FileNodeIDValues.ObjectSpaceManifestListReferenceFND).ToArray();
            foreach (FileNode node in objectSpaceManifestListReferences)
            {
                ObjectSpaceManifestListReferenceFND objectSpaceManifestListReference = node.fnd as ObjectSpaceManifestListReferenceFND;
                ObjectSpaceManifestList objectSpaceManifestList = new MS_ONESTORE.ObjectSpaceManifestList();
                objectSpaceManifestList.DoDeserializeFromByteArray(byteArray, objectSpaceManifestListReference.refField);
                this.ObjectSpaceManifestList.Add(objectSpaceManifestList);
            }

            this.FileDataStoreListReference.AddRange(this.FileNodeSequence.Where(obj => obj.FileNodeID == FileNodeIDValues.FileDataStoreListReferenceFND).ToArray());
            FileNode[] objectSpaceManifestRootArray = this.FileNodeSequence.Where(obj => obj.FileNodeID == FileNodeIDValues.ObjectSpaceManifestRootFND).ToArray();
            if (objectSpaceManifestRootArray.Length == 1)
            {
                this.ObjectSpaceManifestRoot = objectSpaceManifestRootArray[0];
            }
        }
    }
}
