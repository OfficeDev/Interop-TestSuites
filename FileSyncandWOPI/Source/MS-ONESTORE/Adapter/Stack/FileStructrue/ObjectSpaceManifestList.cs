namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// This class is used to represent the Object Space Manifest List.
    /// </summary>
    public class ObjectSpaceManifestList
    {
        public ObjectSpaceManifestList()
        {
            this.FileNodeListFragments = new List<FileNodeListFragment>();
            this.RevisionManifestList = new List<MS_ONESTORE.RevisionManifestList>();
            this.FileNodeSequence = new List<FileNode>();
        }

        /// <summary>
        /// Gets or sets the FileNodeListFragment list.
        /// </summary>
        public List<FileNodeListFragment> FileNodeListFragments { get; set; }
        /// <summary>
        /// Gets or sets the ObjectSpaceManifestListStartFND FileNode 
        /// </summary>
        public FileNode ObjectSpaceManifestListStart { get; set; }

        /// <summary>
        /// Gets or sets the value of the Revision Mamifest List.
        /// </summary>
        public List<RevisionManifestList> RevisionManifestList { get; set; }

        /// <summary>
        /// Gets or sest the sequence of FileNode in Object Group.
        /// </summary>
        public List<FileNode> FileNodeSequence { get; set; }

        /// This method is used to deserialize the Root File Node List from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="reference">Specify the reference of Root File Node List.</param>
        public void DoDeserializeFromByteArray(byte[] byteArray, FileNodeChunkReference reference)
        {
            FileNodeListFragment fragment = new FileNodeListFragment(reference.CbValue);
            fragment.DoDeserializeFromByteArray(byteArray, (int)reference.StpValue);
            this.FileNodeListFragments.Add(fragment);
            this.FileNodeSequence.AddRange(fragment.rgFileNodes.Where(f => f.FileNodeID != FileNodeIDValues.ChunkTerminatorFND).ToArray());
            FileChunkReference64x32 nextFragmentRef = fragment.nextFragment;
            while (nextFragmentRef.IsfcrNil() == false && nextFragmentRef.IsfcrZero() == false)
            {
                FileNodeListFragment nextFragment = new FileNodeListFragment(nextFragmentRef.Cb);
                nextFragment.DoDeserializeFromByteArray(byteArray, (int)nextFragmentRef.Stp);
                nextFragmentRef = nextFragment.nextFragment;
                this.FileNodeListFragments.Add(nextFragment);
                this.FileNodeSequence.AddRange(nextFragment.rgFileNodes.Where(f => f.FileNodeID != FileNodeIDValues.ChunkTerminatorFND).ToArray());
            }

            FileNode[] objectSpaceManifestListStartArray = this.FileNodeSequence.Where(obj => obj.FileNodeID == FileNodeIDValues.ObjectSpaceManifestListStartFND).ToArray();
            if (objectSpaceManifestListStartArray.Length == 1)
            {
                this.ObjectSpaceManifestListStart = objectSpaceManifestListStartArray[0];
            }
            FileNode[] revisionManifestListRefArray = this.FileNodeSequence.Where(obj => obj.FileNodeID == FileNodeIDValues.RevisionManifestListReferenceFND).ToArray();          
            foreach (FileNode revisionManifestListNode in revisionManifestListRefArray)
            {
                RevisionManifestListReferenceFND revisionManifestListReferenceFND = revisionManifestListNode.fnd as RevisionManifestListReferenceFND;
                RevisionManifestList revisionManifestList = new MS_ONESTORE.RevisionManifestList();
                revisionManifestList.DoDeserializeFromByteArray(byteArray, revisionManifestListReferenceFND.refField);
                this.RevisionManifestList.Add(revisionManifestList);
            }
        }
    }
}
