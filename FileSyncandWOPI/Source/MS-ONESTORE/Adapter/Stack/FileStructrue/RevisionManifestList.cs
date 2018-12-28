namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// This class is used to represent the Revision Manifest List.
    /// </summary>
    public class RevisionManifestList
    {
        /// <summary>
        /// Gets or sets the FileNodeListFragment list.
        /// </summary>
        public List<FileNodeListFragment> FileNodeListFragments { get; set; }
        /// <summary>
        /// Gets or sets the Revision Manifest list.
        /// </summary>
        public List<RevisionManifest> RevisionManifests { get; set; }

        /// <summary>
        /// Gets or sets the RevisionRoleDeclarationFND FileNode list.
        /// </summary>
        public List<FileNode> RevisionRoleDeclaration { get; set; }

        /// <summary>
        /// Gets or sets the RevisionRoleAndContextDeclarationFND FileNode list.
        /// </summary>
        public List<FileNode> RevisionRoleAndContextDeclaration { get; set; }

        /// <summary>
        /// Gets or sets the Object Group List.
        /// </summary>
        public List<ObjectGroupList> ObjectGroupList { get; set; }
        /// <summary>
        /// Gets or sest the sequence of FileNode in Object Group.
        /// </summary>
        public List<FileNode> FileNodeSequence { get; set; }

        public RevisionManifestList()
        {
            this.FileNodeListFragments = new List<MS_ONESTORE.FileNodeListFragment>();
            this.RevisionManifests = new List<RevisionManifest>();
            this.RevisionRoleDeclaration = new List<FileNode>();
            this.RevisionRoleAndContextDeclaration = new List<FileNode>();
            this.ObjectGroupList = new List<MS_ONESTORE.ObjectGroupList>();
            this.FileNodeSequence = new List<FileNode>();
        }

        /// This method is used to deserialize the Revision Manifest List from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="reference">Specify the reference of Revision Manifest List.</param>
        public void DoDeserializeFromByteArray(byte[] byteArray, FileNodeChunkReference reference)
        {
            FileNodeListFragment fragment = new FileNodeListFragment(reference.CbValue);
            fragment.DoDeserializeFromByteArray(byteArray, (int)reference.StpValue);
            this.FileNodeListFragments.Add(fragment);
            this.FileNodeSequence.AddRange(fragment.rgFileNodes.Where(f=>f.FileNodeID!=FileNodeIDValues.ChunkTerminatorFND).ToArray());
            FileChunkReference64x32 nextFragmentRef = fragment.nextFragment;
            while (nextFragmentRef.IsfcrNil() == false)
            {
                FileNodeListFragment nextFragment = new FileNodeListFragment(nextFragmentRef.Cb);
                nextFragment.DoDeserializeFromByteArray(byteArray, (int)nextFragmentRef.Stp);
                nextFragmentRef = nextFragment.nextFragment;
                this.FileNodeListFragments.Add(nextFragment);
                this.FileNodeSequence.AddRange(nextFragment.rgFileNodes.Where(f => f.FileNodeID != FileNodeIDValues.ChunkTerminatorFND).ToArray());
            }

            RevisionManifest revisionManifest = null;
            foreach(FileNode fileNode in this.FileNodeSequence)
            {
                if(fileNode.FileNodeID==FileNodeIDValues.RevisionRoleDeclarationFND)
                {
                    this.RevisionRoleDeclaration.Add(fileNode);
                    continue;
                }
                if(fileNode.FileNodeID== FileNodeIDValues.RevisionRoleAndContextDeclarationFND)
                {
                    this.RevisionRoleAndContextDeclaration.Add(fileNode);
                    continue;
                }
                if (fileNode.FileNodeID == FileNodeIDValues.RevisionManifestStart6FND ||
                   fileNode.FileNodeID == FileNodeIDValues.RevisionManifestStart7FND ||
                   fileNode.FileNodeID == FileNodeIDValues.RevisionManifestStart4FND)
                {
                    revisionManifest = new RevisionManifest();
                    revisionManifest.FileNodeSequence.Add(fileNode);
                }
                else if (fileNode.FileNodeID == FileNodeIDValues.RevisionManifestEndFND)
                {
                    revisionManifest.FileNodeSequence.Add(fileNode);
                    this.RevisionManifests.Add(revisionManifest);
                    revisionManifest = new RevisionManifest();
                }
                else if (fileNode.FileNodeID != FileNodeIDValues.ChunkTerminatorFND &&
                    fileNode.FileNodeID != FileNodeIDValues.RevisionManifestListStartFND)
                {
                    revisionManifest.FileNodeSequence.Add(fileNode);
                    if (fileNode.FileNodeID == FileNodeIDValues.ObjectGroupListReferenceFND)
                    {
                        ObjectGroupListReferenceFND objectGroupListRef = fileNode.fnd as ObjectGroupListReferenceFND;
                        ObjectGroupList objectGroupList = new ObjectGroupList();
                        objectGroupList.DoDeserializeFromByteArray(byteArray, objectGroupListRef.Ref);
                        this.ObjectGroupList.Add(objectGroupList);
                    }
                }
            } 
        }
    }

    /// <summary>
    /// This class is used to represent the Revision Manifest.
    /// </summary>
    public class RevisionManifest
    {
        public RevisionManifest()
        {
            this.FileNodeSequence = new List<FileNode>();
        }

        /// <summary>
        /// Gets or sets the sequence of FileNode in Revision Manifest.
        /// </summary>
        public List<FileNode> FileNodeSequence { get; set; }
    }
}
