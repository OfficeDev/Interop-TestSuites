namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// The class is used to represent the Object Group List.
    /// </summary>
    public class ObjectGroupList
    {
        public ObjectGroupList()
        {
            this.FileNodeSequence = new List<FileNode>();
            this.FileNodeListFragments = new List<FileNodeListFragment>();
        }
        /// <summary>
        /// Gets or sest the sequence of FileNode in Object Group.
        /// </summary>
        public List<FileNode> FileNodeSequence { get; set; }
        /// <summary>
        /// Gets or sets the FileNodeListFragment list.
        /// </summary>
        public List<FileNodeListFragment> FileNodeListFragments { get; set; }

        /// This method is used to deserialize the Object Group from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="reference">Specify the reference of Object Group.</param>
        public void DoDeserializeFromByteArray(byte[] byteArray, FileNodeChunkReference reference)
        {
            FileNodeListFragment fragment = new FileNodeListFragment(reference.CbValue);
            fragment.DoDeserializeFromByteArray(byteArray, (int)reference.StpValue);
            this.FileNodeListFragments.Add(fragment);
            this.FileNodeSequence.AddRange(fragment.rgFileNodes.Where(f => f.FileNodeID != FileNodeIDValues.ChunkTerminatorFND).ToArray());
            FileChunkReference64x32 nextFragmentRef = fragment.nextFragment;
            while (nextFragmentRef.IsfcrNil() == false)
            {
                FileNodeListFragment nextFragment = new FileNodeListFragment(nextFragmentRef.Cb);
                nextFragment.DoDeserializeFromByteArray(byteArray, (int)nextFragmentRef.Stp);
                nextFragmentRef = nextFragment.nextFragment;
                this.FileNodeListFragments.Add(nextFragment);
                this.FileNodeSequence.AddRange(nextFragment.rgFileNodes.Where(f => f.FileNodeID != FileNodeIDValues.ChunkTerminatorFND).ToArray());
            }
        }
    }
}
