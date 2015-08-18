namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System.Collections.Generic;

    /// <summary>
    /// The hierarchySync element contains the result of the hierarchy synchronization download operation.
    /// hierarchySync        = *folderChange 
    ///                 [deletions] 
    ///                 state 
    ///                 IncrSyncEnd
    /// </summary>
    public class HierarchySync : SyntacticalBase
    {
        /// <summary>
        /// The end marker of hierarchySync.
        /// </summary>
        public const Markers EndMarker = Markers.PidTagIncrSyncEnd;

        /// <summary>
        /// A list of folderChange value.
        /// </summary>
        private List<FolderChange> folderChangeList;

        /// <summary>
        /// A deletions value.
        /// </summary>
        private Deletions deletions;

        /// <summary>
        /// The ICS state.
        /// </summary>
        private State state;

        /// <summary>
        /// Initializes a new instance of the HierarchySync class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public HierarchySync(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets folderChangeList.
        /// </summary>
        public List<FolderChange> FolderChangeList
        {
            get { return this.folderChangeList; }
            set { this.folderChangeList = value; }
        }

        /// <summary>
        /// Gets or sets deletions.
        /// </summary>
        public Deletions Deletions
        {
            get { return this.deletions; }
            set { this.deletions = value; }
        }

        /// <summary>
        /// Gets or sets state.
        /// </summary>
        public State State
        {
            get { return this.state; }
            set { this.state = value; }
        }

        /// <summary>
        /// Gets a value indicating whether parents are before their children.
        /// </summary>
        public bool IsParentBeforeChildren
        {
            get
            {
                // If parents are before their children,
                // for a specific folder change, its ParentSourceKey
                // must be zero if its parent is the sync folder
                // or can be found before the position of the folder change.
                if (this.FolderChangeList != null && this.FolderChangeList.Count > 0)
                {
                    if (this.FolderChangeList.Count > 1)
                    {
                        int i = 0;
                        for (i = 0; i < this.FolderChangeList.Count; i++)
                        {
                            byte[] parent = this.FolderChangeList[i].GetParentSourceKey();
                            if (!this.IsZeroSourceKey(parent))
                            {
                                break;
                            }
                        }

                        if (i == this.FolderChangeList.Count)
                        {
                            return false;
                        }
                    }

                    if (this.FolderChangeList.Count > 1)
                    {
                        for (int i = this.FolderChangeList.Count - 1; i > 0; i--)
                        {
                            byte[] parent = this.FolderChangeList[i].GetParentSourceKey();
                            if (this.IsZeroSourceKey(parent))
                            {
                                continue;
                            }
                            else
                            {
                                bool hasParent = false;
                                for (int j = 0; j < i; j++)
                                {
                                    if (this.IsBufferEqual(parent, this.FolderChangeList[j].GetSourceKey()))
                                    {
                                        hasParent = true;
                                        break;
                                    }
                                }

                                if (!hasParent)
                                {
                                    return false;
                                }
                            }
                        }

                        return true;
                    }

                    return false;
                }

                return false;
            }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized hierarchySync.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized hierarchySync, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return (FolderChange.Verify(stream)
                || Deletions.Verify(stream)
                || State.Verify(stream))
                && stream.VerifyMarker(
                        Markers.PidTagIncrSyncEnd,
                        (int)stream.Length - MarkersHelper.PidTagLength - (int)stream.Position);
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            this.folderChangeList = new List<FolderChange>();
            while (FolderChange.Verify(stream))
            {
                this.folderChangeList.Add(new FolderChange(stream));
            }

            if (Deletions.Verify(stream))
            {
                this.deletions = new Deletions(stream);
            }

            this.state = new State(stream);
            if (!stream.ReadMarker(Markers.PidTagIncrSyncEnd))
            {
                AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
            }
        }

        /// <summary>
        /// Indicates whether two byte arrays are equal byte by byte.
        /// </summary>
        /// <param name="b1">The first byte array to compare.</param>
        /// <param name="b2">The second byte array to compare.</param>
        /// <returns>True, if two byte arrays have same length, are equal byte by byte.
        /// Else false.
        /// </returns>
        private bool IsBufferEqual(byte[] b1, byte[] b2)
        {
            if (b1.Length != b2.Length)
            {
                return false;
            }

            for (int i = 0; i < b1.Length; i++)
            {
                if (b1[i] != b2[i])
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Indicates whether a byte array is empty or filled with 0.
        /// </summary>
        /// <param name="b">The byte array.</param>
        /// <returns>Whether the byte array is empty or filled with 0.</returns>
        private bool IsZeroSourceKey(byte[] b)
        {
            for (int i = 0; i < b.Length; i++)
            {
                if (b[i] != 0)
                {
                    return false;
                }
            }

            return true;
        }
    }
}