namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;
    using System.IO;

    /// <summary>
    /// A set of ID values with a REPLID.
    /// </summary>
    [SerializableObjectAttribute(true, true)]
    public class REPLID_IDSET : IDSET
    {
        /// <summary>
        /// A set of ID values with REPLID
        /// </summary>
        [SerializableFieldAttribute(1)]
        private List<REPLID_IDSETElement> idsetList;

        /// <summary>
        /// Get message ids.
        /// </summary>
        /// <returns>A list of message ids.</returns>
        public List<ulong> GetIds()
        {
            List<ulong> idlist = new List<ulong>();
            if (this.idsetList != null)
            {
                for (int i = 0; i < this.idsetList.Count; i++)
                {
                    idlist.AddRange(this.idsetList[i].GetIds());
                }
            }

            return idlist;
        }

        /// <summary>
        /// Indicates whether contains IDSET.
        /// </summary>
        /// <param name="idset">A REPLID_IDSET.</param>
        /// <returns>If contains IDSET, return true, else false.</returns>
        public override bool Contains(IDSET idset)
        {
            REPLID_IDSET ridset = idset as REPLID_IDSET;
            if (this.idsetList == null || this.idsetList.Count == 0)
            {
                if (ridset != null)
                {
                    return false;
                }

                return true;
            }
            else
            {
                if (ridset != null)
                {
                    foreach (REPLID_IDSETElement ele1 in ridset.idsetList)
                    {
                        foreach (REPLID_IDSETElement ele2 in this.idsetList)
                        {
                            foreach (GLOBCNTRange rng1 in ele1.GLOBSET.GLOBCNTRangeList)
                            {
                                bool hasRange = false;
                                foreach (GLOBCNTRange rng2 in ele2.GLOBSET.GLOBCNTRangeList)
                                {
                                    if (rng2.Contains(rng1))
                                    {
                                        hasRange = true;
                                        break;
                                    }
                                }

                                if (!hasRange)
                                {
                                    return false;
                                }
                            }
                        }
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Deserialize from a stream.
        /// </summary>
        /// <param name="stream">A stream contains serialize.</param>
        /// <param name="size">Must be -1.</param>
        /// <returns>The number of bytes read from the stream.</returns>
        public override int Deserialize(Stream stream, int size)
        {
            int usedBytes = 0;
            AdapterHelper.Site.Assert.AreEqual(-1, size, "The size value MUST be -1, the actual value is {0}.", size);

            this.idsetList = new List<REPLID_IDSETElement>();
            this.IsAllGLOBCNTInGLOBSET = true;
            this.IsAllGLOBCNTRanged = true;
            this.IsDisjointGLOBCNTMadeIntoSingleton = true;
            this.HasAllDuplicateGLOBCNTRemoved = true;
            this.HasGLOBCNTGroupedIntoRanges = true;
            while (stream.Position < stream.Length)
            {
                REPLID_IDSETElement ele = new REPLID_IDSETElement();
                usedBytes += ele.Deserialize(stream, -1);
                this.IsAllGLOBCNTInGLOBSET = IsAllGLOBCNTInGLOBSET
                    && ele.GLOBSET.IsAllGLOBCNTInGLOBSET;
                this.IsAllGLOBCNTRanged = IsAllGLOBCNTRanged
                    && ele.GLOBSET.IsAllGLOBCNTRanged;
                this.IsDisjointGLOBCNTMadeIntoSingleton = IsDisjointGLOBCNTMadeIntoSingleton
                    && ele.GLOBSET.IsDisjointGLOBCNTMadeIntoSingleton;
                this.HasAllDuplicateGLOBCNTRemoved = HasAllDuplicateGLOBCNTRemoved
                    && ele.GLOBSET.HasAllDuplicateGLOBCNTRemoved;
                this.HasGLOBCNTGroupedIntoRanges = HasGLOBCNTGroupedIntoRanges
                    && ele.GLOBSET.HasGLOBCNTGroupedIntoRanges;
                this.idsetList.Add(ele);
            }

            return usedBytes;
        }
    }
}