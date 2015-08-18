namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A set of ID values with a REPLGUID.
    /// </summary>
    [SerializableObjectAttribute(false, true)]
    public class REPLGUID_IDSET : IDSET
    {
        /// <summary>
        /// A GUID value that represents a REPLGUID. When combined with 
        /// all GLOBCNT values represented in the GLOBSET field, produces 
        /// a set of GIDs. The GUID values can be converted into a REPLID to 
        /// produce a set of IDs.
        /// </summary>
        [SerializableFieldAttribute(1)]
        private List<REPLGUID_IDSETElement> idsetList;

        /// <summary>
        /// Gets or sets the idsetList.
        /// </summary>
        public List<REPLGUID_IDSETElement> IDSetList
        {
            get
            {
                return this.idsetList;
            }

            set
            {
                this.idsetList = value;
            }
        }

        /// <summary>
        /// Get LongTermIds.
        /// </summary>
        /// <returns>A list of LongTermIds.</returns>
        public List<LongTermId> GetLongTermIds()
        {
            List<LongTermId> ids = new List<LongTermId>();
            if (this.idsetList != null)
            {
                for (int i = 0; i < this.idsetList.Count; i++)
                {
                    ids.AddRange(this.idsetList[i].GetLongTermIds());
                }
            }

            return ids;
        }

        /// <summary>
        /// Indicates whether contains IDSET.
        /// </summary>
        /// <param name="idset">A REPLGUID_IDSET.</param>
        /// <returns>If contains IDSET, return true, else false.</returns>
        public override bool Contains(IDSET idset)
        {
            REPLGUID_IDSET ridset = idset as REPLGUID_IDSET;
            if (this.idsetList == null
                || this.idsetList.Count == 0)
            {
                if (ridset != null)
                {
                    return false;
                }

                return true;
            }

            if (ridset != null)
            {
                foreach (REPLGUID_IDSETElement ele1 in ridset.idsetList)
                {
                    foreach (REPLGUID_IDSETElement ele2 in this.idsetList)
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

            this.idsetList = new List<REPLGUID_IDSETElement>();
            this.IsAllGLOBCNTInGLOBSET = true;
            this.IsAllGLOBCNTRanged = true;
            this.IsDisjointGLOBCNTMadeIntoSingleton = true;
            this.HasAllDuplicateGLOBCNTRemoved = true;
            this.HasGLOBCNTGroupedIntoRanges = true;
            while (stream.Position < stream.Length)
            {
                REPLGUID_IDSETElement ele = new REPLGUID_IDSETElement();
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