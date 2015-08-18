namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// A set of ID values.
    /// </summary>
    [SerializableObjectAttribute(false, false)]
    public abstract class IDSET : SerializableBase
    {
        /// <summary>
        /// Gets or sets a value indicating whether all GLOBCNTs in GLOBSET when serializing.
        /// </summary>
        public bool IsAllGLOBCNTInGLOBSET
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets a value indicating whether all Duplicate GLOBCNTs removed when serializing.
        /// </summary>
        public bool HasAllDuplicateGLOBCNTRemoved
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets a value indicating whether all GLOBCNTs are arranged from lowest to highest.
        /// </summary>
        public bool IsAllGLOBCNTRanged
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets a value indicating whether GLOBCNT values are grouped into consecutive ranges 
        /// with a low GLOBCNT value and a high GLOBCNT value.
        /// </summary>
        public bool HasGLOBCNTGroupedIntoRanges
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets a value indicating whether GLOBCNT value which is disjoint is made into a singleton range 
        /// with the low and high GLOBCNT values being the same.
        /// </summary>
        public bool IsDisjointGLOBCNTMadeIntoSingleton
        {
            get;
            set;
        }

        /// <summary>
        /// Indicates whether contains an IDSET.
        /// </summary>
        /// <param name="idset">The IDSET.</param>
        /// <returns>True ,if contains, else false.</returns>
        public abstract bool Contains(IDSET idset);
    }
}