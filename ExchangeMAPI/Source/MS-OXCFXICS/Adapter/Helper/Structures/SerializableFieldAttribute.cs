namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// An Attribute specifies a serializable field. 
    /// </summary>
    [AttributeUsage(AttributeTargets.Field |
            AttributeTargets.Property |
            AttributeTargets.Enum |
            AttributeTargets.Struct)]
    public sealed class SerializableFieldAttribute : Attribute
    {
        /// <summary>
        /// Specifies the order of the assigned field while serializing.
        /// </summary>
        private int order;

        /// <summary>
        /// NOT USED, always be -1, specifies the minimum byte count while serializing.
        /// </summary>
        private int minAllocSize;

        /// <summary>
        /// NOT USED, always false, specifies whether to loop filling 
        /// </summary>
        private bool circleFill;

        /// <summary>
        /// The max allocate size.
        /// </summary>
        private int maxAllocSize;

        /// <summary>
        /// NOT USED, indicates the byte order while serializing, always true.
        /// </summary>
        private bool littleEndian;

        /// <summary>
        /// Initializes a new instance of the SerializableFieldAttribute class.
        /// </summary>
        /// <param name="order">The order of the field in the owner's 
        /// serialization or deserialization.</param>
        public SerializableFieldAttribute(int order)
        {
            this.order = order;
            this.littleEndian = true;
            this.minAllocSize = -1;
            this.maxAllocSize = -1;
            this.circleFill = false;
        }

        /// <summary>
        /// Gets a value indicating whether to loop filling.
        /// NOT USED
        /// </summary>
        public bool CircleFill
        {
            get { return this.circleFill; }
        }

        /// <summary>
        /// Gets the minAllocSize.
        /// NOT USED, always be -1, specifies the minimum byte count while serializing.
        /// </summary>
        public int MinAllocSize
        {
            get { return this.minAllocSize; }
        }

        /// <summary>
        /// Gets the maxAllocSize.
        /// NOT USED, always be -1, specifies the maximum byte count while serializing.
        /// </summary>
        public int MaxAllocSize
        {
            get { return this.maxAllocSize; }
        }

        /// <summary>
        /// Gets a value indicating whether the byte order while serializing, always true.
        /// NOT USED
        /// </summary>
        public bool LittleEndian
        {
            get { return this.littleEndian; }
        }

        /// <summary>
        /// Gets the order of the field in the owner's serialization or deserialization.
        /// </summary>
        public int Order
        {
            get
            {
                return this.order;
            }
        }
    }
}