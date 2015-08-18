namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Diagnostics;
    using System.Reflection;

    /// <summary>
    /// Base class for lexical objects
    /// </summary>
    public abstract class LexicalBase : IStreamSerializable, IStreamDeserializable
    {
        /// <summary>
        /// The length of GUID structure.
        /// </summary>
        public static readonly int GuidLength = Guid.Empty.ToByteArray().Length;

        /// <summary>
        /// Indicates whether a stream MUST NOT be split within a single atom.
        /// </summary>
        private bool isNotSplitedInSingleItem;

        /// <summary>
        /// Initializes a new instance of the LexicalBase class.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        protected LexicalBase(FastTransferStream stream)
        {
            this.Deserialize(stream);

            // No exception set flag
            this.isNotSplitedInSingleItem = true;
        }

        /// <summary>
        /// Gets or sets a value indicating whether a stream MUST NOT be split within a single atom.
        /// </summary>
        protected bool IsNotSplitedInSingleItem
        {
            get { return this.isNotSplitedInSingleItem; }
            set { this.isNotSplitedInSingleItem = value; }
        }

        /// <summary>
        /// Deserialize a LexicalBase instance from a FastTransferStream.
        /// </summary>
        /// <typeparam name="T">A subclass of LexicalBase</typeparam>
        /// <param name="stream">A FastTransferStream</param>
        /// <returns>The deserialized lexical object</returns>
        public static T DeserializeTo<T>(FastTransferStream stream)
            where T : LexicalBase
        {
            Type subType = typeof(T);
            object tmp = subType.Assembly.CreateInstance(
                    subType.FullName,
                    false,
                    BindingFlags.CreateInstance,
                    null,
                    new object[] { stream },
                    null,
                    null);
            T t = tmp as T;
            Debug.Assert(t != null, "Assure deserialize success.");
            return t;
        }

        /// <summary>
        /// The method NOT USED
        /// </summary>
        /// <returns>A FastTransferStream</returns>
        public virtual FastTransferStream Serialize()
        {
            AdapterHelper.Site.Assert.Fail("This method is not implemented.");
            return null;
        }

        /// <summary>
        /// Deserialize from a FastTransferStream
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        public virtual void Deserialize(FastTransferStream stream)
        {
            this.ConsumeNext(stream);
        }

        /// <summary>
        /// Deserialize next object from a FastTransferStream
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        public virtual void ConsumeNext(FastTransferStream stream)
        { 
        }
    }
}