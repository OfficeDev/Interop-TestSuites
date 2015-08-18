namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// The FolderId field
    /// </summary>
    public class FolderId
    {
        /// <summary>
        /// Initializes a new instance of the FolderId class.
        /// </summary>
        /// <param name="value">The folder Id value</param>
        public FolderId(ulong value)
        {
            byte[] byteArray = BitConverter.GetBytes(value);
            this.ReplicaId = new byte[2];
            Array.Copy(byteArray, 6, this.ReplicaId, 0, 2);
            this.GlobalCounter = new byte[6];
            Array.Copy(byteArray, 0, this.GlobalCounter, 0, 6);
        }

        /// <summary>
        /// Gets or sets the ReplicaId field
        /// </summary>
        public byte[] ReplicaId
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the GlobalCounter field
        /// </summary>
        public byte[] GlobalCounter
        {
            get;
            set;
        }
    }
}