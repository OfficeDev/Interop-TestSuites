namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;
    using System.IO;

    /// <summary>
    /// Contains a set of XIDs that represent change numbers of 
    /// messaging objects in different replicas. 
    /// </summary>
    [SerializableObjectAttribute(false, true)]
    public class PredecessorChangeList : SerializableBase
    {
        /// <summary>
        /// A SizedXid list.
        /// </summary>
        [SerializableFieldAttribute(3)]
        private List<SizedXid> sizedXidList;

        /// <summary>
        /// Gets or sets the sizedXidList.
        /// </summary>
        public List<SizedXid> SizedXidList
        {
            get
            {
                return this.sizedXidList;
            }

            set
            {
                this.sizedXidList = value;
            }
        }

        /// <summary>
        /// Gets a value indicating whether XIDs with the same NamespaceGuid MUST have the same length of LocalId fields.
        /// </summary>
        public bool IsSameGuidHasSameLengthXID
        {
            get
            {
                if (this.sizedXidList != null && this.sizedXidList.Count > 0)
                {
                    Dictionary<Guid, int> guidDict = new Dictionary<Guid, int>();
                    foreach (SizedXid xid in this.sizedXidList)
                    {
                        if (guidDict.ContainsKey(xid.XID.NamespaceGuid))
                        {
                            if (xid.XID.LocalId.Length != guidDict[xid.XID.NamespaceGuid])
                            {
                                return false;
                            }
                        }
                        else
                        {
                            guidDict.Add(xid.XID.NamespaceGuid, xid.XID.LocalId.Length);
                        }
                    }
                }

                return true;
            }
        }

        /// <summary>
        /// Deserialize fields in this class from a stream.
        /// </summary>
        /// <param name="stream">Stream contains a serialized instance of this class.</param>
        /// <param name="size">The number of bytes can read if -1, no limitation. MUST be -1.</param>
        /// <returns>Bytes have been read from the stream.</returns>
        public override int Deserialize(Stream stream, int size)
        {
            int bytesRead = 0;

            this.sizedXidList = new List<SizedXid>();
            while (((size >= 0) && bytesRead < size)
                || (size < 0 && stream.Position != stream.Length))
            {
                SizedXid xid = new SizedXid();
                bytesRead += xid.Deserialize(stream, -1);
                this.sizedXidList.Add(xid);
            }

            return bytesRead;
        }
    }
}