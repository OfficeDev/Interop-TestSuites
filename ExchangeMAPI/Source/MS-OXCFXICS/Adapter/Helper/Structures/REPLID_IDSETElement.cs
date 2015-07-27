//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;
    using System.IO;

    /// <summary>
    /// A set of ID values with a REPLID.
    /// </summary>
    [SerializableObjectAttribute(true, true)]
    public class REPLID_IDSETElement : SerializableBase
    {
        /// <summary>
        ///    A REPLID value that when combined with all GLOBCNT values 
        ///    represented in the GLOBSET field, produces a set of IDs.
        /// </summary>
        [SerializableFieldAttribute(1)]
        private ushort repLID;

        /// <summary>
        /// A serialized GLOBSET.
        /// </summary>
        [SerializableFieldAttribute(2)]
        private GLOBSET globset;

        /// <summary>
        /// Gets or sets the GLOBSET.
        /// </summary>
        public GLOBSET GLOBSET
        {
            get
            {
                return this.globset;
            }

            set
            {
                this.globset = value;
            }
        }

        /// <summary>
        /// Gets message ids.
        /// </summary>
        /// <returns>A list of message ids.</returns>
        public List<ulong> GetIds()
        {
            List<ulong> idlist = new List<ulong>();
            if (this.globset != null && this.globset.GLOBCNTList != null)
            {
                List<GLOBCNT> cntList = this.globset.GLOBCNTList;
                for (int i = 0; i < cntList.Count; i++)
                {
                    byte[] repl = BitConverter.GetBytes(this.repLID);
                    byte[] id = new byte[8];
                    Array.Copy(repl, id, 2);
                    byte[] cnt = StructureSerializer.Serialize(cntList[i]);
                    Array.Copy(cnt, 0, id, 2, 6);
                    idlist.Add(BitConverter.ToUInt64(id, 0));
                }
            }

            return idlist;
        }

        /// <summary>
        /// Deserialize from a stream.
        /// </summary>
        /// <param name="stream">A stream contains serialize.</param>
        /// <param name="size">Must be -1.</param>
        /// <returns>The number of bytes read from the stream.</returns>
        public override int Deserialize(Stream stream, int size)
        {
            AdapterHelper.Site.Assert.AreEqual(-1, size, "The size value MUST be -1, the actual value is {0}.", size);

            this.globset = new GLOBSET();
            this.repLID = StreamHelper.ReadUInt16(stream);
            return this.globset.Deserialize(stream, -1) + 2;
        }
    }
}