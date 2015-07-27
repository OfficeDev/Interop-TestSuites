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
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A set of ID values with a REPLGUID.
    /// </summary>
    [SerializableObjectAttribute(false, true)]
    public class REPLGUID_IDSETElement : SerializableBase
    {
        /// <summary>
        /// A GUID value that represents a REPLGUID. When combined with 
        /// all GLOBCNT values represented in the GLOBSET field, produces 
        /// a set of GIDs. The GUID values can be converted into a REPLID to 
        /// produce a set of IDs.
        /// </summary>
        [SerializableFieldAttribute(1)]
        private Guid replguid;

        /// <summary>
        /// A serialized GLOBSET.
        /// </summary>
        [SerializableFieldAttribute(2)]
        private GLOBSET globset;

        /// <summary>
        /// Gets or sets the REPLGUID.
        /// </summary>
        public Guid REPLGUID
        {
            get
            {
                return this.replguid;
            }

            set
            {
                this.replguid = value;
            }
        }

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
        /// Gets LongTermIds.
        /// </summary>
        /// <returns>A list of LongTermIds.</returns>
        public List<LongTermId> GetLongTermIds()
        {
            List<LongTermId> idlist = new List<LongTermId>();
            if (this.globset != null && this.globset.GLOBCNTList != null)
            {
                List<GLOBCNT> cntList = this.globset.GLOBCNTList;
                for (int i = 0; i < cntList.Count; i++)
                {
                    LongTermId lid = new LongTermId
                    {
                        DatabaseGuid = this.replguid.ToByteArray(),
                        GlobalCounter = StructureSerializer.Serialize(cntList[i])
                    };
                    idlist.Add(lid);
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

            this.replguid = StreamHelper.ReadGuid(stream);
            this.globset = new GLOBSET();
            return 0x10 + this.globset.Deserialize(stream, -1);
        }
    }
}