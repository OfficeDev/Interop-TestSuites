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
    using System.Collections.Generic;
    using System.IO;

    /// <summary>
    /// The deletions element contains information about IDs of messaging objects 
    /// that had been deleted, expired, or moved out of the synchronization scope 
    /// since the last synchronization, as specified in the initial ICS state. 
    /// deletions            = IncrSyncDel propList
    /// </summary>
    public class Deletions : SyntacticalBase
    {
        /// <summary>
        /// A propList value.
        /// </summary>
        private PropList propList;

        /// <summary>
        /// Initializes a new instance of the Deletions class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public Deletions(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the propList.
        /// </summary>
        public PropList PropList
        {
            get { return this.propList; }
            set { this.propList = value; }
        }

        /// <summary>
        /// Gets the deleted IDSET.
        /// </summary>
        public List<ulong> IdsetDeleted
        {
            get
            {
                // Get PidTagIdsetDeleted value in the PropList.
                byte[] buffer = this.PropList.GetPropValue(0x67E5, 0x0102) as byte[];
                using (MemoryStream ms = new MemoryStream(buffer, false))
                {
                    REPLID_IDSET idset = new REPLID_IDSET();
                    idset.Deserialize(ms, -1);
                    return idset.GetIds();
                }
            }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized deletions.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized deletions, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return stream.VerifyMarker(Markers.PidTagIncrSyncDel);
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            if (stream.ReadMarker(Markers.PidTagIncrSyncDel))
            {
                this.propList = new PropList(stream);
                return;
            }

            AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
        }
    }
}