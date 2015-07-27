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
    using System.IO;

    /// <summary>
    /// The state element contains the final ICS state of 
    /// the synchronization download operation. 
    /// </summary>
    public class State : SyntacticalBase
    {
        /// <summary>
        /// The start marker of state.
        /// </summary>
        public const Markers StartMarker = Markers.PidTagIncrSyncStateBegin;

        /// <summary>
        /// The end marker of state.
        /// </summary>
        public const Markers EndMarker = Markers.PidTagIncrSyncStateEnd;

        /// <summary>
        /// The propList in the syntactical definition of state.
        /// </summary>
        private PropList props;

        /// <summary>
        /// Initializes a new instance of the State class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public State(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets the IdsetGiven.
        /// </summary>
        public REPLGUID_IDSET IdsetGiven
        {
            get
            {
                AdapterHelper.Site.Assert.IsNull(this.GetIdset(0x4017, 0x0102), "The value of MetaTagIdsetGiven should be null.");

                return this.GetIdset(0x4017, 0x0003);
            }
        }

        /// <summary>
        /// Gets the CnsetSeen.
        /// </summary>
        public REPLGUID_IDSET CnsetSeen
        {
            get
            {
                return this.GetIdset(0x6796, 0x0102);
            }
        }

        /// <summary>
        /// Gets the CnsetSeenFAI.
        /// </summary>
        public REPLGUID_IDSET CnsetSeenFAI
        {
            get
            {
                return this.GetIdset(0x67da, 0x0102);
            }
        }

        /// <summary>
        /// Gets the CnsetSeenFAI.
        /// </summary>
        public REPLGUID_IDSET CnsetRead
        {
            get
            {
                return this.GetIdset(0x67D2, 0x0102);
            }
        }

        /// <summary>
        /// Gets the propList.
        /// </summary>
        public PropList PropList
        {
            get
            {
                return this.props;
            }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized state.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized state, return true, else false</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return stream.VerifyMarker(StartMarker)
                && stream.VerifyMarker(
                    EndMarker,
                    (int)stream.Length - MarkersHelper.PidTagLength - (int)stream.Position);
        }

        /// <summary>
        /// Gets the REPLGUID_IDSET.
        /// </summary>
        /// <param name="id">The id to be calculated.</param>
        /// <param name="type">The type to be calculated.</param>
        /// <returns>A REPLGUID_IDSET.</returns>
        public REPLGUID_IDSET GetIdset(ushort id, ushort type)
        {
            if (this.PropList != null)
            {
                byte[] buffer = this.PropList.GetPropValue(id, type) as byte[];
                if (buffer != null && buffer.Length > 0)
                {
                    using (MemoryStream ms = new MemoryStream(buffer, false))
                    {
                        REPLGUID_IDSET idset = null;
                        idset = new REPLGUID_IDSET();
                        idset.Deserialize(ms, -1);
                        return idset;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            if (stream.ReadMarker(StartMarker))
            {
                this.props = new PropList(stream);
                if (stream.ReadMarker(EndMarker))
                {
                    return;
                }
            }

            AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
        }
    }
}