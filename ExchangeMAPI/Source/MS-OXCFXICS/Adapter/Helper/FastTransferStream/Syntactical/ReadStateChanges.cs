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
    /// <summary>
    /// The readStateChanges element contains information about MIDs of 
    /// Message objects that had their read state changed since the last 
    /// synchronization, as specified by the initial ICS state. 
    /// readStateChanges     = IncrSyncRead propList
    /// </summary>
    public class ReadStateChanges : SyntacticalBase
    {
        /// <summary>
        /// A propList value.
        /// </summary>
        private PropList propList;

        /// <summary>
        /// Initializes a new instance of the ReadStateChanges class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public ReadStateChanges(FastTransferStream stream)
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
        /// Verify that a stream's current position contains a serialized readStateChanges.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized readStateChanges, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return stream.VerifyMarker(Markers.PidTagIncrSyncRead);
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            this.Deserialize<PropList>(
                stream, 
                Markers.PidTagIncrSyncRead,
                out this.propList);
        }
    }
}