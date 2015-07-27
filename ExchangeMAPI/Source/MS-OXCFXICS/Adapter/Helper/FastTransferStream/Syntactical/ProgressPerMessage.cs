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

    /// <summary>
    /// The ProgressPerMessage element contains data that describes 
    /// the approximate size of message change data that follows.
    /// ProgressPerMessage   = IncrSyncProgressPerMsg propList
    /// </summary>
    public class ProgressPerMessage : SyntacticalBase
    {
        /// <summary>
        /// A list of propList.
        /// </summary>
        private PropList propList;

        /// <summary>
        /// Initializes a new instance of the ProgressPerMessage class.
        /// </summary>
        /// <param name="stream">a FastTransferStream</param>
        public ProgressPerMessage(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets a value indicating whether the message followed is a FAI message.
        /// </summary>
        public bool IsFollowedFAIMessage
        {
            get
            {
                // Gets the second property value in the property list.
                ushort fai = Convert.ToUInt16(
                    (this.PropList.PropValues[1] as FixedPropTypePropValue).FixedValue);

                return fai != 0;
            }
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
        /// Verify that a stream's current position contains a serialized ProgressPerMessage.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized ProgressPerMessage, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return stream.VerifyMarker(Markers.PidTagIncrSyncProgressPerMsg);
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            this.Deserialize<PropList>(
                stream,
                Markers.PidTagIncrSyncProgressPerMsg,
                out this.propList);
        }
    }
}