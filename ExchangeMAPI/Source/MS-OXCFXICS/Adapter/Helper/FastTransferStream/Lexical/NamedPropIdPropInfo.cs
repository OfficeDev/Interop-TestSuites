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
    /// The NamedPropIdPropInfo class.
    /// </summary>
    public class NamedPropIdPropInfo : PropInfo
    {
        /// <summary>
        /// The namedPropInfo in lexical definition.
        /// </summary>
        private NamedPropInfo namedPropInfo;

        /// <summary>
        /// Initializes a new instance of the NamedPropIdPropInfo class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public NamedPropIdPropInfo(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the propertySet.
        /// </summary>
        public Guid PropertySet
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the flag.
        /// </summary>
        public byte Flag
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the NamedPropInfo.
        /// </summary>
        public NamedPropInfo NamedPropInfo
        {
            get { return this.namedPropInfo; }
            set { this.namedPropInfo = value; }
        }
        
        /// <summary>
        /// Verify that a stream's current position contains a serialized NamedPropIdPropInfo.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized NamedPropIdPropInfo, return true, else false.</returns>
        public static new bool Verify(FastTransferStream stream)
        {
            return LexicalTypeHelper.IsNamedPropertyID(stream.VerifyUInt16());
        }

        /// <summary>
        /// Deserialize a NamedPropIdPropInfo instance from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        /// <returns>A NamedPropIdPropInfo instance </returns>
        public static new LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            return new NamedPropIdPropInfo(stream);
        }

        /// <summary>
        /// Deserialize next object from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void ConsumeNext(FastTransferStream stream)
        {
            base.ConsumeNext(stream);
            this.namedPropInfo = NamedPropInfo.DeserializeFrom(stream)
                    as NamedPropInfo;
        }
    }
}