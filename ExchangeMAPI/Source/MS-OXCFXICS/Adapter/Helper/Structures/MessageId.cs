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
    /// The MessageId field
    /// </summary>
    public class MessageId
    {
        /// <summary>
        /// Initializes a new instance of the MessageId class.
        /// </summary>
        /// <param name="value">The message Id value.</param>
        public MessageId(ulong value)
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