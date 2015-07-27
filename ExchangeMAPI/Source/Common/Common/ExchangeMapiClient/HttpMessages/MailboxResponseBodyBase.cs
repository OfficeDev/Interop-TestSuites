//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// A base class of response body for Mailbox Server Endpoint.
    /// </summary>
    public abstract class MailboxResponseBodyBase
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the status of the request.
        /// </summary>
        public uint StatusCode { get; protected set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.
        /// </summary>
        public uint AuxiliaryBufferSize { get; protected set; }

        /// <summary>
        /// Gets or sets an array of bytes that constitute the auxiliary payload data returned from the server. 
        /// </summary>
        public byte[] AuxiliaryBuffer { get; protected set; }
    }
}
