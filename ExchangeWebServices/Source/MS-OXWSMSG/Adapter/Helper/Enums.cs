//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    /// <summary>
    /// An enumeration that describes roles of a message. 
    /// </summary>
    public enum Role
    {
        /// <summary>
        /// Represent the sender of a message.
        /// </summary>
        Sender,

        /// <summary>
        /// Represent the recipient of a message.
        /// </summary>
        Recipient1,

        /// <summary>
        /// Represent the recipient of a message.
        /// </summary>
        Recipient2,
    }
}