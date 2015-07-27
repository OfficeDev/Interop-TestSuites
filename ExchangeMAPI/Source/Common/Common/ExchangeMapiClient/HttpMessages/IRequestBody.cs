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
    /// A interface of request body for Mailbox Server Endpoint.
    /// </summary>
    public interface IRequestBody
    {
        /// <summary>
        /// Serialize the instance to byte array.
        /// </summary>
        /// <returns>The serialized data to be returned.</returns>
        byte[] Serialize();
    }
}
