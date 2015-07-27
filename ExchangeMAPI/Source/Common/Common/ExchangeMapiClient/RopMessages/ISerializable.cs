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
    /// This interface define the methods that is needed to serialize an ROP object
    /// </summary>
    public interface ISerializable
    {
        /// <summary>
        /// Serialize into a bytes array.
        /// </summary>
        /// <returns>The bytes array serialized</returns>
        byte[] Serialize();

        /// <summary>
        /// Return the size in bytes of the object serialized
        /// </summary>
        /// <returns>The size in bytes of the object serialized</returns>
        int Size();
    }
}