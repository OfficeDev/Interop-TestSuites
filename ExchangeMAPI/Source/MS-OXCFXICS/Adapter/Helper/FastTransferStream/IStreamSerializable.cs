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
    ///  An interface serializable object must implement.
    /// </summary>
    public interface IStreamSerializable
    {
        /// <summary>
        /// Serialize object to a FastTransferStream
        /// </summary>
        /// <returns>A FastTransferStream contains the serialized object</returns>
        FastTransferStream Serialize();
    }
}