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
    /// This interface define the methods that is needed to deserialize a bytes array into an ROP object
    /// </summary>
    public interface IDeserializable
    {
        /// <summary>
        /// Deserialize input bytes ropBytes into a ROP object
        /// </summary>
        /// <param name="ropBytes">The bytes array to deserialize</param>
        /// <param name="startIndex">The start index of the byte array</param>
        /// <returns>The bytes deserialized</returns>
        int Deserialize(byte[] ropBytes, int startIndex);
    }
}