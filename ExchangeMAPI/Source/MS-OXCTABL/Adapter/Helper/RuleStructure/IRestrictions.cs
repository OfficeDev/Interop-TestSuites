//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCTABL
{
    #region IRestrictions
    /// <summary>
    /// Interface of Restrictions
    /// </summary>
    public interface IRestrictions
    {
        /// <summary>
        /// Gets unsigned 8-bit integer. This value indicates the Type of restriction.
        /// </summary>
        RestrictionType RestrictType
        {
            get;
        }

        /// <summary>
        /// Get the total Size of Restriction
        /// </summary>
        /// <returns>The Size of Restriction buffer</returns>
        int Size();

        /// <summary>
        /// Get serialized byte array for this struct
        /// </summary>
        /// <returns>Serialized byte array</returns>
        byte[] Serialize();

        /// <summary>
        /// Deserialized byte array to an Restriction instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of an Restriction instance</param>
        /// <returns>Bytes count that deserialized in buffer</returns>
        uint Deserialize(byte[] buffer);
    }
    #endregion
}