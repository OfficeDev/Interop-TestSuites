//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    /// <summary>
    /// This class represents a type that can only set to null value.
    /// </summary>
    public sealed class Null
    {
        /// <summary>
        /// Prevents a default instance of the Null class from being created
        /// </summary>
        private Null()
        {
        }

        /// <summary>
        /// Gets the value null.
        /// </summary>
        /// <value>the only value null.</value>
        public static Null Value
        {
            get
            {
                return null;
            }
        }
    }
}