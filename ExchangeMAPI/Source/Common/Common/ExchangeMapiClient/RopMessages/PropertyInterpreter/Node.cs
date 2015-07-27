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
    /// Abstract class defines the interface Parse,
    /// which must be implemented by all sub classes 
    /// </summary>
    public abstract class Node
    {
        /// <summary>
        /// Parse bytes in context into a Node
        /// </summary>
        /// <param name="context">The Context</param>
        public abstract void Parse(Context context);
    }
}