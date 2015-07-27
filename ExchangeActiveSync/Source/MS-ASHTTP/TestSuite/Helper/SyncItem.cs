//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASHTTP
{
    /// <summary>
    /// Wrapper class of ApplicationData from Sync command response.
    /// </summary>
    public class SyncItem
    {
        /// <summary>
        /// Gets or sets the ServerId.
        /// </summary>
        public string ServerId { get; set; }

        /// <summary>
        /// Gets or sets the Subject of the item.
        /// </summary>
        public string Subject { get; set; }
    }
}