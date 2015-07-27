//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// The class is used to represent the editor.
    /// </summary>
    [Serializable]
    public class Editor
    {
        /// <summary>
        /// Gets or sets an int64 representing the editor’s timeout in its UTC "ticks".
        /// </summary>
        public long Timeout { get; set; }
        
        /// <summary>
        /// Gets or sets a unique id for the editor.
        /// </summary>
        public string CacheID { get; set; }
        
        /// <summary>
        /// Gets or sets the friendly name.
        /// </summary>
        public string FriendlyName { get; set; }
        
        /// <summary>
        /// Gets or sets the login name.
        /// </summary>
        public string LoginName { get; set; }
        
        /// <summary>
        /// Gets or sets the SIP address.
        /// </summary>
        public string SIPAddress { get; set; }
        
        /// <summary>
        /// Gets or sets the email address.
        /// </summary>
        public string EmailAddress { get; set; }
        
        /// <summary>
        /// Gets or sets a value indicating whether the user is an editor or reader.
        /// </summary>
        public bool HasEditorPermission { get; set; }
        
        /// <summary>
        /// Gets or sets a value which has up to 3 custom key/value pairs.
        /// </summary>
        public Dictionary<string, string> Metadata { get; set; }
    }
}