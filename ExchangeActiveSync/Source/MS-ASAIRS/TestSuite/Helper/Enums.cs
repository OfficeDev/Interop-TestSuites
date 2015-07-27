//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    /// <summary>
    /// The type of email.
    /// </summary>
    public enum EmailType
    {
        /// <summary>
        /// The email type is plain text
        /// </summary>
        Plaintext,

        /// <summary>
        /// The email type is html
        /// </summary>
        HTML,

        /// <summary>
        /// The email attachment is a normal attachment
        /// </summary>
        NormalAttachment,

        /// <summary>
        /// The email attachment is an e-mail message
        /// </summary>
        EmbeddedAttachment,

        /// <summary>
        /// The email attachment is an embedded Object Linking and Embedding (OLE) object
        /// </summary>
        AttachOLE
    }
}