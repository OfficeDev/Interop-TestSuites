//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Microsoft.Protocols.TestTools;
    using Microsoft.Protocols.TestSuites.MS_ONESTORE;

    /// <summary>
    /// The interface of MS-ONESTOREAdapter class.
    /// </summary>
    public interface IMS_ONESTOREAdapter : IAdapter
    {
        /// <summary>
        /// Load and parse the OneNote revision-based file.
        /// </summary>
        /// <returns>Return the instacne of OneNoteRevisionStoreFile.</returns>
        OneNoteRevisionStoreFile LoadOneNoteFile(string fileName);

        /// <summary>
        /// Load and parse the OneNote alternative packaging file.
        /// </summary>
        /// <returns>Return the instacne of OneNoteAlternativePackagingFile.</returns>
        AlternativePackaging LoadOneNoteFileWithAlternativePackaging(string fileName);
    }
}