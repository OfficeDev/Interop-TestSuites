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
    /// The structure to save how the FastTransfer stream are download
    /// </summary>
    public struct StreamBufferContent
    {
        #region CopyFlag

        /// <summary>
        /// The subfolders of the folder specified download
        /// </summary>
        public bool SubfoldersDownload;                  // 111is true if CopyFoldercopySubfolder flag is set

        /// <summary>
        /// The folder properties download
        /// </summary>
        public bool FolderPropertiesDownload;            // is true if CopyFolderNoGhostContent flag is set

        /// <summary>
        /// PidTagNewFXFolder property is download
        /// </summary>
        public bool PidTagNewFXFolderDownload;           // 111is true if CopyFolderNoGhostContent flag is not set but the folder is Ghost folder actually

        /// <summary>
        /// The folder content download
        /// </summary>
        public bool FolderContentDownload;               // 111is true if CopyFolderNoGhostContent flag is set and the folder is not Ghost folder

        /// <summary>
        /// The  message bodies download in the compressed RTF format
        /// </summary>
        public bool MessageBodyDownloadInRTFFormat;      // is true if bestBody flag not set

        /// <summary>
        /// Message and change identification information is download
        /// </summary>
        public bool MessageAndChangeIdDownlaod;          // is true if CopyMessageSendEntryId flag is set

        /// <summary>
        /// Download objects that the client has no permissions 
        /// </summary>
        public bool NoPermissionsObjectDownload;         // is false if CopyMessageMove flag is set

        /// <summary>
        /// The body of embedded messages download in original format
        /// </summary>
        public bool EmbedMessageBodyInOriginalFormat;    // is true if bestBody flag is set

        /// <summary>
        /// The message body download in original format
        /// </summary>
        public bool MessageBodyDownloadInOriginalFormat; // is true if bestBody flag is set

        #endregion

        #region Sendoption

        /// <summary>
        /// Properties download in Unicode 
        /// </summary>
        public bool PropertiesDownloadInUnicode;         // is true if 1.unicode is set and the property saved in Unicode in the server. 2. Unicode|forceUnicode is set

        /// <summary>
        /// Properties download in code page set on the current connection
        /// </summary>
        public bool PropertiesDownloadInCodePage;        // is true if 1.unicode is set and the property saved not in Unicode in the server 2.The sendOption is not  Unicode and Unicode|forceUnicode 

        /// <summary>
        /// [In messageChange]A server use messageChangeFull
        /// </summary>
        public bool MessageChangeFull;                   // 111is true if partialItem is not set

        /// <summary>
        /// [In messageChange]A server use MessageChangePartial
        /// </summary>
        public bool MessageChangePartial;                // 111is true if priaialItem is set

        /// <summary>
        /// errorInfo  element is used
        /// </summary>
        public bool ErrorInfoUsed;                       // 111is true if RecoverMode is set

        #endregion
    }
}