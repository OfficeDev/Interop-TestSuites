//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Text;

    /// <summary>
    /// Refers to an attachment Id
    /// </summary>
    public struct AttachmentId
    {
        /// <summary> 
        /// The attachment Id bytes length 
        /// </summary>
        public short AttachmentIdLength;

        /// <summary>
        /// Attachment Id bytes
        /// </summary>
        public byte[] Id;
    }

    /// <summary>
    /// A parsed itemid's id  
    /// </summary>
    public class ItemIdId
    {
        /// <summary>
        /// Indicates whether Run Length Encoding (RLE) is used
        /// </summary>
        private byte compressionByte = 0;

        /// <summary>
        /// Indicates the type of the Id
        /// </summary>
        private IdStorageType storageType;

        /// <summary>
        /// Moniker Length
        /// </summary>
        private short? monikerLength;

        /// <summary>
        /// Moniker Bytes
        /// </summary>
        private byte[] monikerBytes;

        /// <summary>
        /// Indicates any special processing to perform on an Id when deserializing it
        /// </summary>
        private IdProcessingInstructionType? processingInstruction;

        /// <summary>
        /// Store Id Bytes Length
        /// </summary>
        private short storeIdLength = 0;

        /// <summary>
        /// Store Id Bytes
        /// </summary>
        private byte[] storeId;

        /// <summary>
        /// Folder Id Bytes Length
        /// </summary>
        private short? folderIdLength;

        /// <summary>
        /// Folder Id Bytes
        /// </summary>
        private byte[] folderId;

        /// <summary>
        /// Indicates how many attachments are in the hierarchy.
        /// </summary>
        private byte? attachmentIdCount;

        /// <summary>
        /// Attachment Ids
        /// </summary>
        private AttachmentId[] attachmentIds;

        /// <summary>
        /// Gets or sets the byte which indicates whether Run Length Encoding (RLE) is used
        /// </summary>
        public byte CompressionByte
        {
            get { return this.compressionByte; }
            set { this.compressionByte = value; }
        }

        /// <summary>
        /// Gets or sets the type of the Id
        /// </summary>
        public IdStorageType StorageType
        {
            get { return this.storageType; }
            set { this.storageType = value; }
        }

        /// <summary>
        /// Gets or sets the moniker length
        /// </summary>
        public short? MonikerLength
        {
            get { return this.monikerLength; }
            set { this.monikerLength = value; }
        }

        /// <summary>
        /// Gets or sets the moniker bytes
        /// </summary>
        public byte[] MonikerBytes
        {
            get { return this.monikerBytes; }
            set { this.monikerBytes = value; }
        }

        /// <summary>
        /// Gets the moniker string for MailboxItemSmtpAddressBased
        /// </summary>
        public string MonikerString
        {
            get
            {
                if (this.storageType == IdStorageType.MailboxItemSmtpAddressBased)
                {
                    return Encoding.UTF8.GetString(this.monikerBytes, 0, this.monikerBytes.Length);
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets the moniker Guid for ConversationIdMailboxGuidBased or MailboxItemMailboxGuidBased
        /// </summary>
        public Guid? MonikerGuid
        {
            get
            {
                if (this.storageType == IdStorageType.ConversationIdMailboxGuidBased || this.storageType == IdStorageType.MailboxItemMailboxGuidBased)
                {
                    return new Guid(Encoding.UTF8.GetString(this.monikerBytes, 0, this.monikerBytes.Length));
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the processing type to perform on an Id when deserializing it
        /// </summary>
        public IdProcessingInstructionType? IdProcessingInstruction
        {
            get { return this.processingInstruction; }
            set { this.processingInstruction = value; }
        }

        /// <summary>
        /// Gets or sets the store Id length
        /// </summary>
        public short StoreIdLength
        {
            get { return this.storeIdLength; }
            set { this.storeIdLength = value; }
        }

        /// <summary>
        /// Gets or sets the store Id bytes
        /// </summary>
        public byte[] StoreId
        {
            get { return this.storeId; }
            set { this.storeId = value; }
        }

        /// <summary>
        /// Gets or sets the folder Id length
        /// </summary>
        public short? FolderIdLength
        {
            get { return this.folderIdLength; }
            set { this.folderIdLength = value; }
        }

        /// <summary>
        /// Gets or sets the folder Id bytes
        /// </summary>
        public byte[] FolderId
        {
            get { return this.folderId; }
            set { this.folderId = value; }
        }

        /// <summary>
        /// Gets or sets the count of attachments.
        /// </summary>
        public byte? AttachmentIdCount
        {
            get { return this.attachmentIdCount; }
            set { this.attachmentIdCount = value; }
        }

        /// <summary>
        /// Gets or sets the attachment Id array
        /// </summary>
        public AttachmentId[] AttachmentIds
        {
            get { return this.attachmentIds; }
            set { this.attachmentIds = value; }
        }
    }
}
