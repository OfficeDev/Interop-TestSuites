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
    using Microsoft.Modeling;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// The enumeration of record the prior operation.
    /// </summary>
    public enum PriorOperation
    {
        /// <summary>
        /// The RopSynchronizationImportMessageMove operation.
        /// </summary>
        RopSynchronizationImportMessageMove,

        /// <summary>
        /// The RopSynchronizationImportMessageChange operation.
        /// </summary>
        RopSynchronizationImportMessageChange,

        /// <summary>
        /// The RopSynchronizationImportHierarchyChange operation.
        /// </summary>
        RopSynchronizationImportHierarchyChange,

        /// <summary>
        /// The RopSynchronizationOpenCollector operation.
        /// </summary>
        RopSynchronizationOpenCollector,

        /// <summary>
        /// The RopFastTransferDestinationConfigure  operation. 
        /// </summary>
        RopFastTransferDestinationConfigure,

        /// <summary>
        /// The RopFastTransferSourceCopyProperties operation
        /// </summary>
        RopFastTransferSourceCopyProperties,

        /// <summary>
        /// The RopFastTransferSourceCopyFolder operation
        /// </summary>
        RopFastTransferSourceCopyFolder,

        /// <summary>
        /// The SynchronizationUploadState operation
        /// </summary>
        SynchronizationUploadState,

        /// <summary>
        /// The RopCreateMessage operation
        /// </summary>
        RopCreateMessage,

        /// <summary>
        /// The RopSynchronizationGetTransferState operation
        /// </summary>
        RopSynchronizationGetTransferState,

        /// <summary>
        /// The RopSynchronizationConfigure operation
        /// </summary>
        RopSynchronizationConfigure,
    
        /// <summary>
        /// The RopFastTransferSourceCopyMessage operation
        /// </summary>
        RopFastTransferSourceCopyMessage
    }

    /// <summary>
    /// The enumeration of record the prior download operation.
    /// </summary>
    public enum PriorDownloadOperation
    {
        /// <summary>
        /// The RopFastTransferSourceCopyMessage operation
        /// </summary>
        RopFastTransferSourceCopyMessage,

        /// <summary>
        /// The RopFastTransferSourceCopyTo operation
        /// </summary>
        RopFastTransferSourceCopyTo,

        /// <summary>
        /// The RopFastTransferSourceCopyProperties operation
        /// </summary>
        RopFastTransferSourceCopyProperties,

        /// <summary>
        /// The RopFastTransferSourceCopyFolder operation
        /// </summary>
        RopFastTransferSourceCopyFolder,

        /// <summary>
        /// The RopSynchronizationConfigure operation
        /// </summary>
        RopSynchronizationConfigure,

        /// <summary>
        /// The RopSynchronizationGetTransferState operation
        /// </summary>
        RopSynchronizationGetTransferState
    }

    /// <summary>
    /// The data for connection
    /// </summary>
    public struct ConnectionData
    {
        /// <summary>
        /// Logon server or not
        /// </summary>
        public int LogonHandleIndex;

        /// <summary>
        /// The local Id count.
        /// </summary>
        public uint LocalIdCount;

        /// <summary>
        /// logOn folder type
        /// </summary>
        public LogonFlags LogonFolderType;

        /// <summary>
        /// Contains folder id
        /// </summary>
        public Sequence<AbstractFolder> FolderContainer;

        /// <summary>
        /// Contains message id
        /// </summary>
        public Sequence<AbstractMessage> MessageContainer;

        /// <summary>
        /// Contains attachment id
        /// </summary>
        public Sequence<AbstractAttachment> AttachmentContainer;

        /// <summary>
        /// The download contexts created on the server
        /// </summary>
        public Sequence<AbstractDownloadInfo> DownloadContextContainer;

        /// <summary>
        /// The upload contexts created on the server
        /// </summary>
        public Sequence<AbstractUploadInfo> UploadContextContainer;
    }

    /// <summary>
    /// The abstract folder structure.
    /// </summary>
    public struct AbstractFolder
    {
        /// <summary>
        /// Folder Id index of parent folder.
        /// </summary>
        public int ParentFolderIdIndex;

        /// <summary>
        /// Folder handle index of parent folder.
        /// </summary>
        public int ParentFolderHandleIndex;

        /// <summary>
        /// The folder Id index.
        /// </summary>
        public int FolderIdIndex;

        /// <summary>
        /// The folder handle index.
        /// </summary>
        public int FolderHandleIndex;

        /// <summary>
        /// The count of subFolder.
        /// </summary>
        public Set<int> SubFolderIds;

        /// <summary>
        /// The count of messages.
        /// </summary>
        public Set<int> MessageIds;

        /// <summary>
        /// The folder properties.
        /// </summary>
        public Set<string> FolderProperties;

        /// <summary>
        /// The change number index.
        /// </summary>
        public int ChangeNumberIndex;

        /// <summary>
        /// The folder permission.
        /// </summary>
        public PermissionLevels FolderPermission;

        /// <summary>
        /// Contains the ICS State have been downloaded.
        /// </summary>
        public MapContainer<int, AbstractUpdatedState> ICSStateContainer;
    }

    /// <summary>
    /// The abstract Updated State structure.
    /// </summary>
    public struct AbstractUpdatedState
    {
        /// <summary>
        /// Contains PidTagIdsetGiven.
        /// </summary>
        public Set<int> IdsetGiven;

        /// <summary>
        /// Contains PidTagCnsetSeen.
        /// </summary>
        public Set<int> CnsetSeen;

        /// <summary>
        /// Contains PidTagCnsetSeenFAI.
        /// </summary>
        public Set<int> CnsetSeenFAI;

        /// <summary>
        /// Contains PidTagCnsetRead. 
        /// </summary>
        public Set<int> CnsetRead;
    }

    /// <summary>
    /// The abstract message structure.
    /// </summary>
    public struct AbstractMessage
    {
        /// <summary>
        /// The folder Id index.
        /// </summary>
        public int FolderIdIndex;

        /// <summary>
        /// The folder handle index.
        /// </summary>
        public int FolderHandleIndex;

        /// <summary>
        /// The message Id index.
        /// </summary>
        public int MessageIdIndex;

        /// <summary>
        /// The message handle index.
        /// </summary>
        public int MessageHandleIndex;

        /// <summary>
        /// The message is FAI or not.
        /// </summary>
        public bool IsFAImessage;

        /// <summary>
        /// The message is read or not.
        /// </summary>
        public bool IsRead;

        /// <summary>
        /// The count of the attachments.
        /// </summary>
        public int AttachmentCount;

        /// <summary>
        /// The message properties.
        /// </summary>
        public Sequence<string> MessageProperties;

        /// <summary>
        /// The change number index.
        /// </summary>
        public int ChangeNumberIndex;

        /// <summary>
        /// The readState Change Number
        /// </summary>
        public int ReadStateChangeNumberIndex;
    }

    /// <summary>
    /// The abstract attachment structure.
    /// </summary>
    public struct AbstractAttachment
    {
        /// <summary>
        /// The attachment handle index.
        /// </summary>
        public int AttachmentHandleIndex;
    }

    /// <summary>
    /// The abstract download context
    /// </summary>
    public struct AbstractDownloadInfo
    {
        /// <summary>
        /// The messaging object type generate the download context.
        /// </summary>
        public ObjectType ObjectType;

        /// <summary>
        /// The relative messaging object handle index generate the download context.
        /// </summary>
        public int RelatedObjectHandleIndex;

        /// <summary>
        /// The download handle index.
        /// </summary>
        public int DownloadHandleIndex;

        /// <summary>
        /// The configuration for the FastTransfer stream.
        /// </summary>
        public FastTransferStreamType AbstractFastTransferStreamType;

        /// <summary>
        /// The stream index for download index.
        /// </summary>
        public int DownloadStreamIndex;

        /// <summary>
        /// The relative SendOptions
        /// </summary>
        public SendOptionAlls Sendoptions;

        /// <summary>
        /// Save the copyFlag type of CopyFolder operation
        /// </summary>
        public CopyFolderCopyFlags CopyFolderCopyFlag;

        /// <summary>
        /// Save the copypFlag type of CopyProperties operation
        /// </summary>
        public CopyPropertiesCopyFlags CopyPropertiesCopyFlag;

        /// <summary>
        /// Save the copyFlag type of CopyTo operation
        /// </summary>
        public CopyToCopyFlags CopyToCopyFlag;

        /// <summary>
        /// Save the copyFlag type of CopyMessage operation
        /// </summary>
        public RopFastTransferSourceCopyMessagesCopyFlags CopyMessageCopyFlag;

        /// <summary>
        /// Variable to indicate which FastTransfer Operation is called
        /// </summary>
        public EnumFastTransferOperation RelatedFastTransferOperation;

        /// <summary>
        /// The relative SynchronizationFlag
        /// </summary>
        public SynchronizationFlag Synchronizationflag;

        /// <summary>
        /// The relative SynchronizationExtraflag
        /// </summary>
        public SynchronizationExtraFlag SynchronizationExtraflag;

        /// <summary>
        /// The relative Property
        /// </summary>
        public Sequence<string> Property;

        /// <summary>
        /// The updated ICS state.
        /// </summary>
        public AbstractUpdatedState UpdatedState;

        /// <summary>
        /// The synchronization type.
        /// </summary>
        public SynchronizationTypes SynchronizationType;

        /// <summary>
        /// Indicates whether descendant sub-objects are copied.
        /// </summary>
        public bool IsLevelTrue;
    }

    /// <summary>
    /// The abstract upload context
    /// </summary>
    public struct AbstractUploadInfo
    {
        /// <summary>
        /// The relative messaging object handle index generate the upload context.
        /// </summary>
        public int RelatedObjectHandleIndex;

        /// <summary>
        /// The related object id index.
        /// </summary>
        public int RelatedObjectIdIndex;

        /// <summary>
        /// The upload handle index.
        /// </summary>
        public int UploadHandleIndex;

        /// <summary>
        /// The synchronization type.
        /// </summary>
        public SynchronizationTypes SynchronizationType;

        /// <summary>
        /// The relative ImportDeleteflags.
        /// </summary>
        public byte ImportDeleteflags;

        /// <summary>
        /// Identify the result whether is newerClientChange or not.
        /// </summary>
        public bool IsnewerClientChange;

        /// <summary>
        /// Variable to indicate which FastTransfer Operation is called
        /// </summary>
        public EnumFastTransferOperation RelatedFastTransferOperation;

        /// <summary>
        /// The updated ICS state.
        /// </summary>
        public AbstractUpdatedState UpdatedState;
    }

    /// <summary>
    /// Helper class of Model
    /// </summary>
    public static class ModelHelper
    {
        /// <summary>
        /// The change number index.
        /// </summary>
        private static int changeNumberIndex = 0;

        /// <summary>
        /// Requirement capture
        /// </summary>
        /// <param name="id">Requirement id</param>
        /// <param name="description">Requirement description</param>
        public static void CaptureRequirement(int id, string description)
        {
            Requirement.Capture(RequirementId.Make("MS-OXCFXICS", id, description));
        }

        /// <summary>
        /// Assign a new change number.
        /// </summary>
        /// <returns>The current change number.</returns>
        public static int GetChangeNumberIndex()
        {
            int tempChangeNumberIndex = ++changeNumberIndex;
            return tempChangeNumberIndex;
        }
    }
}