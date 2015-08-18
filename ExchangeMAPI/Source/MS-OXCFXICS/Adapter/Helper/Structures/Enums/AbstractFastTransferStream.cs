namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using Microsoft.Modeling;

    /// <summary>
    /// Abstract FastTransferStream.
    /// </summary>
    public struct AbstractFastTransferStream
    {
        /// <summary>
        /// The configure for FastTransfer stream.
        /// </summary>
        public FastTransferStreamType StreamType;

        /// <summary>
        /// The hierarchySync syntactical structure.
        /// </summary>
        public AbstractHierarchySync AbstractHierarchySync;

        /// <summary>
        /// The contentsSync syntactical structure.
        /// </summary>
        public AbstractContentsSync AbstractContentsSync;

        /// <summary>
        ///  The FolderContent syntactical structure. 
        /// </summary>
        public AbstractFolderContent AbstractFolderContent;

        /// <summary>
        ///  The TopFolder syntactical structure.
        /// </summary>
        public AbstractTopFolder AbstractTopFolder;

        /// <summary>
        ///  The ICS upload state. 
        /// </summary>
        public AbstractState AbstractState;

        /// <summary>
        /// The MessageContent structure.
        /// </summary>
        public AbstractMessageContent AbstractMessageContent;

        /// <summary>
        /// The MessageList structure.
        /// </summary>
        public AbstractMessageList AbstractMessageList;

        /// <summary>
        /// The server out message change partial.
        /// </summary>
        public bool IsSameReadstateChangeNumber;
    }

    /// <summary>
    /// Represents abstracted state for the state in fast transfer stream.
    /// </summary>
    public struct AbstractState
    {
        /// <summary>
        /// Abstract ICS State Index.
        /// </summary>
        public int AbstractICSStateIndex;

        /// <summary>
        /// Contains indexed of IDSET givens.
        /// </summary>
        public Set<int> IdSetGiven;

        /// <summary>
        /// State changeNumber.
        /// </summary>
        public bool IsNewCnsetReadPropertyChangeNumber;

        /// <summary>
        /// PidTagCnsetSeenFAIProperty changeNumber.
        /// </summary>
        public bool IsNewCnsetSeenFAIPropertyChangeNumber;

        /// <summary>
        /// PidTagCnsetSeenProperty ChangeNumber.
        /// </summary>
        public bool IsNewCnsetSeenPropertyChangeNumber;
    }

    /// <summary>
    /// FolderContent root element structure.
    /// </summary>
    public struct AbstractFolderContent
    {
        /// <summary>
        /// PidTagEcWarning property is output.
        /// </summary>
        public bool IsPidTagEcWarningOut;
        
        /// <summary>
        /// Indicate if the object that client has no permission not output.
        /// </summary>
        public bool IsNoPermissionObjNotOut;
        
        /// <summary>
        /// PidTagNewFXFolder is output.
        /// </summary>
        public bool IsPidTagNewFXFolderOut;
        
        /// <summary>
        /// Folder Content not out.
        /// </summary>
        public bool IsFolderContentNotOut;
        
        /// <summary>
        /// Folder properties is not output.
        /// </summary>
        public bool IsFolderPropertiesNotOut;
        
        /// <summary>
        /// Folder content's SubFolder elements is preceded by a PidTagFXDelProp meta-property. 
        /// </summary>
        public bool IsSubFolderPrecededByPidTagFXDelProp;

        /// <summary>
        /// Folder message element in the folder content.
        /// </summary>
        public AbstractFolderMessage AbsFolderMessage;
    }

    /// <summary>
    /// TopFolder root element structure.
    /// </summary>
    public struct AbstractTopFolder
    {
        /// <summary>
        /// The folderContent element in TopFolder.
        /// </summary>
        public AbstractFolderContent AbsFolderContent;

        /// <summary>
        /// Subfolder download from the server.
        /// </summary>
        public bool SubFolderInScope;
    }

    /// <summary>
    /// Hierarchy Sync.
    /// </summary>
    public struct AbstractHierarchySync
    {
        /// <summary>
        /// About folder change information.
        /// </summary>
        public AbstractFolderChange FolderchangeInfo;

        /// <summary>
        /// Specifies the count of subfolders.
        /// </summary>
        public int FolderCount;

        /// <summary>
        /// The final ICS state.
        /// </summary>
        public AbstractState FinalICSState;
        
        /// <summary>
        /// Indicates the parent folder change appear before child.
        /// </summary>
        public bool IsParentFolderBeforeChild;

        /// <summary>
        /// The abstracted deletions in hierarchySync object.
        /// </summary>
        public AbstractDeletion AbstractDeletion;
    }

    /// <summary>
    /// The contentsSync
    /// </summary>
    public struct AbstractContentsSync
    {
        /// <summary>
        /// Whether is progessTotalSameAsContentsync.
        /// </summary>
        public bool IsprogessTotalPresent;

        /// <summary>
        /// The container for messages info.
        /// </summary>
        public Set<AbstractMessageChangeInfo> MessageInfo;

        /// <summary>
        /// The abstracted deletions in hierarchySync object.
        /// </summary>
        public AbstractDeletion AbstractDeletion;

        /// <summary>
        /// Structure readStatechanges.
        /// </summary>
        public bool IsReadStateChangesExist;

        /// <summary>
        /// The final ICS state.
        /// </summary>
        public AbstractState FinalICSState;

        /// <summary>
        /// Identify Sort By MessageDeliveryTime  property whether is existent.
        /// </summary>
        public bool IsSortByMessageDeliveryTime;

        /// <summary>
        /// Identify sort by LastModificationTime property whether is existent.
        /// </summary>
        public bool IsSortByLastModificationTime;
    }

    /// <summary>
    /// The ProgressPerMessage.
    /// </summary>
    public struct AbstractMessageChangeInfo
    {
        /// <summary>
        /// Whether this element is present or not.
        /// </summary>
        public bool IsProgressPerMessagePresent;

        /// <summary>
        /// Followed FAI message or not.
        /// </summary>
        public bool FollowedFAIMessage;

        /// <summary>
        /// Indicate it contains message change full or partial.
        /// </summary>
        public bool IsMessageChangeFull;

        /// <summary>
        /// Identify PidTagMid property whether is existent.
        /// </summary>
        public bool IsPidTagMidExist;
        
        /// <summary>
        /// Identify PidTagMessageSize property whether is existent.
        /// </summary>
        public bool IsPidTagMessageSizeExist;
        
        /// <summary>
        /// Identify PidTagChangeNumber property whether is existent.
        /// </summary>
        public bool IsPidTagChangeNumberExist;

        /// <summary>
        /// The message Id index.
        /// </summary>
        public int MessageIdIndex;

        /// <summary>
        /// Whether RTF format.
        /// </summary>
        public bool IsRTFformat;
    }

    /// <summary>
    /// Deletion structure.
    /// </summary>
    public struct AbstractDeletion
    {
        /// <summary>
        /// Identify deletion structure whether exist.
        /// </summary>
        public bool IsDeletionPresent;

        /// <summary>
        /// Contains indexes of deleted ids.
        /// </summary>
        public Set<int> IdSetDeleted;

        /// <summary>
        /// Indicates whether PidTagIdsetNoLongerInScope exists.
        /// </summary>
        public bool IsPidTagIdsetNoLongerInScopeExist;

        /// <summary>
        /// Indicates whether PidTagIdsetExpired exists.
        /// </summary>
        public bool IsPidTagIdsetExpiredExist;
    }

    /// <summary>
    /// FolderChange structure.
    /// </summary>
    public struct AbstractFolderChange
    {
        /// <summary>
        /// Identify PidTagFolderId property whether is existent.
        /// </summary>
        public bool IsPidTagFolderIdExist;

        /// <summary>
        /// Identify PidTagParentFolderId  property whether is existent.
        /// </summary>
        public bool IsPidTagParentFolderIdExist;

        /// <summary>
        /// Identify PidTagSourceKey value is zero or not.
        /// </summary>
        public bool IsPidTagSourceKeyValueZero;

        /// <summary>
        /// Identify PidTagParentSourceKey value is zero or not.
        /// </summary>
        public bool IsPidTagParentSourceKeyValueZero;
    }

    /// <summary>
    /// MessageContent root element structure.
    /// </summary>
    public struct AbstractMessageContent
    {
        /// <summary>
        /// The objects that the client has no permission to access is not out from the server.
        /// </summary>
        public bool IsNoPermissionMessageNotOut;
        
        /// <summary>
        /// MessageChildren element.
        /// </summary>
        public AbstractMessageChildren AbsMessageChildren;
       
        /// <summary>
        /// String properties are out by the server in Unicode.
        /// </summary>
        public bool StringPropertiesInUnicode;

        /// <summary>
        /// String properties are output by the server in Unicode Code page property.
        /// </summary>
        public bool StringPropertiesInUnicodeCodePage;

        /// <summary>
        /// String properties are output by the server in other format stored on server by Code page property.
        /// </summary>
        public bool StringPropertiesInOtherCodePage;

        /// <summary>
        /// Boolean value indicates whether has a RTF body.
        /// </summary>
        public bool IsRTFFormat;
    }

    /// <summary>
    /// The MessageList root element structure.
    /// </summary>
    public struct AbstractMessageList
    {
        /// <summary>
        /// PidTagEcWarning property is output.
        /// </summary>
        public bool IsPidTagEcWarningOut;

        // Public Boolean isMessageBodyOutInOriginalFormat.

        /// <summary>
        /// The message element in MessageList.
        /// </summary>
        public AbsMessage AbsMessage;
    }

    /// <summary>
    /// The Message structure.
    /// </summary>
    public struct AbsMessage
    {
        /// <summary>
        /// The MessageContent element in Message.
        /// </summary>
        public AbstractMessageContent AbsMessageContent;
    }

    /// <summary>
    /// The FolderMessage structure.
    /// </summary>
    public struct AbstractFolderMessage
    {
        /// <summary>
        /// Folder content's folderMessage group is preceded by a PidTagFXDelProp meta-property. 
        /// </summary>
        public bool IsFolderMessagesPrecededByPidTagFXDelProp;
        
        /// <summary>
        /// The message element in folder message structure.
        /// </summary>
        public AbstractMessageList MessageList;
    }

    /// <summary>
    /// MessageChildren element structure.
    /// </summary>
    public struct AbstractMessageChildren
    {
        /// <summary>
        /// Attachment elements is preceded by a PidTagFXDelProp meta-property. 
        /// </summary>
        public bool AttachmentPrecededByPidTagFXDelProp;
        
        /// <summary>
        /// Recipient elements is preceded by a PidTagFXDelProp meta-property.
        /// </summary>
        public bool RecipientPrecededByPidTagFXDelProp;
    }
}