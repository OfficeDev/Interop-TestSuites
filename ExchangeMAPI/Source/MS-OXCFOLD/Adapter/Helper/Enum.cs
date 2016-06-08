namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    /// <summary>
    /// Property ID Enum
    /// </summary>
    public enum FolderPropertyId : ushort
    {
        /// <summary>
        /// Indicates the operations available to the client for the object.
        /// </summary>
        PidTagAccess = 0x0FF4,

        /// <summary>
        /// Contains a structure that identifies the last change to the object.
        /// </summary>
        PidTagChangeKey = 0x65E2,

        /// <summary>
        /// Contains the time, in UTC, that the object was created.
        /// </summary>
        PidTagCreationTime = 0x3007,

        /// <summary>
        /// Contains the time, in UTC, of the last modification to the object.
        /// </summary>
        PidTagLastModificationTime = 0x3008,

        /// <summary>
        /// Specifies the number of rows under the header row.
        /// </summary>
        PidTagContentCount = 0x3602,

        /// <summary>
        /// Specifies the number of unread messages in a folder, as computed by the message store.
        /// </summary>
        PidTagContentUnreadCount = 0x3603,

        /// <summary>
        /// Specifies the time, in UTC, when the item or folder was soft deleted.
        /// </summary>
        PidTagDeletedOn = 0x668F,

        /// <summary>
        /// Contains the name-service EntryID of a directory object that refers to a public folder.
        /// </summary>
        PidTagAddressBookEntryId = 0x663B,

        /// <summary>
        /// Contains the Folder ID (FID) ([MS-OXCDATA] section 2.2.1.1) of the folder.
        /// </summary>
        PidTagFolderId = 0x6748,

        /// <summary>
        /// Contains the EntryID of the folder where messages reside.
        /// </summary>
        PidTagParentEntryId = 0x0E09,

        /// <summary>
        /// Contains a number that monotonically increases every time a subfolder is added to, or deleted from, this folder.
        /// </summary>
        PidTagHierarchyChangeNumber = 0x663E,

        /// <summary>
        /// Contains the size, in bytes, consumed by the Message object on the server.
        /// </summary>
        PidTagMessageSize = 0x0E08,

        /// <summary>
        /// Specifies the 64-bit version of the PidTagNormalMessageSize property (section 2.877).
        /// </summary>
        PidTagMessageSizeExtended = 0x0E08,

        /// <summary>
        /// Specifies whether a folder has subfolders.
        /// </summary>
        PidTagSubfolders = 0x360A,

        /// <summary>
        /// Specifies the hide or show status of a folder.
        /// </summary>
        PidTagAttributeHidden = 0x10F4,

        /// <summary>
        /// Contains a comment about the purpose or content of the Address Book object.
        /// </summary>
        PidTagComment = 0x3004,

        /// <summary>
        /// Contains a string value that describes the type of Message object that a folder contains.
        /// </summary>
        PidTagContainerClass = 0x3613,

        /// <summary>
        /// Contains identifiers of the subfolders that are contained in the folder.
        /// </summary>
        PidTagContainerHierarchy = 0x360E,

        /// <summary>
        /// Contains the display name of the folder.
        /// </summary>
        PidTagDisplayName = 0x3001,

        /// <summary>
        /// Contains the display name of the folder.
        /// </summary>
        PidTagFolderAssociatedContents = 0x3610,

        /// <summary>
        /// Specifies the type of a folder that includes the Root folder, Generic folder, and Search folder.
        /// </summary>
        PidTagFolderType = 0x3601,

        /// <summary>
        /// Specifies a user's folder permissions.
        /// </summary>
        PidTagRights = 0x6639,

        /// <summary>
        /// Contains a permissions list for a folder.
        /// </summary>
        PidTagAccessControlListData = 0x3FE0,

        /// <summary>
        /// Contains the permissions for the specified user.
        /// </summary>
        PidTagMemberRights = 0x6673,

        /// <summary>
        /// Contains the information to identify many different types of messaging objects.
        /// </summary>
        PidTagEntryId = 0x0fff,

        /// <summary>
        /// Contains the time, in UTC, that a RopCreateMessage remote operation.
        /// </summary>
        PidTagLocalCommitTime = 0x6709,

        /// <summary>
        /// Contains the time of the most recent message change within the folder container, excluding messages changed within subfolders.
        /// </summary>
        PidTagLocalCommitTimeMax = 0x670A,

        /// <summary>
        /// Contains the total count of messages that have been deleted from a folder, excluding messages deleted within subfolders.
        /// </summary>
        PidTagDeletedCountTotal = 0x670B,

        /// <summary>
        /// Contains a computed value to specify the type or state of a folder.
        /// </summary>
        PidTagFolderFlags = 0x66A8,

        /// <summary>
        /// Specifies the time, in UTC, to trigger the client in cached mode to synchronize the folder hierarchy.
        /// </summary>
        PidTagHierRev = 0x4082
    }

    /// <summary>
    /// Message property Id enum.
    /// </summary>
    public enum MessagePropertyId : short
    {
        /// <summary>
        /// Specifies a message class.
        /// </summary>
        PidTagMessageClass = 0x001A,

        /// <summary>
        /// Corresponds to the message-id field. Data type: PtypString.
        /// </summary>
        PidTagInternetMessageId = 0x1035,

        /// <summary>
        /// Specifies a message ID.
        /// </summary>
        PidTagMid = 0x674A,

        /// <summary>
        /// Contains the EntryID of the folder where messages reside.
        /// </summary>
        PidTagParentEntryId = 0x0E09
    }

    /// <summary>
    /// Restrict type enum.
    /// </summary>
    public enum RestrictType : byte
    {
        /// <summary>
        /// AndRestriction Structure.
        /// </summary>
        AndRestriction = 0x00,

        /// <summary>
        /// NotRestriction Structure.
        /// </summary>
        NotRestriction = 0x02,

        /// <summary>
        /// ContentRestriction Structure.
        /// </summary>
        ContentRestriction = 0x03,

        /// <summary>
        /// PropertyRestriction Structure.
        /// </summary>
        PropertyRestriction = 0x04,

        /// <summary>
        /// ExistRestriction Structure.
        /// </summary>
        ExistRestriction = 0x08
    }

    /// <summary>
    /// The FuzzyLevelLow enum.
    /// </summary>
    public enum FuzzyLevelLowValues : short
    {
        /// <summary>
        /// The value stored in the TaggedValue field and the value of the column property tag match one another in their entirety.
        /// </summary>
        FL_FULLSTRING = 0x0000,

        /// <summary>
        /// The value stored in the TaggedValue field matches some portion of the value of the column property tag.
        /// </summary>
        FL_SUBSTRING = 0x0001,

        /// <summary>
        /// The value stored in the TaggedValue field matches a starting portion of the value of the column property tag.
        /// </summary>
        FL_PREFIX = 0x0002
    }

    /// <summary>
    /// The FuzzyLevelHigh enum.
    /// </summary>
    public enum FuzzyLevelHighValues : short
    {
        /// <summary>
        /// The comparison is case insensitive.
        /// </summary>
        FL_IGNORECASE = 0x0001,

        /// <summary>
        /// The comparison ignores Unicode-defined nonspacing characters such as diacritical marks.
        /// </summary>
        FL_IGNORENONSPACE = 0x0002,

        /// <summary>
        /// The comparison results in a match whenever possible, ignoring case and nonspacing characters.
        /// </summary>
        FL_LOOSE = 0x0004
    }

    /// <summary>
    /// The folder permissions that can be granted to a specified user.  
    /// </summary>
    public enum PidTagMemberRightsEnum : uint
    {
        /// <summary>
        /// No permission.
        /// </summary>
        None = 0,

        /// <summary>
        /// The server allow the specified user's client to read any Message object in the folder.
        /// </summary>
        ReadAny = 0x00000001,

        /// <summary>
        /// The server allow the specified user's client to create new Message objects in the folder.
        /// </summary>
        Create = 0x00000002,

        /// <summary>
        /// The server allow the specified user's client to modify a Message object that was created by that user in the folder.
        /// </summary>
        EditOwned = 0x00000008,

        /// <summary>
        /// The server allow the specified user's client to delete any Message object that was created by that user in the folder.
        /// </summary>
        DeleteOwned = 0x00000010,

        /// <summary>
        /// The server allow the specified user's client to modify any Message object in the folder.
        /// </summary>
        EditAny = 0x00000020,

        /// <summary>
        /// The server allow the specified user's client to delete any Message object in the folder.
        /// </summary>
        DeleteAny = 0x00000040,

        /// <summary>
        /// The server allow the specified user's client to create new folders within the folder.
        /// </summary>
        CreateSubFolder = 0x00000080,

        /// <summary>
        /// The server allow the specified user's client to modify properties set on the folder itself, including the folder permissions.
        /// </summary>
        FolderOwner = 0x00000100,

        /// <summary>
        /// The server include the specified user in any list of administrative contacts associated with the folder.
        /// </summary>
        FolderContact = 0x00000200,

        /// <summary>
        /// The server allow the specified user's client to see the folder in the folder hierarchy table and MUST allow the specified user's client to open the folder by using a RopOpenFolder ROP request.
        /// </summary>
        FolderVisible = 0x00000400,
        
        /// <summary>
        /// The server allow the specified user's client to retrieve brief information about the appointments on the calendar through the Availability Web Service Protocol.
        /// </summary>
        FreeBusySimple = 0x00000800,

        /// <summary>
        /// The server allow the specified user's client to retrieve detailed information about the appointments on the calendar through the Availability Web Service Protocol.
        /// </summary>
        FreeBusyDetailed = 0x00001000,

        /// <summary>
        /// All permission except FreeBusySimple and FreeBusyDetailed.
        /// </summary>
        FullPermission = 0x000007FB
    }
}