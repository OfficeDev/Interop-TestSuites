namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    /// <summary>
    /// The folder permissions that can be granted to a specified user.  
    /// </summary>
    public enum PidTagMemberRights : uint
    {
        /// <summary>
        /// No permission.
        /// </summary>
        None = 0,

        /// <summary>
        /// The server allows the specified user's client to read any Message object in the folder.
        /// </summary>
        ReadAny = 0x00000001,

        /// <summary>
        /// The server allows the specified user's client to create new Message objects in the folder.
        /// </summary>
        Create = 0x00000002,

        /// <summary>
        /// The server allows the specified user's client to modify a Message object that was created by that user in the folder.
        /// </summary>
        EditOwned = 0x00000008,

        /// <summary>
        /// The server allows the specified user's client to delete any Message object that was created by that user in the folder.
        /// </summary>
        DeleteOwned = 0x00000010,

        /// <summary>
        /// The server allows the specified user's client to modify any Message object in the folder.
        /// </summary>
        EditAny = 0x00000020,

        /// <summary>
        /// The server allows the specified user's client to delete any Message object in the folder.
        /// </summary>
        DeleteAny = 0x00000040,

        /// <summary>
        /// The server allows the specified user's client to create new folders within the folder.
        /// </summary>
        CreateSubFolder = 0x00000080,

        /// <summary>
        /// The server allows the specified user's client to modify properties set on the folder itself, including the folder permissions.
        /// </summary>
        FolderOwner = 0x00000100,

        /// <summary>
        /// The server include the specified user in any list of administrative contacts associated with the folder.
        /// </summary>
        FolderContact = 0x00000200,

        /// <summary>
        /// The server allows the specified user's client to see the folder in the folder hierarchy table and MUST allow the specified user's client to open the folder by using a RopOpenFolder ROP request.
        /// </summary>
        FolderVisible = 0x00000400,

        /// <summary>
        /// The server allows the specified user's client to retrieve brief information about the appointments on the calendar through the Availability Web Service Protocol.
        /// </summary>
        FreeBusySimple = 0x00000800,

        /// <summary>
        /// The server allows the specified user's client to retrieve detailed information about the appointments on the calendar through the Availability Web Service Protocol.
        /// </summary>
        FreeBusyDetailed = 0x00001000,

        /// <summary>
        /// All permission except FreeBusySimple and FreeBusyDetailed.
        /// </summary>
        FullPermission = 0x000007FB
    }

    /// <summary>
    /// The long id of properties.
    /// </summary>
    public enum PropertyLID : uint
    {
        /// <summary>
        /// The long id of PidLidSmartNoAttach.
        /// </summary>
        PidLidSmartNoAttach = 0x00008514,

        /// <summary>
        /// The long id of PidLidPrivate.
        /// </summary>
        PidLidPrivate = 0x00008506,

        /// <summary>
        /// The long id of PidLidSideEffects.
        /// </summary>
        PidLidSideEffects = 0x00008510,

        /// <summary>
        /// The long id of PidLidCommonStart.
        /// </summary>
        PidLidCommonStart = 0x00008516,

        /// <summary>
        /// The long id of PidLidCommonEnd.
        /// </summary>
        PidLidCommonEnd = 0x00008517,

        /// <summary>
        /// The long id of PidLidCategories.
        /// </summary>
        PidLidCategories = 0x00009000,

        /// <summary>
        /// The long id of PidLidClassification.
        /// </summary>
        PidLidClassification = 0x000085B6,

        /// <summary>
        /// The long id of PidLidClassificationDescription.
        /// </summary>
        PidLidClassificationDescription = 0x000085B7,

        /// <summary>
        /// The long id of PidLidClassified.
        /// </summary>
        PidLidClassified = 0x000085B5,

        /// <summary>
        /// The long id of PidLidInfoPathFormName.
        /// </summary>
        PidLidInfoPathFormName = 0x85B1,

        /// <summary>
        /// The long id of PidLidAgingDontAgeMe.
        /// </summary>
        PidLidAgingDontAgeMe = 0x0000850E,

        /// <summary>
        /// The long id of PidLidCurrentVersion.
        /// </summary>
        PidLidCurrentVersion = 0x00008552,

        /// <summary>
        /// The long id of PidLidCurrentVersionName.
        /// </summary>
        PidLidCurrentVersionName = 0x00008554
    }

    /// <summary>
    /// The flag indicate the test cases expect to get which object type's properties(message's properties or attachment's properties).
    /// </summary>
    public enum GetPropertiesFlags
    {
        /// <summary>
        /// Not get any property.
        /// </summary>
        None,

        /// <summary>
        /// Get message's properties.
        /// </summary>
        MessageProperties,

        /// <summary>
        /// Get attachment's properties.
        /// </summary>
        AttachmentProperties
    }

    /// <summary>
    /// The value of flags in property PidTagAttachMethod
    /// </summary>
    public enum PidTagAttachMethodFlags
    {
        /// <summary>
        /// The attachment has just been created.
        /// </summary>
        afNone = 0x00000000,

        /// <summary>
        /// The PidTagAttachDataBinary property contains the attachment data.
        /// </summary>
        afByValue = 0x00000001,

        /// <summary>
        /// The PidTagAttachLongPathname property contains a fully qualified path identifying the attachment to recipients with access to a common file server.
        /// </summary>
        afByReference = 0x00000002,

        /// <summary>
        /// The PidTagAttachLongPathname property contains a fully qualified path identifying the attachment.
        /// </summary>
        afByReferenceOnly = 0x00000004,

        /// <summary>
        /// The attachment is an embedded message that is accessed via the RopOpenEmbeddedMessage ROP.
        /// </summary>
        afEmbeddedMessage = 0x00000005,

        /// <summary>
        /// The PidTagAttachDataObject property contains data in an application-specific format.
        /// </summary>
        afStorage = 0x00000006,
    }
}