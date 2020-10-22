namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// The version of SUT.
    /// </summary>
    public enum SutVersion
    {
        /// <summary>
        /// The SUT is Exchange Server 2007 SP3
        /// </summary>
        ExchangeServer2007,

        /// <summary>
        /// The SUT is Exchange Server 2010 SP3
        /// </summary>
        ExchangeServer2010,

        /// <summary>
        /// The SUT is Exchange Server 2013
        /// </summary>
        ExchangeServer2013,

        /// <summary>
        /// The SUT is Exchange Server 2016
        /// </summary>
        ExchangeServer2016,

        /// <summary>
        /// The SUT is Exchange Server 2019
        /// </summary>
        ExchangeServer2019
    }

    /// <summary>
    ///  Command Name
    /// </summary>
    public enum CommandName
    {
        /// <summary>
        /// Sync command.
        /// </summary>
        Sync = 0,

        /// <summary>
        /// SendMail command.
        /// </summary>
        SendMail = 1,

        /// <summary>
        /// SmartForward command.
        /// </summary>
        SmartForward = 2,

        /// <summary>
        /// SmartReply command.
        /// </summary>
        SmartReply = 3,

        /// <summary>
        /// GetAttachment command.
        /// </summary>
        GetAttachment = 4,

        /// <summary>
        /// FolderSync command.
        /// </summary>
        FolderSync = 9,

        /// <summary>
        /// FolderCreate command.
        /// </summary>
        FolderCreate = 10,

        /// <summary>
        /// FolderDelete command.
        /// </summary>
        FolderDelete = 11,

        /// <summary>
        /// FolderUpdate command.
        /// </summary>
        FolderUpdate = 12,

        /// <summary>
        /// MoveItems command.
        /// </summary>
        MoveItems = 13,

        /// <summary>
        /// GetItemEstimate command.
        /// </summary>
        GetItemEstimate = 14,

        /// <summary>
        /// MeetingResponse command.
        /// </summary>
        MeetingResponse = 15,

        /// <summary>
        /// Search command.
        /// </summary>
        Search = 16,

        /// <summary>
        /// Settings command.
        /// </summary>
        Settings = 17,

        /// <summary>
        /// Ping command.
        /// </summary>
        Ping = 18,

        /// <summary>
        /// ItemOperations command.
        /// </summary>
        ItemOperations = 19,

        /// <summary>
        /// Provision command.
        /// </summary>
        Provision = 20,

        /// <summary>
        /// ResolveRecipients command.
        /// </summary>
        ResolveRecipients = 21,

        /// <summary>
        /// ValidateCert command.
        /// </summary>
        ValidateCert = 22,

        /// <summary>
        /// Autodiscover command.
        /// </summary>
        Autodiscover = 5,

        /// <summary>
        /// GetHierarchy command.
        /// </summary>
        GetHierarchy = 6,

        /// <summary>
        /// Find command.
        /// </summary>
        Find = 23,

        /// <summary>
        /// Not exist command.
        /// </summary>
        NotExist
    }

    /// <summary>
    /// command parameterName 
    /// </summary>
    public enum CmdParameterName
    {
        /// <summary>
        /// The parameter of AttachmentName.
        /// </summary>
        AttachmentName = 0,

        /// <summary>
        /// The parameter of CollectionId.
        /// </summary>
        CollectionId = 1,

        /// <summary>
        /// The parameter of CollectionName.
        /// </summary>
        CollectionName = 2,

        /// <summary>
        /// The parameter of ItemId.
        /// </summary>
        ItemId = 3,

        /// <summary>
        /// The parameter of LongId.
        /// </summary>
        LongId = 4,

        /// <summary>
        /// The parameter of ParentId.
        /// </summary>
        ParentId = 5,

        /// <summary>
        /// The parameter of Occurrence.
        /// </summary>
        Occurrence = 6,

        /// <summary>
        /// The parameter of Options.
        /// </summary>
        Options = 7,

        /// <summary>
        /// The parameter of User.
        /// </summary>
        User = 8,

        /// <summary>
        /// The parameter of SaveInSent.
        /// </summary>
        SaveInSent = 9
    }

    /// <summary>
    /// The store to search
    /// </summary>
    public enum SearchName
    {
        /// <summary>
        /// Search the mailbox
        /// </summary>
        Mailbox,

        /// <summary>
        /// Search a Windows SharePoint Services or UNC library
        /// </summary>
        DocumentLibrary,

        /// <summary>
        /// Search the Global Address List
        /// </summary>
        GAL
    }

    /// <summary>
    /// Specifies the type of the folder that was updated (renamed or moved) or added
    /// </summary>
    public enum FolderType : int
    {
        /// <summary>
        /// User-created folder (generic)
        /// </summary>
        UserCreatedGeneric = 1,

        /// <summary>
        /// Default Inbox folder
        /// </summary>
        Inbox = 2,

        /// <summary>
        /// Default Drafts folder
        /// </summary>
        Drafts = 3,

        /// <summary>
        /// Default Deleted Items folder
        /// </summary>
        DeletedItems = 4,

        /// <summary>
        /// Default Sent Items folder
        /// </summary>
        SentItems = 5,

        /// <summary>
        /// Default Outbox folder
        /// </summary>
        Outbox = 6,

        /// <summary>
        /// Default Tasks folder
        /// </summary>
        Tasks = 7,

        /// <summary>
        /// Default Calendar folder
        /// </summary>
        Calendar = 8,

        /// <summary>
        /// Default Contacts folder
        /// </summary>
        Contacts = 9,

        /// <summary>
        /// Default Notes folder
        /// </summary>
        Notes = 10,

        /// <summary>
        /// Default Journal folder
        /// </summary>
        Journal = 11,

        /// <summary>
        /// User-created Mail folder
        /// </summary>
        UserCreatedMail = 12,

        /// <summary>
        /// User-created Calendar folder
        /// </summary>
        UserCreatedCalendar = 13,

        /// <summary>
        /// User-created Contacts folder
        /// </summary>
        UserCreatedContacts = 14,

        /// <summary>
        /// User-created Tasks folder
        /// </summary>
        UserCreatedTasks = 15,

        /// <summary>
        /// User-created journal folder
        /// </summary>
        UserCreatedJournal = 16,

        /// <summary>
        /// User-created Notes folder
        /// </summary>
        UserCreatedNotes = 17,

        /// <summary>
        /// Unknown folder type
        /// </summary>
        Unknown = 18,

        /// <summary>
        /// Recipient information cache
        /// </summary>
        RecipientInformationCache = 19
    }

    /// <summary>
    /// Transport Type for HTTP or HTTPS
    /// </summary>
    public enum ProtocolTransportType
    {
        /// <summary>
        /// HTTP transport.
        /// </summary>
        HTTP,

        /// <summary>
        /// HTTPS transport.
        /// </summary>
        HTTPS
    }

    /// <summary>
    /// Content Type that indicate the body's format
    /// </summary>
    public enum ContentTypeEnum
    {
        /// <summary>
        /// WBXML format
        /// </summary>
        Wbxml,

        /// <summary>
        /// XML format
        /// </summary>
        Xml,

        /// <summary>
        /// HTML format
        /// </summary>
        Html
    }

    /// <summary>
    /// Query value type.
    /// </summary>
    public enum QueryValueType
    {
        /// <summary>
        /// Plain text.
        /// </summary>
        PlainText,

        /// <summary>
        /// Base64 encode.
        /// </summary>
        Base64,
    }

    /// <summary>
    /// Delivery method for Fetch element in ItemOperations.
    /// </summary>
    public enum DeliveryMethodForFetch
    {
        /// <summary>
        /// The inline method of delivering binary content is including data encoded with base64 encoding inside the WBXML. 
        /// </summary>
        Inline,

        /// <summary>
        /// The multipart method of delivering content is a multipart structure with the WBXML being the first part, and the requested data populating the subsequent parts.
        /// </summary>
        MultiPart
    }
}