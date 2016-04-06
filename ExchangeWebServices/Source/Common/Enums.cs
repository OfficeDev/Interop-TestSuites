namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// The protocol transport type which is used to transfer messages between the client and SUT.
    /// </summary>
    public enum TransportProtocol
    {
        /// <summary>
        /// The transport is SOAP over HTTP.
        /// </summary>
        HTTP,

        /// <summary>
        /// The transport is SOAP over HTTPS.
        /// </summary>
        HTTPS
    }

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
        /// The SUT is Exchange Server 2013 SP1
        /// </summary>
        ExchangeServer2013,

        /// <summary>
        /// The SUT is Exchange Server 2016
        /// </summary>
        ExchangeServer2016
    }

    /// <summary> 
    /// Indicates which type of storage is used for the item/folder represented by this Id. 
    /// </summary> 
    public enum IdStorageType : byte
    {
        /// <summary> 
        /// The Id represents an item or folder in a mailbox and it contains a primary SMTP address. 
        /// </summary> 
        MailboxItemSmtpAddressBased = 0,

        /// <summary> 
        /// The Id represents a folder in a PublicFolder store. 
        /// </summary>
        PublicFolder = 1,

        /// <summary> 
        /// The Id represents an item in a PublicFolder store. 
        /// </summary> 
        PublicFolderItem = 2,

        /// <summary> 
        /// The Id represents an item or folder in a mailbox and contains a mailbox GUID. 
        /// </summary> 
        MailboxItemMailboxGuidBased = 3,

        /// <summary> 
        /// The Id represents a conversation in a mailbox and contains a mailbox GUID. 
        /// </summary> 
        ConversationIdMailboxGuidBased = 4,

        /// <summary> 
        /// The Id represents (by objectGuid) an object in the Active Directory. 
        /// </summary> 
        ActiveDirectoryObject = 5
    }

    /// <summary> 
    /// Indicates any special processing to perform on an Id when deserializing it. 
    /// </summary> 
    public enum IdProcessingInstructionType : byte
    {
        /// <summary> 
        /// No special processing. The Id represents a PR_ENTRY_ID 
        /// </summary> 
        Normal = 0,

        /// <summary> 
        /// The Id represents an OccurenceStoreObjectId and therefore 
        /// must be deserialized as a StoreObjectId. 
        /// </summary> 
        Recurrence = 1,

        /// <summary> 
        /// The Id represents a series. 
        /// </summary> 
        Series = 2
    }
}