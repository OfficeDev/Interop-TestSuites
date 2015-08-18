namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    /// <summary>
    /// The request type of MS-OXCMAPIHTTP
    /// </summary>
    public enum RequestType
    {
        /// <summary>
        /// The connect request type
        /// </summary>
        Connect,

        /// <summary>
        /// The Execute request type
        /// </summary>
        Execute,

        /// <summary>
        /// The Disconnect type
        /// </summary>
        Disconnect,

        /// <summary>
        /// The NotificationWait request type
        /// </summary>
        NotificationWait,

        /// <summary>
        /// The PING request type
        /// </summary>
        PING,

        /// <summary>
        /// The Bind request type
        /// </summary>
        Bind,

        /// <summary>
        /// The Unbind request type
        /// </summary>
        Unbind,

        /// <summary>
        /// The CompareMIds request type
        /// </summary>
        CompareMIds,

        /// <summary>
        /// The DNToMId request type
        /// </summary>
        DNToMId,

        /// <summary>
        /// The GetMatches request type
        /// </summary>
        GetMatches,

        /// <summary>
        /// The GetPropList request type
        /// </summary>
        GetPropList,

        /// <summary>
        /// The GetProps request type
        /// </summary>
        GetProps,

        /// <summary>
        /// The GetSpecialTable request type
        /// </summary>
        GetSpecialTable,

        /// <summary>
        /// The GetTemplateInfo request type
        /// </summary>
        GetTemplateInfo,

        /// <summary>
        /// The ModLinkAtt request type
        /// </summary>
        ModLinkAtt,

        /// <summary>
        /// The ModProps request type
        /// </summary>
        ModProps,

        /// <summary>
        /// The QueryColumns request type
        /// </summary>
        QueryColumns,

        /// <summary>
        /// The QueryRows request type
        /// </summary>
        QueryRows,

        /// <summary>
        /// The ResolveNames request type
        /// </summary>
        ResolveNames,

        /// <summary>
        /// The ResortRestriction request type
        /// </summary>
        ResortRestriction,

        /// <summary>
        /// The SeekEntries request type
        /// </summary>
        SeekEntries,

        /// <summary>
        /// The UpdateStat request type
        /// </summary>
        UpdateStat,

        /// <summary>
        /// The GetMailboxUrl request type
        /// </summary>
        GetMailboxUrl,

        /// <summary>
        /// The GetAddressBookUrl request type
        /// </summary>
        GetAddressBookUrl
    }

    /// <summary>
    /// The server endpoint type
    /// </summary>
    public enum ServerEndpoint
    {
        /// <summary>
        /// The endpoint used for Mailbox server
        /// </summary>
        MailboxServerEndpoint,

        /// <summary>
        /// The endpoint used for address book server
        /// </summary>
        AddressBookServerEndpoint
    }

    #region Address Book
    /// <summary>
    /// NspiGetSpecialTable flag values are used as bit flags in the NspiGetSpecialTable method to specify optional behavior to a server. 
    /// </summary>
    public enum NspiGetSpecialTableFlags
    {
        /// <summary>
        /// Specify none to 0.
        /// </summary>
        None = 0x00000000,

        /// <summary>
        /// Specify that the server MUST return the table of the available address creation templates. 
        /// Specify that this flag causes the server to ignore the NspiUnicodeStrings flag.
        /// </summary>
        NspiAddressCreationTemplates = 0x00000002,

        /// <summary>
        /// Specifies that the server MUST return all strings as Unicode representations 
        /// rather than as multibyte strings in the client's code page. 
        /// </summary>
        NspiUnicodeStrings = 0x00000004,
    }

    /// <summary>
    /// The NspiGetTemplateInfo flag values are used as bit flags in the NspiGetTemplateInfo method to specify optional behavior to a server. 
    /// </summary>
    public enum NspiGetTemplateInfoFlags
    {
        /// <summary>
        /// Specifies that the server is to return the value that represents a template.
        /// </summary>
        TI_TEMPLATE = 0x00000001,

        /// <summary>
        /// Specifies that the server is to return the value of the script that is associated with a template.
        /// </summary>
        TI_SCRIPT = 0x00000004,

        /// <summary>
        /// Specifies that the server is to return the e-mail type that is associated with a template.
        /// </summary>
        TI_EMT = 0x00000010,

        /// <summary>
        /// Specifies that the server is to return the name of the help file that is associated with a template.
        /// </summary>
        TI_HELPFILE_NAME = 0x00000020,

        /// <summary>
        /// Specifies that the server is to return the contents of the help file that is associated with a template.
        /// </summary>
        TI_HELPFILE_CONTENTS = 0x00000040,
    }

    /// <summary>
    /// The values are used to specify a specific sort orders for tables. 
    /// </summary>
    public enum TableSortOrders
    {
        /// <summary>
        /// The table is sorted ascending on the PidTagDisplayName property, as specified in [MS-OXCFOLD] section 2.2.2.2.2.3. 
        /// All Exchange NSPI servers MUST support this sort order for at least one LCID.
        /// </summary>
        SortTypeDisplayName = 0x00000000,

        /// <summary>
        /// The table is sorted ascending on the PidTagAddressBookPhoneticDisplayName property, as specified in [MS-OXOABK] section 2.2.3.9. 
        /// Exchange NSPI servers SHOULD support this sort order. Exchange NSPI servers MAY support this only for some LCIDs.
        /// </summary>
        SortTypePhoneticDisplayName = 0x00000003,

        /// <summary>
        /// The table is sorted ascending on the PidTagDisplayName property. 
        /// The client MUST set this value only when using the NspiGetMatches method to open a non-writable table on an object-valued property.
        /// </summary>
        SortTypeDisplayName_RO = 0x000003E8,

        /// <summary>
        /// The table is sorted ascending on the PidTagDisplayName property. 
        /// The client MUST set this value only when using the NspiGetMatches method to open a writable table on an object-valued property.
        /// </summary>
        SortTypeDisplayName_W = 0x000003E9,
    }

    /// <summary>
    /// The values are used to specify display types. 
    /// </summary>
    public enum DisplayTypeValues : uint
    {
        /// <summary>
        /// A typical messaging user.
        /// </summary>
        DT_MAILUSER = 0x00000000,

        /// <summary>
        /// A distribution list.
        /// </summary>
        DT_DISTLIST = 0x00000001,

        /// <summary>
        /// A forum, such as a bulletin board service or a public or shared folder.
        /// </summary>
        DT_FORUM = 0x00000002,

        /// <summary>
        /// An automated agent, such as Quote-Of-The-Day or a weather chart display.
        /// </summary>
        DT_AGENT = 0x00000003,

        /// <summary>
        /// An Address Book object defined for a large group, such as helpdesk, accounting, coordinator, 
        /// or department. Department objects usually have this display type.
        /// </summary>
        DT_ORGANIZATION = 0x00000004,

        /// <summary>
        /// A private, personally administered distribution list.
        /// </summary>
        DT_PRIVATE_DISTLIST = 0x00000005,

        /// <summary>
        /// An Address Book object known to be from a foreign or remote messaging system.
        /// </summary>
        DT_REMOTE_MAILUSER = 0x00000006,

        /// <summary>
        /// An address book hierarchy table container. 
        /// An Exchange NSPI server MUST NOT return this display type except as part of an EntryID of an object in the address book hierarchy table.
        /// </summary>
        DT_CONTAINER = 0x00000100,

        /// <summary>
        /// A display template object. An Exchange NSPI server MUST NOT return this display type.
        /// </summary>
        DT_TEMPLATE = 0x00000101,

        /// <summary>
        /// An address creation template. 
        /// An Exchange NSPI server MUST NOT return this display type except as part of an EntryID of an object in the Address Creation Table.
        /// </summary>
        DT_ADDRESS_TEMPLATE = 0x00000102,

        /// <summary>
        /// A search template. An Exchange NSPI server MUST NOT return this display type. 
        /// </summary>
        DT_SEARCH = 0x00000200
    }

    /// <summary>
    /// The language code identifier (LCID) specified in this section is associated with the minimal required sort order for Unicode strings. 
    /// </summary>
    public enum DefaultLCID
    {
        /// <summary>
        /// Represents the default LCID that is used for comparison of Unicode string representations.
        /// </summary>
        NSPI_DEFAULT_LOCALE = 0x00000409,
    }

    /// <summary>
    /// The required code pages listed in this section are associated with the string handling in the Exchange Server NSPI Protocol, 
    /// and they appear in input parameters to methods in the Exchange Server NSPI Protocol. 
    /// </summary>
    public enum RequiredCodePages : uint
    {
        /// <summary>
        /// Represents the Teletex code page.
        /// </summary>
        CP_TELETEX = 0x00004F25,

        /// <summary>
        /// Represents the Unicode code page.
        /// </summary>
        CP_WINUNICODE = 0x000004B0,
    }

    /// <summary>
    /// The positioning Minimal Entry IDs are used to specify objects in the address book as a function of their positions in tables.
    /// </summary>
    public enum MinimalEntryIDs
    {
        /// <summary>
        /// Specifies the position before the first row in the current address book container.
        /// </summary>
        MID_BEGINNING_OF_TABLE = 0x00000000,

        /// <summary>
        /// Specifies the position after the last row in the current address book container.
        /// </summary>
        MID_END_OF_TABLE = 0x00000002,

        /// <summary>
        /// Specifies the current position in a table. This Minimal Entry ID is only valid in the NspiUpdateStat method. 
        /// In all other cases, it is an invalid Minimal Entry ID, guaranteed to not specify any object in the address book.
        /// </summary>
        MID_CURRENT = 0x00000001,
    }

    /// <summary>
    /// The property type values are used to specify property types.
    /// </summary>
    public enum PropertyTypeValues : uint
    {
        /// <summary>
        /// 2 bytes, a 16-bit integer.
        /// </summary>
        PtypInteger16 = 0x00000002,

        /// <summary>
        /// 4 bytes, a 32-bit integer.
        /// </summary>
        PtypInteger32 = 0x00000003,

        /// <summary>
        /// 1 byte, restricted to 1 or 0.
        /// </summary>
        PtypBoolean = 0x0000000B,

        /// <summary>
        /// Variable size, a string of multi-byte characters in externally specified encoding with terminating null character (single 0 byte).
        /// </summary>
        PtypString8 = 0x0000001E,

        /// <summary>
        /// Variable size, a COUNT followed by that many bytes.
        /// </summary>
        PtypBinary = 0x00000102,

        /// <summary>
        /// Variable size, a string of Unicode characters in UTF-16LE encoding with terminating null character (2 bytes of zero).
        /// </summary>
        PtypString = 0x0000001F,

        /// <summary>
        /// 16 bytes, a GUID with Data1, Data2, and Data3 fields in little-endian format.
        /// </summary>
        PtypGuid = 0x00000048,

        /// <summary>
        /// 8 bytes, a 64-bit integer representing the number of 100-nanosecond intervals since January 1, 1601.
        /// </summary>
        PtypTime = 0x00000040,

        /// <summary>
        /// 4 bytes, a 32-bit integer encoding error information.
        /// </summary>
        PtypErrorCode = 0x0000000A,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypInteger16 values.
        /// </summary>
        PtypMultipleInteger16 = 0x00001002,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypInteger32 values.
        /// </summary>
        PtypMultipleInteger32 = 0x00001003,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypString8 values.
        /// </summary>
        PtypMultipleString8 = 0x0000101E,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypBinary values.
        /// </summary>
        PtypMultipleBinary = 0x00001102,

        /// <summary>
        /// Variable size, a COUNT followed by that PtypString values.
        /// </summary>
        PtypMultipleString = 0x0000101F,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypGuid values.
        /// </summary>
        PtypMultipleGuid = 0x00001048,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypTime values.
        /// </summary>
        PtypMultipleTime = 0x00001040,

        /// <summary>
        /// Single 32-bit value, referencing an address list. 
        /// </summary>
        PtypEmbeddedTable = 0x0000000D,

        /// <summary>
        /// Clients MUST NOT specify this property type in any method's input parameters.
        /// The server MUST specify this property type in any method's output parameters to indicate that a property has a value that cannot be expressed in the Exchange Server NSPI Protocol.
        /// </summary>
        PtypNull = 0x00000001,

        /// <summary>
        /// Clients specify this property type in a method's input parameter to indicate that the client will accept any property type the server chooses when returning propvalues.
        /// Servers MUST NOT specify this property type in any method's output parameters except the method NspiGetIDsFromNames.
        /// </summary>
        PtypUnspecified = 0x00000000
    }

    /// <summary>
    /// Ambiguous name resolution (ANR) Minimal Entry IDs are used to specify the outcome of the ANR process. 
    /// </summary>
    public enum ANRMinEntryIDs
    {
        /// <summary>
        /// The ANR process is unable to map a string to any objects in the address book.
        /// </summary>
        MID_UNRESOLVED = 0x00000000,

        /// <summary>
        /// The ANR process maps a string to multiple objects in the address book.
        /// </summary>
        MID_AMBIGUOUS = 0x0000001,

        /// <summary>
        /// The ANR process maps a string to a single object in the address book.
        /// </summary>
        MID_RESOLVED = 0x0000002,
    }

    /// <summary>
    /// The property flag values that are used as bit flags in GetPropList, GetProps, and QueryRows methods to specify optional behavior to a server.
    /// </summary>
    public enum RetrievePropertyFlags
    {
        /// <summary>
        /// Client requires that the server MUST NOT include proptags with the PtypEmbeddedTable property type 
        /// in any lists of proptags that the server creates on behalf of the client.
        /// </summary>
        fSkipObjects = 0x00000001,

        /// <summary>
        /// Client requires that the server MUST return Entry ID values in Ephemeral Entry ID form.
        /// </summary>
        fEphID = 0x00000002,
    }

    /// <summary>
    /// The QueryColumns flag value is used as a bit flag in the QueryColumns method to specify optional behavior to a server. 
    /// </summary>
    public enum NspiQueryColumnsFlag : uint
    {
        /// <summary>
        /// Specifies that the server MUST return all proptags that specify values with string 
        /// representations as having the PtypString property type.
        /// </summary>
        NspiUnicodeProptypes = 0x80000000,
    }

    /// <summary>
    /// The property ID values are used to specify property ID.
    /// </summary>
    public enum PropertyID : uint
    {
        /// <summary>
        /// The property ID of PidTagDisplayName.
        /// </summary>
        PidTagDisplayName = 0x3001,

        /// <summary>
        /// The property ID of PidTagAddressBookX509Certificate.
        /// </summary>
        PidTagAddressBookX509Certificate = 0x8c6a,

        /// <summary>
        /// The property ID of PidTagUserX509Certificate.
        /// </summary>
        PidTagUserX509Certificate = 0x3A70,

        /// <summary>
        /// The property ID of PidTagEntryId.
        /// </summary>
        PidTagEntryId = 0x0FFF,

        /// <summary>
        /// The property ID of PidTagDisplayType.
        /// </summary>
        PidTagDisplayType = 0x3900,

        /// <summary>
        /// The property ID of PidTagAddressBookMember.
        /// </summary>
        PidTagAddressBookMember = 0x8009,

        /// <summary>
        /// The property ID of PidTagAddressBookPublicDelegates.
        /// </summary>
        PidTagAddressBookPublicDelegates = 0x8015,
    }    
    #endregion
}