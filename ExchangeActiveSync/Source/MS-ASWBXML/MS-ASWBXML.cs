namespace Microsoft.Protocols.TestSuites.MS_ASWBXML
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Xml;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The implementation of MS-ASWBXML
    /// </summary>
    public class MS_ASWBXML
    {
        /// <summary>
        /// WBXML version
        /// </summary>
        private const byte VersionByte = 0x03;

        /// <summary>
        /// Public Identifier
        /// </summary>
        private const byte PublicIdentifierByte = 0x01;

        /// <summary>
        /// Encoding. 0x6A == UTF-8
        /// </summary>
        private const byte CharsetByte = 0x6A;

        /// <summary>
        /// String table length. This is not used in MS-ASWBXML, so this value is always 0.
        /// </summary>
        private const byte StringTableLengthByte = 0x00;

        /// <summary>
        /// XmlDocument that contain the xml
        /// </summary>
        private XmlDocument xmlDoc = new XmlDocument();

        /// <summary>
        /// Code pages.
        /// </summary>
        private CodePage[] codePages;

        /// <summary>
        /// Current code page.
        /// </summary>
        private int currentCodePage = 0;

        /// <summary>
        /// Default code page.
        /// </summary>
        private int defaultCodePage = -1;

        /// <summary> 
        /// The DataCollection in encoding process 
        /// </summary> 
        private Dictionary<string, int> encodeDataCollection = new Dictionary<string, int>();

        /// <summary> 
        /// The DataCollection in decoding process 
        /// </summary> 
        private Dictionary<string, int> decodeDataCollection = new Dictionary<string, int>();

        /// <summary>
        /// An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.
        /// </summary>
        private ITestSite site;

        /// <summary>
        /// Initializes a new instance of the MS_ASWBXML class.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public MS_ASWBXML(ITestSite site)
        {
            this.site = site;

            // Loads up code pages. There are 26 code pages as per MS-ASWBXML
            this.codePages = new CodePage[26];

            // Code Page 0: AirSync
            this.codePages[0] = new CodePage { Namespace = "AirSync", Xmlns = "airsync" };

            this.codePages[0].AddToken(0x05, "Sync");
            this.codePages[0].AddToken(0x06, "Responses");
            this.codePages[0].AddToken(0x07, "Add");
            this.codePages[0].AddToken(0x08, "Change");
            this.codePages[0].AddToken(0x09, "Delete");
            this.codePages[0].AddToken(0x0A, "Fetch");
            this.codePages[0].AddToken(0x0B, "SyncKey");
            this.codePages[0].AddToken(0x0C, "ClientId");
            this.codePages[0].AddToken(0x0D, "ServerId");
            this.codePages[0].AddToken(0x0E, "Status");
            this.codePages[0].AddToken(0x0F, "Collection");
            this.codePages[0].AddToken(0x10, "Class");
            this.codePages[0].AddToken(0x12, "CollectionId");
            this.codePages[0].AddToken(0x13, "GetChanges");
            this.codePages[0].AddToken(0x14, "MoreAvailable");
            this.codePages[0].AddToken(0x15, "WindowSize");
            this.codePages[0].AddToken(0x16, "Commands");
            this.codePages[0].AddToken(0x17, "Options");
            this.codePages[0].AddToken(0x18, "FilterType");
            this.codePages[0].AddToken(0x1B, "Conflict");
            this.codePages[0].AddToken(0x1C, "Collections");
            this.codePages[0].AddToken(0x1D, "ApplicationData");
            this.codePages[0].AddToken(0x1E, "DeletesAsMoves");
            this.codePages[0].AddToken(0x20, "Supported");
            this.codePages[0].AddToken(0x21, "SoftDelete");
            this.codePages[0].AddToken(0x22, "MIMESupport");
            this.codePages[0].AddToken(0x23, "MIMETruncation");
            this.codePages[0].AddToken(0x24, "Wait");
            this.codePages[0].AddToken(0x25, "Limit");
            this.codePages[0].AddToken(0x26, "Partial");
            this.codePages[0].AddToken(0x27, "ConversationMode");
            this.codePages[0].AddToken(0x28, "MaxItems");
            this.codePages[0].AddToken(0x29, "HeartbeatInterval");

            // Code Page 1: Contacts
            this.codePages[1] = new CodePage { Namespace = "Contacts", Xmlns = "contacts" };

            this.codePages[1].AddToken(0x05, "Anniversary");
            this.codePages[1].AddToken(0x06, "AssistantName");
            this.codePages[1].AddToken(0x07, "AssistantPhoneNumber");
            this.codePages[1].AddToken(0x08, "Birthday");
            this.codePages[1].AddToken(0x0C, "Business2PhoneNumber");
            this.codePages[1].AddToken(0x0D, "BusinessAddressCity");
            this.codePages[1].AddToken(0x0E, "BusinessAddressCountry");
            this.codePages[1].AddToken(0x0F, "BusinessAddressPostalCode");
            this.codePages[1].AddToken(0x10, "BusinessAddressState");
            this.codePages[1].AddToken(0x11, "BusinessAddressStreet");
            this.codePages[1].AddToken(0x12, "BusinessFaxNumber");
            this.codePages[1].AddToken(0x13, "BusinessPhoneNumber");
            this.codePages[1].AddToken(0x14, "CarPhoneNumber");
            this.codePages[1].AddToken(0x15, "Categories");
            this.codePages[1].AddToken(0x16, "Category");
            this.codePages[1].AddToken(0x17, "Children");
            this.codePages[1].AddToken(0x18, "Child");
            this.codePages[1].AddToken(0x19, "CompanyName");
            this.codePages[1].AddToken(0x1A, "Department");
            this.codePages[1].AddToken(0x1B, "Email1Address");
            this.codePages[1].AddToken(0x1C, "Email2Address");
            this.codePages[1].AddToken(0x1D, "Email3Address");
            this.codePages[1].AddToken(0x1E, "FileAs");
            this.codePages[1].AddToken(0x1F, "FirstName");
            this.codePages[1].AddToken(0x20, "Home2PhoneNumber");
            this.codePages[1].AddToken(0x21, "HomeAddressCity");
            this.codePages[1].AddToken(0x22, "HomeAddressCountry");
            this.codePages[1].AddToken(0x23, "HomeAddressPostalCode");
            this.codePages[1].AddToken(0x24, "HomeAddressState");
            this.codePages[1].AddToken(0x25, "HomeAddressStreet");
            this.codePages[1].AddToken(0x26, "HomeFaxNumber");
            this.codePages[1].AddToken(0x27, "HomePhoneNumber");
            this.codePages[1].AddToken(0x28, "JobTitle");
            this.codePages[1].AddToken(0x29, "LastName");
            this.codePages[1].AddToken(0x2A, "MiddleName");
            this.codePages[1].AddToken(0x2B, "MobilePhoneNumber");
            this.codePages[1].AddToken(0x2C, "OfficeLocation");
            this.codePages[1].AddToken(0x2D, "OtherAddressCity");
            this.codePages[1].AddToken(0x2E, "OtherAddressCountry");
            this.codePages[1].AddToken(0x2F, "OtherAddressPostalCode");
            this.codePages[1].AddToken(0x30, "OtherAddressState");
            this.codePages[1].AddToken(0x31, "OtherAddressStreet");
            this.codePages[1].AddToken(0x32, "PagerNumber");
            this.codePages[1].AddToken(0x33, "RadioPhoneNumber");
            this.codePages[1].AddToken(0x34, "Spouse");
            this.codePages[1].AddToken(0x35, "Suffix");
            this.codePages[1].AddToken(0x36, "Title");
            this.codePages[1].AddToken(0x37, "WebPage");
            this.codePages[1].AddToken(0x38, "YomiCompanyName");
            this.codePages[1].AddToken(0x39, "YomiFirstName");
            this.codePages[1].AddToken(0x3A, "YomiLastName");
            this.codePages[1].AddToken(0x3C, "Picture");
            this.codePages[1].AddToken(0x3D, "Alias");
            this.codePages[1].AddToken(0x3E, "WeightedRank");

            // Code Page 2: Email
            this.codePages[2] = new CodePage { Namespace = "Email", Xmlns = "email" };
            this.codePages[2].AddToken(0x0F, "DateReceived");
            this.codePages[2].AddToken(0x11, "DisplayTo");
            this.codePages[2].AddToken(0x12, "Importance");
            this.codePages[2].AddToken(0x13, "MessageClass");
            this.codePages[2].AddToken(0x14, "Subject");
            this.codePages[2].AddToken(0x15, "Read");
            this.codePages[2].AddToken(0x16, "To");
            this.codePages[2].AddToken(0x17, "Cc");
            this.codePages[2].AddToken(0x18, "From");
            this.codePages[2].AddToken(0x19, "ReplyTo");
            this.codePages[2].AddToken(0x1A, "AllDayEvent");
            this.codePages[2].AddToken(0x1B, "Categories");
            this.codePages[2].AddToken(0x1C, "Category");
            this.codePages[2].AddToken(0x1D, "DtStamp");
            this.codePages[2].AddToken(0x1E, "EndTime");
            this.codePages[2].AddToken(0x1F, "InstanceType");
            this.codePages[2].AddToken(0x20, "BusyStatus");
            this.codePages[2].AddToken(0x21, "Location");
            this.codePages[2].AddToken(0x22, "MeetingRequest");
            this.codePages[2].AddToken(0x23, "Organizer");
            this.codePages[2].AddToken(0x24, "RecurrenceId");
            this.codePages[2].AddToken(0x25, "Reminder");
            this.codePages[2].AddToken(0x26, "ResponseRequested");
            this.codePages[2].AddToken(0x27, "Recurrences");
            this.codePages[2].AddToken(0x28, "Recurrence");
            this.codePages[2].AddToken(0x29, "Type");
            this.codePages[2].AddToken(0x2A, "Until");
            this.codePages[2].AddToken(0x2B, "Occurrences");
            this.codePages[2].AddToken(0x2C, "Interval");
            this.codePages[2].AddToken(0x2D, "DayOfWeek");
            this.codePages[2].AddToken(0x2E, "DayOfMonth");
            this.codePages[2].AddToken(0x2F, "WeekOfMonth");
            this.codePages[2].AddToken(0x30, "MonthOfYear");
            this.codePages[2].AddToken(0x31, "StartTime");
            this.codePages[2].AddToken(0x32, "Sensitivity");
            this.codePages[2].AddToken(0x33, "TimeZone");
            this.codePages[2].AddToken(0x34, "GlobalObjId");
            this.codePages[2].AddToken(0x35, "ThreadTopic");
            this.codePages[2].AddToken(0x39, "InternetCPID");
            this.codePages[2].AddToken(0x3A, "Flag");
            this.codePages[2].AddToken(0x3B, "Status");
            this.codePages[2].AddToken(0x3C, "ContentClass");
            this.codePages[2].AddToken(0x3D, "FlagType");
            this.codePages[2].AddToken(0x3E, "CompleteTime");
            this.codePages[2].AddToken(0x3F, "DisallowNewTimeProposal");

            // Code Page 3: AirNotify
            this.codePages[3] = new CodePage { Namespace = string.Empty, Xmlns = string.Empty };

            // Code Page 4: Calendar
            this.codePages[4] = new CodePage { Namespace = "Calendar", Xmlns = "calendar" };

            this.codePages[4].AddToken(0x05, "Timezone");
            this.codePages[4].AddToken(0x06, "AllDayEvent");
            this.codePages[4].AddToken(0x07, "Attendees");
            this.codePages[4].AddToken(0x08, "Attendee");
            this.codePages[4].AddToken(0x09, "Email");
            this.codePages[4].AddToken(0x0A, "Name");
            this.codePages[4].AddToken(0x0D, "BusyStatus");
            this.codePages[4].AddToken(0x0E, "Categories");
            this.codePages[4].AddToken(0x0F, "Category");
            this.codePages[4].AddToken(0x11, "DtStamp");
            this.codePages[4].AddToken(0x12, "EndTime");
            this.codePages[4].AddToken(0x13, "Exception");
            this.codePages[4].AddToken(0x14, "Exceptions");
            this.codePages[4].AddToken(0x15, "Deleted");
            this.codePages[4].AddToken(0x16, "ExceptionStartTime");
            this.codePages[4].AddToken(0x17, "Location");
            this.codePages[4].AddToken(0x18, "MeetingStatus");
            this.codePages[4].AddToken(0x19, "OrganizerEmail");
            this.codePages[4].AddToken(0x1A, "OrganizerName");
            this.codePages[4].AddToken(0x1B, "Recurrence");
            this.codePages[4].AddToken(0x1C, "Type");
            this.codePages[4].AddToken(0x1D, "Until");
            this.codePages[4].AddToken(0x1E, "Occurrences");
            this.codePages[4].AddToken(0x1F, "Interval");
            this.codePages[4].AddToken(0x20, "DayOfWeek");
            this.codePages[4].AddToken(0x21, "DayOfMonth");
            this.codePages[4].AddToken(0x22, "WeekOfMonth");
            this.codePages[4].AddToken(0x23, "MonthOfYear");
            this.codePages[4].AddToken(0x24, "Reminder");
            this.codePages[4].AddToken(0x25, "Sensitivity");
            this.codePages[4].AddToken(0x26, "Subject");
            this.codePages[4].AddToken(0x27, "StartTime");
            this.codePages[4].AddToken(0x28, "UID");
            this.codePages[4].AddToken(0x29, "AttendeeStatus");
            this.codePages[4].AddToken(0x2A, "AttendeeType");
            this.codePages[4].AddToken(0x33, "DisallowNewTimeProposal");
            this.codePages[4].AddToken(0x34, "ResponseRequested");
            this.codePages[4].AddToken(0x35, "AppointmentReplyTime");
            this.codePages[4].AddToken(0x36, "ResponseType");
            this.codePages[4].AddToken(0x37, "CalendarType");
            this.codePages[4].AddToken(0x38, "IsLeapMonth");
            this.codePages[4].AddToken(0x39, "FirstDayOfWeek");
            this.codePages[4].AddToken(0x3A, "OnlineMeetingConfLink");
            this.codePages[4].AddToken(0x3B, "OnlineMeetingExternalLink");
            this.codePages[4].AddToken(0x3C, "ClientUid");

            // Code Page 5: Move
            this.codePages[5] = new CodePage { Namespace = "Move", Xmlns = "move" };

            this.codePages[5].AddToken(0x05, "MoveItems");
            this.codePages[5].AddToken(0x06, "Move");
            this.codePages[5].AddToken(0x07, "SrcMsgId");
            this.codePages[5].AddToken(0x08, "SrcFldId");
            this.codePages[5].AddToken(0x09, "DstFldId");
            this.codePages[5].AddToken(0x0A, "Response");
            this.codePages[5].AddToken(0x0B, "Status");
            this.codePages[5].AddToken(0x0C, "DstMsgId");

            // Code Page 6: GetItemEstimate
            this.codePages[6] = new CodePage { Namespace = "GetItemEstimate", Xmlns = "getitemestimate" };

            this.codePages[6].AddToken(0x05, "GetItemEstimate");
            this.codePages[6].AddToken(0x07, "Collections");
            this.codePages[6].AddToken(0x08, "Collection");
            this.codePages[6].AddToken(0x09, "Class");
            this.codePages[6].AddToken(0x0A, "CollectionId");
            this.codePages[6].AddToken(0x0C, "Estimate");
            this.codePages[6].AddToken(0x0D, "Response");
            this.codePages[6].AddToken(0x0E, "Status");

            // Code Page 7: FolderHierarchy
            this.codePages[7] = new CodePage { Namespace = "FolderHierarchy", Xmlns = "folderhierarchy" };

            this.codePages[7].AddToken(0x05, "Folders");
            this.codePages[7].AddToken(0x06, "Folder");
            this.codePages[7].AddToken(0x07, "DisplayName");
            this.codePages[7].AddToken(0x08, "ServerId");
            this.codePages[7].AddToken(0x09, "ParentId");
            this.codePages[7].AddToken(0x0A, "Type");
            this.codePages[7].AddToken(0x0C, "Status");
            this.codePages[7].AddToken(0x0E, "Changes");
            this.codePages[7].AddToken(0x0F, "Add");
            this.codePages[7].AddToken(0x10, "Delete");
            this.codePages[7].AddToken(0x11, "Update");
            this.codePages[7].AddToken(0x12, "SyncKey");
            this.codePages[7].AddToken(0x13, "FolderCreate");
            this.codePages[7].AddToken(0x14, "FolderDelete");
            this.codePages[7].AddToken(0x15, "FolderUpdate");
            this.codePages[7].AddToken(0x16, "FolderSync");
            this.codePages[7].AddToken(0x17, "Count");

            // Code Page 8: MeetingResponse
            this.codePages[8] = new CodePage { Namespace = "MeetingResponse", Xmlns = "meetingresponse" };

            this.codePages[8].AddToken(0x05, "CalendarId");
            this.codePages[8].AddToken(0x06, "CollectionId");
            this.codePages[8].AddToken(0x07, "MeetingResponse");
            this.codePages[8].AddToken(0x08, "RequestId");
            this.codePages[8].AddToken(0x09, "Request");
            this.codePages[8].AddToken(0x0A, "Result");
            this.codePages[8].AddToken(0x0B, "Status");
            this.codePages[8].AddToken(0x0C, "UserResponse");
            this.codePages[8].AddToken(0x0E, "InstanceId");
            this.codePages[8].AddToken(0x12, "SendResponse");

            // Code Page 9: Tasks
            this.codePages[9] = new CodePage { Namespace = "Tasks", Xmlns = "tasks" };

            this.codePages[9].AddToken(0x08, "Categories");
            this.codePages[9].AddToken(0x09, "Category");
            this.codePages[9].AddToken(0x0A, "Complete");
            this.codePages[9].AddToken(0x0B, "DateCompleted");
            this.codePages[9].AddToken(0x0C, "DueDate");
            this.codePages[9].AddToken(0x0D, "UtcDueDate");
            this.codePages[9].AddToken(0x0E, "Importance");
            this.codePages[9].AddToken(0x0F, "Recurrence");
            this.codePages[9].AddToken(0x10, "Type");
            this.codePages[9].AddToken(0x11, "Start");
            this.codePages[9].AddToken(0x12, "Until");
            this.codePages[9].AddToken(0x13, "Occurrences");
            this.codePages[9].AddToken(0x14, "Interval");
            this.codePages[9].AddToken(0x15, "DayOfMonth");
            this.codePages[9].AddToken(0x16, "DayOfWeek");
            this.codePages[9].AddToken(0x17, "WeekOfMonth");
            this.codePages[9].AddToken(0x18, "MonthOfYear");
            this.codePages[9].AddToken(0x19, "Regenerate");
            this.codePages[9].AddToken(0x1A, "DeadOccur");
            this.codePages[9].AddToken(0x1B, "ReminderSet");
            this.codePages[9].AddToken(0x1C, "ReminderTime");
            this.codePages[9].AddToken(0x1D, "Sensitivity");
            this.codePages[9].AddToken(0x1E, "StartDate");
            this.codePages[9].AddToken(0x1F, "UtcStartDate");
            this.codePages[9].AddToken(0x20, "Subject");
            this.codePages[9].AddToken(0x22, "OrdinalDate");
            this.codePages[9].AddToken(0x23, "SubOrdinalDate");
            this.codePages[9].AddToken(0x24, "CalendarType");
            this.codePages[9].AddToken(0x25, "IsLeapMonth");
            this.codePages[9].AddToken(0x26, "FirstDayOfWeek");

            // Code Page 10: ResolveRecipients
            this.codePages[10] = new CodePage { Namespace = "ResolveRecipients", Xmlns = "resolverecipients" };

            this.codePages[10].AddToken(0x05, "ResolveRecipients");
            this.codePages[10].AddToken(0x06, "Response");
            this.codePages[10].AddToken(0x07, "Status");
            this.codePages[10].AddToken(0x08, "Type");
            this.codePages[10].AddToken(0x09, "Recipient");
            this.codePages[10].AddToken(0x0A, "DisplayName");
            this.codePages[10].AddToken(0x0B, "EmailAddress");
            this.codePages[10].AddToken(0x0C, "Certificates");
            this.codePages[10].AddToken(0x0D, "Certificate");
            this.codePages[10].AddToken(0x0E, "MiniCertificate");
            this.codePages[10].AddToken(0x0F, "Options");
            this.codePages[10].AddToken(0x10, "To");
            this.codePages[10].AddToken(0x11, "CertificateRetrieval");
            this.codePages[10].AddToken(0x12, "RecipientCount");
            this.codePages[10].AddToken(0x13, "MaxCertificates");
            this.codePages[10].AddToken(0x14, "MaxAmbiguousRecipients");
            this.codePages[10].AddToken(0x15, "CertificateCount");
            this.codePages[10].AddToken(0x16, "Availability");
            this.codePages[10].AddToken(0x17, "StartTime");
            this.codePages[10].AddToken(0x18, "EndTime");
            this.codePages[10].AddToken(0x19, "MergedFreeBusy");
            this.codePages[10].AddToken(0x1A, "Picture");
            this.codePages[10].AddToken(0x1B, "MaxSize");
            this.codePages[10].AddToken(0x1C, "Data");
            this.codePages[10].AddToken(0x1D, "MaxPictures");

            // Code Page 11: ValidateCert
            this.codePages[11] = new CodePage { Namespace = "ValidateCert", Xmlns = "ValidateCert" };

            this.codePages[11].AddToken(0x05, "ValidateCert");
            this.codePages[11].AddToken(0x06, "Certificates");
            this.codePages[11].AddToken(0x07, "Certificate");
            this.codePages[11].AddToken(0x08, "CertificateChain");
            this.codePages[11].AddToken(0x09, "CheckCrl");
            this.codePages[11].AddToken(0x0A, "Status");

            // Code Page 12: Contacts2
            this.codePages[12] = new CodePage { Namespace = "Contacts2", Xmlns = "contacts2" };

            this.codePages[12].AddToken(0x05, "CustomerId");
            this.codePages[12].AddToken(0x06, "GovernmentId");
            this.codePages[12].AddToken(0x07, "IMAddress");
            this.codePages[12].AddToken(0x08, "IMAddress2");
            this.codePages[12].AddToken(0x09, "IMAddress3");
            this.codePages[12].AddToken(0x0A, "ManagerName");
            this.codePages[12].AddToken(0x0B, "CompanyMainPhone");
            this.codePages[12].AddToken(0x0C, "AccountName");
            this.codePages[12].AddToken(0x0D, "NickName");
            this.codePages[12].AddToken(0x0E, "MMS");

            // Code Page 13: Ping
            this.codePages[13] = new CodePage { Namespace = "Ping", Xmlns = "ping" };

            this.codePages[13].AddToken(0x05, "Ping");
            this.codePages[13].AddToken(0x07, "Status");
            this.codePages[13].AddToken(0x08, "HeartbeatInterval");
            this.codePages[13].AddToken(0x09, "Folders");
            this.codePages[13].AddToken(0x0A, "Folder");
            this.codePages[13].AddToken(0x0B, "Id");
            this.codePages[13].AddToken(0x0C, "Class");
            this.codePages[13].AddToken(0x0D, "MaxFolders");

            // Code Page 14: Provision
            this.codePages[14] = new CodePage { Namespace = "Provision", Xmlns = "provision" };

            this.codePages[14].AddToken(0x05, "Provision");
            this.codePages[14].AddToken(0x06, "Policies");
            this.codePages[14].AddToken(0x07, "Policy");
            this.codePages[14].AddToken(0x08, "PolicyType");
            this.codePages[14].AddToken(0x09, "PolicyKey");
            this.codePages[14].AddToken(0x0A, "Data");
            this.codePages[14].AddToken(0x0B, "Status");
            this.codePages[14].AddToken(0x0C, "RemoteWipe");
            this.codePages[14].AddToken(0x0D, "EASProvisionDoc");
            this.codePages[14].AddToken(0x0E, "DevicePasswordEnabled");
            this.codePages[14].AddToken(0x0F, "AlphanumericDevicePasswordRequired");
            this.codePages[14].AddToken(0x10, "RequireStorageCardEncryption");
            this.codePages[14].AddToken(0x11, "PasswordRecoveryEnabled");
            this.codePages[14].AddToken(0x13, "AttachmentsEnabled");
            this.codePages[14].AddToken(0x14, "MinDevicePasswordLength");
            this.codePages[14].AddToken(0x15, "MaxInactivityTimeDeviceLock");
            this.codePages[14].AddToken(0x16, "MaxDevicePasswordFailedAttempts");
            this.codePages[14].AddToken(0x17, "MaxAttachmentSize");
            this.codePages[14].AddToken(0x18, "AllowSimpleDevicePassword");
            this.codePages[14].AddToken(0x19, "DevicePasswordExpiration");
            this.codePages[14].AddToken(0x1A, "DevicePasswordHistory");
            this.codePages[14].AddToken(0x1B, "AllowStorageCard");
            this.codePages[14].AddToken(0x1C, "AllowCamera");
            this.codePages[14].AddToken(0x1D, "RequireDeviceEncryption");
            this.codePages[14].AddToken(0x1E, "AllowUnsignedApplications");
            this.codePages[14].AddToken(0x1F, "AllowUnsignedInstallationPackages");
            this.codePages[14].AddToken(0x20, "MinDevicePasswordComplexCharacters");
            this.codePages[14].AddToken(0x21, "AllowWiFi");
            this.codePages[14].AddToken(0x22, "AllowTextMessaging");
            this.codePages[14].AddToken(0x23, "AllowPOPIMAPEmail");
            this.codePages[14].AddToken(0x24, "AllowBluetooth");
            this.codePages[14].AddToken(0x25, "AllowIrDA");
            this.codePages[14].AddToken(0x26, "RequireManualSyncWhenRoaming");
            this.codePages[14].AddToken(0x27, "AllowDesktopSync");
            this.codePages[14].AddToken(0x28, "MaxCalendarAgeFilter");
            this.codePages[14].AddToken(0x29, "AllowHTMLEmail");
            this.codePages[14].AddToken(0x2A, "MaxEmailAgeFilter");
            this.codePages[14].AddToken(0x2B, "MaxEmailBodyTruncationSize");
            this.codePages[14].AddToken(0x2C, "MaxEmailHTMLBodyTruncationSize");
            this.codePages[14].AddToken(0x2D, "RequireSignedSMIMEMessages");
            this.codePages[14].AddToken(0x2E, "RequireEncryptedSMIMEMessages");
            this.codePages[14].AddToken(0x2F, "RequireSignedSMIMEAlgorithm");
            this.codePages[14].AddToken(0x30, "RequireEncryptionSMIMEAlgorithm");
            this.codePages[14].AddToken(0x31, "AllowSMIMEEncryptionAlgorithmNegotiation");
            this.codePages[14].AddToken(0x32, "AllowSMIMESoftCerts");
            this.codePages[14].AddToken(0x33, "AllowBrowser");
            this.codePages[14].AddToken(0x34, "AllowConsumerEmail");
            this.codePages[14].AddToken(0x35, "AllowRemoteDesktop");
            this.codePages[14].AddToken(0x36, "AllowInternetSharing");
            this.codePages[14].AddToken(0x37, "UnapprovedInROMApplicationList");
            this.codePages[14].AddToken(0x38, "ApplicationName");
            this.codePages[14].AddToken(0x39, "ApprovedApplicationList");
            this.codePages[14].AddToken(0x3A, "Hash");
            this.codePages[14].AddToken(0x3B, "AccountOnlyRemoteWipe");

            // Code Page 15: Search
            this.codePages[15] = new CodePage { Namespace = "Search", Xmlns = "search" };

            this.codePages[15].AddToken(0x05, "Search");
            this.codePages[15].AddToken(0x07, "Store");
            this.codePages[15].AddToken(0x08, "Name");
            this.codePages[15].AddToken(0x09, "Query");
            this.codePages[15].AddToken(0x0A, "Options");
            this.codePages[15].AddToken(0x0B, "Range");
            this.codePages[15].AddToken(0x0C, "Status");
            this.codePages[15].AddToken(0x0D, "Response");
            this.codePages[15].AddToken(0x0E, "Result");
            this.codePages[15].AddToken(0x0F, "Properties");
            this.codePages[15].AddToken(0x10, "Total");
            this.codePages[15].AddToken(0x11, "EqualTo");
            this.codePages[15].AddToken(0x12, "Value");
            this.codePages[15].AddToken(0x13, "And");
            this.codePages[15].AddToken(0x14, "Or");
            this.codePages[15].AddToken(0x15, "FreeText");
            this.codePages[15].AddToken(0x17, "DeepTraversal");
            this.codePages[15].AddToken(0x18, "LongId");
            this.codePages[15].AddToken(0x19, "RebuildResults");
            this.codePages[15].AddToken(0x1A, "LessThan");
            this.codePages[15].AddToken(0x1B, "GreaterThan");
            this.codePages[15].AddToken(0x1E, "UserName");
            this.codePages[15].AddToken(0x1F, "Password");
            this.codePages[15].AddToken(0x20, "ConversationId");
            this.codePages[15].AddToken(0x21, "Picture");
            this.codePages[15].AddToken(0x22, "MaxSize");
            this.codePages[15].AddToken(0x23, "MaxPictures");

            // Code Page 16: GAL
            this.codePages[16] = new CodePage { Namespace = "GAL", Xmlns = "gal" };

            this.codePages[16].AddToken(0x05, "DisplayName");
            this.codePages[16].AddToken(0x06, "Phone");
            this.codePages[16].AddToken(0x07, "Office");
            this.codePages[16].AddToken(0x08, "Title");
            this.codePages[16].AddToken(0x09, "Company");
            this.codePages[16].AddToken(0x0A, "Alias");
            this.codePages[16].AddToken(0x0B, "FirstName");
            this.codePages[16].AddToken(0x0C, "LastName");
            this.codePages[16].AddToken(0x0D, "HomePhone");
            this.codePages[16].AddToken(0x0E, "MobilePhone");
            this.codePages[16].AddToken(0x0F, "EmailAddress");
            this.codePages[16].AddToken(0x10, "Picture");
            this.codePages[16].AddToken(0x11, "Status");
            this.codePages[16].AddToken(0x12, "Data");

            // Code Page 17: AirSyncBase
            this.codePages[17] = new CodePage { Namespace = "AirSyncBase", Xmlns = "airsyncbase" };

            this.codePages[17].AddToken(0x05, "BodyPreference");
            this.codePages[17].AddToken(0x06, "Type");
            this.codePages[17].AddToken(0x07, "TruncationSize");
            this.codePages[17].AddToken(0x08, "AllOrNone");
            this.codePages[17].AddToken(0x0A, "Body");
            this.codePages[17].AddToken(0x0B, "Data");
            this.codePages[17].AddToken(0x0C, "EstimatedDataSize");
            this.codePages[17].AddToken(0x0D, "Truncated");
            this.codePages[17].AddToken(0x0E, "Attachments");
            this.codePages[17].AddToken(0x0F, "Attachment");
            this.codePages[17].AddToken(0x10, "DisplayName");
            this.codePages[17].AddToken(0x11, "FileReference");
            this.codePages[17].AddToken(0x12, "Method");
            this.codePages[17].AddToken(0x13, "ContentId");
            this.codePages[17].AddToken(0x14, "ContentLocation");
            this.codePages[17].AddToken(0x15, "IsInline");
            this.codePages[17].AddToken(0x16, "NativeBodyType");
            this.codePages[17].AddToken(0x17, "ContentType");
            this.codePages[17].AddToken(0x18, "Preview");
            this.codePages[17].AddToken(0x19, "BodyPartPreference");
            this.codePages[17].AddToken(0x1A, "BodyPart");
            this.codePages[17].AddToken(0x1B, "Status");
            this.codePages[17].AddToken(0x1C, "Add");
            this.codePages[17].AddToken(0x1D, "Delete");
            this.codePages[17].AddToken(0x1E, "ClientId");
            this.codePages[17].AddToken(0x1F, "Content");
            this.codePages[17].AddToken(0x20, "Location");
            this.codePages[17].AddToken(0x21, "Annotation");
            this.codePages[17].AddToken(0x22, "Street");
            this.codePages[17].AddToken(0x23, "City");
            this.codePages[17].AddToken(0x24, "State");
            this.codePages[17].AddToken(0x25, "Country");
            this.codePages[17].AddToken(0x26, "PostalCode");
            this.codePages[17].AddToken(0x27, "Latitude");
            this.codePages[17].AddToken(0x28, "Longitude");
            this.codePages[17].AddToken(0x29, "Accuracy");
            this.codePages[17].AddToken(0x2A, "Altitude");
            this.codePages[17].AddToken(0x2B, "AltitudeAccuracy");
            this.codePages[17].AddToken(0x2C, "LocationUri");
            this.codePages[17].AddToken(0x2D, "InstanceId");


            // Code Page 18: Settings
            this.codePages[18] = new CodePage { Namespace = "Settings", Xmlns = "settings" };

            this.codePages[18].AddToken(0x05, "Settings");
            this.codePages[18].AddToken(0x06, "Status");
            this.codePages[18].AddToken(0x07, "Get");
            this.codePages[18].AddToken(0x08, "Set");
            this.codePages[18].AddToken(0x09, "Oof");
            this.codePages[18].AddToken(0x0A, "OofState");
            this.codePages[18].AddToken(0x0B, "StartTime");
            this.codePages[18].AddToken(0x0C, "EndTime");
            this.codePages[18].AddToken(0x0D, "OofMessage");
            this.codePages[18].AddToken(0x0E, "AppliesToInternal");
            this.codePages[18].AddToken(0x0F, "AppliesToExternalKnown");
            this.codePages[18].AddToken(0x10, "AppliesToExternalUnknown");
            this.codePages[18].AddToken(0x11, "Enabled");
            this.codePages[18].AddToken(0x12, "ReplyMessage");
            this.codePages[18].AddToken(0x13, "BodyType");
            this.codePages[18].AddToken(0x14, "DevicePassword");
            this.codePages[18].AddToken(0x15, "Password");
            this.codePages[18].AddToken(0x16, "DeviceInformation");
            this.codePages[18].AddToken(0x17, "Model");
            this.codePages[18].AddToken(0x18, "IMEI");
            this.codePages[18].AddToken(0x19, "FriendlyName");
            this.codePages[18].AddToken(0x1A, "OS");
            this.codePages[18].AddToken(0x1B, "OSLanguage");
            this.codePages[18].AddToken(0x1C, "PhoneNumber");
            this.codePages[18].AddToken(0x1D, "UserInformation");
            this.codePages[18].AddToken(0x1E, "EmailAddresses");
            this.codePages[18].AddToken(0x1F, "SMTPAddress");
            this.codePages[18].AddToken(0x20, "UserAgent");
            this.codePages[18].AddToken(0x21, "EnableOutboundSMS");
            this.codePages[18].AddToken(0x22, "MobileOperator");
            this.codePages[18].AddToken(0x23, "PrimarySmtpAddress");
            this.codePages[18].AddToken(0x24, "Accounts");
            this.codePages[18].AddToken(0x25, "Account");
            this.codePages[18].AddToken(0x26, "AccountId");
            this.codePages[18].AddToken(0x27, "AccountName");
            this.codePages[18].AddToken(0x28, "UserDisplayName");
            this.codePages[18].AddToken(0x29, "SendDisabled");
            this.codePages[18].AddToken(0x2B, "RightsManagementInformation");

            // Code Page 19: DocumentLibrary
            this.codePages[19] = new CodePage { Namespace = "DocumentLibrary", Xmlns = "documentlibrary" };

            this.codePages[19].AddToken(0x05, "LinkId");
            this.codePages[19].AddToken(0x06, "DisplayName");
            this.codePages[19].AddToken(0x07, "IsFolder");
            this.codePages[19].AddToken(0x08, "CreationDate");
            this.codePages[19].AddToken(0x09, "LastModifiedDate");
            this.codePages[19].AddToken(0x0A, "IsHidden");
            this.codePages[19].AddToken(0x0B, "ContentLength");
            this.codePages[19].AddToken(0x0C, "ContentType");

            // Code Page 20: ItemOperations
            this.codePages[20] = new CodePage { Namespace = "ItemOperations", Xmlns = "itemoperations" };

            this.codePages[20].AddToken(0x05, "ItemOperations");
            this.codePages[20].AddToken(0x06, "Fetch");
            this.codePages[20].AddToken(0x07, "Store");
            this.codePages[20].AddToken(0x08, "Options");
            this.codePages[20].AddToken(0x09, "Range");
            this.codePages[20].AddToken(0x0A, "Total");
            this.codePages[20].AddToken(0x0B, "Properties");
            this.codePages[20].AddToken(0x0C, "Data");
            this.codePages[20].AddToken(0x0D, "Status");
            this.codePages[20].AddToken(0x0E, "Response");
            this.codePages[20].AddToken(0x0F, "Version");
            this.codePages[20].AddToken(0x10, "Schema");
            this.codePages[20].AddToken(0x11, "Part");
            this.codePages[20].AddToken(0x12, "EmptyFolderContents");
            this.codePages[20].AddToken(0x13, "DeleteSubFolders");
            this.codePages[20].AddToken(0x14, "UserName");
            this.codePages[20].AddToken(0x15, "Password");
            this.codePages[20].AddToken(0x16, "Move");
            this.codePages[20].AddToken(0x17, "DstFldId");
            this.codePages[20].AddToken(0x18, "ConversationId");
            this.codePages[20].AddToken(0x19, "MoveAlways");

            // Code Page 21: ComposeMail
            this.codePages[21] = new CodePage { Namespace = "ComposeMail", Xmlns = "composemail" };

            this.codePages[21].AddToken(0x05, "SendMail");
            this.codePages[21].AddToken(0x06, "SmartForward");
            this.codePages[21].AddToken(0x07, "SmartReply");
            this.codePages[21].AddToken(0x08, "SaveInSentItems");
            this.codePages[21].AddToken(0x09, "ReplaceMime");
            this.codePages[21].AddToken(0x0B, "Source");
            this.codePages[21].AddToken(0x0C, "FolderId");
            this.codePages[21].AddToken(0x0D, "ItemId");
            this.codePages[21].AddToken(0x0E, "LongId");
            this.codePages[21].AddToken(0x0F, "InstanceId");
            this.codePages[21].AddToken(0x10, "Mime");
            this.codePages[21].AddToken(0x11, "ClientId");
            this.codePages[21].AddToken(0x12, "Status");
            this.codePages[21].AddToken(0x13, "AccountId");
            this.codePages[21].AddToken(0x15, "Forwardees");
            this.codePages[21].AddToken(0x16, "Forwardee");
            this.codePages[21].AddToken(0x17, "ForwardeeName");
            this.codePages[21].AddToken(0x18, "ForwardeeEmail");

            // Code Page 22: Email2
            this.codePages[22] = new CodePage { Namespace = "Email2", Xmlns = "email2" };

            this.codePages[22].AddToken(0x05, "UmCallerID");
            this.codePages[22].AddToken(0x06, "UmUserNotes");
            this.codePages[22].AddToken(0x07, "UmAttDuration");
            this.codePages[22].AddToken(0x08, "UmAttOrder");
            this.codePages[22].AddToken(0x09, "ConversationId");
            this.codePages[22].AddToken(0x0A, "ConversationIndex");
            this.codePages[22].AddToken(0x0B, "LastVerbExecuted");
            this.codePages[22].AddToken(0x0C, "LastVerbExecutionTime");
            this.codePages[22].AddToken(0x0D, "ReceivedAsBcc");
            this.codePages[22].AddToken(0x0E, "Sender");
            this.codePages[22].AddToken(0x0F, "CalendarType");
            this.codePages[22].AddToken(0x10, "IsLeapMonth");
            this.codePages[22].AddToken(0x11, "AccountId");
            this.codePages[22].AddToken(0x12, "FirstDayOfWeek");
            this.codePages[22].AddToken(0x13, "MeetingMessageType");
            this.codePages[22].AddToken(0x15, "IsDraft");
            this.codePages[22].AddToken(0x16, "Bcc");
            this.codePages[22].AddToken(0x17, "Send");

            // Code Page 23: Notes
            this.codePages[23] = new CodePage { Namespace = "Notes", Xmlns = "notes" };

            this.codePages[23].AddToken(0x05, "Subject");
            this.codePages[23].AddToken(0x06, "MessageClass");
            this.codePages[23].AddToken(0x07, "LastModifiedDate");
            this.codePages[23].AddToken(0x08, "Categories");
            this.codePages[23].AddToken(0x09, "Category");

            // Code Page 24: RightsManagement
            this.codePages[24] = new CodePage { Namespace = "RightsManagement", Xmlns = "rightsmanagement" };

            this.codePages[24].AddToken(0x05, "RightsManagementSupport");
            this.codePages[24].AddToken(0x06, "RightsManagementTemplates");
            this.codePages[24].AddToken(0x07, "RightsManagementTemplate");
            this.codePages[24].AddToken(0x08, "RightsManagementLicense");
            this.codePages[24].AddToken(0x09, "EditAllowed");
            this.codePages[24].AddToken(0x0A, "ReplyAllowed");
            this.codePages[24].AddToken(0x0B, "ReplyAllAllowed");
            this.codePages[24].AddToken(0x0C, "ForwardAllowed");
            this.codePages[24].AddToken(0x0D, "ModifyRecipientsAllowed");
            this.codePages[24].AddToken(0x0E, "ExtractAllowed");
            this.codePages[24].AddToken(0x0F, "PrintAllowed");
            this.codePages[24].AddToken(0x10, "ExportAllowed");
            this.codePages[24].AddToken(0x11, "ProgrammaticAccessAllowed");
            this.codePages[24].AddToken(0x12, "Owner");
            this.codePages[24].AddToken(0x13, "ContentExpiryDate");
            this.codePages[24].AddToken(0x14, "TemplateID");
            this.codePages[24].AddToken(0x15, "TemplateName");
            this.codePages[24].AddToken(0x16, "TemplateDescription");
            this.codePages[24].AddToken(0x17, "ContentOwner");
            this.codePages[24].AddToken(0x18, "RemoveRightsManagementProtection");

            //Code page 25: Find
            this.codePages[25] = new CodePage { Namespace = "Find", Xmlns = "Find" };
            this.codePages[25].AddToken(0x05, "Find");
            this.codePages[25].AddToken(0x06, "SearchId");
            this.codePages[25].AddToken(0x07, "ExecuteSearch");
            this.codePages[25].AddToken(0x08, "MailBoxSearchCriterion");
            this.codePages[25].AddToken(0x09, "Query");
            this.codePages[25].AddToken(0x0A, "Status");
            this.codePages[25].AddToken(0x0B, "FreeText");
            this.codePages[25].AddToken(0x0C, "Options");
            this.codePages[25].AddToken(0x0D, "Range");
            this.codePages[25].AddToken(0x0E, "DeepTraversal");
            this.codePages[25].AddToken(0x11, "Response");
            this.codePages[25].AddToken(0x12, "Result");
            this.codePages[25].AddToken(0x13, "Properties");
            this.codePages[25].AddToken(0x14, "Preview");
            this.codePages[25].AddToken(0x15, "HasAttachments");
            this.codePages[25].AddToken(0x16, "Total");
            this.codePages[25].AddToken(0x17, "DisplayCc");
            this.codePages[25].AddToken(0x18, "DisplayBcc");
            this.codePages[25].AddToken(0x19, "GalSearchCriterion");
            this.codePages[25].AddToken(0x20, "MaxPictures");
            this.codePages[25].AddToken(0x21, "MaxSize");
            this.codePages[25].AddToken(0x22, "Picture");

        }

        /// <summary>
        /// Gets the DataCollection in encoding process
        /// </summary>
        public Dictionary<string, int> EncodeDataCollection
        {
            get { return this.encodeDataCollection; }
        }

        /// <summary>
        /// Gets the DataCollection in decoding process
        /// </summary>
        public Dictionary<string, int> DecodeDataCollection
        {
            get { return this.decodeDataCollection; }
        }

        /// <summary>
        /// Loads byte array and decode to xml string.
        /// </summary>
        /// <param name="byteWBXML">The bytes to be decoded</param>
        /// <returns>The decoded xml string.</returns>
        public string DecodeToXml(byte[] byteWBXML)
        {
            this.xmlDoc = new XmlDocument();

            ByteQueue bytes = new ByteQueue(byteWBXML);

            // Remove the version from bytes
            bytes.Dequeue();

            // Remove public identifier from bytes
            bytes.DequeueMultibyteInt();

            // Gets the Character set from bytes
            int charset = bytes.DequeueMultibyteInt();
            if (charset != 0x6A)
            {
                return string.Empty;
            }

            // String table length. MS-ASWBXML does not use string tables, it should be 0.
            int stringTableLength = bytes.DequeueMultibyteInt();
            this.site.Assert.AreEqual<int>(0, stringTableLength, "MS-ASWBXML does not use string tables, therefore String table length should be 0.");

            // Initializes the DecodeDataCollection and begins to record
            if (null == this.decodeDataCollection)
            {
                this.decodeDataCollection = new Dictionary<string, int>();
            }
            else
            {
                this.decodeDataCollection.Clear();
            }

            // Adds the declaration
            XmlDeclaration xmlDec = this.xmlDoc.CreateXmlDeclaration("1.0", "utf-8", null);
            this.xmlDoc.InsertBefore(xmlDec, null);

            XmlNode currentNode = this.xmlDoc;

            while (bytes.Count > 0)
            {
                byte currentByte = bytes.Dequeue();

                switch ((GlobalTokens)currentByte)
                {
                    case GlobalTokens.SWITCH_PAGE:
                        int newCodePage = (int)bytes.Dequeue();
                        if (newCodePage >= 0 && newCodePage < 26)
                        {
                            this.currentCodePage = newCodePage;
                        }
                        else
                        {
                            this.site.Assert.Fail("Code page value which defined in MS-ASWBXML should be between 0-25, the actual value is : {0}.", newCodePage);
                        }

                        break;
                    case GlobalTokens.END:
                        if (currentNode.ParentNode != null)
                        {
                            currentNode = currentNode.ParentNode;
                        }
                        else
                        {
                            return string.Empty;
                        }

                        break;
                    case GlobalTokens.OPAQUE:
                        int cdataLength = bytes.DequeueMultibyteInt();
                        XmlCDataSection newOpaqueNode;
                        if (currentNode.Name == "ConversationId"
                            || currentNode.Name == "ConversationIndex")
                        {
                            newOpaqueNode = this.xmlDoc.CreateCDataSection(bytes.DequeueBase64String(cdataLength));
                        }
                        else
                        {
                            newOpaqueNode = this.xmlDoc.CreateCDataSection(bytes.DequeueString(cdataLength));
                        }

                        currentNode.AppendChild(newOpaqueNode);
                        break;
                    case GlobalTokens.STR_I:
                        XmlNode newTextNode = this.xmlDoc.CreateTextNode(bytes.DequeueString());
                        currentNode.AppendChild(newTextNode);
                        break;

                    case GlobalTokens.ENTITY:
                    case GlobalTokens.EXT_0:
                    case GlobalTokens.EXT_1:
                    case GlobalTokens.EXT_2:
                    case GlobalTokens.EXT_I_0:
                    case GlobalTokens.EXT_I_1:
                    case GlobalTokens.EXT_I_2:
                    case GlobalTokens.EXT_T_0:
                    case GlobalTokens.EXT_T_1:
                    case GlobalTokens.EXT_T_2:
                    case GlobalTokens.LITERAL:
                    case GlobalTokens.LITERAL_A:
                    case GlobalTokens.LITERAL_AC:
                    case GlobalTokens.LITERAL_C:
                    case GlobalTokens.PI:
                    case GlobalTokens.STR_T:
                        return string.Empty;

                    default:
                        bool hasAttributes = (currentByte & 0x80) > 0;
                        bool hasContent = (currentByte & 0x40) > 0;

                        byte token = (byte)(currentByte & 0x3F);

                        if (hasAttributes)
                        {
                            return string.Empty;
                        }

                        string strTag = this.codePages[this.currentCodePage].GetTag(token) ?? string.Format(CultureInfo.CurrentCulture, "UNKNOWN_TAG_{0,2:X}", token);

                        XmlNode newNode;
                        try
                        {
                            newNode = this.xmlDoc.CreateElement(this.codePages[this.currentCodePage].Xmlns, strTag, this.codePages[this.currentCodePage].Namespace);
                        }
                        catch (XmlException)
                        {
                            return string.Empty;
                        }

                        try
                        {
                            string codepageName = this.codePages[this.currentCodePage].Xmlns;
                            string combinedTagAndToken = string.Format(CultureInfo.CurrentCulture, @"{0}|{1}|{2}|{3}", this.decodeDataCollection.Count, codepageName, strTag, token);
                            this.decodeDataCollection.Add(combinedTagAndToken, this.currentCodePage);
                        }
                        catch (ArgumentException)
                        {
                            return string.Empty;
                        }

                        newNode.Prefix = string.Empty;
                        currentNode.AppendChild(newNode);

                        if (hasContent)
                        {
                            currentNode = newNode;
                        }

                        break;
                }
            }

            using (StringWriter stringWriter = new StringWriter())
            {
                XmlTextWriter textWriter = new XmlTextWriter(stringWriter) { Formatting = Formatting.Indented };
                this.xmlDoc.WriteTo(textWriter);
                textWriter.Flush();

                return stringWriter.ToString();
            }
        }

        /// <summary>
        /// Loads xml string and encodes to bytes.
        /// </summary>
        /// <param name="xmlValue">The xml string.</param>
        /// <returns>The encoded bytes.</returns>
        public byte[] EncodeToWBXML(string xmlValue)
        {
            this.xmlDoc.LoadXml(xmlValue);

            List<byte> byteList = new List<byte>();

            // Initializes the EncodeDataCollection
            if (this.encodeDataCollection == null)
            {
                this.encodeDataCollection = new Dictionary<string, int>();
            }
            else
            {
                this.encodeDataCollection.Clear();
            }

            byteList.Add(VersionByte);
            byteList.Add(PublicIdentifierByte);
            byteList.Add(CharsetByte);
            byteList.Add(StringTableLengthByte);

            foreach (XmlNode node in this.xmlDoc.ChildNodes)
            {
                byteList.AddRange(this.EncodeNode(node));
            }

            return byteList.ToArray();
        }

        /// <summary>
        /// Encodes a string.
        /// </summary>
        /// <param name="value">The string to encode.</param>
        /// <returns>The encoded bytes.</returns>
        private static byte[] EncodeString(string value)
        {
            List<byte> byteList = new List<byte>();

            char[] charArray = value.ToCharArray();

            for (int i = 0; i < charArray.Length; i++)
            {
                byteList.Add((byte)charArray[i]);
            }

            byteList.Add(0x00);

            return byteList.ToArray();
        }

        /// <summary>
        /// Encodes multi byte integer
        /// </summary>
        /// <param name="value">Then integer to encode.</param>
        /// <returns>The encoded bytes</returns>
        private static byte[] EncodeMultibyteInteger(int value)
        {
            List<byte> byteList = new List<byte>();

            while (value > 0)
            {
                byte addByte = (byte)(value & 0x7F);

                if (byteList.Count > 0)
                {
                    addByte |= 0x80;
                }

                byteList.Insert(0, addByte);

                value >>= 7;
            }

            return byteList.ToArray();
        }

        /// <summary>
        /// Encodes opaque data.
        /// </summary>
        /// <param name="opaqueBytes">The opaque data</param>
        /// <returns>The encoded bytes.</returns>
        private static byte[] EncodeOpaque(byte[] opaqueBytes)
        {
            List<byte> byteList = new List<byte>();

            byteList.AddRange(EncodeMultibyteInteger(opaqueBytes.Length));
            byteList.AddRange(opaqueBytes);

            return byteList.ToArray();
        }

        /// <summary>
        /// Encodes a node.
        /// </summary>
        /// <param name="node">The node need to encode.</param>
        /// <returns>The encoded bytes.</returns>
        private byte[] EncodeNode(XmlNode node)
        {
            List<byte> byteList = new List<byte>();
            switch (node.NodeType)
            {
                case XmlNodeType.Element:
                    if (node.Attributes != null && node.Attributes.Count > 0)
                    {
                        this.ParseXmlnsAttributes(node);
                    }

                    if (this.SetCodePageByXmlns(node.NamespaceURI))
                    {
                        byteList.Add((byte)GlobalTokens.SWITCH_PAGE);
                        byteList.Add((byte)this.currentCodePage);
                    }

                    // Gets token in this.codePages
                    byte wbxmlMapToken = this.codePages[this.currentCodePage].GetToken(node.LocalName);
                    byte token = wbxmlMapToken;
                    if (node.HasChildNodes)
                    {
                        token |= 0x40;
                    }

                    byteList.Add(token);

                    string codepageName = this.codePages[this.currentCodePage].Xmlns;
                    string combinedTagAndToken = string.Format(CultureInfo.CurrentCulture, @"{0}|{1}|{2}|{3}", this.encodeDataCollection.Count, codepageName, node.LocalName, wbxmlMapToken);
                    this.encodeDataCollection.Add(combinedTagAndToken, this.currentCodePage);

                    if (node.HasChildNodes)
                    {
                        foreach (XmlNode child in node.ChildNodes)
                        {
                            byteList.AddRange(this.EncodeNode(child));
                        }

                        byteList.Add((byte)GlobalTokens.END);
                    }

                    break;
                case XmlNodeType.Text:
                    byteList.Add((byte)GlobalTokens.STR_I);
                    byteList.AddRange(EncodeString(node.Value));
                    break;
                case XmlNodeType.CDATA:
                    byteList.Add((byte)GlobalTokens.OPAQUE);

                    byte[] cdataValue = System.Text.Encoding.ASCII.GetBytes(node.Value);
                    if (node.ParentNode.Name == "ConversationId"
                        || node.ParentNode.Name == "ConversationIndex")
                    {
                        cdataValue = System.Convert.FromBase64String(node.Value);
                    }

                    byteList.AddRange(EncodeOpaque(cdataValue));
                    break;
            }

            return byteList.ToArray();
        }

        /// <summary>
        /// Gets code page index by namespace.
        /// </summary>
        /// <param name="nameSpace">The namespace</param>
        /// <returns>The index of code page</returns>
        private int GetCodePageByNamespace(string nameSpace)
        {
            for (int i = 0; i < this.codePages.Length; i++)
            {
                if (string.Equals(this.codePages[i].Namespace, nameSpace, StringComparison.CurrentCultureIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        /// <summary>
        /// Switches to the code page by prefix.
        /// </summary>
        /// <param name="namespaceUri">The prefix</param>
        /// <returns>True, if successful.</returns>
        private bool SetCodePageByXmlns(string namespaceUri)
        {
            if (string.IsNullOrEmpty(namespaceUri))
            {
                if (this.currentCodePage != this.defaultCodePage)
                {
                    this.currentCodePage = this.defaultCodePage;
                    return true;
                }

                return false;
            }

            if (string.Equals(this.codePages[this.currentCodePage].Xmlns, namespaceUri, StringComparison.CurrentCultureIgnoreCase))
            {
                return false;
            }

            for (int i = 0; i < this.codePages.Length; i++)
            {
                if (string.Equals(this.codePages[i].Namespace, namespaceUri, StringComparison.CurrentCultureIgnoreCase))
                {
                    this.currentCodePage = i;
                    return true;
                }
            }

            throw new InvalidDataException(string.Format("Unknown Xmlns: {0}.", namespaceUri));
        }

        /// <summary>
        /// Parses namespaceUri attribute
        /// </summary>
        /// <param name="node">The xml node to parse</param>
        private void ParseXmlnsAttributes(XmlNode node)
        {
            if (node.Attributes == null)
            {
                return;
            }

            foreach (XmlAttribute attribute in node.Attributes)
            {
                int codePage = this.GetCodePageByNamespace(attribute.Value);

                if (!string.IsNullOrEmpty(attribute.Value) && (attribute.Value.StartsWith("http://www.w3.org/2001/XMLSchema-instance", StringComparison.CurrentCultureIgnoreCase) || attribute.Value.StartsWith("http://www.w3.org/2001/XMLSchema", StringComparison.CurrentCultureIgnoreCase)))
                {
                    break;
                }

                if (string.Equals(attribute.Name, "XMLNS", StringComparison.CurrentCultureIgnoreCase))
                {
                    this.defaultCodePage = codePage;
                }
                else if (string.Equals(attribute.Prefix, "XMLNS", StringComparison.CurrentCultureIgnoreCase))
                {
                    this.codePages[codePage].Xmlns = attribute.LocalName;
                }
            }
        }
    }
}