namespace Microsoft.Protocols.TestSuites.MS_ASCAL
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Calendar = Microsoft.Protocols.TestSuites.Common.DataStructures.Calendar;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// A static class contains all helper methods used in test cases.
    /// </summary>
    internal static class TestSuiteHelper
    {
        /// <summary>
        /// Private field for seed initialize to random number
        /// </summary>
        private static int seed = new Random().Next(0, 100000);

        /// <summary>
        /// Builds a Sync add request by using the specified sync key, folder collection ID and add application data.
        /// In general, returns the XMl formatted Sync request as follows:
        /// <!--
        /// <?xml version="1.0" encoding="utf-8"?>
        /// <Sync xmlns="AirSync">
        ///   <Collections>
        ///     <Collection>
        ///       <SyncKey>0</SyncKey>
        ///       <CollectionId>5</CollectionId>
        ///       <GetChanges>1</GetChanges>
        ///       <WindowSize>152</WindowSize>
        ///        <Commands>
        ///            <Add>
        ///                <ServerId>5:1</ServerId>
        ///                <ApplicationData>
        ///                    ...
        ///                </ApplicationData>
        ///            </Add>
        ///        </Commands>
        ///     </Collection>
        ///   </Collections>
        /// </Sync>
        /// -->
        /// </summary>
        /// <param name="collectionId">Specify the server ID of the folder to be synchronized, which can be returned by ActiveSync FolderSync command(Refer to [MS-ASCMD]2.2.3.30.5)</param>
        /// <param name="syncKey">Specify the sync key obtained from the last sync response(Refer to [MS-ASCMD]2.2.3.166.4)</param>
        /// <param name="addCalendars">Contains the data used to specify the Add element for Sync command(Refer to [MS-ASCMD]2.2.3.7.2)</param>
        /// <returns>Returns the SyncRequest instance</returns>
        internal static SyncRequest CreateSyncAddRequest(string collectionId, string syncKey, Request.SyncCollectionAddApplicationData addCalendars)
        {
            SyncRequest syncAddRequest;
            Request.SyncCollectionAdd add = new Request.SyncCollectionAdd
            {
                ClientId = TestSuiteHelper.Next(),
                ApplicationData = addCalendars
            };

            List<object> commandList = new List<object> { add };

            // The Sync request include the GetChanges element of the Collection element will set to 0 (FALSE)
            syncAddRequest = TestSuiteHelper.CreateSyncRequest(collectionId, syncKey, false);
            syncAddRequest.RequestData.Collections[0].Commands = commandList.ToArray();

            return syncAddRequest;
        }

        /// <summary>
        /// Builds a Sync delete request by using the specified sync key, folder collection ID, item serverID and deletesAsMoves option.
        /// In general, returns the XML formatted Sync request as follows:
        /// <!--
        /// <?xml version="1.0" encoding="utf-8"?>
        /// <Sync xmlns="AirSync">
        ///   <Collections>
        ///     <Collection>
        ///       <SyncKey>0</SyncKey>
        ///       <CollectionId>5</CollectionId>
        ///       <GetChanges>1</GetChanges>
        ///       <DeletesAsMoves>1</DeletesAsMoves>
        ///       <WindowSize>100</WindowSize>
        ///       <Commands>
        ///         <Delete>
        ///           <ServerId>5:1</ServerId>
        ///         </Delete>
        ///       </Commands>
        ///     </Collection>
        ///   </Collections>
        /// </Sync>
        /// -->
        /// </summary>
        /// <param name="collectionId">Specify the serverId of the folder to be synchronized, which can be returned by ActiveSync FolderSync command(Refer to [MS-ASCMD]2.2.3.30.5)</param>
        /// <param name="syncKey">Specify the sync key obtained from the last sync response(Refer to [MS-ASCMD]2.2.3.166.4)</param>
        /// <param name="serverId">Specify a unique identifier that was assigned by the server for a mailItem, which can be returned by ActiveSync Sync command</param>
        /// <returns>Returns the SyncRequest instance</returns>
        internal static SyncRequest CreateSyncDeleteRequest(string collectionId, string syncKey, string serverId)
        {
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                SyncKey = syncKey,
                CollectionId = collectionId,
                WindowSize = "100",
                DeletesAsMoves = false,
                DeletesAsMovesSpecified = true
            };

            Request.SyncCollectionDelete deleteData = new Request.SyncCollectionDelete { ServerId = serverId };

            syncCollection.Commands = new object[] { deleteData };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
        }

        /// <summary>
        /// Build a generic Sync request without command references by using the specified sync key, folder collection ID and body preference option.
        /// In general, returns the XML formatted Sync request as follows:
        /// <!--
        /// <?xml version="1.0" encoding="utf-8"?>
        /// <Sync xmlns="AirSync">
        ///   <Collections>
        ///     <Collection>
        ///       <SyncKey>0</SyncKey>
        ///       <CollectionId>5</CollectionId>
        ///       <DeletesAsMoves>1</DeletesAsMoves>
        ///       <GetChanges>1</GetChanges>
        ///       <WindowSize>152</WindowSize>
        ///       <Options>
        ///         <MIMESupport>0</MIMESupport>
        ///         <airsyncbase:BodyPreference>
        ///           <airsyncbase:Type>2</airsyncbase:Type>
        ///         </airsyncbase:BodyPreference>
        ///       </Options>
        ///     </Collection>
        ///   </Collections>
        /// </Sync>
        /// -->
        /// </summary>
        /// <param name="collectionId">Specify the server ID of the folder to be synchronized, which can be returned by ActiveSync FolderSync command(Refer to [MS-ASCMD]2.2.3.30.5)</param>
        /// <param name="syncKey">Specify the sync key obtained from the last sync response(Refer to [MS-ASCMD]2.2.3.166.4)</param>
        /// <param name="getChanges">Sets sync collection information related to the GetChanges</param>
        /// <returns>Returns the SyncRequest instance</returns>
        internal static SyncRequest CreateSyncRequest(string collectionId, string syncKey, bool getChanges)
        {
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                WindowSize = "512",
                SyncKey = syncKey,
                CollectionId = collectionId,
                Supported = null
            };

            if (getChanges)
            {
                syncCollection.GetChanges = true;
                syncCollection.GetChangesSpecified = true;
            }

            Request.Options syncOptions = new Request.Options();
            List<object> syncOptionItems = new List<object>();
            List<Request.ItemsChoiceType1> syncOptionItemsName = new List<Request.ItemsChoiceType1>
            {
                Request.ItemsChoiceType1.BodyPreference
            };

            syncOptionItems.Add(
                new Request.BodyPreference()
                {
                    Type = 2,
                    TruncationSize = 0,
                    TruncationSizeSpecified = false,
                    Preview = 0,
                    PreviewSpecified = false
                });
            syncOptions.Items = syncOptionItems.ToArray();
            syncOptions.ItemsElementName = syncOptionItemsName.ToArray();
            syncCollection.Options = new Request.Options[] { syncOptions };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
        }

        /// <summary>
        /// Builds a Search request on the Mailbox store by using the specified keyword and folder collection ID
        /// In general, returns the XML formatted search request as follows:
        /// <!--
        /// <?xml version="1.0" encoding="utf-8"?>
        /// <Search xmlns="Search" xmlns:airsync="AirSync">
        /// <Store>
        ///   <Name>Mailbox</Name>
        ///     <Query>
        ///       <And>
        ///         <airsync:CollectionId>5</airsync:CollectionId>
        ///         <FreeText>Presentation</FreeText>
        ///       </And>
        ///     </Query>
        ///     <Options>
        ///       <RebuildResults />
        ///       <Range>0-9</Range>
        ///       <DeepTraversal/>
        ///     </Options>
        ///   </Store>
        /// </Search>
        /// -->
        /// </summary>
        /// <param name="storeName">Specify the store for which to search. Refer to [MS-ASCMD] section 2.2.3.110.2</param>
        /// <param name="keyword">Specify a string value for which to search. Refer to [MS-ASCMD] section 2.2.3.73</param>
        /// <param name="collectionId">Specify the folder in which to search. Refer to [MS-ASCMD] section 2.2.3.30.4</param>
        /// <returns>Returns a SearchRequest instance</returns>
        internal static SearchRequest CreateSearchRequest(string storeName, string keyword, string collectionId)
        {
            if (null == keyword)
            {
                throw new ArgumentNullException("keyword", "keyword: Specify the value to search");
            }

            if (null == collectionId)
            {
                throw new ArgumentNullException("collectionId", "folderCollectionId: Specify the folder ID to search");
            }

            Request.SearchStore searchStore = new Request.SearchStore
            {
                Name = storeName,
                Options = new Request.Options1()
            };

            Dictionary<Request.ItemsChoiceType6, object> items = new Dictionary<Request.ItemsChoiceType6, object>
            {
                {
                    Request.ItemsChoiceType6.DeepTraversal, string.Empty
                },
                {
                    Request.ItemsChoiceType6.RebuildResults, string.Empty
                },
                {
                    Request.ItemsChoiceType6.Range, "0-9"
                }
            };

            searchStore.Options.Items = items.Values.ToArray<object>();
            searchStore.Options.ItemsElementName = items.Keys.ToArray<Request.ItemsChoiceType6>();

            // Build up query condition by using the keyword and folder CollectionID
            Request.queryType queryItem = new Request.queryType
            {
                Items = new object[] { collectionId, keyword },

                ItemsElementName = new Request.ItemsChoiceType2[]
                {
                    Request.ItemsChoiceType2.CollectionId,
                    Request.ItemsChoiceType2.FreeText
                }
            };

            searchStore.Query = new Request.queryType
            {
                Items = new object[] { queryItem },
                ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And }
            };

            return Common.CreateSearchRequest(new Request.SearchStore[] { searchStore });
        }

        /// <summary>
        /// Builds a ItemOperations request to fetch the whole content of a single mail item
        /// by using the specified collectionId, emailServerId,bodyPreference and bodyPartPreference
        /// In general, returns the XML formatted ItemOperations request as follows:
        /// <!--
        /// <?xml version="1.0" encoding="utf-8"?>
        /// <ItemOperations xmlns="ItemOperations" xmlns:airsync="AirSync" xmlns:airsyncbase="AirSyncBase">
        ///    <Fetch>
        ///       <Store>Mailbox</Store>
        ///       <airsync:CollectionId>5</airsync:CollectionId>
        ///       <airsync:ServerId>5:1</airsync:ServerId>
        ///       <Options>
        ///          <airsync:MIMESupport>2</airsync:MIMESupport>
        ///          <airsyncbase:BodyPreference>
        ///             <airsyncbase:Type>4</airsyncbase:Type>
        ///          </airsyncbase:BodyPreference>
        ///          <airsyncbase:BodyPreference>
        ///             <airsyncbase:Type>2</airsyncbase:Type>
        ///          </airsyncbase:BodyPreference>
        ///       </Options>
        ///    </Fetch>
        /// </ItemOperations>
        /// -->
        /// </summary>
        /// <param name="collectionId">Specify the folder of mailItem, which can be returned by ActiveSync FolderSync command(Refer to [MS-ASCMD]2.2.3.30.2)</param>
        /// <param name="serverIds">Specify a unique identifier that was assigned by the server for a mailItem, which can be returned by ActiveSync Sync command(Refer to [MS-ASCMD]2.2.3.151.5)</param>
        /// <param name="schema">Sets the schema information</param>
        /// <returns>Returns the ItemOperationsRequest instance</returns>
        internal static ItemOperationsRequest CreateItemOperationsFetchRequest(string collectionId, List<string> serverIds, Request.Schema schema)
        {
            Request.ItemOperations itemOperations = new Request.ItemOperations();
            List<object> items = new List<object>();

            Request.ItemOperationsFetch fetchElement = new Request.ItemOperationsFetch();

            if (serverIds != null)
            {
                foreach (string item in serverIds)
                {
                    fetchElement.CollectionId = collectionId;
                    fetchElement.ServerId = item;
                    items.Add(fetchElement);
                }
            }

            itemOperations.Items = items.ToArray();

            foreach (object item in itemOperations.Items)
            {
                Request.ItemOperationsFetch fetch = item as Request.ItemOperationsFetch;
                if (fetch != null)
                {
                    fetch.Store = SearchName.Mailbox.ToString();
                    Request.ItemOperationsFetchOptions fetchOptions = new Request.ItemOperationsFetchOptions();

                    List<object> fetchOptionItems = new List<object>();
                    List<Request.ItemsChoiceType5> fetchOptionItemsName = new List<Request.ItemsChoiceType5>
                    {
                        Request.ItemsChoiceType5.BodyPreference,
                        Request.ItemsChoiceType5.Schema
                    };

                    fetchOptionItems.Add(
                        new Request.BodyPreference()
                        {
                            AllOrNone = false,
                            AllOrNoneSpecified = false,
                            TruncationSize = 0,
                            TruncationSizeSpecified = false,
                            Preview = 0,
                            PreviewSpecified = false,
                            Type = 2,
                        });
                    fetchOptionItems.Add(schema);

                    fetchOptions.Items = fetchOptionItems.ToArray();
                    fetchOptions.ItemsElementName = fetchOptionItemsName.ToArray();
                    fetch.Options = fetchOptions;
                }
            }

            return Common.CreateItemOperationsRequest(itemOperations.Items);
        }

        /// <summary>
        /// Builds an ItemOperations request to empty a folder of all its items
        /// by using the specified collectionId and options
        /// In general, returns the XML formatted ItemOperations request as follows:
        /// <!--
        /// <?xml version="1.0" encoding="utf-8"?>
        /// <ItemOperations xmlns="ItemOperations" xmlns:airsync="AirSync" xmlns:airsyncbase="AirSyncBase">
        ///    <EmptyFolderContents>
        ///       <airsync:CollectionId>5</airsync:CollectionId>
        ///       <Options>
        ///          <DeleteSubFolders ></DeleteSubFolders >
        ///       </Options>
        ///    </EmptyFolderContents>
        /// </ItemOperations>
        /// -->
        /// </summary>
        /// <param name="collectionId">Specify the folder to be emptied, which can be returned by ActiveSync FolderSync command(Refer to [MS-ASCMD]2.2.3.30.2)</param>
        /// <returns>Returns the ItemOperationsRequest instance</returns>
        internal static ItemOperationsRequest CreateItemOperationsEmptyFolderContentsRequest(string collectionId)
        {
            Request.ItemOperationsEmptyFolderContents emptyFolderContents = new Request.ItemOperationsEmptyFolderContents
            {
                CollectionId = collectionId,
                Options = new Request.ItemOperationsEmptyFolderContentsOptions { DeleteSubFolders = string.Empty }
            };

            return Common.CreateItemOperationsRequest(new object[] { emptyFolderContents });
        }

        /// <summary>
        /// Create iCalendar format string from one calendar instance
        /// </summary>
        /// <param name="calendar">Calendar information</param>
        /// <param name="method">Specify normal appointments from meeting requests, responses, and cancellations, it can be set to 'REQUEST', 'REPLY', or 'CANCEL'</param>
        /// <param name="replyMethod">Specify REPLY method, it can be set to 'ACCEPTED', 'TENTATIVE', or 'DECLINED'</param>
        /// <param name="organizerEmailAddress">The organizer email address</param>
        /// <param name="attendeeEmailAddress">The attendee email address</param>
        /// <returns>iCalendar formatted string</returns>
        internal static string CreateiCalendarFormatContent(Calendar calendar, string method, string replyMethod, string organizerEmailAddress, string attendeeEmailAddress)
        {
            StringBuilder icalendar = new StringBuilder();
            icalendar.AppendLine("BEGIN: VCALENDAR");
            icalendar.AppendLine("PRODID:-//Microsoft Protocols TestSuites");
            icalendar.AppendLine("VERSION:2.0");

            switch (method.ToUpper(CultureInfo.CurrentCulture))
            {
                case "REQUEST":
                    icalendar.AppendLine("METHOD:REQUEST");
                    icalendar.AppendLine("X-MS-OLK-FORCEINSPECTOROPEN:TRUE");
                    break;

                case "CANCEL":
                    icalendar.AppendLine("METHOD:CANCEL");
                    break;

                case "REPLY":
                    icalendar.AppendLine("METHOD:REPLY");
                    break;
            }

            icalendar.AppendLine("BEGIN:VTIMEZONE");
            icalendar.AppendLine("TZID:Universal Time");
            icalendar.AppendLine("BEGIN:STANDARD");
            icalendar.AppendLine("DTSTART:16011104T020000");
            icalendar.AppendLine("TZOFFSETFROM:-0000");
            icalendar.AppendLine("TZOFFSETTO:+0000");
            icalendar.AppendLine("END:STANDARD");
            icalendar.AppendLine("BEGIN:DAYLIGHT");
            icalendar.AppendLine("DTSTART:16010311T020000");
            icalendar.AppendLine("TZOFFSETFROM:-0000");
            icalendar.AppendLine("TZOFFSETTO:+0000");
            icalendar.AppendLine("END:DAYLIGHT");
            icalendar.AppendLine("END:VTIMEZONE");

            icalendar.AppendLine("BEGIN:VEVENT");
            icalendar.AppendLine("UID:" + calendar.UID);
            icalendar.AppendLine("DTSTAMP:" + ((DateTime)calendar.DtStamp).ToUniversalTime().ToString("yyyyMMddTHHmmss"));
            icalendar.AppendLine("DESCRIPTION:" + calendar.Subject);

            switch (method.ToUpper(CultureInfo.CurrentCulture))
            {
                case "REQUEST":
                    icalendar.AppendLine("SUMMARY:" + calendar.Subject);
                    icalendar.AppendLine("ATTENDEE;CN=\"\";RSVP=" + (calendar.ResponseRequested == true ? "TRUE" : "FALSE") + ":mailto:" + attendeeEmailAddress);

                    icalendar.AppendLine("ORGANIZER:MAILTO:" + organizerEmailAddress);

                    break;

                case "CANCEL":
                    icalendar.AppendLine("STATUS:CANCELLED");
                    icalendar.AppendLine("SUMMARY:" + "Cancelled:" + calendar.Subject);
                    icalendar.AppendLine("ATTENDEE;CN=\"\";RSVP=" + (calendar.ResponseRequested == true ? "TRUE" : "FALSE") + ":mailto:" + attendeeEmailAddress);

                    icalendar.AppendLine("ORGANIZER:MAILTO:" + organizerEmailAddress);

                    break;

                case "REPLY":
                    icalendar.AppendLine("SUMMARY:" + replyMethod + calendar.Subject);
                    icalendar.AppendLine("ATTENDEE;PARTSTAT=" + replyMethod.ToUpper(CultureInfo.CurrentCulture) + ":mailto:" + attendeeEmailAddress);

                    break;
            }

            icalendar.AppendLine("LOCATION:" + (calendar.Location ?? "My Office"));

            if (calendar.AllDayEvent == 1)
            {
                icalendar.AppendLine("DTSTART;VALUE=DATE:" + ((DateTime)calendar.StartTime).ToUniversalTime().Date.ToString("yyyyMMdd"));
                icalendar.AppendLine("DTEND;VALUE=DATE:" + ((DateTime)calendar.EndTime).ToUniversalTime().Date.ToString("yyyyMMdd"));
                icalendar.AppendLine("X-MICROSOFT-CDO-ALLDAYEVENT:TRUE");
            }
            else
            {
                icalendar.AppendLine("DTSTART;TZID=\"Universal Time\":" + ((DateTime)calendar.StartTime).ToUniversalTime().ToString("yyyyMMddTHHmmss"));
                icalendar.AppendLine("DTEND;TZID=\"Universal Time\":" + ((DateTime)calendar.EndTime).ToUniversalTime().ToString("yyyyMMddTHHmmss"));
            }

            if (calendar.Recurrence != null)
            {
                switch (calendar.Recurrence.Type)
                {
                    case 0:
                        icalendar.AppendLine("RRULE:FREQ=DAILY;COUNT=" + calendar.Recurrence.Occurrences.ToString());
                        break;

                    case 1:
                        icalendar.AppendLine("RRULE:FREQ=WEEKLY;BYDAY=MO;COUNT=" + calendar.Recurrence.Occurrences.ToString());
                        break;

                    case 2:
                        icalendar.AppendLine("RRULE:FREQ=MONTHLY;COUNT=" + calendar.Recurrence.Occurrences.ToString() + ";BYMONTHDAY=1");
                        break;

                    case 3:
                        icalendar.AppendLine("RRULE:FREQ=MONTHLY;COUNT=" + calendar.Recurrence.Occurrences.ToString() + ";BYDAY=1MO");
                        break;

                    case 5:
                        icalendar.AppendLine("RRULE:FREQ=YEARLY;COUNT=" + calendar.Recurrence.Occurrences.ToString() + ";BYMONTHDAY=1;BYMONTH=1");
                        break;

                    case 6:
                        icalendar.AppendLine("RRULE:FREQ=YEARLY;COUNT=" + calendar.Recurrence.Occurrences.ToString() + ";BYDAY=2MO;BYMONTH=1");
                        break;
                }
            }

            if (calendar.Exceptions != null)
            {
                icalendar.AppendLine("EXDATE;TZID=\"Universal Time\":" + calendar.Exceptions.Exception[0].ExceptionStartTime);
            }

            switch (calendar.BusyStatus)
            {
                case 0:
                    icalendar.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:FREE");
                    break;

                case 1:
                    icalendar.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:TENTATIVE");
                    break;

                case 2:
                    icalendar.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:BUSY");
                    break;

                case 3:
                    icalendar.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:OOF");
                    break;
            }

            if (calendar.DisallowNewTimeProposal == true)
            {
                icalendar.AppendLine("X-MICROSOFT-DISALLOW-COUNTER:TRUE");
            }

            switch (method.ToUpper(CultureInfo.CurrentCulture))
            {
                case "REQUEST":
                case "CANCEL":
                    if (calendar.Reminder.HasValue)
                    {
                        icalendar.AppendLine("BEGIN:VALARM");
                        icalendar.AppendLine("TRIGGER:-PT" + calendar.Reminder + "M");
                        icalendar.AppendLine("ACTION:DISPLAY");
                        icalendar.AppendLine("END:VALARM");
                    }

                    break;
            }

            icalendar.AppendLine("END:VEVENT");

            if (calendar.Exceptions != null)
            {
                icalendar.AppendLine("BEGIN:VEVENT");
                icalendar.AppendLine("DTSTART;TZID=\"Universal Time\":" + calendar.Exceptions.Exception[0].StartTime);
                icalendar.AppendLine("DTEND;TZID=\"Universal Time\":" + calendar.Exceptions.Exception[0].EndTime);
                icalendar.AppendLine("UID:" + calendar.UID);
                icalendar.AppendLine("DTSTAMP:" + ((DateTime)calendar.DtStamp).ToUniversalTime().ToString("yyyyMMddTHHmmss"));
                icalendar.AppendLine("SUMMARY:" + calendar.Exceptions.Exception[0].Subject);
                icalendar.AppendLine("RECURRENCE-ID;TZID=\"Universal Time\":" + calendar.Exceptions.Exception[0].ExceptionStartTime);
                icalendar.AppendLine("LOCATION:" + calendar.Exceptions.Exception[0].Location);

                if (calendar.Exceptions.Exception[0].AllDayEvent == 1)
                {
                    icalendar.AppendLine("X-MICROSOFT-CDO-ALLDAYEVENT:TRUE");
                }

                if (calendar.Exceptions.Exception[0].ReminderSpecified==true)
                {
                    icalendar.AppendLine("BEGIN:VALARM");
                    icalendar.AppendLine("TRIGGER:-PT" + calendar.Exceptions.Exception[0].Reminder + "M");
                    icalendar.AppendLine("ACTION:DISPLAY");
                    icalendar.AppendLine("END:VALARM");
                }

                icalendar.AppendLine("END:VEVENT");
            }

            icalendar.AppendLine("END:VCALENDAR");
            return icalendar.ToString();
        }

        /// <summary>
        /// Create a meeting request mime
        /// </summary>
        /// <param name="from">The from address of mail</param>
        /// <param name="to">The to address of the mail</param>
        /// <param name="subject">The subject of the mail</param>
        /// <param name="body">The body content of the mail</param>
        /// <param name="icalendarContent">The content of iCalendar required by this meeting</param>
        /// <returns>Returns the corresponding sample meeting mime</returns>
        internal static string CreateMeetingRequestMime(string from, string to, string subject, string body, string icalendarContent)
        {
            string meetingRequestMime =
@"From: {0}
To: {1}
Subject: {2}
Content-Type: multipart/alternative;
    boundary=""_001_MSASCAL_""
MIME-Version: 1.0

--_001_MSASCAL_
Content-Type: text/plain; charset=""us-ascii""

{3}

--_001_MSASCAL_
Content-Type: text/calendar; charset=""us-ascii""; method=REQUEST

{4}

--_001_MSASCAL_--

";
            return Common.FormatString(meetingRequestMime, from, to, subject, body, icalendarContent);
        }

        /// <summary>
        /// Builds a SendMail request by using the specified client Id, copyToSentItems option and mail mime content.
        /// In general, returns the XML formatted Sync request as follows:
        /// <!--
        /// <?xml version="1.0" encoding="utf-8"?>
        /// <SendMail xmlns="ComposeMail">
        ///   <ClientId>3</ClientId>
        ///   <SaveInSentItems/>
        ///   <Mime>....</Mime>
        /// </SendMail>
        /// -->
        /// </summary>
        /// <param name="clientId">Specify the client Id</param>
        /// <param name="saveInSentItems">Specify whether needs to store a mail copy to sent items</param>
        /// <param name="mime">Specify the mail mime</param>
        /// <returns>Returns the SendMailRequest instance</returns>
        internal static SendMailRequest CreateSendMailRequest(string clientId, bool saveInSentItems, string mime)
        {
            Request.SendMail sendMail = new Request.SendMail
            {
                SaveInSentItems = saveInSentItems ? string.Empty : null,
                ClientId = clientId,
                Mime = mime
            };

            SendMailRequest sendMailRequest = Common.CreateSendMailRequest();
            sendMailRequest.RequestData = sendMail;
            return sendMailRequest;
        }

        /// <summary>
        /// Builds a initial Sync request by using the specified collection Id.
        /// In order to sync the content of a folder, an initial sync key for the folder MUST be obtained from the server.
        /// The client obtains the key by sending an initial Sync request with a SyncKey element value of zero and the CollectionId element
        /// In general, returns the XML formatted Sync request as follows:
        /// <!--
        /// <?xml version="1.0" encoding="utf-8"?>
        /// <Sync xmlns="AirSync">
        ///   <Collections>
        ///     <Collection>
        ///       <SyncKey>0</SyncKey>
        ///       <CollectionId>5</CollectionId>
        ///     </Collection>
        ///   </Collections>
        /// </Sync>
        /// -->
        /// </summary>
        /// <param name="collectionId">Identify the folder as the collection being synchronized, which can be returned by ActiveSync FolderSync command(Refer to [MS-ASCMD]2.2.3.30.5)</param>
        /// <param name="supported">Specifies which contact and calendar elements in a Sync request are managed by the client and therefore not ghosted.</param>
        /// <returns>Returns the SyncRequest instance</returns>
        internal static SyncRequest InitializeSyncRequest(string collectionId, Request.Supported supported)
        {
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                Supported = supported,
                WindowSize = "512",
                CollectionId = collectionId,
                SyncKey = "0"
            };

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType1> itemsElementName = new List<Request.ItemsChoiceType1>
            {
                Request.ItemsChoiceType1.BodyPreference
            };

            items.Add(
                new Request.BodyPreference()
                {
                    TruncationSize = 0,
                    TruncationSizeSpecified = false,
                    Type = 2,
                    Preview = 0,
                    PreviewSpecified = false,
                });

            Request.Options option = new Request.Options
            {
                Items = items.ToArray(),
                ItemsElementName = itemsElementName.ToArray()
            };

            syncCollection.Options = new Request.Options[] { option };

            SyncRequest request = Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });

            return request;
        }

        /// <summary>
        /// Create a Request.Attendees instance with required child elements.
        /// </summary>
        /// <param name="attendeeEmailAddress">The attendee email address</param>
        /// <param name="attendeeName">The attendee display name</param>
        /// <returns>The created Request.Attendees instance.</returns>
        internal static Request.Attendees CreateAttendeesRequired(string[] attendeeEmailAddress, string[] attendeeName)
        {
            List<Request.AttendeesAttendee> attendeelist = new List<Request.AttendeesAttendee>();

            for (int i = 0; i < attendeeEmailAddress.Length; i++)
            {
                Request.AttendeesAttendee attendee = new Request.AttendeesAttendee
                {
                    Email = attendeeEmailAddress[i],
                    Name = attendeeName[i],
                    AttendeeType = 1,
                    AttendeeTypeSpecified = true
                };
                attendeelist.Add(attendee);
            }

            Request.Attendees attendees = new Request.Attendees
            {
                Attendee = attendeelist.ToArray()
            };

            return attendees;
        }

        /// <summary>
        /// Create a Request.Body instance.
        /// </summary>
        /// <param name="bodyType">the format type of the body content of the item</param>
        /// <param name="content">the data of the message body (2) or the message part of the calendar item, contact, document, e-mail, or task</param>
        /// <returns>The created Request.Body instance.</returns>
        internal static Request.Body CreateCalendarBody(byte bodyType, string content)
        {
            Request.Body calendarBody = new Request.Body
            {
                Type = bodyType,
                Data = content
            };

            return calendarBody;
        }

        /// <summary>
        /// Create a Request.Categories instance.
        /// </summary>
        /// <param name="calendarCategory">Specifies a category that is assigned to the calendar item or exception item.</param>
        /// <returns>The created Request.Categories instance.</returns>
        internal static Request.Categories CreateCalendarCategories(string[] calendarCategory)
        {
            Request.Categories calendarCategories = new Request.Categories();

            if (calendarCategory != null)
            {
                calendarCategories.Category = calendarCategory;
            }

            return calendarCategories;
        }

        /// <summary>
        /// Create a Request.ExceptionsException instance with ExceptionStartTime elements.
        /// </summary>
        /// <param name="exceptionStartTime">Specifies the start time of the original recurring meeting</param>
        /// <returns>The created Request.ExceptionsException instance.</returns>
        internal static Request.ExceptionsException CreateExceptionRequired(string exceptionStartTime)
        {
            Request.ExceptionsException exception = new Request.ExceptionsException
            {
                ExceptionStartTime = exceptionStartTime
            };

            return exception;
        }

        /// <summary>
        /// Check if the response message only contains the specified element in the specified xml tag.
        /// </summary>
        /// <param name="rawResponseXml">The raw xml of the response returned by SUT</param>
        /// <param name="tagName">The name of the specified xml tag.</param>
        /// <param name="elementName">The element name that the raw xml should contain.</param>
        /// <returns>If the response only contains the specified element, return true; otherwise, false.</returns>
        internal static bool IsOnlySpecifiedElement(XmlElement rawResponseXml, string tagName, string elementName)
        {
            bool isOnlySpecifiedElement = false;
            if (rawResponseXml != null)
            {
                XmlNodeList nodes = rawResponseXml.GetElementsByTagName(tagName);
                foreach (XmlNode node in nodes)
                {
                    if (node.HasChildNodes)
                    {
                        XmlNodeList children = node.ChildNodes;
                        if (children.Count > 0)
                        {
                            foreach (XmlNode child in children)
                            {
                                if (string.Equals(child.Name, elementName))
                                {
                                    isOnlySpecifiedElement = true;
                                }
                                else
                                {
                                    isOnlySpecifiedElement = false;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            isOnlySpecifiedElement = false;
                        }
                    }
                    else
                    {
                        isOnlySpecifiedElement = false;
                    }
                }
            }

            return isOnlySpecifiedElement;
        }

        /// <summary>
        /// Get the next value of the client ID
        /// </summary>
        /// <returns>Return the id value</returns>
        internal static string Next()
        {
            return (++seed).ToString();
        }
    }
}