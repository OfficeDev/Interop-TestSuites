namespace Microsoft.Protocols.TestSuites.MS_ASEMAIL
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Text;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// A static class contains all helper methods used in test cases.
    /// </summary>
    internal static class TestSuiteHelper
    {
        /// <summary>
        /// private field for seed initialize to random number.
        /// </summary>
        private static int seek = new Random().Next(0, 100000);

        /// <summary>
        /// Get client id for send mail
        /// </summary>
        /// <returns>The client id</returns>
        public static string GetClientId()
        {
            return (++seek).ToString();
        }

        #region Create sample mime string for SendMail operation
        /// <summary>
        /// Create a sample plain text mime
        /// </summary>
        /// <param name="from">The from address of mail</param>
        /// <param name="to">The to address of the mail</param>
        /// <param name="cc">The cc address of the mail</param>
        /// <param name="bcc">The bcc address of the mail</param>
        /// <param name="subject">The subject of the mail</param>
        /// <param name="body">The body content of the mail</param>
        /// <param name="sender">The sender of the mail</param>
        /// <param name="replyTo">The replyTo of the mail</param>
        /// <returns>Returns the corresponding sample plain text mime</returns>
        internal static string CreatePlainTextMime(string from, string to, string cc, string bcc, string subject, string body, string sender = null, string replyTo = null)
        {
            cc = string.IsNullOrEmpty(cc) ? string.Empty : string.Format("Cc: {0}\r\n", cc);
            bcc = string.IsNullOrEmpty(bcc) ? string.Empty : string.Format("Bcc: {0}\r\n", bcc);
            sender = string.IsNullOrEmpty(sender) ? string.Empty : string.Format("Sender: {0}\r\n", sender);
            replyTo = string.IsNullOrEmpty(replyTo) ? string.Empty : string.Format("Reply-To: {0}\r\n", replyTo);

            string plainTextMime =
@"From: {0}
To: {1}
"
+ sender + cc + bcc + replyTo + @"Subject: {2}
Content-Type: text/plain; charset=""us-ascii""
MIME-Version: 1.0

{3}
";
            return Common.FormatString(plainTextMime, from, to, subject, body);
        }

        /// <summary>
        /// Create a MIME with two electronic voice mail attachments.
        /// Note: When writing X-AttachmentOrder header value, the attachment name retrieved from [secondVoiceAttachmentPath]
        /// will be added after the name retrieved from [firstVoiceAttachmentPath].
        /// </summary>
        /// <param name="from">The from address of mail</param>
        /// <param name="to">The to address of the mail</param>
        /// <param name="subject">The subject of the mail</param>
        /// <param name="body">The body content of the mail</param>
        /// <param name="callNumber">The telephone number of voice mail</param>
        /// <param name="firstVoiceFilePath">The path of first voice file</param>
        /// <param name="secondVoiceFilePath">The path of second voice file</param>
        /// <returns>Returns the corresponding sample voice email mime</returns>
        internal static string CreateVoiceAttachmentMime(
            string from,
            string to,
            string subject,
            string body,
            string callNumber,
            string firstVoiceFilePath,
            string secondVoiceFilePath)
        {
            // For simplicity, we don't calculate the duration of voice mp3
            string voiceAttachmentMime =
@"From: {0}
To: {1}
Subject: {2}
Content-Class: voice
X-CallingTelephoneNumber: {3}
X-VoiceMessageDuration: 2
X-AttachmentOrder: {4};{5}
Content-Type: multipart/mixed;
    boundary=""_001_MSASEMAIL_NextPart_""
MIME-Version: 1.0

--_001_MSASEMAIL_NextPart_
Content-Type: text/plain; charset=""us-ascii""

{6}

--_001_MSASEMAIL_NextPart_
Content-Type: audio/{7}
Content-Disposition: attachment; filename=""{8}""
Content-Transfer-Encoding: base64

{9}


--_001_MSASEMAIL_NextPart_
Content-Type: audio/{10}
Content-Disposition: attachment; filename=""{11}""
Content-Transfer-Encoding: base64

{12}

--_001_MSASEMAIL_NextPart_--
";
            string firstVoiceFileName = Path.GetFileName(firstVoiceFilePath);
            string firstVoiceContent = Convert.ToBase64String(File.ReadAllBytes(firstVoiceFilePath));
            string firstVoiceFileExt = Path.GetExtension(firstVoiceFilePath);

            string secondVoiceFileName = Path.GetFileName(secondVoiceFilePath);
            string secondVoiceContent = Convert.ToBase64String(File.ReadAllBytes(secondVoiceFilePath));
            string secondVoiceFileExt = Path.GetExtension(secondVoiceFilePath);

            return Common.FormatString(
                voiceAttachmentMime,
                from,
                to,
                subject,
                callNumber,
                firstVoiceFileName,
                secondVoiceFileName,
                body,
                firstVoiceFileExt.TrimStart('.'),
                firstVoiceFileName,
                firstVoiceContent,
                secondVoiceFileExt.TrimStart('.'),
                secondVoiceFileName,
                secondVoiceContent);
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
    boundary=""_001_MSASEMAIL_NextPart_""
MIME-Version: 1.0

--_001_MSASEMAIL_NextPart_
Content-Type: text/plain; charset=""us-ascii""

{3}

--_001_MSASEMAIL_NextPart_
Content-Type: text/calendar; charset=""us-ascii""; method=REQUEST

{4}

--_001_MSASEMAIL_NextPart_--

";
            return Common.FormatString(meetingRequestMime, from, to, subject, body, icalendarContent);
        }

        /// <summary>
        /// Create iCalendar format string from one calendar instance
        /// </summary>
        /// <param name="calendar">Calendar information</param>
        /// <returns>iCalendar formatted string</returns>
        internal static string CreateiCalendarFormatContent(Calendar calendar)
        {
            StringBuilder ical = new StringBuilder();
            ical.AppendLine("BEGIN: VCALENDAR");
            ical.AppendLine("PRODID:-//Microosft Protocols TestSuites");
            ical.AppendLine("VERSION:2.0");
            ical.AppendLine("METHOD:REQUEST");

            ical.AppendLine("X-MS-OLK-FORCEINSPECTOROPEN:TRUE");
            ical.AppendLine("BEGIN:VTIMEZONE");
            ical.AppendLine("TZID:UTC");
            ical.AppendLine("BEGIN:STANDARD");
            ical.AppendLine("DTSTART:16010101T000000");
            ical.AppendLine("TZOFFSETFROM:-0000");
            ical.AppendLine("TZOFFSETTO:-0000");
            ical.AppendLine("END:STANDARD");
            ical.AppendLine("END:VTIMEZONE");

            ical.AppendLine("BEGIN:VEVENT");
            ical.AppendLine("UID:" + calendar.UID);

            ical.AppendLine("DTSTAMP:" + ((DateTime)calendar.DtStamp).ToString("yyyyMMddTHHmmss"));
            ical.AppendLine("DESCRIPTION:" + calendar.Subject);
            ical.AppendLine("SUMMARY:" + calendar.Subject);
            ical.AppendLine("ATTENDEE;CN=\"\";RSVP=" + (calendar.ResponseRequested == true ? "TRUE" : "FALSE") + ":mailto:" + calendar.Attendees.Attendee[0].Email);
            ical.AppendLine("ORGANIZER:MAILTO:" + calendar.OrganizerEmail);
            ical.AppendLine("LOCATION:" + (calendar.Location ?? "My Office"));

            if (calendar.AllDayEvent == 1)
            {
                ical.AppendLine("DTSTART;VALUE=DATE:" + ((DateTime)calendar.StartTime).Date.ToString("yyyyMMdd"));
                ical.AppendLine("DTEND;VALUE=DATE:" + ((DateTime)calendar.EndTime).Date.ToString("yyyyMMdd"));
                ical.AppendLine("X-MICROSOFT-CDO-ALLDAYEVENT:TRUE");
            }
            else
            {
                ical.AppendLine("DTSTART;TZID=\"UTC\":" + ((DateTime)calendar.StartTime).ToString("yyyyMMddTHHmmss"));
                ical.AppendLine("DTEND;TZID=\"UTC\":" + ((DateTime)calendar.EndTime).ToString("yyyyMMddTHHmmss"));
            }

            if (calendar.Recurrence != null)
            {
                switch (calendar.Recurrence.Type)
                {
                    case 1:
                        ical.AppendLine("RRULE:FREQ=WEEKLY;BYDAY=MO;UNTIL=" + calendar.Recurrence.Until);
                        break;
                    case 2:
                        ical.AppendLine("RRULE:FREQ=MONTHLY;COUNT=3;BYMONTHDAY=1");
                        break;
                    case 3:
                        ical.AppendLine("RRULE:FREQ=MONTHLY;COUNT=3;BYDAY=1MO");
                        break;
                    case 5:
                        ical.AppendLine("RRULE:FREQ=YEARLY;COUNT=3;BYMONTHDAY=1;BYMONTH=1");
                        break;
                    case 6:
                        ical.AppendLine("RRULE:FREQ=YEARLY;COUNT=3;BYDAY=2MO;BYMONTH=1");
                        break;
                }
            }

            if (calendar.Exceptions != null)
            {
                ical.AppendLine("EXDATE;TZID=\"UTC\":" + ((DateTime)calendar.StartTime).AddDays(7).ToString("yyyyMMddTHHmmss"));
                ical.AppendLine("RECURRENCE-ID;TZID=\"UTC\":" + ((DateTime)calendar.StartTime).ToString("yyyyMMddTHHmmss"));
            }

            switch (calendar.BusyStatus)
            {
                case 0:
                    ical.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:FREE");
                    break;
                case 1:
                    ical.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:TENTATIVE");
                    break;
                case 2:
                    ical.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:BUSY");
                    break;
                case 3:
                    ical.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:OOF");
                    break;
            }

            if (calendar.DisallowNewTimeProposal == true)
            {
                ical.AppendLine("X-MICROSOFT-DISALLOW-COUNTER:TRUE");
            }

            if (calendar.Reminder.HasValue)
            {
                ical.AppendLine("BEGIN:VALARM");
                ical.AppendLine("TRIGGER:-PT" + calendar.Reminder + "M");
                ical.AppendLine("ACTION:DISPLAY");
                ical.AppendLine("END:VALARM");
            }

            ical.AppendLine("END:VEVENT");
            ical.AppendLine("END:VCALENDAR");
            return ical.ToString();
        }

        /// <summary>
        /// Create meeting response mail iCalendarContent
        /// </summary>
        /// <param name="startTime">Meeting start time</param>
        /// <param name="endTime">Meeting end time</param>
        /// <param name="uid">Meeting uid</param>
        /// <param name="subject">Meeting subject</param>
        /// <param name="location">Meeting location</param>
        /// <param name="replyto">Meeting organizer email to reply</param>
        /// <param name="from">Meeting attendee email address </param>
        /// <returns>Meeting Response content in iCalendar format</returns>
        internal static string CreateMeetingResponseiCalendarFormatContent(DateTime startTime, DateTime endTime, string uid, string subject, string location, string replyto, string from)
        {
            StringBuilder ical = new StringBuilder();
            ical.AppendLine("BEGIN: VCALENDAR");
            ical.AppendLine("PRODID:-//Microosft Protocols TestSuites");
            ical.AppendLine("VERSION:2.0");
            ical.AppendLine("METHOD:REPLY");
            ical.AppendLine("BEGIN:VEVENT");
            ical.AppendLine("DTSTART:" + startTime.ToString("yyyyMMddTHHmmss"));
            ical.AppendLine("DTSTAMP:" + startTime.ToString("yyyyMMddTHHmmss"));
            ical.AppendLine("DTEND:" + endTime.ToString("yyyyMMddTHHmmss"));
            ical.AppendLine("UID:" + uid);
            ical.AppendLine("DESCRIPTION:" + subject);
            ical.AppendLine("SUMMARY:" + subject);
            ical.AppendLine("ORGANIZER:MAILTO:" + replyto);
            ical.AppendLine("ATTENDEE;PARTSTAT=ACCEPTED;CN=\"\";RSVP=TRUE:" + from);
            ical.AppendLine("LOCATION:" + location);
            ical.AppendLine("END:VEVENT");
            ical.AppendLine("END:VCALENDAR");
            return ical.ToString();
        }

        #endregion

        #region Create some [MS-ASCMD] request required by test case
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
        ///       <Range>0-4</Range>
        ///       <DeepTraversal/>
        ///     </Options>
        ///   </Store>
        /// </Search>
        /// -->
        /// </summary>
        /// <param name="keyword">Specify a string value for which to search. Refer to [MS-ASCMD] section 2.2.3.73</param>
        /// <param name="folderCollectionId">Specify the folder in which to search. Refer to [MS-ASCMD] section 2.2.3.30.4</param>
        /// <returns>Returns a SearchRequest instance</returns>
        internal static SearchRequest CreateSearchRequest(string keyword, string folderCollectionId)
        {
            Request.SearchStore searchStore = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Options = new Request.Options1
                {
                    Items = new object[] { string.Empty, string.Empty },

                    ItemsElementName = new Request.ItemsChoiceType6[]
                    {
                        Request.ItemsChoiceType6.RebuildResults,
                        Request.ItemsChoiceType6.DeepTraversal
                    }
                }
            };

            // Build up query condition by using the keyword and folder CollectionID
            Request.queryType queryItem = new Request.queryType
            {
                Items = new object[] { folderCollectionId, keyword },

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
        /// Builds a Find request on the Mailbox store by using the specified keyword and folder collection ID
        /// In general, returns the XML formatted find request as follows:
        /// <!--
        /// <?xml version="1.0" encoding="utf-8"?>
        /// <Find xmlns="Find">
        /// <SearchId>30483e1c-e7a6-4096-ba2d-3c49caf77bd7</SearchId>
        /// <ExecuteSearch>
        ///   <MailBoxSearchCriterion>
        ///     <Query>
        ///       <Class xmlns="AirSync">Email</Class>
        ///       <airsync:CollectionId>5</airsync:CollectionId>
        ///       <FreeText>MSASEMAIL_S01_TC32_subject_051909_480252</FreeText>
        ///     </Query>
        ///     <Options>  
        ///       <Range>0-4</Range>
        ///       <DeepTraversal/>
        ///     </Options>
        ///   </MailBoxSearchCriterion>
        /// </ExecuteSearch>
        /// </Find>
        /// -->
        /// </summary>
        /// <param name="keyword">Specify a string value for which to search. Refer to [MS-ASCMD] section 2.2.3.73</param>
        /// <param name="folderCollectionId">Specify the folder in which to search. Refer to [MS-ASCMD] section 2.2.3.30.4</param>
        /// <returns>Returns a FindRequest instance</returns>
        internal static FindRequest CreateFindRequest(string folderCollectionId, string keyword)
        {
            Request.Find find = new Request.Find
            {
                SearchId = Guid.NewGuid().ToString(),
                ExecuteSearch = new Request.FindExecuteSearch
                {
                    Item = new Request.FindExecuteSearchMailBoxSearchCriterion
                    {
                        Query = new Request.queryType2
                        {
                            ItemsElementName = new Request.ItemsChoiceType11[] { Request.ItemsChoiceType11.Class, Request.ItemsChoiceType11.CollectionId, Request.ItemsChoiceType11.FreeText },
                            Items = new string[] { "Email", folderCollectionId, keyword}
                        },
                        Options = new Request.FindExecuteSearchMailBoxSearchCriterionOptions
                        {
                            Range = "0-5",
                            DeepTraversal = new Request.EmptyTag {}
                        }
                    },

                },
            };

            return Common.CreateFindRequest(find);
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
        /// <param name="serverId">Specify a unique identifier that is assigned by the server for a mailItem, which can be returned by ActiveSync Sync command(Refer to [MS-ASCMD]2.2.3.151.5)</param>
        /// <param name="bodyPreference">Sets preference information related to the type and size of information for body (Refer to [MS-ASAIRS] 2.2.2.7)</param>
        /// <param name="bodyPartPreference">Sets preference information related to the type and size of information for a message part (Refer to [MS-ASAIRS] 2.2.2.6)</param>
        /// <param name="schema">Sets the schema information.</param>
        /// <returns>Returns the ItemOperationsRequest instance</returns>
        internal static ItemOperationsRequest CreateItemOperationsFetchRequest(
            string collectionId,
            string serverId,
            Request.BodyPreference bodyPreference,
            Request.BodyPartPreference bodyPartPreference,
            Request.Schema schema)
        {
            // Set the ItemOperations:Fetch Options (See [MS-ASCMD] 2.2.3.115.2 Options)
            //  :: To get the whole content of the specified mail item, just ignore Fetch and Range element.
            //     See [MS-ASCMD] 2.2.3.145 Schema and 2.2.3.130.1 Range
            //  :: UserName, Password is not required by MS-ASEMAIl test case
            Request.ItemOperationsFetchOptions fetchOptions = new Request.ItemOperationsFetchOptions();
            List<object> fetchOptionItems = new List<object>();
            List<Request.ItemsChoiceType5> fetchOptionItemsName = new List<Request.ItemsChoiceType5>();

            if (null != bodyPreference)
            {
                fetchOptionItemsName.Add(Request.ItemsChoiceType5.BodyPreference);
                fetchOptionItems.Add(bodyPreference);

                // when body format is mime (Refer to  [MS-ASAIRS] 2.2.2.22 Type)
                if (bodyPreference.Type == 0x4)
                {
                    fetchOptionItemsName.Add(Request.ItemsChoiceType5.MIMESupport);

                    // Magic number '2' indicate server send MIME data for all messages but not S/MIME messages only
                    // (Refer to [MS-ASCMD] 2.2.3.100.1 MIMESupport)
                    fetchOptionItems.Add((byte)0x2);
                }
            }

            if (null != bodyPartPreference)
            {
                fetchOptionItemsName.Add(Request.ItemsChoiceType5.BodyPartPreference);
                fetchOptionItems.Add(bodyPartPreference);
            }

            if (null != schema)
            {
                fetchOptionItemsName.Add(Request.ItemsChoiceType5.Schema);
                fetchOptionItems.Add(schema);
            }

            fetchOptions.Items = fetchOptionItems.ToArray();
            fetchOptions.ItemsElementName = fetchOptionItemsName.ToArray();

            // *Only to fetch email item in mailbox* by using airsync:CollectionId & ServerId
            // So ignore LongId/LinkId/FileReference/RemoveRightsManagementProtection
            Request.ItemOperationsFetch fetchElement = new Request.ItemOperationsFetch()
            {
                CollectionId = collectionId,
                ServerId = serverId,
                Store = SearchName.Mailbox.ToString(),
                Options = fetchOptions
            };

            return Common.CreateItemOperationsRequest(new object[] { fetchElement });
        }

        /// <summary>
        /// Build a generic Sync request without command references by using the specified sync key, folder collection ID and body preference option.
        /// If syncKey is 
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
        ///       <WindowSize>100</WindowSize>
        ///       <Options>
        ///         <MIMESupport>0</MIMESupport>
        ///         <airsyncbase:BodyPreference>
        ///           <airsyncbase:Type>2</airsyncbase:Type>
        ///           <airsyncbase:TruncationSize>5120</airsyncbase:TruncationSize>
        ///         </airsyncbase:BodyPreference>
        ///       </Options>
        ///     </Collection>
        ///   </Collections>
        /// </Sync>
        /// -->
        /// </summary>
        /// <param name="syncKey">Specify the sync key obtained from the last sync response(Refer to [MS-ASCMD]2.2.3.166.4)</param>
        /// <param name="collectionId">Specify the server ID of the folder to be synchronized, which can be returned by ActiveSync FolderSync command(Refer to [MS-ASCMD]2.2.3.30.5)</param>
        /// <param name="bodyPreference">Sets preference information related to the type and size of information for body (Refer to [MS-ASAIRS] 2.2.2.7)</param>
        /// <returns>Returns the SyncRequest instance</returns>
        internal static SyncRequest CreateSyncRequest(string syncKey, string collectionId, Request.BodyPreference bodyPreference)
        {
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                SyncKey = syncKey,
                CollectionId = collectionId
            };

            if (syncKey != "0")
            {
                syncCollection.GetChanges = true;
                syncCollection.GetChangesSpecified = true;
            }

            syncCollection.WindowSize = "100";

            Request.Options syncOptions = new Request.Options();
            List<object> syncOptionItems = new List<object>();
            List<Request.ItemsChoiceType1> syncOptionItemsName = new List<Request.ItemsChoiceType1>();

            if (null != bodyPreference)
            {
                syncOptionItemsName.Add(Request.ItemsChoiceType1.BodyPreference);
                syncOptionItems.Add(bodyPreference);

                // when body format is mime (Refer to [MS-ASAIRS] 2.2.2.22 Type)
                if (bodyPreference.Type == 0x4)
                {
                    syncOptionItemsName.Add(Request.ItemsChoiceType1.MIMESupport);

                    // Magic number '2' indicate server send MIME data for all messages but not S/MIME messages only
                    syncOptionItems.Add((byte)0x2);
                }
            }

            syncOptions.Items = syncOptionItems.ToArray();
            syncOptions.ItemsElementName = syncOptionItemsName.ToArray();
            syncCollection.Options = new Request.Options[] { syncOptions };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
        }

        /// <summary>
        /// Builds a Sync change request by using the specified sync key, folder collection ID and change application data.
        /// In general, returns the XML formatted Sync request as follows:
        /// <!--
        /// <?xml version="1.0" encoding="utf-8"?>
        /// <Sync xmlns="AirSync">
        ///   <Collections>
        ///     <Collection>
        ///       <SyncKey>0</SyncKey>
        ///       <CollectionId>5</CollectionId>
        ///       <GetChanges>1</GetChanges>
        ///       <WindowSize>100</WindowSize>
        ///        <Commands>
        ///            <Change>
        ///                <ServerId>5:1</ServerId>
        ///                <ApplicationData>
        ///                    ...
        ///                </ApplicationData>
        ///            </Change>
        ///        </Commands>
        ///     </Collection>
        ///   </Collections>
        /// </Sync>
        /// -->
        /// </summary>
        /// <param name="syncKey">Specify the sync key obtained from the last sync response(Refer to [MS-ASCMD]2.2.3.166.4)</param>
        /// <param name="collectionId">Specify the server ID of the folder to be synchronized, which can be returned by ActiveSync FolderSync command(Refer to [MS-ASCMD]2.2.3.30.5)</param>
        /// <param name="changeData">Contains the data used to specify the Change element for Sync command(Refer to [MS-ASCMD]2.2.3.11)</param>
        /// <returns>Returns the SyncRequest instance</returns>
        internal static SyncRequest CreateSyncChangeRequest(string syncKey, string collectionId, Request.SyncCollectionChange changeData)
        {
            Request.SyncCollection syncCollection = CreateSyncCollection(syncKey, collectionId);
            syncCollection.Commands = new object[] { changeData };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
        }

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
        ///       <WindowSize>100</WindowSize>
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
        /// <param name="syncKey">Specify the sync key obtained from the last sync response(Refer to [MS-ASCMD]2.2.3.166.4)</param>
        /// <param name="collectionId">Specify the server ID of the folder to be synchronized, which can be returned by ActiveSync FolderSync command(Refer to [MS-ASCMD]2.2.3.30.5)</param>
        /// <param name="addData">Contains the data used to specify the Add element for Sync command(Refer to [MS-ASCMD]2.2.3.7.2)</param>
        /// <returns>Returns the SyncRequest instance</returns>
        internal static SyncRequest CreateSyncAddRequest(string syncKey, string collectionId, Request.SyncCollectionAdd addData)
        {
            Request.SyncCollection syncCollection = TestSuiteHelper.CreateSyncCollection(syncKey, collectionId);
            syncCollection.Commands = new object[] { addData };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
        }

        /// <summary>
        /// Create a sync add request.
        /// </summary>
        /// <param name="syncKey">Specify the sync key obtained from the last sync response</param>
        /// <param name="collectionId">Specify the server ID of the folder to be synchronized, which can be returned by ActiveSync FolderSync command.</param>
        /// <param name="applicationData">Contains the data used to specify the Add element for Sync command.</param>
        /// <returns>Returns the SyncRequest instance.</returns>
        internal static SyncRequest CreateSyncAddRequest(string syncKey, string collectionId, Request.SyncCollectionAddApplicationData applicationData)
        {
            SyncRequest syncAddRequest = TestSuiteHelper.CreateSyncRequest(syncKey, collectionId, null);
            Request.SyncCollectionAdd add = new Request.SyncCollectionAdd
            {
                ClientId = GetClientId(),
                ApplicationData = applicationData
            };

            List<object> commandList = new List<object> { add };

            syncAddRequest.RequestData.Collections[0].Commands = commandList.ToArray();

            return syncAddRequest;
        }

        /// <summary>
        /// Create an instance of SyncCollection
        /// </summary>
        /// <param name="syncKey">Specify the synchronization key obtained from the last sync command response.</param>
        /// <param name="collectionId">Specify the serverId of the folder to be synchronized, which can be returned by ActiveSync FolderSync command</param>
        /// <returns>An instance of SyncCollection</returns>
        internal static Request.SyncCollection CreateSyncCollection(string syncKey, string collectionId)
        {
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                SyncKey = syncKey,
                GetChanges = true,
                GetChangesSpecified = true,
                CollectionId = collectionId,
                WindowSize = "100"
            };

            return syncCollection;
        }

        /// <summary>
        /// Create an instance of SyncCollectionChange
        /// </summary>
        /// <param name="read">The value is TRUE indicates the email has been read; a value of FALSE indicates the email has not been read</param>
        /// <param name="serverId">The server id of the email</param>
        /// <param name="flag">The flag instance</param>
        /// <param name="categories">The list of categories</param>
        /// <returns>An instance of SyncCollectionChange</returns>
        internal static Request.SyncCollectionChange CreateSyncChangeData(bool read, string serverId, Request.Flag flag, Collection<object> categories)
        {
            Request.SyncCollectionChange changeData = new Request.SyncCollectionChange
            {
                ServerId = serverId,
                ApplicationData = new Request.SyncCollectionChangeApplicationData()
            };

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType7> itemsElementName = new List<Request.ItemsChoiceType7>();
            items.Add(read);
            itemsElementName.Add(Request.ItemsChoiceType7.Read);

            if (null != flag)
            {
                items.Add(flag);
                itemsElementName.Add(Request.ItemsChoiceType7.Flag);
            }

            if (null != categories)
            {
                Request.Categories mailCategories = new Request.Categories();
                List<string> category = new List<string>();
                foreach (object categoryObject in categories)
                {
                    category.Add(categoryObject.ToString());
                }

                mailCategories.Category = category.ToArray();
                items.Add(mailCategories);
                itemsElementName.Add(Request.ItemsChoiceType7.Categories2);
            }

            changeData.ApplicationData.Items = items.ToArray();
            changeData.ApplicationData.ItemsElementName = itemsElementName.ToArray();
            return changeData;
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
        /// <param name="copyToSentItems">Specify whether needs to store a mail copy to sent items</param>
        /// <param name="mime">Specify the mail mime</param>
        /// <returns>Returns the SendMailRequest instance</returns>
        internal static SendMailRequest CreateSendMailRequest(string clientId, bool copyToSentItems, string mime)
        {
            Request.SendMail sendMail = new Request.SendMail();

            // If true, save a copy to sent items folder, if false, doesn't save a copy to sent items folder
            if (copyToSentItems)
            {
                sendMail.SaveInSentItems = string.Empty;
            }

            sendMail.ClientId = clientId;
            sendMail.Mime = mime;

            SendMailRequest sendMailRequest = Common.CreateSendMailRequest();
            sendMailRequest.RequestData = sendMail;
            return sendMailRequest;
        }

        /// <summary>
        /// Builds a SmartReply request by using the specified source folder Id, source server Id and reply mime information.
        /// In general, returns the XML formatted Sync request as follows:
        /// <!--
        /// <?xml version="1.0" encoding="utf-8"?>
        /// <SmartReply xmlns="ComposeMail">
        ///   <ClientId>d7b99822-685a-4a40-8dfb-87a114926986</ClientId>
        ///   <Source>
        ///     <FolderId>5</FolderId>
        ///     <ItemId>5:1</ItemId>
        ///   </Source>
        ///   <Mime>...</Mime>
        /// </SmartReply>
        /// -->
        /// </summary>
        /// <param name="sourceFolderId">Specify the folder id of original mail item being replied</param>
        /// <param name="sourceServerId">Specify the server Id of original mail item being replied</param>
        /// <param name="replyMime">The total reply mime</param>
        /// <returns>Returns the SmartReplyRequest instance</returns>
        internal static SmartReplyRequest CreateSmartReplyRequest(string sourceFolderId, string sourceServerId, string replyMime)
        {
            SmartReplyRequest request = new SmartReplyRequest
            {
                RequestData = new Request.SmartReply
                {
                    ClientId = System.Guid.NewGuid().ToString(),
                    Source = new Request.Source { FolderId = sourceFolderId, ItemId = sourceServerId },
                    Mime = replyMime
                }
            };

            request.SetCommandParameters(new Dictionary<CmdParameterName, object>
            {
                {
                    CmdParameterName.CollectionId, sourceFolderId
                },
                {
                    CmdParameterName.ItemId, sourceServerId
                }
            });

            return request;
        }

        /// <summary>
        /// Builds a SmartForward request by using the specified source folder Id, source server Id and forward mime information.
        /// In general, returns the XML formatted Sync request as follows:
        /// <!--
        /// <?xml version="1.0" encoding="utf-8"?>
        /// <SmartForward xmlns="ComposeMail">
        ///   <ClientId>d7b99822-685a-4a40-8dfb-87a114926986</ClientId>
        ///   <Source>
        ///     <FolderId>5</FolderId>
        ///     <ItemId>5:1</ItemId>
        ///   </Source>
        ///   <Mime>...</Mime>
        /// </SmartForward>
        /// -->
        /// </summary>
        /// <param name="sourceFolderId">Specify the folder id of original mail item being forwarded</param>
        /// <param name="sourceServerId">Specify the server Id of original mail item being forwarded</param>
        /// <param name="forwardMime">The total forward mime</param>
        /// <returns>Returns the SmartReplyRequest instance</returns>
        internal static SmartForwardRequest CreateSmartForwardRequest(string sourceFolderId, string sourceServerId, string forwardMime)
        {
            SmartForwardRequest request = new SmartForwardRequest
            {
                RequestData = new Request.SmartForward
                {
                    ClientId = System.Guid.NewGuid().ToString(),
                    Source = new Request.Source { FolderId = sourceFolderId, ItemId = sourceServerId },
                    Mime = forwardMime
                }
            };

            request.SetCommandParameters(new Dictionary<CmdParameterName, object>
            {
                {
                    CmdParameterName.CollectionId, sourceFolderId
                },
                {
                    CmdParameterName.ItemId, sourceServerId
                }
            });

            return request;
        }
        #endregion

        /// <summary>
        /// Get the specified email item from the sync add response by using the subject as the search criteria.
        /// </summary>
        /// <param name="syncResult">The sync result.</param>
        /// <param name="subject">The email subject.</param>
        /// <returns>Return the specified email item.</returns>
        internal static Sync GetSyncAddItem(SyncStore syncResult, string subject)
        {
            Sync item = null;

            if (syncResult.AddElements != null)
            {
                foreach (Sync syncItem in syncResult.AddElements)
                {
                    if (syncItem.Email.Subject == subject)
                    {
                        item = syncItem;
                        break;
                    }

                    if (syncItem.Calendar.Subject == subject)
                    {
                        item = syncItem;
                        break;
                    }
                }
            }

            return item;
        }

        /// <summary>
        /// Get the specified email item from the sync change response by using the subject as the search criteria.
        /// </summary>
        /// <param name="syncResult">The sync result.</param>
        /// <param name="serverId">The email server id.</param>
        /// <returns>Return the specified email item.</returns>
        internal static Sync GetSyncChangeItem(SyncStore syncResult, string serverId)
        {
            Sync item = null;
            if (syncResult.ChangeElements != null)
            {
                foreach (Sync syncItem in syncResult.ChangeElements)
                {
                    if (syncItem.ServerId == serverId)
                    {
                        item = syncItem;
                        break;
                    }
                }
            }

            return item;
        }

        /// <summary>
        /// Get the Status code from the sync change response string
        /// </summary>
        /// <param name="changeResponseXml">Sync Response string in Xml format</param>
        /// <returns>Status code from change response</returns>
        internal static string GetStatusCode(string changeResponseXml)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(changeResponseXml);
            XmlNamespaceManager nameSpaceManager = new XmlNamespaceManager(doc.NameTable);
            nameSpaceManager.AddNamespace("e", "AirSync");

            // Get status code from <Change> element
            XmlNode status = doc.SelectSingleNode("//e:Change/e:Status", nameSpaceManager);
            if (status == null)
            {
                return null;
            }
            else
            {
                return status.InnerText;
            }
        }

        /// <summary>
        /// Get the email item from the ItemOperations response by using the subject as the search criteria.
        /// </summary>
        /// <param name="itemOperationsResult">The ItemOperations command result.</param>
        /// <param name="subject">The email subject.</param>
        /// <returns>Return the right email item from the sync response.</returns>
        internal static ItemOperations GetItemOperationsItem(ItemOperationsStore itemOperationsResult, string subject)
        {
            ItemOperations item = null;
            if (itemOperationsResult.Items != null)
            {
                foreach (ItemOperations itemOperationsItem in itemOperationsResult.Items)
                {
                    if (itemOperationsItem.Email.Subject == subject)
                    {
                        item = itemOperationsItem;
                        break;
                    }
                }
            }

            return item;
        }

        /// <summary>
        /// Create a Request.Attendees instance by the attendee email address.
        /// </summary>
        /// <param name="attendeeEmail">The email address of the attendee</param>
        /// <returns>The created Request.Attendees instance.</returns>
        internal static Request.Attendees CreateAttendees(string attendeeEmail)
        {
            List<Request.AttendeesAttendee> attendeelist = new List<Request.AttendeesAttendee>
            {
                new Request.AttendeesAttendee
                {
                    Email = attendeeEmail,
                    Name = attendeeEmail,
                    AttendeeStatus = 0,
                    AttendeeTypeSpecified = true,
                    AttendeeType = 1
                }
            };

            Request.Attendees attendees = new Request.Attendees { Attendee = attendeelist.ToArray() };
            return attendees;
        }

        /// <summary>
        /// Convert a byte array into an equivalent hex string. E.g. [0xXX, 0xYY]=>"XXYY"
        /// </summary>
        /// <param name="bytes">The byte array need to be converted</param>
        /// <returns>A hex string representation corresponding to the byte array</returns>
        internal static string BytesToHex(byte[] bytes)
        {
            return BytesToHex(bytes, 0, bytes.Length);
        }

        /// <summary>
        /// Convert a byte array into an equivalent hex string at the specified starting position and count
        /// </summary>
        /// <param name="bytes">The byte array need to be converted</param>
        /// <param name="startIndex">start position</param>
        /// <param name="count">The bytes count need to be converted</param>
        /// <returns>A hex string representation corresponding to the byte array</returns>
        internal static string BytesToHex(byte[] bytes, int startIndex, int count)
        {
            if (null == bytes || 0 == bytes.Length || startIndex < 0 || startIndex >= bytes.Length)
            {
                return string.Empty;
            }

            int endIndex = startIndex + count > bytes.Length ? bytes.Length : startIndex + count;

            char[] result = new char[(endIndex - startIndex) * 2];
            int resultIndex = 0;
            byte nib;

            // (0x1A) => (0x1, 0xA) => (0x1 + 0x30, 0xA + 0x37) => ('1','A')
            for (int i = startIndex; i < endIndex; i++)
            {
                nib = (byte)(bytes[i] >> 4);
                result[resultIndex++] = (char)(nib > 9 ? nib + 0x37 : nib + 0x30);
                nib = (byte)(bytes[i] & 0xF);
                result[resultIndex++] = (char)(nib > 9 ? nib + 0x37 : nib + 0x30);
            }

            return new string(result);
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
        /// Generate an outlook id.
        /// </summary>
        /// <param name="createTime">The time indicates when the outlook id is created.</param>
        /// <returns>An outlook id.</returns>
        internal static string GenerateOutlookID(DateTime createTime)
        {
            // EncodedGlobalId        = Header GlobalIdData
            // ThirdPartyGlobalId     = 1*UTF8-octets     ; Assuming UTF-8 is the encoding
            // 
            // Header = ByteArrayID InstanceDate CreationDateTime Padding DataSize
            // 
            // ByteArrayID           = "040000008200E00074C5B7101A82E008"
            // InstanceDate          = InstanceYear InstanceMonth InstanceDay
            // InstanceYear          = 4*4HEXDIG     ; UInt16
            // InstanceMonth         = 2*2HEXDIG     ; UInt8
            // InstanceDay           = 2*2HEXDIG     ; UInt8
            // CreationDateTime      = FileTime 
            // FileTime              = 16*16HEXDIG   ; UInt64
            // Padding               = 16*16HEXDIG   ; "0000000000000000" recommended
            // DataSize              = 8*8HEXDIG     ; UInt32 little-endian
            // GlobalIdData          = 2*HEXDIG
            StringBuilder uidBuilder = new StringBuilder();

            // ByteArrayID
            uidBuilder.Append("040000008200E00074C5B7101A82E008");

            // InstanceDate
            uidBuilder.Append("00000000");
            byte[] timeBytes = BitConverter.GetBytes(createTime.ToFileTimeUtc());

            // CreationDateTime
            uidBuilder.Append(TestSuiteHelper.BytesToHex(timeBytes));

            // Padding
            uidBuilder.Append("0000000000000000");
            byte[] gloabalIdData = Guid.NewGuid().ToByteArray();
            byte[] dataSizeBytes = BitConverter.GetBytes(gloabalIdData.Length);

            // DataSize
            uidBuilder.Append(TestSuiteHelper.BytesToHex(dataSizeBytes));

            // GlobalIdData
            uidBuilder.Append(TestSuiteHelper.BytesToHex(gloabalIdData));

            return uidBuilder.ToString();
        }

        /// <summary>
        /// Set the value of common meeting properties
        /// </summary>
        /// <param name="subject">The subject of the meeting.</param>
        /// <param name="attendeeEmailAddress">The email address of attendee.</param>
        /// <returns>The key and value pairs of common meeting properties.</returns>
        internal static Dictionary<Request.ItemsChoiceType8, object> SetMeetingProperties(string subject, string attendeeEmailAddress, ITestSite testSite)
        {
            Dictionary<Request.ItemsChoiceType8, object> propertiesToValueMap = new Dictionary<Request.ItemsChoiceType8, object>
            {
                {
                    Request.ItemsChoiceType8.Subject, subject
                }
            };

            // Set the subject element.

            // MeetingStauts is set to 1, which means it is a meeting and the user is the meeting organizer.
            byte meetingStatus = 1;
            propertiesToValueMap.Add(Request.ItemsChoiceType8.MeetingStatus, meetingStatus);

            // Set the UID
            string uID = Guid.NewGuid().ToString();
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", testSite).Equals("16.0")|| Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", testSite).Equals("16.1"))
            {
                propertiesToValueMap.Add(Request.ItemsChoiceType8.ClientUid, uID);
            }
            else
            {
                propertiesToValueMap.Add(Request.ItemsChoiceType8.UID, uID);
            }

            // Set the TimeZone
            string timeZone = Common.GetTimeZone("(UTC) Coordinated Universal Time", 0);
            propertiesToValueMap.Add(Request.ItemsChoiceType8.Timezone, timeZone);

            // Set the attendee to user2
            Request.Attendees attendees = TestSuiteHelper.CreateAttendees(attendeeEmailAddress);
            propertiesToValueMap.Add(Request.ItemsChoiceType8.Attendees, attendees);

            return propertiesToValueMap;
        }

        /// <summary>
        /// Create one sample calendar object
        /// </summary>
        /// <param name="subject">Meeting subject</param>
        /// <param name="organizerEmailAddress">Meeting organizer email address</param>
        /// <param name="attendeeEmailAddress">Meeting attendee email address</param>
        /// <returns>One sample calendar object</returns>
        internal static Calendar CreateDefaultCalendar(string subject, string organizerEmailAddress, string attendeeEmailAddress)
        {
            Calendar calendar = new Calendar
            {
                Timezone = Common.GetTimeZone("(UTC) Coordinated Universal Time", 0),
                DtStamp = DateTime.UtcNow,
                StartTime = DateTime.UtcNow.AddHours(1),
                Subject = subject,
                UID = Guid.NewGuid().ToString(),
                OrganizerEmail = organizerEmailAddress,
                OrganizerName = organizerEmailAddress,
                Location = "My Office",
                EndTime = DateTime.UtcNow.AddHours(2),
                Sensitivity = 0,
                BusyStatus = 1,
                AllDayEvent = 0
            };

            List<Response.AttendeesAttendee> attendeelist = new List<Response.AttendeesAttendee>
            {
                new Response.AttendeesAttendee
                {
                    Email = attendeeEmailAddress,
                    Name = attendeeEmailAddress,
                    AttendeeStatus = 0,
                    AttendeeType = 1
                }
            };

            calendar.Attendees = new Response.Attendees { Attendee = attendeelist.ToArray() };

            return calendar;
        }

        /// <summary>
        /// Create iCalendar format string from one calendar instance for cancel one occurrence of a meeting request
        /// </summary>
        /// <param name="calendar">Calendar information</param>
        /// <returns>iCalendar formatted string</returns>
        internal static string CreateiCalendarFormatCancelContent(Calendar calendar)
        {
            StringBuilder ical = new StringBuilder();
            ical.AppendLine("BEGIN: VCALENDAR");
            ical.AppendLine("PRODID:-//Microosft Protocols TestSuites");
            ical.AppendLine("VERSION:2.0");
            ical.AppendLine("METHOD:CANCEL");
            ical.AppendLine("X-MS-OLK-FORCEINSPECTOROPEN:TRUE");
            ical.AppendLine("BEGIN:VTIMEZONE");
            ical.AppendLine("TZID:UTC");
            ical.AppendLine("BEGIN:STANDARD");
            ical.AppendLine("DTSTART:16010101T000000");
            ical.AppendLine("TZOFFSETFROM:-0000");
            ical.AppendLine("TZOFFSETTO:-0000");
            ical.AppendLine("END:STANDARD");
            ical.AppendLine("BEGIN:DAYLIGHT");
            ical.AppendLine("DTSTART:16010311T020000");
            ical.AppendLine("TZOFFSETFROM:-0000");
            ical.AppendLine("TZOFFSETTO:+0000");
            ical.AppendLine("END:DAYLIGHT");
            ical.AppendLine("END:VTIMEZONE");
            ical.AppendLine("BEGIN:VEVENT");
            ical.AppendLine("ATTENDEE;CN=\"\";RSVP=" + (calendar.ResponseRequested == true ? "TRUE" : "FALSE") + ":mailto:" + calendar.Attendees.Attendee[0].Email);
            ical.AppendLine("PUBLIC");
            ical.AppendLine("CREATED:" + ((DateTime)calendar.DtStamp).ToString("yyyyMMddTHHmmss"));
            ical.AppendLine("DESCRIPTION:" + calendar.Subject);
            ical.AppendLine("DTEND;TZID=\"UTC\":" + calendar.Exceptions.Exception[0].EndTime);
            ical.AppendLine("DTSTART;TZID=\"UTC\":" + calendar.Exceptions.Exception[0].StartTime);
            ical.AppendLine("LOCATION:" + calendar.Location);
            ical.AppendLine("ORGANIZER:MAILTO:" + calendar.OrganizerEmail);
            ical.AppendLine("RECURRENCE-ID;TZID=\"UTC\":" + calendar.Exceptions.Exception[0].ExceptionStartTime);
            ical.AppendLine("SUMMARY: Canceled: " + calendar.Subject);
            ical.AppendLine("UID:" + calendar.UID);
            switch (calendar.BusyStatus)
            {
                case 0:
                    ical.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:FREE");
                    break;
                case 1:
                    ical.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:TENTATIVE");
                    break;
                case 2:
                    ical.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:BUSY");
                    break;
                case 3:
                    ical.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:OOF");
                    break;
            }

            if (calendar.DisallowNewTimeProposal == true)
            {
                ical.AppendLine("X-MICROSOFT-DISALLOW-COUNTER:TRUE");
            }

            if (calendar.Recurrence != null)
            {
                switch (calendar.Recurrence.Type)
                {
                    case 1:
                        ical.AppendLine("RRULE:FREQ=WEEKLY;BYDAY=MO;UNTIL=" + calendar.Recurrence.Until);
                        break;
                    case 2:
                        ical.AppendLine("RRULE:FREQ=MONTHLY;COUNT=3;BYMONTHDAY=1");
                        break;
                    case 3:
                        ical.AppendLine("RRULE:FREQ=MONTHLY;COUNT=3;BYDAY=1MO");
                        break;
                    case 5:
                        ical.AppendLine("RRULE:FREQ=YEARLY;COUNT=3;BYMONTHDAY=1;BYMONTH=1");
                        break;
                    case 6:
                        ical.AppendLine("RRULE:FREQ=YEARLY;COUNT=3;BYDAY=2MO;BYMONTH=1");
                        break;
                }
            }

            ical.AppendLine("END:VEVENT");
            ical.AppendLine("END:VCALENDAR");
            return ical.ToString();
        }
    }
}