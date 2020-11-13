namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System.Xml;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// A static class contains all helper methods used in test cases.
    /// </summary>
    internal static class TestSuiteHelper
    {
        #region Build command request

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

        /// <summary>
        /// Build a Sync command request.
        /// </summary>
        /// <param name="syncKey">The current sync key.</param>
        /// <param name="collectionId">The collection id which to sync with.</param>
        /// <param name="commands">The sync commands.</param>
        /// <param name="bodyPreferences">The bodyPreference in the options element.</param>
        /// <param name="bodyPartPreferences">The bodyPartPreference in the options element.</param>
        /// <returns>A Sync command request.</returns>
        internal static SyncRequest CreateSyncRequest(string syncKey, string collectionId, object[] commands, Request.BodyPreference[] bodyPreferences, Request.BodyPartPreference[] bodyPartPreferences)
        {
            SyncRequest request = new SyncRequest
            {
                RequestData =
                {
                    Collections = new Request.SyncCollection[]
                    {
                        new Request.SyncCollection()
                        {
                            SyncKey = syncKey,
                            CollectionId = collectionId
                        }
                    }
                }
            };

            request.RequestData.Collections[0].Commands = commands;

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType1> itemsElementName = new List<Request.ItemsChoiceType1>();

            if (bodyPreferences != null)
            {
                foreach (Request.BodyPreference bodyPreference in bodyPreferences)
                {
                    items.Add(bodyPreference);
                    itemsElementName.Add(Request.ItemsChoiceType1.BodyPreference);

                    // Include the MIMESupport element in request to retrieve the MIME body
                    if (bodyPreference.Type == 4)
                    {
                        items.Add((byte)2);
                        itemsElementName.Add(Request.ItemsChoiceType1.MIMESupport);
                    }
                }
            }

            if (bodyPartPreferences != null)
            {
                foreach (Request.BodyPartPreference bodyPartPreference in bodyPartPreferences)
                {
                    items.Add(bodyPartPreference);
                    itemsElementName.Add(Request.ItemsChoiceType1.BodyPartPreference);
                }
            }

            if (items.Count > 0)
            {
                request.RequestData.Collections[0].Options = new Request.Options[]
                {
                    new Request.Options()
                    {
                        ItemsElementName = itemsElementName.ToArray(),
                        Items = items.ToArray()
                    }
                };
            }

            return request;
        }

        /// <summary>
        /// Build an ItemOperations command request.
        /// </summary>
        /// <param name="collectionId">The collection id.</param>
        /// <param name="serverId">The server id of the mail.</param>
        /// <param name="fileReference">The file reference of the attachment.</param>
        /// <param name="bodyPreferences">The bodyPreference in the options element.</param>
        /// <param name="bodyPartPreferences">The bodyPartPreference in the options element.</param>
        /// <returns>An ItemOperations command request.</returns>
        internal static ItemOperationsRequest CreateItemOperationsRequest(string collectionId, string serverId, string fileReference, Request.BodyPreference[] bodyPreferences, Request.BodyPartPreference[] bodyPartPreferences)
        {
            ItemOperationsRequest request = new ItemOperationsRequest { RequestData = new Request.ItemOperations() };
            Request.ItemOperationsFetch fetch = new Request.ItemOperationsFetch { Store = SearchName.Mailbox.ToString() };

            if (fileReference != null)
            {
                fetch.FileReference = fileReference;
            }
            else
            {
                fetch.CollectionId = collectionId;
                fetch.ServerId = serverId;
            }

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType5> itemsElementName = new List<Request.ItemsChoiceType5>();

            if (bodyPreferences != null)
            {
                foreach (Request.BodyPreference bodyPreference in bodyPreferences)
                {
                    items.Add(bodyPreference);
                    itemsElementName.Add(Request.ItemsChoiceType5.BodyPreference);

                    // Include the MIMESupport element in request to retrieve the MIME body
                    if (bodyPreference.Type == 4)
                    {
                        items.Add((byte)2);
                        itemsElementName.Add(Request.ItemsChoiceType5.MIMESupport);
                    }
                }
            }

            if (bodyPartPreferences != null)
            {
                foreach (Request.BodyPartPreference bodyPartPreference in bodyPartPreferences)
                {
                    items.Add(bodyPartPreference);
                    itemsElementName.Add(Request.ItemsChoiceType5.BodyPartPreference);
                }
            }

            if (items.Count > 0)
            {
                fetch.Options = new Request.ItemOperationsFetchOptions()
                {
                    ItemsElementName = itemsElementName.ToArray(),
                    Items = items.ToArray()
                };
            }

            request.RequestData.Items = new object[] { fetch };

            return request;
        }

        /// <summary>
        /// Build a Search request.
        /// </summary>
        /// <param name="query">The query string.</param>
        /// <param name="collectionId">The collection id of searched folder.</param>
        /// <param name="conversationId">The conversation for which to search.</param>
        /// <param name="bodyPreferences">The bodyPreference in the options element.</param>
        /// <param name="bodyPartPreferences">The bodyPartPreference in the options element.</param>
        /// <returns>A Search command request.</returns>
        internal static SearchRequest CreateSearchRequest(string query, string collectionId, string conversationId, Request.BodyPreference[] bodyPreferences, Request.BodyPartPreference[] bodyPartPreferences)
        {
            SearchRequest request = new SearchRequest
            {
                RequestData =
                {
                    Items = new Request.SearchStore[]
                    {
                        new Request.SearchStore()
                        {
                            Name = SearchName.Mailbox.ToString(),
                            Query = new Request.queryType()
                            {
                                Items = new object[]
                                {
                                    new Request.queryType()
                                    {
                                        Items = new object[]
                                        {
                                            collectionId,
                                            query,
                                            conversationId
                                        },
                                        ItemsElementName = new Request.ItemsChoiceType2[]
                                        {
                                            Request.ItemsChoiceType2.CollectionId,
                                            Request.ItemsChoiceType2.FreeText,
                                            Request.ItemsChoiceType2.ConversationId
                                        }
                                    }
                                },
                                ItemsElementName = new Request.ItemsChoiceType2[]
                                {
                                    Request.ItemsChoiceType2.And
                                }
                            }
                        }
                    }
                }
            };

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType6> itemsElementName = new List<Request.ItemsChoiceType6>();

            if (bodyPreferences != null)
            {
                foreach (Request.BodyPreference bodyPreference in bodyPreferences)
                {
                    items.Add(bodyPreference);
                    itemsElementName.Add(Request.ItemsChoiceType6.BodyPreference);

                    // Include the MIMESupport element in request to retrieve the MIME body
                    if (bodyPreference.Type == 4)
                    {
                        items.Add((byte)2);
                        itemsElementName.Add(Request.ItemsChoiceType6.MIMESupport);
                    }
                }
            }

            if (bodyPartPreferences != null)
            {
                foreach (Request.BodyPartPreference bodyPartPreference in bodyPartPreferences)
                {
                    items.Add(bodyPartPreference);
                    itemsElementName.Add(Request.ItemsChoiceType6.BodyPartPreference);
                }
            }

            items.Add(string.Empty);
            itemsElementName.Add(Request.ItemsChoiceType6.RebuildResults);
            items.Add("0-9");
            itemsElementName.Add(Request.ItemsChoiceType6.Range);
            items.Add(string.Empty);
            itemsElementName.Add(Request.ItemsChoiceType6.DeepTraversal);

            request.RequestData.Items[0].Options = new Request.Options1()
            {
                ItemsElementName = itemsElementName.ToArray(),
                Items = items.ToArray()
            };

            return request;
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

        #endregion

        /// <summary>
        /// Create an instance of SyncCollection.
        /// </summary>
        /// <param name="syncKey">Specify the synchronization key obtained from the last sync command response.</param>
        /// <param name="collectionId">Specify the serverId of the folder to be synchronized, which can be returned by ActiveSync FolderSync command.</param>
        /// <returns>An instance of SyncCollection.</returns>
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

        #region Create MIME for SendMail command
        /// <summary>
        /// Create MIME for SendMail command.
        /// </summary>
        /// <param name="type">The email message body type.</param>
        /// <param name="from">The email address of sender.</param>
        /// <param name="to">The email address of recipient.</param>
        /// <param name="subject">The email subject.</param>
        /// <param name="body">The email body content.</param>
        /// <returns>A MIME for SendMail command.</returns>
        internal static string CreateMIME(EmailType type, string from, string to, string subject, string body)
        {
            string mime = null;
            string winmailData = null;

            // Create a plain text MIME
            if (type == EmailType.Plaintext)
            {
                mime =
@"From: {0}
To: {1}
Subject: {2}
Content-Type: text/plain; charset=""us-ascii""
MIME-Version: 1.0

{3}
";
            }

            // Create an HTML MIME
            if (type == EmailType.HTML)
            {
                mime =
@"From: {0}
To: {1}
Subject: {2}
Content-Type: text/html; charset=""us-ascii""
MIME-Version: 1.0

<html>
<body>
<font color=""blue"">{3}</font>
</body>
</html>
";
            }

            // Create a MIME with normal attachment
            if (type == EmailType.NormalAttachment)
            {
                mime =
@"From: {0}
To: {1}
Subject: {2}
Content-Type: multipart/mixed; boundary=""_boundary_""; type=""text/html""
MIME-Version: 1.0

--_boundary_
Content-Type: text/html; charset=""iso-8859-1""
Content-Transfer-Encoding: quoted-printable

<html><body>{3}<img width=""128"" height=""94"" id=""Picture_x0020_1"" src=""cid:i=
mage001.jpg@01CC1FB3.2053ED80"" alt=""Description: cid:ebdc14bd-deb4-4816-b=
00b-6e2a46097d17""></body></html>

--_boundary_
Content-Type: image/jpeg; name=""number1.jpg""
Content-ID: {4}
Content-Description: number1.jpg
Content-Disposition: inline; size=4; filename=""number1.jpg""
Content-Location: <cid:ebdc14bd-deb4-4816-b00b-6e2a46097d17>
Content-Transfer-Encoding: base64

MQ==

--_boundary_--
";
            }

            // Create a MIME with embedded attachment
            if (type == EmailType.EmbeddedAttachment)
            {
                mime =
@"From: {0}
To: {1}
Subject: {2}
Content-Type: multipart/mixed; boundary=""_boundary_""; type=""text/html""
MIME-Version: 1.0

--_boundary_
Content-Type: text/html; charset=""iso-8859-1""
Content-Transfer-Encoding: quoted-printable

<html><body>{3}</body></html>

--_boundary_
Content-Type: message/rfc822; name=""Embedded mail""
Content-Description: Embedded mail
Content-Disposition: attachment; size=4; filename=""Embedded mail""
Content-Transfer-Encoding: base64

MQ==

--_boundary_--
";
            }

            // Create a MIME with OLE attachment
            if (type == EmailType.AttachOLE)
            {
                winmailData = Convert.ToBase64String(File.ReadAllBytes("winmail.dat"));

                // Split lines, for the maximum length of each line in MIME is no more than 76 characters
                for (int i = 1; i < winmailData.Length / 76; i++)
                {
                    winmailData = winmailData.Insert((76 * i) - 1, "\r\n");
                }

                // The string "contoso.com" is just a sample domain name, it has no relationship to the domain configured in deployment.ptfconfig file, and any changes of this string will lead to the update of winmail.dat file.
                mime =
@"From: {0}
To: {1}
Subject: {2}
MIME-Version: 1.0
Content-Type: multipart/mixed;
    boundary=""_boundary_""
X-MS-TNEF-Correlator: <15CFAB655027B944AD65A26C7A6F2D7A0126D4994B34@DC01.contoso.com>

{3} 

--_boundary_
Content-Type: text/plain;
    charset=""us-ascii""
Content-Transfer-Encoding: 7bit

{3}  

--_boundary_
Content-Type: application/ms-tnef;
    name=""winmail.dat""
Content-Transfer-Encoding: base64
Content-Disposition: attachment;
    filename=""winmail.dat""

{5}

--_boundary_--";
            }

            return Common.FormatString(mime, from, to, subject, body, Guid.NewGuid().ToString(), winmailData);
        }
        #endregion

        /// <summary>
        /// Get the specified email item from the sync add response by using the subject.
        /// </summary>
        /// <param name="syncStore">The sync result.</param>
        /// <param name="subject">The email subject.</param>
        /// <returns>Return the specified email item.</returns>
        internal static DataStructures.Sync GetSyncAddItem(DataStructures.SyncStore syncStore, string subject)
        {
            DataStructures.Sync item = null;

            if (syncStore.AddElements != null)
            {
                foreach (DataStructures.Sync syncItem in syncStore.AddElements)
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

                    if (syncItem.Contact.FileAs == subject)
                    {
                        item = syncItem;
                        break;
                    }
                }
            }

            return item;
        }

        /// <summary>
        /// Get the email item from the Search response by using the subject as the search criteria.
        /// </summary>
        /// <param name="searchStore">The Search command result.</param>
        /// <param name="subject">The email subject.</param>
        /// <returns>The email item corresponds to the specified subject.</returns>
        internal static DataStructures.Search GetSearchItem(DataStructures.SearchStore searchStore, string subject)
        {
            DataStructures.Search searchItem = null;
            if (searchStore.Results.Count > 0)
            {
                foreach (DataStructures.Search item in searchStore.Results)
                {
                    if (item.Email.Subject == subject)
                    {
                        searchItem = item;
                        break;
                    }

                    if (item.Calendar.Subject == subject)
                    {
                        searchItem = item;
                        break;
                    }
                }
            }

            return searchItem;
        }

        /// <summary>
        /// Truncate data according to the specified length.
        /// </summary>
        /// <param name="originalData">The original data.</param>
        /// <param name="length">The length of the byte array.</param>
        /// <returns>The truncated data.</returns>
        internal static string TruncateData(string originalData, int length)
        {
            byte[] bytes = System.Text.Encoding.ASCII.GetBytes(originalData);
            byte[] truncatedBytes = new byte[length];
            for (int i = 0; i < length; i++)
            {
                truncatedBytes[i] = bytes[i];
            }

            return System.Text.Encoding.ASCII.GetString(truncatedBytes);
        }

        /// <summary>
        /// Get the inner text of specified element.
        /// </summary>
        /// <param name="lastRawResponse">The raw xml response.</param>
        /// <param name="parentNodeName">The parent element of the specified node.</param>
        /// <param name="nodeName">The name of the node.</param>
        /// <param name="subject">The subject of the specified item.</param>
        /// <returns>The inner text of the specified element.</returns>
        internal static string GetDataInnerText(XmlElement lastRawResponse, string parentNodeName, string nodeName, string subject)
        {
            string data = null;
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(lastRawResponse.OuterXml);
            XmlNodeList subjectElementNodes = doc.SelectNodes("//*[name()='Subject']");
            for (int i = 0; i < subjectElementNodes.Count; i++)
            {
                if (subjectElementNodes[i].InnerText == subject)
                {
                    XmlNodeList bodyElementNodes = doc.SelectNodes("//*[name()='" + parentNodeName + "']");
                    XmlNodeList dataElementNodes = bodyElementNodes[i].SelectNodes("*[name()='" + nodeName + "']");
                    data = dataElementNodes[0].InnerText;
                    break;
                }
            }

            return data;
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
                    AttendeeTypeSpecified=true,
                    AttendeeType = 1
                }
            };

            Request.Attendees attendees = new Request.Attendees { Attendee = attendeelist.ToArray() };
            return attendees;
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


    }
}