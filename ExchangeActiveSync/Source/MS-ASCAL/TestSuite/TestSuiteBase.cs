namespace Microsoft.Protocols.TestSuites.MS_ASCAL
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Calendar = Microsoft.Protocols.TestSuites.Common.DataStructures.Calendar;
    using ItemOperationsStore = Microsoft.Protocols.TestSuites.Common.DataStructures.ItemOperationsStore;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;
    using SyncItem = Microsoft.Protocols.TestSuites.Common.DataStructures.Sync;
    using SyncStore = Microsoft.Protocols.TestSuites.Common.DataStructures.SyncStore;

    /// <summary>
    /// The base class of scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Variables

        /// <summary>
        /// Gets or sets a value indicating whether ActiveSync protocol version is 12.1 or not.
        /// </summary>
        public bool IsActiveSyncProtocolVersion121 { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether ActiveSync protocol version is 14.0 or not.
        /// </summary>
        public bool IsActiveSyncProtocolVersion140 { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether ActiveSync protocol version is 14.1 or not.
        /// </summary>
        public bool IsActiveSyncProtocolVersion141 { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether ActiveSync protocol version is 16.0 or not.
        /// </summary>
        public bool IsActiveSyncProtocolVersion160 { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether ActiveSync protocol version is 16.1 or not.
        /// </summary>
        public bool IsActiveSyncProtocolVersion161 { get; set; }

        /// <summary>
        /// Gets the protocol adapter
        /// </summary>
        protected IMS_ASCALAdapter CALAdapter { get; private set; }

        /// <summary>
        /// Gets the default ActiveSync Protocol Version
        /// </summary>
        protected string ActiveSyncProtocolVersion { get; private set; }

        /// <summary>
        /// Gets or sets the information of User1.
        /// </summary>
        protected UserInformation User1Information { get; set; }

        /// <summary>
        /// Gets or sets the information of User2.
        /// </summary>
        protected UserInformation User2Information { get; set; }

        /// <summary>
        /// Gets or sets the information of current user.
        /// </summary>
        protected UserInformation CurrentUserInformation { get; set; }

        /// <summary>
        /// Gets or sets the calendar subject element value.
        /// </summary>
        protected string SubjectName { get; set; }

        /// <summary>
        /// Gets or sets the location subject element value.
        /// </summary>
        protected string Location { get; set; }

        /// <summary>
        /// Gets or sets the content element value.
        /// </summary>
        protected string Content { get; set; }

        /// <summary>
        /// Gets or sets the category element value.
        /// </summary>
        protected string Category { get; set; }

        /// <summary>
        /// Gets or sets the startTime element value.
        /// </summary>
        protected DateTime StartTime { get; set; }

        /// <summary>
        /// Gets or sets the endTime element value.
        /// </summary>
        protected DateTime EndTime { get; set; }

        /// <summary>
        /// Gets or sets the pastTime element value.
        /// </summary>
        protected DateTime PastTime { get; set; }

        /// <summary>
        /// Gets or sets the futureTime element value.
        /// </summary>
        protected DateTime FutureTime { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// Create a Request.Recurrence instance with child elements.
        /// </summary>
        /// <param name="recurrenceType">The recurrence type</param>
        /// <param name="occurrences">The value of Occurrences element.</param>
        /// <param name="interval">The value of Interval element.</param>
        /// <returns>The created Request.Recurrence instance without Until element.</returns>
        public Request.Recurrence CreateCalendarRecurrence(byte recurrenceType, int occurrences, int interval)
        {
            Request.Recurrence recurrence = new Request.Recurrence { Type = recurrenceType };

            if (occurrences != 0)
            {
                recurrence.Occurrences = ushort.Parse(occurrences.ToString());
                recurrence.OccurrencesSpecified = true;
            }
            else
            {
                recurrence.OccurrencesSpecified = false;
            }

            recurrence.Interval = ushort.Parse(interval.ToString());

            // The value of WeekOfMonth MUST between 1 and 5. WeekOfMonth is required to be included when Type is either 3 or 6.
            if (recurrence.Type == 3 || recurrence.Type == 6)
            {
                recurrence.WeekOfMonth = byte.Parse("1");
                recurrence.WeekOfMonthSpecified = true;
            }

            // DayOfWeek is required to be included when Type is 1, 3 or 6.
            if (recurrence.Type == 1 || recurrence.Type == 3 || recurrence.Type == 6)
            {
                recurrence.DayOfWeek = ushort.Parse("1");
                recurrence.DayOfWeekSpecified = true;
            }

            // The value of MonthOfYear MUST between 1 and 12. MonthOfYear is required to be included when Type is either 5 or 6.
            if (recurrence.Type == 5 || recurrence.Type == 6)
            {
                recurrence.MonthOfYear = byte.Parse("4");
                recurrence.MonthOfYearSpecified = true;
            }

            // The value of DayOfMonth MUST between 1 and 31. DayOfMonth is required to be included when Type is either 2 or 5.
            if (recurrence.Type == 2 || recurrence.Type == 5)
            {
                recurrence.DayOfMonth = byte.Parse("1");
                recurrence.DayOfMonthSpecified = true;
            }

            // CalendarType is only included when Type is either 2, 3, 5, 6.
            // The CalendarType element is not supported when the MS-ASProtocolVersion header is set to 12.1
            if (!this.IsActiveSyncProtocolVersion121 && (recurrence.Type == 2 || recurrence.Type == 3 || recurrence.Type == 5 || recurrence.Type == 6))
            {
                recurrence.CalendarType = byte.Parse("1");
                recurrence.CalendarTypeSpecified = true;
            }

            return recurrence;
        }

        /// <summary>
        /// Create a Request.Recurrence instance including specified CalendarType element value.
        /// </summary>
        /// <param name="recurrenceInstance">The recurrence instance</param>
        /// <param name="calendarType">The calendarType value</param>
        /// <returns>The created Request.Recurrence instance.</returns>
        public Request.Recurrence CreateRecurrenceIncludingCalendarType(Request.Recurrence recurrenceInstance, byte calendarType)
        {
            Request.Recurrence recurrence = recurrenceInstance;

            if (!this.IsActiveSyncProtocolVersion121)
            {
                recurrence.CalendarType = calendarType;
                recurrence.CalendarTypeSpecified = true;
            }

            return recurrence;
        }

        /// <summary>
        /// Create a Calendar instance with 'Subject', 'OrganizerName', 'OrganizerEmail', 'Location'
        /// 'TimeZone', 'Body' and 'UID' elements.
        /// </summary>
        /// <returns>The collection of created Calendar items.</returns>
        public Dictionary<Request.ItemsChoiceType8, object> CreateDefaultCalendar()
        {
            Dictionary<Request.ItemsChoiceType8, object> calendarItem = new Dictionary<Request.ItemsChoiceType8, object>
            {
                {
                    Request.ItemsChoiceType8.Subject, this.SubjectName
                },
                {
                    Request.ItemsChoiceType8.Timezone,
                    "IP7//ygAVQBUAEMAKwAwADgAOgAwADAAKQAgAEIAZQBpAGoAaQBuAGcALAAgAEMAaABvAG4AZwBxAGkAbgBnACwAIAAAAAAAAAAAAAAAAAAAAAAAAAAAACgAVQBUAEMAKwAwADgAOgAwADAAKQAgAEIAZQBpAGoAaQBuAGcALAAgAEMAaABvAG4AZwBxAGkAbgBnACwAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=="
                },
                {
                    Request.ItemsChoiceType8.Body, TestSuiteHelper.CreateCalendarBody(1, this.Content)
                }
            };

            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) != "12.1"
                && Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) != "14.0"
                && Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) != "14.1")
            {
                calendarItem.Add(Request.ItemsChoiceType8.ClientUid, Guid.NewGuid().ToString());
            }
            else
            {
                calendarItem.Add(Request.ItemsChoiceType8.OrganizerEmail, Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain));
                calendarItem.Add(Request.ItemsChoiceType8.OrganizerName, this.User1Information.UserName);
                calendarItem.Add(Request.ItemsChoiceType8.Location1, this.Location);
                calendarItem.Add(Request.ItemsChoiceType8.UID, Guid.NewGuid().ToString());
            }

            return calendarItem;
        }

        /// <summary>
        /// Delete all the items in a folder.
        /// </summary>
        /// <param name="createdItemsCollection">The collection of items which should be deleted.</param>
        public void DeleteItemsInFolder(Collection<CreatedItems> createdItemsCollection)
        {
            foreach (CreatedItems itemsToFolder in createdItemsCollection)
            {
                SyncStore result = this.SyncChanges(itemsToFolder.CollectionId);

                if (result.AddElements != null)
                {
                    SyncRequest deleteRequest;
                    foreach (SyncItem item in result.AddElements)
                    {
                        if (itemsToFolder.CollectionId == this.CurrentUserInformation.CalendarCollectionId)
                        {
                            foreach (string subject in itemsToFolder.ItemSubject)
                            {
                                if (item.Calendar != null)
                                {
                                    if (item.Calendar.Subject.Equals(subject, StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        deleteRequest = TestSuiteHelper.CreateSyncDeleteRequest(itemsToFolder.CollectionId, result.SyncKey, item.ServerId);
                                        SyncStore deleteResult = this.CALAdapter.Sync(deleteRequest);
                                        Site.Assert.AreEqual<byte>(1, deleteResult.CollectionStatus, "Item should be deleted.");
                                    }
                                }
                            }
                        }

                        if (itemsToFolder.CollectionId == this.CurrentUserInformation.InboxCollectionId || itemsToFolder.CollectionId == this.CurrentUserInformation.DeletedItemsCollectionId || itemsToFolder.CollectionId == this.CurrentUserInformation.SentItemsCollectionId)
                        {
                            foreach (string subject in itemsToFolder.ItemSubject)
                            {
                                if (item.Email != null)
                                {
                                    if (item.Email.Subject.Equals(subject, StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        deleteRequest = TestSuiteHelper.CreateSyncDeleteRequest(itemsToFolder.CollectionId, result.SyncKey, item.ServerId);
                                        SyncStore deleteResult = this.CALAdapter.Sync(deleteRequest);
                                        Site.Assert.AreEqual<byte>(1, deleteResult.CollectionStatus, "Item should be deleted.");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Delete all items from the specified collection
        /// </summary>
        /// <param name="collectionId">The specified collection id</param>
        public void DeleteAllItems(string collectionId)
        {
            ItemOperationsRequest itemOperationsRequest = TestSuiteHelper.CreateItemOperationsEmptyFolderContentsRequest(collectionId);
            ItemOperationsStore emptyFolderContentsResponse = this.CALAdapter.ItemOperations(itemOperationsRequest);

            // Verify itemOperations response, if the command executes successfully, the Status in response should be 1.
            Site.Assert.AreEqual<string>(
                "1",
                emptyFolderContentsResponse.Status,
                "If the command executes successfully, the Status in response should be 1.");
        }

        /// <summary>
        /// Add a calendar to server, and if add is success then sync CalendarFolder.
        /// </summary>
        /// <param name="items">The dictionary store calendar item's element name and element value</param>
        /// <returns>Return the sync response</returns>
        public SyncStore AddSyncCalendar(Dictionary<Request.ItemsChoiceType8, object> items)
        {
            // Create a default calendar instance with Subject, TimeZone, Body, OrganizerEmail, OrganizerName, Location and UID elements
            Dictionary<Request.ItemsChoiceType8, object> calendar = this.CreateDefaultCalendar();

            // Add elements
            if (items != null)
            {
                foreach (KeyValuePair<Request.ItemsChoiceType8, object> item in items)
                {
                    if (calendar.ContainsKey(item.Key))
                    {
                        calendar[item.Key] = item.Value;
                    }
                    else
                    {
                        calendar.Add(item.Key, item.Value);
                    }
                }
            }

            Request.SyncCollectionAddApplicationData addCalendar = new Request.SyncCollectionAddApplicationData
            {
                Items = calendar.Values.ToArray<object>(),
                ItemsElementName = calendar.Keys.ToArray<Request.ItemsChoiceType8>()
            };

            // Sync to get the SyncKey
            SyncStore initializeSyncResponse = this.InitializeSync(this.CurrentUserInformation.CalendarCollectionId, null);

            // Add the calendar item
            SyncRequest syncRequest = TestSuiteHelper.CreateSyncAddRequest(this.CurrentUserInformation.CalendarCollectionId, initializeSyncResponse.SyncKey, addCalendar);
            SyncStore syncCalendarResponse = this.CALAdapter.Sync(syncRequest);

            // Verify sync response, if the Sync command executes successfully, the Status in response should be 1.
            Site.Assert.AreEqual<byte>(
                1,
                syncCalendarResponse.CollectionStatus,
                "If the Sync command executes successfully, the Status in response should be 1.");

            if (syncCalendarResponse.AddResponses != null)
            {
                foreach (Response.SyncCollectionsCollectionResponsesAdd response in syncCalendarResponse.AddResponses)
                {
                    if (response.Status.Equals(byte.Parse("1")))
                    {
                        // Sync command to do an initialization Sync, and get the server changes through sync command
                        syncCalendarResponse = this.SyncChanges(this.CurrentUserInformation.CalendarCollectionId);
                    }
                    else
                    {
                        return syncCalendarResponse;
                    }
                }
            }

            return syncCalendarResponse;
        }

        /// <summary>
        /// Sync changes between client and server
        /// </summary>
        /// <param name="collectionId">Specify the folder collection Id which needs to be synced.</param>
        /// <returns>Return the sync response</returns>
        public SyncStore SyncChanges(string collectionId)
        {
            SyncStore syncResponse;

            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 0;

            do
            {
                Thread.Sleep(waitTime);

                // Sync to get the SyncKey
                SyncStore initializeSyncResponse = this.InitializeSync(collectionId, null);

                // Get the server changes through sync command
                SyncRequest syncRequest = TestSuiteHelper.CreateSyncRequest(collectionId, initializeSyncResponse.SyncKey, true);
                syncResponse = this.CALAdapter.Sync(syncRequest);
                if (syncResponse != null)
                {
                    if (syncResponse.CollectionStatus == 1)
                    {
                        break;
                    }
                }

                counter++;
            }
            while (counter < retryCount);

            // Verify sync response
            Site.Assert.AreEqual<byte>(
                1,
                syncResponse.CollectionStatus,
                "If the Sync command executes successfully, the Status in response should be 1.");

            return syncResponse;
        }

        /// <summary>
        /// Initialize the sync with server
        /// </summary>
        /// <param name="collectionId">Specify the folder collection Id which needs to be synced.</param>
        /// <param name="supported">Specifies which contact and calendar elements in a Sync request are managed by the client and therefore not ghosted.</param>
        /// <returns>Return Sync response</returns>
        public SyncStore InitializeSync(string collectionId, Request.Supported supported)
        {
            // Obtains the key by sending an initial Sync request with a SyncKey element value of zero and the CollectionId element
            SyncRequest initializeSyncRequest = TestSuiteHelper.InitializeSyncRequest(collectionId, supported);
            SyncStore initializeSyncResponse = this.CALAdapter.Sync(initializeSyncRequest);

            // Verify sync result
            Site.Assert.AreEqual<byte>(
                1,
                initializeSyncResponse.CollectionStatus,
                "If the Sync command executes successfully, the Status in response should be 1.");

            return initializeSyncResponse;
        }

        /// <summary>
        /// Get the specified email item from the sync change response
        /// </summary>
        /// <param name="collectionId">The email folder server id.</param>
        /// <param name="calendarSubject">The subject value of calendar.</param>
        /// <returns>Return the specified email item.</returns>
        public SyncItem GetChangeItem(string collectionId, string calendarSubject)
        {
            SyncItem resultItem = null;
            SyncStore syncResponse;

            if (collectionId == this.CurrentUserInformation.CalendarCollectionId)
            {
                // Get the server changes through sync command
                syncResponse = this.SyncChanges(collectionId);

                if (syncResponse != null && syncResponse.CollectionStatus == 1)
                {
                    foreach (SyncItem item in syncResponse.AddElements)
                    {
                        if (item.Calendar.Subject == calendarSubject)
                        {
                            resultItem = item;
                            break;
                        }
                    }
                }
            }
            else if (collectionId == this.CurrentUserInformation.InboxCollectionId || collectionId == this.CurrentUserInformation.DeletedItemsCollectionId)
            {
                int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
                int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
                int counter = 0;

                // Get the mailbox changes
                do
                {
                    Thread.Sleep(waitTime);

                    // Get the server changes through sync command, get the new added email item
                    syncResponse = this.SyncChanges(collectionId);

                    if (syncResponse != null && syncResponse.CollectionStatus == 1)
                    {
                        foreach (SyncItem item in syncResponse.AddElements)
                        {
                            if (item.Email.Subject == calendarSubject)
                            {
                                resultItem = item;
                                break;
                            }
                        }
                    }

                    counter++;
                }
                while ((syncResponse == null || resultItem == null) && counter < retryCount);
            }

            return resultItem;
        }

        /// <summary>
        /// A method that send out string command request that contains multiple CalendarType elements in Recurrence.
        /// </summary>
        /// <param name="type">value of Type element in Recurrence.</param>
        /// <returns>Command response of sending </returns>
        public SyncStore AddCalendarWithMultipleCalendarType(string type)
        {
            SyncStore initialSyncResponse = this.InitializeSync(this.CurrentUserInformation.CalendarCollectionId, null);
            SendStringResponse sendStringResponse = new SendStringResponse();

            switch (type)
            {
                case "2":

                    // EAS_User02@contoso.com in string request is a sample email address of an attendee, nothing will be sent to this mailbox.
                    sendStringResponse = this.CALAdapter.SendStringRequest("<?xml version=\"1.0\" encoding=\"utf-8\"?><Sync xmlns=\"AirSync\"><Collections><Collection><SyncKey>" + initialSyncResponse.SyncKey + "</SyncKey><CollectionId>1</CollectionId><WindowSize>512</WindowSize><Options><BodyPreference xmlns=\"AirSyncBase\"><Type>2</Type></BodyPreference></Options><Commands><Add><ClientId>54353</ClientId><ApplicationData><Subject xmlns=\"Calendar\">TestMail</Subject><MeetingStatus xmlns=\"Calendar\">1</MeetingStatus><UID xmlns=\"Calendar\">040000008200E00074C5B7101A82E00800000000B0CD1F52EBBDC901000000000000000010000000B05E442FCB2CA443BF3D99B51A729FE6</UID><Timezone xmlns=\"Calendar\">IP7//ygAVQBUAEMAKwAwADgAOgAwADAAKQAgAEIAZQBpAGoAaQBuAGcALAAgAEMAaABvAG4AZwBxAGkAbgBnACwAIAAAAAAAAAAAAAAAAAAAAAAAAAAAACgAVQBUAEMAKwAwADgAOgAwADAAKQAgAEIAZQBpAGoAaQBuAGcALAAgAEMAaABvAG4AZwBxAGkAbgBnACwAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</Timezone><Attendees xmlns=\"Calendar\"><Attendee><Email>EAS_User02@contoso.com</Email><Name>EAS_User02</Name><AttendeeStatus>3</AttendeeStatus><AttendeeType>1</AttendeeType></Attendee></Attendees><Recurrence xmlns=\"Calendar\"><Type>" + type + "</Type><Occurrences>3</Occurrences><Interval>0</Interval><DayOfMonth>1</DayOfMonth><CalendarType>1</CalendarType><CalendarType>2</CalendarType></Recurrence><ResponseRequested xmlns=\"Calendar\">1</ResponseRequested><Reminder xmlns=\"Calendar\">10</Reminder></ApplicationData></Add></Commands></Collection></Collections></Sync>");
                    break;

                case "3":

                    // EAS_User02@contoso.com in string request is a sample email address of an attendee, nothing will be sent to this mailbox.
                    sendStringResponse = this.CALAdapter.SendStringRequest("<?xml version=\"1.0\" encoding=\"utf-8\"?><Sync xmlns=\"AirSync\"><Collections><Collection><SyncKey>" + initialSyncResponse.SyncKey + "</SyncKey><CollectionId>1</CollectionId><WindowSize>512</WindowSize><Options><BodyPreference xmlns=\"AirSyncBase\"><Type>2</Type></BodyPreference></Options><Commands><Add><ClientId>54353</ClientId><ApplicationData><Subject xmlns=\"Calendar\">TestMail</Subject><MeetingStatus xmlns=\"Calendar\">1</MeetingStatus><UID xmlns=\"Calendar\">040000008200E00074C5B7101A82E00800000000B0CD1F52EBBDC901000000000000000010000000B05E442FCB2CA443BF3D99B51A729FE6</UID><Timezone xmlns=\"Calendar\">IP7//ygAVQBUAEMAKwAwADgAOgAwADAAKQAgAEIAZQBpAGoAaQBuAGcALAAgAEMAaABvAG4AZwBxAGkAbgBnACwAIAAAAAAAAAAAAAAAAAAAAAAAAAAAACgAVQBUAEMAKwAwADgAOgAwADAAKQAgAEIAZQBpAGoAaQBuAGcALAAgAEMAaABvAG4AZwBxAGkAbgBnACwAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</Timezone><Attendees xmlns=\"Calendar\"><Attendee><Email>EAS_User02@contoso.com</Email><Name>EAS_User02</Name><AttendeeStatus>3</AttendeeStatus><AttendeeType>1</AttendeeType></Attendee></Attendees><Recurrence xmlns=\"Calendar\"><Type>" + type + "</Type><Occurrences>3</Occurrences><Interval>0</Interval><WeekOfMonth>1</WeekOfMonth><DayOfWeek>1</DayOfWeek><CalendarType>1</CalendarType><CalendarType>2</CalendarType></Recurrence><ResponseRequested xmlns=\"Calendar\">1</ResponseRequested><Reminder xmlns=\"Calendar\">10</Reminder></ApplicationData></Add></Commands></Collection></Collections></Sync>");
                    break;

                case "5":

                    // EAS_User02@contoso.com in string request is a sample email address of an attendee, nothing will be sent to this mailbox.
                    sendStringResponse = this.CALAdapter.SendStringRequest("<?xml version=\"1.0\" encoding=\"utf-8\"?><Sync xmlns=\"AirSync\"><Collections><Collection><SyncKey>" + initialSyncResponse.SyncKey + "</SyncKey><CollectionId>1</CollectionId><WindowSize>512</WindowSize><Options><BodyPreference xmlns=\"AirSyncBase\"><Type>2</Type></BodyPreference></Options><Commands><Add><ClientId>54353</ClientId><ApplicationData><Subject xmlns=\"Calendar\">TestMail</Subject><MeetingStatus xmlns=\"Calendar\">1</MeetingStatus><UID xmlns=\"Calendar\">040000008200E00074C5B7101A82E00800000000B0CD1F52EBBDC901000000000000000010000000B05E442FCB2CA443BF3D99B51A729FE6</UID><Timezone xmlns=\"Calendar\">IP7//ygAVQBUAEMAKwAwADgAOgAwADAAKQAgAEIAZQBpAGoAaQBuAGcALAAgAEMAaABvAG4AZwBxAGkAbgBnACwAIAAAAAAAAAAAAAAAAAAAAAAAAAAAACgAVQBUAEMAKwAwADgAOgAwADAAKQAgAEIAZQBpAGoAaQBuAGcALAAgAEMAaABvAG4AZwBxAGkAbgBnACwAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</Timezone><Attendees xmlns=\"Calendar\"><Attendee><Email>EAS_User02@contoso.com</Email><Name>EAS_User02</Name><AttendeeStatus>3</AttendeeStatus><AttendeeType>1</AttendeeType></Attendee></Attendees><Recurrence xmlns=\"Calendar\"><Type>" + type + "</Type><Occurrences>3</Occurrences><Interval>0</Interval><DayOfMonth>1</DayOfMonth><MonthOfYear>1</MonthOfYear><CalendarType>1</CalendarType><CalendarType>2</CalendarType></Recurrence><ResponseRequested xmlns=\"Calendar\">1</ResponseRequested><Reminder xmlns=\"Calendar\">10</Reminder></ApplicationData></Add></Commands></Collection></Collections></Sync>");
                    break;

                case "6":

                    // EAS_User02@contoso.com in string request is a sample email address of an attendee, nothing will be sent to this mailbox.
                    sendStringResponse = this.CALAdapter.SendStringRequest("<?xml version=\"1.0\" encoding=\"utf-8\"?><Sync xmlns=\"AirSync\"><Collections><Collection><SyncKey>" + initialSyncResponse.SyncKey + "</SyncKey><CollectionId>1</CollectionId><WindowSize>512</WindowSize><Options><BodyPreference xmlns=\"AirSyncBase\"><Type>2</Type></BodyPreference></Options><Commands><Add><ClientId>54353</ClientId><ApplicationData><Subject xmlns=\"Calendar\">TestMail</Subject><MeetingStatus xmlns=\"Calendar\">1</MeetingStatus><UID xmlns=\"Calendar\">040000008200E00074C5B7101A82E00800000000B0CD1F52EBBDC901000000000000000010000000B05E442FCB2CA443BF3D99B51A729FE6</UID><Timezone xmlns=\"Calendar\">IP7//ygAVQBUAEMAKwAwADgAOgAwADAAKQAgAEIAZQBpAGoAaQBuAGcALAAgAEMAaABvAG4AZwBxAGkAbgBnACwAIAAAAAAAAAAAAAAAAAAAAAAAAAAAACgAVQBUAEMAKwAwADgAOgAwADAAKQAgAEIAZQBpAGoAaQBuAGcALAAgAEMAaABvAG4AZwBxAGkAbgBnACwAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</Timezone><Attendees xmlns=\"Calendar\"><Attendee><Email>EAS_User02@contoso.com</Email><Name>EAS_User02</Name><AttendeeStatus>3</AttendeeStatus><AttendeeType>1</AttendeeType></Attendee></Attendees><Recurrence xmlns=\"Calendar\"><Type>" + type + "</Type><Occurrences>3</Occurrences><Interval>0</Interval><WeekOfMonth>1</WeekOfMonth><MonthOfYear>1</MonthOfYear><CalendarType>1</CalendarType><CalendarType>2</CalendarType></Recurrence><ResponseRequested xmlns=\"Calendar\">1</ResponseRequested><Reminder xmlns=\"Calendar\">10</Reminder></ApplicationData></Add></Commands></Collection></Collections></Sync>");
                    break;
            }

            SyncStore response = new SyncStore();
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(sendStringResponse.ResponseDataXML);
            XmlNodeList nodes = doc.DocumentElement.GetElementsByTagName("Collection");

            foreach (XmlNode node in nodes)
            {
                foreach (XmlNode item in node.ChildNodes)
                {
                    if (item.Name == "SyncKey")
                    {
                        response.SyncKey = item.InnerText;
                    }

                    if (item.Name == "CollectionId")
                    {
                        response.CollectionId = item.InnerText;
                    }

                    if (item.Name == "Status")
                    {
                        response.CollectionStatus = byte.Parse(item.InnerText);
                    }

                    if (item.Name == "Responses")
                    {
                        foreach (XmlNode add in item)
                        {
                            if (add.Name == "Add")
                            {
                                Response.SyncCollectionsCollectionResponsesAdd res = new Response.SyncCollectionsCollectionResponsesAdd();

                                foreach (XmlNode additem in add)
                                {
                                    if (additem.Name == "ClientId")
                                    {
                                        res.ClientId = additem.InnerText;
                                    }

                                    if (additem.Name == "ServerId")
                                    {
                                        res.ServerId = additem.InnerText;
                                    }

                                    if (additem.Name == "Status")
                                    {
                                        res.Status = additem.InnerText;
                                    }
                                }

                                response.AddResponses.Add(res);
                            }
                        }
                    }
                }
            }

            return response;
        }

        /// <summary>
        /// Call sync command to update properties of an existing calendar item.
        /// </summary>
        /// <param name="serverId">Server Id of the calendar item.</param>
        /// <param name="collectionId">Collection Id of the folder that calendar item is contained in.</param>
        /// <param name="syncKey">Sync key value.</param>
        /// <param name="items">The dictionary store calendar item's element name and element value, which will be changed.</param>
        /// <returns>Return Sync Change response.</returns>
        public SyncStore UpdateCalendarProperty(string serverId, string collectionId, string syncKey, Dictionary<Request.ItemsChoiceType7, object> items)
        {
            Request.SyncCollectionChangeApplicationData syncChangeData = new Request.SyncCollectionChangeApplicationData
            {
                ItemsElementName = items.Keys.ToArray<Request.ItemsChoiceType7>(),
                Items = items.Values.ToArray<object>()
            };

            Request.SyncCollectionChange syncChange = new Request.SyncCollectionChange
            {
                ApplicationData = syncChangeData,
                ServerId = serverId
            };

            SyncRequest syncChangeRequest = new SyncRequest
            {
                RequestData = new Request.Sync { Collections = new Request.SyncCollection[1] }
            };

            syncChangeRequest.RequestData.Collections[0] = new Request.SyncCollection
            {
                Commands = new object[] { syncChange },
                SyncKey = syncKey,
                CollectionId = collectionId
            };

            SyncStore syncChanageResponse = this.CALAdapter.Sync(syncChangeRequest);
            return syncChanageResponse;
        }

        /// <summary>
        /// Record the user name, folder collectionId and subjects the current test case impacts.
        /// </summary>
        /// <param name="userName">The user that current test case used.</param>
        /// <param name="folderCollectionId">The collectionId of folders that the current test case impacts.</param>
        /// <param name="itemSubjects">The subject of items that the current test case impacts.</param>
        protected void RecordCaseRelativeItems(string userName, string folderCollectionId, params string[] itemSubjects)
        {
            // Record the item in the specified folder.
            CreatedItems createdItems = new CreatedItems { CollectionId = folderCollectionId };
            foreach (string subject in itemSubjects)
            {
                createdItems.ItemSubject.Add(subject);
            }

            // Record the created items of User1.
            if (userName == this.User1Information.UserName)
            {
                this.User1Information.UserCreatedItems.Add(createdItems);
            }

            // Record the created items of User2.
            if (userName == this.User2Information.UserName)
            {
                this.User2Information.UserCreatedItems.Add(createdItems);
            }
        }

        /// <summary>
        /// This method is used to change user to call ActiveSync commands and resynchronize the folder collection hierarchy.
        /// </summary>
        /// <param name="userInformation">The information of the user that will switch to.</param>
        /// <param name="isFolderSyncNeeded">A boolean value indicates whether needs to synchronize the folder hierarchy.</param>
        protected void SwitchUser(UserInformation userInformation, bool isFolderSyncNeeded)
        {
            this.CALAdapter.SwitchUser(userInformation.UserName, userInformation.UserPassword, userInformation.UserDomain);

            if (isFolderSyncNeeded)
            {
                FolderSyncResponse folderSyncResponse = this.CALAdapter.FolderSync();

                // Get the folder collectionId of User1
                if (userInformation.UserName == this.User1Information.UserName)
                {
                    if (string.IsNullOrEmpty(this.User1Information.InboxCollectionId))
                    {
                        this.User1Information.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, this.Site);
                    }

                    if (string.IsNullOrEmpty(this.User1Information.CalendarCollectionId))
                    {
                        this.User1Information.CalendarCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Calendar, this.Site);
                    }

                    if (string.IsNullOrEmpty(this.User1Information.SentItemsCollectionId))
                    {
                        this.User1Information.SentItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.SentItems, this.Site);
                    }

                    if (string.IsNullOrEmpty(this.User1Information.DeletedItemsCollectionId))
                    {
                        this.User1Information.DeletedItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.DeletedItems, this.Site);
                    }
                }

                // Get the folder collectionId of User2
                if (userInformation.UserName == this.User2Information.UserName)
                {
                    if (string.IsNullOrEmpty(this.User2Information.InboxCollectionId))
                    {
                        this.User2Information.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, this.Site);
                    }

                    if (string.IsNullOrEmpty(this.User2Information.CalendarCollectionId))
                    {
                        this.User2Information.CalendarCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Calendar, this.Site);
                    }

                    if (string.IsNullOrEmpty(this.User2Information.SentItemsCollectionId))
                    {
                        this.User2Information.SentItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.SentItems, this.Site);
                    }

                    if (string.IsNullOrEmpty(this.User2Information.DeletedItemsCollectionId))
                    {
                        this.User2Information.DeletedItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.DeletedItems, this.Site);
                    }
                }

                this.CurrentUserInformation = userInformation;
            }
        }

        /// <summary>
        /// Using mail with mime content to send a meeting request or cancel request
        /// </summary>
        /// <param name="calendarItem">Calendar information</param>
        /// <param name="subjectName">The subject name of meeting request mail</param>
        /// <param name="organizerEmailAddress">The organizer email address</param>
        /// <param name="attendeeEmailAddress">The attendee email address</param>
        /// <param name="method">Specify normal appointments from meeting requests, responses, and cancellations, it can be set to 'REQUEST', 'REPLY', or 'CANCEL'</param>
        /// <param name="replyMethod">Specify REPLY method, it can be set to 'ACCEPTED', 'TENTATIVE', or 'DECLINED'</param>
        protected void SendMimeMeeting(Calendar calendarItem, string subjectName, string organizerEmailAddress, string attendeeEmailAddress, string method, string replyMethod)
        {
            string icalendarContent = string.Empty;
            switch (method.ToUpper(CultureInfo.CurrentCulture))
            {
                case "REQUEST":
                case "CANCEL":
                    icalendarContent = TestSuiteHelper.CreateiCalendarFormatContent(calendarItem, method, replyMethod, organizerEmailAddress, attendeeEmailAddress);
                    break;

                case "REPLY":
                    icalendarContent = TestSuiteHelper.CreateiCalendarFormatContent(calendarItem, method, replyMethod, attendeeEmailAddress, organizerEmailAddress);
                    break;
            }

            string body = string.Empty;

            string mime = TestSuiteHelper.CreateMeetingRequestMime(
                organizerEmailAddress,
                attendeeEmailAddress,
                subjectName,
                body,
                icalendarContent);

            // Send a meeting request
            SendMailRequest sendMailRequest = TestSuiteHelper.CreateSendMailRequest(TestSuiteHelper.Next(), false, mime);
            SendMailResponse sendMailResponse = this.CALAdapter.SendMail(sendMailRequest);

            Site.Assert.AreEqual<string>(
                string.Empty,
                sendMailResponse.ResponseDataXML,
                 "The server should return an empty xml response data to indicate SendMail command success.");
        }

        /// <summary>
        /// Call MeetingResponse command to respond the meeting request
        /// </summary>
        /// <param name="userResponse">The value indicates whether the meeting is being accepted, tentatively accepted, or declined</param>
        /// <param name="collectionId">Specify the server id of mailbox</param>
        /// <param name="serverId">Specify a unique identifier that was assigned by the server for a mailItem</param>
        /// <param name="instanceId">Specify the start time of the appointment or meeting instance to be modified. The format of the instanceId value
        /// is a string in dateTime ([MS-ASDTYPE] section 2.3) format with the punctuation separators, for example, 2010-04-08T18:16:00.000Z</param>
        /// <returns>a bool value</returns>
        protected bool MeetingResponse(byte userResponse, string collectionId, string serverId, string instanceId)
        {
            bool isSuccess = false;

            // Create a MeetingResponse request item
            Request.MeetingResponseRequest meetingResponseRequestItem = new Request.MeetingResponseRequest
            {
                UserResponse = userResponse,
                CollectionId = collectionId,
                RequestId = serverId
            };

            if (!string.IsNullOrEmpty(instanceId))
            {
                meetingResponseRequestItem.InstanceId = instanceId;
            }

            // Create a MeetingResponse request
            MeetingResponseRequest meetingResponseRequest = Common.CreateMeetingResponseRequest(new Request.MeetingResponseRequest[] { meetingResponseRequestItem });
            MeetingResponseResponse meetingResponseResponse = this.CALAdapter.MeetingResponse(meetingResponseRequest);

            if (meetingResponseResponse.ResponseData.Result[0].Status == "1")
            {
                isSuccess = true;
            }

            return isSuccess;
        }

        #endregion

        #region Test case initialize and cleanup

        /// <summary>
        /// Initialize the test Case
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.InitializePropertiesValue();
        }

        /// <summary>
        /// Clean up the test Case
        /// </summary>
        protected override void TestCleanup()
        {
            base.TestCleanup();
            if (this.User1Information.UserCreatedItems.Count != 0)
            {
                // Switch to User1
                this.SwitchUser(this.User1Information, false);
                this.DeleteItemsInFolder(this.User1Information.UserCreatedItems);
            }

            if (this.User2Information.UserCreatedItems.Count != 0)
            {
                // Switch to User2
                this.SwitchUser(this.User2Information, false);
                this.DeleteItemsInFolder(this.User2Information.UserCreatedItems);
            }
        }

        #endregion

        #region Private method

        /// <summary>
        /// Initialize properties value.
        /// </summary>
        private void InitializePropertiesValue()
        {
            this.IsActiveSyncProtocolVersion121 = false;
            this.IsActiveSyncProtocolVersion140 = false;
            this.IsActiveSyncProtocolVersion141 = false;
            this.IsActiveSyncProtocolVersion160 = false;
            this.IsActiveSyncProtocolVersion161 = false;

            if (this.CALAdapter == null)
            {
                this.CALAdapter = TestClassBase.BaseTestSite.GetAdapter<IMS_ASCALAdapter>();
            }

            // Get the information of User1.
            this.User1Information = new UserInformation()
            {
                UserName = Common.GetConfigurationPropertyValue("OrganizerUserName", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("OrganizerUserPassword", this.Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", this.Site)
            };

            // Get the information of User2.
            this.User2Information = new UserInformation()
            {
                UserName = Common.GetConfigurationPropertyValue("AttendeeUserName", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("AttendeeUserPassword", this.Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", this.Site)
            };
            this.CurrentUserInformation = this.User1Information;
            this.ActiveSyncProtocolVersion = Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site);

            if (this.ActiveSyncProtocolVersion.Equals("12.1"))
            {
                this.IsActiveSyncProtocolVersion121 = true;
            }
            else if (this.ActiveSyncProtocolVersion.Equals("14.0"))
            {
                this.IsActiveSyncProtocolVersion140 = true;
            }
            else if (this.ActiveSyncProtocolVersion.Equals("14.1"))
            {
                this.IsActiveSyncProtocolVersion141 = true;
            }
            else if (this.ActiveSyncProtocolVersion.Equals("16.0"))
            {
                this.IsActiveSyncProtocolVersion160 = true;
            }
            else if (this.ActiveSyncProtocolVersion.Equals("16.1"))
            {
                this.IsActiveSyncProtocolVersion161 = true;
            }
            
            if (Common.GetSutVersion(this.Site) != SutVersion.ExchangeServer2007 || this.IsActiveSyncProtocolVersion121)
            {
                this.SwitchUser(this.User1Information, true);
            }

            this.SubjectName = Common.GenerateResourceName(this.Site, "subject");
            this.Location = Common.GenerateResourceName(this.Site, "location");
            this.Content = Common.GenerateResourceName(this.Site, "content");
            this.Category = Common.GenerateResourceName(this.Site, "category");
            
            // Set StartTime as tomorrow, EndTime as 1 hour after StartTime, PastTime as 4 days ago, FutureTime as 5 days from now.
            this.StartTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 0, 0).AddDays(1);
            this.EndTime = this.StartTime.AddHours(1);
            this.PastTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 9, 0, 0).AddDays(-4);
            this.FutureTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 9, 0, 0).AddDays(5);
        }

        #endregion
    }
}