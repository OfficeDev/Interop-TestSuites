namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Text;
    using System.Threading;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;
    using SyncItem = Microsoft.Protocols.TestSuites.Common.DataStructures.Sync;
    using SyncStore = Microsoft.Protocols.TestSuites.Common.DataStructures.SyncStore;

    /// <summary>
    /// A bass class for scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Variables
        /// <summary>
        /// Set the value as true to record that user has changed DeviceID
        /// </summary>
        private bool changeDeviceIDSpecified;

        /// <summary>
        /// Set the value as true to record that user has changed PolicyKey
        /// </summary>
        private bool changePolicyKeySpecified;
        #endregion

        #region Properties
        /// <summary>
        /// Gets the value of the client Id.
        /// </summary>
        protected static string ClientId
        {
            get
            {
                return Guid.NewGuid().ToString();
            }
        }

        /// <summary>
        /// Gets the MS-ASCMD protocol adapter.
        /// </summary>
        protected IMS_ASCMDAdapter CMDAdapter { get; private set; }

        /// <summary>
        /// Gets MS-ASCMD SUT Control adapter
        /// </summary>
        protected IMS_ASCMDSUTControlAdapter CMDSUTControlAdapter { get; private set; }

        /// <summary>
        /// Gets the value of the last SyncKey returned by Sync operation.
        /// </summary>
        protected string LastSyncKey { get; private set; }

        /// <summary>
        /// Gets the related information of User1.
        /// </summary>
        protected UserInformation User1Information { get; private set; }

        /// <summary>
        /// Gets the related information of User2.
        /// </summary>
        protected UserInformation User2Information { get; private set; }

        /// <summary>
        /// Gets the related information of User3.
        /// </summary>
        protected UserInformation User3Information { get; private set; }

        /// <summary>
        /// Gets the related information of User7.
        /// </summary>
        protected UserInformation User7Information { get; private set; }

        /// <summary>
        /// Gets the related information of User8.
        /// </summary>
        protected UserInformation User8Information { get; private set; }

        /// <summary>
        /// Gets the related information of User9.
        /// </summary>
        protected UserInformation User9Information { get; private set; }

        /// <summary>
        /// Gets the value of the last SyncKey returned by FolderSync.
        /// </summary>
        protected string LastFolderSyncKey { get; private set; }

        /// <summary>
        /// Gets or sets a value indicating whether the oof settings are changed.
        /// </summary>
        protected bool IsOofSettingsChanged { get; set; }

        #endregion

        #region Protected Static Methods

        /// <summary>
        /// Record the user name, folder collectionId and subjects the current test case impacts.
        /// </summary>
        /// <param name="userInformation">The user that current test case used.</param>
        /// <param name="folderCollectionId">The collectionId of folders that the current test case impact.</param>
        /// <param name="itemSubjects">The subject of items that the current test case impact.</param>
        protected static void RecordCaseRelativeItems(UserInformation userInformation, string folderCollectionId, params string[] itemSubjects)
        {
            CreatedItems createdItems = new CreatedItems { CollectionId = folderCollectionId };
            foreach (string itemSubject in itemSubjects)
            {
                createdItems.ItemSubject.Add(itemSubject);
            }

            userInformation.UserCreatedItems.Add(createdItems);
        }

        /// <summary>
        /// Remove items with subject that the current test case doesn't need.
        /// </summary>
        /// <param name="userInformation">The user that current test case used.</param>
        /// <param name="folderCollectionId">The collectionId of folders that the current test case impact.</param>
        /// <param name="itemSubjects">The subject of items that the current test case impact.</param>
        /// <returns>Return true when item removed, return false with wrong parameter.</returns>
        protected static bool RemoveRecordCaseRelativeItems(UserInformation userInformation, string folderCollectionId, params string[] itemSubjects)
        {
            bool removeSuccess = false;
            foreach (CreatedItems userItem in userInformation.UserCreatedItems)
            {
                if (userItem.CollectionId.Equals(folderCollectionId))
                {
                    foreach (string subject in itemSubjects)
                    {
                        removeSuccess = userItem.ItemSubject.Remove(subject);
                    }
                }
            }

            return removeSuccess;
        }

        /// <summary>
        /// Record the user name, folder collection ID that need to be deleted.
        /// </summary>
        /// <param name="userInformation">The user that current test case used.</param>
        /// <param name="folderCollectionID">The folder collection ID that need to be deleted.</param>
        protected static void RecordCaseRelativeFolders(UserInformation userInformation, params string[] folderCollectionID)
        {
            foreach (string folderID in folderCollectionID)
            {
                userInformation.UserCreatedFolders.Add(folderID);
            }
        }

        /// <summary>
        /// Create iCalendar format string from one calendar instance.
        /// </summary>
        /// <param name="calendar">The instance of Calendar class.</param>
        /// <returns>iCalendar formatted string.</returns>
        protected static string CreateiCalendarFormatContent(Calendar calendar)
        {
            StringBuilder ical = new StringBuilder();
            ical.AppendLine("BEGIN: VCALENDAR");
            ical.AppendLine("PRODID:-//Microosft Protocols TestSuites");
            ical.AppendLine("VERSION:2.0");
            ical.AppendLine("METHOD:REQUEST");

            ical.AppendLine("X-MS-OLK-FORCEINSPECTOROPEN:TRUE");
            ical.AppendLine("BEGIN:VTIMEZONE");
            ical.AppendLine("TZID:Universal Time");
            ical.AppendLine("BEGIN:STANDARD");
            ical.AppendLine("DTSTART:16011104T020000");
            ical.AppendLine("RRULE:FREQ=YEARLY;BYDAY=1SU;BYMONTH=11");
            ical.AppendLine("TZOFFSETFROM:-0000");
            ical.AppendLine("TZOFFSETTO:+0000");
            ical.AppendLine("END:STANDARD");
            ical.AppendLine("BEGIN:DAYLIGHT");
            ical.AppendLine("DTSTART:16010311T020000");
            ical.AppendLine("RRULE:FREQ=YEARLY;BYDAY=2SU;BYMONTH=3");
            ical.AppendLine("TZOFFSETFROM:-0000");
            ical.AppendLine("TZOFFSETTO:+0000");
            ical.AppendLine("END:DAYLIGHT");
            ical.AppendLine("END:VTIMEZONE");

            ical.AppendLine("BEGIN:VEVENT");
            ical.AppendLine("UID:" + calendar.UID);
            if (calendar.DtStamp != null)
            {
                ical.AppendLine("DTSTAMP:" + ((DateTime)calendar.DtStamp).ToString("yyyyMMddTHHmmss"));
            }
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
                ical.AppendLine("DTSTART;TZID=\"Universal Time\":" + ((DateTime)calendar.StartTime).ToString("yyyyMMddTHHmmss"));
                ical.AppendLine("DTEND;TZID=\"Universal Time\":" + ((DateTime)calendar.EndTime).ToString("yyyyMMddTHHmmss"));
            }

            if (calendar.Recurrence != null)
            {
                switch (calendar.Recurrence.Type)
                {
                    case 1:
                        ical.AppendLine("RRULE:FREQ=WEEKLY;BYDAY=MO;COUNT=3");
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
                ical.AppendLine("EXDATE;TZID=\"Universal Time\":" + ((DateTime)calendar.StartTime).AddDays(7).ToString("yyyyMMddTHHmmss"));
                ical.AppendLine("RECURRENCE-ID;TZID=\"Universal Time\":" + ((DateTime)calendar.StartTime).ToString("yyyyMMddTHHmmss"));
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
        /// Get the application data of an item specified by field.
        /// </summary>
        /// <param name="syncResponse">The Sync command response.</param>
        /// <param name="field">The element name of the item.</param>
        /// <param name="fieldValue">The value of the item</param>
        /// <returns>If the item exists, return its application data; otherwise, return null.</returns>
        protected static Response.SyncCollectionsCollectionCommandsAddApplicationData GetAddApplicationData(SyncResponse syncResponse, Response.ItemsChoiceType8 field, string fieldValue)
        {
            Response.SyncCollectionsCollectionCommandsAddApplicationData addData = null;

            Response.SyncCollectionsCollectionCommands commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            if (commands != null)
            {
                foreach (Response.SyncCollectionsCollectionCommandsAdd item in commands.Add)
                {
                    for (int i = 0; i < item.ApplicationData.ItemsElementName.Length; i++)
                    {
                        if (item.ApplicationData.ItemsElementName[i] == field &&
                            item.ApplicationData.Items[i].ToString() == fieldValue)
                        {
                            addData = item.ApplicationData;
                            break;
                        }
                    }
                }
            }

            return addData;
        }

        /// <summary>
        /// Create MIME for SendMail command.
        /// </summary>
        /// <param name="from">The email address of sender.</param>
        /// <param name="to">The email address of recipient.</param>
        /// <param name="subject">The email subject.</param>
        /// <param name="body">The email body content.</param>
        /// <returns>A MIME for SendMail command.</returns>
        protected static string CreateMIME(string from, string to, string subject, string body)
        {
            string mime =
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
Content-Description: number1.jpg
Content-Disposition: attachment; size=4; filename=""number1.jpg""
Content-Location: <cid:ebdc14bd-deb4-4816-b00b-6e2a46097d17>
Content-Transfer-Encoding: base64

MQ==

--_boundary_--
";

            return Common.FormatString(mime, from, to, subject, body);
        }

        /// <summary>
        /// Get an object related to a specified ItemsChoiceType10 value.
        /// </summary>
        /// <param name="syncResponse">A Sync response.</param>
        /// <param name="element">An element of ItemsChoiceType10 type, which specifies which object in the Sync response to be retrieved.</param>
        /// <returns>The object to be retrieved, if it exists in the Sync response; otherwise, return null.</returns>
        protected static object GetCollectionItem(SyncResponse syncResponse, Response.ItemsChoiceType10 element)
        {
            if (syncResponse.ResponseData.Item != null)
            {
                Response.SyncCollectionsCollection syncCollection = ((Response.SyncCollections)syncResponse.ResponseData.Item).Collection[0];

                for (int i = 0; i < syncCollection.ItemsElementName.Length; i++)
                {
                    if (syncCollection.ItemsElementName[i] == element)
                    {
                        return syncCollection.Items[i];
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Use this method to get a ServerId.
        /// </summary>
        /// <param name="folderSyncResponse">An instance of the FolderSyncResponse.</param>
        /// <param name="folderName">Folder name.</param>
        /// <returns>The collection ID of the specified folder.</returns>
        protected static string GetCollectionId(FolderSyncResponse folderSyncResponse, string folderName)
        {
            string serverId = string.Empty;
            foreach (Response.FolderSyncChangesAdd changes in folderSyncResponse.ResponseData.Changes.Add)
            {
                if (changes.DisplayName == folderName)
                {
                    serverId = changes.ServerId;
                    break;
                }
            }

            return serverId;
        }

        /// <summary>
        /// This method is used to create MoveItems request.
        /// </summary>
        /// <param name="srcMsgId">The server ID of the item to be moved.</param>
        /// <param name="srcFldId">The server ID of the source folder.</param>
        /// <param name="dstFldId">The server ID of the destination folder.</param>
        /// <returns>The MoveItems request.</returns>
        protected static MoveItemsRequest CreateMoveItemsRequest(string srcMsgId, string srcFldId, string dstFldId)
        {
            Request.MoveItemsMove moveItemsMove = new Request.MoveItemsMove
            {
                DstFldId = dstFldId,
                SrcFldId = srcFldId,
                SrcMsgId = srcMsgId
            };

            return Common.CreateMoveItemsRequest(new Request.MoveItemsMove[] { moveItemsMove });
        }

        /// <summary>
        /// Create a Sync delete operation request which would be used to move deleted items to the Deleted Items folder.
        /// </summary>
        /// <param name="syncKey">The synchronization state of a collection.</param>
        /// <param name="collectionId">The server ID of the folder.</param>
        /// <param name="serverId">An server ID of the item which will be deleted.</param>
        /// <returns>The Sync delete operation request.</returns>
        protected static SyncRequest CreateSyncDeleteRequest(string syncKey, string collectionId, string serverId)
        {
            Request.SyncCollection collection = new Request.SyncCollection
            {
                SyncKey = syncKey,
                GetChanges = true,
                CollectionId = collectionId,
                Commands = new object[] { new Request.SyncCollectionDelete { ServerId = serverId } }
            };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { collection });
        }

        /// <summary>
        /// Create a Sync delete operation request which would be used to delete items permanently.
        /// </summary>
        /// <param name="syncKey">The synchronization state of a collection.</param>
        /// <param name="collectionId">The server ID of the folder.</param>
        /// <param name="serverId">An server ID of the item which will be deleted.</param>
        /// <returns>The Sync delete operation request.</returns>
        protected static SyncRequest CreateSyncPermanentDeleteRequest(string syncKey, string collectionId, string serverId)
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
        /// This method is used to create GetItemEstimate request.
        /// </summary>
        /// <param name="syncKey">The synchronization state of a collection.</param>
        /// <param name="collectionId">The collection id of the folder.</param>
        /// <param name="options">Contains elements that filter the results.</param>
        /// <returns>The GetItemEstimate request.</returns>
        protected static GetItemEstimateRequest CreateGetItemEstimateRequest(string syncKey, string collectionId, Request.Options[] options, bool? ConversationMode = null)
        {
            List<Request.ItemsChoiceType10> itemsElementNames = new List<Request.ItemsChoiceType10>()
            {
                Request.ItemsChoiceType10.SyncKey,
                Request.ItemsChoiceType10.CollectionId
            };
            List<object> items = new List<object>()
            {
                syncKey,
                collectionId
            };

            if (ConversationMode != null)
            {
                itemsElementNames.Add(Request.ItemsChoiceType10.ConversationMode);
                items.Add(ConversationMode.Value);
            }
         
            Request.GetItemEstimateCollection collection = new Request.GetItemEstimateCollection
            {
                ItemsElementName = itemsElementNames.ToArray(),
                Items = items.ToArray()
            };

            if (options != null)
            {
                //itemsElementNames.Add(Request.ItemsChoiceType10.Options);
                //items.Add(options);
                collection.Options = options;
            }
            return Common.CreateGetItemEstimateRequest(new Request.GetItemEstimateCollection[] { collection });
        }

        /// <summary>
        /// This method is used to search serverId from SyncResponse.
        /// </summary>
        /// <param name="responseSync">An instance of the SyncResponse.</param>
        /// <param name="field">The name of the field.</param>
        /// <param name="fieldValue">The value of item.</param>
        /// <returns>The server ID of the specified item.</returns>
        protected static string FindServerId(SyncResponse responseSync, string field, string fieldValue)
        {
            Response.SyncCollections syncCollections = (Response.SyncCollections)responseSync.ResponseData.Item;
            if (syncCollections == null)
            {
                return null;
            }

            Response.SyncCollectionsCollectionCommands commands = null;
            for (int index = 0; index < syncCollections.Collection[0].ItemsElementName.Length; index++)
            {
                if (syncCollections.Collection[0].ItemsElementName[index] == Response.ItemsChoiceType10.Commands)
                {
                    commands = (Response.SyncCollectionsCollectionCommands)syncCollections.Collection[0].Items[index];
                    break;
                }
            }

            if (commands == null || commands.Add == null)
            {
                return null;
            }

            foreach (Response.SyncCollectionsCollectionCommandsAdd add in commands.Add)
            {
                for (int itemIndex = 0; itemIndex < add.ApplicationData.ItemsElementName.Length; itemIndex++)
                {
                    if (add.ApplicationData.ItemsElementName[itemIndex].ToString().Contains(field) && add.ApplicationData.Items[itemIndex].ToString().ToLower().Replace(" ", "") == fieldValue.ToLower().Replace(" ", ""))
                    {
                        return add.ServerId;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// This method is used to search serverId from SyncResponse.
        /// </summary>
        /// <param name="responseSync">An instance of the SyncResponse.</param>
        /// <param name="field">The value of the item name.</param>
        /// <param name="fieldValue">The value of item.</param>
        /// <returns>The collection of serverIds</returns>
        protected static Collection<string> FindServerIdList(SyncResponse responseSync, string field, string fieldValue)
        {
            Response.SyncCollections syncCollections = (Response.SyncCollections)responseSync.ResponseData.Item;
            if (syncCollections == null)
            {
                return null;
            }

            Response.SyncCollectionsCollectionCommands commands = null;
            for (int index = 0; index < syncCollections.Collection[0].ItemsElementName.Length; index++)
            {
                if (syncCollections.Collection[0].ItemsElementName[index] == Response.ItemsChoiceType10.Commands)
                {
                    commands = (Response.SyncCollectionsCollectionCommands)syncCollections.Collection[0].Items[index];
                    break;
                }
            }

            Collection<string> serverId = new Collection<string>();
            foreach (Response.SyncCollectionsCollectionCommandsAdd add in commands.Add)
            {
                for (int itemIndex = 0; itemIndex < add.ApplicationData.ItemsElementName.Length; itemIndex++)
                {
                    if (add.ApplicationData.ItemsElementName[itemIndex].ToString().Contains(field) && add.ApplicationData.Items[itemIndex].ToString().Contains(fieldValue))
                    {
                        serverId.Add(add.ServerId);
                        break;
                    }
                }
            }

            return serverId;
        }

        /// <summary>
        /// Get data from sync response BODY element
        /// </summary>
        /// <param name="syncResponse">An instance of the SyncResponse</param>
        /// <param name="serverId">The value of the ServerId element.</param>
        /// <returns>The value of BODY element.</returns>
        protected static string GetDataFromResponseBodyElement(SyncResponse syncResponse, string serverId)
        {
            Response.Body body = (Response.Body)TestSuiteBase.GetElementValueFromSyncResponse(syncResponse, serverId, Response.ItemsChoiceType8.Body);
            return body.Data;
        }

        /// <summary>
        /// Get element value from Sync response
        /// </summary>
        /// <param name="syncResponse">The Sync response</param>
        /// <param name="serverId">The specified serverId</param>
        /// <param name="elementType">The element type</param>
        /// <returns>The element value</returns>
        protected static object GetElementValueFromSyncResponse(SyncResponse syncResponse, string serverId, Response.ItemsChoiceType8 elementType)
        {
            Response.SyncCollections syncCollections = (Response.SyncCollections)syncResponse.ResponseData.Item;
            Response.SyncCollectionsCollectionCommands commands = null;
            for (int index = 0; index < syncCollections.Collection[0].ItemsElementName.Length; index++)
            {
                if (syncCollections.Collection[0].ItemsElementName[index] == Response.ItemsChoiceType10.Commands)
                {
                    commands = (Response.SyncCollectionsCollectionCommands)syncCollections.Collection[0].Items[index];
                    break;
                }
            }

            foreach (Response.SyncCollectionsCollectionCommandsAdd add in commands.Add)
            {
                if (add.ServerId == serverId)
                {
                    for (int itemIndex = 0; itemIndex < add.ApplicationData.ItemsElementName.Length; itemIndex++)
                    {
                        if (add.ApplicationData.ItemsElementName[itemIndex] == elementType)
                        {
                            return add.ApplicationData.Items[itemIndex];
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Get the attachments of an email.
        /// </summary>
        /// <param name="syncResponse">The Sync command response.</param>
        /// <param name="subject">The email subject.</param>
        /// <returns>The attachments of the email.</returns>
        protected Response.AttachmentsAttachment[] GetEmailAttachments(SyncResponse syncResponse, string subject)
        {
            Response.AttachmentsAttachment[] attachments = null;

            // Get the application data of the email, to which the attachments belong.
            Response.SyncCollectionsCollectionCommandsAddApplicationData addData = TestSuiteBase.GetAddApplicationData(syncResponse, Response.ItemsChoiceType8.Subject1, subject);
            Site.Assert.IsNotNull(addData, string.Format("The email with subject '{0}' should exist.", subject));

            for (int i = 0; i < addData.ItemsElementName.Length; i++)
            {
                if (addData.ItemsElementName[i] == Response.ItemsChoiceType8.Attachments)
                {
                    Response.Attachments attachmentCollection = addData.Items[i] as Response.Attachments;
                    if (attachmentCollection != null)
                    {
                        attachments = new Response.AttachmentsAttachment[attachmentCollection.Items.Length];
                        for (int j = 0; j < attachmentCollection.Items.Length; j++)
                        {
                            attachments[j] = attachmentCollection.Items[j] as Response.AttachmentsAttachment;
                        }
                    }

                    break;
                }
            }
            return attachments;
        }

        /// <summary>
        /// Create an empty Sync request.
        /// </summary>
        /// <param name="collectionId">The value of the folder collectionId.</param>
        /// <returns>The empty Sync request.</returns>
        protected static SyncRequest CreateEmptySyncRequest(string collectionId)
        {
            return TestSuiteBase.CreateEmptySyncRequest(collectionId, -1);
        }

        /// <summary>
        /// Create an empty Sync request with filter type.
        /// </summary>
        /// <param name="collectionId">The value of the folder collectionId.</param>
        /// <param name="filterType">The value of the FilterType.</param>
        /// <returns>The empty Sync request.</returns>
        protected static SyncRequest CreateEmptySyncRequest(string collectionId, int filterType)
        {
            Request.SyncCollection collection = new Request.SyncCollection
            {
                SyncKey = "0",
                CollectionId = collectionId,
                Commands = null
            };

            if (0 <= filterType && filterType <= 8)
            {
                Request.Options option = new Request.Options
                {
                    Items = new object[] { (byte)filterType },
                    ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.FilterType }
                };
                collection.Options = new Request.Options[] { option };
            }

            return Common.CreateSyncRequest(new Request.SyncCollection[] { collection });
        }

        /// <summary>
        /// This method is used to create a Sync add request
        /// </summary>
        /// <param name="syncKey">The value of the SyncKey element.</param>
        /// <param name="collectionId">The value of the CollectionID element.</param>
        /// <param name="syncCollectionAdd">An instance of the SyncCollectionAdd.</param>
        /// <returns>The value of the SyncRequest.</returns>
        protected static SyncRequest CreateSyncAddRequest(string syncKey, string collectionId, Request.SyncCollectionAdd syncCollectionAdd)
        {
            Request.SyncCollection collection = new Request.SyncCollection
            {
                SyncKey = syncKey,
                GetChanges = true,
                CollectionId = collectionId,
                Commands = new object[] { syncCollectionAdd }
            };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { collection });
        }

        /// <summary>
        /// Create a SendMail request.
        /// </summary>
        /// <param name="from">The value of the From element.</param>
        /// <param name="to">The value of the To element.</param>
        /// <param name="cc">The value of the Cc element.</param>
        /// <param name="bcc">The value of the Bcc element.</param>
        /// <param name="subject">The value of the Subject element.</param>
        /// <param name="content">The value of the Content element.</param>
        /// <returns>The SendMail request.</returns>
        protected static SendMailRequest CreateSendMailRequest(string from, string to, string cc, string bcc, string subject, string content)
        {
            string clientID = System.Guid.NewGuid().ToString();
            string mime = Common.CreatePlainTextMime(from, to, cc, bcc, subject, content);

            return Common.CreateSendMailRequest(null, clientID, mime);
        }

        /// <summary>
        /// Create a SmartReply request.
        /// </summary>
        /// <param name="folderID">The value of the FolderID element</param>
        /// <param name="serverID">The value of the ServerId element.</param>
        /// <param name="from">The value of the From element.</param>
        /// <param name="to">The value of the To element.</param>
        /// <param name="cc">The value of the Cc element.</param>
        /// <param name="bcc">The value of the Bcc element.</param>
        /// <param name="subject">The value of the Subject element.</param>
        /// <param name="content">The value of the Content element.</param>
        /// <returns>the SmartReply request.</returns>
        protected static SmartReplyRequest CreateSmartReplyRequest(string folderID, string serverID, string from, string to, string cc, string bcc, string subject, string content)
        {
            SmartReplyRequest request = new SmartReplyRequest
            {
                RequestData = new Request.SmartReply
                {
                    ClientId = System.Guid.NewGuid().ToString(),
                    Source = new Request.Source { FolderId = folderID, ItemId = serverID },
                    Mime = Common.CreatePlainTextMime(@from, to, cc, bcc, subject, content)
                }
            };

            request.SetCommandParameters(new Dictionary<CmdParameterName, object>
            {
                {
                    CmdParameterName.CollectionId, folderID
                },
                {
                    CmdParameterName.ItemId, serverID
                }
            });

            return request;
        }

        /// <summary>
        /// Create a MeetingResponse request.
        /// </summary>
        /// <param name="userResponse">The way the user response the meeting.</param>
        /// <param name="collectionID">The collection id of the folder that contains the meeting request.</param>
        /// <param name="requestID">The server ID of the meeting request message item.</param>
        /// <param name="instanceID">The instance ID of the recurring meeting to be modified.</param>
        /// <returns>The MeetingResponse request.</returns>
        protected static MeetingResponseRequest CreateMeetingResponseRequest(byte userResponse, string collectionID, string requestID, string instanceID)
        {
            Request.MeetingResponseRequest request = new Request.MeetingResponseRequest
            {
                CollectionId = collectionID,
                RequestId = requestID,
                UserResponse = userResponse
            };

            if (!string.IsNullOrEmpty(instanceID))
            {
                request.InstanceId = instanceID;
            }

            return Common.CreateMeetingResponseRequest(new Request.MeetingResponseRequest[] { request });
        }

        /// <summary>
        /// Generate a Provision request.
        /// </summary>
        /// <returns>Provision request.</returns>
        protected static ProvisionRequest GenerateDefaultProvisionRequest()
        {
            Request.ProvisionPoliciesPolicy policy = new Request.ProvisionPoliciesPolicy
            {
                PolicyType = @"MS-EAS-Provisioning-WBXML"
            };

            Request.ProvisionPolicies policies = new Request.ProvisionPolicies { Policy = policy };
            Request.Provision requestData = new Request.Provision
            {
                Policies = policies,
                DeviceInformation = TestSuiteBase.GenerateDeviceInformation(),
                RemoteWipe = null
            };

            ProvisionRequest provisionRequest = new ProvisionRequest { RequestData = requestData };

            return provisionRequest;
        }

        /// <summary>
        /// Generate one device information.
        /// </summary>
        /// <returns>Device information.</returns>
        protected static Request.DeviceInformation GenerateDeviceInformation()
        {
            Request.DeviceInformationSet deviceInfoSet = new Request.DeviceInformationSet
            {
                FriendlyName = "test device",
                UserAgent = "test user agent",
                IMEI = "123456789012345",
                MobileOperator = "Microsoft",
                Model = "test model",
                OS = "windows7",
                OSLanguage = "en-us",
                PhoneNumber = "88888888888"
            };

            Request.DeviceInformation deviceInfo = new Request.DeviceInformation { Set = deviceInfoSet };

            return deviceInfo;
        }

        /// <summary>
        /// Create one OOF request with empty Settings
        /// </summary>
        /// <returns>The OOF settings request</returns>
        protected static SettingsRequest CreateDefaultOofRequest()
        {
            SettingsRequest settingsRequest = new SettingsRequest
            {
                RequestData = new Request.Settings { Oof = new Request.SettingsOof() }
            };
            return settingsRequest;
        }

        /// <summary>
        /// Create default OOF message
        /// </summary>
        /// <param name="bodyType">The OOF message body type</param>
        /// <param name="enabled">The enabled value</param>
        /// <param name="replyMessage">The replyMessage</param>
        /// <returns>The OOF Message</returns>
        protected static Request.OofMessage CreateOofMessage(string bodyType, string enabled, string replyMessage)
        {
            Request.OofMessage oofMessage = new Request.OofMessage
            {
                BodyType = bodyType,
                Enabled = enabled,
                ReplyMessage = replyMessage
            };
            return oofMessage;
        }

        /// <summary>
        /// Get AppliesToExternalKnown OOF message from SettingsResponse
        /// </summary>
        /// <param name="settingsResponse">The Settings response</param>
        /// <returns>The appliesToExternalKnown OOF message</returns>
        protected static Response.OofMessage GetAppliesToExternalKnownOofMessage(SettingsResponse settingsResponse)
        {
            for (int messageIndex = 0; messageIndex < settingsResponse.ResponseData.Oof.Get.OofMessage.Length; messageIndex++)
            {
                if (settingsResponse.ResponseData.Oof.Get.OofMessage[messageIndex].AppliesToExternalKnown != null)
                {
                    return settingsResponse.ResponseData.Oof.Get.OofMessage[messageIndex];
                }
            }

            return null;
        }

        /// <summary>
        /// Get AppliesToExternalUnknown OOF message from SettingsResponse
        /// </summary>
        /// <param name="settingsResponse">The settings response</param>
        /// <returns>The appliesToExternalUnknown OOF message</returns>
        protected static Response.OofMessage GetAppliesToExternalUnknownOofMessage(SettingsResponse settingsResponse)
        {
            for (int messageIndex = 0; messageIndex < settingsResponse.ResponseData.Oof.Get.OofMessage.Length; messageIndex++)
            {
                if (settingsResponse.ResponseData.Oof.Get.OofMessage[messageIndex].AppliesToExternalUnknown != null)
                {
                    return settingsResponse.ResponseData.Oof.Get.OofMessage[messageIndex];
                }
            }

            return null;
        }

        /// <summary>
        /// Get AppliesToInternal OOF message from SettingsResponse
        /// </summary>
        /// <param name="settingsResponse">The settings response</param>
        /// <returns>The appliesToInternal OOF message</returns>
        protected static Response.OofMessage GetAppliesToInternalOofMessage(SettingsResponse settingsResponse)
        {
            for (int messageIndex = 0; messageIndex < settingsResponse.ResponseData.Oof.Get.OofMessage.Length; messageIndex++)
            {
                if (settingsResponse.ResponseData.Oof.Get.OofMessage[messageIndex].AppliesToInternal != null)
                {
                    return settingsResponse.ResponseData.Oof.Get.OofMessage[messageIndex];
                }
            }

            return null;
        }
        #endregion

        #region Test case initialize and cleanup
        /// <summary>
        /// Initialize the test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.CMDAdapter = Site.GetAdapter<IMS_ASCMDAdapter>();
            this.CMDSUTControlAdapter = Site.GetAdapter<IMS_ASCMDSUTControlAdapter>();

            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.User1Information = new UserInformation()
            {
                UserName = Common.GetConfigurationPropertyValue("User1Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User1Password", this.Site),
                UserDomain = domain
            };

            this.User2Information = new UserInformation()
            {
                UserName = Common.GetConfigurationPropertyValue("User2Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User2Password", this.Site),
                UserDomain = domain
            };

            this.User3Information = new UserInformation()
            {
                UserName = Common.GetConfigurationPropertyValue("User3Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User3Password", this.Site),
                UserDomain = domain
            };

            this.User7Information = new UserInformation()
            {
                UserName = Common.GetConfigurationPropertyValue("User7Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User7Password", this.Site),
                UserDomain = domain
            };

            this.User8Information = new UserInformation()
            {
                UserName = Common.GetConfigurationPropertyValue("User8Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User8Password", this.Site),
                UserDomain = domain
            };

            this.User9Information = new UserInformation()
            {
                UserName = Common.GetConfigurationPropertyValue("User9Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User9Password", this.Site),
                UserDomain = domain
            };

            if (Common.GetSutVersion(this.Site) != SutVersion.ExchangeServer2007 || string.Equals(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "12.1"))
            {
                this.SwitchUser(this.User1Information);
            }
        }

        /// <summary>
        /// Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            this.ClearUserCreateItems(this.User1Information);
            this.ClearUserCreateItems(this.User2Information);
            this.ClearUserCreateItems(this.User3Information);
            this.ClearUserCreateItems(this.User7Information);
            this.ClearUserCreateItems(this.User8Information);
            this.ClearUserCreateItems(this.User9Information);

            if (this.changeDeviceIDSpecified)
            {
                // Restore DeviceID.
                this.CMDAdapter.ChangeDeviceType(Common.GetConfigurationPropertyValue("DeviceType", this.Site));
                this.CMDAdapter.ChangeDeviceID(Common.GetConfigurationPropertyValue("DeviceID", this.Site));
            }

            this.changeDeviceIDSpecified = false;

            if (this.changePolicyKeySpecified)
            {
                this.CMDAdapter.ChangePolicyKey(string.Empty);
            }

            this.changePolicyKeySpecified = false;

            if (this.IsOofSettingsChanged)
            {
                this.RevertOofSettings();
            }

            this.IsOofSettingsChanged = false;

            base.TestCleanup();
        }
        #endregion

        #region Protected Methods
        /// <summary>
        /// Get the value of SUT's AccessRights property.
        /// </summary>
        /// <param name="serverComputerName">The computer name of the server.</param>
        /// <param name="userInfo">The user information used to communicate with server.</param>
        /// <returns>The value of SUT's AccessRights property.</returns>
        protected string GetMailboxFolderPermission(string serverComputerName, UserInformation userInfo)
        {
            if (Common.GetSutVersion(this.Site) == SutVersion.ExchangeServer2007)
            {
                this.Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Exchange Server 2007 does not support ActiveSync protocol version 14.0.");

                this.Site.Assume.AreNotEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Exchange Server 2007 does not support ActiveSync protocol version 14.1");
            }

            return this.CMDSUTControlAdapter.GetMailboxFolderPermission(serverComputerName, userInfo.UserName, userInfo.UserPassword, userInfo.UserDomain);
        }

        /// <summary>
        /// Set SUT's AccessRights property to a specified value.
        /// </summary>
        /// <param name="serverComputerName">The computer name of the server.</param>
        /// <param name="userInfo">The user information used to communicate with server.</param>
        /// <param name="permission">The new value of AccessRights.</param>
        protected void SetMailboxFolderPermission(string serverComputerName, UserInformation userInfo, string permission)
        {
            if (Common.GetSutVersion(this.Site) == SutVersion.ExchangeServer2007)
            {
                this.Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Exchange Server 2007 does not support ActiveSync protocol version 14.0.");

                this.Site.Assume.AreNotEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Exchange Server 2007 does not support ActiveSync protocol version 14.1");
            }

            this.CMDSUTControlAdapter.SetMailboxFolderPermission(serverComputerName, userInfo.UserName, userInfo.UserPassword, userInfo.UserDomain, permission);
        }

        /// <summary>
        /// Check if an email has been in the specified folder with options.
        /// </summary>
        /// <param name="collectionId">The collection Id of the folder.</param>
        /// <param name="subject">The email subject.</param>
        /// <param name="options">The Options element to filter the items in the Sync response.</param>
        /// <returns>The Sync command response.</returns>
        protected SyncResponse CheckEmail(string collectionId, string subject, Request.Options[] options)
        {
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(collectionId);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;

            if (options != null)
            {
                syncRequest.RequestData.Collections[0].Options = options;
            }

            SyncResponse syncResponse = this.Sync(syncRequest);

            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", subject);
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 0;
            while ((counter < retryCount) && string.IsNullOrEmpty(serverId))
            {
                Thread.Sleep(waitTime);
                syncResponse = this.Sync(syncRequest);
                serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", subject);
                counter++;
            }

            Site.Assert.IsTrue(!string.IsNullOrEmpty(serverId), string.Format("The email with subject '{0}' should be found.", subject));

            return syncResponse;
        }

        /// <summary>
        /// Check the meeting forward notification mail which is sent from server to User.
        /// </summary>
        /// <param name="userInformation">The user who received notification</param>
        /// <param name="notificationSubject">The notification mail subject</param>
        protected void CheckMeetingForwardNotification(UserInformation userInformation, string notificationSubject)
        {
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            SyncResponse syncResult = this.SyncChanges(userInformation.DeletedItemsCollectionId);
            string serverID = TestSuiteBase.FindServerId(syncResult, "Subject", notificationSubject);
            while (serverID == null && counter < retryCount)
            {
                Thread.Sleep(waitTime);
                syncResult = this.SyncChanges(userInformation.DeletedItemsCollectionId);
                if (syncResult.ResponseDataXML != null)
                {
                    serverID = TestSuiteBase.FindServerId(syncResult, "Subject", notificationSubject);
                }

                counter++;
            }

            if (serverID != null)
            {
                TestSuiteBase.RecordCaseRelativeItems(userInformation, userInformation.DeletedItemsCollectionId, notificationSubject);
            }
        }

        /// <summary>
        /// Create folder.
        /// </summary>
        /// <param name="folderType">The folder type.</param>
        /// <param name="folderName">The folder name.</param>
        /// <param name="parentFolderID">The parent folder serverID.</param>
        /// <returns>Created folder serverID.</returns>
        protected string CreateFolder(byte folderType, string folderName, string parentFolderID)
        {
            // Call method FolderSync to synchronize the collection hierarchy.
            FolderSyncResponse folderSyncResponse = this.FolderSync();

            // Call method FolderCreate to create a new folder as a child folder of the specified parent folder, return ServerId for FolderCreate.
            FolderCreateRequest folderCreateRequest = Common.CreateFolderCreateRequest(folderSyncResponse.ResponseData.SyncKey, folderType, folderName, parentFolderID);
            FolderCreateResponse folderCreateResponse = this.CMDAdapter.FolderCreate(folderCreateRequest);
            Site.Assert.AreEqual(1, Convert.ToInt32(folderCreateResponse.ResponseData.Status), "If the FolderCreate command executes successfully, the Status in response should be 1");

            // Record created folder collectionID.
            string folderId = folderCreateResponse.ResponseData.ServerId;
            return folderId;
        }

        /// <summary>
        /// Create a request to add a contact with job title.
        /// </summary>
        /// <param name="firstName">The first name of the contact.</param>
        /// <param name="middleName">The middle name of the contact.</param>
        /// <param name="lastName">The last name of the contact.</param>
        /// <param name="fileAs">The filing string for the contact.</param>
        /// <param name="jobTitle">The job title of the contact.</param>
        /// <returns>The request to add a contact.</returns>
        protected Request.SyncCollectionAdd CreateAddContactCommand(string firstName, string middleName, string lastName, string fileAs, string jobTitle)
        {
            Request.SyncCollectionAdd appData = new Request.SyncCollectionAdd
            {
                ClientId = TestSuiteBase.ClientId,
                ApplicationData = new Request.SyncCollectionAddApplicationData()
            };

            firstName = Common.GenerateResourceName(this.Site, firstName);
            middleName = Common.GenerateResourceName(this.Site, middleName);
            lastName = Common.GenerateResourceName(this.Site, lastName);

            if (string.IsNullOrEmpty(jobTitle))
            {
                appData.ApplicationData.ItemsElementName = new Request.ItemsChoiceType8[] { Request.ItemsChoiceType8.FileAs, Request.ItemsChoiceType8.FirstName, Request.ItemsChoiceType8.MiddleName, Request.ItemsChoiceType8.LastName };
                appData.ApplicationData.Items = new object[] { fileAs, firstName, middleName, lastName };
            }
            else
            {
                jobTitle = Common.GenerateResourceName(this.Site, jobTitle);
                appData.ApplicationData.ItemsElementName = new Request.ItemsChoiceType8[] { Request.ItemsChoiceType8.FileAs, Request.ItemsChoiceType8.FirstName, Request.ItemsChoiceType8.MiddleName, Request.ItemsChoiceType8.LastName, Request.ItemsChoiceType8.JobTitle };
                appData.ApplicationData.Items = new object[] { fileAs, firstName, middleName, lastName, jobTitle };
            }

            if ("12.1" != Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site))
            {
                appData.Class = "Contacts";
            }

            return appData;
        }

        /// <summary>
        /// Send an email with a normal attachment
        /// </summary>
        /// <param name="subject">The subject of the mail.</param>
        /// <param name="body">The body of the item.</param>
        protected void SendEmailWithAttachment(string subject, string body)
        {
            string from = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string to = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            string mime = TestSuiteBase.CreateMIME(from, to, subject, body);

            SendMailRequest request = new SendMailRequest
            {
                RequestData = { ClientId = TestSuiteBase.ClientId, Mime = mime }
            };

            SendMailResponse sendMailResponse = this.CMDAdapter.SendMail(request);
            Site.Assert.AreEqual<string>(
                 string.Empty,
                 sendMailResponse.ResponseDataXML,
                 "The server should return an empty XML body to indicate SendMail command is executed successfully.");
        }

        /// <summary>
        /// Synchronizes the folder hierarchy.
        /// </summary>
        /// <returns>The response of FolderSync method.</returns>
        protected FolderSyncResponse FolderSync()
        {
            FolderSyncResponse folderSyncResponse = this.CMDAdapter.FolderSync(Common.CreateFolderSyncRequest("0"));
            return folderSyncResponse;
        }

        /// <summary>
        /// Create a SmartForward request.
        /// </summary>
        /// <param name="folderId">The value of the FolderId element.</param>
        /// <param name="serverId">The value of the ServerId element.</param>
        /// <param name="from">The value of the From element.</param>
        /// <param name="to">The value of the To element.</param>
        /// <param name="cc">The value of the Cc element.</param>
        /// <param name="bcc">The value of the Bcc element.</param>
        /// <param name="subject">The value of the Subject element.</param>
        /// <param name="content">The value of the Content element.</param>
        /// <returns>the SmartForward request.</returns>
        protected SmartForwardRequest CreateSmartForwardRequest(string folderId, string serverId, string from, string to, string cc, string bcc, string subject, string content)
        {
            SmartForwardRequest request = new SmartForwardRequest
            {
                RequestData = new Request.SmartForward
                {
                    ClientId = TestSuiteBase.ClientId,
                    Source = new Request.Source { FolderId = folderId, ItemId = serverId }
                }
            };

            string mime = Common.CreatePlainTextMime(from, to, cc, bcc, subject, content);
            request.RequestData.Mime = mime;
            request.SetCommandParameters(new Dictionary<CmdParameterName, object>
            {
                {
                    CmdParameterName.CollectionId, folderId
                },
                {
                    CmdParameterName.ItemId, serverId
                }
            });

            return request;
        }

        /// <summary>
        /// Get the Status value in a response returned by the SendStringRequest operation.
        /// </summary>
        /// <param name="responseDataXml">The response data string returned by the SendStringRequest operation.</param>
        /// <returns>The status code for SendStringRequest operation.</returns>
        protected string GetStatusCode(string responseDataXml)
        {
            if (responseDataXml != null && !string.IsNullOrEmpty(responseDataXml))
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(responseDataXml);

                XmlNodeList nodes = doc.GetElementsByTagName("Status");
                Site.Assert.IsNotNull(nodes, "The server response should contain a Status element.");
                return nodes[0].InnerText;
            }

            return null;
        }

        /// <summary>
        /// Synchronizes changes in a collection between the client and the server.
        /// </summary>
        /// <param name="request">A SyncRequest object that contains the request information.</param>
        /// <param name="isResyncNeeded">A boolean value indicate whether need to re-sync when the response contains MoreAvailable.</param>
        /// <returns>The Sync command response.</returns>
        protected SyncResponse Sync(SyncRequest request, bool isResyncNeeded = true)
        {
            SyncResponse response = this.CMDAdapter.Sync(request, isResyncNeeded);

            // Get the SyncKey returned in the last Sync command response.
            if (response != null
                && response.ResponseData != null
                && response.ResponseData.Item != null)
            {
                Response.SyncCollections syncCollections = response.ResponseData.Item as Response.SyncCollections;
                if (syncCollections != null)
                {
                    foreach (Response.SyncCollectionsCollection syncCollection in syncCollections.Collection)
                    {
                        for (int i = 0; i < syncCollection.ItemsElementName.Length; i++)
                        {
                            if (syncCollection.ItemsElementName[i] == Response.ItemsChoiceType10.SyncKey)
                            {
                                this.LastSyncKey = syncCollection.Items[i] as string;
                            }
                        }
                    }
                }
            }
            else
            {
                this.LastSyncKey = null;
            }

            return response;
        }

        /// <summary>
        /// Establishes a synchronization relationship with the server and initializes the synchronization state.
        /// </summary>
        /// <param name="collectionId">The value of the folder collectionId.</param>
        /// <returns>The Sync command response.</returns>
        protected SyncResponse GetInitialSyncResponse(string collectionId)
        {
            // Call method FolderSync to synchronize the collection hierarchy.
            this.FolderSync();

            // Call method Sync to synchronize changes.
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(collectionId);
            return this.Sync(syncRequest);
        }

        /// <summary>
        /// Change user to call FolderSync command to synchronize the collection hierarchy.
        /// </summary>
        /// <param name="userInformation">The user information that contains case related information</param>
        protected void SwitchUser(UserInformation userInformation)
        {
            this.CMDAdapter.SwitchUser(userInformation.UserName, userInformation.UserPassword, userInformation.UserDomain);
            if (string.IsNullOrEmpty(userInformation.InboxCollectionId))
            {
                FolderSyncResponse folderSyncResponse = this.CMDAdapter.FolderSync(Common.CreateFolderSyncRequest("0"));
                Site.Assert.AreEqual<int>(
                     1,
                     int.Parse(folderSyncResponse.ResponseData.Status),
                     "The server should return a status code 1 in the FolderSync command response to indicate success.");
                this.LastFolderSyncKey = folderSyncResponse.ResponseData.SyncKey;

                userInformation.TasksCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Tasks, this.Site);
                userInformation.NotesCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Notes, this.Site);
                userInformation.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, this.Site);
                userInformation.ContactsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Contacts, this.Site);
                userInformation.CalendarCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Calendar, this.Site);
                userInformation.SentItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.SentItems, this.Site);
                userInformation.DeletedItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.DeletedItems, this.Site);
                userInformation.RecipientInformationCacheCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.RecipientInformationCache, this.Site);
            }
        }

        /// <summary>
        /// Send a meeting request.
        /// </summary>
        /// <param name="subject">The subject of email</param>
        /// <param name="calendar">The meeting calendar</param>
        protected void SendMeetingRequest(string subject, Calendar calendar)
        {
            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0")&&!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"))
            {
                string emailBody = Common.GenerateResourceName(Site, "content");
                string icalendarFormatContent = TestSuiteBase.CreateiCalendarFormatContent(calendar);

                string meetingEmailMime = Common.CreateMeetingRequestMime(
                    calendar.OrganizerEmail,
                    calendar.Attendees.Attendee[0].Email,
                    subject,
                    emailBody,
                    icalendarFormatContent);
                string clientID = TestSuiteBase.ClientId;

                SendMailRequest sendMailRequest = Common.CreateSendMailRequest(clientID, false, meetingEmailMime);

                this.SwitchUser(this.User1Information);
                SendMailResponse response = this.CMDAdapter.SendMail(sendMailRequest);

                Site.Assert.AreEqual<string>(
                     string.Empty,
                     response.ResponseDataXML,
                     "The server should return an empty xml response data to indicate SendMail command success.");
            }
        }

        /// <summary>
        /// Send a weekly meeting request.
        /// </summary>
        /// <param name="meetingRequestSubject">The subject of the meeting request.</param>
        /// <param name="recipientEmailAddress">The email address of the recipient.</param>
        protected void SendWeeklyRecurrenceMeetingRequest(string meetingRequestSubject, string recipientEmailAddress)
        {
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, recipientEmailAddress, null);
            int offset = new Random().Next(1, 100000);

            calendar.DtStamp = DateTime.UtcNow.AddDays(offset);
            calendar.StartTime = DateTime.UtcNow.AddDays(offset);
            calendar.EndTime = DateTime.UtcNow.AddDays(offset).AddHours(2);

            // Set recurrence to weekly
            Response.Recurrence recurrence = new Response.Recurrence
            {
                Type = 1,
                Interval = 1,
                DayOfWeek = 2,
                DayOfWeekSpecified = true,
                Until = DateTime.UtcNow.AddDays(offset + 20).ToString("yyyyMMddTHHmmssZ")
            };

            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.1"))
            {
                recurrence.FirstDayOfWeek = 1;
                recurrence.FirstDayOfWeekSpecified = true;
            }

            calendar.Recurrence = recurrence;

            this.SendMeetingRequest(meetingRequestSubject, calendar);
        }

        /// <summary>
        /// Send a plain text email between User1 and User2.
        /// </summary>
        /// <param name="accountID">The account from which an email is sent.</param>
        /// <param name="emailSubject">Email subject.</param>
        /// <param name="senderName">The sender of the email.</param>
        /// <param name="recipientName">The receiver of the email.</param>
        /// <param name="content">The email content.</param>
        /// <returns>SendMail command response from the server.</returns>
        protected SendMailResponse SendPlainTextEmail(string accountID, string emailSubject, string senderName, string recipientName, string content)
        {
            string from = string.Empty;
            string to = string.Empty;

            // Switch to user1 mailbox.
            if (string.IsNullOrEmpty(senderName))
            {
                from = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
                this.SwitchUser(this.User1Information);
            }
            else
            {
                if (senderName == this.User1Information.UserName)
                {
                    this.SwitchUser(this.User1Information);
                    from = Common.GetMailAddress(senderName, this.User1Information.UserDomain);
                }
                else if (senderName == this.User2Information.UserName)
                {
                    this.SwitchUser(this.User2Information);
                    from = Common.GetMailAddress(senderName, this.User2Information.UserDomain);
                }
                else
                {
                    Site.Assert.Fail("The sender's name is not existed in the current context.");
                }
            }

            if (string.IsNullOrEmpty(recipientName) || recipientName == this.User2Information.UserName)
            {
                to = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            }
            else if (recipientName == this.User1Information.UserName)
            {
                to = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            }
            else
            {
                Site.Assert.Fail("The receiver's name is not existed in the current context.");
            }

            string body = string.IsNullOrEmpty(content)
                            ? Common.GenerateResourceName(Site, "Default Email")
                            : content;

            string mime = Common.CreatePlainTextMime(from, to, null, null, emailSubject, body);
            SendMailRequest sendMailRequest = Common.CreateSendMailRequest(TestSuiteBase.ClientId, false, mime);
            if (accountID != null)
            {
                sendMailRequest.RequestData.AccountId = accountID;
            }

            return this.CMDAdapter.SendMail(sendMailRequest);
        }

        /// <summary>
        /// User2 sends mail to User1 and does FolderSync in User1's mailbox.
        /// </summary>
        /// <returns>The subject of the sent message.</returns>
        protected string SendMailAndFolderSync()
        {
            #region User2 calls method SendMail to send MIME-formatted e-mail messages to User1
            this.SwitchUser(this.User2Information);
            string subject = Common.GenerateResourceName(Site, "subject");
            string content = "The content of the body.";
            SendMailRequest sendMailRequest = TestSuiteBase.CreateSendMailRequest(Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain), Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain), string.Empty, subject, content);
            SendMailResponse sendMailResponse = this.CMDAdapter.SendMail(sendMailRequest);
            Site.Assert.AreEqual<string>(
                string.Empty,
                sendMailResponse.ResponseDataXML,
                "The server should return an empty xml response data to indicate SendMail command success.");
            #endregion

            #region User1 calls method FolderSync to synchronize the collection hierarchy, return folder collectionIds
            this.SwitchUser(this.User1Information);
            #endregion

            #region Record user name, folder collectionId and item subject that are used in case
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, subject);
            #endregion

            return subject;
        }

        /// <summary>
        /// Synchronize the changes of the specific folder.
        /// </summary>
        /// <param name="collectionID">Folder's collectionID.</param>
        /// <param name="isResyncNeeded">A boolean value indicate whether need to re-sync when the response contains MoreAvailable.</param>
        /// <returns>Sync response from the server.</returns>
        protected SyncResponse SyncChanges(string collectionID, bool isResyncNeeded = true)
        {
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(collectionID), isResyncNeeded);
            return this.SyncChanges(this.LastSyncKey, collectionID, isResyncNeeded);
        }

        /// <summary>
        /// Synchronize the changes from last synchronization in specific folder.
        /// </summary>
        /// <param name="syncKey">The sync key.</param>
        /// <param name="collectionID">Folder's collectionID.</param>
        /// <param name="isResyncNeeded">A boolean value indicate whether need to re-sync when the response contains MoreAvailable.</param>
        /// <returns>Sync response from the server.</returns>
        protected SyncResponse SyncChanges(string syncKey, string collectionID, bool isResyncNeeded = true)
        {
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(collectionID);
            syncRequest.RequestData.Collections[0].SyncKey = syncKey;
            Request.Options options = new Request.Options();
            Request.BodyPreference bodyPreference = new Request.BodyPreference { Type = 1 };
            options.Items = new object[] { bodyPreference };
            options.ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.BodyPreference };
            syncRequest.RequestData.Collections[0].Options = new Request.Options[] { options };
            syncRequest.RequestData.Collections[0].GetChanges = true;
            syncRequest.RequestData.Collections[0].GetChangesSpecified = true;
            SyncResponse syncResponse = this.Sync(syncRequest, isResyncNeeded);
            return syncResponse;
        }

        /// <summary>
        /// Get email with special subject.
        /// </summary>
        /// <param name="folderID">The folderID that store mail items.</param>
        /// <param name="subject">Email subject.</param>
        /// <returns>Sync result.</returns>
        protected SyncResponse GetMailItem(string folderID, string subject)
        {
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            SyncResponse syncResult = this.SyncChanges(folderID);
            string serverID = TestSuiteBase.FindServerId(syncResult, "Subject", subject);
            while (serverID == null && counter < retryCount)
            {
                Thread.Sleep(waitTime);
                syncResult = this.SyncChanges(folderID);
                if (syncResult.ResponseDataXML != null)
                {
                    serverID = TestSuiteBase.FindServerId(syncResult, "Subject", subject);
                }

                counter++;
            }

            Site.Assert.IsNotNull(serverID, "The email item with subject '{0}' should be found, retry count: {1}.", subject, counter);
            return syncResult;
        }

        /// <summary>
        /// Confirm the item with specified subject exist in folder.
        /// </summary>
        /// <param name="folderID">The folder item located.</param>
        /// <param name="subject">The item subject.</param>
        /// <returns>The item serverID.</returns>
        protected string GetItemServerIdFromSpecialFolder(string folderID, string subject)
        {
            SyncResponse syncFolderResponse = this.GetMailItem(folderID, subject);
            string itemServerID = TestSuiteBase.FindServerId(syncFolderResponse, "Subject", subject);
            Site.Assert.IsFalse(string.IsNullOrEmpty(itemServerID), "The item's server id should not be null.");

            return itemServerID;
        }

        /// <summary>
        /// Record user has changed device information.
        /// </summary>
        protected void RecordDeviceInfoChanged()
        {
            this.changeDeviceIDSpecified = true;
        }

        /// <summary>
        /// Record user has changed PolicyKey.
        /// </summary>
        protected void RecordPolicyKeyChanged()
        {
            this.changePolicyKeySpecified = true;
        }

        /// <summary>
        /// Delete folder with specified collectionID.
        /// </summary>
        /// <param name="collectionIDList">The collectionID of folder that need to be deleted.</param>
        protected void DeleteFolder(Collection<string> collectionIDList)
        {
            if (collectionIDList.Count > 0)
            {
                foreach (string collectionID in collectionIDList)
                {
                    FolderSyncResponse folderSyncResponse = this.FolderSync();
                    FolderDeleteRequest folderDeleteRequest = Common.CreateFolderDeleteRequest(folderSyncResponse.ResponseData.SyncKey, collectionID);
                    FolderDeleteResponse folderDeleteResponse = this.CMDAdapter.FolderDelete(folderDeleteRequest);
                    Site.Assert.AreEqual<int>(1, int.Parse(folderDeleteResponse.ResponseData.Status), "The created Folder should be deleted.");
                }
            }
        }

        /// <summary>
        /// Clear a user's ActiveSync device, this user should have permission to delete ActiveSync device.
        /// </summary>
        /// <param name="userName">The user name.</param>
        /// <param name="userPassword">The user password.</param>
        /// <param name="domain">The user domain.</param>
        protected void ClearDevice(string userName, string userPassword, string domain)
        {
            bool isDeviceDeleted = this.CMDSUTControlAdapter.DeleteDevice(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                userName,
                userPassword,
                domain);
            Site.Assert.IsTrue(isDeviceDeleted, "The user's ActiveSync device should be deleted");
        }

        /// <summary>
        /// Get FolderCreate command response.
        /// </summary>
        /// <param name="syncKey">The folder SyncKey.</param>
        /// <param name="folderType">The folder type.</param>
        /// <param name="folderName">The folder name.</param>
        /// <param name="parentFolderID">The parent folder serverID.</param>
        /// <returns>Return response from the server.</returns>
        protected FolderCreateResponse GetFolderCreateResponse(string syncKey, byte folderType, string folderName, string parentFolderID)
        {
            // Call method FolderCreate to create a new folder as a child folder of the specified parent folder, return response.
            FolderCreateRequest folderCreateRequest = Common.CreateFolderCreateRequest(syncKey, folderType, folderName, parentFolderID);
            FolderCreateResponse folderCreateResponse = this.CMDAdapter.FolderCreate(folderCreateRequest);
            return folderCreateResponse;
        }

        /// <summary>
        /// Get OOF settings
        /// </summary>
        /// <returns>The OOF setting response</returns>
        protected SettingsResponse GetOofSettings()
        {
            SettingsRequest settingsRequest = new SettingsRequest
            {
                RequestData = new Request.Settings
                {
                    Oof = new Request.SettingsOof { Item = new Request.SettingsOofGet { BodyType = "TEXT" } }
                }
            };

            SettingsResponse settingsResponse = this.CMDAdapter.Settings(settingsRequest);
            return settingsResponse;
        }

        /// <summary>
        /// Create one sample calendar object.
        /// </summary>
        /// <param name="subject">Meeting subject.</param>
        /// <param name="attendeeEmailAddress">Meeting attendee email address.</param>
        /// <param name="createdCalendar">The calendar object</param>
        /// <returns>One sample calendar object.</returns> 
        protected Calendar CreateCalendar(string subject, string attendeeEmailAddress, Calendar createdCalendar)
        {
            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = this.SetMeetingProperties(subject, attendeeEmailAddress, this.Site);
            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                elementsToValueMap.Add(Request.ItemsChoiceType8.ResponseRequested, true);
            }

            if (createdCalendar != null)
            {
                if (createdCalendar.DtStamp != null)
                {
                    elementsToValueMap.Add(Request.ItemsChoiceType8.DtStamp, DateTime.Parse(createdCalendar.DtStamp.ToString()).ToString("yyyyMMddTHHmmssZ"));
                }
                if (createdCalendar.StartTime != null)
                {
                    elementsToValueMap.Add(Request.ItemsChoiceType8.StartTime, DateTime.Parse(createdCalendar.StartTime.ToString()).ToString("yyyyMMddTHHmmssZ"));
                }
                if (createdCalendar.EndTime != null)
                {
                    elementsToValueMap.Add(Request.ItemsChoiceType8.EndTime, DateTime.Parse(createdCalendar.EndTime.ToString()).ToString("yyyyMMddTHHmmssZ"));
                }
            }
            // Call Sync command with Add element to add a meeting
            Request.SyncCollectionAddApplicationData applicationData = new Request.SyncCollectionAddApplicationData
            {
                Items = new object[elementsToValueMap.Count],
                ItemsElementName = new Request.ItemsChoiceType8[elementsToValueMap.Count]
            };

            if (elementsToValueMap.Count > 0)
            {
                elementsToValueMap.Values.CopyTo(applicationData.Items, 0);
                elementsToValueMap.Keys.CopyTo(applicationData.ItemsElementName, 0);
            }

            Request.SyncCollectionAdd calendarData = new Request.SyncCollectionAdd()
            {
                ClientId = TestSuiteBase.ClientId,
                ApplicationData = applicationData,
            };
            this.GetInitialSyncResponse(this.User1Information.CalendarCollectionId);
            SyncRequest syncAddRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, calendarData);
            SyncResponse syncResponse = this.Sync(syncAddRequest);

            SyncStore getChangeResult = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync calendar = this.GetSyncAddItem(getChangeResult, subject);

            return calendar.Calendar;
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Delete all the user created items
        /// </summary>
        /// <param name="userInformation">The user information which contains user created items</param>
        private void DeleteItemsInFolder(UserInformation userInformation)
        {
            foreach (CreatedItems userItem in userInformation.UserCreatedItems)
            {
                SyncRequest emptySyncRequest = TestSuiteBase.CreateEmptySyncRequest(userItem.CollectionId);
                SyncResponse emptySyncResult = this.Sync(emptySyncRequest);
                SyncStore emptySyncResponse = Common.LoadSyncResponse(emptySyncResult);
                int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
                int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
                int counter = 0;

                do
                {
                    Thread.Sleep(waitTime);

                    emptySyncResult = this.Sync(emptySyncRequest);
                    emptySyncResponse = Common.LoadSyncResponse(emptySyncResult);
                    if (emptySyncResponse != null)
                    {
                        if (emptySyncResponse.CollectionStatus == 1)
                        {
                            break;
                        }
                    }

                    counter++;
                }
                while (counter < retryCount / 10);

                if (emptySyncResponse.AddElements != null)
                {
                    SyncRequest deleteRequest;
                    foreach (SyncItem item in emptySyncResponse.AddElements)
                    {
                        deleteRequest = CreateSyncPermanentDeleteRequest(emptySyncResponse.SyncKey, userItem.CollectionId, item.ServerId);
                        SyncResponse resultDelete = this.Sync(deleteRequest);
                        SyncStore deleteResult = Common.LoadSyncResponse(resultDelete);
                        Site.Assert.AreEqual<byte>(1, deleteResult.CollectionStatus, "Item should be deleted.");
                    }
                }

                SyncResponse result = this.SyncChanges(userItem.CollectionId);
                string syncKey = this.LastSyncKey;
                if (result.ResponseData != null && result.ResponseData.Item != null)
                {
                    List<Request.SyncCollectionDelete> deleteData = new List<Request.SyncCollectionDelete>();
                    foreach (string subject in userItem.ItemSubject)
                    {
                        if (userItem.CollectionId == userInformation.ContactsCollectionId)
                        {
                            foreach (string itemServerID in TestSuiteBase.FindServerIdList(result, "FileAs", subject))
                            {
                                Request.SyncCollectionDelete deleteItem = new Request.SyncCollectionDelete
                                {
                                    ServerId = itemServerID
                                };

                                deleteData.Add(deleteItem);
                            }
                        }
                        else
                        {
                            foreach (string itemServerID in TestSuiteBase.FindServerIdList(result, "Subject", subject))
                            {
                                Request.SyncCollectionDelete deleteItem = new Request.SyncCollectionDelete
                                {
                                    ServerId = itemServerID
                                };

                                deleteData.Add(deleteItem);
                            }
                        }
                    }

                    if (deleteData.Count > 0)
                    {
                        Request.SyncCollection syncCollection = new Request.SyncCollection
                        {
                            SyncKey = syncKey,
                            GetChanges = true,
                            GetChangesSpecified = true,
                            CollectionId = userItem.CollectionId,
                            Commands = deleteData.ToArray(),
                            DeletesAsMoves = false,
                            DeletesAsMovesSpecified = true
                        };

                        SyncRequest syncRequest = Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
                        SyncResponse deleteResult = this.Sync(syncRequest);
                        string deleteResultStatus = this.GetStatusCode(deleteResult.ResponseDataXML);

                        Site.Assert.AreEqual<string>(
                            "1",
                            deleteResultStatus,
                            "The value of 'Status' should be 1 which indicates the Sync command executes successfully.",
                            deleteResultStatus);
                    }
                }
            }
        }

        /// <summary>
        /// Clear all user created items
        /// </summary>
        /// <param name="userInformation">The user related information</param>
        private void ClearUserCreateItems(UserInformation userInformation)
        {
            // Clean up specified user created items.
            if (userInformation.UserCreatedItems.Count > 0)
            {
                this.SwitchUser(userInformation);
                this.DeleteItemsInFolder(userInformation);
            }

            if (userInformation.UserCreatedFolders.Count > 0)
            {
                this.SwitchUser(userInformation);
                this.DeleteFolder(userInformation.UserCreatedFolders);
            }
        }

        /// <summary>
        /// Revert the settings of Oof message.
        /// </summary>
        private void RevertOofSettings()
        {
            #region Creates Setting request for three types of audiences
            SettingsRequest settingsRequest = CreateDefaultOofRequest();
            Request.SettingsOofSet settingsOofSet = new Request.SettingsOofSet();
            settingsOofSet.OofState = Request.OofState.Item0;
            settingsOofSet.OofStateSpecified = true;

            string bodyType = "TEXT";
            string enabled = "1";

            Request.OofMessage oofMessageWithNothing = CreateOofMessage(bodyType, enabled, null);
            oofMessageWithNothing.AppliesToInternal = string.Empty;
            oofMessageWithNothing.AppliesToExternalKnown = string.Empty;
            oofMessageWithNothing.AppliesToExternalUnknown = string.Empty;

            settingsOofSet.OofMessage = new Request.OofMessage[] { oofMessageWithNothing };
            settingsRequest.RequestData.Oof.Item = settingsOofSet;
            #endregion

            SettingsResponse settingsResponseAfterSet = this.CMDAdapter.Settings(settingsRequest);
            Site.Assert.AreEqual<string>("1", settingsResponseAfterSet.ResponseData.Oof.Status, "Server should response status 1, if set enabled successful");

            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            Response.OofMessage internalOofMessage = null;
            Response.OofMessage externalKnownOofMessage = null;
            Response.OofMessage externalUnknownOofMessage = null;
            do
            {
                Thread.Sleep(waitTime);
                SettingsResponse settingsResponse = this.GetOofSettings();
                internalOofMessage = GetAppliesToInternalOofMessage(settingsResponse);
                externalKnownOofMessage = GetAppliesToExternalKnownOofMessage(settingsResponse);
                externalUnknownOofMessage = GetAppliesToExternalUnknownOofMessage(settingsResponse);

                counter++;
            }
            while (counter < retryCount &&
                (internalOofMessage.AppliesToInternal == null ||
                internalOofMessage.Enabled != "0" ||
                externalKnownOofMessage.AppliesToExternalKnown == null ||
                externalKnownOofMessage.Enabled != "0" ||
                externalUnknownOofMessage.AppliesToExternalUnknown == null ||
                externalUnknownOofMessage.Enabled != "0"));

            Site.Assert.AreEqual<string>("0", internalOofMessage.Enabled, "The oof message settings for internal users should be disenabled. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>("0", externalKnownOofMessage.Enabled, "The oof message settings for known external users should be disenabled. Retry count: {0}", counter);
        }

        /// <summary>
        /// Set the value of common meeting properties
        /// </summary>
        /// <param name="subject">The subject of the meeting.</param>
        /// <param name="attendeeEmailAddress">The email address of attendee.</param>
        /// <returns>The key and value pairs of common meeting properties.</returns>
        private Dictionary<Request.ItemsChoiceType8, object> SetMeetingProperties(string subject, string attendeeEmailAddress, ITestSite testSite)
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
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", testSite).Equals("16.0") || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", testSite).Equals("16.1"))
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

            List<Request.AttendeesAttendee> attendeelist = new List<Request.AttendeesAttendee>
            {
                new Request.AttendeesAttendee
                {
                    Email = attendeeEmailAddress,
                    Name = attendeeEmailAddress,
                    AttendeeStatus = 0,
                    AttendeeTypeSpecified=true,
                    AttendeeType = 1
                }
            };

            // Set the attendee to user2
            Request.Attendees attendees = new Request.Attendees { Attendee = attendeelist.ToArray() };
            propertiesToValueMap.Add(Request.ItemsChoiceType8.Attendees, attendees);

            return propertiesToValueMap;
        }

        /// <summary>
        /// Get the specified email item.
        /// </summary>
        /// <param name="emailSubject">The subject of the email item.</param>
        /// <param name="folderCollectionId">The serverId of the default folder.</param>
        /// <param name="bodyPreference">The preference information related to the type and size of information that is returned from fetching.</param>
        /// <returns>The result of getting the specified email item.</returns>
        protected SyncStore GetSyncResult(string emailSubject, string folderCollectionId, Request.BodyPreference bodyPreference)
        {
            SyncStore syncItemResult;
            Sync item = null;
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            do
            {
                Thread.Sleep(waitTime);

                // Get the new added email item
                this.GetInitialSyncResponse(folderCollectionId);
                SyncResponse response = this.SyncChanges(this.LastSyncKey, folderCollectionId);
                syncItemResult = Common.LoadSyncResponse(response); ;
                if (syncItemResult != null && syncItemResult.CollectionStatus == 1)
                {
                    item = this.GetSyncAddItem(syncItemResult, emailSubject);
                }

                counter++;
            }
            while ((syncItemResult == null || item == null) && counter < retryCount);

            Site.Assert.IsNotNull(item, "The email item with subject {0} should be found. Retry count: {1}", emailSubject, counter);

            // Verify sync result
            Site.Assert.AreEqual<byte>(
                1,
                syncItemResult.CollectionStatus,
                "If the Sync command executes successfully, the status in response should be 1.");

            return syncItemResult;
        }

        /// <summary>
        /// Get the specified email item from the sync add response by using the subject as the search criteria.
        /// </summary>
        /// <param name="syncResult">The sync result.</param>
        /// <param name="subject">The email subject.</param>
        /// <returns>Return the specified email item.</returns>
        private Sync GetSyncAddItem(SyncStore syncResult, string subject)
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
        #endregion
    }
}