namespace Microsoft.Protocols.TestSuites.MS_ASEMAIL
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Net.Mail;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// The base class of scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Properties
        /// <summary>
        /// Gets protocol Interface of MS-ASEMAIL
        /// </summary>
        protected IMS_ASEMAILAdapter EMAILAdapter { get; private set; }

        /// <summary>
        /// Gets or sets the information of User1.
        /// </summary>
        protected UserInformation User1Information { get; set; }

        /// <summary>
        /// Gets or sets the information of User2.
        /// </summary>
        protected UserInformation User2Information { get; set; }

        /// <summary>
        /// Gets or sets the information of User3.
        /// </summary>
        protected UserInformation User3Information { get; set; }

        /// <summary>
        /// Gets or sets the information of User4.
        /// </summary>
        protected UserInformation User4Information { get; set; }

        /// <summary>
        /// Gets or sets the information of User5.
        /// </summary>
        protected UserInformation User5Information { get; set; }
        #endregion

        #region Create sync delete operation request
        /// <summary>
        /// Create a Sync delete operation request which would be used to delete items permanently.
        /// </summary>
        /// <param name="syncKey">The synchronization state of a collection.</param>
        /// <param name="collectionId">The server ID of the folder.</param>
        /// <param name="serverId">The server ID of the item which will be deleted.</param>
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
        #endregion

        #region Test case initialize and cleanup
        /// <summary>
        /// Override the base TestInitialize function
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();

            if (this.EMAILAdapter == null)
            {
                this.EMAILAdapter = Site.GetAdapter<IMS_ASEMAILAdapter>();
            }

            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);

            this.User1Information = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User1Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User1Password", this.Site),
                UserDomain = domain
            };

            this.User2Information = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User2Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User2Password", this.Site),
                UserDomain = domain
            };

            this.User3Information = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User3Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User3Password", this.Site),
                UserDomain = domain
            };

            this.User4Information = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User4Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User4Password", this.Site),
                UserDomain = domain
            };

            this.User5Information = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User5Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User5Password", this.Site),
                UserDomain = domain
            };

            if (Common.GetSutVersion(this.Site) != SutVersion.ExchangeServer2007 || string.Equals(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "12.1"))
            {
                this.SwitchUser(this.User1Information, true);
            }
        }

        /// <summary>
        /// Override the base TestCleanup function
        /// </summary>
        protected override void TestCleanup()
        {
            if (this.User1Information.UserCreatedItems.Count != 0)
            {
                this.SwitchUser(this.User1Information, false);
                this.DeleteItemsInFolder(this.User1Information);
            }

            if (this.User2Information.UserCreatedItems.Count != 0)
            {
                this.SwitchUser(this.User2Information, false);
                this.DeleteItemsInFolder(this.User2Information);
            }

            if (this.User3Information.UserCreatedItems.Count != 0)
            {
                this.SwitchUser(this.User3Information, false);
                this.DeleteItemsInFolder(this.User3Information);
            }

            if (this.User4Information.UserCreatedItems.Count != 0)
            {
                this.SwitchUser(this.User4Information, false);
                this.DeleteItemsInFolder(this.User4Information);
            }

            if (this.User5Information.UserCreatedItems.Count != 0)
            {
                this.SwitchUser(this.User5Information, false);
                this.DeleteItemsInFolder(this.User5Information);
            }

            base.TestCleanup();
        }
        #endregion

        #region Initialize sync with server
        /// <summary>
        /// Sync changes between client and server
        /// </summary>
        /// <param name="syncKey">The synchronization key returned by last request.</param>
        /// <param name="collectionId">Identify the folder as the collection being synchronized.</param>
        /// <param name="bodyPreference">Sets preference information related to the type and size of information for body</param>
        /// <returns>Return change result</returns>
        protected SyncStore SyncChanges(string syncKey, string collectionId, Request.BodyPreference bodyPreference)
        {
            // Get changes from server use initial syncKey
            SyncRequest syncRequest = TestSuiteHelper.CreateSyncRequest(syncKey, collectionId, bodyPreference);
            SyncStore syncResult = this.EMAILAdapter.Sync(syncRequest);

            return syncResult;
        }

        /// <summary>
        /// Initialize the sync with server
        /// </summary>
        /// <param name="collectionId">Specify the folder collection Id which needs to be synced.</param>
        /// <returns>Return change result</returns>
        protected SyncStore InitializeSync(string collectionId)
        {
            // Obtains the key by sending an initial Sync request with a SyncKey element value of zero and the CollectionId element
            SyncRequest syncRequest = Common.CreateInitialSyncRequest(collectionId);
            SyncStore syncResult = this.EMAILAdapter.Sync(syncRequest);

            // Verify sync change result
            Site.Assert.AreEqual<byte>(
                1,
                syncResult.CollectionStatus,
                "If the Sync command executes successfully, the Status in response should be 1.");

            return syncResult;
        }
        #endregion

        #region SwitchUser
        /// <summary>
        /// Change user to call ActiveSync operations and resynchronize the collection hierarchy.
        /// </summary>
        /// <param name="userInformation">The information of the user.</param>
        /// <param name="isFolderSyncNeeded">A Boolean value that indicates whether needs to synchronize the folder hierarchy.</param>
        protected void SwitchUser(UserInformation userInformation, bool isFolderSyncNeeded)
        {
            this.EMAILAdapter.SwitchUser(userInformation.UserName, userInformation.UserPassword, userInformation.UserDomain);

            if (isFolderSyncNeeded)
            {
                // Call FolderSync command to synchronize the collection hierarchy.
                FolderSyncRequest folderSyncRequest = Common.CreateFolderSyncRequest("0");
                FolderSyncResponse folderSyncResponse = this.EMAILAdapter.FolderSync(folderSyncRequest);

                // Verify FolderSync command response.
                Site.Assert.AreEqual<int>(
                    1,
                    int.Parse(folderSyncResponse.ResponseData.Status),
                    "If the FolderSync command executes successfully, the Status in response should be 1.");
                if (string.IsNullOrEmpty(userInformation.InboxCollectionId))
                {
                    userInformation.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, this.Site);
                }

                if (string.IsNullOrEmpty(userInformation.DeletedItemsCollectionId))
                {
                    userInformation.DeletedItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.DeletedItems, this.Site);
                }

                if (string.IsNullOrEmpty(userInformation.CalendarCollectionId))
                {
                    userInformation.CalendarCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Calendar, this.Site);
                }
            }
        }
        #endregion

        #region Send email
        /// <summary>
        /// Send a plain text email.
        /// </summary>
        /// <param name="subject">The subject of email</param>
        /// <param name="cc">The cc address of the mail</param>
        /// <param name="bcc">The bcc address of the mail</param>
        /// <param name="to">The to address of the mail</param>
        /// <param name="sender">The sender address of the mail</param>
        /// <param name="replyTo">The replyTo address of the mail</param>
        /// <param name="from">The from address of the mail</param>
        protected void SendPlaintextEmail(
            string subject,
            string cc,
            string bcc,
            string to,
            string sender,
            string replyTo,
            string from)
        {
            string emailBody = Common.GenerateResourceName(Site, "content");

            string emailMime = TestSuiteHelper.CreatePlainTextMime(
                string.IsNullOrEmpty(from) ? Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain) : from,
                string.IsNullOrEmpty(to) ? Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain) : to,
                cc,
                bcc,
                subject,
                emailBody,
                sender,
                replyTo);
            string clientId = TestSuiteHelper.GetClientId();

            SendMailRequest sendMailRequest = TestSuiteHelper.CreateSendMailRequest(clientId, false, emailMime);
            this.SwitchUser(this.User1Information, false);
            SendMailResponse response = this.EMAILAdapter.SendMail(sendMailRequest);

            Site.Assert.AreEqual<string>(
                 string.Empty,
                 response.ResponseDataXML,
                 "The server should return an empty xml response data to indicate SendMail command was executed successfully.");
        }

        /// <summary>
        /// Send a plain text email.
        /// </summary>
        /// <param name="subject">The subject of email</param>
        /// <param name="cc">The cc address of the mail</param>
        /// <param name="bcc">The bcc address of the mail</param>
        protected void SendPlaintextEmail(string subject, string cc, string bcc)
        {
            string emailBody = Common.GenerateResourceName(Site, "content");
            string emailMime = TestSuiteHelper.CreatePlainTextMime(
                Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain),
                Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain),
                cc,
                bcc,
                subject,
                emailBody);
            string clientId = TestSuiteHelper.GetClientId();

            SendMailRequest sendMailRequest = TestSuiteHelper.CreateSendMailRequest(clientId, false, emailMime);
            this.SwitchUser(this.User1Information, false);
            SendMailResponse response = this.EMAILAdapter.SendMail(sendMailRequest);

            Site.Assert.AreEqual<string>(
                 string.Empty,
                 response.ResponseDataXML,
                 "The server should return an empty xml response data to indicate SendMail command executes successfully.");

            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
        }

        /// <summary>
        /// Send a meeting request email.
        /// </summary>
        /// <param name="subject">The subject of email</param>
        /// <param name="calendar">The meeting calendar</param>
        protected void SendMeetingRequest(string subject, Calendar calendar)
        {
            string emailBody = Common.GenerateResourceName(Site, "content");
            string icalendarFormatContent = TestSuiteHelper.CreateiCalendarFormatContent(calendar);

            string meetingEmailMime = TestSuiteHelper.CreateMeetingRequestMime(
                calendar.OrganizerEmail,
                calendar.Attendees.Attendee[0].Email,
                subject,
                emailBody,
                icalendarFormatContent);
            string clientId = TestSuiteHelper.GetClientId();

            SendMailRequest sendMailRequest = TestSuiteHelper.CreateSendMailRequest(clientId, false, meetingEmailMime);
            this.SwitchUser(this.User1Information, false);
            SendMailResponse response = this.EMAILAdapter.SendMail(sendMailRequest);

            Site.Assert.AreEqual<string>(
                 string.Empty,
                 response.ResponseDataXML,
                 "The server should return an empty xml response data to indicate SendMail command success.");
        }

        /// <summary>
        /// Create a default calendar object in the current login user calendar folder
        /// </summary>
        /// <param name="subject">The calendar subject</param>
        /// <param name="organizerEmailAddress">The organizer email address</param>
        /// <param name="attendeeEmailAddress">The attendee email address</param>
        /// <param name="calendarUID">The uid of calendar</param>
        /// <param name="timestamp">The DtStamp of calendar</param>
        /// <param name="startTime">The StartTime of calendar</param>
        /// <param name="endTime">The EndTime of calendar</param>
        /// <returns>Returns the Calendar instance</returns>
        protected Calendar CreateDefaultCalendar(
            string subject,
            string organizerEmailAddress,
            string attendeeEmailAddress,
            string calendarUID,
            DateTime? timestamp,
            DateTime? startTime,
            DateTime? endTime)
        {
            #region Configure the default calendar application data
            Request.SyncCollectionAdd syncAddCollection = new Request.SyncCollectionAdd();
            string clientId = TestSuiteHelper.GetClientId();
            syncAddCollection.ClientId = clientId;
            syncAddCollection.ApplicationData = new Request.SyncCollectionAddApplicationData();

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType8> itemsElementName = new List<Request.ItemsChoiceType8>();

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                items.Add(true);
                itemsElementName.Add(Request.ItemsChoiceType8.ResponseRequested);
            }
            #region TIME/Subject/Location/UID
            items.Add(string.Format("{0:yyyyMMddTHHmmss}Z", null == startTime ? DateTime.UtcNow.AddDays(5) : startTime.Value));
            itemsElementName.Add(Request.ItemsChoiceType8.StartTime);

            items.Add(string.Format("{0:yyyyMMddTHHmmss}Z", null == endTime ? DateTime.UtcNow.AddDays(5).AddMinutes(30) : endTime.Value));
            itemsElementName.Add(Request.ItemsChoiceType8.EndTime);

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0")&&!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"))
            {
                items.Add(string.Format("{0:yyyyMMddTHHmmss}Z", null == timestamp ? DateTime.UtcNow.AddDays(5) : timestamp.Value));
                itemsElementName.Add(Request.ItemsChoiceType8.DtStamp);
            }

            items.Add(subject);
            itemsElementName.Add(Request.ItemsChoiceType8.Subject);

            items.Add(calendarUID ?? Guid.NewGuid().ToString());
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0") || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"))
            {
                itemsElementName.Add(Request.ItemsChoiceType8.ClientUid);
            }
            else
            {
                itemsElementName.Add(Request.ItemsChoiceType8.UID);
            }

            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0")|| Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"))
            {
                Request.Location location = new Request.Location();
                location.DisplayName = "OFFICE";
                items.Add(location);
                itemsElementName.Add(Request.ItemsChoiceType8.Location);
            }
            else
            {
                items.Add("OFFICE");
                itemsElementName.Add(Request.ItemsChoiceType8.Location1);
            }
            #endregion

            #region Attendee/Organizer
            Request.AttendeesAttendee attendee = new Request.AttendeesAttendee
            {
                Email = attendeeEmailAddress,
                Name = new MailAddress(attendeeEmailAddress).User,
                AttendeeStatus = 0x0,
                AttendeeTypeSpecified = true,
                AttendeeType = 0x1
            };

            // 0x0 = Response unknown

            // 0x1 = Required
            items.Add(new Request.Attendees() { Attendee = new Request.AttendeesAttendee[] { attendee } });
            itemsElementName.Add(Request.ItemsChoiceType8.Attendees);

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0")&& !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"))
            {
                items.Add(organizerEmailAddress);
                itemsElementName.Add(Request.ItemsChoiceType8.OrganizerEmail);
                items.Add(new MailAddress(organizerEmailAddress).DisplayName);
                itemsElementName.Add(Request.ItemsChoiceType8.OrganizerName);
            }
            #endregion

            #region Sensitivity/BusyStatus/AllDayEvent
            // 0x0 == Normal
            items.Add((byte)0x0);
            itemsElementName.Add(Request.ItemsChoiceType8.Sensitivity);

            // 0x1 == Tentative
            items.Add((byte)0x1);
            itemsElementName.Add(Request.ItemsChoiceType8.BusyStatus);

            // 0x0 not an all-day event
            items.Add((byte)0x0);
            itemsElementName.Add(Request.ItemsChoiceType8.AllDayEvent);
            #endregion

            syncAddCollection.ApplicationData.Items = items.ToArray();
            syncAddCollection.ApplicationData.ItemsElementName = itemsElementName.ToArray();
            #endregion

            #region Execute the Sync command to upload the calendar
            SyncStore initSyncResponse = this.InitializeSync(this.User1Information.CalendarCollectionId);
            SyncRequest uploadCalendarRequest = TestSuiteHelper.CreateSyncAddRequest(initSyncResponse.SyncKey, this.User1Information.CalendarCollectionId, syncAddCollection);
            this.EMAILAdapter.Sync(uploadCalendarRequest);
            #endregion

            #region Get the new added calendar item
            SyncStore getItemResponse = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync calendarItem = TestSuiteHelper.GetSyncAddItem(getItemResponse, subject);
            Site.Assert.IsNotNull(calendarItem, "The item with subject {0} should be found in the folder {1}.", subject, FolderType.Calendar.ToString());
            #endregion

            return calendarItem.Calendar;
        }

        /// <summary>
        /// Send a meeting response email
        /// </summary>
        /// <param name="calendar">The meeting calendar</param>
        protected void SendMeetingResponse(Calendar calendar)
        {
            // Create reply mail to organizer
            string emailBody = Common.GenerateResourceName(Site, "content");
            string icalendarResponseContent = TestSuiteHelper.CreateMeetingResponseiCalendarFormatContent(
                (DateTime)calendar.DtStamp,
                (DateTime)calendar.EndTime,
                calendar.UID,
                calendar.Subject,
                calendar.Location,
                calendar.OrganizerEmail,
                calendar.Attendees.Attendee[0].Email);

            // Create reply mail mime content
            string meetingResponseEmailMime = TestSuiteHelper.CreateMeetingRequestMime(
                calendar.Attendees.Attendee[0].Email,
                calendar.OrganizerEmail,
                calendar.Subject,
                emailBody,
                icalendarResponseContent);

            string clientId = TestSuiteHelper.GetClientId();
            SendMailRequest sendMailRequest = TestSuiteHelper.CreateSendMailRequest(clientId, false, meetingResponseEmailMime);
            this.SwitchUser(this.User2Information, true);
            SendMailResponse response = this.EMAILAdapter.SendMail(sendMailRequest);

            Site.Assert.AreEqual<string>(
                 string.Empty,
                 response.ResponseDataXML,
                 "The server should return an empty xml response data to indicate SendMail command success.");
        }
        #endregion

        #region Get Sync add result
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
                SyncStore initSyncResult = this.InitializeSync(folderCollectionId);
                syncItemResult = this.SyncChanges(initSyncResult.SyncKey, folderCollectionId, bodyPreference);
                if (syncItemResult != null && syncItemResult.CollectionStatus == 1)
                {
                    item = TestSuiteHelper.GetSyncAddItem(syncItemResult, emailSubject);
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

        #endregion

        #region Update email
        /// <summary>
        /// Update email
        /// </summary>
        /// <param name="collectionId">The collectionId of the folder which contains the item to be updated.</param>
        /// <param name="syncKey">The syncKey which is returned from server</param>
        /// <param name="read">The value is TRUE indicates the email has been read; a value of FALSE indicates the email has not been read</param>
        /// <param name="serverId">The server id of the email</param>
        /// <param name="flag">The flag instance</param>
        /// <param name="categories">The array of categories</param>
        /// <returns>Return update email result</returns>
        protected SyncStore UpdateEmail(string collectionId, string syncKey, bool? read, string serverId, Request.Flag flag, Collection<string> categories)
        {
            Request.SyncCollectionChange changeData = new Request.SyncCollectionChange
            {
                ServerId = serverId,
                ApplicationData = new Request.SyncCollectionChangeApplicationData()
            };

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType7> itemsElementName = new List<Request.ItemsChoiceType7>();

            if (null != read)
            {
                items.Add(read);
                itemsElementName.Add(Request.ItemsChoiceType7.Read);
            }

            if (null != flag)
            {
                items.Add(flag);
                itemsElementName.Add(Request.ItemsChoiceType7.Flag);
            }

            if (null != categories)
            {
                Request.Categories2 mailCategories = new Request.Categories2 { Category = new string[categories.Count] };
                categories.CopyTo(mailCategories.Category, 0);
                items.Add(mailCategories);
                itemsElementName.Add(Request.ItemsChoiceType7.Categories2);
            }

            changeData.ApplicationData.Items = items.ToArray();
            changeData.ApplicationData.ItemsElementName = itemsElementName.ToArray();

            SyncRequest syncRequest = TestSuiteHelper.CreateSyncChangeRequest(syncKey, collectionId, changeData);
            SyncStore result = this.EMAILAdapter.Sync(syncRequest);
            Site.Assert.AreEqual<byte>(
                1,
                result.CollectionStatus,
                "The server returns a Status 1 in the Sync command response indicate sync command success.");

            return result;
        }

        /// <summary>
        /// Update email with more data
        /// </summary>
        /// <param name="collectionId">The collectionId of the folder which contains the item to be updated.</param>
        /// <param name="syncKey">The syncKey which is returned from server</param>
        /// <param name="read">The value is TRUE indicates the email has been read; a value of FALSE indicates the email has not been read</param>
        /// <param name="serverId">The server id of the email</param>
        /// <param name="flag">The flag instance</param>
        /// <param name="categories">The list of categories</param>
        /// <param name="additionalElement">Additional flag element</param>
        /// <param name="insertTag">Additional element will insert before this tag</param>
        /// <returns>Return update email result</returns>
        protected SendStringResponse UpdateEmailWithMoreData(string collectionId, string syncKey, bool read, string serverId, Request.Flag flag, Collection<object> categories, string additionalElement, string insertTag)
        {
            // Create normal sync request
            Request.SyncCollectionChange changeData = TestSuiteHelper.CreateSyncChangeData(read, serverId, flag, categories);
            SyncRequest syncRequest = TestSuiteHelper.CreateSyncChangeRequest(syncKey, collectionId, changeData);

            // Calls Sync command to update email with invalid sync request
            SendStringResponse result = this.EMAILAdapter.InvalidSync(syncRequest, additionalElement, insertTag);
            return result;
        }
        #endregion

        #region Add a meeting to the server
        /// <summary>
        /// Add a meeting to the server.
        /// </summary>
        /// <param name="calendarCollectionId">The collectionId of the folder which the item should be added.</param>
        /// <param name="elementsToValueMap">The key and value pairs of common meeting properties.</param>
        protected void SyncAddMeeting(string calendarCollectionId, Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap)
        {
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

            SyncStore iniSync = this.InitializeSync(calendarCollectionId);
            SyncRequest syncAddRequest = TestSuiteHelper.CreateSyncAddRequest(iniSync.SyncKey, calendarCollectionId, applicationData);

            SyncStore syncAddResponse = this.EMAILAdapter.Sync(syncAddRequest);
            Site.Assert.AreEqual<int>(
                1,
                int.Parse(syncAddResponse.AddResponses[0].Status),
                "The sync add operation should be successful.");
        }
        #endregion

        #region Record the userName, folder collectionId and item subject
        /// <summary>
        /// Record the user name, folder collectionId and subjects the current test case impacts.
        /// </summary>
        /// <param name="userName">The user that current test case used.</param>
        /// <param name="folderCollectionId">The collectionId of folders that the current test case impact.</param>
        /// <param name="itemSubjects">The subject of items that the current test case impact.</param>
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
            else if (userName == this.User2Information.UserName)
            {
                this.User2Information.UserCreatedItems.Add(createdItems);
            }
            else if (userName == this.User3Information.UserName)
            {
                this.User3Information.UserCreatedItems.Add(createdItems);
            }
            else if (userName == this.User4Information.UserName)
            {
                this.User4Information.UserCreatedItems.Add(createdItems);
            }
            else if (userName == this.User5Information.UserName)
            {
                this.User5Information.UserCreatedItems.Add(createdItems);
            }
        }

        #endregion

        #region Private method
        /// <summary>
        /// Delete all the items in a folder.
        /// </summary>
        /// <param name="userInformation">The user information which contains user created items</param>
        private void DeleteItemsInFolder(UserInformation userInformation)
        {
            foreach (CreatedItems createdItems in userInformation.UserCreatedItems)
            {
                SyncStore syncStore = this.InitializeSync(createdItems.CollectionId);
                SyncStore result = this.SyncChanges(syncStore.SyncKey, createdItems.CollectionId, null);
                string syncKey = result.SyncKey;
                int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
                int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
                int counter = 0;
                do
                {
                    Thread.Sleep(waitTime);
                    if (result != null)
                    {
                        if (result.CollectionStatus == 1)
                        {
                            break;
                        }
                    }

                    counter++;
                }
                while (counter < retryCount / 10);
                if (result.AddElements != null)
                {
                    SyncRequest deleteRequest;
                    foreach (Sync syncItem in result.AddElements)
                    {
                        if (createdItems.CollectionId == userInformation.CalendarCollectionId)
                        {
                            foreach (string subject in createdItems.ItemSubject)
                            {
                                if (syncItem.Calendar != null)
                                {
                                    if (syncItem.Calendar.Subject.Equals(subject, StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        deleteRequest = CreateSyncPermanentDeleteRequest(syncKey, createdItems.CollectionId, syncItem.ServerId);
                                        SyncStore deleteSyncResult = this.EMAILAdapter.Sync(deleteRequest);
                                        syncKey = deleteSyncResult.SyncKey;
                                        Site.Assert.AreEqual<byte>(1, deleteSyncResult.CollectionStatus, "Item should be deleted.");
                                    }
                                }
                            }
                        }
                        else
                        {
                            List<Request.SyncCollectionDelete> deleteData = new List<Request.SyncCollectionDelete>();
                            foreach (string subject in createdItems.ItemSubject)
                            {
                                string serverId = null;
                                if (result != null)
                                {
                                    foreach (Sync item in result.AddElements)
                                    {
                                        if (item.Email.Subject != null && item.Email.Subject.Contains(subject))
                                        {
                                            serverId = item.ServerId;
                                            break;
                                        }

                                        if (item.Calendar.Subject != null && item.Calendar.Subject.Contains(subject))
                                        {
                                            serverId = item.ServerId;
                                            break;
                                        }
                                    }
                                }

                                if (serverId != null)
                                {
                                    deleteData.Add(new Request.SyncCollectionDelete() { ServerId = serverId });
                                }
                            }

                            if (deleteData.Count > 0)
                            {
                                Request.SyncCollection syncCollection = TestSuiteHelper.CreateSyncCollection(syncKey, createdItems.CollectionId);
                                syncCollection.Commands = deleteData.ToArray();
                                syncCollection.DeletesAsMoves = false;
                                syncCollection.DeletesAsMovesSpecified = true;

                                SyncRequest syncRequest = Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
                                SyncStore deleteResult = this.EMAILAdapter.Sync(syncRequest);
                                syncKey = deleteResult.SyncKey;
                                Site.Assert.AreEqual<byte>(
                                    1,
                                    deleteResult.CollectionStatus,
                                    "The value of Status should be 1 to indicate the Sync command executes successfully.");
                            }
                        }
                    }
                }
            }
        }
        #endregion
    }
}