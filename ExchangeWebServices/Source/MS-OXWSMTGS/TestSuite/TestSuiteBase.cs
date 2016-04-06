namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A bass class for scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Fields
        /// <summary>
        /// Indicate the value of Location element. 
        /// </summary>
        private string location;

        /// <summary>
        /// The name of the user whose role is organizer of a meeting or appointment.
        /// </summary>
        private string organizer;

        /// <summary>
        /// The corresponding password of the user whose role is organizer of a meeting or appointment.
        /// </summary>
        private string organizerPassword;

        /// <summary>
        /// The name of the user whose role is attendeeType of a meeting.
        /// </summary>
        private string attendee;

        /// <summary>
        /// The corresponding password of the user whose role is attendeeType of a meeting.
        /// </summary>
        private string attendeePassword;

        /// <summary>
        /// The name of the user that delegate user exists.
        /// </summary>
        private string delegateUser;

        /// <summary>
        /// The corresponding password of the delegate user.
        /// </summary>
        private string delegatePassword;

        /// <summary>
        /// The wait time for meeting request, meeting response or meeting cancellation message to be received.
        /// </summary>
        private int waitTime;

        /// <summary>
        /// Indicate the value of Subject element.
        /// </summary>
        private string subject;

        /// <summary>
        /// Indicate the value of MeetingWorkspaceUrl element.
        /// </summary>
        private string meetingWorkspaceUrl;

        /// <summary>
        /// Indicate the value of NetShowLocation element.
        /// </summary>
        private string netShowUrl;

        /// <summary>
        /// Indicate the updated value of Location element.
        /// </summary>
        private string locationUpdate;

        /// <summary>
        /// The email address of the user whose role is organizer of a meeting or appointment.
        /// </summary>
        private string organizerEmailAddress;

        /// <summary>
        /// The email address of the user whose role is attendeeType of a meeting.
        /// </summary>
        private string attendeeEmailAddress;

        /// <summary>
        /// The email address of the user that delegate user exists.
        /// </summary>
        private string delegateEmailAddress;

        /// <summary>
        /// The email address of the resource.
        /// </summary>
        private string roomEmailAddress;

        /// <summary>
        /// Indicate the value of conferenceType element.
        /// </summary>
        private int conferenceType;

        /// <summary>
        /// Indicate the value of Interval element, defined in [MS-OXWSCDATA].
        /// </summary>
        private int patternInterval;

        /// <summary>
        /// Indicate the value of NumberOfOccurrences element, defined in [MS-OXWSCDATA].
        /// </summary>
        private int numberOfOccurrences;

        /// <summary>
        /// Indicate the value of LegacyFreeBusy element, defined in [MS-OXWSCDATA].
        /// </summary>
        private LegacyFreeBusyType legacyFreeBusy;

        /// <summary>
        /// Indicate the value of InstanceIndex element, defined in [MS-OXWSCORE].
        /// </summary>
        private int instanceIndex;

        /// <summary>
        /// Indicate the updated value of Subject element used in UpdateItem.
        /// </summary>
        private string subjectUpdate;

        /// <summary>
        /// Indicate the interval between the calendar start/end time and current time.
        /// </summary>
        private int timeInterval;

        /// <summary>
        /// Indicate the upper bound of a For loop for waiting for a valid Server response. 
        /// </summary>
        private int upperBound;

        /// <summary>
        /// Indicate the current value of BaseShape property.
        /// </summary>
        private DefaultShapeNamesType baseShape;

        /// <summary>
        /// Indicate the current value of MessageDisposition property.
        /// </summary>
        private MessageDispositionType messageDisposition;

        /// <summary>
        /// Domain of server.
        /// </summary>
        private string domain;
        #endregion

        #region Properties
        /// <summary>
        /// Gets or sets the MS-OXWSMTGS protocol adapter.
        /// </summary>
        protected IMS_OXWSMTGSAdapter MTGSAdapter { get; set; }

        /// <summary>
        /// Gets or sets the MS-OXWSSRCH protocol adapter which used to supply FindItem method.
        /// </summary>
        protected IMS_OXWSSRCHAdapter SRCHAdapter { get; set; }

        /// <summary>
        /// Gets or sets the MS-OXWSFOLD protocol adapter.
        /// </summary>
        protected IMS_OXWSFOLDAdapter FOLDAdapter { get; set; }

        /// <summary>
        /// Gets the value of Location element.
        /// </summary>
        protected string Location
        {
            get { return this.location; }
        }

        /// <summary>
        /// Gets the value of Domain element. 
        /// </summary>
        protected string Domain
        {
            get { return this.domain; }
        }

        /// <summary>
        /// Gets the name of the user whose role is organizer of a meeting or appointment.
        /// </summary>
        protected string Organizer
        {
            get { return this.organizer; }
        }

        /// <summary>
        /// Gets the corresponding password of the user whose role is organizer of a meeting or appointment.
        /// </summary>
        protected string OrganizerPassword
        {
            get { return this.organizerPassword; }
        }

        /// <summary>
        /// Gets the name of the user whose role is attendeeType of a meeting.
        /// </summary>
        protected string Attendee
        {
            get { return this.attendee; }
        }

        /// <summary>
        /// Gets the corresponding password of the user whose role is attendeeType of a meeting.
        /// </summary>
        protected string AttendeePassword
        {
            get { return this.attendeePassword; }
        }

        /// <summary>
        /// Gets the name of the user that delegate user exists.
        /// </summary>
        protected string DelegateUser
        {
            get { return this.delegateUser; }
        }

        /// <summary>
        /// Gets the corresponding password of the delegate user.
        /// </summary>
        protected string DelegatePassword
        {
            get { return this.delegatePassword; }
        }

        /// <summary>
        /// Gets the wait time for meeting request, meeting response or meeting cancellation message to be received.
        /// </summary>
        protected int WaitTime
        {
            get { return this.waitTime; }
        }

        /// <summary>
        /// Gets the value of Subject element.
        /// </summary>
        protected string Subject
        {
            get { return this.subject; }
        }

        /// <summary>
        /// Gets the value of MeetingWorkspaceUrl element.
        /// </summary>
        protected string MeetingWorkspace
        {
            get { return this.meetingWorkspaceUrl; }
        }

        /// <summary>
        /// Gets the value of NetShowUrl element.
        /// </summary>
        protected string NetShowLocation
        {
            get { return this.netShowUrl; }
        }

        /// <summary>
        /// Gets the updated value of Location element.
        /// </summary>
        protected string LocationUpdate
        {
            get { return this.locationUpdate; }
        }

        /// <summary>
        /// Gets the email address of the user whose role is organizer of a meeting or appointment.
        /// </summary>
        protected string OrganizerEmailAddress
        {
            get { return this.organizerEmailAddress; }
        }

        /// <summary>
        /// Gets the email address of the user whose role is attendeeType of a meeting.
        /// </summary>
        protected string AttendeeEmailAddress
        {
            get { return this.attendeeEmailAddress; }
        }

        /// <summary>
        /// Gets the email address of the user that delegate user exists.
        /// </summary>
        protected string DelegateEmailAddress
        {
            get { return this.delegateEmailAddress; }
        }

        /// <summary>
        /// Gets the email address of the resource.
        /// </summary>
        protected string RoomEmailAddress
        {
            get { return this.roomEmailAddress; }
        }

        /// <summary>
        /// Gets the value of conferenceType element.
        /// </summary>
        protected int ConferenceType
        {
            get { return this.conferenceType; }
        }

        /// <summary>
        /// Gets the value of Interval element.
        /// </summary>
        protected int PatternInterval
        {
            get { return this.patternInterval; }
        }

        /// <summary>
        /// Gets the value of NumberOfOccurrences element.
        /// </summary>
        protected int NumberOfOccurrences
        {
            get { return this.numberOfOccurrences; }
        }

        /// <summary>
        /// Gets the value of LegacyFreeBusy element.
        /// </summary>
        protected LegacyFreeBusyType LegacyFreeBusy
        {
            get { return this.legacyFreeBusy; }
        }

        /// <summary>
        /// Gets the value of InstanceIndex element.
        /// </summary>
        protected int InstanceIndex
        {
            get { return this.instanceIndex; }
        }

        /// <summary>
        /// Gets the updated value of Subject element used in UpdateItem.
        /// </summary>
        protected string SubjectUpdate
        {
            get { return this.subjectUpdate; }
        }

        /// <summary>
        /// Gets the interval between the calendar start/end time and current time.
        /// </summary>
        protected int TimeInterval
        {
            get { return this.timeInterval; }
        }

        /// <summary>
        /// Gets the upper bound of a loop for waiting for a valid Server response.
        /// </summary>
        protected int UpperBound
        {
            get { return this.upperBound; }
        }

        /// <summary>
        /// Gets the current value of BaseShape property.
        /// </summary>
        protected DefaultShapeNamesType BaseShape
        {
            get { return this.baseShape; }
        }

        /// <summary>
        /// Gets the current value of MessageDisposition property.
        /// </summary>
        protected MessageDispositionType MessageDisposition
        {
            get { return this.messageDisposition; }
        }

        /// <summary>
        /// Gets or sets the folder id to be deleted.
        /// </summary>
        protected FolderIdType FolderToDelete
        {
            get;
            set;
        }
        #endregion

        #region Static methods
        /// <summary>
        /// Get an attendeeType or resource instance.
        /// </summary>
        /// <param name="emailAddress">The email address related to an attendeeType or resource.</param>
        /// <returns>An instance of AttendeeType</returns>
        protected static AttendeeType GetAttendeeOrResource(string emailAddress)
        {
            EmailAddressType email = new EmailAddressType();
            email.EmailAddress = emailAddress;
            AttendeeType attendeeType = new AttendeeType();
            attendeeType.Mailbox = email;
            return attendeeType;
        }
        #endregion

        #region Test case initialize and clean up
        /// <summary>
        ///  Initialize the test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.MTGSAdapter = Site.GetAdapter<IMS_OXWSMTGSAdapter>();
            this.SRCHAdapter = Site.GetAdapter<IMS_OXWSSRCHAdapter>();
            this.FOLDAdapter = Site.GetAdapter<IMS_OXWSFOLDAdapter>();

            #region Get property values
            this.domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.location = Common.GetConfigurationPropertyValue("Location", this.Site);
            this.subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("MeetingSubject", this.Site));
            this.meetingWorkspaceUrl = Common.GetConfigurationPropertyValue("MeetingWorkspaceUrl", this.Site);
            this.netShowUrl = Common.GetConfigurationPropertyValue("NetShowUrl", this.Site);
            this.organizer = Common.GetConfigurationPropertyValue("OrganizerName", this.Site);
            this.organizerPassword = Common.GetConfigurationPropertyValue("OrganizerPassword", this.Site);
            this.attendee = Common.GetConfigurationPropertyValue("AttendeeName", this.Site);
            this.attendeePassword = Common.GetConfigurationPropertyValue("AttendeePassword", this.Site);
            this.delegateUser = Common.GetConfigurationPropertyValue("DelegateName", this.Site);
            this.delegatePassword = Common.GetConfigurationPropertyValue("DelegatePassword", this.Site);
            this.locationUpdate = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("LocationUpdate", this.Site));
            this.organizerEmailAddress = this.organizer + "@" + this.domain;
            this.attendeeEmailAddress = this.attendee + "@" + this.domain;
            this.delegateEmailAddress = this.delegateUser + "@" + this.domain;
            this.roomEmailAddress = Common.GetConfigurationPropertyValue("RoomName", this.Site) + "@" + this.domain;
            this.waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            
            this.conferenceType = 2;
            this.patternInterval = int.Parse(Common.GetConfigurationPropertyValue("PatternInterval", this.Site));
            this.numberOfOccurrences = int.Parse(Common.GetConfigurationPropertyValue("NumberOfOccurrences", this.Site));
            this.instanceIndex = int.Parse(Common.GetConfigurationPropertyValue("InstanceIndex", this.Site));
            this.subjectUpdate = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("SubjectUpdate", this.Site));
            this.timeInterval = int.Parse(Common.GetConfigurationPropertyValue("TimeInterval", this.Site));
            this.baseShape = DefaultShapeNamesType.AllProperties;
            this.messageDisposition = MessageDispositionType.SendAndSaveCopy;
            this.legacyFreeBusy = LegacyFreeBusyType.Busy;
            this.FolderToDelete = null;

            #region The upper bound of request loop
            if (!int.TryParse(Common.GetConfigurationPropertyValue("RetryCount", this.Site), out this.upperBound))
            {
                Site.Assert.Fail("The value of RetryCount property in ptf file was not converted to an integer value.");
            }
            #endregion
            #endregion
        }

        /// <summary>
        /// Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            if (this.FolderToDelete != null)
            {
                this.FOLDAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);

                // DeleteFolder request.
                DeleteFolderType deleteFolderRequest = this.GetDeleteFolderRequest(DisposalType.HardDelete, this.FolderToDelete);

                // Delete the specified folder.
                DeleteFolderResponseType deleteFolderResponse = this.FOLDAdapter.DeleteFolder(deleteFolderRequest);

                // Check the response.
                Common.CheckOperationSuccess(deleteFolderResponse, 1, this.Site);
            }

            base.TestCleanup();
        }
        #endregion

        #region Test case base methods
        /// <summary>
        /// Create a single calendar item.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="calendarItem">The calendar item to be created.</param>
        /// <param name="createOperation">An enumeration describes how meeting requests are handled.</param>
        /// <returns>If the method succeeds, returns an instance of ItemInfoResponseMessageType; otherwise, returns null.</returns>
        protected ItemInfoResponseMessageType CreateSingleCalendarItem(Role role, ItemType calendarItem, CalendarItemCreateOrDeleteOperationType createOperation)
        {
            ItemInfoResponseMessageType[] items = this.CreateMultipleCalendarItems(role, new ItemType[] { calendarItem }, createOperation);

            if (items != null && items.Length > 0)
            {
                return items[0];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Create multiple calendar items
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="calendarItems">The calendar items to be created.</param>
        /// <param name="createOperation">An enumeration describes how meeting requests are handled.</param>
        /// <returns>If the method succeeds, returns an array of ItemInfoResponseMessageType instances; otherwise, returns null.</returns>
        protected ItemInfoResponseMessageType[] CreateMultipleCalendarItems(Role role, ItemType[] calendarItems, CalendarItemCreateOrDeleteOperationType createOperation)
        {
            CreateItemType request = this.GetCreateItemType(calendarItems, createOperation);
            this.SwitchMTGSUser(role);
            CreateItemResponseType response = this.MTGSAdapter.CreateItem(request);
            Site.Assert.IsTrue(IsValidResponse(response), "The response messages returned by the CreateItem operation should not be null.");

            return GetItemInfoResponseMessageItems(response.ResponseMessages.Items);
        }

        /// <summary>
        /// Copy a single calendar item into a specified folder.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="itemId">The item id of the calendar item to be copied.</param>
        /// <param name="targetFolderId">The folder, to which the calendar item will be copied.</param>
        /// <returns>If the method succeeds, returns an instance of ItemInfoResponseMessageType; otherwise, returns null.</returns>
        protected ItemInfoResponseMessageType CopySingleCalendarItem(Role role, BaseItemIdType itemId, TargetFolderIdType targetFolderId)
        {
            ItemInfoResponseMessageType[] items = this.CopyMultipleCalendarItems(role, new BaseItemIdType[] { itemId }, targetFolderId);
            if (items != null && items.Length > 0)
            {
                return items[0];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Copy multiple calendar items into a specified folder.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="itemIds">The item ids of the calendar item to be copied.</param>
        /// <param name="targetFolderId">The folder, to which the calendar item will be copied.</param>
        /// <returns>If the method succeeds, returns an array of ItemInfoResponseMessageType instances; otherwise, returns null.</returns>
        protected ItemInfoResponseMessageType[] CopyMultipleCalendarItems(Role role, BaseItemIdType[] itemIds, TargetFolderIdType targetFolderId)
        {
            CopyItemType request = GetCopyItemType(itemIds, targetFolderId);
            this.SwitchMTGSUser(role);
            CopyItemResponseType response = this.MTGSAdapter.CopyItem(request);
            Site.Assert.IsTrue(IsValidResponse(response), "The response messages returned by the CopyItem operation should succeed.");
            
            foreach (ResponseMessageType responseMessage in response.ResponseMessages.Items)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1188");

                // Verify MS-OXWSMSG requirement: MS-OXWSMTGS_R1188    
                Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Success,
                    responseMessage.ResponseClass,
                    1188,
                    @"[In Messages] A successful CopyItem operation returns a CopyItemResponse element, as specified in [MS-OXWSCORE] section 3.1.4.1.2.2, with the ResponseClass attribute of the CopyItemResponseMessage element, as specified in [MS-OXWSCDATA] section 2.2.4.12, set to ""Success"".");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1189");

                // Verify MS-OXWSMSG requirement: MS-OXWSMTGS_R1189
                Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                    ResponseCodeType.NoError,
                    responseMessage.ResponseCode,
                    1189,
                    @"[In Messages] The ResponseCode element, as specified in [MS-OXWSCDATA] section 2.2.4.43, of the CopyItemResponseMessage element is set to ""NoError"".");
            }

            return GetItemInfoResponseMessageItems(response.ResponseMessages.Items);
        }

        /// <summary>
        /// Move a single calendar item into a specified folder.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="itemId">The item id of the calendar item to be moved.</param>
        /// <param name="targetFolderId">The folder, to which the calendar item will be moved.</param>
        /// <returns>If the method succeeds, returns an instance of ItemInfoResponseMessageType; otherwise, returns null.</returns>
        protected ItemInfoResponseMessageType MoveSingleCalendarItem(Role role, BaseItemIdType itemId, TargetFolderIdType targetFolderId)
        {
            ItemInfoResponseMessageType[] items = this.MoveMultipleCalendarItems(role, new BaseItemIdType[] { itemId }, targetFolderId);

            if (items != null && items.Length > 0)
            {
                return items[0];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Move multiple calendar items into a specified folder.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="itemIds">The item ids of the calendar items to be moved.</param>
        /// <param name="targetFolderId">The folder, to which the calendar item will be moved.</param>
        /// <returns>If the method succeeds, returns an array of ItemInfoResponseMessageType instances; otherwise, returns null.</returns>
        protected ItemInfoResponseMessageType[] MoveMultipleCalendarItems(Role role, BaseItemIdType[] itemIds, TargetFolderIdType targetFolderId)
        {
            MoveItemType request = GetMoveItemType(itemIds, targetFolderId);
            this.SwitchMTGSUser(role);
            MoveItemResponseType response = this.MTGSAdapter.MoveItem(request);
            Site.Assert.IsTrue(IsValidResponse(response), "The response messages returned by the MoveItem operation should succeed.");

            foreach (ResponseMessageType item in response.ResponseMessages.Items)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1226");

                // Verify MS-OXWSMSG requirement: MS-OXWSMTGS_R1226
                Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Success,
                    item.ResponseClass,
                    1226,
                    @"[In Messages] A successful MoveItem operation returns a MoveItemResponse element, as specified in [MS-OXWSCORE] section 3.1.4.7.2.2, with the ResponseClass attribute of the MoveItemResponseMessage element, as specified in [MS-OXWSCDATA] section 2.2.4.12, set to ""Success"". ");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1227");

                // Verify MS-OXWSMSG requirement: MS-OXWSMTGS_R1227
                Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                    ResponseCodeType.NoError,
                    item.ResponseCode,
                    1227,
                    @"[In Messages] The ResponseCode element, as specified in [MS-OXWSCDATA] section 2.2.4.43, of the MoveItemResponseMessage element is set to ""NoError"".");
            }

            return GetItemInfoResponseMessageItems(response.ResponseMessages.Items);
        }

        /// <summary>
        /// Update a single specified calendar item.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="itemChangeInfo">The update information.</param>
        /// <param name="updateOperation">An enumeration that describes how meeting updates are communicated after a calendar related item is updated.</param>
        /// <returns>If the method succeeds, returns an instance of UpdateItemResponseMessageType; otherwise, returns null.</returns>
        protected UpdateItemResponseMessageType UpdateSingleCalendarItem(Role role, AdapterHelper itemChangeInfo, CalendarItemUpdateOperationType updateOperation)
        {
            UpdateItemResponseMessageType[] items = this.UpdateMultipleCalendarItems(role, new AdapterHelper[] { itemChangeInfo }, updateOperation);

            if (items != null && items.Length > 0)
            {
                return items[0];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Update multiple specified calendar items.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="itemsChangeInfo">The update information.</param>
        /// <param name="updateOperation">An enumeration that describes how meeting updates are communicated after a calendar related item is updated.</param>
        /// <returns>If the method succeeds, returns an array of UpdateItemResponseMessageType instances; otherwise, returns null.</returns>
        protected UpdateItemResponseMessageType[] UpdateMultipleCalendarItems(Role role, AdapterHelper[] itemsChangeInfo, CalendarItemUpdateOperationType updateOperation)
        {
            UpdateItemType request = this.GetUpdateItemType(itemsChangeInfo, updateOperation);
            this.SwitchMTGSUser(role);
            UpdateItemResponseType response = this.MTGSAdapter.UpdateItem(request);
            Site.Assert.IsTrue(IsValidResponse(response), "The response messages returned by the UpdateItem operation should succeed.");

            List<UpdateItemResponseMessageType> items = new List<UpdateItemResponseMessageType>();
            foreach (ResponseMessageType item in response.ResponseMessages.Items)
            {
                if (ResponseClassType.Success == item.ResponseClass)
                {
                    items.Add(item as UpdateItemResponseMessageType);
                }
            }

            if (items.Count > 0)
            {
                return items.ToArray();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Delete a single calendar item.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="itemId">The item id of the calendar item to be deleted.</param>
        /// <param name="deleteOperation">An enumeration describes how meeting cancellations are handled.</param>
        /// <returns>If the method succeeds, returns an instance of ResponseMessageType; otherwise, returns null.</returns>
        protected ResponseMessageType DeleteSingleCalendarItem(Role role, BaseItemIdType itemId, CalendarItemCreateOrDeleteOperationType deleteOperation)
        {
            ResponseMessageType[] items = this.DeleteMultipleCalendarItems(role, new BaseItemIdType[] { itemId }, deleteOperation);

            if (items != null && items.Length > 0)
            {
                return items[0];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Delete multiple calendar items.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="itemIds">The item ids of the calendar items to be deleted.</param>
        /// <param name="deleteOperation">An enumeration describes how meeting cancellations are handled.</param>
        /// <returns>If the method succeeds, returns an array of ResponseMessageType instances; otherwise, returns null.</returns>
        protected ResponseMessageType[] DeleteMultipleCalendarItems(Role role, BaseItemIdType[] itemIds, CalendarItemCreateOrDeleteOperationType deleteOperation)
        {
            DeleteItemType request = GetDeleteItemType(itemIds, deleteOperation);
            this.SwitchMTGSUser(role);
            DeleteItemResponseType response = this.MTGSAdapter.DeleteItem(request);
            Site.Assert.IsTrue(IsValidResponse(response), "The invocation to DeleteItem operation should be successful.");

            List<ResponseMessageType> calendar = new List<ResponseMessageType>();
            foreach (ResponseMessageType item in response.ResponseMessages.Items)
            {
                if (ResponseClassType.Success == item.ResponseClass)
                {
                    calendar.Add(item);
                }
            }

            if (calendar.Count > 0)
            {
                return calendar.ToArray();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get a single calendar item.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="itemId">The item id of the calendar item to be gotten.</param>
        /// <returns>If the method succeeds, returns an instance of ItemInfoResponseMessageType; otherwise, returns null.</returns>
        protected ItemInfoResponseMessageType GetSingleCalendarItem(Role role, BaseItemIdType itemId)
        {
            ItemInfoResponseMessageType[] items = this.GetMultipleCalendarItems(role, new BaseItemIdType[] { itemId });

            if (items != null && items.Length > 0)
            {
                return items[0];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get multiple calendar items.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="itemIds">The item ids of the calendar items to be gotten.</param>
        /// <returns>If the method succeeds, returns an array of ItemInfoResponseMessageType instances; otherwise, returns null.</returns>
        protected ItemInfoResponseMessageType[] GetMultipleCalendarItems(Role role, BaseItemIdType[] itemIds)
        {
            GetItemType request = this.GetGetItemType(itemIds);
            this.SwitchMTGSUser(role);
            GetItemResponseType response = this.MTGSAdapter.GetItem(request);
            Site.Assert.IsTrue(IsValidResponse(response), "The invocation to GetItem operation should be successful.");

            if (Common.IsRequirementEnabled(8852, this.Site))
            {
                foreach (ResponseMessageType responseMsg in response.ResponseMessages.Items)
                {
                    if (responseMsg.ResponseClass == ResponseClassType.Success)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R8852");

                        // Verify MS-OXWSMTGS requirement: MS-OXWSCDATA_R8852
                        // calendar:ConflictingMeetingCount is included in the request and the operation executes successfully, this requirement can be captured.
                        this.Site.CaptureRequirement(
                            "MS-OXWSCDATA",
                            8852,
                            @"[In Appendix C: Product Behavior] Implementation does support value ""calendar:ConflictingMeetingCount"" specifies the ConflictingMeetingCount property. (Exchange 2010 and above follow this behavior.)");
                    }
                }
            }

            return GetItemInfoResponseMessageItems(response.ResponseMessages.Items);
        }

        /// <summary>
        /// Search a calendar item with a specified UID.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="folder">The folder to search.</param>
        /// <param name="value">The string to search.</param>
        /// <param name="uid">The UID, which the target calendar item contains.</param>
        /// <param name="fieldURI">The property to find.</param>
        /// <returns>If the operation succeeds, return an instance of ItemType; otherwise, return null.</returns>
        protected ItemType SearchSingleItem(Role role, DistinguishedFolderIdNameType folder, string value, string uid, UnindexedFieldURIType fieldURI = UnindexedFieldURIType.itemItemClass)
        {
            // Find items in a specified folder
            ItemIdType[] items = this.SearchItemIds(role, folder, value, fieldURI);
            if (items == null || items.Length == 0)
            {
                return null;
            }

            return this.GetSpecifiedItem(items, role, uid);
        }

        /// <summary>
        /// Search a calendar item with a specified UID, which should be deleted.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="folder">The folder to search.</param>
        /// <param name="value">The string to search.</param>
        /// <param name="uid">The UID, which the target calendar item contains.</param>
        /// <returns>If the item has been deleted successfully, return null; otherwise, return the specified item.</returns>
        protected ItemType SearchDeletedSingleItem(Role role, DistinguishedFolderIdNameType folder, string value, string uid)
        {
            FindItemResponseType findResponse = null;
            int counter = 0;

            FindItemType findRequest = GetFindItemType(folder, value);
            this.SwitchSRCHUser(role);

            while (counter < this.UpperBound)
            {
                System.Threading.Thread.Sleep(this.WaitTime);

                findResponse = this.SRCHAdapter.FindItem(findRequest);
                if (IsValidResponse(findResponse))
                {
                    foreach (ResponseMessageType item in findResponse.ResponseMessages.Items)
                    {
                        FindItemResponseMessageType findItem = item as FindItemResponseMessageType;
                        if (findItem != null && findItem.RootFolder.Item != null)
                        {
                            ArrayOfRealItemsType realItems = findItem.RootFolder.Item as ArrayOfRealItemsType;
                            if (realItems != null && (realItems.Items == null || realItems.Items.Length == 0))
                            {
                                return null;
                            }
                        }
                    }
                }

                counter++;
            }

            return this.GetSpecifiedItem(GetFindItemIds(findResponse), role, uid);
        }

        /// <summary>
        /// Switch the current role to the new one, with the identity of the new role to communicate with Exchange Web Service.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        protected void SwitchMTGSUser(Role role)
        {
            switch (role)
            {
                case Role.Organizer:
                    this.MTGSAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);
                    break;
                case Role.Attendee:
                    this.MTGSAdapter.SwitchUser(this.Attendee, this.AttendeePassword, this.Domain);
                    break;
                case Role.Delegate:
                    this.MTGSAdapter.SwitchUser(this.DelegateUser, this.DelegatePassword, this.Domain);
                    break;
            }
        }

        /// <summary>
        /// Switch the current role to the new one, with the identity of the new role to communicate with Exchange Web Service.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        protected void SwitchSRCHUser(Role role)
        {
            switch (role)
            {
                case Role.Organizer:
                    this.SRCHAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);
                    break;
                case Role.Attendee:
                    this.SRCHAdapter.SwitchUser(this.Attendee, this.AttendeePassword, this.Domain);
                    break;
                case Role.Delegate:
                    this.SRCHAdapter.SwitchUser(this.DelegateUser, this.DelegatePassword, this.Domain);
                    break;
            }
        }

        /// <summary>
        /// Clear one or more folders by a role.
        /// </summary>
        /// <param name="role">A role, of which the folders will be cleaned up.</param>
        /// <param name="folders">The folders to be cleaned up.</param>
        protected void CleanupFoldersByRole(Role role, List<DistinguishedFolderIdNameType> folders)
        {
            foreach (DistinguishedFolderIdNameType folder in folders)
            {
                if (folder == DistinguishedFolderIdNameType.calendar)
                {
                    this.CleanupFolder(role, DistinguishedFolderIdNameType.calendar);
                }

                if (folder == DistinguishedFolderIdNameType.drafts)
                {
                    this.CleanupFolder(role, DistinguishedFolderIdNameType.drafts);
                }

                if (folder == DistinguishedFolderIdNameType.sentitems)
                {
                    this.CleanupFolder(role, DistinguishedFolderIdNameType.sentitems);
                }

                if (folder == DistinguishedFolderIdNameType.inbox)
                {
                    this.CleanupFolder(role, DistinguishedFolderIdNameType.inbox);
                }

                if (folder  == DistinguishedFolderIdNameType.deleteditems)
                {
                    this.CleanupFolder(role, DistinguishedFolderIdNameType.deleteditems);
                }
            }
        }

        /// <summary>
        /// Generate the request message for operation "DeleteFolder".
        /// </summary>
        /// <param name="deleteType">How folders are to be deleted.</param>
        /// <param name="folderIds">An array of folder identifier of the folders need to be deleted</param>
        /// <returns>Delete folder request instance that will send to server.</returns>
        protected DeleteFolderType GetDeleteFolderRequest(DisposalType deleteType, params BaseFolderIdType[] folderIds)
        {
            Site.Assert.IsNotNull(folderIds, "Folders id should not be null!");
            Site.Assert.AreNotEqual<int>(0, folderIds.Length, "Folders id should contains at least one Id!");
            DeleteFolderType deleteFolderRequest = new DeleteFolderType();

            // Specify the delete type.
            deleteFolderRequest.DeleteType = deleteType;
            int folderCount = folderIds.Length;

            // Set the request's folderId field.
            deleteFolderRequest.FolderIds = new BaseFolderIdType[folderCount];
            for (int folderIdIndex = 0; folderIdIndex < folderCount; folderIdIndex++)
            {
                deleteFolderRequest.FolderIds[folderIdIndex] = folderIds[folderIdIndex];
            }

            return deleteFolderRequest;
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Verify the validity of a response returned by an operation.
        /// </summary>
        /// <param name="response">A response returned by an operation.</param>
        /// <returns>If the response is valid, return true; otherwise, false.</returns>
        private static bool IsValidResponse(BaseResponseMessageType response)
        {
            if (response != null
                && response.ResponseMessages != null
                && response.ResponseMessages.Items != null
                && response.ResponseMessages.Items.Length > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Get an array of items of ItemInfoResponseMessage type.
        /// </summary>
        /// <param name="items">The items returned in the response.</param>
        /// <returns>An array of items of ItemInfoResponseMessageType</returns>
        private static ItemInfoResponseMessageType[] GetItemInfoResponseMessageItems(ResponseMessageType[] items)
        {
            List<ItemInfoResponseMessageType> calendar = new List<ItemInfoResponseMessageType>();
            foreach (ResponseMessageType item in items)
            {
                if (ResponseClassType.Success == item.ResponseClass)
                {
                    calendar.Add(item as ItemInfoResponseMessageType);
                }
            }

            if (calendar.Count > 0)
            {
                return calendar.ToArray();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get items of ItemIdType type from a FindItem response message.
        /// </summary>
        /// <param name="findResponse">A response returned by FindItem operation.</param>
        /// <returns>An array of ItemId.</returns>
        private static ItemIdType[] GetFindItemIds(BaseResponseMessageType findResponse)
        {
            List<ItemIdType> listOfItemIds = null;

            if (findResponse != null
                && findResponse.ResponseMessages != null
                && findResponse.ResponseMessages.Items != null
                && findResponse.ResponseMessages.Items.Length > 0)
            {
                ArrayOfRealItemsType items = null;
                foreach (ResponseMessageType item in findResponse.ResponseMessages.Items)
                {
                    FindItemResponseMessageType resonseMessage = item as FindItemResponseMessageType;
                    if (resonseMessage != null)
                    {
                        items = resonseMessage.RootFolder.Item as ArrayOfRealItemsType;
                        break;
                    }
                }

                if (items != null && items.Items != null && items.Items.Length > 0)
                {
                    foreach (ItemType item in items.Items)
                    {
                        if (item.ItemId != null && !string.IsNullOrEmpty(item.ItemId.Id))
                        {
                            if (listOfItemIds == null)
                            {
                                listOfItemIds = new List<ItemIdType>();
                            }

                            listOfItemIds.Add(item.ItemId);
                        }
                    }
                }
            }

            if (listOfItemIds != null)
            {
                return listOfItemIds.ToArray();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get the argument of FindItemType type for FindItem operation.
        /// </summary>
        /// <param name="folderType">The type of folder to search.</param>
        /// <param name="value">A string that specifies the value for a search restriction.</param>
        /// <param name="fieldURI">The property to find.</param>
        /// <returns>An instance of FindItemType type.</returns>
        private static FindItemType GetFindItemType(DistinguishedFolderIdNameType folderType, string value, UnindexedFieldURIType fieldURI = UnindexedFieldURIType.itemItemClass)
        {
            FindItemType findRequest = new FindItemType();

            // Specify that only the items in the folder are returned.
            findRequest.Traversal = ItemQueryTraversalType.Shallow;

            #region Specify all properties to return in FindItem response.
            findRequest.ItemShape = new ItemResponseShapeType();
            findRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            #endregion

            #region Specify a folder that is the root of the search: inbox, calendar, deleteditems, drafts and sentitems.
            // Specify searching items in the distinguished folder represents by input parameter folder.
            DistinguishedFolderIdType disFolderId = new DistinguishedFolderIdType();
            disFolderId.Id = folderType;
            findRequest.ParentFolderIds = new BaseFolderIdType[] { disFolderId };
            #endregion

            #region Specifies a search restriction or query
            PathToUnindexedFieldType indexedField = new PathToUnindexedFieldType();
            indexedField.FieldURI = fieldURI;
            ContainsExpressionType expressionType = new ContainsExpressionType();
            expressionType.Item = indexedField;

            // Specify that the comparison is between the substring of the property value and the constant.
            expressionType.ContainmentMode = ContainmentModeType.Substring;

            // Indicate the ContainmentMode property is serialized in the SOAP message.
            expressionType.ContainmentModeSpecified = true;

            // Specify that the comparison ignores casing and non-spacing characters
            expressionType.ContainmentComparison = ContainmentComparisonType.IgnoreCaseAndNonSpacingCharacters;

            // Indicate the ContainmentComparison property is serialized in the SOAP message.
            expressionType.ContainmentComparisonSpecified = true;
            expressionType.Constant = new ConstantValueType();
            expressionType.Constant.Value = value;

            RestrictionType restriction = new RestrictionType();
            restriction.Item = expressionType;
            if (!string.IsNullOrEmpty(value))
            {
                findRequest.Restriction = restriction;
            }
            #endregion

            return findRequest;
        }

        /// <summary>
        /// Get the argument of CopyItemType type for CopyItem operation.
        /// </summary>
        /// <param name="itemIds">Array of identifiers that represent calendar related items.</param>
        /// <param name="toFolderId">Specify the folder that calendar related item to be copied.</param>
        /// <returns>An instance of CopyItemType type.</returns>
        private static CopyItemType GetCopyItemType(BaseItemIdType[] itemIds, TargetFolderIdType toFolderId)
        {
            if (itemIds != null)
            {
                CopyItemType copyItem = new CopyItemType();
                copyItem.ItemIds = itemIds;
                copyItem.ToFolderId = toFolderId;
                return copyItem;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get the argument of MoveItemType type for MoveItem operation.
        /// </summary>
        /// <param name="itemIds">Array of identifiers that represent calendar related items.</param>
        /// <param name="toFolderId">Specify the folder that calendar related item to be moved.</param>
        /// <returns>An instance of MoveItemType type.</returns>
        private static MoveItemType GetMoveItemType(BaseItemIdType[] itemIds, TargetFolderIdType toFolderId)
        {
            if (itemIds != null)
            {
                MoveItemType moveItem = new MoveItemType();
                moveItem.ItemIds = itemIds;
                moveItem.ToFolderId = toFolderId;
                return moveItem;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get the argument of DeleteItemType type for DeleteItem operation.
        /// </summary>
        /// <param name="itemIds">Array of identifiers that represent calendar related items.</param>
        /// <param name="deleteOperation">An enumeration describes how meeting cancellations are handled.</param>
        /// <returns>An instance of DeleteItemType type.</returns>
        private static DeleteItemType GetDeleteItemType(BaseItemIdType[] itemIds, CalendarItemCreateOrDeleteOperationType deleteOperation)
        {
            if (itemIds != null)
            {
                DeleteItemType deleteItem = new DeleteItemType();
                deleteItem.ItemIds = itemIds;
                deleteItem.DeleteType = DisposalType.HardDelete;
                deleteItem.SendMeetingCancellations = deleteOperation;

                // Indicate the SendMeetingCancellations property is serialized in the SOAP message.
                deleteItem.SendMeetingCancellationsSpecified = true;
                return deleteItem;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get the argument of CreateItemType type for CreateItem operation.
        /// </summary>
        /// <param name="items">Array of calendar related item to be created.</param>
        /// <param name="createOperation">An enumeration describes how meeting requests are handled.</param>
        /// <returns>An instance of CreateItemType type.</returns>
        private CreateItemType GetCreateItemType(ItemType[] items, CalendarItemCreateOrDeleteOperationType createOperation)
        {
            if (items != null)
            {
                NonEmptyArrayOfAllItemsType allItemArray = new NonEmptyArrayOfAllItemsType();
                allItemArray.Items = items;

                CreateItemType createItem = new CreateItemType();
                createItem.Items = allItemArray;

                // Indicate the MessageDisposition property is serialized in the SOAP message.
                createItem.MessageDispositionSpecified = true;
                createItem.MessageDisposition = this.messageDisposition;

                // Indicate the SendMeetingInvitations property is serialized in the SOAP message.
                createItem.SendMeetingInvitationsSpecified = true;
                createItem.SendMeetingInvitations = createOperation;

                return createItem;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get the argument of UpdateItemType type for UpdateItem operation.
        /// </summary>
        /// <param name="itemChanges">Contain the information to be changed for calendar related item.</param>
        /// <param name="updateOperation">An enumeration that describes how meeting updates are communicated after a calendar related item is updated.</param>
        /// <returns>An instance of UpdateItemType type.</returns>
        private UpdateItemType GetUpdateItemType(AdapterHelper[] itemChanges, CalendarItemUpdateOperationType updateOperation)
        {
            if (itemChanges != null)
            {
                UpdateItemType updateItem = new UpdateItemType();
                updateItem.ItemChanges = new ItemChangeType[itemChanges.Length];
                updateItem.SendMeetingInvitationsOrCancellations = updateOperation;

                // Indicate the SendMeetingInvitationsOrCancellations property is serialized in the SOAP message.
                updateItem.SendMeetingInvitationsOrCancellationsSpecified = true;

                // Indicate the MessageDisposition property is serialized in the SOAP message.
                updateItem.MessageDisposition = this.MessageDisposition;
                updateItem.MessageDispositionSpecified = true;

                for (int i = 0; i < itemChanges.Length; i++)
                {
                    PathToUnindexedFieldType pathToUnindexedField = new PathToUnindexedFieldType();
                    pathToUnindexedField.FieldURI = itemChanges[i].FieldURI;

                    SetItemFieldType setItemField = new SetItemFieldType();
                    setItemField.Item = pathToUnindexedField;
                    setItemField.Item1 = itemChanges[i].Item;

                    ItemChangeType itemChange = new ItemChangeType();
                    itemChange.Item = itemChanges[i].ItemId;
                    itemChange.Updates = new ItemChangeDescriptionType[] { setItemField };

                    updateItem.ItemChanges[i] = itemChange;
                }

                return updateItem;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get the argument of GetItemType type for GetItem operation.
        /// </summary>
        /// <param name="itemIds">Array of identifiers that represent calendar related items.</param>
        /// <returns>An instance of GetItemType type.</returns>
        private GetItemType GetGetItemType(BaseItemIdType[] itemIds)
        {
            if (itemIds != null)
            {
                GetItemType getItem = new GetItemType();
                getItem.ItemIds = itemIds;
                getItem.ItemShape = new ItemResponseShapeType();
                getItem.ItemShape.BaseShape = this.BaseShape;

                List<PathToUnindexedFieldType> additionalProperties = new List<PathToUnindexedFieldType>();
                if (Common.IsRequirementEnabled(718, this.Site))
                {
                    additionalProperties.Add(new PathToUnindexedFieldType() { FieldURI = UnindexedFieldURIType.calendarStartTimeZone });

                    if (Common.IsRequirementEnabled(710, this.Site))
                    {
                        additionalProperties.Add(new PathToUnindexedFieldType() { FieldURI = UnindexedFieldURIType.calendarStartTimeZoneId });
                    }
                }

                if (Common.IsRequirementEnabled(719, this.Site))
                {
                    additionalProperties.Add(new PathToUnindexedFieldType() { FieldURI = UnindexedFieldURIType.calendarEndTimeZone });

                    if (Common.IsRequirementEnabled(711, this.Site))
                    {
                        additionalProperties.Add(new PathToUnindexedFieldType() { FieldURI = UnindexedFieldURIType.calendarEndTimeZoneId });
                    }
                }

                if (Common.IsRequirementEnabled(696, this.Site)
                    || Common.IsRequirementEnabled(697, this.Site)
                    || Common.IsRequirementEnabled(707, this.Site)
                    || Common.IsRequirementEnabled(80011, this.Site))
                {
                    additionalProperties.Add(new PathToUnindexedFieldType() { FieldURI = UnindexedFieldURIType.calendarEnhancedLocation });
                }

                if (Common.IsRequirementEnabled(8852, this.Site))
                {
                    additionalProperties.Add(new PathToUnindexedFieldType() { FieldURI = UnindexedFieldURIType.calendarConflictingMeetingCount });
                    additionalProperties.Add(new PathToUnindexedFieldType() { FieldURI = UnindexedFieldURIType.calendarAdjacentMeetingCount });
                    additionalProperties.Add(new PathToUnindexedFieldType() { FieldURI = UnindexedFieldURIType.calendarConflictingMeetings });
                    additionalProperties.Add(new PathToUnindexedFieldType() { FieldURI = UnindexedFieldURIType.calendarAdjacentMeetings });
                }

                if (additionalProperties.Count > 0)
                {
                    getItem.ItemShape.AdditionalProperties = additionalProperties.ToArray();
                }

                return getItem;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get an item with a specified uid.
        /// </summary>
        /// <param name="items">The items of ItemIdType type in some folder, which should be one of folders: inbox, deleteditems, sentitems, calendar, and drafts.</param>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="uid">The UID, which the target item contains.</param>
        /// <returns>If the operation succeeds, the item that has the same uid specified; otherwise, null.</returns>
        private ItemType GetSpecifiedItem(ItemIdType[] items, Role role, string uid)
        {
            if (items == null)
            {
                return null;
            }

            GetItemType getItemRequest = this.GetGetItemType(items);
            this.SwitchMTGSUser(role);
            GetItemResponseType getResponse = this.MTGSAdapter.GetItem(getItemRequest);

            if (!IsValidResponse(getResponse))
            {
                Site.Assert.Fail("The result of GetItem operation should not be null or empty.");
            }

            ItemInfoResponseMessageType[] infoItems = new ItemInfoResponseMessageType[getResponse.ResponseMessages.Items.Length];
            getResponse.ResponseMessages.Items.CopyTo(infoItems, 0);

            foreach (ItemInfoResponseMessageType item in infoItems)
            {
                if (item == null || item.Items == null || item.Items.Items == null)
                {
                    continue;
                }

                if (item.Items.Items[0] is CalendarItemType)
                {
                    CalendarItemType calendar = item.Items.Items[0] as CalendarItemType;
                    if (calendar.UID == uid)
                    {
                        return calendar;
                    }
                }
                else if (item.Items.Items[0] is MeetingCancellationMessageType)
                {
                    MeetingCancellationMessageType meetingCancellationMessage = item.Items.Items[0] as MeetingCancellationMessageType;
                    if (meetingCancellationMessage.UID == uid)
                    {
                        return meetingCancellationMessage;
                    }
                }
                else if (item.Items.Items[0] is MeetingRequestMessageType)
                {
                    MeetingRequestMessageType meetingRequestMessage = item.Items.Items[0] as MeetingRequestMessageType;
                    if (meetingRequestMessage.UID == uid)
                    {
                        return meetingRequestMessage;
                    }
                }
                else if (item.Items.Items[0] is MeetingResponseMessageType)
                {
                    MeetingResponseMessageType meetingResponseMessage = item.Items.Items[0] as MeetingResponseMessageType;
                    if (meetingResponseMessage.UID == uid)
                    {
                        return meetingResponseMessage;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Search all ItemIds of the specified items.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="folder">The folder to search.</param>
        /// <param name="value">The string to search.</param>
        /// <param name="fieldURI">The property to find.</param>
        /// <returns>If the operation succeeds, return all ItemIds of the items found in the specified folder; otherwise, return null.</returns>
        private ItemIdType[] SearchItemIds(Role role, DistinguishedFolderIdNameType folder, string value, UnindexedFieldURIType fieldURI = UnindexedFieldURIType.itemItemClass)
        {
            FindItemResponseType findItemResponse = this.SearchItems(role, folder, value, fieldURI);

            if (findItemResponse != null)
            {
                return GetFindItemIds(findItemResponse);
            } 
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Search all ItemIds of the specified items.
        /// </summary>
        /// <param name="role">The role used to communicate with Exchange Web Service.</param>
        /// <param name="folder">The folder to search.</param>
        /// <param name="value">The string to search.</param>
        /// <param name="fieldURI">The property to find.</param>
        /// <returns>If the operation succeeds, return a response of FindItemResponseType; otherwise, return null.</returns>
        private FindItemResponseType SearchItems(Role role, DistinguishedFolderIdNameType folder, string value, UnindexedFieldURIType fieldURI = UnindexedFieldURIType.itemItemClass)
        {
            FindItemResponseType findResponse = null;
            int counter = 0;
            bool isValid = true;

            // Find the meeting request message in Inbox folder of the attendeeType.
            FindItemType findRequest = GetFindItemType(folder, value, fieldURI);
            this.SwitchSRCHUser(role);
            while (counter < this.UpperBound)
            {
                System.Threading.Thread.Sleep(this.WaitTime);

                if (!isValid)
                {
                    isValid = true;
                }
                
                findResponse = this.SRCHAdapter.FindItem(findRequest);
                if (IsValidResponse(findResponse))
                {
                    foreach (ResponseMessageType item in findResponse.ResponseMessages.Items)
                    {
                        FindItemResponseMessageType findItem = item as FindItemResponseMessageType;
                        if (findItem == null || findItem.RootFolder.Item == null)
                        {
                            isValid = false;
                            break;
                        }
                        else
                        {
                            ArrayOfRealItemsType realItems = findItem.RootFolder.Item as ArrayOfRealItemsType;
                            if (realItems != null && realItems.Items == null)
                            {
                                isValid = false;
                                break;
                            }
                        }
                    }

                    if (isValid)
                    {
                        break;
                    }
                }

                counter++;
            }

            Site.Log.Add(LogEntryKind.Debug, string.Format("The retry count of FindItem operation is {0}", counter));

            if (counter == this.UpperBound && !isValid)
            {
                return null;
            }
            else
            {
                return findResponse;
            }
        }

        /// <summary>
        /// Clear a specified folder of a specified role.
        /// </summary>
        /// <param name="role">A role, of which the folders will be cleaned up.</param>
        /// <param name="folder">The folder to be cleaned up.</param>
        private void CleanupFolder(Role role, DistinguishedFolderIdNameType folder)
        {
            ItemIdType[] itemIdsInCalendar = this.SearchItemIds(role, folder, string.Empty);
            if (itemIdsInCalendar != null && itemIdsInCalendar.Length > 0)
            {
                DeleteItemType deleteItemType = GetDeleteItemType(itemIdsInCalendar, CalendarItemCreateOrDeleteOperationType.SendToNone);
                this.SwitchMTGSUser(role);
                this.MTGSAdapter.DeleteItem(deleteItemType);
            }
        }

        #endregion
    }
}