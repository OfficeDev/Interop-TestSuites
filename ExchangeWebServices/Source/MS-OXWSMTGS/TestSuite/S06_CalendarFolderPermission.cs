namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test types related to calendar folder permission.
    /// </summary>
    [TestClass]
    public class S06_CalendarFolderPermission : TestSuiteBase
    {
        #region Class initialize and clean up
        /// <summary>
        /// Initialize the test class.
        /// </summary>
        /// <param name="context">Context to initialize.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            TestClassBase.Initialize(context);
        }

        /// <summary>
        /// Clean up the test class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test cases
        /// <summary>
        /// This test case is designed to test user can create, read, edit, and delete all items in the folder, and create subfolders
        /// if granted Owner permission to him.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S06_TC01_CalendarPermissionLevelOwner()
        {
            this.MTGSAdapter = new MS_OXWSMTGSAdapter();
            this.MTGSAdapter.Initialize(this.Site);
            this.SRCHAdapter = new MS_OXWSSRCHAdapter();
            this.SRCHAdapter.Initialize(this.Site);
            this.FOLDAdapter = new MS_OXWSFOLDAdapter();
            this.FOLDAdapter.Initialize(this.Site);

            #region Organizer creates a calendar folder in the inbox folder and grant Owner permission to attendee.

            // Switch to user organizer.
            this.FOLDAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);

            CalendarPermissionSetType calendarPermissionSet = new CalendarPermissionSetType();
            calendarPermissionSet.CalendarPermissions = new CalendarPermissionType[1];
            calendarPermissionSet.CalendarPermissions[0] = new CalendarPermissionType();
            calendarPermissionSet.CalendarPermissions[0].CalendarPermissionLevel = CalendarPermissionLevelType.Owner;
            calendarPermissionSet.CalendarPermissions[0].UserId = new UserIdType();
            calendarPermissionSet.CalendarPermissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("AttendeeName", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "CustomFolderByOrganizer" }, new string[] { "IPF.Appointment" }, null);
            ((CalendarFolderType)createFolderRequest.Folders[0]).PermissionSet = calendarPermissionSet;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderIdByOrganizer = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.FolderToDelete = newFolderIdByOrganizer;
            #endregion

            #region Organizer gets the new created folder.

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderIdByOrganizer);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType foldersResponseInfo = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            Site.Assert.AreEqual<string>(newFolderIdByOrganizer.Id, foldersResponseInfo.Folders[0].FolderId.Id, "The calendar folder should be created successfully in inbox.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R507");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R507
            // The element with type CalendarFolderType is returned and pass the schema validation, this requirement can be captured.
            this.Site.CaptureRequirementIfIsInstanceOfType(
                foldersResponseInfo.Folders[0],
                typeof(CalendarFolderType),
                507,
                @"[In t:CalendarFolderType Complex Type] This complex type extends the BaseFolderType complex type, as specified in [MS-OXWSFOLD] section 2.2.4.6.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R143");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R143
            // The element with type CalendarFolderType is returned and pass the schema validation, this requirement can be captured.
            this.Site.CaptureRequirementIfIsInstanceOfType(
                foldersResponseInfo.Folders[0],
                typeof(CalendarFolderType),
                143,
                @"[In t:CalendarFolderType Complex Type] [its schema is] <xs:complexType name=""CalendarFolderType"">
                  <xs:complexContent>
                    <xs:extension
                      base=""t:BaseFolderType""
                      >
                    <xs:sequence>
                      <xs:element name=""SharingEffectiveRights""
                      type=""t:CalendarPermissionReadAccessType""
                      minOccurs=""0""
                    />
                    <xs:element name=""PermissionSet""
                      type=""t:CalendarPermissionSetType""
                      minOccurs=""0""
                    />
                  </xs:sequence>
                </xs:extension>
              </xs:complexContent>
            </xs:complexType>");

            CalendarFolderType folderInfo = (CalendarFolderType)foldersResponseInfo.Folders[0];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R237");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R237
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.PermissionSet,
                237,
                @"[In t:CalendarPermissionSetType Complex Type] [its schema is] <xs:complexType name=""CalendarPermissionSetType"">
                  <xs:sequence>
                    <xs:element name=""CalendarPermissions""
                      type=""t:ArrayOfCalendarPermissionsType""
                     />
                    <xs:element name=""UnknownEntries""
                      type=""t:ArrayOfUnknownEntriesType""
                      minOccurs=""0""
                     />
                  </xs:sequence>
                </xs:complexType>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R146");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R146
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.PermissionSet,
                146,
                @"[In t:CalendarFolderType Complex Type] The type of PermissionSet is t:CalendarPermissionSetType (section 2.2.4.5).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R147");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R147
            this.Site.CaptureRequirementIfAreEqual<CalendarPermissionLevelType>(
                CalendarPermissionLevelType.Owner,
                folderInfo.PermissionSet.CalendarPermissions[2].CalendarPermissionLevel,
                147,
                @"[In t:CalendarFolderType Complex Type]PermissionSet: Specifies all permissions that are configured for a Calendar folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R534");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R534
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.PermissionSet.CalendarPermissions,
                534,
                @"[In t:CalendarPermissionType Complex Type] This complex type extends the BasePermissionType complex type, as specified in [MS-OXWSFOLD] section 2.2.4.3.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R238");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R238
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.PermissionSet.CalendarPermissions,
                238,
                @"[In t:CalendarPermissionSetType Complex Type] The type of CalendarPermissions is t:ArrayofCalendarPermissionsType ([MS-OXWSCDATA] section 2.2.4.4).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R243");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R243
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.PermissionSet.CalendarPermissions,
                243,
                @"[In t:CalendarPermissionType Complex Type] [its schema is] <xs:complexType name=""CalendarPermissionType"">
                  <xs:complexContent>
                    <xs:extension
                      base=""t:BasePermissionType""
                    >
                      <xs:sequence>
                        <xs:element name=""ReadItems""
                          type=""t:CalendarPermissionReadAccessType""
                          minOccurs=""0""
                          maxOccurs=""1""
                         />
                        <xs:element name=""CalendarPermissionLevel""
                          type=""t:CalendarPermissionLevelType""
                          minOccurs=""1""
                          maxOccurs=""1""
                         />
                      </xs:sequence>
                    </xs:extension>
                  </xs:complexContent>
                </xs:complexType>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R244");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R244
            this.Site.CaptureRequirementIfIsTrue(
                folderInfo.PermissionSet.CalendarPermissions[2].ReadItemsSpecified,
                244,
                @"[In t:CalendarPermissionType Complex Type] The type of ReadItems is t:CalendarPermissionReadAccessType (section 2.2.5.5).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R245");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R245
            // As organizer grants Owner to attendee, so ReadItems should be FullDetails.
            this.Site.CaptureRequirementIfAreEqual<CalendarPermissionReadAccessType>(
                CalendarPermissionReadAccessType.FullDetails,
                folderInfo.PermissionSet.CalendarPermissions[2].ReadItems,
                245,
                @"[In t:CalendarPermissionType Complex Type] ReadItems: Specifies the permissions a user has to view items in a Calendar folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R246");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R246
            this.Site.CaptureRequirementIfAreEqual<CalendarPermissionLevelType>(
                CalendarPermissionLevelType.Owner,
                folderInfo.PermissionSet.CalendarPermissions[2].CalendarPermissionLevel,
                246,
                @"[In t:CalendarPermissionType Complex Type] The type of CalendarPermissionLevel is t:CalendarPermissionLevelType (section 2.2.5.4).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R247");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R247
            this.Site.CaptureRequirementIfAreEqual<CalendarPermissionLevelType>(
                CalendarPermissionLevelType.Owner,
                folderInfo.PermissionSet.CalendarPermissions[2].CalendarPermissionLevel,
                247,
                @"[In t:CalendarPermissionType Complex Type]CalendarPermissionLevel: Specifies the permission level that a user has on a Calendar folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R248");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R248
            this.Site.CaptureRequirementIfAreEqual<CalendarPermissionLevelType>(
                CalendarPermissionLevelType.Owner,
                folderInfo.PermissionSet.CalendarPermissions[2].CalendarPermissionLevel,
                248,
                @"[In t:CalendarPermissionType Complex Type] CalendarPermissionLevel: This element MUST exist if calendar permissions are set on a Calendar folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R97");

            // Verify MS-OXWSFOLD requirement: MS-OXWSMTGS_R97
            this.Site.CaptureRequirementIfIsTrue(
                folderInfo.PermissionSet.CalendarPermissions[2].IsFolderContactSpecified && folderInfo.PermissionSet.CalendarPermissions[2].IsFolderContact,
                97,
                @"[In t:CalendarPermissionLevelType Simple Type] Owner: The user is both folder owner and folder contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R239");

            // Verify MS-OXWSFOLD requirement: MS-OXWSMTGS_R239
            // This requirement can be captured directly after the verifications in above.
            this.Site.CaptureRequirement(
                239,
                @"[In t:CalendarPermissionSetType Complex Type] CalendarPermissions: Contains an array of calendar permissions for a folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R86");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R86
            // Element in type CalendarPermissionLevelType is returned, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                86,
                @"[In t:CalendarPermissionLevelType Simple Type] [its schema is] <xs:simpleType name=""CalendarPermissionLevelType"">
                  <xs:restriction
                    base=""xs:string""
                  >
                    <xs:enumeration
                      value=""Author""
                     />
                    <xs:enumeration
                      value=""Contributor""
                     />
                    <xs:enumeration
                      value=""Custom""
                     />
                    <xs:enumeration
                      value=""Editor""
                     />
                    <xs:enumeration
                      value=""FreeBusyTimeAndSubjectAndLocation""
                     />
                    <xs:enumeration
                      value=""FreeBusyTimeOnly""
                     />
                    <xs:enumeration
                      value=""None""
                     />
                    <xs:enumeration
                      value=""NoneditingAuthor""
                     />
                    <xs:enumeration
                      value=""Owner""
                     />
                    <xs:enumeration
                      value=""PublishingAuthor""
                     />
                    <xs:enumeration
                      value=""PublishingEditor""
                     />
                    <xs:enumeration
                      value=""Reviewer""
                     />
                  </xs:restriction>
                </xs:simpleType>");
            #endregion

            #region Organizer creates a calendar item in the created folder.
            CalendarItemType calendarItembyOrganizer = new CalendarItemType();
            calendarItembyOrganizer.UID = Guid.NewGuid().ToString();
            calendarItembyOrganizer.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByOrganizer");
            calendarItembyOrganizer.LegacyFreeBusyStatus = LegacyFreeBusyType.Busy;
            calendarItembyOrganizer.LegacyFreeBusyStatusSpecified = true;
            calendarItembyOrganizer.Location = this.Location;
            calendarItembyOrganizer.Body = new BodyType();
            calendarItembyOrganizer.Body.BodyType1 = BodyTypeType.Text;
            calendarItembyOrganizer.Body.Value = Common.GenerateResourceName(this.Site, "CalendarBody");
            NonEmptyArrayOfAllItemsType allItemArray = new NonEmptyArrayOfAllItemsType();
            allItemArray.Items = new ItemType[] { calendarItembyOrganizer };
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = allItemArray;
            createItemRequest.SendMeetingInvitationsSpecified = true;
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            createItemRequest.SavedItemFolderId.Item = newFolderIdByOrganizer;
            CreateItemResponseType response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            ItemInfoResponseMessageType itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            Site.Assert.IsNotNull(itemInfoResponse, "The calendar item should be created successfully.");
            ItemIdType itemIdByOrganizer = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee creates a folder in the created folder.

            // Switch to user attendee.
            this.FOLDAdapter.SwitchUser(this.Attendee, this.AttendeePassword, this.Domain);

            createFolderRequest = this.GetCreateFolderRequest(newFolderIdByOrganizer.Id, new string[] { "CustomFolderByAttendee" }, new string[] { "IPF.Appointment" }, null);
            createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            bool canCreateSubFolder = ResponseClassType.Success == createFolderResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Attendee creates a calendar item in the created folder.
            this.SwitchMTGSUser(Role.Attendee);
            CalendarItemType calendarItembyAttendee = new CalendarItemType();
            calendarItembyAttendee.UID = Guid.NewGuid().ToString();
            calendarItembyAttendee.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByAttendee");
            createItemRequest.Items.Items[0] = calendarItembyAttendee;
            response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            bool canCreateOwnItem = null != itemInfoResponse;
            Site.Assert.IsTrue(canCreateOwnItem, "The calendar item should be created successfully.");
            ItemIdType itemIdByAttendee = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee gets the calendar item created by hisself.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByAttendee);
            bool canReadOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by hisself.
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByAttendee;

            // Update the created calendar item.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by hisself.
            ResponseMessageType responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByAttendee, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            #region Attendee gets the calendar item created by organizer.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByOrganizer);
            bool canReadNotOwnItem = null != itemInfoResponse;

            if (canReadNotOwnItem)
            {
                CalendarItemType temp = itemInfoResponse.Items.Items[0] as CalendarItemType;
                Site.Assert.AreEqual<string>(calendarItembyOrganizer.Subject, temp.Subject, "User should have permission to view subject of appointment.");
                Site.Assert.AreEqual<string>(calendarItembyOrganizer.Location, temp.Location, "User should have permission to view location of appointment.");
                Site.Assert.AreEqual<bool>(calendarItembyOrganizer.LegacyFreeBusyStatusSpecified, temp.LegacyFreeBusyStatusSpecified, "User should have permission to view free/busy time of appointment.");
                Site.Assert.AreEqual<LegacyFreeBusyType>(calendarItembyOrganizer.LegacyFreeBusyStatus, temp.LegacyFreeBusyStatus, "User should have permission to view free/busy time of appointment.");
                Site.Assert.IsNotNull(temp.Body, "User should have permission to view body of appointment.");
                Site.Assert.AreEqual<BodyTypeType>(calendarItembyOrganizer.Body.BodyType1, temp.Body.BodyType1, "User should have permission to view body of appointment.");
                Site.Assert.AreEqual<string>(calendarItembyOrganizer.Body.Value, temp.Body.Value, "User should have permission to view body of appointment.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R107, the value of ReadItems is {0}", folderInfo.PermissionSet.CalendarPermissions[2].ReadItems);

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R107
                this.Site.CaptureRequirementIfIsTrue(
                    folderInfo.PermissionSet.CalendarPermissions[2].ReadItemsSpecified && folderInfo.PermissionSet.CalendarPermissions[2].ReadItems == CalendarPermissionReadAccessType.FullDetails,
                    107,
                    @"[In t:CalendarPermissionReadAccessType Simple Type] FullDetails: The user has permission to view all items in the calendar, including free/busy time and subject, location, and details of appointments.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R102");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R102
                this.Site.CaptureRequirement(
                    102,
                    @"[In t:CalendarPermissionReadAccessType Simple Type] [its schema is] <xs:simpleType name=""CalendarPermissionReadAccessType"">
                      <xs:restriction
                        base=""xs:string""                    
                      >
                        <xs:enumeration
                          value=""None""
                         />
                        <xs:enumeration
                          value=""TimeOnly""
                         />
                        <xs:enumeration
                          value=""TimeAndSubjectAndLocation""
                         />
                        <xs:enumeration
                          value=""FullDetails""
                         />
                      </xs:restriction>
                    </xs:simpleType>");
            }
            #endregion

            #region Attendee updates the calendar item created by organizer.
            calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByOrganizer;

            // Update the created calendar item.
            updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateNotOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by organizer.
            responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByOrganizer, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteNotOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R96");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R96
            bool isVerifiedR96 = canCreateOwnItem && canReadOwnItem && canUpdateOwnItem && canDeleteOwnItem && canReadNotOwnItem
                && canUpdateNotOwnItem && canDeleteNotOwnItem && canCreateSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR96,
                "Can create own item is {0};\n" +
                "Can read own item is {1};\n" +
                "Can edit own item is {2};\n " +
                "Can delete own item is {3};\n" +
                "Can read not own item is {4};\n" +
                "Can edit not own item is {5};\n" +
                "Can delete not own item is {6};\n" +
                "Can create subfolders is {7};\n",
                canCreateOwnItem,
                canReadOwnItem,
                canUpdateOwnItem,
                canDeleteOwnItem,
                canReadNotOwnItem,
                canUpdateNotOwnItem,
                canDeleteNotOwnItem,
                canCreateSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR96,
                96,
                @"[In t:CalendarPermissionLevelType Simple Type] Owner: The user can create, read, edit, and delete all items in the folder, and create subfolders.");
        }

        /// <summary>
        /// This test case is designed to test user can create and read all items in the folder, and edit and delete only items that the user
        /// creates if granted Author permission to him.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S06_TC02_CalendarPermissionLevelAuthor()
        {
            this.MTGSAdapter = new MS_OXWSMTGSAdapter();
            this.MTGSAdapter.Initialize(this.Site);
            this.SRCHAdapter = new MS_OXWSSRCHAdapter();
            this.SRCHAdapter.Initialize(this.Site);
            this.FOLDAdapter = new MS_OXWSFOLDAdapter();
            this.FOLDAdapter.Initialize(this.Site);

            #region Organizer creates a calendar folder in the inbox folder and grant Author permission to attendee.

            // Switch to user organizer.
            this.FOLDAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);

            CalendarPermissionSetType calendarPermissionSet = new CalendarPermissionSetType();
            calendarPermissionSet.CalendarPermissions = new CalendarPermissionType[1];
            calendarPermissionSet.CalendarPermissions[0] = new CalendarPermissionType();
            calendarPermissionSet.CalendarPermissions[0].CalendarPermissionLevel = CalendarPermissionLevelType.Author;
            calendarPermissionSet.CalendarPermissions[0].UserId = new UserIdType();
            calendarPermissionSet.CalendarPermissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("AttendeeName", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "CustomFolderByOrganizer" }, new string[] { "IPF.Appointment" }, null);
            ((CalendarFolderType)createFolderRequest.Folders[0]).PermissionSet = calendarPermissionSet;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderIdByOrganizer = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.FolderToDelete = newFolderIdByOrganizer;
            #endregion

            #region Organizer gets the new created folder.

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderIdByOrganizer);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType foldersResponseInfo = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            Site.Assert.AreEqual<string>(newFolderIdByOrganizer.Id, foldersResponseInfo.Folders[0].FolderId.Id, "The calendar folder should be created successfully in inbox.");
            #endregion

            #region Organizer creates a calendar item in the created folder.
            CalendarItemType calendarItembyOrganizer = new CalendarItemType();
            calendarItembyOrganizer.UID = Guid.NewGuid().ToString();
            calendarItembyOrganizer.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByOrganizer");
            NonEmptyArrayOfAllItemsType allItemArray = new NonEmptyArrayOfAllItemsType();
            allItemArray.Items = new ItemType[] { calendarItembyOrganizer };
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = allItemArray;
            createItemRequest.SendMeetingInvitationsSpecified = true;
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            createItemRequest.SavedItemFolderId.Item = newFolderIdByOrganizer;
            CreateItemResponseType response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            ItemInfoResponseMessageType itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            Site.Assert.IsNotNull(itemInfoResponse, "The calendar item should be created successfully.");
            ItemIdType itemIdByOrganizer = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee creates a folder in the created folder.

            // Switch to user attendee.
            this.FOLDAdapter.SwitchUser(this.Attendee, this.AttendeePassword, this.Domain);

            createFolderRequest = this.GetCreateFolderRequest(newFolderIdByOrganizer.Id, new string[] { "CustomFolderByAttendee" }, new string[] { "IPF.Appointment" }, null);
            createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            bool canCreateSubFolder = ResponseClassType.Success == createFolderResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Attendee creates a calendar item in the created folder.
            this.SwitchMTGSUser(Role.Attendee);
            CalendarItemType calendarItembyAttendee = new CalendarItemType();
            calendarItembyAttendee.UID = Guid.NewGuid().ToString();
            calendarItembyAttendee.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByAttendee");
            createItemRequest.Items.Items[0] = calendarItembyAttendee;
            response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            bool canCreateOwnItem = null != itemInfoResponse;
            Site.Assert.IsTrue(canCreateOwnItem, "The calendar item should be created successfully.");
            ItemIdType itemIdByAttendee = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee gets the calendar item created by hisself.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByAttendee);
            bool canReadOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by hisself.
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByAttendee;

            // Update the created calendar item.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by hisself.
            ResponseMessageType responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByAttendee, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            #region Attendee gets the calendar item created by organizer.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByOrganizer);
            bool canReadNotOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by organizer.
            calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByOrganizer;

            // Update the created calendar item.
            updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateNotOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by organizer.
            responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByOrganizer, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteNotOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R88");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R88
            bool isVerifiedR88 = canCreateOwnItem && canReadOwnItem && canUpdateOwnItem && canDeleteOwnItem && canReadNotOwnItem
                && !canUpdateNotOwnItem && !canDeleteNotOwnItem && !canCreateSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR88,
                "Can create own item is {0};\n" +
                "Can read own item is {1};\n" +
                "Can edit own item is {2};\n " +
                "Can delete own item is {3};\n" +
                "Can read not own item is {4};\n" +
                "Can edit not own item is {5};\n" +
                "Can delete not own item is {6};\n" +
                "Can create subfolders is {7};\n",
                canCreateOwnItem,
                canReadOwnItem,
                canUpdateOwnItem,
                canDeleteOwnItem,
                canReadNotOwnItem,
                canUpdateNotOwnItem,
                canDeleteNotOwnItem,
                canCreateSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR88,
                88,
                @"[In t:CalendarPermissionLevelType Simple Type] Author: The user can create and read all items in the folder, and edit and delete only items that the user creates.");
        }

        /// <summary>
        /// This test case is designed to test user can create items in the folder and the contents of the folder do not appear
        /// if granted Contributor permission to him.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S06_TC03_CalendarPermissionLevelContributor()
        {
            this.MTGSAdapter = new MS_OXWSMTGSAdapter();
            this.MTGSAdapter.Initialize(this.Site);
            this.SRCHAdapter = new MS_OXWSSRCHAdapter();
            this.SRCHAdapter.Initialize(this.Site);
            this.FOLDAdapter = new MS_OXWSFOLDAdapter();
            this.FOLDAdapter.Initialize(this.Site);

            #region Organizer creates a calendar folder in the inbox folder and grant Contributor permission to attendee.

            // Switch to user organizer.
            this.FOLDAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);

            CalendarPermissionSetType calendarPermissionSet = new CalendarPermissionSetType();
            calendarPermissionSet.CalendarPermissions = new CalendarPermissionType[1];
            calendarPermissionSet.CalendarPermissions[0] = new CalendarPermissionType();
            calendarPermissionSet.CalendarPermissions[0].CalendarPermissionLevel = CalendarPermissionLevelType.Contributor;
            calendarPermissionSet.CalendarPermissions[0].UserId = new UserIdType();
            calendarPermissionSet.CalendarPermissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("AttendeeName", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "CustomFolderByOrganizer" }, new string[] { "IPF.Appointment" }, null);
            ((CalendarFolderType)createFolderRequest.Folders[0]).PermissionSet = calendarPermissionSet;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderIdByOrganizer = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.FolderToDelete = newFolderIdByOrganizer;
            #endregion

            #region Organizer gets the new created folder.

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderIdByOrganizer);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType foldersResponseInfo = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            Site.Assert.AreEqual<string>(newFolderIdByOrganizer.Id, foldersResponseInfo.Folders[0].FolderId.Id, "The calendar folder should be created successfully in inbox.");
            #endregion

            #region Organizer creates a calendar item in the created folder.
            CalendarItemType calendarItembyOrganizer = new CalendarItemType();
            calendarItembyOrganizer.UID = Guid.NewGuid().ToString();
            calendarItembyOrganizer.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByOrganizer");
            NonEmptyArrayOfAllItemsType allItemArray = new NonEmptyArrayOfAllItemsType();
            allItemArray.Items = new ItemType[] { calendarItembyOrganizer };
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = allItemArray;
            createItemRequest.SendMeetingInvitationsSpecified = true;
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            createItemRequest.SavedItemFolderId.Item = newFolderIdByOrganizer;
            CreateItemResponseType response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            ItemInfoResponseMessageType itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            Site.Assert.IsNotNull(itemInfoResponse, "The calendar item should be created successfully.");
            ItemIdType itemIdByOrganizer = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee creates a folder in the created folder.

            // Switch to user attendee.
            this.FOLDAdapter.SwitchUser(this.Attendee, this.AttendeePassword, this.Domain);

            createFolderRequest = this.GetCreateFolderRequest(newFolderIdByOrganizer.Id, new string[] { "CustomFolderByAttendee" }, new string[] { "IPF.Appointment" }, null);
            createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            bool canCreateSubFolder = ResponseClassType.Success == createFolderResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Attendee creates a calendar item in the created folder.
            this.SwitchMTGSUser(Role.Attendee);
            CalendarItemType calendarItembyAttendee = new CalendarItemType();
            calendarItembyAttendee.UID = Guid.NewGuid().ToString();
            calendarItembyAttendee.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByAttendee");
            createItemRequest.Items.Items[0] = calendarItembyAttendee;
            response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            bool canCreateOwnItem = null != itemInfoResponse;
            Site.Assert.IsTrue(canCreateOwnItem, "The calendar item should be created successfully.");
            ItemIdType itemIdByAttendee = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee gets the calendar item created by hisself.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByAttendee);
            bool canReadOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by hisself.
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByAttendee;

            // Update the created calendar item.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by hisself.
            ResponseMessageType responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByAttendee, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            #region Attendee gets the calendar item created by organizer.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByOrganizer);
            bool canReadNotOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by organizer.
            calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByOrganizer;

            // Update the created calendar item.
            updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateNotOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by organizer.
            responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByOrganizer, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteNotOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R486");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R486
            bool isVerifiedR486 = canCreateOwnItem && !canReadOwnItem && !canUpdateOwnItem && !canDeleteOwnItem && !canReadNotOwnItem
                && !canUpdateNotOwnItem && !canDeleteNotOwnItem && !canCreateSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR486,
                "Can create own item is {0};\n" +
                "Can read own item is {1};\n" +
                "Can edit own item is {2};\n " +
                "Can delete own item is {3};\n" +
                "Can read not own item is {4};\n" +
                "Can edit not own item is {5};\n" +
                "Can delete not own item is {6};\n" +
                "Can create subfolders is {7};\n",
                canCreateOwnItem,
                canReadOwnItem,
                canUpdateOwnItem,
                canDeleteOwnItem,
                canReadNotOwnItem,
                canUpdateNotOwnItem,
                canDeleteNotOwnItem,
                canCreateSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR486,
                486,
                @"[In t:CalendarPermissionLevelType Simple Type] Contributor: The user can create items in the folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R487");

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR486,
                487,
                @"[In t:CalendarPermissionLevelType Simple Type] Contributor: The contents of the folder do not appear.");
        }

        /// <summary>
        /// This test case is designed to test user can create, read, edit and delete all items in the folder
        /// if granted Editor permission to him.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S06_TC04_CalendarPermissionLevelEditor()
        {
            this.MTGSAdapter = new MS_OXWSMTGSAdapter();
            this.MTGSAdapter.Initialize(this.Site);
            this.SRCHAdapter = new MS_OXWSSRCHAdapter();
            this.SRCHAdapter.Initialize(this.Site);
            this.FOLDAdapter = new MS_OXWSFOLDAdapter();
            this.FOLDAdapter.Initialize(this.Site);

            #region Organizer creates a calendar folder in the inbox folder and grant Editor permission to attendee.

            // Switch to user organizer.
            this.FOLDAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);

            CalendarPermissionSetType calendarPermissionSet = new CalendarPermissionSetType();
            calendarPermissionSet.CalendarPermissions = new CalendarPermissionType[1];
            calendarPermissionSet.CalendarPermissions[0] = new CalendarPermissionType();
            calendarPermissionSet.CalendarPermissions[0].CalendarPermissionLevel = CalendarPermissionLevelType.Editor;
            calendarPermissionSet.CalendarPermissions[0].UserId = new UserIdType();
            calendarPermissionSet.CalendarPermissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("AttendeeName", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "CustomFolderByOrganizer" }, new string[] { "IPF.Appointment" }, null);
            ((CalendarFolderType)createFolderRequest.Folders[0]).PermissionSet = calendarPermissionSet;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderIdByOrganizer = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.FolderToDelete = newFolderIdByOrganizer;
            #endregion

            #region Organizer gets the new created folder.

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderIdByOrganizer);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType foldersResponseInfo = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            Site.Assert.AreEqual<string>(newFolderIdByOrganizer.Id, foldersResponseInfo.Folders[0].FolderId.Id, "The calendar folder should be created successfully in inbox.");
            #endregion

            #region Organizer creates a calendar item in the created folder.
            CalendarItemType calendarItembyOrganizer = new CalendarItemType();
            calendarItembyOrganizer.UID = Guid.NewGuid().ToString();
            calendarItembyOrganizer.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByOrganizer");
            NonEmptyArrayOfAllItemsType allItemArray = new NonEmptyArrayOfAllItemsType();
            allItemArray.Items = new ItemType[] { calendarItembyOrganizer };
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = allItemArray;
            createItemRequest.SendMeetingInvitationsSpecified = true;
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            createItemRequest.SavedItemFolderId.Item = newFolderIdByOrganizer;
            CreateItemResponseType response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            ItemInfoResponseMessageType itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            Site.Assert.IsNotNull(itemInfoResponse, "The calendar item should be created successfully.");
            ItemIdType itemIdByOrganizer = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee creates a folder in the created folder.

            // Switch to user attendee.
            this.FOLDAdapter.SwitchUser(this.Attendee, this.AttendeePassword, this.Domain);

            createFolderRequest = this.GetCreateFolderRequest(newFolderIdByOrganizer.Id, new string[] { "CustomFolderByAttendee" }, new string[] { "IPF.Appointment" }, null);
            createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            bool canCreateSubFolder = ResponseClassType.Success == createFolderResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Attendee creates a calendar item in the created folder.
            this.SwitchMTGSUser(Role.Attendee);
            CalendarItemType calendarItembyAttendee = new CalendarItemType();
            calendarItembyAttendee.UID = Guid.NewGuid().ToString();
            calendarItembyAttendee.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByAttendee");
            createItemRequest.Items.Items[0] = calendarItembyAttendee;
            response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            bool canCreateOwnItem = null != itemInfoResponse;
            Site.Assert.IsTrue(canCreateOwnItem, "The calendar item should be created successfully.");
            ItemIdType itemIdByAttendee = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee gets the calendar item created by hisself.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByAttendee);
            bool canReadOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by hisself.
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByAttendee;

            // Update the created calendar item.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by hisself.
            ResponseMessageType responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByAttendee, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            #region Attendee gets the calendar item created by organizer.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByOrganizer);
            bool canReadNotOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by organizer.
            calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByOrganizer;

            // Update the created calendar item.
            updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateNotOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by organizer.
            responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByOrganizer, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteNotOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R91");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R91
            bool isVerifiedR91 = canCreateOwnItem && canReadOwnItem && canUpdateOwnItem && canDeleteOwnItem && canReadNotOwnItem
                && canUpdateNotOwnItem && canDeleteNotOwnItem && !canCreateSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR91,
                "Can create own item is {0};\n" +
                "Can read own item is {1};\n" +
                "Can edit own item is {2};\n " +
                "Can delete own item is {3};\n" +
                "Can read not own item is {4};\n" +
                "Can edit not own item is {5};\n" +
                "Can delete not own item is {6};\n" +
                "Can create subfolders is {7};\n",
                canCreateOwnItem,
                canReadOwnItem,
                canUpdateOwnItem,
                canDeleteOwnItem,
                canReadNotOwnItem,
                canUpdateNotOwnItem,
                canDeleteNotOwnItem,
                canCreateSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR91,
                91,
                @"[In t:CalendarPermissionLevelType Simple Type] Editor: The user can create, read, edit and delete all items in the folder.");
        }

        /// <summary>
        /// This test case is designed to test user has no permissions on the folder
        /// if granted None permission to him.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S06_TC05_CalendarPermissionLevelNone()
        {
            this.MTGSAdapter = new MS_OXWSMTGSAdapter();
            this.MTGSAdapter.Initialize(this.Site);
            this.SRCHAdapter = new MS_OXWSSRCHAdapter();
            this.SRCHAdapter.Initialize(this.Site);
            this.FOLDAdapter = new MS_OXWSFOLDAdapter();
            this.FOLDAdapter.Initialize(this.Site);

            #region Organizer creates a calendar folder in the inbox folder and grant None permission to attendee.

            // Switch to user organizer.
            this.FOLDAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);

            CalendarPermissionSetType calendarPermissionSet = new CalendarPermissionSetType();
            calendarPermissionSet.CalendarPermissions = new CalendarPermissionType[1];
            calendarPermissionSet.CalendarPermissions[0] = new CalendarPermissionType();
            calendarPermissionSet.CalendarPermissions[0].CalendarPermissionLevel = CalendarPermissionLevelType.None;
            calendarPermissionSet.CalendarPermissions[0].UserId = new UserIdType();
            calendarPermissionSet.CalendarPermissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("AttendeeName", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "CustomFolderByOrganizer" }, new string[] { "IPF.Appointment" }, null);
            ((CalendarFolderType)createFolderRequest.Folders[0]).PermissionSet = calendarPermissionSet;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderIdByOrganizer = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.FolderToDelete = newFolderIdByOrganizer;
            #endregion

            #region Organizer gets the new created folder.

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderIdByOrganizer);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType foldersResponseInfo = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            Site.Assert.AreEqual<string>(newFolderIdByOrganizer.Id, foldersResponseInfo.Folders[0].FolderId.Id, "The calendar folder should be created successfully in inbox.");

            CalendarFolderType folderInfo = (CalendarFolderType)foldersResponseInfo.Folders[0];

            #endregion

            #region Organizer creates a calendar item in the created folder.
            CalendarItemType calendarItembyOrganizer = new CalendarItemType();
            calendarItembyOrganizer.UID = Guid.NewGuid().ToString();
            calendarItembyOrganizer.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByOrganizer");
            NonEmptyArrayOfAllItemsType allItemArray = new NonEmptyArrayOfAllItemsType();
            allItemArray.Items = new ItemType[] { calendarItembyOrganizer };
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = allItemArray;
            createItemRequest.SendMeetingInvitationsSpecified = true;
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            createItemRequest.SavedItemFolderId.Item = newFolderIdByOrganizer;
            CreateItemResponseType response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            ItemInfoResponseMessageType itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            Site.Assert.IsNotNull(itemInfoResponse, "The calendar item should be created successfully.");
            ItemIdType itemIdByOrganizer = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee creates a folder in the created folder.

            // Switch to user attendee.
            this.FOLDAdapter.SwitchUser(this.Attendee, this.AttendeePassword, this.Domain);

            createFolderRequest = this.GetCreateFolderRequest(newFolderIdByOrganizer.Id, new string[] { "CustomFolderByAttendee" }, new string[] { "IPF.Appointment" }, null);
            createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            bool canCreateSubFolder = ResponseClassType.Success == createFolderResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Attendee creates a calendar item in the created folder.
            this.SwitchMTGSUser(Role.Attendee);
            CalendarItemType calendarItembyAttendee = new CalendarItemType();
            calendarItembyAttendee.UID = Guid.NewGuid().ToString();
            calendarItembyAttendee.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByAttendee");
            createItemRequest.Items.Items[0] = calendarItembyAttendee;
            response = this.MTGSAdapter.CreateItem(createItemRequest);
            bool canCreateOwnItem = response.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success;
            #endregion

            #region Attendee gets the calendar item created by organizer.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByOrganizer);
            bool canReadNotOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by organizer.
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByOrganizer;

            // Update the created calendar item.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateNotOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by organizer.
            ResponseMessageType responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByOrganizer, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteNotOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R94");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R94
            bool isVerifiedR94 = !canCreateOwnItem && !canReadNotOwnItem
                && !canUpdateNotOwnItem && !canDeleteNotOwnItem && !canCreateSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR94,
                "Can create own item is {0};\n" +
                "Can read not own item is {1};\n" +
                "Can edit not own item is {2};\n" +
                "Can delete not own item is {3};\n" +
                "Can create subfolders is {4};\n",
                canCreateOwnItem,
                canReadNotOwnItem,
                canUpdateNotOwnItem,
                canDeleteNotOwnItem,
                canCreateSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR94,
                94,
                @"[In t:CalendarPermissionLevelType Simple Type] None: The user has no permissions on the folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R104, the value of ReadItems is {0}", folderInfo.PermissionSet.CalendarPermissions[2].ReadItems);

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R104
            // This requirement can be captured after R94 is captured.
            this.Site.CaptureRequirementIfIsTrue(
                folderInfo.PermissionSet.CalendarPermissions[2].ReadItemsSpecified && folderInfo.PermissionSet.CalendarPermissions[2].ReadItems == CalendarPermissionReadAccessType.None,
                104,
                @"[In t:CalendarPermissionReadAccessType Simple Type] None: The user does not have permission to view items in the calendar.");
        }

        /// <summary>
        /// This test case is designed to test user can read all items in the folder
        /// if granted Reviewer permission to him.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S06_TC06_CalendarPermissionLevelReviewer()
        {
            this.MTGSAdapter = new MS_OXWSMTGSAdapter();
            this.MTGSAdapter.Initialize(this.Site);
            this.SRCHAdapter = new MS_OXWSSRCHAdapter();
            this.SRCHAdapter.Initialize(this.Site);
            this.FOLDAdapter = new MS_OXWSFOLDAdapter();
            this.FOLDAdapter.Initialize(this.Site);

            #region Organizer creates a calendar folder in the inbox folder and grant Reviewer permission to attendee.

            // Switch to user organizer.
            this.FOLDAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);

            CalendarPermissionSetType calendarPermissionSet = new CalendarPermissionSetType();
            calendarPermissionSet.CalendarPermissions = new CalendarPermissionType[1];
            calendarPermissionSet.CalendarPermissions[0] = new CalendarPermissionType();
            calendarPermissionSet.CalendarPermissions[0].CalendarPermissionLevel = CalendarPermissionLevelType.Reviewer;
            calendarPermissionSet.CalendarPermissions[0].UserId = new UserIdType();
            calendarPermissionSet.CalendarPermissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("AttendeeName", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "CustomFolderByOrganizer" }, new string[] { "IPF.Appointment" }, null);
            ((CalendarFolderType)createFolderRequest.Folders[0]).PermissionSet = calendarPermissionSet;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderIdByOrganizer = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.FolderToDelete = newFolderIdByOrganizer;
            #endregion

            #region Organizer gets the new created folder.

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderIdByOrganizer);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType foldersResponseInfo = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            Site.Assert.AreEqual<string>(newFolderIdByOrganizer.Id, foldersResponseInfo.Folders[0].FolderId.Id, "The calendar folder should be created successfully in inbox.");
            #endregion

            #region Organizer creates a calendar item in the created folder.
            CalendarItemType calendarItembyOrganizer = new CalendarItemType();
            calendarItembyOrganizer.UID = Guid.NewGuid().ToString();
            calendarItembyOrganizer.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByOrganizer");
            NonEmptyArrayOfAllItemsType allItemArray = new NonEmptyArrayOfAllItemsType();
            allItemArray.Items = new ItemType[] { calendarItembyOrganizer };
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = allItemArray;
            createItemRequest.SendMeetingInvitationsSpecified = true;
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            createItemRequest.SavedItemFolderId.Item = newFolderIdByOrganizer;
            CreateItemResponseType response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            ItemInfoResponseMessageType itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            Site.Assert.IsNotNull(itemInfoResponse, "The calendar item should be created successfully.");
            ItemIdType itemIdByOrganizer = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee creates a folder in the created folder.

            // Switch to user attendee.
            this.FOLDAdapter.SwitchUser(this.Attendee, this.AttendeePassword, this.Domain);

            createFolderRequest = this.GetCreateFolderRequest(newFolderIdByOrganizer.Id, new string[] { "CustomFolderByAttendee" }, new string[] { "IPF.Appointment" }, null);
            createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            bool canCreateSubFolder = ResponseClassType.Success == createFolderResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Attendee creates a calendar item in the created folder.
            this.SwitchMTGSUser(Role.Attendee);
            CalendarItemType calendarItembyAttendee = new CalendarItemType();
            calendarItembyAttendee.UID = Guid.NewGuid().ToString();
            calendarItembyAttendee.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByAttendee");
            createItemRequest.Items.Items[0] = calendarItembyAttendee;
            response = this.MTGSAdapter.CreateItem(createItemRequest);
            bool canCreateOwnItem = response.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success;
            #endregion

            #region Attendee gets the calendar item created by organizer.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByOrganizer);
            bool canReadNotOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by organizer.
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByOrganizer;

            // Update the created calendar item.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateNotOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by organizer.
            ResponseMessageType responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByOrganizer, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteNotOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R100");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R100
            bool isVerifiedR100 = !canCreateOwnItem && canReadNotOwnItem
                && !canUpdateNotOwnItem && !canDeleteNotOwnItem && !canCreateSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR100,
                "Can create own item is {0};\n" +
                "Can read not own item is {1};\n" +
                "Can edit not own item is {2};\n" +
                "Can delete not own item is {3};\n" +
                "Can create subfolders is {4};\n",
                canCreateOwnItem,
                canReadNotOwnItem,
                canUpdateNotOwnItem,
                canDeleteNotOwnItem,
                canCreateSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR100,
                100,
                @"[In t:CalendarPermissionLevelType Simple Type] Reviewer: The user can read all items in the folder.");
        }

        /// <summary>
        /// This test case is designed to test user has custom access permission on the folder
        /// if granted Custom permission to him.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S06_TC07_CalendarPermissionLevelCustom()
        {
            this.MTGSAdapter = new MS_OXWSMTGSAdapter();
            this.MTGSAdapter.Initialize(this.Site);
            this.SRCHAdapter = new MS_OXWSSRCHAdapter();
            this.SRCHAdapter.Initialize(this.Site);
            this.FOLDAdapter = new MS_OXWSFOLDAdapter();
            this.FOLDAdapter.Initialize(this.Site);

            #region Organizer creates a calendar folder in the inbox folder and grant Custom permission to attendee.

            // Switch to user organizer.
            this.FOLDAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);

            CalendarPermissionSetType calendarPermissionSet = new CalendarPermissionSetType();
            calendarPermissionSet.CalendarPermissions = new CalendarPermissionType[1];
            calendarPermissionSet.CalendarPermissions[0] = new CalendarPermissionType();
            calendarPermissionSet.CalendarPermissions[0].CalendarPermissionLevel = CalendarPermissionLevelType.Custom;
            calendarPermissionSet.CalendarPermissions[0].CanCreateItems = false;
            calendarPermissionSet.CalendarPermissions[0].CanCreateItemsSpecified = true;
            calendarPermissionSet.CalendarPermissions[0].CanCreateSubFolders = false;
            calendarPermissionSet.CalendarPermissions[0].CanCreateSubFoldersSpecified = true;
            calendarPermissionSet.CalendarPermissions[0].IsFolderVisible = false;
            calendarPermissionSet.CalendarPermissions[0].IsFolderVisibleSpecified = true;
            calendarPermissionSet.CalendarPermissions[0].IsFolderContact = false;
            calendarPermissionSet.CalendarPermissions[0].IsFolderContactSpecified = true;
            calendarPermissionSet.CalendarPermissions[0].IsFolderOwner = false;
            calendarPermissionSet.CalendarPermissions[0].IsFolderOwnerSpecified = true;
            calendarPermissionSet.CalendarPermissions[0].EditItems = PermissionActionType.None;
            calendarPermissionSet.CalendarPermissions[0].EditItemsSpecified = true;
            calendarPermissionSet.CalendarPermissions[0].DeleteItems = PermissionActionType.None;
            calendarPermissionSet.CalendarPermissions[0].DeleteItemsSpecified = true;
            calendarPermissionSet.CalendarPermissions[0].UserId = new UserIdType();
            calendarPermissionSet.CalendarPermissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("AttendeeName", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "CustomFolderByOrganizer" }, new string[] { "IPF.Appointment" }, null);
            ((CalendarFolderType)createFolderRequest.Folders[0]).PermissionSet = calendarPermissionSet;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderIdByOrganizer = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.FolderToDelete = newFolderIdByOrganizer;
            #endregion

            #region Organizer gets the new created folder.

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderIdByOrganizer);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType foldersResponseInfo = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            Site.Assert.AreEqual<string>(newFolderIdByOrganizer.Id, foldersResponseInfo.Folders[0].FolderId.Id, "The calendar folder should be created successfully in inbox.");

            #endregion

            #region Organizer creates a calendar item in the created folder.
            CalendarItemType calendarItembyOrganizer = new CalendarItemType();
            calendarItembyOrganizer.UID = Guid.NewGuid().ToString();
            calendarItembyOrganizer.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByOrganizer");
            NonEmptyArrayOfAllItemsType allItemArray = new NonEmptyArrayOfAllItemsType();
            allItemArray.Items = new ItemType[] { calendarItembyOrganizer };
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = allItemArray;
            createItemRequest.SendMeetingInvitationsSpecified = true;
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            createItemRequest.SavedItemFolderId.Item = newFolderIdByOrganizer;
            CreateItemResponseType response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            ItemInfoResponseMessageType itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            Site.Assert.IsNotNull(itemInfoResponse, "The calendar item should be created successfully.");
            ItemIdType itemIdByOrganizer = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee creates a folder in the created folder.

            // Switch to user attendee.
            this.FOLDAdapter.SwitchUser(this.Attendee, this.AttendeePassword, this.Domain);

            createFolderRequest = this.GetCreateFolderRequest(newFolderIdByOrganizer.Id, new string[] { "CustomFolderByAttendee" }, new string[] { "IPF.Appointment" }, null);
            createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            bool canCreateSubFolder = ResponseClassType.Success == createFolderResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Attendee creates a calendar item in the created folder.
            this.SwitchMTGSUser(Role.Attendee);
            CalendarItemType calendarItembyAttendee = new CalendarItemType();
            calendarItembyAttendee.UID = Guid.NewGuid().ToString();
            calendarItembyAttendee.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByAttendee");
            createItemRequest.Items.Items[0] = calendarItembyAttendee;
            response = this.MTGSAdapter.CreateItem(createItemRequest);
            bool canCreateOwnItem = response.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success;
            #endregion

            #region Attendee gets the calendar item created by organizer.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByOrganizer);
            bool canReadNotOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by organizer.
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByOrganizer;

            // Update the created calendar item.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateNotOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by organizer.
            ResponseMessageType responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByOrganizer, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteNotOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "VerifyMS-OXWSMTGS_R90");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R90
            bool isVerifiedR90 = !canCreateOwnItem && !canReadNotOwnItem
                && !canUpdateNotOwnItem && !canDeleteNotOwnItem && !canCreateSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR90,
                "Can create own item is {0};\n" +
                "Can read not own item is {1};\n" +
                "Can edit not own item is {2};\n" +
                "Can delete not own item is {3};\n" +
                "Can create subfolders is {4};\n",
                canCreateOwnItem,
                canReadNotOwnItem,
                canUpdateNotOwnItem,
                canDeleteNotOwnItem,
                canCreateSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR90,
                90,
                @"[In t:CalendarPermissionLevelType Simple Type] Custom: The user has custom access permissions on the folder.");
        }

        /// <summary>
        /// This test case is designed to test user can create and read all items in the folder, and delete only items that the user creates
        /// creates if granted NoneditingAuthor permission to him.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S06_TC08_CalendarPermissionLevelNoneditingAuthor()
        {
            this.MTGSAdapter = new MS_OXWSMTGSAdapter();
            this.MTGSAdapter.Initialize(this.Site);
            this.SRCHAdapter = new MS_OXWSSRCHAdapter();
            this.SRCHAdapter.Initialize(this.Site);
            this.FOLDAdapter = new MS_OXWSFOLDAdapter();
            this.FOLDAdapter.Initialize(this.Site);

            #region Organizer creates a calendar folder in the inbox folder and grant NoneditingAuthor permission to attendee.

            // Switch to user organizer.
            this.FOLDAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);

            CalendarPermissionSetType calendarPermissionSet = new CalendarPermissionSetType();
            calendarPermissionSet.CalendarPermissions = new CalendarPermissionType[1];
            calendarPermissionSet.CalendarPermissions[0] = new CalendarPermissionType();
            calendarPermissionSet.CalendarPermissions[0].CalendarPermissionLevel = CalendarPermissionLevelType.NoneditingAuthor;
            calendarPermissionSet.CalendarPermissions[0].UserId = new UserIdType();
            calendarPermissionSet.CalendarPermissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("AttendeeName", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "CustomFolderByOrganizer" }, new string[] { "IPF.Appointment" }, null);
            ((CalendarFolderType)createFolderRequest.Folders[0]).PermissionSet = calendarPermissionSet;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderIdByOrganizer = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.FolderToDelete = newFolderIdByOrganizer;
            #endregion

            #region Organizer gets the new created folder.

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderIdByOrganizer);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType foldersResponseInfo = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            Site.Assert.AreEqual<string>(newFolderIdByOrganizer.Id, foldersResponseInfo.Folders[0].FolderId.Id, "The calendar folder should be created successfully in inbox.");
            #endregion

            #region Organizer creates a calendar item in the created folder.
            CalendarItemType calendarItembyOrganizer = new CalendarItemType();
            calendarItembyOrganizer.UID = Guid.NewGuid().ToString();
            calendarItembyOrganizer.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByOrganizer");
            NonEmptyArrayOfAllItemsType allItemArray = new NonEmptyArrayOfAllItemsType();
            allItemArray.Items = new ItemType[] { calendarItembyOrganizer };
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = allItemArray;
            createItemRequest.SendMeetingInvitationsSpecified = true;
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            createItemRequest.SavedItemFolderId.Item = newFolderIdByOrganizer;
            CreateItemResponseType response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            ItemInfoResponseMessageType itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            Site.Assert.IsNotNull(itemInfoResponse, "The calendar item should be created successfully.");
            ItemIdType itemIdByOrganizer = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee creates a folder in the created folder.

            // Switch to user attendee.
            this.FOLDAdapter.SwitchUser(this.Attendee, this.AttendeePassword, this.Domain);

            createFolderRequest = this.GetCreateFolderRequest(newFolderIdByOrganizer.Id, new string[] { "CustomFolderByAttendee" }, new string[] { "IPF.Appointment" }, null);
            createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            bool canCreateSubFolder = ResponseClassType.Success == createFolderResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Attendee creates a calendar item in the created folder.
            this.SwitchMTGSUser(Role.Attendee);
            CalendarItemType calendarItembyAttendee = new CalendarItemType();
            calendarItembyAttendee.UID = Guid.NewGuid().ToString();
            calendarItembyAttendee.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByAttendee");
            createItemRequest.Items.Items[0] = calendarItembyAttendee;
            response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            bool canCreateOwnItem = null != itemInfoResponse;
            Site.Assert.IsTrue(canCreateOwnItem, "The calendar item should be created successfully.");
            ItemIdType itemIdByAttendee = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee gets the calendar item created by hisself.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByAttendee);
            bool canReadOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by hisself.
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByAttendee;

            // Update the created calendar item.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by hisself.
            ResponseMessageType responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByAttendee, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            #region Attendee gets the calendar item created by organizer.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByOrganizer);
            bool canReadNotOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by organizer.
            calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByOrganizer;

            // Update the created calendar item.
            updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateNotOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by organizer.
            responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByOrganizer, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteNotOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R95");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R95
            bool isVerifiedR95 = canCreateOwnItem && canReadOwnItem && !canUpdateOwnItem && canDeleteOwnItem && canReadNotOwnItem
                && !canUpdateNotOwnItem && !canDeleteNotOwnItem && !canCreateSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR95,
                "Can create own item is {0};\n" +
                "Can read own item is {1};\n" +
                "Can edit own item is {2};\n " +
                "Can delete own item is {3};\n" +
                "Can read not own item is {4};\n" +
                "Can edit not own item is {5};\n" +
                "Can delete not own item is {6};\n" +
                "Can create subfolders is {7};\n",
                canCreateOwnItem,
                canReadOwnItem,
                canUpdateOwnItem,
                canDeleteOwnItem,
                canReadNotOwnItem,
                canUpdateNotOwnItem,
                canDeleteNotOwnItem,
                canCreateSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR95,
                95,
                @"[In t:CalendarPermissionLevelType Simple Type] NoneditingAuthor: The user can create and read all items in the folder, and delete only items that the user creates.");
        }

        /// <summary>
        /// This test case is designed to test user can create and read all items in the folder, edit and delete only items that the user creates, and create subfolders
        /// creates if granted PublishingAuthor permission to him.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S06_TC09_CalendarPermissionLevelPublishingAuthor()
        {
            this.MTGSAdapter = new MS_OXWSMTGSAdapter();
            this.MTGSAdapter.Initialize(this.Site);
            this.SRCHAdapter = new MS_OXWSSRCHAdapter();
            this.SRCHAdapter.Initialize(this.Site);
            this.FOLDAdapter = new MS_OXWSFOLDAdapter();
            this.FOLDAdapter.Initialize(this.Site);

            #region Organizer creates a calendar folder in the inbox folder and grant PublishingAuthor permission to attendee.

            // Switch to user organizer.
            this.FOLDAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);

            CalendarPermissionSetType calendarPermissionSet = new CalendarPermissionSetType();
            calendarPermissionSet.CalendarPermissions = new CalendarPermissionType[1];
            calendarPermissionSet.CalendarPermissions[0] = new CalendarPermissionType();
            calendarPermissionSet.CalendarPermissions[0].CalendarPermissionLevel = CalendarPermissionLevelType.PublishingAuthor;
            calendarPermissionSet.CalendarPermissions[0].UserId = new UserIdType();
            calendarPermissionSet.CalendarPermissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("AttendeeName", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "CustomFolderByOrganizer" }, new string[] { "IPF.Appointment" }, null);
            ((CalendarFolderType)createFolderRequest.Folders[0]).PermissionSet = calendarPermissionSet;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderIdByOrganizer = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.FolderToDelete = newFolderIdByOrganizer;
            #endregion

            #region Organizer gets the new created folder.

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderIdByOrganizer);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType foldersResponseInfo = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            Site.Assert.AreEqual<string>(newFolderIdByOrganizer.Id, foldersResponseInfo.Folders[0].FolderId.Id, "The calendar folder should be created successfully in inbox.");
            #endregion

            #region Organizer creates a calendar item in the created folder.
            CalendarItemType calendarItembyOrganizer = new CalendarItemType();
            calendarItembyOrganizer.UID = Guid.NewGuid().ToString();
            calendarItembyOrganizer.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByOrganizer");
            NonEmptyArrayOfAllItemsType allItemArray = new NonEmptyArrayOfAllItemsType();
            allItemArray.Items = new ItemType[] { calendarItembyOrganizer };
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = allItemArray;
            createItemRequest.SendMeetingInvitationsSpecified = true;
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            createItemRequest.SavedItemFolderId.Item = newFolderIdByOrganizer;
            CreateItemResponseType response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            ItemInfoResponseMessageType itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            Site.Assert.IsNotNull(itemInfoResponse, "The calendar item should be created successfully.");
            ItemIdType itemIdByOrganizer = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee creates a folder in the created folder.

            // Switch to user attendee.
            this.FOLDAdapter.SwitchUser(this.Attendee, this.AttendeePassword, this.Domain);

            createFolderRequest = this.GetCreateFolderRequest(newFolderIdByOrganizer.Id, new string[] { "CustomFolderByAttendee" }, new string[] { "IPF.Appointment" }, null);
            createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            bool canCreateSubFolder = ResponseClassType.Success == createFolderResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Attendee creates a calendar item in the created folder.
            this.SwitchMTGSUser(Role.Attendee);
            CalendarItemType calendarItembyAttendee = new CalendarItemType();
            calendarItembyAttendee.UID = Guid.NewGuid().ToString();
            calendarItembyAttendee.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByAttendee");
            createItemRequest.Items.Items[0] = calendarItembyAttendee;
            response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            bool canCreateOwnItem = null != itemInfoResponse;
            Site.Assert.IsTrue(canCreateOwnItem, "The calendar item should be created successfully.");
            ItemIdType itemIdByAttendee = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee gets the calendar item created by hisself.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByAttendee);
            bool canReadOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by hisself.
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByAttendee;

            // Update the created calendar item.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by hisself.
            ResponseMessageType responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByAttendee, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            #region Attendee gets the calendar item created by organizer.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByOrganizer);
            bool canReadNotOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by organizer.
            calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByOrganizer;

            // Update the created calendar item.
            updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateNotOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by organizer.
            responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByOrganizer, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteNotOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R98");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R98
            bool isVerifiedR98 = canCreateOwnItem && canReadOwnItem && canUpdateOwnItem && canDeleteOwnItem && canReadNotOwnItem
                && !canUpdateNotOwnItem && !canDeleteNotOwnItem && canCreateSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR98,
                "Can create own item is {0};\n" +
                "Can read own item is {1};\n" +
                "Can edit own item is {2};\n " +
                "Can delete own item is {3};\n" +
                "Can read not own item is {4};\n" +
                "Can edit not own item is {5};\n" +
                "Can delete not own item is {6};\n" +
                "Can create subfolders is {7};\n",
                canCreateOwnItem,
                canReadOwnItem,
                canUpdateOwnItem,
                canDeleteOwnItem,
                canReadNotOwnItem,
                canUpdateNotOwnItem,
                canDeleteNotOwnItem,
                canCreateSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR98,
                98,
                @"[In t:CalendarPermissionLevelType Simple Type] PublishingAuthor: The user can create and read all items in the folder, edit and delete only items that the user creates, and create subfolders.");
        }

        /// <summary>
        /// This test case is designed to test user can create, read, edit, and delete all items in the folder, and create subfolders
        /// creates if granted PublishingEditor permission to him.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S06_TC10_CalendarPermissionLevelPublishingEditor()
        {
            this.MTGSAdapter = new MS_OXWSMTGSAdapter();
            this.MTGSAdapter.Initialize(this.Site);
            this.SRCHAdapter = new MS_OXWSSRCHAdapter();
            this.SRCHAdapter.Initialize(this.Site);
            this.FOLDAdapter = new MS_OXWSFOLDAdapter();
            this.FOLDAdapter.Initialize(this.Site);

            #region Organizer creates a calendar folder in the inbox folder and grant PublishingEditor permission to attendee.

            // Switch to user organizer.
            this.FOLDAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);

            CalendarPermissionSetType calendarPermissionSet = new CalendarPermissionSetType();
            calendarPermissionSet.CalendarPermissions = new CalendarPermissionType[1];
            calendarPermissionSet.CalendarPermissions[0] = new CalendarPermissionType();
            calendarPermissionSet.CalendarPermissions[0].CalendarPermissionLevel = CalendarPermissionLevelType.PublishingEditor;
            calendarPermissionSet.CalendarPermissions[0].UserId = new UserIdType();
            calendarPermissionSet.CalendarPermissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("AttendeeName", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "CustomFolderByOrganizer" }, new string[] { "IPF.Appointment" }, null);
            ((CalendarFolderType)createFolderRequest.Folders[0]).PermissionSet = calendarPermissionSet;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderIdByOrganizer = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.FolderToDelete = newFolderIdByOrganizer;
            #endregion

            #region Organizer gets the new created folder.

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderIdByOrganizer);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType foldersResponseInfo = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            Site.Assert.AreEqual<string>(newFolderIdByOrganizer.Id, foldersResponseInfo.Folders[0].FolderId.Id, "The calendar folder should be created successfully in inbox.");
            #endregion

            #region Organizer creates a calendar item in the created folder.
            CalendarItemType calendarItembyOrganizer = new CalendarItemType();
            calendarItembyOrganizer.UID = Guid.NewGuid().ToString();
            calendarItembyOrganizer.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByOrganizer");
            NonEmptyArrayOfAllItemsType allItemArray = new NonEmptyArrayOfAllItemsType();
            allItemArray.Items = new ItemType[] { calendarItembyOrganizer };
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = allItemArray;
            createItemRequest.SendMeetingInvitationsSpecified = true;
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            createItemRequest.SavedItemFolderId.Item = newFolderIdByOrganizer;
            CreateItemResponseType response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            ItemInfoResponseMessageType itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            Site.Assert.IsNotNull(itemInfoResponse, "The calendar item should be created successfully.");
            ItemIdType itemIdByOrganizer = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee creates a folder in the created folder.

            // Switch to user attendee.
            this.FOLDAdapter.SwitchUser(this.Attendee, this.AttendeePassword, this.Domain);

            createFolderRequest = this.GetCreateFolderRequest(newFolderIdByOrganizer.Id, new string[] { "CustomFolderByAttendee" }, new string[] { "IPF.Appointment" }, null);
            createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            bool canCreateSubFolder = ResponseClassType.Success == createFolderResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Attendee creates a calendar item in the created folder.
            this.SwitchMTGSUser(Role.Attendee);
            CalendarItemType calendarItembyAttendee = new CalendarItemType();
            calendarItembyAttendee.UID = Guid.NewGuid().ToString();
            calendarItembyAttendee.Subject = Common.GenerateResourceName(this.Site, "CalendarItemCreatedByAttendee");
            createItemRequest.Items.Items[0] = calendarItembyAttendee;
            response = this.MTGSAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, response.ResponseMessages.Items[0].ResponseClass, "The calendar item should be created successfully.");
            itemInfoResponse = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            bool canCreateOwnItem = null != itemInfoResponse;
            Site.Assert.IsTrue(canCreateOwnItem, "The calendar item should be created successfully.");
            ItemIdType itemIdByAttendee = itemInfoResponse.Items.Items[0].ItemId;
            #endregion

            #region Attendee gets the calendar item created by hisself.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByAttendee);
            bool canReadOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by hisself.
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByAttendee;

            // Update the created calendar item.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by hisself.
            ResponseMessageType responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByAttendee, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            #region Attendee gets the calendar item created by organizer.
            itemInfoResponse = this.GetSingleCalendarItem(Role.Attendee, itemIdByOrganizer);
            bool canReadNotOwnItem = null != itemInfoResponse;
            #endregion

            #region Attendee updates the calendar item created by organizer.
            calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = itemIdByOrganizer;

            // Update the created calendar item.
            updatedItem = this.UpdateSingleCalendarItem(Role.Attendee, itemChangeInfo, CalendarItemUpdateOperationType.SendToNone);
            bool canUpdateNotOwnItem = null != updatedItem;
            #endregion

            #region Attendee deletes the calendar item created by organizer.
            responseMessage = this.DeleteSingleCalendarItem(Role.Attendee, itemIdByOrganizer, CalendarItemCreateOrDeleteOperationType.SendToNone);
            bool canDeleteNotOwnItem = null != responseMessage && responseMessage.ResponseClass == ResponseClassType.Success;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R99");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R99
            bool isVerifiedR99 = canCreateOwnItem && canReadOwnItem && canUpdateOwnItem && canDeleteOwnItem && canReadNotOwnItem
                && canUpdateNotOwnItem && canDeleteNotOwnItem && canCreateSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR99,
                "Can create own item is {0};\n" +
                "Can read own item is {1};\n" +
                "Can edit own item is {2};\n " +
                "Can delete own item is {3};\n" +
                "Can read not own item is {4};\n" +
                "Can edit not own item is {5};\n" +
                "Can delete not own item is {6};\n" +
                "Can create subfolders is {7};\n",
                canCreateOwnItem,
                canReadOwnItem,
                canUpdateOwnItem,
                canDeleteOwnItem,
                canReadNotOwnItem,
                canUpdateNotOwnItem,
                canDeleteNotOwnItem,
                canCreateSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR99,
                99,
                @"[In t:CalendarPermissionLevelType Simple Type] PublishingEditor: The user can create, read, edit, and delete all items in the folder, and create subfolders.");
        }

        /// <summary>
        /// This test case is designed to test element SharingEffectiveRights is read only.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S06_TC11_SharingEffectiveRightsIsReadOnly()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1338, this.Site), "Exchange 2007 does not support the SharingEffectiveRights element.");

            #region Organizer creates a calendar folder with setting SharingEffectiveRights element.
            
            // Switch to user organizer.
            this.FOLDAdapter.SwitchUser(this.Organizer, this.OrganizerPassword, this.Domain);

            CalendarPermissionSetType calendarPermissionSet = new CalendarPermissionSetType();
            calendarPermissionSet.CalendarPermissions = new CalendarPermissionType[1];
            calendarPermissionSet.CalendarPermissions[0] = new CalendarPermissionType();
            calendarPermissionSet.CalendarPermissions[0].CalendarPermissionLevel = CalendarPermissionLevelType.Owner;
            calendarPermissionSet.CalendarPermissions[0].UserId = new UserIdType();
            calendarPermissionSet.CalendarPermissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("AttendeeName", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "CustomFolderByOrganizer" }, new string[] { "IPF.Appointment" }, null);
            ((CalendarFolderType)createFolderRequest.Folders[0]).PermissionSet = calendarPermissionSet;
            ((CalendarFolderType)createFolderRequest.Folders[0]).SharingEffectiveRights = CalendarPermissionReadAccessType.FullDetails;
            ((CalendarFolderType)createFolderRequest.Folders[0]).SharingEffectiveRightsSpecified = true;
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1046");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1046
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createFolderResponse.ResponseMessages.Items[0].ResponseCode,
                1046,
                @"[In t:CalendarFolderType Complex Type]SharingEffectiveRights: This element is read-only.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R144");

            this.Site.CaptureRequirementIfIsInstanceOfType(
               ((CalendarFolderType)createFolderRequest.Folders[0]).SharingEffectiveRights,
               typeof(CalendarPermissionReadAccessType),
               144,
               @"[In t:CalendarFolderType Complex Type] The type of  SharingEffectiveRights is t:CalendarPermissionReadAccessType (section 2.2.5.5).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1338");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1338
            // Server handles the SharingEffectiveRights element, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1338,
                @"[In Appendix C: Product Behavior] Implementation does support the SharingEffectiveRights element. (Exchange 2010 and above follow this behavior.)");
            #endregion
        }

        #endregion

        /// <summary>
        /// Generate the request message for operation "GetFolder".
        /// </summary>
        /// <param name="shapeName">The properties to include in the response.</param>
        /// <param name="folderIds">An array of folder identifiers.</param>   
        /// <returns>Get folder request instance that will send to server.</returns>
        protected GetFolderType GetGetFolderRequest(DefaultShapeNamesType shapeName, params BaseFolderIdType[] folderIds)
        {
            Site.Assert.IsNotNull(folderIds, "Folders id should not be null!");
            Site.Assert.AreNotEqual<int>(0, folderIds.Length, "Folders id should contains at least one Id!");
            GetFolderType getFolderRequest = new GetFolderType();

            // Specify how many folders need to be gotten.
            int folderCount = folderIds.Length;

            // Set the request's folderId.
            getFolderRequest.FolderIds = new BaseFolderIdType[folderCount];

            for (int folderIdIndex = 0; folderIdIndex < folderCount; folderIdIndex++)
            {
                getFolderRequest.FolderIds[folderIdIndex] = folderIds[folderIdIndex];
            }

            // Set folder shape.
            getFolderRequest.FolderShape = new FolderResponseShapeType();
            getFolderRequest.FolderShape.BaseShape = shapeName;
            return getFolderRequest;
        }

        /// <summary>
        /// Generate the request message for operation "CreateFolder".
        /// </summary>
        /// <param name="parentFolderId">The folder identifier for the parent folder.</param>
        /// <param name="folderNames">An array of display name of the folders to be created.</param>
        /// <param name="folderClasses">An array of folder class value of the folders to be created.</param>
        /// <param name="permissionSet">An array of permission set value of the folder.</param>
        /// <returns>Create folder request instance that will send to server.</returns>
        protected CreateFolderType GetCreateFolderRequest(string parentFolderId, string[] folderNames, string[] folderClasses, PermissionSetType[] permissionSet)
        {
            CreateFolderType createFolderRequest = new CreateFolderType();
            createFolderRequest.ParentFolderId = new TargetFolderIdType();

            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            DistinguishedFolderIdNameType distinguishedFolderIdName = new DistinguishedFolderIdNameType();
            bool isSuccess = Enum.TryParse<DistinguishedFolderIdNameType>(parentFolderId, true, out distinguishedFolderIdName);

            if (isSuccess)
            {
                distinguishedFolderId.Id = distinguishedFolderIdName;
                createFolderRequest.ParentFolderId.Item = distinguishedFolderId;
            }
            else
            {
                FolderIdType id = new FolderIdType();
                id.Id = parentFolderId;
                createFolderRequest.ParentFolderId.Item = id;
            }

            createFolderRequest = this.ConfigureFolderProperty(folderNames, folderClasses, permissionSet, createFolderRequest);

            return createFolderRequest;
        }

        /// <summary>
        /// Set related folder properties of create folder request
        /// </summary>
        /// <param name="displayNames">Display names of folders that will be set into create folder request.</param>
        /// <param name="folderClasses">Folder class values of folders that will be set into create folder request.</param>
        /// <param name="folderPermissions">Folder permission values of folders that will be set into create folder request. </param>
        /// <param name="createFolderRequest">Create folder request instance that needs to set property values.</param>
        /// <returns>Create folder request instance that have folder property value configured.</returns>
        protected CreateFolderType ConfigureFolderProperty(string[] displayNames, string[] folderClasses, PermissionSetType[] folderPermissions, CreateFolderType createFolderRequest)
        {
            Site.Assert.IsNotNull(displayNames, "Display names should not be null!");
            Site.Assert.IsNotNull(folderClasses, "Folder classes should not be null!");
            Site.Assert.AreEqual<int>(displayNames.Length, folderClasses.Length, "Folder names count should equals to folder class value count!");
            if (folderPermissions != null)
            {
                Site.Assert.AreEqual<int>(displayNames.Length, folderPermissions.Length, "Folder names count should equals to folder permission value count!");
            }

            int folderCount = displayNames.Length;
            createFolderRequest.Folders = new BaseFolderType[folderCount];
            for (int folderPropertyIndex = 0; folderPropertyIndex < folderCount; folderPropertyIndex++)
            {
                string folderResourceName = Common.GenerateResourceName(this.Site, displayNames[folderPropertyIndex]);

                if (folderClasses[folderPropertyIndex] == "IPF.Appointment")
                {
                    CalendarFolderType calendarFolder = new CalendarFolderType();
                    calendarFolder.DisplayName = folderResourceName;
                    createFolderRequest.Folders[folderPropertyIndex] = calendarFolder;
                }
                else if (folderClasses[folderPropertyIndex] == "IPF.Contact")
                {
                    ContactsFolderType contactFolder = new ContactsFolderType();
                    contactFolder.DisplayName = folderResourceName;
                    if (folderPermissions != null)
                    {
                        contactFolder.PermissionSet = folderPermissions[folderPropertyIndex];
                    }

                    createFolderRequest.Folders[folderPropertyIndex] = contactFolder;
                }
                else if (folderClasses[folderPropertyIndex] == "IPF.Task")
                {
                    TasksFolderType taskFolder = new TasksFolderType();
                    taskFolder.DisplayName = folderResourceName;
                    if (folderPermissions != null)
                    {
                        taskFolder.PermissionSet = folderPermissions[folderPropertyIndex];
                    }

                    createFolderRequest.Folders[folderPropertyIndex] = taskFolder;
                }
                else if (folderClasses[folderPropertyIndex] == "IPF.Search")
                {
                    SearchFolderType searchFolder = new SearchFolderType();
                    searchFolder.DisplayName = folderResourceName;

                    // Set search parameters.
                    searchFolder.SearchParameters = new SearchParametersType();
                    searchFolder.SearchParameters.Traversal = SearchFolderTraversalType.Deep;
                    searchFolder.SearchParameters.TraversalSpecified = true;
                    searchFolder.SearchParameters.BaseFolderIds = new DistinguishedFolderIdType[1];
                    DistinguishedFolderIdType inboxType = new DistinguishedFolderIdType();
                    inboxType.Id = new DistinguishedFolderIdNameType();
                    inboxType.Id = DistinguishedFolderIdNameType.inbox;
                    searchFolder.SearchParameters.BaseFolderIds[0] = inboxType;

                    // Use the following search filter 
                    searchFolder.SearchParameters.Restriction = new RestrictionType();
                    PathToUnindexedFieldType path = new PathToUnindexedFieldType();
                    path.FieldURI = UnindexedFieldURIType.itemSubject;
                    RestrictionType restriction = new RestrictionType();
                    ExistsType isEqual = new ExistsType();
                    isEqual.Item = path;
                    restriction.Item = isEqual;
                    searchFolder.SearchParameters.Restriction = restriction;

                    if (folderPermissions != null)
                    {
                        searchFolder.PermissionSet = folderPermissions[folderPropertyIndex];
                    }

                    createFolderRequest.Folders[folderPropertyIndex] = searchFolder;
                }
                else
                {
                    // Set Display Name and Folder Class for the folder to be created.
                    FolderType folder = new FolderType();
                    folder.DisplayName = folderResourceName;
                    folder.FolderClass = folderClasses[folderPropertyIndex];

                    if (folderPermissions != null)
                    {
                        folder.PermissionSet = folderPermissions[folderPropertyIndex];
                    }

                    createFolderRequest.Folders[folderPropertyIndex] = folder;
                }
            }

            return createFolderRequest;
        }
    }
}
