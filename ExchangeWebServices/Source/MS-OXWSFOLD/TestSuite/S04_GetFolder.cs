//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to verify GetFolder operation.
    /// </summary>
    [TestClass]
    public class S04_GetFolder : TestSuiteBase
    {
        #region Class initialize and clean up

        /// <summary>
        /// Initialize the test class.
        /// </summary>
        /// <param name="testContext">Context to initialize.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
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
        /// This test case verifies requirements related to getting calendar folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S04_TC01_GetCalendarFolder()
        {
            #region Get the Calendar folder.

            DistinguishedFolderIdType folder = new DistinguishedFolderIdType();
            folder.Id = (DistinguishedFolderIdNameType)DistinguishedFolderIdNameType.calendar;

            // GetFolder request.
            GetFolderType getCalendarFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.IdOnly, folder);

            // Set to get additional property : folder class
            getCalendarFolderRequest.FolderShape.AdditionalProperties = new BasePathToElementType[]
            {
                  new PathToUnindexedFieldType()
                        {
                            FieldURI = UnindexedFieldURIType.folderFolderClass
                        },
                  new PathToUnindexedFieldType()
                        {
                            FieldURI = UnindexedFieldURIType.folderEffectiveRights
                        }
            };

            // Get the Calendar folder.
            GetFolderResponseType getCalendarFolderResponse = this.FOLDAdapter.GetFolder(getCalendarFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getCalendarFolderResponse, 1, this.Site);

            // Variable to save the folder.
            FolderInfoResponseMessageType foldersResponseInfo = (FolderInfoResponseMessageType)getCalendarFolderResponse.ResponseMessages.Items[0];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R29");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R29
            this.Site.CaptureRequirementIfIsInstanceOfType(
                foldersResponseInfo.Folders[0],
                typeof(CalendarFolderType),
                29,
                @"[In t:ArrayOfFoldersType Complex Type]The type of element CalendarFolder is t:CalendarFolderType ([MS-OXWSMTGS] section 2.2.4.3).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R2901");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R2901
            this.Site.CaptureRequirementIfIsInstanceOfType(
                foldersResponseInfo.Folders[0],
                typeof(CalendarFolderType),
                2901,
                @"[In t:ArrayOfFoldersType Complex Type]CalendarFolder represents a Calendar folder in a mailbox.");

            CalendarFolderType folderInfo = (CalendarFolderType)foldersResponseInfo.Folders[0];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R81");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R81
            // Since the schema has been validate and this element is not null, this requirement will be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.EffectiveRights,
                81,
                @"[In t:BaseFolderType Complex Type]The type of element EffectiveRights is t:EffectiveRightsType ([MS-OXWSCDATA] section 2.2.4.25).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R5934");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R5934
            // Effective rights is returned in response and schema is verified in adapter so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.EffectiveRights,
                5934,
                @"[In t:BaseFolderType Complex Type]This property[EffectiveRights] is returned in a response.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R70.");

            // Verify MS-OXWSFOLD_R70.
            Site.CaptureRequirementIfAreEqual<string>(
                "IPF.Appointment",
                folderInfo.FolderClass,
                70,
                @"[In t:BaseFolderType Complex Type]This value[FolderClass] MUST be ""IPF.Appointment"" for Calendar folders.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R42104");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R42104
            // Additional property folder class is returned from server and schema is verified in adapter, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.FolderClass,
                42104,
                @"[In t:FolderResponseShapeType Complex Type]The element [AdditionalProperties] with type [t:NonEmptyArrayOfPathsToElementType] specifies the identity of additional properties to be returned in a response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R42105");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R42105
            // Additional property folder class is returned from server and schema is verified in adapter, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.FolderClass,
                42105,
                @"[In t:FolderResponseShapeType Complex Type][In t:NonEmptyArrayOfPathsToElementType Complex Type] The element [t:Path] with type [t:Path] specifies a property to be returned in a response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3864");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3864
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getCalendarFolderResponse.ResponseMessages.Items[0].ResponseClass,
                3864,
                @"[In GetFolder Operation]A successful GetFolder operation request returns a GetFolderResponse element with the ResponseClass attribute of the GetFolderResponseMessage element set to ""Success"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R38644");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R38644
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                getCalendarFolderResponse.ResponseMessages.Items[0].ResponseCode,
                38644,
                @"[In GetFolder Operation]A successful GetFolder operation request returns a GetFolderResponse element with the ResponseCode element of the GetFolderResponse element set to ""NoError"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R575");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R575
            // Folder information is returned from server and schema has verified in adapter, so this requirement can be covered.
            this.Site.CaptureRequirementIfIsNotNull(
                foldersResponseInfo,
                575,
                @"[In m:GetFolderType Complex Type]The GetFolderType complex type specifies a request message to get a folder in a server database.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R428");

            // GetFolder operation in this test case gets a distinguished folder : Calendar folder.
            this.Site.CaptureRequirement(
                428,
                @"[In t:NonEmptyArrayOfBaseFolderIdsType Complex Type]DistinguishedFolderId specifies a distinguished folder identifier.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R8402");

            // GetFolder operation in this test case gets a distinguished folder : Calendar folder.
            this.Site.CaptureRequirement(
                8402,
                @"[In t:BaseFolderType Complex Type]DistinguishedFolderId specifies an identifier for a folder that can be referenced by name.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R96");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R96
            // Folder information is returned from server and schema has verified in adapter, so this requirement can be covered.
            this.Site.CaptureRequirementIfIsNotNull(
                foldersResponseInfo,
                96,
                @"[In m:FolderInfoResponseMessageType Complex Type]The type of element Folders is t:ArrayOfFoldersType (section 2.2.4.1).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R9602");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R9602
            // Folder information is returned from server and schema has verified in adapter, so this requirement can be covered.
            this.Site.CaptureRequirementIfIsNotNull(
                foldersResponseInfo,
                9602,
                @"[In m:FolderInfoResponseMessageType Complex Type][Folders] Represents the folders that are returned with the response message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R8401");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R8401
            // Distinguished folder id is set in request and schema is verified in adapter, so if folder information is returned successfully this can be covered.
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getCalendarFolderResponse.ResponseMessages.Items[0].ResponseClass,
                8401,
                @"[In t:BaseFolderType Complex Type]The type of element DistinguishedFolderId is t:DistinguishedFolderIdNameType ([MS-OXWSCDATA] section 2.2.3.10).");

            #endregion
        }

        /// <summary>
        /// This test case verifies requirements related to getting contacts folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S04_TC02_GetContactsFolder()
        {
            #region Get the contacts folder.

            DistinguishedFolderIdType folder = new DistinguishedFolderIdType();
            folder.Id = DistinguishedFolderIdNameType.contacts;

            // GetFolder request.
            GetFolderType getContactsFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folder);

            // Get the Contacts folder.
            GetFolderResponseType getContactsFolderResponse = this.FOLDAdapter.GetFolder(getContactsFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getContactsFolderResponse, 1, this.Site);

            // Variable to save the folder.
            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getContactsFolderResponse.ResponseMessages.Items[0];

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R31");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R31
            this.Site.CaptureRequirementIfIsInstanceOfType(
                allFolders.Folders[0],
                typeof(ContactsFolderType),
                31,
                @"[In t:ArrayOfFoldersType Complex Type]The type of element ContactsFolder is t:ContactsFolderType ([MS-OXWSCONT] section 3.1.4.1.2.2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3301");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3301
            this.Site.CaptureRequirementIfIsInstanceOfType(
                allFolders.Folders[0],
                typeof(ContactsFolderType),
                3301,
                @"[In t:ArrayOfFoldersType Complex Type]ContactsFolder represents a Contacts folder in a mailbox.");

            ContactsFolderType folderInfo = (ContactsFolderType)allFolders.Folders[0];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R69");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R69
            // Folder class value is returned from server and schema has verified in adapter, so this requirement can be covered.
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.FolderClass,
                69,
                @"[In t:BaseFolderType Complex Type]The type of element FolderClass is xs:string [XMLSCHEMA2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R585.");

            // Verify MS-OXWSFOLD_R585.
            Site.CaptureRequirementIfAreEqual<string>(
                "IPF.Contact",
                folderInfo.FolderClass,
                585,
                @"[In t:BaseFolderType Complex Type]This value[FolderClass] MUST be ""IPF.Contact"" for Contacts folders.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R7101");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R7101
            this.Site.CaptureRequirementIfAreEqual<string>(
                "Contacts",
                folderInfo.DisplayName,
                7101,
                @"[In t:BaseFolderType Complex Type]DisplayName specifies the display name of the folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R71");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R71
            // Folder display name is returned from server and schema has verified in adapter, so this requirement can be covered.
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.DisplayName,
                71,
                @"[In t:BaseFolderType Complex Type]The type of element DisplayName is xs:string.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R66");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R66
            // Folder id is returned from server and schema has verified in adapter, so this requirement can be covered.
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.FolderId,
                66,
                @"[In t:BaseFolderType Complex Type]The type of element FolderId is t:FolderIdType ([MS-OXWSCDATA] section 2.2.4.31).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R6602");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R6602
            // Folder id is returned from server and schema is verified in adapter so this requirement can be covered.
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.FolderId,
                6602,
                @"[In t:BaseFolderType Complex Type]FolderId specifies the folder identifier and change key.");
        }

        /// <summary>
        /// This test case verifies requirements related to getting journal folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S04_TC03_GetJournalFolder()
        {
            #region Get the Journal folder.

            DistinguishedFolderIdType folder = new DistinguishedFolderIdType();
            folder.Id = DistinguishedFolderIdNameType.journal;

            // GetFolder request.
            GetFolderType getJournalFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folder);

            // Get the journal folder.
            GetFolderResponseType getJournalFolderResponse = this.FOLDAdapter.GetFolder(getJournalFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getJournalFolderResponse, 1, this.Site);

            // Variable to save the folder.
            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getJournalFolderResponse.ResponseMessages.Items[0];
            BaseFolderType folderInfo = (BaseFolderType)allFolders.Folders[0];

            #endregion

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R586.");

            // Verify MS-OXWSFOLD_R586.
            Site.CaptureRequirementIfAreEqual<string>(
                "IPF.Journal",
                folderInfo.FolderClass,
                586,
                @"[In t:BaseFolderType Complex Type]This value[FolderClass] MUST be ""IPF.Journal"" for journal folders.");
        }

        /// <summary>
        /// This test case verifies requirements related to getting inbox folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S04_TC04_GetInboxFolder()
        {
            #region Get the Inbox folder.

            DistinguishedFolderIdType folder = new DistinguishedFolderIdType();
            folder.Id = DistinguishedFolderIdNameType.inbox;

            // GetFolder request.
            GetFolderType getInboxFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folder);

            // Get the inbox folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getInboxFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Variable to save the folder.
            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            BaseFolderType folderInfo = (BaseFolderType)allFolders.Folders[0];

            #endregion

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R587.");

            // Verify MS-OXWSFOLD_R587.
            Site.CaptureRequirementIfAreEqual<string>(
                "IPF.Note",
                folderInfo.FolderClass,
                587,
                @"[In t:BaseFolderType Complex Type]This value[FolderClass] MUST be ""IPF.Note"" for mail folders.");
        }

        /// <summary>
        /// This test case verifies requirements related to getting notes folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S04_TC05_GetNotesFolder()
        {
            #region Get the notes folder.

            DistinguishedFolderIdType folder = new DistinguishedFolderIdType();
            folder.Id = DistinguishedFolderIdNameType.notes;

            // GetFolder request.
            GetFolderType getNotesFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folder);

            // Get the notes folder.
            GetFolderResponseType getNotesFolderResponse = this.FOLDAdapter.GetFolder(getNotesFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getNotesFolderResponse, 1, this.Site);

            // Variable to save the folder.
            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getNotesFolderResponse.ResponseMessages.Items[0];
            BaseFolderType folderInfo = (BaseFolderType)allFolders.Folders[0];

            #endregion

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R588.");

            // Verify MS-OXWSFOLD_R588.
            Site.CaptureRequirementIfAreEqual<string>(
                "IPF.StickyNote",
                folderInfo.FolderClass,
                588,
                @"[In t:BaseFolderType Complex Type]This value[FolderClass] MUST be ""IPF.StickyNote"" for note folders.");
        }

        /// <summary>
        /// This test case verifies requirements related to getting tasks folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S04_TC06_GetTasksFolder()
        {
            #region Get the tasks folder.

            DistinguishedFolderIdType folder = new DistinguishedFolderIdType();
            folder.Id = DistinguishedFolderIdNameType.tasks;

            // GetFolder request.
            GetFolderType getTasksFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folder);

            // Get the tasks folder.
            GetFolderResponseType getTasksFolderResponse = this.FOLDAdapter.GetFolder(getTasksFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getTasksFolderResponse, 1, this.Site);

            // Variable to save the folder.
            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getTasksFolderResponse.ResponseMessages.Items[0];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R35");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R35
            this.Site.CaptureRequirementIfIsInstanceOfType(
                allFolders.Folders[0],
                typeof(TasksFolderType),
                35,
                @"[In t:ArrayOfFoldersType Complex Type]The type of element TasksFolder is t:TasksFolderType ([MS-OXWSTASK] section 2.2.4.2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3501");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3501
            this.Site.CaptureRequirementIfIsInstanceOfType(
                allFolders.Folders[0],
                typeof(TasksFolderType),
                3501,
                @"[In t:ArrayOfFoldersType Complex Type]TasksFolder represents a Tasks folder that is contained in a mailbox.");

            TasksFolderType folderInfo = (TasksFolderType)allFolders.Folders[0];

            #endregion

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R589.");

            // Verify MS-OXWSFOLD_R589.
            Site.CaptureRequirementIfAreEqual<string>(
                "IPF.Task",
                folderInfo.FolderClass,
                589,
                @"[In t:BaseFolderType Complex Type]This value[FolderClass] MUST be ""IPF.Task"" for Tasks folders.");
        }

        /// <summary>
        /// This test case verifies the requirements related to UnreadCount property of folder via checking the property of Inbox before and after sending a mail.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S04_TC07_UnReadCount()
        {
            #region Get the Inbox folder

            DistinguishedFolderIdType folder = new DistinguishedFolderIdType();
            folder.Id = DistinguishedFolderIdNameType.inbox;

            // GetFolder request.
            GetFolderType getInboxRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folder);

            // Get the Inbox folder.
            GetFolderResponseType getInboxResponse = this.FOLDAdapter.GetFolder(getInboxRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxResponse, 1, this.Site);

            // Variable to save the folder.
            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getInboxResponse.ResponseMessages.Items[0];
            BaseFolderType folderInfo = (BaseFolderType)allFolders.Folders[0];

            // Variable to save the count of unread messages before sent mail to the specific account.
            // Save the unread message count.
            int count = ((FolderType)folderInfo).UnreadCount;

            #endregion

            #region Create an unread message.

            string itemName = Common.GenerateResourceName(this.Site, "Test Mail");

            // Send a mail to User1
            ItemIdType itemId = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), DistinguishedFolderIdNameType.inbox.ToString(), itemName);
            this.NewCreatedItemIds.Add(itemId);

            #endregion

            #region Get the Inbox folder

            // Set the request's folderId field.
            folder.Id = DistinguishedFolderIdNameType.inbox;

            // GetFolder request.
            GetFolderType getInboxAfterMailReceivedRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folder);

            // Get the Inbox folder.
            GetFolderResponseType getInboxAfterMailReceivedResponse = this.FOLDAdapter.GetFolder(getInboxAfterMailReceivedRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxAfterMailReceivedResponse, 1, this.Site);

            // Variable to save the folder.
            allFolders = (FolderInfoResponseMessageType)getInboxAfterMailReceivedResponse.ResponseMessages.Items[0];
            folderInfo = (BaseFolderType)allFolders.Folders[0];
            FolderType folderType = (FolderType)folderInfo;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R99");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R99
            // Unread count value is returned from server, and schema is verified in adapter so this requirement can be covered.
            this.Site.CaptureRequirementIfIsNotNull(
                folderType.UnreadCount,
                99,
                @"[In t:FolderType Complex Type]The type of element UnreadCount is xs:int [XMLSCHEMA2].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R10010");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R10010
            // After sending a new message to mail box, the number of unread messages should be (count + 1).
            this.Site.CaptureRequirementIfAreEqual<int>(
                1 + count,
                folderType.UnreadCount,
                10010,
                @"[In t:FolderType Complex Type]This element[UnreadCount] MUST exist in responses.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R9901");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R9901
            this.Site.CaptureRequirementIfAreEqual<int>(
                1 + count,
                folderType.UnreadCount,
                9901,
                @"[In t:FolderType Complex Type]UnreadCount specifies the number of unread items in a folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R10011");

            // Since R9901 and R100 are captured successfully, this requirement will be captured.
            this.Site.CaptureRequirement(
                10011,
                @"[In t:FolderType Complex Type]This element[UnreadCount] MUST equal the sum of all MessageType complex types ([MS-OXWSMSG] section 2.2.4.1) and PostItemType complex types ([MS-OXWSPOST] section 2.2.4.1) that have the IsRead property set to ""false"".");
        }

        /// <summary>
        /// This test case verifies the requirements related to extended property of folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S04_TC08_GetExtendedProperty()
        {
            #region Create a new folder in the inbox folder

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Set folder class to null.
            createFolderRequest.Folders[0].FolderClass = null;

            // Set folder extended property and its value.
            PathToExtendedFieldType publishInAddressBook = new PathToExtendedFieldType();

            // Set extended property Id and type.
            publishInAddressBook.PropertyTag = "0x671E";
            publishInAddressBook.PropertyType = MapiPropertyTypeType.Boolean;

            ExtendedPropertyType pubAddressbook = new ExtendedPropertyType();
            pubAddressbook.ExtendedFieldURI = publishInAddressBook;

            // Set extended property value.
            pubAddressbook.Item = "1";

            ExtendedPropertyType[] extendedProperties = new ExtendedPropertyType[1];
            extendedProperties[0] = pubAddressbook;

            createFolderRequest.Folders[0].ExtendedProperty = extendedProperties;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R7801");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R7801
            // Folder created successfully with extended property this requirement can be captured.
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                createFolderResponse.ResponseMessages.Items[0].ResponseCode,
                7801,
                @"[In t:BaseFolderType Complex Type]This element [ExtendedProperty] is present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R980302");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R980302
            //  Folder created successfully without permission set so this requirement can be covered.
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                createFolderResponse.ResponseMessages.Items[0].ResponseCode,
                980302,
                @"[In t:FolderType Complex Type]This element [PermissionSet] is not present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R589202");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R589202
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                createFolderResponse.ResponseMessages.Items[0].ResponseCode,
                589202,
                @"[In t:BaseFolderType Complex Type]This element [FolderClass] is not present, server responses NO_ERROR.");

            #region Get the new created folder

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);
            getNewFolderRequest.FolderShape.AdditionalProperties = new BasePathToElementType[1];
            getNewFolderRequest.FolderShape.AdditionalProperties[0] = publishInAddressBook;

            // Get the new created folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);

            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0];
            FolderType folderInfo = (FolderType)allFolders.Folders[0];

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R423");

            // Folder ids is set in request and response is returned from server so this requirement can be captured.
            this.Site.CaptureRequirement(
                423,
                @"[In m:GetFolderType Complex Type]FolderIds is an array of one or more folder identifiers.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R426");

            // Folder id that contains id and change key value is set in request and response is returned from server so this requirement can be captured.
            this.Site.CaptureRequirement(
                426,
                @"[In t:NonEmptyArrayOfBaseFolderIdsType Complex Type]FolderId specifies the folder identifier and change key. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R77");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R77
            // Extended property value is returned from server, and schema is verified in adapter so this requirement can be covered.
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.ExtendedProperty,
                77,
                @"[In t:BaseFolderType Complex Type]The type of element ExtendedProperty is t:ExtendedPropertyType ([MS-OXWSXPROP] section 2.1.5).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R7701");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R7701 
            // In create operation this property value is set to 1 and type to Boolean this means the value is "true" in string type. 
            this.Site.CaptureRequirementIfAreEqual<string>(
                "true",
                ((ExtendedPropertyType)folderInfo.ExtendedProperty.GetValue(0)).Item.ToString(),
                7701,
                @"[In t:BaseFolderType Complex Type]ExtendedProperty specifies the set of extended properties on a folder.");
        }

        /// <summary>
        /// This test case verifies requirements related to getting folder with the base folder shape set to IdOnly.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S04_TC09_GetFolderIdOnly()
        {
            #region Get the contacts folder.

            DistinguishedFolderIdType folder = new DistinguishedFolderIdType();
            folder.Id = DistinguishedFolderIdNameType.contacts;

            // GetFolder request.
            GetFolderType getContactsFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.IdOnly, folder);

            // Get the Contacts folder.
            GetFolderResponseType getContactsFolderResponse = this.FOLDAdapter.GetFolder(getContactsFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getContactsFolderResponse, 1, this.Site);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R42103");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R42103
            bool isVerifiedR42103 = Common.IsIdOnly((XmlElement)this.FOLDAdapter.LastRawResponseXml, "t:ContactsFolder", "t:FolderId");
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR42103,
                42103,
                @"[In t:DefaultShapeNamesType Simple Type] A value of ""IdOnly"" [in DefaultShapeNamesType] specifies only the item or folder ID. include in the response.");
        }
        #endregion
    }
}