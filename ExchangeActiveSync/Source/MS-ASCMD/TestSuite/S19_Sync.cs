namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test the Sync command.
    /// </summary>
    [TestClass]
    public class S19_Sync : TestSuiteBase
    {
        #region Class initialize and clean up
        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test cases
        /// <summary>
        /// This test case is used to verify the requirements related to a successful Sync command.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC01_Sync_Success()
        {
            SyncResponse syncResponse = this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));
            string syncKey = this.LastSyncKey;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4599");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4599
            Site.CaptureRequirementIfIsNotNull(
                syncKey,
                4599,
                @"[In SyncKey(Sync)] The server sends a response that includes a new synchronization key value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4606");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4606
            Site.CaptureRequirementIfIsNotNull(
                syncKey,
                4606,
                @"[In SyncKey(Sync)] the server sends a new synchronization key value in its response to the client.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R648");

            // Verify MS-ASCMD requirement: MS-ASCMD_R648
            bool isVerifyR648 = CheckElementOfItemsChoiceType10(syncResponse, Response.ItemsChoiceType10.SyncKey);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR648,
                648,
                @"[In Sync] The server responds with an initial value of the synchronization key, which the client MUST then use to get the initial set of objects from the server.");

            #region Add a contact item.
            string contactFileAS = Common.GenerateResourceName(Site, "FileAS");
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);
            string originalClientId = addData.ClientId;
            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, addData);
            syncResponse = this.Sync(syncRequest);

            // The value of status is returned by Sync.
            Response.SyncCollectionsCollectionResponses syncCollectionResponse = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.IsNotNull(syncCollectionResponse, "The Response element should not be null.");
            Site.Assert.AreEqual<int>(1, int.Parse(syncCollectionResponse.Add[0].Status), "Status code 1 should be returned to indicate the Sync add command success.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, contactFileAS);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R957");

            // Verify MS-ASCMD requirement: MS-ASCMD_R957
            bool isVerifyR957 = !string.IsNullOrEmpty(syncCollectionResponse.Add[0].ClientId) && (syncCollectionResponse.Add[0].ClientId == originalClientId) && !string.IsNullOrEmpty(syncCollectionResponse.Add[0].ServerId);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR957,
                957,
                @"[In ClientId (Sync)] The server response contains an Add element that contains the original client ID and a new server ID that was assigned for the object, which replaces the client ID as the permanent object identifier.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R942");

            // Verify MS-ASCMD requirement: MS-ASCMD_R942
            Site.CaptureRequirementIfIsNull(
                syncCollectionResponse.Add[0].Class,
                942,
                @"[In Class(Sync)] The Class element is not included in Sync Add responses when the class of the collection matches the item class.");

            uint statusValue = Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status));

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4423");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4423
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                statusValue,
                4423,
                @"[In Status(Sync)] [When the scope is Global], [the cause of the status value 1 is] Server successfully completed command.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R748");

            // Verify MS-ASCMD requirement: MS-ASCMD_R748
            bool isVerifyR748 = !string.IsNullOrEmpty(syncCollectionResponse.Add[0].ClientId) && !string.IsNullOrEmpty(syncCollectionResponse.Add[0].ServerId);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR748,
                748,
                @"[In Add(Sync)] [When a new item is being sent from the client to the server] The server then responds with an Add element in a Responses element, which specifies the client ID and the server ID that was assigned to the new item.");
            #endregion

            #region Synchronize the change caused by the Add operation.
            this.FolderSync();
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId);
            this.Sync(syncRequest);

            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            Response.SyncCollectionsCollectionCommands syncCollectionCommands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            Site.Assert.IsNotNull(syncCollectionCommands, "The commands element should exist in the Sync response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R749");

            // Verify MS-ASCMD requirement: MS-ASCMD_R749
            Site.CaptureRequirementIfIsNotNull(
                syncCollectionCommands.Add,
                749,
                @"[In Add(Sync)] When the client sends a Sync command request to the server and a new item has been added to the server collection since the last synchronization, the server responds with an Add element in a Commands element.");

            // Send an empty request to server, and then receive an empty response from server.
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R656");

            // Verify MS-ASCMD requirement: MS-ASCMD_R656
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                syncResponse.ResponseDataXML,
                656,
                @"[In Sync] In such a case [there are no changes to any of the collections that are specified in the Sync request], the client can receive an empty response from the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R667");

            // Verify MS-ASCMD requirement: MS-ASCMD_R667
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                syncResponse.ResponseDataXML,
                667,
                @"[In Empty Sync Request] If no changes are detected on the server, the Sync response includes only HTTP headers, and no XML payload, and is referred to as an empty Sync response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R679");

            // Verify MS-ASCMD requirement: MS-ASCMD_R679
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                syncResponse.ResponseDataXML,
                679,
                @"[In Empty Sync Response] The server sends a Sync response (section 2.2.2.19) with only HTTP headers, and no XML payload, if there are no pending server changes to report to the client.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the MoreAvailable element will be returned in the Sync command response if there are more changes than the number that are requested in the WindowSize element.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC02_Sync_MoreAvailable()
        {
            // Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
            FolderCreateRequest folderCreateRequest = Common.CreateFolderCreateRequest(this.LastFolderSyncKey, (byte)FolderType.UserCreatedContacts, Common.GenerateResourceName(Site, "FolderCreate"), this.User1Information.ContactsCollectionId);
            FolderCreateResponse folderCreateResponse = this.CMDAdapter.FolderCreate(folderCreateRequest);
            Site.Assert.AreEqual<string>("1", folderCreateResponse.ResponseData.Status, "The server should return a status code 1 in the FolderCreate command response to indicate success.");

            // Record created folder collectionID
            string folderId = folderCreateResponse.ResponseData.ServerId;
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderId);

            // Synchronize the collection hierarchy and changes in a collection between the client and the server.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(folderId));

            #region Add two contact items
            string firstContactFileAs = Common.GenerateResourceName(Site, "FileAS", 1);
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", firstContactFileAs, null);
            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, folderId, addData);
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The first contact item should be added successfully.");

            this.FolderSync();

            string secondContactFileAs = Common.GenerateResourceName(Site, "FileAS", 2);
            addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", secondContactFileAs, null);
            syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, folderId, addData);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The second contact item should be added successfully.");
            #endregion

            #region Synchronize the changes in the new created folder.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(folderId);
            this.Sync(syncRequest);

            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncRequest.RequestData.WindowSize = "1";
            syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The Collections element should not be null in the Sync response.");
            bool isChecked = CheckElementOfItemsChoiceType10(syncResponse, Response.ItemsChoiceType10.MoreAvailable);

            #region Capture Code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4771");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4771
            Site.CaptureRequirementIfIsTrue(
                isChecked,
                4771,
                @"[In WindowSize] If the number of changes on the server is greater than the value of the WindowSize element, the server returns a MoreAvailable element in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5041");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5041
            Site.CaptureRequirementIfIsTrue(
                isChecked,
                5041,
                @"[In Synchronizing Inbox, Calendar, Contacts, and Tasks Folders] If more items remain to be synchronized, the airsync:MoreAvailable element (section 2.2.3.106) is returned in the Sync command response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4765");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4765
            Site.CaptureRequirementIfIsTrue(
                isChecked,
                4765,
                @"[In WindowSize] If the server does not send all the updates in a single message, the Sync response message contains the MoreAvailable element (section 2.2.3.106), which indicates that there are additional updates on the server to be downloaded to the client.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4779");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4779
            Site.CaptureRequirementIfIsTrue(
                isChecked,
                4779,
                @"[In WindowSize] When the server has filled the global WindowSize and collections that have changes did not fit in the response, the server can return a MoreAvailable element");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5015");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5015
            Site.CaptureRequirementIfIsTrue(
                isChecked,
                5015,
                @"[In Synchronizing a Folder Hierarchy] If the number of items returned is larger than the value specified by the airsync:WindowSize element, the airsync:MoreAvailable element (section 2.2.3.106) is returned in the Sync command response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3451");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3451
            Site.CaptureRequirementIfIsTrue(
                isChecked,
                3451,
                @"[In MoreAvailable] It[MoreAvailable element] appears only if the client request contained a WindowSize element and there are still changes to be returned to the client.");

            bool firstContactRetrieved = false;
            Response.SyncCollectionsCollectionCommands commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            if (commands != null)
            {
                Response.SyncCollectionsCollectionCommandsAddApplicationData applicationData =
                    commands.Add[0].ApplicationData;
                for (int i = 0; i < applicationData.ItemsElementName.Length; i++)
                {
                    if (applicationData.ItemsElementName[i] == Response.ItemsChoiceType8.FileAs &&
                        applicationData.Items[i].ToString() == firstContactFileAs)
                    {
                        firstContactRetrieved = true;
                        break;
                    }
                }
            }

            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The Collections element should not be null in the Sync response.");

            bool secondContactRetrieved = false;
            commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            if (commands != null)
            {
                Response.SyncCollectionsCollectionCommandsAddApplicationData applicationData = commands.Add[0].ApplicationData;
                for (int i = 0; i < applicationData.ItemsElementName.Length; i++)
                {
                    if (applicationData.ItemsElementName[i] == Response.ItemsChoiceType8.FileAs &&
                        applicationData.Items[i].ToString() == secondContactFileAs)
                    {
                        secondContactRetrieved = true;
                        break;
                    }
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5787");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5787
            bool isVerifyR5787 = isChecked && firstContactRetrieved && secondContactRetrieved;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR5787,
                5787,
                @"[In WindowSize] A WindowSize element value less than 100 can be useful if the client can display the initial set of objects while additional ones are still being retrieved from the server.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the Sync command sequence for folder synchronization.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC03_Sync_Email_Sequence()
        {
            #region Send a MIME-formatted e-mail from user1 to user2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            #region Synchronizes the changes in the Inbox folder.
            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);

            string olderSyncKey = this.LastSyncKey;
            SyncResponse syncResponse = this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items in the Sync response should not be null.");

            string syncKey = this.LastSyncKey;

            #region Capture Code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R654");

            // Verify MS-ASCMD requirement: MS-ASCMD_R654
            Site.CaptureRequirementIfAreNotEqual<string>(
                "0",
                syncKey,
                654,
                @"[In Sync] The server response also contains a synchronization key that is to be used for the next synchronization session for the folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5051");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5051
            Site.CaptureRequirementIfAreNotEqual<string>(
                "0",
                syncKey,
                5051,
                @"[In Synchronizing Inbox, Calendar, Contacts, and Tasks Folders] [Command sequence for folder synchronization, order 1:] The server responds with the synchronization key for the collection, to be used in successive synchronizations.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5053");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5053
            Site.CaptureRequirementIfAreNotEqual<string>(
                olderSyncKey,
                syncKey,
                5053,
                @"[In Synchronizing Inbox, Calendar, Contacts, and Tasks Folders] [Command sequence for folder synchronization, order 2*:] The server responds with new synchronization keys for each collection.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5032");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5032
            bool isVerifyR5032 = !string.IsNullOrEmpty(syncKey);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR5032,
                5032,
                @"[In Synchronizing Inbox, Calendar, Contacts, and Tasks Folders] In order to synchronize the content of each of the folders, an initial synchronization key for each folder MUST be obtained from the server.");

            #endregion

            #endregion

            #region Synchronize changes in the Inbox folder.
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User2Information.InboxCollectionId);
            this.Sync(syncRequest);

            syncRequest.RequestData.Collections[0].GetChanges = true;
            syncRequest.RequestData.Collections[0].GetChangesSpecified = true;
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            Response.SyncCollectionsCollectionCommands syncCollCommands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5060");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5060
            bool isVerified5060 = syncCollCommands.Add.Length > 0;
            Site.CaptureRequirementIfIsTrue(
                isVerified5060,
                5060,
                @"[In Synchronizing Inbox, Calendar, Contacts, and Tasks Folders] [Command sequence for folder synchronization, order 4*:] The server responds with airsync:Add (section 2.2.3.7.2), airsync:Change (section 2.2.3.23), or airsync:Delete (section 2.2.3.40.2) elements for items in the collection.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if the synchronization key is invalid, then the status code in the server response will be 3.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC04_Sync_Status3()
        {
            // Synchronize the changes with a request containing an invalid SyncKey.
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            string invalidSyncKey = Guid.NewGuid().ToString();
            syncRequest.RequestData.Collections[0].SyncKey = invalidSyncKey;
            SyncResponse syncResponse = this.Sync(syncRequest);
            uint statusCode = Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status));

            #region Capture Code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4415");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4415
            Site.CaptureRequirementIfAreEqual<uint>(
                3,
                statusCode,
                4415,
                @"[In Status(Sync)] If the[Sync command] request failed, the Status element contains a code that indicates the type of failure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4417");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4417
            Site.CaptureRequirementIfAreEqual<uint>(
                3,
                statusCode,
                4417,
                @"[In Status(Sync)] If the[Sync command] operation [on the collection] failed, the Status element contains a code that indicates the type of failure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4425");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4425
            Site.CaptureRequirementIfAreEqual<uint>(
                3,
                statusCode,
                4425,
                @"[In Status(Sync)] [When the scope is Global], [the cause of the status value 3 is] Invalid or mismatched synchronization key.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if the request does not comply with the specification requirements, then the status code in the server response will be 4.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC05_Sync_Status4()
        {
            SyncRequest syncRequest = new SyncRequest();
            SyncResponse syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4430");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4430
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(syncResponse.ResponseData.Status),
                4430,
                @"[In Status(Sync)] [When the scope is Item], [the cause the status value 4 is] There was a semantic error in the synchronization request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4431");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4431
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(syncResponse.ResponseData.Status),
                4431,
                @"[In Status(Sync)] [When the scope is Item], [the cause of the status value 4 is] The client is issuing a request that does not comply with the specification requirements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5778");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5778
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(syncResponse.ResponseData.Status),
                5778,
                @"[In Status(Sync)] [When the scope is Global], [the cause the status value 4 is] There was a semantic error in the synchronization request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5779");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5779
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(syncResponse.ResponseData.Status),
                5779,
                @"[In Status(Sync)] [When the scope is Global], [the cause of the status value 4 is] The client is issuing a request that does not comply with the specification requirements.");

            #region Add a contact item to the recipient information cache.
            string contactFileAS = Common.GenerateResourceName(Site, "FileAS");
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);
            syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.RecipientInformationCacheCollectionId, addData);
            syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R756");

            // Verify MS-ASCMD requirement: MS-ASCMD_R756
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(syncResponse.ResponseData.Status),
                756,
                @"[In Add(Sync)] If a client attempts to add an item to the recipient information cache, a Status element with a value of 4 is returned as a child of the Sync element.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if an empty Sync command request is received and the cached set of notify-able collections is missing, then the status code in the server response will be 13.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC06_Sync_Status13()
        {
            SyncResponse syncResponse = this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId));
            Site.Assert.AreEqual<string>("1", this.GetStatusCode(syncResponse.ResponseDataXML), "The Status value of the Sync command should be 1, when the RequestData element in the Sync command request is not null.");

            // Synchronize the changes with a request, of which the request data are null.
            SyncRequest syncRequest = new SyncRequest { RequestData = null };
            syncResponse = this.Sync(syncRequest, false);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4456");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4456
            Site.CaptureRequirementIfAreEqual<int>(
                13,
                int.Parse(syncResponse.ResponseData.Status),
                4456,
                @"[In Status(Sync)] [When the scope is Item], [the cause of the status value 13 is] An empty or partial Sync command request is received and the cached set of notify-able collections is missing.");
        }

        /// <summary>
        /// This test case is used to verify Sync command, if the HeartbeatInterval element is outside the range set, or smaller than the minimum allowable value, then the status code in the server response will be 14.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC07_Sync_HeartbeatInterval_Status14()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The HeartbeatInterval tag is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Verify the upper and lower bounds of the value of the HeartbeatInterval element.
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            syncRequest.RequestData.HeartbeatInterval = "3541";
            SyncResponse syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4460");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4460
            bool isVerify4460 = int.Parse(syncResponse.ResponseData.Status) == 14 && syncResponse.ResponseData.Item.ToString() == "3540" && syncResponse.ResponseDataXML.ToString().Contains("Limit");
            Site.CaptureRequirementIfIsTrue(
                isVerify4460,
                4460,
                @"[In Status(Sync)] [When the scope is Item], [the cause of the status value 14 is] If the HeartbeatInterval element value [or Wait element value] included in the Sync request is larger than the maximum allowable value, the response contains a Limit element that specifies the maximum allowed value.");

            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            syncRequest.RequestData.HeartbeatInterval = "59";
            syncResponse = this.Sync(syncRequest);
            bool isValidResponse = int.Parse(syncResponse.ResponseData.Status) == 14 && syncResponse.ResponseData.Item.ToString() == "60" && syncResponse.ResponseDataXML.ToString().Contains("Limit");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4461");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4461
            Site.CaptureRequirementIfIsTrue(
                isValidResponse,
                4461,
                @"[In Status(Sync)] [When the scope is Item], [the cause of the status value 14 is] If the HeartbeatInterval element value [or Wait value] included in the Sync request is smaller than the minimum allowable value, the response contains a Limit element that specifies the minimum allowed value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3225");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3225
            // According to above two steps, this requirement can be covered directly.
            Site.CaptureRequirement(
                3225,
                @"[In Limit] A status code 14 indicates that the Limit element specifies the minimum or maximum wait-interval that is acceptable.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5780");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5780
            // According to above two steps, this requirement can be covered directly.
            Site.CaptureRequirement(
                5780,
                @"[In Status(Sync)] [When the scope is Item], [the cause of the status value 14 is] The Sync request was processed successfully but the [wait interval (Wait element value (section 2.2.3.182)) or] heartbeat interval (HeartbeatInterval element value (section 2.2.3.79.2)) that is specified by the client is outside the range set by the server administrator.");
            #endregion
        }

        /// <summary>
        /// This test cases is used to verify the server will return a status value 4 if call Sync command with ConversationMode element for collections that do not store e-mails.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC08_Sync_ConversationMode_Status4()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The ConversationMode element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId);
            this.Sync(syncRequest);

            Request.Options option = new Request.Options
            {
                Items = new object[] { (byte)1 },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.FilterType }
            };

            Request.SyncCollection collection = new Request.SyncCollection
            {
                SyncKey = this.LastSyncKey,
                Options = new Request.Options[] { option },
                CollectionId = this.User1Information.ContactsCollectionId,
                Commands = null,
                GetChanges = true,
                GetChangesSpecified = true,
                ConversationMode = true,
                ConversationModeSpecified = true
            };

            syncRequest.RequestData.Collections = new Request.SyncCollection[] { collection };
            SyncResponse syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2117");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2117
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(syncResponse.ResponseData.Status),
                2117,
                @"[In ConversationMode(Sync)] Specifying the ConversationMode element for collections that do not store emails results in an invalid XML error, Status element (section 2.2.3.162.16) value 4.");
        }

        /// <summary>
        /// This test case is used to verify Sync command, if options for the same Class within the same Collection are redefined, a Status element value of 4 is returned.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC09_Sync_Class_Redefined()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Class element is not supported as a child element of the Options element when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId));

            // Add a calendar item
            string calendarSubject = Common.GenerateResourceName(Site, "canlendarSubject");
            string calendarTo = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string location = Common.GenerateResourceName(Site, "Room11");
            Request.SyncCollectionAdd addData = this.CreateAddCalendarCommand(calendarTo, calendarSubject, location, string.Empty);

            Request.Options firstOption = new Request.Options
            {
                Items = new object[] { "Calendar" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            Request.Options secondOption = new Request.Options
            {
                Items = new object[] { "Calendar" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, addData);
            syncRequest.RequestData.Collections[0].Options = new Request.Options[] { firstOption, secondOption };
            SyncResponse syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R930");

            // Verify MS-ASCMD requirement: MS-ASCMD_R930
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(syncResponse.ResponseData.Status),
                930,
                @"[In Class(Sync)] A Status element (section 2.2.3.162.16) value of 4 is returned if options for the same Class within the same Collection are redefined.");
        }

        /// <summary>
        /// This test case is used to verify Sync command, if no child elements for the Commands element are defined, server does not return a status error.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC10_Sync_NoChildElementsForCommands()
        {
            // Synchronize the changes by sending a request containing a null Commands element.
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            SyncResponse syncResponse = this.Sync(syncRequest);

            if (Common.IsRequirementEnabled(2058, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2058");

                // Verify MS-ASCMD requirement: MS-ASCMD_R2058
                Site.CaptureRequirementIfAreEqual<uint>(
                    1,
                    Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                    2058,
                    @"[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [no child elements for the Commands element]. (Exchange 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to verify if Sync command for e-mail, the status should be correspond to the value of FilterType.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC11_Sync_Email_FilterType()
        {
            #region Call Sync command to verify server supports to filter email when FilterType set to 0.
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId, (byte)0);
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3035");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3035
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                3035,
                @"[In FilterType(Sync)] Yes. [Applies to Email, if FilterType is 0, Status element value is 1.]");

            #endregion

            #region Call Sync command to verify server supports to filter email when FilterType set to 1.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId, (byte)1);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3036");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3036
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                3036,
                @"[In FilterType(Sync)] Yes. [Applies to Email, if FilterType is 1, Status element value is 1.]");
            #endregion

            #region Call Sync command to verify server supports to filter email when FilterType set to 2.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId, (byte)2);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3037");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3037
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                3037,
                @"[In FilterType(Sync)] Yes. [Applies to Email, if FilterType is 2, Status element value is 1.]");
            #endregion

            #region Call Sync command to verify server supports to filter email when FilterType set to 3.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId, (byte)3);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3038");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3038
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                3038,
                @"[In FilterType(Sync)] Yes. [Applies to Email, if FilterType is 3, Status element value is 1.]");
            #endregion

            #region Call Sync command to verify server supports to filter email when FilterType set to 4.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId, (byte)4);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3039");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3039
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                3039,
                @"[In FilterType(Sync)] Yes. [Applies to Email, if FilterType is 4, Status element value is 1.]");
            #endregion

            #region Call Sync command to verify server supports to filter email when FilterType set to 5.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId, (byte)5);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3040");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3040
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                3040,
                @"[In FilterType(Sync)] Yes. [Applies to Email, if FilterType is 5, Status element value is 1.]");
            #endregion

            #region Call Sync command to verify server does not support to filter email when FilterType set to 6.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId, (byte)6);
            syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3041");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3041
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3041,
                @"[In FilterType(Sync)] No, [Applies to email, if FilterType is 6, status is not 1.]");
            #endregion

            #region Call Sync command to verify server does not support to filter email when FilterType set to 7.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId, (byte)7);
            syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3042");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3042
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3042,
                @"[In FilterType(Sync)] No, [Applies to email, if FilterType is 7, status is not 1.]");
            #endregion

            #region Call Sync command to verify server does not support to filter email when FilterType set to 8.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId, (byte)8);
            syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3043");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3043
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3043,
                @"[In FilterType(Sync)] No, [Applies to email, if FilterType is 8, status is not 1.]");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3080");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3080
            Site.CaptureRequirementIfAreEqual<string>(
                "4",
                syncResponse.ResponseData.Status,
                3080,
                @"[In FilterType(Sync)] The server returns a Status element (section 2.2.3.162.16) value of 4 if a FilterType element value of 8 is included in an email Sync request.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if Sync command for Calendar, the status should be correspond to the value of FilterType. 
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC12_Sync_Calendar_FilterType()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Class element is not supported in a Sync command response when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Recurrences cannot be added in protocol version 16.0");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Recurrences cannot be added in protocol version 16.1");
           
            #region Add a new calendar
            string calendarSubject = Common.GenerateResourceName(Site, "calendarSubject");
            DateTime startTime = DateTime.Now.AddDays(1.0);
            DateTime endTime = startTime.AddMinutes(10.0);

            Request.SyncCollectionAdd calendarData = new Request.SyncCollectionAdd
            {
                ClientId = TestSuiteBase.ClientId,
                ApplicationData =
                    new Request.SyncCollectionAddApplicationData
                    {
                        ItemsElementName =
                            new Request.ItemsChoiceType8[]
                            {
                                Request.ItemsChoiceType8.Subject, 
                                Request.ItemsChoiceType8.StartTime, 
                                Request.ItemsChoiceType8.EndTime
                            },
                        Items =
                            new object[]
                            {
                                calendarSubject, 
                                startTime.ToString("yyyyMMddTHHmmssZ"),
                                endTime.ToString("yyyyMMddTHHmmssZ")
                             }
                    },
                Class = "Calendar"
            };

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId));

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, calendarData);
            SyncResponse syncResponse = this.Sync(syncRequest);

            Response.SyncCollectionsCollectionResponses responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The calendar should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, calendarSubject);
            #endregion

            #region Call Sync command to verify server supports to filter calendar when FilterType set to 0.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 0);
            this.Sync(syncRequest);

            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3044");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3044
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                3044,
                @"[In FilterType(Sync)] Yes. [Applies to calendar, if FilterType is 0, Status element value is 1.]");

            Response.SyncCollectionsCollectionCommands commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            int itemsWithFilter = commands.Add.Length;

            syncResponse = this.SyncChanges(this.User1Information.CalendarCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            int itemsWithoutFilter = commands.Add.Length;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3076");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3076
            Site.CaptureRequirementIfAreEqual<int>(
                itemsWithFilter,
                itemsWithoutFilter,
                3076,
                @"[In FilterType(Sync)] If the FilterType element is omitted, all objects are sent from the server without regard for their age.");
            #endregion

            #region Call Sync command to verify server does not support to filter calendar when FilterType set to 1.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 1);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3045");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3045
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3045,
                @"[In FilterType(Sync)] No, [Applies to calendar, if FilterType is 1, status is not 1.]");

            #endregion

            #region Call Sync command to verify server does not support to filter calendar when FilterType set to 2.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 2);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3046");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3046
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3046,
                @"[In FilterType(Sync)] No, [Applies to calendar, if FilterType is 2, status is not 1.]");

            #endregion

            #region Call Sync command to verify server does not support to filter calendar when FilterType set to 3.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 3);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3047");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3047
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3047,
                @"[In FilterType(Sync)] No, [Applies to calendar, if FilterType is 3, status is not 1.]");

            #endregion

            #region Call Sync command to verify server supports to filter calendar when FilterType set to 4.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 4);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3048");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3048
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                3048,
                @"[In FilterType(Sync)] Yes, [Applies to calendar, if FilterType is 4, Status element value is 1.]");

            // Create a future calendar
            this.GetInitialSyncResponse(this.User1Information.CalendarCollectionId);
            calendarSubject = Common.GenerateResourceName(this.Site, "canlendarSubject");
            startTime = DateTime.Now.AddDays(15.0);
            endTime = startTime.AddMinutes(10.0);

            calendarData = new Request.SyncCollectionAdd
            {
                ClientId = TestSuiteBase.ClientId,
                ApplicationData =
                    new Request.SyncCollectionAddApplicationData
                    {
                        ItemsElementName =
                            new Request.ItemsChoiceType8[]
                            {
                                Request.ItemsChoiceType8.Subject,
                                Request.ItemsChoiceType8.StartTime, 
                                Request.ItemsChoiceType8.EndTime
                            },
                        Items =
                            new object[]
                            {
                                calendarSubject, 
                                startTime.ToString("yyyyMMddTHHmmssZ"),
                                endTime.ToString("yyyyMMddTHHmmssZ")
                            }
                    },
                Class = "Calendar"
            };

            syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, calendarData);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The calendar should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, calendarSubject);

            syncResponse = this.SyncChanges(this.User1Information.CalendarCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            bool isVerifyR3065 = !string.IsNullOrEmpty(TestSuiteBase.FindServerId(syncResponse, "Subject", calendarSubject));

            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 4);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            isVerifyR3065 = isVerifyR3065 && !string.IsNullOrEmpty(TestSuiteBase.FindServerId(syncResponse, "Subject", calendarSubject));

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3065");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3065
            Site.CaptureRequirementIfIsTrue(
                isVerifyR3065,
                3065,
                @"[In FilterType(Sync)] Calendar items that are in the future [or that have recurrence but no end date] are sent to the client regardless of the FilterType element value.");

            // create a recurrence calendar without EndTime
            Request.Recurrence recurrence = new Request.Recurrence
            {
                Type = 1,
                OccurrencesSpecified = false,
                DayOfWeek = 2,
                DayOfWeekSpecified = true,
                IsLeapMonthSpecified = false
            };

            string recurrenceCalendarSubject = Common.GenerateResourceName(Site, "recurrenceCanlendarSubject");

            Request.SyncCollectionAdd recurrenceCalendarData = new Request.SyncCollectionAdd
            {
                ClientId = TestSuiteBase.ClientId,
                ApplicationData =
                    new Request.SyncCollectionAddApplicationData
                    {
                        ItemsElementName =
                            new Request.ItemsChoiceType8[] { Request.ItemsChoiceType8.Subject, Request.ItemsChoiceType8.Recurrence, Request.ItemsChoiceType8.UID },
                        Items = new object[] { recurrenceCalendarSubject, recurrence, Guid.NewGuid().ToString() }
                    },
                Class = "Calendar"
            };

            syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, recurrenceCalendarData);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The calendar should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, recurrenceCalendarSubject);

            syncResponse = this.SyncChanges(this.User1Information.CalendarCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            bool isVerifyR5878 = !string.IsNullOrEmpty(TestSuiteBase.FindServerId(syncResponse, "Subject", recurrenceCalendarSubject));

            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 4);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            isVerifyR5878 = isVerifyR5878 && !string.IsNullOrEmpty(TestSuiteBase.FindServerId(syncResponse, "Subject", recurrenceCalendarSubject));

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5878");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5878
            Site.CaptureRequirementIfIsTrue(
                isVerifyR5878,
                5878,
                @"[In FilterType(Sync)] Calendar items [that are in the future or] that have recurrence but no end date are sent to the client regardless of the FilterType element value.");
            #endregion

            #region Call Sync command to verify server supports to filter calendar when FilterType set to 5.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 5);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3049");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3049
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                3049,
                @"[In FilterType(Sync)] Yes. [Applies to calendar, if FilterType is 5, Status element value is 1.]");
            #endregion

            #region Call Sync command to verify server supports to filter calendar when FilterType set to 6.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 6);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3050");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3050
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                3050,
                @"[In FilterType(Sync)] Yes. [Applies to calendar, if FilterType is 6, Status element value is 1.]");
            #endregion

            #region Call Sync command to verify server supports to filter calendar when FilterType set to 7.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 7);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3051");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3051
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                3051,
                @"[In FilterType(Sync)] Yes. [Applies to calendar, if FilterType is 7, Status element value is 1.]");

            commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            int original = commands.Add.Length;

            // Create a overdue calendar
            this.GetInitialSyncResponse(this.User1Information.CalendarCollectionId);
            calendarSubject = Common.GenerateResourceName(this.Site, "canlendarSubject");
            startTime = DateTime.Now.AddMonths(-7);
            endTime = startTime.AddHours(1.0);

            calendarData = new Request.SyncCollectionAdd
            {
                ClientId = TestSuiteBase.ClientId,
                ApplicationData =
                    new Request.SyncCollectionAddApplicationData
                    {
                        ItemsElementName =
                            new Request.ItemsChoiceType8[]
                            {
                                Request.ItemsChoiceType8.Subject,
                                   Request.ItemsChoiceType8.StartTime,
                                   Request.ItemsChoiceType8.EndTime,
                            },
                        Items =
                            new object[]
                            {
                                calendarSubject, 
                                startTime.ToString("yyyyMMddTHHmmssZ"),
                                endTime.ToString("yyyyMMddTHHmmssZ")
                            }
                    },
                Class = "Calendar"
            };

            syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, calendarData);
            syncResponse = this.Sync(syncRequest);
            responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The calendar should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, calendarSubject);

            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 7);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            int current = commands.Add.Length;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2987");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2987
            Site.CaptureRequirementIfAreEqual<int>(
                original,
                current,
                2987,
                @"[In FilterType(Sync)] If a FilterType element is specified, the server sends only objects that are dated within the specified time window.");
            #endregion

            #region Call Sync command to verify server does not support to filter calendar when FilterType set to 8.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 8);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3052");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3052
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3052,
                @"[In FilterType(Sync)] No, [The result of including a FilterType element value of 8 for a Calendar item is undefined.]");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if Sync command for Tasks, the status should be correspond to the value of FilterType. 
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC13_Sync_Tasks_FilterType()
        {
            #region Call Sync command to verify server supports to filter task when FilterType set to 0.
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.TasksCollectionId, 0);
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3053");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3053
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                3053,
                @"[In FilterType(Sync)] Yes. [Applies to tasks, if FilterType is 0, Status element value is 1.]");
            #endregion

            #region Call Sync command to verify server does not support to filter task when FilterType set to 1.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.TasksCollectionId, 1);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData, "The ResponseData should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3054");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3054
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3054,
                @"[In FilterType(Sync)] No, [Applies to tasks, if FilterType is 1, status is not 1.]");
            #endregion

            #region Call Sync command to verify server does not support to filter task when FilterType set to 2.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.TasksCollectionId, 2);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData, "The ResponseData should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3055");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3055
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3055,
                @"[In FilterType(Sync)] No, [Applies to tasks, if FilterType is 2, status is not 1.]");

            #endregion

            #region Call Sync command to verify server does not support to filter task when FilterType set to 3.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.TasksCollectionId, 3);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData, "The ResponseData should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3056");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3056
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3056,
                @"[In FilterType(Sync)] No, [Applies to tasks, if FilterType is 3, status is not 1.]");

            #endregion

            #region Call Sync command to verify server does not support to filter task when FilterType set to 4.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.TasksCollectionId, 4);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData, "The ResponseData should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3057");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3057
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3057,
                @"[In FilterType(Sync)] No, [Applies to tasks, if FilterType is 4, status is not 1.]");

            #endregion

            #region Call Sync command to verify server does not support to filter task when FilterType set to 5.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.TasksCollectionId, 5);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData, "The ResponseData should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3058");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3058
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3058,
                @"[In FilterType(Sync)] No, [Applies to tasks, if FilterType is 5, status is not 1.]");

            #endregion

            #region Call Sync command to verify server does not support to filter task when FilterType set to 6.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.TasksCollectionId, 6);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData, "The ResponseData should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3059");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3059
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3059,
                @"[In FilterType(Sync)] No, [Applies to tasks, if FilterType is 6, status is not 1.]");

            #endregion

            #region Call Sync command to verify server does not support to filter task when FilterType set to 7.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.TasksCollectionId, 7);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData, "The ResponseData should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3060");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3060
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3060,
                @"[In FilterType(Sync)] No, [Applies to tasks, if FilterType is 7, status is not 1.]");
            #endregion

            #region Call Sync command to verify server supports to filter task when FilterType set to 8.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.TasksCollectionId, 8);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData, "The ResponseData should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3061");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3061
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncResponse.ResponseData.Status,
                3061,
                @"[In FilterType(Sync)] Yes. [Applies to tasks, if FilterType is 8, Status element value is 1.]");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the server does not return a protocol status error if including more than one FilterType elements in Sync command request.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC14_Sync_MoreThanOneFilterTypes()
        {
            #region Send a MIME-formatted email to User2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            this.SwitchUser(this.User2Information);

            this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);

            Request.Options option = new Request.Options
            {
                Items = new object[] { (byte)0, (byte)1 },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.FilterType, Request.ItemsChoiceType1.FilterType }
            };

            SyncResponse syncResponse = this.SyncChangesWithOption(this.User2Information.InboxCollectionId, option);
            Site.Assert.IsNotNull(syncResponse.ResponseData, "The ResponseData should not be null.");
            Site.Assert.IsNotNull(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject), "The email should be in the Inbox folder of User2.");

            if (Common.IsRequirementEnabled(3075, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3075");

                // Verify MS-ASCMD requirement: MS-ASCMD_R3075
                Site.CaptureRequirementIfAreEqual<uint>(
                    1,
                    Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                    3075,
                    @"[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one FilterType element as the child of the Options element ]. (Exchange 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to verify Sync command for contacts, if FilterType element is not included in a contact Sync request, no error is thrown. 
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC15_Sync_IncludeFilterTypeOrNot()
        {
            // Synchronize changes without FilterType value.
            SyncResponse syncResponseWithoutFilterType = this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponseWithoutFilterType, Response.ItemsChoiceType10.Status)), "The Status code of the Sync response without FilterType should be 1.");

            // Synchronize changes with a FilterType value.
            SyncResponse syncResponseWithFilterType = this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId, (byte)0));
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponseWithFilterType, Response.ItemsChoiceType10.Status)), "The Status code of the Sync response with FilterType should be 1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3079");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3079
            Site.CaptureRequirementIfAreEqual<string>(
                syncResponseWithoutFilterType.ResponseData.Status,
                syncResponseWithFilterType.ResponseData.Status,
                3079,
                @"[In FilterType(Sync)] Reply is the same whether FilterType element is or not included in a contact Sync request, and no error is thrown.");
        }

        /// <summary>
        /// This test case is used to verify Sync command, if no additional changes are available, MoreAvailable element is omitted.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC16_Sync_WindowSize_MoreAvailable_Omitted()
        {
            // Call method FolderSync to synchronize the collection hierarchy.
            FolderSyncResponse folderSyncResponse = this.FolderSync();

            // Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
            FolderCreateRequest folderCreateRequest = Common.CreateFolderCreateRequest(folderSyncResponse.ResponseData.SyncKey, (byte)FolderType.UserCreatedContacts, Common.GenerateResourceName(Site, "FolderCreate"), this.User1Information.ContactsCollectionId);
            FolderCreateResponse folderCreateResponse = this.CMDAdapter.FolderCreate(folderCreateRequest);
            Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderCreate command response to indicate success.");

            // Record created folder collectionID
            string folderId = folderCreateResponse.ResponseData.ServerId;
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderId);

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(folderId));

            #region Add a contact item with Sync operation.
            string contactFileAS = Common.GenerateResourceName(Site, "FileAS");
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, folderId, addData);
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Response.SyncCollectionsCollectionResponses responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The status code of Sync add operation should be 1.");
            #endregion

            #region Synchronize the changes in the Contacts folder.
            this.FolderSync();
            syncResponse = this.SyncChanges(folderId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, folderId, contactFileAS);

            // Check whether the MoreAvailable element appears in response.
            bool isVerifyR3456 = CheckElementOfItemsChoiceType10(syncResponse, Response.ItemsChoiceType10.MoreAvailable);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3456");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3456
            Site.CaptureRequirementIfIsFalse(
                isVerifyR3456,
                3456,
                @"[In MoreAvailable] It[ MoreAvailable element] is omitted if no additional changes are available.");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if the WindowSize element is omitted, the server behaves as if a WindowSize element with a value of 100 were submitted.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC17_Sync_WindowSize_MoreAvailable()
        {
            // Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
            FolderCreateRequest folderCreateRequest = Common.CreateFolderCreateRequest(this.LastFolderSyncKey, (byte)FolderType.UserCreatedContacts, Common.GenerateResourceName(Site, "FolderCreate"), this.User1Information.ContactsCollectionId);
            FolderCreateResponse folderCreateResponse = this.CMDAdapter.FolderCreate(folderCreateRequest);
            Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderCreate command response to indicate success.");

            // Record created folder collectionID
            string folderId = folderCreateResponse.ResponseData.ServerId;
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderId);

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(folderId));

            string contactFileAS = Common.GenerateResourceName(Site, "FileAS", 1);
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, folderId, addData);
            this.Sync(syncRequest);

            this.FolderSync();

            // Call method Sync to add 101 contact items.
            for (int i = 1; i < 101; i++)
            {
                contactFileAS = Common.GenerateResourceName(this.Site, "FileAS", (uint)(i + 1));
                addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);

                syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, folderId, addData);
                this.Sync(syncRequest);
            }

            #region Synchronize the changes in the Contacts folder.
            SyncResponse syncResponse = this.SyncChanges(folderId, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The Collections element should exist in the Sync response.");

            Response.SyncCollectionsCollectionCommands collectionCommands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;

            bool isVerifyR3459 = CheckElementOfItemsChoiceType10(syncResponse, Response.ItemsChoiceType10.MoreAvailable) && collectionCommands.Add.Length == 100;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3459");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3459
            Site.CaptureRequirementIfIsTrue(
                isVerifyR3459,
                3459,
                @"[In MoreAvailable] If the WindowSize element is omitted, the server behaves as if a WindowSize element with a value of 100 was submitted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4767");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4767
            Site.CaptureRequirementIfIsTrue(
                isVerifyR3459,
                4767,
                @"[In WindowSize] If the WindowSize element is omitted, the server behaves as if a WindowSize element with a value of 100 were submitted.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the Sync command, when handling S/MIME content in the response, the server MUST include the airsyncbase:Body element, which is a child of the ApplicationData element.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC18_Sync_SMIME()
        {
            #region Call method SendMail to send MIME-formatted e-mail messages to the server.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            this.SwitchUser(this.User2Information);

            #region Synchronize the changes in the Inbox folder.
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User2Information.InboxCollectionId);
            this.Sync(syncRequest);

            Request.BodyPreference bodyPreference = new Request.BodyPreference { Type = 4 };

            Request.Options options = new Request.Options
            {
                Items = new object[] { bodyPreference, (byte)1 },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.BodyPreference, Request.ItemsChoiceType1.MIMESupport }
            };

            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncRequest.RequestData.Collections[0].Options = new Request.Options[] { options };
            syncRequest.RequestData.Collections[0].GetChanges = true;
            syncRequest.RequestData.Collections[0].GetChangesSpecified = true;
            SyncResponse syncResponse = this.Sync(syncRequest);

            int counter = 0;
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject);
            while ((counter < retryCount) && string.IsNullOrEmpty(serverId))
            {
                System.Threading.Thread.Sleep(waitTime);
                syncResponse = this.Sync(syncRequest);
                serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject);
                counter++;
            }

            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            Response.SyncCollectionsCollectionCommands syncCollectionsCollectionCommands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            bool isR3408Satisfied = false;
            foreach (Response.ItemsChoiceType8 element in syncCollectionsCollectionCommands.Add[0].ApplicationData.ItemsElementName)
            {
                if (element == Response.ItemsChoiceType8.Body)
                {
                    isR3408Satisfied = true;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3408");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3408
            Site.CaptureRequirementIfIsTrue(
                isR3408Satisfied,
                3408,
                @"[In MIMESupport(Sync)] When handling S/MIME content in the response, the server MUST include the airsyncbase:Body element ([MS-ASAIRS] section 2.2.2.4), which is a child of the ApplicationData element (section 2.2.3.11).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if the client requested server changes but had no changes to send to the server, the Response element is omitted.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC19_Sync_NoResponsesElement()
        {
            // Call method FolderSync to synchronize the collection hierarchy.
            FolderSyncResponse folderSyncResponse = this.FolderSync();

            // Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
            FolderCreateRequest folderCreateRequest = Common.CreateFolderCreateRequest(folderSyncResponse.ResponseData.SyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderCreate"), this.User1Information.InboxCollectionId);
            FolderCreateResponse folderCreateResponse = this.CMDAdapter.FolderCreate(folderCreateRequest);
            Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderCreate command response to indicate success.");

            // Record created folder collectionID
            string folderId = folderCreateResponse.ResponseData.ServerId;
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderId);

            SyncRequest request = TestSuiteBase.CreateEmptySyncRequest(folderId);
            this.Sync(request);
            request.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            SyncResponse syncResponse = this.Sync(request);
            Response.SyncCollectionsCollectionResponses collectionResponse = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3835");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3835
            // When there are no changes on the server, The Responses object in the Sync response is null, which means the Sync response does not contain the Responses element.
            Site.CaptureRequirementIfIsNull(
                collectionResponse,
                3835,
                @"[In Responses] It[ Responses element] is omitted otherwise (for example, if the client requested server changes but had no changes to send to the server).");
        }

        /// <summary>
        /// This test case is used to verify the server does not return a protocol status error if call Sync command with more than one Class elements.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC20_Sync_MoreThanOneClasses()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Class element is not supported as a child element of the Options element when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send a MIME-formatted email to User2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null);

            Request.Options option = new Request.Options
            {
                Items = new object[] { "Email", "Email" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class, Request.ItemsChoiceType1.Class }
            };

            SyncResponse syncResponse = this.SyncChangesWithOption(this.User2Information.InboxCollectionId, option);

            Site.Assert.IsNotNull(syncResponse.ResponseData, "The ResponseData should not be null.");
            Site.Assert.IsTrue(!string.IsNullOrEmpty(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject)), "The email should be in the Inbox folder of User2.");

            if (Common.IsRequirementEnabled(938, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R938");

                // Verify MS-ASCMD requirement: MS-ASCMD_R938
                Site.CaptureRequirementIfAreEqual<uint>(
                    1,
                    Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                    938,
                    @"[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one Class element as child elements of the Options element]. (Exchange 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to verify Sync command, if request includes more than one MaxItem element, the server does not return error.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC21_Sync_MoreThanOneMaxItems()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The MaxItems element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            Request.Options option = new Request.Options
            {
                Items = new object[] { "2", "3" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.MaxItems, Request.ItemsChoiceType1.MaxItems }
            };

            SyncResponse syncResponse = this.SyncChangesWithOption(this.User1Information.RecipientInformationCacheCollectionId, option);

            if (Common.IsRequirementEnabled(3280, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3280");

                // Verify MS-ASCMD requirement: MS-ASCMD_R3280
                Site.CaptureRequirementIfAreEqual<uint>(
                    1,
                    Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                    3280,
                    @"[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one MaxItems element as the child element of the Options element]. (Exchange 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to verify Sync command, if request includes more than one MIMESupport element, the server does not return error.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC22_Sync_MoreThanOneMIMESupport()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Options element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send a MIME-formatted email to User2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null);

            Request.Options option = new Request.Options
            {
                Items = new object[] { (byte)0, (byte)1 },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.MIMESupport, Request.ItemsChoiceType1.MIMESupport }
            };

            SyncResponse syncResponse = this.SyncChangesWithOption(this.User2Information.InboxCollectionId, option);
            Site.Assert.IsNotNull(syncResponse.ResponseData, "The ResponseData should not be null.");
            Site.Assert.IsNotNull(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject), "The email should be in the Inbox folder of User2.");

            if (Common.IsRequirementEnabled(3404, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3404");

                // Verify MS-ASCMD requirement: MS-ASCMD_R3404
                Site.CaptureRequirementIfAreEqual<uint>(
                    1,
                    Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                    3404,
                    @"[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one MIMESupport element as the child element of the Options element]. (Exchange 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to verify Sync command, if request includes more than one MIMETruncation elements, the server does not return error.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC23_Sync_MoreThanOneMIMETruncation()
        {
            #region Send a MIME-formatted email to User2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null);

            // Call Sync with two MIMETruncation options
            Request.Options option = new Request.Options
            {
                Items = new object[] { (byte)2, (byte)3 },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.MIMETruncation, Request.ItemsChoiceType1.MIMETruncation }
            };

            SyncResponse syncResponse = this.SyncChangesWithOption(this.User2Information.InboxCollectionId, option);
            Site.Assert.IsNotNull(syncResponse.ResponseData, "The ResponseData should not be null.");
            Site.Assert.IsNotNull(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject), "The email should be in the Inbox folder of User2.");

            if (Common.IsRequirementEnabled(3432, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3432");

                // Verify MS-ASCMD requirement: MS-ASCMD_R3432
                Site.CaptureRequirementIfAreEqual<uint>(
                    1,
                    Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                    3432,
                    @"[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one MIMETruncation element as the child element of the Options element is undefined]. (Exchange 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to verify Sync command, when adding a new calendar item, of which the EndTime element is missing, the status should be equal to 6.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC24_Sync_Status6_EndTimeMissing()
        {
            // Synchronizes the changes in a collection between the client and the server.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId));

            #region Add a calendar item with Sync command.
            string calendarSubject = Common.GenerateResourceName(Site, "canlendarSubject");
            string calendarTo = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            string location = Common.GenerateResourceName(Site, "Room11");
            Request.SyncCollectionAdd addData = this.CreateAddCalendarCommand(calendarTo, calendarSubject, location, string.Empty);

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, addData);
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Response.SyncCollectionsCollectionResponses collectionResponse = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R752");

            // Verify MS-ASCMD requirement: MS-ASCMD_R752
            Site.CaptureRequirementIfAreEqual<int>(
                6,
                int.Parse(collectionResponse.Add[0].Status),
                752,
                @"[In Add(Sync)] [When the client adds a calendar item] A Status element value of 6 is returned in the Sync response if the EndTime element is not included.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4440");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4440
            Site.CaptureRequirementIfAreEqual<int>(
                6,
                int.Parse(collectionResponse.Add[0].Status),
                4440,
                @"[In Status(Sync)] [When the scope is Item], [the cause of the status value 6 is] The client has sent a malformed or invalid item.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, when adding an email item to the server, the status should be equal to 6.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC25_Sync_Status6_AddEmail()
        {
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId));

            #region Add an email item with Sync command.
            Request.SyncCollectionAdd addData = CreateAddEmailCommand(Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain), TestSuiteBase.ClientId);
            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.InboxCollectionId, addData);
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Response.SyncCollectionsCollectionResponses collectionResponse = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R755");

            // Verify MS-ASCMD requirement: MS-ASCMD_R755
            Site.CaptureRequirementIfAreEqual<int>(
                6,
                int.Parse(collectionResponse.Add[0].Status),
                755,
                @"[In Add(Sync)] If a client attempts to add emails to the server, a Status element with a value of 6 is returned as a child of the Add element.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the requirements related to GetChanges element.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC26_Sync_GetChanges()
        {
            #region Call Sync command with the GetChanges set to true and SyncKey element set 0.
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId);
            syncRequest.RequestData.Collections[0].GetChanges = true;
            syncRequest.RequestData.Collections[0].GetChangesSpecified = true;
            SyncResponse syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5838");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5838
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(syncResponse.ResponseData.Status),
                5838,
                @"[In GetChanges] A Status element (section 2.2.3.162.16) value of 4 is returned if the GetChanges element is [present and empty or] set to 1 (TRUE) when the SyncKey element value is 0 (zero).");
            #endregion

            #region Call Sync command with the GetChanges set to empty and SyncKey element set 0.
            string request = syncRequest.GetRequestDataSerializedXML().Replace("<GetChanges>1</GetChanges>", "<GetChanges></GetChanges>");
            SendStringResponse response = this.CMDAdapter.SendStringRequest(CommandName.Sync, null, request);

            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
            doc.LoadXml(response.ResponseDataXML);
            System.Xml.XmlNodeList nodes = doc.GetElementsByTagName("Status");
            Site.Assert.IsNotNull(nodes, "The Status element should exist in response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3128");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3128
            Site.CaptureRequirementIfAreEqual<string>(
                "4",
                nodes[0].InnerText,
                3128,
                @"[In GetChanges] A Status element (section 2.2.3.162.16) value of 4 is returned if the GetChanges element is present and empty [or set to 1 (TRUE)] when the SyncKey element value is 0 (zero).");
            #endregion

            #region Call Sync command with the GetChanges set to 0 and SyncKey element set 0.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId);
            syncRequest.RequestData.Collections[0].GetChanges = false;
            syncRequest.RequestData.Collections[0].GetChangesSpecified = true;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5839");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5839
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                5839,
                @"[In GetChanges] No error is returned if the GetChanges element is [absent or] set to 0 (FALSE) when the SyncKey value is 0 (zero).");
            #endregion

            #region Call Sync command with SyncKey element set 0.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3129");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3129
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                3129,
                @"[In GetChanges] No error is returned if the GetChanges element is absent [or set to 0 (FALSE)] when the SyncKey value is 0 (zero).");
            #endregion

            #region Add a new contact with Sync operation.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            string contactFileAS = Common.GenerateResourceName(Site, "FileAS");
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);

            syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, addData);
            syncResponse = this.Sync(syncRequest, false);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The new contact should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, contactFileAS);
            #endregion

            #region Call Sync command with the syncKey element set 0.
            this.FolderSync();
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId);
            SyncResponse syncResponseWithSyncKeyIs0 = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponseWithSyncKeyIs0.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            string syncKey = this.LastSyncKey;
            #endregion

            #region Call Sync command with the syncKey element set to non-zero value.
            syncRequest.RequestData.Collections[0].SyncKey = syncKey;
            SyncResponse syncResponseWithSyncKetIsNot0 = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponseWithSyncKetIsNot0.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            #endregion

            #region Call Sync command with an empty GetChanges element.
            syncRequest.RequestData.Collections[0].SyncKey = syncKey;
            syncRequest.RequestData.Collections[0].GetChangesSpecified = false;
            syncResponse = this.Sync(syncRequest, false);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3127");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3127
            Site.CaptureRequirementIfIsNotNull(
                syncResponse.ResponseData.Item,
                3127,
                @"[In GetChanges] A value of 1 (TRUE) is assumed when the GetChanges element is empty.");
            #endregion

            #region Call Sync command with the GetChanges element set to true and the syncKey element set to non-zero value.
            syncRequest.RequestData.Collections[0].SyncKey = syncKey;
            syncRequest.RequestData.Collections[0].GetChanges = true;
            syncRequest.RequestData.Collections[0].GetChangesSpecified = true;
            syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3126");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3126
            Site.CaptureRequirementIfIsNotNull(
                TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands),
                3126,
                @"[In GetChanges] A value of 1 (TRUE) indicates that the client wants the server changes to be returned.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5556");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5556
            // After above R3126 captured, it means the server returns the changes.
            // Then if the Command element is not null in the Sync response when the SyncKey set to non-zero value, the R5556 can be captured.
            Site.CaptureRequirementIfIsNotNull(
                TestSuiteBase.GetCollectionItem(syncResponseWithSyncKetIsNot0, Response.ItemsChoiceType10.Commands),
                5556,
                @"[In GetChanges] If the SyncKey element has a non-zero value, then the request is handled as if the GetChanges element were set to 1 (TRUE).");
            #endregion

            #region Call Sync command with the GetChanges element set to false and the syncKey element set to non-zero value.
            syncRequest.RequestData.Collections[0].SyncKey = syncKey;
            syncRequest.RequestData.Collections[0].GetChanges = false;
            syncRequest.RequestData.Collections[0].GetChangesSpecified = true;
            syncResponse = this.Sync(syncRequest, false);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3125");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3125
            Site.CaptureRequirementIfIsNull(
                syncResponse.ResponseData.Item,
                3125,
                @"[In GetChanges] If the client does not want the server changes returned, the request MUST include the GetChanges element with a value of 0 (FALSE).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5555");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5555
            // After above R3125 captured, it means the server does not return any changes.
            // Then if the Command element is null in the Sync response when the SyncKey set to 0, the R5555 can be captured.
            Site.CaptureRequirementIfIsNull(
                TestSuiteBase.GetCollectionItem(syncResponseWithSyncKeyIs0, Response.ItemsChoiceType10.Commands),
                5555,
                @"[In GetChanges] If the SyncKey element has a value of 0, then the request is handled as if the GetChanges element were set to 0 (FALSE).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if there are changes since the last synchronization, the server response includes a Commands element that contains additions, deletions, and changes.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC27_Sync_Change()
        {
            Site.Assume.AreEqual<string>("Base64", Common.GetConfigurationPropertyValue("HeaderEncodingType", this.Site), "The device ID should be same across all requests, when the HeaderEncodingType is PlainText.");
            CMDAdapter.ChangeDeviceID(Common.GetConfigurationPropertyValue("DeviceID", this.Site));
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            #region Add a new contact with Sync operation.
            string contactFileAS = Common.GenerateResourceName(Site, "FileAS");
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, addData);
            SyncResponse syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Response.SyncCollectionsCollectionResponses collectionResponse = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.IsNotNull(collectionResponse, "The responses element should exist in the Sync response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5255");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5255
            Site.CaptureRequirementIfIsNotNull(
                collectionResponse.Add,
                5255,
                @"[In Responses] Element Responses in Sync command response (section 2.2.2.19), the child elements is Add (section 2.2.3.7.2)[, Fetch (section 2.2.3.63.2) ](If the operation succeeded.)");

            Site.Assert.AreEqual<int>(1, int.Parse(collectionResponse.Add[0].Status), "The new contact should be added correctly.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, contactFileAS);
            this.FolderSync();
            #endregion

            #region Change DeviceID and synchronize the changes in the Contacts folder.
            string syncKey = this.LastSyncKey;
            CMDAdapter.ChangeDeviceID("Device2");
            this.RecordDeviceInfoChanged();

            this.GetInitialSyncResponse(this.User1Information.ContactsCollectionId);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            string serverId = TestSuiteBase.FindServerId(syncResponse, "FileAs", contactFileAS);
            Site.Assert.IsNotNull(serverId, "The added contact should be synchronized down to the current device.");
            #endregion

            #region Change the added Contact information and then synchronize the change to the server.
            string updatedContactFileAS = Common.GenerateResourceName(Site, "UpdatedFileAS");
            Request.SyncCollectionChange appDataChange = CreateChangedContact(serverId, new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.FileAs }, new object[] { updatedContactFileAS });
            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, appDataChange);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The status code of Sync change operation should be 1.");
            #endregion

            #region Restore DeviceID and synchronize the changes in the Contacts folder.
            
            CMDAdapter.ChangeDeviceID(Common.GetConfigurationPropertyValue("DeviceID", this.Site));

            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId);
            syncRequest.RequestData.Collections[0].SyncKey = syncKey;
            syncRequest.RequestData.Collections[0].GetChanges = true;
            syncRequest.RequestData.Collections[0].GetChangesSpecified = true;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The Sync command should be conducted successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, updatedContactFileAS);

            Response.SyncCollectionsCollectionCommands syncCollectionCommands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            Site.Assert.IsNotNull(syncCollectionCommands, "The commands element should exist in the Sync response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3119");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3119
            Site.CaptureRequirementIfIsNotNull(
                syncCollectionCommands.Change,
                3119,
                @"[In GetChanges] If there have been changes since the last synchronization, the server response includes a Commands element (section 2.2.3.32) that contains additions, deletions, and changes.");

            Site.Assert.IsTrue(((Response.SyncCollections)syncResponse.ResponseData.Item).Collection.Length >= 1, "The length of Collections element should not less than 1. The actual value is {0}.", ((Response.SyncCollections)syncResponse.ResponseData.Item).Collection.Length);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4602");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4602
            // If the Assert statement above is passed, the requirement is captured.
            Site.CaptureRequirement(
                4602,
                @"[In SyncKey(Sync)] If the synchronization is successful, the server responds by sending all objects in the collection.");

            List<string> elements = new List<string>
            {
                Request.ItemsChoiceType4.FirstName.ToString(),
                Request.ItemsChoiceType4.MiddleName.ToString(),
                Request.ItemsChoiceType4.LastName.ToString()
            };

            Response.SyncCollectionsCollectionCommands commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            Site.Assert.IsNotNull(commands, "The Commands element should not be null.");
            Site.Assert.IsNotNull(commands.Change, "The Change element should not be null.");
            Site.Assert.IsTrue(commands.Change.Length > 0, "The Change element should have one sub-element at least.");
            Site.Assert.IsNotNull(commands.Change[0].ApplicationData, "The ApplicationData element of the first Change element should not be null.");
            Site.Assert.IsNotNull(commands.Change[0].ApplicationData.ItemsElementName, "The ItemsElementName element of the ApplicationData element of the first Change element should not be null.");
            Site.Assert.IsTrue(commands.Change[0].ApplicationData.ItemsElementName.Length > 0, "The ItemsElementName element of the ApplicationData element of the first Change element should have one sub-element at least.");

            bool isVerifyR879 = false;
            foreach (Response.ItemsChoiceType7 itemElementName in commands.Change[0].ApplicationData.ItemsElementName)
            {
                if (elements.Contains(itemElementName.ToString()))
                {
                    isVerifyR879 = true;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R879");

            // Verify MS-ASCMD requirement: MS-ASCMD_R879
            Site.CaptureRequirementIfIsFalse(
                isVerifyR879,
                879,
                @"[In Change] In all other cases, if an in-schema property is not specified in a change request, the property is actively deleted from the item on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R881");

            // Verify MS-ASCMD requirement: MS-ASCMD_R881
            Site.CaptureRequirementIfIsFalse(
                isVerifyR879,
                881,
                @"[In Change] Otherwise [if a client dose not be aware of this [if an in-schema property is not specified in a change request, the property is actively deleted from the item on the server] when it [client] is sending Sync requests], data can be unintentionally removed.");
            #endregion

            #region Send a MIME-formatted e-mail from user1 to user2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);

            this.SwitchUser(this.User2Information);
            serverId = TestSuiteBase.FindServerId(this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null), "Subject", emailSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region Call Sync command to set a flag to an email.
            string updatedEmailSubject = Common.GenerateResourceName(Site, "updatedSubject");

            DateTime startDate = DateTime.Now.AddDays(5.0);
            DateTime dueDate = startDate.AddHours(1.0);

            // Define email flag
            Request.Flag emailFlag = new Request.Flag
            {
                StartDate = startDate,
                StartDateSpecified = true,
                DueDate = dueDate,
                DueDateSpecified = true
            };

            Request.SyncCollectionChangeApplicationData applicationData = new Request.SyncCollectionChangeApplicationData
            {
                ItemsElementName = new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.Flag, Request.ItemsChoiceType7.Subject },
                Items = new object[] { emailFlag, updatedEmailSubject }
            };

            appDataChange = new Request.SyncCollectionChange { ApplicationData = applicationData, ServerId = serverId };

            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User2Information.InboxCollectionId, appDataChange);
            syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R875");

            // Verify MS-ASCMD requirement: MS-ASCMD_R875
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                875,
                @"[In Change] If all the other elements are sent, extra bandwidth is used, but no errors occur.");

            // Define updated email flag
            Request.Flag updatedEmailFlag = new Request.Flag
            {
                StartDate = startDate,
                StartDateSpecified = true,
                DueDate = dueDate,
                DueDateSpecified = true
            };

            applicationData = new Request.SyncCollectionChangeApplicationData
            {
                ItemsElementName = new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.Flag },
                Items = new object[] { updatedEmailFlag }
            };

            appDataChange = new Request.SyncCollectionChange { ApplicationData = applicationData, ServerId = serverId };

            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User2Information.InboxCollectionId, appDataChange);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The email should be updated successfully.");

            syncResponse = this.SyncChanges(this.User2Information.InboxCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R873");

            // Verify MS-ASCMD requirement: MS-ASCMD_R873
            bool isVerifyR873 = !string.IsNullOrEmpty(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject));
            Site.CaptureRequirementIfIsTrue(
                isVerifyR873,
                873,
                @"[In Change] Certain in-schema properties remain untouched in the following three cases: If there is only an email:Flag ([MS-ASEMAIL] section 2.2.2.27) [, email:Read ([MS-ASEMAIL] section 2.2.2.47), or email:Categories ([MS-ASEMAIL] section 2.2.2.9)] change (that is, if only an email:Flag, email:Categories or email:Read element is present), all other properties will remain unchanged.");
            #endregion

            #region Send a MIME-formatted e-mail from user1 to user2.
            this.SwitchUser(this.User1Information);
            emailSubject = Common.GenerateResourceName(this.Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);

            this.SwitchUser(this.User2Information);
            serverId = TestSuiteBase.FindServerId(this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null), "Subject", emailSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region Call Sync command to set a Read element to an email.
            applicationData = new Request.SyncCollectionChangeApplicationData
            {
                ItemsElementName = new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.Read, Request.ItemsChoiceType7.Subject },
                Items = new object[] { true, updatedEmailSubject }
            };

            appDataChange = new Request.SyncCollectionChange { ApplicationData = applicationData, ServerId = serverId };

            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User2Information.InboxCollectionId, appDataChange);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The email should be updated successfully.");

            applicationData = new Request.SyncCollectionChangeApplicationData
            {
                ItemsElementName = new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.Read },
                Items = new object[] { true }
            };

            appDataChange = new Request.SyncCollectionChange { ApplicationData = applicationData, ServerId = serverId };

            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User2Information.InboxCollectionId, appDataChange);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The email should be updated successfully.");

            syncResponse = this.SyncChanges(this.User2Information.InboxCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5825");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5825
            bool isVerifyR5825 = !string.IsNullOrEmpty(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject));
            Site.CaptureRequirementIfIsTrue(
                isVerifyR5825,
                5825,
                @"[In Change] Certain in-schema properties remain untouched in the following three cases: If there is only an [email:Flag ([MS-ASEMAIL] section 2.2.2.27),] email:Read ([MS-ASEMAIL] section 2.2.2.47) [, or email:Categories ([MS-ASEMAIL] section 2.2.2.9)] change (that is, if only an email:Flag, email:Categories or email:Read element is present), all other properties will remain unchanged.");
            #endregion

            #region Send a MIME-formatted e-mail from user1 to user2.
            this.SwitchUser(this.User1Information);
            emailSubject = Common.GenerateResourceName(this.Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);

            this.SwitchUser(this.User2Information);
            serverId = TestSuiteBase.FindServerId(this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null), "Subject", emailSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region Call Sync command to set a Categories element to an email.
            Request.Categories categories = new Request.Categories { Category = new string[] { "company" } };

            applicationData = new Request.SyncCollectionChangeApplicationData
            {
                ItemsElementName = new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.Categories, Request.ItemsChoiceType7.Subject },
                Items = new object[] { categories, updatedEmailSubject }
            };

            appDataChange = new Request.SyncCollectionChange { ApplicationData = applicationData, ServerId = serverId };

            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User2Information.InboxCollectionId, appDataChange);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The email should be updated successfully.");

            applicationData = new Request.SyncCollectionChangeApplicationData
            {
                ItemsElementName = new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.Categories },
                Items = new object[] { categories }
            };

            appDataChange = new Request.SyncCollectionChange { ApplicationData = applicationData, ServerId = serverId };

            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User2Information.InboxCollectionId, appDataChange);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The email should be updated successfully.");

            syncResponse = this.SyncChanges(this.User2Information.InboxCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5826");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5826
            bool isVerifyR5826 = !string.IsNullOrEmpty(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject));
            Site.CaptureRequirementIfIsTrue(
                isVerifyR5826,
                5826,
                @"[In Change] Certain in-schema properties remain untouched in the following three cases: If there is only [an email:Flag ([MS-ASEMAIL] section 2.2.2.27), email:Read ([MS-ASEMAIL] section 2.2.2.47), or] email:Categories ([MS-ASEMAIL] section 2.2.2.9) change (that is, if only an email:Flag, email:Categories or email:Read element is present), all other properties will remain unchanged.");
            #endregion

            #region Call Sync Add operation to add a new contact.
            this.SwitchUser(this.User1Information);
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            string contactFileAs = Common.GenerateResourceName(Site, "FileAS");

            addData.ClientId = TestSuiteBase.ClientId;
            addData.ApplicationData = new Request.SyncCollectionAddApplicationData
            {
                ItemsElementName =
                    new Request.ItemsChoiceType8[]
                    {
                        Request.ItemsChoiceType8.FileAs, Request.ItemsChoiceType8.FirstName,
                        Request.ItemsChoiceType8.MiddleName, Request.ItemsChoiceType8.LastName,
                        Request.ItemsChoiceType8.Picture
                    },
                Items =
                    new object[]
                    {
                        contactFileAs, "FirstName", "MiddleName", "LastName",
                        Convert.ToBase64String(File.ReadAllBytes("number1.jpg"))
                    }
            };

            if ("12.1" != Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site))
            {
                addData.Class = "Contacts";
            }

            syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, addData);
            syncResponse = this.Sync(syncRequest, false);

            collectionResponse = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.IsNotNull(collectionResponse, "The responses element should exist in the Sync response.");
            Site.Assert.AreEqual<int>(1, int.Parse(collectionResponse.Add[0].Status), "The new contact should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, contactFileAs);

            syncResponse = this.SyncChanges(this.User1Information.ContactsCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            serverId = TestSuiteBase.FindServerId(syncResponse, "FileAs", contactFileAs);
            Site.Assert.IsNotNull(serverId, "The added contact should be synchronized down to the current device.");

            Response.SyncCollectionsCollectionCommandsAddApplicationData responseApplicationData = TestSuiteBase.GetAddApplicationData(syncResponse, Response.ItemsChoiceType8.FileAs, contactFileAS);
            Site.Assert.IsNotNull(responseApplicationData, "The ApplicationData of the calendar should not be null.");

            Response.Body body = null;
            for (int i = 0; i < responseApplicationData.ItemsElementName.Length; i++)
            {
                if (responseApplicationData.ItemsElementName[i] == Response.ItemsChoiceType8.Body)
                {
                    body = responseApplicationData.Items[i] as Response.Body;
                    break;
                }
            }

            Site.Assert.IsNotNull(body, "The Body element should be in the ApplicationData of the contact item.");
            string originalPicture = (string)TestSuiteBase.GetElementValueFromSyncResponse(syncResponse, serverId, Response.ItemsChoiceType8.Picture);
            Site.Assert.IsNotNull(originalPicture, "The picture of the contact should exist.");
            #endregion

            #region Call Sync change operation to update the FileAs element of the contact.
            string updatedContactFileAs = Common.GenerateResourceName(Site, "updatedContactFileAs");

            Request.SyncCollectionChangeApplicationData changeApplicationData = new Request.SyncCollectionChangeApplicationData
            {
                ItemsElementName = new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.FileAs },
                Items = new object[] { updatedContactFileAs }
            };

            appDataChange = new Request.SyncCollectionChange
            {
                ApplicationData = changeApplicationData,
                ServerId = serverId
            };

            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, appDataChange);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The FileAs of the contact should be updated successfully.");
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, contactFileAs);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, updatedContactFileAs);

            syncResponse = this.SyncChanges(this.User1Information.ContactsCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            responseApplicationData = TestSuiteBase.GetAddApplicationData(syncResponse, Response.ItemsChoiceType8.FileAs, updatedContactFileAs);
            Site.Assert.IsNotNull(responseApplicationData, "The ApplicationData of the updated contact should not be null.");

            Response.Body currentBody = null;
            for (int i = 0; i < responseApplicationData.ItemsElementName.Length; i++)
            {
                if (responseApplicationData.ItemsElementName[i] == Response.ItemsChoiceType8.Body)
                {
                    currentBody = responseApplicationData.Items[i] as Response.Body;
                    break;
                }
            }

            Site.Assert.IsNotNull(currentBody, "The Body element should be in the ApplicationData of the updated contact item.");
            string currentPicture = (string)TestSuiteBase.GetElementValueFromSyncResponse(syncResponse, serverId, Response.ItemsChoiceType8.Picture);
            Site.Assert.IsNotNull(currentPicture, "The picture of the updated contact should exist.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R878");

            // Verify MS-ASCMD requirement: MS-ASCMD_R878
            bool isVerifyR878 = body.Type == currentBody.Type && body.EstimatedDataSize == currentBody.EstimatedDataSize && originalPicture == currentPicture;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR878,
                878,
                @"[In Change] [Certain in-schema properties remain untouched in the following three cases:] If the airsyncbase:Body, airsyncbase:Data, or contacts:Picture elements are not present, the corresponding properties will remain unchanged.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if operation failed, Change element as a child element of Response element should appear.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC28_Sync_Change_InvalidServerId()
        {
            // Synchronize the changes in the Contacts folder.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            #region Use an invalid ServerId to try to change a contact.
            string invalidServerId = Guid.NewGuid().ToString();
            string updatedContactFileAS = Common.GenerateResourceName(Site, "UpdatedFileAS");
            Request.SyncCollectionChange appDataChange = CreateChangedContact(invalidServerId, new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.FileAs }, new object[] { updatedContactFileAS });
            SyncRequest syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, appDataChange);
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            Response.SyncCollectionsCollectionResponses syncCollectionResponse = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.IsNotNull(syncCollectionResponse, "The responses element should exist in the Sync response.");

            #region Capture Code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5256");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5256
            Site.CaptureRequirementIfIsNotNull(
                syncCollectionResponse.Change,
                5256,
                @"[In Responses] Element Responses in Sync command response (section 2.2.2.19), the child elements is Change (section 2.2.3.24) (If the operation failed.)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3834");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3834
            Site.CaptureRequirementIfIsNotNull(
                syncCollectionResponse,
                3834,
                @"[In Responses] It[Responses element] is present only if the server has processed operation from the client.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4419");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4419
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                syncCollectionResponse.Change[0].Status,
                4419,
                @"[In Status(Sync)] If the operation[the Change operation, the Add operation, or the Fetch operation] failed, the Status element contains a code that indicates the type of failure.");

            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if both HeartbeatInterval and Wait elements are included, server should return a status value of 4.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC29_Sync_HeartbeatIntervalAndWait_StatusIs4()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The HeartbeatInterval tag is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            syncRequest.RequestData.Wait = "10";
            syncRequest.RequestData.HeartbeatInterval = "1000";
            SyncResponse syncResponse = this.Sync(syncRequest);

            #region Capture Code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3162");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3162
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(syncResponse.ResponseData.Status),
                3162,
                @"[In HeartbeatInterval(Sync)] If both[HeartbeatInterval element and the Wait element ] elements are included, the server response will contain a Status element value of 4.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4748");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4748
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(syncResponse.ResponseData.Status),
                4748,
                @"[In Wait] If both [Wait element and the HeartbeatInterval element ] elements are included, the server response will contain a Status element value of 4.");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if a Wait element value of less than 1 is sent, the server will return a Limit element value of 1.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC30_Sync_Limit_WaitIsLessThan1()
        {
            // Synchronize the changes in the Inbox folder with a request containing a Wait value which is less than the lower bound of the domain.
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            syncRequest.RequestData.Wait = "0";
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Site.Assert.AreEqual<int>(14, int.Parse(syncResponse.ResponseData.Status), "The Status code should be 14.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3227");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3227
            // When the Status value is 14, the value of the Item element represents the value of the Limit element.
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                (string)syncResponse.ResponseData.Item,
                3227,
                @"[In Limit] If a Wait element value of less than 1 is sent, the server returns a Limit element value of 1, indicating the minimum value of the Wait element is 1.");
        }

        /// <summary>
        /// This test case is used to verify Sync command, if a Wait element value of greater than 59 is sent, the server will return a Limit element value of 59.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC31_Sync_Limit_WaitIsGreaterThan59()
        {
            // Synchronize the changes in the Inbox folder with a request containing a Wait value which is greater than the upper bound of the domain.
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            syncRequest.RequestData.Wait = "60";
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Site.Assert.AreEqual<int>(14, int.Parse(syncResponse.ResponseData.Status), "The Status code should be 14.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3228");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3228
            // When the Status value is 14, the value of the Item element represents the value of the Limit element.
            Site.CaptureRequirementIfAreEqual<string>(
                "59",
                (string)syncResponse.ResponseData.Item,
                3228,
                @"[In Limit] If a Wait element value greater than 59 is sent, the server returns a Limit element value of 59, indicating the maximum value of the Wait element is 59.");
        }

        /// <summary>
        /// This test is used to verify Sync command, if a HeartbeatInterval element value of less than 60 is sent, the server will return a Limit element value of 60.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC32_Sync_Limit_HeartbeatIntervalIsLessThan60()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The HeartbeatInterval tag is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            syncRequest.RequestData.HeartbeatInterval = "59";
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Site.Assert.AreEqual<int>(14, int.Parse(syncResponse.ResponseData.Status), "The Status code should be 14.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3229");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3229
            // When the Status value is 14, the value of the Item element represents the value of the Limit element.
            Site.CaptureRequirementIfAreEqual<string>(
                "60",
                (string)syncResponse.ResponseData.Item,
                3229,
                @"[In Limit] If a HeartbeatInterval element value of less than 60 is sent, the server returns a Limit element value of 60, indicating the minimum value of the HeartbeatInterval element is 60.");
        }

        /// <summary>
        /// This test is used to verify Sync command, if a HeartbeatInterval element value of greater than 3540 is sent, the server will return a Limit element value of 3540.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC33_Sync_Limit_HeartbeatIntervalIsLargerThan3540()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The HeartbeatInterval tag is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            syncRequest.RequestData.HeartbeatInterval = "3541";
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3160");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3160
            // When the Status value is 14, the value of the Item element represents the value of the Limit element.
            bool isVerifyR3160 = int.Parse(syncResponse.ResponseData.Status) == 14 && !string.IsNullOrEmpty((string)syncResponse.ResponseData.Item) && (string)syncResponse.ResponseData.Item == "3540";
            Site.CaptureRequirementIfIsTrue(
                isVerifyR3160,
                3160,
                @"[In HeartbeatInterval(Sync)] When the client requests an interval that is outside the acceptable range, the server will send a response that includes a Status element (section 2.2.3.162.16) value of 14 and a Limit element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3230");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3230
            // When the Status value is 14, the value of the Item element represents the value of the Limit element.
            Site.CaptureRequirementIfAreEqual<string>(
                "3540",
                (string)syncResponse.ResponseData.Item,
                3230,
                @"[In Limit] If a HeartbeatInterval element value greater than 3540 is sent, the server returns a Limit element value of 3540, indicating the maximum value of HeartbeatInterval element is 3540.");
        }

        /// <summary>
        /// This test case is used to verify Sync command, if DeletesAsMoves is true, the deleted item is moved to the Deleted Items folder.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC34_Sync_DeletesAsMovesIsTrue()
        {
            #region Send a MIME-formatted email to User2.
            string emailSubject1 = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject1, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject1);

            // Check DeletedItems folder
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User2Information.DeletedItemsCollectionId);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNull(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject1), "The email should not be found in the DeletedItems folder.");

            // Check Inbox folder
            syncResponse = this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject1, null);
            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject1);
            Site.Assert.IsNotNull(serverId, "The email should be found in the inbox folder.");

            #region Delete the email with DeletesAsMoves set to true from the Inbox folder.
            Request.SyncCollection collection = new Request.SyncCollection
            {
                SyncKey = this.LastSyncKey,
                CollectionId = this.User2Information.InboxCollectionId,
                Commands = new object[] { new Request.SyncCollectionDelete { ServerId = serverId } },
                DeletesAsMoves = true,
                DeletesAsMovesSpecified = true
            };

            Request.Sync syncRequestData = new Request.Sync { Collections = new Request.SyncCollection[] { collection } };

            SyncRequest syncRequestDelete = new SyncRequest { RequestData = syncRequestData };
            syncResponse = this.Sync(syncRequestDelete);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The Sync delete operation should be successful.");
            #endregion

            #region Verify if the email has been deleted from the Inbox folder and placed into the DeletedItems folder.
            // Check Inbox folder.
            syncResponse = this.SyncChanges(this.User2Information.InboxCollectionId, false);
            Site.Assert.IsNull(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject1), "The email deleted should not be found in the Inbox folder.");
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject1);

            // Check DeletedItems folder
            this.CheckEmail(this.User2Information.DeletedItemsCollectionId, emailSubject1, null);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.DeletedItemsCollectionId, emailSubject1);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2160");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2160
            // When the Assert statements above are passed, it means the deleted email is moved from the inbox folder to Deleted Items folder, then this requirement can be  captured directly.
            Site.CaptureRequirement(
                2160,
                @"[In DeletesAsMoves] A value of 1 (TRUE), which is the default, indicates that any deleted items are moved to the Deleted Items folder.");
            #endregion

            #region Send a MIME-formatted email to User2.
            this.SwitchUser(this.User1Information);
            string emailSubject2 = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject2, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject1);

            // Check DeletedItems folder
            syncResponse = this.SyncChanges(this.User2Information.DeletedItemsCollectionId);
            Site.Assert.IsNull(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject2), "The email should not be found in the DeletedItems folder.");

            // Check Inbox folder
            syncResponse = this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject2, null);
            serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject2);
            Site.Assert.IsNotNull(serverId, "The email should be found in the inbox folder.");

            #region Delete the email and DeletesAsMoves is not present in the request
            collection = new Request.SyncCollection
            {
                SyncKey = this.LastSyncKey,
                CollectionId = this.User2Information.InboxCollectionId,
                Commands = new object[] { new Request.SyncCollectionDelete { ServerId = serverId } }
            };

            syncRequestData = new Request.Sync { Collections = new Request.SyncCollection[] { collection } };

            syncRequestDelete = new SyncRequest { RequestData = syncRequestData };
            syncResponse = this.Sync(syncRequestDelete);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The Sync delete operation should be successful.");

            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject1);
            #endregion

            #region Verify if the second email has been deleted from the Inbox folder and placed into the DeletedItems folder.
            // Check Inbox folder
            syncResponse = this.SyncChanges(this.User2Information.InboxCollectionId);
            Site.Assert.IsNull(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject2), "The deleted email should not be found in the Inbox folder.");

            // Check DeletedItems folder
            syncResponse = this.SyncChanges(this.User2Information.DeletedItemsCollectionId);
            Site.Assert.IsNotNull(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject2), "The deleted email should be found in the DeletedItems folder.");

            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject1);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.DeletedItemsCollectionId, emailSubject1, emailSubject2);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5874");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5874
            // When the Assert statements above are passed, it means the deleted email is moved from the inbox folder to Deleted Items folder, then this requirement can be  captured directly.
            Site.CaptureRequirement(
                5874,
                @"[In DeletesAsMoves] If element DeleteAsMoves is empty, the delete items are moved to the Deleted Items folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5875");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5875
            // When the Assert statements above are passed, it means the deleted email is moved from the inbox folder to Deleted Items folder, then this requirement can be  captured directly.
            Site.CaptureRequirement(
                5875,
                @"[In DeletesAsMoves] If element DeleteAsMoves is not present, the delete items are moved to the Deleted Items folder.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if DeletesAsMoves is false, the deletion is permanent.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC35_Sync_DeletesAsMovesIsFalse()
        {
            #region Send a MIME-formatted email to User2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);

            SyncResponse syncResponse = this.SyncChanges(this.User2Information.DeletedItemsCollectionId);
            Site.Assert.IsNull(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject), "The email should not be found in the DeletedItems folder.");

            syncResponse = this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject);
            Site.Assert.IsTrue(!string.IsNullOrEmpty(serverId), "The email should be found in the Inbox folder.");

            #region Delete the added email item.
            Request.SyncCollectionDelete appDataDelete = new Request.SyncCollectionDelete { ServerId = serverId };

            Request.SyncCollection collection = new Request.SyncCollection
            {
                SyncKey = this.LastSyncKey,
                GetChanges = true,
                CollectionId = this.User2Information.InboxCollectionId,
                Commands = new object[] { appDataDelete },
                DeletesAsMoves = false,
                DeletesAsMovesSpecified = true
            };

            Request.Sync syncRequestData = new Request.Sync { Collections = new Request.SyncCollection[] { collection } };

            SyncRequest syncRequestDelete = new SyncRequest { RequestData = syncRequestData };
            this.Sync(syncRequestDelete);
            #endregion

            #region Verify if the email has been deleted from the Inbox folder and not placed into the DeletedItems folder.
            syncResponse = this.SyncChanges(this.User2Information.InboxCollectionId);
            Site.Assert.IsNull(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject), "The email deleted should not be found in the Inbox folder.");
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);

            syncResponse = this.SyncChanges(this.User2Information.DeletedItemsCollectionId);
            Site.Assert.IsNull(TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject), "The email deleted should not be found in the DeletedItems folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2158");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2158
            // If the deleted email can not be found in both Inbox and Deleted Items folder, this requirement can be captured directly.
            Site.CaptureRequirement(
                2158,
                @"[In DeletesAsMoves] If the DeletesAsMoves element is set to false, the deletion is permanent.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command for Fetch operation.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC36_Sync_Fetch()
        {
            #region Send a MIME-formatted email to User2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            SyncResponse syncResponse = this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null);

            #region Fetch the email.
            Request.SyncCollectionFetch appDataFetch = new Request.SyncCollectionFetch
            {
                ServerId = TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject)
            };

            Request.SyncCollection collection = new Request.SyncCollection
            {
                SyncKey = this.LastSyncKey,
                GetChanges = true,
                GetChangesSpecified = true,
                CollectionId = this.User2Information.InboxCollectionId,
                Commands = new object[] { appDataFetch }
            };

            Request.Sync syncRequestData = new Request.Sync { Collections = new Request.SyncCollection[] { collection } };

            SyncRequest syncRequestForFetch = new SyncRequest { RequestData = syncRequestData };
            syncResponse = this.Sync(syncRequestForFetch);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            Response.SyncCollectionsCollectionResponses collectionResponse = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.IsNotNull(collectionResponse, "The responses element should exist in the Sync response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5433");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5433
            Site.CaptureRequirementIfIsNotNull(
                collectionResponse.Fetch,
                5433,
                @"[In Responses] Element Responses in Sync command response (section 2.2.2.19), the child elements is [Add (section 2.2.3.7.2),] Fetch (section 2.2.3.63.2) (If the operation succeeded.)");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if the WindowSize element is set to 512, the server can send Sync response messages that contain less than 512 updates.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC37_Sync_WindowSize_512()
        {
            // Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
            FolderCreateRequest folderCreateRequest = Common.CreateFolderCreateRequest(this.LastFolderSyncKey, (byte)FolderType.UserCreatedContacts, Common.GenerateResourceName(Site, "FolderCreate"), this.User1Information.ContactsCollectionId);
            FolderCreateResponse folderCreateResponse = this.CMDAdapter.FolderCreate(folderCreateRequest);
            Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderCreate command response to indicate success.");

            // Record created folder collectionID
            string folderId = folderCreateResponse.ResponseData.ServerId;
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderId);

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(folderId));

            // Call method Sync to add 513 contact items.
            string contactFileAS = Common.GenerateResourceName(Site, "FileAS", 1);
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, folderId, addData);
            this.Sync(syncRequest, false);
            this.FolderSync();

            for (int i = 2; i <= 513; i++)
            {
                contactFileAS = Common.GenerateResourceName(this.Site, "FileAS", (uint)i);
                addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);

                syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, folderId, addData);
                this.Sync(syncRequest, false);
            }

            #region Synchronize the changes in the Contacts folder.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(folderId);
            this.Sync(syncRequest, false);
            string lastSyncKey = this.LastSyncKey;

            syncRequest.RequestData.Collections[0].SyncKey = lastSyncKey;
            SyncResponse syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The Collections element should exist in the Sync response.");

            bool isVerifyR3460WithoutWindowSize = CheckElementOfItemsChoiceType10(syncResponse, Response.ItemsChoiceType10.MoreAvailable);

            syncRequest.RequestData.Collections[0].SyncKey = lastSyncKey;
            syncRequest.RequestData.WindowSize = "512";
            syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The Collections element should exist in the Sync response.");

            Response.SyncCollectionsCollectionCommands collectionCommands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            Site.Assert.IsNotNull(collectionCommands, "The commands element should exist in the Sync response.");

            bool isVerifyR3460WithWindowSize = CheckElementOfItemsChoiceType10(syncResponse, Response.ItemsChoiceType10.MoreAvailable);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3460");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3460
            Site.CaptureRequirementIfIsTrue(
                isVerifyR3460WithWindowSize && isVerifyR3460WithoutWindowSize,
                3460,
                @"[In MoreAvailable] The MoreAvailable element is returned by the server if there are more than 512 changes, regardless of whether the WindowSize element is included in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3460");
            int itemCount = 0;
            if (collectionCommands.Add != null)
            {
                itemCount += collectionCommands.Add.Length;
            }

            if (collectionCommands.Change != null)
            {
                itemCount += collectionCommands.Change.Length;
            }

            if (collectionCommands.Delete != null)
            {
                itemCount += collectionCommands.Delete.Length;
            }

            if (collectionCommands.SoftDelete != null)
            {
                itemCount += collectionCommands.SoftDelete.Length;
            }

            bool isVerifiedR4764 = itemCount < 512;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR4764,
                4764,
                @"[In WindowSize] However, if the WindowSize element is set to 512, the server can send Sync response messages that contain less than 512 updates.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command for Delete operation.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC38_Sync_Delete()
        {
            Site.Assume.AreEqual<string>("Base64", Common.GetConfigurationPropertyValue("HeaderEncodingType", this.Site), "The device ID should be same across all requests, when the HeaderEncodingType is PlainText.");

            // Synchronize the changes in the Contacts folder.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            #region Add a new contact.
            string contactFileAS = Common.GenerateResourceName(Site, "FileAS");
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, addData);
            this.Sync(syncRequest);
            this.FolderSync();

            // Synchronize the changes in the Contacts folder.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            string syncKey = this.LastSyncKey;
            #endregion

            // Changed DeviceID
            CMDAdapter.ChangeDeviceID("Device2");

            // Synchronize the changes in the Contacts folder.
            this.GetInitialSyncResponse(this.User1Information.ContactsCollectionId);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Site.Assert.IsTrue(!string.IsNullOrEmpty(TestSuiteBase.FindServerId(syncResponse, "FileAs", contactFileAS)), "The added contact should be synchronized down to the current device.");

            #region Delete the added contact item.
            syncRequest = TestSuiteBase.CreateSyncDeleteRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, TestSuiteBase.FindServerId(syncResponse, "FileAs", contactFileAS));
            syncRequest.RequestData.Collections[0].DeletesAsMoves = false;
            syncRequest.RequestData.Collections[0].DeletesAsMovesSpecified = true;
            this.Sync(syncRequest);
            #endregion

            // Restore DeviceID
            CMDAdapter.ChangeDeviceID(Common.GetConfigurationPropertyValue("DeviceID", this.Site));

            #region Get the changes in the Contacts folder.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId);
            syncRequest.RequestData.Collections[0].SyncKey = syncKey;
            syncRequest.RequestData.Collections[0].GetChanges = true;
            syncRequest.RequestData.Collections[0].GetChangesSpecified = true;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Site.Assert.IsNull(TestSuiteBase.FindServerId(syncResponse, "FileAs", contactFileAS), "The deleted contact should not exist in the Contacts folder.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the requirements related to Supported element.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC39_Sync_Supported()
        {
            #region Add a new contact.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            string contactFileAS = Common.GenerateResourceName(Site, "FileAS");
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, "Vice President");

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, addData);
            SyncResponse syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Response.SyncCollectionsCollectionResponses responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The new contact should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, contactFileAS);
            #endregion

            #region Call Sync command without Supported element to indicate to the server that elements that can be ghosted are considered not ghosted.
            this.FolderSync();
            syncResponse = this.SyncChanges(this.User1Information.ContactsCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            string serverId = TestSuiteBase.FindServerId(syncResponse, "FileAs", contactFileAS);
            Site.Assert.IsTrue(!string.IsNullOrEmpty(serverId), "The contact should exist in the Contact folder.");
            #endregion

            #region Call Sync change operation to change the JobTitle element.
            Request.SyncCollectionChange appDataChange2 = CreateChangedContact(serverId, new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.JobTitle }, new object[] { "President" });
            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, appDataChange2);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The Sync command should succeed.");

            syncResponse = this.SyncChanges(this.User1Information.ContactsCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            bool isVerifyR5906 = true;
            Response.SyncCollectionsCollectionCommandsAddApplicationData applicationData1 = TestSuiteBase.GetAddApplicationData(syncResponse, Response.ItemsChoiceType8.FileAs, contactFileAS);
            Site.Assert.IsNotNull(applicationData1, "The application data of the contact item should not be null.");
            Site.Assert.IsNotNull(applicationData1.ItemsElementName, "The ItemsElementName should not be null.");
            Site.Assert.IsTrue(applicationData1.ItemsElementName.Length > 0, "The length of ItemsElementName should be greater then 0.");
            for (int i = 0; i < applicationData1.ItemsElementName.Length; i++)
            {
                if (applicationData1.ItemsElementName[i] == Response.ItemsChoiceType8.FirstName || applicationData1.ItemsElementName[i] == Response.ItemsChoiceType8.MiddleName || applicationData1.ItemsElementName[i] == Response.ItemsChoiceType8.LastName)
                {
                    isVerifyR5906 = false;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5906");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5906
            Site.CaptureRequirementIfIsTrue(
                isVerifyR5906,
                5906,
                @"[In Supported] The status of properties that can be ghosted is determined by the client's usage of the Supported element in the initial Sync command request for the containing folder, according to the following rules:1. If the client does not include a Supported element in the initial Sync command request for a folder, then all of the elements that can be ghosted are considered not ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5912");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5912
            Site.CaptureRequirementIfIsTrue(
                isVerifyR5906,
                5912,
                @"[In Supported] [When an existing item is modified via the Change element (section 2.2.3.24) in a Sync command request, the result of omitting an element that can be ghosted changes depending on the status of the element.] If the element is not ghosted, any existing value for that element is deleted.");
            #endregion

            #region Add a new contact.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            contactFileAS = Common.GenerateResourceName(this.Site, "FileAS");
            addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, "Vice President");

            syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, addData);
            syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The new contact should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, contactFileAS);
            #endregion

            #region Call Sync command to indicate to the server that JobTitle element is not ghosted.
            Request.Supported supported = new Request.Supported
            {
                Items = new string[] { string.Empty },
                ItemsElementName = new Request.ItemsChoiceType[] { Request.ItemsChoiceType.JobTitle }
            };

            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId);
            syncRequest.RequestData.Collections[0].Supported = supported;
            this.Sync(syncRequest);

            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncRequest.RequestData.Collections[0].GetChanges = true;
            syncRequest.RequestData.Collections[0].GetChangesSpecified = true;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            serverId = TestSuiteBase.FindServerId(syncResponse, "FileAs", contactFileAS);
            Site.Assert.IsTrue(!string.IsNullOrEmpty(serverId), "The contact should exist in the Contact folder.");
            #endregion

            #region Call Sync change operation to change the FileAS element
            string updatedContactFileAs = Common.GenerateResourceName(Site, "UpdatedFileAs");
            Request.SyncCollectionChange appDataChange3 = CreateChangedContact(serverId, new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.FileAs }, new object[] { updatedContactFileAs });
            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, appDataChange3);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The Sync command should be conducted successfully.");
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, contactFileAS);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, updatedContactFileAs);

            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncRequest.RequestData.Collections[0].Supported = supported;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            bool isVerifyR5907 = true;
            Response.SyncCollectionsCollectionCommandsAddApplicationData applicationData2 = TestSuiteBase.GetAddApplicationData(syncResponse, Response.ItemsChoiceType8.FileAs, updatedContactFileAs);
            Site.Assert.IsNotNull(applicationData2, "The application data of the contact item should not be null.");
            Site.Assert.IsNotNull(applicationData2.ItemsElementName, "The ItemsElementName should not be null.");
            Site.Assert.IsTrue(applicationData2.ItemsElementName.Length > 0, "The length of ItemsElementName should be greater then 0.");
            for (int i = 0; i < applicationData2.ItemsElementName.Length; i++)
            {
                if (applicationData2.ItemsElementName[i] == Response.ItemsChoiceType8.JobTitle && applicationData2.Items[i].ToString() == "Vice President")
                {
                    isVerifyR5907 = false;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5907");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5907
            Site.CaptureRequirementIfIsTrue(
                isVerifyR5907,
                5907,
                @"[In Supported] [The status of properties that can be ghosted is determined by the client's usage of the Supported element in the initial Sync command request for the containing folder, according to the following rules:] 2. If the client includes a Supported element that contains child elements in the initial Sync command request for a folder, then each child element of that Supported element is considered not ghosted.");

            bool isFirstNameGhosted = false;
            bool isMiddleNameGhosted = false;
            bool isLastNameGhosted = false;
            foreach (Response.ItemsChoiceType8 name in applicationData2.ItemsElementName)
            {
                if (name == Response.ItemsChoiceType8.FirstName)
                {
                    isFirstNameGhosted = true;
                }
                else if (name == Response.ItemsChoiceType8.MiddleName)
                {
                    isMiddleNameGhosted = true;
                }
                else if (name == Response.ItemsChoiceType8.LastName)
                {
                    isLastNameGhosted = true;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5911");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5911
            Site.CaptureRequirementIfIsTrue(
                isFirstNameGhosted && isMiddleNameGhosted && isLastNameGhosted,
                5911,
                @"[In Supported] [When an existing item is modified via the Change element (section 2.2.3.24) in a Sync command request, the result of omitting an element that can be ghosted changes depending on the status of the element.] If the element is ghosted, any existing value for that element is preserved.");
            #endregion

            #region Add a new contact.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            contactFileAS = Common.GenerateResourceName(this.Site, "FileAS");
            addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, "Vice President");

            syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, addData);
            syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The new contact should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, contactFileAS);
            #endregion

            #region Call Sync command with empty Supported element.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId);
            syncRequest.RequestData.Collections[0].Supported = null;
            this.Sync(syncRequest);

            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncRequest.RequestData.Collections[0].GetChanges = true;
            syncRequest.RequestData.Collections[0].GetChangesSpecified = true;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            serverId = TestSuiteBase.FindServerId(syncResponse, "FileAs", contactFileAS);
            Site.Assert.IsTrue(!string.IsNullOrEmpty(serverId), "The contact should exist in the Contact folder.");
            #endregion

            #region Call Sync change operation to change the JobTitle element
            Request.SyncCollectionChange appDataChange4 = CreateChangedContact(serverId, new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.JobTitle }, new object[] { "President" });
            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, appDataChange4);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The Sync command should be conducted successfully.");

            syncResponse = this.SyncChanges(this.User1Information.ContactsCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Response.SyncCollectionsCollectionCommandsAddApplicationData applicationData3 = TestSuiteBase.GetAddApplicationData(syncResponse, Response.ItemsChoiceType8.FileAs, contactFileAS);
            Site.Assert.IsNotNull(applicationData3, "The application data of the contact item should not be null.");

            isFirstNameGhosted = false;
            isMiddleNameGhosted = false;
            isLastNameGhosted = false;
            foreach (Response.ItemsChoiceType8 itemElementName in applicationData3.ItemsElementName)
            {
                switch (itemElementName)
                {
                    case Response.ItemsChoiceType8.FirstName:
                        isFirstNameGhosted = true;
                        break;
                    case Response.ItemsChoiceType8.MiddleName:
                        isMiddleNameGhosted = true;
                        break;
                    case Response.ItemsChoiceType8.LastName:
                        isLastNameGhosted = true;
                        break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5909");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5909
            Site.CaptureRequirementIfIsFalse(
                isFirstNameGhosted && isMiddleNameGhosted && isLastNameGhosted,
                5909,
                @"[In Supported] [The status of properties that can be ghosted is determined by the client's usage of the Supported element in the initial Sync command request for the containing folder, according to the following rules:] 3. If the client includes an empty Supported element in the initial Sync command request for a folder, then all elements that can be ghosted are considered ghosted.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if the client has changed an item for which the conflict policy indicates that the server's changes take precedence, server will return a status value 7.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC40_Sync_Conflict()
        {
            Site.Assume.AreEqual<string>("Base64", Common.GetConfigurationPropertyValue("HeaderEncodingType", this.Site), "The device ID should be same across all requests, when the HeaderEncodingType is PlainText.");

            // Synchronize the changes in the Contacts folder.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            #region Add a new contact item.
            string contactFileAS = Common.GenerateResourceName(Site, "FileAS");
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, addData);
            this.Sync(syncRequest, false);
            this.FolderSync();

            SyncResponse syncResponse = this.SyncChanges(this.User1Information.ContactsCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            string serverId = TestSuiteBase.FindServerId(syncResponse, "FileAs", contactFileAS);
            Site.Assert.IsNotNull(serverId, "The contact should exist in the Contact folder.");
            string syncKey = this.LastSyncKey;
            #endregion

            // Changed DeviceID
            CMDAdapter.ChangeDeviceID("Device2");
            this.RecordDeviceInfoChanged();

            #region Change the added contact item.
            this.FolderSync();
            syncResponse = this.SyncChanges(this.User1Information.ContactsCollectionId);

            string updatedContactFileAs = Common.GenerateResourceName(Site, "UpdatedFileAS");
            Request.SyncCollectionChange appDataChange1 = CreateChangedContact(TestSuiteBase.FindServerId(syncResponse, "FileAs", contactFileAS), new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.FileAs }, new object[] { updatedContactFileAs });
            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, appDataChange1);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The status code of Sync change operation should be 1.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, updatedContactFileAs);
            #endregion

            #region Restore DeviceID and change the changed contact item again without Conflict option.
            CMDAdapter.ChangeDeviceID(Common.GetConfigurationPropertyValue("DeviceID", this.Site));
            this.FolderSync();

            string conflictUpdatedFileAs = Common.GenerateResourceName(Site, "ConflictUpdatedFileAS");
            Request.SyncCollectionChange appDataChange2 = CreateChangedContact(serverId, new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.FileAs }, new object[] { conflictUpdatedFileAs });
            syncRequest = CreateSyncChangeRequest(syncKey, this.User1Information.ContactsCollectionId, appDataChange2);
            syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The collections in the Sync response should not be null.");

            bool isReplacedByServerData = false;
            Response.SyncCollectionsCollectionCommands commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            Site.Assert.IsNotNull(commands.Change, "The Change element should not be null.");
            Site.Assert.IsTrue(commands.Change.Length > 0, "The Change element should contain one item at least.");
            Site.Assert.IsNotNull(commands.Change[0].ApplicationData, "The ApplicationData should not be null.");
            Site.Assert.IsNotNull(commands.Change[0].ApplicationData.ItemsElementName, "The ItemsElementName of the ApplicationData should not be null.");
            Site.Assert.IsTrue(commands.Change[0].ApplicationData.ItemsElementName.Length > 0, "The ItemsElementName should contains one item at least.");
            for (int i = 0; i < commands.Change[0].ApplicationData.ItemsElementName.Length; i++)
            {
                if (commands.Change[0].ApplicationData.ItemsElementName[i].ToString() == "FileAs" && commands.Change[0].ApplicationData.Items[i].ToString() == updatedContactFileAs)
                {
                    isReplacedByServerData = true;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2066");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2066
            Site.CaptureRequirementIfIsTrue(
                isReplacedByServerData,
                2066,
                @"[In Conflict] If the Conflict element is not present, the server object will replace the client object when a conflict occurs.");
            #endregion

            #region Change the changed contact item again with Conflict option of value 1.
            Request.Options option = new Request.Options
            {
                Items = new object[] { (byte)1 },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Conflict }
            };

            syncRequest.RequestData.Collections[0].Options = new Request.Options[] { option };
            syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The collections in the Sync response should not be null.");

            Response.SyncCollectionsCollectionResponses responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.IsNotNull(responses, "The Response element should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2069");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2069
            Site.CaptureRequirementIfAreEqual<int>(
                7,
                int.Parse(responses.Change[0].Status),
                2069,
                @"[In Conflict] If the value is 1 and there is a conflict, a Status element (section 2.2.3.162.16) value of 7 is returned to inform the client that the object that the client sent to the server was discarded.");

            isReplacedByServerData = false;
            commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            for (int i = 0; i < commands.Change[0].ApplicationData.ItemsElementName.Length; i++)
            {
                if (commands.Change[0].ApplicationData.ItemsElementName[i] == Response.ItemsChoiceType7.FileAs && commands.Change[0].ApplicationData.Items[i].ToString() == updatedContactFileAs)
                {
                    isReplacedByServerData = true;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4444");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4444
            Site.CaptureRequirementIfIsTrue(
                isReplacedByServerData,
                4444,
                @"[In Status(Sync)] [When the scope is Item], [the cause of the status value 7 is] The client has changed an item for which the conflict policy indicates that the server's changes take precedence.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2068");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2068
            Site.CaptureRequirementIfIsTrue(
                isReplacedByServerData,
                2068,
                @"[In Conflict] A value of 1 means to keep the server object.");
            #endregion

            #region Change the changed contact item again with Conflict option of value 0.
            option.Items = new object[] { (byte)0 };
            option.ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Conflict };

            syncRequest.RequestData.Collections[0].Options = new Request.Options[] { option };
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The collections in the Sync response should not be null.");
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The Sync change operation should be successful.");

            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, conflictUpdatedFileAs);

            syncResponse = this.SyncChanges(this.User1Information.ContactsCollectionId);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The collections in the Sync response should not be null");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2067");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2067
            Site.CaptureRequirementIfIsNotNull(
                TestSuiteBase.FindServerId(syncResponse, "FileAs", conflictUpdatedFileAs),
                2067,
                @"[In Conflict] A value of 0 (zero) means to keep the client object.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if it includes more than one Conflict elements as the child of an Options element, the server does not return any error code.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC41_Sync_MoreThanOneConflicts()
        {
            Request.Options option = new Request.Options
            {
                Items = new object[] { (byte)1, (byte)1 },
                ItemsElementName =
                    new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Conflict, Request.ItemsChoiceType1.Conflict }
            };

            Request.SyncCollection collection = new Request.SyncCollection
            {
                Options = new Request.Options[] { option },
                SyncKey = "0",
                CollectionId = this.User1Information.ContactsCollectionId
            };

            Request.Sync syncRequestData = new Request.Sync { Collections = new Request.SyncCollection[] { collection } };

            SyncRequest syncRequest = new SyncRequest { RequestData = syncRequestData };
            SyncResponse syncResponse = this.Sync(syncRequest);

            if (Common.IsRequirementEnabled(2075, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2075");

                // Verify MS-ASCMD requirement: MS-ASCMD_R2075
                Site.CaptureRequirementIfAreEqual<uint>(
                    1,
                    Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                    2075,
                    @"[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one Conflict element as the child of an Options element]. (Exchange 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to verify the combination of the classes in the Sync command.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC42_Sync_WithCombinationClasses()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Class element is not supported as a child element of the Options element when the ActiveSyncProtocolVersion is 12.1.");

            #region Send a MIME-formatted email to User2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            #region Call Sync to get both Email and SMS items
            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);

            Request.Options option1 = new Request.Options
            {
                Items = new object[] { "Email" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            Request.Options option2 = new Request.Options
            {
                Items = new object[] { "SMS" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            SyncResponse syncResponse = this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, new Request.Options[] { option1, option2 });
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items in the Sync response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R934");

            // Verify MS-ASCMD requirement: MS-ASCMD_R934
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)),
                934,
                @"[In Class(Sync)] Only SMS messages and email messages can be synchronized at the same time.");
            #endregion

            #region Call Sync to get both Calendar and Tasks items
            Request.Options option3 = new Request.Options
            {
                Items = new object[] { "Calendar" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            Request.Options option4 = new Request.Options
            {
                Items = new object[] { "Tasks" },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.Class }
            };

            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User2Information.CalendarCollectionId);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].Options = new Request.Options[] { option3, option4 };
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R935");

            // Verify MS-ASCMD requirement: MS-ASCMD_R935
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(syncResponse.ResponseData.Status),
                935,
                @"[In Class(Sync)] A request for any other combination of classes will fail with a status value of 4.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the requirements related to MIMETruncation element.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC43_Sync_MIMETruncation()
        {
            #region Send a MIME-formatted email to User2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            string emailContent = new string('X', 102500);
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, emailContent);
            #endregion

            #region Switch the current user to User2.
            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region Call Sync with MIMETruncation set to 8 to get complete MIME data.
            Request.BodyPreference bodyPreference = new Request.BodyPreference { Type = 4 };

            Request.Options option = new Request.Options
            {
                Items = new object[] { (byte)2, bodyPreference, (byte)8 },
                ItemsElementName =
                    new Request.ItemsChoiceType1[]
                    {
                        Request.ItemsChoiceType1.MIMESupport, Request.ItemsChoiceType1.BodyPreference,
                        Request.ItemsChoiceType1.MIMETruncation
                    }
            };

            SyncResponse syncResponse = this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, new Request.Options[] { option });

            Response.Body mailBody = GetMailBody(syncResponse, emailSubject);
            Site.Assert.IsNotNull(mailBody, "The body of the received email should not be null.");
            int dataLength = mailBody.Data.Length;

            option = new Request.Options
            {
                Items = new object[] { (byte)2, bodyPreference },
                ItemsElementName =
                    new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.MIMESupport, Request.ItemsChoiceType1.BodyPreference }
            };

            syncResponse = this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, new Request.Options[] { option });

            mailBody = GetMailBody(syncResponse, emailSubject);
            Site.Assert.IsNotNull(mailBody, "The body of the received email should not be null.");
            int original = mailBody.Data.Length;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3426");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3426
            Site.CaptureRequirementIfAreEqual<int>(
                original,
                dataLength,
                3426,
                @"[In MIMETruncation] Value 8 means Do not truncate; send complete MIME data.");
            #endregion

            #region Call Sync with MIMETruncation set to 0 to truncate all body text.
            option = new Request.Options
            {
                Items = new object[] { (byte)2, bodyPreference, (byte)0 },
                ItemsElementName =
                    new Request.ItemsChoiceType1[]
                    {
                        Request.ItemsChoiceType1.MIMESupport, Request.ItemsChoiceType1.BodyPreference,
                        Request.ItemsChoiceType1.MIMETruncation
                    }
            };

            syncResponse = this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, new Request.Options[] { option });

            mailBody = GetMailBody(syncResponse, emailSubject);
            Site.Assert.IsNotNull(mailBody, "The body of the received email should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3418");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3418
            Site.CaptureRequirementIfIsNull(
                mailBody.Data,
                3418,
                @"[In MIMETruncation] Value 0 means Truncate all body text.");
            #endregion

            #region Call Sync with MIMETruncation set to 1 to truncate text over 4096 charaters.
            int dataSize = this.GetEmailBodyDataSize(emailSubject, 1);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3419");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3419
            Site.CaptureRequirementIfAreEqual<int>(
                4096,
                dataSize,
                3419,
                @"[In MIMETruncation] Value 1 means Truncate text over 4,096 characters.");
            #endregion

            #region Call Sync with MIMETruncation set to 2 to truncate text over 5120 charaters.
            dataSize = this.GetEmailBodyDataSize(emailSubject, 2);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3420");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3420
            Site.CaptureRequirementIfAreEqual<int>(
                5120,
                dataSize,
                3420,
                @"[In MIMETruncation] Value 2 means Truncate text over 5,120 characters.");
            #endregion

            #region Call Sync with MIMETruncation set to 3 to truncate text over 7168 charaters.
            dataSize = this.GetEmailBodyDataSize(emailSubject, 3);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3421");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3421
            Site.CaptureRequirementIfAreEqual<int>(
                7168,
                dataSize,
                3421,
                @"[In MIMETruncation] Value 3 means Truncate text over 7,168 characters.");
            #endregion

            #region Call Sync with MIMETruncation set to 4 to truncate text over 10240 charaters.
            dataSize = this.GetEmailBodyDataSize(emailSubject, 4);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3422");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3422
            Site.CaptureRequirementIfAreEqual<int>(
                10240,
                dataSize,
                3422,
                @"[In MIMETruncation] Value 4 means Truncate text over 10,240 characters.");
            #endregion

            #region Call Sync with MIMETruncation set to 5 to truncate text over 20480 charaters.
            dataSize = this.GetEmailBodyDataSize(emailSubject, 5);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3423");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3423
            Site.CaptureRequirementIfAreEqual<int>(
                20480,
                dataSize,
                3423,
                @"[In MIMETruncation] Value 5 means Truncate text over 20,480 characters.");
            #endregion

            #region Call Sync with MIMETruncation set to 6 to truncate text over 51200 charaters.
            dataSize = this.GetEmailBodyDataSize(emailSubject, 6);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3424");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3424
            Site.CaptureRequirementIfAreEqual<int>(
                51200,
                dataSize,
                3424,
                @"[In MIMETruncation] Value 6 means Truncate text over 51,200 characters.");
            #endregion

            #region Call Sync with MIMETruncation set to 7 to truncate text over 102400 charaters.
            dataSize = this.GetEmailBodyDataSize(emailSubject, 7);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3425");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3425
            Site.CaptureRequirementIfAreEqual<int>(
                102400,
                dataSize,
                3425,
                @"[In MIMETruncation] Value 7 means Truncate text over 102,400 characters.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the requirements related to MIMESupport element.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC44_Sync_MIMESupport()
        {
            #region Send an email to User2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            string emailContent = new string('X', 102500);
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, emailContent);
            #endregion

            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null);

            Request.BodyPreference bodyPreference = new Request.BodyPreference { Type = 4 };

            Request.Options option = new Request.Options
            {
                Items = new object[] { (byte)2, bodyPreference, (byte)1 },
                ItemsElementName =
                    new Request.ItemsChoiceType1[]
                    {
                        Request.ItemsChoiceType1.MIMESupport, Request.ItemsChoiceType1.BodyPreference,
                        Request.ItemsChoiceType1.MIMETruncation
                    }
            };

            SyncResponse syncResponse = this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, new Request.Options[] { option });
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The Status of the Sync command response should be 1.");
            Response.Body mailBody = GetMailBody(syncResponse, emailSubject);
            Site.Assert.IsNotNull(mailBody, "The body of the received email should not be null.");
            Site.Assert.AreEqual<byte>(4, mailBody.Type, "The type of the Body should be 4.");
            Site.Assert.IsTrue(mailBody.EstimatedDataSizeSpecified, "The EstimatedDataSize element should be present.");
            Site.Assert.IsTrue(mailBody.TruncatedSpecified, "The Truncated element should be present.");
            Site.Assert.IsNotNull(mailBody.Data, "The Data element of the Body should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3409");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3409
            // When the Assert statements above are passed, the requirement is captured.
            Site.CaptureRequirement(
                3409,
                @"[In MIMESupport(Sync)] [The airsyncbase:Body element] MUST contain the following child elements [the airsyncbase:Type element, the airsyncbase:EstimatedDataSize element, the airsyncbase:Truncated element, the airsyncbase:Data element] in an S/MIME synchronization response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3410");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3410
            // When the Assert statements above are passed, the requirement is captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                4,
                mailBody.Type,
                3410,
                @"[In MIMESupport(Sync)] The airsyncbase:Type element ([MS-ASAIRS] section 2.2.2.22.1) with a value of 4 to inform the device that the data is a MIME BLOB.");
        }

        /// <summary>
        /// This test case is designed to verify the ConversationMode element related requirements.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC45_Sync_ConversationMode()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The airsync:ConversationMode element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send a MIME-formatted email to User2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            #region Switch current user to User2
            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null);
            #endregion

            #region Call Sync command with ConversationMode set to true.
            Request.SyncCollection collection = new Request.SyncCollection
            {
                SyncKey = this.LastSyncKey,
                Options = new Request.Options[]
                {
                    new Request.Options
                    {
                        Items = new object[] { (byte)1 },
                        ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.FilterType }
                    }
                },
                CollectionId = this.User2Information.InboxCollectionId,
                Commands = null,
                GetChanges = true,
                GetChangesSpecified = true,
                ConversationMode = true,
                ConversationModeSpecified = true
            };

            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User2Information.InboxCollectionId);
            this.Sync(syncRequest);
            collection.SyncKey = this.LastSyncKey;
            syncRequest.RequestData.Collections = new Request.SyncCollection[] { collection };
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject);
            Site.Assert.IsTrue(!string.IsNullOrEmpty(serverId), "The new email should be included in the response of the Sync command with ConversationMode.");

            Response.SyncCollectionsCollectionCommands commandsByConversationMode = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            Site.Assert.IsNotNull(commandsByConversationMode.Add, "The Add element returned in the Sync command response should not be null.");
            List<string> serverIdsByConversationMode = new List<string>();

            foreach (Response.SyncCollectionsCollectionCommandsAdd add in commandsByConversationMode.Add)
            {
                serverIdsByConversationMode.Add(add.ServerId);
            }

            // Call Sync with FilterType set to 1 to get items that are dated within 1 day in Inbox folder.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User2Information.InboxCollectionId, (byte)1);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Response.SyncCollectionsCollectionCommands commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            Site.Assert.IsNotNull(commands.Add, "The Add element returned in the Sync command response should not be null.");
            List<string> serverIds = new List<string>();

            foreach (Response.SyncCollectionsCollectionCommandsAdd add in commands.Add)
            {
                serverIds.Add(add.ServerId);
            }

            Site.Assert.IsTrue(serverIdsByConversationMode.Count > 0, "There should be one item in serverIdsByConversationMode at least.");

            bool isVerifyR2107 = true;
            foreach (string item in serverIdsByConversationMode)
            {
                if (!serverIds.Contains(item))
                {
                    isVerifyR2107 = false;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2107");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2107
            Site.CaptureRequirementIfIsTrue(
                isVerifyR2107,
                2107,
                @"[In ConversationMode(Sync)] Setting the ConversationMode element value to 1 (TRUE) results in retrieving all emails that match the conversations received within the date filter specified.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if the CollectionId element is not included in the Sync request, then the status code in the response is 4.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC46_Sync_Supported_Status4()
        {
            #region Add a new contact.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            string contactFileAS = Common.GenerateResourceName(Site, "FileAS");
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, "Vice President");

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.ContactsCollectionId, addData);
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Response.SyncCollectionsCollectionResponses responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The new contact should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.ContactsCollectionId, contactFileAS);
            this.FolderSync();
            #endregion

            #region Call Sync command without specifying CollectionId.
            Request.Supported supported = new Request.Supported
            {
                ItemsElementName = new Request.ItemsChoiceType[] { Request.ItemsChoiceType.JobTitle }
            };

            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId);
            syncRequest.RequestData.Collections[0].CollectionId = null;
            syncRequest.RequestData.Collections[0].Supported = supported;
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData, "The response data returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4545");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4545
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(syncResponse.ResponseData.Status),
                4545,
                @"[In Supported] A Status element (section 2.2.3.162.16) value of 4 is returned in the Sync response if the CollectionId element is not included in the Sync request.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if the Wait element is outside the range set, or smaller than the minimum allowable value, then the status code in the server response is 14.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC47_Sync_Wait_Status14()
        {
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            syncRequest.RequestData.Wait = "60";
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5781");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5781
            Site.CaptureRequirementIfIsTrue(
                syncResponse.ResponseData.Status == "14" && syncResponse.ResponseData.Item.ToString() == "59" && syncResponse.ResponseDataXML.ToString().Contains("Limit"),
                5781,
                @"[In Status(Sync)] [When the scope is Item], [the cause of the status value 14 is] If the [HeartbeatInterval element value] or Wait element value included in the Sync request is larger than the maximum allowable value, the response contains a Limit element that specifies the maximum allowed value.");

            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            syncRequest.RequestData.Wait = "0";
            syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5782");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5782
            Site.CaptureRequirementIfIsTrue(
                syncResponse.ResponseData.Status == "14" && syncResponse.ResponseData.Item.ToString() == "1" && syncResponse.ResponseDataXML.ToString().Contains("Limit"),
                5782,
                @"[In Status(Sync)] [When the scope is Item], [the cause of the status value 14 is] If the [HeartbeatInterval element value or] Wait value included in the Sync request is smaller than the minimum allowable value, the response contains a Limit element that specifies the minimum allowed value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4746");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4746
            // According to above two steps, this requirement can be covered directly.
            Site.CaptureRequirement(
                4746,
                @"[In Wait] When the client requests a wait interval that is outside the acceptable range, the server will send a response that includes a Status element (section 2.2.3.162.16) value of 14 and a Limit element (section 2.2.3.88).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4459");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4459
            // According to above two steps, this requirement can be covered directly.
            Site.CaptureRequirement(
                4459,
                @"[In Status(Sync)] [When the scope is Item], [the cause of the status value 14 is] The Sync request was processed successfully but the wait interval (Wait element value (section 2.2.3.182)) [or heartbeat interval (HeartbeatInterval element value (section 2.2.3.79.2))] that is specified by the client is outside the range set by the server administrator.");
        }

        /// <summary>
        /// This test case is used to verify the limit value of the sum of the number of Add, Change, Delete, and Fetch elements of Sync command.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC48_Sync_Collection_LimitValue()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5671, this.Site) || Common.IsRequirementEnabled(5673, this.Site), "Update Rollup 6 for Exchange 2010 SP2 and Exchange 2013 use the specified limit values by default.");

            #region Create 51 Add elements, 50 Change elements, 50 Delete elements and 50 Fetch elements for the Commands element of Sync command.
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.ContactsCollectionId));

            List<object> commands = new List<object>();

            for (int i = 0; i < 50; i++)
            {
                string contactFileAS = Common.GenerateResourceName(Site, "FileAS");
                Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);
                commands.Add(addData);

                string serverId = this.User1Information.ContactsCollectionId + ":" + Guid.NewGuid().ToString();
                string updatedContactFileAS = Common.GenerateResourceName(Site, "UpdatedFileAS");
                Request.SyncCollectionChange changeData = CreateChangedContact(serverId, new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.FileAs }, new object[] { updatedContactFileAS });
                commands.Add(changeData);

                Request.SyncCollectionDelete deleteData = new Request.SyncCollectionDelete { ServerId = serverId };
                commands.Add(deleteData);

                Request.SyncCollectionFetch fetchData = new Request.SyncCollectionFetch { ServerId = serverId };
                commands.Add(fetchData);
            }

            string contactFileASFor51Add = Common.GenerateResourceName(Site, "FileAS");
            Request.SyncCollectionAdd addDataFor51Add = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileASFor51Add, null);
            commands.Add(addDataFor51Add);
            #endregion

            #region Call Sync command on the Contacts folder
            Request.SyncCollection collection = new Request.SyncCollection
            {
                SyncKey = this.LastSyncKey,
                GetChanges = true,
                CollectionId = this.User1Information.ContactsCollectionId,
                Commands = commands.ToArray()
            };

            SyncRequest syncRequest = Common.CreateSyncRequest(new Request.SyncCollection[] { collection });
            SyncResponse syncResponse = this.Sync(syncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5659");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5659
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(syncResponse.ResponseData.Status),
                5659,
                @"[In Limiting Size of Command Requests] In Sync (section 2.2.2.19) command request, when the limit value of Add, Change, Delete and Fetch elements is bigger than 200 (minimum 1, maximum 2,147,483,647), the error returned by server is Status element (section 2.2.3.162.16) value of 4.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the requirements related to SoftDelete element.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC49_Sync_SoftDelete()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Class element is not supported in a Sync command response when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Add a new calendar before 14 days
            string calendarSubject = Common.GenerateResourceName(Site, "calendarSubject");
            string location = Common.GenerateResourceName(Site, "Room");
            DateTime startTime = DateTime.Now.AddDays(-16);
            DateTime endTime = startTime.AddHours(1.0);
            Request.SyncCollectionAdd calendarData;

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0")&&!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"))
            {
                calendarData = new Request.SyncCollectionAdd
                {
                    ClientId = TestSuiteBase.ClientId,
                    ApplicationData = new Request.SyncCollectionAddApplicationData
                    {
                        ItemsElementName =
                            new Request.ItemsChoiceType8[]
                        {
                            Request.ItemsChoiceType8.Subject, Request.ItemsChoiceType8.Location1,
                            Request.ItemsChoiceType8.StartTime, Request.ItemsChoiceType8.EndTime,
                            Request.ItemsChoiceType8.UID
                        },
                        Items =
                            new object[]
                        {
                            calendarSubject, location, startTime.ToString("yyyyMMddTHHmmssZ"),
                            endTime.ToString("yyyyMMddTHHmmssZ"), Guid.NewGuid().ToString()
                        }
                    },
                    Class = "Calendar"
                };
            }
            else
            {
                calendarData = new Request.SyncCollectionAdd
                {
                    ClientId = TestSuiteBase.ClientId,
                    ApplicationData = new Request.SyncCollectionAddApplicationData
                    {
                        ItemsElementName =
                            new Request.ItemsChoiceType8[]
                        {
                            Request.ItemsChoiceType8.Subject, Request.ItemsChoiceType8.Location,
                            Request.ItemsChoiceType8.StartTime, Request.ItemsChoiceType8.EndTime
                        },
                        Items =
                            new object[]
                        {
                            calendarSubject, 
                            new Request.Location
                            {
                            LocationUri=location
                            }, 
                            startTime.ToString("yyyyMMddTHHmmssZ"),
                            endTime.ToString("yyyyMMddTHHmmssZ")
                        }
                    },
                    Class = "Calendar"
                };
            }

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId));

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, calendarData);
            SyncResponse syncResponse = this.Sync(syncRequest, true);
            Response.SyncCollectionsCollectionResponses responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The calendar should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, calendarSubject);
            #endregion

            #region Add a new calendar within 14 days
            calendarSubject = Common.GenerateResourceName(this.Site, "calendarSubject");
            location = Common.GenerateResourceName(this.Site, "Room");
            startTime = DateTime.Now.AddDays(-10);
            endTime = startTime.AddHours(1.0);

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0")&& !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"))
            {
                calendarData = new Request.SyncCollectionAdd
                {
                    ClientId = TestSuiteBase.ClientId,
                    ApplicationData = new Request.SyncCollectionAddApplicationData
                    {
                        ItemsElementName =
                            new Request.ItemsChoiceType8[]
                        {
                            Request.ItemsChoiceType8.Subject, Request.ItemsChoiceType8.Location1,
                            Request.ItemsChoiceType8.StartTime, Request.ItemsChoiceType8.EndTime,
                            Request.ItemsChoiceType8.UID
                        },
                        Items =
                            new object[]
                        {
                            calendarSubject, location, startTime.ToString("yyyyMMddTHHmmssZ"),
                            endTime.ToString("yyyyMMddTHHmmssZ"), Guid.NewGuid().ToString()
                        }
                    },
                    Class = "Calendar"
                };
            }
            else
            {
                calendarData = new Request.SyncCollectionAdd
                {
                    ClientId = TestSuiteBase.ClientId,
                    ApplicationData = new Request.SyncCollectionAddApplicationData
                    {
                        ItemsElementName =
                            new Request.ItemsChoiceType8[]
                        {
                            Request.ItemsChoiceType8.Subject, Request.ItemsChoiceType8.Location,
                            Request.ItemsChoiceType8.StartTime, Request.ItemsChoiceType8.EndTime
                        },
                        Items =
                            new object[]
                        {
                            calendarSubject, 
                            new Request.Location
                            {
                            LocationUri=location
                            }, 
                            startTime.ToString("yyyyMMddTHHmmssZ"),
                            endTime.ToString("yyyyMMddTHHmmssZ")
                        }
                    },
                    Class = "Calendar"
                };
            }

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId));

            syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, calendarData);
            syncResponse = this.Sync(syncRequest, true);
            responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The calendar should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, calendarSubject);
            #endregion

            #region Call Sync command with FilterType of 4 to get the number of the filtered-out SoftDelete elements.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 4);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The Item element in the Sync response should not be null.");

            Response.SyncCollectionsCollectionCommands commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            Site.Assert.IsNotNull(commands.SoftDelete, "The SoftDelete element in the commands should not be null.");
            int filteredOut = commands.SoftDelete.Length;
            #endregion

            #region Call Sync with FilterType of 0 to get all items in the Calendar folder.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 0);
            this.Sync(syncRequest, false);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            Site.Assert.IsNotNull(commands.Add, "The Add element in the commands should not be null.");
            int total = commands.Add.Length;
            #endregion

            #region Call Sync command with FilterType of 4 to get the number of the filtered-in Add elements.
            syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId, 4);
            this.Sync(syncRequest, false);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");

            commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            Site.Assert.IsNotNull(commands.Add, "The Add element in the commands should not be null.");
            int filterdIn = commands.Add.Length;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3967");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3967
            Site.CaptureRequirementIfAreEqual<int>(
                total - filterdIn,
                filteredOut,
                3967,
                @"[In SoftDelete] The SoftDelete element contains any items that are filtered out of the Sync query due to being outside the FilterType date range [, or no longer specified as part of the SyncOptions instructions].");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the Class element in the Sync Add response.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC50_Sync_Class()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Class element is not supported in a Sync command response when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Add a new Note item
            string noteSubject = Common.GenerateResourceName(Site, "noteSubject");
            Request.Body noteBody = new Request.Body { Type = 1, Data = "Content of the body." };
            Request.Categories4 categories = new Request.Categories4 { Category = new string[] { "blue category" } };

            Request.SyncCollectionAdd noteData = new Request.SyncCollectionAdd
            {
                ClientId = TestSuiteBase.ClientId,
                ApplicationData = new Request.SyncCollectionAddApplicationData
                {
                    ItemsElementName =
                        new Request.ItemsChoiceType8[]
                        {
                            Request.ItemsChoiceType8.Subject1, 
                            Request.ItemsChoiceType8.Body,
                            Request.ItemsChoiceType8.Categories2, 
                            Request.ItemsChoiceType8.MessageClass
                        },
                    Items =
                        new object[]
                        {
                            noteSubject, 
                            noteBody, 
                            categories,
                            "IPM.StickyNote"
                        }
                },
                Class = "Calendar"
            };

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.NotesCollectionId));

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.NotesCollectionId, noteData);
            SyncResponse syncResponse = this.Sync(syncRequest, false);

            Response.SyncCollectionsCollectionResponses responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R941");

            // Verify MS-ASCMD requirement: MS-ASCMD_R941
            Site.CaptureRequirementIfAreEqual<string>(
                "Calendar",
                responses.Add[0].Class,
                941,
                @"[In Class(Sync)] As a child element of the Add element in the Sync command response, the Class element<20> identifies the class of the item being added to the collection.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_9945");

            // Verify MS-ASCMD requirement: MS-ASCMD_R9945
            Site.CaptureRequirementIfAreEqual<string>(
                "Calendar",
                responses.Add[0].Class,
                9945,
                @"[In Class(Sync)] In all contexts of a Sync command request or Sync command response, the valid Class element value is Calendar.");
        }

        /// <summary>
        /// This test case is used to verify Sync command, if Exceptions element is not specified or not present in the sync change request, the original Exceptions element will remain unchanged.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC51_Sync_Change_Exceptions()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Class element is not supported in a Sync command response when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Recurrences cannot be added in protocol version 16.0");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Recurrences cannot be added in protocol version 16.1");

            this.Sync(TestSuiteBase.CreateEmptySyncRequest(this.User1Information.CalendarCollectionId));

            #region Call Sync Add operation to add a new recurrence calendar.
            string recurrenceCalendarSubject = Common.GenerateResourceName(Site, "calendarSubject");
            string location = Common.GenerateResourceName(Site, "Room");
            DateTime currentDate = DateTime.Now.AddDays(1);
            DateTime startTime = new DateTime(currentDate.Year, currentDate.Month, currentDate.Day, 10, 0, 0);
            DateTime endTime = startTime.AddHours(10);

            Request.ExceptionsException exception = new Request.ExceptionsException();
            exception.ExceptionStartTime = startTime.AddDays(2).ToString("yyyyMMddTHHmmssZ");
            Request.Exceptions exceptions = new Request.Exceptions() { Exception = new Request.ExceptionsException[] { exception } };

            Request.Recurrence recurrence = new Request.Recurrence
            {
                Type = 0
            };

            Request.SyncCollectionAdd recurrenceCalendarData = new Request.SyncCollectionAdd
            {
                ClientId = TestSuiteBase.ClientId,
                ApplicationData =
                    new Request.SyncCollectionAddApplicationData
                    {
                        ItemsElementName =
                            new Request.ItemsChoiceType8[] 
                            { 
                                Request.ItemsChoiceType8.Subject, Request.ItemsChoiceType8.Location1,
                                Request.ItemsChoiceType8.StartTime, Request.ItemsChoiceType8.EndTime,
                                Request.ItemsChoiceType8.Recurrence, Request.ItemsChoiceType8.Exceptions,
                                Request.ItemsChoiceType8.UID
                            },
                        Items =
                        new object[] 
                        { 
                            recurrenceCalendarSubject, location, 
                            startTime.ToString("yyyyMMddTHHmmssZ"),
                            endTime.ToString("yyyyMMddTHHmmssZ"),
                            recurrence, exceptions, Guid.NewGuid().ToString()
                        }
                    },
                Class = "Calendar"
            };

            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, recurrenceCalendarData);
            SyncResponse syncResponse = this.Sync(syncRequest);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Response.SyncCollectionsCollectionResponses responses = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.AreEqual<int>(1, int.Parse(responses.Add[0].Status), "The calendar should be added successfully.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, recurrenceCalendarSubject);

            syncResponse = this.SyncChanges(this.User1Information.CalendarCollectionId);
            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", recurrenceCalendarSubject);
            Site.Assert.IsNotNull(serverId, "The recurrence calendar should be found.");
            #endregion

            #region Change the subject of the added recurrence calendar.
            string updatedCalendarSubject = Common.GenerateResourceName(Site, "updatedCalendarSubject");

            Request.SyncCollectionChangeApplicationData changeCalednarData = new Request.SyncCollectionChangeApplicationData();
            changeCalednarData.ItemsElementName = new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.Subject, Request.ItemsChoiceType7.Recurrence };
            changeCalednarData.Items = new object[] { updatedCalendarSubject, recurrence };

            Request.SyncCollectionChange appDataChange = new Request.SyncCollectionChange
            {
                ApplicationData = changeCalednarData,
                ServerId = serverId
            };

            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, appDataChange);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The FileAs of the contact should be updated successfully.");
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, recurrenceCalendarSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, updatedCalendarSubject);

            syncResponse = this.SyncChanges(this.User1Information.CalendarCollectionId);
            serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", updatedCalendarSubject);
            Site.Assert.IsNotNull(serverId, "The recurrence calendar should be found.");

            Response.SyncCollectionsCollectionCommands commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            Site.Assert.IsNotNull(commands.Add, "The Add element should not be null.");

            foreach (Response.SyncCollectionsCollectionCommandsAdd item in commands.Add)
            {
                if (item.ServerId == serverId)
                {
                    for (int i = 0; i < item.ApplicationData.ItemsElementName.Length; i++)
                    {
                        if (item.ApplicationData.ItemsElementName[i] == Response.ItemsChoiceType8.Exceptions)
                        {
                            Response.Exceptions currentExceptions = item.ApplicationData.Items[i] as Response.Exceptions;
                            Site.Assert.IsNotNull(currentExceptions, "The Exceptions element should exist.");

                            Response.ExceptionsException currentException = currentExceptions.Exception[0];

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R877");

                            // Verify MS-ASCMD requirement: MS-ASCMD_R877
                            Site.CaptureRequirementIfAreEqual<string>(
                                exception.ExceptionStartTime.ToString(),
                                currentException.ExceptionStartTime.ToString(),
                                877,
                                @"[In Change] If a calendar:Exception ([MS-ASCAL] section 2.2.2.19) node within the calendar:Exceptions node is not present, that particular exception will remain unchanged.");

                            break;
                        }
                    }

                    break;
                }
            }
            #endregion

            #region Change the subject of the added recurrence calendar again.
            string allNewCalendarSubject = Common.GenerateResourceName(Site, "updatedCalendarSubject");

            changeCalednarData.ItemsElementName = new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.Subject, Request.ItemsChoiceType7.Recurrence, Request.ItemsChoiceType7.Exceptions, Request.ItemsChoiceType7.UID };
            changeCalednarData.Items = new object[] { allNewCalendarSubject, recurrence, null, Guid.NewGuid().ToString() };

            appDataChange = new Request.SyncCollectionChange
            {
                ApplicationData = changeCalednarData,
                ServerId = serverId
            };

            syncRequest = CreateSyncChangeRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, appDataChange);
            syncResponse = this.Sync(syncRequest);
            Site.Assert.AreEqual<uint>(1, Convert.ToUInt32(TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Status)), "The FileAs of the contact should be updated successfully.");
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, updatedCalendarSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, allNewCalendarSubject);

            syncResponse = this.SyncChanges(this.User1Information.CalendarCollectionId);
            serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", allNewCalendarSubject);
            Site.Assert.IsNotNull(serverId, "The recurrence calendar should be found.");

            commands = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Commands) as Response.SyncCollectionsCollectionCommands;
            Site.Assert.IsNotNull(commands.Add, "The Add element should not be null.");

            foreach (Response.SyncCollectionsCollectionCommandsAdd item in commands.Add)
            {
                if (item.ServerId == serverId)
                {
                    for (int i = 0; i < item.ApplicationData.ItemsElementName.Length; i++)
                    {
                        if (item.ApplicationData.ItemsElementName[i] == Response.ItemsChoiceType8.Exceptions)
                        {
                            Response.Exceptions currentExceptions = item.ApplicationData.Items[i] as Response.Exceptions;
                            Site.Assert.IsNotNull(currentExceptions, "The Exceptions element should exist.");

                            Response.ExceptionsException currentException = currentExceptions.Exception[0];

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R876");

                            // Verify MS-ASCMD requirement: MS-ASCMD_R876
                            Site.CaptureRequirementIfAreEqual<string>(
                                exception.ExceptionStartTime.ToString(),
                                currentException.ExceptionStartTime.ToString(),
                                876,
                                @"[In Change] [Certain in-schema properties remain untouched in the following three cases:] If a calendar:Exceptions ([MS-ASCAL] section 2.2.2.20) node is not specified, the properties for that calendar:Exceptions node will remain unchanged.");

                            break;
                        }
                    }

                    break;
                }
            }
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Sync command, if the client issued a fetch or change operation that has a CollectionId value that is no longer valid on the server, then the status code in the server response will be 8.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S19_TC52_Sync_Status8()
        {
            #region Send a MIME-formatted email to User2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            SyncResponse syncResponse = this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, null);
            string itemServerId = TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject);

            #region Delete the email form Inbox of User2.
            SyncRequest syncRequest = TestSuiteBase.CreateSyncDeleteRequest(this.LastSyncKey, this.User2Information.InboxCollectionId, itemServerId);
            syncRequest.RequestData.Collections[0].DeletesAsMoves = false;
            syncRequest.RequestData.Collections[0].DeletesAsMovesSpecified = true;
            syncResponse = this.Sync(syncRequest);
            #endregion

            #region Fetch the email.
            Request.SyncCollectionFetch appDataFetch = new Request.SyncCollectionFetch
            {
                ServerId = itemServerId
            };

            Request.SyncCollection collection = new Request.SyncCollection
            {
                SyncKey = this.LastSyncKey,
                GetChanges = true,
                GetChangesSpecified = true,
                CollectionId = this.User2Information.InboxCollectionId,
                Commands = new object[] { appDataFetch }
            };

            Request.Sync syncRequestData = new Request.Sync { Collections = new Request.SyncCollection[] { collection } };

            SyncRequest syncRequestForFetch = new SyncRequest { RequestData = syncRequestData };
            syncResponse = this.Sync(syncRequestForFetch);
            Response.SyncCollectionsCollectionResponses collectionResponse = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;

            this.Site.CaptureRequirementIfAreEqual<int>(
                8,
                int.Parse(collectionResponse.Fetch.Status),
                4447,
                @"[In Status(Sync)] [When the scope is Item], [the cause of the status value 8 is] The client issued a fetch [or change] operation that has a CollectionId (section 2.2.3.30.5) value that is no longer valid on the server (for example, the item was deleted).");

            #endregion
        }
        #endregion

        #region Private Static Methods
        /// <summary>
        /// Create a Sync change operation request.
        /// </summary>
        /// <param name="syncKey">The synchronization state of a collection.</param>
        /// <param name="collectionId">The server ID of the folder.</param>
        /// <param name="syncCollectionChange">An instance of the SyncCollectionChange.</param>
        /// <returns>The Sync change operation request.</returns>
        private static SyncRequest CreateSyncChangeRequest(string syncKey, string collectionId, Request.SyncCollectionChange syncCollectionChange)
        {
            Request.SyncCollection collection = new Request.SyncCollection
            {
                SyncKey = syncKey,
                GetChanges = true,
                CollectionId = collectionId,
                Commands = new object[] { syncCollectionChange }
            };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { collection });
        }

        /// <summary>
        /// Create a change contact request.
        /// </summary>
        /// <param name="serverId">Server id of the contact.</param>
        /// <param name="itemElementName">The name of the item element.</param>
        /// <param name="items">The value of the item element.</param>
        /// <returns>The change contact request.</returns>
        private static Request.SyncCollectionChange CreateChangedContact(string serverId, Request.ItemsChoiceType7[] itemElementName, object[] items)
        {
            Request.SyncCollectionChange appData = new Request.SyncCollectionChange
            {
                ServerId = serverId,
                ApplicationData = new Request.SyncCollectionChangeApplicationData
                {
                    ItemsElementName = itemElementName,
                    Items = items
                }
            };
            return appData;
        }

        /// <summary>
        /// Create an add Email request.
        /// </summary>
        /// <param name="to">The value of the To element.</param>
        /// <param name="clientId">The value of the ClientId element.</param>
        /// <returns>The add Email request.</returns>
        private static Request.SyncCollectionAdd CreateAddEmailCommand(string to, string clientId)
        {
            Request.SyncCollectionAdd appData = new Request.SyncCollectionAdd
            {
                ClientId = clientId,
                ApplicationData = new Request.SyncCollectionAddApplicationData
                {
                    ItemsElementName = new Request.ItemsChoiceType8[] { Request.ItemsChoiceType8.To },
                    Items = new object[] { to }
                },
                Class = "Email"
            };
            return appData;
        }

        /// <summary>
        /// Check whether an element of ItemsChoiceType10 type appears in a Sync response.
        /// </summary>
        /// <param name="syncResponse">A Sync response.</param>
        /// <param name="element">The element to be checked.</param>
        /// <returns>Return true, if the element exists in the Sync response; otherwise, return false.</returns>
        private static bool CheckElementOfItemsChoiceType10(SyncResponse syncResponse, Response.ItemsChoiceType10 element)
        {
            if (TestSuiteBase.GetCollectionItem(syncResponse, element) != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Get the body of an email.
        /// </summary>
        /// <param name="syncResponse">A Sync command response.</param>
        /// <param name="emailSubject">The email subject.</param>
        /// <returns>The body part of the email.</returns>
        private static Response.Body GetMailBody(SyncResponse syncResponse, string emailSubject)
        {
            Response.Body mailBody = null;
            Response.SyncCollectionsCollectionCommandsAddApplicationData applicationData = TestSuiteBase.GetAddApplicationData(syncResponse, Response.ItemsChoiceType8.Subject1, emailSubject);
            for (int i = 0; i < applicationData.ItemsElementName.Length; i++)
            {
                if (applicationData.ItemsElementName[i] == Response.ItemsChoiceType8.Body)
                {
                    mailBody = applicationData.Items[i] as Response.Body;
                    break;
                }
            }

            return mailBody;
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Get the size of the Data of an email's Body.
        /// </summary>
        /// <param name="emailSubject">The email's subject.</param>
        /// <param name="valueOfMIMETruncation">The value of MIMETruncation, which is in the range of [1,7].</param>
        /// <returns>The size of the Data of the email's Body.</returns>
        private int GetEmailBodyDataSize(string emailSubject, byte valueOfMIMETruncation)
        {
            Site.Assert.IsTrue(!string.IsNullOrEmpty(emailSubject), "The email subject should not be null or empty.");
            Site.Assert.IsTrue(0 < valueOfMIMETruncation && valueOfMIMETruncation < 8, "The value of MIMETruncation should be in the range of [1,7]");

            Request.BodyPreference bodyPreference = new Request.BodyPreference { Type = 4 };

            Request.Options option = new Request.Options
            {
                Items = new object[] { (byte)2, bodyPreference, valueOfMIMETruncation },
                ItemsElementName =
                    new Request.ItemsChoiceType1[]
                    {
                        Request.ItemsChoiceType1.MIMESupport, Request.ItemsChoiceType1.BodyPreference,
                        Request.ItemsChoiceType1.MIMETruncation
                    }
            };

            SyncResponse syncResponse = this.CheckEmail(this.User2Information.InboxCollectionId, emailSubject, new Request.Options[] { option });

            Response.Body mailBody = GetMailBody(syncResponse, emailSubject);
            Site.Assert.IsNotNull(mailBody, "The body of the received email should not be null.");
            Site.Assert.IsNotNull(mailBody.Data, "The Data of the received email's body should not be null.");

            return mailBody.Data.Length;
        }

        /// <summary>
        /// Create an add Calendar request.
        /// </summary>
        /// <param name="to">Recipient of the calendar.</param>
        /// <param name="subject">Subject of the calendar.</param>
        /// <param name="location">Location of the calendar.</param>
        /// <param name="endTime">End time of the calendar.</param>
        /// <returns>The add Calendar request.</returns>
        private Request.SyncCollectionAdd CreateAddCalendarCommand(string to, string subject, string location, string endTime)
        {
            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0"))
            {
                Request.SyncCollectionAdd appData = new Request.SyncCollectionAdd
                {
                    ClientId = TestSuiteBase.ClientId,
                    ApplicationData = new Request.SyncCollectionAddApplicationData
                    {
                        ItemsElementName =
                            new Request.ItemsChoiceType8[]
                        {
                            Request.ItemsChoiceType8.To, Request.ItemsChoiceType8.Subject,
                            Request.ItemsChoiceType8.Location1,
                            Request.ItemsChoiceType8.EndTime
                        },
                        Items = new object[] { to, subject, location, endTime }
                    },
                    Class = "Calendar"
                };
                return appData;
            }

            else
            {
                Request.SyncCollectionAdd appData = new Request.SyncCollectionAdd
                {
                    ClientId = TestSuiteBase.ClientId,
                    ApplicationData = new Request.SyncCollectionAddApplicationData
                    {
                        ItemsElementName =
                            new Request.ItemsChoiceType8[]
                        {
                            Request.ItemsChoiceType8.To, Request.ItemsChoiceType8.Subject,
                            Request.ItemsChoiceType8.Location,
                            Request.ItemsChoiceType8.EndTime
                        },
                        Items = new object[] { 
                        to, 
                        subject, 
                        new Request.Location
                        {
                        LocationUri=location
                        }, 
                        endTime }
                    },
                    Class = "Calendar"
                };
                return appData;
            }
        }

        /// <summary>
        /// Synchronize changes with an Option object.
        /// </summary>
        /// <param name="collectionId">The target collection Id.</param>
        /// <param name="option">The Option object used in Sync request.</param>
        /// <returns>The Sync response.</returns>
        private SyncResponse SyncChangesWithOption(string collectionId, Request.Options option)
        {
            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(collectionId);
            this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            syncRequest.RequestData.Collections[0].Options = new Request.Options[] { option };
            SyncResponse syncResponse = this.Sync(syncRequest);
            return syncResponse;
        }
        #endregion
    }
}