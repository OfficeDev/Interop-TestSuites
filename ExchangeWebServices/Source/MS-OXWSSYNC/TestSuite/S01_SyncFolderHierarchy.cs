namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test SyncFolderHierarchy operation on the following folders: inbox folder, calendar folder, contacts folder, tasks folder and search folder.
    /// </summary>
    [TestClass]
    public class S01_SyncFolderHierarchy : TestSuiteBase
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
        /// Client calls SyncFolderHierarchy operation to sync inbox folder.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S01_TC01_SyncFolderHierarchy_InboxFolder()
        {
            #region Step 1. Client invokes SyncFolderHierarchy operation to get initial syncState of inbox folder.
            DistinguishedFolderIdNameType inboxFolder = DistinguishedFolderIdNameType.inbox;

            // Set DefaultShapeNamesType to AllProperties and include SyncFolderId and SyncState element in the request
            SyncFolderHierarchyType request = TestSuiteHelper.CreateSyncFolderHierarchyRequest(inboxFolder, DefaultShapeNamesType.AllProperties, true, true);

            SyncFolderHierarchyResponseType response = this.SYNCAdapter.SyncFolderHierarchy(request);
            SyncFolderHierarchyResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);
            #endregion

            #region Step 2. Client invokes CreateFolder operation to create two level folders in inbox folder.
            // Generate the created folder name
            string firstLevelFolderName = Common.GenerateResourceName(this.Site, inboxFolder + "FirstLevelFolder");
            string secondLevelFolderName = Common.GenerateResourceName(this.Site, inboxFolder + "SecondLevelFolder");

            // Create two level folders in inbox folder.
            bool isSubFolderCreated = this.FOLDSUTControlAdapter.CreateSubFolders(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                inboxFolder.ToString(),
                firstLevelFolderName,
                secondLevelFolderName);
            Site.Assert.IsTrue(
                isSubFolderCreated,
                string.Format("The sub folders in '{0}' should be created successfully.", inboxFolder));
            #endregion

            #region Step 3. Client invokes SyncFolderHierarchy operation to sync the operation result in Step 2 and verify related requirements.
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There are two folders created on server, so the changes between server and client should not be null");
            SyncFolderHierarchyChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there are two folders created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be folders information returned in SyncFolderHierarchy response since there are two folders created on server.");

            bool isFolderCreated = false;
            for (int i = 0; i < changes.Items.Length; i++)
            {
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Expected value: ItemsElementName: {0}, folder type: {1}; actual value: ItemsElementName: {2}, folder type: {3}",
                    ItemsChoiceType.Create,
                    typeof(FolderType),
                    changes.ItemsElementName[i],
                    (changes.Items[i] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType());

                if (changes.ItemsElementName[i] == ItemsChoiceType.Create && (changes.Items[i] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType() == typeof(FolderType))
                {
                    isFolderCreated = true;
                }
                else
                {
                    isFolderCreated = false;
                    break;
                }
            }

            // If the ItemsElementName of Changes is Create and the type of Item is FolderType, it indicates a regular folder has been created on server and synced on client, 
            // then requirements MS-OXWSSYNC_R84 and MS-OXWSSYNC_R10010 can be captured
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R84");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R84
            Site.CaptureRequirementIfIsTrue(
                isFolderCreated,
                84,
                @"[In t:SyncFolderHierarchyChangesType Complex Type] [The element Create] specifies a folder that has been created on the server and has to be created on the client.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R10010");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R10010
            Site.CaptureRequirementIfIsTrue(
                isFolderCreated,
                10010,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] [The element Folder] specifies a regular folder to create in the client message store.");
            #endregion

            #region Step 4. Client invokes UpdateFolder operation to change the second level folder's name.
            // Generate a new folder name
            string newFolderName = Common.GenerateResourceName(this.Site, inboxFolder + "NewFolderName");

            // Update the name of the second level sub folder.
            bool updatedFolder = this.SYNCSUTControlAdapter.FindAndUpdateFolderName(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                inboxFolder.ToString(),
                secondLevelFolderName,
                newFolderName);
            Site.Assert.IsTrue(
                updatedFolder,
                string.Format("The folder name '{0}' should be updated to '{1}'.", secondLevelFolderName, newFolderName));
            #endregion

            #region  Step 5. Client invokes SyncFolderHierarchy operation with previous syncState to sync the operation result in Step 4 and verify related requirements.
            // Assert the SyncState is not null
            Site.Assert.IsNotNull(responseMessage.SyncState, "The synchronization should not be null.");
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one folder updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there is one folder updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be folder information returned in SyncFolderHierarchy response since there is one folder updated on server.");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one FolderType folder was updated in previous step, so the count of ItemsElementName array in SyncFolderHierarchy responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one FolderType folder was updated in previous step, so the count of Items array in SyncFolderHierarchy responseMessage.Changes should be 1.");

            // If the type of item in SyncFolderHierarchy response is FolderType, then requirement MS-OXWSSYNC_R99 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R99");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R99
            Site.CaptureRequirementIfIsInstanceOfType(
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item,
                typeof(FolderType),
                99,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] The type of Folder is t:FolderType ([MS-OXWSFOLD] section 2.2.4.12).");

            bool isFolderNameUpdated = changes.ItemsElementName[0] == ItemsChoiceType.Update &&
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.DisplayName == newFolderName;

            // If the ItemsElementName of Changes is Update and the item's display name is a new value, it indicates the folder has been updated on server and synced on client, 
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R85. Expected value: ItemsElementName: {0}, folder display name: {1}; actual value: ItemsElementName: {2}, folder display name: {3}",
                ItemsChoiceType.Update,
                newFolderName,
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.DisplayName);

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R85
            Site.CaptureRequirementIfIsTrue(
                isFolderNameUpdated,
                85,
                @"[In t:SyncFolderHierarchyChangesType Complex Type] [The element Update] specifies a folder that has been changed on the server and has to be changed on the client.");

            bool isFolderUpdated = changes.ItemsElementName[0] == ItemsChoiceType.Update &&
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType() == typeof(FolderType);

            // If the ItemsElementName of Changes is Update and the type of Item is FolderType, it indicates a regular folder has been updated on server and synced on client, 
            // then requirement MS-OXWSSYNC_R10020 can be captured.
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R10020. Expected value: ItemsElementName: {0}, folder type: {1}; actual value: ItemsElementName: {2}, folder type: {3}",
                ItemsChoiceType.Update,
                typeof(FolderType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType());

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R10020
            Site.CaptureRequirementIfIsTrue(
                isFolderUpdated,
                10020,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] [The element Folder] specifies a regular folder to update in the client message store.");
            #endregion

            #region Step 6. Client invokes DeleteFolder operation to delete the second level folder that created in step 2.
            // Delete the second level sub folder.
            bool isDeleted = this.SYNCSUTControlAdapter.FindAndDeleteSubFolder(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                inboxFolder.ToString(),
                newFolderName);
            Site.Assert.IsTrue(isDeleted, string.Format("The folder named '{0}' should be deleted from '{1}' successfully.", newFolderName, inboxFolder));
            #endregion

            #region Step 7. Client invokes SyncFolderHierarchy operation with previous syncState to sync the operation result of Step 6 and verify related requirements.
            // Assert the SyncState is not null
            Site.Assert.IsNotNull(responseMessage.SyncState, "The synchronization should not be null.");
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one folder deleted on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there is one folder deleted on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be folder information returned in SyncFolderHierarchy response since there is one folder deleted on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one FolderType folder was deleted in previous step, so the count of Items array in SyncFolderHierarchy responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one FolderType folder was deleted in previous step, so the count of ItemsElementName array in SyncFolderHierarchy responseMessage.Changes should be 1.");

            bool isFolderDeleted = (changes.ItemsElementName[0] == ItemsChoiceType.Delete) && (changes.Items[0].GetType() == typeof(SyncFolderHierarchyDeleteType));
            if (Common.IsRequirementEnabled(37811002, this.Site))
            {
                // If the ItemsElementName is Delete and the item type in changes is SyncFolderHierarchyDeleteType, it indicates a folder has been deleted on server and synced on client. 
                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXWSSYNC_R87. Expected value: ItemsElementName: {0}, change items type: {1}; actual value: ItemsElementName: {2}, change items type: {3}",
                    ItemsChoiceType.Delete,
                    typeof(SyncFolderHierarchyDeleteType),
                    changes.ItemsElementName[0],
                    changes.Items[0].GetType());

                // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R37811002
                Site.CaptureRequirementIfIsTrue(
                    isFolderDeleted,
                    37811002,
                    @"[In Appendix C: Product Behavior] Implementation does include Delete element. (Exchange 2010 and above follow this behavior.)");
            #endregion
            }
        }

        /// <summary>
        /// Client calls SyncFolderHierarchy operation to sync calendar folder.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S01_TC02_SyncFolderHierarchy_CalendarFolder()
        {
            #region Step 1. Client invokes SyncFolderHierarchy operation to get initial syncState of calendar folder.
            DistinguishedFolderIdNameType calendarFolder = DistinguishedFolderIdNameType.calendar;

            // Set DefaultShapeNamesType to IdOnly and include SyncFolderId and SyncState element in the request
            SyncFolderHierarchyType request = TestSuiteHelper.CreateSyncFolderHierarchyRequest(calendarFolder, DefaultShapeNamesType.IdOnly, true, true);
            SyncFolderHierarchyResponseType response = this.SYNCAdapter.SyncFolderHierarchy(request);
            SyncFolderHierarchyResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);
            #endregion

            #region Step 2. Client invokes CreateFolder operation to create two level folders in calendar folder.
            // Generate the created folder name
            string firstLevelFolderName = Common.GenerateResourceName(this.Site, calendarFolder + "FirstLevelFolder");
            string secondLevelFolderName = Common.GenerateResourceName(this.Site, calendarFolder + "SecondLevelFolder");

            // Create two level sub folders in calendar folder.
            bool isSubFolderCreated = this.FOLDSUTControlAdapter.CreateSubFolders(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                calendarFolder.ToString(),
                firstLevelFolderName,
                secondLevelFolderName);
            Site.Assert.IsTrue(isSubFolderCreated, string.Format("The new sub folders in '{0}' should be created successfully.", calendarFolder));
            #endregion

            #region Step 3. Client invokes SyncFolderHierarchy operation to sync the operation result in Step 2 and verify related requirements.
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There are two folders created on server, so the changes between server and client should not be null");
            SyncFolderHierarchyChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there are two folders created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be folders information returned in SyncFolderHierarchy response since there are two folders created on server.");

            bool isCalendarFolderCreated = false;
            for (int i = 0; i < changes.Items.Length; i++)
            {
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Expected value: ItemsElementName: {0}, folder type: {1}; actual value: ItemsElementName: {2}, folder type: {3}",
                    ItemsChoiceType.Create,
                    typeof(CalendarFolderType),
                    changes.ItemsElementName[i],
                    (changes.Items[i] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType());

                if (changes.ItemsElementName[i] == ItemsChoiceType.Create &&
                        (changes.Items[i] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType() == typeof(CalendarFolderType))
                {
                    isCalendarFolderCreated = true;
                }
                else
                {
                    isCalendarFolderCreated = false;
                    break;
                }
            }

            // If the ItemsElementName of Changes is Create and the type of Item is CalendarFolderType, 
            // it indicates a calendar folder has been created on server and synced on client, then requirement MS-OXWSSYNC_R1021 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R1021");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1021
            Site.CaptureRequirementIfIsTrue(
                isCalendarFolderCreated,
                1021,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] [The element CalendarFolder] specifies a calendar folder to create in the client message store.");
            #endregion

            #region Step 4. Client invokes UpdateFolder operation to change the second level folder's name.
            // Generate a new folder name
            string newFolderName = Common.GenerateResourceName(this.Site, calendarFolder + "NewFolderName");

            // Update the name of the second level sub folder.
            bool updatedFolder = this.SYNCSUTControlAdapter.FindAndUpdateFolderName(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                calendarFolder.ToString(),
                secondLevelFolderName,
                newFolderName);
            Site.Assert.IsTrue(updatedFolder, string.Format("The folder name '{0}' should be updated to '{1}'.", secondLevelFolderName, newFolderName));
            #endregion

            #region Step 5. Client invokes SyncFolderHierarchy operation with previous syncState to sync the operation result in Step 4 and verify related requirements.
            // Assert the SyncState is not null
            Site.Assert.IsNotNull(responseMessage.SyncState, "The synchronization should not be null.");
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one folder updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there is one folder updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be folders information returned in SyncFolderHierarchy response since there is one folder updated on server.");

            bool isIdOnly = Common.IsIdOnly((XmlElement)this.SYNCAdapter.LastRawResponseXml, "t:CalendarFolder", "t:FolderId");

            // If there is only one FolderId element in the item of changes, then requirement MS-OXWSSYNC_R2573 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R2573");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R2573
            Site.CaptureRequirementIfIsTrue(
                isIdOnly,
                2573,
                @"[In m:SyncFolderHierarchyType Complex Type] FolderShape element BaseShape, value=IdOnly, specifies only the item or folder ID.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one CalendarFolderType folder was updated in previous step, so the count of Items array in SyncFolderHierarchy responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one CalendarFolderType folder was updated in previous step, so the count of ItemsElementName array in SyncFolderHierarchy responseMessage.Changes should be 1.");

            // If the type of item in SyncFolderHierarchy response is CalendarFolderType, then requirement MS-OXWSSYNC_R101 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R101");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R101
            Site.CaptureRequirementIfIsInstanceOfType(
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item,
                typeof(CalendarFolderType),
                101,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] The type of CalendarFolder element is t:CalendarFolderType ([MS-OXWSMTGS] section 2.2.4.8).");

            bool isCalendarFolderUpdated = (changes.ItemsElementName[0] == ItemsChoiceType.Update) &&
                ((changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType() == typeof(CalendarFolderType));

            // If the ItemsElementName of Changes is Update and the type of Item is CalendarFolderType, it indicates a calendar folder has been updated on server and synced on client, 
            // then requirement MS-OXWSSYNC_R1022 can be captured.
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1022. Expected value: ItemsElementName: {0}, folder type: {1}; actual value: ItemsElementName: {2}, folder type: {3}",
                ItemsChoiceType.Update,
                typeof(CalendarFolderType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType());

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1022
            Site.CaptureRequirementIfIsTrue(
                isCalendarFolderUpdated,
                1022,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] [The element CalendarFolder] specifies a calendar folder to update in the client message store.");
            #endregion

            #region Step 6. Client invokes DeleteFolder operation to delete the second level folder that created in step 2.
            // Delete the second level sub folder
            bool isDeleted = this.SYNCSUTControlAdapter.FindAndDeleteSubFolder(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                calendarFolder.ToString(),
                newFolderName);
            Site.Assert.IsTrue(isDeleted, string.Format("The folder named '{0}' should be deleted from '{1}' successfully.", newFolderName, calendarFolder));
            #endregion

            #region Step 7. Client invokes SyncFolderHierarchy operation with previous syncState to sync the operation result of Step 6 and verify related requirements.
            // Assert the SyncState is not null
            Site.Assert.IsNotNull(responseMessage.SyncState, "The synchronization should not be null.");
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one folder deleted on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert ItemsElementName is not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there is one folder deleted on server.");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one CalendarFolderType folder was deleted in previous step, so the count of ItemsElementName array in SyncFolderHierarchy responseMessage.Changes should be 1.");
            bool isIncrementalSync = changes.ItemsElementName[0] == ItemsChoiceType.Delete && responseMessage.SyncState != null;

            // If the ItemsElementName is Delete and the SyncState element is not null, then requirement MS-OXWSSYNC_R503 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R503. Expected value: ItemsElementName:{0}; actual value: ItemsElementName: {1}", ItemsChoiceType.Delete, changes.ItemsElementName[0]);

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R503
            Site.CaptureRequirementIfIsTrue(
                isIncrementalSync,
                503,
                @"[In Abstract Data Model] If the optional SyncState element of the SyncFolderHierarchyType complex type (section 3.1.4.1.3.6) is included in a SyncFolderHierarchy operation (section 3.1.4.1) request, the server MUST return incremental synchronization information from the last synchronization request.");
            #endregion
        }

        /// <summary>
        /// Client calls SyncFolderHierarchy operation to sync contacts folder.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S01_TC03_SyncFolderHierarchy_ContactsFolder()
        {
            #region Step 1. Client invokes SyncFolderHierarchy operation to get initial syncState of contacts folder.
            DistinguishedFolderIdNameType contactFolder = DistinguishedFolderIdNameType.contacts;

            // Set DefaultShapeNamesType to IdOnly and include SyncFolderId and SyncState element in the request
            SyncFolderHierarchyType request = TestSuiteHelper.CreateSyncFolderHierarchyRequest(contactFolder, DefaultShapeNamesType.Default, true, true);
            SyncFolderHierarchyResponseType response = this.SYNCAdapter.SyncFolderHierarchy(request);
            SyncFolderHierarchyResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);
            #endregion

            #region Step 2. Client invokes CreateFolder operation to create two level folders in contacts folder.
            // Generate the created folder name
            string firstLevelFolderName = Common.GenerateResourceName(this.Site, contactFolder + "FirstLevelFolder");
            string secondLevelFolderName = Common.GenerateResourceName(this.Site, contactFolder + "SecondLevelFolder");

            // Create two level sub folders in contact folder.
            bool isSubFolderCreated = this.FOLDSUTControlAdapter.CreateSubFolders(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                contactFolder.ToString(),
                firstLevelFolderName,
                secondLevelFolderName);
            Site.Assert.IsTrue(isSubFolderCreated, string.Format("The new sub folders in '{0}' should be created successfully.", contactFolder));
            #endregion

            #region Step 3. Client invokes SyncFolderHierarchy operation to sync the operation result in Step 2 and verify related requirements.
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There are two folders created on server, so the changes between server and client should not be null");
            SyncFolderHierarchyChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there are two folders created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be folders information returned in SyncFolderHierarchy response since there are two folders created on server.");

            bool isContactsFolderCreated = false;
            for (int i = 0; i < changes.Items.Length; i++)
            {
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Expected value: ItemsElementName: {0}, folder type: {1}; actual value: ItemsElementName: {2}, folder type: {3}",
                    ItemsChoiceType.Create,
                    typeof(ContactsFolderType),
                    changes.ItemsElementName[i],
                    (changes.Items[i] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType());

                if (changes.ItemsElementName[i] == ItemsChoiceType.Create &&
                    (changes.Items[i] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType() == typeof(ContactsFolderType))
                {
                    isContactsFolderCreated = true;
                }
                else
                {
                    isContactsFolderCreated = false;
                    break;
                }
            }

            // If the ItemsElementName of Changes is Create and the type of Item is ContactsFolderType, it indicates a contacts folder has been created on server and synced on client, 
            // then requirement MS-OXWSSYNC_R1041 can be captured
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R1041");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1041
            Site.CaptureRequirementIfIsTrue(
                isContactsFolderCreated,
                1041,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] [The element ContactsFolder] specifies a contacts folder to create in the client message store.");
            #endregion

            #region Step 4. Client invokes UpdateFolder operation to change the second level folder's name.
            // Generate a new folder name
            string newFolderName = Common.GenerateResourceName(this.Site, contactFolder + "NewFolderName");

            // Update the name of the second level sub folder.
            bool updatedFolder = this.SYNCSUTControlAdapter.FindAndUpdateFolderName(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                contactFolder.ToString(),
                secondLevelFolderName,
                newFolderName);
            Site.Assert.IsTrue(updatedFolder, string.Format("The folder name '{0}' should be updated to '{1}'.", secondLevelFolderName, newFolderName));
            #endregion

            #region Step 5. Client invokes SyncFolderHierarchy operation with previous SyncState to sync the operation result in Step 4 and verify related requirements.
            // Assert the SyncState is not null
            Site.Assert.IsNotNull(responseMessage.SyncState, "The synchronization should not be null.");
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one folder updated on server, so the changes between server and client should not be null");
            SyncFolderHierarchyChangesType changesAfterUpdateFolder = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changesAfterUpdateFolder.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there is one folder updated on server.");
            Site.Assert.IsNotNull(changesAfterUpdateFolder.Items, "There should be folders information returned in SyncFolderHierarchy response since there is one folder updated on server.");

            Site.Assert.AreEqual<int>(1, changesAfterUpdateFolder.Items.Length, "Just one ContactsFolderType folder was updated in previous step, so the count of Items array in SyncFolderHierarchy responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changesAfterUpdateFolder.ItemsElementName.Length, "Just one ContactsFolderType folder was updated in previous step, so the count of ItemsElementName array in SyncFolderHierarchy responseMessage.Changes should be 1.");

            // If the type of item in SyncFolderHierarchy response is ContactsFolderType, then requirement MS-OXWSSYNC_R103 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R103");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R103
            Site.CaptureRequirementIfIsInstanceOfType(
                (changesAfterUpdateFolder.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item,
                typeof(ContactsFolderType),
                103,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] The type of ContactsFolder is t:ContactsFolderType ([MS-OXWSCONT] section 3.1.4.1.1.6).");

            bool isContactsFolderUpdated = (changesAfterUpdateFolder.ItemsElementName[0] == ItemsChoiceType.Update) &&
                ((changesAfterUpdateFolder.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType() == typeof(ContactsFolderType));

            // If the ItemsElementName of Changes is Update and the type of Item is ContactsFolderType, 
            // it indicates a contacts folder has been updated on server and synced on client, then requirement MS-OXWSSYNC_R1042 can be captured.
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1042. Expected value: ItemsElementName: {0}, folder type: {1}; actual value: ItemsElementName: {2}, folder type: {3}",
                ItemsChoiceType.Update,
                typeof(ContactsFolderType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType());

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1042
            Site.CaptureRequirementIfIsTrue(
                isContactsFolderUpdated,
                1042,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] [The element ContactsFolder] specifies a contacts folder to update in the client message store.");

            bool isLastFolderIncluded = responseMessage.IncludesLastFolderInRange &&
                ((changesAfterUpdateFolder.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType() == typeof(ContactsFolderType));

            // Since the last folder that updated is a contacts folder, if the IncludesLastFolderInRange element in SyncFolderHierarchy response is TRUE 
            // and the items in Changes contains ContactsFolderType item, then requirement MS-OXWSSYNC_R4601 can be captured.
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R4601. Expected value: IncludesLastFolderInRange: 'true', folder type: {0}; actual value: IncludesLastFolderInRange: {1}, folder type: {2}",
                typeof(ContactsFolderType),
                responseMessage.IncludesLastFolderInRange,
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType());

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R4601
            Site.CaptureRequirementIfIsTrue(
                isLastFolderIncluded,
                4601,
                @"[In m:SyncFolderHierarchyResponseMessageType Complex Type] [The element IncludesLastFolderInRange] If this element is included in the response, the value is always ""true"".");
            #endregion

            #region Step 6. Client invokes FindAndDeleteSubFolder to delete the second level folder that created in step 2.
            // Delete the second level sub folder.
            bool isDeleted = this.SYNCSUTControlAdapter.FindAndDeleteSubFolder(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                contactFolder.ToString(),
                newFolderName);

            Site.Assert.IsTrue(isDeleted, string.Format("The folder named '{0}' should be deleted from '{1}' successfully.", newFolderName, contactFolder));
            #endregion

            #region Step 7. Client invokes SyncFolderHierarchy operation with previous SyncState to sync the operation result of Step 6 and verify related requirements.
            // Assert the SyncState is not null
            Site.Assert.IsNotNull(responseMessage.SyncState, "The synchronization should not be null.");
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one folder deleted on server, so the changes between server and client should not be null");
            SyncFolderHierarchyChangesType changesAfterDeleteFolder = responseMessage.Changes;

            // Assert ItemsElementName is not null
            Site.Assert.IsNotNull(changesAfterDeleteFolder.Items, "There should be folder information returned in SyncFolderHierarchy response since there is one folder deleted on server.");
            Site.Assert.AreEqual<int>(1, changesAfterDeleteFolder.Items.Length, "Just one ContactsFolderType folder was deleted in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // If the folder id in SyncFolderHierarchy response after DeleteFolder operation is same with it in response after UpdateFolder operation (since there is no other 
            // operation between these two operations, the id should be same), then requirement MS-OXWSSYNC_R117 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R117");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R117
            Site.CaptureRequirementIfAreEqual<string>(
                (changesAfterUpdateFolder.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.FolderId.Id,
                (changesAfterDeleteFolder.Items[0] as SyncFolderHierarchyDeleteType).FolderId.Id,
                117,
                @"[In t:SyncFolderHierarchyDeleteType Complex Type] [The element FolderId] specifies the identifier of the folder to delete from the client message store.");
            #endregion
        }

        /// <summary>
        /// Client calls SyncFolderHierarchy operation to sync TaskFolder.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S01_TC04_SyncFolderHierarchy_TasksFolder()
        {
            #region Step 1. Client invokes SyncFolderHierarchy operation to get initial syncState of tasks folder.
            DistinguishedFolderIdNameType taskFolder = DistinguishedFolderIdNameType.tasks;

            // Set DefaultShapeNamesType to AllProperties and include SyncFolderId and SyncState element in the request
            SyncFolderHierarchyType request = TestSuiteHelper.CreateSyncFolderHierarchyRequest(taskFolder, DefaultShapeNamesType.AllProperties, true, true);
            SyncFolderHierarchyResponseType response = this.SYNCAdapter.SyncFolderHierarchy(request);
            SyncFolderHierarchyResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);
            #endregion

            #region Step 2. Client invokes CreateFolder operation to create two level folders in tasks folder.
            // Generate the created folder name
            string firstLevelFolderName = Common.GenerateResourceName(this.Site, taskFolder + "FirstLevelFolder");
            string secondLevelFolderName = Common.GenerateResourceName(this.Site, taskFolder + "SecondLevelFolder");

            // Create two level sub folders in task folder.
            bool isSubFolderCreated = this.FOLDSUTControlAdapter.CreateSubFolders(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                taskFolder.ToString(),
                firstLevelFolderName,
                secondLevelFolderName);
            Site.Assert.IsTrue(isSubFolderCreated, string.Format("The new sub folders in '{0}' should be created successfully.", taskFolder));
            #endregion

            #region Step 3. Client invokes SyncFolderHierarchy operation to sync the operation result in Step 2 and verify related requirements.
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R49");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R49
            Site.CaptureRequirementIfIsNotNull(
                responseMessage.Changes,
                49,
                @"[In m:SyncFolderHierarchyResponseMessageType Complex Type] [The element Changes] specifies the differences between the folders on the client and the folders on the server.");

            SyncFolderHierarchyChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there are two folders created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be folders information returned in SyncFolderHierarchy response since there are two folders created on server.");

            bool isTasksFolderCreated = false;
            for (int i = 0; i < changes.Items.Length; i++)
            {
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Expected value: ItemsElementName: {0}, folder type: {1}; actual value: ItemsElementName: {2}, folder type: {3}",
                    ItemsChoiceType.Create,
                    typeof(TasksFolderType),
                    changes.ItemsElementName[i],
                    (changes.Items[i] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType());

                if (changes.ItemsElementName[i] == ItemsChoiceType.Create &&
                        (changes.Items[i] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType() == typeof(TasksFolderType))
                {
                    isTasksFolderCreated = true;
                }
                else
                {
                    isTasksFolderCreated = false;
                    break;
                }
            }

            // If the ItemsElementName of Changes is Create and the type of Item is TasksFolderType, 
            // it indicates a tasks folder has been created on server and synced on client, then requirement MS-OXWSSYNC_R1081 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R1081");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1081
            Site.CaptureRequirementIfIsTrue(
                isTasksFolderCreated,
                1081,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] [The element TasksFolder] specifies a tasks folder to create in the client message store.");
            #endregion

            #region Step 4. Client invokes UpdateFolder to change the second level folder's name.
            // Generate a new folder name
            string newFolderName = Common.GenerateResourceName(this.Site, taskFolder + "NewFolderName");

            // Update the name of the second level sub folder.
            bool updatedFolder = this.SYNCSUTControlAdapter.FindAndUpdateFolderName(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                taskFolder.ToString(),
                secondLevelFolderName,
                newFolderName);
            Site.Assert.IsTrue(updatedFolder, string.Format("The folder name '{0}' should be updated to '{1}'.", secondLevelFolderName, newFolderName));
            #endregion

            #region Step 5. Client invokes SyncFolderHierarchy operation with previous SyncState to sync the operation result in Step 4 and verify related requirements.
            // Assert the SyncState is not null
            Site.Assert.IsNotNull(responseMessage.SyncState, "The synchronization should not be null.");
            request.SyncState = responseMessage.SyncState;
            request.FolderShape.AdditionalProperties = new BasePathToElementType[] 
            { 
                new PathToUnindexedFieldType() 
                { 
                    FieldURI = UnindexedFieldURIType.folderDisplayName 
                } 
            };
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one folder updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there is one folder updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be folders information returned in SyncFolderHierarchy response since there is one folder updated on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one TasksFolderType folder was updated in previous step, so the count of Items array in SyncFolderHierarchy responseMessage.Changes should be 1.");

            // If the type of item in SyncFolderHierarchy response is TasksFolderType, then requirement MS-OXWSSYNC_R107 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R107");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R107
            Site.CaptureRequirementIfIsInstanceOfType(
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item,
                typeof(TasksFolderType),
                107,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] The type of TasksFolder is  t:TasksFolderType ([MS-OXWSTASK] section 2.2.4.5).");

            // If the AdditionalProperties element is included in SyncFolderHierarchy request and the FieldURI is point to folder display name, 
            // the additional property DisplayName should be returned in response, then requirement MS-OXWSSYNC_R2574 and MS-OXWSSYNC_R2575 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R2574");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R2574
            Site.CaptureRequirementIfIsNotNull(
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.DisplayName,
                2574,
                @"[In m:SyncFolderHierarchyType Complex Type] FolderShape element AdditionalProperties, specifies the identity of additional properties to be returned in a response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R2575");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R2575
            Site.CaptureRequirementIfIsNotNull(
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.DisplayName,
                2575,
                @"[In m:SyncFolderHierarchyType Complex Type]FolderShape element AdditionalProperties, element t:Path, Specifies a property to be returned in a response.");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one TasksFolderType folder was updated in previous step, so the count of ItemsElementName array in SyncFolderHierarchy responseMessage.Changes should be 1.");
            bool isContactsFolderUpdated = (changes.ItemsElementName[0] == ItemsChoiceType.Update) &&
                ((changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType() == typeof(TasksFolderType));

            // If the ItemsElementName of Changes is Update and the type of Item is TasksFolderType, it indicates a tasks folder has been updated on server and synced on client, then requirement MS-OXWSSYNC_R1082 can be captured.
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1082. Expected value: ItemsElementName: {0}, folder type: {1}; actual value: ItemsElementName: {2}, folder type: {3}",
                ItemsChoiceType.Update,
                typeof(TasksFolderType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType());

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1082
            Site.CaptureRequirementIfIsTrue(
                isContactsFolderUpdated,
                1082,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] [The element TasksFolder] specifies a tasks folder to update in the client message store.");
            #endregion

            #region Step 6. Client invokes DeleteFolder to delete the second level folder that created in step 2.
            // Delete the second level sub folder.
            bool isDeleted = this.SYNCSUTControlAdapter.FindAndDeleteSubFolder(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                taskFolder.ToString(),
                newFolderName);
            Site.Assert.IsTrue(isDeleted, string.Format("The folder named '{0}' should be deleted from '{1}' successfully.", newFolderName, taskFolder));
            #endregion

            #region Step 7. Client invokes SyncFolderHierarchy operation with previous syncState to sync the operation result of Step 6 and verify related requirements.

            // Assert the SyncState is not null
            Site.Assert.IsNotNull(responseMessage.SyncState, "The synchronization should not be null.");
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // If the SyncState element in SyncFolderHierarchy response is not null, it indicates the synchronization state is returned in response.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R44");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R44
            Site.CaptureRequirementIfIsNotNull(
                responseMessage.SyncState,
                44,
                @"[In m:SyncFolderHierarchyResponseMessageType Complex Type] [The element SyncState] specifies a form of the synchronization data, which is encoded with base64 encoding, that is used to identify the synchronization state.");

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one folder deleted on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there is one folder deleted on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be folders information returned in SyncFolderHierarchy response since there is one folder deleted on server.");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one TasksFolderType folder was deleted in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.ItemsElementName[0] == ItemsChoiceType.Delete,
                string.Format("The responseMessage.Changes.ItemsElementName should be 'Delete', the actual value is '{0}'", changes.ItemsElementName[0]));

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one TasksFolderType folder was deleted in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.Items[0].GetType() == typeof(SyncFolderHierarchyDeleteType),
                string.Format("The responseMessage.Changes.Items should be an instance of '{0}'.", typeof(SyncFolderHierarchyDeleteType)));
            #endregion
        }

        /// <summary>
        /// Client calls SyncFolderHierarchy operation to sync SearchFolder.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S01_TC05_SyncFolderHierarchy_SearchFolder()
        {
            #region Step 1. Client invokes SyncFolderHierarchy operation to get initial syncState of search folder and verify related requirements.
            DistinguishedFolderIdNameType searchFolder = DistinguishedFolderIdNameType.searchfolders;

            // Call SyncFolderHierarchy operation with invalid SyncState to verify the error code: ErrorInvalidSyncStateData
            SyncFolderHierarchyType requestWithInvalidSyncState = TestSuiteHelper.CreateSyncFolderHierarchyRequest(searchFolder, DefaultShapeNamesType.AllProperties, false, false);

            // The SyncState element data, encoded with base64 encoding, is set to an invalid value
            requestWithInvalidSyncState.SyncState = TestSuiteBase.InvalidSyncState;
            SyncFolderHierarchyResponseType responseWithInvalidSyncState = this.SYNCAdapter.SyncFolderHierarchy(requestWithInvalidSyncState);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R5188");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R5188
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidSyncStateData,
                responseWithInvalidSyncState.ResponseMessages.Items[0].ResponseCode,
                "MS-OXWSCDATA",
                5188,
                @"[In m:ResponseCodeType Simple Type] [ErrorInvalidSyncStateData: ] This is returned by the SyncFolderHierarchy method if the SyncState property data is invalid.");

            // Set DefaultShapeNamesType to AllProperties and don't include SyncFolderId element in the request
            SyncFolderHierarchyType request = TestSuiteHelper.CreateSyncFolderHierarchyRequest(searchFolder, DefaultShapeNamesType.AllProperties, false, false);
            SyncFolderHierarchyResponseType response = this.SYNCAdapter.SyncFolderHierarchy(request);
            SyncFolderHierarchyResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R2602");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R2602
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                responseMessage.ResponseCode,
                2602,
                @"[In m:SyncFolderHierarchyType Complex Type] This element [SyncFolderId] is not present, server responses NO_ERROR.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R2632");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R2632
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                responseMessage.ResponseCode,
                2632,
                @"[In m:SyncFolderHierarchyType Complex Type] This element [SyncState] not present, server responses NO_ERROR.");

            // Set DefaultShapeNamesType to AllProperties and include SyncFolderId element in the request
            request = TestSuiteHelper.CreateSyncFolderHierarchyRequest(searchFolder, DefaultShapeNamesType.AllProperties, true, true);
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R2601");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R2601
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                responseMessage.ResponseCode,
                2601,
                @"[In m:SyncFolderHierarchyType Complex Type] This element [SyncFolderId] is present, server responses NO_ERROR.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R2631");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R2631
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                responseMessage.ResponseCode,
                2631,
                @"[In m:SyncFolderHierarchyType Complex Type] This element [SyncState] is present, server responses NO_ERROR.");
            #endregion

            #region Step 2. Client invokes CreateFolder operation to create a search folder.
            // Generate the created folder name
            string firstLevelFolderName = Common.GenerateResourceName(this.Site, searchFolder + "FirstLevelFolder");

            // Create a search folder.
            bool isSubFolderCreated = this.FOLDSUTControlAdapter.CreateSearchFolder(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                firstLevelFolderName,
                SearchText);
            Site.Assert.IsTrue(isSubFolderCreated, string.Format("The new search folder named '{0}' should be created successfully.", firstLevelFolderName));
            #endregion

            #region Step 3. Client invokes SyncFolderHierarchy operation to sync the operation result in Step 2 and verify related requirements.
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one folder created on server, so the changes between server and client should not be null");
            SyncFolderHierarchyChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there is one folder created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be folders information returned in SyncFolderHierarchy response since there is one folder created on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one SearchFolderType folder was created in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // If the type of item in SyncFolderHierarchy response is SearchFolderType, then requirement MS-OXWSSYNC_R105 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R105");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R105
            Site.CaptureRequirementIfIsInstanceOfType(
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item,
                typeof(SearchFolderType),
                105,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] The type of SearchFolder is t:SearchFolderType ([MS-OXWSSRCH] section 2.2.4.32).");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one SearchFolderType folder was created in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            bool isSearchFolderCreated = changes.ItemsElementName[0] == ItemsChoiceType.Create &&
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType() == typeof(SearchFolderType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1061. Expected value: ItemsElementName: {0}, folder type: {1}; actual value: ItemsElementName: {2}, folder type: {3}",
                ItemsChoiceType.Create,
                typeof(SearchFolderType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Create and the type of Item is SearchFolderType, it indicates a search folder has been created on server and synced on client, 
            // then requirement MS-OXWSSYNC_R1061 can be captured.
            Site.CaptureRequirementIfIsTrue(
                isSearchFolderCreated,
                1061,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] [The element SearchFolder] specifies a search folder to create in the client message store.");
            #endregion

            #region Step 4. Client invokes UpdateFolder to change the folder's name.
            // Generate a new folder name
            string newFolderName = Common.GenerateResourceName(this.Site, searchFolder + "NewFolderName");

            // Update the name of the new created folder.
            bool updatedFolder = this.SYNCSUTControlAdapter.FindAndUpdateFolderName(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                searchFolder.ToString(),
                firstLevelFolderName,
                newFolderName);
            Site.Assert.IsTrue(updatedFolder, string.Format("The folder name '{0}' should be updated to '{1}'.", firstLevelFolderName, newFolderName));
            #endregion

            #region Step 5. Client invokes SyncFolderHierarchy operation with previous SyncState to sync the operation result in Step 4 and verify related requirements.
            // Assert the SyncState is not null
            Site.Assert.IsNotNull(responseMessage.SyncState, "The synchronization should not be null.");
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one folder updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there is one folder updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be folders information returned in SyncFolderHierarchy response since there is one folder updated on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one SearchFolderType folder was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one SearchFolderType folder was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            bool isSearchFolderUpdated = (changes.ItemsElementName[0] == ItemsChoiceType.Update) &&
                ((changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType() == typeof(SearchFolderType));

            // If the ItemsElementName of Changes is Update and the type of Item is SearchFolderType, it indicates a search folder has been updated on server and synced on client, 
            // then requirement MS-OXWSSYNC_R1062 can be captured.
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1062. Expected value: ItemsElementName: {0}, folder type: {1}; actual value: ItemsElementName: {2}, folder type: {3}",
                ItemsChoiceType.Update,
                typeof(SearchFolderType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderHierarchyCreateOrUpdateType).Item.GetType());

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1062
            Site.CaptureRequirementIfIsTrue(
                isSearchFolderUpdated,
                1062,
                @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] [The element SearchFolder] specifies a search folder to update in the client message store.");

            // In SyncFolderHierarchy request, the SyncFolderId is set to "searchfolders", if the item in Changes.Items[0] of SyncFolderHierarchy response is SearchFolderType,
            // then requirement MS-OXWSSYNC_R259 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R259");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R259
            Site.CaptureRequirementIfIsTrue(
                isSearchFolderUpdated,
                259,
                @"[In m:SyncFolderHierarchyType Complex Type] [The element SyncFolderId] specifies the target folder for the operation [SyncFolderHierarchy].");

            // Call SyncFolderHierarchy again without SyncState to verify that all synchronization is returned.
            SyncFolderHierarchyType requestWithoutSyncState = TestSuiteHelper.CreateSyncFolderHierarchyRequest(searchFolder, DefaultShapeNamesType.AllProperties, true, false);
            response = this.SYNCAdapter.SyncFolderHierarchy(requestWithoutSyncState);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one folder deleted on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderHierarchy response since there is one folder deleted on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be folders information returned in SyncFolderHierarchy response since there is one folder deleted on server.");

            Site.Assert.AreEqual<ItemsChoiceType>(
            ItemsChoiceType.Create,
            changes.ItemsElementName[0],
            "After updating the folder, if the SyncState element is not specified when calling SyncFolderHierarchy, the changes between folders on the client and the folders on the server should be 'Create'");

            bool renamed = false;
            for (int index = 0; index < changes.Items.Length; index++)
            {
                if (string.Compare((changes.Items[index] as SyncFolderHierarchyCreateOrUpdateType).Item.DisplayName, newFolderName, System.StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    renamed = true;
                }
            }

            Site.Assert.IsTrue(
                renamed,
                 "After updating the folder, the display name of the folder should be the expected one.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R262");

            // If the value of ItemsElementName of Changes is Create and the display name of the folder is updated, it indicates the folder in its current state is returned as if it has never been synchronized,
            // then requirement MS-OXWSSYNC_R262 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R262
            Site.CaptureRequirement(
                262,
                @"[In m:SyncFolderHierarchyType Complex Type] If this element [SyncState] is not specified, all items in their current state are returned as if the items have never been synchronized. ");
            #endregion

            #region Step 6. Client invokes DeleteFolder to delete the folder that created in step 2.
            // Delete the created search folder.
            bool isDeleted = this.SYNCSUTControlAdapter.FindAndDeleteSubFolder(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                searchFolder.ToString(),
                newFolderName);
            Site.Assert.IsTrue(isDeleted, string.Format("The folder named '{0}' should be deleted from '{1}' successfully.", newFolderName, searchFolder));
            #endregion

            #region Step 7. Client invokes SyncFolderHierarchy operation with previous syncState to sync the operation result of Step 6 and verify related requirements.
            // Assert the SyncState is not null
            Site.Assert.IsNotNull(responseMessage.SyncState, "The synchronization should not be null.");
            request.SyncState = responseMessage.SyncState;
            response = this.SYNCAdapter.SyncFolderHierarchy(request);
            TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);
            #endregion
        }
        #endregion
    }
}