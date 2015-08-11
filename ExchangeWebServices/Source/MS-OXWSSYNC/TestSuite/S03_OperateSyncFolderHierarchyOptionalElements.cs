namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test SyncFolderHierarchy operation with/without all optional elements in request on the following folders: inbox folder, calendar folder, contacts folder, tasks folder and search folder.
    /// </summary>
    [TestClass]
    public class S03_OperateSyncFolderHierarchyOptionalElements : TestSuiteBase
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
        /// Client calls SyncFolderHierarchy operation without optional elements to get the synchronization information of 
        /// 5 type of folders (inbox, calendar, tasks, contacts, search).
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S03_TC01_SyncFolderHierarchy_WithoutOptionalElements()
        {
            #region Step 1. Add inbox folder and search folder into list.
            // Add inbox folder and search folder into list
            this.FolderIdNameType.Add(DistinguishedFolderIdNameType.inbox);
            this.FolderIdNameType.Add(DistinguishedFolderIdNameType.searchfolders);
            #endregion

            #region Step 2. Client invokes CreateSubFolder to create folders.
            // Generate the created folder name
            string firstLevelFolderName = Common.GenerateResourceName(this.Site, "FirstLevelFolder");
            string secondLevelFolderName = Common.GenerateResourceName(this.Site, "SecondLevelFolder");

            // Create folders under inbox, calendar, contacts, tasks and search folder
            this.CreateMultipleFolders(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                this.FolderIdNameType,
                firstLevelFolderName,
                secondLevelFolderName,
                TestSuiteBase.SearchText);
            #endregion

            #region Step 3. Client invokes SyncFolderHierarchy operation to sync the operation result in Step 2.
            // Call SyncFolderHierarchy operation to sync the create folder result 
            this.GetSyncFolderHierarchyResponseMessage();
            #endregion

            #region Step 4. Client invokes FindAndUpdateFolderName to change the folder's name.
            // Generate a new folder name
            string newFolderName = Common.GenerateResourceName(this.Site, "NewFolderName");

            // Update the specific folder's name to a new one
            this.UpdateMultipleFolders(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                this.FolderIdNameType,
                firstLevelFolderName,
                secondLevelFolderName,
                newFolderName);
            #endregion

            #region  Step 5. Client invokes SyncFolderHierarchy operation with previous syncState to sync the operation result in Step 4.
            // Call SyncFolderHierarchy operation to sync the update folder result
            this.GetSyncFolderHierarchyResponseMessage();
            #endregion

            #region Step 6. Client invokes FindAndDeleteSubFolder to delete the folder that created in step 2.
            // Delete the sub folder that created in step 2
            this.DeleteMultipleFolders(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                this.FolderIdNameType,
                newFolderName);
            #endregion

            #region Step 7. Client invokes SyncFolderHierarchy operation with previous syncState to sync the operation result of Step 6.
            // Call SyncFolderHierarchy operation to sync the delete folder result
            this.GetSyncFolderHierarchyResponseMessage();
            #endregion
        }

        /// <summary>
        /// Client calls SyncFolderHierarchy operation with all optional elements to get the synchronization information of 
        /// 5 type of folders (inbox, calendar, tasks, contacts, search).
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S03_TC02_SyncFolderHierarchy_WithAllOptionalElements()
        {
            #region Step 1. Client invokes SyncFolderHierarchy operation to server to get initial syncState.
            // Add inbox folder and search folder into list
            this.FolderIdNameType.Add(DistinguishedFolderIdNameType.inbox);
            this.FolderIdNameType.Add(DistinguishedFolderIdNameType.searchfolders);

            // Get the initial syncState
            SyncFolderHierarchyResponseMessageType[] responseMessage = new SyncFolderHierarchyResponseMessageType[this.FolderIdNameType.Count];
            SyncFolderHierarchyResponseType[] response = new SyncFolderHierarchyResponseType[this.FolderIdNameType.Count];

            for (int i = 0; i < this.FolderIdNameType.Count; i++)
            {
                SyncFolderHierarchyType request = TestSuiteHelper.CreateSyncFolderHierarchyRequest(this.FolderIdNameType[i], DefaultShapeNamesType.Default, true, true);
                response[i] = this.SYNCAdapter.SyncFolderHierarchy(request);
                responseMessage[i] = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response[i]);
            }
            #endregion

            #region Step 2. Client invokes CreateSubFolder to create folders.
            // Generate the created folder name
            string firstLevelFolderName = Common.GenerateResourceName(this.Site, "FirstLevelFolder");
            string secondLevelFolderName = Common.GenerateResourceName(this.Site, "SecondLevelFolder");

            // Create folders under inbox, calendar, contacts, tasks and search folder
            this.CreateMultipleFolders(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                this.FolderIdNameType,
                firstLevelFolderName,
                secondLevelFolderName,
                TestSuiteBase.SearchText);
            #endregion

            #region Step 3. Client invokes SyncFolderHierarchy operation to sync the operation result in Step 2.
            // Call SyncFolderHierarchy operation to sync the create folder operation result 
            for (int i = 0; i < this.FolderIdNameType.Count; i++)
            {
                responseMessage[i] = this.GetSyncFolderHierarchyResponseMessage(responseMessage[i], this.FolderIdNameType[i]);
            }
            #endregion

            #region Step 4. Client invokes FindAndUpdateFolderName to change the folder's name.
            // Generate a new folder name
            string newFolderName = Common.GenerateResourceName(this.Site, "NewFolderName");

            // Update the specific folder's name to a new one
            this.UpdateMultipleFolders(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                this.FolderIdNameType,
                firstLevelFolderName,
                secondLevelFolderName,
                newFolderName);
            #endregion

            #region  Step 5. Client invokes SyncFolderHierarchy operation with previous syncState to sync the operation result in Step 4.
            // Call SyncFolderHierarchy operation to sync the update folder operation result
            for (int i = 0; i < this.FolderIdNameType.Count; i++)
            {
                responseMessage[i] = this.GetSyncFolderHierarchyResponseMessage(responseMessage[i], this.FolderIdNameType[i]);
            }
            #endregion

            #region Step 6. Client invokes FindAndDeleteSubFolder to delete the folder that created in step 2.
            // Delete the sub folder that created in step 2
            this.DeleteMultipleFolders(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                this.FolderIdNameType,
                newFolderName);
            #endregion

            #region Step 7. Client invokes SyncFolderHierarchy operation with previous syncState to sync the operation result of Step 6.
            // Call SyncFolderHierarchy operation to sync the delete folder operation result
            for (int i = 0; i < this.FolderIdNameType.Count; i++)
            {
                responseMessage[i] = this.GetSyncFolderHierarchyResponseMessage(responseMessage[i], this.FolderIdNameType[i]);
            }
            #endregion
        }
        #endregion
    }
}