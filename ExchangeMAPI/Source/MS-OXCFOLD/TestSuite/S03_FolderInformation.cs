namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is used to verify the properties contained in a folder object or ROP operations related to a search folder.
    /// </summary>
    [TestClass]
    public class S03_FolderInformation : TestSuiteBase
    {
        /// <summary>
        /// Server response handle list.
        /// </summary>
        private List<List<uint>> responseHandles;

        #region Test class initialization.

        /// <summary>
        /// Initialize the test suite
        /// </summary>
        /// <param name="testContext">The test context instance</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Reset the test environment
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        /// <summary>
        /// This test case is designed to validate that the static search by the StaticSearch flag.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC01_StaticSearchVerification()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create the general folder [MSOXCFOLDSubfolder1] under the root folder.

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = (byte)FolderType.Genericfolder,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x00,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder1),
                Comment = Encoding.ASCII.GetBytes(Constants.Subfolder1)
            };
            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint subfolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            ulong subfolderId1 = createFolderResponse.FolderId;

            #endregion

            #region Step 2. The client creates two general messages under the folder [MSOXCFOLDSubfolder1].

            uint messageHandle1 = 0;
            ulong messageId1 = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageId1, ref messageHandle1);

            uint messageHandle2 = 0;
            ulong messageId2 = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageId2, ref messageHandle2);
            #endregion

            #region Step 3. The client calls RopCreateFolder to create the search folder [MSOXCFOLDSearchFolder2] under the root folder.

            createFolderRequest.FolderType = (byte)FolderType.Searchfolder;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.SearchFolder2);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.SearchFolder2);
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint searchFolderHandle2 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            #endregion

            #region Step 4. The client calls RopSetSearchCriteria to establish search criteria for [MSOXCFOLDSearchFolder2].

            RopSetSearchCriteriaRequest setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex
            };
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)MessagePropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            ExistRestriction existRestriction = new ExistRestriction
            {
                PropTag = propertyTag
            };
            setSearchCriteriaRequest.RestrictionDataSize = (ushort)existRestriction.Size();
            setSearchCriteriaRequest.RestrictionData = existRestriction.Serialize();
            setSearchCriteriaRequest.FolderIds = new ulong[] { subfolderId1 };
            setSearchCriteriaRequest.FolderIdCount = (ushort)setSearchCriteriaRequest.FolderIds.Length;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.ContentIndexedSearch | (uint)SetSearchFlags.StaticSearch | (uint)SetSearchFlags.RestartSearch;
            RopSetSearchCriteriaResponse setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle2, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 5. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder2].
            RopGetSearchCriteriaRequest getSearchCriteriaRequest = new RopGetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopGetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                UseUnicode = 0x00,
                IncludeRestriction = 0x01,
                IncludeFolders = 0x01
            };
            RopGetSearchCriteriaResponse getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle2, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");

            #region Verify the requirements: MS-OXCFOLD_R1200 and MS-OXCFOLD_R1233.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1200");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1200
            // SHALLOW_SEARCH bit in the RopSetSearchCriteria ROP request means the search includes only the search folder containers that are specified in the FolderIds field.
            // If the bit SEARCH_RECURSIVE in the RopGetSearchCriteria ROP response is not set means only the search folder containers that are specified in the last RopSetSearchCriteria ROP request are being searched.
            // So, if the bit SEARCH_RECURSIVE in getSearchCriteriaResponse is not set, R1200 can be verified. 
            Site.CaptureRequirementIfAreNotEqual<uint>(
                (uint)GetSearchFlags.Recursive,
                getSearchCriteriaResponse.SearchFlags & (uint)GetSearchFlags.Recursive,
                1200,
                @"[In RopSetSearchCriteria ROP Request Buffer] SearchFlags (4 bytes): If neither bit [RECURSIVE_SEARCH or SHALLOW_SEARCH] is set, the default is SHALLOW_SEARCH.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1233");

            List<ulong> folderIdsInGetSearchCriteriaResponse = new List<ulong>();
            folderIdsInGetSearchCriteriaResponse.AddRange(getSearchCriteriaResponse.FolderIds);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1233
            this.Site.CaptureRequirementIfIsFalse(
                folderIdsInGetSearchCriteriaResponse.Contains(createFolderResponse.FolderId),
                1233,
                @"[In Setting Up a Search Folder] [A search folder cannot be included in its own search scope] Therefore, the FolderIds field MUST NOT include the FID of the search folder.");
            #endregion

            #endregion

            #region Step 6. The client calls RopGetContentsTable to retrieve the contents table for the search folder [MSOXCFOLDSearchFolder2].
            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest();
            RopGetContentsTableResponse getContentsTableResponse;
            getContentsTableRequest.RopId = (byte)RopId.RopGetContentsTable;
            getContentsTableRequest.LogonId = Constants.CommonLogonId;
            getContentsTableRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            getContentsTableRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            getContentsTableRequest.TableFlags = (byte)FolderTableFlags.None;

            int count = 0;
            do
            {
                getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, searchFolderHandle2, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
                if (getContentsTableResponse.RowCount != 2)
                {
                    Thread.Sleep(this.WaitTime);
                }
                else
                {
                    break;
                }

                count++;
            }
            while (count < this.RetryCount);

            uint rowCountBeforeHardDel = getContentsTableResponse.RowCount;
            Site.Assert.AreEqual<uint>(2, rowCountBeforeHardDel, "The two general messages created in step 2 were fetched in search folder.");

            #endregion

            #region Step 7. The client calls RopHardDeleteMessage ROP operation to hard delete a message in the general folder [MSOXCFOLDSubfolder1] under the root folder.

            ulong[] messageIds = new ulong[] { messageId1 };
            RopHardDeleteMessagesRequest hardDeleteMessagesRequest = new RopHardDeleteMessagesRequest
            {
                RopId = (byte)RopId.RopHardDeleteMessages,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                NotifyNonRead = 0x00,
                MessageIdCount = (ushort)messageIds.Length,
                MessageIds = messageIds
            };
            RopHardDeleteMessagesResponse hardDeleteMessagesResponse = this.Adapter.HardDeleteMessages(hardDeleteMessagesRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, hardDeleteMessagesResponse.ReturnValue, "RopHardDeleteMessages ROP operation performs successfully!");
            Site.Assert.AreEqual<uint>(0x00, hardDeleteMessagesResponse.PartialCompletion, "RopHardDeleteMessages ROP operation is complete!");

            #endregion

            #region Step 8. The client calls RopGetContentsTable to retrieve the contents table for the search folder [MSOXCFOLDSearchFolder2].

            count = 0;
            bool searchFolderNotChange = false;
            do
            {
                getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, searchFolderHandle2, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
                if (getContentsTableResponse.RowCount != rowCountBeforeHardDel)
                {
                    Thread.Sleep(this.WaitTime);
                }
                else
                {
                    searchFolderNotChange = true;
                    break;
                }

                count++;
            }
            while (count < this.RetryCount);

            #region Verify the requirements: MS-OXCFOLD_R795, MS-OXCFOLD_R784, MS-OXCFOLD_R1084, MS-OXCFOLD_R549, MS-OXCFOLD_R1093, MS-OXCFOLD_R1094.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R795");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R795
            // The search folder contents didn't change after the message has been hard deleted, it indicates the current search is static search.
            Site.CaptureRequirementIfIsTrue(
                searchFolderNotChange,
                795,
                @"[In RopGetSearchCriteria ROP Response Buffer] SearchFlags (4 bytes): SEARCH_STATIC (0x00010000) means that the search is static.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R784");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R784
            Site.CaptureRequirementIfIsTrue(
                searchFolderNotChange,
                784,
                @"[In RopSetSearchCriteria ROP Request Buffer] SearchFlags (4 bytes): STATIC_SEARCH (0x00040000) means that the search is static, if set.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1084");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1084
            Site.CaptureRequirementIfIsTrue(
                searchFolderNotChange,
                1084,
                @"[In Processing a RopSetSearchCriteria ROP Request] For static search folders, the contents of the search folder are not updated after the initial population is complete.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R549");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R549
            Site.CaptureRequirementIfIsTrue(
                searchFolderNotChange,
                549,
                @"[In Processing a RopSetSearchCriteria ROP Request] A static search causes the search folder to be populated once with all messages that match the search criteria at the point in time when the search is started or restarted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1093");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1093
            Site.CaptureRequirementIfIsTrue(
                searchFolderNotChange,
                1093,
                @"[In Processing a RopSetSearchCriteria ROP Request] The server MUST NOT update the search folder after the initial population when new messages that match the search criteria arrive in the search scope.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1094");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1094
            Site.CaptureRequirementIfIsTrue(
                searchFolderNotChange,
                1094,
                @"[In Processing a RopSetSearchCriteria ROP Request] Or the server MUST NOT update the search folder after the initial population when existing messages that fit the search criteria are deleted.");

            #endregion

            #endregion

            #region Step 9. The client calls RopSetSearchCriteria to restart the search for [MSOXCFOLDSearchFolder2].

            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.ContentIndexedSearch | (uint)SetSearchFlags.RestartSearch;
            setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle2, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 10. The client calls RopGetContentsTable to retrieve the contents table for the search folder [MSOXCFOLDSearchFolder2].

            count = 0;
            do
            {
                getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, searchFolderHandle2, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
                if (getContentsTableResponse.RowCount != 1)
                {
                    Thread.Sleep(this.WaitTime);
                }
                else
                {
                    break;
                }

                count++;
            }
            while (count < this.RetryCount);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R551");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R551
            // There should be 1 message found and copy to the search folder, so the RowCount should be 1.
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                getContentsTableResponse.RowCount,
                551,
                @"[In Processing a RopSetSearchCriteria ROP Request] To trigger an update, another RopSetSearchCriteria ROP request with the RESTART_SEARCH bit set in the SearchFlags field, as specified in section 2.2.1.4.1, is required.");
            #endregion

            #region Step 11. The client calls RopSetSearchCriteria to establish search criteria for [MSOXCFOLDSearchFolder2] by a new RestrictionData.

            setSearchCriteriaRequest = new RopSetSearchCriteriaRequest();
            setSearchCriteriaRequest.RopId = (byte)RopId.RopSetSearchCriteria;
            setSearchCriteriaRequest.LogonId = Constants.CommonLogonId;
            setSearchCriteriaRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            ContentRestriction contentRestriction = new ContentRestriction
            {
                FuzzyLevelLow = FuzzyLevelLowValues.FL_PREFIX,
                FuzzyLevelHigh = FuzzyLevelHighValues.FL_IGNORECASE
            };
            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)MessagePropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            contentRestriction.PropertyTag = propertyTag;
            TaggedPropertyValue taggedProperty = new TaggedPropertyValue
            {
                PropertyTag = propertyTag,
                Value = Encoding.Unicode.GetBytes("IPM.Task" + Constants.StringNullTerminated)
            };
            contentRestriction.TaggedValue = taggedProperty;
            setSearchCriteriaRequest.RestrictionDataSize = (ushort)contentRestriction.Size();
            setSearchCriteriaRequest.RestrictionData = contentRestriction.Serialize();
            setSearchCriteriaRequest.FolderIds = new ulong[] { subfolderId1 };
            setSearchCriteriaRequest.FolderIdCount = (ushort)setSearchCriteriaRequest.FolderIds.Length;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.ContentIndexedSearch | (uint)SetSearchFlags.RestartSearch | (uint)SetSearchFlags.StaticSearch;
            setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle2, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 12. The client calls RopGetContentsTable to retrieve the contents table for the search folder [MSOXCFOLDSearchFolder2].

            count = 0;
            do
            {
                getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, searchFolderHandle2, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
                if (getContentsTableResponse.RowCount != 0)
                {
                    Thread.Sleep(this.WaitTime);
                }
                else
                {
                    break;
                }

                count++;
            }
            while (count < this.RetryCount);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R519");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R519
            // For no message should be found so the search folder should empty, which means the RowCount is 0.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getContentsTableResponse.RowCount,
                519,
                @"[In Processing a RopSetSearchCriteria ROP Request] When new search criteria are applied, the server modifies the search folder to include only the messages that match the new search criteria.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the  SetSearchCriteria operation by using dynamic way.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC02_DynamicSearchVerification()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create the general folder [MSOXCFOLDSubfolder1] under the root folder.

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest();
            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = Constants.CommonLogonId;
            createFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            createFolderRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            createFolderRequest.FolderType = 0x01;
            createFolderRequest.UseUnicodeStrings = 0x0;
            createFolderRequest.OpenExisting = 0x00;
            createFolderRequest.Reserved = 0x0;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder1);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder1);
            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint subfolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            ulong subfolderId1 = createFolderResponse.FolderId;

            #endregion

            #region Step 2. The client creates a non-FAI message under the general folder [MSOXCFOLDSubfolder1].

            uint messageNonFAIHandle1 = 0;
            ulong messageNonFAIId1 = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageNonFAIId1, ref messageNonFAIHandle1);

            #endregion

            #region Step 3. The client calls RopCreateFolder to create the search folder [MSOXCFOLDSearchFolder1] under the root folder.

            createFolderRequest.FolderType = 0x02;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.SearchFolder);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.SearchFolder);
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint searchFolderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            #endregion

            #region Step 4. The client calls RopSetSearchCriteria to establish search criteria for [MSOXCFOLDSearchFolder1].

            RopSetSearchCriteriaRequest setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex
            };
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)MessagePropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            ExistRestriction existRestriction = new ExistRestriction
            {
                PropTag = propertyTag
            };
            setSearchCriteriaRequest.RestrictionDataSize = (ushort)existRestriction.Size();
            setSearchCriteriaRequest.RestrictionData = existRestriction.Serialize();
            setSearchCriteriaRequest.FolderIds = new ulong[] { subfolderId1 };
            setSearchCriteriaRequest.FolderIdCount = (ushort)setSearchCriteriaRequest.FolderIds.Length;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.NonContentIndexedSearch | (uint)SetSearchFlags.RestartSearch;
            RopSetSearchCriteriaResponse setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R783");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R783
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                setSearchCriteriaResponse.ReturnValue,
                783,
                @"[In RopSetSearchCriteria ROP Request Buffer] SearchFlags (4 bytes): NON_CONTENT_INDEXED_SEARCH (0x00020000) means that the search does not use a content-indexed search.");
            #endregion

            #region Step 5. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder1].

            RopGetSearchCriteriaRequest getSearchCriteriaRequest = new RopGetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopGetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                UseUnicode = 0x00,
                IncludeRestriction = 0x01,
                IncludeFolders = 0x01
            };
            RopGetSearchCriteriaResponse getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R785");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R785
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)GetSearchFlags.Running,
                getSearchCriteriaResponse.SearchFlags & (uint)GetSearchFlags.Running,
                785,
                @"[In RopGetSearchCriteria ROP Response Buffer] SearchFlags (4 bytes): SEARCH_RUNNING (0x00000001) means that the search is running, which means that the initial population of the search folder still being compiled.");
            #endregion

            #region Step 6. The client calls RopGetContentsTable to retrieve the contents table for the search folder [MSOXCFOLDSearchFolder1].

            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest();
            RopGetContentsTableResponse getContentsTableResponse;
            getContentsTableRequest.RopId = (byte)RopId.RopGetContentsTable;
            getContentsTableRequest.LogonId = Constants.CommonLogonId;
            getContentsTableRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            getContentsTableRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            getContentsTableRequest.TableFlags = (byte)FolderTableFlags.None;

            int count = 0;
            do
            {
                getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, searchFolderHandle, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
                if (getContentsTableResponse.RowCount != 1)
                {
                    Thread.Sleep(this.WaitTime);
                }
                else
                {
                    break;
                }

                count++;
            }
            while (count < this.RetryCount);

            #region Verify the requirements: MS-OXCFOLD_R552, MS-OXCFOLD_R526.

            Site.Assert.AreEqual<uint>((uint)GetSearchFlags.Running, getSearchCriteriaResponse.SearchFlags & (uint)GetSearchFlags.Running, "The RopSearchCriteria ROP operation has not been complete.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R526");

            // If the client has received the success code from the server after sending a RopSetSearchCriteria request.
            // If the client has received a 'SEARCH_RUNNING' SearchFlags in the RopGetSearchCriteria response after sending a RopGetSearchCriteria response later.
            // If the client has got the messages which according to the search criteria and search scope that are specified in the RopSetSearchCriteria ROP request in the RopGetContentsTable response after sending a RopGetContentsTable request later.
            // Satisfy the above conditions, then this requirement can be verified directly.
            Site.CaptureRequirement(
                526,
                @"[In Processing a RopSetSearchCriteria ROP Request] The server can return the RopSetSearchCriteria ROP response before the search folder is fully updated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R552");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R552
            // There should have 1 message macthed the search criteria, so the RowCount should be 1.
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                getContentsTableResponse.RowCount,
                552,
                @"[In Processing a RopSetSearchCriteria ROP Request] A dynamic search causes the search folder to be initially populated with all messages that match the search criteria at the point in time when the search is started or restarted.");
            #endregion

            #endregion

            #region Step 7. The client calls RopSetSearchCriteria without setting the SearchFlags to establish search criteria for [MSOXCFOLDSearchFolder1].

            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.None;
            setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 8. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder1].

            getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 9. The client calls RopGetContentsTable to retrieve the contents table for the search folder [MSOXCFOLDSearchFolder1].

            count = 0;
            do
            {
                getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, searchFolderHandle, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
                if (getContentsTableResponse.RowCount != 1)
                {
                    Thread.Sleep(this.WaitTime);
                }
                else
                {
                    break;
                }

                count++;
            }
            while (count < this.RetryCount);
            #endregion

            #region Step 10. The client creates a non-FAI message under the general folder [MSOXCFOLDSubfolder1].

            uint messageNonFAIHandle3 = 0;
            ulong messageNonFAIId3 = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageNonFAIId3, ref messageNonFAIHandle3);

            #endregion

            #region Step 11. The client calls RopGetContentsTable to retrieve the contents table for the search folder [MSOXCFOLDSearchFolder1].
            count = 0;
            do
            {
                getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, searchFolderHandle, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
                if (getContentsTableResponse.RowCount != 2)
                {
                    Thread.Sleep(this.WaitTime);
                }
                else
                {
                    break;
                }

                count++;
            }
            while (count < this.RetryCount);

            #region Verify the requirements: MS-OXCFOLD_R1095, MS-OXCFOLD_R1210, MS-OXCFOLD_R1096, and MS-OXCFOLD_R1097.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1096");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1096
            // There should have 2 messages found so the RowCount should be 2.
            Site.CaptureRequirementIfAreEqual<uint>(
                2,
                getContentsTableResponse.RowCount,
                1096,
                @"[In Processing a RopSetSearchCriteria ROP Request] For dynamic search folders, the contents of the search folder MUST continue to be updated as messages start to match or cease to match the search criteria. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1210");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1210
            Site.CaptureRequirementIfAreEqual<uint>(
                2,
                getContentsTableResponse.RowCount,
                1210,
                @"[In RopSetSearchCriteria ROP Request Buffer] SearchFlags (4 bytes): STATIC_SEARCH (0x00040000) means that the search is dynamic, if not set.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1095");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1095
            // There should have 2 messages found so the RowCount should be 2.
            Site.CaptureRequirementIfAreEqual<uint>(
                2,
                getContentsTableResponse.RowCount,
                1095,
                @"[In Processing a RopSetSearchCriteria ROP Request] For dynamic search folders, the contents of the search folder MUST continue to be updated as messages move around the mailbox. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1097");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1097
            Site.CaptureRequirementIfAreEqual<uint>(
                2,
                getContentsTableResponse.RowCount,
                1097,
                @"[In Processing a RopSetSearchCriteria ROP Request] The server continues to update the search folder with messages that enter or exit the search criteria.");

            #endregion

            #endregion
        }

        /// <summary>
        /// This case is designed to validate that the RopSetSearchCriteria operation performs successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC03_RopSetSearchCriteriaSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client create a non-FAI message under the root folder.

            ulong messageId = 0;
            uint messageHandle = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref messageId, ref messageHandle);

            #endregion

            #region Step 2. The client calls RopCreateFolder to create a search folder [MSOXCFOLDSearchFolder1] under the root folder.

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = 0x02,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x00,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.SearchFolder),
                Comment = Encoding.ASCII.GetBytes(Constants.SearchFolder)
            };

            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint searchFolderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            #endregion

            #region Step 3. The client calls RopSetSearchCriteria to establish search criteria for [MSOXCFOLDSearchFolder1].

            ulong[] folderIds = new ulong[]
            {
                this.DefaultFolderIds[0]
            };
            RopSetSearchCriteriaRequest setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                RestrictionDataSize = 0x0005
            };
            byte[] restrictionData = { 0x08, 0x1f, 0x00, 0x1a, 0x00 };
            setSearchCriteriaRequest.RestrictionData = restrictionData;
            setSearchCriteriaRequest.FolderIdCount = (ushort)folderIds.Length;
            setSearchCriteriaRequest.FolderIds = folderIds;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.RestartSearch;

            RopSetSearchCriteriaResponse setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSetSearchCriteria ROP operation performs successfully!");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R512");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R512.
            // Method SetSearchCriteria succeeds indicates that the server responds with a RopSetSearchCriteria ROP response buffer.
            Site.CaptureRequirement(
                512,
                @"[In Processing a RopSetSearchCriteria ROP Request] The server responds with a RopSetSearchCriteria ROP response buffer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R98");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R98
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                setSearchCriteriaResponse.ReturnValue,
                98,
                @"[In RopSetSearchCriteria ROP] The RopSetSearchCriteria ROP ([MS-OXCROPS] section 2.2.4.4) establishes search criteria for a search folder.");

            #region Verify the requirements: MS-OXCFOLD_R766, MS-OXCFOLD_R409, and MS-OXCFOLD_R47.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R766");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R766
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                setSearchCriteriaResponse.ReturnValue,
                766,
                @"[In RopSetSearchCriteria ROP Request Buffer] SearchFlags (4 bytes): RESTART_SEARCH (0x00000002) means that the search is initiated, if this is the first RopSetSearchCriteria ROP request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R409");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R409
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                setSearchCriteriaResponse.ReturnValue,
                409,
                @"[In Setting Up a Search Folder] The client creates a search folder by using the RopCreateFolder ROP ([MS-OXCROPS] section 2.2.4.2) with the FolderType field set to the value 2, as specified in section 2.2.1.2.1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R47");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R47
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                setSearchCriteriaResponse.ReturnValue,
                47,
                @"[InRopCreateFolder ROP Request Buffer] FolderType (1 byte): The value 2 indicates the folder type is Search folder.");

            #endregion

            #endregion

            #region Step 4. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder1].

            RopGetSearchCriteriaRequest getSearchCriteriaRequest = new RopGetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopGetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                UseUnicode = 0x00,
                IncludeRestriction = 0x01,
                IncludeFolders = 0x01
            };

            RopGetSearchCriteriaResponse getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");

            #region Verify the requirements: MS-OXCFOLD_R1200, MS-OXCFOLD_R772.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1200");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1200
            Site.CaptureRequirementIfAreEqual<ushort>(
                (ushort)folderIds.Length,
                getSearchCriteriaResponse.FolderIdCount,
                1200,
                @"[In RopSetSearchCriteria ROP Request Buffer] SearchFlags (4 bytes): If neither bit [RECURSIVE_SEARCH or SHALLOW_SEARCH] is set, the default is SHALLOW_SEARCH.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R772");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R772
            Site.CaptureRequirementIfAreEqual<ushort>(
                (ushort)folderIds.Length,
                getSearchCriteriaResponse.FolderIdCount,
                772,
                @"[In RopSetSearchCriteria ROP Request Buffer] SearchFlags (4 bytes): SHALLOW_SEARCH (0x00000008) means that the search includes only the search folder containers that are specified in the FolderIds field.");

            #endregion
            #endregion

            #region Step 5. The client calls RopSetSearchCriteria with no folder ID specified for search folder [MSOXCFOLDSearchFolder1] which is initialized in previous request.
            setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                RestrictionData = restrictionData,
                SearchFlags = (uint)SetSearchFlags.RestartSearch,
                FolderIdCount = 0,
                FolderIds = null
            };
            setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");
            #endregion

            #region Step 6. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder1].

            getSearchCriteriaRequest = new RopGetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopGetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                UseUnicode = 0x00,
                IncludeRestriction = 0x01,
                IncludeFolders = 0x01
            };

            RopGetSearchCriteriaResponse getSearchCriteriaResponseForNoFolderID = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1241");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1241
            this.Site.CaptureRequirementIfIsTrue(
                this.CompareFolderIDs(getSearchCriteriaResponse.FolderIds, getSearchCriteriaResponseForNoFolderID.FolderIds),
                1241,
                @"[In Processing a RopSetSearchCriteria ROP Request] If the client does not specify FIDs in a subsequent RopSetSearchCriteria ROP request, the server uses the FIDs that were specified in the previous request.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopGetSearchCriteria operation performs successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC04_RopGetSearchCriteriaSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create a search folder [MSOXCFOLDSearchFolder1] under the root folder.

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = 0x02,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x00,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.SearchFolder),
                Comment = Encoding.ASCII.GetBytes(Constants.SearchFolder)
            };

            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint searchFolderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            #endregion

            #region Step 2. The client calls RopSetSearchCriteria to establish search criteria for [MSOXCFOLDSearchFolder1].

            ulong[] folderIds = new ulong[]
            {
                this.DefaultFolderIds[0],
                this.DefaultFolderIds[1],
                this.DefaultFolderIds[3],
                this.DefaultFolderIds[4],
                this.DefaultFolderIds[5],
                this.DefaultFolderIds[6],
                this.DefaultFolderIds[7],
                this.DefaultFolderIds[8],
                this.DefaultFolderIds[9],
                this.DefaultFolderIds[10],
                this.DefaultFolderIds[11],
            };
            byte[] restrictionData = { 0x08, 0x1f, 0x00, 0x1a, 0x00 };
            RopSetSearchCriteriaRequest setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                RestrictionDataSize = 0x0005,
                RestrictionData = restrictionData,
                FolderIdCount = (ushort)folderIds.Length,
                FolderIds = folderIds,
                SearchFlags = (uint)SetSearchFlags.RestartSearch
            };

            RopSetSearchCriteriaResponse setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSetSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 3. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder1].

            RopGetSearchCriteriaRequest getSearchCriteriaRequest = new RopGetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopGetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                UseUnicode = 0x00,
                IncludeRestriction = 0x01,
                IncludeFolders = 0x01
            };

            RopGetSearchCriteriaResponse getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");
            RopGetSearchCriteriaResponse getSearchCriteriaResponse1 = getSearchCriteriaResponse;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1100");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1100.
            // The method GetSearchCriteria succeeds indicates that the server responds with a RopGetSearchCriteria ROP response buffer.
            Site.CaptureRequirement(
                1100,
                @"[In Processing a RopGetSearchCriteria ROP Request] The server responds with a RopGetSearchCriteria ROP response buffer.");

            #region Verify the requirements: MS-OXCFOLD_R129, MS-OXCFOLD_R131, MS-OXCFOLD_R959, MS-OXCFOLD_R2141, MS-OXCFOLD_R1103, and MS-OXCFOLD_R1104.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R129");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R129
            Site.CaptureRequirementIfIsTrue(
                getSearchCriteriaResponse1.RestrictionData != null && getSearchCriteriaResponse1.RestrictionData.Length > 0,
                129,
                @"[In RopGetSearchCriteria ROP Request Buffer] IncludeRestriction (1 byte): A Boolean value that is nonzero (TRUE) if the restriction data is required in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R131");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R131
            Site.CaptureRequirementIfIsTrue(
                getSearchCriteriaResponse1.FolderIds != null && getSearchCriteriaResponse1.FolderIds.Length > 0,
                131,
                @"[In RopGetSearchCriteria ROP Request Buffer] IncludeFolders (1 byte): A Boolean value that is nonzero (TRUE) if the list of folders being searched is required in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1103");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1103
            Site.CaptureRequirementIfIsTrue(
                getSearchCriteriaResponse1.RestrictionData != null && getSearchCriteriaResponse1.RestrictionData.Length > 0,
                1103,
                @"[In Processing a RopGetSearchCriteria ROP Request] The server returns the search criteria only if the IncludeRestriction field of the ROP [RopGetSearchCriteria] request buffer is set to nonzero (TRUE), as specified in section 2.2.1.5.1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1104, the count of folders that are being searched is {0}.", getSearchCriteriaResponse1.FolderIds.Length);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1104
            bool isVerifyR1104 = getSearchCriteriaResponse1.FolderIds.Length > 0;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1104,
                1104,
                @"[In Processing a RopGetSearchCriteria ROP Request] The server returns a list of the folders that are being searched only if the IncludeFolders field of the ROP [RopGetSearchCriteria] request buffer is set to nonzero (TRUE).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R959, field RestrictionData is {0}, the value of the RestrictionDataSize field is {1}.", getSearchCriteriaResponse1.RestrictionData, getSearchCriteriaResponse1.RestrictionDataSize);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R959
            bool isVerifyR959 = getSearchCriteriaResponse1.RestrictionDataSize != 0 && getSearchCriteriaResponse1.RestrictionData != null && getSearchCriteriaResponse1.RestrictionData.Length > 0;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR959,
                959,
                @"[In RopGetSearchCriteria ROP Response Buffer] This field [RestrictionData] is present only if the value of the RestrictionDataSize field is nonzero (TRUE).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R2141, the field FolderIds is {0}, the value of the FolderIdCount field is {1}.", getSearchCriteriaResponse1.FolderIds, getSearchCriteriaResponse1.FolderIdCount);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R2141
            bool isVerifyR2141 = getSearchCriteriaResponse1.FolderIdCount != 0 && getSearchCriteriaResponse1.FolderIds != null && getSearchCriteriaResponse1.FolderIds.Length > 0;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2141,
                2141,
                @"[InRopGetSearchCriteria ROP Response Buffer] This field [FolderIds] is present only if the value of the FolderIdCount field is nonzero (TRUE).");

            #endregion

            #endregion

            #region Step 4. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder1] without setting IncludeRestriction and IncludeFolders.

            getSearchCriteriaRequest.IncludeRestriction = 0x00;
            getSearchCriteriaRequest.IncludeFolders = 0x00;

            getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");
            RopGetSearchCriteriaResponse getSearchCriteriaResponse2 = getSearchCriteriaResponse;

            #region Verify the requirements: MS-OXCFOLD_R130, MS-OXCFOLD_R132, MS-OXCFOLD_R136, and MS-OXCFOLD_R139.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R130");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R130
            Site.CaptureRequirementIfIsNull(
                getSearchCriteriaResponse2.RestrictionData,
                130,
                @"[In RopGetSearchCriteria ROP Request Buffer] IncludeRestriction (1 byte): [A Boolean value that is] zero (FALSE) otherwise [if the restriction data is not required in the response].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R132");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R132
            Site.CaptureRequirementIfIsNull(
                getSearchCriteriaResponse2.FolderIds,
                132,
                @"[In RopGetSearchCriteria ROP Request Buffer] IncludeFolders (1 byte): [A Boolean value that is] zero (FALSE) otherwise [if the list of folders being searched is not required in the response].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R136");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R136
            Site.CaptureRequirementIfAreEqual<ushort>(
                0,
                getSearchCriteriaResponse2.RestrictionDataSize,
                136,
                @"[InRopGetSearchCriteria ROP Response Buffer] RestrictionDataSize (2 bytes): If the IncludeRestriction field of the request buffer was set to zero (FALSE), the value of RestrictionDataSize will be 0.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R139");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R139
            Site.CaptureRequirementIfAreEqual<ushort>(
                0,
                getSearchCriteriaResponse2.FolderIdCount,
                139,
                @"[InRopGetSearchCriteria ROP Response Buffer] FolderIdCount (2 bytes): If the IncludeFolders field of the request buffer was set to zero (FALSE), the FolderIdCount field will be set to 0.");

            #endregion

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopGetHierarchyTable operation performs successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC05_RopGetHierarchyTableSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = 0x01,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x00,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder1),
                Comment = Encoding.ASCII.GetBytes(Constants.Subfolder1)
            };

            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint subfolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 2. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] folder under [MSOXCFOLDSubfolder1].

            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder2);

            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, subfolderHandle1, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint subfolderHandle2 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 3. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder3] folder under [MSOXCFOLDSubfolder2].

            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder3);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder3);

            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, subfolderHandle2, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            ulong subfolderId3 = createFolderResponse.FolderId;
            #endregion

            #region Step 4. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder4] folder under root folder.

            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder4);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder4);

            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            #endregion

            #region Step 5. The client calls RopGetHierarchyTable to retrieve the hierarchy table for the root folder with 'TableFlags' set as 'None'.

            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest
            {
                RopId = (byte)RopId.RopGetHierarchyTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };
            RopGetHierarchyTableResponse getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs successfully!");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1160");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1160.
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(RopGetHierarchyTableResponse),
                getHierarchyTableResponse.GetType(),
                1160,
                @"[In Processing a RopGetHierarchyTable ROP Request] The server responds with a RopGetHierarchyTable ROP response buffer.");

            #region Verify the requirements: MS-OXCFOLD_R802, MS-OXCFOLD_R318, and MS-OXCFOLD_R1163.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R318");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R318
            // There has 2 folders directly under the root folder, so the RowCount should be 2.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000002,
                getHierarchyTableResponse.RowCount,
                318,
                @"[In RopGetHierarchyTable ROP Response Buffer] RowCount (4 bytes): An integer that specifies the number of rows in the hierarchy table.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R802");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R802
            // There has 2 folders directly under the root folder, so the RowCount should be 2.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000002,
                getHierarchyTableResponse.RowCount,
                802,
                @"[In RopGetHierarchyTable ROP Request Buffer] TableFlags (1 byte): If this bit [Depth (0x04)] is not set, the hierarchy table lists only the folder's immediate child folders.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1163");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1163
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000002,
                getHierarchyTableResponse.RowCount,
                1163,
                @"[In Processing a RopGetHierarchyTable ROP Request] The server returns a hierarchy table on which table operations can be performed.");

            #endregion

            #endregion

            #region Step 6. The client calls RopGetHierarchyTable to retrieve the hierarchy table for the root folder with 'TableFlags' set as 'Depth'.

            getHierarchyTableRequest.TableFlags = (byte)FolderTableFlags.Depth;
            getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs successfully!");

            #region Verify the requirements: MS-OXCFOLD_R648, MS-OXCFOLD_R801 and MS-OXCFOLD_R100002.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R801");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R801
            // There has total 4 folders under the root folder, so the RowCount should be 4. 
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000004,
                getHierarchyTableResponse.RowCount,
                801,
                @"[In RopGetHierarchyTable ROP Request Buffer] TableFlags (1 byte): If this bit [Depth (0x04)] is set, the hierarchy table lists folders from all levels under the folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R100002");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R100002
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                getHierarchyTableResponse.ReturnValue,
                100002,
                @"[In RopGetHierarchyTable ROP] The folder can be either [a public folder or] a private mailbox folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R648. isAllowAccessSubFolders");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R648.
            // MS-OXCFOLD_R801 is verified, the current user is the owner of the root folder and its subfolders, MS-OXCFOLD_R648 can be verified directly.
            Site.CaptureRequirement(
                648,
                @"[In Processing a RopGetHierarchyTable ROP Request]The Table object that is returned MUST allow access to the subfolders of the Folder object on which the RopGetHierarchyTable ROP is executed.");

            #endregion

            #endregion

            #region Step 7. The client calls RopGetHierarchyTable to retrieve the hierarchy table for the [MSOXCFOLDSubfolder2].

            getHierarchyTableRequest.TableFlags = (byte)FolderTableFlags.SoftDeletes | (byte)FolderTableFlags.Depth;
            getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, subfolderHandle2, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs successfully!");
            Site.Assert.AreEqual<uint>(0x0000, getHierarchyTableResponse.RowCount, "Cannot get any folder if the folder has not been soft-deleted.");

            #endregion

            #region Step 8. The client calls RopDeleteFolder to soft-delete [MSOXCFOLDSubfolder3] under the root folder.

            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest
            {
                RopId = (byte)RopId.RopDeleteFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                DeleteFolderFlags = (byte)DeleteFolderFlags.DelFolders | (byte)DeleteFolderFlags.DelMessages,
                FolderId = subfolderId3
            };
            RopDeleteFolderResponse deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, subfolderHandle2, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(0, deleteFolderResponse.ReturnValue, "RopDeleteFolder ROP operation performs successfully!");
            Site.Assert.AreEqual<byte>(0x00, deleteFolderResponse.PartialCompletion, "RopDeleteFolder ROP operation is complete.");

            #endregion

            #region Step 9. The client calls RopGetHierarchyTable to retrieve the hierarchy table for the [MSOXCFOLDSubfolder2] after it has been soft-deleted.

            getHierarchyTableRequest.TableFlags = (byte)FolderTableFlags.SoftDeletes;
            getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, subfolderHandle2, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs successfully!");

            if (Common.IsRequirementEnabled(806, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R806");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R806.
                // There has 1 folder softed deleted, so the RowCount should be 1.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x0001,
                    getHierarchyTableResponse.RowCount,
                    806,
                    @"[In RopGetHierarchyTable ROP Request Buffer] TableFlags (1 byte): If this bit [SoftDeletes (0x20)] is set, the hierarchy table lists only the folders that are soft deleted.");
            }
            #endregion

            #region Step 10. The client calls RopGetHierarchyTable with setting the 'UseUnicode' TableFlags to retrieve the hierarchy table for the root folder.

            getHierarchyTableRequest.TableFlags = (byte)FolderTableFlags.Depth | (byte)FolderTableFlags.UseUnicode;
            getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs successfully!");
            uint getHierarchyTableHandle1 = this.responseHandles[0][getHierarchyTableResponse.OutputHandleIndex];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1219");
        
            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1219
            // There left 3 folders under the root folder, so the RowCount should be 3.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000003,
                getHierarchyTableResponse.RowCount,
                1219,
                @"[In RopGetHierarchyTable ROP Request Buffer] TableFlags (1 byte): If this bit [SoftDeletes (0x20)] is not set, the hierarchy table lists only the existing folders.");
            #endregion

            #region Step 11. The client get the properties information from the rows of the table.

            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagDisplayName,
                PropertyType = (ushort)PropertyType.PtypString
            };
            PropertyTag[] propertyTags = new PropertyTag[] { propertyTag };
            List<PropertyRow> propertyRows = this.GetTableRowValue(getHierarchyTableHandle1, (ushort)getHierarchyTableResponse.RowCount, propertyTags);
            Site.Assert.IsNotNull(propertyRows, "The PidTagDisplayName property value could not be retrieved from the hierarchy table object of the [MSOXCFOLDRootFolder].");

            string pidTagDisplayNameUseUnicode = Encoding.Unicode.GetString(propertyRows[0].PropertyValues[0].Value);

            #region Verify the requirement: MS-OXCFOLD_R807 and MS-OXCFOLD_R311.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R807");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R807
            Site.CaptureRequirementIfAreEqual<string>(
                Constants.Subfolder1,
                pidTagDisplayNameUseUnicode,
                807,
                @"[In RopGetHierarchyTable ROP Request Buffer] TableFlags (1 byte): If this bit [UseUnicode (0x40)] is set, the columns that contain string data are returned in Unicode format.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R311");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R311.
            // The PidTagDisplayName property value was get successfully by the table object handle get from step 10, MS-OXCFOLD_R311 can be verified directly.
            Site.CaptureRequirement(
                311,
                @"[In RopGetHierarchyTable ROP Request Buffer] OutputHandleIndex (1 byte): The output Server object for this operation [RopGetHierarchyTable ROP] is a Table object that represents the hierarchy table.");
            #endregion
            #endregion

            #region Step 12. The client calls RopGetHierarchyTable with setting the 'UseUnicode' TableFlags to retrieve the hierarchy table for the root folder.

            getHierarchyTableRequest.TableFlags = (byte)FolderTableFlags.Depth;
            getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs successfully!");
            uint getHierarchyTableHandle = this.responseHandles[0][getHierarchyTableResponse.OutputHandleIndex];

            #endregion

            #region Step 13. The client get the properties information from the rows of the table.

            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagDisplayName,
                PropertyType = (ushort)PropertyType.PtypString8
            };
            propertyTags = new PropertyTag[] { propertyTag };
            propertyRows = this.GetTableRowValue(getHierarchyTableHandle, (ushort)getHierarchyTableResponse.RowCount, propertyTags);
            Site.Assert.IsNotNull(propertyRows, "The PidTagDisplayName property value could not be retrieved from the hierarchy table object of the [MSOXCFOLDRootFolder].");

            string pidTagDisplayNameNoUseUnicodeFlag = Encoding.ASCII.GetString(propertyRows[0].PropertyValues[0].Value);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R808");
        
            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R808
            this.Site.CaptureRequirementIfAreEqual<string>(
                Constants.Subfolder1,
                pidTagDisplayNameNoUseUnicodeFlag,
                808,
                @"[In RopGetHierarchyTable ROP Request Buffer] TableFlags (1 byte): If this bit [UseUnicode (0x40)] is not set, the string data is encoded in the code page of the Logon object.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopGetContentsTabe operation performs successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC06_RopGetContentsTableSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client creates a FAI message in the root folder.

            uint messageFAIHandle = 0;
            ulong mesageFAIId = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, 0x01, ref mesageFAIId, ref messageFAIHandle);

            #endregion

            #region Step 2. The client creates two non-FAI message in the root folder.

            uint messageNonFAIHandle1 = 0;
            ulong mesageNonFAIId1 = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref mesageNonFAIId1, ref messageNonFAIHandle1);

            uint messageNonFAIHandle2 = 0;
            ulong mesageNonFAIId2 = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref mesageNonFAIId2, ref messageNonFAIHandle2);

            #endregion

            #region Step 3. The client calls RopGetContentsTable to retrieve the contents table for the root folder without 'Associated' flag.

            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };
            RopGetContentsTableResponse getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1167");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1167
            // Method GetContentsTable succeeds indicates that the server responds with a RopGetContentsTable ROP response buffer.
            Site.CaptureRequirement(
                1167,
                @"[In Processing a RopGetContentsTable ROP Request] The server responds with a RopGetContentsTable ROP response buffer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1170");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1170
            // Method GetContentsTable succeeds indicates that the server returns a contents table on which table operations can be performed.
            Site.CaptureRequirement(
                1170,
                @"[In Processing a RopGetContentsTable ROP Request] The server returns a contents table on which table operations can be performed.");

            #region Verify the requirements: MS-OXCFOLD_R663, MS-OXCFOLD_R100502,MS-OXCFOLD_R665 and MS-OXCFOLD_R333.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R100502");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R100502.
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                getContentsTableResponse.ReturnValue,
                100502,
                @"[In RopGetContentsTable ROP] This ROP [RopGetContentsTable] applies to [both public folders and] private mailboxes.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R663");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R663
            // There has 2 non-FAI messages under the root folder, so the RowCount should be 2.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000002,
                getContentsTableResponse.RowCount,
                663,
                @"[In RopGetContentsTable ROP Request Buffer] If this bit [Associated] is not set, the contents table lists only the non-FAI messages.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R333");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R333
            // There has 2 messages under the root folder, so the RowCount should be 2.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000002,
                getContentsTableResponse.RowCount,
                333,
                @"[In RopGetContentsTable ROP Response Buffer] RowCount (4 bytes): An integer that specifies the number of rows in the contents table.");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R665");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R665
            // There has only 2 messages existed, so the RowCount should be 2.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000002,
                getContentsTableResponse.RowCount,
                665,
                @"[In RopGetContentsTable ROP Request Buffer]If this bit [SoftDeletes] is not set, the contents table lists only the existing messages.");

            #endregion

            #endregion

            #region Step 4. The client calls RopGetContentsTable to retrieve the contents table for the root folder with 'Associated' flag.

            getContentsTableRequest.TableFlags = 0x02;
            getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, this.RootFolderHandle, ref this.responseHandles);
            uint tableHandle = this.responseHandles[0][getContentsTableResponse.OutputHandleIndex];
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R662");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R662
            // There has 1 FAI messages contained in the folder, so the RowCount should be 1.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000001,
                getContentsTableResponse.RowCount,
                662,
                @"[In RopGetContentsTable ROP Request Buffer] If this bit [Associated] is set, the contents table lists only the FAI messages.");
            #endregion

            #region Step 5. The client calls RopGetContentsTable to retrieve the contents table for the root folder with 'ConversationMembers' flag.

            getContentsTableRequest.TableFlags = 0x80;
            getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, this.RootFolderHandle, ref this.responseHandles);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10251: the return value of the getContentsTableResponse is {0}", getContentsTableResponse.ReturnValue);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10251
            // The'ConversationMembers' flag is set in the getContentsTableRequest, so if the getContentsTableResponse returns success. R10251 can be verified. 
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getContentsTableResponse.ReturnValue,
                10251,
                @"[In RopGetContentsTable ROP Request Buffer] This bit [ConversationMembers] is supported on the Root folder of a mailbox.");
            #endregion

            #region Step 6. Delete one non-FAI message created in step 2 in root folder.
            object ropResponse = null;
            ulong[] messageIds = new ulong[] { mesageNonFAIId1 };

            RopDeleteMessagesRequest deleteMessagesRequest = new RopDeleteMessagesRequest();
            RopDeleteMessagesResponse deleteMessagesResponse;
            deleteMessagesRequest.RopId = (byte)RopId.RopDeleteMessages;
            deleteMessagesRequest.LogonId = Constants.CommonLogonId;
            deleteMessagesRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            deleteMessagesRequest.WantAsynchronous = 0x00;

            // The server does not generate a non-read receipt for the deleted messages.
            deleteMessagesRequest.NotifyNonRead = 0x00;
            deleteMessagesRequest.MessageIdCount = (ushort)messageIds.Length;
            deleteMessagesRequest.MessageIds = messageIds;
            this.Adapter.DoRopCall(deleteMessagesRequest, this.RootFolderHandle, ref ropResponse, ref this.responseHandles);
            deleteMessagesResponse = (RopDeleteMessagesResponse)ropResponse;
            Site.Assert.AreEqual<uint>(
                0,
                deleteMessagesResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            #endregion

            #region Step 7. The client calls RopGetContentsTable to retrieve the contents table for the root folder with 'Soft-delete' set.
            if (Common.IsRequirementEnabled(1017, this.Site))
            {
                getContentsTableRequest.TableFlags = 0x20;
                getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, this.RootFolderHandle, ref this.responseHandles);

                Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1017: the expected RowCount is {0}, and actually, the RowCount in the getContentsTableResponse is {1}.", 1, getContentsTableResponse.RowCount);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1017
                // One non-FAI message has been soft-deleted in step 5, so the return count of the getContentsTableResponse should be 1.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x00000001,
                    getContentsTableResponse.RowCount,
                    1017,
                    @"[In RopGetContentsTable ROP Request Buffer] If this bit [SoftDeletes] is set, the contents table lists only the messages that are soft deleted.");
            }
            #endregion
          
            #region Step 8. Sets the properties PidTagMid visible on the content table.
            RopSetColumnsRequest setColumnsRequest;
            PropertyTag[] propertyTags = new PropertyTag[1];
            propertyTags[0].PropertyId = (ushort)MessagePropertyId.PidTagMid;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypInteger64;

            setColumnsRequest.RopId = 0x12;
            setColumnsRequest.LogonId = 0x00;
            setColumnsRequest.InputHandleIndex = 0x00;
            setColumnsRequest.PropertyTagCount = (ushort)propertyTags.Length;
            setColumnsRequest.PropertyTags = propertyTags;
            setColumnsRequest.SetColumnsFlags = 0x00; // Sync

            this.Adapter.DoRopCall(setColumnsRequest, tableHandle, ref ropResponse, ref this.responseHandles);
            #endregion

            #region Step 9. Gets the message ID of the message in content table.

            RopQueryRowsRequest queryRowsRequest;
            RopQueryRowsResponse queryRowsResponse;

            queryRowsRequest.RopId = 0x15;
            queryRowsRequest.LogonId = 0x00;
            queryRowsRequest.InputHandleIndex = 0x00;
            queryRowsRequest.QueryRowsFlags = 0x00;
            queryRowsRequest.ForwardRead = 0x01;
            queryRowsRequest.RowCount = 1;

            this.Adapter.DoRopCall(queryRowsRequest, tableHandle, ref ropResponse, ref this.responseHandles);
            queryRowsResponse = (RopQueryRowsResponse)ropResponse;
            ulong messageID = BitConverter.ToUInt64(queryRowsResponse.RowData.PropertyRows[0].PropertyValues[0].Value, 0);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R661");
        
            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R661
            this.Site.CaptureRequirementIfAreEqual<ulong>(
                mesageFAIId,
                messageID,
                661,
                @"[In Processing a RopGetContentsTable ROP Request] The Table object that is returned provides information about messages that are directly under the Folder object on which this ROP [RopGetContentsTable ROP] is executed.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the read-only properties of folder object.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC07_GetReadOnlyProperties()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.

            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);

            #endregion

            #region Step 2. The client creates a Non-FAI message and saves it in [MSOXCFOLDSubfolder1].

            uint messageHandle1 = 0;
            ulong messageId1 = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageId1, ref messageHandle1);

            #endregion

            #region Step 3. The client create a FAI message and saves it in [MSOXCFOLDSubfolder1].

            uint messageHandle2 = 0;
            ulong messageId2 = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, 0x01, ref messageId2, ref messageHandle2);

            #endregion

            #region Step 4. The client gets the read-only properties from [MSOXCFOLDSubfolder1].

            PropertyTag[] propertyTagArray = new PropertyTag[11];

            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagContentCount,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            propertyTagArray[0] = propertyTag;

            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagContentUnreadCount,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            propertyTagArray[1] = propertyTag;

            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagDeletedOn,
                PropertyType = (ushort)PropertyType.PtypTime
            };
            propertyTagArray[2] = propertyTag;

            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagFolderId,
                PropertyType = (ushort)PropertyType.PtypInteger64
            };
            propertyTagArray[3] = propertyTag;

            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagParentEntryId,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            propertyTagArray[4] = propertyTag;

            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagHierarchyChangeNumber,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            propertyTagArray[5] = propertyTag;

            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagSubfolders,
                PropertyType = (ushort)0x000B
            };
            propertyTagArray[6] = propertyTag;

            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagMessageSize,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            propertyTagArray[7] = propertyTag;

            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagMessageSizeExtended,
                PropertyType = (ushort)PropertyType.PtypInteger64
            };
            propertyTagArray[8] = propertyTag;

            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagDeletedCountTotal,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            propertyTagArray[9] = propertyTag;

            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagHierRev,
                PropertyType = (ushort)PropertyType.PtypTime
            };
            propertyTagArray[10] = propertyTag;

            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest();
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;
            getPropertiesSpecificRequest.RopId = (byte)RopId.RopGetPropertiesSpecific;
            getPropertiesSpecificRequest.LogonId = Constants.CommonLogonId;
            getPropertiesSpecificRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            getPropertiesSpecificRequest.PropertySizeLimit = 0xFFFF;
            getPropertiesSpecificRequest.PropertyTagCount = (ushort)propertyTagArray.Length;
            getPropertiesSpecificRequest.PropertyTags = propertyTagArray;

            getPropertiesSpecificResponse = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse1 = getPropertiesSpecificResponse;
            #endregion

            #region Verify the requirements: MS-OXCFOLD_R10347, MS-OXCFOLD_R10345, MS-OXCFOLD_R346, MS-OXCFOLD_R10351 and MS-OXCFOLD_R1030, MS-OXCFOLD_R10027, MS-OXCFOLD_R10353, MS-OXCFOLD_R10354, MS-OXCFOLD_R352001 and MS-OXCFOLD_R352002.
            uint pidTagMessageSize = BitConverter.ToUInt32(getPropertiesSpecificResponse1.RowData.PropertyValues[7].Value, 0);
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10353");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10353
            Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                pidTagMessageSize,
                10353,
                @"[In PidTagMessageSize Property] The PidTagMessageSize property ([MS-OXPROPS] section 2.787) specifies the aggregate size of messages in the folder.");
            ulong pidTagMessageSizeExtended = BitConverter.ToUInt64(getPropertiesSpecificResponse1.RowData.PropertyValues[8].Value, 0);
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10354: the 64 bit value of the pidTagMessageSize property is {0}, and the value of the pidTagMessageSizeExtended is {1}", pidTagMessageSize, pidTagMessageSizeExtended);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10354
            Site.CaptureRequirementIfAreEqual<ulong>(
                (ulong)pidTagMessageSize,
                pidTagMessageSizeExtended,
                10354,
                @"[In PidTagMessageSizeExtended Property] The PidTagMessageSizeExtended property ([MS-OXPROPS] section 2.788) specifies the 64-bit version of the PidTagMessageSize property (section 2.2.2.2.1.7).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R354");

            // Client call RopGetPropertiesSpecific to get the PidTagMessageSizeExtended property.
            // If the value of PidTagMessageSizeExtended property can return and type of this property is ulong, then R354 will be verified.
            Site.CaptureRequirementIfIsInstanceOfType(
                pidTagMessageSizeExtended,
                typeof(ulong),
                354,
                @"[In PidTagMessageSizeExtended Property] Type: PtypInteger64 ([MS-OXCDATA] section 2.11.1)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7009");

            Site.CaptureRequirementIfIsInstanceOfType(
                pidTagMessageSizeExtended,
                typeof(ulong),
                Constants.MSOXPROPS,
                7009,
                @"[In PidTagMessageSizeExtended] Data type: PtypInteger64, 0x0014.");

            // [MSOXCFOLDSubfolder1] should have no unread message.
            uint unreadMessagesCountExpect = 0;

            // The index of PidTagContentUnreadCount is 1 according to the RopGetPropertiesSpecific ROP request.
            uint unreadMessagesCountActual = BitConverter.ToUInt32(getPropertiesSpecificResponse1.RowData.PropertyValues[1].Value, 0);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10347");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10347.
            Site.CaptureRequirementIfAreEqual<uint>(
                unreadMessagesCountExpect,
                unreadMessagesCountActual,
                10347,
                @"[In PidTagContentUnreadCount Property] The PidTagContentUnreadCount property ([MS-OXPROPS] section 2.639) specifies the number of unread messages in a folder, as computed by the message store.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10345");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10345
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                BitConverter.ToInt32(getPropertiesSpecificResponse1.RowData.PropertyValues[0].Value, 0),
                10345,
                @"[In PidTagContentCount Property] The PidTagContentCount property ([MS-OXPROPS] section 2.637) specifies the number of messages in a folder, as computed by the message store.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R346");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R346
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                BitConverter.ToInt32(getPropertiesSpecificResponse1.RowData.PropertyValues[0].Value, 0),
                346,
                @"[In PidTagContentCount Property] The value does not include FAI entries in the folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10351");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10351
            Site.CaptureRequirementIfAreEqual<ulong>(
                subfolderId1,
                BitConverter.ToUInt64(getPropertiesSpecificResponse1.RowData.PropertyValues[3].Value, 0),
                10351,
                @"[In PidTagFolderId Property] The PidTagFolderId property ([MS-OXPROPS] section 2.691) contains a FID structure ([MS-OXCDATA] section 2.2.1.1) that uniquely identifies a folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1030");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1030
            Site.CaptureRequirementIfAreEqual<bool>(
                false,
                BitConverter.ToBoolean(getPropertiesSpecificResponse1.RowData.PropertyValues[6].Value, 0),
                1030,
                @"[In PidTagSubfolders Property] [if the folder does not have subfolders,] the value is zero otherwise.");

            uint pidTagDeletedCountTotal = BitConverter.ToUInt32(getPropertiesSpecificResponse1.RowData.PropertyValues[9].Value, 0);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R3006");
        
            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R3006
            // Because no message has been deleted on [MSOXCFOLDSubfolder1]. So if the PidTagDeletedCountTotal property is 0, then R3006 will be verified.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                pidTagDeletedCountTotal,
                3006,
                @"[In PidTagDeletedCountTotal Property] The PidTagDeletedCountTotal property ([MS-OXPROPS] section 2.660) specifies the total number of messages that have been deleted from a folder, excluding messages that have been deleted from the folder's subfolders.");
 
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R352001");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R352001
            Site.CaptureRequirementIfAreEqual<ushort>(
                (ushort)PropertyType.PtypTime,
                propertyTagArray[10].PropertyType,
                352001,
                "[In PidTagHierRev Property] Type: PtypTime ([MS-OXCDATA] section 2.11.1)");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R352002");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R352002
            // If the property value in getPropertiesSpecificResponse1 is not null, then the PidTagHierRev property is returned from server.
            this.Site.CaptureRequirementIfIsTrue(
                getPropertiesSpecificResponse1.RowData.PropertyValues[10].Value.Length > 0,
                352002,
                @"[In PidTagHierRev Property] The PidTagHierRev property ([MS-OXPROPS] section 2.712) specifies the time, in Coordinated Universal Time (UTC), to trigger the client in cached mode to synchronize the folder hierarchy.");
            #endregion

            #region Step 5. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] under [MSOXCFOLDSubfolder1].

            uint subfolderHandle2 = 0;
            ulong subfolderId2 = 0;
            this.CreateFolder(subfolderHandle1, Constants.Subfolder2, ref subfolderId2, ref subfolderHandle2);

            #endregion

            #region Step 6. The client gets the read-only properties from [MSOXCFOLDSubfolder1] after creating [MSOXCFOLDSubfolder2] under [MSOXCFOLDSubfolder1].

            getPropertiesSpecificResponse = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse2 = getPropertiesSpecificResponse;

            #endregion

            #region Verify the requirements: MS-OXCFOLD_R1029, MS-OXCFOLD_R10355.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1029");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1029
            Site.CaptureRequirementIfIsTrue(
                BitConverter.ToBoolean(getPropertiesSpecificResponse.RowData.PropertyValues[6].Value, 0),
                1029,
                @"[In PidTagSubfolders Property] The value of this property [PidTagSubfolders] is nonzero if the folder has subfolders.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10355");

            // MS-OXCFOLD_R1029 and MS-OXCFOLD_R1030 is verified, MS-OXCFOLD_R10355 can be verified directly.
            Site.CaptureRequirement(
                10355,
                @"[In PidTagSubfolders Property] The PidTagSubfolders property ([MS-OXPROPS] section 2.1022) specifies whether the folder has any subfolders.");

            #endregion

            #region Step 7. The client calls RopDeleteFolder to delete [MSOXCFOLDSubfolder2] under [MSOXCFOLDSubfolder1].

            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest
            {
                RopId = (byte)RopId.RopDeleteFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                DeleteFolderFlags = (byte)DeleteFolderFlags.DelFolders | (byte)DeleteFolderFlags.DelMessages,
                FolderId = subfolderId2
            };
            RopDeleteFolderResponse deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, deleteFolderResponse.ReturnValue, "RopDeleteFolderResponse ROP operation performs successfully!");
            #endregion

            #region Step 8. The client gets the read-only properties from [MSOXCFOLDSubfolder1] after deleting [MSOXCFOLDSubfolder2] under [MSOXCFOLDSubfolder1].

            getPropertiesSpecificResponse = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");

            #region Verify the requirements: MS-OXCFOLD_R10028, MS-OXCFOLD_R352.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10028, the value of property PidTagHierarchyChangeNumber after getting the read-only properties from [MSOXCFOLDSubfolder1] is {0}, the value of property PidTagHierarchyChangeNumber after getting the read-only properties from [MSOXCFOLDSubfolder1] after creating [MSOXCFOLDSubfolder2] under [MSOXCFOLDSubfolder1] is {1}, the value of property PidTagHierarchyChangeNumber after getting the read-only properties from [MSOXCFOLDSubfolder1] after deleting [MSOXCFOLDSubfolder2] under [MSOXCFOLDSubfolder1] is {2}.", getPropertiesSpecificResponse1.RowData.PropertyValues[5].Value, getPropertiesSpecificResponse2.RowData.PropertyValues[5].Value, getPropertiesSpecificResponse.RowData.PropertyValues[5].Value);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10028
            bool changeNumberIncreased =
                BitConverter.ToInt32(getPropertiesSpecificResponse1.RowData.PropertyValues[5].Value, 0) < BitConverter.ToInt32(getPropertiesSpecificResponse2.RowData.PropertyValues[5].Value, 0) &&
                BitConverter.ToInt32(getPropertiesSpecificResponse2.RowData.PropertyValues[5].Value, 0) < BitConverter.ToInt32(getPropertiesSpecificResponse.RowData.PropertyValues[5].Value, 0);

            Site.CaptureRequirementIfIsTrue(
                changeNumberIncreased,
                10028,
                @"[In PidTagHierarchyChangeNumber Property] The PidTagHierarchyChangeNumber property ([MS-OXPROPS] section 2.711) specifies the number of subfolders in the folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R352, the value of property PidTagHierarchyChangeNumber after getting the read-only properties from [MSOXCFOLDSubfolder1] is {0}, the value of property PidTagHierarchyChangeNumber after getting the read-only properties from [MSOXCFOLDSubfolder1] after creating [MSOXCFOLDSubfolder2] under [MSOXCFOLDSubfolder1] is {1}, the value of property PidTagHierarchyChangeNumber after getting the read-only properties from [MSOXCFOLDSubfolder1] after deleting [MSOXCFOLDSubfolder2] under [MSOXCFOLDSubfolder1] is {2}.", getPropertiesSpecificResponse1.RowData.PropertyValues[5].Value, getPropertiesSpecificResponse2.RowData.PropertyValues[5].Value, getPropertiesSpecificResponse.RowData.PropertyValues[5].Value);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R352.
            Site.CaptureRequirementIfIsTrue(
                changeNumberIncreased,
                352,
                @"[In PidTagHierarchyChangeNumber Property] The value of this property [PidTagHierarchyChangeNumber] monotonically increases every time a subfolder is added to or deleted from the folder.");

            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the read and write properties of folder object.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC08_SetAndGetProperties()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = 0x01,
                UseUnicodeStrings = 0x1,
                OpenExisting = 0x00,
                Reserved = 0x0,
                DisplayName = Encoding.Unicode.GetBytes(Constants.Subfolder1),
                Comment = Encoding.Unicode.GetBytes(Constants.Subfolder1)
            };

            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            ulong subfolderId1 = createFolderResponse.FolderId;
            uint subfolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 2. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] under the root folder.

            createFolderRequest.FolderType = 0x02;
            createFolderRequest.DisplayName = Encoding.Unicode.GetBytes(Constants.Subfolder2);
            createFolderRequest.Comment = Encoding.Unicode.GetBytes(Constants.StringNullTerminated);

            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint subfolderHandle2 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 3. The client creates a Non-FAI message and saves it in [MSOXCFOLDSubfolder1].

            uint messageHandle1 = 0;
            ulong messageId1 = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageId1, ref messageHandle1);

            #endregion

            #region Step 4. The client create a FAI message and saves it in [MSOXCFOLDSubfolder2].

            uint messageHandle2 = 0;
            ulong messageId2 = 0;
            this.CreateSaveMessage(subfolderHandle2, subfolderId1, 0x01, ref messageId2, ref messageHandle2);

            #endregion

            #region Step 5. The client get the properties from the [MSOXCFOLDSubfolder1] under the root folder.

            PropertyTag[] tags = new PropertyTag[7];
            PropertyTag tag;

            tag.PropertyId = (ushort)FolderPropertyId.PidTagAttributeHidden;
            tag.PropertyType = (ushort)PropertyType.PtypBoolean;
            tags[0] = tag;

            tag.PropertyId = (ushort)FolderPropertyId.PidTagComment;
            tag.PropertyType = (ushort)PropertyType.PtypString;
            tags[1] = tag;

            tag.PropertyId = (ushort)FolderPropertyId.PidTagContainerClass;
            tag.PropertyType = (ushort)PropertyType.PtypString;
            tags[2] = tag;

            tag.PropertyId = (ushort)FolderPropertyId.PidTagDisplayName;
            tag.PropertyType = (ushort)PropertyType.PtypString;
            tags[3] = tag;

            tag.PropertyId = (ushort)FolderPropertyId.PidTagFolderType;
            tag.PropertyType = (ushort)PropertyType.PtypInteger32;
            tags[4] = tag;

            tag.PropertyId = (ushort)FolderPropertyId.PidTagRights;
            tag.PropertyType = (ushort)PropertyType.PtypInteger32;
            tags[5] = tag;

            tag.PropertyId = (ushort)FolderPropertyId.PidTagAccessControlListData;
            tag.PropertyType = (ushort)PropertyType.PtypBinary;
            tags[6] = tag;

            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest();
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = new RopGetPropertiesSpecificResponse();
            getPropertiesSpecificRequest.RopId = (byte)RopId.RopGetPropertiesSpecific;
            getPropertiesSpecificRequest.LogonId = Constants.CommonLogonId;
            getPropertiesSpecificRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            getPropertiesSpecificRequest.PropertySizeLimit = 0xFFFF;
            getPropertiesSpecificRequest.PropertyTagCount = (ushort)tags.Length;
            getPropertiesSpecificRequest.PropertyTags = tags;
            getPropertiesSpecificRequest.WantUnicode = 0x01;
            getPropertiesSpecificResponse = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse1 = getPropertiesSpecificResponse;

            #region Verify the requirements: MS_OXCFOLD_R1033, MS-OXCFOLD_R10359, and MS-OXCFOLD_R811.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1033");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1033.
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                getPropertiesSpecificResponse.RowData.PropertyValues[0].Value[0],
                1033,
                @"[In PidTagAttributeHidden Property] The value is zero otherwise [If the folder is not hidden].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10359");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10359
            Site.CaptureRequirementIfAreEqual<string>(
                Constants.Subfolder1,
                Encoding.Unicode.GetString(getPropertiesSpecificResponse.RowData.PropertyValues[1].Value),
                10359,
                @"[In PidTagComment Property] The PidTagComment property ([MS-OXPROPS] section 2.628) contains a comment about the purpose or content of the folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R811");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R811
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                BitConverter.ToInt32(getPropertiesSpecificResponse.RowData.PropertyValues[4].Value, 0),
                811,
                @"[In PidTagFolderType Property] FOLDER_GENERIC (1): A generic folder that contains messages and other folders.");
            #endregion

            #endregion

            #region Step 6. The client set the read/write properties for the folder.

            TaggedPropertyValue[] taggedPropertyValueArray = new TaggedPropertyValue[6];
            PropertyTag tempPropertyTag = new PropertyTag();
            int size = 0;

            // Set PidTagAttributeHidden property for the folder.
            taggedPropertyValueArray[0] = new TaggedPropertyValue();
            tempPropertyTag.PropertyId = (ushort)FolderPropertyId.PidTagAttributeHidden;
            tempPropertyTag.PropertyType = (ushort)PropertyType.PtypBoolean;
            taggedPropertyValueArray[0].PropertyTag = tempPropertyTag;
            taggedPropertyValueArray[0].Value = new byte[1] { 0x01 };

            // Set PidTagComment property for the folder.
            taggedPropertyValueArray[1] = new TaggedPropertyValue();
            tempPropertyTag.PropertyId = (ushort)FolderPropertyId.PidTagComment;
            tempPropertyTag.PropertyType = (ushort)PropertyType.PtypString;
            taggedPropertyValueArray[1].PropertyTag = tempPropertyTag;
            taggedPropertyValueArray[1].Value = Encoding.Unicode.GetBytes(Constants.Subfolder3);

            // Set PidTagContainerClass property for the folder.
            taggedPropertyValueArray[2] = new TaggedPropertyValue();
            tempPropertyTag.PropertyId = (ushort)FolderPropertyId.PidTagContainerClass;
            tempPropertyTag.PropertyType = (ushort)PropertyType.PtypString;
            taggedPropertyValueArray[2].PropertyTag = tempPropertyTag;
            taggedPropertyValueArray[2].Value = Encoding.Unicode.GetBytes("IPF.Note\0");

            // Set PidTagDisplayName property for the folder.
            taggedPropertyValueArray[3] = new TaggedPropertyValue();
            tempPropertyTag.PropertyId = (ushort)FolderPropertyId.PidTagDisplayName;
            tempPropertyTag.PropertyType = (ushort)PropertyType.PtypString;
            taggedPropertyValueArray[3].PropertyTag = tempPropertyTag;
            taggedPropertyValueArray[3].Value = Encoding.Unicode.GetBytes(Constants.Subfolder3);

            // Set PidTagFolderType property for the folder.
            taggedPropertyValueArray[4] = new TaggedPropertyValue();
            tempPropertyTag.PropertyId = (ushort)FolderPropertyId.PidTagFolderType;
            tempPropertyTag.PropertyType = (ushort)PropertyType.PtypInteger32;
            taggedPropertyValueArray[4].PropertyTag = tempPropertyTag;
            taggedPropertyValueArray[4].Value = BitConverter.GetBytes(0x00000001);

            // Set PidTagRights property for the folder.
            taggedPropertyValueArray[5] = new TaggedPropertyValue();
            tempPropertyTag.PropertyId = (ushort)FolderPropertyId.PidTagRights;
            tempPropertyTag.PropertyType = (ushort)PropertyType.PtypInteger32;
            taggedPropertyValueArray[5].PropertyTag = tempPropertyTag;
            taggedPropertyValueArray[5].Value = BitConverter.GetBytes(0x00000400);

            for (int i = 0; i < taggedPropertyValueArray.Length; i++)
            {
                size += taggedPropertyValueArray[i].Size();
            }

            RopSetPropertiesRequest setPropertiesRequest = new RopSetPropertiesRequest
            {
                RopId = (byte)RopId.RopSetProperties,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                PropertyValueSize = (ushort)(size + 2),
                PropertyValueCount = (ushort)taggedPropertyValueArray.Length,
                PropertyValues = taggedPropertyValueArray
            };

            RopSetPropertiesResponse setPropertiesResponse = this.Adapter.SetFolderObjectProperties(setPropertiesRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, setPropertiesResponse.ReturnValue, "RopSetProperties ROP operation performs successfully!");

            #endregion

            #region Step 7. The client get the properties from the [MSOXCFOLDSubfolder1] under the root folder after execute the RopSetProperties operation.

            getPropertiesSpecificResponse = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");

            #region Verify the requirements: MS_OXCFOLD_R1032, MS-OXCFOLD_R10356.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1032");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1032.
            Site.CaptureRequirementIfAreNotEqual<byte>(
                0,
                getPropertiesSpecificResponse.RowData.PropertyValues[0].Value[0],
                1032,
                @"[In PidTagAttributeHidden Property] The value of this property [PidTagAttributeHidden] is nonzero if the folder is hidden.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10356");

            // MS-OXCFOLD_R1032 and MS-OXCFOLD_R1033 were verified, MS-OXCFOLD_R10356 can be verified directly.
            Site.CaptureRequirement(
                10356,
                @"[In PidTagAttributeHidden Property] The PidTagAttributeHidden property ([MS-OXPROPS] section 2.602) specifies whether the folder is hidden.");

            string pidTagContainerClass = System.Text.Encoding.Unicode.GetString(getPropertiesSpecificResponse.RowData.PropertyValues[2].Value);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1038: the PidTagContainerClass property value is {0}", pidTagContainerClass.Replace("\0", string.Empty));
        
            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1038
            this.Site.CaptureRequirementIfIsTrue(
                pidTagContainerClass.StartsWith("IPF."),
                1038,
                @"[In PidTagContainerClass Property] The value of this property [PidTagContainerClass] MUST begin with ""IPF."".");

            // According to step above, the folder object includes the PidTagContainerClass property and the value of the PidTagContainerClass property follow the definition in Open Specification.
            // So R10036 will be verfied.
            this.Site.CaptureRequirement(
                10036,
                @"[In PidTagContainerClass Property] The PidTagContainerClass property ([MS-OXPROPS] section 2.633) specifies the type of Message object that the folder contains.");
            #endregion

            #endregion

            #region Step 8. The client get the properties from the [MSOXCFOLDSubfolder2] under the root folder.

            getPropertiesSpecificResponse = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, subfolderHandle2, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse3 = getPropertiesSpecificResponse;

            #region Verify the requirement: MS-OXCFOLD_R812 and  MS-OXCFOLD_R1034.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R812");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R812
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                BitConverter.ToInt32(getPropertiesSpecificResponse.RowData.PropertyValues[4].Value, 0),
                812,
                @"[In PidTagFolderType Property] FOLDER_SEARCH (2): A folder that contains the results of a search, in the form of links to messages that meet search criteria.");

            bool isVerifyR1034 = Encoding.Unicode.GetString(getPropertiesSpecificResponse1.RowData.PropertyValues[1].Value) == Constants.Subfolder1 &&
                Encoding.Unicode.GetString(getPropertiesSpecificResponse3.RowData.PropertyValues[1].Value) == Constants.StringNullTerminated;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1034, the property [PidTagComment] is present {0} only if the client sets it.", isVerifyR1034 ? string.Empty : "not");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1034.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR1034,
                1034,
                @"[In PidTagComment Property] This property [PidTagComment] is present only if the client sets it when the folder is created.");

            #endregion

            #endregion

            #region Step 8. The client set the same read/write properties for the [MSOXCFOLDSubfolder2] as [MSOXCFOLDSubfolder1].
            taggedPropertyValueArray = new TaggedPropertyValue[1];
            tempPropertyTag = new PropertyTag();

            // Set PidTagDisplayName property for the folder.
            taggedPropertyValueArray[0] = new TaggedPropertyValue();
            tempPropertyTag.PropertyId = (ushort)FolderPropertyId.PidTagDisplayName;
            tempPropertyTag.PropertyType = (ushort)PropertyType.PtypString;
            taggedPropertyValueArray[0].PropertyTag = tempPropertyTag;
            taggedPropertyValueArray[0].Value = Encoding.Unicode.GetBytes(Constants.Subfolder3);

            setPropertiesRequest = new RopSetPropertiesRequest
            {
                RopId = (byte)RopId.RopSetProperties,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                PropertyValueSize = (ushort)(taggedPropertyValueArray[0].Size() + 2),
                PropertyValueCount = (ushort)taggedPropertyValueArray.Length,
                PropertyValues = taggedPropertyValueArray
            };
            setPropertiesResponse = this.Adapter.SetFolderObjectProperties(setPropertiesRequest, subfolderHandle2, ref this.responseHandles);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1039");

            // Only set one property in RopSetProperties ROP request, use 0 as index here.
            int expectedIndex = 0;
            bool isVerifiedR1039 = false;
            if (setPropertiesResponse.PropertyProblems != null)
            {
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "The PropertyProblems of index {0} in RopSetProperties ROP response return value is: {1}",
                    expectedIndex,
                    setPropertiesResponse.PropertyProblems[expectedIndex].ErrorCode);
                isVerifiedR1039 = setPropertiesResponse.PropertyProblems[expectedIndex].ErrorCode != Constants.SuccessCode;
            }
            else
            {
                Site.Log.Add(LogEntryKind.Debug, "The RopSetProperties ROP response return value is: {0}", setPropertiesResponse.ReturnValue);
                isVerifiedR1039 = setPropertiesResponse.ReturnValue != Constants.SuccessCode;
            }

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1039.
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR1039,
                1039,
                @"[In PidTagDisplayName Property] Sibling folders MUST have unique display names.");
            #endregion

            #region Step 9. The client calls RopOpenFolder to get the Root folder handle.

            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderId = this.DefaultFolderIds[0],
                OpenModeFlags = (byte)FolderOpenModeFlags.OpenSoftDeleted
            };

            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, this.LogonHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, openFolderResponse.ReturnValue, "RopOpenFolder ROP operation performs successfully!");
            uint rootFolderHandle = this.responseHandles[0][openFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 10. The client get the properties from the Root folder.

            getPropertiesSpecificResponse = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, rootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R810");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R810.
            Site.CaptureRequirementIfAreEqual<int>(
                0,
                BitConverter.ToInt32(getPropertiesSpecificResponse.RowData.PropertyValues[4].Value, 0),
                810,
                @"[In PidTagFolderType Property] FOLDER_ROOT (0): The Root folder of the folder hierarchy table; that is, a folder that has no parent folder.");

            RopGetPropertiesAllResponse getPropertiesAllResponse = this.Adapter.GetFolderPropertiesAll(this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, getPropertiesAllResponse.ReturnValue, "RopGetPropertiesAllResponse ROP operation performs successfully!");
            #endregion

            #region Step 11. The client get property PidTagAddressBookEntryId
            if (Common.IsRequirementEnabled(350002, this.Site))
            {
                PropertyTag[] propertyTagArray = new PropertyTag[1];
                PropertyTag propertyTag = new PropertyTag
                {
                    PropertyId = (ushort)FolderPropertyId.PidTagAddressBookEntryId,
                    PropertyType = (ushort)PropertyType.PtypBinary
                };
                propertyTagArray[0] = propertyTag;

                getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest();
                getPropertiesSpecificRequest.RopId = (byte)RopId.RopGetPropertiesSpecific;
                getPropertiesSpecificRequest.LogonId = Constants.CommonLogonId;
                getPropertiesSpecificRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
                getPropertiesSpecificRequest.PropertySizeLimit = 0xFFFF;
                getPropertiesSpecificRequest.PropertyTagCount = (ushort)propertyTagArray.Length;
                getPropertiesSpecificRequest.PropertyTags = propertyTagArray;

                getPropertiesSpecificResponse = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, subfolderHandle1, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R350");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R350.
                // Flag value 0x01 indicates there is error.
                Site.CaptureRequirementIfAreEqual<byte>(
                    0x01,
                    getPropertiesSpecificResponse.RowData.Flag,
                    350,
                    @"[In PidTagAddressBookEntryId Property] This property is set only for public folders.");
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the RopGetHierarchyTable operation responds with error codes.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC09_RopGetHierarchyTableFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopGetHierarchyTable with a logon object handle rather than a folder handle.

            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest
            {
                RopId = (byte)RopId.RopGetHierarchyTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };

            RopGetHierarchyTableResponse getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, this.LogonHandle, ref this.responseHandles);

            #region Verify the requirements: MS-OXCFOLD_R657, MS-OXCFOLD_R658.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R657");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R657
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getHierarchyTableResponse.ReturnValue,
                657,
                @"[In Processing a RopGetHierarchyTable ROP Request]The value of error code ecNotSupported is 0x80040102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R658");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R658
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getHierarchyTableResponse.ReturnValue,
                658,
                @"[In Processing a RopGetHierarchyTable ROP Request] When the error code is ecNotSupported, it indicates that the object that this ROP [RopGetHierarchyTable ROP] was called on is not a Folder object.");

            #endregion

            #endregion

            #region Step 2. The client calls RopGetHierarchyTable with invalid 'TableFlags'.

            getHierarchyTableRequest.TableFlags = 0x03;
            getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, this.RootFolderHandle, ref this.responseHandles);

            #region Verify the requirements: MS-OXCFOLD_R314002, MS-OXCFOLD_R314001.

            if (Common.IsRequirementEnabled(314002, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R314002");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R314002
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040102,
                    getHierarchyTableResponse.ReturnValue,
                    314002,
                    @"[In Appendix A: Product Behavior] If the client sets an invalid bit in the TableFlags field, implementation fails the ROP[RopGetHierarchyTable] with an error code of ecNotSupported(0x80040102). (Microsoft Exchange 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(314001, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R314001");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R314001
                Site.CaptureRequirementIfAreEqual<uint>(
                   Constants.SuccessCode,
                    getHierarchyTableResponse.ReturnValue,
                    314001,
                    @"[In Appendix A: Product Behavior] If the client sets an invalid bit in the TableFlags field, implementation does not fail the ROP[RopGetHierarchyTable]. <19> Section 3.2.5.13:  Exchange 2007 ignores invalid bits instead of failing the ROP.");
            }

            #endregion

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopGetContentsTable operation responds with error codes.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC10_RopGetContentsTableFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client creates a FAI message in the root folder.

            uint messageFAIHandle = 0;
            ulong mesageFAIId = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, 0x01, ref mesageFAIId, ref messageFAIHandle);

            #endregion

            #region Step 2. The client creates a non-FAI message in the root folder.

            uint messageNonFAIHandle = 0;
            ulong mesageNonFAIId = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref mesageNonFAIId, ref messageNonFAIHandle);

            #endregion

            #region Step 3. The client calls RopGetContentsTable with a logon object handle rather than a folder handle.

            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };

            // Use logon object handle in which case is purposed get test error code ecNotSupported [0x80040102].  
            RopGetContentsTableResponse getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, this.LogonHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(0x80040102, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation failed with ecNotSupported [0x80040102].");

            #region Verify the requirements: MS-OXCFOLD_R671, MS-OXCFOLD_R672.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R671");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R671
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getContentsTableResponse.ReturnValue,
                671,
                @"[In Processing a RopGetContentsTable ROP Request] The value of error code ecNotSupported is 0x80040102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R672");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R672
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getContentsTableResponse.ReturnValue,
                672,
                @"[In Processing a RopGetContentsTable ROP Request] When the error code is ecNotSupported, it indicates that the object that this ROP [RopGetContentsTable ROP] was called on is not a Folder object.");

            #endregion

            #endregion

            #region Step 4. The client calls RopGetContentsTable operation with invalid 'TableFlags'.

            // Set an invalid TableFlags in the RopGetContentsTable ROP operation request.
            getContentsTableRequest.TableFlags = 0x04;
            getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, this.RootFolderHandle, ref this.responseHandles);

            #region Verify the requirements: MS-OXCFOLD_R1178, MS-OXCFOLD_R330002, MS-OXCFOLD_R1179, and MS-OXCFOLD_R330001.

            if (Common.IsRequirementEnabled(330002, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R330002");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R330002
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    getContentsTableResponse.ReturnValue,
                    330002,
                    @"[In Appendix A: Product Behavior] If the client sets an invalid bit in the TableFlags field, implementation fails the ROP [RopGetContentsTable]. (Microsoft Exchange Server 2010 and above follow this behavior.)");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1178");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1178
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    getContentsTableResponse.ReturnValue,
                    1178,
                    @"[In Processing a RopGetContentsTable ROP Request] The value of error code ecInvalidParam is 0x80070057.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1179");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1179
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    getContentsTableResponse.ReturnValue,
                    1179,
                    @"[In Processing a RopGetContentsTable ROP Request] When the error code is ecInvalidParam, it indicates that an invalid value was specified in a field.");
            }

            if (Common.IsRequirementEnabled(330001, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R330001");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R330001
                Site.CaptureRequirementIfAreEqual<uint>(
                    Constants.SuccessCode,
                    getContentsTableResponse.ReturnValue,
                    330001,
                    @"[In Appendix A: Product Behavior] If the client sets an invalid bit in the TableFlags field, implementation does not fail the ROP [RopGetContentsTable]. <22> Section 3.2.5.14:  Exchange 2007 ignores invalid bits instead of failing the ROP.");
            }

            #endregion

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopSetSearchCriteria operation responds with error codes.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC11_RopSetSearchCriteriaFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopSetSearchCriteria to establish search criteria for the root folder which is a generic folder rather a search folder.

            ulong[] folderIds = new ulong[]
                {
                    this.DefaultFolderIds[0],
                    this.DefaultFolderIds[1],
                    this.DefaultFolderIds[3],
                    this.DefaultFolderIds[4],
                    this.DefaultFolderIds[5],
                    this.DefaultFolderIds[6],
                    this.DefaultFolderIds[7],
                    this.DefaultFolderIds[8],
                    this.DefaultFolderIds[9],
                    this.DefaultFolderIds[10],
                    this.DefaultFolderIds[11],
                };

            RopSetSearchCriteriaRequest setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                RestrictionDataSize = 0x0005
            };
            byte[] restrictionData = { 0x08, 0x1f, 0x00, 0x1a, 0x00 };
            setSearchCriteriaRequest.RestrictionData = restrictionData;
            setSearchCriteriaRequest.FolderIdCount = (ushort)folderIds.Length;
            setSearchCriteriaRequest.FolderIds = folderIds;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.StopSearch;
            RopSetSearchCriteriaResponse setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, this.RootFolderHandle, ref this.responseHandles);

            #region Verify requirements: MS-OXCFOLD_R569, MS-OXCFOLD_R570, and MS-OXCFOLD_R46.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R569");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R569
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000461,
                setSearchCriteriaResponse.ReturnValue,
                569,
                @"[In Processing a RopSetSearchCriteria ROP Request]The value of error code ecNotSearchFolder is 0x00000461.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R570");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R570
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000461,
                setSearchCriteriaResponse.ReturnValue,
                570,
                @"[In Processing a RopSetSearchCriteria ROP Request] When the error code is ecNotSearchFolder, it indicates the object is not a search folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R46");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R46
            Site.CaptureRequirementIfAreNotEqual<uint>(
                Constants.SuccessCode,
                setSearchCriteriaResponse.ReturnValue,
                46,
                @"[InRopCreateFolder ROP Request Buffer] FolderType (1 byte): The value 1 indicates the folder type is Generic folder.");

            #endregion

            #endregion

            #region Step 2. The client calls RopSetSearchCriteria with a logon object handle rather a search folder handle.

            // Use logon object handle in which case is purposed to get error code ecNotSupported [0x80040102].  
            setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, this.LogonHandle, ref this.responseHandles);

            #region Verify requirement: MS-OXCFOLD_R573, MS-OXCFOLD_R574.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R573");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R573
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                setSearchCriteriaResponse.ReturnValue,
                573,
                @"[In Processing a RopSetSearchCriteria ROP Request]The value of error code ecNotSupported is 0x80040102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R574");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R574
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                setSearchCriteriaResponse.ReturnValue,
                574,
                @"[In Processing a RopSetSearchCriteria ROP Request] When the error code is ecNotSupported, it indicates the object that this ROP [RopSetSearchCriteria] was called on is not a Folder object.");

            #endregion

            #endregion

            #region Step 3. The client calls RopCreateFolder to create a search folder named [MSOXCFOLDSearchFolder1] under the root folder.

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = (byte)FolderType.Searchfolder,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x00,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.SearchFolder),
                Comment = Encoding.ASCII.GetBytes(Constants.SearchFolder)
            };
            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            uint searchFolderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            ulong searchFolderID = createFolderResponse.FolderId;
            #endregion

            #region Step 4. The client calls RopSetSearchCriteria use the handle of [MSOXCFOLDSearchFolder1] and a SearchFlags with invalid bit.

            setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                RestrictionData = restrictionData,
                SearchFlags = uint.MaxValue,
                FolderIdCount = (ushort)folderIds.Length,
                FolderIds = folderIds
            };
            setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);

            #region Verify the requirements: MS-OXCFOLD_R116001, MS-OXCFOLD_R1091, MS-OXCFOLD_R1092, and MS-OXCFOLD_R116002.

            if (Common.IsRequirementEnabled(116001, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R116001");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R116001
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    setSearchCriteriaResponse.ReturnValue,
                    116001,
                    @"[In Appendix A: Product Behavior] If the client sets an invalid bit in the SearchFlags field, implementation does fail the RopSetSearchCriteria ROP operation with ecInvalidParam (0x80070057). (Microsoft Exchange 2010 and above follow this behavior.)");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1091");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1091
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    setSearchCriteriaResponse.ReturnValue,
                    1091,
                    @"[In Processing a RopSetSearchCriteria ROP Request] The value of error code ecInvalidParam is 0x80070057.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1092");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1092
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    setSearchCriteriaResponse.ReturnValue,
                    1092,
                    @"[In Processing a RopSetSearchCriteria ROP Request] When the error code is ecInvalidParam, it indicates the SearchFlags field contains an invalid value.");
            }

            if (Common.IsRequirementEnabled(116002, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R116002");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R116002
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    0x80070057,
                    setSearchCriteriaResponse.ReturnValue,
                    116002,
                    @"[In Appendix A: Product Behavior] If the client sets an invalid bit in the SearchFlags field, implementation does not fail the RopSetSearchCriteria ROP operation with ecInvalidParam (0x80070057). <17> Section 3.2.5.4: Exchange 2007 silently ignores invalid bits and does not return the ecInvalidParam error code.");
            }

            #endregion

            #endregion

            #region Step 5. The client calls RopSetSearchCriteria with no folder ID specified for search folder [MSOXCFOLDSearchFolder1] which is not initialized.
            setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                RestrictionData = restrictionData,
                SearchFlags = (uint)SetSearchFlags.RestartSearch,
                FolderIdCount = 0,
                FolderIds = null
            };
            setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);

            #region Verify the requirements: MS-OXCFOLD_R565, MS-OXCFOLD_R566 and MS-OXCFOLD_R1240.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R565");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R565.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040605,
                setSearchCriteriaResponse.ReturnValue,
                565,
                @"[In Processing a RopSetSearchCriteria ROP Request] The value of error code ecNotInitialized is 0x80040605.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1240: the return value of RopSetSearchCriteria request with no FIDs is {0}", setSearchCriteriaResponse.ReturnValue);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1240.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040605,
                setSearchCriteriaResponse.ReturnValue,
                1240,
                @"[In Processing a RopSetSearchCriteria ROP Request] If the client does not specify FIDs, as specified in ([MS-OXCDATA] section 2.2.1.1), in the initial RopSetSearchCriteria ROP request, the server fails the ROP with ecNotInitialized (0x80040605).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R566");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R566.
            // The clients call RopSetSearchCriteria ROP on a not initialized search folder without setting folder IDs, and the error code 0x80040605 was captured, MS-OXCFOLD_R566 can be verified directly.
            Site.CaptureRequirement(
                566,
                @"[In Processing a RopSetSearchCriteria ROP Request] When the error code is ecNotInitialized, it indicates that no FIDs were specified for this search folder.");
            #endregion
            #endregion

            #region Step 6. The client calls RopSetSearchCriteria with [MSOXCFOLDSearchFolder1] was included in its own search scope.

            setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                RestrictionDataSize = 0x0005,
                RestrictionData = restrictionData,
                SearchFlags = (uint)SetSearchFlags.RestartSearch,
                FolderIds = new ulong[] { searchFolderID },
                FolderIdCount = 1,
            };
            #region Verify the requirements: MS-OXCFOLD_R124201, MS-OXCFOLD_R124202, MS-OXCFOLD_R1243 and MS-OXCFOLD_R1244.

            if (Common.IsRequirementEnabled(124201, this.Site))
            {
                setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R124201: the return value of the RopSetSearchCriteria is {0}", setSearchCriteriaResponse.ReturnValue);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R124201
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x00000490,
                    setSearchCriteriaResponse.ReturnValue,
                    124201,
                    @"[In Appendix A: Product Behavior] Implementation does fail the ROP with ecSearchFolderScopeViolation (0x00000490), if the client sets the search scope to include the search folder itself. (Exchange 2013 and above follows this behavior.)");
                
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1243: the return value of the RopSetSearchCriteria is {0}", setSearchCriteriaResponse.ReturnValue);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1243
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x00000490,
                    setSearchCriteriaResponse.ReturnValue,
                    1243,
                    @"[In Processing a RopSetSearchCriteria ROP Request] The value of error code ecSearchFolderScopeViolation is 0x00000490.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1244: the return value of the RopSetSearchCriteria is {0}", setSearchCriteriaResponse.ReturnValue);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1244
                // The client calls RopSetSearchCriteria with [MSOXCFOLDSearchFolder1] was included in its own search scope, so if the return value is ecSearchFolderScopeViolation, R1244 can be verified.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x00000490,
                    setSearchCriteriaResponse.ReturnValue,
                    1244,
                    @"[In Processing a RopSetSearchCriteria ROP Request] When the error code is ecSearchFolderScopeViolation, it indicates the search folder was included in its own search scope.");
            }

            if (Common.IsRequirementEnabled(124202, this.Site))
            {
                setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R124202: the return value of the RopSetSearchCriteria is {0}", setSearchCriteriaResponse.ReturnValue);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R124202
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    setSearchCriteriaResponse.ReturnValue,
                    124202,
                    @"[In Appendix A: Product Behavior] Implementation does not fail the RopSetSearchCriteria ROP when the search folder is included in its own search scope. <16> Section 3.2.5.4:  Exchange 2007, and Exchange 2010 do not fail the RopSetSearchCriteria ROP when the search folder is included in its own search scope.");
            }
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopGetSearchCriteria operation responds with error codes.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC12_RopGetSearchCriteriaFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopGetSearchCriteria to retrieve the hierarchy table for the root folder which is a generic folder rather than a search folder.

            RopGetSearchCriteriaRequest getSearchCriteriaRequest = new RopGetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopGetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                UseUnicode = 0x00,
                IncludeRestriction = 0x00,
                IncludeFolders = 0xFF
            };
            RopGetSearchCriteriaResponse getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, this.RootFolderHandle, ref this.responseHandles);

            #region Verify the requirements: MS-OXCFOLD_R579, MS-OXCFOLD_R580.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R579");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R579
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000461,
                getSearchCriteriaResponse.ReturnValue,
                579,
                @"[In Processing a RopGetSearchCriteria ROP Request]The value of error code ecNotSearchFolder is 0x00000461.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R580");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R580
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000461,
                getSearchCriteriaResponse.ReturnValue,
                580,
                @"[In Processing a RopGetSearchCriteria ROP Request] When the error code is ecNotSearchFolder, it indicates that the object is not a search folder.");

            #endregion

            #endregion

            #region Step 2. The client calls RopGetSearchCriteria to retrieve the hierarchy table using a logon object handle.

            // Use logon object handle in which case is purposed to get error code ecNotSupported [0x80040102].  
            getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, this.LogonHandle, ref this.responseHandles);

            #region Verify MS-OXCFOLD_R581, MS-OXCFOLD_R577

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R581");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R581
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getSearchCriteriaResponse.ReturnValue,
                581,
                @"[In Processing a RopGetSearchCriteria ROP Request]The value of error code ecNotSupported is 0x80040102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R582");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R582
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getSearchCriteriaResponse.ReturnValue,
                582,
                @"[In Processing a RopGetSearchCriteria ROP Request] When the error code is ecNotSupported, it indicates that the object that this ROP [RopGetSearchCriteria ROP] was called on is not a Folder object.");

            #endregion

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the RestrictionData field related to the search folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC13_RestrictionDataValidation()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = (byte)FolderType.Genericfolder,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x01,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder1),
                Comment = Encoding.ASCII.GetBytes(Constants.Subfolder1)
            };

            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            ulong folderId = createFolderResponse.FolderId;
            uint folderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 2. The client creates a non-FAI message under [MSOXCFOLDSubfolder1].

            ulong messageId1 = 0;
            uint messageHandle1 = 0;
            this.CreateSaveMessage(folderHandle, folderId, ref messageId1, ref messageHandle1);

            #endregion

            #region Step 3. The client calls RopCreateFolder to create a search folder [MSOXCFOLDSearchFolder1] under the root folder.

            createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = (byte)FolderType.Searchfolder,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x01,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.SearchFolder),
                Comment = Encoding.ASCII.GetBytes(Constants.SearchFolder)
            };

            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint searchFolderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 4. The client calls RopSetSearchCriteria to establish search criteria for [MSOXCFOLDSearchFolder1].

            ulong[] folderIds = new ulong[]
            {
                folderId
            };
            RopSetSearchCriteriaRequest setSearchCriteriaRequest = new RopSetSearchCriteriaRequest();
            setSearchCriteriaRequest.RopId = (byte)RopId.RopSetSearchCriteria;
            setSearchCriteriaRequest.LogonId = Constants.CommonLogonId;
            setSearchCriteriaRequest.InputHandleIndex = Constants.CommonInputHandleIndex;

            ContentRestriction contentRestriction = new ContentRestriction
            {
                FuzzyLevelLow = FuzzyLevelLowValues.FL_PREFIX,
                FuzzyLevelHigh = FuzzyLevelHighValues.FL_IGNORECASE
            };
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)MessagePropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            contentRestriction.PropertyTag = propertyTag;
            TaggedPropertyValue taggedProperty = new TaggedPropertyValue
            {
                PropertyTag = propertyTag,
                Value = Encoding.Unicode.GetBytes("IPM.Task" + Constants.StringNullTerminated)
            };
            contentRestriction.TaggedValue = taggedProperty;

            NotRestriction notRestriction = new NotRestriction
            {
                Restricts = contentRestriction
            };

            setSearchCriteriaRequest.RestrictionDataSize = (ushort)notRestriction.Size();
            byte[] restrictionData = notRestriction.Serialize();
            setSearchCriteriaRequest.RestrictionData = restrictionData;
            setSearchCriteriaRequest.FolderIdCount = (ushort)folderIds.Length;
            setSearchCriteriaRequest.FolderIds = folderIds;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.RestartSearch;

            RopSetSearchCriteriaResponse setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSetSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 5. The client calls RopGetSearchCriteria with the UseUnicode field set to 0x01 to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder1].

            RopGetSearchCriteriaRequest getSearchCriteriaRequest = new RopGetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopGetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                UseUnicode = 0x01,
                IncludeRestriction = 0x01,
                IncludeFolders = 0x01
            };

            RopGetSearchCriteriaResponse getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");
            Restriction restrictTemp = RestrictsFactory.Deserialize(getSearchCriteriaResponse.RestrictionData);
            NotRestriction notRestrictionTemp = (NotRestriction)restrictTemp;
            ContentRestriction contentRestrictionTemp = (ContentRestriction)notRestrictionTemp.Restricts;
            contentRestrictionTemp = (ContentRestriction)notRestrictionTemp.Restricts;
            string str = Encoding.Unicode.GetString(contentRestrictionTemp.TaggedValue.Value);

            #region Verify the requirements: MS-OXCFOLD_R127, MS-OXCFOLD_R137.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R127");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R127
            Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Task" + Constants.StringNullTerminated,
                str,
                127,
                @"[In RopGetSearchCriteria ROP Request Buffer] UseUnicode (1 byte): A Boolean value that is nonzero (TRUE) if the value of the RestrictionData field of the ROP response is to be in Unicode format.");

            #endregion

            #endregion

            #region Step 6. The client calls RopGetSearchCriteria with the UseUnicode field set to 0x00 to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder1].

            getSearchCriteriaRequest.UseUnicode = 0x00;

            getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");
            restrictTemp = RestrictsFactory.Deserialize(getSearchCriteriaResponse.RestrictionData);
            notRestrictionTemp = (NotRestriction)restrictTemp;
            contentRestrictionTemp = (ContentRestriction)notRestrictionTemp.Restricts;
            str = Encoding.Unicode.GetString(contentRestrictionTemp.TaggedValue.Value);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R128");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R128
            Site.CaptureRequirementIfAreNotEqual<string>(
                "IPM.Task" + Constants.StringNullTerminated,
                str,
                128,
                @"[In RopGetSearchCriteria ROP Request Buffer] UseUnicode (1 byte): [A Boolean value that is] zero (FALSE) otherwise [if the value of the RestrictionData field of the ROP response is not to be in Unicode format].");
            #endregion

            #region Step 7. The client calls RopSetSearchCriteria to restart searching criteria for [MSOXCFOLDSearchFolder1].

            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.RestartSearch;
            setSearchCriteriaRequest.RestrictionDataSize = 0;
            setSearchCriteriaRequest.RestrictionData = null;
            setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSetSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 8. The client calls RopGetSearchCriteria with the UseUnicode field set to 0x00 to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder1].

            getSearchCriteriaRequest.UseUnicode = 0x01;

            getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");
            restrictTemp = RestrictsFactory.Deserialize(getSearchCriteriaResponse.RestrictionData);
            notRestrictionTemp = (NotRestriction)restrictTemp;
            contentRestrictionTemp = (ContentRestriction)notRestrictionTemp.Restricts;
            str = Encoding.Unicode.GetString(contentRestrictionTemp.TaggedValue.Value);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R105");

            // Get the portion information of RestrictionData field is this data is same as the data set by the previous RopSetSearchCriteria request, this requirement can be verified. 
            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R105
            Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Task" + Constants.StringNullTerminated,
                str,
                105,
                @"[In RopSetSearchCriteria ROP Request Buffer] RestrictionDataSize (2 bytes): If the value of the RestrictionDataSize field is zero, the search criteria that were used most recently for the search folder container are used again.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the static search by the ContentIndexedSearch flag.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC14_ContentIndexedSearchVerification()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create the general folder [MSOXCFOLDSubfolder1] under the root folder.

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = (byte)FolderType.Genericfolder,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x00,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder1),
                Comment = Encoding.ASCII.GetBytes(Constants.Subfolder1)
            };
            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint subfolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            ulong subfolderId1 = createFolderResponse.FolderId;

            #endregion

            #region Step 2. The client creates a non-FAI message under the general folder [MSOXCFOLDSubfolder1].

            uint messageNonFAIHandle1 = 0;
            ulong messageNonFAIId1 = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageNonFAIId1, ref messageNonFAIHandle1);

            #endregion

            #region Step 3. The client calls RopCreateFolder to create the general folder [MSOXCFOLDSubfolder2] under the general folder [MSOXCFOLDSubfolder1].

            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint subfolderHandle2 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            ulong subfolderId2 = createFolderResponse.FolderId;

            #endregion

            #region Step 4. The client creates a non-FAI message under the general folder [MSOXCFOLDSubfolder2].

            uint messageNonFAIHandle2 = 0;
            ulong messageNonFAIId2 = 0;
            this.CreateSaveMessage(subfolderHandle2, subfolderId2, ref messageNonFAIId2, ref messageNonFAIHandle2);

            #endregion

            #region Step 5. The client calls RopCreateFolder to create the search folder [MSOXCFOLDSearchFolder1] under the root folder.

            createFolderRequest.FolderType = (byte)FolderType.Searchfolder;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.SearchFolder);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.SearchFolder);
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint searchFolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 6. The client calls RopSetSearchCriteria to establish search criteria for [MSOXCFOLDSearchFolder1].

            RopSetSearchCriteriaRequest setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex
            };
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)MessagePropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            ExistRestriction existRestriction = new ExistRestriction
            {
                PropTag = propertyTag
            };
            setSearchCriteriaRequest.RestrictionDataSize = (ushort)existRestriction.Size();
            setSearchCriteriaRequest.RestrictionData = existRestriction.Serialize();
            setSearchCriteriaRequest.FolderIds = new ulong[] { subfolderId1 };
            setSearchCriteriaRequest.FolderIdCount = (ushort)setSearchCriteriaRequest.FolderIds.Length;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.ContentIndexedSearch | (uint)SetSearchFlags.RestartSearch | (uint)SetSearchFlags.ForGroundSearch | (uint)SetSearchFlags.RecursiveSearch;
            RopSetSearchCriteriaResponse setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R782");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R782
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                setSearchCriteriaResponse.ReturnValue,
                782,
                @"[In RopSetSearchCriteria ROP Request Buffer] SearchFlags (4 bytes): CONTENT_INDEXED_SEARCH (0x00010000) means that the search uses a content-indexed search.");

            #endregion

            #region Step 7. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder1].

            RopGetSearchCriteriaRequest getSearchCriteriaRequest = new RopGetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopGetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                UseUnicode = 0x00,
                IncludeRestriction = 0x01,
                IncludeFolders = 0x01
            };
            RopGetSearchCriteriaResponse getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");
            ulong priviewFolderId = getSearchCriteriaResponse.FolderIds[0];

            #endregion

            #region Step 8. The client calls RopGetContentsTable to retrieve the contents table for the search folder [MSOXCFOLDSearchFolder1].

            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest();
            RopGetContentsTableResponse getContentsTableResponse;
            getContentsTableRequest.RopId = (byte)RopId.RopGetContentsTable;
            getContentsTableRequest.LogonId = Constants.CommonLogonId;
            getContentsTableRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            getContentsTableRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            getContentsTableRequest.TableFlags = (byte)FolderTableFlags.None;

            int count = 0;
            do
            {
                getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, searchFolderHandle1, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
                if (getContentsTableResponse.RowCount != 2)
                {
                    Thread.Sleep(this.WaitTime);
                }
                else
                {
                    break;
                }

                count++;
            }
            while (count < this.RetryCount);

            #region Verify the requirements: MS-OXCFOLD_R515, MS-OXCFOLD_R770, and MS-OXCFOLD_R789.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R515");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R515
            // There has 2 messages matched the search criteria, so the RowCount should be 2.
            Site.CaptureRequirementIfAreEqual<uint>(
                2,
                getContentsTableResponse.RowCount,
                515,
                @"[In Processing a RopSetSearchCriteria ROP Request] The server fills the search folder according to the search criteria and search scope that are specified in the RopSetSearchCriteria ROP request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R770");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R770
            Site.CaptureRequirementIfAreEqual<uint>(
                2,
                getContentsTableResponse.RowCount,
                770,
                @"[In RopSetSearchCriteria ROP Request Buffer]SearchFlags (4 bytes): RECURSIVE_SEARCH (0x00000004) means that the search includes the search folder containers and all of their child folders.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R789");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R789
            // There has 2 messages matched search criteria, so the RowCount should be 2.
            Site.CaptureRequirementIfAreEqual<uint>(
                2,
                getContentsTableResponse.RowCount,
                789,
                @"[In RopGetSearchCriteria ROP Response Buffer] SearchFlags (4 bytes): If this bit [SEARCH_RECURSIVE (0x00000004)] is set, the specified search folder containers and all their child search folder containers are searched for matching entries.");

            #endregion

            #endregion

            #region Step 9. The client creates a non-FAI message under the general folder [MSOXCFOLDSubfolder1].

            uint messageNonFAIHandle3 = 0;
            ulong messageNonFAIId3 = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageNonFAIId3, ref messageNonFAIHandle3);

            #endregion

            #region Step 10. The client calls RopSetSearchCriteria without setting the SearchFlags field to establish search criteria for [MSOXCFOLDSearchFolder1].

            setSearchCriteriaRequest.FolderIdCount = 0;
            setSearchCriteriaRequest.FolderIds = null;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.None;
            setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 11. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder1].

            getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R109");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R109
            Site.CaptureRequirementIfAreEqual<ulong>(
                priviewFolderId,
                getSearchCriteriaResponse.FolderIds[0],
                109,
                @"[In RopSetSearchCriteria ROP Request Buffer] FolderIdCount (2 bytes): If the FolderIdCount field is set to zero, the folders that were used in the most recent search are used again.");
            #endregion

            #region Step 12. The client calls RopCreateFolder to create the search folder [MSOXCFOLDSearchFolder2] under the root folder.

            createFolderRequest.FolderType = (byte)FolderType.Searchfolder;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.SearchFolder2);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.SearchFolder2);
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint searchFolderHandle2 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 13. The client calls RopSetSearchCriteria to establish search criteria for [MSOXCFOLDSearchFolder2].

            setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex
            };
            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)MessagePropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            existRestriction = new ExistRestriction
            {
                PropTag = propertyTag
            };
            setSearchCriteriaRequest.RestrictionDataSize = (ushort)existRestriction.Size();
            setSearchCriteriaRequest.RestrictionData = existRestriction.Serialize();
            setSearchCriteriaRequest.FolderIds = new ulong[] { subfolderId1 };
            setSearchCriteriaRequest.FolderIdCount = (ushort)setSearchCriteriaRequest.FolderIds.Length;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.ContentIndexedSearch | (uint)SetSearchFlags.RestartSearch;
            setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle2, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 14. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder2].

            getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle2, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 15. The client calls RopGetContentsTable to retrieve the contents table for the search folder [MSOXCFOLDSearchFolder2].
            count = 0;
            do
            {
                getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, searchFolderHandle2, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
                if (getContentsTableResponse.RowCount != 2)
                {
                    Thread.Sleep(this.WaitTime);
                }
                else
                {
                    break;
                }

                count++;
            }
            while (count < this.RetryCount);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R960");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R960
            // There has 2 messages matched the search criteria, so the RowCount should be 2.
            Site.CaptureRequirementIfAreEqual<uint>(
                2,
                getContentsTableResponse.RowCount,
                960,
                @"[In RopGetSearchCriteria ROP Response Buffer] If this bit [SEARCH_RECURSIVE] is not set, only the search folder containers that are specified in the last RopSetSearchCriteria ROP request (section 2.2.1.4.1) are being searched.");

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the dynamic search by the NonContentIndexedSearch flag.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC15_NonContentIndexedSearchVerification()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create the general folder [MSOXCFOLDSubfolder1] under the root folder.

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = 0x01,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x00,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder1),
                Comment = Encoding.ASCII.GetBytes(Constants.Subfolder1)
            };
            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint subfolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            ulong subfolderId1 = createFolderResponse.FolderId;

            #endregion

            #region Step 2. The client creates a non-FAI message under the general folder [MSOXCFOLDSubfolder1].

            uint messageNonFAIHandle1 = 0;
            ulong messageNonFAIId1 = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageNonFAIId1, ref messageNonFAIHandle1);

            #endregion

            #region Step 3. The client calls RopCreateFolder to create the search folder [MSOXCFOLDSearchFolder1] under the root folder.

            createFolderRequest.FolderType = (byte)FolderType.Searchfolder;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.SearchFolder);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.SearchFolder);
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint searchFolderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 4. The client calls RopSetSearchCriteria to establish search criteria for [MSOXCFOLDSearchFolder1].

            RopSetSearchCriteriaRequest setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex
            };
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)MessagePropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            ExistRestriction existRestriction = new ExistRestriction
            {
                PropTag = propertyTag
            };
            setSearchCriteriaRequest.RestrictionDataSize = (ushort)existRestriction.Size();
            setSearchCriteriaRequest.RestrictionData = existRestriction.Serialize();
            setSearchCriteriaRequest.FolderIds = new ulong[] { subfolderId1 };
            setSearchCriteriaRequest.FolderIdCount = (ushort)setSearchCriteriaRequest.FolderIds.Length;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.NonContentIndexedSearch | (uint)SetSearchFlags.RestartSearch;
            RopSetSearchCriteriaResponse setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R783");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R783
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                setSearchCriteriaResponse.ReturnValue,
                783,
                @"[In RopSetSearchCriteria ROP Request Buffer] SearchFlags (4 bytes): NON_CONTENT_INDEXED_SEARCH (0x00020000) means that the search does not use a content-indexed search.");
            #endregion

            #region Step 5. The client calls RopGetContentsTable to retrieve the contents table for the search folder [MSOXCFOLDSearchFolder1].

            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest();
            RopGetContentsTableResponse getContentsTableResponse;
            getContentsTableRequest.RopId = (byte)RopId.RopGetContentsTable;
            getContentsTableRequest.LogonId = Constants.CommonLogonId;
            getContentsTableRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            getContentsTableRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            getContentsTableRequest.TableFlags = (byte)FolderTableFlags.None;

            int count = 0;
            do
            {
                getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, searchFolderHandle, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
                if (getContentsTableResponse.RowCount != 1)
                {
                    Thread.Sleep(this.WaitTime);
                }
                else
                {
                    break;
                }

                count++;
            }
            while (count < this.RetryCount);

            #endregion

            #region Step 6. The client calls RopSetSearchCriteria to stop establish search criteria for [MSOXCFOLDSearchFolder1].

            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.StopSearch;
            setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 7. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder1].

            RopGetSearchCriteriaRequest getSearchCriteriaRequest = new RopGetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopGetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                UseUnicode = 0x00,
                IncludeRestriction = 0x01,
                IncludeFolders = 0x01
            };
            RopGetSearchCriteriaResponse getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");

            #region Verify the requirements: MS-OXCFOLD_R673, MS-OXCFOLD_R54301.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R673");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R673
            Site.CaptureRequirementIfAreNotEqual<uint>(
                (uint)GetSearchFlags.Running,
                getSearchCriteriaResponse.SearchFlags & (uint)GetSearchFlags.Running,
                673,
                @"[In RopSetSearchCriteria ROP Request Buffer] SearchFlags (4 bytes):STOP_SEARCH (0x00000001) means that the search is aborted.");

            if (Common.IsRequirementEnabled(54301, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R54301");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R54301
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    (uint)GetSearchFlags.Running,
                    getSearchCriteriaResponse.SearchFlags & (uint)GetSearchFlags.Running,
                    54301,
                    @"[In Processing a RopSetSearchCriteria ROP Request] Implementation does stop the initial population of the search folder if the STOP_SEARCH bit is set in the SearchFlags field. (Exchange 2007 and above follow this behavior).");
            }

            #endregion

            #endregion

            #region Step 8. The client calls RopSetSearchCriteria to restart establish search criteria for [MSOXCFOLDSearchFolder1].

            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.RestartSearch;
            setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 9. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder1].

            getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");
            Site.Assert.AreEqual<uint>((uint)GetSearchFlags.Running, getSearchCriteriaResponse.SearchFlags & (uint)GetSearchFlags.Running, "The SearchFlags field contains the 'SEARCH_RUNNING' bits.");

            #region Verify the requirements: MS-OXCFOLD_R546, MS-OXCFOLD_R768, and MS-OXCFOLD_R767.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R767");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R767
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)GetSearchFlags.Running,
                getSearchCriteriaResponse.SearchFlags & (uint)GetSearchFlags.Running,
                767,
                @"[In RopSetSearchCriteria ROP Request Buffer] SearchFlags (4 bytes): RESTART_SEARCH (0x00000002) means that the search is initiated, if the search is restarted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R768");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R768
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)GetSearchFlags.Running,
                getSearchCriteriaResponse.SearchFlags & (uint)GetSearchFlags.Running,
                768,
                @"[In RopSetSearchCriteria ROP Request Buffer] SearchFlags (4 bytes): RESTART_SEARCH (0x00000002) means that the search is initiated, if the search is inactive.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R546");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R546
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)GetSearchFlags.Running,
                getSearchCriteriaResponse.SearchFlags & (uint)GetSearchFlags.Running,
                546,
                @"[In Processing a RopSetSearchCriteria ROP Request] If the RESTART_SEARCH bit is set in the SearchFlags field, the server restarts the population of the search folder.");

            #endregion

            #endregion
        }

        /// <summary>
        /// This test case is designed to test readonly property PidTagFolderFlags.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC16_GetPropertyPidTagFolderFlags()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. Call RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.
            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);
            #endregion

            #region Step 2. Creates a none-FAI message in [MSOXCFOLDSubfolder1].
            ulong messageId = 0;
            uint messageHandle = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageId, ref messageHandle);
            #endregion

            #region Step 3. Create a FAI message and saves it in [MSOXCFOLDSubfolder1].
            uint messageHandle2 = 0;
            ulong messageId2 = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, 0x01, ref messageId2, ref messageHandle2);
            #endregion

            #region Step 4. The client calls RopCreateFolder to create the search folder [MSOXCFOLDSearchFolder1] under the root folder.
            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = (byte)FolderType.Searchfolder,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x00,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.SearchFolder),
                Comment = Encoding.ASCII.GetBytes(Constants.SearchFolder)
            };
            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint searchFolderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            #endregion            

            #region Step 5. The client calls RopSetSearchCriteria to establish search criteria for [MSOXCFOLDSearchFolder1].
            RopSetSearchCriteriaRequest setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex
            };
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)MessagePropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            ExistRestriction existRestriction = new ExistRestriction
            {
                PropTag = propertyTag
            };
            setSearchCriteriaRequest.RestrictionDataSize = (ushort)existRestriction.Size();
            setSearchCriteriaRequest.RestrictionData = existRestriction.Serialize();
            setSearchCriteriaRequest.FolderIds = new ulong[] { subfolderId1 };
            setSearchCriteriaRequest.FolderIdCount = (ushort)setSearchCriteriaRequest.FolderIds.Length;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.NonContentIndexedSearch | (uint)SetSearchFlags.RestartSearch;
            RopSetSearchCriteriaResponse setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");
            #endregion

            #region Step 6. The client calls RopGetContentsTable to retrieve the contents table for the search folder [MSOXCFOLDSearchFolder1].
            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest();
            RopGetContentsTableResponse getContentsTableResponse;
            getContentsTableRequest.RopId = (byte)RopId.RopGetContentsTable;
            getContentsTableRequest.LogonId = Constants.CommonLogonId;
            getContentsTableRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            getContentsTableRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            getContentsTableRequest.TableFlags = (byte)FolderTableFlags.None;

            int count = 0;
            do
            {
                getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, searchFolderHandle, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
                if (getContentsTableResponse.RowCount != 1)
                {
                    Thread.Sleep(this.WaitTime);
                }
                else
                {
                    break;
                }

                count++;
            }
            while (count < this.RetryCount);
            #endregion

            #region Step 7. The client creates rules on [MSOXCFOLDSubfolder1] folder.
            RuleData[] sampleRuleDataArray;
            object ropResponse = null;
            sampleRuleDataArray = this.CreateSampleRuleDataArrayForAdd();

            RopModifyRulesRequest modifyRulesRequest = new RopModifyRulesRequest()
            {
                RopId = (byte)RopId.RopModifyRules,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                ModifyRulesFlags = 0x00,
                RulesCount = (ushort)sampleRuleDataArray.Length,
                RulesData = sampleRuleDataArray,
            };

            modifyRulesRequest.RopId = (byte)RopId.RopModifyRules;
            modifyRulesRequest.LogonId = Constants.CommonLogonId;
            modifyRulesRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            sampleRuleDataArray = this.CreateSampleRuleDataArrayForAdd();
            modifyRulesRequest.ModifyRulesFlags = 0x00;
            modifyRulesRequest.RulesCount = (ushort)sampleRuleDataArray.Length;
            modifyRulesRequest.RulesData = sampleRuleDataArray;
            this.Adapter.DoRopCall(modifyRulesRequest, subfolderHandle1, ref ropResponse, ref this.responseHandles);
            RopModifyRulesResponse modifyRulesResponse = (RopModifyRulesResponse)ropResponse;
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, modifyRulesResponse.ReturnValue, "RopModifyRules ROP operation performs successfully!");           
            #endregion

            #region Step 8. The client calls RopOpenFolder to open [MSOXCFOLDSubfolder1] folder.

            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderId = subfolderId1,
                OpenModeFlags = (byte)FolderOpenModeFlags.None
            };
            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, openFolderResponse.ReturnValue, "RopOpenFolder ROP operation performs successfully!");
            #endregion
            
            #region Step 9. The client gets the PidTagFolderFlags propertie from [MSOXCFOLDSubfolder1].
            PropertyTag[] propertyTagArray = new PropertyTag[1];
            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagFolderFlags,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            propertyTagArray[0] = propertyTag;
                        
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest();
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;
            getPropertiesSpecificRequest.RopId = (byte)RopId.RopGetPropertiesSpecific;
            getPropertiesSpecificRequest.LogonId = Constants.CommonLogonId;
            getPropertiesSpecificRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            getPropertiesSpecificRequest.PropertySizeLimit = 0xFFFF;
            getPropertiesSpecificRequest.PropertyTagCount = (ushort)propertyTagArray.Length;
            getPropertiesSpecificRequest.PropertyTags = propertyTagArray;

            getPropertiesSpecificResponse = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");
            uint pidTagFolderFlags = BitConverter.ToUInt32(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value, 0);

            #region Verify MS-OXCFOLD_R1035110, MS-OXCFOLD_R1035111, MS-OXCFOLD_R1035103 and MS-OXCFOLD_R1035107
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, @"Verify MS-OXCFOLD_R1035110: [In PidTagFolderFlags Property] The PidTagFolderId property ([MS-OXPROPS] section 2.692) contains a computed value that specifies the type or state of a folder.");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1035110
            Site.CaptureRequirement(
                1035110,
                @"[In PidTagFolderFlags Property] The PidTagFolderId property ([MS-OXPROPS] section 2.692) contains a computed value that specifies the type or state of a folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, @"Verify MS-OXCFOLD_R1035111: [In PidTagFolderFlags Property] The value is a bitwise OR of zero or more values [1: IPM, 2: SEARCH, 4: NORMAL, 8: RULES] from the following table.");

            bool isR1035111Verified = false;
            if (((pidTagFolderFlags & 1) == 1) || ((pidTagFolderFlags & 2) == 2) || ((pidTagFolderFlags & 4) == 4) || ((pidTagFolderFlags & 8) == 8))
            {
                isR1035111Verified = true;
            }

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1035111
            Site.CaptureRequirementIfIsTrue(
                isR1035111Verified,
                1035111,
                @"[In PidTagFolderFlags Property] The value is a bitwise OR of zero or more values [1: IPM, 2: SEARCH, 4: NORMAL, 8: RULES] from the following table.");
            

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, @"Verify MS-OXCFOLD_R1035103: [In PidTagFolderFlags Property] [The folder flag named ""IPM"" specified] the folder belongs to the IPM subtree portion of the mailbox.");

            bool isR1035103Verified = false;
            if ((pidTagFolderFlags & 1) == 1)
            {
                isR1035103Verified = true;
            }

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1035103
            Site.CaptureRequirementIfIsTrue(
                isR1035103Verified,
                1035103,
                @"[In PidTagFolderFlags Property] [The folder flag named ""IPM"" specified] the folder belongs to the IPM subtree portion of the mailbox.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, @"Verify MS-OXCFOLD_R1035107: [In PidTagFolderFlags Property] [The folder flag named ""NORMAL"" specified] the folder is a generic folder that contains messages and other folders.");

            bool isR1035107Verified = false;
            if ((pidTagFolderFlags & 4) == 4)
            {
                isR1035107Verified = true;
            }

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1035107
            Site.CaptureRequirementIfIsTrue(
                isR1035107Verified,
                1035107,
                @"[In PidTagFolderFlags Property] [The folder flag named ""NORMAL"" specified] the folder is a generic folder that contains messages and other folders.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate RopSetSearchCriteria is regardless of the STATIC_SEARCH bit.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S03_TC17_RopSetSearchCriteriaRegardlessOfSTATIC_SEARCH()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create the general folder [MSOXCFOLDSubfolder1] under the root folder.

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = (byte)FolderType.Genericfolder,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x00,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder1),
                Comment = Encoding.ASCII.GetBytes(Constants.Subfolder1)
            };
            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint subfolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            ulong subfolderId1 = createFolderResponse.FolderId;

            #endregion

            #region Step 2. The client creates two general messages under the folder [MSOXCFOLDSubfolder1].

            uint messageHandle1 = 0;
            ulong messageId1 = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageId1, ref messageHandle1);

            uint messageHandle2 = 0;
            ulong messageId2 = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageId2, ref messageHandle2);
            #endregion

            #region Step 3. The client calls RopCreateFolder to create the search folder [MSOXCFOLDSearchFolder2] under the root folder.

            createFolderRequest.FolderType = (byte)FolderType.Searchfolder;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.SearchFolder2);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.SearchFolder2);
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint searchFolderHandle2 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            #endregion

            #region Step 4. The client calls RopSetSearchCriteria to establish search criteria for [MSOXCFOLDSearchFolder2] without setting STATIC_SEARCH bit.

            RopSetSearchCriteriaRequest setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex
            };
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)MessagePropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            ExistRestriction existRestriction = new ExistRestriction
            {
                PropTag = propertyTag
            };
            setSearchCriteriaRequest.RestrictionDataSize = (ushort)existRestriction.Size();
            setSearchCriteriaRequest.RestrictionData = existRestriction.Serialize();
            setSearchCriteriaRequest.FolderIds = new ulong[] { subfolderId1 };
            setSearchCriteriaRequest.FolderIdCount = (ushort)setSearchCriteriaRequest.FolderIds.Length;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.ContentIndexedSearch;
            RopSetSearchCriteriaResponse setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle2, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 5. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSearchFolder2].
            RopGetSearchCriteriaRequest getSearchCriteriaRequest = new RopGetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopGetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                UseUnicode = 0x00,
                IncludeRestriction = 0x01,
                IncludeFolders = 0x01
            };
            RopGetSearchCriteriaResponse getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle2, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");

            if (Common.IsRequirementEnabled(1238001, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1238001");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1238001
                Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)GetSearchFlags.Static,
                    getSearchCriteriaResponse.SearchFlags & (uint)GetSearchFlags.Static,
                    1238001,
                    @"[In Appendix A: Product Behavior] Implementation does be regardless of the value of the STATIC_SEARCH bit in the RopSetSearchCriteria ROP request. <14> Section 3.2.5.4:  A content-indexed search is always static on the initial release version of Exchange 2010 and Exchange 2007 regardless of the value of the STATIC_SEARCH bit in the RopSetSearchCriteria request.");
            }

            if (Common.IsRequirementEnabled(1238002, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1238002");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1238002
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    (uint)GetSearchFlags.Static,
                    getSearchCriteriaResponse.SearchFlags & (uint)GetSearchFlags.Static,
                    1238002,
                    @"[In Appendix A: Product Behavior] Implementation does not be regardless of the value of the STATIC_SEARCH bit in the RopSetSearchCriteria ROP request.(Exchange 2013 and above follow this hebavior).");
            }

            #endregion
        }

        #region Private methods
        /// <summary>
        /// Compare two folder Id arrays.
        /// </summary>
        /// <param name="firstFolderIDArray">The first folder Id array.</param>
        /// <param name="secondFolderIDArray">The second folder Id array.</param>
        /// <returns>A Boolean indicates whether the two folder Id arrays are same.</returns>
        private bool CompareFolderIDs(ulong[] firstFolderIDArray, ulong[] secondFolderIDArray)
        {
            if (firstFolderIDArray.Length != secondFolderIDArray.Length)
            {
                return false;
            }

            for (int i = 0; i < firstFolderIDArray.Length; i++)
            {
                if (firstFolderIDArray[i] != secondFolderIDArray[i])
                {
                    return false;
                }
            }

            return true;
        }
        #endregion
    }
}