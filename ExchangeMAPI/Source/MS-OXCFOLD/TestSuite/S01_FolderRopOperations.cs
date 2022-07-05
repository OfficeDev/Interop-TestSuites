namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is used to verify the ROP operations related to folder object.
    /// </summary>
    [TestClass]
    public class S01_FolderRopOperations : TestSuiteBase
    {
        /// <summary>
        /// Server response handle list.
        /// </summary>
        private List<List<uint>> responseHandles;

        #region Test Suite Initialization

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
        /// This test case is used to verify the RopOpenFolder operation with the following steps:
        /// Create a private folder and soft delete it, and then open the folder with different open flags.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC01_OpenSoftDeletedFolder()
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

            // Invoke the CreateFolder operation with valid parameters, use root folder handle to indicate that the new folder will be created under the root folder.
            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "Client call RopCreateFolder ROP operation should succeed.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R476");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R476.
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(RopCreateFolderResponse),
                createFolderResponse.GetType(),
                476,
                @"[In Processing a RopCreateFolder ROP Request] The server responds with a RopCreateFolder ROP response buffer.");

            uint createFolderReturnValue1 = createFolderResponse.ReturnValue;
            uint newFolderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            ulong newFolderId = createFolderResponse.FolderId;

            #endregion

            #region Step 2. The client gets the PidTagParentEntryId property of the [MSOXCFOLDSubfolder1].
            PropertyTag[] propertyTagArray = new PropertyTag[1];
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagParentEntryId,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            propertyTagArray[0] = propertyTag;

            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest();
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse1;
            getPropertiesSpecificRequest.RopId = (byte)RopId.RopGetPropertiesSpecific;
            getPropertiesSpecificRequest.LogonId = Constants.CommonLogonId;
            getPropertiesSpecificRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            getPropertiesSpecificRequest.PropertySizeLimit = 0xFFFF;
            getPropertiesSpecificRequest.PropertyTagCount = (ushort)propertyTagArray.Length;
            getPropertiesSpecificRequest.PropertyTags = propertyTagArray;

            getPropertiesSpecificResponse1 = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, newFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse1.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");
            #endregion

            #region Step 3. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] under the root folder with set the 'OpenExisting' flag to 0x01.

            createFolderRequest.OpenExisting = 0x01;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder2);

            // Invoke the CreateFolder operation with valid parameters, use root folder handle to indicate that the new folder will be created under the root folder.
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "Client call RopCreateFolder ROP operation should succeed.");

            uint createFolderReturnValue2 = createFolderResponse.ReturnValue;
            uint newFolderHandle2 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            ulong newFolderId2 = createFolderResponse.FolderId;

            #endregion

            #region Step 4. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder3] under the root folder.

            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder3);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder3);

            // Invoke the CreateFolder operation with valid parameters, use root folder handle to indicate that the new folder will be created under root folder.
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");

            uint newFolderHandle3 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            ulong newFolderId3 = createFolderResponse.FolderId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R483, the return value of creating the first folder is {0}, the return value of creating the second folder is {1}.", createFolderReturnValue1, createFolderReturnValue2);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R483
            bool isVerifyR483 = createFolderReturnValue1 == Constants.SuccessCode && createFolderReturnValue2 == Constants.SuccessCode;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR483,
                483,
                @"[In Processing a RopCreateFolder ROP Request] If a folder with the same name does not exist, the server creates a new folder, regardless of the value of the OpenExisting field.");
            #endregion

            #region Step 5. The client gets the PidTagParentEntryId property of the [MSOXCFOLDSubfolder3].
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse2 = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, newFolderHandle3, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse1.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");

            bool isEqual = Common.CompareByteArray(getPropertiesSpecificResponse1.RowData.PropertyValues[0].Value, getPropertiesSpecificResponse2.RowData.PropertyValues[0].Value);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10027");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10027.
            // MSOXCFOLDSubfolder1 and MSOXCFOLDSubfolder3 are under the same folder, if property PidTagParentEntryId for them are equal,
            // this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isEqual,
                10027,
                @"[In PidTagParentEntryId Property] The PidTagParentEntryId property ([MS-OXPROPS] section 2.858) contains a Folder EntryID structure, as specified in [MS-OXCDATA] section 2.2.4.1, that specifies the entry ID of the folder that contains the message or subfolder.");

            #endregion

            #region Step 6. The client creates rules on [MSOXCFOLDSubfolder1] folder.
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
            this.Adapter.DoRopCall(modifyRulesRequest, newFolderHandle, ref ropResponse, ref this.responseHandles);
            RopModifyRulesResponse modifyRulesResponse = (RopModifyRulesResponse)ropResponse;
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, modifyRulesResponse.ReturnValue, "RopModifyRules ROP operation performs successfully!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug,  @"Verify MS-OXCFOLD_R43");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R43.
            // The RopModifyRules ROP operation performs successfully indicates the input handle was a valid folder object handle. 
            // This input handle was get from the RopCreateFolder ROP response in step 1, and then MS-OXCFOLD_R43 can be verified directly.
            Site.CaptureRequirement(
                43,
                @"[InRopCreateFolder ROP Request Buffer] OutputHandleIndex (1 byte): The output Server object for this operation is a Folder object that represents the folder that was created.");
            #endregion

            #region Step 7. The client calls RopOpenFolder to open [MSOXCFOLDSubfolder1] folder.

            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderId = newFolderId,
                OpenModeFlags = (byte)FolderOpenModeFlags.None
            };
            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, newFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, openFolderResponse.ReturnValue, "RopOpenFolder ROP operation performs successfully!");

            newFolderHandle = this.responseHandles[0][openFolderResponse.OutputHandleIndex];
            openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, newFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, openFolderResponse.ReturnValue, "RopOpenFolder ROP operation performs successfully!");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, @"Verify MS-OXCFOLD_R16");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R16.
            // The RopOpenFolder ROP operation performs successfully indicates the input handle was a valid folder object handle. 
            // This input handle was get from the first RopOpenFolder ROP response in this step, and then MS-OXCFOLD_R16 can be verified directly.
            Site.CaptureRequirement(
                16,
                @"[In RopOpenFolder ROP Request Buffer] OutputHandleIndex (1 byte): The output Server object for this operation [RopOpenFolder ROP] is a Folder object that represents the folder that was opened.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R21");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R21.
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                openFolderResponse.ReturnValue,
                21,
                @"[In RopOpenFolder ROP Request Buffer] OpenModeFlags (1 byte): If this bit [OpenSoftDeleted] is not set, the operation opens an existing folder.<1>");

            if (Common.IsRequirementEnabled(25001, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R25001");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R25001
                Site.CaptureRequirementIfAreEqual<byte>(
                    0,
                    openFolderResponse.HasRules,
                    25001,
                    @"[In Appendix A: Product Behavior] If rules are associated with the folder, implementation does set a zero value to HasRules field. <2> Section 2.2.1.1.2: Exchange 2003 and Exchange 2007, Exchange 2016 and Exchange 2019 return zero (FALSE) in the HasRules field, even when there are rules on the Inbox folder.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R460");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R460
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(RopOpenFolderResponse),
                openFolderResponse.GetType(),
                460,
                @"[In Processing a RopOpenFolder ROP Request] The server responds with a RopOpenFolder ROP response buffer.");

            #region Verify the requirements: MS-OXCFOLD_R25002 and MS-OXCFOLD_R90001.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R90001");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R90001
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                openFolderResponse.ReturnValue,
                90001,
                @"[In RopOpenFolder ROP] The folder can be [either a public folder or] a private mailbox folder.");

            if (Common.IsRequirementEnabled(25002, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R25002");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R25002
                Site.CaptureRequirementIfAreNotEqual<byte>(
                    0x00,
                    openFolderResponse.HasRules,
                    25002,
                    @"[In Appendix A: Product Behavior] If rules are associated with the folder, implementation does set a nonzero value to HasRules field. (Microsoft Exchange Server 2010 follows this behavior.)");
            }

            #endregion

            #endregion

            #region Step 8. The client calls RopOpenFolder to open [MSOXCFOLDSubfolder2] folder.

            openFolderRequest.FolderId = newFolderId2;
            openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, newFolderHandle2, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, openFolderResponse.ReturnValue, "RopOpenFolder ROP operation performs successfully!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R26");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R26
            // According to the Open Specification, if rules are not associated with the folder, the HasRules field is set to zero (FALSE).
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                openFolderResponse.HasRules,
                26,
                @"[In RopOpenFolder ROP Response Buffer] HasRules (1 byte): otherwise [If rules are not associated with the folder], this field [HasRules] is set to zero (FALSE).");
            #endregion

            #region Step 9. The client calls RopDeleteFolder to softly delete [MSOXCFOLDSubfolder3] folder.

            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest
            {
                RopId = (byte)RopId.RopDeleteFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                DeleteFolderFlags = (byte)DeleteFolderFlags.DelFolders,
                FolderId = newFolderId3
            };
            RopDeleteFolderResponse deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, deleteFolderResponse.ReturnValue, "RopDeleteFolder ROP operation performs successfully.");
            Site.Assert.AreEqual<byte>(0x00, deleteFolderResponse.PartialCompletion, "The RopDeleteFolder ROP operation is completed");

            #endregion

            #region Step 10. The client calls RopOpenFolder to open [MSOXCFOLDSubfolder3] which was soft-deleted and open [MSOXCFOLDSubfolder2] which is an existing folder with OpenSoftDeleted flag.

            openFolderRequest.FolderId = newFolderId3;
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.OpenSoftDeleted;

            // Invoke the OpenFolder operation with OpenSoftDeleted set to open a soft-deleted folder.
            openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, newFolderHandle3, ref this.responseHandles);
            uint openSoftDeletedFolderReturnValue = openFolderResponse.ReturnValue;

            openFolderRequest.FolderId = newFolderId2;
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.OpenSoftDeleted;

            // Invoke the OpenFolder operation with OpenSoftDeleted set to open an existing folder.
            openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, newFolderHandle2, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, openFolderResponse.ReturnValue, "Client call RopOpenFolder ROP operation shoould be succeed.");
            uint openExistingFolderReturnValue = openFolderResponse.ReturnValue;
            uint softDeletedFolderHandle = this.responseHandles[0][openFolderResponse.OutputHandleIndex];

            #region Verify the requirement: MS-OXCFOLD_R498, MS-OXCFOLD_R22001, MS-OXCFOLD_R22003 and MS-OXCFOLD_R22004.

            if (Common.IsRequirementEnabled(22001, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(
                    LogEntryKind.Debug, 
                    @"Verify MS-OXCFOLD_R22001, 
                    The return value of the RopOpenFolder ROP response is {0}.",
                    openSoftDeletedFolderReturnValue);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R22001.
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    0,
                    openSoftDeletedFolderReturnValue,
                    22001,
                    @"[In Appendix A: Product Behavior] If OpenSoftDeleted bit in OpenModeFlags is set, implementation ignores the OpenSoftDeleted bit and always opens an existing folder in processing the RopOpenFolder ROP request. <1> Section 2.2.1.1.1: Exchange 2013 and Exchange 2016 ignores the OpenSoftDeleted bit and always opens an existing folder.");
            }

            if (Common.IsRequirementEnabled(22003, this.Site))
            {
                if (0x10 != deleteFolderRequest.DeleteFolderFlags)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R498");

                    // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R498.
                    // Based on DeleteHardDelete is not set, the folder here is soft deleted.
                    // Justify this point via RopOpenFolder response's ReturnValue is equal to 0x00000000.
                    Site.CaptureRequirementIfAreEqual<uint>(
                        Constants.SuccessCode,
                        openFolderResponse.ReturnValue,
                        498,
                        @"[In Processing a RopDeleteFolder ROP Request]If the DELETE_HARD_DELETE bit is not set, the folder becomes soft deleted.");
                }
            
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R22003, the return value of OpenSoftDeletedFolder is {0}.", openSoftDeletedFolderReturnValue);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R22003.
                Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)Constants.SuccessCode,
                    openSoftDeletedFolderReturnValue,
                    22003,
                    @"[In Appendix A: Product Behavior] If OpenSoftDeleted bit in OpenModeFlags is set, implementation does open either an existing folder or a soft-deleted folder in processing the RopOpenFolder ROP request. (Microsoft Exchange Server 2007 and Microsoft Exchange Server 2010 follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(22004, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R22004, the return value of OpenSoftDeletedFolder is {0}, the return value of OpenExistingFolder is {1}.", openSoftDeletedFolderReturnValue, openExistingFolderReturnValue);

                bool isVerifyR22004 =
                    openSoftDeletedFolderReturnValue == Constants.SuccessCode &&
                    openExistingFolderReturnValue == Constants.SuccessCode;

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R22004.
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR22004,
                    22004,
                    @"[In Appendix A: Product Behavior] Implementation provide access to soft-deleted folders. (Microsoft Exchange Server 2007 and Microsoft Exchange Server 2010 follow this behavior.)");

                #region The client calls RopGetPropertiesSpecific to get PidTagDeletedOn property.
                List<PropertyTag> propertyTags = new List<PropertyTag>();
                PropertyTag pidTagDeletedOn = new PropertyTag((ushort)FolderPropertyId.PidTagDeletedOn, (ushort)PropertyType.PtypTime);
                propertyTags.Add(pidTagDeletedOn);
                RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = this.GetSpecificProperties(softDeletedFolderHandle, propertyTags);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10348");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10348
                // If the property value in RopGetPropertiesSpecificResponse is not null, then the PidTagDeletedOn property is returned from server.
                // So R10348 will be verified.
                this.Site.CaptureRequirementIfIsTrue(
                    getPropertiesSpecificResponse.RowData.PropertyValues[0].Value.Length > 0,
                    10348,
                    @"[In PidTagDeletedOn Property] The PidTagDeletedOn property ([MS-OXPROPS] section 2.670) specifies the time when the folder was soft deleted.");
                #endregion
            }

            #endregion

            #endregion

            #region Step 11. The client calls RopOpenFolder to open the [MSOXCFOLDSubfolder2] with OpenModeFlags flag set with invalid value.
            openFolderRequest.FolderId = newFolderId2;
            byte[] invalidOpenModeFlags = { 0x01, 0x02, 0x08, 0x10, 0x20, 0x40, 0x80 };
            for (int i = 0; i < invalidOpenModeFlags.Length; i++)
            {
                openFolderRequest.OpenModeFlags = invalidOpenModeFlags[i];
                openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, newFolderHandle3, ref this.responseHandles);

                switch (invalidOpenModeFlags[i])
                {
                    case 0x01:
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R46501");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R46501
                        Site.CaptureRequirementIfAreEqual<uint>(
                            Constants.SuccessCode,
                            openFolderResponse.ReturnValue,
                            46501,
                            @"[In Processing a RopOpenFolder ROP Request] The server MUST ignore invalid bit ""0x01"" that is set in the OpenModeFlags field.");
                        break;
                    case 0x02:
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R46502");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R46502
                        Site.CaptureRequirementIfAreEqual<uint>(
                            Constants.SuccessCode,
                            openFolderResponse.ReturnValue,
                            46502,
                            @"[In Processing a RopOpenFolder ROP Request] The server MUST ignore invalid bit ""0x02"" that is set in the OpenModeFlags field.");
                        break;
                    case 0x08:
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R46503");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R46503
                        Site.CaptureRequirementIfAreEqual<uint>(
                            Constants.SuccessCode,
                            openFolderResponse.ReturnValue,
                            46503,
                            @"[In Processing a RopOpenFolder ROP Request] The server MUST ignore invalid bit ""0x08"" that is set in the OpenModeFlags field.");
                        break;
                    case 0x10:
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R46504");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R46504
                        Site.CaptureRequirementIfAreEqual<uint>(
                            Constants.SuccessCode,
                            openFolderResponse.ReturnValue,
                            46504,
                            @"[In Processing a RopOpenFolder ROP Request] The server MUST ignore invalid bit ""0x10"" that is set in the OpenModeFlags field.");
                        break;
                    case 0x20:
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R46505");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R46505
                        Site.CaptureRequirementIfAreEqual<uint>(
                            Constants.SuccessCode,
                            openFolderResponse.ReturnValue,
                            46505,
                            @"[In Processing a RopOpenFolder ROP Request] The server MUST ignore invalid bit ""0x20"" that is set in the OpenModeFlags field.");
                        break;
                    case 0x40:
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R46506");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R46506
                        Site.CaptureRequirementIfAreEqual<uint>(
                            Constants.SuccessCode,
                            openFolderResponse.ReturnValue,
                            46506,
                            @"[In Processing a RopOpenFolder ROP Request] The server MUST ignore invalid bit ""0x40"" that is set in the OpenModeFlags field.");
                        break;
                    case 0x80:
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R46507");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R46507
                        Site.CaptureRequirementIfAreEqual<uint>(
                            Constants.SuccessCode,
                            openFolderResponse.ReturnValue,
                            46507,
                            @"[In Processing a RopOpenFolder ROP Request] The server MUST ignore invalid bit ""0x80"" that is set in the OpenModeFlags field.");
                        break;
                }
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopMoveFolder operation responds with error codes.  
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC02_RopMoveFolderFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.

            uint subFolderHandle1 = 0;
            ulong subFolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subFolderId1, ref subFolderHandle1);

            #endregion

            #region Step 2. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] under the root folder.

            uint subFolderHandle2 = 0;
            ulong subFolderId2 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder2, ref subFolderId2, ref subFolderHandle2);

            #endregion

            #region Step 3. The client creates a message in the root folder.

            uint messageHandleInRootFolder = 0;
            ulong messageIdInRootFolder = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref messageIdInRootFolder, ref messageHandleInRootFolder);

            #endregion

            #region Step 4. The client calls RopMoveFolder to move a non-exist folder from the root folder to [MSOXCFOLDSubfolder1] folder.

            // Initialize a server object handle table.
            List<uint> handleList = new List<uint>
            {
                // Add the root folder handle to the server object handle table, and its index value is 0x00.
                // Add the Subfolder1 handle to the server object handle table, and its index value is  0x01.
                this.RootFolderHandle, subFolderHandle1
            };
            
            RopMoveFolderRequest moveFolderRequest = new RopMoveFolderRequest
            {
                RopId = (byte)RopId.RopMoveFolder,
                LogonId = Constants.CommonLogonId,
                SourceHandleIndex = 0x00,
                DestHandleIndex = 0x01,
                WantAsynchronous = 0x00,
                UseUnicode = 0x00,
                FolderId = long.MaxValue,
                NewFolderName = Encoding.ASCII.GetBytes(Constants.Subfolder2)
            };

            // Set a non-exist folder ID in order to make server return an error code "ecNotFound".
            RopMoveFolderResponse moveFolderResponse = this.Adapter.MoveFolder(moveFolderRequest, handleList, ref this.responseHandles);
            handleList.Clear();

            #region Verify the requirements: MS-OXCFOLD_R599, MS-OXCFOLD_R600.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R599");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R599
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                moveFolderResponse.ReturnValue,
                599,
                @"[In Processing a RopMoveFolder ROP Request]The value of error code ecNotFound is 0x8004010F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R600");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R600
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                moveFolderResponse.ReturnValue,
                600,
                @"[In Processing a RopMoveFolder ROP Request] When the error code is ecNotFound, it indicates that there is no folder with the specified ID.");

            #endregion

            #endregion

            #region Step 5. The client calls RopMoveFolder to move a message from the root folder to [MSOXCFOLDSubfolder1] folder.

            // The parameter "messageHandleInRootFolder" stores a message object handle to refer to a message object, add this parameter to the server object handle table is purposed to test error code ecNotSupported [0x80040102].  
            // Add the message handle in the root folder to the server object handle table, and its index value is 0x00.
            handleList.Add(messageHandleInRootFolder);

            // Add the Subfolder1 handle to the server object handle table, and its index value is  0x01.
            handleList.Add(subFolderHandle1);

            moveFolderRequest.RopId = (byte)RopId.RopMoveFolder;
            moveFolderRequest.LogonId = Constants.CommonLogonId;
            moveFolderRequest.SourceHandleIndex = 0x00;
            moveFolderRequest.DestHandleIndex = 0x01;
            moveFolderRequest.WantAsynchronous = 0x00;
            moveFolderRequest.UseUnicode = 0x00;
            moveFolderRequest.FolderId = messageIdInRootFolder;
            moveFolderRequest.NewFolderName = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            moveFolderResponse = this.Adapter.MoveFolder(moveFolderRequest, handleList, ref this.responseHandles);
            handleList.Clear();

            #region Verify the requirements: MS-OXCFOLD_R601, MS-OXCFOLD_R602.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R601");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R601.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                moveFolderResponse.ReturnValue,
                601,
                @"[In Processing a RopMoveFolder ROP Request]The value of error code ecNotSupported is 0x80040102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R602");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R602
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                moveFolderResponse.ReturnValue,
                602,
                @"[In Processing a RopMoveFolder ROP Request] When the error code is ecNotSupported, it indicates that either the source object or the destination object is not a Folder object.");

            #endregion

            #endregion

            #region Step 6. The client calls RopMoveFolder to move target folder failed by bad format.

            // Add the source folder handle to the server object handle table, and its index value is 0x00.
            handleList.Add(this.RootFolderHandle);

            // Add the destination folder handle to the server object handle table, and its index value is 0x01.
            handleList.Add(subFolderHandle2);

            moveFolderRequest = new RopMoveFolderRequest();
            object ropResponse = new object();
            moveFolderRequest.RopId = (byte)RopId.RopMoveFolder;
            moveFolderRequest.LogonId = Constants.CommonLogonId;
            moveFolderRequest.SourceHandleIndex = 0x00;
            moveFolderRequest.DestHandleIndex = 0x01;
            moveFolderRequest.WantAsynchronous = 0x00;
            moveFolderRequest.UseUnicode = 0x00;
            moveFolderRequest.FolderId = subFolderId1;
            moveFolderRequest.NewFolderName = Encoding.Unicode.GetBytes(Constants.Subfolder3);
            FormatException formatException = null;

            try
            {
                this.Adapter.DoRopCall(moveFolderRequest, handleList, ref ropResponse, ref this.responseHandles);
            }
            catch (FormatException e)
            {
                formatException = e;
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R184");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R184.
            // The program throw a format exception indicates the value of the NewFolderName field is not formatted in Unicode.
            Site.CaptureRequirementIfIsNotNull(
                formatException,
                184,
                @"[In RopMoveFolder ROP Request Buffer] UseUnicode (1 byte): it [UseUnicode] is zero (FALSE) otherwise [if the value of the NewFolderName field is not formatted in Unicode].");

            handleList.Clear();

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopCreateFolder operation performs successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC03_RopCreateFolderSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder with 'OpeningExisting' flag set to zero and without setting the Comment flag.

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
                Comment = Encoding.ASCII.GetBytes(Constants.StringNullTerminated)
            };

            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");

            ulong folderId = createFolderResponse.FolderId;

            #region Verify the requirements: MS-OXCFOLD_R3801 and MS-OXCFOLD_R478.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R3801");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R3801
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                createFolderResponse.ReturnValue,
                3801,
                @"[In RopCreateFolder ROP] The folder can be [either a public folder or] a private mailbox folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R478");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R478
            // The client calls RopCreateFolder operation without setting the Comment field and the server return a success code replace with an error code, so this requirement can be verified.
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                createFolderResponse.ReturnValue,
                478,
                @"[In Processing a RopCreateFolder ROP Request] A folder description, specified in the Comment field of the ROP request buffer, is optional.");
            #endregion

            #endregion

            #region Step 2. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder with 'OpeningExisting' flag set to non-zero.

            // Marked the 'OpeningExisting' as true.
            createFolderRequest.OpenExisting = 0x01;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder1);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder1);
            RopCreateFolderResponse createExistingFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createExistingFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");

            #region Verify the requirements: MS-OXCFOLD_R482, MS-OXCFOLD_R50 and MS-OXCFOLD_R1197.
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1197, regardless of the existence of the folder with the same name, the IsExistingFolder in response should be always set to {0}, actually, when folder with the same name exists the IsExistingFolder in response is {1}, when folder with the same name does not exist the IsExistingFolder in response is {2}.", 0, createExistingFolderResponse.IsExistingFolder, createFolderResponse.IsExistingFolder);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1197
            // The IsExistingFolder always be set to zero by the server regardless of the existence of the same name folder exists or not. 
            bool isVerifyR1197 = createFolderResponse.IsExistingFolder == 0 && createExistingFolderResponse.IsExistingFolder == 0;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1197,
                1197,
                @"[RopCreateFolder ROP Response Buffer] The server always sets this field [IsExistingFolder] to zero for a folder that is created in a private mailbox.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R482, the FolderId in this RopCreateFolder response is {0}, and the folder ID of the folder which has already existed is {1}.", createFolderResponse.FolderId, folderId);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R482
            // The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] which has been created previous under the root folder with 'OpeningExisting' flag set to non-zero. So if the FolderId in the response is equal to the folder ID of the MSOXCFOLDSubfolder1, R482 can be verified. 
            Site.CaptureRequirementIfAreEqual<ulong>(
                folderId,
                createExistingFolderResponse.FolderId,
                482,
                @"[In Processing a RopCreateFolder ROP Request] If a folder with the same name already exists and the OpenExisting field is set to nonzero (TRUE), the server opens the existing folder, behaving as if it is processing the RopOpenFolder ROP ([MS-OXCROPS] section 2.2.4.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R50, the FolderId in this RopCreateFolder response is {0}, and the folder ID of the folder which has already existed is {1}.", createFolderResponse.FolderId, folderId);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R50
            // The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] which has been created previous under the root folder with 'OpeningExisting' flag set to non-zero. So if the FolderId in the response is equal to the folder ID of the MSOXCFOLDSubfolder1, R50 can be verified.            
            Site.CaptureRequirementIfAreEqual<ulong>(
                folderId,
                createExistingFolderResponse.FolderId,
                50,
                @"[InRopCreateFolder ROP Request Buffer] OpenExisting (1 byte): A Boolean value that is nonzero (TRUE) if a pre-existing folder, whose name is identical to the name specified in the DisplayName field, is to be opened.");

            #endregion
            #endregion

            #region Step 3. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] under the root folder.

            createFolderRequest.UseUnicodeStrings = 0x01;
            createFolderRequest.OpenExisting = 0x01;
            createFolderRequest.DisplayName = Encoding.Unicode.GetBytes(Constants.Subfolder2);
            createFolderRequest.Comment = Encoding.Unicode.GetBytes(Constants.Subfolder2);
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint folderHandle2 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 4. The client gets the properties from server.

            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest();
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;
            getPropertiesSpecificRequest.RopId = (byte)RopId.RopGetPropertiesSpecific;
            getPropertiesSpecificRequest.LogonId = Constants.CommonLogonId;
            getPropertiesSpecificRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            getPropertiesSpecificRequest.PropertySizeLimit = 0xFFFF;

            PropertyTag[] tags = new PropertyTag[3];
            PropertyTag tag;

            // Get the property: PidTagDisplayName.
            tag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagDisplayName,
                PropertyType = (ushort)PropertyType.PtypString
            };
            tags[0] = tag;

            // Get the property: PidTagComment.
            tag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagComment,
                PropertyType = (ushort)PropertyType.PtypString
            };
            tags[1] = tag;

            // Get the property: PidTagMessageSizeExtended.
            tag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagMessageSizeExtended,
                PropertyType = (ushort)PropertyType.PtypInteger64
            };
            tags[2] = tag;

            getPropertiesSpecificRequest.PropertyTagCount = (ushort)tags.Length;
            getPropertiesSpecificRequest.PropertyTags = tags;
            getPropertiesSpecificResponse = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, folderHandle2, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getPropertiesSpecificResponse.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");

            string pidTagDisplayNameValue = System.Text.UnicodeEncoding.Unicode.GetString(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value);
            string pidTagCommentValue = System.Text.UnicodeEncoding.Unicode.GetString(getPropertiesSpecificResponse.RowData.PropertyValues[1].Value);
            ulong pidTagMessageSizeExtended = BitConverter.ToUInt64(getPropertiesSpecificResponse.RowData.PropertyValues[2].Value, 0);

            #region Verify the requirement: MS-OXCFOLD_R477, MS-OXCFOLD_R48, MS-OXCFOLD_R54 and MS-OXCFOLD_R56.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R477");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R477
            Site.CaptureRequirementIfAreEqual<string>(
                Constants.Subfolder2,
                pidTagDisplayNameValue,
                477,
                @"[In Processing a RopCreateFolder ROP Request] A folder name in the DisplayName field of the ROP request buffer, as specified in section 2.2.1.2.1, MUST be specified to create a folder.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                @"Verify MS-OXCFOLD_R48, the expected value of folder display name is {0}, the actual value of folder display name parsed use unicode is {1}; the expected value of folder comment is {2}, the actual value of folder comment parsed use unicode is {3};",
                Constants.Subfolder2.Trim(Constants.StringNullTerminated.ToCharArray()),
                pidTagDisplayNameValue.Trim(Constants.StringNullTerminated.ToCharArray()),
                Constants.Subfolder2.Trim(Constants.StringNullTerminated.ToCharArray()),
                pidTagCommentValue.Trim(Constants.StringNullTerminated.ToCharArray()));

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R48
            bool isVerifyR48 = Constants.Subfolder2 == pidTagDisplayNameValue && Constants.Subfolder2 == pidTagCommentValue;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR48,
                48,
                @"[InRopCreateFolder ROP Request Buffer] UseUnicodeStrings (1 byte): A Boolean value that is nonzero (TRUE) if the values of the DisplayName and Comment fields are formatted in Unicode.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R54");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R54
            Site.CaptureRequirementIfAreEqual<string>(
                Constants.Subfolder2,
                pidTagDisplayNameValue,
                54,
                @"[In RopCreateFolder ROP Request Buffer] DisplayName (variable): This name [DisplayName] becomes the value of the new folder's PidTagDisplayName property (section 2.2.2.2.2.5).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R56");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R56
            Site.CaptureRequirementIfAreEqual<string>(
                Constants.Subfolder2,
                pidTagCommentValue,
                56,
                @"[InRopCreateFolder ROP Request Buffer] Comment (variable): This string [Comment] becomes the value of the new folder's PidTagComment property (section 2.2.2.2.2.2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10345");
        
            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10345
            this.Site.CaptureRequirementIfAreEqual<ulong>(
                0,
                pidTagMessageSizeExtended,
                10354,
                @"[In PidTagMessageSizeExtended Property] The PidTagMessageSizeExtended property ([MS-OXPROPS] section 2.797) specifies the 64-bit version of the PidTagMessageSize property (section 2.2.2.2.1.7).");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopCopyFolder operation responds with error codes. 
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC04_RopCopyFolderFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] folder in the root folder.

            uint subFolderHandle1 = 0;
            ulong subFolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subFolderId1, ref subFolderHandle1);

            #endregion

            #region Step 2. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] folder in the root folder.

            uint subFolderHandle2 = 0;
            ulong subFolderId2 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder2, ref subFolderId2, ref subFolderHandle2);

            #endregion

            #region Step 3. The client saves a message in the root folder.

            uint messageHandleInRootFolder = 0;
            ulong messageIdInRootFolder = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref messageIdInRootFolder, ref messageHandleInRootFolder);

            #endregion

            #region Step 4. The client calls RopCopyFolder to copy a non-existing folder from the root folder to [MSOXCFOLDSubfolder1] folder.

            // Initialize a server object handle table.
            List<uint> handleList = new List<uint>
            {
                // Add the root folder handle to the server object handle table, and its index value is 0x00.
                // Add the Subfolder1 handle to the server object handle table, and its index value is 0x01.
                this.RootFolderHandle, subFolderHandle1
            };
            
            // Call RopCopyFolder operation with non-exist folder Id.
            RopCopyFolderRequest copyFolderRequest = new RopCopyFolderRequest
            {
                RopId = (byte)RopId.RopCopyFolder,
                LogonId = Constants.CommonLogonId,
                SourceHandleIndex = 0x00,
                DestHandleIndex = 0x01,
                WantAsynchronous = 0x00,
                UseUnicode = 0x00,
                WantRecursive = 0xFF,
                FolderId = long.MaxValue,
                NewFolderName = Encoding.ASCII.GetBytes(Constants.Subfolder2)
            };
            RopCopyFolderResponse copyFolderResponse = this.Adapter.CopyFolder(copyFolderRequest, handleList, ref this.responseHandles);
            handleList.Clear();

            #region Verify the requirements: MS-OXCFOLD_R611, MS-OXCFOLD_R612.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R611");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R611
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                copyFolderResponse.ReturnValue,
                611,
                @"[In Processing a RopCopyFolder ROP Request]The value of error code ecNotFound is 0x8004010F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R612");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R612
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                copyFolderResponse.ReturnValue,
                612,
                @"[In Processing a RopCopyFolder ROP Request] When the error code is ecNotFound, it indicates that there is no folder with the specified ID.");

            #endregion

            #endregion

            #region Step 5. The client calls RopCopyFolder to copy a message from the root folder to [MSOXCFOLDSubfolder1] folder.

            // Add the message handle to the server object handle table, and its index value is 0x00.
            handleList.Add(messageHandleInRootFolder);

            // Add the Subfolder1 handle to the server object handle table, and its index value is 0x01.
            handleList.Add(subFolderHandle1);

            // Call RopCopyFolder operation with a message object Id.
            copyFolderRequest.RopId = (byte)RopId.RopCopyFolder;
            copyFolderRequest.LogonId = Constants.CommonLogonId;
            copyFolderRequest.SourceHandleIndex = 0x00;
            copyFolderRequest.DestHandleIndex = 0x01;
            copyFolderRequest.WantAsynchronous = 0x00;
            copyFolderRequest.UseUnicode = 0x00;
            copyFolderRequest.WantRecursive = 0xFF;
            copyFolderRequest.FolderId = messageHandleInRootFolder;
            copyFolderRequest.NewFolderName = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            RopCopyFolderResponse copyFolderResponseNoFolderSource = this.Adapter.CopyFolder(copyFolderRequest, handleList, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0x80040102, copyFolderResponseNoFolderSource.ReturnValue, "The returnValue should be 0x80040102 when the source object is not a Folder object.");
            handleList.Clear();
            #endregion

            #region Step 6. The client calls RopCopyFolder to copy the [MSOXCFOLDSubfolder1] folder to a message.
            // Add the Subfolder1 handle to the server object handle table, and its index value is 0x00.
            handleList.Add(subFolderHandle1);

            // Add the message handle to the server object handle table, and its index value is 0x01.
            handleList.Add(messageHandleInRootFolder);

            // Call RopCopyFolder operation with a message object Id.
            copyFolderRequest.RopId = (byte)RopId.RopCopyFolder;
            copyFolderRequest.LogonId = Constants.CommonLogonId;
            copyFolderRequest.SourceHandleIndex = 0x00;
            copyFolderRequest.DestHandleIndex = 0x01;
            copyFolderRequest.WantAsynchronous = 0x00;
            copyFolderRequest.UseUnicode = 0x00;
            copyFolderRequest.WantRecursive = 0xFF;
            copyFolderRequest.FolderId = messageHandleInRootFolder;
            copyFolderRequest.NewFolderName = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            RopCopyFolderResponse copyFolderResponseNoFolderDest = this.Adapter.CopyFolder(copyFolderRequest, handleList, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0x80040102, copyFolderResponseNoFolderSource.ReturnValue, "The returnValue should be 0x80040102 when the source object is not a Folder object.");
            handleList.Clear();

            #region Verify the requirements: MS-OXCFOLD_R613, MS-OXCFOLD_R614.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R613");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R613
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                copyFolderResponseNoFolderDest.ReturnValue,
                613,
                @"[In Processing a RopCopyFolder ROP Request]The value of error code ecNotSupported is 0x80040102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R614");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R614
            // According to step above, the returnValue is 0x80040102, when the source object or the destination object is not a Folder object.
            // So MS-OXCFOLD_R614 will be verified.
            Site.CaptureRequirement(
                614,
                @"[In Processing a RopCopyFolder ROP Request] When the error code is ecNotSupported, it indicates that either the source object or the destination object is not a Folder object.");

            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopDeleteFolder operation with an invalid DeleteFolderFlags.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC05_RopDeleteFolderWithInvalidDeleteFolderFlags()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.

            // Create the subfolder under the root folder.
            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);

            #endregion

            #region Step 2. The client calls RopDeleteFolder to soft-delete [MSOXCFOLDSubfolder1] and sets the 'DeleteFolderFlags' flag with invalid bits.

            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest();
            RopDeleteFolderResponse deleteFolderResponse;
            deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
            deleteFolderRequest.LogonId = Constants.CommonLogonId;
            deleteFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;

            if (Common.IsRequirementEnabled(123402, this.Site))
            {
                byte[] invalidDeleteFolderFlags = { 0x02, 0x08, 0x20, 0x40, 0x80 };
                foreach (byte invalidDeleteFolderFlag in invalidDeleteFolderFlags)
                {
                    deleteFolderRequest.FolderId = subfolderId1;
                    deleteFolderRequest.DeleteFolderFlags = invalidDeleteFolderFlag;
                    deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

                    Site.Assert.AreEqual<uint>(
                        Constants.SuccessCode,
                        deleteFolderResponse.ReturnValue,
                        @"[In Processing a RopDeleteFolder ROP Request] The server MUST ignore invalid bit {0} that is set in the DeleteFolderFlags field of the ROP request buffer, the ROP performs successfully.",
                        invalidDeleteFolderFlag);

                    Site.Assert.AreEqual<byte>(
                        0,
                        deleteFolderResponse.PartialCompletion,
                        @"[In Processing a RopDeleteFolder ROP Request] The server MUST ignore invalid bit {0} that is set in the DeleteFolderFlags field of the ROP request buffer, the sub target delete successfully.",
                        invalidDeleteFolderFlag);

                    // Recreate the [MSOXCFOLDSubfolder1] under the root folder.
                    this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R123402");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R123402.
                // [In Processing a RopDeleteFolder ROP Request] The server ignored the invalid bits [0x02, 0x08, 0x20, 0x40, and 0x80], MS-OXCFOLD_R103402 can be verified.
                Site.CaptureRequirement(
                    123402,
                    @"[In Appendix A: Product Behavior] Implementation does ignore invalid bits instead of failing the ROP [RopDeleteFolder], if the client sets an invalid bit in the DeleteFolderFlags field of the ROP request buffer. <15> Section 3.2.5.3:  Exchange 2010 and later ignore invalid bits instead of failing the ROP.");
            }

            if (Common.IsRequirementEnabled(123401, this.Site))
            {
                byte invalidDeleteFolderFlag = 0x02;
                deleteFolderRequest.FolderId = subfolderId1;
                deleteFolderRequest.DeleteFolderFlags = invalidDeleteFolderFlag;
                deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R123401");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R123401.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    deleteFolderResponse.ReturnValue,
                    123401,
                    @"[In Appendix A: Product Behavior] Implementation does fail the ROP [RopDeleteFolder] with an ecInvalidParam (0x80070057) error, if the client sets an invalid bit in the DeleteFolderFlags field of the ROP request buffer. (Exchange 2007 follows this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopDeleteFolder operation to delete folder which contains a message.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC06_RopDeleteFolderWithDelMessages()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.

            // Create the subfolder under the root folder.
            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);

            #endregion

            #region Step 2. The client saves a message in [MSOXCFOLDSubfolder1].

            // Create Message in the [MSOXCFOLDSubfolder1] subfolder.
            ulong messageId = 0;
            uint messageHandle = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageId, ref messageHandle);

            #endregion

            #region Step 3. The client calls RopDeleteFolder to hard-delete [MSOXCFOLDSubfolder1] which contains a message.
            RopDeleteFolderRequest deleteFolderRequest;
            deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
            deleteFolderRequest.LogonId = Constants.CommonLogonId;
            deleteFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DeleteHardDelete;
            deleteFolderRequest.FolderId = subfolderId1;
            RopDeleteFolderResponse deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(
                0,
                deleteFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            #region Step 4. The client calls RopOpenFolder to open [MSOXCFOLDSubfolder1] under the root folder.

            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderId = subfolderId1,
                OpenModeFlags = (byte)FolderOpenModeFlags.OpenSoftDeleted
            };
            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(
                0,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            #region Verify the requirements: MS-OXCFOLD_R816, MS-OXCFOLD_R817 and MS-OXCFOLD_R371.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R816, the return value of DeleteFolder is {0}, the return value of OpenFolder is {1}.", deleteFolderResponse.ReturnValue, openFolderResponse.ReturnValue);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R816
            bool isVerifyR816 = deleteFolderResponse.ReturnValue == Constants.SuccessCode && openFolderResponse.ReturnValue == Constants.SuccessCode;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR816,
                816,
                @"[In Processing a RopDeleteFolder ROP Request] [If the DEL_MESSAGES bit is not set and the folder contains Message objects, neither the folder nor any of its Message objects will be deleted], the ReturnValue field of the ROP response, as specified in section 2.2.1.3.2, will be set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R817, the PartialCompletion field of DeleteFolder is {0}, the return value of OpenFolder is {1}.", deleteFolderResponse.PartialCompletion, openFolderResponse.ReturnValue);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R817
            bool isVerifyR817 = deleteFolderResponse.PartialCompletion != 0x00 && openFolderResponse.ReturnValue == Constants.SuccessCode;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR817,
                817,
                @"[In Processing a RopDeleteFolder ROP Request] [If the DEL_MESSAGES bit is not set and the folder contains Message objects, neither the folder nor any of its Message objects will be deleted], the PartialCompletion field will be set to a nonzero (TRUE) value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R371");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R371
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                openFolderResponse.ReturnValue,
                371,
                @"[In Processing a RopDeleteFolder ROP Request] If the DEL_MESSAGES bit is not set and the folder contains Message objects, neither the folder nor any of its Message objects will be deleted.");

            #endregion

            #endregion

            #region Step 5. The client calls RopDeleteFolder to delete [MSOXCFOLDSubfolder1] which contains a message with 'DeleteFolderFlags' was set to DelMessages.

            deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DelMessages | (byte)DeleteFolderFlags.DeleteHardDelete;
            deleteFolderRequest.FolderId = subfolderId1;
            deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(
                0,
                deleteFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            #region Verify the requirement: MS-OXCFOLD_R762 and MS-OXCFOLD_R82.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R762");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R762
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                deleteFolderResponse.PartialCompletion,
                762,
                @"[In RopDeleteFolder ROP Request Buffer] DeleteFolderFlags (1 byte): DEL_MESSAGES (0x01) means that the folder and all of the Message objects in the folder are deleted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R82");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R82.
            // MS-OXCFOLD_R762 and MS-OXCFOLD_R763 were verified the RopDeleteFolder ROP operates on non-empty folders, MS-OXCFOLD_R82 can be verified directly.
            Site.CaptureRequirement(
                82,
                @"[In RopDeleteFolder ROP Request Buffer] DeleteFolderFlags (1 byte): By default, the RopDeleteFolder ROP operates only on empty folders, but it [RopDeleteFolder ROP] can be used successfully on non-empty folders by setting the DEL_FOLDERS bit and the DEL_MESSAGES bit.");

            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopDeleteFolder operation to delete folder which contains a folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC07_RopDeleteFolderWithDelFolders()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.

            // Create the subfolder under the root folder.
            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);

            #endregion

            #region Step 2. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] under [MSOXCFOLDSubfolder1].

            // Create the subfolder under the [MSOXCFOLDSubfolder1] folder.
            uint subfolderHandle2 = 0;
            ulong subfolderId2 = 0;
            this.CreateFolder(subfolderHandle1, Constants.Subfolder2, ref subfolderId2, ref subfolderHandle2);

            #endregion

            #region Step 3. The client calls RopDeleteFolder to hard-delete [MSOXCFOLDSubfolder1] which contains a subfolder without setting DEL_FOLDERS.
            RopDeleteFolderRequest deleteFolderRequest;
            deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
            deleteFolderRequest.LogonId = Constants.CommonLogonId;
            deleteFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DeleteHardDelete;
            deleteFolderRequest.FolderId = subfolderId1;
            RopDeleteFolderResponse deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(
                0,
                deleteFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            #region Step 4. The client calls RopOpenFolder to open [MSOXCFOLDSubfolder1] under the root folder.
            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest();
            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = Constants.CommonLogonId;
            openFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            openFolderRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            openFolderRequest.FolderId = subfolderId1;
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Invoke the OpenFolder operation with OpenSoftDeleted set to open a hard-deleted folder.
            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(
                0,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            #region Verify the requirements: MS-OXCFOLD_R1080, MS-OXCFOLD_R1081, MS-OXCFOLD_R1079 and MS-OXCFOLD_R948.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1080, the return value of DeleteFolder is {0}, the return value of OpenFolder is {1}.", deleteFolderResponse.ReturnValue, openFolderResponse.ReturnValue);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1080
            bool isVerifyR1080 = deleteFolderResponse.ReturnValue == Constants.SuccessCode && openFolderResponse.ReturnValue == Constants.SuccessCode;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1080,
                1080,
                @"[In Processing a RopDeleteFolder ROP Request] [If the DEL_FOLDERS bit is not set and the folder contains subfolders, neither the folder nor any of its subfolders will be deleted], the ReturnValue field of the ROP response, as specified in section 2.2.1.3.2, will be set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1081, the PartialCompletion field of DeleteFolder is {0}, the return value of OpenFolder is {1}.", deleteFolderResponse.PartialCompletion, openFolderResponse.ReturnValue);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1081
            bool isVerifyR1081 = deleteFolderResponse.PartialCompletion != 0x00 && openFolderResponse.ReturnValue == Constants.SuccessCode;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1081,
                1081,
                @"[In Processing a RopDeleteFolder ROP Request] And [if the DEL_FOLDERS bit is not set and the folder contains subfolders, neither the folder nor any of its subfolders will be deleted], the PartialCompletion field will be set to a nonzero (TRUE) value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1079");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1079
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                openFolderResponse.ReturnValue,
                1079,
                @"[In Processing a RopDeleteFolder ROP Request] If the DEL_FOLDERS bit is not set and the folder contains subfolders, neither the folder nor any of its subfolders will be deleted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R948");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R948
            Site.CaptureRequirementIfAreNotEqual<byte>(
                0x00,
                deleteFolderResponse.PartialCompletion,
                948,
                @"[In RopDeleteFolder ROP Response Buffer] If the ROP fails for a subset of targets, the value of this field [PartialCompletion] is nonzero (TRUE).");

            #endregion
            #endregion

            #region Step 5. The client calls RopDeleteFolder to delete [MSOXCFOLDSubfolder1] which contains an empty folder.

            deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
            deleteFolderRequest.LogonId = Constants.CommonLogonId;
            deleteFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.None;
            deleteFolderRequest.FolderId = subfolderId1;
            deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(
                0,
                deleteFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");
            #endregion

            #region Step 6. The client calls RopDeleteFolder to delete [MSOXCFOLDSubfolder1] which contains an empty folder with 'DeleteFolderFlags' was set to DelFolders.

            deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DelFolders | (byte)DeleteFolderFlags.DeleteHardDelete;
            deleteFolderRequest.FolderId = subfolderId1;
            deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(
                0,
                deleteFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            byte flagPartialCompletionInResponse3 = deleteFolderResponse.PartialCompletion;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R763");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R763
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                flagPartialCompletionInResponse3,
                763,
                @"[In RopDeleteFolder ROP Request Buffer] DeleteFolderFlags (1 byte): DEL_FOLDERS (0x04) means that the folder and all of its subfolders are deleted.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopDeleteFolder operation to hard delete folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC08_RopDeleteFolderToHardDelete()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.

            // Create the subfolder under the root folder.
            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);

            #endregion

            #region Step 2. The client calls RopDeleteFolder to hard-delete [MSOXCFOLDSubfolder1] which is an empty folder.
            RopDeleteFolderRequest deleteFolderRequest;
            deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
            deleteFolderRequest.LogonId = Constants.CommonLogonId;
            deleteFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DeleteHardDelete | (byte)DeleteFolderFlags.DelFolders | (byte)DeleteFolderFlags.DelMessages;
            deleteFolderRequest.FolderId = subfolderId1;
            RopDeleteFolderResponse deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(
                0,
                deleteFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            #region Step 3. The client calls RopOpenFolder to open [MSOXCFOLDSubfolder1] which was hard-deleted in step 2.

            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderId = subfolderId1,
                OpenModeFlags = (byte)FolderOpenModeFlags.OpenSoftDeleted
            };

            // Invoke the OpenFolder operation with OpenSoftDeleted set to open a hard-deleted folder.
            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            #region Verify the requirement: MS-OXCFOLD_R497, MS-OXCFOLD_R97,MS-OXCFOLD_R46201001 and MS-OXCFOLD_R46201002 .

            if (Common.IsRequirementEnabled(46201001, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R46201001");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R46201001
                Site.CaptureRequirementIfAreEqual<uint>(
                    Constants.SuccessCode,
                    openFolderResponse.ReturnValue,
                    46201001,
                    @"[In Appendix A: Product Behavior] If the specified folder has been hard deleted, implementation does not fail the RopOpenFolder ROP ([MS-OXCROPS] section 2.2.4.1), but no folder is opened. <12> Section 3.2.5.1: If the specified folder has been hard deleted, Exchange 2007 does not fail the RopOpenFolder ROP ([MS-OXCROPS] section 2.2.4.1), but no folder is opened.");
            }
       
            if (Common.IsRequirementEnabled(46201002, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R497");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R497
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    Constants.SuccessCode,
                    openFolderResponse.ReturnValue,
                    497,
                    @"[In Processing a RopDeleteFolder ROP Request] If the DELETE_HARD_DELETE bit of the DeleteFolderFlags field of the ROP request buffer is set, as specified in section 2.2.1.3.1, the folder MUST be removed and can no longer be accessed by the client with subsequent ROPs.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R46201002");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R46201002
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    Constants.SuccessCode,
                    openFolderResponse.ReturnValue,
                    46201002,
                    @"[In Appendix A: Product Behavior] If the specified folder has been hard deleted, implementation does fail the RopOpenFolder ROP ([MS-OXCROPS] section 2.2.4.1), but no folder is opened. (Microsoft Exchange Server 2010 and above follow this behavior.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R97");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R97
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                deleteFolderResponse.PartialCompletion,
                97,
                @"[In RopDeleteFolder ROP Response Buffer] PartialCompletion (1 byte): otherwise [if the ROP successes for a subset of targets], the value is zero (FALSE).");

            if (Common.IsRequirementEnabled(764, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R764");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R764
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    Constants.SuccessCode,
                    openFolderResponse.ReturnValue,
                    764,
                    @"[In Appendix A: Product Behavior] If this bit [DELETE_HARD_DELETE] is set, implement does hard delete the folder. &lt;4&gt; Section 2.2.1.3.1:  For Exchange 2003 and later, if DELETE_HARD_DELETE is set, the folder is hard deleted.");
            }
            #endregion

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopDeleteFolder operation to soft delete folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC09_RopDeleteFolderToSoftDelete()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.

            // Create the subfolder under the root folder.
            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);

            #endregion

            #region Step 2. The client calls RopDeleteFolder to soft-delete [MSOXCFOLDSubfolder1] under the root folder.
            RopDeleteFolderRequest deleteFolderRequest;
            deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
            deleteFolderRequest.LogonId = Constants.CommonLogonId;
            deleteFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DelFolders;
            deleteFolderRequest.FolderId = subfolderId1;
            RopDeleteFolderResponse deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(
                0,
                deleteFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            #region Step 3. The client calls RopOpenFolder to open [MSOXCFOLDSubfolder1] which was soft-deleted in step 2.

            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderId = subfolderId1,
                OpenModeFlags = (byte)FolderOpenModeFlags.OpenSoftDeleted
            };

            // Invoke the OpenFolder operation with OpenSoftDeleted set to open a hard-deleted folder.
            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            if (Common.IsRequirementEnabled(814, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R814");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R814
                Site.CaptureRequirementIfAreEqual<uint>(
                    Constants.SuccessCode,
                    openFolderResponse.ReturnValue,
                    814,
                    @"[In RopDeleteFolder ROP Request Buffer] DeleteFolderFlags (1 byte): If this bit [DELETE_HARD_DELETE (0x10)] is not set, the folder is soft deleted.");
            }
            #endregion

            #region Step 4. The client calls RopDeleteFolder to hard-delete [MSOXCFOLDSubfolder1] under the root folder.
            if (Common.IsRequirementEnabled(814, this.Site))
            {
                deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DelFolders | (byte)DeleteFolderFlags.DeleteHardDelete;
                deleteFolderRequest.FolderId = subfolderId1;
                deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

                Site.Assert.AreEqual<uint>(
                    0,
                    deleteFolderResponse.ReturnValue,
                    "If ROP succeeds, ReturnValue of its response will be 0 (success)");
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopDeleteFolder operation with the DeleteFolderFlags not set.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC10_RopDeleteFolderWithNotSetDeleteFolderFlags()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder again.

            // Create the [MSOXCFOLDSubfolder1] under the root folder.
            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);
            #endregion

            #region Step 2. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] under the [MSOXCFOLDSubfolder1].

            // Create the [MSOXCFOLDSubfolder2] under the [MSOXCFOLDSubfolder1].
            uint subfolderHandle2 = 0;
            ulong subfolderId2 = 0;
            this.CreateFolder(subfolderHandle1, Constants.Subfolder2, ref subfolderId2, ref subfolderHandle2);

            #endregion

            #region Step 3. The client calls RopDeleteFolder with 'DeleteFolderFlags' not set to delete [MSOXCFOLDSubfolder1] which has a subfolder.
            RopDeleteFolderRequest deleteFolderRequest;
            deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
            deleteFolderRequest.LogonId = Constants.CommonLogonId;
            deleteFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            deleteFolderRequest.DeleteFolderFlags = 0;
            deleteFolderRequest.FolderId = subfolderId1;
            RopDeleteFolderResponse deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, deleteFolderResponse.ReturnValue, "The RopDeleteFolder operation is successful!");
            Site.Assert.AreNotEqual<byte>(0x00, deleteFolderResponse.PartialCompletion, "Although the RopDeleteFolder operation is successful, the folder is not deleted!");

            #endregion

            #region Step 4. The client calls RopDeleteFolder with 'DeleteFolderFlags' not set to delete [MSOXCFOLDSubfolder2] which has no subfolder.

            deleteFolderRequest.DeleteFolderFlags = 0;
            deleteFolderRequest.FolderId = subfolderId2;
            deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, subfolderHandle1, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, deleteFolderResponse.ReturnValue, "The RopDeleteFolder operation is successful!");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1076");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1076.
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(RopDeleteFolderResponse),
                deleteFolderResponse.GetType(),
                1076,
                @"[In Processing a RopDeleteFolder ROP Request] The server responds with a RopDeleteFolder ROP response buffer.");

            #region Verify the requirement: MS-OXCFOLD_R76, MS-OXCFOLD_R94401.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R94401");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R94401
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                deleteFolderResponse.PartialCompletion,
                94401,
                @"[In RopDeleteFolder ROP] The folder can be [either a public folder or] a private mailbox folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R76");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R76
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                deleteFolderResponse.PartialCompletion,
                76,
                @"[In RopDeleteFolder ROP] By default, the RopDeleteFolder ROP operates only on empty folders.");

            #endregion

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopMoveFolder operation performs successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC11_RopMoveFolderSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] as destination folder under the root folder.

            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);

            #endregion

            #region Step 2. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] as source folder under the root folder.

            uint subfolderHandle2 = 0;
            ulong subfolderId2 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder2, ref subfolderId2, ref subfolderHandle2);

            #endregion

            #region Step 3. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder3] for the source folder.

            uint subfolderHandle3 = 0;
            ulong subfolderId3 = 0;
            this.CreateFolder(subfolderHandle2, Constants.Subfolder3, ref subfolderId3, ref subfolderHandle3);

            #endregion

            #region Step 4. The client calls RopGetHierarchyTable to retrieve the hierarchy table for the [MSOXCFOLDSubfolder2].

            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest
            {
                RopId = (byte)RopId.RopGetHierarchyTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.Depth
            };
            RopGetHierarchyTableResponse getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, subfolderHandle2, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "The RopGetHierarchyTable ROP operation performs successfully.");
            Site.Assert.AreEqual<uint>(1, getHierarchyTableResponse.RowCount, "The folder which will be moved is stored in the source folder now.");

            #endregion

            #region Step 5. The client calls RopGetHierarchyTable to retrieve the hierarchy table for the destination folder.

            getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, subfolderHandle1, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "The RopGetHierarchyTable operation is successful!");
            Site.Assert.AreEqual<uint>(0, getHierarchyTableResponse.RowCount, "The destination folder has no folder now!");

            #endregion

            #region Step 6. The client calls RopMoveFolder to move [MSOXCFOLDSubfolder2] folder from the root folder to [MSOXCFOLDSubfolder1] synchronously.

            // Initialize a server object handle table.
            List<uint> handleList = new List<uint>
            {
                // Add the source folder handle to the server object handle table, and its index value is 0x00.
                // Add the destination folder handle to the server object handle table, and its index value is 0x01.
                subfolderHandle2, subfolderHandle1
            };
            
            RopMoveFolderRequest moveFolderRequest = new RopMoveFolderRequest
            {
                RopId = (byte)RopId.RopMoveFolder,
                LogonId = Constants.CommonLogonId,
                SourceHandleIndex = 0x00,
                DestHandleIndex = 0x01,
                WantAsynchronous = 0x00,
                UseUnicode = 0x01,
                FolderId = subfolderId3,
                NewFolderName = Encoding.Unicode.GetBytes(Constants.Subfolder4)
            };
            RopMoveFolderResponse moveFolderResponse = this.Adapter.MoveFolder(moveFolderRequest, handleList, ref this.responseHandles);
            handleList.Clear();

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, moveFolderResponse.ReturnValue, "The RopMoveFolder operation is successful!");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1111");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1111
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(RopMoveFolderResponse),
                moveFolderResponse.GetType(),
                1111,
                @"[In Processing a RopMoveFolder ROP Request] The server responds with a RopMoveFolder ROP response buffer.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R423");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R423.
            // The RopMoveFolder ROP operation performed successfully, MS-OXCFOLD_R423 can be verified directly.
            Site.CaptureRequirement(
                423,
                @"[In Moving a Folder and Its Contents] To move a folder from one parent folder to another, the client sends a RopMoveFolder ROP request ([MS-OXCROPS] section 2.2.4.7).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R180");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R180.
            // The WantAsynchronous was set to zero in RopMoveFolder ROP Request and the server responds a RopMoveFolder ROP response indicates the ROP operation performed synchronously, MS-OXCFOLD_R180 can be verified directly.
            Site.CaptureRequirement(
                180,
                @"[In RopMoveFolder ROP Request Buffer] WantAsynchronous (1 byte): [A Boolean value that is] zero (FALSE) if the ROP is to be processed synchronously.");

            #region Verify the requirement: MS-OXCFOLD_R17101 and MS-OXCFOLD_R185.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R17101");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R17101
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                moveFolderResponse.PartialCompletion,
                17101,
                @"[In RopMoveFolder ROP] The move can be within a private mailbox [or a public folder, or between a private mailbox and a public folder].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R185");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R185
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                moveFolderResponse.PartialCompletion,
                185,
                @"[In RopMoveFolder ROP Request Buffer] FolderId (8 bytes): A FID structure ([MS-OXCDATA] section 2.2.1.1) that specifies the folder to be moved.");

            #endregion
            #endregion

            #region Step 7. The client gets the properties from server.

            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest();
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;
            getPropertiesSpecificRequest.RopId = (byte)RopId.RopGetPropertiesSpecific;
            getPropertiesSpecificRequest.LogonId = Constants.CommonLogonId;
            getPropertiesSpecificRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            getPropertiesSpecificRequest.PropertySizeLimit = 0xFFFF;

            PropertyTag[] tags = new PropertyTag[1];
            PropertyTag tag;

            // Get the property: PidTagDisplayName.
            tag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagDisplayName,
                PropertyType = (ushort)PropertyType.PtypString
            };
            tags[0] = tag;

            getPropertiesSpecificRequest.PropertyTagCount = (ushort)tags.Length;
            getPropertiesSpecificRequest.PropertyTags = tags;
            getPropertiesSpecificResponse = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, subfolderHandle3, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getPropertiesSpecificResponse.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");

            string pidTagDisplayNameValue = System.Text.UnicodeEncoding.Unicode.GetString(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value);

            #region Verify the requirements: MS-OXCFOLD_R183, MS-OXCFOLD_R186.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R183");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R183
            Site.CaptureRequirementIfAreEqual<string>(
                Constants.Subfolder4,
                pidTagDisplayNameValue,
                183,
                @"[In RopMoveFolder ROP Request Buffer] UseUnicode (1 byte): A Boolean value that is nonzero (TRUE) if the value of the NewFolderName field is formatted in Unicode.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R186");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R186
            Site.CaptureRequirementIfAreEqual<string>(
                Constants.Subfolder4,
                pidTagDisplayNameValue,
                186,
                @"[In RopMoveFolder ROP Request Buffer] NewFolderName (variable): A null-terminated string that specifies the new name for the moved folder.");

            #endregion

            #endregion

            #region Step 8. The client calls RopGetHierarchyTable to retrieve the hierarchy table for the source after moving the target folder in the step 7.

            getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, subfolderHandle2, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "The RopGetHierarchyTable operation is successful!");
            Site.Assert.AreEqual<uint>(0, getHierarchyTableResponse.RowCount, "The source folder has no folder now, the target folder has been moved to destination folder!");

            #endregion

            #region Step 9. The client calls RopGetHierarchyTable to retrieve the hierarchy table for the destination after moving the target folder in the step 7, in order to see whether the target folder is in it or not.

            getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, subfolderHandle1, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "The RopGetHierarchyTable operation is successful!");
            Site.Assert.AreEqual<uint>(1, getHierarchyTableResponse.RowCount, "The destination folder has a folder which has been moved from the source folder now!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R190");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R190
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00,
                moveFolderResponse.PartialCompletion,
                190,
                @"[In RopMoveFolder ROP Response Buffer] PartialCompletion (1 byte): Otherwise [if the ROP successes for a subset of targets], the value of [of this field PartialCompletion] is zero (FALSE).");
            #endregion

            #region Step 10. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder5] under [MSOXCFOLDSubfolder2].

            uint subfolderHandle5 = 0;
            ulong subfolderId5 = 0;
            this.CreateFolder(subfolderHandle2, Constants.Subfolder5, ref subfolderId5, ref subfolderHandle5);

            #endregion

            #region Step 11. The client creates a message under the [MSOXCFOLDSubfolder5].

            uint messageHandle = 0;
            ulong messageId = 0;
            this.CreateSaveMessage(subfolderHandle5, subfolderId5, ref messageId, ref messageHandle);

            #endregion

            #region Step 12. The client calls RopMoveFolder to move target folder.

            // Add the source folder handle to the server object handle table, and its index value is 0x00.
            handleList.Add(subfolderHandle2);

            // Add the destination folder handle to the server object handle table, and its index value is 0x01.
            handleList.Add(subfolderHandle1);

            moveFolderRequest = new RopMoveFolderRequest
            {
                RopId = (byte)RopId.RopMoveFolder,
                LogonId = Constants.CommonLogonId,
                SourceHandleIndex = 0x00,
                DestHandleIndex = 0x01,
                WantAsynchronous = 0x00,
                UseUnicode = 0x01,
                FolderId = subfolderId5,
                NewFolderName = Encoding.Unicode.GetBytes(Constants.Subfolder6)
            };
            moveFolderResponse = this.Adapter.MoveFolder(moveFolderRequest, handleList, ref this.responseHandles);
            handleList.Clear();

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, moveFolderResponse.ReturnValue, "The RopMoveFolder operation is successful!");
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, moveFolderResponse.PartialCompletion, "All target folders were moved successfully.");

            #endregion

            #region Step 13. The client calls RopGetHierarchyTable to retrieve the hierarchy table for the destination after moving the target folder in the step 12, in order to see whether the target folder is in it or not.

            getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, subfolderHandle1, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "The RopGetHierarchyTable operation is successful!");
            Site.Assert.AreEqual<uint>(2, getHierarchyTableResponse.RowCount, "The destination folder has a folder which has been moved from the source folder now!");

            #endregion

            #region Step 14. The client gets the properties from server.
             tags = new PropertyTag[1];

            // Get the property: PidTagDisplayName.
            tag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagDisplayName,
                PropertyType = (ushort)PropertyType.PtypString
            };
            tags[0] = tag;

            getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                PropertySizeLimit = 0xFFFF,
                PropertyTagCount = (ushort)tags.Length,
                PropertyTags = tags
            };
            getPropertiesSpecificResponse = this.Adapter.GetFolderObjectSpecificProperties(getPropertiesSpecificRequest, subfolderHandle5, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getPropertiesSpecificResponse.ReturnValue, "RopGetPropertiesSpecific ROP operation performs successfully!");
            #endregion

            #region Step 15. The client calls RopGetContentsTable to retrieve the contents table for the target folder.

            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };
            RopGetContentsTableResponse getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, subfolderHandle5, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R170");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R170
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                getContentsTableResponse.RowCount,
                170,
                @"[In RopMoveFolder ROP] All contents of the folder are moved with it.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopCopyFolder operation performs successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC12_RopCopyFolderSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] as destination folder under the root folder.

            uint destinationFolderHandle = 0;
            ulong destinationFolderId = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref destinationFolderId, ref destinationFolderHandle);

            #endregion

            #region Step 2. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] as target folder under the root folder.

            uint targetFolderHandle = 0;
            ulong targetFolderId = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder2, ref targetFolderId, ref targetFolderHandle);

            #endregion

            #region Step 3. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder3] under [MSOXCFOLDSubfolder2].

            uint subTargetFolderHandle = 0;
            ulong subTargetFolderId = 0;
            this.CreateFolder(targetFolderHandle, Constants.Subfolder3, ref subTargetFolderId, ref subTargetFolderHandle);

            #endregion

            #region Step 4. The client creates a message in [MSOXCFOLDSubfolder2].

            ulong messageId = 0;
            uint messageHandle = 0;
            this.CreateSaveMessage(targetFolderHandle, targetFolderId, ref messageId, ref messageHandle);

            #endregion

            #region Step 5. The client calls RopGetHierarchyTable to retrieve the hierarchy table for the destination folder [MSOXCFOLDSubfolder1].

            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest
            {
                RopId = (byte)RopId.RopGetHierarchyTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.Depth
            };
            RopGetHierarchyTableResponse getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, destinationFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs successful!");
            Site.Assert.AreEqual<uint>(0, getHierarchyTableResponse.RowCount, "The destination folder has no folder now!");

            #endregion

            #region Step 6. The client calls RopCopyFolder to copy target folder [MSOXCFOLDSubfolder2] from the root folder to destination folder [MSOXCFOLDSubfolder1] synchronously.

            // Initialize a server object handle table.
            List<uint> handleList = new List<uint>
            {
                // Add the root folder handle to the server object handle table, and its index value is 0x00.
                // Add the Subfolder1 handle to the server object handle table, and its index value is 0x01.
                this.RootFolderHandle, destinationFolderHandle
            };

            // Call the RopCopyFolder operation to copy the folder.
            RopCopyFolderRequest copyFolderRequest = new RopCopyFolderRequest()
            {
                RopId = (byte)RopId.RopCopyFolder,
                LogonId = Constants.CommonLogonId,
                SourceHandleIndex = 0x00,
                DestHandleIndex = 0x01,
                WantAsynchronous = 0x00,
                WantRecursive = 0xFF,
                FolderId = targetFolderId,
                UseUnicode = 0x01,
                NewFolderName = Encoding.Unicode.GetBytes(Constants.Subfolder4)
            };
            RopCopyFolderResponse copyFolderResponse = this.Adapter.CopyFolder(copyFolderRequest, handleList, ref this.responseHandles);
            handleList.Clear();

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, copyFolderResponse.ReturnValue, "RopCopyFolder ROP operation performs successful!");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1118");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1118.
            // The RopCopyFolder ROP operation performed successfully, this requirement can be captured directly.
            Site.CaptureRequirement(
                1118,
                @"[In Processing a RopCopyFolder ROP Request] The server responds with a RopCopyFolder ROP response buffer.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1048");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1048.
            // The RopCopyFolder ROP operation performed successfully, this requirement can be captured directly.
            Site.CaptureRequirement(
                1048,
                @"[In Copying a Folder and Its Contents] To copy a folder from one parent folder to another, the client sends a RopCopyFolder ROP request ([MS-OXCROPS] section 2.2.4.8).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R207");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R207.
            // The WantAsynchronous was set to zero and the server responds a RopCopyFolder ROP response indicates the ROP is processed synchronously, MS-OXCFOLD_R155 can be verified directly.
            Site.CaptureRequirement(
                207,
                @"[In RopCopyFolder ROP Request Buffer] WantAsynchronous (1 byte): [A Boolean value that is] zero (FALSE) if the ROP is to be processed synchronously.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R212");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R212.
            // The client request the server to copy a folder and named the new folder to format in Unicode, the server responds a success value, and MS-OXCFOLD_R212 can be verified directly.
            Site.CaptureRequirement(
                212,
                @"[In RopCopyFolder ROP Request Buffer] UseUnicode (1 byte): A Boolean value that is nonzero (TRUE) if the value of the NewFolderName field is formatted in Unicode.");
            #endregion

            #region Step 7. The client calls RopGetHierarchyTable to retrieve the hierarchy table for the destination folder [MSOXCFOLDSubfolder1] after copying operation completed.

            getHierarchyTableRequest.TableFlags = (byte)FolderTableFlags.Depth;
            getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, destinationFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs successful!");
            Site.Assert.AreEqual<uint>(2, getHierarchyTableResponse.RowCount, "The destination folder has two folder now!");

            uint getHierarchyTableHandle = this.responseHandles[0][getHierarchyTableResponse.OutputHandleIndex];

            #region Verify the requirement: MS-OXCFOLD_R210, MS-OXCFOLD_R219, MS-OXCFOLD_R429 and MS-OXCFOLD_R1122.
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R210, the return value of CopyFolder is {0}, the RowCount of GetHierarchyTable is {1}.", copyFolderResponse.ReturnValue, getHierarchyTableResponse.RowCount);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R210.
            bool isVerifyR210 = copyFolderResponse.ReturnValue == Constants.SuccessCode && getHierarchyTableResponse.RowCount > 1;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR210,
                210,
                @"[In RopCopyFolder ROP Request Buffer] WantRecursive (1 byte): A Boolean value that is nonzero (TRUE) if the folder is to be copied recursively-that is, all of the folder's subfolders are copied to the new folder and the subfolders' subfolders are copied to the new folder and so on.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R219");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R219
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                copyFolderResponse.PartialCompletion,
                219,
                @"[In RopCopyFolder ROP Response Buffer] PartialCompletion (1 byte): Otherwise [if the ROP successes for a subset of targets], the value [of PartialCompletion field] is zero (FALSE).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R429, the return value of CopyFolder is {0}, the RowCount of GetHierarchyTable is {1}.", copyFolderResponse.ReturnValue, getHierarchyTableResponse.RowCount);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R429
            bool isVerifyR429 = copyFolderResponse.ReturnValue == Constants.SuccessCode && getHierarchyTableResponse.RowCount > 1;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR429,
                429,
                @"[In Copying a Folder and Its Contents] If the WantRecursive field is set to nonzero (TRUE), as specified in section 2.2.1.8.1, the subfolders that are contained in the source folder are also duplicated in the new folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1122, the return value of CopyFolder is {0}, the RowCount of GetHierarchyTable is {1}.", copyFolderResponse.ReturnValue, getHierarchyTableResponse.RowCount);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1122
            bool isVerifyR1122 = copyFolderResponse.ReturnValue == Constants.SuccessCode && getHierarchyTableResponse.RowCount > 1;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1122,
                1122,
                @"[In Processing a RopCopyFolder ROP Request] If the WantRecursive field of the RopCopyFolder ROP request buffer is set to nonzero (TRUE), as specified in section 2.2.1.8.1, the subfolders contained in the source folder are also duplicated in the new folder in a recursive manner.");
            #endregion

            #region Get the PidTagDisplayName and PidTagFolderId properties value of the [MSOXCFOLDSubfolder1] from the rows of the hierarchy table object.

            PropertyTag[] propertyTags = new PropertyTag[2];
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagDisplayName,
                PropertyType = (ushort)PropertyType.PtypString
            };
            propertyTags[0] = propertyTag;

            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)FolderPropertyId.PidTagFolderId,
                PropertyType = (ushort)PropertyType.PtypInteger64
            };
            propertyTags[1] = propertyTag;

            List<PropertyRow> propertyRows = this.GetTableRowValue(getHierarchyTableHandle, (ushort)getHierarchyTableResponse.RowCount, propertyTags);
            Site.Assert.IsNotNull(propertyRows, "The PidTagDisplayName and PidTagFolderId properties value could not be retrieved from the hierarchy table object of the [MSOXCFOLDSubfolder1].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R397");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R397.
            Site.CaptureRequirementIfIsNotNull(
                propertyRows[0].PropertyValues[1].Value,
                397,
                "[In Opening a Folder] The FID can be retrieved from the hierarchy table that contains the folder's information by including the PidTagFolderId property (section 2.2.2.2.1.6) in a RopSetColumns request ([MS-OXCROPS] section 2.2.5.1.1).");

            ulong copyFolderNewId = BitConverter.ToUInt64(propertyRows[0].PropertyValues[1].Value, 0);

            #endregion
            #endregion

            #region Step 8. The client opens the new copied folder object.

            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderId = copyFolderNewId,
                OpenModeFlags = (byte)FolderOpenModeFlags.OpenSoftDeleted
            };
            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, destinationFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, openFolderResponse.ReturnValue, "RopOpenFolder ROP operation performs successful!");

            uint copyFolderNewHandle = this.responseHandles[0][openFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 9. The client calls RopGetContentsTable to retrieve the contents of the new copied folder object.

            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None | (byte)FolderTableFlags.UseUnicode
            };
            RopGetContentsTableResponse getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, copyFolderNewHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
            uint getContentsTableHandle = this.responseHandles[0][getContentsTableResponse.OutputHandleIndex];

            #region Verify the requirement: MS-OXCFOLD_R198, MS-OXCFOLD_R605.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R605");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R605
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                getContentsTableResponse.RowCount,
                605,
                @"[In Processing a RopCopyFolder ROP Request] All messages contained in the source folder MUST be duplicated in the new folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R198");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R198
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                getContentsTableResponse.RowCount,
                198,
                @"[In RopCopyFolder ROP] All contents of the folder are copied with it.");

            #endregion

            #region Get the PidTagMessageClass property value from the rows of the contents table object of the new copied folder.

            propertyTags = new PropertyTag[1];
            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)MessagePropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            propertyTags[0] = propertyTag;

            propertyRows = this.GetTableRowValue(getContentsTableHandle, (ushort)getContentsTableResponse.RowCount, propertyTags);
            Site.Assert.IsNotNull(propertyRows, "The PidTagMessageClass property value could not be retrieved from the contents table object of the new copied folder.");

            string newFolderName = Encoding.Unicode.GetString(propertyRows[0].PropertyValues[0].Value);
            #endregion

            #region Verify the requirement: MS-OXCFOLD_R1020 and MS-OXCFOLD_R311.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1020");

            // The client create a message with message class "IPM.Note" in the step 4.
            // If the property PidTagMessageClass is encoded to "IPM.Note" by Unicode encoding, this requirement can be verified.
            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1020
            Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Note\0",
                newFolderName,
                1020,
                @"[In RopGetContentsTable ROP Request Buffer] If this bit [UseUnicode] is set, the columns that contain string data are returned in Unicode format.");
            
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R311");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R326.
            // The PidTagMessageClass property value was get successfully from the contents table by the table object handle, MS-OXCFOLD_R326 can be verified directly.
            Site.CaptureRequirement(
                326,
                @"[In RopGetContentsTable ROP Request Buffer] OutputHandleIndex (1 byte): The output Server object for this operation [RopGetContentsTable ROP] is a Table object that represents the contents table.");
            #endregion

            #endregion

            #region Step 10. The client calls RopGetContentsTable that the UseUnicode bit of TableFlags is not set.

            getContentsTableRequest = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };
            
            getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, copyFolderNewHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
            getContentsTableHandle = this.responseHandles[0][getContentsTableResponse.OutputHandleIndex];

            #region Get the PidTagMessageClass property value from the rows of the contents table object of the new copied folder.

            propertyTags = new PropertyTag[1];
            propertyTag = new PropertyTag
            {
                PropertyId = (ushort)MessagePropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString8
            };
            propertyTags[0] = propertyTag;

            propertyRows = this.GetTableRowValue(getContentsTableHandle, (ushort)getContentsTableResponse.RowCount, propertyTags);
            Site.Assert.IsNotNull(propertyRows, "The PidTagMessageClass property value could be retrieved from the contents table object of the new copied folder.");

            newFolderName = Encoding.ASCII.GetString(propertyRows[0].PropertyValues[0].Value);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1021");
        
            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1021
            this.Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Note\0",
                newFolderName,
                1021,
                @"[In RopGetContentsTable ROP Request Buffer] If this bit [UseUnicode] is not set, the string data is encoded in the code page of the Logon object.");
            #endregion
            #endregion

            #region Step 11. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder5] as target folder under the root folder.

            // Create Subfolder5 as a target folder under the root folder.
            uint targetFolderHandle2 = 0;
            ulong targetFolderId2 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder5, ref targetFolderId2, ref targetFolderHandle2);

            #endregion

            #region Step 12. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder6] as target folder under the Subfolder5.

            // Create Subfolder6 as a sub-target folder under the Subfolder5.
            uint subTargetFolderHandle2 = 0;
            ulong subTargetFolderId2 = 0;
            this.CreateFolder(targetFolderHandle2, Constants.Subfolder6, ref subTargetFolderId2, ref subTargetFolderHandle2);

            #endregion

            #region Step 13. The client calls RopCopyFolder to copy target folder [MSOXCFOLDSubfolder5] from the root folder to destination folder [MSOXCFOLDSubfolder1] synchronously with setting the 'WantRecursive' flag to 0x00.

            // Add the root folder handle to the server object handle table, and it index value is 0x00.
            handleList.Add(this.RootFolderHandle);

            // Add the Subfolder1 handle to the server object handle table, and it index value is 0x01.
            handleList.Add(destinationFolderHandle);
            copyFolderRequest.UseUnicode = 0x00;
            copyFolderRequest.WantRecursive = 0x00;
            copyFolderRequest.FolderId = targetFolderId;
            copyFolderRequest.NewFolderName = Encoding.ASCII.GetBytes(Constants.Subfolder7);
            copyFolderResponse = this.Adapter.CopyFolder(copyFolderRequest, handleList, ref this.responseHandles);
            handleList.Clear();

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, copyFolderResponse.ReturnValue, "RopCopyFolder ROP operation performs successful!");

            #endregion

            #region Step 14. The client calls RopGetHierarchyTable to retrieve the hierarchy table for the destination folder [MSOXCFOLDSubfolder1] after the second copy operation completed.

            getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, destinationFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs successful!");
            Site.Assert.AreEqual<uint>(3, getHierarchyTableResponse.RowCount, "The destination folder has three folder after copying twice!");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R211");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R211.
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                copyFolderResponse.PartialCompletion,
                211,
                @"[In RopCopyFolder ROP Request Buffer] WantRecursive (1 byte): [A Boolean value that is nonzero (TRUE) if the folder is to be copied recursively-that is, all of the folder's subfolders are copied to the new folder and the subfolders' subfolders are copied to the new folder and so on.] The value is zero (FALSE) otherwise.");
            #endregion

            #region Step 15. The client calls RopCopyFolder with UseUnicode setting false and NewFolderName encoding in unicode.
            object ropResponse = null;
            copyFolderRequest.UseUnicode = 0x00;
            copyFolderRequest.NewFolderName = Encoding.Unicode.GetBytes(Constants.Subfolder7);
            FormatException formatException = null;

            try
            {
                this.Adapter.DoRopCall(copyFolderRequest, handleList, ref ropResponse, ref this.responseHandles);
            }
            catch (FormatException e)
            {
                formatException = e;
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R213");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R213.
            // The program throw a format exception indicates the value of the NewFolderName field is not formatted in Unicode.
            Site.CaptureRequirementIfIsNotNull(
                formatException,
                213,
                @"[In RopCopyFolder ROP Request Buffer] UseUnicode (1 byte): it [UseUnicode] is zero (FALSE) otherwise [if the value of the NewFolderName field is not formatted in Unicode].");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopOpenFolder operation responds with error codes.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC13_RopOpenFolderFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopOpenFolder with Non-existing folder ID.
            
            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderId = ulong.MaxValue, // Set a non-existing folder ID to open this non-existing folder, in order to make server return an error code.
                OpenModeFlags = (byte)FolderOpenModeFlags.OpenSoftDeleted // Set this bit to indicate that server opens either an existing folder or a soft-deleted folder.
            };

            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, this.LogonHandle, ref this.responseHandles);

            #region Verify the requirements: MS-OXCFOLD_R468, MS-OXCFOLD_R469.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R468");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R468
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                openFolderResponse.ReturnValue,
                468,
                @"[In Processing a RopOpenFolder ROP Request]The value of error code ecNotFound is 0x8004010F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R469");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R469
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                openFolderResponse.ReturnValue,
                469,
                @"[In Processing a RopOpenFolder ROP Request] When the error code is ecNotFound, it indicates the FID ([MS-OXCDATA] section 2.2.1.1) does not correspond to a folder in the database.");

            #endregion

            #endregion

            #region Step 2. The client calls RopGetContentsTable to retrieve the contents table for the root folder.

            RopGetContentsTableRequest getContentsTableRequest;
            getContentsTableRequest.RopId = (byte)RopId.RopGetContentsTable;
            getContentsTableRequest.LogonId = Constants.CommonLogonId;
            getContentsTableRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            getContentsTableRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;

            // Set this bit to make the server retrieve the hierarchy table lists folders from all levels under the folder.
            getContentsTableRequest.TableFlags = 0x00;
            RopGetContentsTableResponse getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, this.RootFolderHandle, ref this.responseHandles);
            uint inboxContentsTableHandle = this.responseHandles[0][getContentsTableResponse.OutputHandleIndex];

            #endregion

            #region Step 3. The client calls RopOpenFolder with non-folder object handle.

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = Constants.CommonLogonId;
            openFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            openFolderRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            openFolderRequest.FolderId = this.RootFolderId;

            // Set this bit to make the server opens either an existing folder or a soft-deleted folder.
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.OpenSoftDeleted;

            // The inboxContentsTableHandle is a table object handle to refer a table object in which case is purposed to test error code ecNotSupported [0x80040102].  
            openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, inboxContentsTableHandle, ref this.responseHandles);

            #region Verify the requirement: MS-OXCFOLD_R472 and MS-OXCFOLD_R473.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R472");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R472
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                openFolderResponse.ReturnValue,
                472,
                @"[In Processing a RopOpenFolder ROP Request]The value of error code ecNotSupported is 0x80040102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R473");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R473
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                openFolderResponse.ReturnValue,
                473,
                @"[In Processing a RopOpenFolder ROP Request] When the error code is ecNotSupported, it indicates the object that this ROP [RopOpenFolder ROP] was called on is not a Folder object or Logon object.");

            #endregion

            #endregion

            #region Step 4. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.

            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);

            #endregion

            #region Step 5. The client calls RopDeleteFolder to softly delete [MSOXCFOLDSubfolder1] folder.

            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest
            {
                RopId = (byte)RopId.RopDeleteFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                DeleteFolderFlags = (byte)DeleteFolderFlags.DelFolders | (byte)DeleteFolderFlags.DelMessages,
                FolderId = subfolderId1
            };
            RopDeleteFolderResponse deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, deleteFolderResponse.ReturnValue, "The RopDeleteFolder operation is successful!");

            #endregion

            #region Step 6. The client calls RopOpenFolder without setting the OpenModeFlags field.

            RopOpenFolderRequest openFolderWithoutSetOpenModeFlagsRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderId = subfolderId1,
                OpenModeFlags = (byte)FolderOpenModeFlags.None
            };

            RopOpenFolderResponse openFolderWithoutSetOpenModeFlagsResponse = this.Adapter.OpenFolder(openFolderWithoutSetOpenModeFlagsRequest, this.RootFolderHandle, ref this.responseHandles);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R471");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R471.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                openFolderWithoutSetOpenModeFlagsResponse.ReturnValue,
                471,
                @"[In Processing a RopOpenFolder ROP Request] When the error code is ecNotFound, it indicates the folder is soft-deleted and the client has not specified the OpenSoftDeleted bit in the OpenModeFlags field.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the RopCreateFolder operation responds with error codes.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC14_RopCreateFolderFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopCreateFolder with a logon object handle.

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = 0x01,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x01,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder1),
                Comment = Encoding.ASCII.GetBytes(Constants.Subfolder1)
            };

            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.LogonHandle, ref this.responseHandles);

            #region Verify the requirements: MS-OXCFOLD_R494, MS-OXCFOLD_R495.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R494");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R494
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                createFolderResponse.ReturnValue,
                494,
                @"[In Processing a RopCreateFolder ROP Request]The value of error code ecNotSupported is 0x80040102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R495");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R495
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                createFolderResponse.ReturnValue,
                495,
                @"[In Processing a RopCreateFolder ROP Request] When the error code is ecNotSupported, it indicates the object that this ROP [RopCreateFolder ROP] was called on an object that is not a Folder object.");

            #endregion

            #endregion

            #region Step 2. The client calls RopCreateFolder with invalid folder type.

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = Constants.CommonLogonId;
            createFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            createFolderRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;

            // Set the folder type to an invalid value in order to make the server return an error code "ecInvalidParam".
            createFolderRequest.FolderType = 0x03;
            createFolderRequest.UseUnicodeStrings = 0x0;
            createFolderRequest.OpenExisting = 0x00;
            createFolderRequest.Reserved = 0x0;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            #region Verify the requirements: MS-OXCFOLD_R486, MS-OXCFOLD_R487.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R486");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R486
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80070057,
                createFolderResponse.ReturnValue,
                486,
                @"[In Processing a RopCreateFolder ROP Request]The value of error code ecInvalidParam is 0x80070057.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R487");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R487
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80070057,
                createFolderResponse.ReturnValue,
                487,
                @"[In Processing a RopCreateFolder ROP Request] When the error code is ecInvalidParam, it indicates the FolderType field contains an invalid value.");

            #endregion

            #endregion

            #region Step 3. The client calls RopCreateFolder with the OpenExisting field set to zero to create [MSOXCFOLDSubfolder3] under root folder for twice.

            ulong folderId = 0;
            uint folderHandle = 0;

            // The first time running the code in the loop is to create a folder
            // The second time running the code in the loop is to create a duplicated folder 
            for (int i = 0; i <= 1; i++)
            {
                createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
                createFolderRequest.LogonId = Constants.CommonLogonId;
                createFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
                createFolderRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
                createFolderRequest.FolderType = 0x01;
                createFolderRequest.UseUnicodeStrings = 0x0;
                createFolderRequest.OpenExisting = 0x00;
                createFolderRequest.Reserved = 0x0;
                createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder3);
                createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder3);
                createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);

                if (createFolderResponse.ReturnValue == Constants.SuccessCode)
                {
                    folderId = createFolderResponse.FolderId;
                    folderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
                }
            }

            #region Verify the requirements: MS-OXCFOLD_R480, MS-OXCFOLD_R1065, MS-OXCFOLD_R491, MS-OXCFOLD_R481, MS-OXCFOLD_R493 and MS-OXCFOLD_R51.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R480");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R480
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040604,
                createFolderResponse.ReturnValue,
                480,
                @"[In Processing a RopCreateFolder ROP Request] In other words, sibling folders cannot have the same name.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1065");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1065
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040604,
                createFolderResponse.ReturnValue,
                1065,
                @"[In Processing a RopCreateFolder ROP Request] The folder name MUST be unique within the parent folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R491");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R491
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040604,
                createFolderResponse.ReturnValue,
                491,
                @"[In Processing a RopCreateFolder ROP Request]The value of error code ecDuplicateName is 0x80040604.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R481");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R481
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040604,
                createFolderResponse.ReturnValue,
                481,
                @"[In Processing a RopCreateFolder ROP Request]If a folder with the same name already exists, and the OpenExisting field is set to zero (FALSE), the server fails the RopCreateFolder ROP request with error code ecDuplicateName.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R493");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R493
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040604,
                createFolderResponse.ReturnValue,
                493,
                @"[In Processing a RopCreateFolder ROP Request] When the error code is ecDuplicateName, it indicates a folder with the same name already exists, and the OpenExisting field was set to zero (FALSE).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R51");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R51
            Site.CaptureRequirementIfAreNotEqual<uint>(
                Constants.SuccessCode,
                createFolderResponse.ReturnValue,
                51,
                @"[InRopCreateFolder ROP Request Buffer] OpenExisting (1 byte): [A Boolean value that is] zero (FALSE) otherwise [if a pre-existing folder, whose name is identical to the name specified in the DisplayName field, is not to be opened.].");

            #endregion

            #endregion

            #region Step 4. The client calls RopDeleteFolder to delete [MSOXCFOLDSubfolder3] under root folder.

            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest
            {
                RopId = (byte)RopId.RopDeleteFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                DeleteFolderFlags = (byte)DeleteFolderFlags.DelMessages | (byte)DeleteFolderFlags.DelFolders,
                FolderId = folderId
            };
            RopDeleteFolderResponse deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(
                0,
                deleteFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            #region Step 5. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder4] under the soft-delete folder.

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = Constants.CommonLogonId;
            createFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            createFolderRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            createFolderRequest.FolderType = 0x01;
            createFolderRequest.UseUnicodeStrings = 0x0;
            createFolderRequest.OpenExisting = 0x00;
            createFolderRequest.Reserved = 0x0;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder4);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder4);
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, folderHandle, ref this.responseHandles);

            #endregion

            #region Verify the requirements: MS-OXCFOLD_R1072, MS-OXCFOLD_R1073.

            if (Common.IsRequirementEnabled(1073, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1072");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1072
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x8004010F,
                    createFolderResponse.ReturnValue,
                    1072,
                    @"[In Processing a RopCreateFolder ROP Request] The value of error code ecNotFound is 0x8004010F.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1073");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1073
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x8004010F,
                    createFolderResponse.ReturnValue,
                    1073,
                    @"[In Processing a RopCreateFolder ROP Request] When the error code is ecNotFound, it indicates the ROP was called on a Folder object that is a soft delete folder.");
            }
            #endregion

            #region Step 6. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder5] under the root folder.

            createFolderRequest = new RopCreateFolderRequest()
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = 0x01,
                UseUnicodeStrings = 0x01,
                OpenExisting = 0x00,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder5),
                Comment = Encoding.ASCII.GetBytes(Constants.Subfolder5)
            };
            FormatException formatException = null;
            object ropResponse = new object();
            try
            {
                this.Adapter.DoRopCall(createFolderRequest, this.RootFolderHandle, ref ropResponse, ref this.responseHandles);
            }
            catch (FormatException e)
            {
                formatException = e;
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R49");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R49.
            // The program throw a format exception indicates the values of the DisplayName and Comment fields are not formatted in Unicode.
            Site.CaptureRequirementIfIsNotNull(
                formatException,
                49,
                @"[In RopCreateFolder ROP Request Buffer] UseUnicodeStrings (1 byte): [A Boolean value that is] zero (FALSE) otherwise [if the values of the DisplayName and Comment fields are not formatted in Unicode.].");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopDeleteFolder operation responds with error codes.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S01_TC15_RopDeleteFolderFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. The client calls RopDeleteFolder with a non-existing folder ID.

            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest
            {
                RopId = (byte)RopId.RopDeleteFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                DeleteFolderFlags = (byte)DeleteFolderFlags.DelFolders | (byte)DeleteFolderFlags.DelMessages, // Set delete folder flags DelMessages and DelFolders.
                FolderId = ulong.MaxValue  // Set a non-existing folder ID in order to make server return an error code "ecNotFound".
            };

            RopDeleteFolderResponse deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.RootFolderHandle, ref this.responseHandles);

            #region Verify the requirements: MS-OXCFOLD_R504, MS-OXCFOLD_R507.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R504");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R504
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                deleteFolderResponse.ReturnValue,
                504,
                @"[In Processing a RopDeleteFolder ROP Request]The value of error code ecNotFound is 0x8004010F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R507");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R507
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                deleteFolderResponse.ReturnValue,
                507,
                @"[In Processing a RopDeleteFolder ROP Request] When the error code is ecNotFound, it indicates folder with the specified ID does not exist.");

            #endregion

            #endregion

            #region Step 2. The client creates [MSOXCFOLDSubfolder1] under the root folder.

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
            uint subFolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            ulong subFolderId1 = createFolderResponse.FolderId;

            #endregion

            #region Step 3. The client creates [MSOXCFOLDSubfolder2] under [MSOXCFOLDSubfolder1].

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = Constants.CommonLogonId;
            createFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            createFolderRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            createFolderRequest.FolderType = 0x01;
            createFolderRequest.UseUnicodeStrings = 0x0;
            createFolderRequest.OpenExisting = 0x00;
            createFolderRequest.Reserved = 0x0;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            this.Adapter.CreateFolder(createFolderRequest, subFolderHandle1, ref this.responseHandles);
            #endregion

            #region Step 4. The client calls RopGetContentsTable to retrieve the contents table for the [MSOXCFOLDSubfolder3].
            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = 0x00
            };
            RopGetContentsTableResponse getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, subFolderHandle1, ref this.responseHandles);
            uint tableObjectHandleOfSubfolder1 = this.responseHandles[0][getContentsTableResponse.OutputHandleIndex];

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "The RopGetContentsTable operation is successful!");

            #endregion

            #region Step 5. The client calls RopDeleteFolder with a non-folder object handle.

            deleteFolderRequest.DeleteFolderFlags = 0;
            deleteFolderRequest.FolderId = subFolderId1;

            // The parameter tableObjectHandleOfSubfolder1 stores a table object handle to refer a table object in which case is purposed to test error code ecNotSupported [0x80040102].  
            deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, tableObjectHandleOfSubfolder1, ref this.responseHandles);

            #endregion

            #region Verify the requirements: MS-OXCFOLD_R510, MS-OXCFOLD_R509.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R510");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R510
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                deleteFolderResponse.ReturnValue,
                510,
                @"[In Processing a RopDeleteFolder ROP Request] The value of error code ecNotSupported is 0x80040102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R509");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R509
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                deleteFolderResponse.ReturnValue,
                509,
                @"[In Processing a RopDeleteFolder ROP Request] When the error code is ecNotSupported, it indicates the object that this ROP [RopDeleteFolder ROP] was called on is not a Folder object.");

            #endregion

            #region Step 6. The client calls RopDeleteFolder with a Root folder object handle and its folder ID.

            deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DelFolders;
            deleteFolderRequest.FolderId = this.DefaultFolderIds[0];
            deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.LogonHandle, ref this.responseHandles);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R2509");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R2509
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                deleteFolderResponse.ReturnValue,
                2509,
                @"[In Processing a RopDeleteFolder ROP Request] When the error code is ecNotSupported, it indicates the object that this ROP [RopDeleteFolder ROP] was called on is an attempt was made to delete the Root folder..");
            #endregion

            #region Step 7. The client calls RopDeleteFolder that include an invalid bit in the DeleteFolderFlags field.
           
            if (Common.IsRequirementEnabled(123401, this.Site))
            {
                // The 0x02 is an invalid value of the DeleteFolderFlags field.
                deleteFolderRequest.DeleteFolderFlags = 0x02;
                deleteFolderRequest.FolderId = subFolderId1;
                deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, subFolderHandle1, ref this.responseHandles);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1235");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1235
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    deleteFolderResponse.ReturnValue,
                    1235,
                    @"[In Processing a RopDeleteFolder ROP Request] The value of error code ecInvalidParam is 0x80070057.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1236");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1236
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    deleteFolderResponse.ReturnValue,
                    1236,
                    @"[In Processing a RopDeleteFolder ROP Request] When the error code is ecInvalidParam, it indicates an invalid value was specified in a field.");
            }
            #endregion
        }

        #region Private methods
        /// <summary>
        /// Get value of specific properties in object.
        /// </summary>
        /// <param name="objectHandle">The object handle.</param>
        /// <param name="propertyList">The specific properties list.</param>
        /// <returns>The response of calling RopGetPropertiesSpecific.</returns>
        private RopGetPropertiesSpecificResponse GetSpecificProperties(uint objectHandle, List<PropertyTag> propertyList)
        {
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = 0x00, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = 0x00, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                PropertySizeLimit = 0xFFFF, // This value specifies the maximum number of the property
                PropertyTagCount = (ushort)propertyList.Count,
                PropertyTags = propertyList.ToArray()
            };
            object response = new object();
            this.Adapter.DoRopCall(getPropertiesSpecificRequest, objectHandle, ref response, ref this.responseHandles);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)response;
            Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse.ReturnValue, "Call RopGetPropertiesSpecific should success.");
            return getPropertiesSpecificResponse;
        }
        #endregion
    }
}