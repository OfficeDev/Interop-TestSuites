namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is used to verify the ROP operations on public folders.
    /// </summary>
    [TestClass]
    public class S04_OperateOnPublicFolder : TestSuiteBase
    {
        /// <summary>
        /// Server response handle list.
        /// </summary>
        private List<List<uint>> responseHandles;

        /// <summary>
        /// Logon public folder handle.
        /// </summary>
        private uint publicLogonHandle;

        /// <summary>
        /// A logon response for a public folder.
        /// </summary>
        private RopLogonResponse logonResponse;

        /// <summary>
        /// An unsigned integer indicates the public folder handle.
        /// </summary>
        private uint publicFoldersHandle;

        /// <summary>
        /// An unsigned integer indicates the root folder handle in public folder.
        /// </summary>
        private uint publicRootFolderHandle;

        /// <summary>
        /// An unsigned long indicates the root folder ID in public folder.
        /// </summary>
        private ulong publicRootFolderId;

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
        /// This test case is designed to validate that the RopCreateFolder operation performs successfully in a public folder. 
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S04_TC01_PublicFolderNonGhostedFolderValidation()
        {
            this.CheckWhetherSupportTransport();
            this.Logon();
            this.PublicFolderInitialization();

            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            #region Step 1. The client calls RopCreateFolder to create a search folder named [MSOXCFOLDSearchFolder1] under the root public folder.
            createFolderRequest = new RopCreateFolderRequest()
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
            object response = null;
            uint result = this.Adapter.DoRopCall(createFolderRequest, this.publicRootFolderHandle, ref response, ref this.responseHandles);

            if (Common.IsRequirementEnabled(10660201, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10660201");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10660201
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    result,
                    10660201,
                    @"[In Appendix A: Product Behavior] If the ROP was called to create a search folder on a public folders message store, the implemetation does return ecError <12> Section 3.2.5.2:  Exchange 2010 and Exchange 2007 return ecError.");
            }

            if (Common.IsRequirementEnabled(10660202, this.Site))
            {
                createFolderResponse = (RopCreateFolderResponse)response;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10660202");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10660202
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040102,
                    createFolderResponse.ReturnValue,
                    10660202,
                    @"[In Appendix A: Product Behavior] If the ROP was called to create a search folder on a public folders message store, the implemetation does return ecNotSupported. <12> Exchange 2013 and Exchange 2016 return ecNotSupported.");
            }
            #endregion

            #region Step 2. The client calls RopCreateFolder to create a generic folder named [MSOXCFOLDSubfolder1].

            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            createFolderResponse = this.CreateFolder(this.publicRootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R3802");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R3802
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                createFolderResponse.ReturnValue,
                3802,
                @"[In RopCreateFolder ROP] The folder can be either a public folder [or a private mailbox folder].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R61");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R61.
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                createFolderResponse.IsExistingFolder,
                61,
                @"[RopCreateFolder ROP Response Buffer] IsExistingFolder (1 byte): The value is zero (FALSE) if a public folder with that name does not exist.");

            byte isExistingFolder = createFolderResponse.IsExistingFolder;
            bool isHasRulesInExistingFolder = createFolderResponse.HasRules == null;
            bool isIsGhostedInExistingFolder = createFolderResponse.IsGhosted == null;
            #endregion

            #region Step 3. The client calls RopOpenFolder to open folder named [MSOXCFOLDSubfolder1].
            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                OpenModeFlags = (byte)FolderOpenModeFlags.None,
                FolderId = subfolderId1
            };
            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, this.publicLogonHandle, ref this.responseHandles);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R90002");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R90002
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                openFolderResponse.ReturnValue,
                90002,
                @"[In RopOpenFolder ROP] The folder can be either a public folder [or a private mailbox folder].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R30");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R30
            Site.CaptureRequirementIfIsNotNull(
                openFolderResponse.IsGhosted,
                30,
                @"[In RopOpenFolder ROP Response Buffer] IsGhosted (1 byte): This field [IsGhosted] is present only for folders that are in a public store.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R29");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R29.
            Site.CaptureRequirementIfAreEqual<byte?>(
                0x00,
                openFolderResponse.IsGhosted,
                29,
                @"[In RopOpenFolder ROP Response Buffer] IsGhosted (1 byte): otherwise [If the server hosts an active replica of the folder], this field [IsGhosted] is set to zero (FALSE).");
            #endregion

            #region Step 4. The client calls RopCreateFolder to create [MSOXCFOLD_PublicFolderMailEnabled] which is an existing folder with 'OpeningExisting' flag set to non-zero.

            string folderDisplayName = Common.GetConfigurationPropertyValue("MailEnabledPublicFolder", this.Site) + Constants.StringNullTerminated;
            createFolderRequest = new RopCreateFolderRequest()
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = (byte)FolderType.Genericfolder,
                UseUnicodeStrings = 0x1,
                OpenExisting = 0x01,
                Reserved = 0x0,
                DisplayName = Encoding.Unicode.GetBytes(folderDisplayName),
                Comment = Encoding.Unicode.GetBytes(folderDisplayName)
            };

            // Invoke the CreateFolder operation with valid parameters, use root folder handle to indicate that the new folder will be created under the root folder.
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.publicFoldersHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");

            subfolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
 
            if (Common.IsRequirementEnabled(60001, this.Site))
            {
                // Regardless of the existence of the named public folder, if the IsExistingFolder in response is always set to zero, R60001 can be verified.
                bool isR60001Verified = isExistingFolder == 0 && createFolderResponse.IsExistingFolder == 0;
                
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R60001, regardless of the existence of the public folder, the IsExistingFolder in response should be always set to {0}, actually, when public folder exists the IsExistingFolder in response is {1}, when public folder does not exist the IsExistingFolder in response is {2}.", 0, createFolderResponse.IsExistingFolder, isExistingFolder);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R60001
                Site.CaptureRequirementIfIsTrue(
                    isR60001Verified,
                    60001,
                    @"[In Appendix A: Product Behavior] Implementation does return zero (FALSE) in the IsExistingFolder field regardless of the existence of the named public folder. <3> Section 2.2.1.2.2: Exchange 2010 Exchange 2013 and Exchange 2016 always return zero (FALSE) in the IsExistingFolder field regardless of the existence of the named public folder.");
            }

            if (Common.IsRequirementEnabled(60002, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R60002");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R60002.
                Site.CaptureRequirementIfAreNotEqual<byte>(
                    0x00,
                    createFolderResponse.IsExistingFolder,
                    60002,
                    @"[In Appendix A: Product Behavior] If a public folder with the name given by the DisplayName field of the request buffer already exists, implementation does set a nonzero (TRUE) value to IsExistingFolder field. (Microsoft Exchange Server 2007 follows this behavior.)");

                bool isHasRules = createFolderResponse.HasRules != null;
                bool isIsGhosted = createFolderResponse.IsGhosted != null;
                bool isR926Verified = isHasRules && isHasRulesInExistingFolder;
                bool isR931Verified = isIsGhosted && isIsGhostedInExistingFolder;

                // Add the debug information.
                Site.Log.Add(
                    LogEntryKind.Debug,
                    @"Verify MS-OXCFOLD_R926: When the IsExistingFolder field is set to a zero, the compare result of HasRules and null is {0}; When the IsExistingFolder field is set to a nonzero, the compare result of HasRules and null is {1}",
                    isHasRulesInExistingFolder,
                    isHasRules);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R926.
                Site.CaptureRequirementIfIsTrue(
                    isR926Verified,
                    926,
                    @"[In RopCreateFolder ROP Response Buffer] This field [HasRules] is present only if the IsExistingFolder field is set to a nonzero (TRUE) value.");

                // Add the debug information.
                Site.Log.Add(
                    LogEntryKind.Debug,
                    @"Verify MS-OXCFOLD_R931: When the IsExistingFolder field is set to a zero, the compare result of IsGhosted and null is {0}; When the IsExistingFolder field is set to a nonzero, the compare result of IsGhosted and null is {1}",
                    isIsGhostedInExistingFolder,
                    isIsGhosted);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R931.
                Site.CaptureRequirementIfIsTrue(
                    isR931Verified,
                    931,
                    @"[In RopCreateFolder ROP Response Buffer] This field [IsGhosted] is present only if the IsExistingFolder field is set to a nonzero (TRUE) value and only for folders that are in a public message store.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R68");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R68.
                Site.CaptureRequirementIfAreEqual<byte?>(
                    0x00,
                    createFolderResponse.IsGhosted,
                    68,
                    @"[In RopCreateFolder ROP Response Buffer] IsGhosted (1 byte): otherwise [If the server hosts an active replica of the folder], this field [IsGhosted] is set to zero (FALSE).");
            }
            #endregion

            #region Step 5. The client gets the PidTagAddressBookEntryId property of the [MSOXCFOLD_PublicFolderMailEnabled].
            if (Common.IsRequirementEnabled(350002, this.Site))
            {
                PropertyTag[] propertyTagArray = new PropertyTag[1];
                PropertyTag propertyTag = new PropertyTag
                {
                    PropertyId = (ushort)FolderPropertyId.PidTagAddressBookEntryId,
                    PropertyType = (ushort)PropertyType.PtypBinary
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

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R350002");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R350002.
                // The property value is returned if Flag value is 0x00.
                Site.CaptureRequirementIfAreEqual<byte>(
                    0x00,
                    getPropertiesSpecificResponse.RowData.Flag,
                    350002,
                    @"[In Appendix A: Product Behavior] The implementation does support the PidTagAddressBookEntryId property. (Exchange 2007 and Exchange 2010 follow this behavior).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R350");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R350.
                // The property value is returned if Flag value is 0x00.
                Site.CaptureRequirementIfAreEqual<byte>(
                    0x00,
                    getPropertiesSpecificResponse.RowData.Flag,
                    350,
                    @"[In PidTagAddressBookEntryId Property] This property is set only for public folders.");
            }
            #endregion

            #region Step 6. The client calls RopGetPropertiesAll on [MSOXCFOLD_PublicFolderGhosted].
            RopGetPropertiesAllResponse getAllPropertiesResponse = this.Adapter.GetFolderPropertiesAll(subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, getAllPropertiesResponse.ReturnValue, "RopGetPropertiesAllResponse ROP operation performs successfully!");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopOpenFolder operation performs successfully in a public folder.
        /// This test case depends on the second SUT. If the second SUT is not present, this test case cannot be executed.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S04_TC02_PublicFolderGhostedFolderValidation()
        {
            this.CheckWhetherSupportTransport();
            this.Logon();
            this.PublicFolderInitialization();
            string ghostedPublicFolder = Common.GetConfigurationPropertyValue("GhostedPublicFolder", this.Site) + Constants.StringNullTerminated;

            // The ghosted folder is only supported when the 2nd SUT exists.
            if (Common.GetConfigurationPropertyValue("Sut2ComputerName", this.Site) == string.Empty)
            {
                Site.Assert.Inconclusive("This case runs only when the second system under test exists.");
            }
            else
            {
                #region Step 1. The client calls OpenFolder to open the ghosted public folder.
                ulong ghostedFolderId = this.GetSubfolderIDByName(this.DefaultFolderIds[1], this.publicLogonHandle, ghostedPublicFolder);

                RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
                {
                    RopId = 0x02,
                    LogonId = Constants.CommonLogonId,
                    InputHandleIndex = Constants.CommonInputHandleIndex,
                    OutputHandleIndex = Constants.CommonOutputHandleIndex,
                    FolderId = ghostedFolderId
                };
                RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, this.publicLogonHandle, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(0, openFolderResponse.ReturnValue, "The open folder operation should succeed.");

                #region Verify requirements.

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R907");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R907
                Site.CaptureRequirementIfAreNotEqual<byte?>(
                    0x00,
                    openFolderResponse.IsGhosted,
                    907,
                    @"[In RopOpenFolder ROP Response Buffer] If the server does not host an active replica of the folder, this field [IsGhosted] is set to a nonzero (TRUE) value.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R909, the value of ServerCount is {0}.", openFolderResponse.ServerCount);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R909.
                // MS-OXCFOLD_R907 was verified and the IsGhosted field is set to a nonzero (TRUE) value, if the ServerCount is not null, MS-OXCFOLD_R909 can be verified.
                Site.CaptureRequirementIfIsNotNull(
                    openFolderResponse.ServerCount,
                    909,
                    @"[In RopOpenFolder ROP Response Buffer] This field [ServerCount] is present only if the IsGhosted field is set to a nonzero (TRUE) value.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R912, the value of CheapServerCount is {0}.", openFolderResponse.CheapServerCount);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R912
                // The CheapServerCount is not null indicates that it presents.
                Site.CaptureRequirementIfIsNotNull(
                    openFolderResponse.CheapServerCount,
                    912,
                    @"[In RopOpenFolder ROP Response Buffer] This field [CheapServerCount] is present only if the IsGhosted field is set to a nonzero (TRUE) value.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R915, the count of Servers is {0}.", openFolderResponse.Servers.Length);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R915
                // The Servers is not null indicates that it presents.
                Site.CaptureRequirementIfIsNotNull(
                    openFolderResponse.Servers,
                    915,
                    @"[In RopOpenFolder ROP Response Buffer] This field [Servers] is present only if the IsGhosted field is set to a nonzero (TRUE) value.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R913: The server that has a replica of the folder {0} is {1}", ghostedPublicFolder.Trim(Constants.StringNullTerminated.ToCharArray()), openFolderResponse.Servers[0].Trim(Constants.StringNullTerminated.ToCharArray()));

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R913.
                // The RopOpenFolder ROP response deserialized successfully, MS-OXCFOLD_R913 can be verified directly.
                Site.CaptureRequirement(
                    913,
                    @"[In RopOpenFolder ROP Response Buffer] Servers (variable): An array of null-terminated ASCII strings, each of which specifies a server that has a replica of the folder.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R908");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R908.
                Site.CaptureRequirementIfAreEqual<int>(
                    openFolderResponse.Servers.Length,
                    (int)openFolderResponse.ServerCount,
                    908,
                    @"[In RopOpenFolder ROP Response Buffer] ServerCount (2 bytes): An integer that specifies the number of servers that have a replica of the folder.");
                #endregion
                #endregion

                #region Step 2. The client calls RopCreateFolder to create a ghosted public folder under the root folder.

                // IsGhosted filed is present only if the IsExistingFolder field is set to a nonzero (TRUE) value and only for folders that are in a public store.
                // If implementation return zero (FALSE) in the IsExistingFolder field when the named folder already exists,
                // IsGhosted field related requirements can't be verified.
                if (Common.IsRequirementEnabled(60002, this.Site))
                {
                    RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
                    {
                        RopId = (byte)RopId.RopCreateFolder,
                        LogonId = Constants.CommonLogonId,
                        InputHandleIndex = Constants.CommonInputHandleIndex,
                        OutputHandleIndex = Constants.CommonOutputHandleIndex,
                        FolderType = (byte)FolderType.Genericfolder,
                        UseUnicodeStrings = 0x0,
                        OpenExisting = 0xff,
                        Reserved = 0x0,
                        DisplayName = Encoding.ASCII.GetBytes(ghostedPublicFolder),
                        Comment = Encoding.ASCII.GetBytes(ghostedPublicFolder)
                    };
                    RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.publicFoldersHandle, ref this.responseHandles);
                    Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "The open folder operation should succeed.");

                    #region Verify requirements.
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R929");

                    // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R929.
                    Site.CaptureRequirementIfAreNotEqual<byte?>(
                        0x00,
                        createFolderResponse.IsGhosted,
                        929,
                        @"[In RopCreateFolder ROP Response Buffer] If the server does not host an active replica of the folder, this field [IsGhosted] is set to a nonzero (TRUE) value.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R933, the value of ServerCount is {0}.", createFolderResponse.ServerCount);

                    // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R933.
                    // MS-OXCFOLD_R929 was verified and the IsGhosted field is set to a nonzero (TRUE) value, if the ServerCount is not null, MS-OXCFOLD_R933 can be verified.
                    Site.CaptureRequirementIfIsNotNull(
                        createFolderResponse.ServerCount,
                        933,
                        @"[In RopCreateFolder ROP Response Buffer] This field [ServerCount] is present only if the IsGhosted field is set to a nonzero (TRUE) value.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R938, the value of CheapServerCount is {0}.", createFolderResponse.CheapServerCount);

                    // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R938.
                    // MS-OXCFOLD_R929 was verified and the IsGhosted field is set to a nonzero (TRUE) value, if the ServerCount is not null, MS-OXCFOLD_R933 can be verified.
                    Site.CaptureRequirementIfIsNotNull(
                        createFolderResponse.CheapServerCount,
                        938,
                        @"[In RopCreateFolder ROP Response Buffer] This field [CheapServerCount] is present only if the IsGhosted field is set to a nonzero (TRUE) value.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R941, the count of Servers is {0}.", createFolderResponse.Servers.Length);

                    // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R941
                    // The Servers is not null indicates that it presents.
                    Site.CaptureRequirementIfIsNotNull(
                        createFolderResponse.Servers,
                        941,
                        @"[In RopCreateFolder ROP Response Buffer] This field [Servers] is present only if the IsGhosted field is set to a nonzero (TRUE) value.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R939: The server that has a replica of the folder {0} is {1}", ghostedPublicFolder.Trim(Constants.StringNullTerminated.ToCharArray()), createFolderResponse.Servers[0].Trim(Constants.StringNullTerminated.ToCharArray()));

                    // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R939.
                    // The RopCreateFolder ROP response deserialized successfully, MS-OXCFOLD_R939 can be verified directly.
                    Site.CaptureRequirement(
                        939,
                        @"[In RopCreateFolder ROP Response Buffer] Servers (variable): An array of null-terminated ASCII strings, each of which specifies a server that has a replica of the folder.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R932");

                    // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R932.
                    Site.CaptureRequirementIfAreEqual<int>(
                        createFolderResponse.Servers.Length,
                        (int)createFolderResponse.ServerCount,
                        932,
                        @"[In RopCreateFolder ROP Response Buffer] ServerCount (2 bytes): An integer that specifies the number of servers that have a replica of the folder.");
                    #endregion
                }
                #endregion
            }
        }

        /// <summary>
        /// This test case is designed to validate that the RopDeleteFolder operation performs successfully in a public folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S04_TC03_RopDeletePublicFolder()
        {
            this.CheckWhetherSupportTransport();
            this.Logon();
            this.PublicFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root public folder.

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

            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs should successfully.");
            ulong subfolderId1 = createFolderResponse.FolderId;
            #endregion

            #region Step 2. The client calls RopDeleteFolder to hard-delete [MSOXCFOLDSubfolder1].

            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest
            {
                RopId = (byte)RopId.RopDeleteFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                DeleteFolderFlags =
                    (byte)DeleteFolderFlags.DelMessages | (byte)DeleteFolderFlags.DelFolders |
                    (byte)DeleteFolderFlags.DeleteHardDelete,
                FolderId = subfolderId1
            };

            RopDeleteFolderResponse deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, deleteFolderResponse.ReturnValue, "RopDeleteFolder ROP operation performs should successfully.");

            #endregion

            #region Step 3. The client calls RopGetHierarchyTable to retrieve the hierarchy table for the root public folder.

            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest
            {
                RopId = (byte)RopId.RopGetHierarchyTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };

            RopGetHierarchyTableResponse getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs should successfully.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R94402");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R94402
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getHierarchyTableResponse.RowCount,
                94402,
                @"[In RopDeleteFolder ROP] The folder can be either a public folder [or a private mailbox folder]. ");
            #endregion

            #region Step 4. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] under the root public folder.

            createFolderRequest = new RopCreateFolderRequest
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

            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully.");
            ulong subfolderId2 = createFolderResponse.FolderId;
            #endregion

            #region Step 5. The client calls RopDeleteFolder to soft-delete [MSOXCFOLDSubfolder2].

            deleteFolderRequest = new RopDeleteFolderRequest
            {
                RopId = (byte)RopId.RopDeleteFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                DeleteFolderFlags = (byte)DeleteFolderFlags.None,
                FolderId = subfolderId2
            };

            deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, deleteFolderResponse.ReturnValue, "RopDeleteFolder ROP operation performs successfully.");

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopMoveCopyMessages operation copies a message in a public folder successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S04_TC04_RopMoveCopyMessagesInPublicFolderSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Logon();
            this.PublicFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root public folder.

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

            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs should successfully.");
            uint subfolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 2. The client creates a message in the root public folder.

            uint messageHandle = 0;
            ulong messageId = 0;
            this.CreateSaveMessage(this.publicRootFolderHandle, this.publicRootFolderId, ref messageId, ref messageHandle);

            #endregion

            #region Step 3. The client calls RopMoveCopyMessages to copy the message from the root public folder to [MSOXCFOLDSubfolder1].

            ulong[] messageIds = new ulong[] { messageId };
            List<uint> handlelist = new List<uint>
            {
                this.publicRootFolderHandle, subfolderHandle1
            };

            RopMoveCopyMessagesRequest moveCopyMessagesRequest = new RopMoveCopyMessagesRequest
            {
                RopId = (byte)RopId.RopMoveCopyMessages,
                LogonId = Constants.CommonLogonId,
                SourceHandleIndex = 0x00,
                DestHandleIndex = 0x01,
                MessageIdCount = (ushort)messageIds.Length,
                MessageIds = messageIds,
                WantAsynchronous = 0x00,
                WantCopy = 0x01
            };

            // WantCopy is non-zero (TRUE) indicates this is a copy operation.
            RopMoveCopyMessagesResponse moveCopyMessagesResponse = this.Adapter.MoveCopyMessages(moveCopyMessagesRequest, handlelist, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, moveCopyMessagesResponse.ReturnValue, "The RopMoveCopyMessages ROP operation performs should successfully.");
            handlelist.Clear();

            #endregion

            #region Step 4. The client calls RopGetContentsTable to retrieve the contents table for [MSOXCFOLDSubfolder1] with 'ConversationMembers' flag.

            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };

            RopGetContentsTableResponse getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs should successfully.");

            #region Verify the requirement: MS-OXCFOLD_R14502 and MS-OXCFOLD_R10252.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10252: the return value of the getContentsTableResponse is {0}", getContentsTableResponse.ReturnValue);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10252
            // The'ConversationMembers' flag is seted in the getContentsTableRequest, so if the getContentsTableResponse returns success. R10252 can be verified. 
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getContentsTableResponse.ReturnValue,
                10252,
                @"[In RopGetContentsTable ROP Request Buffer] This bit [ConversationMembers] is supported on public folders.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R14502");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R14502
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                getContentsTableResponse.RowCount,
                14502,
                @"[In RopMoveCopyMessages ROP] This ROP applies to both public folders [and private mailboxes].");
            #endregion
            #endregion
         }

        /// <summary>
        /// This test case is designed to validate that the RopMoveFolder operation performs successfully in a public folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S04_TC05_RopMovePublicFolder()
        {
            this.CheckWhetherSupportTransport();
            this.Logon();
            this.PublicFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root public folder.

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

            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs should successfully.");
            uint subfolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 2. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] under the root public folder.

            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder2);

            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs should successfully.");
            ulong subfolderId2 = createFolderResponse.FolderId;

            #endregion

            #region Step 3. The client calls RopMoveFolder to move the [MSOXCFOLDSubfolder2] from root public folder to [MSOXCFOLDSubfolder1].
            // Initialize a list of server object handles.
            List<uint> handleList = new List<uint>
            {
                // Add the source folder handle to the list of server object handles, and its index value is 0x00.
                // Add the destination folder handle to the server object handle table, and its index value is 0x01.
                this.publicRootFolderHandle, subfolderHandle1
            };

            RopMoveFolderRequest moveFolderRequest = new RopMoveFolderRequest
            {
                RopId = (byte)RopId.RopMoveFolder,
                LogonId = Constants.CommonLogonId,
                SourceHandleIndex = 0x00,
                DestHandleIndex = 0x01,
                WantAsynchronous = 0x00,
                UseUnicode = 0x01,
                FolderId = subfolderId2,
                NewFolderName = Encoding.Unicode.GetBytes(Constants.Subfolder3)
            };

            RopMoveFolderResponse moveFolderResponse = this.Adapter.MoveFolder(moveFolderRequest, handleList, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, moveFolderResponse.ReturnValue, "The RopMoveFolder operation performs should successfully.");
            handleList.Clear();

            #endregion

            #region Step 4. The client calls RopGetHierarchyTable to retrieve the hierarchy table of the root public folder.

            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest
            {
                RopId = (byte)RopId.RopGetHierarchyTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.Depth
            };

            RopGetHierarchyTableResponse getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs should successfully.");
            RopGetHierarchyTableResponse getHierarchyTableResponse1 = getHierarchyTableResponse;

            #endregion

            #region Step 5. The client calls RopGetHierarchyTable to retrieve the hierarchy table of the [MSOXCFOLDSubfolder1].

            getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs should successfully.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCFOLD_R17102: The subfolder count of target folder before RopMoveFolder was {0}, The subfolder count of source folder after RopMoveFolder was {1}.",
                getHierarchyTableResponse1.RowCount,
                getHierarchyTableResponse.RowCount);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R17102
            bool isVerifyR17102 = getHierarchyTableResponse1.RowCount == 2 && getHierarchyTableResponse.RowCount == 1;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR17102,
                17102,
                @"[In RopMoveFolder ROP] The move can be within [a private mailbox] or a public folder, [or between a private mailbox and a public folder].");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopCopyFolder operation performs successfully in a public folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S04_TC06_RopCopyPublicFolder()
        {
            if (!Common.IsRequirementEnabled(19702002, this.Site))
            {
                this.NeedCleanup = false;
                Site.Assert.Inconclusive("The server does not support the RopCopyFolder ROP ([MS-OXCROPS] section 2.2.4.8) for public folders.");
            }

            this.CheckWhetherSupportTransport();
            this.Logon();
            this.PublicFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root public folder.

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

            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs should successfully.");
            uint subfolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 2. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] under the root public folder.

            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder2);

            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs should successfully.");
            ulong subfolderId2 = createFolderResponse.FolderId;

            #endregion

            #region Step 3. The client calls RopCopyFolder to copy the [MSOXCFOLDSubfolder2] from the root public folder to [MSOXCFOLDSubfolder1].
            // Initialize a server object handle table.
            List<uint> handleList = new List<uint>
            {
                this.publicRootFolderHandle, subfolderHandle1
            };

            RopCopyFolderRequest copyFolderRequest = new RopCopyFolderRequest
            {
                RopId = (byte)RopId.RopCopyFolder,
                LogonId = Constants.CommonLogonId,
                SourceHandleIndex = 0x00,
                DestHandleIndex = 0x01,
                WantAsynchronous = 0x00,
                UseUnicode = 0x00,
                WantRecursive = 0xFF,
                FolderId = subfolderId2,
                NewFolderName = Encoding.ASCII.GetBytes(Constants.Subfolder3)
            };

            RopCopyFolderResponse copyFolderResponse = this.Adapter.CopyFolder(copyFolderRequest, handleList, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, copyFolderResponse.ReturnValue, "RopCopyFolder ROP operation performs successful!");
            handleList.Clear();

            #endregion

            #region Step 4. The client calls RopGetHierarchyTable to retrieve the hierarchy table of [MSOXCFOLDSubfolder1].

            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest
            {
                RopId = (byte)RopId.RopGetHierarchyTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };

            RopGetHierarchyTableResponse getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs successfully!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R19702002");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R19702002
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                getHierarchyTableResponse.RowCount,
                19702002,
                @"[In Appendix A: Product Behavior] Implementation does support the  RopCopyFolder ROP ([MS-OXCROPS] section 2.2.4.8) for public folders. (Microsoft Exchange Server 2007 and Microsoft Exchange Server 2010 follow this behavior.)");

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopEmptyFolder operation performs successfully in a public folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S04_TC07_RopEmptyPublicFolder()
        {
            if (!Common.IsRequirementEnabled(97501002, this.Site))
            {
                this.NeedCleanup = false;
                Site.Assert.Inconclusive("The server does not support the RopEmptyFolder ROP ([MS-OXCROPS] section 2.2.4.9) for public folders.");
            }

            this.CheckWhetherSupportTransport();
            this.Logon();
            this.PublicFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root public folder.

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

            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully.");
            uint subfolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            #endregion

            #region Step 2. The client creates a message in [MSOXCFOLDSubfolder1].

            uint messageHandle = 0;
            ulong messageId = 0;
            this.CreateSaveMessage(this.publicRootFolderHandle, this.publicRootFolderId, ref messageId, ref messageHandle);

            #endregion

            #region Step 3. The client creates a subfolder in [MSOXCFOLDSubfolder1].

            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder2);

            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully.");
            #endregion

            #region Step 4. The client calls RopEmptyFolder to empty [MSOXCFOLDSubfolder1].

            RopEmptyFolderRequest emptyFolderRequest = new RopEmptyFolderRequest
            {
                RopId = (byte)RopId.RopEmptyFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                WantDeleteAssociated = 0x00
            };

            // Invoke RopEmptyFolder operation to soft delete Subfolder3 from Subfolder1 without deleting Subfolder1.
            RopEmptyFolderResponse emptyFolderResponse = this.Adapter.EmptyFolder(emptyFolderRequest, subfolderHandle1, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(0, emptyFolderResponse.ReturnValue, "RopEmptyFolder ROP operation performs successfully on [MSOXCFOLDSubfolder1].");

            #endregion

            #region Step 5. The client calls RopGetContentsTable to retrieve the contents table of [MSOXCFOLDSubfolder1].

            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };

            RopGetContentsTableResponse getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");

            #endregion

            #region Step 6. The client calls RopGetHierarchyTable to retrieve the hierarchy table of [MSOXCFOLDSubfolder1].

            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest
            {
                RopId = (byte)RopId.RopGetHierarchyTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };

            RopGetHierarchyTableResponse getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs successfully!");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCFOLD_R97501002: The message count of the target folder after RopEmptyFolder is {0}, the subfolder count of the target folder after RopEmptyFolder is {1}",
                getContentsTableResponse.RowCount,
                getHierarchyTableResponse.RowCount);

            bool isVerifyR97501002 = getContentsTableResponse.RowCount == 0 && getHierarchyTableResponse.RowCount == 0;

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R97501002.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR97501002,
                97501002,
                @"[In Appendix A: Product Behavior] Implementation does support the RopEmptyFolder ROP ([MS-OXCROPS] section 2.2.4.9) for public folders. (Microsoft Exchange Server 2007 and Microsoft Exchange Server 2010 follow this behavior.)");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopHardDeleteMessagesAndSubfolders operation 
        /// hard deletes a message or a folder in public folder successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S04_TC08_RopHardDeleteMessagesAndSubfoldersInPublicFolder()
        {
            if (!Common.IsRequirementEnabled(98301002, this.Site))
            {
                this.NeedCleanup = false;
                Site.Assert.Inconclusive("The server does not support the RopHardDeleteMessagesAndSubfolders ROP for public folders.");
            }

            this.CheckWhetherSupportTransport();
            this.Logon();
            this.PublicFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root public folder.

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest();
            RopCreateFolderResponse createFolderResponse = new RopCreateFolderResponse();
            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = Constants.CommonLogonId;
            createFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            createFolderRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;
            createFolderRequest.UseUnicodeStrings = 0x0;
            createFolderRequest.OpenExisting = 0x01;
            createFolderRequest.Reserved = 0x0;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder1);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder1);

            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.publicRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully.");
            uint subfolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            #endregion

            #region Step 2. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder2] under [MSOXCFOLDSubfolder1].

            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder2);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(Constants.Subfolder2);

            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully.");
            #endregion

            #region Step 3. The client creates a message in [MSOXCFOLDSubfolder1].

            uint messageHandle = 0;
            ulong messageId = 0;
            this.CreateSaveMessage(this.publicRootFolderHandle, this.publicRootFolderId, ref messageId, ref messageHandle);

            #endregion

            #region Step 4. The client calls RopHardDeleteMessagesAndSubfolders applying to [MSOXCFOLDSubfolder1].

            object ropResponse = new object();
            RopHardDeleteMessagesAndSubfoldersRequest hardDeleteMessagesAndSubfoldersRequest = new RopHardDeleteMessagesAndSubfoldersRequest();
            RopHardDeleteMessagesAndSubfoldersResponse hardDeleteMessagesAndSubfoldersResponse;
            hardDeleteMessagesAndSubfoldersRequest.RopId = (byte)RopId.RopHardDeleteMessagesAndSubfolders;
            hardDeleteMessagesAndSubfoldersRequest.LogonId = Constants.CommonLogonId;
            hardDeleteMessagesAndSubfoldersRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            hardDeleteMessagesAndSubfoldersRequest.WantAsynchronous = 0x00;
            hardDeleteMessagesAndSubfoldersRequest.WantDeleteAssociated = 0xFF;

            this.Adapter.DoRopCall(hardDeleteMessagesAndSubfoldersRequest, subfolderHandle1, ref ropResponse, ref this.responseHandles);
            hardDeleteMessagesAndSubfoldersResponse = (RopHardDeleteMessagesAndSubfoldersResponse)ropResponse;

            Site.Assert.AreEqual<uint>(0, hardDeleteMessagesAndSubfoldersResponse.ReturnValue, "RopHardDeleteMessagesAndSubfolders ROP operation performs successfully.");
            Site.Assert.AreEqual<uint>(0, hardDeleteMessagesAndSubfoldersResponse.PartialCompletion, "If delete all subsets of targets succeeds, PartialCompletion of its response will be 0 (success)");

            #endregion

            #region Step 5. The client calls GetContentsTable to retrieve the contents table of [MSOXCFOLDSubfolder1].

            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };

            RopGetContentsTableResponse getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");

            #endregion

            #region Step 6. The client calls GetHierarchyTable to retrieve the hierarchy table of [MSOXCFOLDSubfolder1].

            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest
            {
                RopId = (byte)RopId.RopGetHierarchyTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };

            RopGetHierarchyTableResponse getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable ROP operation performs successfully!");

            // Add the debug information.
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCFOLD_R98301002: The message count of the target folder after RopEmptyFolder is {0}, the subfolder count of the target folder after RopEmptyFolder is {1}",
                getContentsTableResponse.RowCount,
                getHierarchyTableResponse.RowCount);

            bool isVerifyR98301002 = getHierarchyTableResponse.RowCount == 0 && getContentsTableResponse.RowCount == 0;

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R98301002.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR98301002,
                98301002,
                @"[In Appendix A: Product Behavior] Implementation does support the RopHardDeleteMessagesAndSubfolders ROP ([MS-OXCROPS] section 2.2.4.10) for public folders. (Microsoft Exchange Server 2007 and Microsoft Exchange Server 2010 follow this behavior.)");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopDeleteMessages operation deletes a message in a public folder successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S04_TC09_RopDeleteMessagesInPublicFolder()
        {
            this.CheckWhetherSupportTransport();
            this.Logon();
            this.PublicFolderInitialization();

            #region Step 1. The client creates a message in the root public folder.
            ulong messageId = 0;
            uint messageHandle = 0;
            this.CreateSaveMessage(this.publicRootFolderHandle, this.publicRootFolderId, ref messageId, ref messageHandle);
            #endregion

            #region Step 2. The client calls RopGetContentsTable to retrieve the contents table of the root public folder.
            RopGetContentsTableRequest getContentsTableRequestFirst = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = 0x00
            };

            // Get the Contents table of the opened folder.
            RopGetContentsTableResponse getContentsTableResponseFirst = this.Adapter.GetContentsTable(getContentsTableRequestFirst, this.publicRootFolderHandle, ref this.responseHandles);

            uint rowCountFirst = getContentsTableResponseFirst.RowCount;
            #endregion

            #region Step 3. The client calls RopDeleteMessages to delete this message in the root public folder.
            RopDeleteMessagesRequest deleteMessagesRequest = new RopDeleteMessagesRequest();

            // Organize the DeleteMessage request.
            ulong[] messageIdsDeleted = new ulong[1];
            messageIdsDeleted[0] = messageId;
            deleteMessagesRequest.RopId = (byte)RopId.RopDeleteMessages;
            deleteMessagesRequest.LogonId = Constants.CommonLogonId;
            deleteMessagesRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            deleteMessagesRequest.MessageIdCount = (ushort)messageIdsDeleted.Length;
            deleteMessagesRequest.MessageIds = messageIdsDeleted;
            RopDeleteMessagesResponse deleteMessagesResponse = this.Adapter.DeleteMessages(deleteMessagesRequest, this.publicRootFolderHandle, ref this.responseHandles);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R98801");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R98801
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                deleteMessagesResponse.ReturnValue,
                98801,
                @"[In RopDeleteMessages ROP] This ROP [RopDeleteMessages] applies to both public folders [and private mailboxes].");
            #endregion

            #region Step 4. The client calls RopGetContentsTable to retrieve the contents table of the root public folder.
            RopGetContentsTableRequest getContentsTableRequestSecond = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = 0x00
            };

            // Get the Contents table of the opened folder.
            RopGetContentsTableResponse getContentsTableResponseSecond = this.Adapter.GetContentsTable(getContentsTableRequestSecond, this.publicRootFolderHandle, ref this.responseHandles);
            uint rowCountSecond = getContentsTableResponseSecond.RowCount;

            Assert.AreEqual<uint>(rowCountFirst - 1, rowCountSecond, "If ROP succeeds, the second RowCount value returned from server should be one less than the first one. ");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopHardDeleteMessage operation hard deletes a message in a public folder successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S04_TC10_RopHardDeleteMessagesInPublicFolder()
        {
            this.CheckWhetherSupportTransport();
            this.Logon();
            this.PublicFolderInitialization();

            #region Step 1. The client creates a message in the root public folder.
            ulong messageId = 0;
            uint messageHandle = 0;
            this.CreateSaveMessage(this.publicRootFolderHandle, this.publicRootFolderId, ref messageId, ref messageHandle);
            #endregion

            #region Step 2. The client calls RopGetContentsTable to retrieve the contents table of the root public folder.
            RopGetContentsTableRequest getContentsTableRequestFirst = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = 0x00
            };

            // Get the Contents table of the opened folder.
            RopGetContentsTableResponse getContentsTableResponseFirst = this.Adapter.GetContentsTable(getContentsTableRequestFirst, this.publicRootFolderHandle, ref this.responseHandles);

            uint rowCountFirst = getContentsTableResponseFirst.RowCount;
            #endregion

            #region Step 3. The client calls RopHardDeleteMessage to delete this message in the root public folder.
            object ropResponse = null;
            ulong[] messageIds = new ulong[] { messageId };

            RopHardDeleteMessagesRequest hardDeleteMessagesRequest = new RopHardDeleteMessagesRequest();
            RopHardDeleteMessagesResponse hardDeleteMessagesResponse = new RopHardDeleteMessagesResponse();
            hardDeleteMessagesRequest.RopId = (byte)RopId.RopHardDeleteMessages;
            hardDeleteMessagesRequest.LogonId = Constants.CommonLogonId;
            hardDeleteMessagesRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            hardDeleteMessagesRequest.WantAsynchronous = 0x00;

            // The server does not generate a non-read receipt for the deleted messages.
            hardDeleteMessagesRequest.NotifyNonRead = 0x00;
            hardDeleteMessagesRequest.MessageIdCount = (ushort)messageIds.Length;
            hardDeleteMessagesRequest.MessageIds = messageIds;
            this.Adapter.DoRopCall(hardDeleteMessagesRequest, this.publicRootFolderHandle, ref ropResponse, ref this.responseHandles);
            hardDeleteMessagesResponse = (RopHardDeleteMessagesResponse)ropResponse;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R99401");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R99401
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                hardDeleteMessagesResponse.ReturnValue,
                99401,
                @"[In RopHardDeleteMessages ROP] This ROP [RopHardDeleteMessages] applies to both public folders [and private mailboxes].");
            #endregion

            #region Step 4. The client calls RopGetContentsTable to retrieve the contents table of the root public folder.
            RopGetContentsTableRequest getContentsTableRequestSecond = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.Depth
            };

            // Get the Contents table of the opened folder.
            RopGetContentsTableResponse getContentsTableResponseSecond = this.Adapter.GetContentsTable(getContentsTableRequestSecond, this.publicRootFolderHandle, ref this.responseHandles);

            uint rowCountSecond = getContentsTableResponseSecond.RowCount;

            Assert.AreEqual<uint>(rowCountFirst - 1, rowCountSecond, "If ROP succeeds, the second RowCount value returned from server should be one less than the first one. ");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopGetHierarchyTable operation performs successfully in a public folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S04_TC11_RopGetHierarchyTableFromPublicFolderSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Logon();
            this.PublicFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root public folder.
            uint folderHandle1 = 0;
            ulong folderId1 = 0;
            this.CreateFolder(this.publicRootFolderHandle, Constants.Subfolder1, ref folderId1, ref folderHandle1);
            #endregion

            #region Step 2. The client creates a message in the public root folder.
            ulong messageId = 0;
            uint messageHandle = 0;
            this.CreateSaveMessage(this.publicRootFolderHandle, this.publicRootFolderId, ref messageId, ref messageHandle);
            #endregion

            #region Step 3. The client calls RopGetHierarchyTable retrieve the hierarchy table for the public root folder.
            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest
            {
                RopId = (byte)RopId.RopGetHierarchyTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = 0x04
            };
            RopGetHierarchyTableResponse getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, this.publicRootFolderHandle, ref this.responseHandles);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R100001");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R100001.
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                getHierarchyTableResponse.ReturnValue,
                100001,
                @"[In RopGetHierarchyTable ROP] The folder can be either a public folder [or a private mailbox folder].");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopGetContentsTable operation performs successfully in a public folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S04_TC12_RopGetContentsTableFromPublicFolder()
        {
            this.CheckWhetherSupportTransport();
            this.Logon();
            this.PublicFolderInitialization();

            #region Step 1. The client calls RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root public folder.
            uint folderHandle1 = 0;
            ulong folderId1 = 0;
            this.CreateFolder(this.publicRootFolderHandle, Constants.Subfolder1, ref folderId1, ref folderHandle1);
            #endregion

            #region Step 2. The client creates a message in the public root folder.
            ulong messageId = 0;
            uint messageHandle = 0;
            this.CreateSaveMessage(this.publicRootFolderHandle, this.publicRootFolderId, ref messageId, ref messageHandle);
            #endregion

            #region Step 3. The client calls RopGetContentsTable retrieve the contents table for the public root folder.
            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = 0xC8
            };
            RopGetContentsTableResponse getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, this.publicRootFolderHandle, ref this.responseHandles);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R100501");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R100501.
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                getContentsTableResponse.ReturnValue,
                100501,
                @"[In RopGetContentsTable ROP] This ROP [RopGetContentsTable] applies to both public folders [and private mailboxes].");
            #endregion
        }

        /// <summary>
        /// Test initialize. Overrides the method TestInitialize defined in base class.
        /// </summary>
        protected override void TestInitialize()
        {
            this.Adapter = Site.GetAdapter<IMS_OXCFOLDAdapter>();
            this.RootFolder = Common.GenerateResourceName(this.Site, Constants.RootFolder) + Constants.StringNullTerminated; 
        }

        /// <summary>
        /// Overrides TestClassBase's TestCleanup()
        /// </summary>
        protected override void TestCleanup()
        {
            if (this.NeedCleanup == true)
            {
                RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest
                {
                    RopId = (byte)RopId.RopDeleteFolder,
                    LogonId = Constants.CommonLogonId,
                    InputHandleIndex = Constants.CommonInputHandleIndex,
                    DeleteFolderFlags =
                        (byte)DeleteFolderFlags.DelMessages | (byte)DeleteFolderFlags.DelFolders |
                        (byte)DeleteFolderFlags.DeleteHardDelete,
                    FolderId = this.publicRootFolderId
                };
                RopDeleteFolderResponse deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.publicFoldersHandle, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(0, deleteFolderResponse.ReturnValue, "RopDeleteFolderResponse ROP operation performs successfully!");

                #region  Roprelease
                RopReleaseRequest releaseRequest = new RopReleaseRequest();
                object ropResponse = null;
                releaseRequest.RopId = (byte)RopId.RopRelease;
                releaseRequest.LogonId = Constants.CommonLogonId;
                releaseRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
                this.Adapter.DoRopCall(releaseRequest, this.LogonHandle, ref ropResponse, ref this.responseHandles);
                #endregion

                this.publicLogonHandle = 0;
                this.responseHandles = null;
                this.Adapter.DoDisconnect();
            }
        }

        /// <summary>
        /// Logon to Public Folder.
        /// </summary>
        private void Logon()
        {
            bool returnStatus = this.Adapter.DoConnect(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(returnStatus, "connection is successful");

            // Logon to Public Folder.
            this.logonResponse = this.Logon(LogonFlags.PublicFolder, out this.publicLogonHandle);
        }

        /// <summary>
        /// Initialize a generic folder under the Inbox folder as a root folder for test.
        /// </summary>
        private void PublicFolderInitialization()
        {
            #region Open the public folder.
            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderId = this.logonResponse.FolderIds[Constants.PublicFolderIndex],
                OpenModeFlags = (byte)FolderOpenModeFlags.None
            };

            // Use the logon object as input handle here.
            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, this.publicLogonHandle, ref this.responseHandles);
            this.publicFoldersHandle = this.responseHandles[0][openFolderResponse.OutputHandleIndex];
            #endregion

            #region Create a generic folder for test.
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
                DisplayName = Encoding.ASCII.GetBytes(this.RootFolder),
                Comment = Encoding.ASCII.GetBytes(this.RootFolder)
            };
            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.publicFoldersHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating Folder should succeed.");
            this.publicRootFolderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            this.publicRootFolderId = createFolderResponse.FolderId;
            #endregion
        }
    }
}