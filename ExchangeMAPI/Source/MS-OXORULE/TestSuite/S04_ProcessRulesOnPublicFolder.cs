namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Security.Policy;
    using System.Text;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario aims to validate server behaviors of:
    /// 1. The operations of RopModifyRules and RopGetRulesTable when adding, modifying, deleting and retrieving standard rules.
    /// 2. The functions of ROPs referenced from MS-OXCMSG and MS-OXCTABL for adding, modifying, deleting and retrieving extended rules.
    /// </summary>
    [TestClass]
    public class S04_ProcessRulesOnPublicFolder : TestSuiteBase
    {
        #region Test Class Initialization
        /// <summary>
        /// Use ClassInitialize to run code before running the first test in the class
        /// </summary>
        /// <param name="context">Context information associated with MS-OXORULE.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            TestClassBase.Initialize(context);
        }

        /// <summary>
        ///  Use ClassCleanup to run code after all tests in a class have run.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        /// <summary>
        /// This test case is designed to add, modify and delete standard rule on public folder.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S04_TC01_AddModifyDeleteStandardRule_OnPublicFolder()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region TestUser1 logs on to the public folder.
            RopOpenFolderResponse openFolderResponse;
            RopLogonResponse logonResponse;
            bool ret = this.OxoruleAdapter.Connect(ConnectionType.PublicFolderServer, this.User1Name, this.User1ESSDN, this.User1Password);
            Site.Assert.IsTrue(ret, "connect to public folder server should be successful");
            uint publicFolderLogonHandler = this.OxoruleAdapter.RopLogon(LogonType.PublicFolder, this.User1ESSDN, out logonResponse);

            // Assert the client to log on to the public folder successfully.
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "Logon the public folder should be successful.");

            // Folder index 1 is the Interpersonal Messages subtree, and this is defined in MS-OXCSTOR.
            uint publicfolderHandler = this.OxoruleAdapter.RopOpenFolder(publicFolderLogonHandler, logonResponse.FolderIds[1], out openFolderResponse);

            // Get the store object's entry ID.
            this.GetStoreObjectEntryID(StoreObjectType.PublicFolder, this.Server, this.User1ESSDN);

            RopCreateFolderResponse createFolderResponse;
            string newFolderName = Common.GenerateResourceName(this.Site, Constants.FolderDisplayName);
            uint newFolderHandle = this.OxoruleAdapter.RopCreateFolder(publicfolderHandler, newFolderName, Constants.FolderComment, out createFolderResponse);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating folder operation should succeed.");
            #endregion

            #region TestUser1 prepares value for rule properties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameDelete);
            RuleData ruleForDelete = AdapterHelper.GenerateValidRuleData(ActionType.OP_DELETE, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, new DeleteMarkReadActionData(), ruleProperties, null);
            #endregion

            #region TestUser1 adds rule OP_Delelte.
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(newFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForDelete });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Delete rule should succeed.");
            #endregion

            #region TestUser1 gets rule ID of the created rule.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(newFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");

            PropertyTag ruleIDTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleId,
                PropertyType = (ushort)PropertyType.PtypInteger64
            };
            RopQueryRowsResponse queryRowsResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, new PropertyTag[1] { ruleIDTag });
            Site.Assert.AreEqual<uint>(0, queryRowsResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");

            // Only one rule added in this folder, so the row count in the rule table should be 1.
            Site.Assert.AreEqual<uint>(1, queryRowsResponse.RowCount, "The rule number in the rule table is {0}", queryRowsResponse.RowCount);
            this.VerifyRuleTable();
            ulong ruleId = BitConverter.ToUInt64(queryRowsResponse.RowData.PropertyRows[0].PropertyValues[0].Value, 0);
            #endregion

            #region TestUser1 modifies the created rule.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, "RuleNameForModify");
            ruleForDelete = AdapterHelper.GenerateValidRuleData(ActionType.OP_DELETE, TestRuleDataType.ForModify, 1, RuleState.ST_ENABLED, new DeleteMarkReadActionData(), ruleProperties, ruleId);
            modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(newFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleForDelete });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Modifying the OP_DELETE rule should be success");
            #endregion

            #region TestUser1 calls RopGetRulesTable with valid TableFlags.
            ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(newFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            #endregion

            #region TestUser1 retrieves rule information for the modified rule.

            PropertyTag ruleNameTag = new PropertyTag { PropertyId = (ushort)PropertyId.PidTagRuleName, PropertyType = (ushort)PropertyType.PtypString };
            
            // Retrieves rows from the rule table.
            queryRowsResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, new PropertyTag[2] { ruleIDTag, ruleNameTag });
            Site.Assert.AreEqual<uint>(0, queryRowsResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");

            // Only one rule added in this folder, so the row count in the rule table should be 1.
            Site.Assert.AreEqual<uint>(1, queryRowsResponse.RowCount, "The rule number in the rule table is {0}", queryRowsResponse.RowCount);
            this.VerifyRuleTable();
            
            ulong ruleIdModified = BitConverter.ToUInt64(queryRowsResponse.RowData.PropertyRows[0].PropertyValues[0].Value, 0);
            bool isSameRuleId = ruleId == ruleIdModified;
            Site.Assert.IsTrue(isSameRuleId, "The original rule Id is {0} and the modified rule Id is {1}", ruleId, ruleIdModified);

            string modifiedRuleName = AdapterHelper.PropertyValueConvertToString(queryRowsResponse.RowData.PropertyRows.ToArray()[0].PropertyValues[1].Value);
            Site.Assert.AreEqual<string>(ruleProperties.Name, modifiedRuleName, "The actual rule name {0} should be equal to the expected rule name {1}.", modifiedRuleName, ruleProperties.Name);
            #endregion

            #region TestUser1 deletes the created rule.
            ruleForDelete = AdapterHelper.GenerateValidRuleData(ActionType.OP_DELETE, TestRuleDataType.ForRemove, 1, RuleState.ST_ENABLED, null, ruleProperties, ruleId);
            modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(newFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleForDelete });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Deleting the OP_DELETE rule should succeed.");
           
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(publicfolderHandler, createFolderResponse.FolderId);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to add, modify and delete extended rule on public folder.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S04_TC02_AddModifyDeleteExtendedRule_OnPublicFolder()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region TestUser1 logs on to the public folder.
            RopOpenFolderResponse openFolderResponse;
            RopLogonResponse logonResponse;
            bool ret = this.OxoruleAdapter.Connect(ConnectionType.PublicFolderServer, this.User1Name, this.User1ESSDN, this.User1Password);
            Site.Assert.IsTrue(ret, "connect to public folder server should be successful");
            uint publicFolderLogonHandler = this.OxoruleAdapter.RopLogon(LogonType.PublicFolder, this.User1ESSDN, out logonResponse);

            // Assert the client to log on to the public folder successfully.
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "Logon the public folder should be successful.");

            // Folder index 1 is the Interpersonal Messages subtree, and this is defined in MS-OXCSTOR.
            uint publicfolderHandler = this.OxoruleAdapter.RopOpenFolder(publicFolderLogonHandler, logonResponse.FolderIds[1], out openFolderResponse);

            // Get the store object's entry ID.
            this.GetStoreObjectEntryID(StoreObjectType.PublicFolder, this.Server, this.User1ESSDN);

            RopCreateFolderResponse createFolderResponse;
            string newFolderName = Common.GenerateResourceName(this.Site, Constants.FolderDisplayName);
            uint newFolderHandle = this.OxoruleAdapter.RopCreateFolder(publicfolderHandler, newFolderName, Constants.FolderComment, out createFolderResponse);
            ulong newFolderID = createFolderResponse.FolderId;
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating folder operation should succeed.");
            #endregion

            #region TestUser1 creates an FAI message.
            RopCreateMessageResponse ropCreateMessageResponse;
            uint extendedRuleMessageHandle = this.OxoruleAdapter.RopCreateMessage(newFolderHandle, newFolderID, Convert.ToByte(true), out ropCreateMessageResponse);
            Site.Assert.AreEqual<uint>(0, ropCreateMessageResponse.ReturnValue, "Creating the first FAI message should succeed.");
            #endregion

            #region TestUser1 adds the extended rule with NamedProperty successfully.
            string ruleConditionSubjectNameForAdd = Constants.RuleConditionSubjectContainString;
            NamedPropertyInfo namedPropertyInfo = new NamedPropertyInfo
            {
                NoOfNamedProps = 2,
                PropId = new uint[2]
                {
                    0x8001, 0x8002
                }
            };
            PropertyName testPropertyName = new PropertyName
            {
                Guid = System.Guid.NewGuid().ToByteArray(),
                Kind = 0x01,
                Name = Encoding.Unicode.GetBytes(Constants.NameOfPropertyName + "\0")
            };

            // 0x01 means the property is identified by the name property.
            testPropertyName.NameSize = (byte)testPropertyName.Name.Length;

            PropertyName secondPropertyName = new PropertyName
            {
                Guid = System.Guid.NewGuid().ToByteArray(),
                Kind = 0x00,
                LID = 0x88888888
            };

            // 0x00 means the property is identified by the LID.
            namedPropertyInfo.NamedProperty = new PropertyName[2] { testPropertyName, secondPropertyName };
            namedPropertyInfo.NamedPropertiesSize = (uint)(testPropertyName.Serialize().Length + secondPropertyName.Serialize().Length);
            string ruleName = Common.GenerateResourceName(this.Site, Constants.ExtendRulename1);
            TaggedPropertyValue[] extendedRulePropertiesForAdd = AdapterHelper.GenerateExtendedRuleTestData(ruleName, 0, (uint)RuleState.ST_ENABLED, Constants.PidTagRuleProvider, ActionType.OP_MARK_AS_READ, new DeleteMarkReadActionData(), ruleConditionSubjectNameForAdd, namedPropertyInfo);

            // Set properties for extended rule FAI message.
            RopSetPropertiesResponse ropSetPropertiesResponse = this.OxoruleAdapter.RopSetProperties(extendedRuleMessageHandle, extendedRulePropertiesForAdd);
            Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "Setting property for Extended rule FAI message should succeed.");

            // Save changes of message.
            RopSaveChangesMessageResponse ropSaveChangesMessagResponse = this.OxoruleAdapter.RopSaveChangesMessage(extendedRuleMessageHandle);
            Site.Assert.AreEqual<uint>(0, ropSaveChangesMessagResponse.ReturnValue, "Saving Extend rule message should succeed.");

            // Specify the properties to be got. 
            PropertyTag[] propertyTagArray = new PropertyTag[1];

            // PidTagRuleMessageProvider
            propertyTagArray[0].PropertyId = (ushort)PropertyId.PidTagRuleMessageProvider;
            propertyTagArray[0].PropertyType = (ushort)PropertyType.PtypString;

            // Get the specific properties of the extended rule.
            RopGetPropertiesSpecificResponse ropGetPropertiesSpecificResponse = this.OxoruleAdapter.RopGetPropertiesSpecific(extendedRuleMessageHandle, propertyTagArray);
            Site.Assert.AreEqual<uint>(0, ropGetPropertiesSpecificResponse.ReturnValue, "Getting specific properties operation should succeed.");
            string pidTagRuleMessageProviderData = AdapterHelper.PropertyValueConvertToString(ropGetPropertiesSpecificResponse.RowData.PropertyValues[0].Value);
            Site.Assert.AreEqual<string>(Constants.PidTagRuleProvider, pidTagRuleMessageProviderData, "The rule provider data should be RuleOrganizer.");
            #endregion

            #region Modify the created rule.
            ruleName = Common.GenerateResourceName(this.Site, Constants.ExtendRulename2);
            TaggedPropertyValue[] extendedRulePropertiesForModify = AdapterHelper.GenerateExtendedRuleTestData(ruleName, 0, (uint)RuleState.ST_ENABLED, Constants.PidTagRuleProvider, ActionType.OP_MARK_AS_READ, new DeleteMarkReadActionData(), ruleConditionSubjectNameForAdd, namedPropertyInfo);

            // Set properties for extended rule FAI message.
            ropSetPropertiesResponse = this.OxoruleAdapter.RopSetProperties(extendedRuleMessageHandle, extendedRulePropertiesForModify);
            Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "Setting property for Extended rule FAI message should succeed.");

            // Save changes of message.
            ropSaveChangesMessagResponse = this.OxoruleAdapter.RopSaveChangesMessage(extendedRuleMessageHandle);
            Site.Assert.AreEqual<uint>(0, ropSaveChangesMessagResponse.ReturnValue, "Saving Extend rule message should succeed.");

            // PidTagSubject
            propertyTagArray[0].PropertyId = (ushort)PropertyId.PidTagRuleMessageName;
            propertyTagArray[0].PropertyType = (ushort)PropertyType.PtypString;

            // Get the specific properties of the extended rule.
            ropGetPropertiesSpecificResponse = this.OxoruleAdapter.RopGetPropertiesSpecific(extendedRuleMessageHandle, propertyTagArray);
            Site.Assert.AreEqual<uint>(0, ropGetPropertiesSpecificResponse.ReturnValue, "Getting specific properties operation should succeed.");
            string messageName = AdapterHelper.PropertyValueConvertToString(ropGetPropertiesSpecificResponse.RowData.PropertyValues[0].Value);
            Site.Assert.AreEqual<string>(ruleName, messageName, "The rule subject should be {0}.", ruleName);
            #endregion

            #region Release the created message to delete the created rule.
            this.OxoruleAdapter.ReleaseRop(extendedRuleMessageHandle);

            // Get the specific properties of the extended rule.
            ropGetPropertiesSpecificResponse = this.OxoruleAdapter.RopGetPropertiesSpecific(extendedRuleMessageHandle, propertyTagArray);
            Site.Assert.IsNull(ropGetPropertiesSpecificResponse.RowData, "The property value of the extended rule should be null!");
            #endregion

            #region Delete the folder.
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(publicfolderHandler, newFolderID);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_TAG rule.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S04_TC03_ServerExecuteRule_Action_OP_TAG_OnPublicFolder()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region TestUser1 logs on to the public folder.
            RopOpenFolderResponse openFolderResponse;
            RopLogonResponse logonResponse;
            bool ret = this.OxoruleAdapter.Connect(ConnectionType.PublicFolderServer, this.User1Name, this.User1ESSDN, this.User1Password);
            Site.Assert.IsTrue(ret, "connect to public folder server should be successful");
            uint publicFolderLogonHandler = this.OxoruleAdapter.RopLogon(LogonType.PublicFolder, this.User1ESSDN, out logonResponse);

            // Assert the client to log on to the public folder successfully.
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "Logon the public folder should be successful.");

            // Folder index 1 is the Interpersonal Messages subtree, and this is defined in MS-OXCSTOR.
            uint publicfolderHandler = this.OxoruleAdapter.RopOpenFolder(publicFolderLogonHandler, logonResponse.FolderIds[1], out openFolderResponse);

            // Get the store object's entry ID.
            this.GetStoreObjectEntryID(StoreObjectType.PublicFolder, this.Server, this.User1ESSDN);

            RopCreateFolderResponse createFolderResponse;
            string newFolderName = Common.GenerateResourceName(this.Site, Constants.FolderDisplayName);
            uint newFolderHandle = this.OxoruleAdapter.RopCreateFolder(publicfolderHandler, newFolderName, Constants.FolderComment, out createFolderResponse);
            ulong newFolderID = createFolderResponse.FolderId;
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating folder operation should succeed.");
            #endregion

            #region TestUser1 prepares value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameTag);
            #endregion

            #region TestUser1 adds an OP_TAG rule to the new created folder.
            TagActionData tagActionData = new TagActionData();
            PropertyTag tagActionDataPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagImportance,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            tagActionData.PropertyTag = tagActionDataPropertyTag;
            tagActionData.PropertyValue = BitConverter.GetBytes(1);

            RuleData ruleOpTag = AdapterHelper.GenerateValidRuleData(ActionType.OP_TAG, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, tagActionData, ruleProperties, null);
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(newFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleOpTag });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding OP_TAG rule should succeed.");
            #endregion

            #region TestUser1 creates a message.
            RopCreateMessageResponse ropCreateMessageResponse;
            uint messageHandle = this.OxoruleAdapter.RopCreateMessage(newFolderHandle, newFolderID, Convert.ToByte(false), out ropCreateMessageResponse);
            Site.Assert.AreEqual<uint>(0, ropCreateMessageResponse.ReturnValue, "Creating message should succeed.");
            #endregion

            #region TestUser1 saves the subject property of the message to trigger the rule.
            string subjectName = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            TaggedPropertyValue subjectProperty = new TaggedPropertyValue();
            PropertyTag pidTagSubjectPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagSubject,
                PropertyType = (ushort)PropertyType.PtypString
            };
            subjectProperty.PropertyTag = pidTagSubjectPropertyTag;
            subjectProperty.Value = Encoding.Unicode.GetBytes(subjectName + "\0");

            // Set properties for the created message to trigger the rule.
            RopSetPropertiesResponse ropSetPropertiesResponse = this.OxoruleAdapter.RopSetProperties(messageHandle, new TaggedPropertyValue[] { subjectProperty });
            Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "Setting property for the created message should succeed.");

            // Save changes of message.
            RopSaveChangesMessageResponse ropSaveChangesMessagResponse = this.OxoruleAdapter.RopSaveChangesMessage(messageHandle);
            Site.Assert.AreEqual<uint>(0, ropSaveChangesMessagResponse.ReturnValue, "Saving the created message should succeed.");

            // Wait for the message to be saved and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message and its properties to verify the rule evaluation.
            PropertyTag[] propertyTagList = new PropertyTag[2];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagImportance;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypInteger32;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;

            uint contentTableHandle = 0;
            uint rowCount = 0;
            RopQueryRowsResponse getMailMessageContent = this.GetExpectedMessage(newFolderHandle, ref contentTableHandle, propertyTagList, ref rowCount, 1, subjectName);
            Site.Assert.AreEqual<uint>(1, rowCount, @"The message number in the specific folder should be 1.");
            Site.Assert.AreEqual<int>(1, BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[(int)rowCount - 1].PropertyValues[0].Value, 0), "If the rule is executed, the PidTagImportance property of the message should be the value set by the rule.");

            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(publicfolderHandler, newFolderID);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to test the operation for adding unsupported rules to the public folder. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S04_TC04_AddNotSupportedRule_OnPublicFolder()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region TestUser2 logs on to the public folder.
            RopOpenFolderResponse openFolderResponse;
            RopLogonResponse logonResponse;
            bool ret = this.OxoruleAdapter.Connect(ConnectionType.PublicFolderServer, this.User2Name, this.User2ESSDN, this.User2Password);
            Site.Assert.IsTrue(ret, "connect to public folder server should be successful");
            uint publicFolderLogonHandler = this.OxoruleAdapter.RopLogon(LogonType.PublicFolder, this.User2ESSDN, out logonResponse);

            // Assert the client to log on to the public folder successfully.
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "Logon the public folder should be successful.");

            // Folder index 1 is the Interpersonal Messages subtree, and this is defined in MS-OXCSTOR.
            uint publicfolderHandler = this.OxoruleAdapter.RopOpenFolder(publicFolderLogonHandler, logonResponse.FolderIds[1], out openFolderResponse);
            ulong publicFolderID = logonResponse.FolderIds[1];
            #endregion

            #region TestUser2 tests the unsupported rules in public folder.
            #region TestUser2 prepares value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameDeferredAction);
            #endregion

            #region TestUser2 adds a Rule set Action Type to OP_DEFER_ACTION.
            DeferredActionData deferredActionData = new DeferredActionData
            {
                Data = Common.GetBytesFromBinaryHexString(Constants.DeferredActionBufferData)
            };
            RuleData deferredActionRuleData = AdapterHelper.GenerateValidRuleData(ActionType.OP_DEFER_ACTION, TestRuleDataType.ForAdd, 5, RuleState.ST_ENABLED, deferredActionData, ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(publicfolderHandler, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { deferredActionRuleData });
            #endregion

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R258: server returned value is {0}", ropModifyRulesResponse.ReturnValue);

            // Verify MS-OXORULE requirement: MS-OXORULE_R258.
            // If the ReturnValue is not 0x00000000, it means the server failed to add a public folder rule with this OP_DEFER_ACTION Type action.
            // So it is not used in a public folder rule.
            bool isVerifyR258 = ropModifyRulesResponse.ReturnValue != 0x00000000;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR258,
                258,
                @"[In ActionBlock Structure] The meaning of action type OP_DEFER_ACTION: MUST NOT be used in a public folder rule (2).");

            #region TestUser2 prepares rules' data.
            MoveCopyActionData moveCopyActionData = new MoveCopyActionData();

            // Get the Inbox folder entry ID.
            byte[] folderEId = this.OxoruleAdapter.GetFolderEntryId(StoreObjectType.PublicFolder, publicfolderHandler, publicFolderID);

            // Get the store object's entry ID.
            byte[] storeEId = this.GetStoreObjectEntryID(StoreObjectType.PublicFolder, this.Server, this.User1ESSDN);
            moveCopyActionData.FolderEID = folderEId;
            moveCopyActionData.StoreEID = storeEId;
            moveCopyActionData.FolderEIDSize = (ushort)folderEId.Length;
            moveCopyActionData.StoreEIDSize = (ushort)storeEId.Length;
            #endregion

            #region TestUser2 prepares value for ruleProperties variable.
            ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameMoveOne);
            #endregion

            #region TestUser2 adds OP_MOVE rule to the public folder.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameMoveOne);
            RuleData ruleForMove = AdapterHelper.GenerateValidRuleData(ActionType.OP_MOVE, TestRuleDataType.ForAdd, 0, RuleState.ST_ENABLED, moveCopyActionData, ruleProperties, null);
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(publicfolderHandler, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForMove });
            #endregion

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R839: server returned value is {0}", modifyRulesResponse.ReturnValue);

            // Verify MS-OXORULE requirement: MS-OXORULE_R839.
            // If the returnValue is not 0x00000000, it means this OP_MOVE Type used in a public folder rule is failed,
            // so it is not used in a public folder rule.
            bool isVerifyR839 = modifyRulesResponse.ReturnValue != 0x00000000;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR839,
                839,
                @"[In ActionBlock Structure] The meaning of action type OP_MOVE: MUST NOT be used in a public folder rule (2).");

            #region TestUser2 adds OP_COPY rule to the public folder.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameMoveTwo);
            RuleData ruleForCopy = AdapterHelper.GenerateValidRuleData(ActionType.OP_COPY, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, moveCopyActionData, ruleProperties, null);
            modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(publicfolderHandler, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForCopy });
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R840: server returned value is {0}", modifyRulesResponse.ReturnValue);

            // Verify MS-OXORULE requirement: MS-OXORULE_R840.
            // If the ReturnValue is not 0x00000000, it means this OP_COPY Type used in a public folder rule is failed,
            // so it is not used in a public folder rule.
            bool isVerifyR840 = modifyRulesResponse.ReturnValue != 0x00000000;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR840,
                840,
                @"[In ActionBlock Structure] The meaning of action type OP_COPY: MUST NOT be used in a public folder rule (2).");
            #endregion
        }
    }
}