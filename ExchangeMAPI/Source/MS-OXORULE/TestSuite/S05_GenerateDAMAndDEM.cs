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
    /// This scenario aims to validate server behaviors about DAM and DEM message. 
    /// </summary>
    [TestClass]
    public class S05_GenerateDAMAndDEM : TestSuiteBase
    {
        #region Test Class Initialization
        /// <summary>
        /// Use ClassInitialize to run code before running the first test in the class.
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
        /// This test case is designed for server to generate one DAM (Deferred Action Message) message in DAF (Deferred Action Folder) folder when there are more than one OP_DEFER_ACTION actions that belong to the same rule provider. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S05_TC01_ServerGenerateOneDAM_ForOP_DEFER_ACTION_BelongToSameRuleProvider()
        {
            this.CheckMAPIHTTPTransportSupported();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(908, this.Site), "This case runs only when only one DAM message is generated in DAF folder when more than one OP_DEFER_ACTION actions belong to the same rule provider.");

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.DAMPidTagRuleNameOne);
            #endregion

            #region TestUser1 creates 2 new rules which can trigger server to generate DAM. The 2 new rules have the same rule provider and the same rule condition.
            // If the action type is "OP_DEFER_ACTION", the ActionData buffer is completely under the control of the client that created the rule.
            // When a message that satisfies the rule condition is received, the server creates a DAM
            // and places the entire content of the ActionBlock field as part of the PidTagClientActions property on the DAM.
            DeferredActionData deferredActionData1SetByClient = new DeferredActionData
            {
                Data = Encoding.Unicode.GetBytes(Constants.DAMPidTagRuleActionOne + "\0")
            };
            DeferredActionData deferredActionData2SetByClient = new DeferredActionData
            {
                Data = Encoding.Unicode.GetBytes(Constants.DAMPidTagRuleActionTwo + "\0")
            };
            ActionType actionType = ActionType.OP_DEFER_ACTION;

            RuleData ruleData1SetByClient = AdapterHelper.GenerateValidRuleData(actionType, TestRuleDataType.ForAdd, 0, RuleState.ST_ENABLED, deferredActionData1SetByClient, ruleProperties, null);
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.DAMPidTagRuleNameTwo);
            RuleData ruleData2SetByClient = AdapterHelper.GenerateValidRuleData(actionType, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, deferredActionData2SetByClient, ruleProperties, null);

            // Call RopModifyRules.
            this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleData1SetByClient, ruleData2SetByClient });
            #endregion

            #region TestUser1 calls RopGetRulesTable.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");
            #endregion

            #region TestUser1 retrieves rule information for the newly added rule.
            PropertyTag[] propertyAllTags = new PropertyTag[]
            {
                new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleName,
                    PropertyType = (ushort)PropertyType.PtypString
                },
                new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleId,
                    PropertyType = (ushort)PropertyType.PtypInteger64
                },
                new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleSequence,
                    PropertyType = (ushort)PropertyType.PtypInteger32
                }
            };

            // Retrieves rows from the rule table.
            RopQueryRowsResponse queryRowResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyAllTags);
            Site.Assert.AreEqual<uint>(0, queryRowResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");

            // Add two rules to the Inbox folder, so if get rule table successfully and the rule count is 2,
            // it means the server is returning a table with the rules added by the test suite.
            Site.Assert.AreEqual(2, queryRowResponse.RowCount, @"There should be 2 rules returned, actual returned row count is {0}.", queryRowResponse.RowCount);
            byte[] pidTagRuleId1 = null;
            if (AdapterHelper.PropertyValueConvertToUint(queryRowResponse.RowData.PropertyRows[0].PropertyValues[2].Value) == 0)
            {
                pidTagRuleId1 = queryRowResponse.RowData.PropertyRows[0].PropertyValues[1].Value;
            }
            else
            {
                pidTagRuleId1 = queryRowResponse.RowData.PropertyRows[1].PropertyValues[1].Value;
            }
            this.VerifyRuleTable();
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these 2 rules.
            string deliveredMessageSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser2 delivers a message to TestUser1 to trigger these 2 rules
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, deliveredMessageSubject);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 checks the generation of DAM, which is placed under DAF folder.
            #region TestUser1 sets the interested columns of the message table in the DAF folder.
            PropertyTag[] propertyTagOfDAM = new PropertyTag[9];
            propertyTagOfDAM = AdapterHelper.GenerateRuleInfoPropertiesOfDAM();

            // Query rows include the property values of the interested columns.
            uint contentsTableHandleOfDAF = 0;
            uint rowCount = 0;
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForDAM;
            RopQueryRowsResponse ropQueryRowsResponseOfDAM = this.GetExpectedMessage(this.DAFFolderHandle, ref contentsTableHandleOfDAF, propertyTagOfDAM, ref rowCount);
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.OtherTarget;
            bool isDAMUnderDAFFolder = rowCount > 0;
            #endregion

            #region TestUser1 verifies the properties' values contained in RopQueryRowsResponse for the generated DAM messages.

            // In this test case, there is only one row returned in the PropertyRows buffer, which represents one generated DAM message. 
            // And there are 9 interested properties for the DAM message returned in the PropertyValues buffer.
            // The returned property values' order for the row is the same with the order they are set through RopSetColumns.
            byte[] pidTagClientActionsOfDAMOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[0].PropertyValues[5].Value;

            // actionDataForRuleOneSetByClient represents the OP_DEFER_ACTION action in the Rule One which has been added in this test case.
            ActionBlock actionDataForRuleOneSetByClient = new ActionBlock(CountByte.TwoBytesCount)
            {
                ActionType = ActionType.OP_DEFER_ACTION,
                ActionFlavor = 0x00000000,
                ActionFlags = 0x00000000,
                ActionDataValue = deferredActionData1SetByClient
            };
            actionDataForRuleOneSetByClient.ActionLength = actionDataForRuleOneSetByClient.ActionDataValue.Size() + 9;

            // actionDataForRuleTwoSetByClient represents the OP_DEFER_ACTION action in the Rule Two which has been added in this test case.
            ActionBlock actionDataForRuleTwoSetByClient = new ActionBlock(CountByte.TwoBytesCount)
            {
                ActionType = ActionType.OP_DEFER_ACTION,
                ActionFlavor = 0x00000000,
                ActionFlags = 0x00000000,
                ActionDataValue = deferredActionData2SetByClient
            };
            actionDataForRuleTwoSetByClient.ActionLength = actionDataForRuleTwoSetByClient.ActionDataValue.Size() + 9;

            // Pack the information about the two OP_DEFER_ACTION actions above into one RuleAction structure.
            RuleAction twoPackedRuleActionsSetByClient = new RuleAction(CountByte.TwoBytesCount)
            {
                NoOfActions = 0x0002,
                Actions = new ActionBlock[2]
            };
            twoPackedRuleActionsSetByClient.Actions[0] = actionDataForRuleOneSetByClient;
            twoPackedRuleActionsSetByClient.Actions[1] = actionDataForRuleTwoSetByClient;

            // pidTagClientActionsOfDAM represents the pidTagClientActions property set on the DAM message. 
            // This property contains the relevant actions which need to be further processed by the client.
            RuleAction pidTagClientActionsOfDAM = AdapterHelper.PropertyValueConvertToRuleAction(pidTagClientActionsOfDAMOfBytes);

            // Verify MS-OXORULE requirement: MS-OXORULE_R908
            // This test case is designed based on that there are more than one OP_DEFER_ACTION actions that belong to the same rule provider.
            // That RowCount of ropQueryRowsResponseOfDAM equals 1 means there is only one DAM generated.
            // That twoPackedRuleActionsSetByClient equals to pidTagClientActionsOfDAM means the server indeed packed 
            // the information about the two OP_DEFER_ACTION actions into one DAM on the property pidTagClientActions.
            bool isVerifyR908 = (ropQueryRowsResponseOfDAM.RowCount == 1) && Common.CompareByteArray(twoPackedRuleActionsSetByClient.Serialize(), pidTagClientActionsOfDAM.Serialize());

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R908: RowCount of ropQueryRowsResponseOfDAM is {0}, twoPackedRuleActionsSetByClient is {1}, pidTagClientActionsOfDAM is {2}", ropQueryRowsResponseOfDAM.RowCount, twoPackedRuleActionsSetByClient, pidTagClientActionsOfDAM);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR908,
                908,
                @"[In Generating a DAM] The implementation does this [pack information about more than one ""OP_DEFER_ACTION"" actions (2) for any given message into one DAM] when there are more than one ""OP_DEFER_ACTION"" actions (2) that belong to the same rule provider. (Exchange 2003 and above follow this behavior.)");

            byte[] pidTagRuleIds = AdapterHelper.PropertyValueConvertToBinary(ropQueryRowsResponseOfDAM.RowData.PropertyRows[0].PropertyValues[6].Value);

            System.Collections.Generic.List<byte> temp = new System.Collections.Generic.List<byte>();
            temp.AddRange(pidTagRuleId1);
            temp.AddRange(pidTagRuleId1);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R368");

            // Verify MS-OXORULE requirement: MS-OXORULE_R368
            Site.CaptureRequirementIfIsTrue(
                Common.CompareByteArray(temp.ToArray(), pidTagRuleIds),
                368,
                @"[In PidTagRuleIds] The PidTagRuleIds property ([MS-OXPROPS] section 2.941) is a buffer contains the PidTagRuleId (section 2.2.1.3.1.1) value (8 bytes) from the first rules (2) that contributed actions (2) in the PidTagClientActions property (section 2.2.6.6), and repeats that value once for each rule (2) that contributed actions (2).");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed for server to generate separate DAM messages in DAF folder when the OP_DEFER_ACTION actions belong to separate rule providers.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S05_TC02_ServerGenerateSeparateDAM_ForOP_DEFER_ACTION_BelongToSeparateRuleProvider()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.DAMPidTagRuleNameOne);
            #endregion

            #region TestUser1 gets a DAF message before generating a DAM.
            #region Step1: TestUser1 calls RopGetContentsTable to get a table of all messages which are placed in the DAF folder.
            uint contentsTableHandleOfDAFBeforeGenerateDAM;

            // Call RopGetContentsTable.
            RopGetContentsTableResponse ropGetContentsTableResponseOfDAFBeforeGenerateDAM = this.OxoruleAdapter.RopGetContentsTable(this.DAFFolderHandle, ContentTableFlag.None, out contentsTableHandleOfDAFBeforeGenerateDAM);
            Site.Assert.AreEqual<uint>(0, ropGetContentsTableResponseOfDAFBeforeGenerateDAM.ReturnValue, "Getting DAF contents table should succeed");
            #endregion

            #region Step2: TestUser1 calls RopSetColumns to set the interested columns of the message table in the DAF folder.

            // Here are 10 interested columns listed as below.
            // Prepare the data in the RopSetColumns request buffer.
            PropertyTag[] propertyTagOfDAMBeforeGenerateDAM = new PropertyTag[9];
            propertyTagOfDAMBeforeGenerateDAM = AdapterHelper.GenerateRuleInfoPropertiesOfDAM();

            // Query rows which include the property values of the interested columns. 
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForDAM;
            RopQueryRowsResponse ropQueryRowsResponseOfDAMBeforeGenerateDAM = this.OxoruleAdapter.QueryPropertiesInTable(contentsTableHandleOfDAFBeforeGenerateDAM, propertyTagOfDAMBeforeGenerateDAM);
            Site.Assert.AreEqual<ushort>(0, ropQueryRowsResponseOfDAMBeforeGenerateDAM.RowCount, "There should be 0 DAM generated in the DAF folder");
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.OtherTarget;
            #endregion
            #endregion

            #region TestUser1 creates two new rules which can trigger server to generate DAM. The two new rules have different rule providers and the same rule condition.
            // If the action type is "OP_DEFER_ACTION", the ActionData buffer is completely under the control of the client that created the rule.
            // When a message that satisfies the rule condition is received, the server creates a DAM
            // and places the entire content of the ActionBlock field as part of the PidTagClientActions property on the DAM.
            DeferredActionData deferredActionData1SetByClient = new DeferredActionData
            {
                Data = Encoding.Unicode.GetBytes(Constants.DAMPidTagRuleActionOne + "\0")
            };
            DeferredActionData deferredActionData2SetByClient = new DeferredActionData
            {
                Data = Encoding.Unicode.GetBytes(Constants.DAMPidTagRuleActionTwo + "\0")
            };

            ruleProperties.Provider = Constants.DAMPidTagRuleProviderOne;
            RuleData ruleData1SetByClient = AdapterHelper.GenerateValidRuleData(ActionType.OP_DEFER_ACTION, TestRuleDataType.ForAdd, 0, RuleState.ST_ENABLED, deferredActionData1SetByClient, ruleProperties, null);

            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.DAMPidTagRuleNameTwo);
            ruleProperties.Provider = Constants.DAMPidTagRuleProviderTwo;
            RuleData ruleData2SetByClient = AdapterHelper.GenerateValidRuleData(ActionType.OP_DEFER_ACTION, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, deferredActionData2SetByClient, ruleProperties, null);

            // Call RopModifyRules.
            this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleData1SetByClient, ruleData2SetByClient });
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these two rules.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser2 delivers a message to trigger these two rules.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);
            #endregion

            #region TestUser1 calls GetNotifyResponse to check whether it has got notification.
            if (Common.IsRequirementEnabled(899, this.Site))
            {
                // Get notification detail from server.
                RopNotifyResponse ropNotifyResponse = this.GetNotifyResponse();

                // Verify requirement MS-OXORULE_R899
                // If the notification data got from the server isn't null, it means the DAF supports notification, so this requirement can be verified.
                bool isVerifyR899 = ropNotifyResponse.NotificationData != null;

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R899.");

                // Verify MS-OXORULE requirement: MS-OXORULE_R899.
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR899,
                    899,
                    @"[In Initialization] The DAF does support notifications on its contents table object on the implementation, as specified in [MS-OXCNOTIF]. (Exchange 2003 and above follow this behavior.)");
            }
            #endregion

            #region TestUser1 gets Message's entry ID.
            #region Step1: TestUser1 gets the message handle and message ID.

            // Prepare the data in the RopSetColumns request buffer
            PropertyTag[] propertyTagOfInboxFolder = new PropertyTag[3];
            propertyTagOfInboxFolder[0].PropertyId = (ushort)PropertyId.PidTagMid;
            propertyTagOfInboxFolder[0].PropertyType = (ushort)PropertyType.PtypInteger64;
            propertyTagOfInboxFolder[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagOfInboxFolder[1].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagOfInboxFolder[2].PropertyId = (ushort)PropertyId.PidTagHasDeferredActionMessages;
            propertyTagOfInboxFolder[2].PropertyType = (ushort)PropertyType.PtypBoolean;

            // Query rows which include the property values of the interested columns. 
            uint contentsTableHandleOfInboxFolder = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse ropQueryRowsResponseOfInboxFolder = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandleOfInboxFolder, propertyTagOfInboxFolder, ref expectedMessageIndex, mailSubject);

            ulong messageId = AdapterHelper.PropertyValueConvertToUint64(ropQueryRowsResponseOfInboxFolder.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);
            RopOpenMessageResponse ropOpenMessageResponse = new RopOpenMessageResponse();

            // Open the message to get the message handle.
            uint messagehandle = this.OxoruleAdapter.RopOpenMessage(this.InboxFolderHandle, this.InboxFolderID, messageId, out ropOpenMessageResponse);

            // Subject, bodyText and originalMessageSender are the properties set by the server on the replied message.
            string subject = AdapterHelper.PropertyValueConvertToString(ropQueryRowsResponseOfInboxFolder.RowData.PropertyRows[expectedMessageIndex].PropertyValues[1].Value);
            #endregion

            #region Step2: TestUser1 gets the message entry ID and the Inbox folder's entry ID.

            // Get message's entry ID.
            byte[] entryId = this.OxoruleAdapter.GetMessageEntryId(this.InboxFolderHandle, this.InboxFolderID, messagehandle, messageId);
            #endregion
            #endregion

            #region TestUser1 checks the generation of DAM which is placed under DAF folder. There should be 2 separate DAMs generated.
            #region Step1: TestUser1 calls RopSetColumns to set the interested columns of the message table in the DAF folder.

            // Prepare the data in the RopSetColumns request buffer.
            PropertyTag[] propertyTagOfDAM = new PropertyTag[6];
            propertyTagOfDAM[0].PropertyId = (ushort)PropertyId.PidTagMessageClass;
            propertyTagOfDAM[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagOfDAM[1].PropertyId = (ushort)PropertyId.PidTagRuleProvider;
            propertyTagOfDAM[1].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagOfDAM[2].PropertyId = (ushort)PropertyId.PidTagClientActions;
            propertyTagOfDAM[2].PropertyType = (ushort)PropertyType.PtypBinary;
            propertyTagOfDAM[3].PropertyId = (ushort)PropertyId.PidTagDeferredActionMessageOriginalEntryId;
            propertyTagOfDAM[3].PropertyType = (ushort)PropertyType.PtypServerId;
            propertyTagOfDAM[4].PropertyId = (ushort)PropertyId.PidTagMid;
            propertyTagOfDAM[4].PropertyType = (ushort)PropertyType.PtypInteger64;
            propertyTagOfDAM[5].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagOfDAM[5].PropertyType = (ushort)PropertyType.PtypString;

            // Query rows which include the property values of the interested columns.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForDAM;
            uint contentsTableHandleOfDAF = 0;
            uint rowCount = 0;
            RopQueryRowsResponse ropQueryRowsResponseOfDAM = this.GetExpectedMessage(this.DAFFolderHandle, ref contentsTableHandleOfDAF, propertyTagOfDAM, ref rowCount);
            Site.Assert.AreEqual<uint>(0, ropQueryRowsResponseOfDAM.ReturnValue, "Query rows operation should succeed.");
            Site.Assert.AreEqual<ushort>(2, ropQueryRowsResponseOfDAM.RowCount, "There should be 2 DAMs generated in the DAF folder");
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.OtherTarget;
            #endregion

            #region Step2: TestUser1 verifies the properties' values contained in RopQueryRowsResponse for the 2 generated DAM messages.

            // In this test case, there are 2 rows returned in the PropertyRows buffer, which represent the 2 generated DAM messages. 
            // And there are 3 interested properties for each DAM message returned in the PropertyValues buffer.
            // The returned property values' order in each row is the same with the order they are set through RopSetColumns.
            string pidTagMessageClassOfDAMOne = string.Empty;
            string pidTagRuleProviderOfDAMOne = string.Empty;
            RuleAction pidTagClientActionsOfDAMOne = new RuleAction();
            byte[] pidTagDeferredActionMessageOriginalEntryIdOfDAMOne = null;
            byte[] pidTagMIDOfDAMOneOfBytes = null;
            string pidTagMessageClassOfDAMTwo = string.Empty;
            string pidTagRuleProviderOfDAMTwo = string.Empty;
            RuleAction pidTagClientActionsOfDAMTwo = new RuleAction();
            byte[] pidTagMIDOfDAMTwoOfBytes = null;
            string ruleProviderOneSetByClient = Constants.DAMPidTagRuleProviderOne;
            string ruleProviderTwoSetByClient = Constants.DAMPidTagRuleProviderTwo;
            for (int i = 0; i < ropQueryRowsResponseOfDAM.RowCount; i++)
            {
                byte[] pidTagRuleProviderOfDAMOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[i].PropertyValues[1].Value;

                // pidTagRuleProviderOfDAM is set to the same value as the PidTagRuleProvider property on the rule set by client that triggers to generate the DAM.
                string pidTagRuleProviderOfDAM = AdapterHelper.PropertyValueConvertToString(pidTagRuleProviderOfDAMOfBytes);

                // This DAM is generated for Rule One
                if (pidTagRuleProviderOfDAM == ruleProviderOneSetByClient)
                {
                    byte[] pidTagMessageClassOfDAMOneOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[i].PropertyValues[0].Value;
                    pidTagMessageClassOfDAMOne = AdapterHelper.PropertyValueConvertToString(pidTagMessageClassOfDAMOneOfBytes);

                    byte[] pidTagRuleProviderOfDAMOneOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[i].PropertyValues[1].Value;
                    pidTagRuleProviderOfDAMOne = AdapterHelper.PropertyValueConvertToString(pidTagRuleProviderOfDAMOneOfBytes);

                    byte[] pidTagClientActionsOfDAMOneOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[i].PropertyValues[2].Value;
                    pidTagClientActionsOfDAMOne = AdapterHelper.PropertyValueConvertToRuleAction(pidTagClientActionsOfDAMOneOfBytes);

                    byte[] pidTagDamOriginalEntryIdOfDAMOneOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[i].PropertyValues[3].Value;
                    pidTagDeferredActionMessageOriginalEntryIdOfDAMOne = AdapterHelper.PropertyValueConvertToBinary(pidTagDamOriginalEntryIdOfDAMOneOfBytes);

                    pidTagMIDOfDAMOneOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[i].PropertyValues[4].Value;
                }
                else if (pidTagRuleProviderOfDAM == ruleProviderTwoSetByClient)
                {
                    // This DAM is generated for Rule Two.
                    byte[] pidTagMessageClassOfDAMTwoOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[i].PropertyValues[0].Value;
                    pidTagMessageClassOfDAMTwo = AdapterHelper.PropertyValueConvertToString(pidTagMessageClassOfDAMTwoOfBytes);

                    byte[] pidTagRuleProviderOfDAMTwoOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[i].PropertyValues[1].Value;
                    pidTagRuleProviderOfDAMTwo = AdapterHelper.PropertyValueConvertToString(pidTagRuleProviderOfDAMTwoOfBytes);

                    byte[] pidTagClientActionsOfDAMTwoOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[i].PropertyValues[2].Value;
                    pidTagClientActionsOfDAMTwo = AdapterHelper.PropertyValueConvertToRuleAction(pidTagClientActionsOfDAMTwoOfBytes);

                    byte[] pidTagDamOriginalEntryIdOfDAMTwoOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[i].PropertyValues[3].Value;
                    AdapterHelper.PropertyValueConvertToBinary(pidTagDamOriginalEntryIdOfDAMTwoOfBytes);

                    pidTagMIDOfDAMTwoOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[i].PropertyValues[4].Value;
                }
            }

            // Reconstruct the 2 ruleActions set by client which should equal PidTagClientActions property of each DAM.
            ActionBlock[] actionBlockForRuleOneSetByClient = new ActionBlock[1];
            actionBlockForRuleOneSetByClient[0] = new ActionBlock(CountByte.TwoBytesCount)
            {
                ActionType = ActionType.OP_DEFER_ACTION,
                ActionFlavor = 0x00000000,
                ActionFlags = 0x00000000,
                ActionDataValue = deferredActionData1SetByClient
            };
            actionBlockForRuleOneSetByClient[0].ActionLength = actionBlockForRuleOneSetByClient[0].ActionDataValue.Size() + 9;

            RuleAction ruleActionForRuleOneSetByClient = new RuleAction(CountByte.TwoBytesCount)
            {
                NoOfActions = 0x0001,
                Actions = actionBlockForRuleOneSetByClient
            };

            ActionBlock[] actionBlockForRuleTwoSetByClient = new ActionBlock[1];
            actionBlockForRuleTwoSetByClient[0] = new ActionBlock(CountByte.TwoBytesCount)
            {
                ActionType = ActionType.OP_DEFER_ACTION,
                ActionFlavor = 0x00000000,
                ActionFlags = 0x00000000,
                ActionDataValue = deferredActionData2SetByClient
            };
            actionBlockForRuleTwoSetByClient[0].ActionLength = actionBlockForRuleTwoSetByClient[0].ActionDataValue.Size() + 9;

            RuleAction ruleActionForRuleTwoSetByClient = new RuleAction(CountByte.TwoBytesCount)
            {
                NoOfActions = 0x0001,
                Actions = actionBlockForRuleTwoSetByClient
            };

            // Add the debug information.
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXORULE_R571: the value of RowCount is {0}, pidTagRuleProviderOfDAMOne is {1}, pidTagRuleProviderOfDAMTwo is {2}",
                ropQueryRowsResponseOfDAM.RowCount,
                pidTagRuleProviderOfDAMOne,
                pidTagRuleProviderOfDAMTwo);

            // Verify MS-OXORULE requirement: MS-OXORULE_R571.
            // That the RowCount of the ropQueryRowsResponseOfDAM equals 2 and the property values of pidTagMessageClass returned for each message under the DAF folder
            // are both "IPC.Microsoft Exchange 4.0.Deferred action" means there are 2 separate DAMs generated.
            // The RuleProvider value returned in pidTagRuleProvider property on each DAM must equal each RuleProvider set by client.
            // The OP_DEFER_ACTION actions returned in pidTagClientActions property on each DAM must equal the each ruleAction set by client.
            bool isVerifyR571 = ropQueryRowsResponseOfDAM.RowCount == 2 &&
                                pidTagMessageClassOfDAMOne.Equals(Constants.DAMMessageClass) &&
                                pidTagMessageClassOfDAMTwo.Equals(Constants.DAMMessageClass) &&
                                pidTagRuleProviderOfDAMOne.Equals(ruleProviderOneSetByClient) &&
                                pidTagRuleProviderOfDAMTwo.Equals(ruleProviderTwoSetByClient) &&
                                Common.CompareByteArray(pidTagClientActionsOfDAMOne.Serialize(), ruleActionForRuleOneSetByClient.Serialize()) &&
                                Common.CompareByteArray(pidTagClientActionsOfDAMTwo.Serialize(), ruleActionForRuleTwoSetByClient.Serialize());

            Site.CaptureRequirementIfIsTrue(
                isVerifyR571,
                571,
                 @"[In Generating a DAM]The server MUST generate separate DAMs for ""OP_DEFER_ACTION"" actions (2) that belong to separate rule providers.");
            #endregion
            #endregion

            #region TestUser1 updates the DAM message with invalid client EntryId, and the updating operation is failed.
            // Call RopUpdateDeferredActionMessages to update the DAM message, and trigger a failed call.
            // Prepare data in the RopUpdateDeferredActionMessages request buffer. Here the EntryId will cause a failed result.
            byte[] clientEntryId = AdapterHelper.ConvertStringToBytes(Constants.InvalidateEntryId);
            byte[] serverEntryId = AdapterHelper.ConvertStringToBytes(Constants.InvalidateEntryId);

            // Call RopUpdateDeferredActionMessages.
            RopUpdateDeferredActionMessagesResponse ropUpdateDeferredActionMessagesResponse = this.OxoruleAdapter.RopUpdateDeferredActionMessages(this.LogonHandle, serverEntryId, clientEntryId);
            Site.Assert.AreNotEqual<uint>(0, ropUpdateDeferredActionMessagesResponse.ReturnValue, "Updating DAM operation should fail.");
            #endregion

            #region TestUser1 updates the DAM message with valid client EntryId, and the updating operation is successful.

            // Prepare data in the RopUpdateDeferredActionMessages request buffer
            clientEntryId = AdapterHelper.ConvertStringToBytes(Constants.ClientEntryId);

            // serverEntryId is set to the value of pidTagDeferredActionMessageOriginalEntryId.
            serverEntryId = pidTagDeferredActionMessageOriginalEntryIdOfDAMOne;

            // Call RopUpdateDeferredActionMessages.
            ropUpdateDeferredActionMessagesResponse = this.OxoruleAdapter.RopUpdateDeferredActionMessages(this.LogonHandle, serverEntryId, clientEntryId);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R371: server return value is {0}", ropUpdateDeferredActionMessagesResponse.ReturnValue);

            // Verify MS-OXORULE requirement: MS-OXORULE_R371
            // If the server return success, it means the server EntryId is valid and the PidTagDeferredActionMessageOriginalEntryId contain the server EntryId, otherwise can't return success.
            bool isVerifyR371 = ropUpdateDeferredActionMessagesResponse.ReturnValue == 0;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR371,
                371,
                "[In PidTagDeferredActionMessageOriginalEntryId Property] The PidTagDeferredActionMessageOriginalEntryId property ([MS-OXPROPS] section 2.652) contains the server EntryID for the DAM message on the server.");
            #endregion

            #region TestUser1 verifies whether the associated properties on each DAM are changed after updating.
            #region Step1: TestUser1 calls RopOpenMessage to open the two DAM messages and to get their message handle.

            // Prepare data in the RopOpenMessage request buffer.
            ulong damOneMessageId = BitConverter.ToUInt64(pidTagMIDOfDAMOneOfBytes, 0);
            ulong damTwoMessageId = BitConverter.ToUInt64(pidTagMIDOfDAMTwoOfBytes, 0);

            // Call RopOpenMessage.
            RopOpenMessageResponse openMessageResponseOfDAM = new RopOpenMessageResponse();
            uint damOneHandle = this.OxoruleAdapter.RopOpenMessage(this.DAFFolderHandle, this.DAFFolderID, damOneMessageId, out openMessageResponseOfDAM);
            Site.Assert.AreEqual<uint>(0, openMessageResponseOfDAM.ReturnValue, "Opening DAM one Message operation should succeed.");
            uint damTwoHandle = this.OxoruleAdapter.RopOpenMessage(this.DAFFolderHandle, this.DAFFolderID, damTwoMessageId, out openMessageResponseOfDAM);
            Site.Assert.AreEqual<uint>(0, openMessageResponseOfDAM.ReturnValue, "Opening DAM two Message operation should succeed.");
            #endregion

            #region Step2: TestUser1 calls RopGetPropertiesSpecific to get the changed properties' value after updating.

            // Prepare data in the RopGetPropertiesSpecific request buffer.
            PropertyTag[] propertyTags = new PropertyTag[3];

            // PidTagDeferredActionMessageOriginalEntryId
            propertyTags[0].PropertyId = (ushort)PropertyId.PidTagDeferredActionMessageOriginalEntryId;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypServerId;

            // PidTagDeferredActionMessageOriginalEntryId
            propertyTags[1].PropertyId = (ushort)PropertyId.PidTagRuleFolderEntryId;
            propertyTags[1].PropertyType = (ushort)PropertyType.PtypBinary;

            // PidTagDeferredActionMessageOriginalEntryId
            propertyTags[2].PropertyId = (ushort)PropertyId.PidTagDamOriginalEntryId;
            propertyTags[2].PropertyType = (ushort)PropertyType.PtypBinary;

            // Call RopGetPropertiesSpecific.
            RopGetPropertiesSpecificResponse ropGetPropertiesSpecificResponseOne = this.OxoruleAdapter.RopGetPropertiesSpecific(damOneHandle, propertyTags);
            Site.Assert.AreEqual<uint>(0, ropGetPropertiesSpecificResponseOne.ReturnValue, "Getting DAM one specific properties operation should succeed.");
            RopGetPropertiesSpecificResponse ropGetPropertiesSpecificResponseTwo = this.OxoruleAdapter.RopGetPropertiesSpecific(damTwoHandle, propertyTags);
            Site.Assert.AreEqual<uint>(0, ropGetPropertiesSpecificResponseOne.ReturnValue, "Getting DAM two specific properties operation should succeed.");
            #endregion

            #region Step3: TestUser1 verifies the properties' value contained in the RopGetPropertiesSpecific response buffer.

            // The returned property values' order in the RopGetPropertiesSpecific response buffer is the same with the order they are set in the request buffer.
            byte[] pidTagDeferredActionMessageOriginalEntryIdOneAfterUpdate = AdapterHelper.PropertyValueConvertToBinary(ropGetPropertiesSpecificResponseOne.RowData.PropertyValues[0].Value);
            byte[] pidTagDeferredActionMessageOriginalEntryIdTwoAfterUpdate = AdapterHelper.PropertyValueConvertToBinary(ropGetPropertiesSpecificResponseTwo.RowData.PropertyValues[0].Value);
            #endregion

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R618: the PidTagDeferredActionMessageOriginalEntryId on the two generated DAMs are {0}, {1}, and the clientEntryId is {2}", pidTagDeferredActionMessageOriginalEntryIdOneAfterUpdate, pidTagDeferredActionMessageOriginalEntryIdTwoAfterUpdate, clientEntryId);

            // Verify MS-OXORULE requirement: MS-OXORULE_R618.
            // If the PidTagDeferredActionMessageOriginalEntryId on the two generated DAMs are both changed to the value of clientEntryId set on RopUpdateDeferredActionMessages operation, it means server has found all DAMs that 
            // have the value of the PidTagDeferredActionMessageOriginalEntryId property that are equal to the value in the ServerEntryId field of the RopUpdateDeferredActionMessages ROP request buffer
            bool isVerifyR618 = Common.CompareByteArray(pidTagDeferredActionMessageOriginalEntryIdOneAfterUpdate, clientEntryId) &&
                Common.CompareByteArray(pidTagDeferredActionMessageOriginalEntryIdTwoAfterUpdate, clientEntryId);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR618,
                618,
                @"[In Receiving a RopUpdateDeferredActionMessages ROP Request] The server also MUST find all DAMs that have the value of the PidTagDeferredActionMessageOriginalEntryId property (section 2.2.6.8) equal to the value in the ServerEntryId field of the RopUpdateDeferredActionMessages ROP request buffer, as specified in section 2.2.3.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R619 , pidTagDeferredActionMessageOriginalEntryIdOneAfterUpdate is {0}, pidTagDeferredActionMessageOriginalEntryIdTwoAfterUpdate is {1}", pidTagDeferredActionMessageOriginalEntryIdOneAfterUpdate, pidTagDeferredActionMessageOriginalEntryIdTwoAfterUpdate);

            // Verify MS-OXORULE requirement: MS-OXORULE_R619.
            // If the PidTagDeferredActionMessageOriginalEntryId on the two generated DAMs are both changed to the clientEntryId, it means the server has changed PidTagDamOriginalEntryId on each DAM.
            bool isVerifyR619 = Common.CompareByteArray(pidTagDeferredActionMessageOriginalEntryIdOneAfterUpdate, clientEntryId) &&
                Common.CompareByteArray(pidTagDeferredActionMessageOriginalEntryIdTwoAfterUpdate, clientEntryId);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR619,
                619,
                @"[In Receiving a RopUpdateDeferredActionMessages ROP Request] The server MUST then change the value of the PidTagDeferredActionMessageOriginalEntryId property on each DAM it finds to the value passed in the ClientEntryId field of the same ROP request buffer.");
            #endregion

            #endregion
        }

        /// <summary>
        /// This test case is designed for server to generate DAM message to verify modify rule. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S05_TC03_ServerGenerateDAM_ForModifyRule()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.DAMPidTagRuleNameOne);
            #endregion

            #region TestUser1 creates a rule which can trigger server to generate DAM.
            // If the action type is "OP_DEFER_ACTION", the ActionData buffer is completely under the control of the client that created the rule.
            // When a message that satisfies the rule condition is received, the server creates a DAM
            // and places the entire content of the ActionBlock field as part of the PidTagClientActions property on the DAM.
            DeferredActionData deferredActionDataSetByClient = new DeferredActionData
            {
                Data = Encoding.Unicode.GetBytes(Constants.DAMPidTagRuleActionOne + "\0")
            };
            ActionType actionType = ActionType.OP_DEFER_ACTION;
            RuleData ruleDataSetByClient = AdapterHelper.GenerateValidRuleData(actionType, TestRuleDataType.ForAdd, 0, RuleState.ST_ENABLED, deferredActionDataSetByClient, ruleProperties, null);

            // Call RopModifyRules.
            this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleDataSetByClient });
            #endregion

            #region TestUser1 calls RopGetRulesTable.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");
            #endregion

            #region TestUser1 retrieves rule information for the newly added rule.
            PropertyTag[] propertyAllTags = new PropertyTag[2];
            propertyAllTags[0].PropertyId = (ushort)PropertyId.PidTagRuleName;
            propertyAllTags[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyAllTags[1].PropertyId = (ushort)PropertyId.PidTagRuleId;
            propertyAllTags[1].PropertyType = (ushort)PropertyType.PtypInteger64;

            // Retrieves rows from the rule table.
            RopQueryRowsResponse queryRowResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyAllTags);
            Site.Assert.AreEqual<uint>(0, queryRowResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");

            // Only one rule added in this folder, so the row count in the rule table should be 1.
            // If the rule table is got successfully and the rule count is correct, it means that the server is returning a table with the rule added by the test suite.
            Site.Assert.AreEqual<uint>(1, queryRowResponse.RowCount, "The rule number in the rule table is {0}", queryRowResponse.RowCount);
            this.VerifyRuleTable();

            ulong ruleId = BitConverter.ToUInt64(queryRowResponse.RowData.PropertyRows[0].PropertyValues[1].Value, 0);
            #endregion

            #region TestUser1 modifies the rule added in step 2.
            deferredActionDataSetByClient = new DeferredActionData
            {
                Data = Encoding.Unicode.GetBytes(Constants.DAMPidTagRuleActionTwo + "\0")
            };
            actionType = ActionType.OP_DEFER_ACTION;
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.DAMPidTagRuleNameTwo);
            ruleDataSetByClient = AdapterHelper.GenerateValidRuleData(actionType, TestRuleDataType.ForModify, 0, RuleState.ST_ENABLED, deferredActionDataSetByClient, ruleProperties, ruleId);

            // Call RopModifyRules.
            this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleDataSetByClient });
            #endregion

            #region TestUser1 calls RopGetRulesTable.
            ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");
            #endregion

            #region TestUser1 retrieves rule information for the newly modified rule.
            // Retrieves rows from the rule table.
            queryRowResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyAllTags);
            Site.Assert.AreEqual<uint>(0, queryRowResponse.ReturnValue, "Retrieving rows from the rule table should succeed");

            // Only one rule added in this folder, so the row count in the rule table should be 1.
            // If the rule table is got successfully and the rule count is correct, it means that the server is returning a table with the rule added by the test suite.
            Site.Assert.AreEqual<uint>(1, queryRowResponse.RowCount, "The rule number in the rule table is {0}", queryRowResponse.RowCount);
            this.VerifyRuleTable();
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger this rule.
            string deliveredMessageSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser2 delivers a message to trigger this rule.
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, deliveredMessageSubject);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 calls RopSetColumns to set the interested columns of the message table in the DAF folder.

            // Prepare the data in the RopSetColumns request buffer.
            PropertyTag[] propertyTagOfDAM = new PropertyTag[2];
            propertyTagOfDAM[0].PropertyId = (ushort)PropertyId.PidTagClientActions;
            propertyTagOfDAM[0].PropertyType = (ushort)PropertyType.PtypBinary;
            propertyTagOfDAM[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagOfDAM[1].PropertyType = (ushort)PropertyType.PtypString;

            // Call RopQueryRows.
            uint contentsTableHandleOfDAF = 0;
            uint rowCount = 0;
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForDAM;
            RopQueryRowsResponse ropQueryRowsResponseOfDAM = this.GetExpectedMessage(this.DAFFolderHandle, ref contentsTableHandleOfDAF, propertyTagOfDAM, ref rowCount);
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.OtherTarget;

            // Get the value of PidTagClientActions property.
            byte[] pidTagClientActionsOfDAMOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[(int)rowCount - 1].PropertyValues[0].Value;

            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R883");
            RuleAction pidTagClientActionsOfDAM = AdapterHelper.PropertyValueConvertToRuleAction(pidTagClientActionsOfDAMOfBytes);

            // pidTagClientActionsOfDAM is the value of the property PidTagClientActions. 
            // If it Action Type is equal to the relevant actions as they were set by the client, R883 can be verified.
            Site.CaptureRequirementIfAreEqual<ActionType>(
                ActionType.OP_DEFER_ACTION,
                pidTagClientActionsOfDAM.Actions[0].ActionType,
                883,
                @"[In PidTagClientActions Property] The server is required to set values in this property according to the relevant actions (2) as they were set by the client when the rule (2) was changed by using the RopModifyRules ROP ([MS-OXCROPS] section 2.2.11.1).");
            #endregion
        }

        /// <summary>
        /// This test case is designed for server to generate a DEM message to get the PidTagDamOriginalEntryIdProperty.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S05_TC04_ServerGenerateDEM_ByOP_MOVE_Error_ForPidTagDamOriginalEntryIdProperty()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Empty the DAF folder and the Inbox folder.
            // Empty the DAF folder.
            this.OxoruleAdapter.RopEmptyFolder(this.DAFFolderHandle, 0x00);

            // Empty the Inbox folder.
            this.OxoruleAdapter.RopEmptyFolder(this.InboxFolderHandle, 0x00);
            #endregion

            #region Prepare value for ruleProperties variable.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForDEM;
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.DEMRule);
            #endregion

            #region TestUser1 creates folder1 in server store.
            RopCreateFolderResponse createFolderResponse;
            this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, Common.GenerateResourceName(this.Site, "User1Folder01"), "TestForOP_MOVE", out createFolderResponse);
            ulong newFolderId = createFolderResponse.FolderId;
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating folder operation should succeed.");
            #endregion

            #region TestUser1 adds one rule in the Inbox folder.
            #region Prepare rules' data.
            MoveCopyActionData moveCopyActionData = new MoveCopyActionData();

            // Get the created folder1 entry id.
            ServerEID serverEID = new ServerEID(BitConverter.GetBytes(newFolderId));
            byte[] folder1EId = serverEID.Serialize();

            // Get the store object's entry id
            byte[] storeEId = this.GetStoreObjectEntryID(StoreObjectType.Mailbox, this.Server, this.User1ESSDN);
            moveCopyActionData.FolderInThisStore = 1;
            moveCopyActionData.FolderEID = folder1EId;
            moveCopyActionData.StoreEID = storeEId;
            moveCopyActionData.FolderEIDSize = (ushort)folder1EId.Length;
            moveCopyActionData.StoreEIDSize = (ushort)storeEId.Length;
            #endregion

            #region Generate rule with generated rule data.
            // Generate rule data.
            RuleData ruleForMove = AdapterHelper.GenerateValidRuleData(ActionType.OP_MOVE, TestRuleDataType.ForAdd, 0, RuleState.ST_ENABLED, moveCopyActionData, ruleProperties, null);

            // Call RopModifyRules ROP to add a rule.
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForMove });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding rule should succeed!");
            #endregion
            #endregion

            #region Delete the newly created folders.
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, newFolderId);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1.
            // TestUser2 delivers a message to TestUser1.
            string deliveredMessageSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, deliveredMessageSubject);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message handle and message ID.
            // Prepare the data in the RopSetColumns request buffer
            PropertyTag[] propertyTagOfInboxFolder = new PropertyTag[2];
            propertyTagOfInboxFolder[0].PropertyId = (ushort)PropertyId.PidTagMid;
            propertyTagOfInboxFolder[0].PropertyType = (ushort)PropertyType.PtypInteger64;
            propertyTagOfInboxFolder[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagOfInboxFolder[1].PropertyType = (ushort)PropertyType.PtypString;

            // Each row includes the property values of the interested columns.
            // Call RopQueryRows.
            uint contentsTableHandleOfInboxFolder = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse ropQueryRowsResponseOfInboxFolder = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandleOfInboxFolder, propertyTagOfInboxFolder, ref expectedMessageIndex, deliveredMessageSubject);

            ulong messageId = AdapterHelper.PropertyValueConvertToUint64(ropQueryRowsResponseOfInboxFolder.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);
            RopOpenMessageResponse ropOpenMessageResponse = new RopOpenMessageResponse();

            // Open the message to get the message handle.
            uint messagehandle = this.OxoruleAdapter.RopOpenMessage(this.InboxFolderHandle, this.InboxFolderID, messageId, out ropOpenMessageResponse);
            #endregion

            #region TestUser1 gets the message entry ID and the Inbox folder's entry ID.
           
            // Get message's entry ID.
            byte[] messageEntryId = this.OxoruleAdapter.GetMessageEntryId(this.InboxFolderHandle, this.InboxFolderID, messagehandle, messageId);

            // Get Inbox folder's entry ID
            byte[] inboxFolderEntryId = this.OxoruleAdapter.GetFolderEntryId(StoreObjectType.Mailbox, this.InboxFolderHandle, this.InboxFolderID);
            #endregion

            #region TestUser1 gets specified properties in DEM.

            // Call RopGetContentsTable.
            PropertyTag[] propertyTagOfDAF = new PropertyTag[5];
            propertyTagOfDAF[0].PropertyId = (ushort)PropertyId.PidTagRuleError;
            propertyTagOfDAF[0].PropertyType = (ushort)PropertyType.PtypInteger32;
            propertyTagOfDAF[1].PropertyId = (ushort)PropertyId.PidTagDamOriginalEntryId;
            propertyTagOfDAF[1].PropertyType = (ushort)PropertyType.PtypBinary;
            propertyTagOfDAF[2].PropertyId = (ushort)PropertyId.PidTagRuleFolderEntryId;
            propertyTagOfDAF[2].PropertyType = (ushort)PropertyType.PtypBinary;
            propertyTagOfDAF[3].PropertyId = (ushort)PropertyId.PidTagMessageClass;
            propertyTagOfDAF[3].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagOfDAF[4].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagOfDAF[4].PropertyType = (ushort)PropertyType.PtypString;

            uint contentsTableHandleOfDAF = 0;
            uint rowCount = 0;
            RopQueryRowsResponse ropQueryRowsResponseOfDAF = this.GetExpectedMessage(this.DAFFolderHandle, ref contentsTableHandleOfDAF, propertyTagOfDAF, ref rowCount);
            Site.Assert.IsTrue(rowCount > 0, @"The message number in the specific folder is {0}", rowCount);

            // Get specific DEM properties.
            uint ruleError = AdapterHelper.PropertyValueConvertToUint(ropQueryRowsResponseOfDAF.RowData.PropertyRows[0].PropertyValues[0].Value);
            byte[] pidTagDamOriginalEntryId = new byte[ropQueryRowsResponseOfDAF.RowData.PropertyRows[0].PropertyValues[1].Value.Length - 2];
            Array.Copy(ropQueryRowsResponseOfDAF.RowData.PropertyRows[0].PropertyValues[1].Value, 2, pidTagDamOriginalEntryId, 0, ropQueryRowsResponseOfDAF.RowData.PropertyRows[0].PropertyValues[1].Value.Length - 2);
            byte[] pidTagRuleFolderEntryId = new byte[ropQueryRowsResponseOfDAF.RowData.PropertyRows[0].PropertyValues[2].Value.Length - 2];
            Array.Copy(ropQueryRowsResponseOfDAF.RowData.PropertyRows[0].PropertyValues[2].Value, 2, pidTagRuleFolderEntryId, 0, ropQueryRowsResponseOfDAF.RowData.PropertyRows[0].PropertyValues[2].Value.Length - 2);
            string messageClassValue = AdapterHelper.PropertyValueConvertToString(ropQueryRowsResponseOfDAF.RowData.PropertyRows[0].PropertyValues[3].Value);

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R387");

            // Verify MS-OXORULE requirement: MS-OXORULE_R387.
            // The destination folder entry ID is invalid in this case, so the cause of the error is error moving or copying the message to the destination folder.
            Site.CaptureRequirementIfAreEqual<uint>(
                6,
                ruleError,
                387,
                @"[In PidTagRuleError Property] The meaning of the value 0x00000006: Error moving or copying the message to the destination folder.");

            if (Common.IsRequirementEnabled(909, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R909");

                // Verify MS-OXORULE requirement: MS-OXORULE_R909.
                // Because the destination folder entry ID is invalid, the server encounters an error when processing the OP_MOVE action.
                // The messageClassValue indicates the message whether is a DEM.
                Site.CaptureRequirementIfAreEqual<string>(
                    "IPC.Microsoft Exchange 4.0.Deferred Error",
                    messageClassValue,
                    909,
                    @"[In Handling Errors During Rule Processing] The implementation does generate a DEM when it encounters an error processing a rule (2) on an incoming message. (Exchange 2003 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(7132, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R7132.");

                // Verify MS-OXORULE requirement: MS-OXORULE_R7132.
                // The messageEntryIdInbyte get from the inboxFolder is the message that was being processed by the server when this error was encountered.
                bool isVerifyR7132 = pidTagDamOriginalEntryId.Length == messageEntryId.Length;
                if (isVerifyR7132)
                {
                    for (int i = 0; i < pidTagDamOriginalEntryId.Length; i++)
                    {
                        if (pidTagDamOriginalEntryId[i] != messageEntryId[i])
                        {
                            isVerifyR7132 = false;
                            break;
                        }
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR7132,
                    7132,
                    @"[[In Appendix A: Product Behavior] Implementation does set PidTagDamOriginalEntryId  to the EntryID of the message that was being processed by the server when this error was encountered (this is, the ""delivered message""). (Exchange 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(7152, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R7152.");

                // Verify MS-OXORULE requirement: MS-OXORULE_R7152.
                // The inboxFolder is the folder where the rule that triggered in this case.
                bool isVerifyR7152 = pidTagRuleFolderEntryId.Length == inboxFolderEntryId.Length;
                if (isVerifyR7152)
                {
                    for (int i = 0; i < pidTagRuleFolderEntryId.Length; i++)
                    {
                        if (pidTagRuleFolderEntryId[i] != inboxFolderEntryId[i])
                        {
                            isVerifyR7152 = false;
                            break;
                        }
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR7152,
                    7152,
                    @"[[In Appendix A: Product Behavior] Implementation does set PidTagRuleFolderEntryId   to the EntryID of the folder where the rule (2) that triggered the generation of this DEM is stored. (Exchange 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(896, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R896");

                // Verify MS-OXORULE requirement: MS-OXORULE_R896.
                // The error caused by the OP_MOVE action while server executing a rule is verified by R387,
                // and the messageClassValue indicates the message whether is a DEM.
                Site.CaptureRequirementIfAreEqual<string>(
                    "IPC.Microsoft Exchange 4.0.Deferred Error",
                    messageClassValue,
                    896,
                    @"[In DEM Syntax] Implementation does create a DEM when an error is encountered while executing a rule (2). (Exchange 2003 and above follow this behavior.)");
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R903");

            // Verify MS-OXORULE requirement: MS-OXORULE_R903.
            // The error caused by the OP_MOVE action is verified by R387,
            // and the messageClassValue indicates the message whether is a DEM.
            Site.CaptureRequirementIfAreEqual<string>(
                "IPC.Microsoft Exchange 4.0.Deferred Error",
                messageClassValue,
                903,
                @"[In Processing DAMs and DEMs] The server places a message in the DAF when it encounters a problem performing an action (2) of a server-side rule (DEM).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R810");

            // Verify MS-OXORULE requirement: MS-OXORULE_R810.
            // If the ropGetContentsTableResponseOfDAF.RowCount is set to 1, it means there is a message in DAF.
            // The messageClassValue indicates whether the message is a DEM.
            bool isVerifyR810 = messageClassValue == "IPC.Microsoft Exchange 4.0.Deferred Error";

            Site.CaptureRequirementIfIsTrue(
                isVerifyR810,
                810,
                @"[In Handling Errors During Rule Processing] The server MUST generate the DEM in the following manner: 1. Create a new message (DEM) in the DAF.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R812");

            // Verify MS-OXORULE requirement: MS-OXORULE_R812.
            // If the value of the PidTagRuleMessageClass is null means the server has saved the DEM.
            Site.CaptureRequirementIfIsNotNull(
                messageClassValue,
                812,
                @"[In Handling Errors During Rule Processing] The server MUST generate the DEM in the following manner: 3. Save the DEM.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed for server to generate a DEM message to get the ActionNumberProperty.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S05_TC05_ServerGenerateDEM_ByOP_MOVE_Error_ForActionNumberProperty()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Empty the DAF folder and the Inbox folder.
            // Empty the DAF folder.
            this.OxoruleAdapter.RopEmptyFolder(this.DAFFolderHandle, 0x00);

            // Empty the Inbox folder.
            this.OxoruleAdapter.RopEmptyFolder(this.InboxFolderHandle, 0x00);
            #endregion

            #region Prepare value for ruleProperties variable.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForDEM;
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.DEMRule);
            #endregion

            #region TestUser1 creates folder1 in server store.
            RopCreateFolderResponse createFolderResponse;
            this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, Common.GenerateResourceName(this.Site, "User1Folder01"), "TestForOP_MOVE", out createFolderResponse);
            ulong newFolderId = createFolderResponse.FolderId;
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating folder operation should succeed.");
            #endregion

            #region TestUser1 adds one rule with two actions in the Inbox folder.
            #region Step1. Prepare rules' data.
            MoveCopyActionData moveCopyActionData = new MoveCopyActionData();

            // Get the created folder1 entry id.
            ServerEID serverEID = new ServerEID(BitConverter.GetBytes(newFolderId));
            byte[] folder1EId = serverEID.Serialize();

            // Get the store object's entry id
            byte[] storeEId = this.GetStoreObjectEntryID(StoreObjectType.Mailbox, this.Server, this.User1ESSDN);
            moveCopyActionData.FolderInThisStore = 1;
            moveCopyActionData.FolderEID = folder1EId;
            moveCopyActionData.StoreEID = storeEId;
            moveCopyActionData.FolderEIDSize = (ushort)folder1EId.Length;
            moveCopyActionData.StoreEIDSize = (ushort)storeEId.Length;

            DeleteMarkReadActionData markAsRead = new DeleteMarkReadActionData();
            #endregion

            #region Step2. Generate rule with generated rule data.
            // Generate rule data.
            RuleData ruleForMove = AdapterHelper.GenerateValidRuleDataWithFlavor(new ActionType[] { ActionType.OP_MARK_AS_READ, ActionType.OP_MOVE }, 0, RuleState.ST_ENABLED, new IActionData[] { markAsRead, moveCopyActionData }, new uint[] { 0, 0 }, ruleProperties);

            // Call RopModifyRules ROP to add a rule.
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForMove });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding new rule should succeed!");
            #endregion
            #endregion

            #region Delete the newly created folder.
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, newFolderId);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1.
            // TestUser2 delivers a message to TestUser1.
            string deliveredMessageSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, deliveredMessageSubject);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets specified properties in DEM.
            PropertyTag[] propertyTagOfDAF = new PropertyTag[7];
            propertyTagOfDAF[0].PropertyId = (ushort)PropertyId.PidTagMessageClass;
            propertyTagOfDAF[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagOfDAF[1].PropertyId = (ushort)PropertyId.PidTagRuleActionType;
            propertyTagOfDAF[1].PropertyType = (ushort)PropertyType.PtypInteger32;
            propertyTagOfDAF[2].PropertyId = (ushort)PropertyId.PidTagRuleActionNumber;
            propertyTagOfDAF[2].PropertyType = (ushort)PropertyType.PtypInteger32;
            propertyTagOfDAF[3].PropertyId = (ushort)PropertyId.PidTagMid;
            propertyTagOfDAF[3].PropertyType = (ushort)PropertyType.PtypInteger64;
            propertyTagOfDAF[4].PropertyId = (ushort)PropertyId.PidTagDamOriginalEntryId;
            propertyTagOfDAF[4].PropertyType = (ushort)PropertyType.PtypBinary;
            propertyTagOfDAF[5].PropertyId = (ushort)PropertyId.PidTagRuleFolderEntryId;
            propertyTagOfDAF[5].PropertyType = (ushort)PropertyType.PtypBinary;
            propertyTagOfDAF[6].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagOfDAF[6].PropertyType = (ushort)PropertyType.PtypString;

            uint contentsTableHandleOfDAF = 0;
            uint rowCount = 0;
            RopQueryRowsResponse ropQueryRowsResponseOfDAF = this.GetExpectedMessage(this.DAFFolderHandle, ref contentsTableHandleOfDAF, propertyTagOfDAF, ref rowCount);
            Site.Assert.IsTrue(rowCount > 0, @"The message number in the specific folder is {0}", rowCount);

            // Get specific DEM properties.
            uint ruleActionType = AdapterHelper.PropertyValueConvertToUint(ropQueryRowsResponseOfDAF.RowData.PropertyRows[0].PropertyValues[1].Value);
            uint ruleActionNumber = AdapterHelper.PropertyValueConvertToUint(ropQueryRowsResponseOfDAF.RowData.PropertyRows[0].PropertyValues[2].Value);
            ulong deferredErrorMessageId = AdapterHelper.PropertyValueConvertToUint64(ropQueryRowsResponseOfDAF.RowData.PropertyRows[0].PropertyValues[3].Value);

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R707");

            // Verify MS-OXORULE requirement: MS-OXORULE_R707.
            // The DEM in this case is generated because the failure of the OP_MOVE action,
            // and the OP_MOVE action block is the second action block in the action data of the rule,
            // so the zero-based index of the OP_MOVE action is 1.
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                ruleActionNumber,
                707,
                @"[In PidTagRuleActionNumber Property] The PidTagRuleActionNumber property ([MS-OXPROPS] section 2.934) MUST be set to the zero-based index of the action (2) that failed. (For example, if specific to an action (2), a property value of 0x00000000 means that the first action (2) failed, 0x00000001 means that the second action (2) failed.)");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R709");

            // Verify MS-OXORULE requirement: MS-OXORULE_R709.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)ActionType.OP_MOVE,
                ruleActionType,
                709,
                @"[In PidTagRuleActionNumber Property] The ActionType field value of the action (2) at this index MUST be the same value as the value of the PidTagRuleActionType property (section 2.2.7.3) in this DEM.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R750");

            // Verify MS-OXORULE requirement: MS-OXORULE_R750.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)ActionType.OP_MOVE,
                ruleActionType,
                750,
                @"[In PidTagRuleActionType Property] This property [PidTagRuleActionType] MUST be set to the value of the ActionType field, as specified in section 2.2.5.1. [if the failure is specific to an action (2).]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R1009");

            // Verify MS-OXORULE requirement: MS-OXORULE_R1009.
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)ActionType.OP_MOVE,
                ruleActionType,
                1009,
                @"[In PidTagRuleActionType Property] The PidTagRuleActionType property ([MS-OXPROPS] section 2.936) specifies the action (2) of the rule (2) that failed.");
            #endregion
            #endregion

            #region Call RopGetPropertiesAll to verify all DEM properties.
            RopOpenMessageResponse openMessageResponse;

            // Call RopOpenMessageResponse to get the DEM handle.
            uint deferredErrorMessageHandle = this.OxoruleAdapter.RopOpenMessage(this.DAFFolderHandle, this.DAFFolderID, deferredErrorMessageId, out openMessageResponse);
            Site.Assert.AreEqual<uint>(0, openMessageResponse.ReturnValue, "Call RopOpenMessage failed!");

            // Call RopGetPropertiesAll to verify all DEM properties in adapter.
            this.OxoruleAdapter.RopGetPropertiesAll(deferredErrorMessageHandle, 0xffff, (ushort)WantUnicode.Want);
            #endregion
        }

        /// <summary>
        /// This test case is designed for server to generate a DEM message to get the RuleIdProperty.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S05_TC06_ServerGenerateDEM_ByOP_MOVE_Error_ForRuleIdProperty()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Empty the DAF folder and the Inbox folder.
            // Empty the DAF folder.
            this.OxoruleAdapter.RopEmptyFolder(this.DAFFolderHandle, 0x00);

            // Empty the Inbox folder.
            this.OxoruleAdapter.RopEmptyFolder(this.InboxFolderHandle, 0x00);
            #endregion

            #region Prepare value for ruleProperties variable.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForDEM;
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.DEMRule);
            #endregion

            #region TestUser1 creates folder1 in server store.
            RopCreateFolderResponse createFolderResponse;
            this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, Common.GenerateResourceName(this.Site, "User1Folder01"), "TestForOP_MOVE", out createFolderResponse);
            ulong newFolderId = createFolderResponse.FolderId;
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating folder operation should succeed.");
            #endregion

            #region Prepare rules' data
            MoveCopyActionData moveCopyActionData = new MoveCopyActionData();

            // Get the created folder1 entry id.
            ServerEID serverEID = new ServerEID(BitConverter.GetBytes(newFolderId));
            byte[] folder1EId = serverEID.Serialize();

            // Get the store object's entry id
            byte[] storeEId = this.GetStoreObjectEntryID(StoreObjectType.Mailbox, this.Server, this.User1ESSDN);
            moveCopyActionData.FolderInThisStore = 1;
            moveCopyActionData.FolderEID = folder1EId;
            moveCopyActionData.StoreEID = storeEId;
            moveCopyActionData.FolderEIDSize = (ushort)folder1EId.Length;
            moveCopyActionData.StoreEIDSize = (ushort)storeEId.Length;
            IActionData[] moveCopyAction = { moveCopyActionData };
            #endregion

            #region Generate test RuleData.
            // Add rule for move without rule Provider Data.
            ruleProperties.ProviderData = string.Empty;
            RuleData ruleForMoveFolder = AdapterHelper.GenerateValidRuleDataWithFlavor(new ActionType[] { ActionType.OP_MOVE }, 0, RuleState.ST_ENABLED, moveCopyAction, new uint[] { 0 }, ruleProperties);

            #endregion

            #region TestUser1 adds OP_MOVE rule to the Inbox folder.
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForMoveFolder });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Move rule should succeed.");
            #endregion

            #region Delete the newly created folder.
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, newFolderId);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion

            #region TestUser2 delivers a message to trigger these rules.

            // TestUser2 deliver a message to trigger these rules.
            string deliveredMessageSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title", 1);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, deliveredMessageSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets specified properties in DEM.
            PropertyTag[] propertyTagOfDAF = new PropertyTag[3];
            propertyTagOfDAF[0].PropertyId = (ushort)PropertyId.PidTagMessageClass;
            propertyTagOfDAF[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagOfDAF[1].PropertyId = (ushort)PropertyId.PidTagRuleId;
            propertyTagOfDAF[1].PropertyType = (ushort)PropertyType.PtypInteger64;
            propertyTagOfDAF[2].PropertyId = (ushort)PropertyId.PidTagRuleProvider;
            propertyTagOfDAF[2].PropertyType = (ushort)PropertyType.PtypString;

            uint contentsTableHandleOfDAF1 = 0;
            uint rowCount1 = 0;
            this.GetExpectedMessage(this.DAFFolderHandle, ref contentsTableHandleOfDAF1, propertyTagOfDAF, ref rowCount1);
            Site.Assert.IsTrue(rowCount1 == 1, @"The message number in the specific folder is {0}", rowCount1);
            #endregion

            #region TestUser2 delivers second message to trigger these rules.

            // TestUser2 deliver a message to trigger these rules.
            deliveredMessageSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title", 2);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, deliveredMessageSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets specified properties in DEM.
            uint contentsTableHandleOfDAF2 = 0;
            uint rowCount2 = 0;
            RopQueryRowsResponse ropQueryRowsResponseOfDAF2 = this.GetExpectedMessage(this.DAFFolderHandle, ref contentsTableHandleOfDAF2, propertyTagOfDAF, ref rowCount2);
            Site.Assert.IsTrue(rowCount2 > 0, @"The message number in the specific folder is {0}", rowCount2);

            // Get specific DEM properties.
            ulong ruleIdInDEM = AdapterHelper.PropertyValueConvertToUint64(ropQueryRowsResponseOfDAF2.RowData.PropertyRows[0].PropertyValues[1].Value);
            string ruleProviderInDEM = AdapterHelper.PropertyValueConvertToString(ropQueryRowsResponseOfDAF2.RowData.PropertyRows[0].PropertyValues[2].Value);
            #endregion

            #region TestUser1 gets specified Properties in rule.
            RopGetRulesTableResponse getRulesTableResponse;

            // Call RopGetRulesTable.
            uint inboxRulesTableHandler = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out getRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, getRulesTableResponse.ReturnValue, "Getting rules table handler should succeed!");

            PropertyTag[] propertyTagOfRule = new PropertyTag[3];
            propertyTagOfRule[0].PropertyId = (ushort)PropertyId.PidTagRuleState;
            propertyTagOfRule[0].PropertyType = (ushort)PropertyType.PtypInteger32;
            propertyTagOfRule[1].PropertyId = (ushort)PropertyId.PidTagRuleId;
            propertyTagOfRule[1].PropertyType = (ushort)PropertyType.PtypInteger64;
            propertyTagOfRule[2].PropertyId = (ushort)PropertyId.PidTagRuleProvider;
            propertyTagOfRule[2].PropertyType = (ushort)PropertyType.PtypString;

            // Call RopQueryRowsResponse.
            RopQueryRowsResponse ropQueryRowsResponseOfInboxfolder = this.OxoruleAdapter.QueryPropertiesInTable(inboxRulesTableHandler, propertyTagOfRule);

            // Get PidTagRuleState property.
            uint ruleState = AdapterHelper.PropertyValueConvertToUint(ropQueryRowsResponseOfInboxfolder.RowData.PropertyRows[0].PropertyValues[0].Value);
            ulong ruleIdInRule = AdapterHelper.PropertyValueConvertToUint64(ropQueryRowsResponseOfInboxfolder.RowData.PropertyRows[0].PropertyValues[1].Value);
            string ruleProviderInRule = AdapterHelper.PropertyValueConvertToString(ropQueryRowsResponseOfInboxfolder.RowData.PropertyRows[0].PropertyValues[2].Value);

            #region Capture Code
            if (Common.IsRequirementEnabled(913, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R913,  the rule state is {0}.", ruleState.ToString());

                // Verify MS-OXORULE requirement: MS-OXORULE_R913.
                bool isVerifyR913 = (ruleState & (uint)RuleState.ST_ERROR) == (uint)RuleState.ST_ERROR;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR913,
                    913,
                    @"[In Handling Errors During Rule Processing] The first time the server finds a server-side rule to be in error and has generated a DEM for it, the implementation does set the ST_ERROR flag in the PidTagRuleState property (section 2.2.1.3.1.3) of that rule (2). (Exchange 2003 and above follow this behavior.)");
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R581");

            // Verify MS-OXORULE requirement: MS-OXORULE_R581.
            // TestUser2 delivered two messages to TestUser1, these operations may cause two failure of the same rule,
            // but if the ST_ERROR flag prevents creating multiple DEMs with the same error information,
            // the ST_ERROR flag must be set and the number of the DEM must be 1 in this case.
            bool isVerifyR581 = ((ruleState & (uint)RuleState.ST_ERROR) == (uint)RuleState.ST_ERROR) && (rowCount2 == 1);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR581,
                581,
                @"[In Handling Errors During Rule Processing (Creating a DEM)] Examination of the ST_ERROR flag on subsequent operations is used to prevent creating multiple DEMs with the same error information.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R711");

            // Verify MS-OXORULE requirement: MS-OXORULE_R711.
            Site.CaptureRequirementIfAreEqual<string>(
                ruleProviderInRule,
                ruleProviderInDEM,
                711,
                @"[In PidTagRuleProvider Property] The PidTagRuleProvider property (section 2.2.1.3.1.5) MUST be set to the same value as the PidTagRuleProvider property on the rule or rules that have caused the DEM to be generated.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R717");

            // Verify MS-OXORULE requirement: MS-OXORULE_R717.
            Site.CaptureRequirementIfAreEqual<ulong>(
                ruleIdInRule,
                ruleIdInDEM,
                717,
                @"[In PidTagRuleId Property] The PidTagRuleId (section 2.2.1.3.1.1) property MUST be set to the same value as the value of the PidTagRuleId property on the rule (2) that has generated this error.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed for server to generate one DAM (Deferred Action Message) message in DAF (Deferred Action Folder) folder.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S05_TC07_ServerGenerateOneDAM_ForOP_DEFER_ACTION()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.DAMPidTagRuleNameOne);
            #endregion

            #region TestUser1 creates one new rules which can trigger server to generate DAM.
            // If the action type is "OP_DEFER_ACTION", the ActionData buffer is completely under the control of the client that created the rule.
            // When a message that satisfies the rule condition is received, the server creates a DAM
            // and places the entire content of the ActionBlock field as part of the PidTagClientActions property on the DAM.
            DeferredActionData deferredActionDataSetByClient = new DeferredActionData
            {
                Data = Encoding.Unicode.GetBytes(Constants.DAMPidTagRuleActionOne + "\0")
            };
            ActionType actionType = ActionType.OP_DEFER_ACTION;

            RuleData ruleData1SetByClient = AdapterHelper.GenerateValidRuleData(actionType, TestRuleDataType.ForAdd, 0, RuleState.ST_ENABLED, deferredActionDataSetByClient, ruleProperties, null);

            // Call RopModifyRules.
            this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleData1SetByClient });
            #endregion

            #region TestUser1 calls RopGetRulesTable.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");
            #endregion

            #region TestUser1 retrieves rule information for the newly added rule.
            PropertyTag[] propertyAllTags = new PropertyTag[]
            {
                new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleName,
                    PropertyType = (ushort)PropertyType.PtypString
                }
            };

            // Retrieves rows from the rule table.
            RopQueryRowsResponse queryRowResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyAllTags);
            Site.Assert.AreEqual<uint>(0, queryRowResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");
            Site.Assert.AreEqual(1, queryRowResponse.RowCount, @"There should be one rule returned, actual returned row count is {0}.", queryRowResponse.RowCount);
            this.VerifyRuleTable();
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger the rule.
            string deliveredMessageSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser2 delivers a message to TestUser1 to trigger the rule.
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, deliveredMessageSubject);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets Message's entry ID.
            #region TestUser1 gets the message handle and message ID.

            // Prepare the data in the RopSetColumns request buffer.
            PropertyTag[] propertyTagOfInboxFolder = new PropertyTag[4];
            propertyTagOfInboxFolder[0].PropertyId = (ushort)PropertyId.PidTagMid;
            propertyTagOfInboxFolder[0].PropertyType = (ushort)PropertyType.PtypInteger64;
            propertyTagOfInboxFolder[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagOfInboxFolder[1].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagOfInboxFolder[2].PropertyId = (ushort)PropertyId.PidTagHasDeferredActionMessages;
            propertyTagOfInboxFolder[2].PropertyType = (ushort)PropertyType.PtypBoolean;
            propertyTagOfInboxFolder[3].PropertyId = (ushort)PropertyId.PidTagRwRulesStream;
            propertyTagOfInboxFolder[3].PropertyType = (ushort)PropertyType.PtypBoolean;

            // Each row includes the property values of the interested columns which are set in above step. 
            uint contentsTableHandleOfInboxFolder = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse ropQueryRowsResponseOfInboxFolder = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandleOfInboxFolder, propertyTagOfInboxFolder, ref expectedMessageIndex, deliveredMessageSubject);
            Site.Assert.AreEqual<uint>(0, ropQueryRowsResponseOfInboxFolder.ReturnValue, "Query rows operation should succeed.");
            ulong messageId = AdapterHelper.PropertyValueConvertToUint64(ropQueryRowsResponseOfInboxFolder.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);
            RopOpenMessageResponse ropOpenMessageResponse = new RopOpenMessageResponse();

            // Open the message to get the message handle.
            uint messagehandle = this.OxoruleAdapter.RopOpenMessage(this.InboxFolderHandle, this.InboxFolderID, messageId, out ropOpenMessageResponse);
            #endregion

            #region TestUser1 gets the message entry id and the Inbox folder's entry ID.

            // Get message's entry ID.
            byte[] messageEId = this.OxoruleAdapter.GetMessageEntryId(this.InboxFolderHandle, this.InboxFolderID, messagehandle, messageId);
            #endregion

            #region TestUser1 gets the message subject and checks if has deffered action messages.

            // Subject, bodyText and originalMessageSender are the properties set by the server on the replied message.
            string subject = AdapterHelper.PropertyValueConvertToString(ropQueryRowsResponseOfInboxFolder.RowData.PropertyRows[expectedMessageIndex].PropertyValues[1].Value);

            // Get the value of PidTagHasDeferredActionMessages.
            byte[] pidTagHasDeferredActionMessagesOfDAMOfBytes = ropQueryRowsResponseOfInboxFolder.RowData.PropertyRows[expectedMessageIndex].PropertyValues[2].Value;
            bool pidTagHasDeferredActionMessagesOfDAM = AdapterHelper.PropertyValueConvertToBool(pidTagHasDeferredActionMessagesOfDAMOfBytes);
            #endregion
            #endregion

            #region TestUser1 checks the generation of DAM, which is placed under DAF folder.
            #region TestUser1 sets the interested columns of the message table in the DAF folder.
            PropertyTag[] propertyTagOfDAM = new PropertyTag[9];
            propertyTagOfDAM = AdapterHelper.GenerateRuleInfoPropertiesOfDAM();

            // Query rows include the property values of the interested columns.
            uint contentsTableHandleOfDAF = 0;
            uint rowCount = 0;
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForDAM;
            RopQueryRowsResponse ropQueryRowsResponseOfDAM = this.GetExpectedMessage(this.DAFFolderHandle, ref contentsTableHandleOfDAF, propertyTagOfDAM, ref rowCount);
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.OtherTarget;
            bool isDAMUnderDAFFolder = rowCount > 0;
            #endregion

            #region TestUser1 verifies the properties' values contained in RopQueryRowsResponse for the generated DAM messages.

            // In this test case, there is only one row returned in the PropertyRows buffer, which represents one generated DAM message. 
            // And there are 9 interested properties for the DAM message returned in the PropertyValues buffer.
            // The returned property values' order for the row is the same with the order they are set through RopSetColumns.
            byte[] pidTagMessageClassOfDAMOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[0].PropertyValues[0].Value;
            byte[] pidTagDamBackPatchedOfDAMOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[0].PropertyValues[1].Value;
            byte[] pidTagDamOriginalEntryIdOfDAMOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[0].PropertyValues[2].Value;
            byte[] pidTagRuleProviderOfDAMOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[0].PropertyValues[3].Value;
            byte[] pidTagRuleFolderEntryIdofDAMOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[0].PropertyValues[4].Value;
            byte[] pidTagClientActionsOfDAMOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[0].PropertyValues[5].Value;
            byte[] pidTagDeferredActionMessageOriginalEntryIdOfDAMOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[0].PropertyValues[7].Value;
            byte[] pidTagMIDOfDAMOfBytes = ropQueryRowsResponseOfDAM.RowData.PropertyRows[0].PropertyValues[8].Value;

            #region Capture Code
            if (Common.IsRequirementEnabled(740, this.Site))
            {
                // For PtypBinary type, the first 2 bytes are the count of the binary bytes,
                // so the actual value of the property with PtypBinary type does not contain the first 2 bytes. 
                byte[] pidTagDamOriginalEntryIdOfDAMOfBytesWithoutSize = new byte[pidTagDamOriginalEntryIdOfDAMOfBytes.Length - 2];
                Array.Copy(pidTagDamOriginalEntryIdOfDAMOfBytes, 2, pidTagDamOriginalEntryIdOfDAMOfBytesWithoutSize, 0, pidTagDamOriginalEntryIdOfDAMOfBytes.Length - 2);

                bool isPidTagDamOriginalEntryIdEqualMessageEId = pidTagDamOriginalEntryIdOfDAMOfBytesWithoutSize.Length == messageEId.Length;
                if (isPidTagDamOriginalEntryIdEqualMessageEId)
                {
                    for (int i = 0; i < pidTagDamOriginalEntryIdOfDAMOfBytesWithoutSize.Length; i++)
                    {
                        if (pidTagDamOriginalEntryIdOfDAMOfBytesWithoutSize[i] != messageEId[i])
                        {
                            isPidTagDamOriginalEntryIdEqualMessageEId = false;
                            break;
                        }
                    }
                }

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R740.");

                // Verify MS-OXORULE requirement: MS-OXORULE_R740.
                // PidTagDamOriginalEntryIdOfDAMOfBytes is the value of the PidTagDamOriginalEntryId property on DAM. 
                // If it equal to the EntryID of the message, R740 can be verified.
                bool isVerifyR740 = isPidTagDamOriginalEntryIdEqualMessageEId;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR740,
                    740,
                    @"[In PidTagDamOriginalEntryId Property] This PidTagDamOriginalEntryId property ([MS-OXPROPS] section 2.650) MUST be set to the EntryID of the delivered (target) message that the client has to process.");
            }

            if (Common.IsRequirementEnabled(741, this.Site))
            {
                // For PtypBinary type, the first 2 bytes are the count of the binary bytes,
                // so the actual value of the property with PtypBinary type does not contain the first 2 bytes. 
                byte[] pidTagRuleFolderEntryIdofDAMOfBytesWithoutSize = new byte[pidTagRuleFolderEntryIdofDAMOfBytes.Length - 2];
                Array.Copy(pidTagRuleFolderEntryIdofDAMOfBytes, 2, pidTagRuleFolderEntryIdofDAMOfBytesWithoutSize, 0, pidTagRuleFolderEntryIdofDAMOfBytes.Length - 2);

                byte[] folderEId = this.OxoruleAdapter.GetFolderEntryId(StoreObjectType.Mailbox, this.InboxFolderHandle, this.InboxFolderID);
                bool isPidTagRuleFolderEntryIdEqualfolderEId = pidTagRuleFolderEntryIdofDAMOfBytesWithoutSize.Length == folderEId.Length;
                if (isPidTagRuleFolderEntryIdEqualfolderEId)
                {
                    for (int i = 0; i < pidTagRuleFolderEntryIdofDAMOfBytesWithoutSize.Length; i++)
                    {
                        if (pidTagRuleFolderEntryIdofDAMOfBytesWithoutSize[i] != folderEId[i])
                        {
                            isPidTagRuleFolderEntryIdEqualfolderEId = false;
                            break;
                        }
                    }
                }

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R741.");

                // Verify MS-OXORULE requirement: MS-OXORULE_R741.
                // pidTagRuleFolderEntryIdofDAMOfBytes is the value of the PidTagRuleFolderEntryId property on DAM. 
                // If it equal to the EntryID of the DAF folder, R741 can be verified.
                bool isVerifyR741 = isPidTagRuleFolderEntryIdEqualfolderEId;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR741,
                    741,
                    @"[In PidTagRuleFolderEntryId Property] The PidTagRuleFolderEntryId property ([MS-OXPROPS] section 2.939) MUST be set to the EntryID of the folder where the rule (2) that triggered the generation of this DAM is stored.");
            }

            RuleAction pidTagClientActionsOfDAM = AdapterHelper.PropertyValueConvertToRuleAction(pidTagClientActionsOfDAMOfBytes);
            Site.Assert.AreEqual<int>(1, pidTagClientActionsOfDAM.Actions.Length, "There should be only one action in the rule action!");
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R365.");

            // pidTagClientActionsOfDAM is the value of the property PidTagClientActions. 
            // If it Action Type is equal to the relevant actions as they were set by the client, R365 can be verified.
            Site.CaptureRequirementIfAreEqual<ActionType>(
                ActionType.OP_DEFER_ACTION,
                pidTagClientActionsOfDAM.Actions[0].ActionType,
                365,
                @"[In PidTagClientActions] The server is required to set values in this property according to the relevant actions (2) as they were set by the client when the rule (2) was created by using the RopModifyRules ROP ([MS-OXCROPS] section 2.2.11.1).");

            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R808.");

            // Verify MS-OXORULE requirement: MS-OXORULE_R808.
            // If the value of properties that were set by server is not null, R808 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                pidTagClientActionsOfDAM.Actions[0].ActionDataValue,
                808,
                @"[In Generating a DAM] The server MUST generate the DAM in the following manner: 2. Set the property values on the DAM as specified in section 2.2.6.");

            // Verify MS-OXORULE requirement: MS-OXORULE_R359.
            string pidTagRuleProviderOfDAM = AdapterHelper.PropertyValueConvertToString(pidTagRuleProviderOfDAMOfBytes);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R359.");

            Site.CaptureRequirementIfAreEqual<string>(
                ruleProperties.Provider,
                pidTagRuleProviderOfDAM,
                359,
                @"[In PidTagRuleProvider] The PidTagRuleProvider property ([MS-OXPROPS] section 2.951) MUST be set to the same value as the PidTagRuleProvider property on the rule or rules that have generated the DAM.");

            // Add the debug information.
            string pidTagMessageClassOfDAM = AdapterHelper.PropertyValueConvertToString(pidTagMessageClassOfDAMOfBytes);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R563: there is {0} DAM generated, PidTagMessageClass is {1}, PidTagRuleProvider is {2}", ropQueryRowsResponseOfDAM.RowCount, pidTagMessageClassOfDAM, pidTagRuleProviderOfDAM);

            // Verify MS-OXORULE requirement: MS-OXORULE_R563.
            // This test case is designed based on the rule condition that is evaluated to be TRUE but the server cannot perform the actions specified in the rule.
            // That RowCount of ropQueryRowsResponseOfDAM equals 1 means there is only one message generated under the DAF folder.
            // That pidTagMessageClassOfDAM equals to "IPC.Microsoft Exchange 4.0.Deferred Action" means this message is a DAM message
            // That pidTagRuleProviderOfDAM equals to ruleProviderSetByClient means this DAM is generated for the rule added in this test case.
            bool isVerifyR563 = ropQueryRowsResponseOfDAM.RowCount == 1 &&
                pidTagMessageClassOfDAM == Constants.DAMMessageClass &&
                pidTagRuleProviderOfDAM == ruleProperties.Provider;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR563,
                563,
                @"[In Generating a DAM]A server MUST generate a DAM when a rule (2) condition evaluates to ""TRUE"" but the server cannot perform the actions (2) specified in the rule (2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R724: the value of PidTagHasDeferredActionMessages is {0}, and PidTagSubject is {1}", pidTagHasDeferredActionMessagesOfDAM, subject);

            // Verify MS-OXORULE requirement: MS-OXORULE_R724.
            // The server has generated a DAM and it is verified in R563.
            bool isVerifyR724 = pidTagHasDeferredActionMessagesOfDAM && subject.Equals(deliveredMessageSubject);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR724,
                724,
                @"[In PidTagHasDeferredActionMessages Property] This property MUST be set to ""TRUE"" if a message has at least one associated DAM.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R564.");

            // Verify MS-OXORULE requirement: MS-OXORULE_R564.
            // The server has generated a DAM and it is verified in R563.
            Site.CaptureRequirementIfIsTrue(
                pidTagHasDeferredActionMessagesOfDAM,
                564,
                @"[In Generating a DAM] When the server generates DAMs for a message, the server MUST set the value of the PidTagHasDeferredActionMessages property (section 2.2.9.1) on the message to ""TRUE"".");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R536.");

            // Verify MS-OXORULE requirement: MS-OXORULE_R536.
            // The server has generated a DAM and it is verified in R563.
            Site.CaptureRequirementIfIsTrue(
                pidTagHasDeferredActionMessagesOfDAM,
                536,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_DEFER_ACTION"": The server MUST also set the PidTagHasDeferredActionMessages property (section 2.2.9.1) to ""TRUE"" on the message.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R426: there is {0} DAM generated, PidTagMessageClass is {1}, PidTagRuleProvider is {2}", ropQueryRowsResponseOfDAM.RowCount, pidTagMessageClassOfDAM, pidTagRuleProviderOfDAM);

            // Verify MS-OXORULE requirement: MS-OXORULE_R426.
            // R426 has verified that the server has generated a DAM.
            bool isVerifyR426 = isVerifyR563;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR426,
                426,
                @"[In Processing DAMs and DEMs] The server places a message in the DAF when it needs the client to perform an action (2) as a result of a client-side rule (DAM).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R932: there is {0} DAM generated, PidTagMessageClass is {1}, PidTagRuleProvider is {2}", ropQueryRowsResponseOfDAM.RowCount, pidTagMessageClassOfDAM, pidTagRuleProviderOfDAM);

            // Verify MS-OXORULE requirement: MS-OXORULE_R932.
            bool isVerifyR932 = isVerifyR563;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR932,
                932,
                @"[In Processing Incoming Messages to a Folder] When executing a rule (2) whose condition evaluates to ""TRUE"" as per the restriction (2) in the PidTagRuleCondition property (section 2.2.1.3.1.9), then the server MUST generate a DAM for the client to process as specified in section 3.2.5.1.2 in the case of a client-side rule.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R354.");

            // Verify MS-OXORULE requirement: MS-OXORULE_R354.
            // Whether the DAM is generated has been verified in R563.
            bool isVerifyR354 = isVerifyR563 && !BitConverter.ToBoolean(pidTagDamBackPatchedOfDAMOfBytes, 0);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR354,
                354,
                @"[In PidTagDamBackPatched property] The PidTagDamBackPatched property ([MS-OXPROPS] section 2.649) MUST be set to ""FALSE"" when the DAM is generated.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R565: there is {0} DAM generated, PidTagMessageClass is {1}, PidTagRuleProvider is {2}, and is DAM under DAF folder is {3}", ropQueryRowsResponseOfDAM.RowCount, pidTagMessageClassOfDAM, pidTagRuleProviderOfDAM, isDAMUnderDAFFolder);

            // Verify MS-OXORULE requirement: MS-OXORULE_R565.
            // If the DAM is generated in the DAF folder, it means the server has created a new message (DAM) successfully.
            // Whether the DAM is generated has been verified in R563.
            bool isVerifyR565 = isVerifyR563 && isDAMUnderDAFFolder;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR565,
                565,
                @"[In Generating a DAM] The server MUST generate the DAM in the following manner: 1. Create a new message (DAM) in the DAF.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R567: there is {0} DAM generated, PidTagMessageClass is {1}, PidTagRuleProvider is {2}, and is DAM under DAF folder is {3}", ropQueryRowsResponseOfDAM.RowCount, pidTagMessageClassOfDAM, pidTagRuleProviderOfDAM, isDAMUnderDAFFolder);

            // Verify MS-OXORULE requirement: MS-OXORULE_R567.
            // If the DAM is generated in the DAF folder, it means the server has saved the DAM successfully.
            // Whether the DAM is generated has been verified in R563.
            bool isVerifyR567 = isVerifyR563 && isDAMUnderDAFFolder;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR567,
                567,
                @"[In Generating a DAM] The server MUST generate the DAM in the following manner: 3. Save the DAM.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R535: there is {0} DAM generated, PidTagMessageClass is {1}, PidTagRuleProvider is {2}, and is DAM under DAF folder is {3}", ropQueryRowsResponseOfDAM.RowCount, pidTagMessageClassOfDAM, pidTagRuleProviderOfDAM, isDAMUnderDAFFolder);

            // Verify MS-OXORULE requirement: MS-OXORULE_R535.
            // Whether the DAM is generated has been verified in R563.
            bool isVerifyR535 = isVerifyR563 && isDAMUnderDAFFolder;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR535,
                535,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_DEFER_ACTION"": The server MUST generate a DAM as specified in section 3.2.5.1.2.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R313: there is {0} DAM generated, PidTagMessageClass is {1}, PidTagRuleProvider is {2}", ropQueryRowsResponseOfDAM.RowCount, pidTagMessageClassOfDAM, pidTagRuleProviderOfDAM);

            // Verify MS-OXORULE requirement: MS-OXORULE_R313.
            // If the DAM has been generated, it means the actions for the rule cannot be executed on the server, and the rule is a client-side rule.
            // Whether the DAM is generated has been verified in R563.
            bool isVerifyR313 = isVerifyR563;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR313,
                313,
                @"[In OP_DEFER_ACTION ActionData Structure] If one or more actions (2) for a specific rule (2) cannot be executed on the server, the rule (2) is required to be a client-side rule, with a value in the ActionType field of ""OP_DEFER_ACTION"".");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R257: there is {0} DAM generated, PidTagMessageClass is {1}, PidTagRuleProvider is {2}", ropQueryRowsResponseOfDAM.RowCount, pidTagMessageClassOfDAM, pidTagRuleProviderOfDAM);

            // Verify MS-OXORULE requirement: MS-OXORULE_R257.
            // If the DAM has been generated, it means the actions for the rule cannot be executed on the server, and the rule is a client-side rule.
            // Whether the DAM is generated has been verified in R563.
            bool isVerifyR257 = isVerifyR563;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR257,
                257,
                @"[In ActionBlock Structure] The meaning of action type OP_DEFER_ACTION: Used for actions (2) that cannot be executed by the server (like playing a sound).");

            // The error code for NotFound is 0x8004010f, which is represented as notFoundError.
            byte[] notFoundError = new byte[4] { 0x0f, 0x01, 0x04, 0x80 };

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R427: the value of PidTagDeferredActionMessageOriginalEntryId is {0}", pidTagDeferredActionMessageOriginalEntryIdOfDAMOfBytes);

            // Verify MS-OXORULE requirement: MS-OXORULE_R427.
            // If the value of PidTagDeferredActionMessageOriginalEntryId doesn't equal the NotFound error,
            // it indicates the server has updated the PidTagDeferredActionMessageOriginalEntryId and set it on DAM. 
            bool isVerifyR427 = !Common.CompareByteArray(pidTagDeferredActionMessageOriginalEntryIdOfDAMOfBytes, notFoundError);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR427,
                427,
                @"[In Processing DAMs and DEMs]When the server creates a DAM, it updates the PidTagDeferredActionMessageOriginalEntryId property (section 2.2.6.8), which is then used by the client in the ServerEntryId field of the RopUpdateDeferredActionMessages ROP request buffer (section 2.2.3).");

            // actionDataForRuleSetByClient represents the OP_DEFER_ACTION action in the Rule which has been added in this test case.
            ActionBlock actionDataForRuleSetByClient = new ActionBlock(CountByte.TwoBytesCount)
            {
                ActionType = ActionType.OP_DEFER_ACTION,
                ActionFlavor = 0x00000000,
                ActionFlags = 0x00000000,
                ActionDataValue = deferredActionDataSetByClient
            };
            actionDataForRuleSetByClient.ActionLength = actionDataForRuleSetByClient.ActionDataValue.Size() + 9;

            // Pack the information about the OP_DEFER_ACTION action to a RuleAction structure.
            RuleAction packedRuleActionsSetByClient = new RuleAction(CountByte.TwoBytesCount)
            {
                NoOfActions = 0x0001,
                Actions = new ActionBlock[1]
            };
            packedRuleActionsSetByClient.Actions[0] = actionDataForRuleSetByClient;

            // pidTagClientActionsOfDAM represents the pidTagClientActions property set on the DAM message. 
            // This property contains the relevant actions which need to be further processed by the client.
            pidTagClientActionsOfDAM = AdapterHelper.PropertyValueConvertToRuleAction(pidTagClientActionsOfDAMOfBytes);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R321");

            // Verify MS-OXORULE requirement: MS-OXORULE_R321.
            // This test case is designed based on a message that satisfies the rule condition that is received.
            // Whether the DAM is created has been verified in R563.
            // Whether the server places the entire content of the ActionBlock as part of the PidTagClientActions property on the DAM
            // has been verified in R570 and R365,
            // so the condition to verify R306 is "isVerifyR563" and the value of the PidTagClientActions property.
            bool isVerifyR321 = isVerifyR563 && (ropQueryRowsResponseOfDAM.RowCount == 1) && Common.CompareByteArray(packedRuleActionsSetByClient.Serialize(), pidTagClientActionsOfDAM.Serialize());

            Site.CaptureRequirementIfIsTrue(
                isVerifyR321,
                321,
                @"[In OP_DEFER_ACTION ActionData Structure] When a message that satisfies the rule (2) condition is received, the server creates a DAM and places the entire content of the ActionBlocks field of the RuleAction structure in the PidTagClientActions property (section 2.2.6.6) on the DAM as specified in sections 3.2.5.1.2, 2.2.6, and 2.2.6.6.");
            #endregion
            #endregion
            #endregion

            #region TestUser1 updates the DAM message.

            // The first two bytes of pidTagDeferredActionMessageOriginalEntryIdOfDAMOfBytes are only used to indicate the byte length of this property,
            // so the actual value of the pidTagDeferredActionMessageOriginalEntryId should exclude the first 2 bytes. 
            byte[] pidTagDeferredActionMessageOriginalEntryIdOfDAM = new byte[pidTagDeferredActionMessageOriginalEntryIdOfDAMOfBytes.Length - 2];
            Array.Copy(pidTagDeferredActionMessageOriginalEntryIdOfDAMOfBytes, 2, pidTagDeferredActionMessageOriginalEntryIdOfDAM, 0, pidTagDeferredActionMessageOriginalEntryIdOfDAM.Length);

            // Prepare data in the RopUpdateDeferredActionMessages request buffer.
            byte[] clientEntryId = AdapterHelper.ConvertStringToBytes(Constants.InvalidateEntryId);

            // serverEntryId is set to the value of pidTagDeferredActionMessageOriginalEntryId.
            byte[] serverEntryId = pidTagDeferredActionMessageOriginalEntryIdOfDAM;

            // Call RopUpdateDeferredActionMessages.
            RopUpdateDeferredActionMessagesResponse ropUpdateDeferredActionMessagesResponse = this.OxoruleAdapter.RopUpdateDeferredActionMessages(this.LogonHandle, serverEntryId, clientEntryId);
            Site.Assert.AreEqual<uint>(0, ropUpdateDeferredActionMessagesResponse.ReturnValue, "Updating deferred action message operation should succeed.");

            #region Capture Code

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R145.");

            // Verify MS-OXORULE requirement: MS-OXORULE_R145.
            // Set InputHandleIndex value to 0x00 when constructing RopModifyRulesRequest.
            // The InputHandleIndex can be got only in adapter code.
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                ropUpdateDeferredActionMessagesResponse.InputHandleIndex,
                145,
                @"[In RopUpdateDeferredActionMessages ROP Response Buffer] InputHandleIndex (1 byte): This value MUST be the same as the index to the input handle in the request buffer for this operation [Processing RopUpdateDeferredActionMessages ROP Response].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R617.");

            // Verify MS-OXORULE requirement: MS-OXORULE_R617.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                ropUpdateDeferredActionMessagesResponse.ReturnValue,
                617,
                @"[In Receiving a RopUpdateDeferredActionMessages ROP Request] If the server successfully parses the data in the ROP request buffer, it MUST return 0x00000000 as the value of the ReturnValue field in the response buffer.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R147.");

            // Verify MS-OXORULE requirement: MS-OXORULE_R147.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                ropUpdateDeferredActionMessagesResponse.ReturnValue,
                147,
                @"[In RopUpdateDeferredActionMessages ROP Response Buffer] ReturnValue: To indicate success, the server returns 0x00000000.");
            #endregion
            #endregion

            #region TestUser1 verifies whether the associated properties on DAM are changed after updating.

            #region TestUser1 calls RopOpenMessage to open the DAM message and to get its message handle.

            // Prepare data in the RopOpenMessage request buffer
            ulong damMessageId = BitConverter.ToUInt64(pidTagMIDOfDAMOfBytes, 0);

            // Call RopOpenMessage.
            RopOpenMessageResponse openMessageResponseOfDAM = new RopOpenMessageResponse();
            uint damHandle = this.OxoruleAdapter.RopOpenMessage(this.DAFFolderHandle, this.DAFFolderID, damMessageId, out openMessageResponseOfDAM);
            Site.Assert.AreEqual<uint>(0, openMessageResponseOfDAM.ReturnValue, "Opening DAF folder should succeed");

            #endregion

            #region TestUser1 calls RopGetPropertiesSpecific to get the changed properties' value after updating.

            // Prepare data in the RopGetPropertiesSpecific request buffer.
            PropertyTag[] propertyTags = new PropertyTag[2];

            // PidTagDamBackPatched
            propertyTags[0].PropertyId = (ushort)PropertyId.PidTagDamBackPatched;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypBoolean;

            // PidTagDamOriginalEntryId
            propertyTags[1].PropertyId = (ushort)PropertyId.PidTagDamOriginalEntryId;
            propertyTags[1].PropertyType = (ushort)PropertyType.PtypBinary;

            // The RopGetPropertiesSpecific call
            RopGetPropertiesSpecificResponse ropGetPropertiesSpecificResponse = this.OxoruleAdapter.RopGetPropertiesSpecific(damHandle, propertyTags);
            Site.Assert.AreEqual<uint>(0, ropGetPropertiesSpecificResponse.ReturnValue, "Getting specific properties operation should succeed.");

            #endregion

            #region TestUser1 verifies the properties' value contained in the RopGetPropertiesSpecific response buffer.

            // The returned property values' order in the RopGetPropertiesSpecific response buffer is the same with the order they are set in the request buffer.
            bool pidTagDamBackPatchedOfDAMAfterUpdate = AdapterHelper.PropertyValueConvertToBool(ropGetPropertiesSpecificResponse.RowData.PropertyValues[0].Value);
            byte[] pidTagDamUpdatedEntryIdOfDAMOfBytes = ropGetPropertiesSpecificResponse.RowData.PropertyValues[1].Value;
            #endregion

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R136");

            // Verify MS-OXORULE requirement: MS-OXORULE_R136.
            // PidTagDamUpdatedEntryIdOfDAMOfBytes represents the value of PidTagDamOriginalEntryId property on changed DAM, if it does not equal the original one, 
            // it means it is updated by the server as a result of a RopUpdateDeferredActionMessages request.
            bool isVerifyR136 = pidTagDamUpdatedEntryIdOfDAMOfBytes != pidTagDamOriginalEntryIdOfDAMOfBytes;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR136,
                136,
                @"[In RopUpdateDeferredActionMessages ROP] The RopUpdateDeferredActionMessages ROP instructs the server to update the PidTagDamOriginalEntryId property (section 2.2.6.3) on one or more DAMs.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R620");

            // Verify MS-OXORULE requirement: MS-OXORULE_R620.
            // pidTagDamBackPatchedOfDAMAfterUpdate represents the value of PidTagDamBackPatched property on changed DAM, 
            // which is updated by the server as a result of a RopUpdateDeferredActionMessages request.
            Site.CaptureRequirementIfIsTrue(
                pidTagDamBackPatchedOfDAMAfterUpdate,
                620,
                @"[In Receiving a RopUpdateDeferredActionMessages ROP Request] The server MUST also set the value of the PidTagDamBackPatched property (section 2.2.6.2) to ""TRUE"" on any DAM that it changed.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R355");

            // Verify MS-OXORULE requirement: MS-OXORULE_R355.
            // pidTagDamBackPatchedOfDAMAfterUpdate represents the value of PidTagDamBackPatched property on changed DAM, 
            // which is updated by the server as a result of a RopUpdateDeferredActionMessages request.
            Site.CaptureRequirementIfIsTrue(
                pidTagDamBackPatchedOfDAMAfterUpdate,
                355,
                @"[In PidTagDamBackPatched property] It [The PidTagDamBackPatched property] MUST be set to ""TRUE"" if the DAM was updated by the server as a result of a RopUpdateDeferredActionMessages request ([MS-OXCROPS] section 2.2.11.3).");
            #endregion
            #endregion
        }
    }
}