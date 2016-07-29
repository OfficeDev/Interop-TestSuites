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
    /// This scenario aims to validate server behaviors of processing Out of Office rule.
    /// </summary>
    [TestClass]
    public class S03_ProcessOutOfOfficeRule : TestSuiteBase
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
        /// This test case is designed to verify that a rule which has ST_EXIT_LEVEL flag but does not have ST_ONLY_WHEN_OOF flag will not be evaluated.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S03_TC01_OOFBehaviorsNotExecuteSubRuleWithoutST_ONLY_WHEN_OOFForST_EXIT_LEVEL()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameTag);
            string setOOFMailAddress = this.User1Name + "@" + this.Domain;
            string userPassword = this.User1Password;
            #endregion

            #region Set TestUser1 to OOF state.
            bool isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, true);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office on for {0} should succeed.", this.User1Name);
            Thread.Sleep(this.WaitForSetOOFComplete);
            #endregion

            #region TestUser1 adds an OP_TAG rule with PidTagRuleState set to ST_ENABLED | ST_EXIT_LEVEL.
            TagActionData tagActionData = new TagActionData();
            PropertyTag tagActionDataPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagImportance,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            tagActionData.PropertyTag = tagActionDataPropertyTag;
            tagActionData.PropertyValue = BitConverter.GetBytes(2);

            RuleData ruleOpTag = AdapterHelper.GenerateValidRuleData(ActionType.OP_TAG, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED | RuleState.ST_EXIT_LEVEL, tagActionData, ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleOpTag });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding OP_TAG rule should succeed");
            #endregion

            #region TestUser1 adds OP_MARK_AS_READ rule with PidTagRuleState set to ST_ENABLED.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameMarkAsRead);
            RuleData ruleForMarkRead = AdapterHelper.GenerateValidRuleData(ActionType.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 2, RuleState.ST_ENABLED, new DeleteMarkReadActionData(), ruleProperties, null);
            ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleForMarkRead });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding mark as read rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region Testuser1 verifies whether the specific property value is set on the received mail.
            PropertyTag[] propertyTagList = new PropertyTag[3];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagImportance;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypInteger32;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[2].PropertyId = (ushort)PropertyId.PidTagMessageFlags;
            propertyTagList[2].PropertyType = (ushort)PropertyType.PtypInteger32;

            uint contentsTableHandle = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);

            // If the PidTagImportance is the value which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property specified in the rule's action buffer structure.
            Site.Assert.AreEqual<int>(2, BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0), "The value of PidTagImportance field should be the specific value set by the client.");
            #endregion

            #region Testuser1 verifies whether the received mail is marked as read.
            bool isSubjectContainsRuleConditionSubject = false;
            int messageFlags = 0;
            mailSubject = AdapterHelper.PropertyValueConvertToString(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[1].Value);
            if (mailSubject.Contains(ruleProperties.ConditionSubjectName))
            {
                isSubjectContainsRuleConditionSubject = true;
                messageFlags = BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[2].Value, 0);
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R560, the value of markAsReadFlag is {0}", messageFlags);

            // Verify MS-OXORULE requirement: MS-OXORULE_R560.
            // 0x00000001 is the bit which indicates the message has been marked as read. So if this bit is not set in messageFlags,
            // it means the message is not marked as read, which means the rule has not been executed.
            bool isVerifyR560 = isSubjectContainsRuleConditionSubject && (messageFlags & 0x00000001) != 0x00000001;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR560,
                560,
                @"[In Interaction Between ST_ONLY_WHEN_OOF and ST_EXIT_LEVEL Flags] When the Out of Office state is set on the mailbox, as specified in [MS-OXWOOF], and a rule (2) condition evaluates to ""TRUE"", if the rule (2) has the ST_EXIT_LEVEL flag specified in section 2.2.1.3.1.3 set, then the server MUST NOT evaluate subsequent rules (2) that do not have the ST_ONLY_WHEN_OOF flag set.");
            #endregion

            #region Set Testuser1 back to normal state (not in OOF state)
            isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, false);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office off for {0} should succeed.", this.User1Name);
            Thread.Sleep(this.WaitForSetOOFComplete);
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify that a rule which has ST_EXIT_LEVEL flag, but its sub-sequence rule that has ST_ONLY_WHEN_OOF flag set will be evaluated.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S03_TC02_OOFBehaviorsExecuteSubRuleWithST_ONLY_WHEN_OOFForST_EXIT_LEVEL()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameTag);
            string setOOFMailAddress = this.User1Name + "@" + this.Domain;
            string userPassword = this.User1Password;
            #endregion

            #region Set TestUser1 to OOF state.
            bool isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, true);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office on for {0} should succeed.", this.User1Name);
            Thread.Sleep(this.WaitForSetOOFComplete);
            #endregion

            #region TestUser1 adds an OP_TAG rule with PidTagRuleState set to ST_ENABLED | ST_EXIT_LEVEL.
            TagActionData tagActionData = new TagActionData();
            PropertyTag tagActionDataPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagImportance,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            tagActionData.PropertyTag = tagActionDataPropertyTag;
            tagActionData.PropertyValue = BitConverter.GetBytes(2);

            RuleData ruleOpTag = AdapterHelper.GenerateValidRuleData(ActionType.OP_TAG, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED | RuleState.ST_EXIT_LEVEL, tagActionData, ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleOpTag });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding OP_TAG rule should succeed");
            #endregion

            #region TestUser1 adds OP_MARK_AS_READ rule with PidTagRuleState set to ST_ENABLED | ST_ONLY_WHEN_OOF.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameMarkAsRead);

            RuleData ruleForMarkRead = AdapterHelper.GenerateValidRuleData(ActionType.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 2, RuleState.ST_ENABLED | RuleState.ST_ONLY_WHEN_OOF, new DeleteMarkReadActionData(), ruleProperties, null);
            ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleForMarkRead });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Mark as read rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName, 1);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region Testuser1 verifies whether the specific property value is set on the received mail.
            PropertyTag[] propertyTagList = new PropertyTag[3];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagImportance;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypInteger32;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[2].PropertyId = (ushort)PropertyId.PidTagMessageFlags;
            propertyTagList[2].PropertyType = (ushort)PropertyType.PtypInteger32;

            uint contentsTableHandle = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);

            // If the PidTagImportance is the value which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property specified in the rule's action buffer structure.
            Site.Assert.AreEqual<int>(2, BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0), "The value of PidTagImportance field should be the specific value set by the client.");
            #endregion

            #region Testuser1 verifies whether the received mail is marked as read.
            int messageFlags = 0;
            bool isSubjectContainsRuleConditionSubject = false;
            if (getMailMessageContent.RowCount != 0)
            {
                isSubjectContainsRuleConditionSubject = true;
                messageFlags = BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[2].Value, 0);
            }

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R807, the value of markAsReadFlag is {0}", messageFlags);

            // Verify MS-OXORULE requirement: MS-OXORULE_R807.
            // messageFlags indicates whether the mail is marked as read. 
            // 0x00000001 is the bit which indicates the message has been marked as read. So if this bit is set in messageFlags,
            // it means the message is marked as read, which means the rule has been executed.
            bool isVerifyR807 = isSubjectContainsRuleConditionSubject && (messageFlags & 0x00000001) == 0x00000001;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR807,
                807,
                @"[In Interaction Between ST_ONLY_WHEN_OOF and ST_EXIT_LEVEL Flags] Subsequent rules (2) that have the ST_ONLY_WHEN_OOF flag set MUST be evaluated.");

            // When is not mark as read, it means rule evaluation will not terminate.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R858");

            // Verify MS-OXORULE requirement: MS-OXORULE_R858
            // If the message is marked as read, it means the rule evaluation is not terminate.
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                messageFlags & 1,
                858,
                "[In PidTagRuleState Property] EL (ST_EXIT_LEVEL, Bitmask 0x00000010): rule (2) evaluation will not terminate after executing this rule (2) if for evaluation of Out of Office rules.");

            #endregion
            #endregion

            #region Set TestUser1 back to normal state (not in OOF state).
            isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, false);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office off for {0} should succeed.", this.User1Name);
            Thread.Sleep(this.WaitForSetOOFComplete);
            #endregion

            #region TestUser1 adds an OP_Forward rule with rule sequence set to 2.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameForward);
            ForwardDelegateActionData forwardActionData = new ForwardDelegateActionData
            {
                RecipientCount = (ushort)0x01
            };
            RecipientBlock recipientBlock = new RecipientBlock
            {
                Reserved = 0x01,
                NoOfProperties = (ushort)0x04u
            };

            #region Prepare the recipient Block of the rule to forward the message to TestUser2.
            TaggedPropertyValue[] recipientProperties = AdapterHelper.GenerateRecipientPropertiesBlock(this.User2Name, this.User2ESSDN);
            recipientBlock.PropertiesData = recipientProperties;

            #endregion

            forwardActionData.RecipientsData = new RecipientBlock[1] { recipientBlock };
            RuleData ruleForward = AdapterHelper.GenerateValidRuleData(ActionType.OP_FORWARD, TestRuleDataType.ForAdd, 2, RuleState.ST_ENABLED, forwardActionData, ruleProperties, null);
            ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleForward });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Forward rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName, 2);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            #endregion

            #region Testuser1 verifies whether the specific property value is set on the received mail.
            PropertyTag[] propertyTagListInNormal = new PropertyTag[3];
            propertyTagListInNormal[0].PropertyId = (ushort)PropertyId.PidTagImportance;
            propertyTagListInNormal[0].PropertyType = (ushort)PropertyType.PtypInteger32;
            propertyTagListInNormal[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagListInNormal[1].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagListInNormal[2].PropertyId = (ushort)PropertyId.PidTagMessageFlags;
            propertyTagListInNormal[2].PropertyType = (ushort)PropertyType.PtypInteger32;

            uint contentTableHandlerNormal = 0;
            expectedMessageIndex = 0;
            RopQueryRowsResponse getMailMessageContentNormal = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandlerNormal, propertyTagListInNormal, ref expectedMessageIndex, mailSubject);

            // If the PidTagImportance is the value which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property specified in the rule's action buffer structure.
            Site.Assert.AreEqual<int>(2, BitConverter.ToInt32(getMailMessageContentNormal.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0), "The value of PidTagImportance field should be the specific value set by the client.");
            #endregion

            #region Testuser1 verifies whether the received mail is marked as read.
            bool isSubjectContainsRuleConditionSubjectNormal = false;
            int messageFlagsNormal = 0;
            string mailSubjectNormal = AdapterHelper.PropertyValueConvertToString(getMailMessageContentNormal.RowData.PropertyRows[expectedMessageIndex].PropertyValues[1].Value);
            if (mailSubjectNormal.Contains(mailSubject))
            {
                isSubjectContainsRuleConditionSubjectNormal = true;
                messageFlagsNormal = BitConverter.ToInt32(getMailMessageContentNormal.RowData.PropertyRows[expectedMessageIndex].PropertyValues[2].Value, 0);
            }

            // Verify MS-OXORULE requirement: MS-OXORULE_R67.
            // If the message is marked as read means the execution of the OOF rule was done by the OOF user. 
            bool isExcuted = (messageFlags & 0x00000001) == 0x00000001;

            // If the message is not marked as read, it means the execution of the OOF rule was not done by the normal user.  
            bool isNotExcute = isSubjectContainsRuleConditionSubjectNormal && (messageFlagsNormal & 0x00000001) != 0x00000001;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R67: the value of markAsReadFlag for OOF user is {0}, and for normal user is {0}", messageFlags, messageFlagsNormal);

            // Indicate whether the OOF rule is executed by the OOF user.
            bool isVerify67 = isExcuted && isNotExcute == true;

            Site.CaptureRequirementIfIsTrue(
                isVerify67,
                67,
                "[In PidTagRuleState] OF (ST_ONLY_WHEN_OOF, Bitmask 0x00000004): The rule (2) is executed only when a user sets the Out of Office (OOF) state on the mailbox, as specified in [MS-OXWOOF] section 2.2.5.2.");
            #endregion

            #region TestUser2 Verifies whether the OP_Forward rule is executed.
            // Let TestUser2 logon to the server.
            this.LogonMailbox(TestUser.TestUser2);

            bool doesUnexpectedMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref contentTableHandlerNormal, propertyTagListInNormal, mailSubject);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R857");

            // Verify MS-OXORULE requirement: MS-OXORULE_R857.
            // For Out of Office rules, R858 has been verified in the above step, so if the OP_Forward rule can not be executed, this requirement can be captured.
            Site.CaptureRequirementIfIsFalse(
                doesUnexpectedMessageExist,
                857,
                "[In PidTagRuleState Property] EL (ST_EXIT_LEVEL, Bitmask 0x00000010): rule (2) evaluation will terminate after executing this rule (2) except for evaluation of Out of Office rules.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the rules behavior of ST_ONLY_WHEN_OOF state. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S03_TC03_OOFBehaviorsForST_ONLY_WHEN_OOF()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameTag);
            string setOOFMailAddress = this.User1Name + "@" + this.Domain;
            string userPassword = this.User1Password;

            // Set the OOF status to false.
            bool isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, false);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office off for {0} should succeed.", this.User1Name);
            #endregion

            #region TestUser1 adds an OP_TAG rule with PidTagRuleState set to ST_ENABLED | ST_ONLY_WHEN_OOF.
            TagActionData tagActionData = new TagActionData();
            PropertyTag tagActionDataPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagImportance,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            tagActionData.PropertyTag = tagActionDataPropertyTag;
            tagActionData.PropertyValue = BitConverter.GetBytes(2);

            RuleData ruleOpTag = AdapterHelper.GenerateValidRuleData(ActionType.OP_TAG, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED | RuleState.ST_ONLY_WHEN_OOF, tagActionData, ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleOpTag });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding OP_TAG rule should succeed");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName, 1);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region Testuser1 verifies whether the specific property value is set on the received mail.
            PropertyTag[] propertyTagList = new PropertyTag[2];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagImportance;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypInteger32;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;

            uint contentsTableHandle = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);

            // If the PidTagImportance is the value which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property specified in the rule's action buffer structure.
            bool isRuleNotExecuteWhenNotInOOFState = BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0) != 2;
            Site.Assert.AreNotEqual<int>(2, BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0), "The value of PidTagImportance field should be the specific value set be the client.");
            #endregion

            #region Set TestUser1 to OOF state
            isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, true);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office on for {0} should succeed.", this.User1Name);
            Thread.Sleep(this.WaitForSetOOFComplete);
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName, 2);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region Testuser1 verifies whether the specific property value is set on the received mail.
            contentsTableHandle = 0;
            expectedMessageIndex = 0;
            getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);

            // If the PidTagImportance is the value which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property specified in the rule's action buffer structure.
            bool isRuleExecuteWhenInOOFState = BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0) == 2;

            #region Capture Code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R517: whether rule is executed or not when it is not in out of office state is {0}, and when it is in out of office state is {1}", isRuleNotExecuteWhenNotInOOFState, isRuleExecuteWhenInOOFState);

            // Verify MS-OXORULE requirement: MS-OXORULE_R517.
            bool isVerifyR517 = isRuleNotExecuteWhenNotInOOFState && isRuleExecuteWhenInOOFState;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR517,
                517,
                @"[In Processing Incoming Messages to a Folder] The server MUST evaluate rules (2) that have the ST_ONLY_WHEN_OOF flag set in the PidTagRuleState property only when the mailbox is in an OOF state as specified in [MS-OXWOOF] section 2.2.4.1.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R553: whether rule is executed or not when it is not in out of office state is {0}, and when it is in out of office state is {1}", isRuleNotExecuteWhenNotInOOFState, isRuleExecuteWhenInOOFState);

            // Verify MS-OXORULE requirement: MS-OXORULE_R553.
            bool isVerifyR553 = isVerifyR517;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR553,
                553,
                @"[In Processing Out of Office Rules] The server evaluates and executes Out of Office rules only when the mailbox is in an Out of Office state, as specified in [MS-OXWOOF] section 2.2.4.1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R582: whether rule is executed or not when it is not in out of office state is {0}, and when it is in out of office state is {1}", isRuleNotExecuteWhenNotInOOFState, isRuleExecuteWhenInOOFState);

            // Verify MS-OXORULE requirement: MS-OXORULE_R582
            bool isVerifyR582 = isVerifyR517;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR582,
                582,
                @"[In Entering and Exiting the Out of Office State] When the mailbox enters the Out of Office state as specified in [MS-OXWOOF] section 2.2.4.1, the server MUST start processing rules (2) marked with the ST_ONLY_WHEN_OOF flag in the PidTagRuleState property (section 2.2.1.3.1.3).");
            #endregion
            #endregion

            #region Set Testuser1 back to normal state (not in OOF state)
            isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, false);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office off for {0} should succeed.", this.User1Name);
            Thread.Sleep(this.WaitForSetOOFComplete);
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName, 3);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region Testuser1 verifies whether the specific property value is set on the received mail.
            contentsTableHandle = 0;
            expectedMessageIndex = 0;
            getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R584");

            // Verify MS-OXORULE requirement: MS-OXORULE_R584
            // If the PidTagImportance is the value which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property specified in the rule's action buffer structure.
            Site.CaptureRequirementIfAreNotEqual<int>(
                2,
                BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0),
                584,
                @"[In Entering and Exiting the Out of Office State] When the mailbox exits the Out of Office state, the server MUST stop processing rules (2) marked with the ST_ONLY_WHEN_OOF flag in the PidTagRuleState property.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to test delivering a message twice to check the result of a rule that has ST_KEEP_OOF_HIST set. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S03_TC04_OOFBehaviorsForST_KEEP_OOF_HIST()
        {
            this.CheckMAPIHTTPTransportSupported();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(583, this.Site), "This case runs only when the server supports to keep a list of rules with ST_KEEP_OOF_HIST flag.");

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameOOFReply);
            string setOOFMailAddress = this.User1Name + "@" + this.Domain;
            string userPassword = this.User1Password;
            #endregion

            #region Set TestUser1 to OOF state.
            bool isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, true);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office on for {0} should succeed.", this.User1Name);
            Thread.Sleep(this.WaitForSetOOFComplete);
            #endregion

            #region Create one reply template for OP_OOF_REPLY action Type.
            ulong replyTemplateMessageID;
            uint replyTemplateMessageHandle;
            TaggedPropertyValue[] addReplyBody = new TaggedPropertyValue[1];
            addReplyBody[0] = new TaggedPropertyValue();
            PropertyTag addReplyBodyPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagBody,
                PropertyType = (ushort)PropertyType.PtypString
            };
            addReplyBody[0].PropertyTag = addReplyBodyPropertyTag;
            string replyMessageBody = Common.GenerateResourceName(this.Site, Constants.MessageOfOOFReply);
            addReplyBody[0].Value = Encoding.Unicode.GetBytes(replyMessageBody + "\0");
            string replyTemplateSubject = Common.GenerateResourceName(this.Site, Constants.ReplyTemplateSubject);
            byte[] replyTemplateGUID = this.OxoruleAdapter.CreateReplyTemplate(this.InboxFolderHandle, this.InboxFolderID, true, replyTemplateSubject, addReplyBody, out replyTemplateMessageID, out replyTemplateMessageHandle);
            #endregion

            #region TestUser1 adds OP_OOF_REPLY rule with PidTagRuleState set to ST_ENABLED | ST_KEEP_OOF_HIST.

            ReplyActionData replyRuleActionData = new ReplyActionData
            {
                ReplyTemplateGUID = replyTemplateGUID,
                ReplyTemplateFID = this.InboxFolderID,
                ReplyTemplateMID = replyTemplateMessageID
            };

            RuleData ruleDataForReplyRule = AdapterHelper.GenerateValidRuleData(ActionType.OP_OOF_REPLY, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED | RuleState.ST_KEEP_OOF_HIST, replyRuleActionData, ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleDataForReplyRule });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding reply rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser2 log on to the server.
            this.LogonMailbox(TestUser.TestUser2);
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title", 1);
            this.DeliverMessageToTriggerRule(this.User1Name, this.User1ESSDN, mailSubject, null);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            PropertyTag[] propertyTagList = new PropertyTag[2];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagBody;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;

            uint contentTableHandler = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse getNormalMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandler, propertyTagList, ref expectedMessageIndex, mailSubject);
            string mailBody = AdapterHelper.PropertyValueConvertToString(getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);
            bool isBodyContainsReplyTemplateBody = mailBody.Contains(replyMessageBody);
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules again.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title", 2);
            this.DeliverMessageToTriggerRule(this.User1Name, this.User1ESSDN, mailSubject, null);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message sent by TestUser2.
            // TestUser1 log on to the server.
            this.LogonMailbox(TestUser.TestUser1);
            uint inboxFolderContentsTableHandle = 0;
            PropertyTag[] propertyTags = new PropertyTag[1];
            propertyTags[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypString;

            uint rowCount = 0;
            RopQueryRowsResponse getInboxMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref inboxFolderContentsTableHandle, propertyTags, ref rowCount, 1, mailSubject);
            Site.Assert.AreEqual<uint>(0, getInboxMailMessageContent.ReturnValue, "getInboxMailMessageContent should succeed.");
            #endregion

            #region TestUser2 verifies whether can receive the replied message.
            
            // TestUser2 log on to the server.
            this.LogonMailbox(TestUser.TestUser2);
            bool doesUnexpectedMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref contentTableHandler, propertyTagList, mailSubject);

            #region Capture code
            if (Common.IsRequirementEnabled(5572, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R5572: the expected reply message index is {0} and the total message count is {1}, and whether the message body contains the reply template body is {2}", expectedMessageIndex, getNormalMailMessageContent.RowCount, isBodyContainsReplyTemplateBody);

                // Verify MS-OXORULE requirement: MS-OXORULE_R5572.
                // The above case shows the rule was not executed twice, which indirectly indicates server adds the normal user into 
                // History List after the rule was executed once.
                Site.CaptureRequirementIfIsFalse(
                    doesUnexpectedMessageExist,
                    5572,
                    @"[In Appendix A: Product Behavior] Implementation does not evaluate the rule (2) if the sender is on the list. (Exchange 2007, Exchange 2010 and Exchange 2016 follow this behavior.)");
            }
            // If R577 is verified, which means the sender was added to the list of recipients.
            this.Site.CaptureRequirement(
                558,
                @"[In Processing Out of Office Rules] If not [the sender is not on the list] and the rule (2) condition evaluates to ""TRUE"", the server MUST add the sender to the list of recipients (2) for the rule (2) in addition to executing the rule (2) action (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R583");

            // Verify MS-OXORULE requirement: MS-OXORULE_R583
            // That is verified means when sending two messages, only the first one will execute the rule that has ST_KEEP_OOF_HIST flag, so the server must have kept a list for rules.
            bool isVerifyR583 = isBodyContainsReplyTemplateBody;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR583,
                583,
                @"[In Entering and Exiting the Out of Office State] The server MUST also keep a list for rules (2) that have the ST_KEEP_OOF_HIST flag in the PidTagRuleState property specified in section 3.2.1.2.");
            #endregion
            #endregion

            #region Set TestUser1 back to normal state (not in OOF state)
            // Testuser1 logon to the server
            this.LogonMailbox(TestUser.TestUser1);

            isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, false);
            Site.Assert.IsTrue(isSetOOFSuccess, "Cancelling Out of Office state for {0} should succeed.", this.User1Name);
            Thread.Sleep(this.WaitForSetOOFComplete);
            #endregion
        }

        /// <summary>
        /// This test case is designed to test whether the server has the same behaviors for OP_OOF_REPLY as OP_REPLY. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S03_TC05_OOFBehaviorsForOP_OOF_REPLY()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameOOFReply);
            #endregion

            #region Create one reply template for OP_OOF_REPLY action Type in TestUser1's Inbox folder.
            ulong replyTemplateMessageID;
            uint replyTemplateMessageHandle;
            TaggedPropertyValue[] addReplyBody = new TaggedPropertyValue[1];
            addReplyBody[0] = new TaggedPropertyValue();
            PropertyTag addReplyBodyPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagBody,
                PropertyType = (ushort)PropertyType.PtypString
            };
            addReplyBody[0].PropertyTag = addReplyBodyPropertyTag;
            string replyMessageBody = Common.GenerateResourceName(this.Site, Constants.MessageOfOOFReply);
            addReplyBody[0].Value = Encoding.Unicode.GetBytes(replyMessageBody + "\0");
            string replyTemplateSubject = Common.GenerateResourceName(this.Site, Constants.ReplyTemplateSubject);
            byte[] replyTemplateGUID = this.OxoruleAdapter.CreateReplyTemplate(this.InboxFolderHandle, this.InboxFolderID, true, replyTemplateSubject, addReplyBody, out replyTemplateMessageID, out replyTemplateMessageHandle);
            #endregion

            #region TestUser1 adds OP_OOF_REPLY rule with PidTagRuleState set to ST_ENABLED.
            ReplyActionData replyRuleActionData = new ReplyActionData
            {
                ReplyTemplateGUID = replyTemplateGUID,
                ReplyTemplateFID = this.InboxFolderID,
                ReplyTemplateMID = replyTemplateMessageID
            };

            RuleData ruleDataForReplyRule = AdapterHelper.GenerateValidRuleDataWithFlavor(ActionType.OP_OOF_REPLY, 1, RuleState.ST_ENABLED, replyRuleActionData, 0x00000000, ruleProperties);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleDataForReplyRule });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding reply rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser2 verifies whether can receive the OOF replied message.
            // Let Testuser2 logon to the server
            this.LogonMailbox(TestUser.TestUser2);

            PropertyTag[] propertyTagList = new PropertyTag[4];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagAutoForwarded;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypBoolean;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagBody;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[2].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[2].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[3].PropertyId = (ushort)PropertyId.PidTagMessageClass;
            propertyTagList[3].PropertyType = (ushort)PropertyType.PtypString;

            uint contentsTableHandle = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse ropQueryRowsResponse = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);

            string mailBodyTestUser2 = AdapterHelper.PropertyValueConvertToString(ropQueryRowsResponse.RowData.PropertyRows[expectedMessageIndex].PropertyValues[1].Value);
            string subject = AdapterHelper.PropertyValueConvertToString(ropQueryRowsResponse.RowData.PropertyRows[expectedMessageIndex].PropertyValues[2].Value);
            bool isBodyContainsReplyTemplateBody = mailBodyTestUser2.Contains(replyMessageBody);
            Site.Assert.IsTrue(isBodyContainsReplyTemplateBody, "Message should contain the template body!");

            #region Capture Code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R255: the replied message count is {0}, and whether the message body contains the reply template body is {1}", ropQueryRowsResponse.RowCount, isBodyContainsReplyTemplateBody);

            // Verify MS-OXORULE requirement: MS-OXORULE_R255
            // Testuser2 sent message to Testuser1,Testuser2 can receive a reply means the server send the reply
            bool isVerifyR255 = ropQueryRowsResponse.RowCount > 0 && isBodyContainsReplyTemplateBody;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR255,
                255,
                @"[In ActionBlock Structure] The meaning of action type OP_OOF_REPLY: Sends an OOF reply to the message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R922: the replied message count is {0}, and whether the message body contains the reply template body is {1}", ropQueryRowsResponse.RowCount, isBodyContainsReplyTemplateBody);

            // Verify MS-OXORULE requirement: MS-OXORULE_R922
            // Testuser2 sent message to Testuser1,Testuser2 can receive a reply means the server send the reply
            bool isVerifyR922 = ropQueryRowsResponse.RowCount > 0 && isBodyContainsReplyTemplateBody;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR922,
                922,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message:] ""OP_OOF_REPLY"": The server MUST behave as specified for the ""OP_REPLY"" action (2). [The server MUST use properties from the reply template and from the original message to create a reply to the message and then send the reply.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R953: the replied message count is {0}, and whether the message body contains the reply template body is {1}", ropQueryRowsResponse.RowCount, isBodyContainsReplyTemplateBody);

            bool isVerifyR953 = ropQueryRowsResponse.RowCount > 0 && isBodyContainsReplyTemplateBody;

            // Verify MS-OXORULE requirement: MS-OXORULE_R953.
            // The ActionFlavor is set to zero in the OP_REPLY rule, so if the messages template is the same with the reply template, means this reply is not use server-defined text in the reply message, and it's a standard reply. 
            Site.CaptureRequirementIfIsTrue(
                isVerifyR953,
                953,
                @"[In Action Flavors] [If the ActionType field value is ""OP_OOF_REPLY"", the ActionFlavor field MUST have one of the values specified in the following table [XXXXXX (ST) (NS) XXXXXXXXXXXXXXXXXXXXXXXX] or zero (0x00000000)] A value of zero (0x00000000) indicates standard reply behavior, as specified in section 3.1.4.2.5.");

            string propertyValue = AdapterHelper.PropertyValueConvertToString(ropQueryRowsResponse.RowData.PropertyRows[expectedMessageIndex].PropertyValues[3].Value);
            if (Common.IsRequirementEnabled(906, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R906: whether the message body contains the reply template body is {0}, and the PidTagMessageClass is {1}", isBodyContainsReplyTemplateBody, propertyValue);

                // Verify MS-OXORULE requirement: MS-OXORULE_R906
                // The value of PidTagMessageClass is a prefix, and the client can append a client-specific value at the end of this property,
                // so if propertyValue is start with "IPM.Note.rules.OOFTemplate", R906 can be verified.
                string prefixOfPidTagMessageClass = "IPM.Note.rules.OOFTemplate";
                bool isVerifyR906 = isBodyContainsReplyTemplateBody && propertyValue.ToUpperInvariant().StartsWith(prefixOfPidTagMessageClass.ToUpperInvariant());

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR906,
                    906,
                    @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_OOF_REPLY"": The implementation does set the value of the PidTagMessageClass property ([MS-OXCMSG] section 2.2.1.3) on the reply message to ""IPM.Note.rules.OOFTemplate"" in addition. (Exchange 2003 and above follow this behavior.)");
            }
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed for RuleState being set to 0x00000100.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S03_TC06_OOFBehaviorsForFlagSameSemanticAsST_ONLY_WHEN_OOF()
        {
            this.CheckMAPIHTTPTransportSupported();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(625, this.Site), "This case runs only when the flag 0x00000100 has the same semantics as ST_ONLY_WHEN_OOF of PidTagRuleState.");

            #region Prepare value for ruleProperties variable
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameTag);
            string setOOFMailAddress = this.User1Name + "@" + this.Domain;
            string userPassword = this.User1Password;

            // Set the OOF status to false.
            bool isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, false);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office off for {0} should succeed.", this.User1Name);
            #endregion

            #region TestUser1 adds an OP_TAG rule with PidTagRuleState set to ST_ENABLED | 0x00000100.
            TagActionData tagActionData = new TagActionData();
            PropertyTag tagActionDataPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagImportance,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            tagActionData.PropertyTag = tagActionDataPropertyTag;
            tagActionData.PropertyValue = BitConverter.GetBytes(2);

            RuleData ruleOpTag = AdapterHelper.GenerateValidRuleData(ActionType.OP_TAG, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED | RuleState.X_Same_Semantic_ST_ONLY_WHEN_OOF, tagActionData, ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleOpTag });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding OP_TAG rule should succeed");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName, 1);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region Testuser1 verifies whether the specific property value is set on the received mail.
            PropertyTag[] propertyTagList = new PropertyTag[2];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagImportance;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypInteger32;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;

            uint contentsTableHandle = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);

            // If the PidTagImportance is the value which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property specified in the rule's action buffer structure.
            bool isRuleNotExecuteWhenNotInOOFState = BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0) != 2;
            Site.Assert.AreNotEqual<int>(2, BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0), "The value of PidTagImportance field should not be set!");
            #endregion

            #region Set TestUser1 to OOF state
            isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, true);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office on for {0} should succeed.", this.User1Name);
            Thread.Sleep(this.WaitForSetOOFComplete);
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName, 2);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region Testuser1 verifies whether the specific property value is set on the received mail.
            contentsTableHandle = 0;
            expectedMessageIndex = 0;
            getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);

            // If the PidTagImportance is the value which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property specified in the rule's action buffer structure.
            bool isRuleExecuteWhenInOOFState = BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0) == 2;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R625: whether rule is executed or not when it is not in out of office state is {0}, and when it is in out of office state is {1}", isRuleNotExecuteWhenNotInOOFState, isRuleExecuteWhenInOOFState);

            // Verify MS-OXORULE requirement: MS-OXORULE_R625
            // isRuleNotExecuteWhenNotInOOFState == true and  isRuleExecuteWhenInOOFState == true means the rule with rule state 0x000001000
            // will be executed only when the mail box is in OOF state, then this requirement can be verified.
            bool isVerifyR625 = isRuleNotExecuteWhenNotInOOFState && isRuleExecuteWhenInOOFState;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR625,
                625,
                @"[In Appendix A: Product Behavior] Bit flag 0x00000100 has the same semantics as the ST_ONLY_WHEN_OOF bit flag on the implementation. [<1> Section 2.2.1.3.2.3: Bit flag 0x00000100 has the same semantics as the ST_ONLY_WHEN_OOF bit flag on Exchange 2007.]");
            #endregion

            #region Set Testuser1 back to normal state (not in OOF state)
            isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, false);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office off for {0} should succeed.", this.User1Name);
            Thread.Sleep(this.WaitForSetOOFComplete);
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName, 3);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region Testuser1 verifies whether the specific property value is set on the received mail.
            contentsTableHandle = 0;
            expectedMessageIndex = 0;
            getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R584");

            // Verify MS-OXORULE requirement: MS-OXORULE_R584
            // If the PidTagImportance is the value which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property specified in the rule's action buffer structure.
            Site.CaptureRequirementIfAreNotEqual<int>(
                2,
                BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0),
                584,
                @"[In Entering and Exiting the Out of Office State] When the mailbox exits the Out of Office state, the server MUST stop processing rules (2) marked with the ST_ONLY_WHEN_OOF flag in the PidTagRuleState property.");
            #endregion
        }

        /// <summary>
        /// This test case is designed for RuleState being set to 0x00000080.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S03_TC07_OOFBehaviorsForFlagDisableSpecificOOFRule()
        {
            this.CheckMAPIHTTPTransportSupported();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(624, this.Site), "This case runs only when implementation does use flag 0x00000080 to disable a specific Out of Office rule.");

            #region Prepare value for ruleProperties variable
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameTag);
            string setOOFMailAddress = this.User1Name + "@" + this.Domain;
            string userPassword = this.User1Password;
            #endregion

            #region TestUser1 adds an OP_TAG rule with PidTagRuleState set to ST_ENABLED | ST_ONLY_WHEN_OOF.
            TagActionData tagActionData = new TagActionData();
            PropertyTag tagActionDataPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagImportance,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            tagActionData.PropertyTag = tagActionDataPropertyTag;
            tagActionData.PropertyValue = BitConverter.GetBytes(2);

            RuleData ruleOpTag = AdapterHelper.GenerateValidRuleData(ActionType.OP_TAG, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED | RuleState.ST_ONLY_WHEN_OOF, tagActionData, ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleOpTag });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding OP_TAG rule should succeed");

            // Call RopGetRulesTable with valid TableFlags.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");

            // Get rule properties.
            PropertyTag[] propertyTags = new PropertyTag[2];
            propertyTags[0].PropertyId = (ushort)PropertyId.PidTagRuleName;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTags[1].PropertyId = (ushort)PropertyId.PidTagRuleId;
            propertyTags[1].PropertyType = (ushort)PropertyType.PtypInteger64;

            RopQueryRowsResponse queryRowsResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyTags);
            Site.Assert.AreEqual<uint>(0, queryRowsResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");

            ulong ruleId = 0;

            // Filter the correct rule.
            for (int i = 0; i < queryRowsResponse.RowCount; i++)
            {
                string ruleName = AdapterHelper.PropertyValueConvertToString(queryRowsResponse.RowData.PropertyRows[i].PropertyValues[0].Value);
                if (ruleName == ruleProperties.Name)
                {
                    ruleId = BitConverter.ToUInt64(queryRowsResponse.RowData.PropertyRows[i].PropertyValues[1].Value, 0);
                    break;
                }
            }

            #endregion

            #region Set TestUser1 to OOF state
            bool isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, true);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office on for {0} should succeed.", this.User1Name);
            Thread.Sleep(this.WaitForSetOOFComplete);
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            string messageSubjectName1 = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName, 1);

            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, messageSubjectName1);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region Testuser1 verifies whether the specific property value is set on the received mail.
            PropertyTag[] propertyTagList = new PropertyTag[2];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagImportance;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypInteger32;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;

            uint contentsTableHandle = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, messageSubjectName1);

            // If the PidTagImportance is the value which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property specified in the rule's action buffer structure.
            bool isRuleExecuteWhenInOOFState = BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0) == 2;
            Site.Assert.IsTrue(isRuleExecuteWhenInOOFState, "The OP_TAG rule should be executed in OOF state!");
            #endregion

            #region Disable the OP_TAG rule.
            ruleOpTag = AdapterHelper.GenerateValidRuleData(ActionType.OP_TAG, TestRuleDataType.ForModify, 1, RuleState.X_DisableSpecificOOFRule | RuleState.ST_ONLY_WHEN_OOF, tagActionData, ruleProperties, ruleId);
            ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleOpTag });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Modifying the OP_TAG rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            string messageSubjectName2 = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName, 2);

            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, messageSubjectName2);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region Testuser1 verifies whether the specific property value is set on the received mail.
            contentsTableHandle = 0;
            expectedMessageIndex = 0;
            getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, messageSubjectName2);

            // If the PidTagImportance is the value which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property specified in the rule's action buffer structure.
            isRuleExecuteWhenInOOFState = BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0) == 2;

            #region Capture Code

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R624");

            // Verify MS-OXORULE requirement: MS-OXORULE_R624.
            Site.CaptureRequirementIfIsFalse(
                isRuleExecuteWhenInOOFState,
                624,
                @"[In Appendix A: Product Behavior] Bit flag 0x00000080 is used to disable a specific Out of Office rule on the implementation. [<1> Section 2.2.1.3.2.3: Bit flag 0x00000080 is used to disable a specific Out of Office rule on Exchange 2007.]");

            if (Common.IsRequirementEnabled(621, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R621");

                // Verify MS-OXORULE requirement: MS-OXORULE_R621.
                // If r624 and r625 are completed, it means it does use the bit flags 0x00000080 and 0x00000100.
                Site.CaptureRequirement(
                    621,
                    "[In Appendix A: Product Behavior] The implementation uses bit flags 0x00000080 and 0x00000100 to store information about Out of Office functionality. [<1> Section 2.2.1.3.1.3: The Exchange 2007 implementation uses bit flags 0x00000080 and 0x00000100 to store information about Out of Office functionality.]");
            }
            #endregion
            #endregion

            #region Set Testuser1 back to normal state (not in OOF state)
            isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, false);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office off for {0} should succeed.", this.User1Name);
            Thread.Sleep(this.WaitForSetOOFComplete);
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_OOF_REPLY rule when the incoming message has PidTagAutoResponseSuppress property set to 0x00000010. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S03_TC08_OOFBehaviorsNotExecuteOOFReplyRuleForOOFReplySuppressMessage()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameOOFReply);
            #endregion

            #region Create one reply template for OP_OOF_REPLY action Type in TestUser1's Inbox folder.
            ulong replyTemplateMessageID;
            uint replyTemplateMessageHandle;
            TaggedPropertyValue[] addReplyBody = new TaggedPropertyValue[1];
            addReplyBody[0] = new TaggedPropertyValue();
            PropertyTag addReplyBodyPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagBody,
                PropertyType = (ushort)PropertyType.PtypString
            };
            addReplyBody[0].PropertyTag = addReplyBodyPropertyTag;
            string replyMessageBody = Common.GenerateResourceName(this.Site, Constants.MessageOfOOFReply);
            addReplyBody[0].Value = Encoding.Unicode.GetBytes(replyMessageBody + "\0");
            string replyTemplateSubject = Common.GenerateResourceName(this.Site, Constants.ReplyTemplateSubject);
            byte[] replyTemplateGUID = this.OxoruleAdapter.CreateReplyTemplate(this.InboxFolderHandle, this.InboxFolderID, true, replyTemplateSubject, addReplyBody, out replyTemplateMessageID, out replyTemplateMessageHandle);
            #endregion

            #region TestUser1 adds OP_OOF_REPLY rule with PidTagRuleState set to ST_ENABLED.
            ReplyActionData replyRuleActionData = new ReplyActionData
            {
                ReplyTemplateGUID = replyTemplateGUID,
                ReplyTemplateFID = this.InboxFolderID,
                ReplyTemplateMID = replyTemplateMessageID
            };

            RuleData ruleDataForReplyRule = AdapterHelper.GenerateValidRuleDataWithFlavor(ActionType.OP_OOF_REPLY, 1, RuleState.ST_ENABLED, replyRuleActionData, 0x00000000, ruleProperties);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleDataForReplyRule });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding reply rule should succeed.");
            #endregion

            #region TestUser2 deliver a message by ROPs to TestUser1 to trigger these rules, and PidTagAutoResponseSuppress on the message has the 0x00000010 bit set.
            // Let Testuser2 logon to the server
            this.LogonMailbox(TestUser.TestUser2);

            TaggedPropertyValue[] mailProperty = new TaggedPropertyValue[1];
            mailProperty[0] = new TaggedPropertyValue();
            PropertyTag mailPropertyPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagAutoResponseSuppress,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            mailProperty[0].PropertyTag = mailPropertyPropertyTag;
            mailProperty[0].Value = BitConverter.GetBytes(0x00000010);
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            uint submitMsgReturnValue = this.DeliverMessageToTriggerRule(this.User1Name, this.User1ESSDN, mailSubject, mailProperty);
            Site.Assert.AreEqual(0, (int)submitMsgReturnValue, "Delivering message should succeed.");
            #endregion

            #region TestUser1 gets the message sent by TestUser2.
            // TestUser1 log on to the server.
            this.LogonMailbox(TestUser.TestUser1);
            uint inboxFolderContentsTableHandle = 0;
            PropertyTag[] propertyTags = new PropertyTag[1];
            propertyTags[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypString;

            uint rowCount = 0;
            RopQueryRowsResponse getInboxMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref inboxFolderContentsTableHandle, propertyTags, ref rowCount, 1, mailSubject);
            Site.Assert.AreEqual<uint>(0, getInboxMailMessageContent.ReturnValue, "Getting the message should succeed.");
            #endregion

            #region TestUser2 verifies whether can receive the OOF replied message.

            // TestUser2 log on to the server.
            this.LogonMailbox(TestUser.TestUser2);
            PropertyTag[] propertyTagList = new PropertyTag[2];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagBody;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;
            uint contentsTableHandle = 0;
            bool doesUnexpectedMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, replyMessageBody, PropertyId.PidTagBody);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R533");

            // Verify MS-OXORULE requirement: MS-OXORULE_R533
            // TestUser2 has sent a message to TestUser1. If TestUser2 doesn't get the replied message, it means the rule is not executed. This requirement can be verified.
            Site.CaptureRequirementIfIsFalse(
                doesUnexpectedMessageExist,
                533,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_OOF_REPLY"": The server MUST NOT send a reply if the PidTagAutoResponseSuppress property on the message has the 0x00000010 bit set.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to test the server behavior for OP_OOF_REPLY when action flavor is NS. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S03_TC09_OOFBehaviorsForOP_OOF_REPLY_ActionFlavor_NS()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameOOFReply);
            string setOOFMailAddress = this.User1Name + "@" + this.Domain;
            string userPassword = this.User1Password;
            #endregion

            #region Set TestUser1 to OOF state.
            bool isSetOOFSuccess = this.SUTSetOOFAdapter.SetUserOOFSettings(setOOFMailAddress, userPassword, true);
            Site.Assert.IsTrue(isSetOOFSuccess, "Turn Out of Office on for {0} should succeed.", this.User1Name);
            Thread.Sleep(this.WaitForSetOOFComplete);
            #endregion

            #region Create one reply template for OP_OOF_REPLY action Type in TestUser1's Inbox folder.
            ulong replyTemplateMessageID;
            uint replyTemplateMessageHandle;

            TaggedPropertyValue[] replyTemplateProperties;
            TaggedPropertyValue[] temp = AdapterHelper.GenerateRecipientPropertiesBlock(this.User2Name, this.User2ESSDN);
            replyTemplateProperties = new TaggedPropertyValue[4];
            replyTemplateProperties[0] = new TaggedPropertyValue();
            PropertyTag addReplyBodyPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagBody,
                PropertyType = (ushort)PropertyType.PtypString
            };
            replyTemplateProperties[0].PropertyTag = addReplyBodyPropertyTag;
            string replyMessageBody = Common.GenerateResourceName(this.Site, Constants.MessageOfOOFReply);
            replyTemplateProperties[0].Value = Encoding.Unicode.GetBytes(replyMessageBody + "\0");
            Array.Copy(temp, 0, replyTemplateProperties, 1, temp.Length - 1);
            string replyTemplateSubject = Common.GenerateResourceName(this.Site, Constants.ReplyTemplateSubject);
            byte[] replyTemplateGUID = this.OxoruleAdapter.CreateReplyTemplate(this.InboxFolderHandle, this.InboxFolderID, true, replyTemplateSubject, replyTemplateProperties, out replyTemplateMessageID, out replyTemplateMessageHandle);
            #endregion

            #region TestUser1 adds OP_OOF_REPLY rule with actionFlavor set to NS(0x00000001).
            ReplyActionData replyRuleActionData = new ReplyActionData
            {
                ReplyTemplateGUID = replyTemplateGUID,
                ReplyTemplateFID = this.InboxFolderID,
                ReplyTemplateMID = replyTemplateMessageID
            };
            uint actionFlavor_NS = (uint)ActionFlavorsReply.NS;

            RuleData ruleDataForReplyRule = AdapterHelper.GenerateValidRuleDataWithFlavor(ActionType.OP_OOF_REPLY, 1, RuleState.ST_ENABLED, replyRuleActionData, actionFlavor_NS, ruleProperties);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleDataForReplyRule });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding reply rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser2 verifies whether can receive the OOF replied message.
            // Let Testuser2 logon to the server
            this.LogonMailbox(TestUser.TestUser2);

            PropertyTag[] propertyTagList = new PropertyTag[4];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagAutoForwarded;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypBoolean;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagBody;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[2].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[2].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[3].PropertyId = (ushort)PropertyId.PidTagMessageClass;
            propertyTagList[3].PropertyType = (ushort)PropertyType.PtypString;

            uint contentsTableHandle = 0;
            int expectedMessageIndex = 0;            

            if (Common.IsRequirementEnabled(10191, this.Site))
            {
                RopQueryRowsResponse ropQueryRowsResponse = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);
                string mailBodyTestUser2 = AdapterHelper.PropertyValueConvertToString(ropQueryRowsResponse.RowData.PropertyRows[expectedMessageIndex].PropertyValues[1].Value);
                string subject = AdapterHelper.PropertyValueConvertToString(ropQueryRowsResponse.RowData.PropertyRows[expectedMessageIndex].PropertyValues[2].Value);
                bool isBodyContainsReplyTemplateBody = mailBodyTestUser2.Contains(replyMessageBody);
                Site.Assert.IsTrue(isBodyContainsReplyTemplateBody, "Message should contain the template body!");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R10191: the replied message count is {0}, and whether the message body contains the reply template body is {1}", ropQueryRowsResponse.RowCount, isBodyContainsReplyTemplateBody);

                // Verify MS-OXORULE requirement: MS-OXORULE_R10191
                // Testuser2 sent message to Testuser1,Testuser2 can receive a reply means the server send the reply
                bool isVerifyR10191 = ropQueryRowsResponse.RowCount > 0 && isBodyContainsReplyTemplateBody;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR10191,
                    10191,
                    @"[In Appendix A: Product Behavior] Implementation does send a reply message if the ActionType is ""OP_OOF_REPLY"" if action flavor is NS. (<6> Section 2.2.5.1.1:  Exchange 2007, Exchange 2010, and Exchange 2013 send a reply message if the ActionType is ""OP_OOF_REPLY"".)");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R1020");

                // Verify MS-OXORULE requirement: MS-OXORULE_R1020
                // The recipients information is added in the reply template when it is created. Since the rule works as expected, this requirement can be captured.
                Site.CaptureRequirement(
                    1020,
                    @"[In Action Flavors] NS (Bitmask 0x00000001): [OP_OOF_REPLY ActionType] [The server SHOULD<6> not send the message to the message sender]The reply template MUST contain recipients (2) in this case [if the NS flag is set].");
            }

            #endregion
        }
    }
}