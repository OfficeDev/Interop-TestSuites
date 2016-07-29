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
    /// This scenario aims to validate server behaviors of processing server-side rules other than Out of Office rule 
    /// because action of OP_OOF_REPLY is complicated enough to be a separate scenario.
    /// </summary>
    [TestClass]
    public class S02_ProcessServerSideRulesOtherthanOOF : TestSuiteBase
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
        /// This test case is designed to validate the execution of OP_BOUNCE rule.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC01_ServerExecuteRule_Action_OP_BOUNCE()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameBounce);
            #endregion

            #region  TestUser1 adds a Bounce rule with Bounce of action data set to CanNotDisplay.
            BounceActionData bounceActionData = new BounceActionData
            {
                Bounce = BounceCode.CanNotDisplay
            };

            RuleData ruleBounce = AdapterHelper.GenerateValidRuleData(ActionType.OP_BOUNCE, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, bounceActionData, ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleBounce });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Bounce rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.

            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // Let TestUser2 logon to the server.
            this.LogonMailbox(TestUser.TestUser2);
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title", 1);
            this.DeliverMessageToTriggerRule(this.User1Name, this.User1ESSDN, mailSubject, null);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 verifies there is no new message in the Inbox folder.
            // Let TestUser1 logon to the server.
            this.LogonMailbox(TestUser.TestUser1);

            // Set PidTagSubject and PidTagMessageFlags visible.
            PropertyTag[] propertyTag = new PropertyTag[2];
            propertyTag[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTag[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTag[1].PropertyId = (ushort)PropertyId.PidTagMessageFlags;
            propertyTag[1].PropertyType = (ushort)PropertyType.PtypInteger32;

            // Get mail message content.
            uint contentsTableHandle = 0;
            bool doesUnexpectedMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref contentsTableHandle, propertyTag, mailSubject);

            #region Capture Code
            if (Common.IsRequirementEnabled(5472, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R5472");

                // Verify MS-OXORULE requirement: MS-OXORULE_R5472.
                // TestUser2 has sent a message to TestUser1. If the message doesn't appear in the user's mailbox, this requirement can be verified.
                Site.CaptureRequirementIfIsFalse(
                    doesUnexpectedMessageExist,
                    5472,
                    @"[In Appendix A: Product Behavior] [""OP_BOUNCE""]Implementation does not support the original message appears in the user's mailbox. (Exchange 2010 and above follow this behavior.)");
            }
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R342");

            // If the message doesn't appear in the user's mailbox, this requirement can be verified.
            Site.CaptureRequirementIfIsFalse(
                doesUnexpectedMessageExist,
                342,
                @"[In OP_BOUNCE ActionData Structure] The meaning of the BounceCode value 0x0000001F: The message was rejected because it cannot be displayed to the user.");
            #endregion
            #endregion

            #region TestUser2 verifies there is a reply message in the Inbox folder.
            // Let TestUser2 logon to the server.
            this.LogonMailbox(TestUser.TestUser2);

            PropertyTag[] propertyTagList = new PropertyTag[2];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagMessageFlags;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypInteger32;

            int expectedMessageIndex = 0;
            uint contentTableHandle = 0;
            RopQueryRowsResponse getMailMessageContent2 = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);
            string mailSubjectForCanNotDisplay = AdapterHelper.PropertyValueConvertToString(getMailMessageContent2.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);
            bool isReceivedReplyMessageForCanNotDisplay = mailSubjectForCanNotDisplay.Contains(mailSubject);
            Site.Assert.IsTrue(isReceivedReplyMessageForCanNotDisplay, "The server should send a reply message to the sender for the Can_Not_Display Bounce rule.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R260");

            // Verify MS-OXORULE requirement: MS-OXORULE_R260.
            // TestUser2 has send a message to TestUser1, if the message cannot be found in TestUser1's mailbox and TestUser2 received the replied message,
            // it means the server has executed the BOUNCE rule.
            bool isVerifyR260 = (!doesUnexpectedMessageExist) && isReceivedReplyMessageForCanNotDisplay;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR260,
                260,
                @"[In ActionBlock Structure] The meaning of action type OP_BOUNCE: Rejects the message back to the sender.");
            #endregion

            #region TestUser1 adds a Bounce rule with Bounce of action data set to TooLarge.
            // Let TestUser1 log on to the server.
            this.LogonMailbox(TestUser.TestUser1);
            bounceActionData = new BounceActionData
            {
                Bounce = BounceCode.TooLarge
            };

            ruleBounce = AdapterHelper.GenerateValidRuleData(ActionType.OP_BOUNCE, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, bounceActionData, ruleProperties, null);
            ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleBounce });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Bounce rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.

            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // Let TestUser2 logon to the server.
            this.LogonMailbox(TestUser.TestUser2);
            mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title", 2);
            this.DeliverMessageToTriggerRule(this.User1Name, this.User1ESSDN, mailSubject, null);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 verifies the server rejects the too large messages.
            // Let TestUser1 logon to the server.
            this.LogonMailbox(TestUser.TestUser1);
            doesUnexpectedMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref contentsTableHandle, propertyTag, mailSubject);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R341");

            // TestUser2 has sent a message to TestUser1. If the message doesn't appear in the user's mailbox, this requirement can be verified.
            Site.CaptureRequirementIfIsFalse(
                doesUnexpectedMessageExist,
                341,
                @"[In OP_BOUNCE ActionData Structure] The meaning of the BounceCode value 0x0000000D: The message was rejected because it was too large.");
            #endregion

            #region TestUser2 verifies there is a reply message in the Inbox folder.
            // Let TestUser2 logon to the server.
            this.LogonMailbox(TestUser.TestUser2);
            expectedMessageIndex = 0;
            contentTableHandle = 0;
            RopQueryRowsResponse getMailMessageContentForTooLarge = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);
            string mailSubjectForTooLarge = AdapterHelper.PropertyValueConvertToString(getMailMessageContentForTooLarge.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);
            Site.Assert.AreEqual<uint>(0, getMailMessageContentForTooLarge.ReturnValue, "Getting the message should succeed, the actual returned value is {0}!", getMailMessageContentForTooLarge.ReturnValue);
            bool isReceivedReplyMessageForTooLarge = mailSubjectForTooLarge.Contains(mailSubject);
            Site.Assert.IsTrue(isReceivedReplyMessageForTooLarge, "The server should send a reply message to the sender for the Too_Large Bounce rule.");

            // Let TestUser1 logon to the server.
            this.LogonMailbox(TestUser.TestUser1);
            #endregion

            // The Denied Bounce rule cannot be verified on Exchange 2007.
            if (Common.IsRequirementEnabled(343, this.Site))
            {
                #region TestUser1 adds a Bounce rule with Bounce of action data set to Denied.
                bounceActionData = new BounceActionData
                {
                    Bounce = BounceCode.Denied
                };

                ruleBounce = AdapterHelper.GenerateValidRuleData(ActionType.OP_BOUNCE, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, bounceActionData, ruleProperties, null);
                ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleBounce });
                Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Bounce rule should succeed.");
                #endregion

                #region TestUser2 delivers a message to TestUser1 to trigger these rules.

                Thread.Sleep(this.WaitForTheRuleToTakeEffect);

                // Let TestUser2 logon to the server.
                this.LogonMailbox(TestUser.TestUser2);
                mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title", 3);
                this.DeliverMessageToTriggerRule(this.User1Name, this.User1ESSDN, mailSubject, null);
                Thread.Sleep(this.WaitForTheRuleToTakeEffect);
                #endregion

                #region TestUser1 gets the messages in its Inbox folder
                // Let TestUser1 logon to the server.
                this.LogonMailbox(TestUser.TestUser1);
                doesUnexpectedMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref contentsTableHandle, propertyTag, mailSubject);

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R343");

                // TestUser2 has sent a message to TestUser1. If the message doesn't appear in the user's mailbox, this requirement can be verified.
                Site.CaptureRequirementIfIsFalse(
                    doesUnexpectedMessageExist,
                    343,
                    @"[In OP_BOUNCE ActionData Structure] The meaning of the BounceCode value 0x00000026: The message delivery was denied for other reasons.");
                #endregion

                #region TestUser2 verifies there is a reply message in the Inbox folder.
                // Let TestUser2 logon to the server.
                this.LogonMailbox(TestUser.TestUser2);
                expectedMessageIndex = 0;
                contentTableHandle = 0;
                RopQueryRowsResponse getMailMessageContentForDenied = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);
                string mailSubjectForDenied = AdapterHelper.PropertyValueConvertToString(getMailMessageContentForDenied.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);
                Site.Assert.AreEqual<uint>(0, getMailMessageContentForDenied.ReturnValue, "Getting the message should succeed, the actual returned value is {0}!", getMailMessageContentForDenied.ReturnValue);
                bool isReceivedReplyMessageForDenied = mailSubjectForDenied.Contains(mailSubject);
                Site.Assert.IsTrue(isReceivedReplyMessageForDenied, "The server should send a reply message to the sender for the Denied Bounce rule.");
                if (Common.IsRequirementEnabled(5462, this.Site))
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R5462");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R5462.
                    bool isVerifyR546 = isReceivedReplyMessageForCanNotDisplay && isReceivedReplyMessageForTooLarge && isReceivedReplyMessageForDenied;
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR546,
                        5462,
                        @"[In Appendix A: Product Behavior] [""OP_BOUNCE""]Implementation does send a reply message to the sender detailing why the sender's message couldn't be delivered to the user's mailbox. (Exchange 2010 and above follow this behavior.)");
                }
                #endregion
            }
            else
            {
                if (Common.IsRequirementEnabled(5462, this.Site))
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R5462");

                    // Verify MS-OXORULE requirement: MS-OXORULE_R5462.
                    bool isVerifyR546 = isReceivedReplyMessageForCanNotDisplay && isReceivedReplyMessageForTooLarge;
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR546,
                        5462,
                        @"[In Appendix A: Product Behavior] [""OP_BOUNCE""]Implementation does send a reply message to the sender detailing why the sender's message couldn't be delivered to the user's mailbox. (Exchange 2010 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_REPLY rule. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC02_ServerExecuteRule_Action_OP_REPLY()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameReply);
            #endregion

            #region Create a reply template in the TestUser1's Inbox folder.
            ulong replyTemplateMessageId;
            uint replyTemplateMessageHandler;
            string replyTemplateSubject = Common.GenerateResourceName(this.Site, Constants.ReplyTemplateSubject);

            TaggedPropertyValue[] replyTemplateProperties = new TaggedPropertyValue[1];
            replyTemplateProperties[0] = new TaggedPropertyValue();
            PropertyTag replyTemplatePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagBody,
                PropertyType = (ushort)PropertyType.PtypString
            };
            replyTemplateProperties[0].PropertyTag = replyTemplatePropertyTag;
            replyTemplateProperties[0].Value = Encoding.Unicode.GetBytes(Constants.ReplyTemplateBody + "\0");

            byte[] guidBytes = this.OxoruleAdapter.CreateReplyTemplate(this.InboxFolderHandle, this.InboxFolderID, false, replyTemplateSubject, replyTemplateProperties, out replyTemplateMessageId, out replyTemplateMessageHandler);
            #endregion

            #region TestUser1 adds a reply rule with 0x00000000 action flavor to TestUser1's Inbox folder.
            ReplyActionData replyActionData = new ReplyActionData
            {
                ReplyTemplateGUID = new byte[guidBytes.Length]
            };
            Array.Copy(guidBytes, 0, replyActionData.ReplyTemplateGUID, 0, guidBytes.Length);

            replyActionData.ReplyTemplateFID = this.InboxFolderID;
            replyActionData.ReplyTemplateMID = replyTemplateMessageId;

            RuleData ruleForReply = AdapterHelper.GenerateValidRuleDataWithFlavor(ActionType.OP_REPLY, 0, RuleState.ST_ENABLED, replyActionData, 0x00000000, ruleProperties);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForReply });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding rule with actionFlavor set to 0x00000000 should succeed.");
            #endregion

            #region TestUser2 sends a mail to the TestUser1 to trigger this rule.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser2 deliver a message to trigger these rules
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 verifies there are messages in the specific folder then get the sentRepresentingEmailAddress.
            // Specify the message properties to be got.
            PropertyTag[] propertyTagList = new PropertyTag[4];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagBody;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[2].PropertyId = (ushort)PropertyId.PidTagSentRepresentingEmailAddress;
            propertyTagList[2].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[3].PropertyId = (ushort)PropertyId.PidTagHasDeferredActionMessages;
            propertyTagList[3].PropertyType = (ushort)PropertyType.PtypBoolean;

            uint contentTableHandler = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse getNormalMailMessageContentOnFlavorRule = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandler, propertyTagList, ref expectedMessageIndex, mailSubject);

            string sentRepresentingEmailAddress = AdapterHelper.PropertyValueConvertToString(getNormalMailMessageContentOnFlavorRule.RowData.PropertyRows[expectedMessageIndex].PropertyValues[2].Value);

            // Get the value of PidTagHasDeferredActionMessages.
            byte[] pidTagHasDeferredActionMessagesOfMessageInInboxFolder = getNormalMailMessageContentOnFlavorRule.RowData.PropertyRows[expectedMessageIndex].PropertyValues[3].Value;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R886.");

            // Verify MS-OXORULE requirement: MS-OXORULE_R886.
            // If a rule has no DAM, its pidTagHasDeferredActionMessages property is false or not existing.
            // So if check whether this property exists and its value is false (0x00) or it does not exist (the value is 0x8004010f), it means the rule has no DAM.
            bool isVerifyR886 = Common.CompareByteArray(pidTagHasDeferredActionMessagesOfMessageInInboxFolder, new byte[] { 0x00 }) ||
                             Common.CompareByteArray(pidTagHasDeferredActionMessagesOfMessageInInboxFolder, new byte[] { 0x0f, 0x01, 0x04, 0x80 });

            Site.CaptureRequirementIfIsTrue(
                isVerifyR886,
                886,
                @"[In PidTagHasDeferredActionMessages Property] This property MUST be set to ""false"" if a message has no associated DAM.");
            #endregion

            #region TestUser2 verifies there are reply messages in the specific folder.
            // Let TestUser2 log on to the server.
            this.LogonMailbox(TestUser.TestUser2);

            PropertyTag[] propertyTagList1 = new PropertyTag[3];
            propertyTagList1[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList1[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList1[1].PropertyId = (ushort)PropertyId.PidTagBody;
            propertyTagList1[1].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList1[2].PropertyId = (ushort)PropertyId.PidTagReceivedByEmailAddress;
            propertyTagList1[2].PropertyType = (ushort)PropertyType.PtypString;

            expectedMessageIndex = 0;
            getNormalMailMessageContentOnFlavorRule = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandler, propertyTagList1, ref expectedMessageIndex, replyTemplateSubject);

            #region Capture Code
            // Subject, bodyText and originalMessageSender are the properties set by the server on the replied message.
            string subject = AdapterHelper.PropertyValueConvertToString(getNormalMailMessageContentOnFlavorRule.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);
            string bodyText = AdapterHelper.PropertyValueConvertToString(getNormalMailMessageContentOnFlavorRule.RowData.PropertyRows[expectedMessageIndex].PropertyValues[1].Value);
            string receivedByEmailAddress = AdapterHelper.PropertyValueConvertToString(getNormalMailMessageContentOnFlavorRule.RowData.PropertyRows[expectedMessageIndex].PropertyValues[2].Value);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R525: the actual value of property PidTagSubject and PidTagReceivedByEmailAddress separately is {0}, {1}", subject, receivedByEmailAddress.ToUpperInvariant());

            // Verify MS-OXORULE requirement: MS-OXORULE_R525.
            // Subject is the property from the reply template, ReceivedByEmailAddress is the property from the original message's SentRepresentingEmailAddress property, message body is the text in the reply template.
            // The message class property is set to IPM.Note in sutcontrolAdapter when being sent.
            bool isVerifyR525 = subject == replyTemplateSubject && receivedByEmailAddress.ToUpperInvariant() == sentRepresentingEmailAddress.ToUpperInvariant() && bodyText.ToUpperInvariant() == Constants.ReplyTemplateBody.ToUpperInvariant();

            Site.CaptureRequirementIfIsTrue(
                isVerifyR525,
                525,
                @"[In Processing Incoming Messages to a Folder] Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message: ""OP_REPLY"": The server MUST use properties from the reply template and from the original message to create a reply to the message and then send the reply.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R253: the actual value of the PidTagSubject is {0}, and the name of the template is {1}", subject, replyTemplateSubject);

            // Verify MS-OXORULE requirement: MS-OXORULE_R253.
            // Subject is the property from the reply template, and the name of the template is replyTemplateSubject,
            // if they are same with each other this requirement can be verified.
            bool isVerifyR253 = subject == replyTemplateSubject;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR253,
                253,
                @"[In ActionBlock Structure] The meaning of action type OP_REPLY: Replies to the message.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R952: the actual value of the PidTagSubject is {0}, and the name of the template is {1}", subject, replyTemplateSubject);
            bool isVerifiedR952 = subject == replyTemplateSubject && bodyText.ToUpperInvariant() == Constants.ReplyTemplateBody.ToUpperInvariant();

            // Verify MS-OXORULE requirement: MS-OXORULE_R952.
            // The ActionFlavor is set to zero in the OP_REPLY rule, so if the messages template is the same with the reply template, means this reply is not use server-defined text in the reply message, and it's a standard reply. 
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR952,
                952,
                @"[In Action Flavors] [If the ActionType field value is ""OP_REPLY"", the ActionFlavor field MUST have one of the values specified in the following table [XXXXXX (ST) (NS) XXXXXXXXXXXXXXXXXXXXXXXX] or zero (0x00000000)] A value of zero (0x00000000) indicates standard reply behavior, as specified in section 3.1.4.2.5).");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_REPLY rule with action flavor NS.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC03_ServerExecuteRule_Action_OP_REPLY_ActionFlavor_NS()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameReply);
            #endregion

            #region Create a reply template in the TestUser1's Inbox folder.
            ulong replyTemplateMessageId;
            uint replyTemplateMessageHandler;

            TaggedPropertyValue[] replyTemplateProperties;

            // Add recipient information .
            TaggedPropertyValue[] temp = AdapterHelper.GenerateRecipientPropertiesBlock(this.User2Name, this.User2ESSDN);
            replyTemplateProperties = new TaggedPropertyValue[4];
            replyTemplateProperties[0] = new TaggedPropertyValue();
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagBody,
                PropertyType = (ushort)PropertyType.PtypString
            };
            replyTemplateProperties[0].PropertyTag = propertyTag;
            replyTemplateProperties[0].Value = Encoding.Unicode.GetBytes(Constants.ReplyTemplateBody + "\0");
            Array.Copy(temp, 0, replyTemplateProperties, 1, temp.Length - 1);

            string replyTemplateSubject = Common.GenerateResourceName(this.Site, Constants.ReplyTemplateSubject);
            byte[] guidBytes = this.OxoruleAdapter.CreateReplyTemplate(this.InboxFolderHandle, this.InboxFolderID, false, replyTemplateSubject, replyTemplateProperties, out replyTemplateMessageId, out replyTemplateMessageHandler);
            #endregion

            #region TestUser1 adds a reply rule with actionFlavor set to NS(0x00000001).
            ReplyActionData replyActionDataWithFlavor = new ReplyActionData
            {
                ReplyTemplateGUID = new byte[guidBytes.Length]
            };
            Array.Copy(guidBytes, 0, replyActionDataWithFlavor.ReplyTemplateGUID, 0, guidBytes.Length);

            replyActionDataWithFlavor.ReplyTemplateFID = this.InboxFolderID;
            replyActionDataWithFlavor.ReplyTemplateMID = replyTemplateMessageId;
            uint actionFlavor_NS = (uint)ActionFlavorsReply.NS;

            RuleData ruleForReplyWithFlavor = AdapterHelper.GenerateValidRuleDataWithFlavor(ActionType.OP_REPLY, 0, RuleState.ST_ENABLED, replyActionDataWithFlavor, actionFlavor_NS, ruleProperties);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForReplyWithFlavor });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding rule with actionFlavor set to NS should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser2 delivers a message to trigger the rules.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message sent by TestUser2.
            uint inboxFolderContentsTableHandle = 0;
            PropertyTag[] propertyTags = new PropertyTag[1];
            propertyTags[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypString;

            uint rowCount = 0;
            RopQueryRowsResponse getInboxMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref inboxFolderContentsTableHandle, propertyTags, ref rowCount, 1, mailSubject);
            Site.Assert.AreEqual<uint>(0, getInboxMailMessageContent.ReturnValue, "Getting the message should succeed.");
            #endregion

            #region TestUser2 verifies there is no new message in the Inbox folder.
            // Let TestUser2 log on to the server.
            this.LogonMailbox(TestUser.TestUser2);

            // Specify the message properties to be got.
            PropertyTag[] propertyTagList = new PropertyTag[1];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;

            uint contentsTableHandle = 0;
            bool doesUnexpectedMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, replyTemplateSubject);

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R288");

            // Verify MS-OXORULE requirement: MS-OXORULE_R288.
            // TestUser2 has sent a message to TestUser1. If TestUser2 doesn't get the replied message, this requirement can be verified.
            Site.CaptureRequirementIfIsFalse(
                doesUnexpectedMessageExist,
                288,
                @"[In Action Flavors] NS (Bitmask 0x00000001): [OP_REPLY ActionType] Do not send the message to the message sender.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R663");

            // Verify MS-OXORULE requirement: MS-OXORULE_R663
            // The recipients (2) information is added in the reply template when it is created. Since the rule works as expected, this requirement can be captured.
            Site.CaptureRequirement(
                663,
                @"[In Action Flavors] NS (Bitmask 0x00000001): [OP_REPLY ActionType] The reply template MUST contain recipients (2) in this case [if the NS flag is set].");

            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_REPLY rule with PidTagAutoResponseSuppress set to 0x00000020.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC04_ServerExecuteRule_Action_OP_REPLY_PidTagAutoResponseSuppress()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameReply);
            #endregion

            #region Create a reply template in the TestUser1's Inbox folder.
            ulong replyTemplateMessageId;
            uint replyTemplateMessageHandler;
            TaggedPropertyValue[] replyTemplateProperties = new TaggedPropertyValue[1];
            replyTemplateProperties[0] = new TaggedPropertyValue();
            PropertyTag replyTemplatePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagBody,
                PropertyType = (ushort)PropertyType.PtypString
            };
            replyTemplateProperties[0].PropertyTag = replyTemplatePropertyTag;
            replyTemplateProperties[0].Value = Encoding.Unicode.GetBytes(Constants.ReplyTemplateBody + "\0");

            string replyTemplateSubject = Common.GenerateResourceName(this.Site, Constants.ReplyTemplateSubject);
            byte[] guidBytes = this.OxoruleAdapter.CreateReplyTemplate(this.InboxFolderHandle, this.InboxFolderID, false, replyTemplateSubject, replyTemplateProperties, out replyTemplateMessageId, out replyTemplateMessageHandler);
            #endregion

            #region TestUser1 adds a reply rule with 0x00000000 action flavor to TestUser1's Inbox folder.
            ReplyActionData replyActionData = new ReplyActionData
            {
                ReplyTemplateGUID = new byte[guidBytes.Length]
            };
            Array.Copy(guidBytes, 0, replyActionData.ReplyTemplateGUID, 0, guidBytes.Length);
            replyActionData.ReplyTemplateFID = this.InboxFolderID;
            replyActionData.ReplyTemplateMID = replyTemplateMessageId;

            RuleData ruleForReply = AdapterHelper.GenerateValidRuleDataWithFlavor(ActionType.OP_REPLY, 0, RuleState.ST_ENABLED, replyActionData, 0x00000000, ruleProperties);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForReply });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding reply rule with actionFlavor set to 0x00000000 should succeed.");
            #endregion

            #region TestUser2 sends a mail to TestUser1 with PidTagAutoResponseSuppress set to 0x00000020.
            // Let TestUser2 log on to the server.
            this.LogonMailbox(TestUser.TestUser2);

            TaggedPropertyValue autoResponseSuppressProperty = new TaggedPropertyValue();
            PropertyTag autoResponseSuppressPropertyPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagAutoResponseSuppress,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            autoResponseSuppressProperty.PropertyTag = autoResponseSuppressPropertyPropertyTag;
            autoResponseSuppressProperty.Value = BitConverter.GetBytes(0x00000020);

            string subject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.DeliverMessageToTriggerRule(this.User1Name, this.User1ESSDN, subject, new TaggedPropertyValue[1] { autoResponseSuppressProperty });
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser2 verifies the server doesn't send a reply to normal user.
            uint contentsTableHandle = 0;
            int expectedMessageIndex = 0;

            // Specify the message properties to be got.
            PropertyTag[] propertyTagarray = new PropertyTag[1];

            // PidTagSubject
            propertyTagarray[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagarray[0].PropertyType = (ushort)PropertyType.PtypString;
            bool doesUnexpectedMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref contentsTableHandle, propertyTagarray, replyTemplateSubject);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R526");

            // Verify MS-OXORULE requirement: MS-OXORULE_R526.
            // This test case is designed based on the PidTagAutoResponseSuppress property on the message that has the 0x00000020 bit set.
            // TestUser2 has sent a message to TestUser1. If TestUser2 doesn't get the replied message, this requirement can be verified.
            Site.CaptureRequirementIfIsFalse(
                doesUnexpectedMessageExist,
                526,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_REPLY"": The server MUST NOT send a reply if the PidTagAutoResponseSuppress property ([MS-OXOMSG] section 2.2.1.77) on the message that has the 0x00000020 bit set.");

            #endregion

            #region TestUser1 verifies there are messages which are the sent message in the Inbox folder.
            // Let TestUser1 logon to the server
            this.LogonMailbox(TestUser.TestUser1);

            PropertyTag[] propertyTags = new PropertyTag[1];
            propertyTags[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypString;

            expectedMessageIndex = 0;
            RopQueryRowsResponse mailMessageContentInUser1 = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTags, ref expectedMessageIndex, subject);
            Site.Assert.AreEqual<uint>(0, mailMessageContentInUser1.ReturnValue, "Getting mail message operation should succeed.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_DELETE rule. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC05_ServerExecuteRule_Action_OP_DELETE()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameCopy);
            #endregion

            #region TestUser1 creates a new folder in server store.
            RopCreateFolderResponse createFolderResponse;
            uint newFolderHandle = this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, Common.GenerateResourceName(this.Site, "User1Folder"), "TestForOP_COPY", out createFolderResponse);
            ulong newFolderId = createFolderResponse.FolderId;
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating folder operation should succeed.");
            #endregion

            #region Prepare rules' data.
            MoveCopyActionData moveCopyActionData = new MoveCopyActionData();

            // Get the created folder entry ID.
            ServerEID serverEID = new ServerEID(BitConverter.GetBytes(newFolderId));
            byte[] folderEId = serverEID.Serialize();

            // Get the store object's entry ID.
            byte[] storeEId = this.GetStoreObjectEntryID(StoreObjectType.Mailbox, this.Server, this.User1ESSDN);
            moveCopyActionData.FolderInThisStore = 1;
            moveCopyActionData.FolderEID = folderEId;
            moveCopyActionData.StoreEID = storeEId;
            moveCopyActionData.FolderEIDSize = (ushort)folderEId.Length;
            moveCopyActionData.StoreEIDSize = (ushort)storeEId.Length;
            RuleData ruleForCopy = AdapterHelper.GenerateValidRuleData(ActionType.OP_COPY, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, moveCopyActionData, ruleProperties, null);
            #endregion

            #region TestUser1 adds OP_COPY rule to the Inbox folder.
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForCopy });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Copy rule should succeed.");
            #endregion

            #region TestUser1 adds OP_DELETE rule for Inbox folder with rule Sequence set to 2.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameDelete);
            RuleData ruleForDelete = AdapterHelper.GenerateValidRuleData(ActionType.OP_DELETE, TestRuleDataType.ForAdd, 2, RuleState.ST_ENABLED, new DeleteMarkReadActionData(), ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleForDelete });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding the delete rule should succeed.");
            #endregion

            #region TestUser1 adds OP_FORWARD rule with rule sequence set to 3.
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

            #region Prepare the recipient Block.
            TaggedPropertyValue[] recipientProperties = AdapterHelper.GenerateRecipientPropertiesBlock(this.User2Name, this.User2ESSDN);

            recipientBlock.PropertiesData = recipientProperties;
            #endregion

            forwardActionData.RecipientsData = new RecipientBlock[1] { recipientBlock };
            RuleData ruleForward = AdapterHelper.GenerateValidRuleData(ActionType.OP_FORWARD, TestRuleDataType.ForAdd, 3, RuleState.ST_ENABLED, forwardActionData, ruleProperties, null);
            ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleForward });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Forward rule should succeed");
            #endregion

            #region Get the rule properties in the Inbox of TestUser1.
            RopGetRulesTableResponse ropGetRulesTableResponse = new RopGetRulesTableResponse();
            uint ruleTableHandler = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table operation should succeed.");

            PropertyTag[] propertyTags = new PropertyTag[2];

            // PidTagRuleFolderEntryId
            propertyTags[0].PropertyId = (ushort)PropertyId.PidTagRuleFolderEntryId;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypBinary;

            // PidTagRuleLevel
            propertyTags[1].PropertyId = (ushort)PropertyId.PidTagRuleLevel;
            propertyTags[1].PropertyType = (ushort)PropertyType.PtypInteger32;

            RopQueryRowsResponse ropQueryRowsResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandler, propertyTags);
            Site.Assert.AreEqual<uint>(0, ropQueryRowsResponse.ReturnValue, "Query rows operation should succeed.");

            // Three rules have been added to the Inbox folder, so the row count in the rule table should be 3.
            Site.Assert.AreEqual<uint>(3, ropQueryRowsResponse.RowCount, "The rule number in the rule table is {0}", ropQueryRowsResponse.RowCount);
            this.VerifyRuleTable();
            #endregion

            #region TestUser2 deliver a message to TestUser1 to trigger these rules
            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 checks whether the origin message is deleted.
            // Set PidTagSubject and PidTagMessageFlags visible.
            PropertyTag[] propertyTag = new PropertyTag[2];
            propertyTag[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTag[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTag[1].PropertyId = (ushort)PropertyId.PidTagMessageFlags;
            propertyTag[1].PropertyType = (ushort)PropertyType.PtypInteger32;

            // Get mail message content.
            uint contentsTableHandle = 0;
            bool doesUnexpectedMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref contentsTableHandle, propertyTag, mailSubject);

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R931");

            // Verify MS-OXORULE requirement: MS-OXORULE_R931.
            // In this test case, the server-side rule is OP_DELETE, and the actions specified in the PidTagRuleActions property associated with this rule is to delete the incoming message. 
            // TestUser2 has sent a message to TestUser1. If the message doesn't appear in TestUser1's mailbox, it means the server has deleted the incoming message. This requirement can be verified.
            Site.CaptureRequirementIfIsFalse(
                doesUnexpectedMessageExist,
                931,
                @"[In Processing Incoming Messages to a Folder] When executing a rule (2) whose condition evaluates to ""TRUE"" as per the restriction (2) in the PidTagRuleCondition property (section 2.2.1.3.1.9), then the server MUST perform the actions (2) specified in the PidTagRuleActions property (section 2.2.1.3.1.10) associated with that rule (2) in the case of a server-side rule.");

            if (Common.IsRequirementEnabled(898, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R898");

                Site.CaptureRequirementIfIsFalse(
                    doesUnexpectedMessageExist,
                    898,
                    @"[In OP_DELETE or OP_MARK_AS_READ Data Buffer Format] For the OP_DELETE action type, the implementation does delete the incoming messages according to the ActionType itself. (Windows Exchange 2003 and above follow this behavior)");
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R549");

            // Verify MS-OXORULE requirement: MS-OXORULE_R549.
            // TestUser2 has sent a message to TestUser1. If the message doesn't appear in TestUser1's mailbox, it means the server has deleted the incoming message. This requirement can be verified.
            Site.CaptureRequirementIfIsFalse(
                doesUnexpectedMessageExist,
                549,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_DELETE"": The server MUST delete the message.");

            #endregion
            #endregion

            #region TestUser2 checks whether a forward message is received.
            this.LogonMailbox(TestUser.TestUser2);
            bool doesUnexpectedMessageExist2 = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref contentsTableHandle, propertyTag, mailSubject);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R550");

            // Verify MS-OXORULE requirement: MS-OXORULE_R550.
            // In this test case, if the forward message doesn't appear in TestUser2's mailbox, it means the subsequent OP_FORWARD rule, which is not an Out of Office rule, is not executed.
            Site.CaptureRequirementIfIsFalse(
                doesUnexpectedMessageExist2,
                550,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_DELETE"": The server MUST stop evaluating subsequent rules (2) on the message except for Out of Office rules.");
            #endregion

            #region Delete the newly created folder.
            this.LogonMailbox(TestUser.TestUser1);
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, newFolderId);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_ MARK_AS_READ rule.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC06_ServerExecuteRule_Action_OP_MARK_AS_READ()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameMarkAsRead);
            #endregion

            #region TestUser1 adds a OP_MARK_AS_READ rule with rule_state set to a value contains all other flags except ST_ENABLED and ST_ONLY_WHEN_OOF.
            RuleState ruleState = RuleState.X | RuleState.ST_ERROR | RuleState.ST_EXIT_LEVEL | RuleState.ST_KEEP_OOF_HIST | RuleState.ST_RULE_PARSE_ERROR | RuleState.ST_SKIP_IF_SCL_IS_SAFE;
            RuleData ruleForMarkRead2 = AdapterHelper.GenerateValidRuleData(ActionType.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 1, ruleState, new DeleteMarkReadActionData(), ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForMarkRead2 });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Mark read rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser2 delivers a message to trigger these rules.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title", 1);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message content to verify the rule evaluation.
            PropertyTag[] propertyTagList = new PropertyTag[2];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagMessageFlags;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypInteger32;

            uint contentTableHandle = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);
            int messageFlags = BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[1].Value, 0);
            string subject = AdapterHelper.PropertyValueConvertToString(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);

            #region Capture Code
            Site.Assert.AreNotEqual<RuleState>(RuleState.ST_ENABLED, ruleState & RuleState.ST_ENABLED, "ST_ENABLED flag is not set in the PidTagRuleState property");
            Site.Assert.AreNotEqual<RuleState>(RuleState.ST_ONLY_WHEN_OOF, ruleState & RuleState.ST_ONLY_WHEN_OOF, "ST_ONLY_WHEN_OOF flag is not set in the PidTagRuleState property");

            // Add the debug information.
            Site.Log.Add(
               LogEntryKind.Debug, "Verify MS-OXORULE_R516, the pidTagSubject of the incoming message is {0}, the pidTagMessageFlags of the incoming message is {1}, the ruleState is {2}", mailSubject, messageFlags, ruleState);

            // Verify MS-OXORULE requirement: MS-OXORULE_R516.
            // Add an OP_MARK_AS_READ rule with the ST_ENABLED flag not set in the PidTagRuleState property.
            // mailSubject and messageFlags represent the subject name and message flag of the incoming message.
            // If mailSubject equals ruleConditionSubjectName, it means the incoming message satisfies the rule condition and can trigger the OP_MARK_AS_READ rule.
            // 0x00000001 is the flag which represents the message has been read. If messageFlags doesn't set this flag, it means the incoming message 
            // isn't marked as read, which indicates the server doesn't evaluate the rule that is not enabled.
            bool isVerifyR516 = (subject == mailSubject) &&
                                    ((messageFlags & 0x00000001) != 0x00000001);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR516,
                516,
                @"[In Processing Incoming Messages to a Folder] The server MUST only evaluate rules (2) that are enabled; that is, rules (2) that have the ST_ENABLED flag set in the PidTagRuleState property (2.2.1.3.1.3).");

            // Add the debug information.
            Site.Log.Add(
               LogEntryKind.Debug, "Verify MS-OXORULE_R63, the pidTagSubject of the incoming message is {0}, the pidTagMessageFlags of the incoming message is {1}, the ruleState is {2}", mailSubject, messageFlags, ruleState);

            // Verify MS-OXORULE requirement: MS-OXORULE_R63.
            // Add an OP_MARK_AS_READ rule without the ST_ENABLED flag and ST_ONLY_WHEN_OOF flag set in the PidTagRuleState property.
            // mailSubject and messageFlags represent the subject name and message flag of the incoming message.
            // If mailSubject equals ruleConditionSubjectName, it means the incoming message satisfies the rule condition and can trigger the OP_MARK_AS_READ rule.
            // 0x00000001 is the flag which represents the message has been read. If messageFlags doesn't set this flag, it means the incoming message 
            // isn't marked as read, which indicates the server doesn't evaluate the rule that is not enabled.
            bool isVerifyR63 = (subject == mailSubject) &&
                                    ((messageFlags & 0x00000001) != 0x00000001);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR63,
                63,
                @"[In PidTagRuleState] EN (ST_ENABLED, Bitmask 0x00000001): If neither this flag nor the ST_ONLY_WHEN_OOF flag are set, the server skips this rule (2) when evaluating rules (2).");
            #endregion

            #endregion

            #region TestUser1 adds a rule for ActionType OP_MARK_AS_READ.
            // Clean all the contents in the Inbox folder of TestUser1.
            this.OxoruleAdapter.RopEmptyFolder(this.InboxFolderHandle, 0);

            // Wait for the empty folder operation to take action.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            RuleData ruleForMarkRead = AdapterHelper.GenerateValidRuleData(ActionType.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, new DeleteMarkReadActionData(), ruleProperties, null);
            ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForMarkRead });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding mark as read rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser2 delivers a message to trigger these rules.
            mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title", 2);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message content to verify the rule evaluation.
            getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);
            int messageFlag = BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[1].Value, 0);
            subject = AdapterHelper.PropertyValueConvertToString(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);

            #region Capture Code
            if (Common.IsRequirementEnabled(914, this.Site))
            {
                // Add an OP_MARK_AS_READ rule. mailSubject represents the subject name of the delivered message.
                // If mailSubject equals ruleConditionSubjectName, it means the delivered message satisfies the rule condition 
                // and can trigger the OP_MARK_AS_READ rule. 0x00000001 is the flag which represents the message has been read.
                // If messageFlags has set this flag, it means the delivered message is marked as read according to the rule,
                // which also indicates the server starts using the newly added rule when processing the delivered message.
                bool isR914Satisfied = subject == mailSubject &&
                                       (messageFlag & 0x00000001) == 0x00000001;

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R914, the pidTagSubject of the delivered message is {0}, the pidTagMessageFlags of the delivered message is {1}", mailSubject, messageFlag);

                // Verify MS-OXORULE requirement: MS-OXORULE_R914.
                Site.CaptureRequirementIfIsTrue(
                    isR914Satisfied,
                    914,
                    @"[In Receiving a RopModifyRules ROP Request] The implementation does start using the newly modified rules (2) when processing messages delivered to that folder as soon as it successfully processes the RopModifyRules ROP request. (Exchange 2003 and above follow this behavior.)");
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R551, the value of PidTagMessageFlags is {0}", messageFlag);

            // Verify MS-OXORULE requirement: MS-OXORULE_R551.
            // 0x00000001 is the flag which represents the message has been read.
            // If messageFlags has set this flag, it means the delivered message is marked as read according to the rule. Otherwise, it means the delivered message is not marked as read.
            bool isVerifyR551 = (messageFlag & 0x00000001) == 0x00000001;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR551,
                551,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_MARK_AS_READ"": the server MUST set the MSGFLAG_READ flag (0x00000001) in the PidTagMessageFlags property ([MS-OXPROPS] section 2.782) on the message.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R902, the value of PidTagMessageFlags is {0}", messageFlag);

            // Verify MS-OXORULE requirement: MS-OXORULE_R902.
            // 0x00000001 is the flag which represents the message has been read.
            // If messageFlags has set this flag, it means the delivered message is marked as read according to the rule. Otherwise, it means the delivered message is not marked as read.
            bool isVerifyR902 = (messageFlag & 0x00000001) == 0x00000001;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR902,
                902,
                @"[In OP_DELETE or OP_MARK_AS_READ ActionData Structure] For the OP_MARK_AS_READ action type, the incoming messages are marked as read according to the ActionType itself.");

            #endregion
            #endregion

            #region TestUser2 delivers a message which does not satisfy rule condition.
            // Clean all the contents in the Inbox folder of TestUser1.
            this.OxoruleAdapter.RopEmptyFolder(this.InboxFolderHandle, 0);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser2 delivers a message to trigger these rules.
            string notSatisfyMessageSubject = Common.GenerateResourceName(this.Site, Constants.ExtendRulename2 + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, notSatisfyMessageSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message content to verify the rule evaluation.
            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            expectedMessageIndex = 0;
            getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandle, propertyTagList, ref expectedMessageIndex, notSatisfyMessageSubject);
            messageFlag = BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[1].Value, 0);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R559, the value of PidTagMessageFlags is {0}", messageFlag);

            // Verify MS-OXORULE requirement: MS-OXORULE_R559.
            // mailSubject represents the subject name of the delivered message.
            // If mailSubject does not equals ruleConditionSubjectName, it means the delivered message does not satisfies the rule condition and cannot trigger the OP_MARK_AS_READ rule.
            // 0x00000001 is the flag which represents the message has been read.
            // If messageFlags has set this flag, it means the delivered message is marked as read according to the rule. Otherwise, it means the delivered message is not marked as read.
            bool isVerifyR559 = (messageFlag & 0x00000001) != 0x00000001 && !notSatisfyMessageSubject.Contains(ruleProperties.ConditionSubjectName);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR559,
                559,
                @"[In Processing Out of Office Rules] If the rule (2) condition evaluates to ""false"", no additional action (2) needs to be taken.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_FORWARD rule.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC07_ServerExecuteRule_Action_OP_FORWARD()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameForwardTM);
            #endregion

            #region TestUser1 adds a rule for ActionType with OP_Forward ActionFlavor set to TM and rule sequence set to 0.
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
            RuleData ruleForwardTM = AdapterHelper.GenerateValidRuleDataWithFlavor(ActionType.OP_FORWARD, 0, RuleState.ST_ENABLED, forwardActionData, (uint)ActionFlavorsForward.TM, ruleProperties);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForwardTM });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Forward rule should succeed.");
            #endregion

            #region TestUser2 adds a rule for ActionType OP_Forward with rule sequence set to 0.
            // Let TestUser2 log on to the server.
            this.LogonMailbox(TestUser.TestUser2);
            forwardActionData = new ForwardDelegateActionData
            {
                RecipientCount = (ushort)0x01
            };
            recipientBlock = new RecipientBlock
            {
                Reserved = 0x01,
                NoOfProperties = (ushort)0x04u
            };

            #region Prepare the recipient Block of the rule to forward the message to TestUser1.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameForward);
            recipientProperties = AdapterHelper.GenerateRecipientPropertiesBlock(this.User1Name, this.User1ESSDN);

            recipientBlock.PropertiesData = recipientProperties;
            #endregion

            forwardActionData.RecipientsData = new RecipientBlock[1] { recipientBlock };
            RuleData ruleForward = AdapterHelper.GenerateValidRuleData(ActionType.OP_FORWARD, TestRuleDataType.ForAdd, 0, RuleState.ST_ENABLED, forwardActionData, ruleProperties, null);
            ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForward });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Forward rule should succeed.");
            #endregion

            #region TestUser1 delivers a message to itself to trigger these rules.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser1 delivers a message to itself to trigger these rules.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title", 1);
            this.SUTAdapter.SendMailToRecipient(this.User1Name, this.User1Password, this.User1Name, mailSubject);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser2 gets the forwarded message to verify the rule evaluation.
            PropertyTag[] propertyTagList = new PropertyTag[2];

            // pidTagSubject and pidTagMessageClass
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagMessageClass;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;
            uint contentsTableHandle = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse testUser2getNormalMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);
            Site.Assert.AreEqual<uint>(0, testUser2getNormalMailMessageContent.ReturnValue, "Getting message property operation should succeed.");

            string subject = AdapterHelper.PropertyValueConvertToString(testUser2getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);
            byte[] propertyValue = testUser2getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[1].Value;

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R537");

            // Verify MS-OXORULE requirement: MS-OXORULE_R537.
            // The subject name of the forwarded message should contain the original received message's subject name. 
            bool isVerifiedR537 = subject.ToUpperInvariant().Contains(mailSubject.ToUpperInvariant());

            // If there exists a message under the recipient's Inbox folder, whose subject name contains the original received message's subject name,
            // it means the server has forwarded the message to the corresponding recipient.
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR537,
                537,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_FORWARD"": The server MUST forward the message to the recipients (2) specified in the action buffer structure (except for messages forwarded to the sender).");

            if (Common.IsRequirementEnabled(802, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R802: the value of PidTagSubject is {0}, and the value of PidTagMessageClass is {1}", mailSubject, AdapterHelper.PropertyValueConvertToString(propertyValue));

                // If there exists a message under the recipient's Inbox folder, whose subject name contains the original received message's subject name,
                // it means the server has forwarded the message to the corresponding recipient, and for SMS text messages, the value of the PidTagMessageClass property is set to "IPM.Note.Mobil.SMS.Alert". 
                bool isVerify802 = subject.ToUpperInvariant().Contains(mailSubject.ToUpperInvariant()) && !AdapterHelper.PropertyValueConvertToString(propertyValue).Equals("IPM.Note.Mobil.SMS.Alert");

                // Verify MS-OXORULE requirement: MS-OXORULE_R802.
                Site.CaptureRequirementIfIsTrue(
                    isVerify802,
                    802,
                    @"[In Appendix A: Product Behavior] Implementation does not support forwarding messages as SMS text messages. [<5> Section 2.2.5.1.1: Exchange 2003 and Exchange 2007 do not support forwarding messages as SMS text messages.]");
            }

            if (Common.IsRequirementEnabled(897, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R897: the PidTagMessageClass of the message is {0}", AdapterHelper.PropertyValueConvertToString(propertyValue));

                // If there exists a message under the recipient's Inbox folder, whose subject name contains the original received message's subject name,
                // it means the server has forwarded the message to the corresponding recipient, and for SMS text messages, the value of the PidTagMessageClass property is set to "IPM.Note.Mobil.SMS.Alert". 
                bool isVerify897 = subject.ToUpperInvariant().Contains(mailSubject.ToUpperInvariant()) && AdapterHelper.PropertyValueConvertToString(propertyValue).Equals("IPM.Note.Mobile.SMS.Alert");

                // Verify MS-OXORULE requirement: MS-OXORULE_R897.
                Site.CaptureRequirementIfIsTrue(
                    isVerify897,
                    897,
                    @"[In Action Flavors] TM (Bitmask 0x00000008): Implementation does forward the message as a Short Message Service (SMS) text message. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion
            #endregion

            #region TestUser1 get the forwarded message to verify the rule evaluation.
            // Let TestUser1 log on to the server.
            this.LogonMailbox(TestUser.TestUser1);

            #region Capture Code
            if (Common.IsRequirementEnabled(799, this.Site))
            {
                // The subject name of the forwarded message should contain the original received message's subject name. 
                uint countOfExepctedMessage = 0;
                RopQueryRowsResponse testUser1getNormalMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref countOfExepctedMessage, 2, mailSubject);
                Site.Assert.AreEqual<uint>(0, testUser1getNormalMailMessageContent.ReturnValue, "Getting message property operation should succeed.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R799");

                // Verify MS-OXORULE requirement: MS-OXORULE_R799.
                // If test user2 has forwarded the message that has been forwarded to the test user1, there should be two expected messages in the test user1's Inbox folder.
                Site.CaptureRequirementIfAreEqual<uint>(
                    2,
                    countOfExepctedMessage,
                    799,
                    @"[In Appendix A: Product Behavior] Implementation does forward messages that have been forwarded to the sender. [<17> Section 3.2.5.1: Exchange 2007 forwards messages that have been forwarded to the sender.]");
            }

            if (Common.IsRequirementEnabled(907, this.Site))
            {
                // The subject name of the forwarded message should contain the original received message's subject name. 
                uint countOfExepctedMessage = 0;
                RopQueryRowsResponse testUser1getNormalMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref countOfExepctedMessage, 1, mailSubject);
                Site.Assert.AreEqual<uint>(0, testUser1getNormalMailMessageContent.ReturnValue, "Getting message property operation should succeed.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R907");

                // Verify MS-OXORULE requirement: MS-OXORULE_R907
                // If test user2 has not forwarded the message that has been forwarded to the test user1, there should be only one expected message in the test user1's Inbox folder.
                Site.CaptureRequirementIfAreEqual<uint>(
                    1,
                    countOfExepctedMessage,
                    907,
                    @"[In Processing Incoming Messages to a Folder] Implementation does not forward messages that were forwarded to the sender. (Exchange 2003, Exchange 2010 and above follow this behavior.)");
            }
            #endregion
            #endregion

            #region TestUser1 adds a rule for ActionType OP_Forward with OP_Forward ActionFlavor set to PR.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameForward);
            forwardActionData = new ForwardDelegateActionData
            {
                RecipientCount = (ushort)0x01
            };
            recipientBlock = new RecipientBlock
            {
                Reserved = 0x01,
                NoOfProperties = (ushort)0x04u
            };

            #region Prepare the recipient Block of the rule to forward the message to TestUser2.
            recipientProperties = AdapterHelper.GenerateRecipientPropertiesBlock(this.User2Name, this.User2ESSDN);

            recipientBlock.PropertiesData = recipientProperties;
            #endregion

            forwardActionData.RecipientsData = new RecipientBlock[1] { recipientBlock };
            RuleData ruleForwardPR = AdapterHelper.GenerateValidRuleDataWithFlavor(ActionType.OP_FORWARD, 0, RuleState.ST_ENABLED, forwardActionData, (uint)ActionFlavorsForward.PR, ruleProperties);
            ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForwardPR });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Forward rule should succeed.");
            #endregion

            #region TestUser2 deletes the previous rule and clean the Inbox folder.
            // Let TestUser2 log on to the server.
            this.LogonMailbox(TestUser.TestUser2);

            // Call RopGetRulesTable with valid TableFlags.
            this.ClearAllRules();

            // Clean all the contents in the Inbox folder of TestUser2.
            this.OxoruleAdapter.RopEmptyFolder(this.InboxFolderHandle, 0);
            #endregion

            #region TestUser1 delivers a message to itself to trigger these rules.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser1 to deliver a message to itself to trigger these rules
            mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title", 2);
            this.SUTAdapter.SendMailToRecipient(this.User1Name, this.User1Password, this.User1Name, mailSubject);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser2 gets the forwarded message to verify the rule evaluation.
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagAutoForwarded;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypBoolean;
            expectedMessageIndex = 0;
            testUser2getNormalMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);
            Site.Assert.AreEqual<uint>(0, testUser2getNormalMailMessageContent.ReturnValue, "Getting message property operation should succeed.");
            bool isAutoForwarded = AdapterHelper.PropertyValueConvertToBool(testUser2getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[1].Value);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R275: The expected message with subject {0} in TestUser2 inbox has the auto forwarded property set to {1}", mailSubject, isAutoForwarded);

            // If there exists a message under the recipient's Inbox folder
            // it means the server has forwarded the message to the corresponding recipient, and if the value of the PidTagAutoForwarded property is set to "true", means it is autoforwards.
            bool isVerify275 = isAutoForwarded;

            // Verify MS-OXORULE requirement: MS-OXORULE_R275
            Site.CaptureRequirementIfIsTrue(
                isVerify275,
                275,
                @"[In Action Flavors] PR (Bitmask 0x00000001): Preserves the sender information and indicates that the message was auto forwarded.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_DELEGATE rule. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC08_ServerExecuteRule_Action_OP_DELEGATE()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameDelegate);
            #endregion

            #region TestUser1 adds an OP_DELEGATE rule.
            ForwardDelegateActionData delegateActionData = new ForwardDelegateActionData
            {
                RecipientCount = (ushort)0x01
            };
            RecipientBlock recipientBlock = new RecipientBlock
            {
                Reserved = 0x01,
                NoOfProperties = (ushort)0x05u
            };

            #region Prepare recipient Block.
            TaggedPropertyValue[] recipientProperties = new TaggedPropertyValue[5];

            TaggedPropertyValue[] temp = AdapterHelper.GenerateRecipientPropertiesBlock(this.User2Name, this.User2ESSDN);
            Array.Copy(temp, 0, recipientProperties, 0, temp.Length);

            // Add PidTagSmtpEmailAdderss.
            recipientProperties[4] = new TaggedPropertyValue();
            PropertyTag pidTagSmtpEmailAdderssPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagSmtpAddress,
                PropertyType = (ushort)PropertyType.PtypString
            };
            recipientProperties[4].PropertyTag = pidTagSmtpEmailAdderssPropertyTag;
            recipientProperties[4].Value = Encoding.Unicode.GetBytes(this.User2Name + "@" + this.Domain + "\0");

            recipientBlock.PropertiesData = recipientProperties;
            #endregion

            delegateActionData.RecipientsData = new RecipientBlock[1] { recipientBlock };
            RuleData ruleDelegate = AdapterHelper.GenerateValidRuleData(ActionType.OP_DELEGATE, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, delegateActionData, ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleDelegate });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding delegate rule should succeed.");
            #endregion

            #region TestUser1 delivers a message to itself to trigger the rule.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser1 delivers a message to itself to trigger the rule.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User1Name, this.User1Password, this.User1Name, mailSubject);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser2 gets the delegate message to verify the rule evaluation.
            // Let TestUser2 log on to the server.
            this.LogonMailbox(TestUser.TestUser2);

            PropertyTag[] propertyTagList = new PropertyTag[7];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagReceivedRepresentingEntryId;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypBinary;
            propertyTagList[2].PropertyId = (ushort)PropertyId.PidTagReceivedRepresentingAddressType;
            propertyTagList[2].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[3].PropertyId = (ushort)PropertyId.PidTagReceivedRepresentingEmailAddress;
            propertyTagList[3].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[4].PropertyId = (ushort)PropertyId.PidTagReceivedRepresentingName;
            propertyTagList[4].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[5].PropertyId = (ushort)PropertyId.PidTagReceivedRepresentingSearchKey;
            propertyTagList[5].PropertyType = (ushort)PropertyType.PtypBinary;
            propertyTagList[6].PropertyId = (ushort)PropertyId.PidTagDelegatedByRule;
            propertyTagList[6].PropertyType = (ushort)PropertyType.PtypBoolean;

            uint contentTableHandler = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse getNormalMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandler, propertyTagList, ref expectedMessageIndex, mailSubject);
            #endregion

            #region Get TestUser1's information from address book

            // Let TestUser1 log on to the server.
            this.LogonMailbox(TestUser.TestUser1);

            PropertyTagArray_r ptags = new PropertyTagArray_r
            {
                Values = 5,
                AulPropTag = AdapterHelper.SerializeRecipientProperties()
            };

            // The Windows NSPI will be invoked when the first parameter is domain name instead of server address.
            PropertyRowSet_r? propertyRows = this.OxoruleAdapter.GetRecipientInfo(this.Domain, this.User1Name, this.Domain, this.User1Password, ptags);
            Site.Assert.IsNotNull(propertyRows, "The recipient information returned by the NSPI service should not be null");
            int user1Index = 0;
            for (int i = 0; i < propertyRows.Value.Rows; i++)
            {
                if (Encoding.Unicode.GetString(propertyRows.Value.PropertyRowSet[i].Props[3].Value.LpszW).ToLower(System.Globalization.CultureInfo.CurrentCulture) == this.User1Name.ToLower(System.Globalization.CultureInfo.CurrentCulture))
                {
                    user1Index = i;
                    break;
                }
            }

            // The two EntryId should be the same.
            AddressBookEntryID addressbookEntryId = new AddressBookEntryID();
            addressbookEntryId.Deserialize(propertyRows.Value.PropertyRowSet[user1Index].Props[0].Value.Bin.Lpb);
            byte[] pidTagReceivedRepresentingEntryIdbytesTemp = getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[1].Value;
            byte[] pidTagReceivedRepresentingEntryIdbytes = new byte[pidTagReceivedRepresentingEntryIdbytesTemp.Length - 2];
            Array.Copy(pidTagReceivedRepresentingEntryIdbytesTemp, 2, pidTagReceivedRepresentingEntryIdbytes, 0, pidTagReceivedRepresentingEntryIdbytes.Length);
            AddressBookEntryID mailEntryID = new AddressBookEntryID();
            mailEntryID.Deserialize(pidTagReceivedRepresentingEntryIdbytes);
            string subject = AdapterHelper.PropertyValueConvertToString(getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R539");

            // Verify MS-OXORULE requirement: MS-OXORULE_R539.
            Site.CaptureRequirementIfAreEqual<string>(
                mailSubject,
                subject,
                539,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_DELEGATE"": the server MUST resend the message to the recipients (2) specified in the action buffer structure.");

            string pidTagEntryIdOfMailboxUser = addressbookEntryId.ValueOfX500DN.ToLower(System.Globalization.CultureInfo.CurrentCulture);
            string pidTagReceivedRepresentingEntryId = mailEntryID.ValueOfX500DN.ToLower(System.Globalization.CultureInfo.CurrentCulture);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R540");

            // Verify MS-OXORULE requirement: MS-OXORULE_R540.
            Site.CaptureRequirementIfAreEqual<string>(
                pidTagReceivedRepresentingEntryId,
                pidTagEntryIdOfMailboxUser,
                540,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_DELEGATE"": The server also MUST set the values of the following properties to match the current user's properties in the address book: The PidTagReceivedRepresentingEntryId property ([MS-OXOMSG] section 2.2.1.25) MUST be set to the same value as the mailbox user's PidTagEntryId property ([MS-OXOABK] section 2.2.3.3).");

            string pidTagReceivedRepresentingAddressType = Encoding.Unicode.GetString(getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[2].Value);

            // The actual value of pidTagReceivedRepresentingAddressType should not contain the last '\0' character.
            pidTagReceivedRepresentingAddressType = pidTagReceivedRepresentingAddressType.Substring(0, pidTagReceivedRepresentingAddressType.Length - 1);

            // In this test case, the mailbox user's PidTagAddressType is "EX".
            string pidTagAddressTypeOfMailboxUser = System.Text.UTF8Encoding.Unicode.GetString(propertyRows.Value.PropertyRowSet[user1Index].Props[1].Value.LpszW);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R541");

            // Verify MS-OXORULE requirement: MS-OXORULE_R541.
            Site.CaptureRequirementIfAreEqual<string>(
                pidTagReceivedRepresentingAddressType,
                pidTagAddressTypeOfMailboxUser,
                541,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_DELEGATE"": The PidTagReceivedRepresentingAddressType property ([MS-OXOMSG] section 2.2.1.23) MUST be set to the same value as the mailbox user's PidTagAddressType property ([MS-OXOABK] section 2.2.3.13).");

            string pidTagReceivedRepresentingEmailAddress = Encoding.Unicode.GetString(getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[3].Value);

            // The actual value of PidTagReceivedRepresentingEmailAddress should not contain the last '\0' character.
            pidTagReceivedRepresentingEmailAddress = pidTagReceivedRepresentingEmailAddress.Substring(0, pidTagReceivedRepresentingEmailAddress.Length - 1).ToUpperInvariant();

            // In this test case, the mailbox user's PidTagEmailAddress is the adminUserDN.
            string pidTagEmailAddressOfMailboxUser = Encoding.Unicode.GetString(propertyRows.Value.PropertyRowSet[user1Index].Props[2].Value.LpszW).ToUpperInvariant();

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R542");

            // Verify MS-OXORULE requirement: MS-OXORULE_R542.
            Site.CaptureRequirementIfAreEqual<string>(
                pidTagReceivedRepresentingEmailAddress,
                pidTagEmailAddressOfMailboxUser,
                542,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_DELEGATE"": The PidTagReceivedRepresentingEmailAddress property ([MS-OXOMSG] section 2.2.1.24) MUST be set to the same value as the mailbox user's PidTagEmailAddress property ([MS-OXOABK] section 2.2.3.14).");

            string pidTagReceivedRepresentingName = AdapterHelper.PropertyValueConvertToString(getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[4].Value).ToLower(System.Globalization.CultureInfo.CurrentCulture);

            // In this test case, the mailbox user's PidTagDisplayName is "administrator".
            string pidTagDisplayNameOfMailboxUser = Encoding.Unicode.GetString(propertyRows.Value.PropertyRowSet[user1Index].Props[3].Value.LpszW).ToLower(System.Globalization.CultureInfo.CurrentCulture);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R543");

            // Verify MS-OXORULE requirement: MS-OXORULE_R543.
            Site.CaptureRequirementIfAreEqual<string>(
                pidTagReceivedRepresentingName,
                pidTagDisplayNameOfMailboxUser,
                543,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_DELEGATE"": The PidTagReceivedRepresentingName property ([MS-OXOMSG] section 2.2.1.26) MUST be set to the same value as the mailbox user's PidTagDisplayName property ([MS-OXCFOLD] section 2.2.2.2.2.5).");

            byte[] pidTagReceivedRepresentingSearchKeyOfbytes = getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[5].Value;
            byte[] pidTagReceivedRepresentingSearchKey = AdapterHelper.PropertyValueConvertToBinary(pidTagReceivedRepresentingSearchKeyOfbytes);
            byte[] pidTagSearchKeyOfMailboxUser = propertyRows.Value.PropertyRowSet[user1Index].Props[4].Value.Bin.Lpb;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R544: the value of PidTagReceivedRepresentingSearchKey is {0}", pidTagSearchKeyOfMailboxUser);

            // Verify MS-OXORULE requirement: MS-OXORULE_R544.
            bool isVerifyR544 = Common.CompareByteArray(pidTagSearchKeyOfMailboxUser, pidTagReceivedRepresentingSearchKey);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR544,
                544,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_DELEGATE"": The PidTagReceivedRepresentingSearchKey property ([MS-OXOMSG] section 2.2.1.27) MUST be set to the same value as the mailbox user's PidTagSearchKey property ([MS-OXCPRPT] section 2.2.1.9).");

            // BitConverter.ToBoolean() is used to convert a byte array to a bool value from the byte array index of 0.
            bool pidTagDelegatedByRule = BitConverter.ToBoolean(getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[6].Value, 0);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R545");

            // Verify MS-OXORULE requirement: MS-OXORULE_R545.
            Site.CaptureRequirementIfIsTrue(
                pidTagDelegatedByRule,
                545,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_DELEGATE"": The PidTagDelegatedByRule property ([MS-OXOMSG] section 2.2.1.84) MUST be set to ""TRUE"".");

            #endregion
            #endregion

            #region TestUser1 calls RopGetRulesTable with valid TableFlags.

            RopGetRulesTableResponse ropGetRulesTableResponse;
            uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");
            #endregion

            #region TestUser1 calls RopQueryRows to retrieve rows from the rule table

            PropertyTag[] propertyTags = new PropertyTag[2];
            propertyTags[0].PropertyId = (ushort)PropertyId.PidTagRuleName;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTags[1].PropertyId = (ushort)PropertyId.PidTagRuleActions;
            propertyTags[1].PropertyType = (ushort)PropertyType.PtypRuleAction;

            // Retrieves rows from the rule table.
            RopQueryRowsResponse queryRowResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyTags);
            Site.Assert.AreEqual<uint>(0, queryRowResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");
            ForwardDelegateActionData forwardDelegateActionDataOfQueryRowResponse = new ForwardDelegateActionData();
            RuleAction ruleAction = new RuleAction();
            for (int i = 0; i < queryRowResponse.RowCount; i++)
            {
                System.Text.UnicodeEncoding converter = new UnicodeEncoding();
                string ruleName = converter.GetString(queryRowResponse.RowData.PropertyRows.ToArray()[i].PropertyValues[0].Value);
                if (ruleName == ruleProperties.Name + "\0")
                {
                    // Verify structure RuleAction 
                    ruleAction.Deserialize(queryRowResponse.RowData.PropertyRows[i].PropertyValues[1].Value);
                    forwardDelegateActionDataOfQueryRowResponse.Deserialize(ruleAction.Actions[0].ActionDataValue.Serialize());
                    break;
                }
            }

            #region Capture Code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R1003");

            // Verify MS-OXORULE requirement: MS-OXORULE_R1003
            this.Site.CaptureRequirementIfIsInstanceOfType(
                forwardDelegateActionDataOfQueryRowResponse.RecipientsData,
                typeof(RecipientBlock[]),
                1003,
                @"[In OP_FORWARD and OP_DELEGATE ActionData Structure] RecipientBlocks (variable): An array of RecipientBlockData structures, each of which specifies information about one recipient (2).");

            for (int i = 0; i < forwardDelegateActionDataOfQueryRowResponse.RecipientsData.Length; i++)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R1006");

                // Verify MS-OXORULE requirement: MS-OXORULE_R1006
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    forwardDelegateActionDataOfQueryRowResponse.RecipientsData[i].PropertiesData,
                    typeof(TaggedPropertyValue[]),
                    1006,
                    @"[In RecipientBlockData Structure] PropertyValues (variable): An array of TaggedPropertyValue structures, each of which contains a property that provides some information about the recipient (2).");
            }
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to test the operation of TestUser1 for adding an OP_TAG standard rule to the server.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC09_ServerExecuteRule_Action_OP_TAG()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameTag);
            #endregion

            #region TestUser1 adds an OP_TAG rule to the Inbox folder.
            TagActionData tagActionData = new TagActionData();
            PropertyTag tagActionDataPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagImportance,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            tagActionData.PropertyTag = tagActionDataPropertyTag;
            tagActionData.PropertyValue = BitConverter.GetBytes(2);

            RuleData ruleOpTag = AdapterHelper.GenerateValidRuleData(ActionType.OP_TAG, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, tagActionData, ruleProperties, null);
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleOpTag });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding OP_TAG rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1.

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");

            // TestUser2 delivers a message to trigger these rules.
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message and its properties to verify the rule evaluation.
            PropertyTag[] propertyTagList = new PropertyTag[2];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagImportance;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypInteger32;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;

            uint contentTableHandle = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R548: the value of PidTagImportance is {0}", BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0));

            // Verify MS-OXORULE requirement: MS-OXORULE_R548.
            // If the PidTagImportance is 2 which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property specified in the rule's action buffer structure.
            bool isVerifyR548 = BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0) == 2;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR548,
                548,
                @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_TAG"": The server MUST set on the message the property specified in the action buffer structure.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R266: the value of PidTagImportance is {0}", BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0));

            // Verify MS-OXORULE requirement: MS-OXORULE_R266.
            // If the PidTagImportance is 2 which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property specified in the rule's action buffer structure.
            bool isVerifyR266 = BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0) == 2;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR266,
                266,
                @"[In ActionBlock Structure] The meaning of action type OP_TAG: Adds or changes a property on the message.");
            #endregion

            #region TestUser1 gets the rule action property to check the rule action property.
            uint ruleHandle;
            RopGetRulesTableResponse ruleTableResponse = new RopGetRulesTableResponse();
            ruleHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ruleTableResponse);

            // Get rule property.
            PropertyTag[] ruleProperty = new PropertyTag[1];
            ruleProperty[0].PropertyId = (ushort)PropertyId.PidTagRuleActions;
            ruleProperty[0].PropertyType = (ushort)PropertyType.PtypRuleAction;
            RopQueryRowsResponse queryRowResponseOfProperty = this.OxoruleAdapter.QueryPropertiesInTable(ruleHandle, ruleProperty);
            Site.Assert.AreEqual<uint>(0, queryRowResponseOfProperty.ReturnValue, "Getting the rule action property should succeed, the actual value is {0}!", queryRowResponseOfProperty.ReturnValue);
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_MOVE rule.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC10_ServerExecuteRule_Action_OP_MOVE()
        {
            this.CheckMAPIHTTPTransportSupported();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(929, this.Site), "This case runs only when the server supports OP_MOVE action when FolderInThisStore is set to 0.");

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameMoveOne);
            #endregion

            #region TestUser1 creates folder1 in server store.
            RopCreateFolderResponse createFolderResponse;
            uint newFolderHandle = this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, Common.GenerateResourceName(this.Site, "User1Folder01"), "TestForOP_MOVE", out createFolderResponse);
            ulong newFolderId = createFolderResponse.FolderId;
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating folder operation should succeed.");
            #endregion

            #region Prepare rules' data
            MoveCopyActionData moveCopyActionData1 = new MoveCopyActionData();

            // Get the created folder1 entry id.
            byte[] folder1EId = this.OxoruleAdapter.GetFolderEntryId(StoreObjectType.Mailbox, newFolderHandle, newFolderId);

            // Get the store object's entry id
            byte[] storeEId = this.GetStoreObjectEntryID(StoreObjectType.Mailbox, this.Server, this.User1ESSDN);
            moveCopyActionData1.FolderInThisStore = 0;
            moveCopyActionData1.FolderEID = folder1EId;
            moveCopyActionData1.StoreEID = storeEId;
            moveCopyActionData1.FolderEIDSize = (ushort)folder1EId.Length;
            moveCopyActionData1.StoreEIDSize = (ushort)storeEId.Length;

            IActionData[] moveCopyActionData = { moveCopyActionData1 };
            #endregion

            #region Generate test RuleData.
            // Add rule for move without rule Provider Data.
            ruleProperties.ProviderData = string.Empty;
            RuleData ruleForMoveFolder = AdapterHelper.GenerateValidRuleDataWithFlavor(new ActionType[] { ActionType.OP_MOVE }, 0, RuleState.ST_ENABLED, moveCopyActionData, new uint[] { 0 }, ruleProperties);

            #endregion

            #region TestUser1 adds OP_MOVE rule to the Inbox folder.
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForMoveFolder });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Move rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.

            // TestUser2 deliver a message to trigger these rules.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message content to verify the rule evaluation.
            uint inboxFolderContentsTableHandle = 0;
            PropertyTag[] propertyTagList = new PropertyTag[1];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;

            bool doesOriginalMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref inboxFolderContentsTableHandle, propertyTagList, mailSubject);
            uint newFolder1ContentsTableHandle = 0;
            uint rowCount = 0;
            RopQueryRowsResponse getNewFolder1MailMessageContent = this.GetExpectedMessage(newFolderHandle, ref newFolder1ContentsTableHandle, propertyTagList, ref rowCount, 1, mailSubject);
            Site.Assert.AreEqual<uint>(0, getNewFolder1MailMessageContent.ReturnValue, "Getting message on the folder should succeed.");

            this.VerifyActionTypeOP_MOVE(getNewFolder1MailMessageContent, doesOriginalMessageExist);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R929");

            // Verify MS-OXORULE requirement: MS-OXORULE_R929.
            // When the server moves the message to the destination store, it means StoreEID is set to the destination store EntryID.
            bool isVerifyR929 = getNewFolder1MailMessageContent.RowCount == 1 && !doesOriginalMessageExist;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR929,
                929,
                @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Standard Rules] StoreEID (variable):  A Store Object EntryID structure, as specified in [MS-OXCDATA] section 2.2.4.3, [In OP_MOVE action data] Identifies the message store.");
            #endregion

            #region Delete the folders created in step2 and step3.
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, newFolderId);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_COPY rule.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC11_ServerExecuteRule_Action_OP_COPY()
        {
            this.CheckMAPIHTTPTransportSupported();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(298, this.Site), "This case runs only when the server supports OP_COPY action when FolderInThisStore is set to 0.");

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameCopy);
            #endregion

            #region TestUser1 creates folder1 in server store.
            RopCreateFolderResponse createFolderResponse;
            uint newFolderHandle = this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, Common.GenerateResourceName(this.Site, "User1Folder"), "TestForOP_COPY", out createFolderResponse);
            ulong newFolderId = createFolderResponse.FolderId;
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating folder operation should succeed.");
            #endregion

            #region Prepare rules' data.
            MoveCopyActionData moveCopyActionData = new MoveCopyActionData();

            // Get the created folder entry ID.
            byte[] folderEId = this.OxoruleAdapter.GetFolderEntryId(StoreObjectType.Mailbox, newFolderHandle, newFolderId);

            // Get the store object's entry ID.
            byte[] storeEId = this.GetStoreObjectEntryID(StoreObjectType.Mailbox, this.Server, this.User1ESSDN);
            moveCopyActionData.FolderInThisStore = 0;
            moveCopyActionData.FolderEID = folderEId;
            moveCopyActionData.StoreEID = storeEId;
            moveCopyActionData.FolderEIDSize = (ushort)folderEId.Length;
            moveCopyActionData.StoreEIDSize = (ushort)storeEId.Length;
            #endregion

            #region Generate test RuleData.
            // Add rule for OP_COPY without rule Provider Data.
            ruleProperties.ProviderData = string.Empty;
            RuleData ruleForCopy = AdapterHelper.GenerateValidRuleData(ActionType.OP_COPY, TestRuleDataType.ForAdd, 0, RuleState.ST_ENABLED, moveCopyActionData, ruleProperties, null);
            #endregion

            #region TestUser1 adds OP_COPY rule to the Inbox folder.
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForCopy });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Copy rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.

            // TestUser2 delivers a message to trigger these rules.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message content to verify the rule evaluation.
            PropertyTag[] propertyTagList = new PropertyTag[1];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;

            uint inboxFolderContentTableHandler = 0;
            uint newFolderContentTableHandler = 0;
            int inboxFolderMessageIndex = 0;
            int newFolderMessageIndex = 0;
            RopQueryRowsResponse getInboxFolderMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref inboxFolderContentTableHandler, propertyTagList, ref inboxFolderMessageIndex, mailSubject);
            Site.Assert.AreEqual<uint>(0, getInboxFolderMailMessageContent.ReturnValue, "Getting message on the inbox folder should succeed.");

            RopQueryRowsResponse getNewFolderMailMessageContent = this.GetExpectedMessage(newFolderHandle, ref newFolderContentTableHandler, propertyTagList, ref newFolderMessageIndex, mailSubject);
            Site.Assert.AreEqual<uint>(0, getNewFolderMailMessageContent.ReturnValue, "Getting message on the newly created folder should succeed.");
            string mailSubject2 = AdapterHelper.PropertyValueConvertToString(getNewFolderMailMessageContent.RowData.PropertyRows[newFolderMessageIndex].PropertyValues[0].Value);

            this.VerifyActionTypeOP_COPY(mailSubject2, getNewFolderMailMessageContent, getInboxFolderMailMessageContent, ruleProperties);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R298");

            // Verify MS-OXORULE requirement: MS-OXORULE_R298.
            // When the server copies the message to the destination store, it means StoreEID is set to the destination store EntryID.
            bool isVerifyR298 = mailSubject2.Equals(mailSubject) && getNewFolderMailMessageContent.RowCount == 1 && getInboxFolderMailMessageContent.RowCount == 1;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR298,
                298,
                @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Standard Rules] StoreEID (variable):  A Store Object EntryID structure, as specified in [MS-OXCDATA] section 2.2.4.3, [In OP_COPY action data] Identifies the message store.");
            #endregion

            #region Delete the newly created folder.
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, newFolderId);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to test the server execute rules in order.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC12_ServerExecuteRules_InOrder()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameForward);
            #endregion

            #region TestUser1 adds a rule for ActionType OP_Forward with rule sequence set to 1.
            ForwardDelegateActionData forwardActionData = new ForwardDelegateActionData
            {
                RecipientCount = (ushort)0x01
            };
            RecipientBlock recipientBlock = new RecipientBlock
            {
                Reserved = 0x01,
                NoOfProperties = (ushort)0x04u
            };

            #region Prepare the recipient Block.
            TaggedPropertyValue[] recipientProperties = AdapterHelper.GenerateRecipientPropertiesBlock(this.User2Name, this.User2ESSDN);

            recipientBlock.PropertiesData = recipientProperties;
            #endregion

            forwardActionData.RecipientsData = new RecipientBlock[1] { recipientBlock };
            RuleData ruleForward = AdapterHelper.GenerateValidRuleData(ActionType.OP_FORWARD, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, forwardActionData, ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForward });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Forward rule should succeed.");
            #endregion

            #region TestUser1 adds a delete rule for Inbox folder with rule Sequence set to 100.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameDelete);
            RuleData ruleForDelete = AdapterHelper.GenerateValidRuleData(ActionType.OP_DELETE, TestRuleDataType.ForAdd, 100, RuleState.ST_ENABLED, new DeleteMarkReadActionData(), ruleProperties, null);
            ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleForDelete });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Delete rule should succeed.");
            #endregion

            #region TestUser1 gets rule table.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table Should succeed");
            #endregion

            #region TestUser1 retrieves rule information for the newly added rule.
            PropertyTag[] propertyTags = new PropertyTag[]
            {
                new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleName,
                    PropertyType = (ushort)PropertyType.PtypString
                }
            };

            // Retrieve rows from the rule table.
            RopQueryRowsResponse queryRowResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyTags);
            Site.Assert.AreEqual<uint>(0, queryRowResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");

            // Two rules have been added to the Inbox folder, so the row count in the rule table should be 2.
            Site.Assert.AreEqual<uint>(2, queryRowResponse.RowCount, "The rule number in the rule table is {0}", queryRowResponse.RowCount);
            this.VerifyRuleTable();

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R778");

            // Add two rules to the Inbox folder. If the rule table is got successfully and the rule count is 2,
            // it means that the server had stored all previously created rules.
            Site.CaptureRequirementIfAreEqual<uint>(
                2,
                queryRowResponse.RowCount,
                778,
                @"[In Returning and Maintaining the Rules Table] When a user creates or modifies a rule using the RopModifyRules ROP request ([MS-OXCROPS] section 2.2.11.1), the server MUST store this and all previously created rules.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R779: the row count of the rule table is {0}", queryRowResponse.RowCount);

            // Add two rules to the Inbox folder. If the rule table is got successfully and the rule count is 2. This requirement can be verified.
            Site.CaptureRequirementIfAreEqual<uint>(
                2,
                queryRowResponse.RowCount,
                779,
                @"[In Returning and Maintaining the Rules Table] [When a user creates or modifies a rule using the RopModifyRules ROP request ([MS-OXCROPS] section 2.2.11.1)] The server MUST also respond to a RopGetRulesTable ROP request ([MS-OXCROPS] section 2.2.11.2) by returning these rules to the client in the form of a rules table.");
            #endregion
            #endregion

            #region TestUser1 delivers a message to itself to trigger these rules.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser1 delivers a message to itself to trigger these rules.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User1Name, this.User1Password, this.User1Name, mailSubject);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the messages in the Inbox to verify the rule evaluation.
            PropertyTag[] propertyTagList = new PropertyTag[2];

            // pidTagSubject and pidTagMessageClass
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagMessageClass;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;
            uint contentsTableHandle = 0;
            bool doesUnexpectedMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, mailSubject);
            #endregion

            #region TestUser2 gets the forwarded message to verify the rule evaluation.
            // Let Testuser2 to logon to the server
            this.LogonMailbox(TestUser.TestUser2);
            uint contentTableHandler = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse testUser2getNormalMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandler, propertyTagList, ref expectedMessageIndex, mailSubject);
            Site.Assert.AreEqual<uint>(0, testUser2getNormalMailMessageContent.ReturnValue, "Getting message property operation should succeed.");

            string testUser2mailSubject = AdapterHelper.PropertyValueConvertToString(testUser2getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R514, the value of mailSubject is {0}", testUser2mailSubject);

            // Verify MS-OXORULE requirement: MS-OXORULE_R514.
            // This test case has added 2 rules (OP_FORWARD and OP_DELETE) in increasing order of the value of the PidTagRuleSequence property.
            // According to this order, OP_FORWARD should be performed before OP_DELETE.
            // If there is a message existing under the forwarded recipient's Inbox folder, whose subject name contains the original received message's subject name,
            // it means the server has forwarded the message to the corresponding recipient before executing the OP_DELETE rule.
            // If the RowCount of getMailMessageContent equals 0, it means the original received message under the administrator's Inbox folder
            // has been deleted. So it proves the server evaluates each rule in the increasing order.
            bool isVerifyR514 = testUser2mailSubject.ToUpperInvariant().Contains(mailSubject.ToUpperInvariant()) && !doesUnexpectedMessageExist;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR514,
                514,
                @"[In Processing Incoming Messages to a Folder] For each message delivered to a folder, the server evaluates each rule (2) in that folder in increasing order of the value of the PidTagRuleSequence property (section 2.2.1.3.1.2) in each rule (2).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_MOVE rule with FolderInThisStore set to 1.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC13_ServerExecuteRule_Action_OP_MOVE_FolderInThisStore()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameMoveOne);
            #endregion

            #region TestUser1 creates folder1 in server store.
            RopCreateFolderResponse createFolderResponse;
            uint newFolderHandle = this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, Common.GenerateResourceName(this.Site, "User1Folder01"), "TestForOP_MOVE", out createFolderResponse);
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

            #region TestUser2 delivers a message to TestUser1 to trigger the rule.

            // TestUser2 deliver a message to trigger the rule.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message content to verify the rule evaluation.
            uint inboxFolderContentsTableHandle = 0;
            PropertyTag[] propertyTagList = new PropertyTag[1];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;

            bool doesOriginalMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref inboxFolderContentsTableHandle, propertyTagList, mailSubject);
            uint newFolder1ContentsTableHandle = 0;
            uint rowCount = 0;
            RopQueryRowsResponse getNewFolder1MailMessageContent = this.GetExpectedMessage(newFolderHandle, ref newFolder1ContentsTableHandle, propertyTagList, ref rowCount, 1, mailSubject);
            Site.Assert.IsFalse(doesOriginalMessageExist, "The original message shouldn't exist anymore.");
            Site.Assert.AreEqual<uint>(0, getNewFolder1MailMessageContent.ReturnValue, "getNewFolder1MailMessageContent should succeed.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R956");

            // Verify MS-OXORULE requirement: MS-OXORULE_R956
            // Since the destination folder is created in the user's mailbox, if the Email can be found in the destination folder, MS-OXORULE_R956 can be verified.
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                getNewFolder1MailMessageContent.RowCount,
                956,
                @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Standard Rules] The destination folder for a Move action in a standard rule can be in the user's mailbox.");
            #endregion

            if (Common.IsRequirementEnabled(294, this.Site))
            {
                #region TestUser1 gets a rule table.

                RopGetRulesTableResponse ropGetRulesTableResponse;
                uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
                Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");
                #endregion

                #region TestUser1 retrieves rule information to check if the rule exists.
                PropertyTag[] propertyTags = new PropertyTag[2];
                propertyTags[0].PropertyId = (ushort)PropertyId.PidTagRuleName;
                propertyTags[0].PropertyType = (ushort)PropertyType.PtypString;
                propertyTags[1].PropertyId = (ushort)PropertyId.PidTagRuleActions;
                propertyTags[1].PropertyType = (ushort)PropertyType.PtypRuleAction;

                // Retrieves rows from the rule table.
                RopQueryRowsResponse queryRowResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyTags);
                Site.Assert.AreEqual<uint>(0, queryRowResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");
                MoveCopyActionData moveActionDataOfQueryRowsResponse = new MoveCopyActionData();
                RuleAction ruleAction = new RuleAction();
                for (int i = 0; i < queryRowResponse.RowCount; i++)
                {
                    System.Text.UnicodeEncoding converter = new UnicodeEncoding();
                    string ruleName = converter.GetString(queryRowResponse.RowData.PropertyRows.ToArray()[i].PropertyValues[0].Value);
                    if (ruleName == ruleProperties.Name + "\0")
                    {
                        // Verify structure RuleAction 
                        ruleAction.Deserialize(queryRowResponse.RowData.PropertyRows[i].PropertyValues[1].Value);
                        moveActionDataOfQueryRowsResponse.Deserialize(ruleAction.Actions[0].ActionDataValue.Serialize());
                        break;
                    }
                }

                #region Capture Code

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R736");

                ServerEID serverEIDOfResponse = new ServerEID(new byte[] { });
                serverEIDOfResponse.Deserialize(moveActionDataOfQueryRowsResponse.FolderEID);
                bool isVerifiedR736 = moveActionDataOfQueryRowsResponse.FolderInThisStore == 0x01 && serverEIDOfResponse.FolderID != null && serverEIDOfResponse.MessageID != null && serverEIDOfResponse.Instance != null;

                // Verify MS-OXORULE requirement: MS-OXORULE_R736
                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR736,
                    736,
                    @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Standard Rules] If the value of the FolderInThisStore field is 0x01, this field [FolderEID] contains a ServerEid structure, as specified in section 2.2.5.1.2.1.1. ");

                FolderID folderId = new FolderID();
                folderId.Deserialize(BitConverter.ToUInt64(serverEIDOfResponse.FolderID, 0));

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R983");

                bool isVerifiedR983 = folderId.ReplicaId != null && folderId.GlobalCounter != null;

                // Verify MS-OXORULE requirement: MS-OXORULE_R983
                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR983,
                    983,
                    @"[In ServerEid Structure] FolderId (8 bytes): A Folder ID structure, as specified in [MS-OXCDATA] section 2.2.1.1, identifies the destination folder.");
                if (Common.IsRequirementEnabled(7032, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R7032");

                    bool isVerifiedR703 = moveActionDataOfQueryRowsResponse.FolderInThisStore == 0x01 && getNewFolder1MailMessageContent.RowCount != 0;

                    // Verify MS-OXORULE requirement: MS-OXORULE_R7032
                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR703,
                        7032,
                        @"[In Appendix A: Product Behavior] Implementation does set this field (FolderInThisStore) to 0x01 if the destination folder is in the user's mailbox. (Exchange 2007 and Exchange 2016 follow this behavior).");
                }
                #endregion
                #endregion
            }

            #region Delete the newly created folder.
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, newFolderId);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_COPY rule with FolderInThisStore set to 1.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC14_ServerExecuteRule_Action_OP_COPY_FolderInThisStore()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameCopy);
            #endregion

            #region TestUser1 creates folder1 in server store.
            RopCreateFolderResponse createFolderResponse;
            uint newFolderHandle = this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, Common.GenerateResourceName(this.Site, "User1Folder01"), "TestForOP_COPY", out createFolderResponse);
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
            // Add rule for OP_COPY without rule Provider Data.
            ruleProperties.ProviderData = string.Empty;
            RuleData ruleForMoveFolder = AdapterHelper.GenerateValidRuleDataWithFlavor(new ActionType[] { ActionType.OP_COPY }, 0, RuleState.ST_ENABLED, moveCopyAction, new uint[] { 0 }, ruleProperties);

            #endregion

            #region TestUser1 adds OP_COPY rule to the Inbox folder.
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForMoveFolder });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Copy rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger the rule.

            // TestUser2 deliver a message to trigger the rule.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message content to verify the rule evaluation.

            uint inboxFolderContentsTableHandle = 0;
            PropertyTag[] propertyTagList = new PropertyTag[1];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;
            uint rowCount = 0;
            RopQueryRowsResponse getInboxMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref inboxFolderContentsTableHandle, propertyTagList, ref rowCount, 1, mailSubject);
            uint newFolder1ContentsTableHandle = 0;
            rowCount = 0;
            RopQueryRowsResponse getNewFolder1MailMessageContent = this.GetExpectedMessage(newFolderHandle, ref newFolder1ContentsTableHandle, propertyTagList, ref rowCount, 1, mailSubject);
            Site.Assert.AreEqual<uint>(0, getNewFolder1MailMessageContent.ReturnValue, "Getting message on the newly created folder should succeed.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R957");

            bool isVerifiedR957 = getInboxMailMessageContent.RowCount != 0 && getNewFolder1MailMessageContent.RowCount != 0;

            // Verify MS-OXORULE requirement: MS-OXORULE_R957
            // Since the destination folder is created in the user's mailbox, if the Emails can be found in the destination folder and Inbox folder, MS-OXORULE_R957 can be verified. 
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR957,
                957,
                @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Standard Rules] The destination folder for a Copy action in a standard rule can be in the user's mailbox.");
            #endregion

            #region Delete the newly created folder.
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, newFolderId);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_MOVE extended rule.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC15_ServerExecuteExtendedRule_Action_OP_MOVE()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameMoveOne);
            #endregion

            #region TestUser1 creates folder1 in server store.
            RopCreateFolderResponse createFolderResponse;
            uint newFolderHandle = this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, Common.GenerateResourceName(this.Site, "User1Folder01"), "TestForOP_MOVE", out createFolderResponse);
            ulong newFolderId = createFolderResponse.FolderId;
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating folder operation should succeed.");
            #endregion

            #region Prepare rules' data
            MoveCopyActionDataOfExtendedRule moveActionData = new MoveCopyActionDataOfExtendedRule();

            // Get the created folder1 entry id.
            byte[] folder1EId = this.OxoruleAdapter.GetFolderEntryId(StoreObjectType.Mailbox, newFolderHandle, newFolderId);

            // Get the store object's entry id
            byte[] storeEId = this.GetStoreObjectEntryID(StoreObjectType.Mailbox, this.Server, this.User1ESSDN);
            moveActionData.FolderEID = folder1EId;
            moveActionData.StoreEID = storeEId;
            moveActionData.FolderEIDSize = (uint)folder1EId.Length;
            moveActionData.StoreEIDSize = (uint)storeEId.Length;
            #endregion

            #region TestUser1 creates an FAI message for the extended rule.
            RopCreateMessageResponse ropCreateMessageResponse;
            uint extendedRuleMessageHandle1 = this.OxoruleAdapter.RopCreateMessage(this.InboxFolderHandle, this.InboxFolderID, Convert.ToByte(true), out ropCreateMessageResponse);
            Site.Assert.AreEqual<uint>(0, ropCreateMessageResponse.ReturnValue, "Creating the FAI message should succeed.");

            NamedPropertyInfo namedPropertyInfo1 = new NamedPropertyInfo
            {
                NoOfNamedProps = 0
            };
            TaggedPropertyValue[] extendedRuleProperties1 = AdapterHelper.GenerateExtendedRuleTestData(ruleProperties.Name, 0, (uint)RuleState.ST_ENABLED, Constants.PidTagRuleProvider, ActionType.OP_MOVE, moveActionData, ruleProperties.ConditionSubjectName, namedPropertyInfo1);

            // Set properties for extended rule FAI message.
            RopSetPropertiesResponse ropSetPropertiesResponse = this.OxoruleAdapter.RopSetProperties(extendedRuleMessageHandle1, extendedRuleProperties1);
            Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "Setting property for Extended rule FAI message should succeed.");

            // Save changes of message.
            RopSaveChangesMessageResponse ropSaveChangesMessagResponse = this.OxoruleAdapter.RopSaveChangesMessage(extendedRuleMessageHandle1);
            Site.Assert.AreEqual(0, (int)ropSaveChangesMessagResponse.ReturnValue, "Saving Extend rule message should succeed.");
            #endregion

            #region TestUser1 gets the extended rule.
            #region Step1: TestUser1 calls RopGetContentsTable to get a table of all messages which are placed in the Inbox folder.
            uint contentsTableHandleOfFAIMessage;

            // Call RopGetContentsTable.
            RopGetContentsTableResponse ropGetContentsTableResponseOfFAIMessage = this.OxoruleAdapter.RopGetContentsTable(this.InboxFolderHandle, ContentTableFlag.Associated, out contentsTableHandleOfFAIMessage);
            Site.Assert.AreEqual<uint>(0, ropGetContentsTableResponseOfFAIMessage.ReturnValue, "Getting contents table should succeed, the actual returned value is {0}", ropGetContentsTableResponseOfFAIMessage.ReturnValue);
            #endregion

            #region Step2: TestUser1 calls RopSetColumns to set the interested columns of the message table in the Inbox folder.

            // Here are 6 interested columns listed as below.
            // Prepare the data in the RopSetColumns request buffer.
            PropertyTag[] propertyTagOfFAIMessage = new PropertyTag[6];
            PropertyTag pidTagRuleMessageNameTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleMessageName,
                PropertyType = (ushort)PropertyType.PtypString
            };
            propertyTagOfFAIMessage[0] = pidTagRuleMessageNameTag;
            PropertyTag pidTagMessageClassTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            propertyTagOfFAIMessage[1] = pidTagMessageClassTag;
            PropertyTag pidTagRuleMessageStatePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleMessageState,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            propertyTagOfFAIMessage[2] = pidTagRuleMessageStatePropertyTag;
            PropertyTag pidTagRuleMessageProviderPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleMessageProvider,
                PropertyType = (ushort)PropertyType.PtypString
            };
            propertyTagOfFAIMessage[3] = pidTagRuleMessageProviderPropertyTag;
            PropertyTag pidTagExtendedRuleMessageActionsPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagExtendedRuleMessageActions,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            propertyTagOfFAIMessage[4] = pidTagExtendedRuleMessageActionsPropertyTag;
            PropertyTag pidTagExtendedRuleMessageConditionPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagExtendedRuleMessageCondition,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            propertyTagOfFAIMessage[5] = pidTagExtendedRuleMessageConditionPropertyTag;

            // Query rows which include the property values of the interested columns. 
            RopQueryRowsResponse ropQueryRowsResponseOfFAIMessage = this.OxoruleAdapter.QueryPropertiesInTable(contentsTableHandleOfFAIMessage, propertyTagOfFAIMessage);
            Site.Assert.AreEqual<uint>(0, ropQueryRowsResponseOfFAIMessage.ReturnValue, "Querying Rows Response of FAI Message should succeed, the actual returned value is {0}", ropQueryRowsResponseOfFAIMessage.ReturnValue);
            MoveCopyActionDataOfExtendedRule moveActionDataOfQueryRowsResponse = new MoveCopyActionDataOfExtendedRule();
            ExtendedRuleActions extendedRuleActions = new ExtendedRuleActions();
            for (int i = 0; i < ropQueryRowsResponseOfFAIMessage.RowCount; i++)
            {
                System.Text.UnicodeEncoding converter = new UnicodeEncoding();
                string messageName = converter.GetString(ropQueryRowsResponseOfFAIMessage.RowData.PropertyRows.ToArray()[i].PropertyValues[0].Value);
                if (messageName == ruleProperties.Name + "\0")
                {
                    byte[] extendedRuleMessageActionBinary = ropQueryRowsResponseOfFAIMessage.RowData.PropertyRows.ToArray()[i].PropertyValues[4].Value;
                    byte[] extendedRuleMessageActionBuffer = new byte[extendedRuleMessageActionBinary.Length - 2];
                    Array.Copy(extendedRuleMessageActionBinary, 2, extendedRuleMessageActionBuffer, 0, extendedRuleMessageActionBinary.Length - 2);
                    extendedRuleActions.Deserialize(extendedRuleMessageActionBuffer);
                    moveActionDataOfQueryRowsResponse.Deserialize(extendedRuleActions.RuleActionBuffer.Actions[0].ActionDataValue.Serialize());
                }
            }

            #region Capture Code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R936");

            // Verify MS-OXORULE requirement: MS-OXORULE_R936
            bool isVerifiedR936 = extendedRuleActions.NamedPropertyInformation.NoOfNamedProps == 0 && extendedRuleActions.NamedPropertyInformation.PropId == null && extendedRuleActions.NamedPropertyInformation.NamedPropertiesSize == 0 && extendedRuleActions.NamedPropertyInformation.NamedProperty == null;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR936,
                936,
                @"[In NamedPropertyInformation Structure] [If no named properties are used in the structure that follows the NamedPropertyInformation structure] no other fields [except NoOfNamedProps] are present.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R978");

            bool isVerifiedR978 = Common.CompareByteArray(folder1EId, moveActionDataOfQueryRowsResponse.FolderEID);

            // Verify MS-OXORULE requirement: MS-OXORULE_R978
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR978,
                978,
                @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Extended Rules] FolderEID (variable): A Folder EntryID structure, as specified in [MS-OXCDATA] section 2.2.4.1, [In OP_MOVE action data] identifies the destination folder.");
            #endregion

            #endregion
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger the rule.

            // TestUser2 deliver a message to trigger the rule.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message content to verify the rule evaluation.
            uint inboxFolderContentsTableHandle = 0;
            PropertyTag[] propertyTagList = new PropertyTag[1];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;

            bool doesOriginalMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref inboxFolderContentsTableHandle, propertyTagList, mailSubject);
            uint newFolder1ContentsTableHandle = 0;
            uint rowCount = 0;
            RopQueryRowsResponse getNewFolder1MailMessageContent = this.GetExpectedMessage(newFolderHandle, ref newFolder1ContentsTableHandle, propertyTagList, ref rowCount, 1, mailSubject);
            Site.Assert.IsFalse(doesOriginalMessageExist, "The original message shouldn't exist anymore.");
            Site.Assert.AreEqual<uint>(0, getNewFolder1MailMessageContent.ReturnValue, "getNewFolder1MailMessageContent should succeed.");

            #region Capture code

            this.VerifyActionTypeOP_MOVE(getNewFolder1MailMessageContent, doesOriginalMessageExist);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R968, since the destination folder is created in the user's mailbox, try to verify the message was move to the folder");

            // Verify MS-OXORULE requirement: MS-OXORULE_R968
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                getNewFolder1MailMessageContent.RowCount,
                968,
                @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Extended Rules] The destination folder for a Move action in an extended rule MUST be in the user's mailbox.");
            #endregion
            #endregion

            #region Delete the newly created folder.
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, newFolderId);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_COPY extended rule.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC16_ServerExecuteExtendedRule_Action_OP_COPY()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameMoveOne);
            #endregion

            #region TestUser1 creates folder1 in server store.
            RopCreateFolderResponse createFolderResponse;
            uint newFolderHandle = this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, Common.GenerateResourceName(this.Site, "User1Folder01"), "TestForOP_COPY", out createFolderResponse);
            ulong newFolderId = createFolderResponse.FolderId;
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating folder operation should succeed.");
            #endregion

            #region Prepare rules' data
            MoveCopyActionDataOfExtendedRule copyActionData = new MoveCopyActionDataOfExtendedRule();

            // Get the created folder1 entry id.
            byte[] folder1EId = this.OxoruleAdapter.GetFolderEntryId(StoreObjectType.Mailbox, newFolderHandle, newFolderId);

            // Get the store object's entry id
            byte[] storeEId = this.GetStoreObjectEntryID(StoreObjectType.Mailbox, this.Server, this.User1ESSDN);
            copyActionData.FolderEID = folder1EId;
            copyActionData.StoreEID = storeEId;
            copyActionData.FolderEIDSize = (uint)folder1EId.Length;
            copyActionData.StoreEIDSize = (uint)storeEId.Length;
            #endregion

            #region TestUser1 creates an FAI message for the extended rule.
            RopCreateMessageResponse ropCreateMessageResponse;
            uint extendedRuleMessageHandle1 = this.OxoruleAdapter.RopCreateMessage(this.InboxFolderHandle, this.InboxFolderID, Convert.ToByte(true), out ropCreateMessageResponse);
            Site.Assert.AreEqual<uint>(0, ropCreateMessageResponse.ReturnValue, "Creating the FAI message should succeed.");

            NamedPropertyInfo namedPropertyInfo1 = new NamedPropertyInfo
            {
                NoOfNamedProps = 0
            };
            TaggedPropertyValue[] extendedRuleProperties1 = AdapterHelper.GenerateExtendedRuleTestData(ruleProperties.Name, 0, (uint)RuleState.ST_ENABLED, Constants.PidTagRuleProvider, ActionType.OP_COPY, copyActionData, ruleProperties.ConditionSubjectName, namedPropertyInfo1);

            // Set properties for extended rule FAI message.
            RopSetPropertiesResponse ropSetPropertiesResponse = this.OxoruleAdapter.RopSetProperties(extendedRuleMessageHandle1, extendedRuleProperties1);
            Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "Setting property for Extended rule FAI message should succeed.");

            // Save changes of message.
            RopSaveChangesMessageResponse ropSaveChangesMessagResponse = this.OxoruleAdapter.RopSaveChangesMessage(extendedRuleMessageHandle1);
            Site.Assert.AreEqual(0, (int)ropSaveChangesMessagResponse.ReturnValue, "Saving Extend rule message should succeed.");
            #endregion

            #region TestUser1 gets the extended rule.
            #region Step1: TestUser1 gets a table of all messages which are placed in the Inbox folder.
            uint contentsTableHandleOfFAIMessage;
            RopGetContentsTableResponse ropGetContentsTableResponseOfFAIMessage = this.OxoruleAdapter.RopGetContentsTable(this.InboxFolderHandle, ContentTableFlag.Associated, out contentsTableHandleOfFAIMessage);
            Site.Assert.AreEqual<uint>(0, ropGetContentsTableResponseOfFAIMessage.ReturnValue, "Getting contents table should succeed, the actual returned value is {0}", ropGetContentsTableResponseOfFAIMessage.ReturnValue);
            #endregion

            #region Step2: TestUser1 sets the interested columns of the message table in the Inbox folder.

            // Here are 6 interested columns listed as below.
            PropertyTag[] propertyTagOfFAIMessage = new PropertyTag[6];
            PropertyTag pidTagRuleMessageNameTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleMessageName,
                PropertyType = (ushort)PropertyType.PtypString
            };
            propertyTagOfFAIMessage[0] = pidTagRuleMessageNameTag;
            PropertyTag pidTagMessageClassTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            propertyTagOfFAIMessage[1] = pidTagMessageClassTag;
            PropertyTag pidTagRuleMessageStatePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleMessageState,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            propertyTagOfFAIMessage[2] = pidTagRuleMessageStatePropertyTag;
            PropertyTag pidTagRuleMessageProviderPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleMessageProvider,
                PropertyType = (ushort)PropertyType.PtypString
            };
            propertyTagOfFAIMessage[3] = pidTagRuleMessageProviderPropertyTag;
            PropertyTag pidTagExtendedRuleMessageActionsPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagExtendedRuleMessageActions,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            propertyTagOfFAIMessage[4] = pidTagExtendedRuleMessageActionsPropertyTag;
            PropertyTag pidTagExtendedRuleMessageConditionPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagExtendedRuleMessageCondition,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            propertyTagOfFAIMessage[5] = pidTagExtendedRuleMessageConditionPropertyTag;

            // Query rows which include the property values of the interested columns.
            RopQueryRowsResponse ropQueryRowsResponseOfFAIMessage = this.OxoruleAdapter.QueryPropertiesInTable(contentsTableHandleOfFAIMessage, propertyTagOfFAIMessage);
            Site.Assert.AreEqual<uint>(0, ropQueryRowsResponseOfFAIMessage.ReturnValue, "Querying Rows Response of FAI Message should succeed, the actual returned value is {0}", ropQueryRowsResponseOfFAIMessage.ReturnValue);
            MoveCopyActionDataOfExtendedRule copyActionDataOfQueryRowsResponse = new MoveCopyActionDataOfExtendedRule();
            for (int i = 0; i < ropQueryRowsResponseOfFAIMessage.RowCount; i++)
            {
                System.Text.UnicodeEncoding converter = new UnicodeEncoding();
                string messageName = converter.GetString(ropQueryRowsResponseOfFAIMessage.RowData.PropertyRows.ToArray()[i].PropertyValues[0].Value);
                if (messageName == ruleProperties.Name + "\0")
                {
                    byte[] extendedRuleMessageActionBinary = ropQueryRowsResponseOfFAIMessage.RowData.PropertyRows.ToArray()[i].PropertyValues[4].Value;
                    byte[] extendedRuleMessageActionBuffer = new byte[extendedRuleMessageActionBinary.Length - 2];
                    Array.Copy(extendedRuleMessageActionBinary, 2, extendedRuleMessageActionBuffer, 0, extendedRuleMessageActionBinary.Length - 2);
                    ExtendedRuleActions extendedRuleActions = new ExtendedRuleActions();
                    extendedRuleActions.Deserialize(extendedRuleMessageActionBuffer);
                    copyActionDataOfQueryRowsResponse.Deserialize(extendedRuleActions.RuleActionBuffer.Actions[0].ActionDataValue.Serialize());
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R977");

            bool isVerifiedR977 = Common.CompareByteArray(folder1EId, copyActionDataOfQueryRowsResponse.FolderEID);

            // Verify MS-OXORULE requirement: MS-OXORULE_R977
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR977,
                977,
                @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Extended Rules] FolderEID (variable): A Folder EntryID structure, as specified in [MS-OXCDATA] section 2.2.4.1, [In OP_COPY action data] identifies the destination folder.");
            #endregion
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger the rule.

            // TestUser2 deliver a message to trigger these rules.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message content to verify the rule evaluation.
            uint inboxFolderContentsTableHandle = 0;
            PropertyTag[] propertyTagList = new PropertyTag[1];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;

            uint rowCount = 0;
            RopQueryRowsResponse getInboxMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref inboxFolderContentsTableHandle, propertyTagList, ref rowCount, 1, mailSubject);
            uint newFolder1ContentsTableHandle = 0;
            rowCount = 0;
            RopQueryRowsResponse getNewFolder1MailMessageContent = this.GetExpectedMessage(newFolderHandle, ref newFolder1ContentsTableHandle, propertyTagList, ref rowCount, 1, mailSubject);
            Site.Assert.AreEqual<uint>(0, getNewFolder1MailMessageContent.ReturnValue, "getNewFolder1MailMessageContent should succeed.");

            #region Capture code

            this.VerifyActionTypeOP_COPY(mailSubject, getNewFolder1MailMessageContent, getInboxMailMessageContent, ruleProperties);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R969, since the destination folder is created in the user's mailbox, try to verify the message was copy to the folder");

            // Verify MS-OXORULE requirement: MS-OXORULE_R969
            bool isVerifiedR969 = getInboxMailMessageContent.RowCount != 0 && getNewFolder1MailMessageContent.RowCount != 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR969,
                969,
                @"[In OP_MOVE and OP_COPY ActionData Structure] [Buffer Format for Extended Rules] The destination folder for a Copy action in an extended rule MUST be in the user's mailbox.");
            #endregion
            #endregion

            #region Delete the newly created folder.
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, newFolderId);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of rule forward as attachment.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC17_ServerExecuteRule_ForwardAsAttachment()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameForwardAT);
            #endregion

            #region TestUser1 adds a rule for ActionType with OP_Forward ActionFlavor set to AT.
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
            RuleData ruleForwardAT = AdapterHelper.GenerateValidRuleDataWithFlavor(ActionType.OP_FORWARD, 0, RuleState.ST_ENABLED, forwardActionData, (uint)ActionFlavorsForward.AT, ruleProperties);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForwardAT });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Forward rule should succeed.");
            #endregion

            #region TestUser1 delivers a message to itself to trigger the rule.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser1 delivers a message to itself to trigger the rule.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User1Name, this.User1Password, this.User1Name, mailSubject);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser2 gets the forwarded message to verify the rule evaluation.

            // Let TestUser2 log on to the server.
            this.LogonMailbox(TestUser.TestUser2);
            PropertyTag[] propertyTagList = new PropertyTag[3];

            // pidTagSubject and pidTagMessageClass
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagMessageClass;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[2].PropertyId = (ushort)PropertyId.pidTagHasAttachment;
            propertyTagList[2].PropertyType = (ushort)PropertyType.PtypBoolean;
            uint contentsTableHandle = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse testUser2getNormalMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, mailSubject);
            Site.Assert.AreEqual<uint>(0, testUser2getNormalMailMessageContent.ReturnValue, "Getting message property operation should succeed.");

            mailSubject = AdapterHelper.PropertyValueConvertToString(testUser2getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value);
            byte[] hasAttachment = testUser2getNormalMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[2].Value;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R279");

            // Verify MS-OXORULE requirement: MS-OXORULE_R279
            bool isVerifiedR279 = hasAttachment.Length == 1 && hasAttachment[0] == 1;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR279,
                279,
                @"[In Action Flavors] AT (Bitmask 0x00000004): Forwards the message as an attachment.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of multiple OP_MOVE rules.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC18_ServerExecuteRule_Action_MultipleOP_MOVE()
        {
            this.CheckMAPIHTTPTransportSupported();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(929, this.Site), "This case runs only when the server supports OP_MOVE action when FolderInThisStore is set to 0.");
            Site.Assume.IsTrue(Common.IsRequirementEnabled(904, this.Site), "This case runs only when the server supports creating multiple copies of the message for OP_MOVE action.");

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameMoveOne);
            #endregion

            #region TestUser1 creates folder1 in server store.
            RopCreateFolderResponse createFolderResponse;
            uint newFolderHandle = this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, Common.GenerateResourceName(this.Site, "User1Folder01"), "TestForOP_MOVE", out createFolderResponse);
            ulong newFolderId = createFolderResponse.FolderId;
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating folder operation should succeed.");
            #endregion

            #region TestUser1 creates folder2 in server store.
            RopCreateFolderResponse createFolder2Response;
            uint newFolder2Handle = this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, Common.GenerateResourceName(this.Site, "User1Folder02"), "TestForOP_MOVE", out createFolder2Response);
            ulong newFolder2Id = createFolder2Response.FolderId;
            Site.Assert.AreEqual<uint>(0, createFolder2Response.ReturnValue, "Creating folder operation should succeed.");
            #endregion

            #region Prepare rules' data
            MoveCopyActionData moveCopyActionData1 = new MoveCopyActionData();

            // Get the created folder1 entry id.
            byte[] folder1EId = this.OxoruleAdapter.GetFolderEntryId(StoreObjectType.Mailbox, newFolderHandle, newFolderId);

            // Get the store object's entry id
            byte[] storeEId = this.GetStoreObjectEntryID(StoreObjectType.Mailbox, this.Server, this.User1ESSDN);
            moveCopyActionData1.FolderInThisStore = 0;
            moveCopyActionData1.FolderEID = folder1EId;
            moveCopyActionData1.StoreEID = storeEId;
            moveCopyActionData1.FolderEIDSize = (ushort)folder1EId.Length;
            moveCopyActionData1.StoreEIDSize = (ushort)storeEId.Length;
            MoveCopyActionData moveCopyActionData2 = new MoveCopyActionData();

            // Get the created folder2 entry ID.
            byte[] folder2EId = this.OxoruleAdapter.GetFolderEntryId(StoreObjectType.Mailbox, newFolder2Handle, newFolder2Id);
            moveCopyActionData2.FolderInThisStore = 0;
            moveCopyActionData2.FolderEID = folder2EId;
            moveCopyActionData2.StoreEID = storeEId;
            moveCopyActionData2.FolderEIDSize = (ushort)folder2EId.Length;
            moveCopyActionData2.StoreEIDSize = (ushort)storeEId.Length;
            IActionData[] moveCopyActionData = { moveCopyActionData1, moveCopyActionData2 };
            #endregion

            #region Generate test RuleData.
            // Add rule for move without rule Provider Data.
            ruleProperties.ProviderData = string.Empty;
            RuleData ruleForMoveFolder = AdapterHelper.GenerateValidRuleDataWithFlavor(new ActionType[] { ActionType.OP_MOVE, ActionType.OP_MOVE }, 0, RuleState.ST_ENABLED, moveCopyActionData, new uint[] { 0, 0 }, ruleProperties);

            #endregion

            #region TestUser1 adds OP_MOVE rule to the Inbox folder.
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForMoveFolder });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Move rule should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.

            // TestUser2 deliver a message to trigger these rules.
            string mailSubject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName + "Title");
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject);

            // Wait for the mail to be received and the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser1 gets the message content to verify the rule evaluation.
            uint inboxFolderContentsTableHandle = 0;
            PropertyTag[] propertyTagList = new PropertyTag[1];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;

            bool doesOriginalMessageExist = this.CheckUnexpectedMessageExist(this.InboxFolderHandle, ref inboxFolderContentsTableHandle, propertyTagList, mailSubject);
            uint newFolder1ContentsTableHandle = 0;
            uint rowCount = 0;
            RopQueryRowsResponse getNewFolder1MailMessageContent = this.GetExpectedMessage(newFolderHandle, ref newFolder1ContentsTableHandle, propertyTagList, ref rowCount, 1, mailSubject);
            Site.Assert.AreEqual<uint>(0, getNewFolder1MailMessageContent.ReturnValue, "Getting message on the first folder should succeed.");
            uint newFolder2ContentsTableHandle = 0;
            rowCount = 0;
            RopQueryRowsResponse getNewFolder2MailMessageContent = this.GetExpectedMessage(newFolder2Handle, ref newFolder2ContentsTableHandle, propertyTagList, ref rowCount, 1, mailSubject);
            Site.Assert.AreEqual<uint>(0, getNewFolder2MailMessageContent.ReturnValue, "Getting message on the second folder should succeed.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R904: the message count in NewFolder1 and NewFolder2 are {0} and {1}", getNewFolder1MailMessageContent.RowCount, getNewFolder2MailMessageContent.RowCount);

            // Verify MS-OXORULE requirement: MS-OXORULE_R904
            bool isVerifyR904 = getNewFolder1MailMessageContent.RowCount == 1 && getNewFolder2MailMessageContent.RowCount == 1 && !doesOriginalMessageExist;

            Site.CaptureRequirementIfIsTrue(
                  isVerifyR904,
                  904,
                  @"[In Processing Incoming Messages to a Folder] [Following is a description of what the server does when it executes each action (2) type, as specified in section 2.2.5.1.1, for an incoming message] ""OP_MOVE"": The implementation does create multiple copies of the message and then delete the original message, if multiple ""OP_MOVE"" operations apply to the same message. (Exchange 2003 and above follow this behavior.)");
            #endregion

            #region Delete the folders created in step2 and step3.
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, newFolderId);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, newFolder2Id);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the rules behavior of ST_SKIP_IF_SCL_IS_SAFE state. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC19_RuleState_ST_SKIP_IF_SCL_IS_SAFE()
        {
            this.CheckMAPIHTTPTransportSupported();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(7411, this.Site), "This test case only runs if implementation does not skip evaluation of this rule (ST_SKIP_IF_SCL_IS_SAFE) when the delivered message's PidTagContentFilterSpamConfidenceLevel property has a value of 0xFFFFFFFF");
            
            #region TestUser1 adds an OP_TAG rule with PidTagRuleState set to ST_ENABLED | ST_SKIP_IF_SCL_IS_SAFE.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameTag);
            TagActionData tagActionData = new TagActionData();
            PropertyTag tagActionDataPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagImportance,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            tagActionData.PropertyTag = tagActionDataPropertyTag;
            tagActionData.PropertyValue = BitConverter.GetBytes(2);

            RuleData ruleOpTag = AdapterHelper.GenerateValidRuleData(ActionType.OP_TAG, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED | RuleState.ST_SKIP_IF_SCL_IS_SAFE, tagActionData, ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleOpTag });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding OP_TAG rule should succeed");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger these rules.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // Let TestUser2 log on to the server
            this.LogonMailbox(TestUser.TestUser2);

            TaggedPropertyValue contentFilterSpamConfidenceLevel = new TaggedPropertyValue();
            PropertyTag contentFilterSpamConfidenceLevelTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagContentFilterSpamConfidenceLevel,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };
            contentFilterSpamConfidenceLevel.PropertyTag = contentFilterSpamConfidenceLevelTag;
            contentFilterSpamConfidenceLevel.Value = BitConverter.GetBytes(0xFFFFFFFF);
            string subject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName);
            this.DeliverMessageToTriggerRule(this.User1Name, this.User1ESSDN, subject, new TaggedPropertyValue[1] { contentFilterSpamConfidenceLevel });

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region Testuser1 verifies whether the specific property value is set on the received mail.
            // Let TestUser2 log on to the server
            this.LogonMailbox(TestUser.TestUser1);

            PropertyTag[] propertyTagList = new PropertyTag[2];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagImportance;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypInteger32;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypString;

            uint contentsTableHandle = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse getMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentsTableHandle, propertyTagList, ref expectedMessageIndex, subject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R7411");

            // Verify MS-OXORULE requirement: MS-OXORULE_R7411
            // If the PidTagImportance is the value which is set on OP_TAG rule, it means the rule tacks action and the rule sets the property 
            // specified in the rule's action buffer structure.
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                BitConverter.ToInt32(getMailMessageContent.RowData.PropertyRows[expectedMessageIndex].PropertyValues[0].Value, 0),
                7411,
                @"[In Appendix A: Product Behavior] Implementation does not skip evaluation of this rule (ST_SKIP_IF_SCL_IS_SAFE) if the delivered message's PidTagContentFilterSpamConfidenceLevel property ([MS-OXPROPS] section 2.638) has a value of 0xFFFFFFFF. (Exchange 2003 and above follow this behavior.)");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the execution of OP_REPLY rule for automatically generated messages,. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S02_TC20_ServerExecuteRule_Action_OP_REPLY_AutoGeneratedMsg()
        {
            this.CheckMAPIHTTPTransportSupported();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5281, this.Site), "This test case only runs when implementation does not avoid sending replies to automatically generated messages to avoid generating endless autoreply loops for \"OP_REPLY\"");

            #region Prepare value for ruleProperties variable
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameReply);
            #endregion

            #region Create a reply template in the TestUser1's Inbox folder.
            ulong replyTemplateMessageId;
            uint replyTemplateMessageHandler;
            string replyTemplateSubject = Common.GenerateResourceName(this.Site, Constants.ReplyTemplateSubject);

            TaggedPropertyValue[] replyTemplateProperties = new TaggedPropertyValue[1];
            replyTemplateProperties[0] = new TaggedPropertyValue();
            PropertyTag replyTemplatePropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagBody,
                PropertyType = (ushort)PropertyType.PtypString
            };
            replyTemplateProperties[0].PropertyTag = replyTemplatePropertyTag;
            replyTemplateProperties[0].Value = Encoding.Unicode.GetBytes(Constants.ReplyTemplateBody + "\0");

            byte[] guidBytes = this.OxoruleAdapter.CreateReplyTemplate(this.InboxFolderHandle, this.InboxFolderID, false, replyTemplateSubject, replyTemplateProperties, out replyTemplateMessageId, out replyTemplateMessageHandler);
            #endregion

            #region TestUser1 adds a reply rule to TestUser1's Inbox folder.
            ReplyActionData replyActionData = new ReplyActionData
            {
                ReplyTemplateGUID = new byte[guidBytes.Length]
            };
            Array.Copy(guidBytes, 0, replyActionData.ReplyTemplateGUID, 0, guidBytes.Length);

            replyActionData.ReplyTemplateFID = this.InboxFolderID;
            replyActionData.ReplyTemplateMID = replyTemplateMessageId;

            RuleData ruleForReply = AdapterHelper.GenerateValidRuleDataWithFlavor(ActionType.OP_REPLY, 0, RuleState.ST_ENABLED, replyActionData, 0x00000000, ruleProperties);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForReply });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding rule with actionFlavor set to 0x00000000 should succeed.");
            #endregion

            #region TestUser2 sends a mail with PidTagAutoForwarded setting to true to the TestUser1 to trigger this rule.
            // Sleep enough time to wait for the rule to take effect.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // Let TestUser2 log on to the server
            this.LogonMailbox(TestUser.TestUser2);

            TaggedPropertyValue autoForwarded = new TaggedPropertyValue();
            PropertyTag autoForwardedTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagAutoForwarded,
                PropertyType = (ushort)PropertyType.PtypBoolean
            };
            autoForwarded.PropertyTag = autoForwardedTag;
            autoForwarded.Value = BitConverter.GetBytes(true);
            string subject = Common.GenerateResourceName(this.Site, ruleProperties.ConditionSubjectName);
            this.DeliverMessageToTriggerRule(this.User1Name, this.User1ESSDN, subject, new TaggedPropertyValue[1] { autoForwarded });

            // Sleep enough time to wait for the rule to be executed on the delivered message.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            #endregion

            #region TestUser2 verifies there are reply messages in the specific folder.
            PropertyTag[] propertyTagList1 = new PropertyTag[3];
            propertyTagList1[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList1[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList1[1].PropertyId = (ushort)PropertyId.PidTagBody;
            propertyTagList1[1].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList1[2].PropertyId = (ushort)PropertyId.PidTagReceivedByEmailAddress;
            propertyTagList1[2].PropertyType = (ushort)PropertyType.PtypString;

            uint contentTableHandler = 0;
            int expectedMessageIndex = 0;
            RopQueryRowsResponse getNormalMailMessageContent = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandler, propertyTagList1, ref expectedMessageIndex, replyTemplateSubject);

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R5281");

            // Verify MS-OXORULE requirement: MS-OXORULE_R5281.
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                getNormalMailMessageContent.RowData.PropertyRows.Count,
                5281,
                @"[In Appendix A: Product Behavior] Implementation does not avoid sending replies to automatically generated messages to avoid generating endless autoreply loops for ""OP_REPLY"". (Exchange 2003 and above follow this behavior.)");
            #endregion
            #endregion
        }

        /// <summary>
        /// Verify action type OP_MOVE.
        /// </summary>
        /// <param name="getFolderMailMessageContent">Message content gotten from the specified folder.</param>
        /// <param name="doesOriginalMessageExist">Whether the original message exists.</param>
        private void VerifyActionTypeOP_MOVE(RopQueryRowsResponse getFolderMailMessageContent, bool doesOriginalMessageExist)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R247: the message count in folder is {0}", getFolderMailMessageContent.RowCount);

            // Verify MS-OXORULE requirement: MS-OXORULE_R247
            bool isVerifyR247 = getFolderMailMessageContent.RowCount == 1 && !doesOriginalMessageExist;

            Site.CaptureRequirementIfIsTrue(
                 isVerifyR247,
                 247,
                 @"[In ActionBlock Structure] The meaning of action type OP_MOVE: Moves the message to a folder.");
        }

        /// <summary>
        /// Verify action type OP_COPY.
        /// </summary>
        /// <param name="mailSubject">The subject of the new mail.</param>
        /// <param name="getNewFolderMailMessageContent">The mail message content gotten from the new folder.</param>
        /// <param name="getInboxMailMessageContent">The mail message content gotten from the Inbox folder.</param>
        /// <param name="ruleProperties">The properties of the current rule.</param>
        private void VerifyActionTypeOP_COPY(string mailSubject, RopQueryRowsResponse getNewFolderMailMessageContent, RopQueryRowsResponse getInboxMailMessageContent, RuleProperties ruleProperties)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R250: the message count in newFolder is {0}, and in inboxFolder is {1}", getNewFolderMailMessageContent.RowCount, getInboxMailMessageContent.RowCount);

            // Verify MS-OXORULE requirement: MS-OXORULE_R250
            bool isVerifyR250 = mailSubject.Contains(ruleProperties.ConditionSubjectName) && getNewFolderMailMessageContent.RowCount == 1;

            Site.CaptureRequirementIfIsTrue(
                  isVerifyR250,
                  250,
                  @"[In ActionBlock Structure] The meaning of action type OP_COPY: Copies the message to a folder.");
        }
    }
}