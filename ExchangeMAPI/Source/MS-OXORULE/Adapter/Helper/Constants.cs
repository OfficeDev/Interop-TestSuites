namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    /// <summary>
    /// This class intends to provide some important constant to retrieve the deployment data.
    /// So that can achieve easy-maintainability and easy-expandability.
    /// </summary>
    public class Constants
    {
        #region Const Global variables

        #region The names of properties from deployment configuration file.
        /// <summary>
        /// The name of domain where server belongs to.
        /// </summary>
        public const string Domain = "Domain";

        /// <summary>
        /// The user name of User1's mailbox.
        /// </summary>
        public const string User1Name = "AdminUserName";

        /// <summary>
        /// The user password of User1's mailbox.
        /// </summary>
        public const string User1Password = "AdminUserPassword";

        /// <summary>
        /// The user name of User2's mailbox.
        /// </summary>
        public const string User2Name = "User2Name";

        /// <summary>
        /// The user password of User2's mailbox.
        /// </summary>
        public const string User2Password = "User2Password";

        /// <summary>
        /// The ESSDN of TestUser1.
        /// </summary>
        public const string User1ESSDN = "AdminUserESSDN";

        /// <summary>
        /// The ESSDN of TestUser2.
        /// </summary>
        public const string User2ESSDN = "User2ESSDN";

        /// <summary>
        /// The computer name of the system under test.
        /// </summary>
        public const string Server = "SutComputerName";

        /// <summary>
        /// Specify the value waiting for the rule to take effect.
        /// </summary>
        public const string WaitForTheRuleToTakeEffect = "WaitForTheRuleToTakeEffect";

        /// <summary>
        /// Specify the value for the OOF web service URL.
        /// </summary>
        public const string SetOOFWebServiceURL = "EwsUrl";

        /// <summary>
        /// Specify the value waiting for the set OOF state complete.
        /// </summary>
        public const string WaitForSetOOFComplete = "WaitForSetOOFComplete";

        /// <summary>
        /// Specify the maximum number of retry times if the client cannot retrieve the expected message in the preconfigured time period for some unknown reasons, for example, limit of server resources.
        /// </summary>
        public const string GetMessageRepeatTime = "GetMessageRepeatTime";

        /// <summary>
        /// Specify the maximum Size allowed for a property value returned.
        /// </summary>
        public const string PropertySizeLimit = "PropertySizeLimit";     
        #endregion

        #region Test Data used in MS_OXORULEAdapter
        /// <summary>
        /// Specify the value for the RopNotify response name, the default value on exchange is "RopNotifyResponse".
        /// </summary>
        public const string NameOfRopNotifyResponse = "RopNotifyResponse";
        #endregion

        #region Test Data used in AdapterHelper
        /// <summary>
        /// Specify the value for the extended rule message class, the default value on exchange is "IPM.ExtendedRule.Message".
        /// </summary>
        public const string ExtendedRuleMessageClass = "IPM.ExtendedRule.Message";

        /// <summary>
        /// Specify the value for PidTagRuleMessageLevel, the default value on exchange is "0".
        /// </summary>
        public const uint ExtendedRuleMessageLevel = 0;

        /// <summary>
        /// Specify the value for extended rule version, the default value on exchange is "1".
        /// </summary>
        public const uint ExtendedRuleVersion = 1;

        /// <summary>
        /// Specify the value for ActionFlavor of Rule Action, the default value on exchange is "0".
        /// </summary>
        public const uint CommonActionFlavor = 0;

        /// <summary>
        /// Specify the value for ActionFlags of Rule Action, the default value on outlook 2003 and above is "0".
        /// </summary>
        public const uint RuleActionFlags = 0;

        /// <summary>
        /// Specify the value for PidTagRuleLevel, the default value on exchange is "0".
        /// </summary>
        public const uint RuleLevel = 0;

        /// <summary>
        /// Specify the value for the PidTagRuleUserFlags, which can be set to a random string for testing such as "1".
        /// </summary>
        public const string PidTagRuleUserFlags1 = "1"; 
        #endregion

        #region Properties use to construct common rule data and relative structures
        /// <summary>
        /// Specify the value for the PidTagRuleProvider, which can be set to a random string for testing such as "RuleOrganizer".
        /// </summary>
        public const string PidTagRuleProvider = "RuleOrganizer";

        /// <summary>
        /// Specify the value for the PidTagRuleProviderData, which can be set to a random string for testing such as "01000000010000002222222270C1E340".
        /// </summary>
        public const string PidTagRuleProviderData = "01000000010000002222222270C1E340";

        /// <summary>
        /// Specify the value for the rule condition subject, which can be set to a random string for testing such as "fdx".
        /// </summary>
        public const string RuleConditionSubjectContainString = "fdx"; 
        #endregion

        #region ReplyTemplate
        /// <summary>
        /// Specify the value for the reply template Subject, which can be set to a random string for testing such as "ReplyTemplateDemo".
        /// </summary>
        public const string ReplyTemplateSubject = "ReplyTemplateDemo";

        /// <summary>
        /// Specify the value for reply template body, which can be set to a random string for testing such as "This is a reply template!".
        /// </summary>
        public const string ReplyTemplateBody = "This is a reply template!";

        /// <summary>
        /// Specify the message class value for reply template, which can be set to a random string for testing using "IPM.Note.rules.ReplyTemplate." as prefix.
        /// </summary>
        public const string ReplyTemplate = "IPM.Note.rules.ReplyTemplate.Microsoft";
        #endregion

        #region OOFTemplate
        /// <summary>
        /// Specify the message class value for OOF reply template, which can be set to a random string for testing using "IPM.Note.rules.OofTemplate." as prefix.
        /// </summary>
        public const string OOFReplyTemplate = "IPM.Note.rules.OofTemplate.Microsoft";

        /// <summary>
        /// The message of OOF reply.
        /// </summary>
        public const string MessageOfOOFReply = "This is my OOF reply";
        #endregion

        #region RuleData used for DAM
        /// <summary>
        /// Specify the value for PidTagRuleCondition of DAM, which can be set to a random string for testing such as "DAM Test".
        /// </summary>
        public const string DAMPidTagRuleConditionSubjectContainString = "DAM Test";

        /// <summary>
        /// Specify the value for PidTagRuleUserFlags of DAM, which can be set to a random string for testing such as "1".
        /// </summary>
        public const string DAMPidTagRuleUserFlags = "1";

        /// <summary>
        /// Specify the value for PidTagRuleProviderData of DAM, which can be set to a random string for testing such as "01000000010000002222222270C1E340".
        /// </summary>
        public const string DAMPidTagRuleProviderData = "01000000010000002222222270C1E340";

        /// <summary>
        /// Specify the value for message class of DAM, the default value on exchange is "IPC.Microsoft Exchange 4.0.Deferred Action".
        /// </summary>
        public const string DAMMessageClass = "IPC.Microsoft Exchange 4.0.Deferred Action";

        #endregion

        #region Different value for properties of rule one
        /// <summary>
        /// Specify the value for the first PidTagRuleProvider of DAM, which can be set to a random string for testing such as "RuleOrganizerOne".
        /// </summary>
        public const string DAMPidTagRuleProviderOne = "RuleOrganizerOne";

        /// <summary>
        /// Specify the value for the first PidTagRuleAction of DAM, which can be set to a random string for testing such as "DAM_RuleAction_One".
        /// </summary>
        public const string DAMPidTagRuleActionOne = "DAM_RuleAction_One";

        /// <summary>
        /// Specify the value for the first PidTagRuleName of DAM, which can be set to a random string for testing such as "DAM_RuleOne".
        /// </summary>
        public const string DAMPidTagRuleNameOne = "DAM_RuleOne"; 
        #endregion

        #region Different value for properties of rule two
        /// <summary>
        /// Specify the value for the second PidTagRuleName of DAM, which can be set to a random string for testing such as "DAM_RuleTwo".
        /// </summary>
        public const string DAMPidTagRuleNameTwo = "DAM_RuleTwo";

        /// <summary>
        /// Specify the value for the second PidTagRuleAction of DAM, which can be set to a random string for testing such as "DAM_RuleAction_Two".
        /// </summary>
        public const string DAMPidTagRuleActionTwo = "DAM_RuleAction_Two";

        /// <summary>
        /// Specify the value for the second PidTagRuleProvider of DAM, which can be set to a random string for testing such as "RuleOrganizerTwo".
        /// </summary>
        public const string DAMPidTagRuleProviderTwo = "RuleOrganizerTwo"; 
        #endregion

        #region RuleData used for DEM
        /// <summary>
        /// Specify the value for PidTagRuleUserFlags of DEM, which can be set to a random string for testing such as "1".
        /// </summary>
        public const string DEMPidTagRuleUserFlags = "1";

        /// <summary>
        /// Specify the value for PidTagRuleCondition of DEM, which can be set to a random string for testing such as "DEM Test".
        /// </summary>
        public const string DEMPidTagRuleConditionSubjectContainString = "DEM Test";

        /// <summary>
        /// Specify the value for PidTagRuleProviderData of DEM, which can be set to a random string for testing such as "01000000010000002222222270C1E340".
        /// </summary>
        public const string DEMPidTagRuleProviderData = "01000000010000002222222270C1E340";

        /// <summary>
        /// Specify the value for the PidTagRuleProvider of DEM, which can be set to a random string for testing such as "RuleOrganizerOne"
        /// </summary>
        public const string DEMPidTagRuleProvider = "RuleOrganizerOne"; 
        #endregion

        #region Rule Name
        /// <summary>
        /// Specify the value for the rule name of OP_MARK_AS_READ rule, which can be set to a random string for testing such as "MarkAsRead".
        /// </summary>
        public const string RuleNameMarkAsRead = "MarkAsRead";

        /// <summary>
        /// Specify the value for the rule name of OP_DELETE rule, which can be set to a random string for testing such as "Delete". 
        /// </summary>
        public const string RuleNameDelete = "Delete";

        /// <summary>
        /// Specify the value for the first rule name of OP_MOVE rule, which can be set to a random string for testing such as "MoveRuleOne".
        /// </summary>
        public const string RuleNameMoveOne = "MoveRuleOne";

        /// <summary>
        /// Specify the value for the second rule name of OP_MOVE rule, which can be set to a random string for testing such as "MoveRuleTwo".
        /// </summary>
        public const string RuleNameMoveTwo = "MoveRuleTwo";

        /// <summary>
        /// Specify the value for the rule name of OP_Copy rule, which can be set to a random string for testing such as "CopyRule".
        /// </summary>
        public const string RuleNameCopy = "CopyRule";

        /// <summary>
        /// Specify the value for the rule name of OP_FORWARD rule, which can be set to a random string for testing such as "ForwardRule".
        /// </summary>
        public const string RuleNameForward = "ForwardRule";

        /// <summary>
        /// Specify the value for the rule name of OP_DEFER_ACTION rule, which can be set to a random string for testing such as "DefferedRule". 
        /// </summary>
        public const string RuleNameDeferredAction = "DeferredRule";

        /// <summary>
        /// Specify the value for the rule name of OP_BOUNCE rule, which can be set to a random string for testing such as "BounceRule".
        /// </summary>
        public const string RuleNameBounce = "BounceRule";

        /// <summary>
        /// Specify the value for the rule name of OP_TAG rule, which can be set to a random string for testing such as "TagRule".
        /// </summary>
        public const string RuleNameTag = "TagRule";

        /// <summary>
        /// Specify the value for the rule name of OP_DELEGATE rule, which can be set to a random string for testing such as "DelegateRule".
        /// </summary>
        public const string RuleNameDelegate = "DelegateRule";

        /// <summary>
        /// Specify the value for the rule name of OP_REPLY rule, which can be set to a random string for testing such as "ReplyRule".
        /// </summary>
        public const string RuleNameReply = "ReplyRule";

        /// <summary>
        /// Specify the value for the rule name of OP_FORWARD rule with AT ActionFlavor flag, which can be set to a random string for testing such as "ForwardRuleAT".
        /// </summary>
        public const string RuleNameForwardAT = "ForwardRuleAT";

        /// <summary>
        /// Specify the value for the rule name of OP_FORWARD rule with TM ActionFlavor flag, which can be set to a random string for testing such as "ForwardRuleTM".
        /// </summary>
        public const string RuleNameForwardTM = "ForwardRuleTM";

        /// <summary>
        /// Specify the value for the rule name of OOF reply, which can be set to a random string for testing such as "OOFReplyRule".
        /// </summary>
        public const string RuleNameOOFReply = "OOFReplyRule";

        /// <summary>
        /// Specify the value for the PidTagRuleName of DEM, which can be set to a random string for testing such as "DAM_RuleTwo".
        /// </summary>
        public const string DEMRule = "DEMRule"; 
        #endregion

        #region Test data used in S01_AddModifyDeleteRetrieveRules
        /// <summary>
        /// An invalidate folder handle.
        /// </summary>
        public const string InvalidateFolderHandler = "4294967295";

        /// <summary>
        /// Specify the value for the DefferActionBufferData, 
        /// the default value on exchange is "01000000010000002222222270C1E34001000000010000002222222270C1E340".
        /// </summary>
        public const string DeferredActionBufferData = "01000000010000002222222270C1E34001000000010000002222222270C1E340";

        /// <summary>
        /// Specify a specific value for the first extend rule name, which can be set to a random string for testing such as "ExtendRulename1".
        /// </summary>
        public const string ExtendRulename1 = "ExtendRulename1";

        /// <summary>
        /// Specify a specific value for the second extend rule name, which can be set to a random string for testing such as "ExtendRulename2".
        /// </summary>
        public const string ExtendRulename2 = "ExtendRulename2";

        /// <summary>
        /// Specify a specific value for the third extend rule name, which can be set to a random string for testing such as "ExtendRulename3".
        /// </summary>
        public const string ExtendRulename3 = "ExtendRulename3";

        /// <summary>
        /// Specify the condition value for the first extend rule, which can be set to a random string for testing such as "extendedrule1".
        /// </summary>
        public const string ExtendRuleCondition1 = "extendedrule1";

        /// <summary>
        /// Specify the condition value for the second extend rule, which can be set to a random string for testing such as "extendedrule2".
        /// </summary>
        public const string ExtendRuleCondition2 = "extendedrule2";

        /// <summary>
        /// Specify the condition value for the third extend rule, which can be set to a random string for testing such as "extendedrule3".
        /// </summary>
        public const string ExtendRuleCondition3 = "extendedrule3";

        /// <summary>
        /// Specify the value for the folder display name, which can be set to a random string for testing such as "TestFolder".
        /// </summary>
        public const string FolderDisplayName = "TestFolder";

        /// <summary>
        /// Specify the value for the folder comment, which can be set to a random string for testing such as "Created for test!"
        /// </summary>
        public const string FolderComment = "Created for test!";

        /// <summary>
        /// Specify the value for the name of NamedProperty, which can be set to a random string for testing such as "Property". The length of the string value should be less than 10 on Exchange 2007 and 2010 for this test suite.
        /// </summary>
        public const string NameOfPropertyName = "Property";
        #endregion

        #region Test Data used in S05_GenerateDAMAndDEM
        /// <summary>
        /// Specify the value for client entry ID who download the message, 
        /// the default value on exchange is "0x00, 0x00, 0x00, 0xe6, 0xd1, 0x1b, 0x07,0x49, 0x83, 0x46, 0x88, 
        /// 0xd0, 0xc9, 0xa7, 0x51, 0x59, 0x8a, 0x42, 0x3e, 0xec, 0xc1".
        /// </summary>
        public const string ClientEntryId = "0x00, 0x00, 0x00, 0xe6, 0xd1, 0x1b, 0x07, 0x49, 0x83, 0x46, 0x88, 0xd0, 0xc9, 0xa7, 0x51, 0x59, 0x8a, 0x42, 0x3e, 0xec, 0xc1";

        /// <summary>
        /// Specify the value for invalidate entry ID, which can be set to a random string consist of hex16 values for testing such as "0x00, 0x01".
        /// </summary>
        public const string InvalidateEntryId = "0x00, 0x01";
        #endregion

        /// <summary>
        /// Size of FlatUID_r structure in byte.
        /// </summary>
        public const int FlatUIDByteSize = 16;

        /// <summary>
        /// The maximum number of rows for the NspiGetMatches method to return in a restricted address book container.
        /// </summary>
        public const uint GetMatchesRequestedRowNumber = 5000;

        /// <summary>
        /// The maximum number of rows for the NspiQueryRows method to return in a restricted address book container.
        /// </summary>
        public const uint QueryRowsRequestedRowNumber = 5000;

        /// <summary>
        /// A string which specifies a user name which doesn't exist.
        /// </summary>
        public const string UnresolvedName = "XXXXXX";

        /// <summary>
        /// A CodePage that server does not recognize (0xFFFFFFFF).
        /// </summary>
        public const string UnrecognizedCodePage = "4294967295";

        /// <summary>
        /// A Minimal Entry ID that server does not recognize (0xFFFFFFFF).
        /// </summary>
        public const string UnrecognizedMID = "4294967295";
        #endregion
    }
}