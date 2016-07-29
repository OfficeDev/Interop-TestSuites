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
    public class S01_AddModifyDeleteRetrieveRules : TestSuiteBase
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
        /// This test case is designed to test the operation adding OP_Mark_as_Read and OP_Delete rules to the server.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S01_TC01_AddMark_as_ReadAndDeleteRule()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region TestUser1 prepares value for rule properties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameMarkAsRead);
            #endregion

            #region TestUser1 creates a new folder in the Inbox folder.
            RopCreateFolderResponse createFolderResponse;
            string testFolderName = Common.GenerateResourceName(this.Site, Constants.FolderDisplayName);
            uint newFolderHandle = this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, testFolderName, Constants.FolderComment, out createFolderResponse);
            #endregion

            #region TestUser1 gets the value of PidTagHasRules of the newly created folder before adding a new rule.
            PropertyTag pidTagHasRules = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagHasRules,
                PropertyType = (ushort)PropertyType.PtypBoolean
            };
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = this.OxoruleAdapter.RopGetPropertiesSpecific(newFolderHandle, new PropertyTag[] { pidTagHasRules });
            Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse.ReturnValue, "Getting PidTagHasRules property should succeed.");

            // Use a variable to note whether the newly created folder has rules before adding the new rule to it.
            bool pidTagHasRulesBeforeAdd = false;
            if (getPropertiesSpecificResponse.RowData.PropertyValues != null && getPropertiesSpecificResponse.RowData.PropertyValues.Count > 0)
            {
                // The flag field set to 0x00 means all the queried property values are present and without error as specified in MS-OXCDATA section 2.8.1.1.
                if (getPropertiesSpecificResponse.RowData.Flag == 0x00)
                {
                    pidTagHasRulesBeforeAdd = BitConverter.ToBoolean(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value, 0);
                }
            }

            if (Common.IsRequirementEnabled(8851, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R8851");

                // Verify MS-OXORULE requirement: MS-OXORULE_8851.
                // There are no rules for the newly created folder before adding rule to it, so the pidtaghasrules property must be false
                Site.CaptureRequirementIfIsFalse(
                    pidTagHasRulesBeforeAdd,
                    8851,
                    @"[[In Appendix A: Product Behavior] Implementation does set PidTagHasRules to ""FALSE"" if no rule (2) is set on a folder. (Exchange 2007, Exchange 2010 and Exchange 2016 follow this behavior.)");
            }
            #endregion

            #region TestUser1 generates test RuleData.
            RuleData ruleForMarkRead = AdapterHelper.GenerateValidRuleData(ActionType.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, new DeleteMarkReadActionData(), ruleProperties, null);
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameDelete);

            // Add rule for delete without rule Provider Data.
            ruleProperties.ProviderData = string.Empty;
            RuleData ruleForDelete = AdapterHelper.GenerateValidRuleData(ActionType.OP_DELETE, TestRuleDataType.ForAdd, 0, RuleState.ST_ENABLED, new DeleteMarkReadActionData(), ruleProperties, null);
            #endregion

            #region TestUser1 adds OP_MARK_AS_READ and OP_DELETE rules to the newly created folder.
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(newFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForDelete, ruleForMarkRead });

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R596");

            // Verify MS-OXORULE requirement: MS-OXORULE_R596.
            // If the return value of the RopModifyRules ROP is 0x00000000, it means server parses the request successfully, and this requirement can be verified.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                modifyRulesResponse.ReturnValue,
                596,
                @"[In Receiving a RopModifyRules ROP Request] If the server successfully parses the data in the request buffer and is able to process all requests for adding, modifying, and deleting rules (2) present in the request buffer, the server MUST return 0x00000000 as the value of the ReturnValue field in the response buffer.");
            #endregion

            #region TestUser1 gets the value of PidTagHasRules of the newly created folder after adding rules.
            if (Common.IsRequirementEnabled(7202, this.Site))
            {
                getPropertiesSpecificResponse = this.OxoruleAdapter.RopGetPropertiesSpecific(newFolderHandle, new PropertyTag[] { pidTagHasRules });
                Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse.ReturnValue, "Getting PidTagHasRules property should succeed.");

                // Use a variable to note whether the newly created folder has rules after add two rules to it.
                bool pidTagHasRulesAfterAdd = false;
                if (getPropertiesSpecificResponse.RowData.PropertyValues != null && getPropertiesSpecificResponse.RowData.PropertyValues.Count > 0)
                {
                    pidTagHasRulesAfterAdd = BitConverter.ToBoolean(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value, 0);
                }

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R7202");
                Site.CaptureRequirementIfIsTrue(
                    pidTagHasRulesAfterAdd,
                    7202,
                    @"[[In Appendix A: Product Behavior] Implementation does set PidTagHasRules to ""TRUE"" if any rules (2) are set on a folder. (Exchange 2007, Exchange 2010 and Exchange 2016 follow this behavior.)");
            }
            #endregion

            #region TestUser1 calls RopGetRulesTable to get the rules table.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(newFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "RopGetRulesTable operation should success.");

            #region Capture Code

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R133");

            // Verify MS-OXORULE requirement: MS-OXORULE_R133.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                ropGetRulesTableResponse.ReturnValue,
                133,
                @"[In RopGetRulesTable ROP Response Buffer] ReturnValue (4 bytes): To indicate success, the server returns 0x00000000.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R611");

            // Verify MS-OXORULE requirement: MS-OXORULE_R611.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                ropGetRulesTableResponse.ReturnValue,
                611,
                @"[In Receiving a RopGetRulesTable ROP Request] If the server successfully parses the data in the ROP request buffer, it MUST return 0x00000000 as the value of the ReturnValue field in the response buffer.");
            #endregion
            #endregion

            #region TestUser1 retrieves rule information for the newly added rule.
            PropertyTag[] propertyTags = this.GenerateRuleInfoProperties();

            // Retrieves rows from the rule table.
            RopQueryRowsResponse queryRowResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyTags);
            Site.Assert.AreEqual<uint>(0, queryRowResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R612, the return table handle actual is:{0}", ruleTableHandle.ToString());

            // Verify MS-OXORULE requirement: MS-OXORULE_R612.
            // There are only two rules on the newly created folder added by the test suite.
            // If the value of the RowCount field is 2, it means that when using the table handle returned from the server to get the rules, the rules can be got.
            // So the returned table handle is a valid handle, and this requirement can be verified.
            Site.CaptureRequirementIfAreEqual<uint>(
                2,
                queryRowResponse.RowCount,
                612,
                @"[In Receiving a RopGetRulesTable ROP Request] If the server successfully parses the data in the ROP request buffer, it MUST return a valid table handle through which the client can access the folder rules (2) using table specific ROPs defined in [MS-OXCTABL].");

            // Get PidTagRuleName's value from the specific RuleData.
            TaggedPropertyValue ruleNameProperty = new TaggedPropertyValue();
            foreach (PropertyValue value in ruleForMarkRead.PropertyValues)
            {
                if (((TaggedPropertyValue)value).PropertyTag.PropertyId == (ushort)PropertyId.PidTagRuleName)
                {
                    ruleNameProperty = (TaggedPropertyValue)value;
                    break;
                }
            }

            // Get the property row of the specific rule.
            PropertyRow propertyRow = null;
            if (queryRowResponse.RowData.PropertyRows != null)
            {
                foreach (PropertyRow propRow in queryRowResponse.RowData.PropertyRows)
                {
                    if (Common.CompareByteArray(propRow.PropertyValues[0].Value, ruleNameProperty.Value))
                    {
                        propertyRow = propRow;
                        break;
                    }
                }
            }

            // Get the rule's rule ID.
            byte[] byteRuleIDValue = this.GetPropertyFromList(PropertyId.PidTagRuleId, propertyRow, propertyTags);
            ulong ruleIDForMarkRead = BitConverter.ToUInt64(byteRuleIDValue, 0);

            // Get the rule's Provider data.
            byte[] byteRuleProviderData = this.GetPropertyFromList(PropertyId.PidTagRuleProviderData, propertyRow, propertyTags);

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R100");

            // Verify MS-OXORULE requirement: MS-OXORULE_R100.
            // If the property value of PidTagRuleProviderData is equal to the value saved in Util.cs file, it means that the server has preserved the RuleProviderData.
            bool isVerifyR100 = Common.GetStringFromBinary(byteRuleProviderData, true).Equals(Constants.PidTagRuleProviderData, StringComparison.CurrentCultureIgnoreCase);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR100,
                100,
                @"[In PidTagRuleProviderData Property] The server is to preserve this value [is contained by PidTagRuleUserFlags Property] if set by the client.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R122.");

            // Verify MS-OXORULE requirement: MS-OXORULE_R122.
            // queryRowResponse will return all the rules in the table. Here will return two rules that has been added above.
            // So if the value of RowCount is 0x0002, it means the server has returned all the standard rules associated with the given folder.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x0002,
                queryRowResponse.RowCount,
                122,
                @"[In RopGetRulesTable ROP] The table returned by the server is required to contain all standard rules associated with a given folder.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R123, the RowCount is {0}", queryRowResponse.RowData.Count);

            // Verify MS-OXORULE requirement: MS-OXORULE_R123.
            // If the value of RowData in queryRowResponse is not null and the count of the RowData is equal to RowCount, it means each row represents one rule.
            bool isVerifyR123 = queryRowResponse.RowData != null && queryRowResponse.RowData.Count == queryRowResponse.RowCount;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR123,
                123,
                @"[In RopGetRulesTable ROP] Each row in the table MUST represent one rule (2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R597, the rule id is:{0}", ruleIDForMarkRead.ToString());

            // Verify MS-OXORULE requirement: MS-OXORULE_R597.
            bool isVerifyR597 = ruleIDForMarkRead != 0x0000000000000000;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR597,
                597,
                @"[In Receiving a RopModifyRules ROP Request] The server MUST assign a value for the PidTagRuleId property (section 2.2.7.8) for each rule (2) that has been added by the RopModifyRules ROP request.");

            byte[] valueRuleUserFlagsRetrieved = this.GetPropertyFromList(PropertyId.PidTagRuleUserFlags, propertyRow, propertyTags);

            // BitConverter.ToInt32() is used to convert a byte array to a int value from the byte array index of 0.
            int ruleUserFlagsRetrieved = BitConverter.ToInt32(valueRuleUserFlagsRetrieved, 0);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R96 , the rule flag set is :{0}, the rule flag get is:{1}", ruleProperties.UserFlag, Convert.ToString(ruleUserFlagsRetrieved));

            // Verify MS-OXORULE requirement: MS-OXORULE_R96.
            // UserFlag represents the PidTagRuleUserFlags set by the client, and ruleUserFlagsRetrieved represents 
            // the PidTagRuleUserFlags set by the server. If they are equal to each other, it means the server preserves this value. 
            Site.CaptureRequirementIfAreEqual<string>(
                ruleProperties.UserFlag,
                Convert.ToString(ruleUserFlagsRetrieved),
                96,
                @"[In PidTagRuleUserFlags Property] The server is to preserve this value [is contained by PidTagRuleUserFlags Property] if set by the client.");
            #endregion
            #endregion

            #region TestUser1 gets ruleActions of the two rules.
            ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(newFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");

            PropertyTag actionsTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleActions,
                PropertyType = (ushort)PropertyType.PtypRuleAction
            };

            // Set the query target to standardardRules.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForStandardRules;

            RopQueryRowsResponse getAllActionsResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, new PropertyTag[1] { actionsTag });
            Site.Assert.AreEqual<uint>(0, getAllActionsResponse.ReturnValue, "Getting the rule actions should succeed.");

            // Two rules have been added to the newly created folder, so the row count in the rule table should be 2.
            Site.Assert.AreEqual<uint>(2, getAllActionsResponse.RowCount, "The rule number in the rule table is {0}", getAllActionsResponse.RowCount);
            this.VerifyRuleTable();
            #endregion

            #region Delete the created folder and clear the status of the adapter.
            RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, createFolderResponse.FolderId);
            Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");

            // Clear the status of the adapter.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.OtherTarget;
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the standard rules created with Unused Flags. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S01_TC02_AddNewStandardRule_InvalidRequestBuffer()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region TestUser1 prepares value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameMarkAsRead);
            #endregion

            #region Generate test RuleData.
            RuleData ruleForMarkRead = AdapterHelper.GenerateValidRuleData(ActionType.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, new DeleteMarkReadActionData(), ruleProperties, null);
            #endregion

            #region TestUser1 gets the returned value from RopModifyRules response.
            RopModifyRulesResponse responseinValidModifyRulesFlag = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_Unused, new RuleData[] { ruleForMarkRead });

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R821");

            // Verify MS-OXORULE requirement: MS-OXORULE_R821.
            // 0x80070057 means error code ecInvalidParam.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80070057,
                responseinValidModifyRulesFlag.ReturnValue,
                821,
                @"[In Receiving a RopModifyRules ROP Request] The value of error code ecInvalidParam: 0x80070057.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R822");

            // Verify MS-OXORULE requirement: MS-OXORULE_R822.
            // All the x bits in the ModifyRulesFlag field of the ROP request is not set to 0, so if error code ecInvalidParam (0x80070057) is returned, R822 can be verified.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80070057,
                responseinValidModifyRulesFlag.ReturnValue,
                822,
                @"[In Receiving a RopModifyRules ROP Request] The meaning of error code ecInvalidParam: One or more of the x bits in the ModifyRulesFlag field of the ROP request is not set to 0.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R904");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R904.
            // If the returnValue of RopModifyRules response is 0x80070057, it means the server returns the error code InvalidParameter.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80070057,
                responseinValidModifyRulesFlag.ReturnValue,
                "MS-OXCDATA",
                 904,
                @"[In Error Codes] The numeric value (hex) for error code InvalidParameter is 0x80070057, %x57.00.07.80.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to test the operation adding extended rules with large message condition to Inbox folder.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S01_TC03_AddExtendedRule_WithLargeMessageCondition()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region TestUser1 creates an FAI message.
            RopCreateMessageResponse ropCreateMessageResponse;
            uint extendedRuleMessageHandle = this.OxoruleAdapter.RopCreateMessage(this.InboxFolderHandle, this.InboxFolderID, Convert.ToByte(true), out ropCreateMessageResponse);
            Site.Assert.AreEqual<uint>(0, ropCreateMessageResponse.ReturnValue, "Creating the first FAI message should succeed.");
            #endregion

            #region TestUser1 adds the extended rule with NamedProperty successfully.
            string ruleConditionSubjectName = Constants.RuleConditionSubjectContainString;
            NamedPropertyInfo namedPropertyInfo = new NamedPropertyInfo
            {
                NoOfNamedProps = 2,
                PropId = new uint[2]
                {
                    0x8001, 0x8002
                }
            };

            // 0x01 means the property is identified by the name property.
            PropertyName testPropertyName = new PropertyName
            {
                Guid = System.Guid.NewGuid().ToByteArray(),
                Kind = 0x01,
                Name = Encoding.Unicode.GetBytes(Constants.NameOfPropertyName + "\0")
            };
            testPropertyName.NameSize = (byte)testPropertyName.Name.Length;

            // 0x00 means the property is identified by the LID.
            PropertyName secondPropertyName = new PropertyName
            {
                Guid = System.Guid.NewGuid().ToByteArray(),
                Kind = 0x00,
                LID = 0x88888888
            };
            namedPropertyInfo.NamedProperty = new PropertyName[2] { testPropertyName, secondPropertyName };
            namedPropertyInfo.NamedPropertiesSize = (uint)(testPropertyName.Serialize().Length + secondPropertyName.Serialize().Length);
            string ruleName = Common.GenerateResourceName(this.Site, Constants.ExtendRulename1);
            TaggedPropertyValue[] extendedRuleProperties = AdapterHelper.GenerateExtendedRuleTestData(ruleName, 0, (uint)RuleState.ST_ENABLED, Constants.PidTagRuleProvider, ActionType.OP_MARK_AS_READ, new DeleteMarkReadActionData(), ruleConditionSubjectName, namedPropertyInfo);

            // Set properties for extended rule FAI message.
            RopSetPropertiesResponse ropSetPropertiesResponse = this.OxoruleAdapter.RopSetProperties(extendedRuleMessageHandle, extendedRuleProperties);
            Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "Setting property for Extended rule FAI message should succeed.");

            // Save changes of message.
            RopSaveChangesMessageResponse ropSaveChangesMessagResponse = this.OxoruleAdapter.RopSaveChangesMessage(extendedRuleMessageHandle);
            Site.Assert.AreEqual<uint>(0, ropSaveChangesMessagResponse.ReturnValue, "Saving Extend rule message should succeed.");

            // Specify the properties to be got. 
            PropertyTag[] propertyTagArray = new PropertyTag[2];

            // PidTagRuleMessageProviderData
            propertyTagArray[0].PropertyId = (ushort)PropertyId.PidTagRuleMessageProvider;
            propertyTagArray[0].PropertyType = (ushort)PropertyType.PtypString;

            // PidTagRuleMessageName
            propertyTagArray[1].PropertyId = (ushort)PropertyId.PidTagRuleMessageName;
            propertyTagArray[1].PropertyType = (ushort)PropertyType.PtypString;

            // Get the specific properties of the extended rule.
            RopGetPropertiesSpecificResponse ropGetPropertiesSpecificResponse = this.OxoruleAdapter.RopGetPropertiesSpecific(extendedRuleMessageHandle, propertyTagArray);
            Site.Assert.AreEqual<uint>(0, ropGetPropertiesSpecificResponse.ReturnValue, "Getting folder id property operation should succeed.");

            #region Capture Code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R184");

            // Verify MS-OXORULE requirement: MS-OXORULE_R184
            string pidTagRuleMessageProviderData = AdapterHelper.PropertyValueConvertToString(ropGetPropertiesSpecificResponse.RowData.PropertyValues[0].Value);
            Site.CaptureRequirementIfAreEqual<string>(
                Constants.PidTagRuleProvider,
                pidTagRuleMessageProviderData,
                184,
                @"[In PidTagRuleMessageProvider Property] This property has the same semantics as the PidTagRuleProvider property (section 2.2.1.3.1.5). [The PidTagRuleMessageProvider property identifies the client application that owns the rule (2).]");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R164");

            // Verify MS-OXORULE requirement: MS-OXORULE_R164
            // The rule name set by client is "ExtendRulename1".
            string pidTagRuleMessageNameValue = AdapterHelper.PropertyValueConvertToString(ropGetPropertiesSpecificResponse.RowData.PropertyValues[1].Value);
            Site.CaptureRequirementIfAreEqual<string>(
                 ruleName,
                 pidTagRuleMessageNameValue,
                 164,
                 @"[In PidTagRuleMessageName Property] This property has the same semantics as the PidTagRuleName property (section 2.2.1.3.1.4). [The PidTagRuleMessageName property specifies the name of the rule (2).]");
            #endregion
            #endregion

            #region TestUser1 retrieves data of the new extended rule.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForExtendedRules;
            RopGetPropertiesAllResponse ropGetExtendRuleMessageResponse = this.OxoruleAdapter.RopGetPropertiesAll(extendedRuleMessageHandle, this.PropertySizeLimitFlag, (ushort)WantUnicode.Want);
            Site.Assert.AreEqual<uint>(0, ropGetExtendRuleMessageResponse.ReturnValue, "Getting all properties operation should succeed.");
            Site.Assert.IsTrue(ropGetExtendRuleMessageResponse.PropertyValues.Length != 0, "Extended Rule data should be found in related FAI message!");
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.OtherTarget;

            ExtendedRuleActions extendedRuleMessageActions = null;

            // Check the properties set on Extended Rule, and find the Extended Rule Actions.
            for (int i = 0; i < ropGetExtendRuleMessageResponse.PropertyValues.Length; i++)
            {
                // propertyId indicates the Id of a property set on Extended Rule.
                ushort propertyId = ropGetExtendRuleMessageResponse.PropertyValues[i].PropertyTag.PropertyId;
                if (propertyId == (ushort)PropertyId.PidTagExtendedRuleMessageActions)
                {
                    byte[] propertyValue = ropGetExtendRuleMessageResponse.PropertyValues[i].Value;
                    extendedRuleMessageActions = AdapterHelper.PropertyValueConvertToExtendedRuleActions(propertyValue);
                    break;
                }
            }

            Site.Assert.AreNotEqual<ExtendedRuleActions>(null, extendedRuleMessageActions, "extendedRuleMessageActions should not be null.");

            // Get the Property Names saved by server in the extendedRuleMessageActions.
            PropertyName[] propertyNames = extendedRuleMessageActions.NamedPropertyInformation.NamedProperty;
            PropertyName testPropertyNameSavedOnServer = new PropertyName();
            PropertyName secondPropertyNameSavedOnServer = new PropertyName();
            if (propertyNames != null && propertyNames.Length > 0)
            {
                for (int i = 0; i < propertyNames.Length; i++)
                {
                    // If the Kind is 0x01, it means this PropertyName is the testPropertyName.
                    if (propertyNames[i].Kind == 0x01)
                    {
                        testPropertyNameSavedOnServer = propertyNames[i];
                    }
                    else if (propertyNames[i].Kind == 0x00)
                    {
                        // If the Kind is 0x00, it means this PropertyName is the secondPropertyName.
                        secondPropertyNameSavedOnServer = propertyNames[i];
                    }
                }
            }

            Site.Assert.IsNotNull(testPropertyNameSavedOnServer.Guid, "testPropertyNameSavedOnServer should not be null.");
            Site.Assert.IsNotNull(secondPropertyNameSavedOnServer.Guid, "secondPropertyNameSavedOnServer should not be null.");
            #endregion

            #region TestUser1 creates a new FAI message.
            extendedRuleMessageHandle = this.OxoruleAdapter.RopCreateMessage(this.InboxFolderHandle, this.InboxFolderID, Convert.ToByte(true), out ropCreateMessageResponse);
            Site.Assert.AreEqual<uint>(0, ropCreateMessageResponse.ReturnValue, "Creating the second FAI message should succeed.");
            #endregion

            #region TestUser1 creates an extended rule with PidTagExtendedRuleMessageCondition property value size larger than the size set by the server.
            propertyTagArray = new PropertyTag[1];
            propertyTagArray[0].PropertyId = (ushort)PropertyId.PidTagExtendedRuleSizeLimit;
            propertyTagArray[0].PropertyType = (ushort)PropertyType.PtypInteger32;

            ropGetPropertiesSpecificResponse = this.OxoruleAdapter.RopGetPropertiesSpecific(this.LogonHandle, propertyTagArray);
            uint pidTagExtendedRuleSizeLimit = Common.ConvertByteArrayToUint(ropGetPropertiesSpecificResponse.RowData.PropertyValues[0].Value);
            
            // According to MS-OXCRPC, "The server SHOULD fail with the RPC status code of RPC_X_BAD_STUB_DATA (0x000006F7) if the request buffer is larger than 0x00040000 bytes in size."
            if (pidTagExtendedRuleSizeLimit < 0x00040000)
            {
                ruleConditionSubjectName = Constants.RuleConditionSubjectContainString;
                namedPropertyInfo = new NamedPropertyInfo
                {
                    NoOfNamedProps = 2,
                    PropId = new uint[2]
                    {
                        0x8001, 0x8002
                    }
                };

                // Generate a string value whose size lager than the one specified by the PidTagExtendedRuleSizeLimit property.
                StringBuilder stringByteValue = new StringBuilder("ExtentRuleSize");
                stringByteValue.Append('a', (int)pidTagExtendedRuleSizeLimit);

                // If the value of Kind is 0x01, it means that the property is identified by the name property.
                testPropertyName = new PropertyName
                {
                    Guid = System.Guid.NewGuid().ToByteArray(),
                    Kind = 0x01,
                    Name = Encoding.Unicode.GetBytes(stringByteValue + "\0")
                };

                testPropertyName.NameSize = (byte)testPropertyName.Name.Length;

                // If the value of Kind is 0x00, it means that the property is identified by the LID.
                secondPropertyName = new PropertyName
                {
                    Guid = System.Guid.NewGuid().ToByteArray(),
                    Kind = 0x00,
                    LID = 0x88888888
                };

                namedPropertyInfo.NamedProperty = new PropertyName[2] { testPropertyName, secondPropertyName };
                namedPropertyInfo.NamedPropertiesSize = (uint)(testPropertyName.Serialize().Length + secondPropertyName.Serialize().Length);
                extendedRuleProperties = AdapterHelper.GenerateExtendedRuleTestData(Common.GenerateResourceName(this.Site, Constants.ExtendRulename1), 0, (uint)RuleState.ST_ENABLED, Constants.PidTagRuleProvider, ActionType.OP_MARK_AS_READ, new DeleteMarkReadActionData(), ruleConditionSubjectName, namedPropertyInfo);
                TaggedPropertyValue pidTagExtendedRuleMessageCondition = new TaggedPropertyValue();
                foreach (TaggedPropertyValue propertyValue in extendedRuleProperties)
                {
                    if (propertyValue.PropertyTag.PropertyId == (ushort)PropertyId.PidTagExtendedRuleMessageCondition)
                    {
                        pidTagExtendedRuleMessageCondition = propertyValue;
                        break;
                    }
                }

                uint pidTagExtendedRuleMessageConditionSize = uint.Parse(pidTagExtendedRuleMessageCondition.Value.Length.ToString());

                // Set properties for extended rule FAI message.
                ropSetPropertiesResponse = this.OxoruleAdapter.RopSetProperties(extendedRuleMessageHandle, extendedRuleProperties);

                // Save changes of message.
                ropSaveChangesMessagResponse = this.OxoruleAdapter.RopSaveChangesMessage(extendedRuleMessageHandle);

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R1016");

                // Verify MS-OXORULE requirement: MS-OXORULE_R1016.
                if (pidTagExtendedRuleMessageConditionSize > pidTagExtendedRuleSizeLimit)
                {
                    Site.CaptureRequirementIfAreNotEqual<uint>(
                        0x0000,
                        ropSaveChangesMessagResponse.ReturnValue,
                        1016,
                        @"[In Processing Incoming Messages to a Folder] If the PidTagExtendedRuleSizeLimit property is set and the size of the PidTagExtendedRuleMessageCondition property (section 2.2.4.1.10) exceeds the value specified by the PidTagExtendedRuleSizeLimit property, the server MUST return an error.");
                }
                else
                {
                    Site.Assert.Fail("The size of the PidTagExtendedRuleMessageCondition property should exceeds the value specified by the PidTagExtendedRuleSizeLimit property.");
                }
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the standard rules created with Unused Flags. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S01_TC04_AddStandardRuleToVerifyUnusedFlag()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region TestUser1 prepares value for the rule Data.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameMarkAsRead);
            #endregion

            #region Generate test RuleData.
            RuleData ruleForMarkRead = AdapterHelper.GenerateValidRuleData(ActionType.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, new DeleteMarkReadActionData(), ruleProperties, null);

            // Rule data with PidTagRuleState's x bit set and not set PidTagRuleUserFlags property.
            RuleData ruleWithXRuleStateFlag = AdapterHelper.GenerateValidRuleData(ActionType.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 1, RuleState.X, new DeleteMarkReadActionData(), ruleProperties, null);

            // Rule data with PidTagRuleState's ER bit set and not set PidTagRuleUserFlags property.
            RuleData ruleWithERRuleStateFlag = AdapterHelper.GenerateValidRuleData(ActionType.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 1, RuleState.ST_ERROR, new DeleteMarkReadActionData(), ruleProperties, null);

            // Rule data with PidTagRuleState's PE bit set and not set PidTagRuleUserFlags property.
            RuleData ruleWithPERuleStateFlag = AdapterHelper.GenerateValidRuleData(ActionType.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 1, RuleState.ST_RULE_PARSE_ERROR, new DeleteMarkReadActionData(), ruleProperties, null);

            ruleProperties.UserFlag = string.Empty;

            // Rule data without UserFlag.
            RuleData ruleWithOutUserFlag = AdapterHelper.GenerateValidRuleData(ActionType.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, new DeleteMarkReadActionData(), ruleProperties, null);
            #endregion

            #region Testuser1 adds a rule with all property set to valid values to the Inbox folder.
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForMarkRead });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding rules should succeed.");
            #endregion

            #region TestUser1 calls RopGetRulesTable with invalid TableFlags and server returns an error.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Invalid, out ropGetRulesTableResponse);

            #region Capture Code
            if (Common.IsRequirementEnabled(836, this.Site))
            {
                // GetRulesTable returns 0x0(Success).
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R836");

                // Verify MS-OXORULE requirement: MS-OXORULE_R836
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x0,
                    ropGetRulesTableResponse.ReturnValue,
                    836,
                    @"[In Appendix A: Product Behavior] Implementation does ignore the x bits and returns ecSuccess in this case [One or more of the x bits on the TableFlags field of RopGetRulesTable ROP Request is set to a nonzero value.]. [<23> Section 3.2.5.3: Exchange 2007 ignores the x bits and returns ecSuccess in this case.]");
            }

            if (Common.IsRequirementEnabled(920, this.Site))
            {
                // GetRulesTable returns 0x80040102(ecNotSupported).
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R920");

                // Verify MS-OXORULE requirement: MS-OXORULE_R920.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040102,
                    ropGetRulesTableResponse.ReturnValue,
                    920,
                    @"[In Receiving a RopGetRulesTable ROP Request] Implementation does return the error code ecNotSupported if one or more of the x bits on the TableFlags field is set to a nonzero value. (Exchange 2003, Exchange 2010, and above follow this behavior.)");
            }
            #endregion
            #endregion

            #region TestUser1 adds a rule to Inbox folder with unused ModifyRulesFlag.
            RopModifyRulesResponse modifyRulesResponseWithUnusedModifyRulesFlag = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_Unused, new RuleData[] { ruleForMarkRead });
            Site.Assert.AreEqual<uint>(0x80070057, modifyRulesResponseWithUnusedModifyRulesFlag.ReturnValue, "Adding rules should return invalidParameter.");
            #endregion

            #region TestUser1 adds a rule to Inbox folder with X rulestate set.
            RopModifyRulesResponse modifyRulesResponseWithXRuleState = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleWithXRuleStateFlag });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponseWithXRuleState.ReturnValue, "Adding rules should succeed.");
            #endregion

            #region TestUser1 adds a rule to Inbox folder with ER rulestate set.
            RopModifyRulesResponse modifyRulesResponseWithERRuleState = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleWithERRuleStateFlag });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponseWithERRuleState.ReturnValue, "Adding rules should succeed.");
            #endregion

            #region TestUser1 adds a rule to Inbox folder with PE rulestate set.
            RopModifyRulesResponse modifyRulesResponseWithPERuleState = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleWithPERuleStateFlag });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponseWithPERuleState.ReturnValue, "Adding rules should succeed.");
            #endregion

            #region TestUser1 adds a rule to Inbox folder without setting rule userflag.
            RopModifyRulesResponse modifyRulesResponseWithOutUserFlag = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleWithOutUserFlag });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponseWithOutUserFlag.ReturnValue, "Adding rules should succeed.");

            #region Capture Code
            if (Common.IsRequirementEnabled(888, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R888: the return value of RopMidfyRules when the x bit set on the RuleState field is {0}", modifyRulesResponseWithXRuleState.ReturnValue);

                // Verify MS-OXORULE requirement: MS-OXORULE_R888.
                // If the modifyRulesResponse with X RuleState set is equals to the response without X RuleState set, it means whether the RuleState's x bit is set, the response is the same. 
                bool isVerifyR888 = modifyRulesResponseWithXRuleState.ReturnValue == modifyRulesResponse.ReturnValue;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR888,
                    888,
                    @"[In PidTagRuleState Property] x: The RopModifyRules ROP replies the same response whether the bit locations marked with x are set to 0 or 1 on the implementation. (Exchange 2003, Exchange 2010 and above follow this behavior)");
            }

            // Add the debug information.
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R844: the return value of RopMidfyRules when the ER flag set on the RuleState field is {0}", modifyRulesResponseWithERRuleState.ReturnValue);

            // Verify MS-OXORULE requirement: MS-OXORULE_R844.
            bool isVerifiedR844 = modifyRulesResponseWithERRuleState.ReturnValue == modifyRulesResponse.ReturnValue;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR844,
                844,
                @"[In PidTagRuleState Property] The RopModifyRules ROP replies the same response whether this Bitmask ER is set to 0 or 1.");

            // Add the debug information.
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R846: the return value of RopMidfyRules when the PE flag set on the RuleState field is {0}", modifyRulesResponseWithPERuleState.ReturnValue);

            // Verify MS-OXORULE requirement: MS-OXORULE_R846.
            bool isVerifiedR846 = modifyRulesResponseWithPERuleState.ReturnValue == modifyRulesResponse.ReturnValue;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR846,
                846,
                @"[In PidTagRuleState Property]  PE (ST_RULE_PARSE_ERROR, Bitmask 0x00000040): The RopModifyRules ROP replies the same response whether this flag is set or not.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify modify and delete rules. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S01_TC05_ModifyOrDeleteRule()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameForward);
            #endregion

            #region TestUser1 prepares the recipient block for Forward rules.
            RecipientBlock recipientBlock = new RecipientBlock
            {
                Reserved = 0x01,
                NoOfProperties = (ushort)0x04u
            };

            TaggedPropertyValue[] recipientProperties = AdapterHelper.GenerateRecipientPropertiesBlock(this.User2Name, this.User2ESSDN);
            recipientBlock.PropertiesData = recipientProperties;
            #endregion

            #region TestUser1 adds rule forward.
            ForwardDelegateActionData forwardActionData = new ForwardDelegateActionData
            {
                RecipientCount = (ushort)0x01,
                RecipientsData = new RecipientBlock[1]
                {
                    recipientBlock
                }
            };

            RuleData ruleForward = AdapterHelper.GenerateValidRuleData(ActionType.OP_FORWARD, TestRuleDataType.ForAdd, 0, RuleState.ST_ENABLED, forwardActionData, ruleProperties, null);
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForward });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Forward rule should be success");
            #endregion

            #region TestUser1 calls RopGetRulesTable with valid TableFlags.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");
            #endregion

            #region TestUser1 retrieves rule information for the newly added rule.
            PropertyTag[] propertyTags = new PropertyTag[2];
            propertyTags[0].PropertyId = (ushort)PropertyId.PidTagRuleName;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTags[1].PropertyId = (ushort)PropertyId.PidTagRuleId;
            propertyTags[1].PropertyType = (ushort)PropertyType.PtypInteger64;

            // Retrieve rows from the rule table.
            RopQueryRowsResponse queryRowResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyTags);
            Site.Assert.AreEqual<uint>(0, queryRowResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");
            
            // Add one rules to the Inbox folder. If the rule table is got successfully and the rule count is 1,
            // it means that the server is returning a table with the rule added by the test suite.
            Site.Assert.IsTrue(queryRowResponse.RowCount == 1, "The rule number in the rule table is {0}", queryRowResponse.RowCount);
            ulong ruleId = BitConverter.ToUInt64(queryRowResponse.RowData.PropertyRows[0].PropertyValues[1].Value, 0);
            this.VerifyRuleTable();

            #region   Capture R127
            bool isVerifiedR127 = Encoding.Unicode.GetString(queryRowResponse.RowData.PropertyRows[0].PropertyValues[0].Value).Equals(ruleProperties.Name + "\0");
  
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R127");

            // Verify MS-OXORULE requirement: MS-OXORULE_R127
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR127,
                127,
                @"[In RopGetRulesTable ROP Request Buffer] [TableFlags] U: (Bitmask 0x40): This bit is set if the client is requesting that string values in the table be returned as Unicode strings.");

            #endregion
            #endregion

            #region TestUser1 prepares the recipient block for Forward rules.
            recipientBlock = new RecipientBlock
            {
                Reserved = 0x01,
                NoOfProperties = (ushort)0x04u
            };

            recipientProperties = AdapterHelper.GenerateRecipientPropertiesBlock(this.User1Name, this.User1ESSDN);
            recipientBlock.PropertiesData = recipientProperties;
            #endregion

            #region TestUser1 modifies the created rule.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameForward);
            forwardActionData = new ForwardDelegateActionData
            {
                RecipientCount = (ushort)0x01,
                RecipientsData = new RecipientBlock[1]
                {
                    recipientBlock
                }
            };

            ruleForward = AdapterHelper.GenerateValidRuleData(ActionType.OP_FORWARD, TestRuleDataType.ForModify, 0, RuleState.ST_ENABLED, forwardActionData, ruleProperties, ruleId);
            modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleForward });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Forward rule should be success");
            #endregion

            #region TestUser1 calls RopGetRulesTable with valid TableFlags.
            ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");
            #endregion

            #region TestUser1 retrieves rule information for the modified rule.
            // Retrieves rows from the rule table.
            queryRowResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyTags);
            Site.Assert.AreEqual<uint>(0, queryRowResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");

            // Only one rule added in the Inbox folder. If the rule table is got successfully and the rule count is 1,
            // it means that the server is returning a table with the rule added by the test suite.
            Site.Assert.AreEqual<uint>(1, queryRowResponse.RowCount, "The rule number in the rule table is {0}", queryRowResponse.RowCount);
            this.VerifyRuleTable();
            
            ulong ruleIdModified = BitConverter.ToUInt64(queryRowResponse.RowData.PropertyRows[0].PropertyValues[1].Value, 0);
            bool isSameRuleId = ruleId == ruleIdModified;
            Site.Assert.IsTrue(isSameRuleId, "The original rule Id is {0} and the modified rule Id is {1}", ruleId, ruleIdModified);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R671");

            bool isVerifiedR671 = isSameRuleId && queryRowResponse.RowCount == 1;

            // Verify MS-OXORULE requirement: MS-OXORULE_R671.
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR671,
                671,
                @"[In RopModifyRules ROP Request Buffer] [ModifyRulesFlag] R (Bitmask 0x01): If this bit is not set, the rules (2) specified in this request represent changes (delete, modify, and add) to the set of rules (2) already existing in this folder.");

            #endregion

            #region TestUser1 adds rule DELEGATE.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameDelegate, 1);
            ForwardDelegateActionData delegateActionData = new ForwardDelegateActionData
            {
                RecipientCount = (ushort)0x01
            };

            #region Prepare the Delegate rule Recipient block
            RecipientBlock delegateRecipientBlock = new RecipientBlock
            {
                Reserved = 0x01,
                NoOfProperties = (ushort)0x05u
            };
            TaggedPropertyValue[] delegateRecipientProperties = new TaggedPropertyValue[5];

            TaggedPropertyValue[] tempArray = AdapterHelper.GenerateRecipientPropertiesBlock(this.User2Name, this.User2ESSDN);
            Array.Copy(tempArray, 0, delegateRecipientProperties, 0, tempArray.Length);

            // Add PidTagSmtpEmailAdderss.
            delegateRecipientProperties[4] = new TaggedPropertyValue();
            PropertyTag delegateRecipientPropertiesPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagSmtpAddress,
                PropertyType = (ushort)PropertyType.PtypString
            };
            delegateRecipientProperties[4].PropertyTag = delegateRecipientPropertiesPropertyTag;
            delegateRecipientProperties[4].Value = Encoding.Unicode.GetBytes(this.User2Name + "@" + this.Domain + "\0");

            delegateRecipientBlock.PropertiesData = delegateRecipientProperties;
            #endregion

            delegateActionData.RecipientsData = new RecipientBlock[1] { delegateRecipientBlock };
            RuleData ruleDelegate = AdapterHelper.GenerateValidRuleData(ActionType.OP_DELEGATE, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, delegateActionData, ruleProperties, null);
            modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleDelegate });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Delegate rule should succeed.");
            #endregion

            #region TestUser1 calls RopGetRulesTable with valid TableFlags.
            ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");
            #endregion

            #region TestUser1 retrieves rule information to check if two rules exist.
            // Retrieves rows from the rule table.
            queryRowResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyTags);
            Site.Assert.AreEqual<uint>(0, queryRowResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");

            // Two rules have been added to the Inbox folder, so the row count in the rule table should be 2.
            Site.Assert.AreEqual<uint>(2, queryRowResponse.RowCount, "The rule number in the rule table  should be 2.");
            bool twoRulesExistBeforeReplaceAll = queryRowResponse.RowCount == 2;
            this.VerifyRuleTable();

            #endregion

            #region TestUser1 adds another rule to replace all the existing rules.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameDelegate, 2);
            delegateActionData = new ForwardDelegateActionData
            {
                RecipientCount = (ushort)0x01
            };

            #region Prepare the Delegate rule Recipient block
            delegateRecipientBlock = new RecipientBlock
            {
                Reserved = 0x01,
                NoOfProperties = (ushort)0x05u
            };
            delegateRecipientProperties = new TaggedPropertyValue[5];

            tempArray = AdapterHelper.GenerateRecipientPropertiesBlock(this.User2Name, this.User2ESSDN);
            Array.Copy(tempArray, 0, delegateRecipientProperties, 0, tempArray.Length);

            // Add PidTagSmtpEmailAdderss
            delegateRecipientProperties[4] = new TaggedPropertyValue();
            delegateRecipientPropertiesPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagSmtpAddress,
                PropertyType = (ushort)PropertyType.PtypString
            };
            delegateRecipientProperties[4].PropertyTag = delegateRecipientPropertiesPropertyTag;
            delegateRecipientProperties[4].Value = Encoding.Unicode.GetBytes(this.User2Name + "@" + this.Domain + "\0");

            delegateRecipientBlock.PropertiesData = delegateRecipientProperties;
            #endregion

            delegateActionData.RecipientsData = new RecipientBlock[1] { delegateRecipientBlock };
            ruleDelegate = AdapterHelper.GenerateValidRuleData(ActionType.OP_DELEGATE, TestRuleDataType.ForAdd, 0, RuleState.ST_ENABLED, delegateActionData, ruleProperties, null);
            modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleDelegate });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Delegate rule should succeed.");
            #endregion

            #region TestUser1 calls RopGetRulesTable with valid TableFlags.
            ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");
            #endregion

            #region TestUser1 retrieves rule information to check if the existing rules are replaced.
            // Retrieves rows from the rule table.
            queryRowResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyTags);
            Site.Assert.AreEqual<uint>(0, queryRowResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");

            // Only one rule exist on Inbox folder, so the row count in the rule table should be 1.
            Site.Assert.AreEqual<uint>((uint)1, queryRowResponse.RowCount, "The rule number in the rule table is {0}", queryRowResponse.RowCount);
            this.VerifyRuleTable();

            // Only one rule exist on Inbox folder, means other rules have been replaced.
            bool oneRuleExistsAfterReplaceAll = queryRowResponse.RowCount == 1;
            ruleId = BitConverter.ToUInt64(queryRowResponse.RowData.PropertyRows[0].PropertyValues[1].Value, 0);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R669: the row count of rule table before replace is {0}, and after replace is {1}", twoRulesExistBeforeReplaceAll, oneRuleExistsAfterReplaceAll);

            // Verify MS-OXORULE requirement: MS-OXORULE_R669
            bool isVerifiedR669 = twoRulesExistBeforeReplaceAll && oneRuleExistsAfterReplaceAll;

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR669,
                669,
                @"[In RopModifyRules ROP Request Buffer] [ModifyRulesFlag] R (Bitmask 0x01): If this bit is set, the rules (2) in this request are to replace the existing set of rules (2) in the folder.");
            #endregion

            #region TestUser1 deletes the last created rule.
            ruleDelegate = AdapterHelper.GenerateValidRuleData(ActionType.OP_DELEGATE, TestRuleDataType.ForRemove, 0, RuleState.ST_ENABLED, delegateActionData, ruleProperties, ruleId);
            modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleDelegate });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Deleting the created rule should be success");
            #endregion

            #region TestUser1 calls RopGetRulesTable with valid TableFlags.
            ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");
            #endregion

            #region TestUser1 retrieves rule information to check if the last created rule is deleted.
            // Retrieves rows from the rule table.
            queryRowResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyTags);
            Site.Assert.AreEqual<uint>(0, queryRowResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");

            // Get rule table succeed and the rule count is 0, means the server is returning a table with no rule.
            Site.Assert.IsTrue(queryRowResponse.RowCount == 0, "The rule number in the rule table is {0}", queryRowResponse.RowCount);

            // Add the debug information.
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R14");

            // Verify MS-OXORULE requirement: MS-OXORULE_R14.
            // According to above steps, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                14,
                @"[In RopModifyRules ROP] The RopModifyRules ROP ([MS-OXCROPS] section 2.2.11.1) creates, modifies, or deletes rules (2) in a folder.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the rule actions of OP_MOVE, OP_COPY, OP_BOUNCE and OP_DEFERRED. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S01_TC06_AddMoveCopyAndBounceDeferredrules()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region TestUser1 prepares value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameBounce);
            #endregion

            #region TestUser1 adds Bounce rule.
            BounceActionData bounceActionData = new BounceActionData
            {
                Bounce = BounceCode.CanNotDisplay
            };

            RuleData ruleBounce = AdapterHelper.GenerateValidRuleData(ActionType.OP_BOUNCE, TestRuleDataType.ForAdd, 0, RuleState.ST_ENABLED, bounceActionData, ruleProperties, null);
            RopModifyRulesResponse ropModifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleBounce });
            Site.Assert.AreEqual<uint>(0, ropModifyRulesResponse.ReturnValue, "Adding Bounce rule should succeed.");
            #endregion

            #region TestUser1 adds Deferred Action rule.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameDeferredAction);
            DeferredActionData deferredActionData = new DeferredActionData
            {
                Data = Common.GetBytesFromBinaryHexString(Constants.DeferredActionBufferData)
            };
            RuleData deferredActionRuleData = AdapterHelper.GenerateValidRuleData(ActionType.OP_DEFER_ACTION, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, deferredActionData, ruleProperties, null);

            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { deferredActionRuleData });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding deferred action rule to public folder should succeed.");
            #endregion

            if (Common.IsRequirementEnabled(294, this.Site))
            {
                #region TestUser1 creates a folder in server store.
                RopCreateFolderResponse createFolderResponse;
                string user1FolderName = Common.GenerateResourceName(this.Site, "User1Folder");
                uint folderHandle = this.OxoruleAdapter.RopCreateFolder(this.InboxFolderHandle, user1FolderName, "TestForOP_COPY", out createFolderResponse);
                ulong folderId = createFolderResponse.FolderId;
                Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating folder operation should succeed.");
                #endregion

                #region TestUser1 prepares rules' data.
                MoveCopyActionData moveCopyActionData = new MoveCopyActionData();

                // Get the created folder entry ID.
                byte[] folderEId = this.OxoruleAdapter.GetFolderEntryId(StoreObjectType.Mailbox, folderHandle, folderId);

                // Get the store object's entry ID.
                byte[] storeEId = this.GetStoreObjectEntryID(StoreObjectType.Mailbox, this.Server, this.User1ESSDN);
                moveCopyActionData.FolderEID = folderEId;
                moveCopyActionData.StoreEID = storeEId;
                moveCopyActionData.FolderEIDSize = (ushort)folderEId.Length;
                moveCopyActionData.StoreEIDSize = (ushort)storeEId.Length;
                #endregion

                #region TestUSer1 generates test RuleData.
                ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameCopy);
                RuleData ruleForCopy = AdapterHelper.GenerateValidRuleData(ActionType.OP_COPY, TestRuleDataType.ForAdd, 2, RuleState.ST_ENABLED, moveCopyActionData, ruleProperties, null);

                // Add rule for move with rule Provider Data.
                ruleProperties.ProviderData = Constants.PidTagRuleProviderData;
                ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameMoveOne);
                RuleData ruleForMove = AdapterHelper.GenerateValidRuleData(ActionType.OP_MOVE, TestRuleDataType.ForAdd, 3, RuleState.ST_ENABLED, moveCopyActionData, ruleProperties, null);
                #endregion

                #region TestUser1 adds OP_MOVE and OP_COPY rules to the Inbox folder.
                modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleForMove, ruleForCopy });

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R680");

                // Verify MS-OXORULE requirement: MS-OXORULE_R680.
                // If the return value of the RopModifyRules response is 0x00, it means operation is successfully.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x00000000,
                    modifyRulesResponse.ReturnValue,
                    680,
                    @"[In RopModifyRules ROP Response Buffer] ReturnValue: To indicate success, the server returns 0x00000000.");

                // Wait for the mail to be received and the rule to take effect.
                Thread.Sleep(this.WaitForTheRuleToTakeEffect);
                #endregion

                RopDeleteFolderResponse deleteFolder = this.OxoruleAdapter.RopDeleteFolder(this.InboxFolderHandle, folderId);
                Site.Assert.AreEqual<uint>(0, deleteFolder.ReturnValue, "Deleting folder should succeed.");
            }

            #region TestUser1 gets rows from the rule table.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "RopGetRulesTable operation should succeed.");

            PropertyTag[] propertyListTag = new PropertyTag[4];

            // PidTagRuleActions property.
            propertyListTag[0].PropertyId = (ushort)PropertyId.PidTagRuleActions;
            propertyListTag[0].PropertyType = (ushort)PropertyType.PtypRuleAction;
            propertyListTag[1].PropertyId = (ushort)PropertyId.PidTagRuleName;
            propertyListTag[1].PropertyType = (ushort)PropertyType.PtypString;
            propertyListTag[2].PropertyId = (ushort)PropertyId.PidTagRuleProvider;
            propertyListTag[2].PropertyType = (ushort)PropertyType.PtypString;
            propertyListTag[3].PropertyId = (ushort)PropertyId.PidTagRuleProviderData;
            propertyListTag[3].PropertyType = (ushort)PropertyType.PtypBinary;

            // Set the query target to standardardRules.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForStandardRules;

            RopQueryRowsResponse getAllActionsResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyListTag);
            Site.Assert.AreEqual<uint>(0, getAllActionsResponse.ReturnValue, "Getting rule properties should succeed.");

            // When R294 is enabled, there should be 4 rules exist on Inbox folder, otherwise only 2 rules exist.
            // If the rule table is got successfully and the rule count is correct, it means that the server is returning a table with the rule added by the test suite.
            if (Common.IsRequirementEnabled(294, this.Site))
            {
                Site.Assert.AreEqual(4, getAllActionsResponse.RowCount, @"There should be 4 rules returned, actual returned row count is {0}.", getAllActionsResponse.RowCount);
            }
            else
            {
                Site.Assert.AreEqual(2, getAllActionsResponse.RowCount, @"There should be 2 rules returned, actual returned row count is {0}.", getAllActionsResponse.RowCount);
            }

            this.VerifyRuleTable();

            // Clear the status of the adapter.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.OtherTarget;
            #endregion
        }

        /// <summary>
        /// This test case is designed to test the operation adding new standard rules OP_Forward and OP_Delegate to the server.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S01_TC07_AddForwardAndDelegateRules()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region TestUser1 prepares value for rule properties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameForward);
            #endregion

            #region TestUser1 prepares the recipient block for Forward rules.
            RecipientBlock recipientBlock = new RecipientBlock
            {
                Reserved = 0x01,
                NoOfProperties = (ushort)0x04u
            };

            TaggedPropertyValue[] recipientProperties = AdapterHelper.GenerateRecipientPropertiesBlock(this.User2Name, this.User2ESSDN);
            recipientBlock.PropertiesData = recipientProperties;
            #endregion

            #region TestUser1 adds rule forward with ActionFlavor set to PR and NC.
            ForwardDelegateActionData forwardActionData = new ForwardDelegateActionData
            {
                RecipientCount = (ushort)0x01,
                RecipientsData = new RecipientBlock[1]
                {
                    recipientBlock
                }
            };
            uint actionForwardFlavorNC_PR = (uint)ActionFlavorsForward.NC | (uint)ActionFlavorsForward.PR;

            RuleData ruleForward = AdapterHelper.GenerateValidRuleDataWithFlavor(ActionType.OP_FORWARD, 100, RuleState.ST_ENABLED, forwardActionData, actionForwardFlavorNC_PR, ruleProperties);
            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForward });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Forward rule should be success");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R276.");

            // Verify MS-OXORULE requirement: MS-OXORULE_R276.
            // If the ReturnValue of modifyRulesResponse response is 0x00000000, it means PR can be combined with the NC ActionFlavor flag.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                modifyRulesResponse.ReturnValue,
                276,
                @"[In Action Flavors] PR (Bitmask 0x00000001): Can be combined with the NC ActionFlavor flag.");
            #endregion

            #region TestUser1 adds rule forward with ActionFlavor set to AT.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameForwardAT);
            ForwardDelegateActionData forwardActionDataAT = new ForwardDelegateActionData
            {
                RecipientCount = (ushort)0x01,
                RecipientsData = new RecipientBlock[1]
                {
                    recipientBlock
                }
            };
            RuleData ruleForwardAT = AdapterHelper.GenerateValidRuleDataWithFlavor(ActionType.OP_FORWARD, 100, RuleState.ST_ENABLED, forwardActionDataAT, (uint)ActionFlavorsForward.AT, ruleProperties);
            modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleForwardAT });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Forward rule should be success");
            #endregion

            #region TestUser1 adds rule forward with ActionFlavor set to TM.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameForwardTM);
            ForwardDelegateActionData forwardActionDataTM = new ForwardDelegateActionData
            {
                RecipientCount = (ushort)0x01,
                RecipientsData = new RecipientBlock[1]
                {
                    recipientBlock
                }
            };
            RuleData ruleForwardTM = AdapterHelper.GenerateValidRuleDataWithFlavor(ActionType.OP_FORWARD, 100, RuleState.ST_ENABLED, forwardActionDataTM, (uint)ActionFlavorsForward.TM, ruleProperties);
            modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleForwardTM });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Forward rule should be success");
            #endregion

            #region TestUser1 adds rule DELEGATE.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameDelegate);
            ForwardDelegateActionData delegateActionData = new ForwardDelegateActionData
            {
                RecipientCount = (ushort)0x01
            };

            #region Prepare the Delegate rule Recipient block.
            RecipientBlock delegateRecipientBlock = new RecipientBlock
            {
                Reserved = 0x01,
                NoOfProperties = (ushort)0x05u
            };
            TaggedPropertyValue[] delegateRecipientProperties = new TaggedPropertyValue[5];

            TaggedPropertyValue[] tempArray = AdapterHelper.GenerateRecipientPropertiesBlock(this.User2Name, this.User2ESSDN);
            Array.Copy(tempArray, 0, delegateRecipientProperties, 0, tempArray.Length);

            // Add PidTagSmtpEmailAdderss
            delegateRecipientProperties[4] = new TaggedPropertyValue();
            PropertyTag delegateRecipientPropertiesPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagSmtpAddress,
                PropertyType = (ushort)PropertyType.PtypString
            };
            delegateRecipientProperties[4].PropertyTag = delegateRecipientPropertiesPropertyTag;
            delegateRecipientProperties[4].Value = Encoding.Unicode.GetBytes(this.User2Name + "@" + this.Domain + "\0");

            delegateRecipientBlock.PropertiesData = delegateRecipientProperties;
            #endregion

            delegateActionData.RecipientsData = new RecipientBlock[1] { delegateRecipientBlock };
            RuleData ruleDelegate = AdapterHelper.GenerateValidRuleData(ActionType.OP_DELEGATE, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, delegateActionData, ruleProperties, null);
            modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleDelegate });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Delegate rule should succeed.");
            #endregion

            #region TestUser1 gets ruleActions of the four rules.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");

            PropertyTag actionsTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagRuleActions,
                PropertyType = (ushort)PropertyType.PtypRuleAction
            };

            // Set the query target to standardardRules.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForStandardRules;

            RopQueryRowsResponse getAllActionsResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, new PropertyTag[1] { actionsTag });
            Site.Assert.AreEqual<uint>(0, getAllActionsResponse.ReturnValue, "Getting the rule actions should succeed.");

            // Four rules have been added to the Inbox folder, so the row count in the rule table should be 4.
            Site.Assert.AreEqual<uint>(4, getAllActionsResponse.RowCount, "The rule number in the rule table is {0}", getAllActionsResponse.RowCount);
            this.VerifyRuleTable();

            // Clear the status of the adapter.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.OtherTarget;
            #endregion
        }

        /// <summary>
        /// This test case is designed to test the operation adding OP_REPLY and OP_OOF_REPLY rules to the server.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S01_TC08_AddReplyAndOOF_ReplyRules()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region TestUser1 prepares value for rule properties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameOOFReply);
            #endregion

            #region Create a Reply template.
            RopCreateMessageResponse ropCreateMessageResponse;
            uint replyTemplateMessageHandler = this.OxoruleAdapter.RopCreateMessage(this.InboxFolderHandle, this.InboxFolderID, Convert.ToByte(true), out ropCreateMessageResponse);
            Site.Assert.AreEqual<uint>(0, ropCreateMessageResponse.ReturnValue, "Creating FAI message should succeed.");

            TaggedPropertyValue[] replyTemplateProperties = new TaggedPropertyValue[3];

            // PidTagMessageClass
            replyTemplateProperties[0] = new TaggedPropertyValue();
            PropertyTag pidTagMessageClassPropertyTag1 = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            replyTemplateProperties[0].PropertyTag = pidTagMessageClassPropertyTag1;
            replyTemplateProperties[0].Value = Encoding.Unicode.GetBytes(Constants.ReplyTemplate + "\0");

            // PidTagReplyTemplateId
            replyTemplateProperties[1] = new TaggedPropertyValue();
            PropertyTag pidTagReplyTemplateIdPropertyTag1 = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagReplyTemplateId,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            replyTemplateProperties[1].PropertyTag = pidTagReplyTemplateIdPropertyTag1;
            Guid newReplyTemplateGuid = System.Guid.NewGuid();
            replyTemplateProperties[1].Value = Common.AddInt16LengthBeforeBinaryArray(newReplyTemplateGuid.ToByteArray());

            // PidTagSubject
            replyTemplateProperties[2] = new TaggedPropertyValue();
            PropertyTag pidTagSubjectPropertyTag1 = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagSubject,
                PropertyType = (ushort)PropertyType.PtypString
            };
            replyTemplateProperties[2].PropertyTag = pidTagSubjectPropertyTag1;
            replyTemplateProperties[2].Value = Encoding.Unicode.GetBytes(Common.GenerateResourceName(this.Site, Constants.ReplyTemplateSubject) + "\0");

            RopSetPropertiesResponse ropSetPropertiesResponse = this.OxoruleAdapter.RopSetProperties(replyTemplateMessageHandler, replyTemplateProperties);
            Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "Setting property for Extended rule FAI message should succeed.");

            // Save changes of message.
            RopSaveChangesMessageResponse ropSaveChangesMessagResponseReply = this.OxoruleAdapter.RopSaveChangesMessage(replyTemplateMessageHandler);
            Site.Assert.AreEqual(0, (int)ropSaveChangesMessagResponseReply.ReturnValue, "Saving Extend rule message should succeed.");

            // Get the newly created message's folder ID.
            PropertyTag fidTagReplyTemplate = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagFolderId,
                PropertyType = (ushort)PropertyType.PtypInteger64
            };
            RopGetPropertiesSpecificResponse ropGetPropertiesSpecificResponseReply = this.OxoruleAdapter.RopGetPropertiesSpecific(replyTemplateMessageHandler, new PropertyTag[1] { fidTagReplyTemplate });
            Site.Assert.AreEqual<uint>(0, ropGetPropertiesSpecificResponseReply.ReturnValue, "Getting folder id operation should succeed.");

            // Get the reply template's guid.
            PropertyTag fidTagOOFReply = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagReplyTemplateId,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            RopGetPropertiesSpecificResponse ropGetGuid = this.OxoruleAdapter.RopGetPropertiesSpecific(replyTemplateMessageHandler, new PropertyTag[1] { fidTagOOFReply });
            Site.Assert.AreEqual<uint>(0, ropGetGuid.ReturnValue, "Getting guid property operation should succeed.");
            #endregion

            #region TestUser1 adds rule OOF_REPLY.
            ReplyActionData oofReplyActionData = new ReplyActionData
            {
                ReplyTemplateGUID = new byte[ropGetGuid.RowData.PropertyValues[0].Value.Length - 2]
            };
            Array.Copy(ropGetGuid.RowData.PropertyValues[0].Value, 2, oofReplyActionData.ReplyTemplateGUID, 0, ropGetGuid.RowData.PropertyValues[0].Value.Length - 2);
            oofReplyActionData.ReplyTemplateFID = BitConverter.ToUInt64(ropGetPropertiesSpecificResponseReply.RowData.PropertyValues[0].Value, 0);
            oofReplyActionData.ReplyTemplateMID = ropSaveChangesMessagResponseReply.MessageId;
            RuleData ruleForOOFReply = AdapterHelper.GenerateValidRuleData(ActionType.OP_OOF_REPLY, TestRuleDataType.ForAdd, 100, RuleState.ST_ONLY_WHEN_OOF, oofReplyActionData, ruleProperties, null);

            RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForOOFReply });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding OOF_REPLY rule should succeed.");
            #endregion

            #region TestUser1 adds rule REPLY.
            ruleProperties.Name = Common.GenerateResourceName(this.Site, Constants.RuleNameReply);
            ReplyActionData replyActionData = new ReplyActionData
            {
                ReplyTemplateGUID = new byte[ropGetGuid.RowData.PropertyValues[0].Value.Length - 2]
            };
            Array.Copy(ropGetGuid.RowData.PropertyValues[0].Value, 2, replyActionData.ReplyTemplateGUID, 0, ropGetGuid.RowData.PropertyValues[0].Value.Length - 2);
            replyActionData.ReplyTemplateFID = BitConverter.ToUInt64(ropGetPropertiesSpecificResponseReply.RowData.PropertyValues[0].Value, 0);
            replyActionData.ReplyTemplateMID = ropSaveChangesMessagResponseReply.MessageId;
            RuleData ruleForReply = AdapterHelper.GenerateValidRuleData(ActionType.OP_REPLY, TestRuleDataType.ForAdd, 0, RuleState.ST_ENABLED, replyActionData, ruleProperties, null);
            modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, new RuleData[] { ruleForReply });
            Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Adding Reply rule should succeed.");
            #endregion

            #region TestUser1 gets ruleActions of the two rules.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreEqual<uint>(0, ropGetRulesTableResponse.ReturnValue, "Getting rule table should succeed.");

            PropertyTag[] propertyTags = new PropertyTag[2];

            // PidTagRuleName
            propertyTags[0].PropertyId = (ushort)PropertyId.PidTagRuleName;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypString;

            // PidTagRuleActions
            propertyTags[1].PropertyId = (ushort)PropertyId.PidTagRuleActions;
            propertyTags[1].PropertyType = (ushort)PropertyType.PtypRuleAction;

            // Set the query target to standardardRules.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForStandardRules;

            RopQueryRowsResponse getAllActionsResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, propertyTags);
            Site.Assert.AreEqual<uint>(0, getAllActionsResponse.ReturnValue, "Getting the rule actions should succeed.");

            // Two rules have been added to the Inbox folder, so the row count in the rule table should be 2.
            Site.Assert.AreEqual<uint>(2, getAllActionsResponse.RowCount, "The rule number in the rule table is {0}", getAllActionsResponse.RowCount);
            this.VerifyRuleTable();

            RuleAction ruleActionForReply = new RuleAction();
            ReplyActionData actionDataForReply = new ReplyActionData();
            RuleAction ruleActionForOOFReply = new RuleAction();
            ReplyActionData actionDataForOOFReply = new ReplyActionData();
            bool hasActionDataForReply = false;
            bool hasActionDataForOOFReply = false;
            for (int i = 0; i < getAllActionsResponse.RowCount; i++)
            {
                System.Text.UnicodeEncoding converter = new UnicodeEncoding();
                string ruleName = converter.GetString(getAllActionsResponse.RowData.PropertyRows[i].PropertyValues[0].Value);
                if (ruleName.Contains(Constants.RuleNameReply) && !hasActionDataForReply)
                {
                    ruleActionForReply.Deserialize(getAllActionsResponse.RowData.PropertyRows[i].PropertyValues[1].Value);
                    actionDataForReply.Deserialize(ruleActionForReply.Actions[0].ActionDataValue.Serialize());
                    hasActionDataForReply = true;
                }

                if (ruleName.Contains(Constants.RuleNameOOFReply) && !hasActionDataForOOFReply)
                {
                    ruleActionForOOFReply.Deserialize(getAllActionsResponse.RowData.PropertyRows[i].PropertyValues[1].Value);
                    actionDataForOOFReply.Deserialize(ruleActionForOOFReply.Actions[0].ActionDataValue.Serialize());
                    hasActionDataForOOFReply = true;
                }

                if (hasActionDataForReply && hasActionDataForOOFReply)
                {
                    break;
                }
            }

            // Add a variable to verify R727 and R928.
            bool isPidTagReplyTemplateIdEqualtemplateGUIDForOOFReply = false;

            if (Common.CompareByteArray(newReplyTemplateGuid.ToByteArray(), actionDataForOOFReply.ReplyTemplateGUID))
            {
                isPidTagReplyTemplateIdEqualtemplateGUIDForOOFReply = true;
            }

            // Add a variable to verify R309
            bool isPidTagReplyTemplateIdEqualtemplateGUIDForReply = false;
            if (Common.CompareByteArray(newReplyTemplateGuid.ToByteArray(), actionDataForReply.ReplyTemplateGUID))
            {
                isPidTagReplyTemplateIdEqualtemplateGUIDForReply = true;
            }

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R727: the actual value of PidTagReplyTemplateId property is {0}", BitConverter.ToString(actionDataForOOFReply.ReplyTemplateGUID));

            // Verify MS-OXORULE requirement: MS-OXORULE_R727.
            // ReplyTemplateGUID of the actionDataForOOFReply is the value of PidTagReplyTemplateId property, and newGuidOOFReplyGUID is the GUID of the reply template.
            bool isVerifyR727 = isPidTagReplyTemplateIdEqualtemplateGUIDForOOFReply;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR727,
                727,
                @"[In PidTagReplyTemplateId Property] The PidTagReplyTemplateId property ([MS-OXPROPS] section 2.909) specifies the GUID for the reply template.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R928: the actual value of PidTagReplyTemplateId property is {0}", BitConverter.ToString(actionDataForOOFReply.ReplyTemplateGUID));

            // Verify MS-OXORULE requirement: MS-OXORULE_R928.
            // ReplyTemplateGUID of the actionDataForOOFReply is the value of PidTagReplyTemplateId property, and newGuidOOFReplyGUID is the GUID of the reply template.
            bool isVerifyR928 = isPidTagReplyTemplateIdEqualtemplateGUIDForOOFReply;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR928,
                928,
                @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Standard Rules] The value of the ReplyTemplateGUID field in OP_OOF_REPLY action data is equal to the value of the PidTagReplyTemplateId property (section 2.2.9.2) that is set on the reply template.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R309: the actual value of PidTagReplyTemplateId property is {0}", BitConverter.ToString(actionDataForReply.ReplyTemplateGUID));

            // Verify MS-OXORULE requirement: MS-OXORULE_R309.
            // ReplyTemplateGUID of the actionDataForReply is the value of PidTagReplyTemplateId property, and newGuidOOFReplyGUID is the GUID of the reply template.
            bool isVerifyR309 = isPidTagReplyTemplateIdEqualtemplateGUIDForReply;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR309,
                309,
                @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Standard Rules] The value of the ReplyTemplateGUID field in OP_REPLY action data is equal to the value of the PidTagReplyTemplateId property (section 2.2.9.2) that is set on the reply template.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R305, the folder ID which contains the reply template is {0}, and actually the value is {1} in the actionDataForReply, and the value is {2} in actionDataForOOFReply", replyActionData.ReplyTemplateFID, actionDataForReply.ReplyTemplateFID, actionDataForOOFReply.ReplyTemplateFID);
            bool isVerifiedR305 = replyActionData.ReplyTemplateFID == actionDataForReply.ReplyTemplateFID && replyActionData.ReplyTemplateFID == actionDataForOOFReply.ReplyTemplateFID;

            // Verify MS-OXORULE requirement: MS-OXORULE_R305.
            // ReplyTemplateFID in the replyActionData is set to the folder ID that contains the reply template, so if the value in the rule actionData is the same as it, R305 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR305,
                305,
                @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Standard Rules] ReplyTemplateFID (8 bytes): A Folder ID structure, as specified in [MS-OXCDATA] section 2.2.1.1, that identifies the folder that contains the reply template.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R307, the Message ID which is used as the reply template is {0}, and actually the value is {1} in the actionDataForReply, and the value is {2} in the actionDataForOOFReply", replyActionData.ReplyTemplateMID, actionDataForReply.ReplyTemplateMID, actionDataForOOFReply.ReplyTemplateMID);
            bool isVerifiedR307 = replyActionData.ReplyTemplateMID == actionDataForReply.ReplyTemplateMID && replyActionData.ReplyTemplateMID == actionDataForOOFReply.ReplyTemplateMID;
            
            // Verify MS-OXORULE requirement: MS-OXORULE_R307.
            // ReplyTemplateMID in the replyActionData is set to the message ID which is used as the reply template, so if the value in the rule actionData is the same as it, R307 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR307,
                307,
                @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Standard Rules] ReplyTemplateMID (8 bytes): A Message ID structure, as specified in [MS-OXCDATA] section 2.2.1.2, that identifies the FAI message being used as the reply template.");
            #endregion

            // Clear the status of the adapter.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.OtherTarget;
            #endregion
        }

        /// <summary>
        /// This test case is designed to test the operation adding extended rules for three times.
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S01_TC09_AddExtendedRuleForThreeTimes()
        {
            this.CheckMAPIHTTPTransportSupported();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(646, this.Site), "This case runs only when the server supports processing more than two extended rules it encounters per folder.");

            #region TestUser1 creates an FAI message for the first extended rule.
            RopCreateMessageResponse ropCreateMessageResponse;
            uint extendedRuleMessageHandle1 = this.OxoruleAdapter.RopCreateMessage(this.InboxFolderHandle, this.InboxFolderID, Convert.ToByte(true), out ropCreateMessageResponse);
            Site.Assert.AreEqual<uint>(0, ropCreateMessageResponse.ReturnValue, "Creating the first FAI message should succeed.");

            NamedPropertyInfo namedPropertyInfo1 = new NamedPropertyInfo
            {
                NoOfNamedProps = 0
            };
            TaggedPropertyValue[] extendedRuleProperties1 = AdapterHelper.GenerateExtendedRuleTestData(Common.GenerateResourceName(this.Site, Constants.ExtendRulename1), 0, (uint)RuleState.ST_ENABLED, Constants.PidTagRuleProvider, ActionType.OP_MARK_AS_READ, new DeleteMarkReadActionData(), Constants.ExtendRuleCondition1, namedPropertyInfo1);

            // Set properties for extended rule FAI message.
            RopSetPropertiesResponse ropSetPropertiesResponse = this.OxoruleAdapter.RopSetProperties(extendedRuleMessageHandle1, extendedRuleProperties1);
            Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "Setting property for Extended rule FAI message should succeed.");

            // Save changes of message.
            RopSaveChangesMessageResponse ropSaveChangesMessagResponse = this.OxoruleAdapter.RopSaveChangesMessage(extendedRuleMessageHandle1);
            Site.Assert.AreEqual(0, (int)ropSaveChangesMessagResponse.ReturnValue, "Saving Extend rule message should succeed.");
            #endregion

            #region TestUser1 retrieves data of the extended rule.
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.ForExtendedRules;
            RopGetPropertiesAllResponse ropGetExtendRuleMessageResponse = this.OxoruleAdapter.RopGetPropertiesAll(extendedRuleMessageHandle1, this.PropertySizeLimitFlag, (ushort)WantUnicode.Want);
            Site.Assert.AreEqual<uint>(0, ropGetExtendRuleMessageResponse.ReturnValue, "Getting all properties operation should succeed.");
            Site.Assert.IsTrue(ropGetExtendRuleMessageResponse.PropertyValues.Length != 0, "Extended Rule data should be found in related FAI message!");
            this.OxoruleAdapter.TargetOfRop = TargetOfRop.OtherTarget;

            ExtendedRuleActions extendedRuleMessageActions = new ExtendedRuleActions();

            // Check the properties set on Extended Rule, and find the Extended Rule Actions.
            for (int i = 0; i < ropGetExtendRuleMessageResponse.PropertyValues.Length; i++)
            {
                // propertyId indicates the Id of a property set on Extended Rule.
                ushort propertyId = ropGetExtendRuleMessageResponse.PropertyValues[i].PropertyTag.PropertyId;
                if (propertyId == (ushort)PropertyId.PidTagExtendedRuleMessageActions)
                {
                    byte[] propertyValue = ropGetExtendRuleMessageResponse.PropertyValues[i].Value;
                    extendedRuleMessageActions = AdapterHelper.PropertyValueConvertToExtendedRuleActions(propertyValue);
                    break;
                }
            }

            // Get the Property Names saved by server in the extendedRuleMessageActions.
            PropertyName[] propertyNames = extendedRuleMessageActions.NamedPropertyInformation.NamedProperty;
            uint[] propertyIds = extendedRuleMessageActions.NamedPropertyInformation.PropId;
            Site.Assert.AreEqual<uint>(0, extendedRuleMessageActions.NamedPropertyInformation.NoOfNamedProps, "The property NoOfNamedProps of NamedPropertyInformation should be zero!");

            #region Capture Code
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R696.");

            // NoOfNamedProps set to 0, it means no named properties are used in the structure that follows the Named Property Information buffer.
            // So if NamedProperty of NamedPropertyInformation in the extendedRuleMessageActions is null. This requirement can be verified. 
            bool isVerifyR696 = propertyNames == null;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR696,
                696,
                @"[In NamedPropertyInformation Structure] NoOfNamedProps (2 bytes): If no named properties are used in the structure that follows the NamedPropertyInformation structure, the value of this field MUST be 0x0000.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R936");

            // NoOfNamedProps set to 0, it means no named properties are used in the structure that follows the Named Property Information buffer.
            // So if NamedProperty, PropIds of NamedPropertyInformation in the extendedRuleMessageActions is null. This requirement can be verified. 
            bool isVerifyR936 = propertyNames == null && propertyIds == null;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR936,
                936,
                @"[In NamedPropertyInformation Structure] [If no named properties are used in the structure that follows the NamedPropertyInformation structure] no other fields [except NoOfNamedProps] are present.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R697");

            // When NoOfNamedProps is set to 0, NamedPropertyInformation in ExtendedRuleActions reduces to a 2-byte WORD value of NoOfNamedProps.
            Site.CaptureRequirement(
                697,
                @"[In NamedPropertyInformation Structure] Note that if there are no named properties to be listed, the NamedPropertyInformation structure reduces to a 2-byte value of 0x0000.");
            #endregion
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger the rule.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser2 delivers a message to TestUser1 to trigger the rule.
            string mailSubject1 = Common.GenerateResourceName(this.Site, Constants.ExtendRuleCondition1);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject1);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            PropertyTag[] propertyTagList = new PropertyTag[2];
            propertyTagList[0].PropertyId = (ushort)PropertyId.PidTagSubject;
            propertyTagList[0].PropertyType = (ushort)PropertyType.PtypString;
            propertyTagList[1].PropertyId = (ushort)PropertyId.PidTagMessageFlags;
            propertyTagList[1].PropertyType = (ushort)PropertyType.PtypInteger32;

            int messageFlag1 = 1;
            uint contentTableHandle1 = 0;
            int expectedMessageIndex1 = 0;
            RopQueryRowsResponse getMailMessageContent1 = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandle1, propertyTagList, ref expectedMessageIndex1, mailSubject1);
            mailSubject1 = AdapterHelper.PropertyValueConvertToString(getMailMessageContent1.RowData.PropertyRows[expectedMessageIndex1].PropertyValues[0].Value);
            messageFlag1 = BitConverter.ToInt32(getMailMessageContent1.RowData.PropertyRows[expectedMessageIndex1].PropertyValues[1].Value, 0);

            #endregion

            #region TestUser1 creates an FAI message for the second extended rule.
            uint extendedRuleMessageHandle2 = this.OxoruleAdapter.RopCreateMessage(this.InboxFolderHandle, this.InboxFolderID, Convert.ToByte(true), out ropCreateMessageResponse);
            Site.Assert.AreEqual<uint>(0, ropCreateMessageResponse.ReturnValue, "Creating the second FAI message should succeed.");
            #endregion

            #region TestUser1 creates the second extended rule with no NamedProperty.
            NamedPropertyInfo namedPropertyInfo2 = new NamedPropertyInfo();
            namedPropertyInfo1.NoOfNamedProps = 0;
            TaggedPropertyValue[] extendedRuleProperties2 = AdapterHelper.GenerateExtendedRuleTestData(Common.GenerateResourceName(this.Site, Constants.ExtendRulename2), 0, (uint)RuleState.ST_ENABLED, Constants.PidTagRuleProvider, ActionType.OP_MARK_AS_READ, new DeleteMarkReadActionData(), Constants.ExtendRuleCondition2, namedPropertyInfo2);

            // Set properties for extended rule FAI message.
            ropSetPropertiesResponse = this.OxoruleAdapter.RopSetProperties(extendedRuleMessageHandle2, extendedRuleProperties2);
            Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "Setting property for Extended rule FAI message should succeed.");

            // Save changes of message.
            ropSaveChangesMessagResponse = this.OxoruleAdapter.RopSaveChangesMessage(extendedRuleMessageHandle2);
            Site.Assert.AreEqual(0, (int)ropSaveChangesMessagResponse.ReturnValue, "Saving Extend rule message should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger the rule.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser2 delivers a message to TestUser1 to trigger the rule.
            string mailSubject2 = Common.GenerateResourceName(this.Site, Constants.ExtendRuleCondition2);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject2);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            int messageFlag2 = 1;
            uint contentTableHandle2 = 0;
            int expectedMessageIndex2 = 0;
            RopQueryRowsResponse getMailMessageContent2 = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandle2, propertyTagList, ref expectedMessageIndex2, mailSubject2);
            mailSubject2 = AdapterHelper.PropertyValueConvertToString(getMailMessageContent2.RowData.PropertyRows[expectedMessageIndex2].PropertyValues[0].Value);
            messageFlag2 = BitConverter.ToInt32(getMailMessageContent2.RowData.PropertyRows[expectedMessageIndex2].PropertyValues[1].Value, 0);
            #endregion

            #region TestUser1 creates an FAI message for the third extended rule.
            uint extendedRuleMessageHandle3 = this.OxoruleAdapter.RopCreateMessage(this.InboxFolderHandle, this.InboxFolderID, Convert.ToByte(true), out ropCreateMessageResponse);
            Site.Assert.AreEqual<uint>(0, ropCreateMessageResponse.ReturnValue, "Creating the second FAI message should succeed.");
            #endregion

            #region TestUser1 creates the third extended rule with no NamedProperty.
            NamedPropertyInfo namedPropertyInfo3 = new NamedPropertyInfo();
            namedPropertyInfo1.NoOfNamedProps = 0;
            TaggedPropertyValue[] extendedRuleProperties3 = AdapterHelper.GenerateExtendedRuleTestData(Common.GenerateResourceName(this.Site, Constants.ExtendRulename3), 0, (uint)RuleState.ST_ENABLED, Constants.PidTagRuleProvider, ActionType.OP_MARK_AS_READ, new DeleteMarkReadActionData(), Constants.ExtendRuleCondition3, namedPropertyInfo3);

            // Set properties for extended rule FAI message.
            ropSetPropertiesResponse = this.OxoruleAdapter.RopSetProperties(extendedRuleMessageHandle3, extendedRuleProperties3);
            Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "Setting property for Extended rule FAI message should succeed.");

            // Save changes of the message.
            ropSaveChangesMessagResponse = this.OxoruleAdapter.RopSaveChangesMessage(extendedRuleMessageHandle3);
            Site.Assert.AreEqual(0, (int)ropSaveChangesMessagResponse.ReturnValue, "Saving Extend rule message should succeed.");
            #endregion

            #region TestUser2 delivers a message to TestUser1 to trigger the rule.
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // TestUser2 delivers a message to TestUser1 to trigger the rule.
            string mailSubject3 = Common.GenerateResourceName(this.Site, Constants.ExtendRuleCondition3);
            this.SUTAdapter.SendMailToRecipient(this.User2Name, this.User2Password, this.User1Name, mailSubject3);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            int messageFlag3 = 1;
            uint contentTableHandle3 = 0;
            int expectedMessageIndex3 = 0;
            RopQueryRowsResponse getMailMessageContent3 = this.GetExpectedMessage(this.InboxFolderHandle, ref contentTableHandle3, propertyTagList, ref expectedMessageIndex3, mailSubject3);
            mailSubject3 = AdapterHelper.PropertyValueConvertToString(getMailMessageContent3.RowData.PropertyRows[expectedMessageIndex3].PropertyValues[0].Value);
            messageFlag3 = BitConverter.ToInt32(getMailMessageContent3.RowData.PropertyRows[expectedMessageIndex3].PropertyValues[1].Value, 0);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R646: the third message is {0} marked as read.", (messageFlag3 & 0x00000001) == 0x00000001 ? string.Empty : "not");

            // Verify MS-OXORULE requirement: MS-OXORULE_R646.
            // 0x00000001 is the flag which represents the message has been read. If messageFlag doesn't set this flag, it means the incoming message 
            // isn't marked as read, which indicates the server doesn't evaluate the third rule.
            bool isVerifyR646 = (messageFlag3 & 0x00000001) == 0x00000000 && (messageFlag2 & 0x00000001) == 0x00000001 && (messageFlag1 & 0x00000001) == 0x00000001;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR646,
                646,
                @"[In Appendix A: Product Behavior] Implementation does process the standard rule for a message but does only process the first two extended rules it encounters per folder. [<15> Section 3.2.4.1: Exchange 2007 by default will process the standard rule for a message but will only process the first two extended rules it encounters per folder.]");
            #endregion
        }

        /// <summary>
        /// This test case is designed to test the operation adding OP_REPLY extended rule. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S01_TC10_AddExtendedRule_OP_REPLY()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameReply);
            #endregion

            #region Create a reply template in the TestUser1's Inbox folder.
            ulong replyTemplateMessageId;
            uint replyTemplateMessageHandle;
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

            byte[] guidByte = this.OxoruleAdapter.CreateReplyTemplate(this.InboxFolderHandle, this.InboxFolderID, false, replyTemplateSubject, replyTemplateProperties, out replyTemplateMessageId, out replyTemplateMessageHandle);
            #endregion

            #region TestUser1 gets the reply template.
            #region Step1: TestUser1 gets a table of the messages.
            uint contentsTableHandleOfFAIMessage;
            RopGetContentsTableResponse ropGetContentsTableResponseOfFAIMessage = this.OxoruleAdapter.RopGetContentsTable(this.InboxFolderHandle, ContentTableFlag.Associated, out contentsTableHandleOfFAIMessage);
            Site.Assert.AreEqual<uint>(0, ropGetContentsTableResponseOfFAIMessage.ReturnValue, "Getting DAF contents table should succeed");
            #endregion

            #region Step2: TestUser1 gets the interested columns of the message.

            // Here are 2 interested columns listed as below.
            PropertyTag[] propertyTagOfFAIMessage = new PropertyTag[2];
            PropertyTag pidTagMessageSubject = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagSubject,
                PropertyType = (ushort)PropertyType.PtypString
            };
            propertyTagOfFAIMessage[0] = pidTagMessageSubject;
            PropertyTag pidTagReplyTemplateId = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagReplyTemplateId,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            propertyTagOfFAIMessage[1] = pidTagReplyTemplateId;

            // Query rows which include the property values of the interested columns. 
            RopQueryRowsResponse ropQueryRowsResponseOfFAIMessage = this.OxoruleAdapter.QueryPropertiesInTable(contentsTableHandleOfFAIMessage, propertyTagOfFAIMessage);
            Site.Assert.AreNotEqual<ushort>(0, ropQueryRowsResponseOfFAIMessage.RowCount, "There should be 0 DAM generated in the DAF folder");
            byte[] replyTemplateId = null;
            for (int i = 0; i < ropQueryRowsResponseOfFAIMessage.RowCount; i++)
            {
                string messageSubject = AdapterHelper.PropertyValueConvertToString(ropQueryRowsResponseOfFAIMessage.RowData.PropertyRows[i].PropertyValues[0].Value);
                if (messageSubject.Equals(replyTemplateSubject, StringComparison.CurrentCultureIgnoreCase))
                {
                    replyTemplateId = ropQueryRowsResponseOfFAIMessage.RowData.PropertyRows[i].PropertyValues[1].Value;
                }
            }

            byte[] pidTagReplyTemplateIdValue = new byte[replyTemplateId.Length - 2];
            Array.Copy(replyTemplateId, 2, pidTagReplyTemplateIdValue, 0, replyTemplateId.Length - 2);

            #endregion
            #endregion

            #region TestUser1 gets the message entry ID and the Inbox folder's entry ID.
            byte[] messageEntryId = this.OxoruleAdapter.GetMessageEntryId(this.InboxFolderHandle, this.InboxFolderID, replyTemplateMessageHandle, replyTemplateMessageId);
            #endregion

            #region Prepare RuleAction data
            ReplyActionDataOfExtendedRule ruleActionData = new ReplyActionDataOfExtendedRule();
            ruleActionData.MessageEIDSize = 0x46;
            ruleActionData.ReplyTemplateMessageEID = messageEntryId;
            ruleActionData.ReplyTemplateGUID = guidByte;
            #endregion

            #region TestUser1 creates an FAI message for the OP_REPLY extended rule.
            RopCreateMessageResponse ropCreateMessageResponse;
            uint extendedRuleMessageHandle1 = this.OxoruleAdapter.RopCreateMessage(this.InboxFolderHandle, this.InboxFolderID, Convert.ToByte(true), out ropCreateMessageResponse);
            Site.Assert.AreEqual<uint>(0, ropCreateMessageResponse.ReturnValue, "Creating the FAI message should succeed.");

            NamedPropertyInfo namedPropertyInfo1 = new NamedPropertyInfo
            {
                NoOfNamedProps = 0
            };
            TaggedPropertyValue[] extendedRuleProperties1 = AdapterHelper.GenerateExtendedRuleTestData(ruleProperties.Name, 0, (uint)RuleState.ST_ENABLED, Constants.PidTagRuleProvider, ActionType.OP_REPLY, ruleActionData, ruleProperties.ConditionSubjectName, namedPropertyInfo1);

            // Set properties for extended rule FAI message.
            RopSetPropertiesResponse ropSetPropertiesResponse = this.OxoruleAdapter.RopSetProperties(extendedRuleMessageHandle1, extendedRuleProperties1);
            Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "Setting property for Extended rule FAI message should succeed.");

            // Save changes of message.
            RopSaveChangesMessageResponse ropSaveChangesMessagResponse = this.OxoruleAdapter.RopSaveChangesMessage(extendedRuleMessageHandle1);
            Site.Assert.AreEqual(0, (int)ropSaveChangesMessagResponse.ReturnValue, "Saving Extend rule message should succeed.");
            #endregion

            #region TestUser1 gets the OP_REPLY extended rule.
            #region Step1: TestUser1 gets a table of FAI messages.

            ropGetContentsTableResponseOfFAIMessage = this.OxoruleAdapter.RopGetContentsTable(this.InboxFolderHandle, ContentTableFlag.Associated, out contentsTableHandleOfFAIMessage);
            Site.Assert.AreEqual<uint>(0, ropGetContentsTableResponseOfFAIMessage.ReturnValue, "Getting contents table should succeed");
            #endregion

            #region Step2: TestUser1 sets the interested columns of the FAI message table.

            // Here are 6 interested columns listed as below.
            propertyTagOfFAIMessage = new PropertyTag[6];
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
            ropQueryRowsResponseOfFAIMessage = this.OxoruleAdapter.QueryPropertiesInTable(contentsTableHandleOfFAIMessage, propertyTagOfFAIMessage);
            Site.Assert.AreEqual<uint>(0, ropQueryRowsResponseOfFAIMessage.ReturnValue, "Querying Rows Response of FAI Message should succeed, the actual returned value is {0}", ropQueryRowsResponseOfFAIMessage.ReturnValue);

            ExtendedRuleActions extendedRuleAction = new ExtendedRuleActions();
            ReplyActionDataOfExtendedRule replyActionData = new ReplyActionDataOfExtendedRule();
            for (int i = 0; i < ropQueryRowsResponseOfFAIMessage.RowCount; i++)
            {
                // Since the PidTagMessageClass property of Extended rule MUST have a value of "IPM.ExtendedRule.Message", use PidTagMessageClass property to get the extended rule data.
                System.Text.UnicodeEncoding converter = new UnicodeEncoding();
                string messageClass = converter.GetString(ropQueryRowsResponseOfFAIMessage.RowData.PropertyRows.ToArray()[i].PropertyValues[1].Value);
                if (messageClass == "IPM.ExtendedRule.Message" + "\0")
                {
                    byte[] extendedRuleMessageActionBinary = ropQueryRowsResponseOfFAIMessage.RowData.PropertyRows[i].PropertyValues[4].Value;
                    byte[] extendedRuleMessageActionBuffer = new byte[extendedRuleMessageActionBinary.Length - 2];

                    // Remove the two length bytes to get the extended rule action data.
                    Array.Copy(extendedRuleMessageActionBinary, 2, extendedRuleMessageActionBuffer, 0, extendedRuleMessageActionBinary.Length - 2);
                    extendedRuleAction.Deserialize(extendedRuleMessageActionBuffer);
                    replyActionData.Deserialize(extendedRuleAction.RuleActionBuffer.Actions[0].ActionDataValue.Serialize());
                }
            }
            #endregion

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R1001: the value of PidTagReplyTemplateId property is {0}, and the ReplyTemplateGUID in the action data is {1}", pidTagReplyTemplateIdValue, replyActionData.ReplyTemplateGUID);

            bool isVerifiedR1001 = Common.CompareByteArray(pidTagReplyTemplateIdValue, replyActionData.ReplyTemplateGUID);

            // Verify MS-OXORULE requirement: MS-OXORULE_R1001.
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR1001,
                1001,
                @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Extended Rules] The value of the ReplyTemplateGUID field in OP_REPLY action data is equal to the value of the PidTagReplyTemplateId property that is set on the reply template.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to test the operation adding OP_OOF_REPLY extended rule. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S01_TC11_AddExtendedRule_OP_OOF_REPLY()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region Prepare value for ruleProperties variable
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameOOFReply);
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

            byte[] guidByte = this.OxoruleAdapter.CreateReplyTemplate(this.InboxFolderHandle, this.InboxFolderID, true, replyTemplateSubject, replyTemplateProperties, out replyTemplateMessageId, out replyTemplateMessageHandler);
            #endregion

            #region TestUser1 gets the reply template.
            #region Step1: TestUser1 gets a table of the messages.
            uint contentsTableHandleOfFAIMessage;
            RopGetContentsTableResponse ropGetContentsTableResponseOfFAIMessage = this.OxoruleAdapter.RopGetContentsTable(this.InboxFolderHandle, ContentTableFlag.Associated, out contentsTableHandleOfFAIMessage);
            Site.Assert.AreEqual<uint>(0, ropGetContentsTableResponseOfFAIMessage.ReturnValue, "Getting DAF contents table should succeed");
            #endregion

            #region Step2: TestUser1 sets the interested columns of the message table.

            // Here are 2 interested columns listed as below.
            PropertyTag[] propertyTagOfFAIMessage = new PropertyTag[2];
            PropertyTag pidTagMessageSubject = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagSubject,
                PropertyType = (ushort)PropertyType.PtypString
            };
            propertyTagOfFAIMessage[0] = pidTagMessageSubject;
            PropertyTag pidTagReplyTemplateId = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagReplyTemplateId,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            propertyTagOfFAIMessage[1] = pidTagReplyTemplateId;

            // Query rows which include the property values of the interested columns. 
            RopQueryRowsResponse ropQueryRowsResponseOfFAIMessage = this.OxoruleAdapter.QueryPropertiesInTable(contentsTableHandleOfFAIMessage, propertyTagOfFAIMessage);
            Site.Assert.AreNotEqual<ushort>(0, ropQueryRowsResponseOfFAIMessage.RowCount, "There should be message generated in TestUser1's Inbox folder");
            byte[] replyTemplateId = null;
            for (int i = 0; i < ropQueryRowsResponseOfFAIMessage.RowCount; i++)
            {
                string messageSubject = AdapterHelper.PropertyValueConvertToString(ropQueryRowsResponseOfFAIMessage.RowData.PropertyRows[i].PropertyValues[0].Value);
                if (messageSubject.Equals(replyTemplateSubject, StringComparison.CurrentCultureIgnoreCase))
                {
                    replyTemplateId = ropQueryRowsResponseOfFAIMessage.RowData.PropertyRows[i].PropertyValues[1].Value;
                }
            }

            byte[] pidTagReplyTemplateIdValue = new byte[replyTemplateId.Length - 2];
            Array.Copy(replyTemplateId, 2, pidTagReplyTemplateIdValue, 0, replyTemplateId.Length - 2);
            #endregion
            #endregion

            #region TestUser1 gets the message entry ID and the Inbox folder's entry ID.
            byte[] messageEntryId = this.OxoruleAdapter.GetMessageEntryId(this.InboxFolderHandle, this.InboxFolderID, replyTemplateMessageHandler, replyTemplateMessageId);
            #endregion

            #region Prepare rules' data
            ReplyActionDataOfExtendedRule ruleActionData = new ReplyActionDataOfExtendedRule();
            ruleActionData.MessageEIDSize = 0x46;
            ruleActionData.ReplyTemplateMessageEID = messageEntryId;
            ruleActionData.ReplyTemplateGUID = guidByte;
            #endregion

            #region TestUser1 creates an FAI message for the OP_OOF_REPLY extended rule.
            RopCreateMessageResponse ropCreateMessageResponse;
            uint extendedRuleMessageHandle1 = this.OxoruleAdapter.RopCreateMessage(this.InboxFolderHandle, this.InboxFolderID, Convert.ToByte(true), out ropCreateMessageResponse);
            Site.Assert.AreEqual<uint>(0, ropCreateMessageResponse.ReturnValue, "Creating the FAI message should succeed.");

            NamedPropertyInfo namedPropertyInfo1 = new NamedPropertyInfo
            {
                NoOfNamedProps = 0
            };
            TaggedPropertyValue[] extendedRuleProperties1 = AdapterHelper.GenerateExtendedRuleTestData(ruleProperties.Name, 0, (uint)RuleState.ST_ENABLED, Constants.PidTagRuleProvider, ActionType.OP_OOF_REPLY, ruleActionData, ruleProperties.ConditionSubjectName, namedPropertyInfo1);

            // Set properties for extended rule FAI message.
            RopSetPropertiesResponse ropSetPropertiesResponse = this.OxoruleAdapter.RopSetProperties(extendedRuleMessageHandle1, extendedRuleProperties1);
            Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "Setting property for Extended rule FAI message should succeed.");

            // Save changes of message.
            RopSaveChangesMessageResponse ropSaveChangesMessagResponse = this.OxoruleAdapter.RopSaveChangesMessage(extendedRuleMessageHandle1);
            Site.Assert.AreEqual(0, (int)ropSaveChangesMessagResponse.ReturnValue, "Saving Extend rule message should succeed.");
            #endregion

            #region TestUser1 gets the OP_OOF_REPLY extended rule.
            #region Step1: TestUser1 gets a table of all messages which are placed in the Inbox folder.

            ropGetContentsTableResponseOfFAIMessage = this.OxoruleAdapter.RopGetContentsTable(this.InboxFolderHandle, ContentTableFlag.Associated, out contentsTableHandleOfFAIMessage);
            Site.Assert.AreEqual<uint>(0, ropGetContentsTableResponseOfFAIMessage.ReturnValue, "Getting DAF contents table should succeed");
            #endregion

            #region Step2: TestUser1 sets the interested columns of the message table in the Inbox folder.

            // Here are 6 interested columns listed as below.
            propertyTagOfFAIMessage = new PropertyTag[6];
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
            ropQueryRowsResponseOfFAIMessage = this.OxoruleAdapter.QueryPropertiesInTable(contentsTableHandleOfFAIMessage, propertyTagOfFAIMessage);
            Site.Assert.AreEqual<uint>(0, ropQueryRowsResponseOfFAIMessage.ReturnValue, "Querying Rows Response of FAI Message should succeed, the actual returned value is {0}", ropQueryRowsResponseOfFAIMessage.ReturnValue);
            ExtendedRuleActions extendedRuleAction = new ExtendedRuleActions();
            ReplyActionDataOfExtendedRule replyActionData = new ReplyActionDataOfExtendedRule();
            for (int i = 0; i < ropQueryRowsResponseOfFAIMessage.RowCount; i++)
            {
                // Since the PidTagMessageClass property of Extended rule MUST have a value of "IPM.ExtendedRule.Message", use PidTagMessageClass property to get the extended rule data.
                System.Text.UnicodeEncoding converter = new UnicodeEncoding();
                string messageClass = converter.GetString(ropQueryRowsResponseOfFAIMessage.RowData.PropertyRows.ToArray()[i].PropertyValues[1].Value);
                if (messageClass == "IPM.ExtendedRule.Message" + "\0")
                {
                    byte[] extendedRuleMessageActionBinary = ropQueryRowsResponseOfFAIMessage.RowData.PropertyRows[i].PropertyValues[4].Value;
                    byte[] extendedRuleMessageActionBuffer = new byte[extendedRuleMessageActionBinary.Length - 2];

                    // Remove the two length bytes to get the extended rule action data.
                    Array.Copy(extendedRuleMessageActionBinary, 2, extendedRuleMessageActionBuffer, 0, extendedRuleMessageActionBinary.Length - 2);
                    extendedRuleAction.Deserialize(extendedRuleMessageActionBuffer);
                    replyActionData.Deserialize(extendedRuleAction.RuleActionBuffer.Actions[0].ActionDataValue.Serialize());
                    break;
                }
            }
            #endregion

            #region Capture Code

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R998: the Message EntryID is {0}, and the ReplyTemplateMessageEID in the action data is {1}", messageEntryId, replyActionData.ReplyTemplateMessageEID);

            bool isVerifiedR998 = Common.CompareByteArray(ruleActionData.ReplyTemplateMessageEID, replyActionData.ReplyTemplateMessageEID);

            // Verify MS-OXORULE requirement: MS-OXORULE_R998.
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR998,
                998,
                @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Extended Rules] ReplyTemplateMessageEID (70 bytes): A Message EntryID structure, as specified in [MS-OXCDATA] section 2.2.4.2, that contains the entry ID of the reply template.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R1000");

            bool isVerifiedR1000 = Common.CompareByteArray(ruleActionData.ReplyTemplateGUID, replyActionData.ReplyTemplateGUID);

            // Verify MS-OXORULE requirement: MS-OXORULE_R1000.
            // IF the value of the ReplyTemplateGUID field in OP_OOF_REPLY action data is equal to the value of the PidTagReplyTemplateId property that is set on the reply template, R1000 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR1000,
                1000,
                @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Extended Rules] ReplyTemplateGUID (16 bytes): A GUID that is generated by the client in the process of creating a reply template.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R1002: the value of PidTagReplyTemplateId property is {0}, and the ReplyTemplateGUID in the action data is {1}", pidTagReplyTemplateIdValue, replyActionData.ReplyTemplateGUID);

            bool isVerifiedR1002 = isVerifiedR998 && isVerifiedR1000 && replyActionData.MessageEIDSize == ruleActionData.MessageEIDSize;

            // Verify MS-OXORULE requirement: MS-OXORULE_R1002.
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR1002,
                1002,
                @"[OP_REPLY and OP_OOF_REPLY ActionData Structure] [Buffer Format for Extended Rules] The value of the ReplyTemplateGUID field in OP_OOF_REPLY action data is equal to the value of the PidTagReplyTemplateId property that is set on the reply template.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify get rule table with invalid parameters. 
        /// </summary>
        [TestCategory("MSOXORULE"), TestMethod()]
        public void MSOXORULE_S01_TC12_GetRuleTableWithInvalidParameters()
        {
            this.CheckMAPIHTTPTransportSupported();

            #region TestUser1 prepares value for ruleProperties variable.
            RuleProperties ruleProperties = AdapterHelper.GenerateRuleProperties(this.Site, Constants.RuleNameMarkAsRead);
            #endregion

            #region Generate test RuleData.
            RuleData ruleForMarkRead = AdapterHelper.GenerateValidRuleData(ActionType.OP_MARK_AS_READ, TestRuleDataType.ForAdd, 1, RuleState.ST_ENABLED, new DeleteMarkReadActionData(), ruleProperties, null);
            #endregion

            #region TestUser1 gets the returned value from RopModifyRules response.
            RopModifyRulesResponse responseOfModifyRules = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_ReplaceAll, new RuleData[] { ruleForMarkRead });
            Site.Assert.AreEqual<uint>(0, responseOfModifyRules.ReturnValue, "Add the Mark_As_Read rule should succeed.");
            #endregion

            #region TestUser1 calls RopGetRulesTable with invalid TableFlags and server returns an error.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Invalid, out ropGetRulesTableResponse);

            if (Common.IsRequirementEnabled(800, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R800");

                // Verify MS-OXORULE requirement: MS-OXORULE_R800.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x00000000,
                    ropGetRulesTableResponse.ReturnValue,
                    800,
                    @"[In Appendix A: Product Behavior] Implementation does ignore the x bits and does not return an error for nonzero values. [<3> Section 2.2.2.1: Exchange 2007 ignores the x bits and does not return an error for nonzero values.]");
            }

            if (Common.IsRequirementEnabled(889, this.Site))
            {
                #region Capture Code
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R889");

                // Verify MS-OXORULE requirement: MS-OXORULE_R889.
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    0x00000000,
                    ropGetRulesTableResponse.ReturnValue,
                    889,
                    @"[In RopGetRulesTable ROP Request Buffer] [TableFlags] x: Implementation does return an error if these bits are nonzero but can ignore them. (Exchange 2003, Exchange 2010 and above follow this behavior.)");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R826");

                // Verify MS-OXORULE requirement: MS-OXORULE_R826
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040102,
                    ropGetRulesTableResponse.ReturnValue,
                    826,
                    @"[In Receiving a RopGetRulesTable ROP Request] The value of error code ecNotSupported: 0x80040102.");
                #endregion
            }
            #endregion

            #region TestUser1 calls RopGetRulesTable with invalid folder handle and server returns an error.
            uint invalidFolderHandler = uint.Parse(Constants.InvalidateFolderHandler);
            this.OxoruleAdapter.RopGetRulesTable(invalidFolderHandler, TableFlags.Normal, out ropGetRulesTableResponse);
            Site.Assert.AreNotEqual((uint)0x00000000, ropGetRulesTableResponse.ReturnValue, "The return value should not be 0x00000000, actual return value is {0}.", ropGetRulesTableResponse.ReturnValue);
            #endregion
        }

        #region Private method

        /// <summary>
        /// Get the bytes value of one property from a list of properties.
        /// </summary>
        /// <param name="propertyName">The property need to get value.</param>
        /// <param name="propertyRow">The property row.</param>
        /// <param name="propertyTags">The properties tags.</param>
        /// <returns>The bytes value of the property.</returns>
        private byte[] GetPropertyFromList(PropertyId propertyName, PropertyRow propertyRow, PropertyTag[] propertyTags)
        {
            byte[] value = new byte[0];
            TaggedPropertyValue[] allProperties = null;
            if (propertyRow != null && propertyRow.PropertyValues != null && propertyRow.PropertyValues.Count > 0)
            {
                allProperties = new TaggedPropertyValue[propertyTags.Length];
                for (int i = 0; i < propertyTags.Length; i++)
                {
                    allProperties[i] = new TaggedPropertyValue
                    {
                        PropertyTag = propertyTags[i],
                        Value = propertyRow.PropertyValues[i].Value
                    };
                }
            }

            foreach (TaggedPropertyValue taggedValue in allProperties)
            {
                if (taggedValue.PropertyTag.PropertyId == (ushort)propertyName)
                {
                    value = taggedValue.Value;
                    break;
                }
            }

            return value;
        }

        /// <summary>
        /// Generate PropertyTag arrays for rule properties.
        /// </summary>
        /// <returns>PropertyTag arrays for rule properties.</returns>
        private PropertyTag[] GenerateRuleInfoProperties()
        {
            PropertyTag[] propertyTags = new PropertyTag[4];

            // PidTagRuleName
            propertyTags[0].PropertyId = (ushort)PropertyId.PidTagRuleName;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypString;

            // PidTagRuleUserFlags
            propertyTags[1].PropertyId = (ushort)PropertyId.PidTagRuleUserFlags;
            propertyTags[1].PropertyType = (ushort)PropertyType.PtypInteger32;

            // PidTagRuleProviderData
            propertyTags[2].PropertyId = (ushort)PropertyId.PidTagRuleProviderData;
            propertyTags[2].PropertyType = (ushort)PropertyType.PtypBinary;

            // PidTagRuleId
            propertyTags[3].PropertyId = (ushort)PropertyId.PidTagRuleId;
            propertyTags[3].PropertyType = (ushort)PropertyType.PtypInteger64;

            return propertyTags;
        }
        #endregion
    }
}