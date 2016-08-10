namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is designed to verify the response buffer formats of Rule ROPs. 
    /// </summary>
    [TestClass]
    public class S09_RuleROPs : TestSuiteBase
    {
        #region Class Initialization and Cleanup

        /// <summary>
        /// Class initialize.
        /// </summary>
        /// <param name="testContext">The session context handle</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Class cleanup.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Cases

        /// <summary>
        /// This method tests the ROP buffers of RopModifyRules, RopGetRulesTable and RopUpdateDeferredActionMessages.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S09_TC01_TestRopModifyRules()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Preparations-Open folder and get it's handle.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Call GetFolderObjectHandle method to get folder object handle..");

            uint folderHandle = GetFolderObjectHandle(ref logonResponse);

            // Step 2: Send the RopModifyRules request and verify success response.
            #region RopModifyRules

            RopModifyRulesRequest modifyRulesRequest;
            RopModifyRulesResponse modifyRulesResponse;
            RuleData[] sampleRuleDataArray;

            modifyRulesRequest.RopId = (byte)RopId.RopModifyRules;

            modifyRulesRequest.LogonId = TestSuiteBase.LogonId;
            modifyRulesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Call CreateSampleRuleDataArrayForAdd method to create Sample RuleData Array.
            sampleRuleDataArray = this.CreateSampleRuleDataArrayForAdd();

            modifyRulesRequest.ModifyRulesFlags = (byte)ModifyRulesFlags.None;
            modifyRulesRequest.RulesCount = (ushort)sampleRuleDataArray.Length;
            modifyRulesRequest.RulesData = sampleRuleDataArray;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopModifyRules request.");

            // Send the RopModifyRules request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                modifyRulesRequest,
                folderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            modifyRulesResponse = (RopModifyRulesResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                modifyRulesResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success0");

            #endregion

            // Step 3: Send the RopGetRulesTable request and verify success response.
            #region RopGetRulesTable

            RopGetRulesTableRequest getRulesTableRequest;
            RopGetRulesTableResponse getRulesTableResponse;

            getRulesTableRequest.RopId = (byte)RopId.RopGetRulesTable;
            getRulesTableRequest.LogonId = TestSuiteBase.LogonId;
            getRulesTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            getRulesTableRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            getRulesTableRequest.TableFlags = (byte)TableFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopGetRulesTable request.");

            // Send the RopGetRulesTable request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getRulesTableRequest,
                folderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getRulesTableResponse = (RopGetRulesTableResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getRulesTableResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success0");

            #endregion

            // Step 4: Send the RopUpdateDeferredActionMessages request and verify success response.
            #region RopUpdateDeferredActionMessages

            RopUpdateDeferredActionMessagesRequest updateDeferredActionMessagesRequest;

            updateDeferredActionMessagesRequest.RopId = (byte)RopId.RopUpdateDeferredActionMessages;
            updateDeferredActionMessagesRequest.LogonId = TestSuiteBase.LogonId;
            updateDeferredActionMessagesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set ServerEntryIdSize,which specifies the size of the ServerEntryId field.
            updateDeferredActionMessagesRequest.ServerEntryIdSize = TestSuiteBase.ServerEntryIdSize;

            byte[] serverId = { ServerEntryId };
            updateDeferredActionMessagesRequest.ServerEntryId = serverId;

            // Set ClientEntryIdSize,which specifies the size of the ClientEntryId field.
            updateDeferredActionMessagesRequest.ClientEntryIdSize = TestSuiteBase.ClientEntryIdSize;

            byte[] clientId = { ClientEntryId };
            updateDeferredActionMessagesRequest.ClientEntryId = clientId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopUpdateDeferredActionMessages request.");

            // Send the RopUpdateDeferredActionMessages request.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                updateDeferredActionMessagesRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion
        }

        #endregion

        #region Common method

        /// <summary>
        /// Create Sample RuleData Array For Add.
        /// </summary>
        /// <returns>Return RuleData array</returns>
        private RuleData[] CreateSampleRuleDataArrayForAdd()
        {
            // Count of PropertyValues.
            int length = 4;
            TaggedPropertyValue[] propertyValues = new TaggedPropertyValue[length];

            for (int i = 0; i < length; i++)
            {
                propertyValues[i] = new TaggedPropertyValue();
            }

            TaggedPropertyValue taggedPropertyValue;

            // MS-OXORULE 2.2.1.3.2
            // When adding a rule, the client MUST NOT pass in PidTagRuleId, it MUST pass in PidTagRuleCondition,
            // PidTagRuleActions and PidTagRuleProvider.

            // PidTagRuleSequence
            taggedPropertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagRuleSequence].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagRuleSequence].PropertyType
                }
            };
            byte[] value3 = { 0x00, 0x00, 0x00, 0x0a };
            taggedPropertyValue.Value = value3;
            propertyValues[3] = taggedPropertyValue;

            // PidTagRuleCondition
            taggedPropertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagRuleCondition].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagRuleCondition].PropertyType
                }
            };
            byte[] value = 
            {
                0x03, 0x01, 0x00, 0x01, 0x00, 0x1f, 0x00, 0x37, 0x00, 0x1f, 0x00,
                0x37, 0x00, 0x50, 0x00, 0x72, 0x00, 0x6f, 0x00, 0x6a, 0x00, 0x65,
                0x00, 0x63, 0x00, 0x74, 0x00, 0x20, 0x00, 0x58, 0x00, 0x00, 0x00
            };
            taggedPropertyValue.Value = value;
            propertyValues[1] = taggedPropertyValue;

            // PidTagRuleActions
            taggedPropertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagRuleActions].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagRuleActions].PropertyType
                }
            };
            byte[] value1 = { 0x01, 0x00, 0x09, 0x00, 0x0B, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
            taggedPropertyValue.Value = value1;
            propertyValues[2] = taggedPropertyValue;

            // PidTagRuleProvider
            taggedPropertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagRuleProvider].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagRuleProvider].PropertyType
                }
            };
            byte[] value2 = Encoding.Unicode.GetBytes("RuleOrganizerContoso\0");
            taggedPropertyValue.Value = value2;
            propertyValues[0] = taggedPropertyValue;

            RuleData sampleRuleData = new RuleData
            {
                RuleDataFlags = (byte)RuleDataFlags.RowAdd,
                PropertyValueCount = (ushort)propertyValues.Length,
                PropertyValues = propertyValues
            };

            RuleData[] sampleRuleDataArray = new RuleData[1];
            sampleRuleDataArray[0] = sampleRuleData;

            return sampleRuleDataArray;
        }

        #endregion 
    }
}