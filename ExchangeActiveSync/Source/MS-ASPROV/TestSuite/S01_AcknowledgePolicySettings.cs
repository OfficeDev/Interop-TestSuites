namespace Microsoft.Protocols.TestSuites.MS_ASPROV
{
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test the acknowledge phase of Provision command.
    /// </summary>
    [TestClass]
    public class S01_AcknowledgePolicySettings : TestSuiteBase
    {
        #region Class initialize and clean up
        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        /// <summary>
        /// This test case is intended to validate the acknowledgement phase of Provision.
        /// </summary>
        [TestCategory("MSASPROV"), TestMethod()]
        public void MSASPROV_S01_TC01_AcknowledgeSecurityPolicySettings()
        {
            #region Switch current user to the user who has custom policy settings.
            // Switch to the user who has been configured with custom policy.
            this.SwitchUser(this.User2Information, false);
            #endregion

            #region Download the policy settings.
            // Download the policy settings.
            ProvisionResponse provisionResponse = this.CallProvisionCommand(string.Empty, "MS-EAS-Provisioning-WBXML", "1");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R394");

            // Verify MS-ASPROV requirement: MS-ASPROV_R394
            // The value of Status is 1, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                provisionResponse.ResponseData.Status,
                394,
                @"[In Status (Provision)] Value 1 means Success.");

            string temporaryPolicyKey = provisionResponse.ResponseData.Policies.Policy.PolicyKey;

            // Get the policy element from the Provision response.
            Response.ProvisionPoliciesPolicy policy = provisionResponse.ResponseData.Policies.Policy;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R310");

            // Verify MS-ASPROV requirement: MS-ASPROV_R310
            // The PolicyType, PolicyKey, Status and Data elements are not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                policy.Data != null && policy.PolicyKey != null && policy.PolicyType != null && policy.Status != 0,
                310,
                @"[In Policy] In the initial Provision command response, the Policy element has only the following child elements: PolicyType (section 2.2.2.43) (required) PolicyKey (section 2.2.2.42) (required) Status (section 2.2.2.54.1) (required) Data (section 2.2.2.24) ( required).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R417");

            // Verify MS-ASPROV requirement: MS-ASPROV_R417
            // The PolicyType, PolicyKey, Status and Data elements are not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                policy.Data != null && policy.PolicyKey != null && policy.PolicyType != null && policy.Status != 0,
                417,
                @"[In Abstract Data Model] In order 1, the server response contains the policy type, policy key, data, and status code.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R378");

            // Verify MS-ASPROV requirement: MS-ASPROV_R378
            // The value of Status is 1, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                policy.Status,
                378,
                @"[In Status (Policy)] Value 1 means Success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R209");

            // Verify MS-ASPROV requirement: MS-ASPROV_R209
            // The Data element is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                policy.Data,
                209,
                @"[In Data (container Data Type)] It [Data element] is a required child element of the Policy element (section 2.2.2.41) in responses to initial Provision command requests, as specified in section 3.2.5.1.1.");

            // Because the user is not allowe to download attachment.
            // So if AttachmentsEnabled element is false then R206 will be verified.
            this.Site.CaptureRequirementIfAreEqual<bool>(
                false,
                policy.Data.EASProvisionDoc.AttachmentsEnabled,
                206,
                @"[In AttachmentsEnabled] Value 0 means Attachments are not allowed to be downloaded.");

            // Because if the Data element is a container Data type, then the Data element should contain the EASProvisionDoc element.
            // So if the Data element contain the EASProvisionDoc element and the PolicyType element is set to "MS-EAS-Provisioning-WBXML", then R888 will be verified.
            this.Site.CaptureRequirementIfIsTrue(
                policy.Data.EASProvisionDoc != null && policy.PolicyType == "MS-EAS-Provisioning-WBXML",
                888,
                @"[In Data (container Data Type)] This element [Data (container Data Type)] requires that the PolicyType element (section 2.2.2.43) is set to ""MS-EAS-Provisioning-WBXML"".");
            
            this.Site.CaptureRequirementIfIsTrue(
                !string.IsNullOrEmpty(policy.PolicyKey),
                486,
                @"[In Provision Command] The server generates, stores, and sends the policy key when it responds to a Provision command request for policy settings.");
            #endregion

            #region Acknowledge the policy settings.
            // Acknowledge the policy settings.
            provisionResponse = this.CallProvisionCommand(temporaryPolicyKey, "MS-EAS-Provisioning-WBXML", "1");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R678");

            // Verify MS-ASPROV requirement: MS-ASPROV_R678
            // The value of Status is 1, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                provisionResponse.ResponseData.Status,
                678,
                @"[In Provision Command Errors] [The meaning of status value] 1 [is] Success.");

            bool isR441Verified = provisionResponse.ResponseData.Policies != null && provisionResponse.ResponseData.Status == 1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R441");

            // Verify MS-ASPROV requirement: MS-ASPROV_R441
            // The Policies element is not null and the value of Status is 1, so this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isR441Verified,
                441,
                @"[In Provision Command Errors] [The cause of status value 1 is] The Policies element contains information about security policies.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R650");

            // Verify MS-ASPROV requirement: MS-ASPROV_R650
            // The acknowledgement Provision succeeds and PolicyKey element is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                temporaryPolicyKey,
                650,
                @"[In Responding to an Initial Request] The value of the PolicyKey element (section 2.2.2.42) is a temporary policy key that will be valid only for an acknowledgment request to acknowledge the policy settings contained in the Data element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R657");

            // Verify MS-ASPROV requirement: MS-ASPROV_R657
            // The command executed successfully using the temporary PolicyKey, so this requirement can be captured.
            Site.CaptureRequirement(
                657,
                @"[In Responding to a Security Policy Settings Acknowledgment] The server MUST ensure that the current policy key sent by the client in a security policy settings acknowledgment matches the temporary policy key issued by the server in the response to the initial request from this client.");

            // Get the policy element from the Provision response.
            policy = provisionResponse.ResponseData.Policies.Policy;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R605");

            // Verify MS-ASPROV requirement: MS-ASPROV_R605
            // The PolicyType, PolicyKey and Status elements are not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                policy.PolicyKey != null && policy.PolicyType != null && policy.Status != 0,
                605,
                @"[In Policy] In the acknowledgment Provision command response, the Policy element has the following child elements: PolicyType (section 2.2.2.43) (required) PolicyKey (section 2.2.2.42) (required) Status (section 2.2.2.54.1) (required).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R419");

            // Verify MS-ASPROV requirement: MS-ASPROV_R419
            // The PolicyType, PolicyKey and Status elements are not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                policy.PolicyKey != null && policy.PolicyType != null && policy.Status != 0,
                419,
                @"[In Abstract Data Model] In order 2, the server response contains the policy type, policy key, and status code to indicate that the server recorded the client's acknowledgement.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R688");

            // Verify MS-ASPROV requirement: MS-ASPROV_R688
            // The value of Status is 1, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                policy.Status,
                688,
                @"[In Provision Command Errors] [The meaning of status value] 1 [is] Success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R466");

            // Verify MS-ASPROV requirement: MS-ASPROV_R466
            // The value of Status is 1, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                policy.Status,
                466,
                @"[In Provision Command Errors] [The cause of status value 1 is] The requested policy data is included in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R703");

            // Verify MS-ASPROV requirement: MS-ASPROV_R703
            // The value of Status is 1, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                policy.Status,
                703,
                @"[In Provision Command Errors] [When the scope is] Policy, [the meaning of status value] 1 [is] Success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R495");

            // Verify MS-ASPROV requirement: MS-ASPROV_R495
            // The value of Status is 1, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                policy.Status,
                495,
                @"[In Provision Command Errors] [When the scope is Policy], [the cause of status value 1 is] The requested policy data is included in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R601");

            // Verify MS-ASPROV requirement: MS-ASPROV_R601
            // The Data element is null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNull(
                policy.Data,
                601,
                @"[In Data (container Data Type)] It [Data element] is not present in responses to acknowledgment requests, as specified in section 3.2.5.1.2.");
            #endregion

            #region Apply the final policy key got from acknowledgement Provision response.
            // Get the final policy key from the Provision command response.
            string finalPolicyKey = provisionResponse.ResponseData.Policies.Policy.PolicyKey;
            
            // Apply the final policy key for the subsequence commands.
            this.PROVAdapter.ApplyPolicyKey(finalPolicyKey);
            #endregion

            #region Call FolderSync command with the final policy key.
            FolderSyncRequest folderSyncRequest = Common.CreateFolderSyncRequest("0");
            FolderSyncResponse folderSynReponse = this.PROVAdapter.FolderSync(folderSyncRequest);
            Site.Assert.AreEqual(folderSynReponse.StatusCode, HttpStatusCode.OK, "Server should return a HTTP expected status code [{0}] after apply Policy Key, actual is [{1}]", HttpStatusCode.OK, folderSynReponse.StatusCode);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R662");

            // Verify MS-ASPROV requirement: MS-ASPROV_R662
            // The FolderSync command executed successfully after the final policy key is applied, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                finalPolicyKey,
                662,
                @"[In Responding to a Security Policy Settings Acknowledgment] The value of the PolicyKey element (section 2.2.2.42) is a permanent policy key that is valid for subsequent command requests from the client.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate command could be executed successfully without acknowledging security policy settings if a security policy is set on the implementation to allow it.
        /// </summary>
        [TestCategory("MSASPROV"), TestMethod()]
        public void MSASPROV_S01_TC02_WithoutAcknowledgingSecurityPolicySettings()
        {
            #region Switch the current user to the user with setting AllowNonProvisionableDevices to true.
            this.SwitchUser(this.User3Information, false);
            #endregion

            #region Apply string.Empty to PolicyKey.
            this.PROVAdapter.ApplyPolicyKey(string.Empty);
            #endregion

            #region Call FolderSync command without Provision.
            FolderSyncRequest folderSyncRequest = Common.CreateFolderSyncRequest("0");
            FolderSyncResponse folderSynReponse = this.PROVAdapter.FolderSync(folderSyncRequest);
            Site.Assert.AreEqual(folderSynReponse.StatusCode, HttpStatusCode.OK, "Server should return a HTTP expected status code [{0}], actual is [{1}]", HttpStatusCode.OK, folderSynReponse.StatusCode);

            if (Common.IsRequirementEnabled(509, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R509");

                // Verify MS-ASPROV requirement: MS-ASPROV_R509
                // The FolderSync command executed successfully without Provision, so this requirement can be captured.
                Site.CaptureRequirement(
                    509,
                    @"[In Appendix B: Product Behavior] The implementation does require that the client device has requested and acknowledged the security policy settings before the client is allowed to synchronize with the server, unless a security policy is set on the implementation to allow it [client is allowed to synchronize with the implementation]. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// This test case is intended to test when a policy setting that was previously set is unset on the server.
        /// </summary>
        [TestCategory("MSASPROV"), TestMethod()]
        public void MSASPROV_S01_TC03_InitialPreviouslySettingUnset()
        {
            #region Download the policy settings.
            // Download the policy settings.
            ProvisionResponse provisionResponse = this.CallProvisionCommand(string.Empty, "MS-EAS-Provisioning-WBXML", "1");

            // Because the user is allowe to download attachment.
            // So if AttachmentsEnabled element is true then R207 will be verified.
            this.Site.CaptureRequirementIfAreEqual<bool>(
                true,
                provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.AttachmentsEnabled,
                207,
                @"[In AttachmentsEnabled] Value 1 means Attachments are allowed to be downloaded.");

            if (Common.IsRequirementEnabled(1044, this.Site))
            {
                // Because the MinDevicePasswordLength is unset in previously set on the server.
                // So if the MinDevicePasswordLength element is emtry string then R1044 will be verified.
                this.Site.CaptureRequirementIfIsTrue(
                    string.IsNullOrEmpty(provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.MinDevicePasswordLength),
                    1044,
                    @"[In Appendix B: Product Behavior] When a policy setting that was previously set is unset on the server, the  implementation does specify the element that represents the setting as an empty tag [or a default value]. (Exchange 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1045, this.Site))
            {
                // Because the DevicePasswordHistory is unset in previously set on the server.
                // So if the DevicePasswordHistory element is default value(0) then R1045 will be verified.
                this.Site.CaptureRequirementIfIsTrue(
                    provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.DevicePasswordHistory == 0,
                    1045,
                    @"[In Appendix B: Product Behavior] When a policy setting that was previously set is unset on the server, the  implementation does specify the element that represents the setting as [an empty tag or] a default value. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion
        }
    }
}