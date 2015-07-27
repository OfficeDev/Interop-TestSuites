//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASPROV
{
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to test the negative status of Provision command.
    /// </summary>
    [TestClass]
    public class S03_ProvisionNegative : TestSuiteBase
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
        /// This test case is intended to validate Status 3 of Policy element.
        /// </summary>
        [TestCategory("MSASPROV"), TestMethod()]
        public void MSASPROV_S03_TC01_VerifyPolicyStatus3()
        {
            #region Call Provision command with invalid policy type.
            // Assign an invalid policy type in the provision request
            string invalidType = "InvalidMS-EAS-Provisioning-WBXML";
            ProvisionResponse provisionResponse = this.CallProvisionCommand(string.Empty, invalidType, "1");

            byte policyStatus = provisionResponse.ResponseData.Policies.Policy.Status;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R380");

            // Verify MS-ASPROV requirement: MS-ASPROV_R380
            // The Status of Policy element is 3, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                3,
                policyStatus,
                380,
                @"[In Status (Policy)] Value 3 means Unknown PolicyType value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R471");

            // Verify MS-ASPROV requirement: MS-ASPROV_R471
            // The Status of Policy element is 3, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                3,
                policyStatus,
                471,
                @"[In Provision Command Errors] [The cause of status value 3 is] The client sent a policy that the server does not recognize.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R502");

            // Verify MS-ASPROV requirement: MS-ASPROV_R502
            // The Status of Policy element is 3, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                3,
                policyStatus,
                502,
                @"[In Provision Command Errors] [When the scope is Policy], [the cause of status value 3 is] The client sent a policy that the server does not recognize.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate Status 5 of Policy element.
        /// </summary>
        [TestCategory("MSASPROV"), TestMethod()]
        public void MSASPROV_S03_TC02_VerifyPolicyStatus5()
        {
            #region Download the policy settings.
            // Download the policy settings.
            ProvisionResponse provisionResponse = this.CallProvisionCommand(string.Empty, "MS-EAS-Provisioning-WBXML", "1");
            string temporaryPolicyKey = provisionResponse.ResponseData.Policies.Policy.PolicyKey;
            #endregion

            #region Acknowledge the policy settings.
            // Acknowledge the policy settings.
            this.CallProvisionCommand(temporaryPolicyKey, "MS-EAS-Provisioning-WBXML", "1");
            #endregion

            #region Switch current user to the user who has custom policy settings.
            // Switch to the user who has been configured with custom policy.
            this.SwitchUser(this.User2Information, false);

            #endregion

            #region Call Provision command with out-of-date PolicyKey.
            provisionResponse = this.CallProvisionCommand(temporaryPolicyKey, "MS-EAS-Provisioning-WBXML", "1");

            byte policyStatus = provisionResponse.ResponseData.Policies.Policy.Status;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R475");

            // Verify MS-ASPROV requirement: MS-ASPROV_R475
            // The Status of Policy element is 5, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                5,
                policyStatus,
                475,
                @"[In Provision Command Errors] [The cause of status value 5 is] The client is trying to acknowledge an out-of-date [or invalid policy].");
            #endregion

            #region Call Provision command with invalid PolicyKey.
            provisionResponse = this.CallProvisionCommand("1234567890", "MS-EAS-Provisioning-WBXML", "1");

            policyStatus = provisionResponse.ResponseData.Policies.Policy.Status;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R761");

            // Verify MS-ASPROV requirement: MS-ASPROV_R761
            // The Status of Policy element is 5, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                5,
                policyStatus,
                761,
                @"[In Provision Command Errors] [The cause of status value 5 is] The client is trying to acknowledge an [out-of-date or] invalid policy.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R382");

            // Verify MS-ASPROV requirement: MS-ASPROV_R382
            // The Status of Policy element is 5, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                5,
                policyStatus,
                382,
                @"[In Status (Policy)] Value 5 means The client is acknowledging the wrong policy key.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R692");

            // Verify MS-ASPROV requirement: MS-ASPROV_R692
            // The Status of Policy element is 5, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                5,
                policyStatus,
                692,
                @"[In Provision Command Errors] [The meaning of status value] 5 [is] Policy key mismatch.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R507");

            // Verify MS-ASPROV requirement: MS-ASPROV_R507
            // The Status of Policy element is 5, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                5,
                policyStatus,
                507,
                @"[In Provision Command Errors] [When the scope is Policy], [the cause of status value 5 is] The client is trying to acknowledge an out-of-date or invalid policy.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R708");

            // Verify MS-ASPROV requirement: MS-ASPROV_R708
            // The Status of Policy element is 5, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                5,
                policyStatus,
                708,
                @"[In Provision Command Errors] [When the scope is] Policy, [the meaning of status value] 5 [is] Policy key mismatch.");

            if (Common.IsRequirementEnabled(695, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R695");

                // Verify MS-ASPROV requirement: MS-ASPROV_R695
                // The Status of Policy element is 5, so this requirement can be captured.
                Site.CaptureRequirementIfAreEqual<byte>(
                    5,
                    policyStatus,
                    695,
                    @"[In Appendix B: Product Behavior] If it does not [current policy key sent by the client in a security policy settings acknowledgment does not match the temporary policy key issued by the server in the response to the initial request from this client], the implementation does return a Status (section 2.2.2.53.2) value of 5, as specified in section 3.2.5.2. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate Status 2 of Provision element.
        /// </summary>
        [TestCategory("MSASPROV"), TestMethod()]
        public void MSASPROV_S03_TC03_VerifyProvisionStatus2()
        {
            #region Create a Provision request with syntax error.
            ProvisionRequest provisionRequest = Common.CreateProvisionRequest(null, new Request.ProvisionPolicies(), null);
            Request.ProvisionPoliciesPolicy policy = new Request.ProvisionPoliciesPolicy
            {
                PolicyType = "MS-EAS-Provisioning-WBXML"
            };

            // The format in which the policy settings are to be provided to the client device.
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.1")
            {
                // Configure the DeviceInformation.
                Request.DeviceInformation deviceInfomation = new Request.DeviceInformation();
                Request.DeviceInformationSet deviceInformationSet = new Request.DeviceInformationSet
                {
                    Model = "ASPROVTest"
                };
                deviceInfomation.Set = deviceInformationSet;
                provisionRequest.RequestData.DeviceInformation = deviceInfomation;
            }

            provisionRequest.RequestData.Policies.Policy = policy;
            string requestBody = provisionRequest.GetRequestDataSerializedXML();
            requestBody = requestBody.Replace(@"<Policies>", string.Empty);
            requestBody = requestBody.Replace(@"</Policies>", string.Empty);
            #endregion

            #region Call Provision command and get the Status of response.
            ProvisionResponse provisionResponse = this.PROVAdapter.SendProvisionStringRequest(requestBody);
            byte provisionStatus = provisionResponse.ResponseData.Status;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R395");

            // Verify MS-ASPROV requirement: MS-ASPROV_R395
            // The Status of Provision element is 2, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                provisionStatus,
                395,
                @"[In Status (Provision)] Value 2 means Protocol error.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R679");

            // Verify MS-ASPROV requirement: MS-ASPROV_R679
            // The Status of Provision element is 2, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                provisionStatus,
                679,
                @"[In Provision Command Errors] [The meaning of status value] 2 [is] Protocol error.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R450");

            // Verify MS-ASPROV requirement: MS-ASPROV_R450
            // The Status of Provision element is 2, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                provisionStatus,
                450,
                @"[In Provision Command Errors] [The cause of status value 2 is] Syntax error in the Provision command request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R497");

            // Verify MS-ASPROV requirement: MS-ASPROV_R497
            // The Status of Provision element is 2, so this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                provisionStatus,
                497,
                @"[In Provision Command Errors] [When the scope is Global], [the cause of status value 2 is] Syntax error in the Provision command request.");

            if (Common.IsRequirementEnabled(697, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R697");

                // Verify MS-ASPROV requirement: MS-ASPROV_R697
                // Status 2 is returned when there is syntax error in the Provision command request, so this requirement can be captured.
                Site.CaptureRequirement(
                    697,
                    @"[In Appendix B: Product Behavior] If the level of compliance does not meet the server's requirements, the implementation does return an appropriate value in the Status (section 2.2.2.53.2) element. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the status code when the policy key is invalid.
        /// </summary>
        [TestCategory("MSASPROV"), TestMethod()]
        public void MSASPROV_S03_TC04_VerifyInvalidPolicyKey()
        {
            #region Call Provision command to download the policy settings.
            // Download the policy setting.
            ProvisionResponse provisionResponse = this.CallProvisionCommand(string.Empty, "MS-EAS-Provisioning-WBXML", "1");
            string temporaryPolicyKey = provisionResponse.ResponseData.Policies.Policy.PolicyKey;
            #endregion

            #region Call Provision command to acknowledge the policy settings and get the valid PolicyKey
            // Acknowledge the policy setting.
            provisionResponse = this.CallProvisionCommand(temporaryPolicyKey, "MS-EAS-Provisioning-WBXML", "1");

            string finalPolicyKey = provisionResponse.ResponseData.Policies.Policy.PolicyKey;
            #endregion

            #region Call FolderSync command with an invalid PolicyKey which is different from the one got from last step.
            // Apply an invalid policy key
            this.PROVAdapter.ApplyPolicyKey(finalPolicyKey.Substring(0, 1));

            // Call folder sync with "0" in initialization phase.
            FolderSyncRequest folderSyncRequest = Common.CreateFolderSyncRequest("0");

            if ("12.1" == Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site))
            {
                string httpErrorCode = null;
                try
                {
                    this.PROVAdapter.FolderSync(folderSyncRequest);
                }
                catch (WebException exception)
                {
                    httpErrorCode = Common.GetErrorCodeFromException(exception);
                }

                Site.Assert.IsFalse(string.IsNullOrEmpty(httpErrorCode), "Server should return expected [449] error code if client do not have policy key");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R511");

                // Verify MS-ASPROV requirement: MS-ASPROV_R511
                // HTTP 449 status is returned when the MS-ASProtocolVersion header is set to 12.1, so this requirement is captured.
                Site.CaptureRequirementIfAreEqual(
                    "449",
                    httpErrorCode,
                    511,
                    @"[In Appendix A: Product Behavior] <2> Section 3.1.5.1: When the MS-ASProtocolVersion header is set to 12.1, the server sends an HTTP 449 response to request a Provision command from the client.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R514");

                // Verify MS-ASPROV requirement: MS-ASPROV_R514
                // HTTP 449 status is returned when the MS-ASProtocolVersion header is set to 12.1, so this requirement is captured.
                Site.CaptureRequirementIfAreEqual(
                    "449",
                    httpErrorCode,
                    514,
                    @"[In Appendix A: Product Behavior] <3> Section 3.2.5.1: When the MS-ASProtocolVersion header is set to 12.1, the server sends an HTTP 449 response to indicate that the client needs to request the security policy settings and obtain a new policy key.");
            }
            else
            {
                FolderSyncResponse folderSyncResponse = this.PROVAdapter.FolderSync(folderSyncRequest);
                Site.Assert.AreEqual(144, folderSyncResponse.ResponseData.Status, "The server should return status 144 to indicate a invalid policy key.");
            }

            if (Common.IsRequirementEnabled(537, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R537");

                // Verify MS-ASPROV requirement: MS-ASPROV_R537
                // If the above capture or assert passed, it means the server did returns a status code when the policy key is mismatched.
                Site.CaptureRequirement(
                    537,
                    @"[In Appendix A: Product Behavior] If the policy key sent by the client does not match the stored policy key, the implementation does return a status code in the next command response indicating that the client needs to send another Provision command to request the security policy settings and obtain a new policy key. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion
        }
    }
}