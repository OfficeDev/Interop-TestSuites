//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is used to test the Provision command.
    /// </summary>
    [TestClass]
    public class S12_Provision : TestSuiteBase
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

        #region Test cases
        /// <summary>
        /// This test case is used to verify when download policies from server, server should return provision policies and a template policy key, and then acknowledge the policies by using template policy. 
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S12_TC01_Provision_DownloadPolicy()
        {
            #region User calls Provision command to download policies from server
            // Calls Provision command to download policies
            ProvisionRequest provisionRequest = TestSuiteBase.GenerateDefaultProvisionRequest();
            ProvisionResponse provisionResponse = this.CMDAdapter.Provision(provisionRequest);

            // Get policyKey, policyType and statusCode from server response
            string policyKey = GetPolicyKeyFromResponse(provisionResponse);
            string policyType = provisionResponse.ResponseData.Policies.Policy.PolicyType;
            Response.ProvisionPoliciesPolicyData data = provisionResponse.ResponseData.Policies.Policy.Data;
            byte statusCode = provisionResponse.ResponseData.Status;
            #endregion

            #region Verify Requirements MS-ASCMD_R5026, MS-ASCMD_R4990, MS-ASCMD_R4992
            // If User calls Provision command to download policies successful, server will return policyKey, policyType, data and statusCode in response, then MS-ASCMD_R5026, MS-ASCMD_R4990, MS-ASCMD_R4992 are verified.
            // The policy settings with the format specified in PolicyType element, are contained in Data element of Provision command response.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5026");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5026
            Site.CaptureRequirementIfIsTrue(
                policyKey != null && policyType != null && data != null && statusCode == 1,
                5026,
                @"[In Downloading Policy Settings] [Provision sequence for downloading policy settings, order 1:] The server responds with the policy type, policy key, data, and status code.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4990");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4990
            Site.CaptureRequirementIfIsTrue(
                policyKey != null && policyType != null && data != null,
                4990,
                @"[In Downloading Policy Settings] The server then responds with the provision:PolicyType, provision:PolicyKey (as specified in [MS-ASPROV] section 2.2.2.41), and provision:Data ([MS-ASPROV] section 2.2.2.23) elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4992");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4992
            Site.CaptureRequirementIfIsTrue(
                policyType != null && data != null,
                4992,
                @"[In Downloading Policy Settings] The policy settings, in the format specified in the provision:PolicyType element, are contained in the provision:Data element.");
            #endregion

            #region User calls Provision command to acknowledge policies.

            // Set acknowledgeStatus value to 1, means accept the policy.
            string acknowledgeStatus = "1";
            ProvisionRequest provisionAcknowledgeRequest = TestSuiteBase.GenerateDefaultProvisionRequest();
            provisionAcknowledgeRequest.RequestData.Policies.Policy.PolicyKey = policyKey;
            provisionAcknowledgeRequest.RequestData.Policies.Policy.Status = acknowledgeStatus;

            // Calls Provision command
            ProvisionResponse provisionAcknowledgeResponse = this.CMDAdapter.Provision(provisionAcknowledgeRequest);

            // Get policyKey, policyType and status code from server response
            policyKey = GetPolicyKeyFromResponse(provisionAcknowledgeResponse);
            policyType = provisionAcknowledgeResponse.ResponseData.Policies.Policy.PolicyType;
            statusCode = provisionAcknowledgeResponse.ResponseData.Policies.Policy.Status;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5028");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5028
            Site.CaptureRequirementIfIsTrue(
                policyKey != null && policyType != null && statusCode == 1,
                5028,
                @"[In Downloading Policy Settings] [Provision sequence for downloading policy settings, order 2:] The server responds with the policy type, policy key, and status code to indicate that the server recorded the client's acknowledgement.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if Provision command request does not include PolicyType element, the server returns status 2.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S12_TC02_Provision_WithoutPolicyTypeElement()
        {
            #region User calls Provision command to download policies without policy type element in request.
            ProvisionRequest provisionRequest = TestSuiteBase.GenerateDefaultProvisionRequest();
            
            // Set the policy type Element value to null
            provisionRequest.RequestData.Policies.Policy.PolicyType = null;
            ProvisionResponse provisionResponse = this.CMDAdapter.Provision(provisionRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4989");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4989
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                provisionResponse.ResponseData.Status,
                4989,
                @"[In Downloading Policy Settings] If the provision:PolicyType element is not included in the initial Provision command request, the server responds with a provision:Status element value of 2.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if the client sends the Provision command with invalid policy key, server will return status value 144.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S12_TC03_Provision_Status144()
        {
            #region User calls Provision command to download policies from server
            // Calls Provision command to download policies
            ProvisionRequest provisionRequest = TestSuiteBase.GenerateDefaultProvisionRequest();
            ProvisionResponse provisionResponse = this.CMDAdapter.Provision(provisionRequest);
            Site.Assert.AreEqual(1, provisionResponse.ResponseData.Status, "If Provision operation executes successfully, server should return status 1");

            // Get policyKey
            string policyKey = GetPolicyKeyFromResponse(provisionResponse);
            #endregion

            #region User calls Provision command to acknowledge policies.
            // Set acknowledgeStatus value to 1, means accept the policy.
            string acknowledgeStatus = "1";
            ProvisionRequest provisionAcknowledgeRequest = TestSuiteBase.GenerateDefaultProvisionRequest();
            provisionAcknowledgeRequest.RequestData.Policies.Policy.PolicyKey = policyKey;
            provisionAcknowledgeRequest.RequestData.Policies.Policy.Status = acknowledgeStatus;

            // Calls Provision command
            ProvisionResponse provisionAcknowledgeResponse = this.CMDAdapter.Provision(provisionAcknowledgeRequest);
            Site.Assert.AreEqual(1, provisionResponse.ResponseData.Status, "If Provision operation executes successfully, server should return status 1");

            // Get policyKey
            string finalPolicyKey = GetPolicyKeyFromResponse(provisionAcknowledgeResponse);
            #endregion

            #region Call FolderSync command with an invalid PolicyKey which is different from the one got from last step.
            this.CMDAdapter.ChangePolicyKey(finalPolicyKey.Substring(0, 1));
            this.RecordPolicyKeyChanged();
            
            if ("12.1" == Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site))
            {
                string httpErrorCode = null;
                try
                {
                    // Call FolderSync command
                    this.CMDAdapter.FolderSync(Common.CreateFolderSyncRequest("0"));
                }
                catch (WebException exception)
                {
                    httpErrorCode = Common.GetErrorCodeFromException(exception);
                }

                Site.Assert.AreEqual("449", httpErrorCode, "[In MS-ASPROV Appendix A: Product Behavior] <2> Section 3.1.5.1: When the MS-ASProtocolVersion header is set to 12.1, the server sends an HTTP 449 response to request a Provision command from the client.");
            }
            else
            {
                // Call FolderSync command
                FolderSyncResponse folderSyncResponse = this.CMDAdapter.FolderSync(Common.CreateFolderSyncRequest("0"));

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4912");

                // Verify MS-ASCMD requirement: MS-ASCMD_R4912
                Site.CaptureRequirementIfAreEqual<byte>(
                    144,
                    folderSyncResponse.ResponseData.Status,
                    4912,
                    @"[In Common Status Codes] [The meaning of the status value 144 is] The device's policy key is invalid.");
            }

            #endregion
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Get PolicyKey from Provision Response
        /// </summary>
        /// <param name="response">Provision Response</param>
        /// <returns>Policy Key, if the response doesn't contain the PolicyKey, returns null</returns>
        private static string GetPolicyKeyFromResponse(ProvisionResponse response)
        {
            if (null != response.ResponseData.Policies)
            {
                Response.ProvisionPoliciesPolicy policyInResponse = response.ResponseData.Policies.Policy;
                if (policyInResponse != null)
                {
                    return policyInResponse.PolicyKey;
                }
            }

            return null;
        }
        #endregion
    }
}