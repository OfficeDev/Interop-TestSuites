namespace Microsoft.Protocols.TestSuites.MS_ASPROV
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the remote wipe directive.
    /// </summary>
    [TestClass]
    public class S02_RemoteWipe : TestSuiteBase
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
        /// This test case is intended to validate a successful remote wipe directive of Provision.
        /// </summary>
        [TestCategory("MSASPROV"), TestMethod()]
        public void MSASPROV_S02_TC01_RemoteWipe()
        {
            #region Apply a unique DeviceType.
            // Switch the user credential to User1 to get user information.
            this.SwitchUser(this.User1Information, true);

            // Apply the unique DeviceType.
            this.DeviceType = string.Format("{0}{1}", "ASPROV", DateTime.Now.ToString("mmssfff"));
            this.PROVAdapter.ApplyDeviceType(this.DeviceType);
            this.CurrentUserInformation.UserName = this.User1Information.UserName;
            this.CurrentUserInformation.UserDomain = this.User1Information.UserDomain;

            #endregion

            #region Acknowledge the policy setting and set the device status on server to be wipe pending
            this.AcknowledgeSecurityPolicySettings();

            // Set the device status on server to be wipe pending.
            string userEmail = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);

            bool dataWiped = PROVSUTControlAdapter.WipeData(this.SutComputerName, userEmail, this.User1Information.UserPassword, this.DeviceType);
            Site.Assert.IsTrue(dataWiped, "The data on the device with DeviceType {0} should be wiped successfully.", this.DeviceType);
            #endregion

            #region Perform an initial remote wipe
            // Send an empty Provision request to indicate a remote wipe operation on client.
            ProvisionRequest emptyRequest = new ProvisionRequest();
            ProvisionResponse provisionResponse = this.PROVAdapter.Provision(emptyRequest);

            Site.Assert.IsNotNull(provisionResponse, "If the Provision command executes successfully, the response from server should not be null.");
            Site.Assert.AreEqual<byte>(1, provisionResponse.ResponseData.Status, "The server should return status code 1 to indicate a success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R653");

            // Verify MS-ASPROV requirement: MS-ASPROV_R653
            // The RemoteWipe element is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                provisionResponse.ResponseData.RemoteWipe,
                653,
                @"[In Responding to an Initial Request] The RemoteWipe MUST only be included if a remote wipe has been requested for the client.");
            #endregion

            #region Perform a failure remote wipe acknowledgment
            // Set the remote wipe status to 2 to indicate a remote wipe failure on client.
            ProvisionRequest wipeRequest = new ProvisionRequest
            {
                RequestData =
                {
                    RemoteWipe = new Microsoft.Protocols.TestSuites.Common.Request.ProvisionRemoteWipe
                    {
                        Status = 2
                    }
                }
            };

            provisionResponse = this.PROVAdapter.Provision(wipeRequest);

            if (Common.IsRequirementEnabled(1042, this.Site))
            {
                this.Site.CaptureRequirementIfAreEqual<byte>(
                    1,
                    provisionResponse.ResponseData.Status,
                    1042,
                    @"[In Appendix A: Product Behavior]  If the client reports failure, the implementation does return a value of 2 in the Status element [and a remote wipe directive]. (<4> Section 3.2.5.1.2.2:  In Exchange 2007 and Exchange 2010, if the client reports failure, the server returns a value of 1 in the Status element.)");
            }

            if (Common.IsRequirementEnabled(1048, this.Site))
            {
                this.Site.CaptureRequirementIfAreEqual<byte>(
                2,
                provisionResponse.ResponseData.Status,
                1048,
                @"[In Appendix A: Product Behavior] If the client reports failure, the implementation does return a value of 2 in the Status element [and a remote wipe directive]. (Exchange 2013 and above follow this behavior.)");
            }

            // Send an empty Provision request to indicate a remote wipe operation on client.
            provisionResponse = this.PROVAdapter.Provision(emptyRequest);
            Site.Assert.AreEqual<byte>(1, provisionResponse.ResponseData.Status, "The server should return status code 1 to indicate a success.");

            if (Common.IsRequirementEnabled(702, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R702");

                // Verify MS-ASPROV requirement: MS-ASPROV_R702
                // The RemoteWipe element is not null, so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    provisionResponse.ResponseData.RemoteWipe,
                    702,
                    @"[In Appendix B: Product Behavior] If the client reports failure, the implementation does return [a value of 2 in the Status element and] a remote wipe directive. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion

            #region Perform a successful remote wipe acknowledgment
            // Set the remote wipe status to 1 to indicate a successful wipe on client.
            wipeRequest.RequestData.RemoteWipe.Status = 1;
            ProvisionResponse wipeResponse = this.PROVAdapter.Provision(wipeRequest);

            if (Common.IsRequirementEnabled(1041, this.Site))
            {
                this.Site.CaptureRequirementIfAreEqual<byte>(
                1,
                wipeResponse.ResponseData.Status,
                1041,
                @"[In Appendix A: Product Behavior] If the client reports success, the implementation does return a value of 1 in the Status element (section 2.2.2.53.2). (<3> Section 3.2.5.1.2.2:  In Exchange 2007 and Exchange 2010, if the client reports success, the server returns a value of 1 in the Status element and a remote wipe directive.)");
            }

            if (Common.IsRequirementEnabled(1047, this.Site))
            {
                this.Site.CaptureRequirementIfAreEqual<byte>(
                    1,
                    wipeResponse.ResponseData.Status,
                    1047,
                    @"[In Appendix A: Product Behavior] If the client reports success, the implementation does return a value of 1 in the Status element (section 2.2.2.53.2). (Exchange 2013 and above follow this behavior.)");
            }

            // Record the provision confirmation mail for user1 to the item collection of User1.
            string confirmationMailSubject = "Remote Device Wipe Confirmation";
            CreatedItems inboxItemForUser1 = Common.RecordCreatedItem(this.User1Information.InboxCollectionId, confirmationMailSubject);
            this.User1Information.UserCreatedItems.Add(inboxItemForUser1);
            CreatedItems sentItemForUser1 = Common.RecordCreatedItem(this.User1Information.SentItemsCollectionId, confirmationMailSubject);
            this.User1Information.UserCreatedItems.Add(sentItemForUser1);
            #endregion

            #region Remove the device from server and perform another initial remote wipe
            // Remove the device from the mobile list after wipe operation is successful.
            bool deviceRemoved = PROVSUTControlAdapter.RemoveDevice(this.SutComputerName, userEmail, this.User1Information.UserPassword, this.DeviceType);
            Site.Assert.IsTrue(deviceRemoved, "The device with DeviceType {0} should be removed successfully.", this.DeviceType);

            // Send an empty Provision request when the client is not requested for a remote wipe.
            provisionResponse = this.PROVAdapter.Provision(emptyRequest);
            Site.Assert.AreEqual<byte>(1, provisionResponse.ResponseData.Status, "The server should return status code 1 to indicate a success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R654");

            // Verify MS-ASPROV requirement: MS-ASPROV_R654
            // The RemoteWipe element is null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNull(
                provisionResponse.ResponseData.RemoteWipe,
                654,
                @"[In Responding to an Initial Request] Otherwise [if a remote wipe has not been requested for the client], it [RemoteWipe] MUST be omitted");
            #endregion
        }
    }
}