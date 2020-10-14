namespace Microsoft.Protocols.TestSuites.MS_ASCON
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to apply a conversation-based filter using GetItemEstimate command.
    /// </summary>
    [TestClass]
    public class S02_GetItemEstimate : TestSuiteBase
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

        #region MSASCON_S02_TC01_GetItemEstimate_Filter
        /// <summary>
        /// This test case is designed to validate filtering a conversation by GetItemEstimate command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S02_TC01_GetItemEstimate_Filter()
        {
            #region Create a conversation and sync to get the created conversation item.
            string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
            this.CreateConversation(conversationSubject);
            #endregion

            #region Initial Sync on Inbox folder.
            // Call Initial Sync command to get the latest SyncKey.
            SyncStore syncStore = this.CONAdapter.Sync(Common.CreateInitialSyncRequest(User1Information.InboxCollectionId));
            #endregion

            #region Send GetItemEstimate request and get response.
            GetItemEstimateResponse getItemEstimateResponse = this.CallGetItemEstimateCommand(syncStore.SyncKey, User1Information.InboxCollectionId);

            // Verify GetItemEstimate command response.
            bool isVerifyR211 = getItemEstimateResponse.ResponseData.Response.Length == 1 && getItemEstimateResponse.ResponseData.Response[0].Status == "1";

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R211");
            Site.Log.Add(LogEntryKind.Debug, "The length of the Response element from GetItemEstimate command response is {0}.", getItemEstimateResponse.ResponseData.Response.Length);
            Site.Log.Add(LogEntryKind.Debug, "The value of the Status element from GetItemEstimate command response is {0}.", getItemEstimateResponse.ResponseData.Response[0].Status);

            // Verify MS-ASCON requirement: MS-ASCON_R211
            // If the GetItemEstimate command executed successfully, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR211,
                211,
                 @"[In Processing a GetItemEstimate Command] When a conversation-based filter is applied to the GetItemEstimate command request, as specified in section 3.1.4.9, the server sends a GetItemEstimate command response ([MS-ASCMD] section 2.2.1.9) that specifies an estimate of the items that meet the filter criteria and need to be synchronized.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R332");

            // Verify MS-ASCON requirement: MS-ASCON_R332
            // If the GetItemEstimate command executed successfully, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR211,
                332,
                 @"[In Applying a Conversation-based Filter] A conversation-based filter can also be applied to the GetItemEstimate command request ([MS-ASCMD] section 2.2.1.9) to get an estimate of the items that both meet the filter criteria and need to be synchronized.");
            #endregion
        }
        #endregion

        #region MSASCON_S02_TC02_GetItemEstimate_Status4
        /// <summary>
        /// This test case is designed to validate Status 4 (protocol error) by GetItemEstimate command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S02_TC02_GetItemEstimate_Status4()
        {
            #region Initial Sync on Calendar folder.
            // Call Initial Sync command to get the latest SyncKey.
            SyncStore syncStore = this.CONAdapter.Sync(Common.CreateInitialSyncRequest(User1Information.CalendarCollectionId));
            #endregion

            #region Call GetItemEstimate command on Calendar folder with setting ConversationMode element in the request.
            GetItemEstimateResponse getItemEstimateResponse = this.CallGetItemEstimateCommand(syncStore.SyncKey, User1Information.CalendarCollectionId);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R336");

            // Verify MS-ASCON requirement: MS-ASCON_R336
            // If the response Status is 4, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(getItemEstimateResponse.ResponseData.Status),
                336,
                @"[In Processing a GetItemEstimate Command] [The meaning of status code] 4 [is] Protocol error. The conversation-based filter cannot be applied to a folder that is not of the Email class.");
            #endregion
        }
        #endregion
    }
}