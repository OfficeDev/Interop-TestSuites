namespace Microsoft.Protocols.TestSuites.MS_ASRM
{
    using System.Globalization;
    using Common.DataStructures;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to test the Search command.
    /// </summary>
    [TestClass]
    public class S03_Search : TestSuiteBase
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
        public static void ClassCleanUp()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region MSASRM_S03_TC01_Search_RightsManagedEmailMessages
        /// <summary>
        /// This test case is designed to call Search command to find a rights-managed e-mail message.
        /// </summary>
        [TestCategory("MSASRM"), TestMethod()]
        public void MSASRM_S03_TC01_Search_RightsManagedEmailMessages()
        {
            this.CheckPreconditions();

            #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed.
            string templateID = this.GetTemplateID("MSASRM_AllRights_AllowedTemplate");
            #endregion

            #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, switches to User2, and calls FolderSync command.
            string subject = this.SendMailAndFolderSync(templateID, false, null);
            #endregion

            #region The client logs on User2's account, calls Search command to search the rights-managed e-mail message from server.
            SearchRequest searchRequest = Common.CreateSearchRequest(subject, this.UserTwoInformation.InboxCollectionId);
            searchRequest.RequestData.Items[0].Options.Items = new object[] { string.Empty, string.Empty, true };
            searchRequest.RequestData.Items[0].Options.ItemsElementName = new Request.ItemsChoiceType6[] 
            {
                Request.ItemsChoiceType6.RebuildResults,
                Request.ItemsChoiceType6.DeepTraversal,
                Request.ItemsChoiceType6.RightsManagementSupport
            };

            SearchStore result = this.ASRMAdapter.Search(searchRequest);

            Site.Assert.AreEqual<int>(1, result.Results.Count, "There should be only 1 item fetched in ItemOperations command response.");
            Search search = result.Results[0];
            Site.Assert.IsNotNull(search, "The returned item should not be null.");
            Site.Assert.IsNotNull(search.Email, "The expected rights-managed e-mail message should not be null.");
            Site.Assert.IsNull(search.Email.Attachments, "The Attachments element in expected rights-managed e-mail message should be null.");
            Site.Assert.IsNotNull(search.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
            #endregion
        }
        #endregion
    }
}