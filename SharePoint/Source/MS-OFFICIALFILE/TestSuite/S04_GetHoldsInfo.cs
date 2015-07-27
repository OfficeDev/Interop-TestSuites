//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This test class is used to test GetHoldsInfo operation.
    /// </summary>
    [TestClass]
    public class S04_GetHoldsInfo : TestSuiteBase
    {
        #region Test Suite Initialization

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">An instance of an object that derives from the Microsoft.VisualStudio.TestTools.UnitTesting.TestContext class.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear up the class.
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.Cleanup();
        }

        #endregion

        /// <summary>
        /// This test case is used to test GetServerInfo on a repository that is configured for routing content with one added hold.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S04_TC01_GetHoldsInfo()
        {
            if (!Common.Common.IsRequirementEnabled(354, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetHoldsInfo operations.");
            }

            // Initial parameters to use the repository that is not configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            // call GetHoldsInfo to get all the holds.
            HoldInfo[] holdsInfo = this.Adapter.GetHoldsInfo();

            if (Common.Common.IsRequirementEnabled(354, this.Site))
            {
                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R354
                Site.CaptureRequirementIfIsNotNull(
                         holdsInfo,
                         "MS-OFFICIALFILE",
                         354,
                         @"[In Appendix C: Product Behavior] Implementation does provide this method [GetHoldsInfo]. (Microsoft SharePoint Server 2010 and above follow this behavior.)");
            }

            // Verify the every item in ArrayOfHoldInfo. We just add one hold on server for testing.
            if (holdsInfo.Length >= 1)
            {
                foreach (HoldInfo holdInfo in holdsInfo)
                {
                    bool isNotNullOrEmpty = !string.IsNullOrEmpty(holdInfo.Url);

                    Site.Assert.IsTrue(
                        isNotNullOrEmpty,
                        string.Format("The URL of the legal hold should be non-empty, actual it is {0}", isNotNullOrEmpty ? "non-empty" : "empty"));

                    // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R150
                    Site.CaptureRequirementIfIsTrue(
                             isNotNullOrEmpty,
                             "MS-OFFICIALFILE",
                             150,
                             @"[In HoldInfo] Url: URL of the legal hold, which MUST be non-empty.");

                    // Verify the every item in ArrayOfHoldInfo. We just add one hold on server for testing.
                    // Add the log information.
                    Site.Log.Add(Microsoft.Protocols.TestTools.LogEntryKind.Comment, "holdsInfo[0].Id = {0}", holdsInfo[0].Id.ToString());

                    Site.Assert.IsTrue(
                       holdInfo.Id > 0,
                       string.Format("For the requirement MS-OFFICIALFILE_R156, the Id value should be a positive integer, actually it is {0}", holdsInfo[0].Id > 0 ? "positive" : "negative"));

                    // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R156
                    Site.CaptureRequirementIfIsTrue(
                             holdInfo.Id > 0,
                             "MS-OFFICIALFILE",
                             156,
                             @"[In HoldInfo] Id: Identifier of the legal hold, which MUST be a positive integer.");

                    // Add the log information.
                    Site.Log.Add(Microsoft.Protocols.TestTools.LogEntryKind.Comment, "holdInfo.ListId = {0}", holdInfo.ListId.ToString());

                    // If the de-serialization succeeds, then the ListId should match the GUID format, then directly capture it.
                    // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R157
                    Site.CaptureRequirement(
                             "MS-OFFICIALFILE",
                             157,
                             @"[In HoldInfo] ListId: Identifier of the storage location of the legal hold, which MUST be a GUID.");

                    // Add the log information.
                    Site.Log.Add(Microsoft.Protocols.TestTools.LogEntryKind.Comment, "holdInfo.WebId = {0}", holdInfo.WebId.ToString());

                    // If the de-serialization succeeds, then the WebId should match the GUID format, then directly capture it.
                    // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R158
                    Site.CaptureRequirement(
                             "MS-OFFICIALFILE",
                             158,
                             @"[In HoldInfo] WebId: Identifier of the repository that contains the legal hold, which MUST be a GUID.");
                }
            }
            else
            {
                Site.Assume.Inconclusive(string.Format("At least one legal hold should be configured on the site {0}", Common.Common.GetConfigurationPropertyValue("EnableContentOrganizerRecordsCenterSite", this.Site)));
            }
        }
    }
}
