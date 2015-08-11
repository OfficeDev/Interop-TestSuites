namespace Microsoft.Protocols.TestSuites.MS_SHDACCWS
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;

    /// <summary>
    /// This scenario is used to judge whether it is the only client currently editing a document stored on a collaboration server.
    /// </summary>
    [TestClass]
    public class S01_VerifyIsSingleClient : TestSuiteBase
    {
        #region Additional test attributes, Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.ClassCleanup();
        }

        #endregion

        #region Test Case

        #region MSSHDACCWS_S01_TC01_CoAuthoringIsAsked
        /// <summary>
        /// Verify that the client can get IsOnlyClientSoapOut messages for IsOnlyClient operation, and the server returns "false" when there was a co-authoring transition request for the document.
        /// </summary>
        [TestCategory("MSSHDACCWS"), TestMethod()]
        public void MSSHDACCWS_S01_TC01_CoAuthoringIsAsked()
        {
            // Set the Co-authoring status for the specified file which is identified by the property "FileIdOfCoAuthoring".
            bool isSetCoauthoringSuccess = SHDACCWSSUTControlAdapter.SUTSetCoAuthoringStatus();
            Site.Assert.IsTrue(isSetCoauthoringSuccess, "The Co-authoring status should be set on the specified file.");

            // Get an identifier of the document which there was a co-authoring transition request for.
            string fileIdOfCoAuthoring = Common.GetConfigurationPropertyValue("fileIdOfCoAuthoring", this.Site);

            // Call method IsOnlyClient with the identifier of the document which there was a co-authoring transition request for.
            bool allCoAuthoringStatus = SHDACCWSAdapter.IsOnlyClient(Guid.Parse(fileIdOfCoAuthoring));
            
            // If server returns "false", then capture MS-SHDACCWS requirement: MS-SHDACCWS_R52.
            this.Site.CaptureRequirementIfIsFalse(
                allCoAuthoringStatus,
                52,
                @"[In IsOnlyClientResponse] IsOnlyClientResult : The value of this element MUST be false if there was a co-authoring transition request for the document.");
        }
        #endregion

        #region MSSHDACCWS_S01_TC02_NoClientAuthoring
        /// <summary>
        /// Verify that the client can get IsOnlyClientSoapOut messages for IsOnlyClient operation, and the server returns "true" when no client is editing the document.
        /// </summary>
        [TestCategory("MSSHDACCWS"), TestMethod()]
        public void MSSHDACCWS_S01_TC02_NoClientAuthoring()
        {
            // Get an identifier of the document that no client is editing it.
            string fileIdOfNormal = Common.GetConfigurationPropertyValue("fileIdOfNormal", this.Site);

            // Call method IsOnlyClient with the identifier of the document that no client is editing it.
            bool normalStatus = SHDACCWSAdapter.IsOnlyClient(Guid.Parse(fileIdOfNormal));

            // If server returns "true", then capture MS-SHDACCWS requirement: MS-SHDACCWS_R55.
            this.Site.CaptureRequirementIfIsTrue(
                normalStatus,
                55,
                @"[In IsOnlyClientResponse] IsOnlyClientResult : If no client currently editing the file, the value [IsOnlyClientResult] MUST be true.");
        }
        #endregion

        #region MSSHDACCWS_S01_TC03_OnlyOneClientAuthoring
        /// <summary>
        /// Verify that the client can get IsOnlyClientSoapOut messages for IsOnlyClient operation, and the server returns "true" when the document is currently edited by one client.
        /// </summary>
        [TestCategory("MSSHDACCWS"), TestMethod()]
        public void MSSHDACCWS_S01_TC03_OnlyOneClientAuthoring()
        {
            // Set specified status of exclusive lock to the specified file which is identified by the property "FileIdOfLock".
            bool isSetExclusiveLockSuccess = SHDACCWSSUTControlAdapter.SUTSetExclusiveLock();
            Site.Assert.IsTrue(isSetExclusiveLockSuccess, "The 'Exclusive' lock status should be set on the specified file.");

            // Get an identifier of the document that is currently edited by one client.
            string fileIdOfLock = Common.GetConfigurationPropertyValue("FileIdOfLock", this.Site);
            
            // Call method IsOnlyClient with the identifier of the document that is currently edited by one client.
            bool lockStatus = SHDACCWSAdapter.IsOnlyClient(Guid.Parse(fileIdOfLock));

            // If server returns "true", then capture MS-SHDACCWS requirement: MS-SHDACCWS_R56.
            this.Site.CaptureRequirementIfIsTrue(
                 lockStatus,
                 56,
                 @"[In IsOnlyClientResponse] IsOnlyClientResult : If only one client is allowed to edit the file with the ""Exclusive Lock"" mode , the value [IsOnlyClientResult] MUST be true.");
        }
        #endregion

        #region MSSHDACCWS_S01_TC04_FileNotExistOnServer
        /// <summary>
        /// Verify that the client can get IsOnlyClientSoapOut messages for IsOnlyClient operation, and the server returns "true" when the document specified by the id can't be found on the server.
        /// </summary>
        [TestCategory("MSSHDACCWS"), TestMethod()]
        public void MSSHDACCWS_S01_TC04_FileNotExistOnServer()
        {
            // Call method IsOnlyClient with the identifier of the document that specified by the id can't be found on the server.
            bool nonExistStatus = SHDACCWSAdapter.IsOnlyClient(Guid.NewGuid());

            // If server returns "true", then capture MS-SHDACCWS requirement: MS-SHDACCWS_R57.
            this.Site.CaptureRequirementIfIsTrue(
                nonExistStatus,
                57,
                @"[In IsOnlyClientResponse] IsOnlyClientResult : If the file specified by the input id can't be found on server, the value [IsOnlyClientResult] MUST be true.");
        }
        #endregion

        #endregion
    }
}