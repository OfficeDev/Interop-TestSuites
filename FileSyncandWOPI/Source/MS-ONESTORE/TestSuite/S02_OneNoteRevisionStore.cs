namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the requirements related with .one file.
    /// </summary>
    [TestClass]
    public class S02_OneNoteRevisionStore : TestSuiteBase
    {
        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestSuiteBase.Cleanup();
        }

        #region Test cases
        /// <summary>
        /// The test case is validate that the requirements related with .one file.
        /// </summary>
        [TestCategory("MSONESTORE"), TestMethod]
        public void MSONESTORE_S02_TC01_LoadOneNoteFileWithFileData()
        {
            string fileName = Common.GetConfigurationPropertyValue("OneFileLocal", Site);

            OneNoteRevisionStoreFile file = this.Adapter.LoadOneNoteFile(fileName);
        }

        /// <summary>
        /// The test case is validate that the requirements related with .onetoc2 file.
        /// </summary>
        [TestCategory("MSONESTORE"), TestMethod]
        public void MSONESTORE_S02_TC02_LoadOnetocFile()
        {
            string fileName = Common.GetConfigurationPropertyValue("OnetocFileLocal", Site);

            OneNoteRevisionStoreFile file = this.Adapter.LoadOneNoteFile(fileName);
        }
        #endregion
    }
}
