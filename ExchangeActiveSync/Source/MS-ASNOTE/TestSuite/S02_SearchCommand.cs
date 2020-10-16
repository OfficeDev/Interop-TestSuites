namespace Microsoft.Protocols.TestSuites.MS_ASNOTE
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to retrieve notes that match the criteria specified by the client.
    /// </summary>
    [TestClass]
    public class S02_SearchCommand : TestSuiteBase
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

        #region MSASNOTE_S02_TC01_Search_GetZeroOrMoreNotes
        /// <summary>
        /// This test case is designed to test when there is zero or more notes that satisfy the search criteria, the server will respond with expected number of notes.
        /// </summary>
        [TestCategory("MSASNOTE"), TestMethod()]
        public void MSASNOTE_S02_TC01_Search_GetZeroOrMoreNotes()
        {
            #region Call method Sync to add two notes to the server
            Dictionary<Request.ItemsChoiceType8, object> addElements = this.CreateNoteElements();
            this.SyncAdd(addElements, 2);
            #endregion

            #region Call method Search to search notes using the given keyword text
                // Search note from server
                SearchStore result = this.NOTEAdapter.Search(this.UserInformation.NotesCollectionId, Common.GeneratePrefixOfResourceName(this.Site) + "_subject", true, 2);

                this.Site.Assert.AreEqual<int>(
                    2,
                    result.Results.Count,
                    @"Two results should be returned in Search response.");

                this.Site.Assert.IsNotNull(
                    result.Results[0].Note,
                    @"The first note class in Search response should not be null.");

                this.Site.Assert.IsNotNull(
                    result.Results[1].Note,
                    @"The second note class in Search response should not be null.");

                #endregion
            
            #region Call method Search to search notes using an invalid keyword text

            result = this.NOTEAdapter.Search(this.UserInformation.NotesCollectionId, Common.GenerateResourceName(this.Site, "notExisting_subject"), false, 0);

            this.Site.Assert.AreEqual<int>(
                0,
                result.Results.Count,
                @"No results should be returned in Search response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R208");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R208
            // Server can return zero or more Notes class blocks which can be seen from two steps above.
            this.Site.CaptureRequirement(
                208,
                @"[In Abstract Data Model] The server returns a Notes class XML block for every note that matches the criteria specified by the client command request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R128");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R128
            // Server can return zero or more Notes class blocks which can be seen from two steps above.
            this.Site.CaptureRequirement(
                128,
                @"[In Abstract Data Model] The server can return zero or more Notes class XML blocks in its response, depending on how many notes match the criteria specified by the client command request.");

                #endregion
            
        }
        #endregion
    }
}