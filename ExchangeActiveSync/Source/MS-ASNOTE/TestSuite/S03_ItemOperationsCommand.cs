namespace Microsoft.Protocols.TestSuites.MS_ASNOTE
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to retrieve notes data from the server.
    /// </summary>
    [TestClass]
    public class S03_ItemOperationsCommand : TestSuiteBase
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

        #region MSASNOTE_S03_TC01_ItemOperations_GetZeroOrMoreNotes
        /// <summary>
        /// This test case is designed to test when there is zero or more notes that satisfy the ItemOperations criteria, the server responds with expected number of notes.
        /// </summary>
        [TestCategory("MSASNOTE"), TestMethod()]
        public void MSASNOTE_S03_TC01_ItemOperations_GetZeroOrMoreNotes()
        {
            #region Call method Sync to add two notes to the server
            Dictionary<Request.ItemsChoiceType8, object> addElements = this.CreateNoteElements();
            this.SyncAdd(addElements, 2);
            #endregion

            #region Call method Sync to synchronize the note item with the server.
            SyncStore result = this.SyncChanges(1);

            #endregion

            #region Call method ItemOperations to fetch all the information about notes using ServerIds
            // serverIds:the server ids of two note items.
            List<string> serverIds = new List<string> { result.AddElements[0].ServerId, result.AddElements[1].ServerId };
            ItemOperationsRequest itemOperationRequest = TestSuiteHelper.CreateItemOperationsFetchRequest(this.UserInformation.NotesCollectionId, serverIds, null, null, null);
            ItemOperationsStore itemOperationsResult = this.NOTEAdapter.ItemOperations(itemOperationRequest);

            Site.Assert.AreEqual<int>(
                2,
                itemOperationsResult.Items.Count,
                @"Two results should be returned in ItemOperations response.");

            Site.Assert.IsNotNull(
                itemOperationsResult.Items[0].Note,
                @"The first note class in ItemOperations response should not be null.");

            Site.Assert.IsNotNull(
                itemOperationsResult.Items[1].Note,
                @"The second note class in ItemOperations response should not be null.");

            #endregion

            #region Call method ItemOperations to fetch all the information about notes using a non-existing ServerIds
            serverIds.Clear();
            serverIds.Add(this.UserInformation.NotesCollectionId + ":notExisting");
            itemOperationRequest = TestSuiteHelper.CreateItemOperationsFetchRequest(this.UserInformation.NotesCollectionId, serverIds, null, null, null);
            itemOperationsResult = this.NOTEAdapter.ItemOperations(itemOperationRequest);

            Site.Assert.IsNull(
                itemOperationsResult.Items[0].Note,
                @"zero Notes class XML blocks is returned in its response");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R208");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R208
            // Server can return zero or more Notes class blocks which can be seen from two steps above.
            Site.CaptureRequirement(
                208,
                @"[In Abstract Data Model] The server returns a Notes class XML block for every note that matches the criteria specified by the client command request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R128");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R128
            // Server can return zero or more Notes class blocks which can be seen from two steps above.
            Site.CaptureRequirement(
                128,
                @"[In Abstract Data Model] The server can return zero or more Notes class XML blocks in its response, depending on how many notes match the criteria specified by the client command request.");

            #endregion
        }
        #endregion

        #region MSASNOTE_S03_TC02_ItemOperations_SchemaViewFetch
        /// <summary>
        /// This test case is designed to test when an itemoperations:Schema element is included in the ItemOperations command request, the elements returned by the server are restricted by the schema.
        /// </summary>
        [TestCategory("MSASNOTE"), TestMethod()]
        public void MSASNOTE_S03_TC02_ItemOperations_SchemaViewFetch()
        {
            #region Call method Sync to add a note to the server
            Dictionary<Request.ItemsChoiceType8, object> addElements = this.CreateNoteElements();
            this.SyncAdd(addElements, 1);
            #endregion

            #region Call method Search to search notes using the given keyword text

            // Search note from server
            SearchStore result = this.NOTEAdapter.Search(this.UserInformation.NotesCollectionId, addElements[Request.ItemsChoiceType8.Subject1].ToString(), true, 1);

            Site.Assert.AreEqual<int>(
                1,
                result.Results.Count,
                @"There should be only one note item returned in sync response.");

            #endregion

            #region Call method ItemOperations to fetch all the information about notes using longIds.
            // longIds:Long id of the created note item.
            List<string> longIds = new List<string> { result.Results[0].LongId };

            Request.BodyPreference bodyReference = new Request.BodyPreference { Type = 1 };
            Request.Schema schema = new Request.Schema
            {
                ItemsElementName = new Request.ItemsChoiceType3[1],
                Items = new object[] { new Request.Body() }
            };
            schema.ItemsElementName[0] = Request.ItemsChoiceType3.Body;

            // serverIds:null
            ItemOperationsRequest itemOperationRequest = TestSuiteHelper.CreateItemOperationsFetchRequest(null, null, longIds, bodyReference, schema);
            ItemOperationsStore itemOperationsResult = this.NOTEAdapter.ItemOperations(itemOperationRequest);
            Site.Assert.IsNotNull(itemOperationsResult, "The ItemOperations result must not be null!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R103");

            Site.Assert.IsNull(itemOperationsResult.Items[0].Note.Subject, "Subject should be null.");
            Site.Assert.IsNull(itemOperationsResult.Items[0].Note.MessageClass, "MessageClass should be null.");
            Site.Assert.IsNull(itemOperationsResult.Items[0].Note.Categories, "Categories should be null.");
            Site.Assert.IsFalse(itemOperationsResult.Items[0].Note.IsLastModifiedDateSpecified, "LastModifiedSpecified should not be present.");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R103
            Site.CaptureRequirementIfIsNotNull(
                itemOperationsResult.Items[0].Note.Body,
                103,
                @"[In ItemOperations Command Response] If an itemoperations:Schema element ([MS-ASCMD] section 2.2.3.145) is included in the ItemOperations command request, then the elements returned in the ItemOperations command response MUST be restricted to the elements that were included as child elements of the ItemOperations:Schema element in the command request.");

            #endregion
        }
        #endregion
    }
}