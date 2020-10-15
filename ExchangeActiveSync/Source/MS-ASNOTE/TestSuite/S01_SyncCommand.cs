namespace Microsoft.Protocols.TestSuites.MS_ASNOTE
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Globalization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to synchronize notes on the server.
    /// </summary>
    [TestClass]
    public class S01_SyncCommand : TestSuiteBase
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

        #region MSASNOTE_S01_TC01_Sync_AddNote
        /// <summary>
        /// This test case is designed to test adding a note using the Sync command.
        /// </summary>
        [TestCategory("MSASNOTE"), TestMethod()]
        public void MSASNOTE_S01_TC01_Sync_AddNote()
        {
            #region Call method Sync to add a note to the server
            Dictionary<Request.ItemsChoiceType8, object> addElements = this.CreateNoteElements();
            this.SyncAdd(addElements, 1);

            #endregion

            #region Call method Sync to synchronize the note item with the server.

            SyncStore result = this.SyncChanges(1);

            Note note = result.AddElements[0].Note;

            Site.Assert.IsNotNull(
                note.Categories,
                @"The Categories element in note class in response should not be null.");

            Site.Assert.IsNotNull(
                note.Categories.Category,
                @"The Category element in note class in response should not be null.");

            Site.Assert.AreEqual<int>(
                1,
                note.Categories.Category.Length,
                "The length of category should be 1 in response");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R211");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R211
            // If the value of the single category element is the same in request and response, then MS-ASNOTE_R211 can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                ((Request.Categories4)addElements[Request.ItemsChoiceType8.Categories2]).Category[0],
                note.Categories.Category[0],
                211,
                @"[In Category] [The Category element] specifies a user-selected label that has been applied to the note.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R123");

            Site.Assert.IsNotNull(
                note.Body,
                @"The Body element in note class in response should not be null.");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R123
            // If Body element is present in response, and the Data element is not null, then MS-ASNOTE_R123 can be captured.
            Site.CaptureRequirementIfIsNotNull(
                note.Body.Data,
                123,
                @"[In Body] When the airsyncbase:Body element is used in a Sync command  response ([MS-ASCMD] section 2.2.2.19), the airsyncbase:Data element ([MS-ASAIRS] section 2.2.2.10.1) is a required child element of the airsyncbase:Body element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R58");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R58
            // If the value of the subject element is the same in request and response, then MS-ASNOTE_R58 can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                addElements[Request.ItemsChoiceType8.Subject1].ToString(),
                note.Subject,
                58,
                @"[In Subject] The Subject element specifies the subject of the note.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R51");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R51
            // If the value of the LastModifiedDate element is specified, then MS-ASNOTE_R51 can be captured.
            Site.CaptureRequirementIfIsTrue(
                note.IsLastModifiedDateSpecified,
                51,
                @"[In LastModifiedDate] The LastModifiedDate element specifies when the note was last changed.");

            #endregion
        }

        #endregion

        #region MSASNOTE_S01_TC02_Sync_ChangeNote_WithoutBodyInRequest
        /// <summary>
        /// This test case is designed to test changing a note's Subject and MessageClass elements without including the note's body in the Sync command.
        /// </summary>
        [TestCategory("MSASNOTE"), TestMethod()]
        public void MSASNOTE_S01_TC02_Sync_ChangeNote_WithoutBodyInRequest()
        {
            #region Call method Sync to add a note to the server
            Dictionary<Request.ItemsChoiceType8, object> addElements = this.CreateNoteElements();
            addElements[Request.ItemsChoiceType8.Categories2] = new Request.Categories4();
            SyncStore addResult = this.SyncAdd(addElements, 1);
            Response.SyncCollectionsCollectionResponsesAdd item = addResult.AddResponses[0];
            #endregion

            #region Call method Sync to change the note's Subject and MessageClass elements.
            // changeElements:Change the note's subject by replacing its subject with a new subject.
            Dictionary<Request.ItemsChoiceType7, object> changeElements = new Dictionary<Request.ItemsChoiceType7, object>();
            string changedSubject = Common.GenerateResourceName(Site, "subject");
            changeElements.Add(Request.ItemsChoiceType7.Subject2, changedSubject);

            // changeElements:Change the note's MessageClass by replacing its MessageClass with a new MessageClass.
            changeElements.Add(Request.ItemsChoiceType7.MessageClass, "IPM.StickyNote.MSASNOTE");
            changeElements = TestSuiteHelper.CombineChangeAndAddNoteElements(addElements, changeElements);

            // changeElements:Remove the note's Body in change command
            changeElements.Remove(Request.ItemsChoiceType7.Body);
            SyncStore changeResult = this.SyncChange(addResult.SyncKey, item.ServerId, changeElements);

            Site.Assert.AreEqual<byte>(
                1,
                changeResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate sync command succeed.");

            // The subject of the note is updated. 
            this.ExistingNoteSubjects.Remove(addElements[Request.ItemsChoiceType8.Subject1].ToString());
            this.ExistingNoteSubjects.Add(changeElements[Request.ItemsChoiceType7.Subject2].ToString());
            #endregion

            #region Call method Sync to synchronize the note item with the server.
            // Synchronize the changes with server
            SyncStore result = this.SyncChanges(addResult.SyncKey, 1);

            bool isNoteFound = TestSuiteHelper.CheckSyncChangeCommands(result, changeElements[Request.ItemsChoiceType7.Subject2].ToString(), this.Site);

            Site.Assert.IsTrue(isNoteFound, "The note with subject:{0} should be returned in Sync command response.", changeElements[Request.ItemsChoiceType7.Subject2].ToString());

            Note note = result.ChangeElements[0].Note;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R113");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R113
            Site.CaptureRequirementIfIsNotNull(
                note.Body,
                113,
                @"[In Sync Command Response] The absence of an airsyncbase:Body element (section 2.2.2.1) within an airsync:Change element is not to be interpreted as an implicit delete.");

            Site.Assert.AreEqual<string>(
                changeElements[Request.ItemsChoiceType7.Subject2].ToString(),
                note.Subject,
                "The subject element in Change Command response should be the same with the changed value of subject in Change Command request.");

            Site.Assert.AreEqual<string>(
                changeElements[Request.ItemsChoiceType7.MessageClass].ToString(),
                note.MessageClass,
                "The MessageClass element in Change Command response should be the same with the changed value of MessageClass in Change Command request.");

            #endregion
        }
        #endregion

        #region MSASNOTE_S01_TC03_Sync_LastModifiedDateIgnored
        /// <summary>
        /// This test case is designed to test the server ignores the element LastModifiedDate if includes it in the request.
        /// </summary>
        [TestCategory("MSASNOTE"), TestMethod()]
        public void MSASNOTE_S01_TC03_Sync_LastModifiedDateIgnored()
        {
            #region Call method Sync to add a note to the server
            Dictionary<Request.ItemsChoiceType8, object> addElements = this.CreateNoteElements();
            string lastModifiedDate = DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture);
            addElements.Add(Request.ItemsChoiceType8.LastModifiedDate, lastModifiedDate);
            System.Threading.Thread.Sleep(1000);
            SyncStore addResult = this.SyncAdd(addElements, 1);
            Response.SyncCollectionsCollectionResponsesAdd item = addResult.AddResponses[0];
            #endregion

            #region Call method Sync to synchronize the note item with the server.
            SyncStore result = this.SyncChanges(1);

            Note note =null;

            for (int i = 0; i < result.AddElements.Count; i++)
            {
                if (addResult.AddElements != null && addResult.AddElements.Count > 0)
                {
                    if (addResult.CollectionStatus==1&& result.AddElements[0].Note.Subject.ToString()==addElements[Request.ItemsChoiceType8.Subject1].ToString())
                    {
                        note = result.AddElements[i].Note;
                        break;
                    }
                }
                else if (addResult.AddResponses != null && addResult.AddResponses.Count > 0)
                {
                    if (addResult.CollectionStatus == 1 && result.AddElements[0].Note.Subject.ToString() == addElements[Request.ItemsChoiceType8.Subject1].ToString())
                    {
                        note = result.AddElements[i].Note;
                        break;
                    }  
                }             
            }
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R84");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R84
            Site.CaptureRequirementIfAreNotEqual<string>(
                lastModifiedDate,
                note.LastModifiedDate.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture),
                84,
                @"[In LastModifiedDate Element] If it is included in a Sync command request, the server will ignore it.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R209");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R209
            // this requirement can be captured directly after MS-ASNOTE_R84. 
            Site.CaptureRequirement(
                209,
                @"[In LastModifiedDate Element] If a Sync command request includes the LastModifiedDate element, the server ignores the element and returns the actual time that the note was last modified.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R126");

            bool isVerifiedR126 = note.Body != null && note.Subject != null && note.MessageClass != null && note.IsLastModifiedDateSpecified && note.Categories != null && note.Categories.Category != null;

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R126
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR126,
                126,
                @"[In Sync Command Response] Any of the elements for the Notes class[airsyncbase:Body, Subject, MessageClass, LastModifiedDate, Categories or Category], as specified in section 2.2.2, can be included in a Sync command response as child elements of the airsync:ApplicationData element ([MS-ASCMD] section 2.2.3.11) within [either] an airsync:Add element ([MS-ASCMD] section 2.2.3.7.2) [or an airsync:Change element ([MS-ASCMD] section 2.2.3.24)].");

            #endregion

            #region Call method Sync to only change the note's LastModifiedDate element, the server will ignore the change, and the note item should be unchanged.
            // changeElements: Change the note's LastModifiedDate by replacing its LastModifiedDate with a new LastModifiedDate.
            Dictionary<Request.ItemsChoiceType7, object> changeElements = new Dictionary<Request.ItemsChoiceType7, object>();
            lastModifiedDate = DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture);
            changeElements.Add(Request.ItemsChoiceType7.LastModifiedDate, lastModifiedDate);
            this.SyncChange(result.SyncKey, item.ServerId, changeElements);

            #endregion

            #region Call method Sync to synchronize the changes with the server.

            SyncStore result2 = this.SyncChanges(result.SyncKey, 1);

            bool isNoteFound;
            if (result2.ChangeElements != null)
            {
                isNoteFound = TestSuiteHelper.CheckSyncChangeCommands(result, addElements[Request.ItemsChoiceType8.Subject1].ToString(), this.Site);

                Site.Assert.IsFalse(isNoteFound, "The note with subject:{0} should not be returned in Sync command response.", addElements[Request.ItemsChoiceType8.Subject1].ToString());
            }
            else
            {
                Site.Log.Add(LogEntryKind.Debug, @"The Change elements are null.");
            }
            #endregion

            #region Call method Sync to change the note's LastModifiedDate and subject.
            // changeElements: Change the note's LastModifiedDate by replacing its LastModifiedDate with a new LastModifiedDate.
            // changeElements: Change the note's subject by replacing its subject with a new subject.
            changeElements = new Dictionary<Request.ItemsChoiceType7, object>();
            lastModifiedDate = DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture);
            changeElements.Add(Request.ItemsChoiceType7.LastModifiedDate, lastModifiedDate);
            string changedSubject = Common.GenerateResourceName(Site, "subject");
            changeElements.Add(Request.ItemsChoiceType7.Subject2, changedSubject);
            changeElements = TestSuiteHelper.CombineChangeAndAddNoteElements(addElements, changeElements);
            this.SyncChange(result.SyncKey, item.ServerId, changeElements);

            #endregion

            #region Call method Sync to synchronize the note item with the server.

            result = this.SyncChanges(result.SyncKey, 1);

            isNoteFound = TestSuiteHelper.CheckSyncChangeCommands(result, changeElements[Request.ItemsChoiceType7.Subject2].ToString(), this.Site);

            Site.Assert.IsTrue(isNoteFound, "The note with subject:{0} should be returned in Sync command response.", changeElements[Request.ItemsChoiceType7.Subject2].ToString());

            // The subject of the note is updated. 
            this.ExistingNoteSubjects.Remove(addElements[Request.ItemsChoiceType8.Subject1].ToString());
            this.ExistingNoteSubjects.Add(changeElements[Request.ItemsChoiceType7.Subject2].ToString());

            note = result.ChangeElements[0].Note;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R210");

            bool isVerifiedR210 = note.Body != null && note.Subject != null && note.MessageClass != null && note.IsLastModifiedDateSpecified && note.Categories != null && note.Categories.Category != null;

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R210
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR210,
                210,
                @"[In Sync Command Response] Any of the elements for the Notes class[airsyncbase:Body, Subject, MessageClass, LastModifiedDate, Categories or Category], as specified in section 2.2.2, can be included in a Sync command response as child elements of the airsync:ApplicationData element ([MS-ASCMD] section 2.2.3.11) within [either an airsync:Add element ([MS-ASCMD] section 2.2.3.7.2) or] an airsync:Change element ([MS-ASCMD] section 2.2.3.24).");

            Site.Assert.AreEqual<string>(
                changeElements[Request.ItemsChoiceType7.Subject2].ToString(),
                note.Subject,
                "The subject element in Change Command response should be the same with the changed value of subject in Change Command request.");

            #endregion
        }
        #endregion

        #region MSASNOTE_S01_TC04_Sync_SupportedError
        /// <summary>
        /// This test case is designed to test when the client includes an airsync:Supported element in a Sync command request, the server returns a status error 4.
        /// </summary>
        [TestCategory("MSASNOTE"), TestMethod()]
        public void MSASNOTE_S01_TC04_Sync_SupportedError()
        {
            #region Call an initial method Sync including the Supported option.
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                CollectionId = this.UserInformation.NotesCollectionId,
                SyncKey = "0",
                Supported = new Request.Supported()
            };
            SyncRequest syncRequest = Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
            SyncStore syncResult = this.NOTEAdapter.Sync(syncRequest, false);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R114");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R114
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                syncResult.Status,
                114,
                @"[In Sync Command Response] If the airsync:Supported element ([MS-ASCMD] section 2.2.3.164) is included in a Sync command request for Notes class data, the server returns a Status element with a value of 4, as specified in [MS-ASCMD] section 2.2.3.162.16.");

            #endregion
        }
        #endregion

        #region MSASNOTE_S01_TC05_Sync_InvalidMessageClass
        /// <summary>
        /// This test case is designed to test when the MessageClass content does not use the standard format in a Sync request, the server responds with a status error 6.
        /// </summary>
        [TestCategory("MSASNOTE"), TestMethod()]
        public void MSASNOTE_S01_TC05_Sync_InvalidMessageClass()
        {
            #region Call method Sync to add a note to the server
            Dictionary<Request.ItemsChoiceType8, object> addElements = this.CreateNoteElements();
            addElements[Request.ItemsChoiceType8.MessageClass] = "IPM.invalidClass";
            SyncRequest syncRequest = TestSuiteHelper.CreateInitialSyncRequest(this.UserInformation.NotesCollectionId);
            SyncStore syncResult = this.NOTEAdapter.Sync(syncRequest, false);

            Site.Assert.AreEqual<byte>(
                1,
                syncResult.CollectionStatus,
                "The server should return a status code 1 in the Sync command response indicate sync command success.");

            List<object> addData = new List<object>();
            Request.SyncCollectionAdd add = new Request.SyncCollectionAdd
            {
                ClientId = System.Guid.NewGuid().ToString(),
                ApplicationData = new Request.SyncCollectionAddApplicationData
                {
                    ItemsElementName = new Request.ItemsChoiceType8[addElements.Count],
                    Items = new object[addElements.Count]
                }
            };

            addElements.Keys.CopyTo(add.ApplicationData.ItemsElementName, 0);
            addElements.Values.CopyTo(add.ApplicationData.Items, 0);
            addData.Add(add);

            syncRequest = TestSuiteHelper.CreateSyncRequest(syncResult.SyncKey, this.UserInformation.NotesCollectionId, addData);
            SyncStore addResult = this.NOTEAdapter.Sync(syncRequest, false);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R119");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R119
            Site.CaptureRequirementIfAreEqual<int>(
                6,
                int.Parse(addResult.AddResponses[0].Status),
                119,
                @"[In MessageClass Element] If a client submits a Sync command request ([MS-ASCMD] section 2.2.2.19) that contains a MessageClass element value that does not conform to the requirements specified in section 2.2.2.5, the server MUST respond with a Status element with a value of 6, as specified in [MS-ASCMD] section 2.2.3.162.16.");

            #endregion
        }
        #endregion

        #region MSASNOTE_S01_TC06_Sync_AddNote_WithBodyTypes
        /// <summary>
        /// This test case is designed to test that the type element of the body in note item has 3 different values:1, 2, 3.
        /// </summary>
        [TestCategory("MSASNOTE"), TestMethod()]
        public void MSASNOTE_S01_TC06_Sync_AddNote_WithBodyTypes()
        {
            #region Call method Sync to add a note to the server
            Dictionary<Request.ItemsChoiceType8, object> addElements = this.CreateNoteElements();
            this.SyncAdd(addElements, 1);

            #endregion

            #region Call method Sync to synchronize the note item with the server and expect to get the body of Type 1.
            SyncStore result = this.SyncChanges(1);

            Note note = result.AddElements[0].Note;

            Site.Assert.AreEqual<string>(
                ((Request.Body)addElements[Request.ItemsChoiceType8.Body]).Data,
                note.Body.Data,
                @"The content of body in response should be equal to that in request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R38");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R38
            // If the content of the body is the same in request and response and the type is 1, then MS-ASNOTE_R38 can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                note.Body.Type,
                38,
                @"[In Body] The value 1 means Plain text.");

            #endregion

            #region Call method Sync to synchronize the note item with the server and expect to get the body of Type 2.
            result = this.SyncChanges(2);

            note = result.AddElements[0].Note;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R39");

            bool isHTML = TestSuiteHelper.IsHTML(note.Body.Data);
            Site.Assert.IsTrue(
                isHTML,
                @"The content of body element in response should be in HTML format. Actual: {0}",
                isHTML);

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R39
            // If the content of the body is in HTML format and the type is 2, then MS-ASNOTE_R39 can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                note.Body.Type,
                39,
                @"[In Body] The value 2 means HTML.");

            #endregion

            #region Call method Sync to synchronize the note item with the server and expect to get the body of Type 3.
            result = this.SyncChanges(3);

            note = result.AddElements[0].Note;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R40");

            try
            {
                byte[] contentBytes = Convert.FromBase64String(note.Body.Data);
                System.Text.Encoding.UTF8.GetString(contentBytes);
            }
            catch (FormatException formatException)
            {
                throw new FormatException("The content of body should be Base64 encoded", formatException);
            }

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R40
            // If the content of the body is in Base64 format and the type is 3, then MS-ASNOTE_R40 can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                3,
                note.Body.Type,
                40,
                @"[In Body] The value 3 means Rich Text Format (RTF).");

            #endregion
        }
        #endregion

        #region MSASNOTE_S01_TC07_Sync_ChangeNote_Categories
        /// <summary>
        /// This test case is designed to test changing a note's Categories element and its child elements.
        /// </summary>
        [TestCategory("MSASNOTE"), TestMethod()]
        public void MSASNOTE_S01_TC07_Sync_ChangeNote_Categories()
        {
            #region Call method Sync to add a note with two child elements in a Categories element to the server
            Dictionary<Request.ItemsChoiceType8, object> addElements = this.CreateNoteElements();
            Request.Categories4 categories = new Request.Categories4 { Category = new string[2] };
            Collection<string> category = new Collection<string> { "blue category", "red category" };
            category.CopyTo(categories.Category, 0);
            addElements[Request.ItemsChoiceType8.Categories2] = categories;
            this.SyncAdd(addElements, 1);
            #endregion

            #region Call method Sync to synchronize the note item with the server and expect to get two child elements in response.
            // Synchronize the changes with server
            SyncStore result = this.SyncChanges(1);

            Note noteAdded = result.AddElements[0].Note;

            Site.Assert.IsNotNull(noteAdded.Categories, "The Categories element in response should not be null.");
            Site.Assert.IsNotNull(noteAdded.Categories.Category, "The category array in response should not be null.");
            Site.Assert.AreEqual(2, noteAdded.Categories.Category.Length, "The length of category array in response should be equal to 2.");
            #endregion

            #region Call method Sync to change the note with MessageClass elements and one child element of Categories element is missing.
            Dictionary<Request.ItemsChoiceType7, object> changeElements = new Dictionary<Request.ItemsChoiceType7, object>
            {
                {
                    Request.ItemsChoiceType7.MessageClass, "IPM.StickyNote.MSASNOTE1"
                }
            };

            categories.Category = new string[1];
            category.Remove("red category");
            category.CopyTo(categories.Category, 0);
            changeElements.Add(Request.ItemsChoiceType7.Categories3, categories);

            SyncStore changeResult = this.SyncChange(result.SyncKey, result.AddElements[0].ServerId, changeElements);

            Site.Assert.AreEqual<byte>(
                1,
                changeResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate sync command succeed.");

            #endregion

            #region Call method Sync to synchronize the note item with the server, and check if one child element is missing in response.
            // Synchronize the changes with server
            result = this.SyncChanges(result.SyncKey, 1);

            bool isNoteFound = TestSuiteHelper.CheckSyncChangeCommands(result, addElements[Request.ItemsChoiceType8.Subject1].ToString(), this.Site);

            Site.Assert.IsTrue(isNoteFound, "The note with subject:{0} should be returned in Sync command response.", addElements[Request.ItemsChoiceType8.Subject1].ToString());

            Note note = result.ChangeElements[0].Note;
            Site.Assert.IsNotNull(note.Categories, "The Categories element in response should not be null.");
            Site.Assert.IsNotNull(note.Categories.Category, "The category array in response should not be null.");
            Site.Assert.IsNotNull(note.Subject, "The Subject element in response should not be null.");
            Site.Assert.AreEqual(1, note.Categories.Category.Length, "The length of category array in response should be equal to 1.");

            bool hasRedCategory = false;

            if (note.Categories.Category[0].Equals("red category", StringComparison.Ordinal))
            {
                hasRedCategory = true;
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R10002");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R10002
            Site.CaptureRequirementIfIsFalse(
                hasRedCategory,
                10002,
                @"[In Sync Command Response] If a child of the Categories element (section 2.2.2.3) that was previously set is missing, the server will delete that property from the note.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R10003");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R10003
            Site.CaptureRequirementIfAreEqual<string>(
                noteAdded.Subject,
                note.Subject,
                10003,
                @"[In Sync Command Response] The absence of a Subject element (section 2.2.2.6) within an airsync:Change element is not to be interpreted as an implicit delete.");
            #endregion

            #region Call method Sync to change the note with MessageClass elements and without Categories element.
            changeElements = new Dictionary<Request.ItemsChoiceType7, object>
            {
                {
                    Request.ItemsChoiceType7.MessageClass, "IPM.StickyNote.MSASNOTE2"
                }
            };

            changeResult = this.SyncChange(result.SyncKey, result.ChangeElements[0].ServerId, changeElements);

            Site.Assert.AreEqual<byte>(
                1,
                changeResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate sync command succeed.");

            #endregion

            #region Call method Sync to synchronize the note item with the server, and check if the Categories element is missing in response.
            // Synchronize the changes with server
            result = this.SyncChanges(result.SyncKey, 1);

            isNoteFound = TestSuiteHelper.CheckSyncChangeCommands(result, addElements[Request.ItemsChoiceType8.Subject1].ToString(), this.Site);

            Site.Assert.IsTrue(isNoteFound, "The note with subject:{0} should be returned in Sync command response.", addElements[Request.ItemsChoiceType8.Subject1].ToString());

            note = result.ChangeElements[0].Note;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R112");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R112
            Site.CaptureRequirementIfIsNull(
                note.Categories,
                112,
                @"[In Sync Command Response] If the Categories element (section 2.2.2.2) that was previously set is missing[in an airsync:Change element in a Sync command request], the server will delete that property from the note.");

            #endregion
        }
        #endregion
    }
}