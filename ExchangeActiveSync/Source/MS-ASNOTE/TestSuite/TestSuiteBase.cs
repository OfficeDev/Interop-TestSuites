namespace Microsoft.Protocols.TestSuites.MS_ASNOTE
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// The base class of scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Variables

        /// <summary>
        /// Gets the list of existing notes' subjects
        /// </summary>
        protected Collection<string> ExistingNoteSubjects { get; private set; }

        /// <summary>
        /// Gets protocol interface of MS-ASNOTE
        /// </summary>
        protected IMS_ASNOTEAdapter NOTEAdapter { get; private set; }

        /// <summary>
        /// Gets or sets the related information of User.
        /// </summary>
        protected UserInformation UserInformation { get; set; }

        #endregion

        #region Test suite initialize and clean up

        /// <summary>
        /// Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            // If implementation doesn't support this specification [MS-ASNOTE], the case will not start.
            if (bool.Parse(Common.GetConfigurationPropertyValue("MS-ASNOTE_Supported", this.Site)))
            {
                if (this.ExistingNoteSubjects != null && this.ExistingNoteSubjects.Count > 0)
                {
                    SyncStore changesResult = this.SyncChanges(1);

                    foreach (string subject in this.ExistingNoteSubjects)
                    {
                        string serverId = null;
                        foreach (Sync add in changesResult.AddElements)
                        {
                            if (add.Note.Subject == subject)
                            {
                                serverId = add.ServerId;
                                break;
                            }
                        }

                        this.Site.Assert.IsNotNull(serverId, "The note with subject {0} should be found.", subject);

                        SyncStore deleteResult = this.SyncDelete(changesResult.SyncKey, serverId);

                        this.Site.Assert.AreEqual<byte>(
                            1,
                            deleteResult.CollectionStatus,
                            "The server should return a status code of 1 in the Sync command response indicate sync command succeed.");
                    }

                    this.ExistingNoteSubjects.Clear();
                }
            }

            base.TestCleanup();
        }

        /// <summary>
        /// Initialize the Test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            if (this.NOTEAdapter == null)
            {
                this.NOTEAdapter = this.Site.GetAdapter<IMS_ASNOTEAdapter>();
            }

            // If implementation doesn't support this specification [MS-ASNOTE], the case will not start.
            if (!bool.Parse(Common.GetConfigurationPropertyValue("MS-ASNOTE_Supported", this.Site)))
            {
                this.Site.Assert.Inconclusive("This test suite is not supported under current SUT, because MS-ASNOTE_Supported value is set to false in MS-ASNOTE_{0}_SHOULDMAY.deployment.ptfconfig file.", Common.GetSutVersion(this.Site));
            }

            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Notes class is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            this.ExistingNoteSubjects = new Collection<string>();

            // Set the information of user.
            this.UserInformation = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("UserName", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("UserPassword", this.Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", this.Site)
            };

            if (Common.GetSutVersion(this.Site) != SutVersion.ExchangeServer2007 || string.Equals(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "12.1"))
            {
                FolderSyncResponse folderSyncResponse = this.NOTEAdapter.FolderSync(Common.CreateFolderSyncRequest("0"));

                // Get the CollectionId from FolderSync command response.
                if (string.IsNullOrEmpty(this.UserInformation.NotesCollectionId))
                {
                    this.UserInformation.NotesCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Notes, this.Site);
                }
            }
        }

        #endregion

        /// <summary>
        /// Create the elements of a note
        /// </summary>
        /// <returns>The dictionary of value and name for note's elements to be created</returns>
        protected Dictionary<Request.ItemsChoiceType8, object> CreateNoteElements()
        {
            Dictionary<Request.ItemsChoiceType8, object> addElements = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(this.Site, "subject");
            addElements.Add(Request.ItemsChoiceType8.Subject1, subject);

            Request.Body noteBody = new Request.Body { Type = 1, Data = "Content of the body." };
            addElements.Add(Request.ItemsChoiceType8.Body, noteBody);

            Request.Categories4 categories = new Request.Categories4 { Category = new string[] { "blue category" } };
            addElements.Add(Request.ItemsChoiceType8.Categories2, categories);

            addElements.Add(Request.ItemsChoiceType8.MessageClass, "IPM.StickyNote");
            return addElements;
        }

        #region Call Sync command to fetch the notes

        /// <summary>
        /// Call Sync command to fetch all notes
        /// </summary>
        /// <param name="bodyType">The type of the body</param>
        /// <returns>Return change result</returns>
        protected SyncStore SyncChanges(byte bodyType)
        {
            SyncRequest syncInitialRequest = TestSuiteHelper.CreateInitialSyncRequest(this.UserInformation.NotesCollectionId);
            SyncStore syncInitialResult = this.NOTEAdapter.Sync(syncInitialRequest, false);

            // Verify sync change result
            this.Site.Assert.AreEqual<byte>(
                1,
                syncInitialResult.CollectionStatus,
                "The server returns a Status 1 in the Sync command response indicate sync command success.",
                syncInitialResult.Status);

            SyncStore syncResult = this.SyncChanges(syncInitialResult.SyncKey, bodyType);

            this.Site.Assert.AreEqual<byte>(
                1,
                syncResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate sync command succeed.");

            this.Site.Assert.IsNotNull(
                syncResult.AddElements,
                "The server should return Add elements in response");

            Collection<Sync> expectedCommands = new Collection<Sync>();
            foreach (Sync sync in syncResult.AddElements)
            {
                this.Site.Assert.IsNotNull(
                    sync,
                    @"The Add element in response should not be null.");

                this.Site.Assert.IsNotNull(
                    sync.Note,
                    @"The note class in response should not be null.");

                if (this.ExistingNoteSubjects.Contains(sync.Note.Subject))
                {
                    expectedCommands.Add(sync);
                }
            }

            this.Site.Assert.AreEqual<int>(
                this.ExistingNoteSubjects.Count,
                expectedCommands.Count,
                @"The number of Add elements returned in response should be equal to the number of expected notes' subjects");

            syncResult.AddElements.Clear();
            foreach (Sync sync in expectedCommands)
            {
                syncResult.AddElements.Add(sync);
            }

            return syncResult;
        }

        /// <summary>
        /// Call Sync command to fetch the change of the notes from previous syncKey
        /// </summary>
        /// <param name="syncKey">The sync key</param>
        /// <param name="bodyType">The type of the body</param>
        /// <returns>Return change result</returns>
        protected SyncStore SyncChanges(string syncKey, byte bodyType)
        {
            Request.BodyPreference bodyPreference = new Request.BodyPreference { Type = bodyType };

            SyncRequest syncRequest = TestSuiteHelper.CreateSyncRequest(syncKey, this.UserInformation.NotesCollectionId, bodyPreference);
            SyncStore syncResult = this.NOTEAdapter.Sync(syncRequest, true);

            return syncResult;
        }

        #endregion

        #region Call Sync command to add a note

        /// <summary>
        /// Call Sync command to add a note
        /// </summary>
        /// <param name="addElements">The elements of a note item to be added</param>
        /// <param name="count">The number of the note</param>
        /// <returns>Return the sync add result</returns>
        protected SyncStore SyncAdd(Dictionary<Request.ItemsChoiceType8, object> addElements, int count)
        {
            SyncRequest syncRequest = TestSuiteHelper.CreateInitialSyncRequest(this.UserInformation.NotesCollectionId);
            SyncStore syncResult = this.NOTEAdapter.Sync(syncRequest, false);

            // Verify sync change result
            this.Site.Assert.AreEqual<byte>(
                1,
                syncResult.CollectionStatus,
                "The server should return a status code 1 in the Sync command response indicate sync command success.");

            List<object> addData = new List<object>();
            string[] subjects = new string[count];

            // Construct every note
            for (int i = 0; i < count; i++)
            {
                Request.SyncCollectionAdd add = new Request.SyncCollectionAdd
                {
                    ClientId = System.Guid.NewGuid().ToString(),
                    ApplicationData = new Request.SyncCollectionAddApplicationData
                    {
                        ItemsElementName = new Request.ItemsChoiceType8[addElements.Count],
                        Items = new object[addElements.Count]
                    }
                };

                // Since only one subject is generated in addElement, if there are multiple notes, generate unique subjects with index for every note.
                if (count > 1)
                {
                    addElements[Request.ItemsChoiceType8.Subject1] = Common.GenerateResourceName(this.Site, "subject", (uint)(i + 1));
                }

                subjects[i] = addElements[Request.ItemsChoiceType8.Subject1].ToString();
                addElements.Keys.CopyTo(add.ApplicationData.ItemsElementName, 0);
                addElements.Values.CopyTo(add.ApplicationData.Items, 0);
                addData.Add(add);
            }

            syncRequest = TestSuiteHelper.CreateSyncRequest(syncResult.SyncKey, this.UserInformation.NotesCollectionId, addData);
            SyncStore addResult = this.NOTEAdapter.Sync(syncRequest, false);

            this.Site.Assert.AreEqual<byte>(
                1,
                addResult.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate sync command succeed.");

            this.Site.Assert.IsNotNull(
                addResult.AddResponses,
                @"The Add elements in Responses element of the Sync response should not be null.");

            this.Site.Assert.AreEqual<int>(
                count,
                addResult.AddResponses.Count,
                @"The actual number of note items should be returned in Sync response as the expected number.");

            for (int i = 0; i < count; i++)
            {
                this.Site.Assert.IsNotNull(
                    addResult.AddResponses[i],
                    @"The Add element in response should not be null.");

                this.Site.Assert.AreEqual<int>(
                    1,
                    int.Parse(addResult.AddResponses[i].Status),
                    "The server should return a Status 1 in the Sync command response indicate sync command succeed.");

                this.ExistingNoteSubjects.Add(subjects[i]);
            }

            return addResult;
        }

        #endregion

        #region Call Sync command to change a note

        /// <summary>
        /// Call Sync command to change a note
        /// </summary>
        /// <param name="syncKey">The sync key</param>
        /// <param name="serverId">The server Id of the note</param>
        /// <param name="changedElements">The changed elements of the note</param>
        /// <returns>Return the sync change result</returns>
        protected SyncStore SyncChange(string syncKey, string serverId, Dictionary<Request.ItemsChoiceType7, object> changedElements)
        {
            Request.SyncCollectionChange change = new Request.SyncCollectionChange
            {
                ServerId = serverId,
                ApplicationData = new Request.SyncCollectionChangeApplicationData
                {
                    ItemsElementName = new Request.ItemsChoiceType7[changedElements.Count],
                    Items = new object[changedElements.Count]
                }
            };

            changedElements.Keys.CopyTo(change.ApplicationData.ItemsElementName, 0);
            changedElements.Values.CopyTo(change.ApplicationData.Items, 0);

            List<object> changeData = new List<object> { change };
            SyncRequest syncRequest = TestSuiteHelper.CreateSyncRequest(syncKey, this.UserInformation.NotesCollectionId, changeData);
            return this.NOTEAdapter.Sync(syncRequest, false);
        }

        #endregion

        #region Call Sync command to delete a note

        /// <summary>
        /// Call Sync command to delete a note
        /// </summary>
        /// <param name="syncKey">The sync key</param>
        /// <param name="serverId">The server id of the note, which is returned by server</param>
        /// <returns>Return the sync delete result</returns>
        private SyncStore SyncDelete(string syncKey, string serverId)
        {
            List<object> deleteData = new List<object> { new Request.SyncCollectionDelete { ServerId = serverId } };
            SyncRequest syncRequest = TestSuiteHelper.CreateSyncRequest(syncKey, this.UserInformation.NotesCollectionId, deleteData);
            SyncStore result = this.NOTEAdapter.Sync(syncRequest, false);
            return result;
        }

        #endregion
    }
}