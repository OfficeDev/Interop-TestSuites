namespace Microsoft.Protocols.TestSuites.MS_ASNOTE
{
    using System.Collections.Generic;
    using System.Text.RegularExpressions;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// A static class contains all helper methods used in test cases.
    /// </summary>
    internal static class TestSuiteHelper
    {
        /// <summary>
        /// This method is used to check the Sync Change commands.
        /// </summary>
        /// <param name="result">The sync result which is returned from server</param>
        /// <param name="subject">The expected note's subject</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <returns>The boolean value which represents whether the note with expected subject is found or not in sync result</returns>
        internal static bool CheckSyncChangeCommands(SyncStore result, string subject, ITestSite site)
        {
            site.Assert.AreEqual<byte>(
                1,
                result.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate sync command succeed.");

            bool isNoteFound = false;
            foreach (Sync sync in result.ChangeElements)
            {
                site.Assert.IsNotNull(
                    sync,
                    @"The Change element in response should not be null.");

                site.Assert.IsNotNull(
                    sync.Note,
                    @"The note class in response should not be null.");

                if (sync.Note.Subject.Equals(subject))
                {
                    isNoteFound = true;
                }
            }

            return isNoteFound;
        }

        /// <summary>
        /// Combines the to-be-changed elements with the added elements to form a dictionary of changed elements for the note
        /// </summary>
        /// <param name="addElements">All the elements of the created note</param>
        /// <param name="changeElements">The to-be-changed elements of the note</param>
        /// <returns>All the elements of the note to be changed</returns>
        internal static Dictionary<Request.ItemsChoiceType7, object> CombineChangeAndAddNoteElements(Dictionary<Request.ItemsChoiceType8, object> addElements, Dictionary<Request.ItemsChoiceType7, object> changeElements)
        {
            foreach (Request.ItemsChoiceType8 addElementName in addElements.Keys)
            {
                Request.ItemsChoiceType7 changeElementName = (Request.ItemsChoiceType7)System.Enum.Parse(typeof(Request.ItemsChoiceType7), addElementName.ToString());
                if(changeElementName== Request.ItemsChoiceType7.Subject1)
                {
                    changeElementName = Request.ItemsChoiceType7.Subject2;
                }
                if (!changeElements.ContainsKey(changeElementName) && addElementName == Request.ItemsChoiceType8.Categories2)
                {
                    changeElements.Add(Request.ItemsChoiceType7.Categories3, addElements[addElementName]);
                }
                else if(!changeElements.ContainsKey(changeElementName) && addElementName == Request.ItemsChoiceType8.Subject1)
                {
                    changeElements.Add(Request.ItemsChoiceType7.Subject2, addElements[addElementName]);
                }
                else if (!changeElements.ContainsKey(changeElementName))
                {
                    changeElements.Add(changeElementName, addElements[addElementName]);
                }
            }

            return changeElements;
        }

        /// <summary>
        /// Builds a initial Sync request by using the specified collection Id.
        /// </summary>
        /// <param name="collectionId">Folder collection Id to be synchronized.</param>
        /// <returns>Returns the SyncRequest instance</returns>
        internal static SyncRequest CreateInitialSyncRequest(string collectionId)
        {
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                CollectionId = collectionId,
                SyncKey = "0"
            };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
        }

        /// <summary>
        /// Builds a ItemOperations request to fetch the whole content of the notes
        /// </summary>
        /// <param name="collectionId">Specifies the notes folder</param>
        /// <param name="serverIds">Specifies a unique identifier that is assigned by the server for the notes</param>
        /// <param name="longIds">Specifies a unique identifier that is assigned by the server to each result returned by a previous Search response.</param>
        /// <param name="bodyPreference">Sets preference information related to the type and size of the body.</param>
        /// <param name="schema">Specifies the schema of the item to be fetched.</param>
        /// <returns>Returns the ItemOperationsRequest instance</returns>
        internal static ItemOperationsRequest CreateItemOperationsFetchRequest(
            string collectionId,
            List<string> serverIds,
            List<string> longIds,
            Request.BodyPreference bodyPreference,
            Request.Schema schema)
        {
            Request.ItemOperationsFetchOptions fetchOptions = new Request.ItemOperationsFetchOptions();
            List<object> fetchOptionItems = new List<object>();
            List<Request.ItemsChoiceType5> fetchOptionItemsName = new List<Request.ItemsChoiceType5>();

            if (null != bodyPreference)
            {
                fetchOptionItemsName.Add(Request.ItemsChoiceType5.BodyPreference);
                fetchOptionItems.Add(bodyPreference);
            }

            if (null != schema)
            {
                fetchOptionItemsName.Add(Request.ItemsChoiceType5.Schema);
                fetchOptionItems.Add(schema);
            }

            fetchOptions.Items = fetchOptionItems.ToArray();
            fetchOptions.ItemsElementName = fetchOptionItemsName.ToArray();

            List<Request.ItemOperationsFetch> fetchElements = new List<Request.ItemOperationsFetch>();

            if (serverIds != null)
            {
                foreach (string serverId in serverIds)
                {
                    Request.ItemOperationsFetch fetchElement = new Request.ItemOperationsFetch()
                    {
                        CollectionId = collectionId,
                        ServerId = serverId,
                        Store = SearchName.Mailbox.ToString(),
                        Options = fetchOptions
                    };
                    fetchElements.Add(fetchElement);
                }
            }

            if (longIds != null)
            {
                foreach (string longId in longIds)
                {
                    Request.ItemOperationsFetch fetchElement = new Request.ItemOperationsFetch()
                    {
                        LongId = longId,
                        Store = SearchName.Mailbox.ToString(),
                        Options = fetchOptions
                    };
                    fetchElements.Add(fetchElement);
                }
            }

            return Common.CreateItemOperationsRequest(fetchElements.ToArray());
        }

        /// <summary>
        /// Builds a generic Sync request without command references by using the specified sync key, folder collection ID and body preference option.
        /// </summary>
        /// <param name="syncKey">Specifies the sync key obtained from the last sync response.</param>
        /// <param name="collectionId">Specifies the server ID of the folder to be synchronized.</param>
        /// <param name="bodyPreference">Sets preference information related to the type and size of information for body.</param>
        /// <returns>Returns the SyncRequest instance</returns>
        internal static SyncRequest CreateSyncRequest(string syncKey, string collectionId, Request.BodyPreference bodyPreference)
        {
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                SyncKey = syncKey,
                CollectionId = collectionId
            };

            // Sets Getchanges only if SyncKey != 0 since fail response is returned when the SyncKey element value is 0.
            if (syncKey != "0")
            {
                syncCollection.GetChanges = true;
                syncCollection.GetChangesSpecified = true;
            }

            syncCollection.WindowSize = "512";

            Request.Options syncOptions = new Request.Options();
            List<object> syncOptionItems = new List<object>();
            List<Request.ItemsChoiceType1> syncOptionItemsName = new List<Request.ItemsChoiceType1>();

            if (null != bodyPreference)
            {
                syncOptionItemsName.Add(Request.ItemsChoiceType1.BodyPreference);
                syncOptionItems.Add(bodyPreference);
            }

            syncOptions.Items = syncOptionItems.ToArray();
            syncOptions.ItemsElementName = syncOptionItemsName.ToArray();
            syncCollection.Options = new Request.Options[] { syncOptions };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
        }

        /// <summary>
        /// Builds a Sync change request by using the specified sync key, folder collection ID and change application data.
        /// </summary>
        /// <param name="syncKey">Specifies the sync key obtained from the last sync response.</param>
        /// <param name="collectionId">Specifies the server ID of the folder to be synchronized.</param>
        /// <param name="data">Contains the data used to specify the Change element for Sync command.</param>
        /// <returns>Returns the SyncRequest instance</returns>
        internal static SyncRequest CreateSyncRequest(string syncKey, string collectionId, List<object> data)
        {
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                SyncKey = syncKey,
                GetChanges = true,
                GetChangesSpecified = true,
                DeletesAsMoves = false,
                DeletesAsMovesSpecified = true,
                CollectionId = collectionId
            };

            Request.Options option = new Request.Options();
            Request.BodyPreference preference = new Request.BodyPreference
            {
                Type = 2,
                Preview = 0,
                PreviewSpecified = true
            };

            option.Items = new object[] { preference };
            option.ItemsElementName = new Request.ItemsChoiceType1[]
                {
                    Request.ItemsChoiceType1.BodyPreference,
                };

            syncCollection.Options = new Request.Options[1];
            syncCollection.Options[0] = option;

            syncCollection.WindowSize = "512";
            syncCollection.Commands = data.ToArray();

            return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
        }

        /// <summary>
        /// Whether the content is in HTML format
        /// </summary>
        /// <param name="content">The string to be checked</param>
        /// <returns>Returns the value to represent whether the content is in HTML format or not</returns>
        internal static bool IsHTML(string content)
        {
            Regex reg = new Regex(@"<(html)>(.*\n*)*<\/\1>", RegexOptions.IgnoreCase);
            bool isHTML = reg.IsMatch(content);
            return isHTML;
        }
    }
}