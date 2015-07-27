//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASTASK
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using SyncItem = Microsoft.Protocols.TestSuites.Common.DataStructures.Sync;
    using SyncStore = Microsoft.Protocols.TestSuites.Common.DataStructures.SyncStore;

    /// <summary>
    /// The base class of scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Variables

        /// <summary>
        /// Gets the list of existing tasks' subjects.
        /// </summary>
        protected Collection<string> ItemsNeedToDelete { get; private set; }

        /// <summary>
        /// Gets the protocol adapter.
        /// </summary>
        protected IMS_ASTASKAdapter TASKAdapter { get; private set; }

        /// <summary>
        /// Gets the default ActiveSync Protocol Version.
        /// </summary>
        protected string ActiveSyncProtocolVersion { get; private set; }

        /// <summary>
        /// Gets or sets the user information.
        /// </summary>
        protected UserInformation UserInformation { get; set; }

        #endregion

        #region Public methods

        /// <summary>
        /// Synchronize changes between client and server.
        /// </summary>
        /// <param name="collectionId">Specify the folder collection Id which needs to be synchronized.</param>
        /// <returns>Return the sync response.</returns>
        public SyncStore SyncChanges(string collectionId)
        {
            SyncStore syncResponse;

            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 0;

            // Synchronize to get the SyncKey.
            SyncStore initializeSyncResponse = this.TASKAdapter.Sync(Common.CreateInitialSyncRequest(collectionId));
            SyncRequest syncRequest = TestSuiteHelper.CreateSyncRequest(initializeSyncResponse.SyncKey, collectionId, null);

            do
            {
                Thread.Sleep(waitTime);

                // Get the server changes through sync command.
                syncResponse = this.TASKAdapter.Sync(syncRequest);
                if (syncResponse != null)
                {
                    if (syncResponse.CollectionStatus == 1)
                    {
                        break;
                    }
                }

                counter++;
            }
            while (counter < retryCount);

            // Verify sync response
            Site.Assert.AreEqual<byte>(
                1,
                syncResponse.CollectionStatus,
                "If the Sync command executes successfully, the Status in response should be 1.");

            return syncResponse;
        }

        /// <summary>
        /// Get the specified task item from the sync change response.
        /// </summary>
        /// <param name="collectionId">The task folder server id.</param>
        /// <param name="taskSubject">The subject value of task.</param>
        /// <returns>Return the specified task item.</returns>
        public SyncItem GetChangeItem(string collectionId, string taskSubject)
        {
            SyncItem resultItem = null;

            if (collectionId == this.UserInformation.TasksCollectionId)
            {
                // Get the server changes through sync command.
                SyncStore syncResponse = this.SyncChanges(collectionId);

                foreach (SyncItem item in syncResponse.AddElements)
                {
                    if (item.Task.Subject == taskSubject)
                    {
                        resultItem = item;
                        break;
                    }
                }
            }

            return resultItem;
        }

        #endregion

        #region Test case initialize and cleanup

        /// <summary>
        /// Initialize the Test Suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();

            if (this.TASKAdapter == null)
            {
                this.TASKAdapter = Site.GetAdapter<IMS_ASTASKAdapter>();
            }

            this.ActiveSyncProtocolVersion = Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site);
            this.UserInformation = new UserInformation();

            if (Common.GetSutVersion(this.Site) != SutVersion.ExchangeServer2007 || string.Equals(this.ActiveSyncProtocolVersion, "12.1"))
            {
                FolderSyncResponse folderSyncResponse = this.TASKAdapter.FolderSync();
                this.UserInformation.TasksCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Tasks, this.Site);
            }

            this.ItemsNeedToDelete = new Collection<string>();
        }

        /// <summary>
        /// Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            if (this.ItemsNeedToDelete.Count > 0)
            {
                SyncStore changesResult = this.SyncChanges(this.UserInformation.TasksCollectionId);

                foreach (string subject in this.ItemsNeedToDelete)
                {
                    string serverId = null;
                    foreach (Sync add in changesResult.AddElements)
                    {
                        if (add.Task.Subject == subject)
                        {
                            serverId = add.ServerId;
                            break;
                        }
                    }

                    Site.Assert.IsNotNull(serverId, "The task with subject {0} should be found.", subject);

                    SyncStore deleteResult = this.SyncDeleteTask(changesResult.SyncKey, serverId);

                    Site.Assert.AreEqual<byte>(
                        1,
                        deleteResult.CollectionStatus,
                        "The server should return a status code of 1 in the Sync command response indicate sync command succeed.");
                }

                this.ItemsNeedToDelete.Clear();
            }

            base.TestCleanup();
        }

        #endregion

        #region Protected methods

        /// <summary>
        /// Call Sync command to add a task.
        /// </summary>
        /// <param name="addElements">The elements of a task item to be added.</param>
        /// <returns>Return the sync response.</returns>
        protected SyncStore SyncAddTask(Dictionary<Request.ItemsChoiceType8, object> addElements)
        {
            SyncStore initializeSyncResponse = this.TASKAdapter.Sync(Common.CreateInitialSyncRequest(this.UserInformation.TasksCollectionId));

            // Verify sync result
            Site.Assert.AreEqual<byte>(
                1,
                initializeSyncResponse.CollectionStatus,
                "The server should return a status code 1 in the Sync command response to indicate the Sync command executes successfully.");

            Dictionary<Request.ItemsChoiceType8, object> task = TestSuiteHelper.CreateTaskElements();

            // Add elements
            if (addElements != null)
            {
                foreach (KeyValuePair<Request.ItemsChoiceType8, object> item in addElements)
                {
                    if (task.ContainsKey(item.Key))
                    {
                        task[item.Key] = item.Value;
                    }
                    else
                    {
                        task.Add(item.Key, item.Value);
                    }
                }
            }

            List<object> addData = new List<object>();
            Request.SyncCollectionAdd add = new Request.SyncCollectionAdd
            {
                ClientId = System.Guid.NewGuid().ToString(),
                ApplicationData = new Request.SyncCollectionAddApplicationData
                {
                    Items = task.Values.ToArray<object>(),
                    ItemsElementName = task.Keys.ToArray<Request.ItemsChoiceType8>()
                }
            };
            addData.Add(add);

            SyncRequest syncRequest = TestSuiteHelper.CreateSyncRequest(initializeSyncResponse.SyncKey, this.UserInformation.TasksCollectionId, addData);
            SyncStore syncResponse = this.TASKAdapter.Sync(syncRequest);

            Site.Assert.AreEqual<byte>(
                1,
                syncResponse.CollectionStatus,
                "The server should return a Status 1 in the Sync command response indicate sync command succeed.");

            Site.Assert.IsNotNull(
                syncResponse.AddResponses,
                @"The Add elements in Responses element of the Sync response should not be null.");

            return syncResponse;
        }

        /// <summary>
        /// Call Sync command to change a task.
        /// </summary>
        /// <param name="syncKey">The sync key.</param>
        /// <param name="serverId">The server Id of the task.</param>
        /// <param name="changedElements">The changed elements of the task.</param>
        /// <returns>Return the sync change result.</returns>
        protected SyncStore SyncChangeTask(string syncKey, string serverId, Dictionary<Request.ItemsChoiceType7, object> changedElements)
        {
            Request.SyncCollectionChange change = new Request.SyncCollectionChange
            {
                ServerId = serverId,
                ApplicationData = new Request.SyncCollectionChangeApplicationData
                {
                    Items = changedElements.Values.ToArray<object>(),
                    ItemsElementName = changedElements.Keys.ToArray<Request.ItemsChoiceType7>()
                }
            };

            SyncRequest syncRequest = new SyncRequest
            {
                RequestData = new Request.Sync { Collections = new Request.SyncCollection[1] }
            };
            syncRequest.RequestData.Collections[0] = new Request.SyncCollection
            {
                Commands = new object[] { change },
                SyncKey = syncKey,
                CollectionId = this.UserInformation.TasksCollectionId
            };
            SyncStore syncResponse = this.TASKAdapter.Sync(syncRequest);

            return syncResponse;
        }

        /// <summary>
        /// Call Sync command to delete a task.
        /// </summary>
        /// <param name="syncKey">The sync key.</param>
        /// <param name="serverId">The server id of the task, which is returned by server.</param>
        /// <returns>Return the sync delete result.</returns>
        protected SyncStore SyncDeleteTask(string syncKey, string serverId)
        {
            List<object> deleteData = new List<object>();
            Request.SyncCollectionDelete delete = new Request.SyncCollectionDelete { ServerId = serverId };
            deleteData.Add(delete);

            SyncRequest syncRequest = TestSuiteHelper.CreateSyncRequest(syncKey, this.UserInformation.TasksCollectionId, deleteData);
            SyncStore syncResponse = this.TASKAdapter.Sync(syncRequest);

            return syncResponse;
        }

        #endregion
    }
}