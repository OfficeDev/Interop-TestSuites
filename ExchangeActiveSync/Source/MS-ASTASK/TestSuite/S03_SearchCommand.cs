namespace Microsoft.Protocols.TestSuites.MS_ASTASK
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using SyncItem = Microsoft.Protocols.TestSuites.Common.DataStructures.Sync;
    using SyncStore = Microsoft.Protocols.TestSuites.Common.DataStructures.SyncStore;

    /// <summary>
    /// This scenario is to test Task class element on the server by using Search command.
    /// </summary>
    [TestClass]
    public class S03_SearchCommand : TestSuiteBase
    {
        #region Test Class initialize and clean up

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
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        /// <summary>
        /// This test case is designed to verify the requirements about processing the Search command.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S03_TC01_RetrieveTaskItemWithSearch()
        {
            #region Call Sync command to create a task item

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");
            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);

            #endregion

            #region Call Sync command to add the task to the server

            // add task
            SyncStore syncResponse = this.SyncAddTask(taskItem);
            Site.Assert.AreEqual<byte>(1, syncResponse.AddResponses[0].Status, "Adding a task item to server should success.");
            SyncItem task = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(task.Task, "The task which subject is {0} should exist in server.", subject);
            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Search command to search task on the server

            Request.Options1 option = new Request.Options1();
            Dictionary<Request.ItemsChoiceType6, object> items = new Dictionary<Request.ItemsChoiceType6, object>
            {
                {
                    Request.ItemsChoiceType6.DeepTraversal, string.Empty
                },
                {
                    Request.ItemsChoiceType6.RebuildResults, string.Empty
                },
                {
                    Request.ItemsChoiceType6.Range, "0-9"
                }
            };
            option.Items = items.Values.ToArray<object>();
            option.ItemsElementName = items.Keys.ToArray<Request.ItemsChoiceType6>();

            var queryItem = new Request.queryType
            {
                Items = new object[] { "Tasks", this.UserInformation.TasksCollectionId, subject },

                ItemsElementName = new Request.ItemsChoiceType5[]
                {
                    Request.ItemsChoiceType5.Class,
                    Request.ItemsChoiceType5.CollectionId,
                    Request.ItemsChoiceType5.FreeText
                }
            };

            Request.queryType queryType = new Request.queryType
            {
                Items = new object[] { queryItem },
                ItemsElementName = new Request.ItemsChoiceType5[] { Request.ItemsChoiceType5.And }
            };

            SearchRequest searchRequest = TestSuiteHelper.CreateSearchRequest(SearchName.Mailbox.ToString(), option, queryType);

            // Search the task
            SearchStore searchResponse = this.TASKAdapter.Search(searchRequest);

            // Verify search response
            Site.Assert.AreNotEqual<int>(0, searchResponse.Range.Length, "The search response should be successful");

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R361");

            // Verify MS-ASTASK requirement: MS-ASTASK_R361
            // If the Task item in response is not null, this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                searchResponse.Results[0].Task,
                361,
                @"[In Search Command Response] When a client uses the Search command request ([MS-ASCMD] section 2.2.2.14) to retrieve Task class items from the server that match the criteria specified by the client, as specified in section 3.1.5.2, the server responds with a Search command response ([MS-ASCMD] section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R363");

            // Since MS-ASTASK_R361 is captured, this requirement can be captured too.
            Site.CaptureRequirement(
                363,
                @"[In Search Command Response] Top-level Task class elements, as specified in section 2.2, are returned as child elements of the search:Properties element ([MS-ASCMD] section 2.2.3.128) in the Search command response.");
        }

        /// <summary>
        /// This test case is designed to verify the requirements about processing the Search command with invalid Task item.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S03_TC02_RetrieveInvalidTaskItemWithSearch()
        {
            #region Call Sync command to create a task item

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");
            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);

            #endregion

            #region Call Sync command to add the task to the server

            // add task
            SyncStore syncResponse = this.SyncAddTask(taskItem);
            Site.Assert.AreEqual<byte>(1, syncResponse.AddResponses[0].Status, "Adding a task item to server should success.");
            SyncItem task = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(task.Task, "The task which subject is {0} should exist in server.", subject);
            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call Search command to search task on the server

            // Send a string create task item request without Type element contains in Recurrence.
            SendStringResponse sendStringResponse = this.TASKAdapter.SendStringRequest("<Search xmlns=\"Search\"><Store><Name>Mailbox</Name><Query><And><Class xmlns=\"AirSync\">Tasks</Class><CollectionId xmlns=\"AirSync\">" + this.UserInformation.TasksCollectionId + "</CollectionId><FreeText>FreeText</FreeText><Subject xmlns=\"Tasks\">user</Subject></And></Query><Options><DeepTraversal /></Options></Store></Search>", CommandName.Search);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(sendStringResponse.ResponseDataXML);
            XmlNamespaceManager xnm = new XmlNamespaceManager(doc.NameTable);
            xnm.AddNamespace("e", "Search");

            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 1;
            XmlNode searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);

            while (counter < retryCount && searchStatus != null && searchStatus.InnerXml.Equals("10"))
            {
                Thread.Sleep(waitTime);
                sendStringResponse = this.TASKAdapter.SendStringRequest("<Search xmlns=\"Search\"><Store><Name>Mailbox</Name><Query><And><Class xmlns=\"AirSync\">Tasks</Class><CollectionId xmlns=\"AirSync\">" + this.UserInformation.TasksCollectionId + "</CollectionId><FreeText>FreeText</FreeText><Subject xmlns=\"Tasks\">user</Subject></And></Query><Options><DeepTraversal /></Options></Store></Search>", CommandName.Search);
                doc.LoadXml(sendStringResponse.ResponseDataXML);
                xnm = new XmlNamespaceManager(doc.NameTable);
                xnm.AddNamespace("e", "Search");
                searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);
                counter++;
            }
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R364");

            // Verify MS-ASTASK requirement: MS-ASTASK_R364
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                Common.GetSearchStatusCode(sendStringResponse),
                364,
                @"[In Search Command Response] If elements that belong to the Task class are included in a Search command request, the server MUST return a Status value of 2 in the Store element, as specified in [MS-ASCMD] section 2.2.3.162.12.");
        }
    }
}