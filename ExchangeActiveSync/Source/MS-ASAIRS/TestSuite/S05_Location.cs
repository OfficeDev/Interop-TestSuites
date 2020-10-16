namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to test the Location element and its sub elements, which is used by the Sync command, Search command and ItemOperations command.
    /// </summary>
    [TestClass]
    public class S05_Location : TestSuiteBase
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

        #region MSASAIRS_S05_TC01_Location
        /// <summary>
        /// This case is designed to test element Location and its sub elements.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S05_TC01_Location()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Location element is supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Location element is supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Location element is supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add an appointment to the server
            Request.SyncCollectionAddApplicationData applicationData = new Request.SyncCollectionAddApplicationData();

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType8> itemsElementName = new List<Request.ItemsChoiceType8>();

            string subject = Common.GenerateResourceName(Site, "Subject");
            items.Add(subject);
            itemsElementName.Add(Request.ItemsChoiceType8.Subject);

            // MeetingStauts is set to 0, which means it is an appointment with no attendees.
            byte meetingStatus = 0;
            items.Add(meetingStatus);
            itemsElementName.Add(Request.ItemsChoiceType8.MeetingStatus);

            Request.Location location = new Request.Location();
            location.Accuracy = (double)1;
            location.AccuracySpecified = true;
            location.Altitude = (double)55.46;
            location.AltitudeAccuracy = (double)1;
            location.AltitudeAccuracySpecified = true;
            location.AltitudeSpecified = true;
            location.Annotation = "Location sample annotation";
            location.City = "Location sample city";
            location.Country = "Location sample country";
            location.DisplayName = "Location sample dislay name";
            location.Latitude = (double)11.56;
            location.LatitudeSpecified = true;
            location.LocationUri = "Location Uri";
            location.Longitude = (double)1.9;
            location.LongitudeSpecified = true;
            location.PostalCode = "Location sample postal code";
            location.State = "Location sample state";
            location.Street = "Location sample street";
            items.Add(location);
            itemsElementName.Add(Request.ItemsChoiceType8.Location);

            applicationData.Items = items.ToArray();
            applicationData.ItemsElementName = itemsElementName.ToArray();
            SyncRequest syncAddRequest = TestSuiteHelper.CreateSyncAddRequest(this.GetInitialSyncKey(this.User1Information.CalendarCollectionId), this.User1Information.CalendarCollectionId, applicationData);

            DataStructures.SyncStore syncAddResponse = this.ASAIRSAdapter.Sync(syncAddRequest);
            Site.Assert.IsTrue(syncAddResponse.AddResponses[0].Status.Equals("1"), "The sync add operation should be success; It is:{0} actually", syncAddResponse.AddResponses[0].Status);

            // Add the appointment to clean up list.
            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the new added calendar item.
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null, null, null);
            #endregion

            #region Call ItemOperations command to reterive the added calendar item.
            this.GetItemOperationsResult(this.User1Information.CalendarCollectionId, syncItem.ServerId, null, null, null, null);
            #endregion

            #region Call Sync command to remove the location of the added calender item.
            // Create empty change items list.
            List<object> changeItems = new List<object>();
            List<Request.ItemsChoiceType7> changeItemsElementName = new List<Request.ItemsChoiceType7>();

            // Create an empty location.
            location = new Request.Location();

            // Add the location field name into the change items element name list.
            changeItemsElementName.Add(Request.ItemsChoiceType7.Location);
            // Add the empty location value to the change items value list.
            changeItems.Add(location);

            // Create sync collection change.
            Request.SyncCollectionChange collectionChange = new Request.SyncCollectionChange
            {
                ServerId = syncItem.ServerId,
                ApplicationData = new Request.SyncCollectionChangeApplicationData
                {
                    ItemsElementName = changeItemsElementName.ToArray(),
                    Items = changeItems.ToArray()
                }
            };

            // Create change sync collection
            Request.SyncCollection collection = new Request.SyncCollection
            {
                SyncKey = this.SyncKey,
                CollectionId = this.User1Information.CalendarCollectionId,
                Commands = new object[] { collectionChange }
            };

            // Create change sync request.
            SyncRequest syncChangeRequest = Common.CreateSyncRequest(new Request.SyncCollection[] { collection });

            // Change the location of the added calender by Sync request.
            DataStructures.SyncStore syncChangeResponse = this.ASAIRSAdapter.Sync(syncChangeRequest);
            Site.Assert.IsTrue(syncChangeResponse.CollectionStatus.Equals(1), "The sync change operation should be success; It is:{0} actually", syncChangeResponse.CollectionStatus);

            #region Call Sync command to get the new changed calendar item that removed the location.
            syncItem = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null, null, null);
            #endregion

            #region Call ItemOperations command to reterive the changed calendar item that removed the location.
            DataStructures.ItemOperations itemOperations = this.GetItemOperationsResult(this.User1Information.CalendarCollectionId, syncItem.ServerId, null, null, null, null);
            #endregion
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R1001013");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R1001013
            Site.CaptureRequirementIfIsNull(
                itemOperations.Calendar.Location1.DisplayName,
                1001013,
                @"[In Location] The client's request can include an empty Location element to remove the location from an item.");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Call Search command to search the added calendar item.
                this.GetSearchResult(subject, this.User1Information.CalendarCollectionId, null, null, null);
                #endregion
            }
        }
        #endregion
    }
}
