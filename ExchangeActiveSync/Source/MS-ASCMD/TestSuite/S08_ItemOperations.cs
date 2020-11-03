namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test the ItemOperations command.
    /// </summary>
    [TestClass]
    public class S08_ItemOperations : TestSuiteBase
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
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test Cases
        /// <summary>
        /// This test case is used to test when the DeleteSubFolders element is not present in ItemOperations request, the server will not delete subfolders.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC01_ItemOperations_DeleteSubFolders()
        {
            #region Call method FolderCreate to create a new folder as the child folder of the specified parent folder.
            // Create a user-created mail folder as a subfolder of Inbox folder.
            string subFolderName = Common.GenerateResourceName(Site, "subfolder");
            FolderCreateRequest folderCreateRequest = Common.CreateFolderCreateRequest(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, subFolderName, this.User1Information.InboxCollectionId);
            FolderCreateResponse folderCreateResponse = this.CMDAdapter.FolderCreate(folderCreateRequest);

            Site.Assert.AreEqual<int>(
                1,
                int.Parse(folderCreateResponse.ResponseData.Status),
                "The Status element of the FolderCreate response should be 1.");

            // Record created folder CollectionId.
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderCreateResponse.ResponseData.ServerId);
            #endregion

            #region Call method FolderSync to synchronize the collection hierarchy and get CollectionId of the created subfolder.
            FolderSyncResponse folderSyncResponse = this.FolderSync();
            string subFolderId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.UserCreatedMail, this.Site);
            Site.Assert.IsNotNull(subFolderId, "The CollectionId of the created child folder should not be null.");
            #endregion

            #region Call ItemOperations command which contains the EmptyFolderContents element without DeleteSubFolders element.
            ItemOperationsRequest itemOperationsRequest = new ItemOperationsRequest
            {
                RequestData = new Request.ItemOperations
                {
                    Items = new object[]
                    {
                        new Request.ItemOperationsEmptyFolderContents
                        {
                            CollectionId = this.User1Information.InboxCollectionId
                        }
                    }
                }
            };

            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);

            Site.Assert.AreEqual<string>("1", itemOperationsResponse.ResponseData.Status, "The Status element of the ItemOperations response should be 1.");

            #endregion

            #region Call method FolderSync to synchronize the collection hierarchy and check whether the subfolder exists.
            folderSyncResponse = this.FolderSync();

            Site.Assert.IsNotNull(folderSyncResponse.ResponseData.Changes, "The Changes element of the FolderSync Response should not be null.");
            Response.FolderSyncChangesAdd[] folderChangesAdds = folderSyncResponse.ResponseData.Changes.Add;

            bool subFolderExists = false;
            foreach (Response.FolderSyncChangesAdd folderChangesAdd in folderChangesAdds)
            {
                if (folderChangesAdd.DisplayName == subFolderName && folderChangesAdd.ParentId == this.User1Information.InboxCollectionId)
                {
                    subFolderId = folderChangesAdd.ServerId;
                    subFolderExists = true;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2167, Expected subfolder name: {0}, parent folder's CollectionId: {1}, subfolder's ServerId: {2}", subFolderName, this.User1Information.InboxCollectionId, subFolderId);

            // Verify MS-ASCMD requirement: MS-ASCMD_R2167
            // If the user-created mail subfolder can be found, and ServerId is the same with the one after created, this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                subFolderExists,
                2167,
                @"[In DeleteSubFolders] If the DeleteSubFolders element is not present in the request, the default behavior is to not delete subfolders.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2244");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2244
            // This requirement can be captured directly after MS-ASCMD_R2167
            Site.CaptureRequirement(
                2244,
                @"[In EmptyFolderContents] If the DeleteSubFolders option is not included in the request, the subfolders of the specified airsync:CollectionId are not deleted.");
            #endregion

            #region Call ItemOperations command which contains the EmptyFolderContents element with DeleteSubFolders element.
            itemOperationsRequest = new ItemOperationsRequest
            {
                RequestData = new Request.ItemOperations
                {
                    Items = new object[]
                    {
                        new Request.ItemOperationsEmptyFolderContents
                        {
                            CollectionId = this.User1Information.InboxCollectionId,
                            Options = new Request.ItemOperationsEmptyFolderContentsOptions
                            {
                                DeleteSubFolders = string.Empty
                            }
                        }
                    }
                }
            };

            itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);

            Site.Assert.AreEqual<string>("1", itemOperationsResponse.ResponseData.Status, "The Status element of the ItemOperations response should be 1.");

            // Remove the recorded user name and its corresponding folder's CollectionId.
            this.User1Information.UserCreatedFolders.Clear();
            #endregion

            #region Call method FolderSync to synchronize the collection hierarchy and check whether the subfolder exists.
            folderSyncResponse = this.FolderSync();

            Site.Assert.IsNotNull(folderSyncResponse.ResponseData.Changes, "The Changes element of the FolderSync Response should not be null.");
            folderChangesAdds = folderSyncResponse.ResponseData.Changes.Add;

            subFolderExists = false;
            foreach (Response.FolderSyncChangesAdd folderChangesAdd in folderChangesAdds)
            {
                if (folderChangesAdd.DisplayName == subFolderName && folderChangesAdd.ParentId == this.User1Information.InboxCollectionId)
                {
                    subFolderExists = true;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2243");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2243
            // If the user-created mail subfolder cannot be found, this requirement can be verified.
            Site.CaptureRequirementIfIsFalse(
                subFolderExists,
                2243,
                @"[In EmptyFolderContents] The element supports a single child element of the Options element, the DeleteSubFolders element (section 2.2.3.44), which determines whether subfolders contained in the folder are deleted.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test when the client specifies an option that is invalid for the parent element in ItemOperations request, the server returns a Status element value of 2.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC02_ItemOperations_Status2()
        {
            #region Call method ItemOperations which contains the Fetch element, with an invalid option in Options element.
            Request.ItemOperationsFetchOptions options = new Request.ItemOperationsFetchOptions
            {
                Items = new object[] { string.Empty },
                ItemsElementName = new Request.ItemsChoiceType5[] { Request.ItemsChoiceType5.Range }
            };

            // Set the Range element to empty string as the child element of option element in Fetch element, to trigger the status code 2.
            options.Items = new object[] { string.Empty };
            options.ItemsElementName = new Request.ItemsChoiceType5[] { Request.ItemsChoiceType5.Range };

            ItemOperationsRequest itemOperationsRequest = CreateItemOperationsRequestWithFetchElement(this.User1Information.InboxCollectionId, SearchName.Mailbox.ToString(), string.Empty, options);
            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3547");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3547
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                itemOperationsResponse.ResponseData.Status,
                3547,
                @"[In Options(ItemOperations)] If the client specifies an option that is invalid for the parent element, the server returns a Status element (section 2.2.3.162.7) value of 2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4154");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4154
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                itemOperationsResponse.ResponseData.Status,
                4154,
                @"[In Status(ItemOperations)] [The meaning of the status value] 2 [is] Protocol error - protocol violation/XML validation error.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test when the Store element is unknown or unsupported in ItemOperations command request, the server returns a Status element value of 9.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC03_ItemOperations_Status9()
        {
            #region Call method ItemOperations which contains the Fetch element, with an unsupported value of Store element.
            ItemOperationsRequest itemOperationsRequest = CreateItemOperationsRequestWithFetchElement(this.User1Information.InboxCollectionId, "InvalidStore", string.Empty, null);
            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Response.ItemOperationsResponseFetch fetch = this.GetFetchElement(itemOperationsResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4161");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4161
            Site.CaptureRequirementIfAreEqual<string>(
                "9",
                fetch.Status,
                4161,
                @"[In Status(ItemOperations)] [The meaning of the status value] 9 [is] The store is unknown or unsupported.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test when the object was not found in ItemOperations request, the server returns a Status element value of 6.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC04_ItemOperations_Status6()
        {
            #region Call method ItemOperations which contains the Fetch element, with an invalid ServerID.
            ItemOperationsRequest itemOperationsRequest = CreateItemOperationsRequestWithFetchElement(this.User1Information.InboxCollectionId, SearchName.Mailbox.ToString(), "InvalidServerID", null);
            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Response.ItemOperationsResponseFetch fetch = this.GetFetchElement(itemOperationsResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4158");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4158
            Site.CaptureRequirementIfAreEqual<string>(
                "6",
                fetch.Status,
                4158,
                @"[In Status(ItemOperations)] [The meaning of the status value] 6 [is] Document library - The object was not found or access denied.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test Fetch element in ItemOperations command. 
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC05_ItemOperations_Fetch()
        {
            #region User2 sends mail to User1 and does FolderSync in User1's mailbox.
            string subject = this.SendMailAndFolderSync();
            #endregion

            #region Call method Sync to synchronize changes of Inbox folder in User1's mailbox between the client and the server.
            SyncResponse syncResponse = this.GetMailItem(this.User1Information.InboxCollectionId, subject);
            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", subject);
            #endregion

            #region Call method ItemOperations which contains the Fetch element.
            ItemOperationsRequest itemOperationsRequest = CreateItemOperationsRequestWithFetchElement(this.User1Information.InboxCollectionId, SearchName.Mailbox.ToString(), serverId, null);
            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);

            Response.ItemOperationsResponseFetch fetch = this.GetFetchElement(itemOperationsResponse);
            Site.Assert.IsNotNull(fetch.Properties, "The Properties element of the ItemOperations response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4153");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4153
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                fetch.Status,
                4153,
                @"[In Status(ItemOperations)] [The meaning of the status value] 1 [is] Success.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test Move element in ItemOperations method.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC06_ItemOperations_Move()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Move element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region User2 sends mail to User1 and does FolderSync in User1's mailbox.
            string subject = this.SendMailAndFolderSync();
            #endregion

            #region Call method Sync to synchronize changes of Inbox folder in User1's mailbox between the client and the server.
            SyncResponse syncResponse = this.GetMailItem(this.User1Information.InboxCollectionId, subject);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, subject);
            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", subject);
            string conversationId = GetConverstationIdFromSyncResposne(syncResponse, serverId);
            #endregion

            #region Call method ItemOperations which contains the Move element.
            ItemOperationsRequest itemOperationsRequest = new ItemOperationsRequest
            {
                RequestData = new Request.ItemOperations
                {
                    Items = new object[]
                    {
                        new Request.ItemOperationsMove
                        {
                            ConversationId = conversationId,
                            DstFldId = this.User1Information.DeletedItemsCollectionId,
                            Options = new Request.ItemOperationsMoveOptions
                            {
                                MoveAlways = string.Empty
                            }
                        }
                    }
                }
            };

            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);

            // Record user name, folder's CollectionId and item subject that are used in this case.
            syncResponse = this.GetMailItem(this.User1Information.DeletedItemsCollectionId, subject);
            serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", subject);
            Site.Assert.IsNotNull(serverId, "The serverId of expected item should not be null");
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, subject);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.DeletedItemsCollectionId, subject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R201");

            // Verify MS-ASCMD requirement: MS-ASCMD_R201
            Site.CaptureRequirementIfAreEqual<string>(
                 "1",
                itemOperationsResponse.ResponseData.Status,
                201,
                @"[In ItemOperations] The Move operation moves a conversation to a destination folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3465");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3465
            Site.CaptureRequirementIfAreEqual<string>(
                 "1",
                itemOperationsResponse.ResponseData.Status,
                3465,
                @"[In Move(ItemOperations)] It [Move element] identifies the body of the request or response as containing the operation that moves a given conversation.<52>");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify fetching an email item with a multipart response.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC07_ItemOperations_MultiplePartResponse()
        {
            #region User2 sends mail to User1 and does FolderSync in User1's mailbox.
            string subject = this.SendMailAndFolderSync();
            #endregion

            #region Call method Sync to synchronize changes of Inbox folder in User1's mailbox between the client and the server.
            SyncResponse syncResponse = this.GetMailItem(this.User1Information.InboxCollectionId, subject);
            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", subject);
            #endregion

            #region Call method ItemOperations which contains the Fetch element and with MultiPart set.
            ItemOperationsRequest itemOperationsRequest = CreateItemOperationsRequestWithFetchElement(this.User1Information.InboxCollectionId, SearchName.Mailbox.ToString(), serverId, null);
            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.MultiPart);

            Response.ItemOperationsResponseFetch fetch = this.GetFetchElement(itemOperationsResponse);
            Site.Assert.IsNotNull(fetch.Properties, "The Properties element of the ItemOperations response should not be null.");

            // Multipart responses have a Content-Type header with the value "application/vnd.ms-sync.multipart".
            Site.Assert.IsNotNull(itemOperationsResponse.Headers, "The headers of ItemOperations response should not be null.");
            Site.Assert.IsNotNull(itemOperationsResponse.Headers["Content-Type"], "The Content-Type header of the ItemOperations response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5505, header value is:{0}", itemOperationsResponse.Headers.ToString());

            // Verify MS-ASCMD requirement: MS-ASCMD_R5505
            Site.CaptureRequirementIfIsTrue(
                itemOperationsResponse.Headers["Content-Type"].ToString().Contains("application/vnd.ms-sync.multipart"),
                5505,
                @"[In Delivery of Content Requested by Fetch] Multipart responses have a Content-Type header with the value ""application/vnd.ms-sync.multipart"".");

            Site.Assert.IsNotNull(itemOperationsResponse.MultipartMetadata, "The metadata of ItemOperations multipart response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5509");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5509
            Site.CaptureRequirementIfAreEqual<int>(
                itemOperationsResponse.MultipartMetadata.PartsMetaData.Length,
                itemOperationsResponse.MultipartMetadata.PartsCount,
                5509,
                @"[In MultiPartResponse] PartsCount (4 bytes): This field [PartsCount] is an integer that indicates the number of PartMetaData structures contained in the PartsMetaData field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5512");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5512
            Site.CaptureRequirementIfAreEqual<int>(
                itemOperationsResponse.MultipartMetadata.PartsCount,
                itemOperationsResponse.MultipartMetadata.PartsMetaData.Length,
                5512,
                @"[In MultiPartResponse] The length of the array is equal to the value of the PartsCount field.");

            byte[][] parts = new byte[itemOperationsResponse.MultipartMetadata.PartsMetaData.Length][];
            for (int partsIndex = 0; partsIndex < itemOperationsResponse.MultipartMetadata.PartsMetaData.Length; partsIndex++)
            {
                parts[partsIndex] = new byte[itemOperationsResponse.MultipartMetadata.PartsMetaData[partsIndex].Count];

                // Get the content of each part
                System.Array.Copy(itemOperationsResponse.RawBody, itemOperationsResponse.MultipartMetadata.PartsMetaData[partsIndex].Start, parts[partsIndex], 0, itemOperationsResponse.MultipartMetadata.PartsMetaData[partsIndex].Count);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5517");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5517
                Site.CaptureRequirementIfIsInstanceOfType(
                    itemOperationsResponse.MultipartMetadata.PartsMetaData[partsIndex].Start,
                    typeof(int),
                    5517,
                    @"[In PartMetaData] Offset (4 bytes): This field [Offset] is an integer that contains the number of bytes from the beginning of the MultiPartResponse structure where a part's data is located.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5518");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5518
                Site.CaptureRequirementIfIsInstanceOfType(
                    itemOperationsResponse.MultipartMetadata.PartsMetaData[partsIndex].Count,
                    typeof(int),
                    5518,
                    @"[In PartMetaData] Length (4 bytes): This field [Length] is an integer that contains the number of bytes that comprise the part's data.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5510");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5510
            Site.CaptureRequirementIfAreEqual<int>(
                parts.Length,
                itemOperationsResponse.MultipartMetadata.PartsCount,
                5510,
                @"[In MultiPartResponse] It [PartsCount] also indicates the number of binary parts contained in the Parts field.");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify if the MIMESupport option is included in the request, the server will return MIME for S/MIME-only messages, all messages, or no messages.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC08_ItemOperations_Fetch_MIMESupport()
        {
            #region User2 sends mail to User1 and does FolderSync in User1's mailbox.
            string subject = this.SendMailAndFolderSync();
            #endregion

            #region Call method Sync to synchronize changes of Inbox folder in User1's mailbox between the client and the server.
            SyncResponse syncResponse = this.GetMailItem(this.User1Information.InboxCollectionId, subject);
            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", subject);
            #endregion

            #region Call method ItemOperations which contains the Fetch element with inline method and MIMESupport setting to 0.
            Request.ItemOperationsFetchOptions options = new Request.ItemOperationsFetchOptions();
            Request.BodyPreference bodyPreference = new Request.BodyPreference { Type = 4 };

            // Set the value of MIMESupport element to 0, which indicates never to send MIME data.
            options.Items = new object[] { (byte)0, bodyPreference };
            options.ItemsElementName = new Request.ItemsChoiceType5[] { Request.ItemsChoiceType5.MIMESupport, Request.ItemsChoiceType5.BodyPreference };
            ItemOperationsRequest itemOperationsRequest = CreateItemOperationsRequestWithFetchElement(this.User1Information.InboxCollectionId, SearchName.Mailbox.ToString(), serverId, options);
            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Response.Body body = this.GetBodyElement(itemOperationsResponse);
            Site.Assert.IsNull(body.Data, "There should be no MIME data returned when MIMESupport setting to 0.");
            #endregion

            #region Call method ItemOperations which contains the Fetch element with inline method and MIMESupport setting to 1.

            // Set the value of MIMESupport element to 1, which indicates to send MIME data for S/MIME messages only, and send regular body for all other messages.
            // In this case, the response of this step contains no MIME data.
            options.Items = new object[] { (byte)1, bodyPreference };
            itemOperationsRequest = CreateItemOperationsRequestWithFetchElement(this.User1Information.InboxCollectionId, SearchName.Mailbox.ToString(), serverId, options);
            itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            body = this.GetBodyElement(itemOperationsResponse);
            Site.Assert.IsNull(body.Data, "There should be no MIME data returned for normal message when MIMESupport setting to 1.");
            #endregion

            #region Call method ItemOperations which contains the Fetch element with inline method and MIMESupport setting to 2.

            // Set the value of MIMESupport element to 2, which indicates to send MIME data for all messages.
            options.Items = new object[] { (byte)2, bodyPreference };
            itemOperationsRequest = CreateItemOperationsRequestWithFetchElement(this.User1Information.InboxCollectionId, SearchName.Mailbox.ToString(), serverId, options);
            itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            body = this.GetBodyElement(itemOperationsResponse);
            Site.Assert.IsNotNull(body.Data, "There should be MIME data returned for normal message when MIMESupport setting to 2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3366");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3366
            Site.CaptureRequirementIfAreEqual<byte>(
                4,
                body.Type,
                3366,
                @"[In MIMESupport(ItemOperations)] The airsyncbase:Type element ([MS-ASAIRS] section 2.2.2.22.1) with a value of 4 to inform the device that the data is a MIME BLOB.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3367");

            Site.Assert.IsTrue(body.EstimatedDataSizeSpecified, "The EstimatedDataSize element should be specified in response when MIMESupport setting to 2.");
            XmlElement lastRawResponse = (XmlElement)this.CMDAdapter.LastRawResponseXml;
            string data = GetElementInnerText(lastRawResponse, "Body", "Data", subject);

            // Verify MS-ASCMD requirement: MS-ASCMD_R3367
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)data.Length,
                body.EstimatedDataSize,
                3367,
                @"[In MIMESupport(ItemOperations)] The airsyncbase:EstimatedDataSize element ([MS-ASAIRS] section 2.2.2.12.2) to specify the rough total size of the data.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3368");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3368
            // If response Body element contains data element, that means response contains full MIME BLOB data then MS-ASCMD_R3368 is verified.
            Site.CaptureRequirementIfIsNotNull(
                body.Data,
                3368,
                @"[In MIMESupport(ItemOperations)] The airsyncbase:Data element ([MS-ASAIRS] section 2.2.2.10.1) that contains the full MIME BLOB.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3360");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3360
            // This requirement can be capture after above steps.
            Site.CaptureRequirement(
                3360,
                @"[In MIMESupport(ItemOperations)] [To support fetching of the full S/MIME message, the Fetch (section 2.2.3.63.1) request MUST include] The airsync:MIMESupport element (section 2.2.3.100.3) to indicate to the server to return MIME for S/MIME-only messages, all messages, or no messages.");
            #endregion

            #region Call method SendStringRequest to call ItemOperations command without specifying delivery method and with MIMESupport setting to 2.
            string request = itemOperationsRequest.GetRequestDataSerializedXML();
            SendStringResponse sendStringResponse = this.CMDAdapter.SendStringRequest(CommandName.ItemOperations, null, request);
            Site.Assert.IsNotNull(sendStringResponse.ResponseDataXML, "The XML data in response should not be null.");
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(sendStringResponse.ResponseDataXML);
            XmlElement xmlElement = (XmlElement)doc.DocumentElement;
            string defaultData = GetElementInnerText(xmlElement, "Body", "Data", subject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5503");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5503
            // If the data gotten by ItemOperations command with inline method is equal to the data got by ItemOperations without specifying delivery method, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                defaultData,
                data,
                5503,
                @"[In Delivery of Content Requested by Fetch] The inline method is the default method for ItemOperations responses.");
            #endregion

            #region Call method ItemOperations which contains the Fetch element with MultiPart method and MIMESupport setting to 2.
            itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.MultiPart);
            body = this.GetBodyElement(itemOperationsResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5366");

            // When the item being fetched is located in a mailbox, server returns the Part element as a child of the airsyncbase:Body.
            // Verify MS-ASCMD requirement: MS-ASCMD_R5366
            Site.CaptureRequirementIfIsNotNull(
                body.Part,
                5366,
                @"[In Part] It [Part element] is a child of the airsyncbase:Body element if the item being fetched is located in a mailbox.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the ItemOperations command cannot perform operations on items in the recipient information cache.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC09_ItemOperations_RecipientInformationCache()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recipient information cache is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call ItemOperations command which contains the EmptyFolderContents element on recipient information cache, and check the Status element for this operation.
            ItemOperationsRequest itemOperationsRequest = new ItemOperationsRequest
            {
                RequestData = new Request.ItemOperations
                {
                    Items = new object[]
                    {
                        new Request.ItemOperationsEmptyFolderContents
                        {
                            CollectionId = this.User1Information.RecipientInformationCacheCollectionId
                        }
                    }
                }
            };

            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);

            Site.Assert.AreEqual<string>("1", itemOperationsResponse.ResponseData.Status, "The Status element of the ItemOperations response should be 1.");

            Site.Assert.IsNotNull(itemOperationsResponse.ResponseData.Response, "The Response element of the ItemOperations response should not be null.");
            Site.Assert.IsNotNull(itemOperationsResponse.ResponseData.Response.EmptyFolderContents, "The EmptyFolderContents element of the ItemOperations response should not be null.");
            Site.Assert.AreEqual<int>(1, itemOperationsResponse.ResponseData.Response.EmptyFolderContents.Length, "The length of EmptyFolderContents element array should be 1.");
            Site.Assert.IsNotNull(itemOperationsResponse.ResponseData.Response.EmptyFolderContents[0], "The EmptyFolderContents element of the ItemOperations response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R196");

            // Verify MS-ASCMD requirement: MS-ASCMD_R196
            // If the status is not 1, which means the operation is not successful.
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                itemOperationsResponse.ResponseData.Response.EmptyFolderContents[0].Status,
                196,
                @"[In ItemOperations] The ItemOperations command cannot perform operations on items in the recipient information cache.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test EmptyFolderContents element in ItemOperations command.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC10_ItemOperations_EmptyFolderContents()
        {
            #region Create a sub-folder in User1's Contacts folder.
            string folderId = this.CreateFolder((byte)FolderType.UserCreatedContacts, Common.GenerateResourceName(Site, "FolderCreate"), this.User1Information.ContactsCollectionId);
            Site.Assert.IsFalse(string.IsNullOrEmpty(folderId), "The sub-folder should be created successfully under the Contacts' folder.");
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderId);
            this.FolderSync();
            #endregion

            #region Add a new contact with Sync operation.
            string contactFileAS = Common.GenerateResourceName(Site, "FileAS");
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);
            this.Sync(TestSuiteBase.CreateEmptySyncRequest(folderId));
            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, folderId, addData);
            SyncResponse syncResponse = this.Sync(syncRequest, false);
            Site.Assert.IsNotNull(syncResponse.ResponseData.Item, "The items returned in the Sync command response should not be null.");
            Response.SyncCollectionsCollectionResponses collectionResponse = TestSuiteBase.GetCollectionItem(syncResponse, Response.ItemsChoiceType10.Responses) as Response.SyncCollectionsCollectionResponses;
            Site.Assert.IsNotNull(collectionResponse, "The responses element should exist in the Sync response.");
            Site.Assert.AreEqual<string>("1", collectionResponse.Add[0].Status, "The new contact should be added correctly.");
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, folderId, contactFileAS);
            #endregion

            #region Call ItemOperations command which contains the EmptyFolderContents element.
            ItemOperationsRequest itemOperationsRequest = new ItemOperationsRequest
            {
                RequestData = new Request.ItemOperations
                {
                    Items = new object[]
                    {
                        new Request.ItemOperationsEmptyFolderContents
                        {
                            CollectionId = folderId
                        }
                    }
                }
            };

            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Site.Assert.AreEqual<string>("1", itemOperationsResponse.ResponseData.Status, "The Status element of the ItemOperations response should be 1.");
            #endregion

            #region Verify if the contact has been removed from User1's Contacts folder.
            syncResponse = this.SyncChanges(folderId);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R199");

            // If the EmptyFolderContents operation enables the client to empty a folder of all its items, all items will be deleted form the specified folder.
            // Verify MS-ASCMD requirement: MS-ASCMD_R199
            Site.CaptureRequirementIfIsNull(
                syncResponse.ResponseData.Item,
                199,
                @"[In ItemOperations] The EmptyFolderContents operation enables the client to empty a folder of all its items.");

            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User1Information, folderId, contactFileAS);
            #endregion
        }

        /// <summary>
        /// This test case is used to verify fetching the specified document library.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC11_ItemOperations_Fetch_DocumentLibrary()
        {
            #region Call method ItemOperations to get data of a shared document with inline method.
            // Initialize a Fetch element.
            Request.ItemOperationsFetch fetch = new Request.ItemOperationsFetch
            {
                LinkId = Common.GetConfigurationPropertyValue("SharedDocument1", this.Site),
                Store = SearchName.DocumentLibrary.ToString(),
                Options = new Request.ItemOperationsFetchOptions
                {
                    ItemsElementName = new Request.ItemsChoiceType5[] { Request.ItemsChoiceType5.UserName, Request.ItemsChoiceType5.Password },
                    Items = new string[] { this.User1Information.UserName, this.User1Information.UserPassword }
                }
            };

            // Create an ItemOperations command request.
            ItemOperationsRequest itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { fetch });

            // Send request and get response.
            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            string data = this.GetDataElement(itemOperationsResponse);
            Site.Assert.IsNotNull(data, "The Data element of the ItemOperations response should not be null.");
            byte[] contents = new byte[0];

            try
            {
                contents = Convert.FromBase64String(data);
            }
            catch (FormatException formatException)
            {
                Site.Assert.Fail("The binary content should be base64 encoded. Detail exception message: {0}", formatException.Message);
            }

            // If data string can be successful decoded with base64, that means the data string is encoded with base64 encoding, then MS-ASCMD_R234, MS-ASCMD_R2301 are verified.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R234");

            // Verify MS-ASCMD requirement: MS-ASCMD_R234
            Site.CaptureRequirement(
                234,
                @"[In Delivery of Content Requested by Fetch] The inline method of delivering binary content is including data encoded with base64 encoding inside the WBXML.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2301");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2301
            Site.CaptureRequirement(
                2301,
                @"[In Fetch(ItemOperations)] Because the ItemOperations response potentially contains large amounts of binary data, this protocol [MS-ASCMD] enables the client to choose a delivery method that is most efficient for its implementation by providing the following two methods to deliver content that is requested by the Fetch element: Inline-The binary content is encoded with base64 encoding and is included inside the WBXML.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5919");

            bool existVersionElement = false;

            for (int i = 0; i < itemOperationsResponse.ResponseData.Response.Fetch[0].Properties.ItemsElementName.Length; i++)
            {
                if (itemOperationsResponse.ResponseData.Response.Fetch[0].Properties.ItemsElementName[i] == Response.ItemsChoiceType3.Version)
                {
                    existVersionElement = true;
                    break;
                }
            }

            // Verify MS-ASCMD requirement: MS-ASCMD_R5919
            Site.CaptureRequirementIfIsTrue(
                existVersionElement,
                5919,
                @"[In Version] The Version element is present only when the ItemOperations command is used to access a [Windows SharePoint Services or] UNC resource.");

            #endregion

            #region Call method ItemOperations to get data of a shared document with MultiPart method.
            // Send request and get response.
            itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.MultiPart);
            data = this.GetDataElement(itemOperationsResponse);
            Site.Assert.IsNull(data, "The Data element of the ItemOperations response should be null.");

            bool existPartElement = false;

            for (int i = 0; i < itemOperationsResponse.ResponseData.Response.Fetch[0].Properties.ItemsElementName.Length; i++)
            {
                if (itemOperationsResponse.ResponseData.Response.Fetch[0].Properties.ItemsElementName[i] == Response.ItemsChoiceType3.Part)
                {
                    existPartElement = true;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2302");

            Site.Assert.IsTrue(itemOperationsResponse.MultipartMetadata.PartsCount >= 2, "There should be at least 2 parts, since the first part is WBXML, the subsequent parts contains non-empty data in txt file. Actual parts count: {0}", itemOperationsResponse.MultipartMetadata.PartsCount);
            int contentLength = 0;
            for (int i = 1; i < itemOperationsResponse.MultipartMetadata.PartsMetaData.Length; i++)
            {
                contentLength += itemOperationsResponse.MultipartMetadata.PartsMetaData[i].Count;
            }

            Site.Assert.IsTrue(contentLength > 0, "The subsequent parts should contains non-empty data in txt file, the length of the data in txt file: {0}.", contentLength);

            // Verify MS-ASCMD requirement: MS-ASCMD_R2302
            // The multipart structure have been implemented in the test suite's adapter, and the WBXML data in first part is converted to XML string, and de-serialized as proxy class in Response element,
            // There should be at least 2 parts, since the first part is WBXML, and is de-serialized into Response element, the subsequent parts contains non-empty data in txt file.
            Site.CaptureRequirement(
                2302,
                @"[In Fetch(ItemOperations)] [Because the ItemOperations response potentially contains large amounts of binary data, this protocol [MS-ASCMD] enables the client to choose a delivery method that is most efficient for its implementation by providing the following two methods to deliver content that is requested by the Fetch element:] Multipart-This method involves a multipart structure in which the WBXML is the first part, and the requested data populates the subsequent parts.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R236");

            // Verify MS-ASCMD requirement: MS-ASCMD_R236
            // This requirement can be captured after above steps.
            Site.CaptureRequirement(
                236,
                @"[In Delivery of Content Requested by Fetch] The multipart method of delivering content is a multipart structure with the WBXML being the first part, and the requested data populating the subsequent parts.");

            #endregion

            #region Call method ItemOperations to get data of a shared document with inline method and Range element specified as "0-0" to fetch the first byte.

            // Initialize a Fetch element.
            fetch = new Request.ItemOperationsFetch
            {
                LinkId = Common.GetConfigurationPropertyValue("SharedDocument1", this.Site),
                Store = SearchName.DocumentLibrary.ToString(),
                Options = new Request.ItemOperationsFetchOptions
                {
                    ItemsElementName = new Request.ItemsChoiceType5[] { Request.ItemsChoiceType5.UserName, Request.ItemsChoiceType5.Password, Request.ItemsChoiceType5.Range },
                    Items = new string[] { this.User1Information.UserName, this.User1Information.UserPassword, "0-0" }
                }
            };

            // Create an ItemOperations command request.
            itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { fetch });

            // Send request and get response.
            itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            data = this.GetDataElement(itemOperationsResponse);
            Site.Assert.IsNotNull(data, "The Data element of the ItemOperations response should not be null.");
            byte[] firstByte = new byte[0];
            try
            {
                firstByte = Convert.FromBase64String(data);
            }
            catch (FormatException formatException)
            {
                Site.Assert.Fail("The binary content should be base64 encoded. Detail exception message: {0}", formatException.Message);
            }

            Site.Assert.AreEqual<int>(1, firstByte.Length, "There should be only 1 byte returned as specified in Range element.");
            int total = 0;

            for (int i = 0; i < itemOperationsResponse.ResponseData.Response.Fetch[0].Properties.ItemsElementName.Length; i++)
            {
                if (itemOperationsResponse.ResponseData.Response.Fetch[0].Properties.ItemsElementName[i] == Response.ItemsChoiceType3.Total)
                {
                    total = int.Parse(itemOperationsResponse.ResponseData.Response.Fetch[0].Properties.Items[i].ToString());
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3716");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3716
            Site.CaptureRequirementIfAreEqual<int>(
                total,
                contents.Length,
                3716,
                @"[In Range(ItemOperations)] If the Range element is omitted in the Fetch request, the whole item is fetched.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5422");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5422
            // If server responds with a successful status code, which including the Range element as "0-0", this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                itemOperationsResponse.ResponseData.Response.Fetch[0].Status,
                5422,
                @"[In Range] The byte range is zero-indexed.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5423");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5423
            // If the returned byte equals to the first byte of the document's content, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                contents[0],
                firstByte[0],
                5423,
                @"[In Range] The first byte is indicated by a 0 (zero).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the status 4 of ItemOperations command indicated the specified URI of document library is bad.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC12_ItemOperations_Status4()
        {
            #region Call method ItemOperations to get data of an invalid shared document with inline fetch.
            // Initialize a Fetch element.
            Request.ItemOperationsFetch itemOperationsFetch = new Request.ItemOperationsFetch
            {
                LinkId = "//invalidURI",
                Store = SearchName.DocumentLibrary.ToString(),
                Options = new Request.ItemOperationsFetchOptions
                {
                    ItemsElementName = new Request.ItemsChoiceType5[] { Request.ItemsChoiceType5.UserName, Request.ItemsChoiceType5.Password },
                    Items = new string[] { this.User1Information.UserName, this.User1Information.UserPassword }
                }
            };

            // Create an ItemOperations command request.
            ItemOperationsRequest itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { itemOperationsFetch });

            // Send request and get response.
            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);

            Response.ItemOperationsResponseFetch fetch = this.GetFetchElement(itemOperationsResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4156");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4156
            Site.CaptureRequirementIfAreEqual<string>(
                "4",
                fetch.Status,
                4156,
                @"[In Status(ItemOperations)] [The meaning of the status value] 4 [is] Document library access - The specified URI is bad.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the status 5 of ItemOperations command indicated access denied.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC13_ItemOperations_Status5()
        {
            #region Call method ItemOperations to get data of a shared document by a user with an invalid password.
            // Initialize a Fetch element.
            Request.ItemOperationsFetch itemOperationsFetch = new Request.ItemOperationsFetch
            {
                LinkId = Common.GetConfigurationPropertyValue("SharedDocument1", this.Site),
                Store = SearchName.DocumentLibrary.ToString(),
                Options = new Request.ItemOperationsFetchOptions
                {
                    ItemsElementName = new Request.ItemsChoiceType5[] { Request.ItemsChoiceType5.UserName, Request.ItemsChoiceType5.Password },
                    Items = new string[] { this.User1Information.UserName, "invalidPassword" }
                }
            };

            // Create an ItemOperations command request.
            ItemOperationsRequest itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { itemOperationsFetch });

            // Send request and get response.
            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Response.ItemOperationsResponseFetch fetch = this.GetFetchElement(itemOperationsResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4157");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4157
            Site.CaptureRequirementIfAreEqual<string>(
                "5",
                fetch.Status,
                4157,
                @"[In Status(ItemOperations)] [The meaning of the status value] 5 [is] Document library - Access denied.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the status 8 of ItemOperations command indicated the byte-range is too large.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC14_ItemOperations_Status8()
        {
            #region Call method ItemOperations to get data of a shared document with the byte range exceeded the actual length of document contents.
            // Initialize a Fetch element.
            Request.ItemOperationsFetch itemOperationsFetch = new Request.ItemOperationsFetch
            {
                LinkId = Common.GetConfigurationPropertyValue("SharedDocument1", this.Site),
                Store = SearchName.DocumentLibrary.ToString(),
                Options = new Request.ItemOperationsFetchOptions
                {
                    ItemsElementName = new Request.ItemsChoiceType5[] { Request.ItemsChoiceType5.UserName, Request.ItemsChoiceType5.Password, Request.ItemsChoiceType5.Range },
                    Items = new string[] { this.User1Information.UserName, this.User1Information.UserPassword, "99-99" }
                }
            };

            // Create an ItemOperations command request.
            ItemOperationsRequest itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { itemOperationsFetch });

            // Send request and get response.
            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Response.ItemOperationsResponseFetch fetch = this.GetFetchElement(itemOperationsResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4160");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4160
            Site.CaptureRequirementIfAreEqual<string>(
                "8",
                fetch.Status,
                4160,
                @"[In Status(ItemOperations)] [The meaning of the status value] 8 [is] The byte-range is invalid or too large.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the status 10 of ItemOperations command indicated the file is empty.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC15_ItemOperations_Status10()
        {
            #region Call method ItemOperations to get data of an empty shared document: SharedDocument2.
            // Initialize a Fetch element.
            Request.ItemOperationsFetch itemOperationsFetch = new Request.ItemOperationsFetch
            {
                LinkId = Common.GetConfigurationPropertyValue("SharedDocument2", this.Site),
                Store = SearchName.DocumentLibrary.ToString(),
                Options = new Request.ItemOperationsFetchOptions
                {
                    ItemsElementName = new Request.ItemsChoiceType5[] { Request.ItemsChoiceType5.UserName, Request.ItemsChoiceType5.Password },
                    Items = new string[] { this.User1Information.UserName, this.User1Information.UserPassword }
                }
            };

            // Create an ItemOperations command request.
            ItemOperationsRequest itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { itemOperationsFetch });

            // Send request and get response.
            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Response.ItemOperationsResponseFetch fetch = this.GetFetchElement(itemOperationsResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4162");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4160
            Site.CaptureRequirementIfAreEqual<string>(
                "10",
                fetch.Status,
                4162,
                @"[In Status(ItemOperations)] [The meaning of the status value] 10 [is] The file is empty.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the status 11 of ItemOperations command indicated the requested data size is too large.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC16_ItemOperations_Status11()
        {
            #region Call method ItemOperations to get data of a shared document with a too large value in Range element exceeded the maximum range value.
            // Initialize a Fetch element.
            Request.ItemOperationsFetch itemOperationsFetch = new Request.ItemOperationsFetch
            {
                LinkId = Common.GetConfigurationPropertyValue("SharedDocument1", this.Site),
                Store = SearchName.DocumentLibrary.ToString(),
                Options = new Request.ItemOperationsFetchOptions
                {
                    ItemsElementName = new Request.ItemsChoiceType5[] { Request.ItemsChoiceType5.UserName, Request.ItemsChoiceType5.Password, Request.ItemsChoiceType5.Range },
                    Items = new string[] { this.User1Information.UserName, this.User1Information.UserPassword, "0-10240000" }
                }
            };

            // Create an ItemOperations command request.
            ItemOperationsRequest itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { itemOperationsFetch });

            // Send request and get response.
            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Response.ItemOperationsResponseFetch fetch = this.GetFetchElement(itemOperationsResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4163");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4163
            Site.CaptureRequirementIfAreEqual<string>(
                "11",
                fetch.Status,
                4163,
                @"[In Status(ItemOperations)] [The meaning of the status value] 11 [is] The requested data size is too large.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the status 15 of ItemOperations command indicated attachment ID is invalid.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC17_ItemOperations_Status15()
        {
            #region Call method ItemOperations to get data of an attachment with invalid attachment ID.
            // Initialize a Fetch element.
            Request.ItemOperationsFetch itemOperationsFetch = new Request.ItemOperationsFetch
            {
                FileReference = "InvalidAttachmentID",
                Store = SearchName.Mailbox.ToString()
            };

            // Create an ItemOperations command request.
            ItemOperationsRequest itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { itemOperationsFetch });

            // Send request and get response.
            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Response.ItemOperationsResponseFetch fetch = this.GetFetchElement(itemOperationsResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4166");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4166
            Site.CaptureRequirementIfAreEqual<string>(
                "15",
                fetch.Status,
                4166,
                @"[In Status(ItemOperations)] [The meaning of the status value] 15 [is] Attachment fetch provider - Attachment or attachment ID is invalid.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the status 2 of ItemOperations command when the client specifies an option that is invalid for the parent element.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC18_ItemOperations_MoreThanOneFileReference_Status2()
        {
            #region Send a mail with normal attachment
            string subject = Common.GenerateResourceName(Site, "NormalAttachment_Subject");
            string body = Common.GenerateResourceName(Site, "NormalAttachment_Body");
            this.SendEmailWithAttachment(subject, body);
            #endregion

            #region Call method Sync to synchronize changes of Inbox folder in User2's mailbox between the client and the server.
            this.SwitchUser(this.User2Information);
            SyncResponse syncResponse = this.GetMailItem(this.User2Information.InboxCollectionId, subject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, subject);

            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", subject);
            Response.Attachments attachments = null;
            Response.SyncCollectionsCollection collection = ((Response.SyncCollections)syncResponse.ResponseData.Item).Collection[0];
            for (int i = 0; i < collection.ItemsElementName.Length; i++)
            {
                if (collection.ItemsElementName[i] == Response.ItemsChoiceType10.Commands)
                {
                    Response.SyncCollectionsCollectionCommands commands = ((Response.SyncCollections)syncResponse.ResponseData.Item).Collection[0].Items[i] as Response.SyncCollectionsCollectionCommands;
                    foreach (Response.SyncCollectionsCollectionCommandsAdd add in commands.Add)
                    {
                        if (add.ServerId.Equals(serverId))
                        {
                            for (int j = 0; j < add.ApplicationData.ItemsElementName.Length; j++)
                            {
                                if (add.ApplicationData.ItemsElementName[j] == Response.ItemsChoiceType8.Attachments)
                                {
                                    attachments = (Response.Attachments)add.ApplicationData.Items[j];
                                    break;
                                }
                            }

                            break;
                        }
                    }

                    break;
                }
            }

            Site.Assert.IsNotNull(attachments, "The attachments element in expected email item should not be null.");
            #endregion

            #region Call method SendStringRequest to call ItemOperations method to get data of an attachment with two FileReference elements
            // Initialize a Fetch element.
            Request.ItemOperationsFetch fetch = new Request.ItemOperationsFetch
            {
                FileReference = ((Response.AttachmentsAttachment)attachments.Items[0]).FileReference,
                Store = SearchName.Mailbox.ToString(),
                CollectionId = ((Response.AttachmentsAttachment)attachments.Items[0]).FileReference
            };

            // Create an ItemOperations command request and construct two FileReference elements
            ItemOperationsRequest itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { fetch });
            string request = itemOperationsRequest.GetRequestDataSerializedXML();
            request = request.Replace(@"CollectionId xmlns=""AirSync""", @"FileReference xmlns=""AirSyncBase""");
            request = request.Replace(@"/CollectionId", @"/FileReference");
            SendStringResponse sendStringResponse = this.CMDAdapter.SendStringRequest(CommandName.ItemOperations, null, request);

            string status = string.Empty;
            string xmlResponse = sendStringResponse.ResponseDataXML;
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlResponse);
            XmlNamespaceManager xmlNameSpaceManager = new XmlNamespaceManager(doc.NameTable);
            xmlNameSpaceManager.AddNamespace("e", "ItemOperations");
            XmlNode statusNode = doc.SelectSingleNode("/e:ItemOperations/e:Status", xmlNameSpaceManager);

            if (statusNode != null)
            {
                status = statusNode.InnerText;
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R765");

            // Verify MS-ASCMD requirement: MS-ASCMD_R765
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                status,
                765,
                @"[In airsyncbase:FileReference] Violation of this constraint [in an ItemOperations request, only one airsyncbase:FileReference identifier can exist per Fetch node] results in a Status element (section 2.2.3.162.7) value of 2 being returned from the server.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the status 156 of ItemOperations command indicated the destination folder must be of type IPF.Note.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC19_ItemOperations_Status156()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Move element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region User2 sends mail to User1 and does FolderSync in User1's mailbox.
            string subject = this.SendMailAndFolderSync();
            #endregion

            #region Call method Sync to synchronize changes of Inbox folder in User1's mailbox between the client and the server.
            SyncResponse syncResponse = this.GetMailItem(this.User1Information.InboxCollectionId, subject);
            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", subject);
            string conversationId = GetConverstationIdFromSyncResposne(syncResponse, serverId);
            #endregion

            #region Call method ItemOperations which contains the Move element.
            ItemOperationsRequest itemOperationsRequest = new ItemOperationsRequest
            {
                RequestData = new Request.ItemOperations
                {
                    Items = new object[]
                    {
                        new Request.ItemOperationsMove
                        {
                            ConversationId = conversationId,
                            DstFldId = this.User1Information.NotesCollectionId,
                            Options = new Request.ItemOperationsMoveOptions
                            {
                                MoveAlways = string.Empty
                            }
                        }
                    }
                }
            };

            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Site.Assert.IsNotNull(itemOperationsResponse.ResponseData.Response, "The Response element of the ItemOperations response should not be null.");
            Site.Assert.AreEqual<string>("1", itemOperationsResponse.ResponseData.Status, @"The Status element of the ItemOperations response should be 1.");
            Site.Assert.IsNotNull(itemOperationsResponse.ResponseData.Response.Move, "The Move element of the ItemOperations response should not be null.");
            Site.Assert.AreEqual<int>(1, itemOperationsResponse.ResponseData.Response.Move.Length, "The length of Move element array should be 1.");
            Site.Assert.IsNotNull(itemOperationsResponse.ResponseData.Response.Move[0], "The first Move element of the ItemOperations response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4942");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4942
            Site.CaptureRequirementIfAreEqual<string>(
                "156",
                itemOperationsResponse.ResponseData.Response.Move[0].Status,
                4942,
                @"[In Common Status Codes] [The meaning of the status value 156 is] The destination folder for the move is invalid.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the status 16 of ItemOperations command indicated the access to the resource is denied.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S08_TC20_ItemOperations_Status16()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Recipient Information cache is not supported when MS-ASProtocolVersion is 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call ItemOperations command which contains the EmptyFolderContents element without DeleteSubFolders element.
            ItemOperationsRequest itemOperationsRequest = new ItemOperationsRequest
            {
                RequestData = new Request.ItemOperations
                {
                    Items = new object[]
                    {
                        new Request.ItemOperationsEmptyFolderContents
                        {
                            CollectionId = "RI"
                        }
                    }
                }
            };

            ItemOperationsResponse itemOperationsResponse = this.CMDAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);

            Site.Assert.AreEqual<string>("1", itemOperationsResponse.ResponseData.Status, "The Status element of the ItemOperations response should be 1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4167");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4167
            Site.CaptureRequirementIfAreEqual<string>(
                "16",
                itemOperationsResponse.ResponseData.Response.EmptyFolderContents[0].Status,
                4167,
                @"[In Status(ItemOperations)] [The meaning of the status value] 16 [is] Access to the resource is denied.");

            #endregion
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Create ItemOperations request with only Fetch element
        /// </summary>
        /// <param name="collectionId">The CollectionId of a folder</param>
        /// <param name="store">The store on which to do ItemOperations method</param>
        /// <param name="serverId">The ServerId of an Item</param>
        /// <param name="options">The Options element in ItemOperations request</param>
        /// <returns>The request of ItemOperations method</returns>
        private static ItemOperationsRequest CreateItemOperationsRequestWithFetchElement(string collectionId, string store, string serverId, Request.ItemOperationsFetchOptions options)
        {
            ItemOperationsRequest itemOperationsRequest = new ItemOperationsRequest
            {
                RequestData = new Request.ItemOperations
                {
                    Items = new object[]
                    {
                        new Request.ItemOperationsFetch
                        {
                            CollectionId = collectionId,
                            Store = store,
                            ServerId = serverId,
                            Options = options
                        }
                    }
                }
            };

            return itemOperationsRequest;
        }

        /// <summary>
        /// Get the value of ConversationId element from SyncResponse
        /// </summary>
        /// <param name="syncResponse">An instance of the SyncResponse.</param>
        /// <param name="serverId">The value of the ServerId element.</param>
        /// <returns>The value of the ConversationId element</returns>
        private static string GetConverstationIdFromSyncResposne(SyncResponse syncResponse, string serverId)
        {
            int index = 0;
            foreach (Response.ItemsChoiceType10 item in ((Response.SyncCollections)syncResponse.ResponseData.Item).Collection[0].ItemsElementName)
            {
                if (item == Response.ItemsChoiceType10.Commands)
                {
                    Response.SyncCollectionsCollectionCommands commands = ((Response.SyncCollections)syncResponse.ResponseData.Item).Collection[0].Items[index] as Response.SyncCollectionsCollectionCommands;
                    foreach (Response.SyncCollectionsCollectionCommandsAdd add in commands.Add)
                    {
                        if (add.ServerId.Equals(serverId))
                        {
                            for (int j = 0; j < add.ApplicationData.ItemsElementName.Length; j++)
                            {
                                if (add.ApplicationData.ItemsElementName[j] == Response.ItemsChoiceType8.ConversationId)
                                {
                                    return add.ApplicationData.Items[j].ToString();
                                }
                            }
                        }
                    }
                }

                index++;
            }

            return string.Empty;
        }

        /// <summary>
        /// Get the inner text of specified element.
        /// </summary>
        /// <param name="lastRawResponse">The raw xml response.</param>
        /// <param name="parentNodeName">The parent element of the specified node.</param>
        /// <param name="nodeName">The name of the node.</param>
        /// <param name="subject">The subject of the specified item.</param>
        /// <returns>The inner text of the specified element.</returns>
        private static string GetElementInnerText(XmlElement lastRawResponse, string parentNodeName, string nodeName, string subject)
        {
            string data = null;
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(lastRawResponse.OuterXml);
            XmlNodeList subjectElementNodes = doc.SelectNodes("//*[name()='Subject']");
            for (int i = 0; i < subjectElementNodes.Count; i++)
            {
                if (subjectElementNodes[i].InnerText == subject)
                {
                    XmlNodeList bodyElementNodes = doc.SelectNodes("//*[name()='" + parentNodeName + "']");
                    XmlNodeList dataElementNodes = bodyElementNodes[i].SelectNodes("*[name()='" + nodeName + "']");
                    data = dataElementNodes[0].InnerText;
                    break;
                }
            }

            return data;
        }

        /// <summary>
        /// Get the Body element in ItemOperations response.
        /// </summary>
        /// <param name="itemOperationsResponse">The response of ItemOperations method</param>
        /// <returns>The Body element</returns>
        private Response.Body GetBodyElement(ItemOperationsResponse itemOperationsResponse)
        {
            Response.ItemOperationsResponseFetch fetch = this.GetFetchElement(itemOperationsResponse);
            Site.Assert.IsNotNull(fetch.Properties, "The Properties element of the ItemOperations response should not be null.");
            Response.Body body = null;
            foreach (object tempObject in itemOperationsResponse.ResponseData.Response.Fetch[0].Properties.Items)
            {
                if (typeof(Response.Body).Equals(tempObject.GetType()))
                {
                    body = (Response.Body)tempObject;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3363");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3363
            // If server response contains Body element, then MS-ASCMD_R3363 is verified.
            Site.CaptureRequirementIfIsNotNull(
                body,
                3363,
                @"[In MIMESupport(ItemOperations)] The server's response MUST include the airsyncbase:Body element ([MS-ASAIRS] section 2.2.2.4), which is a child element of the Properties element (section 2.2.3.128.1).");

            return body;
        }

        /// <summary>
        /// Get the Data element in ItemOperations response.
        /// </summary>
        /// <param name="itemOperationsResponse">The response of ItemOperations method</param>
        /// <returns>The Data element</returns>
        private string GetDataElement(ItemOperationsResponse itemOperationsResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5886");

            Site.Assert.IsTrue(this.User1Information.UserPassword.Length <= 256, "The expected length of password should be not larger than 256, the actual length is: {0}", this.User1Information.UserPassword.Length);

            // Verify MS-ASCMD requirement: MS-ASCMD_R5886
            // The value of Status element equals 1 represent the server accepts password and returns a successful response.
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                itemOperationsResponse.ResponseData.Status,
                5886,
                @"[In Password(ItemOperations)] The server accepts password values is not larger than 256 characters in length.");

            Response.ItemOperationsResponseFetch fetch = this.GetFetchElement(itemOperationsResponse);
            Site.Assert.IsNotNull(fetch.Properties, "The Properties element of the ItemOperations response should not be null.");
            string data = null;

            for (int i = 0; i < fetch.Properties.ItemsElementName.Length; i++)
            {
                if (fetch.Properties.ItemsElementName[i] == Response.ItemsChoiceType3.Data)
                {
                    data = fetch.Properties.Items[i].ToString();
                    break;
                }
            }

            return data;
        }

        /// <summary>
        /// Get the only Fetch element in ItemOperations response.
        /// </summary>
        /// <param name="itemOperationsResponse">The response of ItemOperations method</param>
        /// <returns>The first Fetch element</returns>
        private Response.ItemOperationsResponseFetch GetFetchElement(ItemOperationsResponse itemOperationsResponse)
        {
            Site.Assert.IsNotNull(itemOperationsResponse.ResponseData.Response, "The Response element of the ItemOperations response should not be null.");
            Site.Assert.AreEqual<string>("1", itemOperationsResponse.ResponseData.Status, @"The Status element of the ItemOperations response should be 1.");
            Site.Assert.IsNotNull(itemOperationsResponse.ResponseData.Response.Fetch, "The Fetch element of the ItemOperations response should not be null.");
            Site.Assert.AreEqual<int>(1, itemOperationsResponse.ResponseData.Response.Fetch.Length, "The length of Fetch element array should be 1.");
            Site.Assert.IsNotNull(itemOperationsResponse.ResponseData.Response.Fetch[0], "The first Fetch element of the ItemOperations response should not be null.");
            return itemOperationsResponse.ResponseData.Response.Fetch[0];
        }
        #endregion
    }
}