namespace Microsoft.Protocols.TestSuites.MS_OXWSBTRF
{
    using System.Collections.ObjectModel;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test export items from a mailbox server and upload items to a mailbox server.
    /// </summary>
    [TestClass]
    public class S01_ExportAndUploadItems : TestSuiteBase
    {
        #region Class initialize and clean up.
        /// <summary>
        /// Initializes the test class. 
        /// </summary>
        /// <param name="testContext">Context to initialize.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clean up the test class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test cases
        /// <summary>
        /// This test case is designed to validate the successful response returned by ExportItems and UploadItems operations 
        /// when the CreateActionType is CreateNew.
        /// </summary>
        [TestCategory("MSOXWSBTRF"), TestMethod()]
        public void MSOXWSBTRF_S01_TC01_ExportAndUploadItems_CreateNew_Success()
        {
            #region Prerequisites.
            // In the initialize step, multiple items in the specified parent folder have been created.
            // If that step executes successfully, the count of CreatedItemId list should be equal to the count of OriginalFolderId list. 
            Site.Assert.AreEqual<int>(
                this.CreatedItemId.Count,
                this.OriginalFolderId.Count,
                string.Format(
                "The exportedItemIds array should contain {0} item ids, actually, it contains {1}", 
                this.OriginalFolderId.Count, 
                this.CreatedItemId.Count));
            #endregion

            #region Call ExportItems operation to export the items from the server.
            // Initialize three ExportItemsType instances.
            ExportItemsType exportItems = new ExportItemsType();

            // Initialize four ItemIdType instances with three different situations to cover the case:
            // 1. The ChangeKey is not present;
            // 2. If the ChangeKey attribute of the ItemIdType complex type is present, its value MUST be either valid or NULL.
            exportItems.ItemIds = new ItemIdType[4]
            {
                new ItemIdType 
                { 
                    // The ChangeKey is not present.
                    Id = this.CreatedItemId[0].Id,
                },
                new ItemIdType 
                { 
                    // The ChangeKey is null.
                    Id = this.CreatedItemId[1].Id,
                    ChangeKey = null
                },
                new ItemIdType 
                {
                    // The ChangeKey is valid.
                    Id = this.CreatedItemId[2].Id,
                    ChangeKey = this.CreatedItemId[2].ChangeKey
                },
                new ItemIdType 
                { 
                    // The ChangeKey is not present.
                    Id = this.CreatedItemId[3].Id,
                }
            };

            // Call ExportItems operation.
            ExportItemsResponseType exportItemsResponse = this.BTRFAdapter.ExportItems(exportItems);
            Site.Assert.IsNotNull(exportItemsResponse, "The ExportItems response should not be null.");

            // Check whether the ExportItems operation is executed successfully.
            foreach (ExportItemsResponseMessageType responseMessage in exportItemsResponse.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Success,
                responseMessage.ResponseClass,
                string.Format(
                    @"The ExportItems operation should be successful. Expected response code: {0}, actual response code: {1}",
                    ResponseClassType.Success,
                    responseMessage.ResponseClass));
            }

            // If the operation executes successfully, the count of items in ExportItems response should be equal to the items in ExportItems request.
            Site.Assert.AreEqual<int>(
                exportItemsResponse.ResponseMessages.Items.Length,
                exportItems.ItemIds.Length,
                string.Format(
                "The exportItems response should contain {0} items, actually, it contains {1}",
                 exportItems.ItemIds.Length,
                 exportItemsResponse.ResponseMessages.Items.Length));

            // Verify ManagementRole part of ExportItems operation
            this.VerifyManagementRolePart();
            #endregion

            #region Verify the ExportItems response related requirements
            // Verify the ExportItemsResponseType related requirements.
            this.VerifyExportItemsSuccessResponse(exportItemsResponse);

            // If the id in ExportItems request is same with the id in ExportItems response, then MS-OXWSBTRF_R169 and MS-OXWSBTRF_R182 can be captured.
            bool isSameItemId = false;
            for (int i = 0; i < exportItemsResponse.ResponseMessages.Items.Length; i++)
            {
                Site.Log.Add(
                   LogEntryKind.Debug,
                   "The exported item's id: '{0}' should be same with the created item's id: {1}.",
                   (exportItemsResponse.ResponseMessages.Items[i] as ExportItemsResponseMessageType).ItemId.Id,
                   this.CreatedItemId[i].Id);
                if ((exportItemsResponse.ResponseMessages.Items[i] as ExportItemsResponseMessageType).ItemId.Id == this.CreatedItemId[i].Id)
                {
                    isSameItemId = true;
                }
                else
                {
                    isSameItemId = false;
                    break;
                }
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R169.");

            // Verify requirement MS-OXWSBTRF_R169
            Site.CaptureRequirementIfIsTrue(
                isSameItemId,
                169,
                @"[In m:ExportItemsResponseMessageType Complex Type][ItemId] specifies the item identifier of a single exported item.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R182.");

            // Verify requirement MS-OXWSBTRF_R182
            Site.CaptureRequirementIfIsTrue(
                isSameItemId,
                182,
                @"[In t:NonEmptyArrayOfItemIdsType Complex Type][ItemId] specifies the item identifier of an item to export from a mailbox.");
            #endregion

            #region Call UploadItems operation and set the CreateAction element to CreateNew to upload the items that exported in last step.
            ExportItemsResponseMessageType[] exportItemsResponseMessages = TestSuiteHelper.GetResponseMessages<ExportItemsResponseMessageType>(exportItemsResponse);

            // Initialize the upload items using the data of previous export items and set the item CreateAction to CreateNew.
            UploadItemsResponseMessageType[] uploadItemsResponse = this.UploadItems(exportItemsResponseMessages, this.OriginalFolderId, CreateActionType.CreateNew, true, true, false);
            #endregion

            #region Verify the UploadItems response related requirements when the CreateAction is CreateNew.
            // If the UploadItems response item's ID is not the same as the previous exported item's ID, then MS-OXWSBTRF_R228 can be captured.
            bool isNotSameItemId = false;
            for (int i = 0; i < uploadItemsResponse.Length; i++)
            {
                // Log the expected and actual value
                Site.Log.Add(
                   LogEntryKind.Debug,
                   "The uploaded item's id: '{0}' should not be same with the exported item's id when the CreateAction is set to CreateNew: {1}.",
                   uploadItemsResponse[i].ItemId.Id,
                   exportItemsResponseMessages[i].ItemId.Id);
                if (exportItemsResponseMessages[i].ItemId.Id != uploadItemsResponse[i].ItemId.Id)
                {
                    isNotSameItemId = true;
                }
                else
                {
                    isNotSameItemId = false;
                    break;
                }
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R228.");

            // Verify Requirement: MS-OXWSBTRF_R228.
            Site.CaptureRequirementIfIsTrue(
                    isNotSameItemId,
                    228,
                    @"[In CreateActionType Simple Type][CreateNew]
                    The <ItemId> element that is returned in the UploadItemsResponseMessageType complex type, as specified in section 3.1.4.2.3.2, 
                    MUST contain the new item identifier.");

            // Call getItem to get the items that was uploaded
            ItemIdType[] itemIds = new ItemIdType[uploadItemsResponse.Length];
            for (int i = 0; i < uploadItemsResponse.Length; i++)
            {
                itemIds[i] = uploadItemsResponse[i].ItemId;
            }

            ItemType[] getItems = this.GetItems(itemIds);

            // Verify the array of items uploaded to a mailbox
            this.VerifyItemsUploadedToMailbox(getItems);

            // If the verification of the items uploaded to mailbox is successful, then this requirement can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R226");

            // Verify requirement MS-OXWSBTRF_R226
            Site.CaptureRequirement(
                226,
                @"[In CreateActionType Simple Type]The Value of CreateNew specifies that a new copy of the original item is uploaded to the mailbox.");

            // If the value of IsAssociated attribute in getItem response is same with the value in uploadItems request, 
            // then requirement MS-OXWSBTRF_R222 can be captured.
            bool isSameIsAssociated = false;
            for (int i = 0; i < getItems.Length; i++)
            {
                // Log the expected and actual value
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "The value of IsAssociated attribute: {0} in getItem response should be: true",
                    getItems[i].IsAssociated);
                if (true == getItems[i].IsAssociated)
                {
                    isSameIsAssociated = true;
                }
                else
                {
                    isSameIsAssociated = false;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R222");

            // Verify requirement MS-OXWSBTRF_R222
            Site.CaptureRequirementIfIsTrue(
                isSameIsAssociated,
                222,
                @"[In m:UploadItemType Complex Type]If it [IsAssociated] is present, it indicates that the item is a folder associated item.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the successful response returned by ExportItems and UploadItems operations 
        /// when the CreateActionType is Update.
        /// </summary>
        [TestCategory("MSOXWSBTRF"), TestMethod()]
        public void MSOXWSBTRF_S01_TC02_ExportAndUploadItems_Update_Success()
        {
            #region Get the exported items.
            // Get the exported items which are prepared for uploading.
            // And set the parameter to true to configure the SOAP header before calling ExportItems operation.
            ExportItemsResponseMessageType[] exportedItems = this.ExportItems(true);
            #endregion

            #region Call UploadItems operation to upload the items that exported to the server in last step.
            // Initialize the upload items using the previous exported data, and set that item CreateAction to Update.
            // Call UploadItems operation with the isAssociated attribute setting to true
            UploadItemsResponseMessageType[] uploadItemsResponse = this.UploadItems(exportedItems, this.OriginalFolderId, CreateActionType.Update, false, false, false);
            #endregion

            #region Verify the UploadItems response related requirements when the CreateAction is Update.
            // If the itemId in uploadItems response is same with it in the uploadItems request, 
            // then MS-OXWSBTRF_R200 and MS-OXWSBTRF_R214 can be captured.
            bool isSameItemId = false;
            for (int i = 0; i < uploadItemsResponse.Length; i++)
            {
                // Log the expected and actual value.
                Site.Log.Add(
                   LogEntryKind.Debug,
                   "The item's id: '{0}' in UploadItems request should be same with the item's id: {1} in UploadItems response when the CreateAction is set to Update.",
                   exportedItems[i].ItemId.Id,
                   uploadItemsResponse[i].ItemId.Id);
                if (exportedItems[i].ItemId.Id == uploadItemsResponse[i].ItemId.Id)
                {
                    isSameItemId = true;
                }
                else
                {
                    isSameItemId = false;
                    break;
                }
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R200.");
            
            // Verify requirement MS-OXWSBTRF_R200
            Site.CaptureRequirementIfIsTrue(
                isSameItemId,
                200,
                @"[In m:UploadItemsResponseMessageType Complex Type]
                The ItemId element specifies the item identifier of an item that was uploaded into a mailbox.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R214.");

            // Verify requirement MS-OXWSBTRF_R200
            Site.CaptureRequirementIfIsTrue(
                isSameItemId,
                214,
                @"[In m:UploadItemType Complex Type][ItemId] specifies the item identifier of the upload item.");

            // Call getItem to get the items that was uploaded.
            ItemIdType[] itemIds = new ItemIdType[uploadItemsResponse.Length];
            for (int i = 0; i < uploadItemsResponse.Length; i++)
            {
                itemIds[i] = uploadItemsResponse[i].ItemId;
            }

            ItemType[] getItems = this.GetItems(itemIds);

            // Verify the array of items uploaded to a mailbox
            this.VerifyItemsUploadedToMailbox(getItems);
            
            // If the ParentFolderId in GetItem response is same with it in UploadItem request, then requirement MS-OXWSBTRF_R211 can be captured.
            bool isSameParentFolderId = false;
            for (int i = 0; i < getItems.Length; i++)
            {
                // Log the expected and actual value.
                Site.Log.Add(
                   LogEntryKind.Debug,
                   "The ParentFolderId: {0} in GetItem response should be same with the ParentFolderId: {1} in UploadItems request the items are updated successfully.",
                   getItems[i].ParentFolderId.Id,
                   this.OriginalFolderId[i]);
                if (this.OriginalFolderId[i] == getItems[i].ParentFolderId.Id)
                {
                    isSameParentFolderId = true;
                }
                else
                {
                    isSameParentFolderId = false;
                    break;
                }
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R211");

            // Verify requirement MS-OXWSBTRF_R211
            Site.CaptureRequirementIfIsTrue(
                isSameParentFolderId,
                211,
                @"[In m:UploadItemType Complex Type][ParentFolderId] specifies the target folder in which to place the upload item.");
            
            // Call exportItems again after update to verify whether the items are updated or not.
            ExportItemsResponseMessageType[] exportedItemsAfterUpload = this.ExportItems(false);

            // If the value of Data element is different, then requirement MS-OXWSBTRF_R229 can be captured.
            bool isNotSameData = false;
            for (int i = 0; i < exportedItemsAfterUpload.Length; i++)
            {
                // Log the expected and actual value.
                Site.Log.Add(
                   LogEntryKind.Debug,
                   "The Data: {0} before UploadItems should not be same with the Data: {1} that after UploadItems when the CreateAction is set to Update.",
                   exportedItems[i].Data,
                   exportedItemsAfterUpload[i].Data);
                if (exportedItems[i].Data != exportedItemsAfterUpload[i].Data)
                {
                    isNotSameData = true;
                }
                else
                {
                    isNotSameData = false;
                    break;
                }
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R229");

            // Verify requirement MS-OXWSBTRF_R229
            Site.CaptureRequirementIfIsTrue(
                isNotSameData,
                229,
                @"[In CreateActionType Simple Type]The Value of Update specifies that the upload will update an item that is already present in the mailbox.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the successful response returned by ExportItems and UploadItems operations when the CreateActionType is UpdateOrCreate.
        /// </summary>
        [TestCategory("MSOXWSBTRF"), TestMethod()]
        public void MSOXWSBTRF_S01_TC03_ExportAndUploadItems_UpdateOrCreate_Success()
        {
            #region Get the exported items.
            // Get the exported items which is prepared for uploading.
            ExportItemsResponseMessageType[] exportedItems = this.ExportItems(false);
            #endregion

            #region Call UploadItems operation and set the CreateAction element to UpdateorCreate and the parent folder to the original one to upload the items that exported in last step.
            // Initialize the uploaded items using the information of previous exported items, set that item's CreateAction to UpdateOrCreate and the ParentFolderId to the same as the parent folder.
            // Set the last parameter to true to configure the SOAP header before calling UploadItems operation
            UploadItemsResponseMessageType[] uploadItemsResponse = this.UploadItems(exportedItems, this.OriginalFolderId, CreateActionType.UpdateOrCreate, true, false, true);
            
            // Call getItem to get the items that was uploaded.
            ItemIdType[] itemIds = new ItemIdType[uploadItemsResponse.Length];
            for (int i = 0; i < uploadItemsResponse.Length; i++)
            {
                itemIds[i] = uploadItemsResponse[i].ItemId;
            }

            ItemType[] getItems = this.GetItems(itemIds);

            // Verify the array of items uploaded to a mailbox
            this.VerifyItemsUploadedToMailbox(getItems);
            #endregion

            #region Call ExportItems again to verify the value of Data after update.
            ExportItemsResponseMessageType[] exportedItemsAfterUpload = this.ExportItems(false);

            // If the value of Data element is different, then requirement MS-OXWSBTRF_R2321 can be captured.
            bool isNotSameData = false;
            for (int i = 0; i < exportedItemsAfterUpload.Length; i++)
            {
                // Log the expected and actual value.
                Site.Log.Add(
                   LogEntryKind.Debug,
                   "The Data: {0} before upload should not be same with the Data: {1} that after upload when the criteria meets to Update.",
                   exportedItems[i].Data,
                   exportedItemsAfterUpload[i].Data);
                if (exportedItems[i].Data != exportedItemsAfterUpload[i].Data)
                {
                    isNotSameData = true;
                }
                else
                {
                    isNotSameData = false;
                    break;
                }
            }

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R2321");

            // Verify requirement MS-OXWSBTRF_R3221
            Site.CaptureRequirementIfIsTrue(
                isNotSameData,
                2321,
                @"[In CreateActionType Simple Type][UpdateOrCreate]If the criteria for a successful update are met, the target item is updated.");
            #endregion

            #region Create another sub folder in the specified parent folder.
            // Create another sub folder in the specified parent folder.
            Collection<string> subFolderIds = new Collection<string>();
            for (int i = 0; i < this.ParentFolderType.Count; i++)
            {
                // Generate the folder name.
                string folderName = Common.GenerateResourceName(this.Site, this.ParentFolderType[i] + "NewFolder");

                // Create another sub folder in the specified parent folder
                subFolderIds.Add(this.CreateSubFolder(this.ParentFolderType[i], folderName));
                Site.Assert.IsNotNull(
                    subFolderIds[i], 
                    string.Format(
                    "The sub folder named '{0}' under '{1}' should be created successfully.", 
                    folderName, 
                    this.ParentFolderType[i].ToString()));
            }
            #endregion

            #region Call UploadItems operation and set the CreateAction element to UpdateorCreate and the parent folder to the new created one to upload the items that exported in last step to a new folder.
            // Initialize the uploaded items using the previous exported items, set the item's CreateAction to UpdateOrCreate and the ParentFolderId to the sub folder ID.
            UploadItemsResponseMessageType[] uploadItemsWithChangedParentFolderIdResponseMessages = this.UploadItems(exportedItemsAfterUpload, subFolderIds, CreateActionType.UpdateOrCreate, true, false, false);
            #endregion

            #region Verify the UploadItemsResponseType related requirements
            // Call getItem to get the items that was uploaded.
            ItemIdType[] newItemIds = new ItemIdType[uploadItemsWithChangedParentFolderIdResponseMessages.Length];
            for (int i = 0; i < uploadItemsWithChangedParentFolderIdResponseMessages.Length; i++)
            {
                newItemIds[i] = uploadItemsWithChangedParentFolderIdResponseMessages[i].ItemId;
            }

            ItemType[] getUpdatedItems = this.GetItems(newItemIds);

            // Verify the array of items uploaded to a mailbox
            this.VerifyItemsUploadedToMailbox(getUpdatedItems);

            // Check the parent folder id is not the original one.
            bool isNotSameParentFolderId = false;
            for (int i = 0; i < getUpdatedItems.Length; i++)
            {
                // Log the expected and actual value.
                Site.Log.Add(
                   LogEntryKind.Debug,
                   "The ParentFolderId: {0} in GetItems response should not be same with the original ParentFolderId: {1} when uploading items to sub folders",
                   getUpdatedItems[i].ParentFolderId.Id,
                   this.OriginalFolderId[i]);
                if (this.OriginalFolderId[i] != getUpdatedItems[i].ParentFolderId.Id)
                {
                    isNotSameParentFolderId = true;
                }
                else
                {
                    isNotSameParentFolderId = false;
                    break;
                }
            }

            // If requirement MS-OXWSBTRF_R195 and MS-OXWSBTRF_R207 are captured and the ParentFolderId is not the original one, then requirement MS-OXWSBTRF_R235 can be captured.
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R235");

            // Verify requirement MS-OXWSBTRF_R235
            Site.CaptureRequirementIfIsTrue(
                isNotSameParentFolderId,
                235,
                @"[In CreateActionType Simple Type][UpdateOrCreate]If the target item is not in the original folder specified by the <ParentFolderId> element in the 
                UploadItemType complex type, a new copy of the original item is uploaded to the mailbox associated with the folder specified by the <ParentFolderId> element.");

            // Check the item id is a new one.
            bool isNotSameItemId = false;
            for (int i = 0; i < uploadItemsWithChangedParentFolderIdResponseMessages.Length; i++)
            {
                // Log the expected and actual value.
                Site.Log.Add(
                   LogEntryKind.Debug,
                   "The Item's id: {0} in UploadItems response should not be same with the Item's id: {1} in ExportItems response when the criteria meets CreateNew",
                   uploadItemsWithChangedParentFolderIdResponseMessages[i].ItemId.Id,
                   exportedItems[i].ItemId.Id);
                if (exportedItems[i].ItemId.Id != uploadItemsWithChangedParentFolderIdResponseMessages[i].ItemId.Id)
                {
                    isNotSameItemId = true;
                }
                else
                {
                    isNotSameItemId = false;
                    break;
                }
            }

            // If the items' ID in UploadItems response is not the same as the previous exported item's ID and the ParentFolderId in GetItem response is not same with it in UploadItem request.
            // Then MS-OXWSBTRF_R2351 can be captured.
            bool isNotSameItemIdAndParentFolderId = isNotSameParentFolderId && isNotSameItemId;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R2351.");

            // Verify Requirement: MS-OXWSBTRF_R2351.
            Site.CaptureRequirementIfIsTrue(
                    isNotSameItemIdAndParentFolderId,
                    2351,
                    @"[In CreateActionType Simple Type][UpdateOrCreate][If the target item is not in the original folder specified by the 
                    <ParentFolderId> element in the UploadItemType complex type]The <ItemId> element that is returned in the 
                    UploadItemsResponseMessageType complex type, as specified in section 3.1.4.2.3.2, MUST contain the new item identifier.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the unsuccessful response returned by ExportItems operation when the request is unsuccessful.
        /// There are three kinds of situation:
        /// 1 The ID attribute of ItemId is empty.
        /// 2 The ChangeKey attribute of ItemId is invalid.
        /// 3 The ID attribute of ItemId is invalid.
        /// </summary>
        [TestCategory("MSOXWSBTRF"), TestMethod()]
        public void MSOXWSBTRF_S01_TC04_ExportItems_Fail()
        {
            #region Get the exported items.
            // In the initialize step, multiple items in the specified parent folder have been created.
            // If that step executes successfully, the count of CreatedItemId list should be equal to the count of OriginalFolderId list. 
            Site.Assert.AreEqual<int>(
                this.CreatedItemId.Count,
                this.OriginalFolderId.Count,
                string.Format(
                "The exportedItemIds array should contain {0} item ids, actually, it contains {1}.", 
                this.OriginalFolderId.Count,
                this.CreatedItemId.Count));
            #endregion

            #region Call ExportItems operation to export the items from the server.
            // Initialize three ExportItemsType instances.
            ExportItemsType exportItems = new ExportItemsType();

            // Initialize ItemIdType instances with three different situations:
            // 1.The ID attribute of ItemId is empty;
            // 2.The ID attribute of ItemId is valid and the ChangeKey is invalid;
            // 3.The ID attribute of ItemId is invalid and the ChangeKey is null.
            exportItems.ItemIds = new ItemIdType[3] 
            {
                new ItemIdType 
                { 
                    Id = string.Empty,
                    ChangeKey = null
                },
                new ItemIdType 
                { 
                    Id = this.CreatedItemId[1].Id,
                    ChangeKey = InvalidChangeKey
                },
                new ItemIdType 
                {
                    Id = InvalidItemId,
                    ChangeKey = null
                }
            };

            // Call ExportItems operation.
            ExportItemsResponseType exportItemsResponse = this.BTRFAdapter.ExportItems(exportItems);

            // Check whether the ExportItems operation is executed successfully.
            foreach (ExportItemsResponseMessageType responseMessage in exportItemsResponse.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                responseMessage.ResponseClass,
                string.Format(
                    @"The ExportItems operation should be unsuccessful. Expected response code: {0}, actual response code: {1}",
                    ResponseClassType.Error,
                    responseMessage.ResponseClass));
            }
            #endregion

            #region Verify ExportItems fail related requirements.
            this.VerifyExportItemsErrorResponse(exportItemsResponse);
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the unsuccessful response returned by UploadItems operation when the request is unsuccessful and the CreateActionType is Update.
        /// There are four kinds of this situation:
        /// 1 The target item is not in the original folder specified by the ParentFolderId element in the UploadItemType.
        /// 2 The ID attribute of ItemId is empty.
        /// 3 The ChangeKey attribute of ItemId is invalid.
        /// 4 The ID attribute of ItemId is invalid.
        /// </summary>
        [TestCategory("MSOXWSBTRF"), TestMethod()]
        public void MSOXWSBTRF_S01_TC05_ExportAndUploadItems_Update_Fail()
        {
            #region Get the exported items.
            // Get the exported items which are prepared for uploading.
            ExportItemsResponseMessageType[] exportedItem = this.ExportItems(false);
            #endregion

            #region Create a new folder to place the upload item.
            // Create another sub folder.
            string[] subFolderIds = new string[this.ParentFolderType.Count];
            for (int i = 0; i < subFolderIds.Length; i++)
            {
                // Generate the folder name.
                string folderName = Common.GenerateResourceName(this.Site, this.ParentFolderType[i] + "NewFolder");

                // Create sub folder in the specified parent folder
                subFolderIds[i] = this.CreateSubFolder(this.ParentFolderType[i], folderName);
                Site.Assert.IsNotNull(
                    subFolderIds[i], 
                    string.Format(
                    "The sub folder named '{0}' under '{1}' should be created successfully.", 
                    folderName, 
                    this.ParentFolderType[i].ToString()));
            }
            #endregion

            #region Call UploadItems operation with CreateAction set to Update and the ParentFolderId set to the new created sub folder.
            // Initialize the uploaded items using the previous exported items, set the item's CreateAction to Update and the ParentFolderId to the sub folder.
            UploadItemsType uploadItemsWithChangedParentFolderId = new UploadItemsType();
            uploadItemsWithChangedParentFolderId.Items = new UploadItemType[this.ItemCount];
            for (int i = 0; i < uploadItemsWithChangedParentFolderId.Items.Length; i++)
            {
                uploadItemsWithChangedParentFolderId.Items[i] = TestSuiteHelper.GenerateUploadItem(
                    exportedItem[i].ItemId.Id,
                    exportedItem[i].ItemId.ChangeKey,
                    exportedItem[i].Data,
                    subFolderIds[i],
                    CreateActionType.Update);
            }

            // Call UploadItems operation.
            UploadItemsResponseType uploadItemsWithChangedParentFolderIdResponse = this.BTRFAdapter.UploadItems(uploadItemsWithChangedParentFolderId);

            // Check whether the ExportItems operation is executed successfully.
            foreach (UploadItemsResponseMessageType responseMessage in uploadItemsWithChangedParentFolderIdResponse.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                responseMessage.ResponseClass,
                string.Format(
                    @"The ExportItems operation should be unsuccessful. Expected response code: {0}, actual response code: {1}",
                    ResponseClassType.Error, 
                    responseMessage.ResponseClass));
            }
            #endregion

            #region Verify UploadItems related requirements.
            // Verify UploadItems related requirements.
            this.VerifyUploadItemsErrorResponse(uploadItemsWithChangedParentFolderIdResponse);

            // If the ResponseClass property equals Error, then MS-OXWSBTRF_R231 can be captured 
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R231.");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R231
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorItemNotFound,
                uploadItemsWithChangedParentFolderIdResponse.ResponseMessages.Items[0].ResponseCode,
                231,
                @"[In CreateActionType Simple Type][Update] 
                if the target item is not in the original folder specified by the <ParentFolderId> element in the UploadItemType complex type, 
                an ErrorItemNotFound error code MUST be returned in the UploadItemsResponseMessageType complex type.");
            #endregion

            #region Call UploadItems operation with CreateAction set to Update and the ItemId set to invalid value.
            // Call UploadItemsFail method to upload three kinds of items and set the CreateAction to Update:
            // 1.The ID attribute of ItemId is empty;
            // 2.The ID attribute of ItemId is valid and the ChangeKey is invalid;
            // 3.The ID attribute of ItemId is invalid and the ChangeKey is null.
            this.UploadInvalidItems(this.OriginalFolderId, CreateActionType.Update);
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate unsuccessful response returned by UploadItems operation when the request is unsuccessful and the CreateActionType is CreateOrUpdate.
        /// There are three kinds of this situation:
        /// 1 The id attribute of ItemId is empty. 
        /// 2 The ChangeKey attribute of ItemId is invalid.
        /// 3 The id attribute of ItemId is invalid.
        /// </summary>
        [TestCategory("MSOXWSBTRF"), TestMethod()]
        public void MSOXWSBTRF_S01_TC06_ExportAndUploadItems_UpdateOrCreate_Fail()
        {
            #region Call UploadItems operation to upload the items.
            // Call UploadItemsFail method to upload three kinds of items and set the CreateAction to UpdateOrCreate:
            // 1.The ID attribute of ItemId is empty;
            // 2.The ID attribute of ItemId is valid and the ChangeKey is invalid;
            // 3.The ID attribute of ItemId is invalid and the ChangeKey is null.
            this.UploadInvalidItems(this.OriginalFolderId, CreateActionType.UpdateOrCreate);
            #endregion
        }
        #endregion
    }
}