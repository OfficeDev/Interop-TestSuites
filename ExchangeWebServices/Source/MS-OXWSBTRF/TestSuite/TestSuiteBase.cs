namespace Microsoft.Protocols.TestSuites.MS_OXWSBTRF
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The bass class of scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Fields
        /// <summary>
        /// Gets the invalid change key.
        /// </summary>
        protected const string InvalidChangeKey = "InvalidChangeKey";

        /// <summary>
        /// Gets the invalid item ID.
        /// </summary>
        protected const string InvalidItemId = "InvalidItemId";

        /// <summary>
        /// The folders created by case.
        /// </summary>
        private List<BaseFolderType> createdfolders = new List<BaseFolderType>();
        #endregion

        #region Properties
        /// <summary>
        /// Gets the default folder name.
        /// </summary>
        protected Collection<DistinguishedFolderIdNameType> ParentFolderType { get; private set; }

        /// <summary>
        /// Gets the original folder ID.
        /// </summary>
        protected Collection<string> OriginalFolderId { get; private set; }

        /// <summary>
        /// Gets the Subject of an item.
        /// </summary>
        protected Collection<string> CreatedItemSubject { get; private set; }

        /// <summary>
        /// Gets the IDs of the created items.
        /// </summary>
        protected Collection<ItemIdType> CreatedItemId { get; private set; }

        /// <summary>
        /// Gets the items count that prepared to export or upload.
        /// </summary>
        protected int ItemCount { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSBTRF protocol adapter instance.
        /// </summary>
        protected IMS_OXWSBTRFAdapter BTRFAdapter { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSFOLD protocol adapter instance.
        /// </summary>
        protected IMS_OXWSFOLDAdapter FOLDAdapter { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSCORE protocol adapter instance.
        /// </summary>
        protected IMS_OXWSCOREAdapter COREAdapter { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSBTRF SUT control adapter instance.
        /// </summary>
        protected IMS_OXWSBTRFSUTControlAdapter BTRFSUTControlAdapter { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the schema validation is successful.
        /// </summary>
        protected bool IsSchemaValidated { get; private set; }

        /// <summary>
        /// Gets the value of field createdfolders.
        /// </summary>
        private List<BaseFolderType> CreatedFolders
        {
            get { return this.createdfolders; }
        }
        #endregion

        #region Test case initialize and clean up
        /// <summary>
        /// Initialize the Test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.BTRFAdapter = Site.GetAdapter<IMS_OXWSBTRFAdapter>();

            // If implementation doesn't support this specification [MS-OXWSBTRF] as specified in section 8, the case will not start.
            if (!bool.Parse(Common.GetConfigurationPropertyValue("MS-OXWSBTRF_Supported", this.Site)))
            {
                SutVersion currentSutVersion = (SutVersion)Enum.Parse(typeof(SutVersion), Common.GetConfigurationPropertyValue("SutVersion", this.Site));
                this.Site.Assert.Inconclusive("This test suite is not supported under current SUT, because MS-OXWSBTRF_Supported value is set to false in MS-OXWSBTRF_{0}_SHOULDMAY.deployment.ptfconfig file.", currentSutVersion);
            }
            else
            {
                this.FOLDAdapter = Site.GetAdapter<IMS_OXWSFOLDAdapter>();
                this.COREAdapter = Site.GetAdapter<IMS_OXWSCOREAdapter>();
                this.BTRFSUTControlAdapter = Site.GetAdapter<IMS_OXWSBTRFSUTControlAdapter>();

                // Add four folder types to ParentFolderType list.
                this.ParentFolderType = new Collection<DistinguishedFolderIdNameType>() 
                {
                    DistinguishedFolderIdNameType.inbox, 
                    DistinguishedFolderIdNameType.calendar, 
                    DistinguishedFolderIdNameType.contacts, 
                    DistinguishedFolderIdNameType.tasks 
                };

                // Initialize the OriginalFolderId collection to store the folder ids that items will be exported from and uploaded to
                this.OriginalFolderId = new Collection<string>();
                for (int i = 0; i < this.ParentFolderType.Count; i++)
                {
                    // Generate the folder name.
                    string folderName = Common.GenerateResourceName(this.Site, this.ParentFolderType[i] + "OriginalFolder");

                    // Create a sub folder in the specified parent folder.
                    string folderId = this.CreateSubFolder(this.ParentFolderType[i], folderName);

                    // Add the new created sub folder's id to OriginalFolderId collection.
                    this.OriginalFolderId.Add(folderId);
                    Site.Assert.IsNotNull(
                        this.OriginalFolderId[i],
                        string.Format(
                        "The sub folder named '{0}' in folder '{1}' should be created successfully.",
                        folderName,
                        this.ParentFolderType[i].ToString()));
                }

                // Initialize the CreatedItemSubject list to store all created item subjects.
                this.CreatedItemSubject = new Collection<string>();

                // Initialize CreatedItemId list to store all created item ids.
                this.CreatedItemId = new Collection<ItemIdType>();

                // Create an ItemIdType array to store the created items' ids.
                ItemIdType[] itemIds = new ItemIdType[this.OriginalFolderId.Count];

                for (int i = 0; i < itemIds.Length; i++)
                {
                    // Generate the item subject.
                    string itemSubject = Common.GenerateResourceName(this.Site, this.ParentFolderType[i] + "Item");

                    // Create items in the created sub folders. 
                    itemIds[i] = this.CreateItem(this.ParentFolderType[i], this.OriginalFolderId[i], itemSubject);
                    Site.Assert.IsNotNull(itemIds[i], string.Format("The item with subject '{0}' should be created successfully!", itemSubject));

                    // If the item id is not empty, add it to the CreatedItemId list.
                    this.CreatedItemId.Add(itemIds[i]);

                    // Add the Subject to the CreatedItemSubject list.
                    this.CreatedItemSubject.Add(itemSubject);
                }

                this.ItemCount = this.CreatedItemId.Count;
                ExchangeServiceBinding.ServiceResponseEvent += new ExchangeServiceBinding.ServiceResponseDelegate(this.ExchangeServiceBinding_ResponseEvent);
            }
        }

        /// <summary>
        /// Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            // Clean up environment.
            for (int i = 0; i < this.CreatedFolders.Count; i++)
            {
                bool isCleaned = this.BTRFSUTControlAdapter.CleanupFolder(
                Common.GetConfigurationPropertyValue("UserName", this.Site),
                Common.GetConfigurationPropertyValue("UserPassword", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                    this.CreatedFolders[i].ParentFolderId.Id.ToString(),
                     this.CreatedFolders[i].DisplayName);
                Site.Assert.IsTrue(
                    isCleaned,
                    string.Format("All the items and sub folders in the folder '{0}' should be cleaned up.", this.CreatedFolders[i].ParentFolderId.Id.ToString()));

                ExchangeServiceBinding.ServiceResponseEvent -= new ExchangeServiceBinding.ServiceResponseDelegate(this.ExchangeServiceBinding_ResponseEvent);
            }

            base.TestCleanup();
        }
        #endregion

        #region Capture methods
        /// <summary>
        /// Verify the ExportItemsResponseType related requirements when the ExportItems operation executes successfully.
        /// </summary>
        /// <param name="exportItemsResponse">The ExportItemsResponseType instance returned from the server.</param>
        protected void VerifyExportItemsSuccessResponse(BaseResponseMessageType exportItemsResponse)
        {
            foreach (ResponseMessageType responseMessage in exportItemsResponse.ResponseMessages.Items)
            {
                ExportItemsResponseMessageType exportResponse = responseMessage as ExportItemsResponseMessageType;
                Site.Assert.IsNotNull(exportResponse, "The response got from server should not be null.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R2107.");

                // Verify requirement: MS-OXWSBTRF_R2107
                Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Success,
                    responseMessage.ResponseClass,
                    2107,
                    @"[In tns:ExportItemsSoapOut Message]If the request is successful, the ExportItems operation returns an ExportItemsResponse element 
                    with the ResponseClass attribute of the ExportItemsResponseMessage element set to ""Success"".");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R170.");

                // If the ItemId in exportResponse is not null, then requirement MS-OXWSBTRF_R170 can be verified.
                // Verify requirement: MS-OXWSBTRF_R170
                Site.CaptureRequirementIfIsNotNull(
                    exportResponse.ItemId,
                    170,
                    @"[In m:ExportItemsResponseMessageType Complex Type]This element[ItemId] MUST be present if the export operation is successful.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R172.");

                // If the data in exportResponse is not null, then requirement MS-OXWSBTRF_R172 and MS-OXWSBTRF_R173 can be verified.
                // Verify requirement: MS-OXWSBTRF_R172
                Site.CaptureRequirementIfIsNotNull(
                        exportResponse.Data,
                        172,
                        @"[In m:ExportItemsResponseMessageType Complex Type][Data] specifies the data of a single exported item.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R173.");

                // Verify requirement: MS-OXWSBTRF_R173
                Site.CaptureRequirementIfIsNotNull(
                        exportResponse.Data,
                        173,
                        @"[In m:ExportItemsResponseMessageType Complex Type]This element[Data] MUST be present if the export operation is successful.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R2108.");

                // Verify requirement: MS-OXWSBTRF_R2108
                Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                    ResponseCodeType.NoError,
                    responseMessage.ResponseCode,
                    2108,
                    @"[In tns:ExportItemsSoapOut Message][If the request is successful]The ResponseCode element of the ExportItemsResponseMessage element is set to ""NoError"".");
            }

            // If the length of items in ExportItems response is equal to this.ItemCount, and the exported items' id are same as created items' id then requirement MS-OXWSBTRF_R179 can be captured.
            Site.Assert.AreEqual<int>(exportItemsResponse.ResponseMessages.Items.Length, this.ItemCount, "The exported items' count should be the same with created items' count");
            bool isSameId = false;
            for (int i = 0; i < this.ItemCount; i++)
            {
                Site.Log.Add(
                   LogEntryKind.Debug,
                   "The exported items' id: '{0}' should be same with the created items' id: {1}.",
                   (exportItemsResponse.ResponseMessages.Items[i] as ExportItemsResponseMessageType).ItemId.Id,
                   this.CreatedItemId[i].Id);
                if ((exportItemsResponse.ResponseMessages.Items[i] as ExportItemsResponseMessageType).ItemId.Id == this.CreatedItemId[i].Id)
                {
                    isSameId = true;
                }
                else
                {
                    isSameId = false;
                    break;
                }
            }

            Site.Assert.IsTrue(isSameId, "The exported items' id should be same to the created items' id. ");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R179.");

            // After verify the length of items in ExportItems response is equal to this.ItemCount, and the exported items' id are same as created items' id then requirement MS-OXWSBTRF_R179 can be captured.
            Site.CaptureRequirement(
                179,
                @"[In ExportItemsType Complex Type][ItemIds] specifies the item identifier array of the items to export.");
        }

        /// <summary>
        /// Verify the ExportItemsResponseType related requirements when ExportItems operation executes unsuccessfully.
        /// </summary>
        /// <param name="exportItemsResponse">The ExportItemsResponseType instance returned from the server.</param>
        protected void VerifyExportItemsErrorResponse(BaseResponseMessageType exportItemsResponse)
        {
            foreach (ResponseMessageType responseMessage in exportItemsResponse.ResponseMessages.Items)
            {
                // If the ExportItems operation is unsuccessful, the ResponseClass should be set to "Error", then requirement MS-OXWSBTRF_R2109 can be captured.
                // Add debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R2109");

                // Verify requirement: MS-OXWSBTRF_R2109.
                Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Error,
                    responseMessage.ResponseClass,
                    2109,
                    @"[In tns:ExportItemsSoapOut Message]If the request is unsuccessful, the ExportItems operation returns an ExportItemsResponse element with the ResponseClass 
                    attribute of the ExportItemsResponseMessage element set to ""Error"".");

                // If the ResponseCode is correspond to the schema, then the ResponseCode should be a value of the ResponseCodeType simple type.
                // Add debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R2110");

                // requirement MS-OXWSBTRF_R2110 can be captured.
                Site.CaptureRequirementIfIsTrue(
                    this.IsSchemaValidated,
                    2110,
                    @"[In tns:ExportItemsSoapOut Message][If the request is unsuccessful]The ResponseCode element of the ExportItemsResponseMessage element is set to a value of the ResponseCodeType simple type, as specified in [MS-OXWSCDATA] section 2.2.5.24.");
            }
        }

        /// <summary>
        /// Verify the UploadItemsResponseType related requirements when UploadItems operation executes successfully.
        /// </summary>
        /// <param name="uploadItemsResponse">The UploadItemsResponseType instance returned from the server.</param>
        protected void VerifyUploadItemsSuccessResponse(BaseResponseMessageType uploadItemsResponse)
        {
            foreach (UploadItemsResponseMessageType responseMessage in uploadItemsResponse.ResponseMessages.Items)
            {
                Site.Assert.IsNotNull(responseMessage, "The response got from server should not be null.");

                // If the length of items in UploadItems response is equal to this.ItemCount, and uploaded Items' subject are the same as created items' subject then requirement MS-OXWSBTRF_R207 can be captured.
                Site.Assert.AreEqual<int>(uploadItemsResponse.ResponseMessages.Items.Length, this.ItemCount, "The count of uploaded items should be the same as the count of created items");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R2007.");

                // Verify requirement: MS-OXWSBTRF_R2007
                Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Success,
                    responseMessage.ResponseClass,
                    2007,
                    @"[In tns:UploadItemsSoapOut Message]If the request [UploadItems request] is successful, the UploadItems operation returns an 
                    UploadItemsResponse element with the ResponseClass attribute of the UploadItemsResponseMessage element set to ""Success"".");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R2008.");

                // Verify requirement: MS-OXWSBTRF_R2008
                Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                    ResponseCodeType.NoError,
                    responseMessage.ResponseCode,
                    2008,
                    @"[In tns:UploadItemsSoapOut Message][If the UploadItems request is successful]
                    The ResponseCode element of the UploadItemsResponseMessage element is set to ""NoError"".");
            }
        }

        /// <summary>
        /// Verify the array of items uploaded to a mailbox
        /// </summary>
        /// <param name="getItems">The items information.</param>
        protected void VerifyItemsUploadedToMailbox(ItemType[] getItems)
        {
            bool isSameItemSubject = false;
            for (int i = 0; i < getItems.Length; i++)
            {
                // Log the expected subject and actual subject
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "The new uploaded item with subject '{0}' should have the same subject '{1}' with the original one.",
                    getItems[i].Subject,
                    this.CreatedItemSubject[i]);
                if (this.CreatedItemSubject[i] == getItems[i].Subject)
                {
                    isSameItemSubject = true;
                }
                else
                {
                    isSameItemSubject = false;
                    break;
                }
            }

            Site.Assert.IsTrue(isSameItemSubject, "The uploaded items' subject should be same to the created items' subject.");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R195");

            // Verify requirement MS-OXWSBTRF_R195
            // After verify the length of items in UploadItems response is equal to this.ItemCount, and uploaded Items' subject are the same as created items' subject then requirement MS-OXWSBTRF_R195 can be captured.
            Site.CaptureRequirement(
                195,
                @"[In t:NonEmptyArrayOfUploadItemsType Complex Type][Item] specifies the array of items to upload in to a mailbox.");

            // Add debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R207");

            // Verify requirement MS-OXWSBTRF_R207
            // After verify the length of items in UploadItems response is equal to this.ItemCount, and uploaded Items' subject are the same as created items' subject then requirement MS-OXWSBTRF_R207 can be captured.
            Site.CaptureRequirement(
                207,
                @"[In m:UploadItemsType Complex Type][The element Items] specifies the collection of items to upload into a mailbox.");
        }

        /// <summary>
        /// Verify the UploadItemsResponseType related requirements when UploadItems executes unsuccessfully.
        /// </summary>
        /// <param name="uploadItemsResponse">The UploadItemsResponseType instance returned from the server.</param>
        protected void VerifyUploadItemsErrorResponse(BaseResponseMessageType uploadItemsResponse)
        {
            foreach (UploadItemsResponseMessageType responseMessage in uploadItemsResponse.ResponseMessages.Items)
            {
                // If the UploadItems operation is unsuccessful, the ResponseClass should be set to "Error", then requirement MS-OXWSBTRF_R2009 can be captured.
                // Add debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R2009");

                // Verify requirement: MS-OXWSBTRF_R2009.
                Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Error,
                    responseMessage.ResponseClass,
                    2009,
                    @"[In tns:UploadItemsSoapOut Message]
                    If the request [UploadItems request] is unsuccessful, the UploadItems operation returns an UploadItemsResponse element 
                    with the ResponseClass attribute of the UploadItemsResponseMessage element set to ""Error"".");

                // If the ResponseCode is correspond to the schema, the ResponseCode should be a value of the ResponseCodeType simple type, then requirement MS-OXWSBTRF_R2010 can be captured.
                // Add debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R2010");

                // Verify requirement: MS-OXWSBTRF_R2010
                Site.CaptureRequirementIfIsTrue(
                    this.IsSchemaValidated,
                    2010,
                    @"[In tns:UploadItemsSoapOut Message][If the request UploadItems request is unsuccessful]The ResponseCode element of the UploadItemsResponseMessage element is set to a value of the ResponseCodeType simple type, as specified in [MS-OXWSCDATA] section 2.2.5.24.");
            }
        }

        /// <summary>
        /// Verify the ManagementRole part of ExportItems operation
        /// </summary>
        protected void VerifyManagementRolePart()
        {
            if (Common.IsRequirementEnabled(2111, this.Site))
            {
                // Add debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R2111");

                // If the proxy can communicate with server successfully, then the WSDL related requirements can be captured.
                // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R2111
                Site.CaptureRequirement(
                    2111,
                    @"[In Appendix C: Product Behavior]Implementation does not implement the ManagementRole part. <wsdl:operation name=""ExportItems"">
   <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/ExportItems"" />
   <wsdl:input>
       <soap:header message=""tns:ExportItemsSoapIn"" part=""Impersonation"" use=""literal""/>
       <soap:header message=""tns:ExportItemsSoapIn"" part=""MailboxCulture"" use=""literal""/>
       <soap:header message=""tns:ExportItemsSoapIn"" part=""RequestVersion"" use=""literal""/>
       <soap:body parts=""request"" use=""literal"" />
   </wsdl:input>
   <wsdl:output>
       <soap:body parts=""ExportItemsResult"" use=""literal"" />
       <soap:header message=""tns:ExportItemsSoapOut"" part=""ServerVersion"" use=""literal""/>
   </wsdl:output>
</wsdl:operation>
(<1> Section 3.1.4.1:  Exchange 2010 does not implement the ManagementRole part.)");
            }
        }
        #endregion

        #region Test case base methods
        #region Get export or upload items
        /// <summary>
        /// Get the ExportItemsResponseMessageType response of ExportItems operation.
        /// </summary>
        /// <param name="configureSOAPHeader">A Boolean value specifies whether configuring the SOAP header before calling operations.</param>
        /// <returns>The array of ExportItemsResponseMessageType response.</returns>
        protected ExportItemsResponseMessageType[] ExportItems(bool configureSOAPHeader)
        {
            #region Prerequisite.
            // In the initialize step, three items in the specified sub folder have been created.
            // If that step executes successfully, the length of CreatedItemId list should be equal to the length of OriginalFolderId list. 
            Site.Assert.AreEqual<int>(
                this.CreatedItemId.Count,
                this.OriginalFolderId.Count,
                string.Format(
                "The exportedItemIds array should contain {0} item ids, actually, it contains {1}",
                this.OriginalFolderId.Count,
                this.CreatedItemId.Count));
            #endregion

            #region Call ExportItems operation to export the items from the server.
            // Initialize the export items' instances.
            ExportItemsType exportItems = new ExportItemsType();

            // Initialize three ItemIdType instances.
            exportItems.ItemIds = new ItemIdType[this.ItemCount];
            for (int i = 0; i < exportItems.ItemIds.Length; i++)
            {
                exportItems.ItemIds[i] = new ItemIdType();
                exportItems.ItemIds[i].Id = this.CreatedItemId[i].Id;
            }

            // Initialize a ExportItemsResponseType instance.
            ExportItemsResponseType exportItemsResponse = this.BTRFAdapter.ExportItems(exportItems);

            // Check whether the operation is executed successfully.
            foreach (ExportItemsResponseMessageType exportResponse in exportItemsResponse.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Success,
                exportResponse.ResponseClass,
                string.Format(
                "Export items should be successful! Expected response code: {0}, actual response code: {1}",
                ResponseCodeType.NoError,
                exportResponse.ResponseCode));
            }

            // If the operation executes successfully, the items in exportItems response should be equal to the value of ItemCount
            Site.Assert.AreEqual<int>(
                exportItemsResponse.ResponseMessages.Items.Length,
                this.ItemCount,
                string.Format(
                "The exportItems response should contain {0} items, actually, it contains {1}",
                this.ItemCount,
                exportItemsResponse.ResponseMessages.Items.Length));
            #endregion

            #region Verify the ExportItems response related requirements.
            // Verify the ExportItemsResponseType related requirements.
            this.VerifyExportItemsSuccessResponse(exportItemsResponse);
            #endregion

            #region Get the ExportItemsResponseMessageType items.
            ExportItemsResponseMessageType[] exportItemsResponseMessages =
                TestSuiteHelper.GetResponseMessages<ExportItemsResponseMessageType>(exportItemsResponse);
            return exportItemsResponseMessages;
            #endregion
        }

        /// <summary>
        /// Get the UploadItemsResponseMessageType response of UploadItems operation when this operation executes successfully.
        /// </summary>
        /// <param name="exportedItems">The items exported from server.</param>
        /// <param name="parentFolderId">Specifies the target folder in which to place the upload item.</param>
        /// <param name="createAction">Specifies the action for uploading items to the folder.</param>
        /// <param name="isAssociatedSpecified">A Boolean value specifies whether IsAssociated attribute is specified.</param>
        /// <param name="isAssociated">Specifies the value of the IsAssociated attribute.</param>
        /// <param name="configureSOAPHeader">A Boolean value specifies whether configuring the SOAP header before calling operations.</param>
        /// <returns>The array of UploadItemsResponseMessageType response.</returns>
        protected UploadItemsResponseMessageType[] UploadItems(
            ExportItemsResponseMessageType[] exportedItems,
            Collection<string> parentFolderId,
            CreateActionType createAction,
            bool isAssociatedSpecified,
            bool isAssociated,
            bool configureSOAPHeader)
        {
            #region Call UploadItems operation to upload the items that exported in last step to the server.
            // Initialize the upload items using the data of previous exported items, and set that item CreateAction to a value of CreateActionType.
            UploadItemsType uploadItems = new UploadItemsType();
            uploadItems.Items = new UploadItemType[this.ItemCount];
            for (int i = 0; i < uploadItems.Items.Length; i++)
            {
                uploadItems.Items[i] = TestSuiteHelper.GenerateUploadItem(
                    exportedItems[i].ItemId.Id,
                    exportedItems[i].ItemId.ChangeKey,
                    exportedItems[i].Data,
                    parentFolderId[i],
                    createAction);
                uploadItems.Items[i].IsAssociatedSpecified = isAssociatedSpecified;
                if (uploadItems.Items[i].IsAssociatedSpecified)
                {
                    uploadItems.Items[i].IsAssociated = isAssociated;
                }
            }

            // Call UploadItems operation.
            UploadItemsResponseType uploadItemsResponse = this.BTRFAdapter.UploadItems(uploadItems);

            // Check whether the operation is executed successfully.
            foreach (UploadItemsResponseMessageType uploadResponse in uploadItemsResponse.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Success,
                uploadResponse.ResponseClass,
                string.Format(
                    @"The UploadItems operation should be successful. Expected response code: {0}, actual response code: {1}",
                    ResponseClassType.Error,
                    uploadResponse.ResponseClass));
            }

            // If the operation executes successfully, the items in UploadItems response should be equal to the value of ItemCount
            Site.Assert.AreEqual<int>(
                uploadItemsResponse.ResponseMessages.Items.Length,
                this.ItemCount,
                string.Format(
                "The uploadItems response should contain {0} items, actually, it contains {1}",
                this.ItemCount,
                uploadItemsResponse.ResponseMessages.Items.Length));
            #endregion

            #region Verify the UploadItemsResponseType related requirements
            // Verify the UploadItemsResponseType related requirements.
            this.VerifyUploadItemsSuccessResponse(uploadItemsResponse);
            #endregion

            #region Call GetResponseMessages to get the UploadItemsResponseMessageType items.
            // Get the UploadItemsResponseMessageType items.
            UploadItemsResponseMessageType[] uploadItemsResponseMessages = TestSuiteHelper.GetResponseMessages<UploadItemsResponseMessageType>(uploadItemsResponse);

            return uploadItemsResponseMessages;
            #endregion
        }

        /// <summary>
        /// Get the UploadItemsResponseMessageType response of UploadItems operation when this operation executes unsuccessfully.
        /// </summary>
        /// <param name="parentFolderId">Specifies the target folder in which to place the upload item.</param>
        /// <param name="createAction">Specifies the action for uploading items to the folder.</param>
        /// <returns>The array of UploadItemsResponseMessageType response.</returns>
        protected UploadItemsResponseMessageType[] UploadInvalidItems(Collection<string> parentFolderId, CreateActionType createAction)
        {
            #region Get the exported items and parent folder ID.
            // Get the exported items which is prepared for uploading.
            ExportItemsResponseMessageType[] exportedItem = this.ExportItems(false);
            #endregion

            #region Call UploadItems operation to upload the items that exported in last step to the server.
            // Initialize the upload items using the data of previous export items, and set that item CreateAction to a value of CreateActionType.
            UploadItemsType uploadInvalidItems = new UploadItemsType();
            uploadInvalidItems.Items = new UploadItemType[4];

            // The ID attribute of ItemId is empty.
            uploadInvalidItems.Items[0] = TestSuiteHelper.GenerateUploadItem(
                string.Empty,
                null,
                exportedItem[0].Data,
                parentFolderId[0],
                createAction);

            // The ID attribute of ItemId is valid and the ChangeKey is invalid.
            uploadInvalidItems.Items[1] = TestSuiteHelper.GenerateUploadItem(
                exportedItem[1].ItemId.Id,
                InvalidChangeKey,
                exportedItem[1].Data,
                parentFolderId[1],
                createAction);

            // The ID attribute of ItemId is invalid and the ChangeKey is null.
            uploadInvalidItems.Items[2] = TestSuiteHelper.GenerateUploadItem(
                InvalidItemId,
                null,
                exportedItem[2].Data,
                parentFolderId[2],
                createAction);

            // The ID attribute of ItemId is invalid and the ChangeKey is null.
            uploadInvalidItems.Items[3] = TestSuiteHelper.GenerateUploadItem(
                InvalidItemId,
                null,
                exportedItem[3].Data,
                parentFolderId[3],
                createAction);

            // Call UploadItems operation.
            UploadItemsResponseType uploadItemsResponse = this.BTRFAdapter.UploadItems(uploadInvalidItems);

            // Check whether the ExportItems operation is executed successfully.
            foreach (UploadItemsResponseMessageType responseMessage in uploadItemsResponse.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                responseMessage.ResponseClass,
                string.Format(
                    @"The UploadItems operation should be unsuccessful. Expected response code: {0}, actual response code: {1}",
                    ResponseClassType.Error,
                    responseMessage.ResponseClass));
            }
            #endregion

            #region Verify the UploadItemsResponseType related requirements
            // Verify the UploadItemsResponseType related requirements.
            this.VerifyUploadItemsErrorResponse(uploadItemsResponse);
            #endregion

            #region Call GetResponseMessages to get the UploadItemsResponseMessageType items.
            // Get the UploadItemsResponseMessageType items.
            UploadItemsResponseMessageType[] uploadItemsResponseMessages = TestSuiteHelper.GetResponseMessages<UploadItemsResponseMessageType>(uploadItemsResponse);

            return uploadItemsResponseMessages;
            #endregion
        }
        #endregion

        #region Construct MS-OXWSFOLD operation request and get the response
        /// <summary>
        /// Create a sub folder in the specified parent folder.
        /// </summary>
        /// <param name="parentFolderType">Type of the parent folder.</param>
        /// <param name="subFolderName">Name of the folder which should be created.</param>
        /// <returns>ID of the new created sub folder.</returns>
        protected string CreateSubFolder(DistinguishedFolderIdNameType parentFolderType, string subFolderName)
        {
            // Variable to specified the created sub folder ID and the folder class name.
            string subFolderId = null;
            string folderClassName = null;

            // Set the folder's class name according to the type of parent folder.
            switch (parentFolderType)
            {
                case DistinguishedFolderIdNameType.contacts:
                    folderClassName = "IPF.Contact";
                    break;
                case DistinguishedFolderIdNameType.calendar:
                    folderClassName = "IPF.Appointment";
                    break;
                case DistinguishedFolderIdNameType.tasks:
                    folderClassName = "IPF.Task";
                    break;
                case DistinguishedFolderIdNameType.inbox:
                    folderClassName = "IPF.Note";
                    break;
                default:
                    Site.Assert.Fail(@"The parent folder type '{0}' is invalid.The valid folder types are: contacts, calendar, tasks and inbox", parentFolderType);
                    break;
            }

            // Initialize the create folder request.
            CreateFolderType createFolderRequest = new CreateFolderType();
            FolderType folderProperties = new FolderType();

            // Set parent folder id.
            createFolderRequest.ParentFolderId = new TargetFolderIdType();
            DistinguishedFolderIdType parentFolder = new DistinguishedFolderIdType();
            parentFolder.Id = parentFolderType;
            createFolderRequest.ParentFolderId.Item = parentFolder;

            // Set Display Name and Folder Class for the folder to be created.
            folderProperties.DisplayName = subFolderName;
            folderProperties.FolderClass = folderClassName;

            // Set permission.
            folderProperties.PermissionSet = new PermissionSetType();
            folderProperties.PermissionSet.Permissions = new PermissionType[1];
            folderProperties.PermissionSet.Permissions[0] = new PermissionType();
            folderProperties.PermissionSet.Permissions[0].CanCreateItems = true;
            folderProperties.PermissionSet.Permissions[0].CanCreateSubFolders = true;
            folderProperties.PermissionSet.Permissions[0].PermissionLevel = new PermissionLevelType();
            folderProperties.PermissionSet.Permissions[0].PermissionLevel = PermissionLevelType.Editor;
            folderProperties.PermissionSet.Permissions[0].UserId = new UserIdType();

            string primaryUserName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            string primaryDomain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            folderProperties.PermissionSet.Permissions[0].UserId.PrimarySmtpAddress = primaryUserName + "@" + primaryDomain;

            createFolderRequest.Folders = new BaseFolderType[1];
            createFolderRequest.Folders[0] = folderProperties;

            // Invoke CreateFolder operation and get the response.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            if (createFolderResponse != null && createFolderResponse.ResponseMessages.Items[0].ResponseClass.ToString() == ResponseClassType.Success.ToString())
            {
                FolderIdType folderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
                subFolderId = folderId.Id;
                FolderType created = new FolderType() { DisplayName = folderProperties.DisplayName, FolderClass = folderClassName, FolderId = folderId, ParentFolderId = new FolderIdType() { Id = parentFolder.Id.ToString() } };
                this.CreatedFolders.Add(created);
            }

            return subFolderId;
        }
        #endregion

        #region Construct MS-OXWSCORE operation request and get the response
        /// <summary>
        /// Create an item in the specified folder.
        /// </summary>
        /// <param name="parentFolderType">Type of the parent folder.</param>
        /// <param name="parentFolderId">ID of the parent folder.</param>
        /// <param name="itemSubject">Subject of the item which should be created.</param>
        /// <returns>ID of the created item.</returns>
        protected ItemIdType CreateItem(DistinguishedFolderIdNameType parentFolderType, string parentFolderId, string itemSubject)
        {
            // Create a request for the CreateItem operation and initialize the ItemType instance.
            CreateItemType createItemRequest = new CreateItemType();
            ItemType item = null;

            // Get different values for item based on different parent folder type.
            switch (parentFolderType)
            {
                case DistinguishedFolderIdNameType.contacts:
                    ContactItemType contact = new ContactItemType();
                    contact.Subject = itemSubject;
                    contact.FileAs = itemSubject;
                    item = contact;
                    break;
                case DistinguishedFolderIdNameType.calendar:
                    // Set the sendMeetingInvitations property.
                    CalendarItemCreateOrDeleteOperationType sendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToNone;
                    createItemRequest.SendMeetingInvitations = (CalendarItemCreateOrDeleteOperationType)sendMeetingInvitations;
                    createItemRequest.SendMeetingInvitationsSpecified = true;
                    CalendarItemType calendar = new CalendarItemType();
                    calendar.Subject = itemSubject;
                    item = calendar;
                    break;
                case DistinguishedFolderIdNameType.inbox:
                    MessageType message = new MessageType();
                    message.Subject = itemSubject;
                    item = message;
                    break;
                case DistinguishedFolderIdNameType.tasks:
                    TaskType taskItem = new TaskType();
                    taskItem.Subject = itemSubject;
                    item = taskItem;
                    break;
                default:
                    Site.Assert.Fail("The parent folder type '{0}' is invalid and the valid folder types are: contacts, calendar, inbox and tasks.", parentFolderType.ToString());
                    break;
            }

            // Set the MessageDisposition property.
            MessageDispositionType messageDisposition = MessageDispositionType.SaveOnly;
            createItemRequest.MessageDisposition = (MessageDispositionType)messageDisposition;
            createItemRequest.MessageDispositionSpecified = true;

            // Specify the folder in which new items are saved.
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            FolderIdType folderId = new FolderIdType();
            folderId.Id = parentFolderId;
            createItemRequest.SavedItemFolderId.Item = folderId;

            // Specify the collection of items to be created.
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = new ItemType[] { item };

            // Initialize the ID of the created item.
            ItemIdType createdItemId = null;

            // Invoke the create item operation and get the response.
            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);

            if (createItemResponse != null && createItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                ItemInfoResponseMessageType info = createItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
                Site.Assert.IsNotNull(info, "The items in CreateItem response should not be null.");

                // Get the ID of the created item.
                createdItemId = info.Items.Items[0].ItemId;
            }

            return createdItemId;
        }

        /// <summary>
        /// Get the items information.
        /// </summary>
        /// <param name="itemIds">The array of item ids.</param>
        /// <returns>The items information.</returns>
        protected ItemType[] GetItems(BaseItemIdType[] itemIds)
        {
            GetItemType getItem = new GetItemType();
            if (itemIds != null)
            {
                getItem.ItemIds = itemIds;
                getItem.ItemShape = new ItemResponseShapeType();
                getItem.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            }

            // Call GetItem operation
            GetItemResponseType getItemResponse = this.COREAdapter.GetItem(getItem);

            // Check whether the GetItem operation is executed successfully.
            foreach (ResponseMessageType responseMessage in getItemResponse.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                    ResponseClassType.Success,
                    responseMessage.ResponseClass,
                    string.Format(
                        "Get items should not be failed! Expected response code: {0}, actual response code: {1}",
                        ResponseCodeType.NoError,
                        responseMessage.ResponseCode));
            }

            // If the operation executes successfully, the items in getItem response should be equal to the value of ItemCount.
            Site.Assert.AreEqual<int>(
                getItemResponse.ResponseMessages.Items.Length,
                this.ItemCount,
                string.Format(
                "The getItem response should contain {0} items, actually it contains {1}",
                this.ItemCount,
                getItemResponse.ResponseMessages.Items.Length));

            // Get the items from successful response.
            ItemType[] getItems = Common.GetItemsFromInfoResponse<ItemType>(getItemResponse);

            return getItems;
        }
        #endregion

        /// <summary>
        /// Handle the server response.
        /// </summary>
        /// <param name="request">The request messages.</param>
        /// <param name="response">The response messages.</param>
        /// <param name="isSchemaValidated">Whether the schema is validated.</param>
        protected void ExchangeServiceBinding_ResponseEvent(
            BaseRequestType request,
            BaseResponseMessageType response,
            bool isSchemaValidated)
        {
            this.IsSchemaValidated = isSchemaValidated;
        }
        #endregion
    }
}