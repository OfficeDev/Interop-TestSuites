//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSTASK
{
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
        /// This flag contains the status of each response message.
        /// </summary>
        private Collection<ResponseClassType> responseClass;

        /// <summary>
        /// This flag contains the detail status information of each response message.
        /// </summary>
        private Collection<ResponseCodeType> responseCode;
        
        #endregion

        #region Properties
        /// <summary>
        /// Gets the MS-OXWSTASK protocol adapter.
        /// </summary>
        protected IMS_OXWSTASKAdapter TASKAdapter { get; private set; }

        /// <summary>
        /// Gets a value which is the status of each response message.
        /// </summary>
        protected Collection<ResponseClassType> ResponseClass
        {
            get { return this.responseClass; }
        }

        /// <summary>
        /// Gets a value which is the detail status information of each response message.
        /// </summary>
        protected Collection<ResponseCodeType> ResponseCode
        {
            get { return this.responseCode; }
        }

        #endregion

        #region Test case initialize
        /// <summary>
        ///  Initial the Test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.TASKAdapter = this.Site.GetAdapter<IMS_OXWSTASKAdapter>();
            this.TASKAdapter.Initialize(this.Site);
        }
        #endregion

        #region Test case base methods

        /// <summary>
        /// Initialize the collections of ResponseClass and ResponseCode.
        /// </summary>
        protected void InitializeCollections()
        {
            // Initialize the collection of response class.
            if (this.responseClass != null)
            {
                this.responseClass.Clear();
            }
            else
            {
                this.responseClass = new Collection<ResponseClassType>();
            }

            // Initialize the collection of response code.
            if (this.responseCode != null)
            {
                this.responseCode.Clear();
            }
            else
            {
                this.responseCode = new Collection<ResponseCodeType>();
            }
        }

        /// <summary>
        /// Creates tasks on the server and extracts the items id from response.
        /// </summary>
        /// <param name="taskItems">The task items which will be created.</param>
        /// <returns>The extracted items id array.</returns>
        protected ItemIdType[] CreateTasks(params TaskType[] taskItems)
        {
            // Configure the CreateItem request. 
            CreateItemType createItemRequest = TestSuiteHelper.GenerateCreateItemRequest(taskItems);

            // Get the CreateItem response from server.
            CreateItemResponseType createItemResponse = this.TASKAdapter.CreateItem(createItemRequest);

            this.VerifyResponseMessage(createItemResponse);

            // Save the ItemId of task item got from the createItem response.
            return Common.GetItemIdsFromInfoResponse(createItemResponse);
        }

        /// <summary>
        /// Get tasks on the server according to items' id.
        /// </summary>
        /// <param name="itemIds">The item id of task which will be gotten.</param>
        /// <returns>The got task items array.</returns>
        protected TaskType[] GetTasks(params ItemIdType[] itemIds)
        {
            GetItemType getItemRequest = TestSuiteHelper.GenerateGetItemRequest(itemIds);

            // Get the GetItem response from the server by using the ItemId got from createItem response.
            GetItemResponseType getItemResponse = this.TASKAdapter.GetItem(getItemRequest);

            this.VerifyResponseMessage(getItemResponse);

            return Common.GetItemsFromInfoResponse<TaskType>(getItemResponse);
        }

        /// <summary>
        /// Copy tasks on the server according to items' id.
        /// </summary>
        /// <param name="itemIds">The item id of tasks which will be copied.</param>
        /// <returns>The extracted items id array.</returns>
        protected ItemIdType[] CopyTasks(params ItemIdType[] itemIds)
        {
            // Define the CopyItem request.
            CopyItemType copyItemRequest = TestSuiteHelper.GenerateCopyItemRequest(itemIds);

            // Call the CopyItem method to copy the task items created in previous steps. 
            CopyItemResponseType copyItemResponse = this.TASKAdapter.CopyItem(copyItemRequest);

            this.VerifyResponseMessage(copyItemResponse);

            // Save the ItemId of task item got from the CopyItem response.
            return Common.GetItemIdsFromInfoResponse(copyItemResponse);
        }

        /// <summary>
        /// Move tasks on the server according to items' id.
        /// </summary>
        /// <param name="itemIds">The item id of tasks which will be moved.</param>
        /// <returns>The extracted items id array.</returns>
        protected ItemIdType[] MoveTasks(params ItemIdType[] itemIds)
        {
            // Define the MoveItem request.
            MoveItemType moveItemRequest = TestSuiteHelper.GenerateMoveItemRequest(itemIds);

            // Call the MoveItem method to move the task items created in previous steps. 
            MoveItemResponseType moveItemResponse = this.TASKAdapter.MoveItem(moveItemRequest);

            this.VerifyResponseMessage(moveItemResponse);

            // Save the ItemId of task item got from the MoveItem response.
            return Common.GetItemIdsFromInfoResponse(moveItemResponse);
        }

        /// <summary>
        /// Delete tasks on the server according to items' id.
        /// </summary>
        /// <param name="itemIds">The item id of tasks which will be deleted.</param>
        protected void DeleteTasks(params ItemIdType[] itemIds)
        {
            // Define the DeleteItem request.
            DeleteItemType deleteItemRequest = TestSuiteHelper.GenerateDeleteItemRequest(itemIds);

            // Call the DeleteItem method to delete the task items created in previous steps. 
            DeleteItemResponseType deleteItemResponse = this.TASKAdapter.DeleteItem(deleteItemRequest);

            this.VerifyResponseMessage(deleteItemResponse);
        }

        /// <summary>
        /// Delete tasks with the condition whether affect task occurrences.
        /// </summary>
        /// <param name="affectedTaskOccurrences">The value indicate whether affect task occurrences.</param>
        /// <param name="itemIds">The item id of tasks which will be deleted.</param>
        protected void DeleteTasks(AffectedTaskOccurrencesType affectedTaskOccurrences, params ItemIdType[] itemIds)
        {
            // Define the DeleteItem request.
            DeleteItemType deleteItemRequest = TestSuiteHelper.GenerateDeleteItemRequest(itemIds);
            deleteItemRequest.AffectedTaskOccurrences = affectedTaskOccurrences;

            // Call the DeleteItem method to delete the task items created in previous steps. 
            DeleteItemResponseType deleteItemResponse = this.TASKAdapter.DeleteItem(deleteItemRequest);

            this.VerifyResponseMessage(deleteItemResponse);
        }

        /// <summary>
        /// Update tasks on the server according to items' id.
        /// </summary>
        /// <param name="itemIds">The item id of tasks which will be updated.</param>
        /// <returns>The extracted items id array.</returns>
        protected ItemIdType[] UpdateTasks(params ItemIdType[] itemIds)
        {
            // Define the UpdateItem request.
            UpdateItemType updateItemRequest = TestSuiteHelper.GenerateUpdateItemRequest(itemIds);

            // Call the UpdateItem method to update the task items created in previous steps. 
            UpdateItemResponseType updateItemResponse = this.TASKAdapter.UpdateItem(updateItemRequest);

            this.VerifyResponseMessage(updateItemResponse);

            // Save the ItemId of task item got from the updateItem response.
            return Common.GetItemIdsFromInfoResponse(updateItemResponse);
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Verify whether the response is a valid response message and extract the ResponseClass and ResponseCode.
        /// </summary>
        /// <param name="baseResponseMessage">The response message.</param>
        private void VerifyResponseMessage(BaseResponseMessageType baseResponseMessage)
        {
            this.InitializeCollections();
            Site.Assert.IsNotNull(baseResponseMessage, @"The response should not be null.");
            Site.Assert.IsNotNull(baseResponseMessage.ResponseMessages, @"The ResponseMessages in response should not be null.");
            Site.Assert.IsNotNull(baseResponseMessage.ResponseMessages.Items, @"The items of ResponseMessages in response should not be null.");
            Site.Assert.IsTrue(baseResponseMessage.ResponseMessages.Items.Length > 0, @"The items of ResponseMessages in response should not be null.");
            int messageCount = baseResponseMessage.ResponseMessages.Items.Length;
            for (int i = 0; i < messageCount; i++)
            {
                ResponseMessageType responseMessage = baseResponseMessage.ResponseMessages.Items[i];
                this.ResponseClass.Add(responseMessage.ResponseClass);

                if (responseMessage.ResponseCodeSpecified)
                {
                    this.responseCode.Add(responseMessage.ResponseCode);
                }

                if (responseMessage.ResponseClass == ResponseClassType.Success && !(baseResponseMessage is DeleteItemResponseType))
                {
                    ItemInfoResponseMessageType itemInfo = responseMessage as ItemInfoResponseMessageType;
                    Site.Assert.IsNotNull(itemInfo.Items, @"The items of ResponseMessages in response should not be null.");
                    Site.Assert.IsNotNull(itemInfo.Items.Items, @"The task items in response should not be null.");
                    Site.Assert.IsTrue(itemInfo.Items.Items.Length > 0, @"There is one task item at least in response.");
                    foreach (ItemType item in itemInfo.Items.Items)
                    {
                        Site.Assert.IsNotNull(item.ItemId, @"The task item id in response should not be null.");
                    }
                }
            }
        }
        #endregion
    }
}
