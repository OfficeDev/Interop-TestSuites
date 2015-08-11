namespace Microsoft.Protocols.TestSuites.MS_OXWSBTRF
{
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class contains all helper methods used in test cases.
    /// </summary>
    public static class TestSuiteHelper
    {
        #region Static method
        /// <summary>
        /// Get items in ResponseMessages.
        /// </summary>
        /// <typeparam name="T">An enumerator type.</typeparam>
        /// <param name="response">A web service call response.</param>
        /// <returns>An array instance of the specified ResponseMessageType items.</returns>
        public static T[] GetResponseMessages<T>(BaseResponseMessageType response)
            where T : ResponseMessageType
        {
            T[] responseMessages = new T[response.ResponseMessages.Items.Length];
            for (int i = 0; i < response.ResponseMessages.Items.Length; i++)
            {
                // Get the uploaded ExportItemsResponseMessageType items
                responseMessages[i] = (T)response.ResponseMessages.Items[i];
            }

            return responseMessages;
        }

        /// <summary>
        /// Set the contents of a single item to upload into a mailbox.
        /// </summary>
        /// <param name="id">The identifier of the upload item.</param>
        /// <param name="changeKey">The ChangeKey of the upload item.</param>
        /// <param name="data">The data of the upload item.</param>
        /// <param name="parentFolderId">The target folder in which to place the upload item.</param>
        /// <param name="createAction">The action for uploading items into a mailbox.</param>
        /// <returns>The upload item.</returns>
        public static UploadItemType GenerateUploadItem(
            string id,
            string changeKey,
            byte[] data,
            string parentFolderId,
            CreateActionType createAction)
        {
            UploadItemType uploadItem = new UploadItemType();
            uploadItem.ItemId = new ItemIdType();
            uploadItem.ItemId.Id = id;
            uploadItem.ItemId.ChangeKey = changeKey;
            uploadItem.Data = data;
            uploadItem.ParentFolderId = new FolderIdType();
            uploadItem.ParentFolderId.Id = parentFolderId;
            uploadItem.CreateAction = createAction;

            return uploadItem;
        }
        #endregion
    }
}