namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class contains all helper methods used in test cases.
    /// </summary>
    public static class TestSuiteHelper
    {
        /// <summary>
        /// Creates a request for the SyncFolderHierarchy operation.
        /// </summary>
        /// <param name="folder">A default folder name.</param>
        /// <param name="defaultShapeNames">Standard sets of properties to return.</param>
        /// <param name="isSyncFolderIdPresent">A Boolean value indicates whether the SyncFolderId element present in the request.</param>
        /// <param name="isSyncStatePresent">A Boolean value indicates whether the SyncState element present in the request.</param>
        /// <returns>An instance of SyncFolderHierarchyType used by the SyncFolderHierarchy operation.</returns>
        public static SyncFolderHierarchyType CreateSyncFolderHierarchyRequest(
            DistinguishedFolderIdNameType folder,
            DefaultShapeNamesType defaultShapeNames,
            bool isSyncFolderIdPresent,
            bool isSyncStatePresent)
        {
            // Create an instance of SyncFolderHierarchyType
            SyncFolderHierarchyType request = new SyncFolderHierarchyType();

            request.FolderShape = new FolderResponseShapeType();
            request.FolderShape.BaseShape = defaultShapeNames;

            // Set the value of SyncFolderId if this element is present in the request.
            if (isSyncFolderIdPresent)
            {
                request.SyncFolderId = new TargetFolderIdType();
                DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
                distinguishedFolderId.Id = folder;
                request.SyncFolderId.Item = distinguishedFolderId;
            }

            // Set the value of SyncState element if this element is present in the request.
            if (isSyncStatePresent)
            {
                request.SyncState = string.Empty;
            }

            return request;
        }

        /// <summary>
        /// Create a request without optional elements for SyncFolderHierarchy operation.
        /// </summary>
        /// <returns>An instance of SyncFolderHierarchyType used by the SyncFolderHierarchy operation.</returns>
        public static SyncFolderHierarchyType CreateSyncFolderHierarchyRequest()
        {
            // Only FolderShape is required element.
            SyncFolderHierarchyType request = new SyncFolderHierarchyType();
            request.FolderShape = new FolderResponseShapeType();
            request.FolderShape.BaseShape = DefaultShapeNamesType.IdOnly;

            return request;
        }

        /// <summary>
        /// Parses items in ResponseMessages
        /// </summary>
        /// <typeparam name="T">the type of the items.</typeparam>
        /// <param name="response">a web service call Response</param>
        /// <returns>An array instance of the specified ResponseMessageType</returns>
        public static T[] ParseResponse<T>(BaseResponseMessageType response)
            where T : ResponseMessageType
        {
            T[] responseMessage;
            responseMessage = new T[response.ResponseMessages.Items.Length];
            for (int i = 0; i < responseMessage.Length; i++)
            {
                responseMessage[i] = response.ResponseMessages.Items[i] as T;
            }

            return responseMessage;
        }

        /// <summary>
        /// Parses the first item in ResponseMessages, ensure the returned value is correct.
        /// </summary>
        /// <typeparam name="T">Type of the item.</typeparam>
        /// <param name="response">A web service call response.</param>
        /// <returns>An instance of the specified ResponseMessageType.</returns>
        public static T EnsureResponse<T>(BaseResponseMessageType response)
            where T : ResponseMessageType
        {
            T[] responseMessage = ParseResponse<T>(response);
            if (responseMessage.Length != 1)
            {
                throw new IncorrectItemCountException("There should be only 1 item in the response.");
            }

            if (responseMessage[0].ResponseClass != ResponseClassType.Success || responseMessage[0].ResponseCode != ResponseCodeType.NoError)
            {
                throw new ResponseErrorException(
                    string.Format(
                    "ResponseClass or ResponseCode error. Response Code: {0}; Response Class: {1}; MessageText: {2}",
                    responseMessage[0].ResponseCode,
                    responseMessage[0].ResponseClass,
                    responseMessage[0].MessageText));
            }

            return responseMessage[0];
        }

        /// <summary>
        /// An exception class provides the exception while the count of items in a collection is illegal .
        /// </summary>
        public class IncorrectItemCountException : ApplicationException
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="IncorrectItemCountException"/> class. 
            /// </summary>
            /// <param name="message">Prompt information when exception occurs.</param>
            public IncorrectItemCountException(string message)
                : base(message)
            {
            }
        }

        /// <summary>
        /// An exception class provides the exception while the response error occurs.
        /// </summary>
        public class ResponseErrorException : ApplicationException
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="ResponseErrorException"/> class. 
            /// </summary>
            /// <param name="message">Prompt information when exception occurs.</param>
            public ResponseErrorException(string message)
                : base(message)
            {
            }
        }
    }    
}