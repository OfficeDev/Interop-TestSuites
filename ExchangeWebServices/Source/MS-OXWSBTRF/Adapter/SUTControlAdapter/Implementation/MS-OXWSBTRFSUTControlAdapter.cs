namespace Microsoft.Protocols.TestSuites.MS_OXWSBTRF
{
    using System;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OXWSBTRF SUT control adapter implementation. 
    /// </summary>
    public class MS_OXWSBTRFSUTControlAdapter : ManagedAdapterBase, IMS_OXWSBTRFSUTControlAdapter
    {
        #region Fields
        /// <summary>
        /// Exchange Service Binding.
        /// </summary> 
        private ExchangeServiceBinding exchangeServiceBinding;

        /// <summary>
        /// User name used to access web service.
        /// </summary>
        private string userName;

        /// <summary>
        /// Password used to access web service.
        /// </summary>
        private string password;

        /// <summary>
        /// Domain of server.
        /// </summary>
        private string domain;

        /// <summary>
        /// Url used to access Web Service.
        /// </summary>
        private string url;
        #endregion

        #region Initialize TestSuite
        /// <summary>
        /// Initialize the test site.
        /// </summary>
        /// <param name="testSite">default testSite</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXWSBTRF";

            // Merge common configuration and SHOULD/MAY configuration files 
            Common.MergeConfiguration(testSite);

            // Get the parameters from configuration files.
            this.userName = Common.GetConfigurationPropertyValue("UserName", testSite);
            this.password = Common.GetConfigurationPropertyValue("UserPassword", testSite);
            this.domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            this.url = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            this.exchangeServiceBinding = new ExchangeServiceBinding(this.url, this.userName, this.password, this.domain, testSite);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
        }
        #endregion

        #region IMS_OXWSBTRFSUTControlAdapter Operations
        /// <summary>
        /// Log on to a mailbox with a specified user account and delete all the subfolders from the specified folder.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="folderName">Name of the folder which should be cleaned up.</param>
        /// <param name="destFolderName">The name of the destination folder which will be deleted.</param>
        /// <returns>If the folder is cleaned up successfully, return true; otherwise, return false.</returns>
        public bool CleanupFolder(string userName, string userPassword, string userDomain, string folderName, string destFolderName)
        {
            // Log on mailbox with specified user mailbox.
            this.exchangeServiceBinding.Credentials = new NetworkCredential(userName, userPassword, userDomain);
            #region Delete all sub folders and the items in these sub folders in the specified parent folder.
            // Parse the parent folder name.
            DistinguishedFolderIdNameType parentFolderName = (DistinguishedFolderIdNameType)Enum.Parse(typeof(DistinguishedFolderIdNameType), folderName, true);

            // Create an array of BaseFolderType.
            BaseFolderType[] folders = null;

            // Create the request and specify the traversal type.
            FindFolderType findFolderRequest = new FindFolderType();
            findFolderRequest.Traversal = FolderQueryTraversalType.Deep;

            // Define the properties to be returned in the response.
            FolderResponseShapeType responseShape = new FolderResponseShapeType();
            responseShape.BaseShape = DefaultShapeNamesType.Default;
            findFolderRequest.FolderShape = responseShape;

            // Identify which folders to search.
            DistinguishedFolderIdType[] folderIDArray = new DistinguishedFolderIdType[1];
            folderIDArray[0] = new DistinguishedFolderIdType();
            folderIDArray[0].Id = parentFolderName;

            // Add the folders to search to the request.
            findFolderRequest.ParentFolderIds = folderIDArray;

            // Invoke FindFolder operation and get the response.
            FindFolderResponseType findFolderResponse = this.exchangeServiceBinding.FindFolder(findFolderRequest);

            // If there are folders found under the specified folder, delete all of them.
            if (findFolderResponse != null && findFolderResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                // Get the folders from the response.
                FindFolderResponseMessageType findFolderResponseMessageType = findFolderResponse.ResponseMessages.Items[0] as FindFolderResponseMessageType;
                Site.Assert.IsNotNull(findFolderResponseMessageType, "The items in FindFolder response should not be null.");

                folders = findFolderResponseMessageType.RootFolder.Folders;
                if (folders.Length != 0)
                {
                    ////Indicates whether the destination folder was found and removed.
                    bool found = false;

                    // Loop to delete all the found folders.
                    foreach (BaseFolderType currentFolder in folders)
                    {
                        if (string.Compare(currentFolder.DisplayName, destFolderName, StringComparison.InvariantCultureIgnoreCase) != 0)
                        {
                            continue;
                        }

                        FolderIdType responseFolderId = currentFolder.FolderId;

                        FolderIdType folderId = new FolderIdType();
                        folderId.Id = responseFolderId.Id;

                        DeleteFolderType deleteFolderRequest = new DeleteFolderType();
                        deleteFolderRequest.DeleteType = DisposalType.HardDelete;
                        deleteFolderRequest.FolderIds = new BaseFolderIdType[1];
                        deleteFolderRequest.FolderIds[0] = folderId;

                        // Invoke DeleteFolder operation and get the response.
                        DeleteFolderResponseType deleteFolderResponse = this.exchangeServiceBinding.DeleteFolder(deleteFolderRequest);

                        Site.Assert.AreEqual<ResponseClassType>(
                            ResponseClassType.Success,
                            deleteFolderResponse.ResponseMessages.Items[0].ResponseClass,
                            "The delete folder operation should be successful.");

                        found = true;
                        break;
                    }

                    Site.Assert.IsTrue(
                           found,
                           "The destination folder can not be found in the assigned parent folder.");
                }
            }
            #endregion

            #region Check whether sub folders are deleted successfully.
            // Invoke the FindFolder operation again.
            findFolderResponse = this.exchangeServiceBinding.FindFolder(findFolderRequest);

            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                findFolderResponse.ResponseMessages.Items[0].ResponseCode,
                string.Format(
                "The delete folder operation should be successful. Expected response code: {0}, actual response code: {1}",
                ResponseCodeType.NoError,
                findFolderResponse.ResponseMessages.Items[0].ResponseCode));

            // Get the found folders from the response.
            FindFolderResponseMessageType findFolderResponseMessage = findFolderResponse.ResponseMessages.Items[0] as FindFolderResponseMessageType;
            folders = findFolderResponseMessage.RootFolder.Folders;

            // If no sub folders that created by case could be found, the folder has been cleaned up successfully.
            foreach (BaseFolderType folder in folders)
            {
                if (string.Compare(folder.DisplayName, destFolderName, StringComparison.InvariantCultureIgnoreCase) != 0)
                {
                    continue;
                }
                else
                {
                    return false;
                }
            }

            return true;
            #endregion
        }
        #endregion
    }
}