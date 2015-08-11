namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUTControlAdapter's interface definition.
    /// </summary>
    public interface IMS_COPYSSUTControlAdapter : IAdapter
    {
        #region Interact with ListsService

        /// <summary>
        /// This method is used to delete the files.
        /// </summary>
        /// <param name="fileUrls">Specify the file URLs that will be deleted. Each file URL is split by ";" symbol</param>
        /// <returns>Return true if the operation succeeds, otherwise return false.</returns>
        [MethodHelp(@"Remove the files from the specified (fileUrls). If deleting the file succeeds, then enter true, otherwise enter false.")]
        bool DeleteFiles(string fileUrls);

        /// <summary>
        /// This method is used to upload a file to the specified full file URL. The file's content will be random generated, and encoded with UTF8.
        /// </summary>
        /// <param name="fileUrl">A parameter represents the absolute URL of a file, where the file will be uploaded. The file URL must point to the destination SUT which is indicated by "SutComputerName" property in "SharePointCommonConfiguration.deployment.ptfconfig" file.</param>
        /// <returns>Return true if the operation succeeds, otherwise returns false.</returns>
        [MethodHelp(@"Upload the text file to the specified (fileUrl). The file must contain content and the binary file must use the UTF8 encoding format. If the file upload succeeds, then enter true, otherwise enter false.")]
        bool UploadTextFile(string fileUrl);

        /// <summary>
        /// A method used to check out a file by specified user credential.
        /// </summary>
        /// <param name="fileUrl">A parameter represents the absolute URL of a file which will be checked out by specified user.</param>
        /// <param name="userName">A parameter represents the user name which will check out the file. The file must be stored in the destination SUT which is indicated by "SutComputerName" property in "SharePointCommonConfiguration.deployment.ptfconfig" file.</param>
        /// <param name="password">A parameter represents the password of the user.</param>
        /// <param name="domain">A parameter represents the domain of the user.</param>
        /// <returns>Return true if the operation succeeds, otherwise returns false.</returns>
        [MethodHelp(@"Check out the file by using the specified user (userName). If the checkout succeeds, then enter true, otherwise enter false.")]
        bool CheckOutFileByUser(string fileUrl, string userName, string password, string domain);

        /// <summary>
        /// A method used to undo checkout for a file by specified user credential.
        /// </summary>
        /// <param name="fileUrl">A parameter represents the absolute URL of a file which will be undo checkout by specified user.</param>
        /// <param name="userName">A parameter represents the user name which will undo checkout the file. The file must be stored in the destination SUT which is indicated by "SutComputerName" property in "SharePointCommonConfiguration.deployment.ptfconfig" file.</param>
        /// <param name="password">A parameter represents the password of the user.</param>
        /// <param name="domain">A parameter represents the domain of the user.</param>
        /// <returns>Return true if the operation succeeds, otherwise returns false.</returns>
        [MethodHelp(@"Undo the Checkout for the file by using the specified user (userName). If the undo checkout succeeds, then enter true, otherwise enter false.")]
        bool UndoCheckOutFileByUser(string fileUrl, string userName, string password, string domain);

        #endregion
    }
}