namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// SUT control managed code adapter interface.
    /// </summary>
    public interface IMS_FSSHTTP_FSSHTTPBManagedCodeSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// This method is used to remove the file from the path of file URI.
        /// </summary>
        /// <param name="fileUrl">Specify the URL in where the file will be removed.</param>
        /// <param name="fileName">Specify the name for the file that will be removed.</param>
        /// <returns>Return true if the operation succeeds, otherwise return false.</returns>
        [MethodHelp(@"Remove the file(fileName) from the specified URL(fileUrl). Enter True, if the file is removed successfully; otherwise, enter False.")]
        bool RemoveFile(string fileUrl, string fileName);

        /// <summary>
        /// This method is used to check out the specified file using the specified credential.
        /// </summary>
        /// <param name="fileUrl">Specify the absolute URL of a file which needs to be checked out.</param>
        /// <param name="userName">Specify the name of the user who checks out the file.</param>
        /// <param name="password">Specify the password of the user.</param>
        /// <param name="domain">Specify the domain of the user.</param>
        /// <returns>Return true if the check out succeeds, otherwise return false.</returns>
        [MethodHelp(@"Check out the file (fileUrl) using the credential (userName, password and domain)." +
                      "If the operation succeeds, enter true; otherwise, enter false.")]
        bool CheckOutFile(string fileUrl, string userName, string password, string domain);


        /// <summary>
        /// This method is used to check in the specified file using the specified credential.
        /// </summary>
        /// <param name="fileUrl">Specify the absolute URL of a file which needs to be checked in.</param>
        /// <param name="userName">Specify the name of the user who checks in the file.</param>
        /// <param name="password">Specify the password of the user.</param>
        /// <param name="domain">Specify the domain of the user.</param>
        /// <param name="checkInComments">Specify the checked in comments.</param>
        /// <returns>Return true if the check in succeeds, otherwise return false.</returns>
        [MethodHelp(@"Check in the file (fileUrl) using the credential (userName, password and domain) and comments (checkInComments)." +
                      "If the operation succeeds, enter true; otherwise, enter false.")]
        bool CheckInFile(string fileUrl, string userName, string password, string domain, string checkInComments);
    }
}