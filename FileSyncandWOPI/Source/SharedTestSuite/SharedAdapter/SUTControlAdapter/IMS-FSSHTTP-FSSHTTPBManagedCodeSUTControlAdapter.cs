namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// SUT control managed code adapter interface.
    /// </summary>
    public interface IMS_FSSHTTP_FSSHTTPBManagedCodeSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// This method is used to upload a file to the specified URI.
        /// </summary>
        /// <param name="fileUrl">Specify the URL where the file will be uploaded to.</param>
        /// <param name="fileName">Specify the name for the file to upload.</param>
        /// <returns>Return true if the operation succeeds, otherwise return false.</returns>
        [MethodHelp(@"Upload the file(fileName) to the specified URL(fileUrl). Enter True, if the upload succeeds; otherwise, enter False.")]
        bool UploadTextFile(string fileUrl, string fileName);

        /// <summary>
        /// This method is used to remove the file from the path of file URI.
        /// </summary>
        /// <param name="fileUrl">Specify the URL in where the file will be removed.</param>
        /// <param name="fileName">Specify the name for the file that will be removed.</param>
        /// <returns>Return true if the operation succeeds, otherwise return false.</returns>
        [MethodHelp(@"Remove the file(fileName) from the specified URL(fileUrl). Enter True, if the file is removed successfully; otherwise, enter False.")]
        bool RemoveFile(string fileUrl, string fileName);
    }
}