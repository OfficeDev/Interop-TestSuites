namespace Microsoft.Protocols.TestSuites.MS_VERSS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-VERSS Adapter class.
    /// </summary>
    public interface IMS_VERSSAdapter : IAdapter
    {
        #region Interact with versionsService
        /// <summary>
        /// This operation is used to get details about all versions of the specified file that the user can access.
        /// </summary>
        /// <param name="fileName">The site-relative path of a file on the protocol server.</param>
        /// <returns>The response message for getting all versions of 
        /// the specified file that the user can access.</returns>
        GetVersionsResponseGetVersionsResult GetVersions(string fileName);

        /// <summary>
        /// This operation is used to restore the specified file to a specific version.
        /// </summary>
        /// <param name="fileName">The site-relative path of the file which will be restored.</param>
        /// <param name="fileVersion">The version number of the file which will be restored.</param>
        /// <returns>The response message for restoring the specified file to a specific version.</returns>
        RestoreVersionResponseRestoreVersionResult RestoreVersion(string fileName, string fileVersion);

        /// <summary>
        /// The DeleteVersion operation is used to delete a specific version of the specified file. 
        /// </summary>
        /// <param name="fileName">The site-relative path of the file name whose version is to be deleted.</param>
        /// <param name="fileVersion">The number of the file version to be deleted.</param>
        /// <returns>The response message for deleting a version of the specified file on the protocol server.</returns>
        DeleteVersionResponseDeleteVersionResult DeleteVersion(string fileName, string fileVersion);

        /// <summary>
        /// This operation is used to delete all the previous versions of the specified file
        /// except the published version and the current version.
        /// </summary>
        /// <param name="fileName">The site-relative path of the file which will be deleted.</param>
        /// <returns>The response message for deleting all previous versions of
        /// the specified file on the protocol server.</returns>
        DeleteAllVersionsResponseDeleteAllVersionsResult DeleteAllVersions(string fileName);

        /// <summary>
        /// Initialize a protocol web service using incorrect authorization information.
        /// </summary>
        void InitializeUnauthorizedService();
        #endregion
    }
}