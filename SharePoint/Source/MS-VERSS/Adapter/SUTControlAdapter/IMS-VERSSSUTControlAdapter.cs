namespace Microsoft.Protocols.TestSuites.MS_VERSS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUT Control Adapter interface of MS-VERSS.
    /// If the interface methods involve list or file, the file is added in the list, 
    /// the list is generated in the site collection and the site collection is configured as 
    /// the SiteCollectionName property in the MS-VERSS_TestSuite.deployment.ptfconfig file.
    /// </summary>
    public interface IMS_VERSSSUTControlAdapter : IAdapter
    {
        #region Interact with ListsService

        /// <summary>
        /// Upload a file to the specified list.
        /// </summary>
        /// <param name="listName">The name of the list used for this protocol testing.</param>
        /// <param name="fileName">The name of the file uploaded for protocol testing.</param>
        /// <param name="uploadFilePath">The path of the file.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully,
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        [MethodHelp("Add the file specified by the path (uploadFilePath) into the specified list (listName)," +
            " and rename it to the specified name (fileName)." + 
            " The return value should be a Boolean value that indicates whether the operation was run successfully," + 
            " TRUE means the operation was run successfully, FALSE means the operation failed.")]
        bool AddFile(string listName, string fileName, string uploadFilePath);

        /// <summary>
        /// Create a sub folder into the specified list.
        /// </summary>
        /// <param name="listName">The name of the list used for this protocol testing.</param>
        /// <param name="folderName">The name of the sub folder used for protocol testing.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully,
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        [MethodHelp("Create a sub folder specified by the name (folderName) into the specified list (listName)," +
            " The return value should be a Boolean value that indicates whether the operation was run successfully," +
            " TRUE means the operation was run successfully, FALSE means the operation failed.")]
        bool AddFolder(string listName, string folderName);

        /// <summary>
        /// Set whether check out is enforced in the list.
        /// </summary>
        /// <param name="listName">The name of the list used for this protocol testing.</param>
        /// <param name="enforce">The Boolean value indicates whether check out is enforced in the list,
        /// TRUE represents that the file must be checked out before being modified.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully, 
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        [MethodHelp("Determine whether check-out is enforced in the specified list (listName). " +
            "The TRUE value of enforce (enforce) means check-out is enforced," + 
            " the FALSE value of enforce (enforce) means check out is not enforced. " +
            "The return value should be a Boolean value that indicates whether" + 
            " the operation was run successfully, TRUE means the operation was run successfully, " +
            "FALSE means the operation failed.")]
        bool SetEnforceCheckout(string listName, bool enforce);

        /// <summary>
        /// Set whether the file is published.
        /// </summary>
        /// <param name="listName">The name of the list used for this protocol testing.</param>
        /// <param name="fileName">The location of the file which will be published.</param>
        /// <param name="publish">The Boolean value indicates the status of the file is published or unpublished.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully, 
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        [MethodHelp("Determine whether the specified file (fileName) in the specified list (listName) is published." +
            " The TRUE value of publish (publish) means the file is published," +
            " the FALSE value of publish (publish) means the file is not published. The return value should be " +
            "a Boolean value that indicates whether the operation was run successfully, " +
            "TRUE means the operation was run successfully, FALSE means the operation failed.")]
        bool SetFilePublish(string listName, string fileName, bool publish);

        /// <summary>
        /// Set whether versioning is enabled in the list.
        /// </summary>
        /// <param name="listName">The name of the list used for this protocol testing.</param>
        /// <param name="enableVersioning">The Boolean value indicates whether versioning is enabled.</param>
        /// <param name="enableMinorVersions">The Boolean indicates whether minor versions are enabled 
        /// when versioning is enabled for the document library.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully, 
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        [MethodHelp("Determine whether versioning is enabled in the specified list (listName). " +
            "The TRUE value of versioning (enableVersioning) means the versioning is enabled," +
            " the FALSE value of versioning (enableVersioning) means the versioning is not enabled." +
            " The TRUE value of minor versions (enableMinorVersions) means that minor versioning is enabled," +
            " the FALSE value of minor versions (enableMinorVersions) means that minor versioning is not enabled." +
            " The return value should be a Boolean value that indicates whether" +
            " the operation was run successfully, TRUE means the operation was run successfully, " +
            "FALSE means the operation failed.")]
        bool SetVersioning(string listName, bool enableVersioning, bool enableMinorVersions);

        /// <summary>
        /// Get all versions of the specified file in the specified list.
        /// </summary>
        /// <param name="listName">The name of the list used for this protocol testing.</param>
        /// <param name="fileName">The location of a file on the protocol server.</param>
        /// <returns> The return value should be a String value that includes all versions of the specified file, 
        /// the format of each version should be major version plus dot plus minor version, such as 0.1, 1.0, etc.
        /// The most recent version should be preceded with @. All the other versions should not have any prefix.
        /// And versions should be separated by ^. For example "0.1^@0.2". 
        /// If the operation failed then return empty string.</returns>
        [MethodHelp("Get all versions of the specified file (fileName) in the specified list (listName). " +
            "The return value should be a string value that includes all versions" +
            " of the specified file, the format of each version should be major version plus dot plus minor version," +
            " such as 0.1, 1.0, etc. The most recent version" +
            " should be preceded with @. All other versions should not have any prefix. " +
            "Versions should be separated by a ^. For example \"0.1^@0.2\". " +
            "If the operation fails, then an empty string is returned.")]
        string GetFileVersions(string listName, string fileName);

        /// <summary>
        /// Get the attributes of the specified version of the file.
        /// </summary>
        /// <param name="listName">The name of the list used for this protocol testing.</param>
        /// <param name="fileName">The location of the file which will be gotten.</param>
        /// <param name="fileVersion">The version number of the file which will be gotten.</param>
        /// <returns>The return value should be a String value that includes the createdByName attribute 
        /// and the size attribute. The createdByName attribute is a String value that indicates the creator of
        /// the specified version of the file. The size attribute is a String value that indicates the size,
        /// in bytes, of the specified version of the file. The createdByName attribute and 
        /// the size attribute are separated by ^. For example "CONTOSO\Administrator^20". 
        /// If the operation failed then return empty string.</returns>
        [MethodHelp("Get the attributes of the specified version (fileVersion) of the file (fileName) in" +
            " the specified list (listName). The return value should be a string" + 
            " value that includes the createdByName attribute and the size attribute. The createdByName " +
            "attribute is a string value that indicates the creator of the specified" + 
            " version of the file. The size attribute is a string value that indicates the size, in bytes," +
            " of the specified version of the file. The createdByName attribute and the size attribute are" +
            " separated by ^. For example \"CONTOSO\\Administrator^20\". If the operation fails, then an empty string is returned.")]
        string GetFileVersionAttributes(string listName, string fileName, string fileVersion);

        /// <summary>
        /// Check whether the specified file with the specified version exists in recycle bin.
        /// </summary>
        /// <param name="fileName">The specified file name.</param>
        /// <param name="version">The specified version.</param>
        /// <returns>The Boolean value indicates whether the specified file with 
        /// the specified version exists in recycle bin.</returns>
        [MethodHelp("Check whether the specified file (fileName) with the specified version (version) exists in the Recycle Bin." +
            " The return value should be a Boolean value" + 
            " that indicates whether the specified file with the specified version exists in the Recycle Bin, " +
            "TRUE means the specified file with the specified version" + 
            " exists in the Recycle Bin, FALSE means the specified file with " +
            "the specified version does not exist in the Recycle Bin.")]
        bool IsFileExistInRecycleBin(string fileName, string version);

        /// <summary>
        /// Delete the items whose original locations were in the specified list from recycle bin.
        /// </summary>
        /// <param name="listName">The name of the list used for this protocol testing.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully, 
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        [MethodHelp("Delete the items whose original locations were in the specified list (listName)" +
            " from the Recycle Bin. The return value should be a Boolean value" + 
            " that indicates whether the operation was run successfully, " +
            "TRUE means the operation was run successfully, FALSE means the operation failed.")]
        bool DeleteItemsInListFromRecycleBin(string listName);

        /// <summary>
        /// Set whether the recycle bin is enabled for site. 
        /// </summary>
        /// <param name="isEnabled">A Boolean value indicates whether the recycle bin is enabled.</param>
        /// <returns>A Boolean value indicates whether the set operation succeed.</returns>
        [MethodHelp("Determine whether the Recycle Bin is enabled in the specified site. " +
            "The TRUE value of the Recycle Bin enabled (isEnabled) means the Recycle Bin is enabled," +
            " the FALSE value of the Recycle Bin enabled (isEnabled) means the Recycle Bin is not enabled. " +
            "The return value should be a Boolean value " + 
            " that indicates whether the operation was run successfully, " +
            "TRUE means the operation was run successfully, FALSE means the operation failed.")]
        bool SetRecycleBinEnable(bool isEnabled);

        /// <summary>
        /// A method used to get the value of Recycle Bin setting of site.
        /// </summary>
        /// <returns>A return value represents the value of Recycle Bin setting of site.</returns>
        [MethodHelp(@"Get the Recycle Bin setting (On or Off) of the site and enter the appropriate value in action results. Accepted values are True or false, enter ""True"" if the setting is on & ""False"" if the setting is off.")]
        bool GetRecycleBin();
        #endregion
    }
}