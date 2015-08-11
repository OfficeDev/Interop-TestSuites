namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUT control adapter of MS-OXWSFOLD.
    /// It includes CreateSubFolders and CreateSearchFolder methods which can be implemented with operations defined in MS-OXWSFOLD.
    /// </summary>
    public interface IMS_OXWSFOLDSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Log on to a mailbox with a specified user account and create two different-level subfolders in the specified parent folder.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="parentFolderName">Name of the parent folder.</param>
        /// <param name="firstLevelSubFolderName">Name of the first level sub folder which will be created under the parent folder.</param>
        /// <param name="secondLevelSubFolderName">Name of the second level sub folder which will be created under the first level sub folder.</param>
        /// <returns>If the two level sub folders are created successfully, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName, userPassword, userDomain) and create a subfolder named (firstLevelSubFolderName)" +
            " under (parentFolderName) folder," +
            " then create another subfolder named (secondLevelSubFolderName) under the (firstLevelSubFolderName) folder.\n" +
            " If the operation succeeds, enter \"TRUE\";" +
            " otherwise, enter \"FALSE\".")]
        bool CreateSubFolders(string userName, string userPassword, string userDomain, string parentFolderName, string firstLevelSubFolderName, string secondLevelSubFolderName);

        /// <summary>
        /// Log on to a mailbox with a specified user account and create a search folder.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="searchFolderName">Name of the search folder which should be created.</param>
        /// <param name="searchText">Search text of the search folder which should be created.</param>
        /// <returns>If the search folder is created successfully, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName, userPassword, userDomain) and create a search folder named (searchFolderName)" +
            " which is used to search for items by using the search text (searchText).\n" + 
            " If the operation succeeds, enter \"TRUE\";" +
            " otherwise, enter \"FALSE\".")]
        bool CreateSearchFolder(string userName, string userPassword, string userDomain, string searchFolderName, string searchText);
    }
}