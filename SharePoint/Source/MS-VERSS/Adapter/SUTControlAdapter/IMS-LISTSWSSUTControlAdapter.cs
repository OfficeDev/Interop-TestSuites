namespace Microsoft.Protocols.TestSuites.MS_VERSS
{
    using System;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUT control adapter of MS-LISTSWS. 
    /// If the interface methods involve list or file, the file is added in the list, 
    /// the list is generated in the site collection and the site collection is configured as 
    /// the SiteCollectionName property in the MS-VERSS_TestSuite.deployment.ptfconfig file.
    /// </summary>
    public interface IMS_LISTSWSSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Creates a list in site collection.
        /// </summary>
        /// <param name="listName">The name of list that will be created in site collection.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully, 
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        [MethodHelp("Add the specified list (listName) to the server." + 
            " The return value should be a Boolean value that indicates whether" + 
            " the operation was run successfully, TRUE means the operation was run successfully," +
            " FALSE means the operation failed.")]
        bool AddList(string listName);

        /// <summary>
        /// Delete the specified list in site collection.
        /// </summary>
        /// <param name="listName">The name of list which will be deleted.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully, 
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        [MethodHelp("Delete the specified list (listName) from the server. The return value should be a Boolean value that" +
            " indicates whether the operation was" + 
            " run successfully, TRUE means the operation was run successfully, FALSE means the operation failed.")]
        bool DeleteList(string listName);

        /// <summary>
        /// Check in file to a document library.
        /// </summary>
        /// <param name="pageUrl">The URL of the file to be checked in.</param>
        /// <param name="comments">A string containing check-in comments.</param>
        /// <param name="checkInType">A string representation of the values:
        /// 0 means check in minor version, 1 means check in major version or 2 means check in as overwrite.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully, 
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        [MethodHelp("Check in the file (pageUrl) with comments (comments). The check-in type (checkInType) should be:" +
            " 0 means check in as a minor version," + 
            " 1 means check in as a major version or 2 means check in as overwrite. " +
            "If minor version is disabled, the \"checkInType\" parameter should be ignored." +
            " The return value should be a Boolean value that indicates whether the operation was" + 
            " run successfully, TRUE means the operation was run successfully, FALSE means the operation failed.")]
        bool CheckInFile(Uri pageUrl, string comments, string checkInType);

        /// <summary>
        /// Check out a file in a document library.
        /// </summary>
        /// <param name="pageUrl">The URL of the file to be checked out.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully,
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        [MethodHelp("Check out the file (pageUrl). The return value should be a Boolean value that indicates" +
            " whether the operation was run successfully," + 
            " TRUE means the operation was run successfully, FALSE means the operation failed.")]
        bool CheckoutFile(Uri pageUrl);

        /// <summary>
        /// Get the id of specified list.
        /// </summary>
        /// <param name="listName">The specified list name.</param>
        /// <returns>The string value indicates the id of the specified list.</returns>
        [MethodHelp("Get the ID of the specified list (listName). The return value should be a string value " +
            "that indicates the ID of the specified list. If the operation fails, an empty string is returned.")]
        string GetListID(string listName);
    }
}