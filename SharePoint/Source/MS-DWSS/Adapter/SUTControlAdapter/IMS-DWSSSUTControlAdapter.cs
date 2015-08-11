namespace Microsoft.Protocols.TestSuites.MS_DWSS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUT control adapter of MS-DWSS.
    /// </summary>
    public interface IMS_DWSSSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Creates a list in site collection.
        /// </summary>
        /// <param name="listName">The name of list that will be created in site collection.</param>
        /// <param name="templateID">A 32-bit integer that specifies the list template to use.</param>
        /// <param name="baseUrl">The site URL for connecting with the specified Document Workspace Site.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully, 
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        [MethodHelp("Add the specified list (listName) to the server." +
            " The return value should be a Boolean value that indicates whether" +
            " the operation was run successfully. TRUE means the operation was run successfully," +
            " FALSE means the operation failed.")]
        bool AddList(string listName, int templateID, string baseUrl);

        /// <summary>
        /// Delete the specified list in the base site.
        /// </summary>
        /// <param name="listName">The name of list which will be deleted.</param>
        /// <param name="baseUrl">The site URL for connecting with the specified Document Workspace Site.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully, 
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        [MethodHelp("Delete the specified list from the server. The return value should be a Boolean value that" +
            " indicates whether the operation was" +
            " run successfully. TRUE means the operation was run successfully, FALSE means the operation failed.")]
        bool DeleteList(string listName, string baseUrl);
    }
}