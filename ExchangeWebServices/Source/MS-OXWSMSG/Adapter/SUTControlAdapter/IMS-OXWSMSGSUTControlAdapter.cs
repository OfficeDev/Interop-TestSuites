namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OXWSMSG SUT control adapter interface.
    /// </summary>
    public interface IMS_OXWSMSGSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Clean up all items in the Calendar, Inbox, Deleted Items, Drafts and Sent Items folders, which contain a specified subject.
        /// </summary>
        /// <param name="userName">The name of the user used to communicate with server</param>
        /// <param name="password">The password of the user used to communicate with server</param>
        /// <param name="domain">The domain of the user used to communicate with server</param>
        /// <param name="subject">Subject of the item to be removed.</param>
        /// <param name="folders">The folders to be cleaned up, which are delimited by ';'.</param>
        /// <returns>If succeed, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName, userPassword, userDomain) and delete the items with the subject (subject) in (folders) folder/folders. Folder names are delimited with the ';' character." +
            "If the operation succeeds, enter \"true\";" + 
            "otherwise, enter \"false\".")]
        bool CleanupFolders(string userName, string password, string domain, string subject, string folders);
    }
}