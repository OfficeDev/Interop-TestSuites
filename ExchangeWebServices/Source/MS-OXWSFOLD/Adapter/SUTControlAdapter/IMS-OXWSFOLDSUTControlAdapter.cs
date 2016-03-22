namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OXWSFOLD SUT control adapter interface
    /// </summary>
    public interface IMS_OXWSFOLDSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Sets managed folder's storeQuota value.
        /// </summary>
        /// <param name="managedFolderName">The managedFolder name of the user.</param>
        [MethodHelp("Sets managed folder's storeQuota value with the managed folder name (managedFolderName).")]
        string SetManagedFolderStoreQuota(string managedFolderName);

        /// <summary>
        /// Do not set managed folder's storeQuota value.
        /// </summary>
        /// <param name="ManagedFolderName">The managedFolder name of the user.</param>
        [MethodHelp("Do not set managed folder's storeQuota value with the managed folder name (managedFolderName).")]
        string DoNotSetManagedFolderStoreQuota(string managedFolderName);
    }
}