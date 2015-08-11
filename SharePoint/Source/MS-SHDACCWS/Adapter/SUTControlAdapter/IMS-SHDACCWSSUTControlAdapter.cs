namespace Microsoft.Protocols.TestSuites.MS_SHDACCWS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUTControlAdapter's implementation.
    /// </summary>
    public interface IMS_SHDACCWSSUTControlAdapter : IAdapter
    {
        #region Interact with ListsService

        /// <summary>
        /// Set the Co-authoring status for the specified file under the specified Document LibraryName list. 
        /// The specified file is identified by the property "FileIdOfCoAuthoring".
        /// </summary>
        /// <returns>True if the operation success, otherwise false.</returns>
        [MethodHelp("Set the specified co-authoring status for the specified file which is identified by the property \"FileIdOfCoAuthoring\". Enter \"TRUE\" if the operation succeeds; otherwise, enter \"FALSE\".")]
        bool SUTSetCoAuthoringStatus();

        /// <summary>
        /// Set status of exclusive lock to the specified file which is identified by the property "FileIdOfLock".
        /// </summary>
        /// <returns>True if the operation success, otherwise false.</returns>
        [MethodHelp("Set the specified status of the exclusive lock to the specified file which is identified by the property \"FileIdOfLock\". Enter \"TRUE\" if the operation succeeds; otherwise, enter \"FALSE\".")]
        bool SUTSetExclusiveLock();

        #endregion
    }
}