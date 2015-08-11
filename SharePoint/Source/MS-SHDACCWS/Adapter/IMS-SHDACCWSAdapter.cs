namespace Microsoft.Protocols.TestSuites.MS_SHDACCWS
{
    using System;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This adapter interface definition of MS-SHDACCWS 
    /// </summary>
    public interface IMS_SHDACCWSAdapter : IAdapter
    {
        #region Interact with versionsService

        /// <summary>
        /// Specifies whether a co-authoring transition request was made for a document.
        /// </summary>
        /// <param name="id">The identifier(Guid) of the document in the server.</param>
        /// <returns>Whether a co-authoring transition request was made for a document.</returns>
        bool IsOnlyClient(Guid id);

        #endregion
    }
}