namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using System.Collections.Generic;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-OXWSFOLD.
    /// </summary>
    public interface IMS_OXWSFOLDAdapter : IAdapter
    {
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT.
        /// </summary>
        IXPathNavigable LastRawRequestXml { get; }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT.
        /// </summary>
        IXPathNavigable LastRawResponseXml { get; }

        /// <summary>
        /// Create a new folder within a specific folder.
        /// </summary>
        /// <param name="request">Request of CreateFolder operation.</param>
        /// <returns>Response of CreateFolder operation.</returns>
        CreateFolderResponseType CreateFolder(CreateFolderType request);

        /// <summary>
        /// Delete a folder from mailbox.
        /// </summary>
        /// <param name="request">Request of DeleteFolder operation.</param>
        /// <returns>Response of DeleteFolder operation.</returns>
        DeleteFolderResponseType DeleteFolder(DeleteFolderType request);

        /// <summary>
        /// Get folders, Calendar folders, Contacts folders, Tasks folders, and search folders.
        /// </summary>
        /// <param name="request">Request of GetFolder operation.</param>
        /// <returns>Response of GetFolder operation.</returns>
        GetFolderResponseType GetFolder(GetFolderType request);

        /// <summary>
        /// Switch the current user to a new one, with the identity of the new user to communicate with Exchange Web Service.
        /// </summary>
        /// <param name="userName">The name of a user.</param>
        /// <param name="userPassword">The password of a user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        void SwitchUser(string userName, string userPassword, string userDomain);
    }
}