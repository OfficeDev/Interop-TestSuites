namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
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
        /// Copy one folder into another one.
        /// </summary>
        /// <param name="request">Request of CopyFolder operation.</param>
        /// <returns>Response of CopyFolder operation.</returns>
        CopyFolderResponseType CopyFolder(CopyFolderType request);

        /// <summary>
        /// Create a new folder within a specific folder.
        /// </summary>
        /// <param name="request">Request of CreateFolder operation.</param>
        /// <returns>Response of CreateFolder operation.</returns>
        CreateFolderResponseType CreateFolder(CreateFolderType request);

        /// <summary>
        /// Create a managed folder in server, which should be added in mailbox in advance by server administrator.
        /// </summary>
        /// <param name="request">Request of CreateManagedFolder operation.</param>
        /// <returns>Response of CreateManagedFolder operation.</returns>
        CreateManagedFolderResponseType CreateManagedFolder(CreateManagedFolderRequestType request);

        /// <summary>
        /// Delete a folder from mailbox.
        /// </summary>
        /// <param name="request">Request of DeleteFolder operation.</param>
        /// <returns>Response of DeleteFolder operation.</returns>
        DeleteFolderResponseType DeleteFolder(DeleteFolderType request);

        /// <summary>
        /// Empty identified folders and can be used to delete the subfolders of the specified folder.
        /// </summary>
        /// <param name="request">Request of EmptyFolder operation.</param>
        /// <returns>Response of EmptyFolder operation.</returns>
        EmptyFolderResponseType EmptyFolder(EmptyFolderType request);

        /// <summary>
        /// Get folders, Calendar folders, Contacts folders, Tasks folders, and search folders.
        /// </summary>
        /// <param name="request">Request of GetFolder operation.</param>
        /// <returns>Response of GetFolder operation.</returns>
        GetFolderResponseType GetFolder(GetFolderType request);

        /// <summary>
        /// Move folders from a specified parent folder to another parent folder.
        /// </summary>
        /// <param name="request">Request of MoveFolder operation.</param>
        /// <returns>Response of MoveFolder operation.</returns>
        MoveFolderResponseType MoveFolder(MoveFolderType request);

        /// <summary>
        /// Modify properties of an existing folder in the server store.
        /// </summary>
        /// <param name="request">Request of UpdateFolder operation.</param>
        /// <returns>Response of UpdateFolder operation.</returns>
        UpdateFolderResponseType UpdateFolder(UpdateFolderType request);

        /// <summary>
        /// Switch the current user to a new one, with the identity of the new user to communicate with Exchange Web Service.
        /// </summary>
        /// <param name="userName">The name of a user.</param>
        /// <param name="userPassword">The password of a user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        void SwitchUser(string userName, string userPassword, string userDomain);

        /// <summary>
        /// Configure the SOAP header before calling operations.
        /// </summary>
        /// <param name="headerValues">Specify the values for SOAP header.</param>
        void ConfigureSOAPHeader(Dictionary<string, object> headerValues);
    }
}