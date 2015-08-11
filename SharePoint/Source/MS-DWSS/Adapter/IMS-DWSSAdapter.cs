namespace Microsoft.Protocols.TestSuites.MS_DWSS
{
    using System.Net;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-DWSS Adapter class.
    /// </summary>
    public interface IMS_DWSSAdapter : IAdapter
    {
        #region Properties

        /// <summary>
        /// Gets or sets the base URL of the Document Workspace Soap Service the client is requesting.
        /// </summary>
        string ServiceUrl { get; set; }

        /// <summary>
        /// Gets or sets the security credentials for Document Workspace Soap Service client authentication.
        /// </summary>
        ICredentials Credentials { get; set; }

        #endregion

        #region DWSS WSDL Operations

        /// <summary>
        /// The operation to determine whether an authenticated user has permission to create a Document Workspace at the specified URL.
        /// </summary>
        /// <param name="dwsUrl">Site-relative URL that specifies where to create the Document Workspace.</param>
        /// <param name="error">An error indication.</param>
        /// <returns>A site-relative URL that specifies where the Document Workspace is created.</returns>
        string CanCreateDwsUrl(string dwsUrl, out Error error);

        /// <summary>
        /// The operation to create a new Document Workspace.
        /// </summary>
        /// <param name="dwsName">Specifies the name of the Document Workspace site, this parameter can be empty.</param>
        /// <param name="users">Specifies the users to be added as contributors in the Document Workspace site, this parameter can be null.</param>
        /// <param name="dwsTitle">Specifies the title of the workspace, this parameter can be empty.</param>
        /// <param name="docs">Specifies information to be stored as a key-value pair in the site metadata, this parameter can be null.</param>
        /// <param name="error">An error indication.</param>
        /// <returns>CreateDws operation response.</returns>
        CreateDwsResultResults CreateDws(string dwsName, UsersItem users, string dwsTitle, DocumentsItem docs, out Error error);

        /// <summary>
        /// The operation to create a folder in the document library of the current Document Workspace site.
        /// </summary>
        /// <param name="folderUrl">Site-relative URL with the full path for the new folder.</param>
        /// <param name="error">An error indication.</param>
        void CreateFolder(string folderUrl, out Error error);

        /// <summary>
        /// The operation to delete a Document Workspace from the protocol server.
        /// </summary>
        /// <param name="error">An error indication.</param>
        void DeleteDws(out Error error);

        /// <summary>
        /// The operation to delete a folder from a document library on the site.
        /// </summary>
        /// <param name="folderUrl">Site-relative URL specifying the folder to delete.</param>
        /// <param name="error">An error indication.</param>
        void DeleteFolder(string folderUrl, out Error error);

        /// <summary>
        /// The operation to obtain a URL for a named document in a Document Workspace.
        /// </summary>
        /// <param name="docId">A unique string that represents a document key.</param>
        /// <param name="error">An error indication.</param>
        /// <returns>An absolute URL that refers to the requested document if the call is successful.</returns>
        string FindDwsDoc(string docId, out Error error);

        /// <summary>
        /// The operation to return general information about the Document Workspace site, as well as its members, documents, links, and tasks.
        /// </summary>
        /// <param name="docUrl">A site-based URL of a document in the document library in the Document Workspace.</param>
        /// <param name="lastUpdate">Contains the lastUpdate value returned in the result of a previous GetDwsData or GetDwsMetaData operation, or an empty string.</param>
        /// <param name="error">An error indication.</param>
        /// <returns>GetDwsData operation response.</returns>
        Results GetDwsData(string docUrl, string lastUpdate, out Error error);

        /// <summary>
        /// The operation to return information about a Document Workspace site and the lists that it contains.
        /// </summary>
        /// <param name="docUrl">A site-relative URL that specifies the list or document to describe in the response.</param>
        /// <param name="docId">A unique string that represents a document key.</param>
        /// <param name="isMinimal">A Boolean value that specifies whether to return information.</param>
        /// <param name="error">An error indication.</param>
        /// <returns>GetDwsMetaData operation response.</returns>
        GetDwsMetaDataResultTypeResults GetDwsMetaData(string docUrl, string docId, bool isMinimal, out Error error);

        /// <summary>
        /// The operation to delete a user from a Document Workspace.
        /// </summary>
        /// <param name="userId">The user identifier of the user to remove from the workspace. This positive integer MUST be in the range from zero through 2,147,483,647, inclusive.</param>
        /// <param name="error">An error indication.</param>
        void RemoveDwsUser(int userId, out Error error);

        /// <summary>
        /// The operation to change the title of a Document Workspace.
        /// </summary>
        /// <param name="dwsTitle">A string contains the new title of the workspace.</param>
        /// <param name="error">An error indication.</param>
        void RenameDws(string dwsTitle, out Error error);

        /// <summary>
        /// The operation to modify the metadata of a Document Workspace. This method is deprecated and should not be called by the protocol client.
        /// </summary>
        /// <param name="updates">A string that contains CAML instructions specifying how to update the workspace information.</param>
        /// <param name="meetingInstance">A string that contains the meeting information, this parameter can be empty.</param>
        /// <param name="error">An error indication.</param>
        /// <returns>UpdateDwsData operation response.</returns>
        string UpdateDwsData(string updates, string meetingInstance, out Error error);

        #endregion
    }
}