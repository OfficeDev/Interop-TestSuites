//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System.Net;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface definition of the MS-WOPI adapter definition.
    /// </summary>
    public interface IMS_WOPIAdapter : IAdapter
    {
        /// <summary>
        /// This method is used to take a lock for editing a file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="lockIdentifierValue">A parameter represents the value which is provided by WOPI client that WOPI server used to identify the lock on the file.</param>
        /// <returns>A return value represents the http response for Lock operation.</returns>
        WOPIHttpResponse Lock(string targetResourceUrl, WebHeaderCollection commonHeaders, string lockIdentifierValue);

        /// <summary>
        /// This method is used to update a file on the WOPI server.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiSize">A parameter represents the size of the request body.</param>
        /// <param name="bodyContents">A parameter represents the body contents of the request.</param>
        /// <param name="lockIdentifierOfFile">A parameter represents the value which is provided by WOPI client that WOPI server used to identify the lock on the file.</param>
        /// <returns>A return value represents the http response for PutFile operation.</returns>
        WOPIHttpResponse PutFile(string targetResourceUrl, WebHeaderCollection commonHeaders, int? xwopiSize, byte[] bodyContents, string lockIdentifierOfFile);

        /// <summary>
        /// This method is used to return information about folder and permissions for the user which is determined by "targetResourceUrl" parameter and the "Authorization" header in "commonHeaders" parameter.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiSessionContextValue">A parameter represents the value of the session context information.</param>
        /// <returns>A return value represents the http response for CheckFolderInfo operation.</returns>
        WOPIHttpResponse CheckFolderInfo(string targetResourceUrl, WebHeaderCollection commonHeaders, string xwopiSessionContextValue);

        /// <summary>
        /// This method is used to return the file information including file properties and permissions for the user who is identified by the token that is sent in the request.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiSessionContextValue">A parameter represents the value of the session context information.</param>
        /// <returns>A return value represents the http response for CheckFileInfo operation.</returns>
        WOPIHttpResponse CheckFileInfo(string targetResourceUrl, WebHeaderCollection commonHeaders, string xwopiSessionContextValue);

        /// <summary>
        /// This method is used to get a file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="maxExpectedSize">A parameter represents the specifying upper bound size of the file being requested.</param>
        /// <returns>A return value represents the http response for GetFile operation.</returns>
        WOPIHttpResponse GetFile(string targetResourceUrl, WebHeaderCollection commonHeaders, int? maxExpectedSize);

        /// <summary>
        /// This method is used to release a lock for editing a file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="lockIdentifierOfFile">A parameter represents the value which is provided by WOPI client that WOPI server used to identify the lock on the file.</param>
        /// <returns>A return value represents the http response for UnLock operation.</returns>
        WOPIHttpResponse UnLock(string targetResourceUrl, WebHeaderCollection commonHeaders, string lockIdentifierOfFile);

        /// <summary>
        /// This method is used to return the contents of a folder on the WOPI server.
        /// </summary>
        /// <param name="targetResourceUrlOfFloder">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <returns>A return value represents the http response for EnumerateChildren operation.</returns>
        WOPIHttpResponse EnumerateChildren(string targetResourceUrlOfFloder, WebHeaderCollection commonHeaders);

        /// <summary>
        /// This method is used to create a file on the WOPI server based on the current file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiSuggestedTarget">A parameter represents the file name in order to create a file.</param>
        /// <param name="xwopiRelativeTarget">A parameter represents the file name of the current file.</param>
        /// <param name="bodyContents">A parameter represents the body contents of the request.</param>
        /// <param name="xwopiOverwriteRelativeTarget">A parameter represents the value that specifies whether the host overwrite the file name if it exists.</param>
        /// <param name="xwopiSize">A parameter represents the size of the file.</param>
        /// <returns>A return value represents the http response for PutRelativeFile operation.</returns>
        WOPIHttpResponse PutRelativeFile(string targetResourceUrl, WebHeaderCollection commonHeaders, string xwopiSuggestedTarget, string xwopiRelativeTarget, byte[] bodyContents, bool? xwopiOverwriteRelativeTarget, int? xwopiSize);

        /// <summary>
        /// This method is used to refresh the existing lock for modifying a file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="lockIdentifierOfRefreshLock">A parameter represents the value which is provided by WOPI client that WOPI server used to identify the lock on the file.</param>
        /// <returns>A return value represents the http response for RefreshLock operation.</returns>
        WOPIHttpResponse RefreshLock(string targetResourceUrl, WebHeaderCollection commonHeaders, string lockIdentifierOfRefreshLock);

        /// <summary>
        /// This method is used to release and retake a lock for editing a file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="lockIdentifierValue">A parameter represents the value which is provided by WOPI client that WOPI server used to identify the lock on the file.</param>
        /// <param name="lockIdentifierOldValue">A parameter represents the value which is previously provided by WOPI client that WOPI server used to identify the lock on the file.</param>
        /// <returns>A return value represents the http response for UnlockAndRelock operation.</returns>
        WOPIHttpResponse UnlockAndRelock(string targetResourceUrl, WebHeaderCollection commonHeaders, string lockIdentifierValue, string lockIdentifierOldValue);

        /// <summary>
        /// This method is used to delete a file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <returns>A return value represents the http response for DeleteFile operation.</returns>
        WOPIHttpResponse DeleteFile(string targetResourceUrl, WebHeaderCollection commonHeaders);

        /// <summary>
        /// This method is used to get a link to a file though which a user is able to operate on a file in a limited way.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiRestrictedLink">A parameter represents the type of restricted link being request by WOPI client.</param>
        /// <returns>A return value represents the http response for GetRestrictedLink operation.</returns>
        WOPIHttpResponse GetRestrictedLink(string targetResourceUrl, WebHeaderCollection commonHeaders, string xwopiRestrictedLink);

        /// <summary>
        /// This method is used to revoke all links to a file through which a number of users are able to operate on a file in a limited way.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiRestrictedLink">A parameter represents the type of restricted link being revoked by WOPI client.</param>
        /// <returns>A return value represents the http response for RevokeRestrictedLink operation.</returns>
        WOPIHttpResponse RevokeRestrictedLink(string targetResourceUrl, WebHeaderCollection commonHeaders, string xwopiRestrictedLink);

        /// <summary>
        /// This method is used to access the WOPI server's implementation of a secure store.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiApplicationId">A parameter represents the value of application ID.</param>
        /// <returns>A return value represents the http response for ReadSecureStore operation.</returns>
        WOPIHttpResponse ReadSecureStore(string targetResourceUrl, WebHeaderCollection commonHeaders, string xwopiApplicationId);
    }
}
