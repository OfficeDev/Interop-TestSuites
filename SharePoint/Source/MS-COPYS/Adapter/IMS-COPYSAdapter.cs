//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This interface definition of MS-COPYS protocol operations. 
    /// </summary>
    public interface IMS_COPYSAdapter : IAdapter
    {
        #region Interact with MS-COPYS service

        /// <summary>
        /// Switch the current credentials of the protocol adapter by specified user. After perform this method, all protocol operations will be performed by specified user.
        /// </summary>
        /// <param name="userName">A parameter represents the user name.</param>
        /// <param name="password">A parameter represents the password of the user.</param>
        /// <param name="domain">A parameter represents the domain of the user.</param>
        void SwitchUser(string userName, string password, string domain);

        /// <summary>
        /// Switch the target service location. The adapter will send the MS-COPYS message to specified service location.
        /// </summary>
        /// <param name="serviceLocation">A parameter represents the service location which host the MS-COPYS service.</param>
        void SwitchTargetServiceLocation(ServiceLocation serviceLocation);

        /// <summary>
        /// A method is used to copy a file when the destination of the operation is on the same protocol server as the source location.
        /// </summary>
        /// <param name="sourceUrl">A parameter represents the location of the file in the source location.</param>
        /// <param name="destinationUrls">A parameter represents a collection of destination location. The operation will try to copy files to that locations.</param>
        /// <returns>A return value represents the result of the operation. It includes status of the copy operation for a destination location.</returns>
        CopyIntoItemsLocalResponse CopyIntoItemsLocal(string sourceUrl, string[] destinationUrls);

        /// <summary>
        /// A method is used to retrieve the contents and metadata for a file from the specified location.
        /// </summary>
        /// <param name="url">A parameter represents the location of the file.</param>
        /// <returns>A return value represents the file contents and metadata.</returns>
        GetItemResponse GetItem(string url);

        /// <summary>
        /// A method used to copy a file to a destination server that is different from the source location.
        /// </summary>
        /// <param name="sourceUrl">A parameter represents the absolute IRI of the file in the source location.</param>
        /// <param name="destinationUrls">A parameter represents a collection of locations on the destination server.</param>
        /// <param name="fields">A parameter represents a collection of the metadata for the file.</param>
        /// <param name="rawStreamValue">A parameter represents the contents of the file. The contents will be encoded in Base64 format and sent in request.</param>
        /// <returns>A return value represents the result of the operation.</returns>
        CopyIntoItemsResponse CopyIntoItems(string sourceUrl, string[] destinationUrls, FieldInformation[] fields, byte[] rawStreamValue);

        #endregion
    }
}