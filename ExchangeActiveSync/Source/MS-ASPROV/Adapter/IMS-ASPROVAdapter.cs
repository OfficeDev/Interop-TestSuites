//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASPROV
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-ASPROV.
    /// </summary>
    public interface IMS_ASPROVAdapter : IAdapter
    {
        /// <summary>
        /// Gets the XML request sent to protocol SUT.
        /// </summary>
        IXPathNavigable LastRawRequestXml { get; }

        /// <summary>
        /// Gets the XML response received from protocol SUT.
        /// </summary>
        IXPathNavigable LastRawResponseXml { get; }

        /// <summary>
        /// Change the user authentication.
        /// </summary>
        /// <param name="userName">The name of a user.</param>
        /// <param name="userPassword">The password of a user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        void SwitchUser(string userName, string userPassword, string userDomain);

        /// <summary>
        /// Apply the specified PolicyKey.
        /// </summary>
        /// <param name="appliedPolicyKey">The policy key to apply.</param>
        void ApplyPolicyKey(string appliedPolicyKey);

        /// <summary>
        /// Apply the specified DeviceType.
        /// </summary>
        /// <param name="appliedDeviceType">The device type to apply.</param>
        void ApplyDeviceType(string appliedDeviceType);

        /// <summary>
        /// Request the security policy settings that the administrator sets from the server.
        /// </summary>
        /// <param name="provisionRequest">The request of Provision command.</param>
        /// <returns>The response of Provision command.</returns>
        ProvisionResponse Provision(ProvisionRequest provisionRequest);

        /// <summary>
        /// Synchronizes the changes in a collection between the client and the server by sending SyncRequest object.
        /// </summary>
        /// <param name="syncRequest">A SyncRequest object that contains the request information.</param>
        /// <returns>A SyncStore object.</returns>
        SyncStore Sync(SyncRequest syncRequest);

        /// <summary>
        /// Find an email with specific subject.
        /// </summary>
        /// <param name="syncRequest">A SyncRequest object that contains the request information.</param>
        /// <param name="subject">The subject of the email to find.</param>
        /// <param name="isRetryNeeded">A boolean indicating whether need retry.</param>
        /// <returns>The email with specific subject.</returns>
        Sync SyncEmail(SyncRequest syncRequest, string subject, bool isRetryNeeded);

        /// <summary>
        /// Synchronizes the collection hierarchy from server.
        /// </summary>
        /// <param name="folderSyncRequest">The request of FolderSync command.</param>
        /// <returns>The response of FolderSync command.</returns>
        FolderSyncResponse FolderSync(FolderSyncRequest folderSyncRequest);

        /// <summary>
        /// Send string request of Provision command to the server and get the response.
        /// </summary>
        /// <param name="provisionRequest">The string request of Provision command.</param>
        /// <returns>The response of Provision command.</returns>
        ProvisionResponse SendProvisionStringRequest(string provisionRequest);
    }
}