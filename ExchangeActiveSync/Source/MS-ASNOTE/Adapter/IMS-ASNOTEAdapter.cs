//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASNOTE
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-ASNOTE.
    /// </summary>
    public interface IMS_ASNOTEAdapter : IAdapter
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
        /// Sync data from the server
        /// </summary>
        /// <param name="syncRequest">Sync command request.</param>
        /// <param name="isResyncNeeded">A bool value indicates whether need to re-sync when the response contains MoreAvailable element.</param>
        /// <returns>The sync result which is returned from server</returns>
        SyncStore Sync(SyncRequest syncRequest, bool isResyncNeeded);

        /// <summary>
        /// Loop to get the results of the specific query request by Search command.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the folder to search.</param>
        /// <param name="subject">The subject of the note to get.</param>
        /// <param name="isLoopNeeded">A boolean value specify whether need the loop</param>
        /// <param name="expectedCount">The expected number of the note to be found.</param>
        /// <returns>The results in response of Search command</returns>
        SearchStore Search(string collectionId, string subject, bool isLoopNeeded, int expectedCount);

        /// <summary>
        /// Fetch all information about exchange object
        /// </summary>
        /// <param name="itemOperationsRequest">ItemOperations command request.</param>
        /// <returns>The ItemOperations result which is returned from server</returns>
        ItemOperationsStore ItemOperations(ItemOperationsRequest itemOperationsRequest);

        /// <summary>
        /// Synchronizes the collection hierarchy
        /// </summary>
        /// <param name="folderSyncRequest">FolderSync command request.</param>
        /// <returns>The FolderSync response which is returned from the server</returns>
        FolderSyncResponse FolderSync(FolderSyncRequest folderSyncRequest);
    }
}