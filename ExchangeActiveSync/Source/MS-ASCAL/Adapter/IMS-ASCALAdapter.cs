//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASCAL
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using ItemOperationsStore = Microsoft.Protocols.TestSuites.Common.DataStructures.ItemOperationsStore;
    using SearchStore = Microsoft.Protocols.TestSuites.Common.DataStructures.SearchStore;
    using SyncStore = Microsoft.Protocols.TestSuites.Common.DataStructures.SyncStore;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-ASCAL
    /// </summary>
    public interface IMS_ASCALAdapter : IAdapter
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
        /// Sync calendars from the server
        /// </summary>
        /// <param name="syncRequest">The request for Sync command</param>
        /// <returns>The Sync response which is returned from server</returns>
        SyncStore Sync(SyncRequest syncRequest);

        /// <summary>
        /// Search calendars using the given keyword text
        /// </summary>
        /// <param name="searchRequest">The request for Search command</param>
        /// <returns>The search data returned from the server</returns>
        SearchStore Search(SearchRequest searchRequest);

        /// <summary>
        /// Fetch all the information about calendars using longIds or ServerIds
        /// </summary>
        /// <param name="itemOperationsRequest">The request for ItemOperations command</param>
        /// <returns>The fetch items information</returns>
        ItemOperationsStore ItemOperations(ItemOperationsRequest itemOperationsRequest);

        /// <summary>
        /// FolderSync command to synchronize the collection hierarchy 
        /// </summary>
        /// <returns>The FolderSync response</returns>
        FolderSyncResponse FolderSync();

        /// <summary>
        /// Send MIME-formatted e-mail messages to the server
        /// </summary>
        /// <param name="sendMailRequest">The request for SendMail command</param>
        /// <returns>The SendMail response which is returned from the server</returns>
        SendMailResponse SendMail(SendMailRequest sendMailRequest);

        /// <summary>
        /// MeetingResponse for accepting or declining a meeting request
        /// </summary>
        /// <param name="meetingResponseRequest">The request for MeetingResponse</param>
        /// <returns>The MeetingResponse response which is returned from server</returns>
        MeetingResponseResponse MeetingResponse(MeetingResponseRequest meetingResponseRequest);

        /// <summary>
        /// Send a Sync command string request and get Sync response from server.
        /// </summary>
        /// <param name="stringRequest">The request for Sync command</param>
        /// <returns>The Sync response which is returned from server</returns>
        SendStringResponse SendStringRequest(string stringRequest);

        /// <summary>
        /// Change user to call ActiveSync command
        /// </summary>
        /// <param name="userName">The name of the user.</param>
        /// <param name="userPassword">The password of the user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        void SwitchUser(string userName, string userPassword, string userDomain);
    }
}