//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using System.Collections.Generic;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-OXWSSYNC.
    /// </summary>
    public interface IMS_OXWSSYNCAdapter : IAdapter
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
        /// Gets synchronization information that enables folders to be synchronized between a client and a server
        /// </summary>
        /// <param name="request">A request to the SyncFolderHierarchy operation</param>
        /// <returns>A response from the SyncFolderHierarchy operation</returns>
        SyncFolderHierarchyResponseType SyncFolderHierarchy(SyncFolderHierarchyType request);

        /// <summary>
        /// Gets synchronization information that enables items to be synchronized between a client and a server
        /// </summary>
        /// <param name="request">A request to the SyncFolderItems operation</param>
        /// <returns>A response from the SyncFolderItems operation</returns>
        SyncFolderItemsResponseType SyncFolderItems(SyncFolderItemsType request);

        /// <summary>
        /// Configures the SOAP header to cover the case that the header contains all optional parts before calling operations.
        /// </summary>
        /// <param name="headerValues">Specify the values for SOAP header.</param>
        void ConfigureSOAPHeader(Dictionary<string, object> headerValues);
    }
}