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
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The MS-WOPI SUT Control Adapter interface.
    /// </summary>
    public interface IMS_WOPISUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Add a file into the specified Document Library list on the SUT.
        /// </summary>
        /// <param name="documentLibraryName">The name of the Document Library list where the file will be added.</param>
        /// <param name="fileName">The name of the file.</param>
        /// <returns>A return value represents the absolute URL of the file on the specified Document library list if succeed, otherwise return null.</returns>
        [MethodHelp("Add a file (fileName) with any content into the document library (documentLibraryName). Enter the absolute URL of the file.")]
        string AddFileToSUT(string documentLibraryName, string fileName);

        /// <summary>
        /// Delete the uploaded files whose URLs are been specified.
        /// </summary>
        /// <param name="currentDoclibraryListName">A parameter represents the list name which is used as search condition to get the list id.</param>
        /// <param name="uploadedfilesUrls">A parameter represents a string which contains all URLs for the uploaded files, separated by ",".</param>
        /// <returns>Returns True indicating Cleanup uploaded files was successful</returns>
         [MethodHelp("Delete the files (uploadedfilesUrls) from the document library (currentDoclibraryListName). The files specified in uploadedfilesUrls are delimited with by ','. If the operation succeeds, enter 'true'; otherwise, enter 'false'.")]
        bool DeleteUploadedFilesOnSUT(string currentDoclibraryListName, string uploadedfilesUrls);
 
        /// <summary>
        /// Trigger a WOPI discovery action between the test client and the WOPI server, the test client will act as WOPI client. The WOPI server should receive the correct response from the test client and record a mapping between the WOPI server and the test client.
        /// </summary>
        /// <param name="testClientName">The name of current test client which is running the test suite.</param>
        /// <returns>Returns true indicating this operation is successful, the WOPI server has known the test client as a valid WOPI client.</returns>
        [MethodHelp("Trigger the WOPI server to discover the WOPI client (testClientName) through the WOPI Discovery operation. If the operation succeeds, enter 'true'; otherwise, enter 'false'.")]
        bool TriggerWOPIDiscovery(string testClientName);

        /// <summary>
        /// Remove a WOPI discovery records history between the specified test client and the WOPI server, the test client will act as WOPI client. After this operation executing successfully, the WOPI server will send a new discovery request to the test client, when trigger the WOPI server to discovery, such as calling SUTTriggerWOPIDiscovery method. 
        /// </summary>
        /// <param name="testClientName">The current test client name.</param>
        /// <returns>Returns true indicating this operation is successful. The WOPI server will send a new discovery request to the test client, when trigger the WOPI server to discovery, such as calling SUTTriggerWOPIDiscovery method.</returns>
        [MethodHelp("Remove the Discovery record for the WOPI client (testClientName) from the WOPI server. If the operation succeeds, enter 'true'; otherwise, enter 'false'.")]
        bool RemoveWOPIDiscoveryRecord(string testClientName);
    }
}
