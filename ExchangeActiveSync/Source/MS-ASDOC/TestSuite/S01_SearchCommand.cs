//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASDOC
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to retrieve Document class items that match the criteria specified by the client through the Search command messages.
    /// </summary>
    [TestClass]
    public class S01_SearchCommand : TestSuiteBase
    {
        #region Class initialize and cleanup

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        /// <summary>
        /// This test case is designed to verify server behavior when using Search command to search a visible folder with LinkId.
        /// </summary>
        [TestCategory("MSASDOC"), TestMethod()]
        public void MSASDOC_S01_TC01_SearchCommand_VisibleFolderWithLinkId()
        {
            #region Client calls Search command to get document class value of a shared visible folder.

            // Get search command response.
            SearchResponse searchResponse = this.SearchCommand(Common.GetConfigurationPropertyValue("SharedVisibleFolder", Site));
            Site.Assert.AreEqual<string>("1", searchResponse.ResponseData.Response.Store.Status, "The folder should be found.");
            Site.Assert.AreEqual<int>(1, searchResponse.ResponseData.Response.Store.Result.Length, "Only one folder information should be returned.");

            #endregion

            for (int i = 0; i < searchResponse.ResponseData.Response.Store.Result[0].Properties.ItemsElementName.Length; i++)
            {
                if (searchResponse.ResponseData.Response.Store.Result[0].Properties.ItemsElementName[i] == Response.ItemsChoiceType6.IsFolder)
                {
                    byte isFolder = (byte)searchResponse.ResponseData.Response.Store.Result[0].Properties.Items[i];

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R47");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R47
                    Site.CaptureRequirementIfAreEqual<byte>(
                        1,
                        isFolder,
                        47,
                        @"[In IsFolder] The value 1 means the item is  a folder.");
                }

                if (searchResponse.ResponseData.Response.Store.Result[0].Properties.ItemsElementName[i] == Response.ItemsChoiceType6.IsHidden)
                {
                    byte isHidden = (byte)searchResponse.ResponseData.Response.Store.Result[0].Properties.Items[i];

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R118");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R118
                    Site.CaptureRequirementIfAreEqual<byte>(
                        0,
                        isHidden,
                        118,
                        @"[In IsHidden]The value 0 means the folder is not hidden.");
                }
            }
        }

        /// <summary>
        /// This test case is designed to verify server behavior when using Search command to search a hidden folder with LinkId.
        /// </summary>
        [TestCategory("MSASDOC"), TestMethod()]
        public void MSASDOC_S01_TC02_SearchCommand_HiddenFolderWithLinkId()
        {
            #region Client calls Search command to get document class value of a shared hidden folder.

            // Get Search command response.
            SearchResponse searchResponse = this.SearchCommand(Common.GetConfigurationPropertyValue("SharedHiddenFolder", Site));
            Site.Assert.AreEqual<string>("1", searchResponse.ResponseData.Response.Store.Status, "The folder should be found.");
            Site.Assert.AreEqual<int>(1, searchResponse.ResponseData.Response.Store.Result.Length, "Only one folder information should be returned.");

            #endregion

            for (int i = 0; i < searchResponse.ResponseData.Response.Store.Result[0].Properties.ItemsElementName.Length; i++)
            {
                if (searchResponse.ResponseData.Response.Store.Result[0].Properties.ItemsElementName[i] == Response.ItemsChoiceType6.IsHidden)
                {
                    byte isHidden = (byte)searchResponse.ResponseData.Response.Store.Result[0].Properties.Items[i];

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R119");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R119
                    Site.CaptureRequirementIfAreEqual<byte>(
                        1,
                        isHidden,
                        119,
                        @"[In IsHidden]The value 1 means the folder  is hidden.");
                }
            }
        }

        /// <summary>
        /// This test case is designed to verify server behavior when using Search command to retrieve data of a document without LinkId.
        /// </summary>
        [TestCategory("MSASDOC"), TestMethod()]
        public void MSASDOC_S01_TC03_SearchCommand_WithoutLinkId()
        {
            #region Client calls Search command without LinkId.

            // Get Search command response.
            SearchResponse searchResponse = this.SearchCommand(null);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R85");

            // Verify MS-ASDOC requirement: MS-ASDOC_R85
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                searchResponse.ResponseData.Response.Store.Status,
                85,
                @"[In Search Command Request] If the LinkId element is not included in a Search command request, then the server MUST respond with protocol error 2.");
        }

        /// <summary>
        /// This test case is designed to verify server behavior when using Search command to retrieve data of zero or more documents.
        /// </summary>
        [TestCategory("MSASDOC"), TestMethod()]
        public void MSASDOC_S01_TC04_SearchCommand_GetZeroOrMoreDocumentClass()
        {
            #region Client calls Search command to get document class value of a folder that doesn't exist.

            // Get Search command response.
            SearchResponse searchGetZeroDocResponse = this.SearchCommand(Common.GetConfigurationPropertyValue("SharedFolder", Site) + Guid.NewGuid().ToString());
            Site.Assert.IsNull(searchGetZeroDocResponse.ResponseData.Response.Store.Result, "Document class value should not be returned.");

            #endregion

            #region Client calls Search command to get document class value of a shared folder which is the root folder.

            // Get Search command response.
            SearchResponse searchGetMultipleDocResponse = this.SearchCommand(Common.GetConfigurationPropertyValue("SharedFolder", Site));
            Site.Assert.AreEqual<string>("1", searchGetMultipleDocResponse.ResponseData.Response.Store.Status, "The folder should be found.");

            // According to MS-ASCMD section 2.2.3.142.2 :If the documentlibrary:LinkId element value in the request points to a folder, the metadata properties of the folder are returned as the first item, and the contents of the folder are returned as subsequent results.
            Site.Assert.IsTrue(searchGetMultipleDocResponse.ResponseData.Response.Store.Result.Length == 5, "Root folder and the contents of the folder should be returned.");

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R112");

            // Server can return zero or more documents class blocks which can be seen from two steps above.
            Site.CaptureRequirement(
                112,
                @"[In Search Command Response] The server can return zero or more Document class blocks in its response, depending on how many document items match the criteria specified in the client command request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R111");

            // Items that match the criteria are returned.
            Site.CaptureRequirement(
                111,
                @"[In Search Command Response] The server MUST return a Document class XML block for every item that matches the criteria specified in the client command request.");
        }

        /// <summary>
        /// This test case is designed to verify server behavior when using Search command to search a visible document.
        /// </summary>
        [TestCategory("MSASDOC"), TestMethod()]
        public void MSASDOC_S01_TC05_SearchCommand_VisibleDocument()
        {
            #region Client calls Search command to get document class value of a shared visible document.

            // Get Search command response.
            SearchResponse searchResponse = this.SearchCommand(Common.GetConfigurationPropertyValue("SharedVisibleDocument", Site));
            Site.Assert.AreEqual<string>("1", searchResponse.ResponseData.Response.Store.Status, "Document class value should be returned.");
            Site.Assert.AreEqual<int>(1, searchResponse.ResponseData.Response.Store.Result.Length, "Only one document information should be returned.");

            #endregion

            for (int i = 0; i < searchResponse.ResponseData.Response.Store.Result[0].Properties.ItemsElementName.Length; i++)
            {
                if (searchResponse.ResponseData.Response.Store.Result[0].Properties.ItemsElementName[i] == Response.ItemsChoiceType6.IsFolder)
                {
                    byte isFolder = (byte)searchResponse.ResponseData.Response.Store.Result[0].Properties.Items[i];

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_46");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R46
                    Site.CaptureRequirementIfAreEqual<byte>(
                        0,
                        isFolder,
                        46,
                        @"[In IsFolder] The value 0 means the item is not a folder.");
                }

                if (searchResponse.ResponseData.Response.Store.Result[0].Properties.ItemsElementName[i] == Response.ItemsChoiceType6.IsHidden)
                {
                    byte isHidden = (byte)searchResponse.ResponseData.Response.Store.Result[0].Properties.Items[i];

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R52");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R52
                    Site.CaptureRequirementIfAreEqual<byte>(
                        0,
                        isHidden,
                        52,
                        @"[In IsHidden]The value 0 means the document  is not hidden.");
                }
            }
        }

        /// <summary>
        /// This test case is designed to verify server behavior when using Search command to search a Hidden document.
        /// </summary>
        [TestCategory("MSASDOC"), TestMethod()]
        public void MSASDOC_S01_TC06_SearchCommand_HiddenDocument()
        {
            #region Client calls Search command to get document class value of a shared hidden document.

            // Get Search command response.
            SearchResponse searchResponse = this.SearchCommand(Common.GetConfigurationPropertyValue("SharedHiddenDocument", Site));
            Site.Assert.AreEqual<string>("1", searchResponse.ResponseData.Response.Store.Status, "Document class value should be returned.");
            Site.Assert.AreEqual<int>(1, searchResponse.ResponseData.Response.Store.Result.Length, "Only one document information should be returned.");

            #endregion

            for (int i = 0; i < searchResponse.ResponseData.Response.Store.Result[0].Properties.ItemsElementName.Length; i++)
            {
                if (searchResponse.ResponseData.Response.Store.Result[0].Properties.ItemsElementName[i] == Response.ItemsChoiceType6.IsHidden)
                {
                    byte isHidden = (byte)searchResponse.ResponseData.Response.Store.Result[0].Properties.Items[i];

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R53");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R53
                    Site.CaptureRequirementIfAreEqual<byte>(
                        1,
                        isHidden,
                        53,
                        @"[In IsHidden]The value 1 means the document  is hidden.");
                }
            }
        }
    }
}