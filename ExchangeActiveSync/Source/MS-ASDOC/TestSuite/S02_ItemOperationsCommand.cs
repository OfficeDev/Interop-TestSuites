namespace Microsoft.Protocols.TestSuites.MS_ASDOC
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.Request;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to retrieve data from the server for one or more individual documents through ItemOperations command messages.
    /// </summary>
    [TestClass]
    public class S02_ItemOperationsCommand : TestSuiteBase
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
        /// This test case is designed to verify server behavior when using ItemOperations to retrieve document data without multipart header.
        /// </summary>
        [TestCategory("MSASDOC"), TestMethod()]
        public void MSASDOC_S02_TC01_ItemOperations_WithoutMultipart()
        {
            #region Client calls ItemOperations command to get data of a shared visible document.

            // Build element Fetch.
            ItemOperationsFetch fetchElement = this.FetchElementBuilder(Common.GetConfigurationPropertyValue("SharedVisibleDocument", Site));

            // Create an ItemOperation command request.
            ItemOperationsRequest itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { fetchElement });

            // Send request and get response.
            ItemOperationsResponse itemOperationResponse = this.ASDOCAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Site.Assert.AreEqual<int>(1, itemOperationResponse.ResponseData.Response.Fetch.Length, "Only one document's information should be returned.");
            Site.Assert.AreEqual<string>("1", itemOperationResponse.ResponseData.Response.Fetch[0].Status, "The call should be successful.");

            #endregion

            // Multipart responses have a Content-Type header with the value "application/vnd.ms-sync.multipart". 
            bool isMultiPartResponse = itemOperationResponse.Headers.ToString().Contains("application/vnd.ms-sync.multipart");

            bool existDataElement = false;

            foreach (Response.ItemsChoiceType3 itemsElementName in itemOperationResponse.ResponseData.Response.Fetch[0].Properties.ItemsElementName)
            {
                if (itemsElementName == Response.ItemsChoiceType3.Data)
                {
                    existDataElement = true;
                    break;
                }
            }

            Site.Assert.IsTrue(existDataElement, "itemoperations:Data element should be returned");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R69, header value is:{0}", itemOperationResponse.Headers.ToString());

            // Verify MS-ASDOC requirement: MS-ASDOC_R69  
            // If itemoperations:Data element is successfully returned, the body has been decoded from base64 encoding.
            Site.CaptureRequirement(
                69,
                @"[In Requesting the Document Body from the Server] The body of the document is returned as [either] text encoded with base64 encoding in the itemoperations:Data element ([MS-ASCMD] section 2.2.3.39.2) of the ItemOperations command response [or as binary data], depending on the content type that the client requested.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R77, header value is:{0}", itemOperationResponse.Headers.ToString());

            // Verify MS-ASDOC requirement: MS-ASDOC_R77
            Site.CaptureRequirementIfIsFalse(
                isMultiPartResponse,
                77,
                @"[In ItemOperations Command Request] If this header[the HTTP header MS-ASAcceptMultiPart: T] is not used, the document data is returned as text.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R98, header value is:{0}", itemOperationResponse.Headers.ToString());

            // Verify MS-ASDOC requirement: MS-ASDOC_R98
            // If itemoperations:Data element is successfully returned, the body has been decoded from base64 encoding.
            Site.CaptureRequirementIfIsFalse(
                isMultiPartResponse,
                98,
                @"[In Retrieving the Document Body] The server responds with an ItemOperations command response ([MS-ASCMD] section 2.2.1.10), which returns the body of the document as text encoded with base64 encoding in the itemoperations:Data element ([MS-ASCMD] section 2.2.3.39.2) of the response if the command request was not a multi-part request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R107, header value is:{0}", itemOperationResponse.Headers.ToString());

            // Verify MS-ASDOC requirement: MS-ASDOC_R107
            // If itemoperations:Data element is successfully returned, the body has been decoded from base64 encoding.
            Site.CaptureRequirementIfIsFalse(
                isMultiPartResponse,
                107,
                @"[In ItemOperations Command Response] Otherwise[an ItemOperations command request for the body of the document was not made using the MS-ASAcceptMultiPart: T header ], the server MUST transmit the document as data encoded with base64 encoding within the itemoperations:Data element ([MS-ASCMD] section 2.2.3.39.2) of the ItemOperations command response.");
        }

        /// <summary>
        /// This test case is designed to verify server behavior when using ItemOperations to retrieve data of zero or more documents.
        /// </summary>
        [TestCategory("MSASDOC"), TestMethod()]
        public void MSASDOC_S02_TC02_ItemOperations_GetZeroOrMoreDocumentClass()
        {
            #region Client calls ItemOperations command to get data of a nonexistent document under a shared folder.

            // Build element Fetch.
            ItemOperationsFetch fetchZeroDoc = this.FetchElementBuilder(Common.GetConfigurationPropertyValue("SharedVisibleDocument", Site) + Guid.NewGuid().ToString());

            // Create an ItemOperation command request that zero document class value will be returned.
            ItemOperationsRequest itemOperationsGetZeroDocRequest = Common.CreateItemOperationsRequest(new object[] { fetchZeroDoc });

            // Send request and get response.
            ItemOperationsResponse itemOperationGetZeroDocResponse = this.ASDOCAdapter.ItemOperations(itemOperationsGetZeroDocRequest, DeliveryMethodForFetch.Inline);

            // Make sure this file doesn't exist.
            Site.Assert.AreEqual<int>(1, itemOperationGetZeroDocResponse.ResponseData.Response.Fetch.Length, "Only one fetch element should be returned.");
            Site.Assert.AreEqual<string>("6", itemOperationGetZeroDocResponse.ResponseData.Response.Fetch[0].Status, "Document information should not be returned.");

            #endregion

            #region Client calls ItemOperations command to get data of a shared visible document and a shared hidden document.

            // Build element Fetch.
            ItemOperationsFetch fetch1 = this.FetchElementBuilder(Common.GetConfigurationPropertyValue("SharedHiddenDocument", Site));
            ItemOperationsFetch fetch2 = this.FetchElementBuilder(Common.GetConfigurationPropertyValue("SharedVisibleDocument", Site));

            ItemOperationsFetch[] fetches = new ItemOperationsFetch[] { fetch1, fetch2 };

            // Create an ItemOperation command request that multiple document class value will be returned.
            ItemOperationsRequest itemOperationGetMoreDocRequest = Common.CreateItemOperationsRequest(fetches);

            // Send request and get response.
            ItemOperationsResponse itemOperationGetMoreDocResponse = this.ASDOCAdapter.ItemOperations(itemOperationGetMoreDocRequest, DeliveryMethodForFetch.Inline);

            // Information of the two documents is returned.
            Site.Assert.AreEqual<int>(2, itemOperationGetMoreDocResponse.ResponseData.Response.Fetch.Length, "Two document class items should be returned.");
            Site.Assert.AreEqual<string>("1", itemOperationGetMoreDocResponse.ResponseData.Response.Fetch[0].Status, "Document information should be returned.");
            Site.Assert.AreEqual<string>("1", itemOperationGetMoreDocResponse.ResponseData.Response.Fetch[1].Status, "Document information should be returned");

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R102");

            // Items that match the criteria are returned.
            Site.CaptureRequirement(
                102,
                @"[In ItemOperations Command Response] The server MUST return a Document class XML block for every item that matches the criteria specified in the client command request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R103");

            // Server can return zero or more documents class blocks which can be seen from two steps above.
            Site.CaptureRequirement(
                103,
                @"[In ItemOperations Command Response] The server can return zero or more Document class blocks in its response, depending on how many document items match the criteria specified in the client command request.");
        }

        /// <summary>
        /// This test case is designed to verify server behavior when using ItemOperations to retrieve document data with multipart header.
        /// </summary>
        [TestCategory("MSASDOC"), TestMethod()]
        public void MSASDOC_S02_TC03_ItemOperations_WithMultipart()
        {
            #region Client calls ItemOperations command to get data of a shared visible document.

            // Build element Fetch.
            ItemOperationsFetch fetch = this.FetchElementBuilder(Common.GetConfigurationPropertyValue("SharedVisibleDocument", Site));

            // Create an ItemOperation command request.
            ItemOperationsRequest itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { fetch });

            // Send request and get response.
            ItemOperationsResponse itemOperationResponse = this.ASDOCAdapter.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.MultiPart);
            Site.Assert.AreEqual<int>(1, itemOperationResponse.ResponseData.Response.Fetch.Length, "Only one document's information should be returned.");
            Site.Assert.AreEqual<string>("1", itemOperationResponse.ResponseData.Response.Fetch[0].Status, "The call should be successful.");

            #endregion

            // Multipart responses have a Content-Type header with the value "application/vnd.ms-sync.multipart". 
            bool isMultiPartResponse = itemOperationResponse.Headers.ToString().Contains("application/vnd.ms-sync.multipart");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R124, header value is:{0}", itemOperationResponse.Headers.ToString());

            // Verify MS-ASDOC requirement: MS-ASDOC_R124 
            // If header contains "application/vnd.ms-sync.multipart" the body is encoded as binary data.
            Site.CaptureRequirementIfIsTrue(
                isMultiPartResponse,
                124,
                @"[In Requesting the Document Body from the Server] The body of the document is returned as [either text encoded with base64 encoding in the itemoperations:Data element ([MS-ASCMD] section 2.2.3.39.2) of the ItemOperations command response or as] binary data, depending on the content type that the client requested.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R76, header value is:{0}", itemOperationResponse.Headers.ToString());

            // Verify MS-ASDOC requirement: MS-ASDOC_R76            
            Site.CaptureRequirementIfIsTrue(
                isMultiPartResponse,
                76,
                @"[In ItemOperations Command Request] A client can use the HTTP header MS-ASAcceptMultiPart: T to specify that the server returns the document data in multipart binary format.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R120, header value is:{0}", itemOperationResponse.Headers.ToString());

            // Verify MS-ASDOC requirement: MS-ASDOC_R120
            Site.CaptureRequirementIfIsTrue(
                isMultiPartResponse,
                120,
                @"[In Retrieving the Document Body] The server responds with an ItemOperations command response ([MS-ASCMD] section 2.2.1.10), which returns the body of the document as binary text in multiple parts if the command request was a multi-part request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R106, header value is:{0}", itemOperationResponse.Headers.ToString());

            // Verify MS-ASDOC requirement: MS-ASDOC_R106
            Site.CaptureRequirementIfIsTrue(
                isMultiPartResponse,
                106,
                @"[In ItemOperations Command Response] If an ItemOperations command request for the body of the document was made using the MS-ASAcceptMultiPart: T header, then the server MUST respond by providing the document body as binary data in multiple parts.");
        }

        /// <summary>
        /// Builds a Fetch element for the ItemOperations request.
        /// </summary>
        /// <param name="linkId">A URL of the item to be fetched.</param>
        /// <returns>A Fetch element of the ItemOperations request.</returns>
        private ItemOperationsFetch FetchElementBuilder(string linkId)
        {
            // Initialize a fetch element.
            ItemOperationsFetch fetch = new ItemOperationsFetch
            {
                LinkId = linkId,
                Store = SearchName.DocumentLibrary.ToString(),
                Options = new ItemOperationsFetchOptions { ItemsElementName = new ItemsChoiceType5[2] }
            };

            // The user account is used to search the document from the document library
            fetch.Options.ItemsElementName[0] = ItemsChoiceType5.UserName;
            fetch.Options.ItemsElementName[1] = ItemsChoiceType5.Password;
            fetch.Options.Items = new string[2];
            fetch.Options.Items[0] = Common.GetConfigurationPropertyValue("UserName", this.Site);
            fetch.Options.Items[1] = Common.GetConfigurationPropertyValue("UserPassword", this.Site);

            return fetch;
        }
    }
}