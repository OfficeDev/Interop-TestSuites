namespace Microsoft.Protocols.TestSuites.MS_LISTSWS
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the following operations with valid or invalid parameters.
    /// <list type="bullet">
    ///     <item>GetAttachmentCollection</item>
    ///     <item>AddAttachment</item>  
    ///     <item>DeleteAttachment</item>
    /// </list>
    /// </summary>
    [TestClass]
    public class S04_OperationOnAttachment : TestClassBase
    {
        #region Private member variables

        /// <summary>
        /// Protocol adapter
        /// </summary>
        private IMS_LISTSWSAdapter listswsAdapterInstance;

        #endregion

        #region Additional test attributes

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">A parameter represents a testContext.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test cases

        #region AddAttachment

        /// <summary>
        /// This test case is used to verify that the server returns error code "0x8007000d"
        /// when adding an attachment to a list item that content length equals to zero in AddAttachment operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC01_AddAttachment_ContentLengthZero()
        {
            #region Create a new list

            bool createListSucceeded = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            createListSucceeded = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(createListSucceeded, "MSLISTSWS_S04_TC01_AddAttachment_ContentLenghZero, created a new list successfully.");
            #endregion

            #region Add a list item

            List<string> listItemIds = TestSuiteHelper.AddListItems(listGuid, 1);
            Site.Assert.IsNotNull(listItemIds, "Called AddListItems successfully.");

            #endregion

            #region Call the AddAttachment with the content length 0

            bool addAttachmentFail = false;
            string errorCode = null;
            try
            {
                // Call the AddAttachment with the content length 0
                string fileName = Common.GetConfigurationPropertyValue("AttachmentName", this.Site);
                byte[] attachment = new byte[0];
                this.listswsAdapterInstance.AddAttachment(listName, listItemIds[0], fileName, attachment);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                addAttachmentFail = true;
            }

            Site.Assert.IsTrue(addAttachmentFail, "MSLISTSWS_S04_TC01_AddAttachment_ContentLenghZero, calling AddAttachmet with content length 0 fails.");

            // Capture R299 if the error code 0x8007000d is returned.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x8007000d",
                errorCode,
                299,
                @"[In AddAttachment] If the protocol client requests to add an attachment of content length zero, the protocol server MUST return a SOAP fault with error code 0x8007000d.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test AddAttachment operation when at least one of its 
        /// input parameters is invalid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC02_AddAttachment_InvalidParameters()
        {
            #region Add a list and add a list item
            string uniqueListName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(uniqueListName);

            // add one item, if not successful or added number not equal to the specified, will fail the test case.
            List<string> listitemIds = TestSuiteHelper.AddListItems(listId, 1);
            #endregion

            // Add an attachment using valid GUID as listName parameter.
            string attachmentFileName = Common.GetConfigurationPropertyValue("AttachmentName", this.Site);
            byte[] attachmentData = TestSuiteHelper.GetAttachmentContent(attachmentFileName);
            string addedListitemId = listitemIds[0];

            bool isSoapFaultReturned = false;
            string errorCode = string.Empty;
            try
            {
                string invalidGuidAndAndNocorrespondListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
                this.listswsAdapterInstance.AddAttachment(
                                                    invalidGuidAndAndNocorrespondListName,
                                                    addedListitemId,
                                                    attachmentFileName,
                                                    attachmentData);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // if the error code 0x82000006 is returned, Capture R2750 
            if (Common.IsRequirementEnabled(2750, this.Site))
            {
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual values: isSoapFaultReturned[{0}],errorCode[{1}] for requirement #R2750",
                    isSoapFaultReturned,
                    !string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

                Site.CaptureRequirementIfIsTrue(
                    isSoapFaultReturned && !string.IsNullOrEmpty(errorCode)
                    && errorCode.Equals("0x82000006", StringComparison.OrdinalIgnoreCase),
                    2750,
                    @"[In AddAttachment]Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks.(SharePoint Foundation 2010 and above follow this behavior.)");
            }

            // Try to add an attachment with listItemID that does not correspond to a ListItem.
            isSoapFaultReturned = false;
            errorCode = string.Empty;
            try
            {
                // Current List only added one list item in this step.
                int notcorrespondListItemId = int.MaxValue;
                this.listswsAdapterInstance.AddAttachment(
                                                    listId,
                                                    notcorrespondListItemId.ToString(),
                                                    attachmentFileName,
                                                    attachmentData);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // If a SOAP fault without error code is returned. capture R296 and R1551 
            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: isSoapFaultReturned[{0}],errorCode[{1}] for requirement #R296 and #R1551",
                    isSoapFaultReturned,
                    string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturned && string.IsNullOrEmpty(errorCode),
                296,
                @"[In AddAttachment] If the listItemID does not correspond to a list item in a specified list on the site, the protocol server MUST return a SOAP fault.");

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturned && string.IsNullOrEmpty(errorCode),
                1551,
                @"[In AddAttachment] [If the listItemID does not correspond to a list item in a specified list on the site, the protocol server MUST return a SOAP fault.]There is no error code returned for this fault.");

            // Add an attachment.
            string attachmentUrl = this.listswsAdapterInstance.AddAttachment(
                                                listId,
                                                addedListitemId,
                                                attachmentFileName,
                                                attachmentData);
            this.Site.Assert.IsTrue(
                                   attachmentUrl.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) > 0,
                                   "Add an attachment into the list item should be successful.");

            // Try to add an attachment with the same fileName into same list item.
            isSoapFaultReturned = false;
            errorCode = string.Empty;
            try
            {
                string duplicateFileName = attachmentFileName;
                this.listswsAdapterInstance.AddAttachment(
                                                listId,
                                                addedListitemId,
                                                duplicateFileName,
                                                attachmentData);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // If an error code 0x81020067 is returned.Capture R297 and R1552 
            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: isSoapFaultReturned[{0}],errorCode[{1}] for requirement #R297 and #R1552",
                    isSoapFaultReturned,
                    errorCode.Equals("0x81020067") ? "0x81020067" : errorCode);

            Site.CaptureRequirementIfIsTrue(
               isSoapFaultReturned && !string.IsNullOrEmpty(errorCode)
               && errorCode.Equals("0x81020067", StringComparison.OrdinalIgnoreCase),
               297,
                @"[In AddAttachment] If the specified fileName is not unique for the specified list item's attachment collection then the protocol server MUST return a SOAP fault with error code 0x81020067.");

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturned && !string.IsNullOrEmpty(errorCode)
                && errorCode.Equals("0x81020067", StringComparison.OrdinalIgnoreCase),
                1552,
                @"[In AddAttachment] [If the specified fileName is not unique for the specified list item's attachment collection then the protocol server MUST return a SOAP fault with error code 0x81020067.] This indicates that the specified fileName is already in use.");

            // Try to add an attachment without listItemID parameter.
            isSoapFaultReturned = false;
            errorCode = string.Empty;
            try
            {
                string emptyListItemId = string.Empty;
                this.listswsAdapterInstance.AddAttachment(
                                                listId,
                                                emptyListItemId,
                                                attachmentFileName,
                                                attachmentData);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // Capture R300 if an error code 0x82000001 is returned.
            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: isSoapFaultReturned[{0}],errorCode[{1}] for requirement #R300",
                    isSoapFaultReturned,
                    errorCode.Equals("0x82000001") ? "0x82000001" : errorCode);

            Site.CaptureRequirementIfIsTrue(
                  isSoapFaultReturned && !string.IsNullOrEmpty(errorCode)
                  && errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                  300,
                  @"[In AddAttachment] If the protocol client request excludes the listItemID parameter, the protocol server MUST return a SOAP fault with error code 0x82000001.");

            // Add second list item on the current list, and now add attachment operation target to the second ListItem.
            List<string> secondAddedListitemIds = TestSuiteHelper.AddListItems(listId, 1);
            string secondListitemId = secondAddedListitemIds[0];
            isSoapFaultReturned = false;
            errorCode = string.Empty;
            try
            {
                this.listswsAdapterInstance.AddAttachment(
                                                listId,
                                                secondListitemId,
                                                attachmentFileName,
                                                null);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // Capture R1554 if a SOAP fault without error code is returned.
            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: isSoapFaultReturned[{0}],errorCode[{1}] for requirement #R1554",
                    isSoapFaultReturned,
                    string.IsNullOrEmpty(errorCode) ? "IsNullOrEmpty" : errorCode);

            Site.CaptureRequirementIfIsTrue(
                    isSoapFaultReturned && string.IsNullOrEmpty(errorCode),
                    1554,
                    @"[In AddAttachment] Otherwise [if the protocol client request includes the listItemID parameter and], if the listName or attachment parameters are missing, the protocol server MUST return a SOAP fault with no error code.");
        }

        /// <summary>
        /// This test case is used to verify that the server returns error code "0x81020073" with FileName parameter is Null in AddAttachment operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC03_AddAttachment_NullFileName()
        {
            #region Create a new list

            bool createListSucceeded = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            createListSucceeded = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(createListSucceeded, "MSLISTSWS_S04_TC03_AddAttachment_NullFileName, created a new list successfully.");
            #endregion

            #region Add a list item
            List<string> listItemIds = TestSuiteHelper.AddListItems(listGuid, 1);
            Site.Assert.IsNotNull(listItemIds, "Called AddListItems successfully.");

            #endregion

            #region AddAttachment with the file name is null

            bool addAttachmentFail = false;
            string errorCode = null;
            try
            {
                // Call the AddAttachment with the file name is null.
                string fileName = Common.GetConfigurationPropertyValue("AttachmentName", this.Site);
                byte[] attachment = TestSuiteHelper.GetAttachmentContent(fileName);
                this.listswsAdapterInstance.AddAttachment(listName, listItemIds[0], null, attachment);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                addAttachmentFail = true;
            }

            Site.Assert.IsTrue(addAttachmentFail, "MSLISTSWS_S04_TC03_AddAttachment_NullFileName, call AddAttachmet with file name is null.");

            // Capture R3011 if the error code 0x81020073 is returned.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x81020073",
                errorCode,
                3011,
                @"[In AddAttachment] Otherwise [if the protocol client request includes the listItemID, listName and attachment parameters], if the fileName parameter is missing, the protocol server MUST return a SOAP fault with error code 0x81020073.");

            #endregion
        }

        /// <summary>
        /// This test case is used to validate AddAttachment operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in the current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC04_AddAttachment_SP3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2426, this.Site), @"Test is executed only when R2426Enabled is set to true.");

            // Add a list
            #region Add a list
            string guid = TestSuiteHelper.CreateList();
            string expectederrorString = "Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).";
            #endregion

            #region Initialize property value

            // Add one List item 
            List<string> ids = TestSuiteHelper.AddListItems(guid, 1);

            #endregion

            #region Invoke AddAttachment and capture requirements
            string errorCode = string.Empty;
            string errorString = string.Empty;
            bool isSoapFaultGenerate = false;
            try
            {
                // Add an attachment using invalid GUID as listName parameter.
                string invalidGuidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
                this.listswsAdapterInstance.AddAttachment(invalidGuidListName, ids[0], null, null);
            }
            catch (SoapException exp)
            {
                isSoapFaultGenerate = true;
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                errorString = TestSuiteHelper.GetErrorString(exp);
            }

            this.Site.Assert.IsTrue(isSoapFaultGenerate, "There should be a soap fault generated for AddAttachment operation, when use invalid GUID.");
            this.Site.Assert.IsTrue(string.IsNullOrEmpty(errorCode), "The Soap exception contains error Code.");

            this.Site.Log.Add(LogEntryKind.Debug, "The actual error string:[{0}] \r\n for requirement #R2426", errorString);

            Site.Log.Add(
                LogEntryKind.Debug,
                "The actual value: errorString[{0}] for requirement #R2426",
                errorString);

            Site.CaptureRequirementIfIsTrue(
                expectederrorString.Equals(errorString, StringComparison.OrdinalIgnoreCase),
                2426,
                @"[In Appendix B: Product Behavior]Implementation does not return a SOAP fault with error code 0x82000006. (<22> Section 3.1.4.1: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");

            #endregion
        }

        /// <summary>
        /// This test case is used to test AddAttachment operation when all its input parameters are valid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC05_AddAttachment_Succeed()
        {
            #region Add a list and add a list item
            string uniqueListName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(uniqueListName);

            // add one item, if not successful or added number not equal to the specified, will fail the test case.
            List<string> listitemIds = TestSuiteHelper.AddListItems(listId, 1);
            #endregion

            // Add an attachment using valid GUID as listName parameter.
            string attachmentFileName = Common.GetConfigurationPropertyValue("AttachmentName", this.Site);
            byte[] attachmentData = TestSuiteHelper.GetAttachmentContent(attachmentFileName);
            string addedListitemId = listitemIds[0];

            // call AddAttachment operation to add an attachment to the added list item.
            string attachmentUrl = this.listswsAdapterInstance.AddAttachment(
                                                    listId,
                                                    addedListitemId,
                                                    attachmentFileName,
                                                    attachmentData);

            // Get attachment to verify whether AddAttachment operation successfully or not.
            GetAttachmentCollectionResponseGetAttachmentCollectionResult attachmentResult = null;
            attachmentResult = this.listswsAdapterInstance.GetAttachmentCollection(listId, addedListitemId);
            if (null == attachmentResult || null == attachmentResult.Attachments || attachmentResult.Attachments.Length == 0)
            {
                this.Site.Assert.Fail("Could Not get the expected added attachment in List, List Id:[{0}]", listId);
            }

            // Get the attachment URL from GetAttachmentCollection response.
            // The URL value is contained in a string array,  which is defined in [MS-LISTSWS].
            // There is at least one text item which contains the "attacgneFileName" in the Text array.
            var attachmentUrlItem = from string urlItem in attachmentResult.Attachments
                                    where urlItem.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) >= 0
                                    select urlItem;

            bool isAttachFileNameContained = attachmentUrlItem.Count() > 0;

            // If the AddAttachment succeeds. Capture R292, R301 and R302 
            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: isAttachFileNameContained[{0}],attachmentUrl[{1}] for requirement #R292 #R301 and #R302",
                    isAttachFileNameContained,
                    attachmentUrl);

            Site.CaptureRequirementIfIsTrue(
                isAttachFileNameContained && attachmentUrl.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) >= 0,
                292,
                @"[AddAttachment]If the specified listName is a valid GUID and corresponds to the identification of a list on the site, use that list.");

            Site.CaptureRequirementIfIsTrue(
                 isAttachFileNameContained && attachmentUrl.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) >= 0,
                 301,
                 @"[In AddAttachment] If the protocol client passes valid input parameters, the protocol server MUST add the attachment to the specified list item in the specified list.");

            Site.CaptureRequirementIfIsTrue(
                isAttachFileNameContained && attachmentUrl.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) >= 0,
                302,
                @"[In AddAttachment operation] If the operation succeeds, an attachment MUST be added to the specified list item in the specified list.");

            // Add a new list item to the list, and use this new added item to call AddAttachment operation.
            listitemIds = TestSuiteHelper.AddListItems(listId, 1);
            string secondaddedListitemId = listitemIds[0];

            // call AddAttachment operation to add an attachment to the added list item.
            attachmentUrl = string.Empty;

            attachmentUrl = this.listswsAdapterInstance.AddAttachment(
                                                uniqueListName,
                                                secondaddedListitemId,
                                                attachmentFileName,
                                                attachmentData);

            // Get attachment to verify whether AddAttachment operation successfully or not
            attachmentResult = null;
            attachmentResult = this.listswsAdapterInstance.GetAttachmentCollection(listId, secondaddedListitemId);
            if (null == attachmentResult || null == attachmentResult.Attachments || attachmentResult.Attachments.Length == 0)
            {
                this.Site.Assert.Fail("Could Not get the expected added attachment in List, List Id:[{0}]", listId);
            }

            // Get the attachment URL from GetAttachmentCollection response.
            // The URL value is contained in a string array,  which is defined in [MS-LISTSWS].
            // There is at least one text item which contain the "attacgneFileName" in the Text array.
            attachmentUrlItem = from string urlItem in attachmentResult.Attachments
                                where urlItem.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) >= 0
                                select urlItem;

            isAttachFileNameContained = attachmentUrlItem.Count() > 0;

            // if the AddAttachment succeeds, Capture R293 and R294 
            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: isAttachFileNameContained[{0}],attachmentUrl[{1}] for requirement #R293 and #R294",
                    isAttachFileNameContained,
                    attachmentUrl);

            Site.CaptureRequirementIfIsTrue(
               isAttachFileNameContained && attachmentUrl.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) >= 0,
                293,
                @"[AddAttachment]If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            Site.CaptureRequirementIfIsTrue(
               isAttachFileNameContained && attachmentUrl.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) >= 0,
                294,
                @"[AddAttachment]If the specified listName [is not a valid GUID or] does not correspond to the identification of a list on the site, check if the listName corresponds to the list title of a list on the site and if so, use that list.");
        }

        #endregion

        #region DeleteAttachment

        /// <summary>
        /// This test case is used to test DeleteAttachment operation when at least one of its input parameters is invalid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC06_DeleteAttachment_InvalidParameters()
        {
            #region Add a list

            string listId = TestSuiteHelper.CreateList();

            #endregion

            #region Add a list item and add attachment into the list item.

            // Add one item, if not successful or added number not equal to the specified, will fail the test case.
            List<string> listitemIds = TestSuiteHelper.AddListItems(listId, 1);
            string attachmentFileName = Common.GetConfigurationPropertyValue("AttachmentName", this.Site);

            // Get the attachment bytes data
            byte[] attachmentData = TestSuiteHelper.GetAttachmentContent(attachmentFileName);
            string addedListitemId = listitemIds[0];

            // call AddAttachment operation to add an attachment to the added list item.
            string attachmentUrl = this.listswsAdapterInstance.AddAttachment(
                                                    listId,
                                                    addedListitemId,
                                                    attachmentFileName,
                                                    attachmentData);
            this.Site.Assert.IsTrue(
                                  attachmentUrl.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) > 0,
                                  "Add an attachment into the list item should be successful.");

            #endregion

            #region Invoke DeleteAttachment and capture requirements

            bool isSoapFaultReturned = false;
            string errorCode = string.Empty;
            try
            {
                string fullUrlOfAttachment = TestSuiteHelper.GetAttachmentFullUrl(listId, addedListitemId, attachmentFileName);
                string invalidGuidAndNocorrespondName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
                this.listswsAdapterInstance.DeleteAttachment(
                                            invalidGuidAndNocorrespondName,
                                            addedListitemId,
                                            fullUrlOfAttachment);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            this.Site.Assert.IsTrue(isSoapFaultReturned, "There is a Soap fault returned for Requirement #1968");

            // Capture R2754 if the error code 0x82000006 is returned.
            if (Common.IsRequirementEnabled(2754, this.Site))
            {
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Returned Error Code [{0}] for requirement R2754",
                    string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

                Site.CaptureRequirementIfIsTrue(
                     errorCode.Equals("0x82000006", StringComparison.OrdinalIgnoreCase),
                     2754,
                     @"[DeleteAttachment]Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks.(SharePoint Foundation 2010 and above follow this behavior.)");
            }

            isSoapFaultReturned = false;
            errorCode = string.Empty;
            try
            {
                string emptyListItemId = string.Empty;
                string fullUrlOfAttachment = TestSuiteHelper.GetAttachmentFullUrl(listId, addedListitemId, attachmentFileName);
                this.listswsAdapterInstance.DeleteAttachment(
                                            listId,
                                            emptyListItemId,
                                            fullUrlOfAttachment);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            this.Site.Assert.IsTrue(isSoapFaultReturned && !string.IsNullOrEmpty(errorCode), "No Soap fault or error code was returned for requirement #R1700.");

            // Capture R1700 and R479 if the error code 0x82000001 is returned.
            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: errorCode[{0}] for requirement #R1700 and #R479",
                    errorCode.Equals("0x82000001") ? "0x82000001" : errorCode);

            Site.CaptureRequirementIfIsTrue(
                 errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                 1700,
                 @"[In DeleteAttachment operation] If the protocol client excludes the listItemID parameter, the protocol server MUST return a SOAP fault with error code 0x82000001.");

            Site.CaptureRequirementIfIsTrue(
               errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                479,
                @"[In DeleteAttachment operation] otherwise [if the operation failed], the protocol server MUST return a SOAP fault with error codes as described in section 3.1.4.10.");

            // Try to delete the attachment with invalid-URL parameter.
            // Current list contains two list items, try to call DeleteAttachment operation with second added ListItem id
            // but the use the first list item's attachment URL.
            List<string> listitemIdsOfSecondAdded = TestSuiteHelper.AddListItems(listId, 1);
            string secondAddedListitemId = listitemIdsOfSecondAdded[0];

            isSoapFaultReturned = false;
            errorCode = string.Empty;
            try
            {
                // Try to use the attachment URL of first added item, and send the target item id is second added.
                string invalidAttachmentPath = TestSuiteHelper.GetAttachmentFullUrl(listId, addedListitemId, attachmentFileName);
                this.listswsAdapterInstance.DeleteAttachment(
                                            listId,
                                            secondAddedListitemId,
                                            invalidAttachmentPath);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            this.Site.Assert.IsTrue(isSoapFaultReturned && !string.IsNullOrEmpty(errorCode), "There is no Soap fault existed or the no ErrorCode return for Requirement #1700");

            // Capture R471 and R1701 if the error code 0x82000001 is returned.
            Site.Log.Add(
                   LogEntryKind.Debug,
                   "The actual value: errorCode[{0}] for requirement #R471 and #R1701",
                   errorCode.Equals("0x82000001") ? "0x82000001" : errorCode);

            Site.CaptureRequirementIfIsTrue(
                 errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                 471,
                 @"[In DeleteAttachment operation] If the specified URL does not correspond to the absolute URL path of an attachment in a specified list item in a specified list on the site, the protocol server MUST return a SOAP fault with error code 0x82000001.");

            Site.CaptureRequirementIfIsTrue(
                 errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                 1701,
                 @"[In DeleteAttachment operation] [If the specified URL does not correspond to the absolute URL path of an attachment in a specified list item in a specified list on the site, the protocol server MUST return a SOAP fault with error code 0x82000001.] This indicates that the URL is missing or is invalid.");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify DeleteAttachment operation when the list ItemId is specified but the list name or URL is excluded.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC07_DeleteAttachment_ListNameOrUrlExclude()
        {
            #region Create a new list

            bool createListSucceeded = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            createListSucceeded = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(createListSucceeded, "MSLISTSWS_S04_TC07_DeleteAttachment_ListNameOrUrlExclude, created a new list successfully.");
            #endregion

            #region Add a list item
            List<string> listItemIds = TestSuiteHelper.AddListItems(listGuid, 1);
            Site.Assert.IsNotNull(listItemIds, "Called AddListItems successfully.");

            #endregion

            #region Add an attachment

            string fileName = Common.GetConfigurationPropertyValue("AttachmentName", this.Site);
            byte[] attachment = TestSuiteHelper.GetAttachmentContent(fileName);
            string addAttachment = this.listswsAdapterInstance.AddAttachment(listName, listItemIds[0], fileName, attachment);
            Site.Assert.IsNotNull(addAttachment, "Called AddAttachment successfully.");

            #endregion

            #region DeleteAttachment when the list name or the URL is null

            // Delete the attachment when the list name is null
            bool listNameError = false;
            try
            {
                this.listswsAdapterInstance.DeleteAttachment(null, listItemIds[0], addAttachment);
            }
            catch (SoapException exp)
            {
                listNameError = TestSuiteHelper.GetErrorCode(exp) == null;
            }

            Site.Assert.IsTrue(listNameError, "Deleting attachment when the list name is null should fail.");

            // Delete the attachment when the URL is null
            bool attachmentUrlError = false;
            try
            {
                this.listswsAdapterInstance.DeleteAttachment(listName, listItemIds[0], null);
            }
            catch (SoapException exp)
            {
                attachmentUrlError = TestSuiteHelper.GetErrorCode(exp) == null;
            }

            Site.Assert.IsTrue(attachmentUrlError, "Deleting attachment when the URL is null should fail.");

            // Delete the attachment when the URL and list name are null
            bool listNameAndAttachmentUrlError = false;
            try
            {
                this.listswsAdapterInstance.DeleteAttachment(null, listItemIds[0], null);
            }
            catch (SoapException exp)
            {
                listNameAndAttachmentUrlError = TestSuiteHelper.GetErrorCode(exp) == null;
            }

            Site.Assert.IsTrue(listNameAndAttachmentUrlError, "Deleting attachment when both of the list name and URL are null should fail.");

            // If listNameError, attachmentUrlError and listNameAndAttachmentUrlError are all true, R3013 will be able to be captured.
            Site.CaptureRequirement(
                3013,
                @"[In DeleteAttachment operation] Otherwise [if the protocol client includes the listItemID parameter], if the protocol client excludes any of the other parameters, the protocol server MUST return a SOAP fault with no error code.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test the DeleteAttachment operation when ListItemID parameter
        /// does not correspond to a list item.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC08_DeleteAttachment_ListItemIDNotCorrespond()
        {
            #region Add a list and add a list item
            string uniqueListName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(uniqueListName);

            // add one item, if not successful or added number not equal to the specified, fail the test case.
            List<string> listitemIds = TestSuiteHelper.AddListItems(listId, 1);
            #endregion

            // Add an attachment using valid GUID as listName parameter.
            string attachmentFileName = Common.GetConfigurationPropertyValue("AttachmentName", this.Site);
            byte[] attachmentData = TestSuiteHelper.GetAttachmentContent(attachmentFileName);
            string addedListitemId = listitemIds[0];

            #region Add an attachment
            // call AddAttachment operation to add an attachment to the added list item.
            string attachmentUrl = this.listswsAdapterInstance.AddAttachment(
                                                    listId,
                                                    addedListitemId,
                                                    attachmentFileName,
                                                    attachmentData);
            this.Site.Assert.IsTrue(
                                  attachmentUrl.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) > 0,
                                  "Adding an attachment into the list item should be successful.");
            #endregion

            #region Invoke DeleteAttachment and capture requirements

            // Delete the attachment with the ListItemID which does not correspond to a list item.
            string fullUrlOfAttachment = TestSuiteHelper.GetAttachmentFullUrl(listId, addedListitemId, attachmentFileName);
            string listItemIDNotCorrespond = (int.Parse(addedListitemId) + 1).ToString();

            try
            {
                this.listswsAdapterInstance.DeleteAttachment(listId, listItemIDNotCorrespond, fullUrlOfAttachment);
                this.Site.Assert.Fail("DeleteAttachment operation must return a SOAP fault if the listItemID parameter does not correspond to a list item.");
            }
            catch (SoapException ex)
            {
                string errorCode = TestSuiteHelper.GetErrorCode(ex);

                // If a SOAP exception is returned and the error code is null, capture R470.
                Site.CaptureRequirementIfIsNull(
                    errorCode,
                    470,
                    @"[In DeleteAttachment operation] If the listItemID does not correspond to a list item in a specified list on the site, the protocol server MUST return a SOAP fault. There is no error code for this fault.");
            }

            #endregion
        }

        /// <summary>
        /// This test case is used to verify DeleteAttachment operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in the current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC09_DeleteAttachment_NonExistentListName_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2440, this.Site), @"Test is executed only when R2440Enabled is set to true.");

            // Add a list
            #region Add a list
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            string guid = TestSuiteHelper.CreateList();
            string expectederrorString = "Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).";
            #endregion

            #region Initialize property value

            // Add one List item
            List<string> ids = TestSuiteHelper.AddListItems(guid, 1);

            #endregion

            #region Invoke AddAttachment and capture requirements
            string errorCode = string.Empty;
            string errorString = string.Empty;
            bool isSoapFaultGenerate = false;
            try
            {
                // Add an attachment using invalid GUID as listName parameter.
                string attchmentFileName = Common.GetConfigurationPropertyValue("AttachmentName", this.Site);
                byte[] attachmentFile = TestSuiteHelper.GetAttachmentContent(attchmentFileName);
                string attachmentUrl = this.listswsAdapterInstance.AddAttachment(guid, ids[0], attchmentFileName, attachmentFile);

                // Delete the attachment with an invalid list name.
                this.listswsAdapterInstance.DeleteAttachment(invalidListName, ids[0], attachmentUrl);
            }
            catch (SoapException exp)
            {
                isSoapFaultGenerate = true;
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                errorString = TestSuiteHelper.GetErrorString(exp);
            }

            this.Site.Assert.IsTrue(isSoapFaultGenerate, "There should be a soap fault generated for AddAttachment operation, when use invalid GUID.");
            this.Site.Assert.IsTrue(string.IsNullOrEmpty(errorCode), "The Soap exception contains error Code.");
            this.Site.Log.Add(LogEntryKind.Debug, "The actual error string:[{0}] \r\n for requirement #R2440", errorString);

            // If the error string equal to the "Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)."
            Site.Log.Add(
                LogEntryKind.Debug,
                "The actual value: errorString[{0}] for requirement #R2440",
                errorString);

            Site.CaptureRequirementIfIsTrue(
                        expectederrorString.IndexOf(errorString, StringComparison.OrdinalIgnoreCase) >= 0,
                        2440,
                        @"[In Appendix B: Product Behavior]Implementation does not return a SOAP fault with error code 0x82000006. (<36> Section 3.1.4.10: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
            #endregion
        }

        /// <summary>
        /// This test case is used to test DeleteAttachment operation when all its input parameters are valid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC10_DeleteAttachment_Succeed()
        {
            #region Add a list and add a list item
            string uniqueListName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(uniqueListName);

            // add one item, if not successful or added number not equal to the specified, fail the test case.
            List<string> listitemIds = TestSuiteHelper.AddListItems(listId, 1);
            #endregion

            // Add an attachment using valid GUID as listName parameter.
            string attachmentFileName = Common.GetConfigurationPropertyValue("AttachmentName", this.Site);
            byte[] attachmentData = TestSuiteHelper.GetAttachmentContent(attachmentFileName);
            string addedListitemId = listitemIds[0];

            #region Add an attachment
            // call AddAttachment operation to add an attachment to the added list item.
            string attachmentUrl = this.listswsAdapterInstance.AddAttachment(
                                                    listId,
                                                    addedListitemId,
                                                    attachmentFileName,
                                                    attachmentData);
            this.Site.Assert.IsTrue(
                                  attachmentUrl.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) > 0,
                                  "Add an attachment into the list item should be successful.");
            #endregion

            #region Invoke DeleteAttachment and capture requirements

            // Delete the attachment with valid GUID as listName parameter.
            string fullUrlOfAttachment = TestSuiteHelper.GetAttachmentFullUrl(listId, addedListitemId, attachmentFileName);
            this.listswsAdapterInstance.DeleteAttachment(listId, addedListitemId, fullUrlOfAttachment);

            // Call GetAttachmentCollection to verify delete successfully or not.
            GetAttachmentCollectionResponseGetAttachmentCollectionResult getAttachmenResult = null;
            getAttachmenResult = this.listswsAdapterInstance.GetAttachmentCollection(listId, addedListitemId);
            if (null == getAttachmenResult || null == getAttachmenResult.Attachments)
            {
                this.Site.Assert.Fail("Could not Get the valid response of GetAttachmentCollection operation");
            }

            // Capture R466, R473 and R478 if the DeleteAttachment succeeds.
            Site.CaptureRequirementIfAreEqual<int>(
                0,
                getAttachmenResult.Attachments.Length,
                466,
                @"[In DeleteAttachment operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, use that list.");

            Site.CaptureRequirementIfAreEqual<int>(
                0,
               getAttachmenResult.Attachments.Length,
                473,
                @"[In DeleteAttachment operation] If the protocol client calls the operation by passing valid input parameters, the protocol server MUST delete the attachment associated with the specified list item in the specified list.");

            Site.CaptureRequirementIfAreEqual<int>(
                0,
                getAttachmenResult.Attachments.Length,
                478,
                @"[In DeleteAttachment operation] If the operation succeeds, the attachment MUST be deleted.");

            // Add an attachment.
            List<string> listitemIdsOfsecond = TestSuiteHelper.AddListItems(listId, 1);
            string secondListitemId = listitemIdsOfsecond[0];
            attachmentUrl = string.Empty;
            attachmentUrl = this.listswsAdapterInstance.AddAttachment(
                                                   listId,
                                                   secondListitemId,
                                                   attachmentFileName,
                                                   attachmentData);
            this.Site.Assert.IsTrue(
                                  !string.IsNullOrEmpty(attachmentUrl)
                                  && attachmentUrl.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) > 0,
                                  "Add an attachment into the list item should be successful.");

            // Delete the attachment with valid list title as listName parameter.
            fullUrlOfAttachment = string.Empty;
            fullUrlOfAttachment = TestSuiteHelper.GetAttachmentFullUrl(listId, secondListitemId, attachmentFileName);
            this.listswsAdapterInstance.DeleteAttachment(uniqueListName, secondListitemId, fullUrlOfAttachment);

            // Call GetAttachmentCollection to verify delete successfully or not.
            getAttachmenResult = null;
            getAttachmenResult = this.listswsAdapterInstance.GetAttachmentCollection(listId, secondListitemId);
            if (null == getAttachmenResult || null == getAttachmenResult.Attachments)
            {
                this.Site.Assert.Fail("Could Not Get the valid response of GetAttachmentCollection operation");
            }

            // Capture R467 and R468 if the DeleteAttachment succeeds.
            Site.CaptureRequirementIfAreEqual<int>(
                0,
               getAttachmenResult.Attachments.Length,
                467,
                @"[In DeleteAttachment operation] If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            Site.CaptureRequirementIfAreEqual<int>(
              0,
             getAttachmenResult.Attachments.Length,
                468,
                @"[In DeleteAttachment operation] If the specified listName does not correspond to the identification of a list on the site, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            #endregion
        }

        #endregion

        #region GetAttachmentCollection

        /// <summary>
        /// This test case is used to test the server behavior when the input parameters are valid but no attachment exists in GetAttachmentCollection operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC11_GetAttachmentCollection_ZeroAttachmentTest()
        {
            #region Add a list and add a list item

            string listId = TestSuiteHelper.CreateList();

            // add one item, if not successful or added number not equal to the specified, will fail the test case.
            List<string> listitemIds = TestSuiteHelper.AddListItems(listId, 1);
            #endregion

            #region Invoke GetAttachmentCollection and capture requirements

            GetAttachmentCollectionResponseGetAttachmentCollectionResult getAttachmentCollectionResult = null;
            bool isSoapFault = false;
            try
            {
                // pick up the added item id, the number is one. 
                getAttachmentCollectionResult = this.listswsAdapterInstance.GetAttachmentCollection(listId, listitemIds[0]);
            }
            catch (SoapException)
            {
                isSoapFault = true;
            }

            #endregion

            #region Requirement Verification

            // If no soap fault, and getAttachmentCollectionResult is not null and the returned 
            // Attachments' length is zero, requirement 543 is captured.
            this.Site.Assert.IsFalse(isSoapFault, "There is unexpected Soap Fault when calling [GetAttachmentCollection] operation.");

            Site.CaptureRequirementIfAreEqual<int>(
                0,
                getAttachmentCollectionResult.Attachments.Length,
                543,
                @"[In GetAttachmentCollection operation] If the protocol client calls the operation on a listItemID containing zero attachments, the protocol server MUST return an empty attachment collection.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test GetAttachmentCollection operation when at least one of parameters are valid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC12_GetAttachmentCollection_InvalidParameters()
        {
            #region Add a list and add a list item
            string uniqueListName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(uniqueListName);

            // add one item, if not successful or added number not equal to the specified, will fail the test case.
            List<string> listitemIds = TestSuiteHelper.AddListItems(listId, 1);
            #endregion

            // Add an attachment using valid GUID as listName parameter.
            string attachmentFileName = Common.GetConfigurationPropertyValue("AttachmentName", this.Site);
            byte[] attachmentData = TestSuiteHelper.GetAttachmentContent(attachmentFileName);
            string addedListitemId = listitemIds[0];

            #region Add an attachment
            // call AddAttachment operation to add an attachment to the added list item.
            string attachmentUrl = this.listswsAdapterInstance.AddAttachment(
                                                    listId,
                                                    addedListitemId,
                                                    attachmentFileName,
                                                    attachmentData);
            this.Site.Assert.IsTrue(
                                  attachmentUrl.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) > 0,
                                  "Add an attachment into the list item should be successful.");
            #endregion

            #region Invoke GetAttachmentCollection and capture requirements

            // Try to get the attachment collection with an invalid listName parameter.
            bool isSoapFaultReturned = false;
            string errorCode = string.Empty;
            try
            {
                // Get a list name that is not a valid GUID and it does not correspond to any existing list
                string invalidGuidAndNocorrespondName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
                this.listswsAdapterInstance.GetAttachmentCollection(invalidGuidAndNocorrespondName, addedListitemId);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // Capture R2758, if the error code 0x82000006 is returned.
            if (Common.IsRequirementEnabled(2758, this.Site))
            {
                Site.Log.Add(
                   LogEntryKind.Debug,
                   "Returned Error Code [{0}] for requirement #R2758",
                   string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

                Site.CaptureRequirementIfIsTrue(
                    isSoapFaultReturned && "0x82000006".Equals(errorCode, StringComparison.OrdinalIgnoreCase),
                    2758,
                    @"[In GetAttachmentCollection operation]Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks.(SharePoint Foundation 2010 and above follow this behavior.) ");
            }

            // Try to get the attachment collection with invalid listItemID parameter.
            isSoapFaultReturned = false;
            errorCode = string.Empty;
            try
            {
                // Get a list item ID that does not correspond to any existed ListItem on current list.
                // Current List only added one item.
                string invalidlistItemID = int.MaxValue.ToString();
                this.listswsAdapterInstance.GetAttachmentCollection(uniqueListName, invalidlistItemID);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // Capture R538, R549 and R1748 if a SOAP fault without error code is returned.
            Site.Log.Add(
                   LogEntryKind.Debug,
                   "The actual value: isSoapFaultReturned[{0}],errorCode[{1}] for requirement #R538, #R549 and #R1748",
                   isSoapFaultReturned,
                   string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturned && string.IsNullOrEmpty(errorCode),
                538,
                @"[In GetAttachmentCollection operation] If the listItemID does not correspond to a list item in a specified list on the site, the protocol server MUST return a SOAP fault.");

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturned && string.IsNullOrEmpty(errorCode),
                1748,
                @"[In GetAttachmentCollection operation] [If the listItemID does not correspond to a list item in a specified list on the site, the protocol server MUST return a SOAP fault.]There is no error code for this fault.");

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturned && string.IsNullOrEmpty(errorCode),
                549,
                @"[In GetAttachmentCollection operation] otherwise [if the operation failed], the protocol server MUST return a SOAP fault with error codes as described in section 3.1.4.14.");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetAttachmentCollection operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in the current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC13_GetAttachmentCollection_NonExistentListName_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2450, this.Site), @"Test is executed only when R2450Enabled is set to true.");

            // Add a list
            #region Add a list
            string invalidListName = "123";
            string guid = TestSuiteHelper.CreateList();
            string expectederrorString = "GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).";
            #endregion

            #region Initialize property value

            // Add one List item
            List<string> ids = TestSuiteHelper.AddListItems(guid, 1);

            #endregion

            #region Invoke AddAttachment and capture requirements
            string errorCode = string.Empty;
            string errorString = string.Empty;
            bool isSoapFaultGenerate = false;
            try
            {
                // Delete the attachment with an invalid listName.
                this.listswsAdapterInstance.GetAttachmentCollection(invalidListName, ids[0]);
            }
            catch (SoapException exp)
            {
                isSoapFaultGenerate = true;
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                errorString = TestSuiteHelper.GetErrorString(exp);
            }

            this.Site.Assert.IsTrue(isSoapFaultGenerate, "There should be a soap fault generated for AddAttachment operation, when use invalid GUID.");
            this.Site.Assert.IsTrue(string.IsNullOrEmpty(errorCode), "The Soap exception contains error Code.");
            this.Site.Log.Add(LogEntryKind.Debug, "The actual error string:[{0}] \r\n for requirement #R2450", errorString);

            // If the error string equal to the "Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)." capture R2450
            Site.Log.Add(
                LogEntryKind.Debug,
                "The actual value: errorString[{0}] for requirement #R2450",
                errorString);

            Site.CaptureRequirementIfIsTrue(
                        expectederrorString.IndexOf(errorString, StringComparison.OrdinalIgnoreCase) >= 0,
                        2450,
                        @"[In Appendix B: Product Behavior]Implementation does not return a SOAP fault with error code 0x82000006. (<42> Section 3.1.4.14: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");

            #endregion
        }

        /// <summary>
        /// The test case is used to verify GetAttachmentCollection operation when list item id is null.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC14_GetAttachmentCollection_NullItemId()
        {
            #region Create a new list

            bool createListSucceeded = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            createListSucceeded = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(createListSucceeded, "MSLISTSWS_S04_TC13_GetAttachmentCollection_NullItemId, created a new list successfully.");
            #endregion

            #region GetAttachmentCollection when the list item ID is null.
            string errorCode = null;
            bool isCaptureR540 = false;
            try
            {
                // Call GetAttachmentCollection when list item id is null.
                this.listswsAdapterInstance.GetAttachmentCollection(listName, null);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isCaptureR540 = true;
            }

            Site.Assert.IsTrue(isCaptureR540, "Calling GetAttachmentCollection with list item id is null fail.");

            // Capture R540 if the error code 0x82000001 is returned.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x82000001",
                errorCode,
                540,
                @"[In GetAttachmentCollection operation] If the protocol client excludes the listItemID parameter, the protocol server MUST return a SOAP fault with error code 0x82000001.");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetAttachmentCollection operation when the list name is null.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC15_GetAttachmentCollection_NullListName()
        {
            #region Create a new list

            bool createListSucceeded = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            createListSucceeded = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(createListSucceeded, "MSLISTSWS_S04_TC14_GetAttachmentCollection_NullListName, created a new list successfully.");
            #endregion

            #region Add a list item
            List<string> listItemIds = TestSuiteHelper.AddListItems(listGuid, 1);
            Site.Assert.IsNotNull(listItemIds, "Called AddListItems successfully.");

            #endregion

            #region Invoke GetAttachmentCollection and capture requirements
            string errorCode = null;
            bool isCaptureR541 = false;
            try
            {
                // Call GetAttachmentCollection without list name.
                this.listswsAdapterInstance.GetAttachmentCollection(null, listItemIds[0]);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isCaptureR541 = errorCode == null;
            }

            // Capture R541 if the protocol server return a SOAP fault with no error code.
            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: isCaptureR541[{0}] for requirement #R541",
                    isCaptureR541);

            Site.CaptureRequirementIfIsTrue(
                isCaptureR541,
                541,
                @"[In GetAttachmentCollection operation] If the protocol client excludes the listName parameter, the protocol server MUST return a SOAP fault with no error code.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test GetAttachmentCollection operation when all its input parameters are valid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC16_GetAttachmentCollection_Succeed()
        {
            #region Add a list

            string uniqueListName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(uniqueListName);

            #endregion

            #region Add a list item and add attachment into the list item.

            // Add one item, if not successful or added number not equal to the specified, will fail the test case.
            List<string> listitemIds = TestSuiteHelper.AddListItems(listId, 1);
            string attachmentFileName = Common.GetConfigurationPropertyValue("AttachmentName", this.Site);

            // Get the attachment bytes data
            byte[] attachmentData = TestSuiteHelper.GetAttachmentContent(attachmentFileName);
            string addedListitemId = listitemIds[0];

            // call AddAttachment operation to add an attachment to the added list item.
            string attachmentUrl = this.listswsAdapterInstance.AddAttachment(
                                                    listId,
                                                    addedListitemId,
                                                    attachmentFileName,
                                                    attachmentData);
            this.Site.Assert.IsTrue(
                                  attachmentUrl.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) > 0,
                                  "Add an attachment into the list item should be successful.");

            #endregion

            #region Invoke GetAttachmentCollection and capture requirements
            // Get the attachment collection with valid GUID as listName parameter.
            bool isSoapFaultReturned = false;
            GetAttachmentCollectionResponseGetAttachmentCollectionResult getAttachmentCollectionResult = null;
            try
            {
                getAttachmentCollectionResult = this.listswsAdapterInstance.GetAttachmentCollection(listId, addedListitemId);
            }
            catch (SoapException)
            {
                isSoapFaultReturned = true;
                string errorMsg = string.Format(
                     "There is unexpected SoapFault thrown when calling GetAttachmentCollection operation with valid list GUID [{0}]",
                     listId);
                this.Site.Log.Add(LogEntryKind.Debug, errorMsg);
            }

            // Capture R534 and R542 if the GetAttachmentCollection succeeds and there is no exception.
            Site.CaptureRequirementIfIsFalse(
                isSoapFaultReturned,
                534,
                @"[In GetAttachmentCollection operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, use that list.");

            Site.CaptureRequirementIfIsFalse(
                isSoapFaultReturned,
                542,
                @"[In GetAttachmentCollection operation] If the protocol client passes valid input parameters, the protocol server MUST return the collection of attachments associated with the list item.");

            // Get the attachment collection with valid list title as listName parameter.
            // Current List use a specified name and record by "uniqueListName"
            isSoapFaultReturned = false;
            try
            {
                this.listswsAdapterInstance.GetAttachmentCollection(uniqueListName, addedListitemId);
            }
            catch (SoapException)
            {
                isSoapFaultReturned = true;
                string errorMsg = string.Format(
                     "There is unexpected SoapFault was thrown when calling GetAttachmentCollection operation with existed list Name [{0}]",
                     uniqueListName);
                this.Site.Log.Add(LogEntryKind.Debug, errorMsg);
            }

            // Capture R535 and R536 if the GetAttachmentCollection succeeds.
            Site.CaptureRequirementIfIsFalse(
                isSoapFaultReturned,
                535,
                @"[In GetAttachmentCollection operation] If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            Site.CaptureRequirementIfIsFalse(
                isSoapFaultReturned,
                536,
                @"[In GetAttachmentCollection operation] If the specified listName does not correspond to the identification of a list on the site, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            // Delete the attachment.
            // Add second list item and without adding any attachments.
            List<string> listIdsOfSecondListItem = TestSuiteHelper.AddListItems(listId, 1);
            string secondListItemId = listIdsOfSecondListItem[0];

            // Get the attachment collection with valid GUID as listName parameter.
            getAttachmentCollectionResult = null;
            isSoapFaultReturned = false;
            try
            {
                getAttachmentCollectionResult = this.listswsAdapterInstance.GetAttachmentCollection(listId, secondListItemId);
            }
            catch (SoapException)
            {
                isSoapFaultReturned = true;
                string errorMsg = string.Format(
                     "There is unexpected SoapFault thrown when calling GetAttachmentCollection operation with valid list GUID [{0}] and valid listitemId[{1}]"
                     + @"to the target ListItem does not contain any attachment",
                     listId,
                     secondListItemId);
                this.Site.Assert.Fail(errorMsg);
            }

            this.Site.Assert.IsTrue(
                    getAttachmentCollectionResult != null && getAttachmentCollectionResult.Attachments != null,
                    "The response of GetAttachmentCollection operation must not be null for Requirement #543.");

            // Capture R543 if the GetAttachmentCollection succeeds and no attachment is returned.
            Site.CaptureRequirementIfAreEqual<int>(
                0,
                getAttachmentCollectionResult.Attachments.Length,
                543,
                @"[In GetAttachmentCollection operation] If the protocol client calls the operation on a listItemID containing zero attachments, the protocol server MUST return an empty attachment collection.");
            #endregion
        }

        /// <summary>
        /// The test case will verify GetAttachmentCollection operation when list item ID is negative.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S04_TC17_GetAttachmentCollection_NegativeListItemID()
        {
            Site.Assume.IsTrue(
                Common.IsRequirementEnabled(30371, this.Site) || Common.IsRequirementEnabled(30372, this.Site),
                @"Test is executed only when R30371Enabled is set to true or R30372Enabled is set to true.");

            #region Create a new list

            bool createListSucceeded = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            createListSucceeded = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(createListSucceeded, "MSLISTSWS_S04_TC16_GetAttachmentCollection_NegativeListItemID, created a new list successfully.");
            #endregion

            #region Invoke GetAttachmentCollection and capture requirements
            string negativeItemID = "-1";
            bool caughtSoapException = false;

            try
            {
                // Call GetAttachmentCollection with negative listItemID.
                this.listswsAdapterInstance.GetAttachmentCollection(listName, negativeItemID);

                Site.Assert.Fail("GetAttachmentCollection operation must return a SOAP fault if the listItemID parameter is negative.");
            }
            catch (SoapException ex)
            {
                caughtSoapException = true;
                string errorCode = TestSuiteHelper.GetErrorCode(ex);
                string errorString = TestSuiteHelper.GetErrorString(ex);

                // Capture R30371 if the protocol server return a SOAP fault with no error code.
                // Capture R30372 if the protocol server return a SOAP fault with error code 0x81020014.
                Site.Log.Add(
                            LogEntryKind.Debug,
                            "The actual value: errorCode[{0}]",
                            string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

                Site.Log.Add(
                            LogEntryKind.Debug,
                            "The actual value: errorString[{0}]",
                            string.IsNullOrEmpty(errorString) ? "NullOrEmpty" : errorString);

                if (Common.IsRequirementEnabled(30371, this.Site))
                {
                    string expectedErrorString = "Value does not fall within the expected range.";

                    bool isR30371Verified = errorCode == null && expectedErrorString.Equals(errorString, StringComparison.OrdinalIgnoreCase);

                    Site.CaptureRequirementIfIsTrue(
                            isR30371Verified,
                            30371,
                            @"[In GetAttachmentCollection operation]  Implementation does return the following SOAP fault with no error code: ""Value does not fall within the expected range."", "
                            + @" if the listItemID is a negative value. (<43> Section 3.1.4.14: wss3 return the following SOAP fault with no error code: ""Value does not fall within the expected range."")");
                }

                if (Common.IsRequirementEnabled(30372, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<string>(
                            "0x81020014",
                            errorCode,
                            30372,
                            @"[In GetAttachmentCollection operation] Implementation does return a SOAP fault with error code 0x81020014, if the listItemID is a negative value. (SharePoint Foundation 2010 and above follow this behavior.)");
                }
            }

            Site.Assert.IsTrue(caughtSoapException, "A SOAP exception should be thrown by the server when invoke the operation 'GetAttachmentCollection' with negative listItemID.");

            #endregion
        }

        #endregion

        #endregion

        #region Override methods
        /// <summary>
        /// This method will run before test case executes
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.listswsAdapterInstance = this.Site.GetAdapter<IMS_LISTSWSAdapter>();

            Common.CheckCommonProperties(this.Site, true);

            #region new initialization
            if (!TestSuiteHelper.GuardEnviromentClean())
            {
                Site.Debug.Fail("The test environment is not clean, refer the log files for details.");
            }

            // Initialize the TestSuiteHelper
            TestSuiteHelper.Initialize(this.Site, this.listswsAdapterInstance);
            #endregion
        }

        /// <summary>
        /// This method will run after test case executes
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            #region new clean up
            TestSuiteHelper.CleanUp();
            #endregion
        }

        #endregion
    }
}