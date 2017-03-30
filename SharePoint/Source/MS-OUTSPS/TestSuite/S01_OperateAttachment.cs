namespace Microsoft.Protocols.TestSuites.MS_OUTSPS
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A test class contains test cases of S01 scenario.
    /// </summary>
    [TestClass]
    public class S01_OperateAttachment : TestSuiteBase
    {
        #region Additional test attributes, Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.ClassCleanup();
        }

        #endregion

        #region Test cases

        #region MSOUTSPS_S01_TC01_OperateAttachment_AppointmentTemplateType
        /// <summary>
        /// This test case is used to verify AddAttachment operation, DeleteAttachment operation, GetAttachment operation,
        /// and GetAttachmentCollection operation with Appointment template type.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S01_TC01_OperateAttachment_AppointmentTemplateType()
        {
            #region Add a list item

            // Add one list into SUT.
            string listId = this.AddListToSUT(TemplateType.Events);

            // Add one list item
            List<string> addedListitems = this.AddItemsToList(listId, 1);
            string addedListitemId = addedListitems[0];
            byte[] attachmentContent = this.GenerateUniqueAttachmentContents(5);
            string attachmentName = this.GetUniqueAttachmentName();

            #endregion

            #region AddAttachment operation

            // Call AddAttachment operation.
            string fileUrl = OutspsAdapter.AddAttachment(
                                    listId,
                                    addedListitemId,
                                    attachmentName,
                                    attachmentContent);

            #endregion

            #region HTTPGET operation

            // Get full URL of an attachment
            string fullUrlOfAttachment = this.GetAttachmentFullUrl(listId, addedListitemId, attachmentName);
            Uri fullUrlOfAttachmentPath;
            if (!Uri.TryCreate(fullUrlOfAttachment, UriKind.RelativeOrAbsolute, out fullUrlOfAttachmentPath))
            {
                this.Site.Assert.Fail("The full url of attachment should be valid Uri format string.");
            }

            // Call HTTPGET operation.
            byte[] getContentsOfAttachment = OutspsAdapter.HTTPGET(fullUrlOfAttachmentPath, "f");

            #endregion

            #region Capture R1062, R1070, R1073, R1075

            // If the length of attachment content in protocol SUT equal to added by AddAttachment operation, then capture R1062, R1070, R1073
            Site.CaptureRequirementIfAreEqual(
                        attachmentContent.Length,
                        getContentsOfAttachment.Length,
                        1062,
                        "[In Message Processing Events and Sequencing Rules][The operation]AddAttachment Adds an attachment to an item.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1070
            Site.CaptureRequirementIfAreEqual(
                        attachmentContent.Length,
                        getContentsOfAttachment.Length,
                        1070,
                        "[In AddAttachment]AddAttachment is used by protocol clients to create a new attachment on an item on the protocol server.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R10730
            Site.CaptureRequirementIfAreEqual(
                        attachmentContent.Length,
                        getContentsOfAttachment.Length,
                        10730,
                        "[In Messages]AddAttachmentResponse specified the response to a request to create a new attachment on an item on the protocol server.");

            // If the AddAttachment operation return valid attachment url, then Capture R1075
            Site.CaptureRequirementIfIsNotNull(
                        fileUrl,
                        1075,
                        "[In AddAttachmentResponse]If an AddAttachmentResponse is received, then the upload was successful.");

            #endregion

            #region GetAttachmentCollection operation

            byte[] attachmentContentSecond = this.GenerateUniqueAttachmentContents(5);
            string attachmentNameSecond = this.GetUniqueAttachmentName();

            // Call the AddAttachment operation.
            OutspsAdapter.AddAttachment(
                          listId,
                          addedListitemId,
                          attachmentNameSecond,
                          attachmentContentSecond);

            // Call GetAttachmentCollection operation.
            GetAttachmentCollectionResponseGetAttachmentCollectionResult getAttachementCollectionResult = OutspsAdapter.GetAttachmentCollection(listId, addedListitemId);

            #endregion

            #region Capture R1065, R11000

            // If add new attachment successfully, total attachment number is 2, Capture R1065, R11000
            bool isContainExpectedAttachment = this.VerifyWhetherContainExpectedNumberAttachment(getAttachementCollectionResult, 2);

            this.Site.CaptureRequirementIfIsTrue(
                isContainExpectedAttachment,
                1065,
                "[In Message Processing Events and Sequencing Rules][The operation]GetAttachmentCollection Gets a list of the attachments on an item.");

            this.Site.CaptureRequirementIfIsTrue(
                isContainExpectedAttachment,
                11000,
                "[In Messages]GetAttachmentCollectionResponse specified the response to a request to get the list of all attachments on a single item in one list.");

            #endregion

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                #region GetListItemChangesSinceToken operation

                // Set CamlQueryOptions and view fields make the "attachment" field present in response.
                CamlQueryOptions camloptions = new CamlQueryOptions();
                camloptions.QueryOptions = new CamlQueryOptionsQueryOptions();
                camloptions.QueryOptions.IncludeAttachmentUrls = bool.TrueString;
                camloptions.QueryOptions.IncludeAttachmentVersion = bool.TrueString;
                CamlViewFields viewfieds = this.GenerateViewFields(false, new List<string> { "Attachments" });

                // Call GetListItemChangesSinceToken operation.
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult getListItemChangesRes = null;
                getListItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                                                                  listId,
                                                                                  null,
                                                                                  null,
                                                                                  viewfieds,
                                                                                  null,
                                                                                  camloptions,
                                                                                  null,
                                                                                  null);

                this.Site.Assert.IsNotNull(getListItemChangesRes, "SUT should return a response contain data.");
                string headerValue = this.GetIfMatchHeaderValueFromResponse(getListItemChangesRes, fullUrlOfAttachment, int.Parse(addedListitemId));

                #endregion

                #region Capture R1241

                // If the header is returned, then R1241 should be covered.
                this.Site.CaptureRequirementIfIsNotNull(
                    headerValue,
                    1241,
                    "[In HTTP PUT]The Attachments property MUST contain a file version if protocol clients have included the IncludeAttachmentUrls and IncludeAttachmentVersion elements specified in [MS-LISTSWS].");

                #endregion

                #region HTTPPUT operation

                // Call HTTPPUT operation to update the attachment contents for the first attachment.
                fullUrlOfAttachmentPath = null;
                if (!Uri.TryCreate(fullUrlOfAttachment, UriKind.RelativeOrAbsolute, out fullUrlOfAttachmentPath))
                {
                    this.Site.Assert.Fail("The full url of attachment should be valid Uri format string.");
                }

                byte[] newattachmentContents = this.GenerateUniqueAttachmentContents(6);
                OutspsAdapter.HTTPPUT(fullUrlOfAttachmentPath, headerValue, newattachmentContents);

                // Verify the updated attachment contents' length
                this.VerifyAttachmentContentsLength(fullUrlOfAttachment, newattachmentContents.Length);

                #endregion

                #region HTTPPUT operation with the header data does not match the current version
                getListItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                                                  listId,
                                                                  null,
                                                                  null,
                                                                  viewfieds,
                                                                  null,
                                                                  camloptions,
                                                                  null,
                                                                  null);

                this.Site.Assert.IsNotNull(getListItemChangesRes, "SUT should return a response contain data.");
                string headerValue2 = this.GetIfMatchHeaderValueFromResponse(getListItemChangesRes, fullUrlOfAttachment, int.Parse(addedListitemId));
                int index = Int32.Parse(headerValue2.Replace("\"", string.Empty).Split(',')[1]);
                string invalidHeaderValue = headerValue2.Split(',')[0] + "," + (index + 1) + "\"";
                try
                {
                    OutspsAdapter.HTTPPUT(fullUrlOfAttachmentPath, invalidHeaderValue, newattachmentContents);
                }
                catch (WebException exception)
                {
                    HttpWebResponse webResponse = exception.Response as HttpWebResponse;

                    // Verify MS-OUTSPS requirement: MS-OUTSPS_R1243
                    this.Site.CaptureRequirementIfAreEqual<int>(
                        412,
                        (int)webResponse.StatusCode,
                        1243,
                        "[In HTTP PUT]Protocol servers MUST respond with an HTTP status code 412 (which indicates a precondition failed) if the header data does not match the current version.");
                }
                #endregion
            }

            OutspsAdapter.DeleteAttachment(listId, addedListitemId, fullUrlOfAttachment);

            #region Capture R10930

            // If the operation does not return SoapException, capture R10930 directly. Because the schema of the DeleteAttachmentResponse define in [MS-LISTSWS]: <s:element name="DeleteAttachmentResponse"><s:complexType/></s:element>, does not contain any complex type definition, so the proxy class marked this operation as void type return.
            this.Site.CaptureRequirement(
                10930,
                "[In Messages]DeleteAttachmentResponse specified the response to a request to delete attachments from an item on the protocol server.");

            #endregion
 
            // Verify whether the attachment was deleted.
            bool isDeleteSucceed = this.VerifyDeleteAttachmentSucceed(fullUrlOfAttachment);
 
            #region Capture R1064, R1091

            // Because the specified attachment was deleted, the HTTPGET operation could not find it, so there will have a soap exception.
            this.Site.CaptureRequirementIfIsTrue(
                  isDeleteSucceed,
                  1064,
                  @"[In Message Processing Events and Sequencing Rules][The operation]DeleteAttachment Deletes an attachment from an item on a list.");

            this.Site.CaptureRequirementIfIsTrue(
                     isDeleteSucceed,
                      1091,
                      @"[In DeleteAttachment]Protocol clients use DeleteAttachment to delete attachments from an item on the protocol server.");

            #endregion
        }

        #endregion

        #region MSOUTSPS_S01_TC02_OperateAttachment_ContactsTemplateType
        /// <summary>
        /// This test case is used to verify AddAttachment operation, DeleteAttachment operation, GetAttachment operation,
        /// and GetAttachmentCollection operation with Contacts template type.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S01_TC02_OperateAttachment_ContactsTemplateType()
        {
            #region Add a list item

            // Add one list into SUT.
            string listId = this.AddListToSUT(TemplateType.Contacts);

            // Add one list item
            List<string> addedListitems = this.AddItemsToList(listId, 1);
            string addedListitemId = addedListitems[0];
            byte[] attachmentContent = this.GenerateUniqueAttachmentContents(5);
            string attachmentName = this.GetUniqueAttachmentName();

            #endregion

            #region AddAttachment operation

            // Call AddAttachment operation.
            OutspsAdapter.AddAttachment(
                                    listId,
                                    addedListitemId,
                                    attachmentName,
                                    attachmentContent);

            #endregion

            #region HTTPGET operation

            // Get full URL of an attachment
            string fullUrlOfAttachment = this.GetAttachmentFullUrl(listId, addedListitemId, attachmentName);
            Uri fullUrlOfAttachmentPath;
            if (!Uri.TryCreate(fullUrlOfAttachment, UriKind.RelativeOrAbsolute, out fullUrlOfAttachmentPath))
            {
                this.Site.Assert.Fail("The full url of attachment should be valid Uri format string.");
            }

            // Call HTTPGET operation.
            byte[] getContentsOfAttachment = OutspsAdapter.HTTPGET(fullUrlOfAttachmentPath, "f");
            this.Site.Assert.AreEqual<int>(
                        attachmentContent.Length,
                        getContentsOfAttachment.Length,
                        "The attachment content's length should equal to added by AddAttachment operation.");

            #endregion

            #region GetAttachmentCollection operation

            byte[] attachmentContentSecond = this.GenerateUniqueAttachmentContents(5);
            string attachmentNameSecond = this.GetUniqueAttachmentName();

            // Call the AddAttachment operation.
            OutspsAdapter.AddAttachment(
                              listId,
                              addedListitemId,
                              attachmentNameSecond,
                              attachmentContentSecond);

            // Call GetAttachmentCollection operation.
            GetAttachmentCollectionResponseGetAttachmentCollectionResult getAttachementCollectionResult = OutspsAdapter.GetAttachmentCollection(listId, addedListitemId);

            // If add new attachment successfully, total attachment number is 2
            this.VerifyWhetherContainExpectedNumberAttachment(getAttachementCollectionResult, 2);

            #endregion

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                #region GetListItemChangesSinceToken operation

                // Set CamlQueryOptions and view fields make the "attachment" field present in response.
                CamlQueryOptions camloptions = new CamlQueryOptions();
                camloptions.QueryOptions = new CamlQueryOptionsQueryOptions();
                camloptions.QueryOptions.IncludeAttachmentUrls = bool.TrueString;
                camloptions.QueryOptions.IncludeAttachmentVersion = bool.TrueString;
                CamlViewFields viewfieds = this.GenerateViewFields(false, new List<string> { "Attachments" });

                // Call GetListItemChangesSinceToken operation.
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult getlistItemChangesRes = null;
                getlistItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                                                                  listId,
                                                                                  null,
                                                                                  null,
                                                                                  viewfieds,
                                                                                  null,
                                                                                  camloptions,
                                                                                  null,
                                                                                  null);

                this.Site.Assert.IsNotNull(getlistItemChangesRes, "SUT should return a response contain data.");
                string headerValue = this.GetIfMatchHeaderValueFromResponse(getlistItemChangesRes, fullUrlOfAttachment, int.Parse(addedListitemId));

                #endregion

                #region HTTPPUT operation

                // Call HTTPPUT operation.
                fullUrlOfAttachmentPath = null;
                if (!Uri.TryCreate(fullUrlOfAttachment, UriKind.RelativeOrAbsolute, out fullUrlOfAttachmentPath))
                {
                    this.Site.Assert.Fail("The full url of attachment should be valid Uri format string.");
                }

                byte[] newattachmentContents = this.GenerateUniqueAttachmentContents(6);
                OutspsAdapter.HTTPPUT(fullUrlOfAttachmentPath, headerValue, newattachmentContents);

                // Verify updated attachment.
                this.VerifyAttachmentContentsLength(fullUrlOfAttachment, newattachmentContents.Length);
                #endregion
            }

            OutspsAdapter.DeleteAttachment(listId, addedListitemId, fullUrlOfAttachment);

            // Verify whether delete attachment succeed.
            this.VerifyDeleteAttachmentSucceed(fullUrlOfAttachment);
        }

        #endregion

        #region MSOUTSPS_S01_TC03_OperateAttachment_DiscussionBoardTemplateType
        /// <summary>
        /// This test case is used to verify AddAttachment operation, DeleteAttachment operation, GetAttachment operation,
        /// and GetAttachmentCollection operation with Discussion_Board template type.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S01_TC03_OperateAttachment_DiscussionBoardTemplateType()
        {
            #region Add a list item

            // Add one list into SUT.
            string listId = this.AddListToSUT(TemplateType.Discussion_Board);

            // Add one list item
            List<string> addedListitems = this.AddItemsToList(listId, 1);
            string addedListitemId = addedListitems[0];
            byte[] attachmentContent = this.GenerateUniqueAttachmentContents(5);
            string attachmentName = this.GetUniqueAttachmentName();

            #endregion

            #region AddAttachment operation

            // Call AddAttachment operation.
            OutspsAdapter.AddAttachment(
                                    listId,
                                    addedListitemId,
                                    attachmentName,
                                    attachmentContent);

            #endregion

            #region HTTPGET operation

            // Get full URL of an attachment
            string fullUrlOfAttachment = this.GetAttachmentFullUrl(listId, addedListitemId, attachmentName);
            Uri fullUrlOfAttachmentPath;
            if (!Uri.TryCreate(fullUrlOfAttachment, UriKind.RelativeOrAbsolute, out fullUrlOfAttachmentPath))
            {
                this.Site.Assert.Fail("The full url of attachment should be valid Uri format string.");
            }

            // Call HTTPGET operation.
            byte[] getContentsOfAttachment = OutspsAdapter.HTTPGET(fullUrlOfAttachmentPath, "f");

            this.Site.Assert.AreEqual<int>(
                        attachmentContent.Length,
                        getContentsOfAttachment.Length,
                        "The attachment content's length should equal to added by AddAttachment operation.");

            #endregion

            #region GetAttachmentCollection operation

            byte[] attachmentContentSecond = this.GenerateUniqueAttachmentContents(5);
            string attachmentNameSecond = this.GetUniqueAttachmentName();

            // Call the AddAttachment operation.
            OutspsAdapter.AddAttachment(
                            listId,
                            addedListitemId,
                            attachmentNameSecond,
                            attachmentContentSecond);

            // Call GetAttachmentCollection operation.
            GetAttachmentCollectionResponseGetAttachmentCollectionResult getAttachementCollectionResult = OutspsAdapter.GetAttachmentCollection(listId, addedListitemId);

            // If the new attachment was added successfully,, total attachment number is 2
            this.VerifyWhetherContainExpectedNumberAttachment(getAttachementCollectionResult, 2);

            #endregion

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                #region GetListItemChangesSinceToken operation

                // Set CamlQueryOptions and view fields make the "attachment" field present in response.
                CamlQueryOptions camloptions = new CamlQueryOptions();
                camloptions.QueryOptions = new CamlQueryOptionsQueryOptions();
                camloptions.QueryOptions.IncludeAttachmentUrls = bool.TrueString;
                camloptions.QueryOptions.IncludeAttachmentVersion = bool.TrueString;
                CamlViewFields viewfieds = this.GenerateViewFields(false, new List<string> { "Attachments" });

                // Call GetListItemChangesSinceToken operation.
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult getlistItemChangesRes = null;
                getlistItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                                                                  listId,
                                                                                  null,
                                                                                  null,
                                                                                  viewfieds,
                                                                                  null,
                                                                                  camloptions,
                                                                                  null,
                                                                                  null);

                this.Site.Assert.IsNotNull(getlistItemChangesRes, "SUT should return a response contain data.");
                string headerValue = this.GetIfMatchHeaderValueFromResponse(getlistItemChangesRes, fullUrlOfAttachment, int.Parse(addedListitemId));

                #endregion

                #region HTTPPUT operation

                // Call HTTPPUT operation.
                fullUrlOfAttachmentPath = null;
                if (!Uri.TryCreate(fullUrlOfAttachment, UriKind.RelativeOrAbsolute, out fullUrlOfAttachmentPath))
                {
                    this.Site.Assert.Fail("The full url of attachment should be valid Uri format string.");
                }

                byte[] newattachmentContents = this.GenerateUniqueAttachmentContents(6);
                OutspsAdapter.HTTPPUT(fullUrlOfAttachmentPath, headerValue, newattachmentContents);

                // Verify updated attachment.
                this.VerifyAttachmentContentsLength(fullUrlOfAttachment, newattachmentContents.Length);
                #endregion
            }

            OutspsAdapter.DeleteAttachment(listId, addedListitemId, fullUrlOfAttachment);

            // Verify whether delete attachment succeed.
            this.VerifyDeleteAttachmentSucceed(fullUrlOfAttachment);
        }

        #endregion

        #region MSOUTSPS_S01_TC04_OperateAttachment_TasksTemplateType
        /// <summary>
        /// This test case is used to verify AddAttachment operation, DeleteAttachment operation, GetAttachment operation,
        /// and GetAttachmentCollection operation with Tasks template type.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S01_TC04_OperateAttachment_TasksTemplateType()
        {
            #region Add a list item

            // Add one list into SUT.
            string listId = this.AddListToSUT(TemplateType.Tasks);

            // Add one list item
            List<string> addedListitems = this.AddItemsToList(listId, 1);
            string addedListitemId = addedListitems[0];
            byte[] attachmentContent = this.GenerateUniqueAttachmentContents(5);
            string attachmentName = this.GetUniqueAttachmentName();

            #endregion

            #region AddAttachment operation

            // Call AddAttachment operation.
           OutspsAdapter.AddAttachment(
                                    listId,
                                    addedListitemId,
                                    attachmentName,
                                    attachmentContent);

            #endregion

            #region HTTPGET operation

            // Get full URL of an attachment
            string fullUrlOfAttachment = this.GetAttachmentFullUrl(listId, addedListitemId, attachmentName);

            // Call HTTPGET operation.
            Uri fullUrlOfAttachmentPath;
            if (!Uri.TryCreate(fullUrlOfAttachment, UriKind.RelativeOrAbsolute, out fullUrlOfAttachmentPath))
            {
                this.Site.Assert.Fail("The full url of attachment should be valid Uri format string.");
            }

            byte[] getContentsOfAttachment = OutspsAdapter.HTTPGET(fullUrlOfAttachmentPath, "f");

            this.Site.Assert.AreEqual<int>(
                        attachmentContent.Length,
                        getContentsOfAttachment.Length,
                        "The attachment content's length should equal to added by AddAttachment operation.");

            #endregion

            #region GetAttachmentCollection operation

            byte[] attachmentContentSecond = this.GenerateUniqueAttachmentContents(5);
            string attachmentNameSecond = this.GetUniqueAttachmentName();

            // Call the AddAttachment operation.
            OutspsAdapter.AddAttachment(
                           listId,
                           addedListitemId,
                           attachmentNameSecond,
                           attachmentContentSecond);

            // Call GetAttachmentCollection operation.
            GetAttachmentCollectionResponseGetAttachmentCollectionResult getAttachementCollectionResult = OutspsAdapter.GetAttachmentCollection(listId, addedListitemId);

            // If add new attachment successfully, total attachment number is 2
            this.VerifyWhetherContainExpectedNumberAttachment(getAttachementCollectionResult, 2);

            #endregion

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                #region GetListItemChangesSinceToken operation

                // Set CamlQueryOptions and view fields make the "attachment" field present in response.
                CamlQueryOptions camloptions = new CamlQueryOptions();
                camloptions.QueryOptions = new CamlQueryOptionsQueryOptions();
                camloptions.QueryOptions.IncludeAttachmentUrls = bool.TrueString;
                camloptions.QueryOptions.IncludeAttachmentVersion = bool.TrueString;
                CamlViewFields viewfieds = this.GenerateViewFields(false, new List<string> { "Attachments" });

                // Call GetListItemChangesSinceToken operation.
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult getlistItemChangesRes = null;
                getlistItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                                                                  listId,
                                                                                  null,
                                                                                  null,
                                                                                  viewfieds,
                                                                                  null,
                                                                                  camloptions,
                                                                                  null,
                                                                                  null);

                this.Site.Assert.IsNotNull(getlistItemChangesRes, "SUT should return a response contain data.");
                string headerValue = this.GetIfMatchHeaderValueFromResponse(getlistItemChangesRes, fullUrlOfAttachment, int.Parse(addedListitemId));

                #endregion

                #region HTTPPUT operation

                // Call HTTPPUT operation.
                fullUrlOfAttachmentPath = null;
                if (!Uri.TryCreate(fullUrlOfAttachment, UriKind.RelativeOrAbsolute, out fullUrlOfAttachmentPath))
                {
                    this.Site.Assert.Fail("The full url of attachment should be valid Uri format string.");
                }

                byte[] newattachmentContents = this.GenerateUniqueAttachmentContents(6);
                OutspsAdapter.HTTPPUT(fullUrlOfAttachmentPath, headerValue, newattachmentContents);

                // Verify updated attachment.
                this.VerifyAttachmentContentsLength(fullUrlOfAttachment, newattachmentContents.Length);
                #endregion
            }

            OutspsAdapter.DeleteAttachment(listId, addedListitemId, fullUrlOfAttachment);

            // Verify whether delete attachment succeed.
            this.VerifyDeleteAttachmentSucceed(fullUrlOfAttachment);
        }

        #endregion

        #region MSOUTSPS_S01_TC05_AddAttachment_Fail

        /// <summary>
        /// This test case is used to verify AddAttachment operation, create a new attachment failed.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S01_TC05_AddAttachment_Fail()
        {
            // Add one list into SUT.
            string listId = this.AddListToSUT(TemplateType.Events);

            // Add one list item
            List<string> addedListitems = this.AddItemsToList(listId, 1);
            string addedListitemId = addedListitems[0];
            byte[] attachmentContent = this.GenerateUniqueAttachmentContents(5);
            string attachmentName = this.GetUniqueAttachmentName();
            string errorCode = null;

            // Call AddAttachment operation.
            OutspsAdapter.AddAttachment(
                                    listId,
                                    addedListitemId,
                                    attachmentName,
                                    attachmentContent);

            try
            {
                OutspsAdapter.AddAttachment(
                                    listId,
                                    addedListitemId,
                                    attachmentName,
                                    attachmentContent);
            }
            catch (SoapException soapException)
            {
                errorCode = Common.ExtractErrorCodeFromSoapFault(soapException);
            }

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1076
            Site.CaptureRequirementIfAreEqual(
                "0x81020067",
                errorCode,
                1076,
                "[In AddAttachmentResponse][If a SOAP exception is received instead of an AddAttachmentResponse, the protocol client SHOULD behave as follows:]If the exception errorcode (see [MS-LISTSWS] section 3.1.4.1) is 0x81020067, this indicates that the item already has an attachment with the same file name.");
        }

        #endregion
 
        #endregion Test cases

        #region Private method 
        
        /// <summary>
        /// A method used to get IF-MATCH header value by specified attachment url.
        /// </summary>
        /// <param name="getListItemChangesRes">A parameter represents a response of "GetListItemChangesSinceToken" operation where method will find the expected IF-MATCH header value.</param>
        /// <param name="attachmentFullUrl">A parameter represents the full url of attachment which is used to match attachment field value.</param>
        /// <param name="listItemId">A parameter represents the list item id of a list item where the attachment is added.</param>
        /// <returns>A return value represents the expected IF-MATCH header value.</returns>
        private string GetIfMatchHeaderValueFromResponse(GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult getListItemChangesRes, string attachmentFullUrl, int listItemId)
        {
            if (null == getListItemChangesRes || null == getListItemChangesRes.listitems || null == getListItemChangesRes.listitems.data
                || null == getListItemChangesRes.listitems.data.Any)
            {
                throw new ArgumentException("Should contain valid item changed data.", "getListItemChangesRes");
            }

            if (listItemId <= 0)
            {
                throw new ArgumentException("The value should be large than Zero.", "listItemId");
            }

            if (string.IsNullOrEmpty(attachmentFullUrl))
            {
                throw new ArgumentException("The value should have non-empty string value", "attachmentFullUrl");
            }

            if (0 == getListItemChangesRes.listitems.data.Any.Length)
            {
                throw new InvalidOperationException("Could not get the attachment changed record.");
            }

            // Get attachment field value.
            XmlNode[] changesRecords = this.GetZrowItems(getListItemChangesRes.listitems.data.Any);
            string attachmentFieldValue = Common.GetZrowAttributeValue(changesRecords, listItemId - 1, "ows_attachments");

            if (!attachmentFieldValue.StartsWith(@";#", StringComparison.OrdinalIgnoreCase))
            {
                this.Site.Assert.Fail("the attachment field value must begin with [;#].");
            }

            if (!attachmentFieldValue.EndsWith(@";#", StringComparison.OrdinalIgnoreCase))
            {
                this.Site.Assert.Fail("the attachment field value must end with [;#].");
            }

            string[] splitValues = attachmentFieldValue.Split(new string[] { @";#" }, StringSplitOptions.RemoveEmptyEntries);

            // The parse logic is described in MS-LISTSWS section 2.2.4.4
            string expectedHeaderValue = string.Empty;
            for (int index = 0; index < splitValues.Length; index++)
            {
                if (splitValues[index].Equals(attachmentFullUrl, StringComparison.OrdinalIgnoreCase)
                    && index + 1 < splitValues.Length)
                {
                    expectedHeaderValue = splitValues[index + 1];
                    break;
                }
            }

            if (string.IsNullOrEmpty(expectedHeaderValue))
            {
                string errorMsg = string.Format("Could not find expected [IF-Match] header value by specified attachment url[{0}]", attachmentFullUrl);
                this.Site.Assert.Fail(errorMsg);
            }

            expectedHeaderValue = string.Format(@"""{0}""", expectedHeaderValue);
            return expectedHeaderValue;
        }

        /// <summary>
        /// A method used to verify the response of GetAttachmentCollection whether contain expected Number Attachment items. If it does not pass the verification, this method will throw a Assert exception.
        /// </summary>
        /// <param name="getAttachementCollectionResult">A parameter represents the response of GetAttachmentCollection operation.</param>
        /// <param name="expectedAttachmentsNumber">A parameter represents the expected attachment items' number which is used to check the response of GetAttachmentCollection operation.</param>
        /// <returns>Return true indicating the response of GetAttachmentCollection contains expected number of attachment items</returns>
        private bool VerifyWhetherContainExpectedNumberAttachment(GetAttachmentCollectionResponseGetAttachmentCollectionResult getAttachementCollectionResult, int expectedAttachmentsNumber)
        {
            this.Site.Assert.IsNotNull(getAttachementCollectionResult, "The GetAttachmentCollection operation should return valid response.");
            this.Site.Assert.IsNotNull(getAttachementCollectionResult.Attachments, "The response of GetAttachmentCollection operation should contain valid attachments data.");

            // Verify the number of attachment collection's items whether equal to expected value.
            this.Site.Assert.AreEqual<int>(
                        expectedAttachmentsNumber,
                        getAttachementCollectionResult.Attachments.Length,
                        "The response of GetAttachmentCollection operation should contain [{0}] expected attachments data, actual:[{1}]",
                        expectedAttachmentsNumber,
                        getAttachementCollectionResult.Attachments.Length);

            return true;
        }

        /// <summary>
        /// A method used to verify attachment contents' length which is got from protocol SUT whether equal to specified value. If it does not pass the verification, this method will throw a Assert exception.
        /// </summary>
        /// <param name="fullUrlOfAttachment">A parameter represents the full url of attachment which is used by HTTPGET operation to get the actual attachment content.</param>
        /// <param name="expectedContentsLength">A parameter represents the expected content length.</param>
        /// <returns>Return true indicating the attachment content length equal to the expected value.</returns>
        private bool VerifyAttachmentContentsLength(string fullUrlOfAttachment, int expectedContentsLength)
        {
           if (string.IsNullOrEmpty(fullUrlOfAttachment))
           {
               throw new ArgumentNullException("fullUrlOfAttachment");
           }

           Uri fullUrlOfAttachmentPath;
           if (!Uri.TryCreate(fullUrlOfAttachment, UriKind.RelativeOrAbsolute, out fullUrlOfAttachmentPath))
           {
               this.Site.Assert.Fail("The full url of attachment should be valid Uri format string.");
           }

           byte[] getContentsOfAttachment = OutspsAdapter.HTTPGET(fullUrlOfAttachmentPath, "f");

           this.Site.Assert.AreEqual<int>(
                                    expectedContentsLength,
                                    getContentsOfAttachment.Length,
                                    "The attachment contents' length should equal to expected value[{0}]",
                                    expectedContentsLength);
           return true;
        }

        /// <summary>
        /// A method used to verify whether the attachment have been deleted or not.
        /// </summary>
        /// <param name="fullUrlOfAttachment">A parameter represents the full url of attachment which is used by HTTPGET operation to get the actual attachment content.</param>
        /// <returns>Return true indicating delete attachment succeed.</returns>
        private bool VerifyDeleteAttachmentSucceed(string fullUrlOfAttachment)
        {
            if (string.IsNullOrEmpty(fullUrlOfAttachment))
            {
                throw new ArgumentNullException("fullUrlOfAttachment");
            }

            HttpStatusCode lowLevelStatusCode = HttpStatusCode.OK;
            Uri fullUrlOfAttachmentPath;
            if (!Uri.TryCreate(fullUrlOfAttachment, UriKind.RelativeOrAbsolute, out fullUrlOfAttachmentPath))
            {
                this.Site.Assert.Fail("The full url of attachment should be valid Uri format string.");
            }

            try
            {
               OutspsAdapter.HTTPGET(fullUrlOfAttachmentPath, "f");
            }
            catch (WebException webException)
            {
                lowLevelStatusCode = this.GetStatusCodeFromWebException(webException);
                Site.Log.Add(LogEntryKind.Debug, "The attachment has not been found" + webException.Message.ToString());
            }

            this.Site.Assert.AreEqual<HttpStatusCode>(
                      HttpStatusCode.NotFound,
                      lowLevelStatusCode,
                      @"The protocol SUT should return ""404"" status code which means ""NotFound"", if the attachment[{0}] have been deleted.",
                      fullUrlOfAttachment);
            return true;
        }

        #endregion Private method 
    }
}