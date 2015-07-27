//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_LISTSWS
{
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using System.Xml.Schema;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Protocol Adapter's Implementation.
    /// </summary>
    public partial class MS_LISTSWSAdapter : ManagedAdapterBase, IMS_LISTSWSAdapter
    {
        #region Private member variables
        /// <summary>
        /// Web service proxy generated from the full WSDL of LISTSWS
        /// </summary>
        private ListsSoap listsProxy;

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the security credentials.
        /// </summary>
        public ICredentials Credentials
        {
            get
            {
                return this.listsProxy.Credentials;
            }

            set
            {
                this.listsProxy.Credentials = value;
            }
        }

        #endregion

        #region MS-LISTSWS adapter operations

        /// <summary>
        /// Add a list in the current site based on the specified name, description, and list template identifier.
        /// </summary>
        /// <param name="listName">The title of the list which will be added.</param>
        /// <param name="description">Text which will be set as description of newly created list.</param>
        /// <param name="templateID">The template ID used to create this list.</param>
        /// <returns>Returns the AddList result.</returns>
        public AddListResponseAddListResult AddList(string listName, string description, int templateID)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            AddListResponseAddListResult result = null;
            try
            {
                // invoke the proxy
                result = this.listsProxy.AddList(listName, description, templateID);

                // Verify the requirements of AddList operation.
                this.VerifyAddListOperation(result);

                // Verify the requirements of the transport.
                this.VerifyTransportRequirements();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// Check in a document item to a document library.
        /// </summary>
        /// <param name="pageUrl">The Url specifies which file will be check in. </param>
        /// <param name="comment">The comments of the check in.</param>
        /// <param name="checkinType">A string representation the check type value <see cref="CheckInTypeValue"/> </param>
        /// <returns>Returns True indicating CheckInFile was successful</returns>
        public bool CheckInFile(string pageUrl, string comment, string checkinType)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            bool result = true;
            try
            {
                result = this.listsProxy.CheckInFile(pageUrl, comment, checkinType);

                // Verify the requirements of CheckInFile operation.
                this.VerifyCheckInFileOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The CheckOutFile operation is used to check out a document in a document library.
        /// </summary>
        /// <param name="pageUrl">The Url specifies which file will be check out.</param>
        /// <param name="checkoutToLocal">"TRUE" means to keep a local version for offline editing</param>
        /// <param name="lastModified">A string in date format that represents the date and time of the last modification by the site to the file; for example, "20 Jun 1982 12:00:00 GMT</param>
        /// <returns>Returns the checkout file result.</returns>
        public bool CheckOutFile(string pageUrl, string checkoutToLocal, string lastModified)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            bool result = false;
            try
            {
                result = this.listsProxy.CheckOutFile(pageUrl, checkoutToLocal, lastModified);

                // Verify the requirements of CheckOutFile operation.
                this.VerifyCheckOutFileOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The AddAttachment operation is used to add an attachment to the specified list item in the specified list.
        /// </summary>
        /// <param name="listName">The GUID or the list title of the list in which the list item to add attachment.</param>
        /// <param name="listItemID">The id of the list item in which the attachment will be added.</param>
        /// <param name="fileName">The name of the file being added as an attachment.</param>
        /// <param name="attachment">Content of the attachment file (byte array).</param>
        /// <returns>The URL of the newly added attachment.</returns>s
        public string AddAttachment(string listName, string listItemID, string fileName, byte[] attachment)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            string attachmentRelativeUrl = null;
            try
            {
                attachmentRelativeUrl = this.listsProxy.AddAttachment(listName, listItemID, fileName, attachment);

                // Verify the requirements of the AddAttachment operation.
                this.VerifyAddAttachmentOperation(attachmentRelativeUrl);

                // Verify the requirements of the transport.
                this.VerifyTransportRequirements();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return attachmentRelativeUrl;
        }

        /// <summary>
        /// The DeleteAttachment operation is used to remove the attachment from the specified list 
        /// item in the specified list.
        /// </summary>
        /// <param name="listName">The name of the list in which the list item to delete existing attachment.</param>
        /// <param name="listItemID">The id of the list item from which the attachment will be deleted.</param>
        /// <param name="url">Absolute URL of the attachment that should be deleted.</param>
        public void DeleteAttachment(string listName, string listItemID, string url)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            try
            {
                this.listsProxy.DeleteAttachment(listName, listItemID, url);

                // Verify the requirements of DeleteAttachment operation.
                this.VerifyDeleteAttachmentOperation();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }
        }

        /// <summary>
        /// The GetListItemChanges operation is used to retrieve the list items that have been inserted or updated 
        /// since the specified date and time and matching the specified filter criteria.
        /// </summary>
        /// <param name="listName">The name of the list from which the list item changes will be got</param>
        /// <param name="viewFields">Indicates which fields of the list item SHOULD be returned</param>
        /// <param name="since">The date and time to start retrieving changes in the list
        /// If the parameter is null, Protocol Server should return all list items
        /// If the date that is passed in is not in UTC format, protocol server will use protocol server's local time zone and converted to UTC time</param>
        /// <param name="contains">Restricts the results returned by giving a specific value to be searched for in the specified list item field</param>
        /// <returns>Return the get list item change result</returns>
        public GetListItemChangesResponseGetListItemChangesResult GetListItemChanges(string listName, CamlViewFields viewFields, string since, CamlContains contains)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            GetListItemChangesResponseGetListItemChangesResult result = null;
            try
            {
                result = this.listsProxy.GetListItemChanges(listName, viewFields, since, contains);

                // Verify the requirements of GetListItemChanges operation.
                this.VerifyGetListItemChangesOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }
            
            return result;
        }

        /// <summary>
        /// The GetAttachmentCollection operation is used to retrieve information about all the lists on the current site.
        /// </summary>
        /// <param name="listName">list name or GUID for returning the result.</param>
        /// <param name="listItemID">The identifier of the content type which will be collected.</param>
        /// <returns>Return attachment collection result.</returns>
        public GetAttachmentCollectionResponseGetAttachmentCollectionResult GetAttachmentCollection(string listName, string listItemID)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            GetAttachmentCollectionResponseGetAttachmentCollectionResult result = null;
            try
            {
                result = this.listsProxy.GetAttachmentCollection(listName, listItemID);

                // Verify the requirements of GetAttachmentCollection operation.
                this.VerifyGetAttachmentCollectionOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The CreateContentType operation is used to create a new content type on a list.
        /// </summary>
        /// <param name="listName">The name of the list for which the content type will be created</param>
        /// <param name="displayName">The XML-encoded name of the content type to be created</param>
        /// <param name="parentType">The identification of a content type from which the content type to be created will inherit</param>
        /// <param name="fields">The container for a list of existing fields to be included in the content type </param>
        /// <param name="contentTypeProperties">The container for properties to set on the content type </param>
        /// <param name="addToView">Specifies whether the fields will be added to the default list view, where "TRUE" MUST correspond to true, and all other values to false</param>
        /// <returns>Return the ID of newly created content type.</returns>
        public string CreateContentType(string listName, string displayName, string parentType, AddOrUpdateFieldsDefinition fields, CreateContentTypeContentTypeProperties contentTypeProperties, string addToView)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");
           
            string result = null;
            try
            {
                result = this.listsProxy.CreateContentType(listName, displayName, parentType, fields, contentTypeProperties, addToView);

                // Verify the requirements of CreateContentType operation.
                this.VerifyCreateContentTypeOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The DeleteContentType operation is used to remove the association between the specified list and the specified content type.
        /// </summary>
        /// <param name="listName">The name of the list in which a content type will be deleted</param>
        /// <param name="contentTypeId">The identifier of the content type which will be deleted</param>
        /// <returns>return delete content type </returns>
        public DeleteContentTypeResponseDeleteContentTypeResult DeleteContentType(string listName, string contentTypeId)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            DeleteContentTypeResponseDeleteContentTypeResult result = null;
            try
            {
                result = this.listsProxy.DeleteContentType(listName, contentTypeId);

                // Verify the requirements of DeleteContentType operation.
                this.VerifyDeleteContentTypeOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The DeleteContentTypeXmlDocument operation is used to delete an XML Document Property from XML Document collection in a content type of a list.
        /// </summary>
        /// <param name="listName">The name of the list in which a content type xml document will be deleted</param>
        /// <param name="contentTypeId">The identifier of the content type for which the xml document will be deleted</param>
        /// <param name="documentUri">The namespace URI of the XML document to remove</param>
        /// <returns>Return deleted content type XmlDocument Result</returns>
        public DeleteContentTypeXmlDocumentResponseDeleteContentTypeXmlDocumentResult DeleteContentTypeXmlDocument(string listName, string contentTypeId, string documentUri)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            DeleteContentTypeXmlDocumentResponseDeleteContentTypeXmlDocumentResult result = null;
            try
            {
                result = this.listsProxy.DeleteContentTypeXmlDocument(listName, contentTypeId, documentUri);

                // Verify the requirements of DeleteContentTypeXmlDocument operation.
                this.VerifyDeleteContentTypeXmlDocumentOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The GetListContentType operation is used to get content type data for a given content type identifier.
        /// </summary>
        /// <param name="listName">The name of the list for which a content type will be got</param>
        /// <param name="contentTypeId">The identifier of the content type of which information will be got</param>
        /// <returns>return List content type</returns>
        public GetListContentTypeResponseGetListContentTypeResult GetListContentType(string listName, string contentTypeId)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            GetListContentTypeResponseGetListContentTypeResult result = null;
            try
            {
                result = this.listsProxy.GetListContentType(listName, contentTypeId);

                // Verify the requirements of GetListContentType operation.
                this.VerifyGetListContentTypeOperation(result);
                SchemaValidation.LastRawResponseXml.GetAttribute("HasExternalDataSource");
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The GetListContentTypes operation is used to retrieve all content types from list.
        /// </summary>
        /// <param name="listName">The name of the list for which a content type will be got</param>
        /// <param name="contentTypeId">The identifier of the content type of which information will be got</param>
        /// <returns>All content types on a list.</returns>
        public GetListContentTypesResponseGetListContentTypesResult GetListContentTypes(string listName, string contentTypeId)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            GetListContentTypesResponseGetListContentTypesResult result = null;
            try
            {
                result = this.listsProxy.GetListContentTypes(listName, contentTypeId);

                // Verify the requirements of GetListContentTypes operation.
                this.VerifyGetListContentTypesOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The ApplyContentTypeToList operation is used to apply an existing site content type to 
        /// the requested list.
        /// </summary>
        /// <param name="webUrl">This parameter is reserved and MUST be ignored.  The value MUST be an empty string if it is present.</param>
        /// <param name="contentTypeId">The identifier of the content type which will be applied to a list</param>
        /// <param name="listName">The name of the list to which the content type will be applied</param>
        /// <returns>ApplyContentTypeToList Result</returns>
        public ApplyContentTypeToListResponseApplyContentTypeToListResult ApplyContentTypeToList(string webUrl, string contentTypeId, string listName)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            ApplyContentTypeToListResponseApplyContentTypeToListResult result = null;
            try
            {
                result = this.listsProxy.ApplyContentTypeToList(webUrl, contentTypeId, listName);

                // Verify the requirements of ApplyContentTypeToList operation.
                this.VerifyApplyContentTypeToListOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The UpdateContentType operation is used to update a content type on a list.
        /// </summary>
        /// <param name="listName">The name of the list of which a content type will be updated</param>
        /// <param name="contentTypeId">The identifier of the content type which will be updated</param>
        /// <param name="contentTypeProperties">The container for properties to set on the content type</param>
        /// <param name="newFields">The new fields that will be used as parameter in UpdateContentType</param>
        /// <param name="updateFields">The fields that will be updated</param>
        /// <param name="deleteFields">The fields that will be used as parameter in UpdateContentType</param>
        /// <param name="addToView">Specifies whether the fields will be added to the default list view, "TRUE" means add to the view and "False" means not.</param>
        /// <returns>Update content type Result</returns>
        public UpdateContentTypeResponseUpdateContentTypeResult UpdateContentType(string listName, string contentTypeId, UpdateContentTypeContentTypeProperties contentTypeProperties, AddOrUpdateFieldsDefinition newFields, AddOrUpdateFieldsDefinition updateFields, DeleteFieldsDefinition deleteFields, string addToView)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            UpdateContentTypeResponseUpdateContentTypeResult result = null;
            try
            {
                result = this.listsProxy.UpdateContentType(listName, contentTypeId, contentTypeProperties, newFields, updateFields, deleteFields, addToView);

                // Verify the requirements of the transport.
                this.VerifyTransportRequirements();

                // Verify the requirements of UpdateContentType operation.
                this.VerifyUpdateContentTypeOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The UpdateContentTypeXmlDocument operation is used to update the XML document of a list content type.
        /// </summary>
        /// <param name="listName">The name of the list of which a content type xml document will be updated</param>
        /// <param name="contentTypeId">The identifier of the content type of the xml document will be updated.</param>
        /// <param name="newDocument">The XML document to be added to the content type XML document collection</param>
        /// <returns>Update Content Type Xml Document Result</returns>
        public System.Xml.XPath.IXPathNavigable UpdateContentTypeXmlDocument(string listName, string contentTypeId, System.Xml.XmlNode newDocument)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            System.Xml.XmlNode result = null;
            try
            {
                result = this.listsProxy.UpdateContentTypeXmlDocument(listName, contentTypeId, newDocument);

                // Verify the requirements of UpdateContentTypeXmlDocument operation.
                this.VerifyUpdateContentTypeXmlDocumentOperation(result);

                // Verify the requirements of the transport.
                this.VerifyTransportRequirements();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The UpdateContentTypesXmlDocument operation is used to update XML Document properties of the content type collection on a list.
        /// </summary>
        /// <param name="listName">The name of the list of which some content type xml documents will be updated</param>
        /// <param name="newDocument">The container element for a list of content type and XML document to update</param>
        /// <returns>UpdateContentTypesXmlDocument Result</returns>
        public UpdateContentTypesXmlDocumentResponseUpdateContentTypesXmlDocumentResult UpdateContentTypesXmlDocument(string listName, UpdateContentTypesXmlDocumentNewDocument newDocument)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            UpdateContentTypesXmlDocumentResponseUpdateContentTypesXmlDocumentResult result = null;
            try
            {
                result = this.listsProxy.UpdateContentTypesXmlDocument(listName, newDocument);

                // Verify the requirements of the UpdateContentTypesXmlDocument operation.
                this.VerifyUpdateContentTypesXmlDocumentOperation(result);

                // Verify the requirements of the transport.
                this.VerifyTransportRequirements();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// AddDiscussionBoardItem  operation is used to add new discussion items to a specified 
        /// Discussion Board.
        /// </summary>
        /// <param name="listName">The name of the discussion board in which the new item will be added</param>
        /// <param name="message">The message to be added to the discussion board. The message MUST be in MIME format and then Base64 encoded</param>
        /// <returns>AddDiscussionBoardItem Result</returns>
        public AddDiscussionBoardItemResponseAddDiscussionBoardItemResult AddDiscussionBoardItem(string listName, byte[] message)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            AddDiscussionBoardItemResponseAddDiscussionBoardItemResult result = null;
            try
            {
                result = this.listsProxy.AddDiscussionBoardItem(listName, message);

                // Verify the requirements of the AddDiscussionBoardItem operation.
                this.VerifyAddDiscussionBoardItemOperation(result);

                // Verify the requirements of the transport.
                this.VerifyTransportRequirements();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The DeleteList operation is used to delete the specified list from the specified site.
        /// </summary>
        /// <param name="listName">The name of the list which will be deleted</param>
        public void DeleteList(string listName)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            try
            {
                this.listsProxy.DeleteList(listName);

                // Verify the requirements of the DeleteList operation.
                this.VerifyDeleteListOperation();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }
        }

        /// <summary>
        /// The GetList operation is used to retrieve properties and fields for a specified list.
        /// </summary>
        /// <param name="listName">The name of the list from which information will be got</param>
        /// <returns>GetList Result</returns>
        public GetListResponseGetListResult GetList(string listName)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            GetListResponseGetListResult result = null;
            try
            {
                result = this.listsProxy.GetList(listName);

                // Verify the requirements of the GetList operation.
                this.VerifyGetListOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// This operation is used to retrieve properties and fields for a specified list and a view.
        /// </summary>
        /// <param name="listName">The name of the list for which information and view will be got.</param>
        /// <param name="viewName">The GUID refers to a view of the list</param>
        /// <returns>GetListAndView Result</returns>
        public GetListAndViewResponseGetListAndViewResult GetListAndView(string listName, string viewName)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            GetListAndViewResponseGetListAndViewResult result = null;
            try
            {
                result = this.listsProxy.GetListAndView(listName, viewName);

                // Verify the requirements of the GetListAndView operation.
                this.VerifyGetListAndViewOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// This operation is used to retrieve information about all the lists on the current site.
        /// </summary>
        /// <returns>GetListCollection Result</returns>
        public GetListCollectionResponseGetListCollectionResult GetListCollection()
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            GetListCollectionResponseGetListCollectionResult result = null;
            try
            {
                result = this.listsProxy.GetListCollection();

                // Verify the requirements of the GetListCollection operation.
                this.VerifyGetListCollectionOperation(result);

                // Verify the requirements of the transport.
                this.VerifyTransportRequirements();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The UpdateListItems operation is used to insert, update, and delete to specified list items in a list.
        /// </summary>
        /// <param name="listName">The name of the list for which list items will be updated</param>
        /// <param name="updates">Specifies the operations to perform on a list item</param>
        /// <returns>return the updated list items</returns>
        public UpdateListItemsResponseUpdateListItemsResult UpdateListItems(string listName, UpdateListItemsUpdates updates)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            UpdateListItemsResponseUpdateListItemsResult result = null;
            try
            {
                result = this.listsProxy.UpdateListItems(listName, updates);

                // Verify the requirements of the UpdateListItems operation.
                this.VerifyUpdateListItemsOperation(result, updates);

                // Verify the requirements of the transport.
                this.VerifyTransportRequirements();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The AddListFromFeature operation is used to add a new list to the specified site based 
        /// on the specified template and feature.
        /// </summary>
        /// <param name="listName">The name of the new list to be added</param>
        /// <param name="description">:  A string that is the description of the list to be created</param>
        /// <param name="featureID">The identifier of the feature's GUID</param>
        /// <param name="templateID">The list template identifier of a template that is already installed on the site</param>
        /// <returns>update list items</returns>
        public AddListFromFeatureResponseAddListFromFeatureResult AddListFromFeature(string listName, string description, string featureID, int templateID)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            AddListFromFeatureResponseAddListFromFeatureResult result = null;
            try
            {
                result = this.listsProxy.AddListFromFeature(listName, description, featureID, templateID);

                // Verify the requirements of AddListFromFeature operation.
                this.VerifyAddListFromFeatureOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// This operation is used to retrieve details about list items in a list that satisfy specified criteria.
        /// </summary>
        /// <param name="listName">The name of the list from which item changes will be got</param>
        /// <param name="viewName">The GUID refers to a view of the list</param>
        /// <param name="query">The query to determine which records from the list are to be returned </param>
        /// <param name="viewFields">Specifies which fields of the list item should be returned</param>
        /// <param name="rowLimit">Specifies the maximum number of rows of data to return in the response</param>
        /// <param name="queryOptions">Specifies various options for modifying the query</param>
        /// <param name="webID">The GUID of the site that contains the list. If not specified, the default Web site based on the SOAP request is used</param>
        /// <returns>Get list item result</returns>
        public GetListItemsResponseGetListItemsResult GetListItems(string listName, string viewName, GetListItemsQuery query, CamlViewFields viewFields, string rowLimit, CamlQueryOptions queryOptions, string webID)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            GetListItemsResponseGetListItemsResult result = null;
            try
            {
                result = this.listsProxy.GetListItems(listName, viewName, query, viewFields, rowLimit, queryOptions, webID);

                // Verify the requirements of the GetListItems operation.
                this.VerifyGetListItemsOperation(result, queryOptions, viewFields);

                // Verify the requirements of the transport.
                this.VerifyTransportRequirements();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The UpdateList operation is used to update list properties and add, remove, or update fields.
        /// </summary>
        /// <param name="listName">The name of the list which will be updated</param>
        /// <param name="listProperties">The properties of the specified list</param>
        /// <param name="newFields">new fields to be added to the list</param>
        /// <param name="updateFields">the fields to be updated in the list</param>
        /// <param name="deleteFields">the fields to be deleted from the list.</param>
        /// <param name="listVersion">A string represents an integer value that specifies the current version of the list</param>
        /// <returns>UpdateList Result</returns>
        public UpdateListResponseUpdateListResult UpdateList(string listName, UpdateListListProperties listProperties, UpdateListFieldsRequest newFields, UpdateListFieldsRequest updateFields, UpdateListFieldsRequest deleteFields, string listVersion)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            UpdateListResponseUpdateListResult result = null;
            try
            {
                result = this.listsProxy.UpdateList(listName, listProperties, newFields, updateFields, deleteFields, listVersion);

                // Verify the requirements of the UpdateList operation.
                this.VerifyUpdateListOperation(result);

                // Verify the requirements of the transport.
                this.VerifyTransportRequirements();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The GetVersionCollection operation is used to get version information for a specified field of a list.Get VersionCollection 
        /// </summary>
        /// <param name="strListID">The identifier of the list whose change versions will be got</param>
        /// <param name="strListItemID">The identifier of the list item whose change versions will be got</param>
        /// <param name="strFieldName">The name of the field whose value version will be got</param>
        /// <returns>return the version</returns>
        public GetVersionCollectionResponseGetVersionCollectionResult GetVersionCollection(string strListID, string strListItemID, string strFieldName)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            GetVersionCollectionResponseGetVersionCollectionResult result = null;
            try
            {
                result = this.listsProxy.GetVersionCollection(strListID, strListItemID, strFieldName);

                // Verify the requirements of the GetVersionCollection operation.
                this.VerifyGetVersionCollectionOperation(result);

                // Verify the requirements of the transport.
                this.VerifyTransportRequirements();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The Undo CheckOut operation is used to undo the checkout of the specified file in a document library.
        /// </summary>
        /// <param name="pageUrl">The parameter specifying where will execute the undo check out</param>
        /// <returns>A return value represents the operation result, True means UndoCheckOut successfully, false means protocol SUT fail to perform UndoCheckOut operation.</returns>
        public bool UndoCheckOut(string pageUrl)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            bool result = false;
            try
            {
                result = this.listsProxy.UndoCheckOut(pageUrl);

                // Verify the requirements of UndoCheckOut operation.
                this.VerifyUndoCheckOutOperation();

                // Verify the requirements of the transport.
                this.VerifyTransportRequirements();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The GetListItemChangesSinceToken operation is used to return changes made to a specified list after the event 
        /// expressed by the change token, if specified, or to return all the list items in the list.
        /// </summary>
        /// <param name="listName">The name of the list from which version collection will be got</param>
        /// <param name="viewName">The GUID refers to a view of the list</param>
        /// <param name="query">The query to determine which records from the list are to be 
        /// returned and the order in which they will be returned</param>
        /// <param name="viewFields">Specifies which fields of the list item will be returned</param>
        /// <param name="rowLimit">Indicate the maximum number of rows of data to return</param>
        /// <param name="queryOptions">Specifies various options for modifying the query</param>
        /// <param name="changeToken">Assigned a string comprising a token returned by a previous 
        /// call to this operation.</param>
        /// <param name="contains">Specifies a value to search for</param>
        /// <returns>A return value represent the list item changes since the specified token</returns>
        public GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult GetListItemChangesSinceToken(string listName, string viewName, GetListItemChangesSinceTokenQuery query, CamlViewFields viewFields, string rowLimit, CamlQueryOptions queryOptions, string changeToken, CamlContains contains)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult result = null;
            try
            {
                result = this.listsProxy.GetListItemChangesSinceToken(listName, viewName, query, viewFields, rowLimit, queryOptions, changeToken, contains);

                // Verify the requirements of the GetListItemChangesSinceToken operation.
                this.VerifyGetListItemChangesSinceTokenOperation(
                    result,
                    queryOptions,
                    viewFields);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The AddWikiPage operation MUST NOT be used. It is not supported.
        /// </summary>
        /// <param name="strListName">The GUID or the title of the Wiki library</param>
        /// <param name="listRelPageUrl">The URL of the page to be added. It's a relative URL to 
        /// the Wiki library's root folder</param>
        /// <param name="wikiContent">describe Wiki content</param>
        /// <returns>The value indicating the error code in the thrown SOAP fault</returns>
        public AddWikiPageResponseAddWikiPageResult AddWikiPage(string strListName, string listRelPageUrl, string wikiContent)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            AddWikiPageResponseAddWikiPageResult result = null;
            try
            {
                result = this.listsProxy.AddWikiPage(strListName, listRelPageUrl, wikiContent);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The GetListContentTypesAndProperties operation is used to retrieve all content types from a list, and specified properties from the list and site property bags
        /// </summary>
        /// <param name="listName">The name will be added.</param>
        /// <param name="contentTypeId">The identifier of the content type which will be used as a match criterion.</param>
        /// <param name="propertyPrefix">The prefix of the requested property keys.</param>
        /// <param name="includeWebProperties">A Boolean value indicating will be returned.</param>
        /// <param name="includeWebPropertiesSpecified">The properties and files from the site property bag will be specified.</param>
        /// <returns>Return the types and properties.</returns>
        public GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResult GetListContentTypesAndProperties(string listName, string contentTypeId, string propertyPrefix, bool includeWebProperties, [System.Xml.Serialization.XmlIgnoreAttribute()] bool includeWebPropertiesSpecified)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResult result = null;
            try
            {
                result = this.listsProxy.GetListContentTypesAndProperties(listName, contentTypeId, propertyPrefix, includeWebProperties, includeWebPropertiesSpecified);

                // Verify the requirements of the GetListContentTypesAndProperties operation
                this.VerifyGetListContentTypesAndPropertiesOperation(result);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The GetListItemChangesWithKnowledge operation is used to get changes made to a specified list after the event expressed by the knowledge parameter, 
        /// if specified, or to return all the list items in the list.
        /// </summary>
        /// <param name="listName">Get list name</param>
        /// <param name="viewName">The GUID of a view of the list</param>
        /// <param name="query">Query the list</param>
        /// <param name="viewFields">Specifies which fields of the list item should be returned</param>
        /// <param name="rowLimit">The maximum number of rows of data to return</param>
        /// <param name="queryOptions">Specifies various options for modifying the query</param>
        /// <param name="syncScope">This parameter MUST be null or empty</param>
        /// <param name="knowledge">Specifies the knowledge data structure in XML format</param>
        /// <param name="contains">Specifies a value to search for</param>
        /// <returns>The value indicating the error code in the thrown SOAP fault.</returns>
        public GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult GetListItemChangesWithKnowledge(
            string listName, 
            string viewName,
            GetListItemChangesWithKnowledgeQuery query, 
            CamlViewFields viewFields,
            string rowLimit,
            CamlQueryOptions queryOptions, 
            string syncScope,
            GetListItemChangesWithKnowledgeKnowledge knowledge, 
            CamlContains contains)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult result = null;
            try
            {
                result = this.listsProxy.GetListItemChangesWithKnowledge(listName, viewName, query, viewFields, rowLimit, queryOptions, syncScope, knowledge, contains);

                // Verify the requirements of the GetListItemChangesWithKnowledge operation.
                this.VerifyGetListItemChangesWithKnowledgeOperation(result, queryOptions, viewFields);
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        /// <summary>
        /// The UpdateListItemsWithKnowledge operation is used to operation is used to insert, update, and delete specified list items in a list.
        /// </summary>
        /// <param name="listName">The name of the list for which list items will be updated</param>
        /// <param name="updates">A parameter represents the operations to perform on a list item</param>
        /// <param name="syncScope">A parameter is reserved and MUST be ignored</param>
        /// <param name="knowledge">Specifies a value to search for</param>
        /// <returns>Return items with knowledge specified</returns>
        public UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult UpdateListItemsWithKnowledge(
            string listName,
            UpdateListItemsWithKnowledgeUpdates updates, 
            string syncScope, 
            System.Xml.XmlNode knowledge)
        {
            this.Site.Assert.IsNotNull(this.listsProxy, "The Proxy instance should not be NULL. If assert failed, the adapter need to be initialized");

            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult result = null;
            try
            {
                result = this.listsProxy.UpdateListItemsWithKnowledge(listName, updates, syncScope, knowledge);

                // Verify the requirements of the UpdateListItems operation.
                this.VerifyUpdateListItemsWithKnowledgeOperation(result, updates);

                // Verify the requirements of the transport.
                this.VerifyTransportRequirements();
            }
            catch (XmlSchemaValidationException exp)
            {
                // Log the errors and warnings
                this.LogSchemaValidationErrors();

                this.Site.Assert.Fail(exp.Message);
            }
            catch (SoapException)
            {
                this.VerifySoapExceptionFault();
                throw;
            }

            return result;
        }

        #endregion

        #region Override methods

        /// <summary>
        /// The Overridden Initialize method
        /// </summary>
        /// <param name="testSite">The ITestSite member of ManagedAdapterBase</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            // Set the protocol name of current test suite
            testSite.DefaultProtocolDocShortName = "MS-LISTSWS";

            // Initialize the ListsSoap.
            this.listsProxy = Proxy.CreateProxy<ListsSoap>(testSite, true, true, true);

            // Initialize the Helper.
            AdapterHelper.Initialize(testSite);

            // Load Common configuration
            this.LoadCommonConfiguration();

            // Load SHOULDMAY configuration 
            this.LoadCurrentSutSHOULDMAYConfiguration();

            Site.DefaultProtocolDocShortName = AdapterHelper.DefaultProtocolDocShortName;
            this.listsProxy.Url = this.GetTargetServiceUrl();
            string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            string password = Common.GetConfigurationPropertyValue("Password", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.listsProxy.Credentials = new NetworkCredential(userName, password, domain);

            Common.AcceptServerCertificate();
            this.SetSoapVersion();

            // Configure the service timeout.
            string soapTimeOut = Common.GetConfigurationPropertyValue("ServiceTimeOut", this.Site);

            // 60000 means the configure SOAP Timeout is in minute.
            this.listsProxy.Timeout = Convert.ToInt32(soapTimeOut) * 60000;
        }

        #endregion

        #region Private helper methods

        /// <summary>
        /// A method used to load Common Configuration
        /// </summary>
        private void LoadCommonConfiguration()
        {
            // Merge the common configuration into local configuration
            string conmmonConfigFileName = AdapterHelper.GetValueFromConfig("CommonConfigurationFileName");

            // Execute the merge the common configuration
            Common.MergeGlobalConfig(conmmonConfigFileName, this.Site);

            Common.CheckCommonProperties(this.Site, true);
        }

        /// <summary>
        /// A method used to load SHOULDMAY Configuration according to the current SUT version
        /// </summary>
        private void LoadCurrentSutSHOULDMAYConfiguration()
        {
            Common.MergeSHOULDMAYConfig(this.Site);
        }

        /// <summary>
        /// Set the SOAP version according to the SoapVersion property.
        /// </summary>
        private void SetSoapVersion()
        {
            SoapVersion soapVersion = Common.GetConfigurationPropertyValue<SoapVersion>("SoapVersion", this.Site);

            switch (soapVersion)
            {
                case SoapVersion.SOAP11:
                    {
                        this.listsProxy.SoapVersion = SoapProtocolVersion.Soap11;
                        break;
                    }

                default:
                    {
                        this.listsProxy.SoapVersion = SoapProtocolVersion.Soap12;
                        break;
                    }
            }
        }

        /// <summary>
        /// A method used to Get target service fully qualified URL, it indicates which site the test suite will run on.
        /// </summary>
        /// <returns>A return value represents the target service fully qualified URL</returns>
        private string GetTargetServiceUrl()
        {
            string fullyServiceURL = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);
            return fullyServiceURL;
        }

        /// <summary>
        /// A method is used to log all the schema validation errors and warnings.
        /// </summary>
        private void LogSchemaValidationErrors()
        {
            // Log all the schema validation warnings if exist
            foreach (ValidationEventArgs warning in SchemaValidation.XmlValidationWarnings)
            {
                this.Site.Log.Add(
                                     LogEntryKind.Warning,
                                    "Schema validation Warning:{0} occurs in the position ({1},{2})",
                                    warning.Exception.Message,
                                    warning.Exception.LineNumber,
                                    warning.Exception.LinePosition);
            }

            // Log all the schema validation errors if exist
            foreach (ValidationEventArgs error in SchemaValidation.XmlValidationErrors)
            {
                this.Site.Log.Add(
                                    LogEntryKind.TestError,
                                    "Schema validation Error:{0} occurs in the position ({1},{2})",
                                    error.Exception.Message,
                                    error.Exception.LineNumber,
                                    error.Exception.LinePosition);
            }
        }

        #endregion
    }
}