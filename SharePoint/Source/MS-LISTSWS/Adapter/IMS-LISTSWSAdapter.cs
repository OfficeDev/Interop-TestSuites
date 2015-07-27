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
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// Adapter interface
    /// </summary>
    public interface IMS_LISTSWSAdapter : IAdapter
    {
        /// <summary>
        /// Gets or sets the security credentials.
        /// </summary>
        ICredentials Credentials { get; set; }

        /// <summary>
        /// Add a list in the current site based on the specified name, description, and list template identifier.
        /// </summary>
        /// <param name="listName">The title of the list which will be added.</param>
        /// <param name="description">Text which will be set as description of newly created list.</param>
        /// <param name="templateID">The template ID used to create this list.</param>
        /// <returns>Returns the AddList result.</returns>
        AddListResponseAddListResult AddList(string listName, string description, int templateID);

        /// <summary>
        /// Check in a document item to a document library.
        /// </summary>
        /// <param name="pageUrl">The Url specifies which file will be check in. </param>
        /// <param name="comment">The comments of the check in.</param>
        /// <param name="checkinType">A string representation the check type value <see cref="CheckInTypeValue"/> </param>
        /// <returns>Returns True indicating CheckInFile was successful</returns>
        bool CheckInFile(string pageUrl, string comment, string checkinType);

        /// <summary>
        /// The CheckOutFile operation is used to check out a document in a document library.
        /// </summary>
        /// <param name="pageUrl">The Url specifies which file will be check out.</param>
        /// <param name="checkoutToLocal">"TRUE" means to keep a local version for offline editing</param>
        /// <param name="lastModified">A string in date format that represents the date and time of the last modification by the site to the file; for example, "20 Jun 1982 12:00:00 GMT</param>
        /// <returns>Returns True indicating CheckOutFile was successful</returns>
        bool CheckOutFile(string pageUrl, string checkoutToLocal, string lastModified);

        /// <summary>
        /// The AddAttachment operation is used to add an attachment to the specified list item in the specified list.
        /// </summary>
        /// <param name="listName">The GUID or the list title of the list in which the list item to add attachment.</param>
        /// <param name="listItemID">The id of the list item in which the attachment will be added.</param>
        /// <param name="fileName">The name of the file being added as an attachment.</param>
        /// <param name="attachment">Content of the attachment file (byte array).</param>
        /// <returns>The URL of the newly added attachment.</returns>
        string AddAttachment(string listName, string listItemID, string fileName, byte[] attachment);

        /// <summary>
        /// The DeleteAttachment operation is used to remove the attachment from the specified list 
        /// item in the specified list.
        /// </summary>
        /// <param name="listName">The name of the list in which the list item to delete existing attachment.</param>
        /// <param name="listItemID">The id of the list item from which the attachment will be deleted.</param>
        /// <param name="url">Absolute URL of the attachment that should be deleted.</param>
        void DeleteAttachment(string listName, string listItemID, string url);

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
        /// <returns>Return the get list item change result.</returns>
        GetListItemChangesResponseGetListItemChangesResult GetListItemChanges(string listName, CamlViewFields viewFields, string since, CamlContains contains);

        /// <summary>
        /// The GetAttachmentCollection operation is used to retrieve information about all the lists on the current site.
        /// </summary>
        /// <param name="listName">list name or GUID for returning the result.</param>
        /// <param name="listItemID">The identifier of the content type which will be collected.</param>
        /// <returns>Return attachment collection result.</returns>
        GetAttachmentCollectionResponseGetAttachmentCollectionResult GetAttachmentCollection(string listName, string listItemID);

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
        string CreateContentType(string listName, string displayName, string parentType, AddOrUpdateFieldsDefinition fields, CreateContentTypeContentTypeProperties contentTypeProperties, string addToView);

        /// <summary>
        /// The DeleteContentType operation is used to remove the association between the specified list and the specified content type.
        /// </summary>
        /// <param name="listName">The name of the list in which a content type will be deleted</param>
        /// <param name="contentTypeId">The identifier of the content type which will be deleted</param>
        /// <returns>return delete content type </returns>
        DeleteContentTypeResponseDeleteContentTypeResult DeleteContentType(string listName, string contentTypeId);

        /// <summary>
        /// The DeleteContentTypeXmlDocument operation is used to delete an XML Document Property from XML Document collection in a content type of a list.
        /// </summary>
        /// <param name="listName">The name of the list in which a content type xml document will be deleted</param>
        /// <param name="contentTypeId">The identifier of the content type for which the xml document will be deleted</param>
        /// <param name="documentUri">The namespace URI of the XML document to remove</param>
        /// <returns>Return deleted content type XmlDocument Result</returns>
        DeleteContentTypeXmlDocumentResponseDeleteContentTypeXmlDocumentResult DeleteContentTypeXmlDocument(string listName, string contentTypeId, string documentUri);

        /// <summary>
        /// The GetListContentType operation is used to get content type data for a given content type identifier.
        /// </summary>
        /// <param name="listName">The name of the list for which a content type will be got</param>
        /// <param name="contentTypeId">The identifier of the content type of which information will be got</param>
        /// <returns>return List content type</returns>
        GetListContentTypeResponseGetListContentTypeResult GetListContentType(string listName, string contentTypeId);

        /// <summary>
        /// The GetListContentTypes operation is used to retrieve all content types from list.
        /// </summary>
        /// <param name="listName">The name of the list for which a content type will be got</param>
        /// <param name="contentTypeId">The identifier of the content type of which information will be got</param>
        /// <returns>All content types on a list.</returns>
        GetListContentTypesResponseGetListContentTypesResult GetListContentTypes(string listName, string contentTypeId);

        /// <summary>
        /// The ApplyContentTypeToList operation is used to apply an existing site content type to 
        /// the requested list.
        /// </summary>
        /// <param name="webUrl">This parameter is reserved and MUST be ignored.  The value MUST be an empty string if it is present.</param>
        /// <param name="contentTypeId">The identifier of the content type which will be applied to a list</param>
        /// <param name="listName">The name of the list to which the content type will be applied</param>
        /// <returns>ApplyContentTypeToList Result</returns>
        ApplyContentTypeToListResponseApplyContentTypeToListResult ApplyContentTypeToList(string webUrl, string contentTypeId, string listName);

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
        UpdateContentTypeResponseUpdateContentTypeResult UpdateContentType(string listName, string contentTypeId, UpdateContentTypeContentTypeProperties contentTypeProperties, AddOrUpdateFieldsDefinition newFields, AddOrUpdateFieldsDefinition updateFields, DeleteFieldsDefinition deleteFields, string addToView);

        /// <summary>
        /// The UpdateContentTypeXmlDocument operation is used to update the XML document of a list content type.
        /// </summary>
        /// <param name="listName">The name of the list of which a content type xml document will be updated</param>
        /// <param name="contentTypeId">The identifier of the content type of the xml document will be updated.</param>
        /// <param name="newDocument">The XML document to be added to the content type XML document collection</param>
        /// <returns>Update Content Type Xml Document Result</returns>
        System.Xml.XPath.IXPathNavigable UpdateContentTypeXmlDocument(string listName, string contentTypeId, System.Xml.XmlNode newDocument);

        /// <summary>
        /// The UpdateContentTypesXmlDocument operation is used to update XML Document properties of the content type collection on a list.
        /// </summary>
        /// <param name="listName">The name of the list of which some content type xml documents will be updated</param>
        /// <param name="newDocument">The container element for a list of content type and XML document to update</param>
        /// <returns>UpdateContentTypesXmlDocument Result</returns>
        UpdateContentTypesXmlDocumentResponseUpdateContentTypesXmlDocumentResult UpdateContentTypesXmlDocument(string listName, UpdateContentTypesXmlDocumentNewDocument newDocument);

        /// <summary>
        /// AddDiscussionBoardItem  operation is used to add new discussion items to a specified 
        /// Discussion Board.
        /// </summary>
        /// <param name="listName">The name of the discussion board in which the new item will be added</param>
        /// <param name="message">The message to be added to the discussion board. The message MUST be in MIME format and then Base64 encoded</param>
        /// <returns>AddDiscussionBoardItem Result</returns>
        AddDiscussionBoardItemResponseAddDiscussionBoardItemResult AddDiscussionBoardItem(string listName, byte[] message);

        /// <summary>
        /// The DeleteList operation is used to delete the specified list from the specified site.
        /// </summary>
        /// <param name="listName">The name of the list which will be deleted</param>
        void DeleteList(string listName);

        /// <summary>
        /// The GetList operation is used to retrieve properties and fields for a specified list.
        /// </summary>
        /// <param name="listName">The name of the list from which information will be got</param>
        /// <returns>GetList Result</returns>
        GetListResponseGetListResult GetList(string listName);

        /// <summary>
        /// This operation is used to retrieve properties and fields for a specified list and a view.
        /// </summary>
        /// <param name="listName">The name of the list for which information and view will be got.</param>
        /// <param name="viewName">The GUID refers to a view of the list</param>
        /// <returns>GetListAndView Result</returns>
        GetListAndViewResponseGetListAndViewResult GetListAndView(string listName, string viewName);

        /// <summary>
        /// This operation is used to retrieve information about all the lists on the current site.
        /// </summary>
        /// <returns>GetListCollection Result</returns>
        GetListCollectionResponseGetListCollectionResult GetListCollection();

        /// <summary>
        /// The UpdateListItems operation is used to insert, update, and delete to specified list items in a list.
        /// </summary>
        /// <param name="listName">The name of the list for which list items will be updated</param>
        /// <param name="updates">Specifies the operations to perform on a list item</param>
        /// <returns>return the updated list items</returns>
        UpdateListItemsResponseUpdateListItemsResult UpdateListItems(string listName, UpdateListItemsUpdates updates);

        /// <summary>
        /// The AddListFromFeature operation is used to add a new list to the specified site based 
        /// on the specified template and feature.
        /// </summary>
        /// <param name="listName">The name of the new list to be added</param>
        /// <param name="description">:  A string that is the description of the list to be created</param>
        /// <param name="featureID">The identifier of the feature's GUID</param>
        /// <param name="templateID">The list template identifier of a template that is already installed on the site</param>
        /// <returns>Add list from feature Result</returns>
        AddListFromFeatureResponseAddListFromFeatureResult AddListFromFeature(string listName, string description, string featureID, int templateID);

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
        /// <returns>return list items</returns>
        GetListItemsResponseGetListItemsResult GetListItems(string listName, string viewName, GetListItemsQuery query, CamlViewFields viewFields, string rowLimit, CamlQueryOptions queryOptions, string webID);

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
        UpdateListResponseUpdateListResult UpdateList(string listName, UpdateListListProperties listProperties, UpdateListFieldsRequest newFields, UpdateListFieldsRequest updateFields, UpdateListFieldsRequest deleteFields, string listVersion);

        /// <summary>
        /// The GetVersionCollection operation is used to get version information for a specified field of a list.Get VersionCollection 
        /// </summary>
        /// <param name="strListID">The identifier of the list whose change versions will be got</param>
        /// <param name="strListItemID">The identifier of the list item whose change versions will be got</param>
        /// <param name="strFieldName">The name of the field whose value version will be got</param>
        /// <returns>return the version</returns>
        GetVersionCollectionResponseGetVersionCollectionResult GetVersionCollection(string strListID, string strListItemID, string strFieldName);

        /// <summary>
        /// The Undo CheckOut operation is used to undo the checkout of the specified file in a document library.
        /// </summary>
        /// <param name="pageUrl">The parameter specifying where will execute the undo check out</param>
        /// <returns>A return value represents the operation result, True means UndoCheckOut successfully, false means protocol SUT fail to perform UndoCheckOut operation.</returns>
        bool UndoCheckOut(string pageUrl);

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
        GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult GetListItemChangesSinceToken(string listName, string viewName, GetListItemChangesSinceTokenQuery query, CamlViewFields viewFields, string rowLimit, CamlQueryOptions queryOptions, string changeToken, CamlContains contains);

        /// <summary>
        /// The AddWikiPage operation MUST NOT be used. It is not supported.
        /// </summary>
        /// <param name="strListName">The GUID or the title of the Wiki library</param>
        /// <param name="listRelPageUrl">The URL of the page to be added. It's a relative URL to the Wiki library's root folder</param>
        /// <param name="wikiContent">describe Wiki content</param>
        /// <returns>The value indicating the error code in the thrown SOAP fault</returns>
        AddWikiPageResponseAddWikiPageResult AddWikiPage(string strListName, string listRelPageUrl, string wikiContent);

        /// <summary>
        /// The GetListContentTypesAndProperties operation is used to retrieve all content types from a list, and specified properties from the list and site property bags
        /// </summary>
        /// <param name="listName">The name will be added.</param>
        /// <param name="contentTypeId">The identifier of the content type which will be used as a match criterion.</param>
        /// <param name="propertyPrefix">The prefix of the requested property keys.</param>
        /// <param name="includeWebProperties">A bool value indicating will be returned.</param>
        /// <param name="includeWebPropertiesSpecified">The properties and files from the site property bag will be specified.</param>
        /// <returns>Return the types and properties.</returns>
        GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResult GetListContentTypesAndProperties(string listName, string contentTypeId, string propertyPrefix, bool includeWebProperties, [System.Xml.Serialization.XmlIgnoreAttribute()] bool includeWebPropertiesSpecified);

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
        GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult GetListItemChangesWithKnowledge(string listName, string viewName, GetListItemChangesWithKnowledgeQuery query, CamlViewFields viewFields, string rowLimit, CamlQueryOptions queryOptions, string syncScope, GetListItemChangesWithKnowledgeKnowledge knowledge, CamlContains contains);

        /// <summary>
        /// The UpdateListItemsWithKnowledge operation is used to operation is used to insert, update, and delete specified list items in a list.
        /// </summary>
        /// <param name="listName">The name of the list for which list items will be updated</param>
        /// <param name="updates">A parameter represents the operations to perform on a list item</param>
        /// <param name="syncScope">A parameter is reserved and MUST be ignored</param>
        /// <param name="knowledge">Specifies a value to search for</param>
        /// <returns>Return items with knowledge specified</returns>
        UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult UpdateListItemsWithKnowledge(string listName, UpdateListItemsWithKnowledgeUpdates updates, string syncScope, System.Xml.XmlNode knowledge);
    }
}