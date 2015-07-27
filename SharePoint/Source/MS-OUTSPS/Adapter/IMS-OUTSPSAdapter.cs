//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OUTSPS
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This adapter interface of MS-OUTSPS
    /// </summary>
    public interface IMS_OUTSPSAdapter : IAdapter
    {
        /// <summary>
        /// Add a list in the current site based on the specified name, description, and list template identifier.
        /// </summary>
        /// <param name="listName">The title of the list which will be added.</param>
        /// <param name="description">Text which will be set as description of newly created list.</param>
        /// <param name="templateId">The template ID used to create this list.</param>
        /// <returns>Returns the AddList result.</returns>
        AddListResponseAddListResult AddList(string listName, string description, int templateId);

        /// <summary>
        /// The AddAttachment operation is used to add an attachment to the specified list item in the specified list.
        /// </summary>
        /// <param name="listName">The GUID or the list title of the list in which the list item to add attachment.</param>
        /// <param name="listItemId">The id of the list item in which the attachment will be added.</param>
        /// <param name="fileName">The name of the file being added as an attachment.</param>
        /// <param name="attachment">Content of the attachment file (byte array).</param>
        /// <returns>The URL of the newly added attachment.</returns>
        string AddAttachment(string listName, string listItemId, string fileName, byte[] attachment);

        /// <summary>
        /// The DeleteAttachment operation is used to remove the attachment from the specified list 
        /// item in the specified list.
        /// </summary>
        /// <param name="listName">The name of the list in which the list item to delete existing attachment.</param>
        /// <param name="listItemId">The id of the list item from which the attachment will be deleted.</param>
        /// <param name="url">Absolute URL of the attachment that should be deleted.</param>
        void DeleteAttachment(string listName, string listItemId, string url);

        /// <summary>
        /// The GetListItemChanges operation is used to retrieve the list items that have been inserted or updated
        /// since the specified date and time and matching the specified filter criteria.
        /// </summary>
        /// <param name="listName">The name of the list from which the list item changes will be got.</param>
        /// <param name="viewFields">Indicates which fields of the list item SHOULD be returned</param>
        /// <param name="since">The date and time to start retrieving changes in the list
        /// If the parameter is null, protocol server should return all list items
        /// If the date that is passed in is not in UTC format, protocol server will use protocol server's local time zone and convert it to UTC time</param>
        /// <param name="contains">Restricts the results returned by giving a specific value to be searched for in the specified list item field</param>
        /// <returns>Return the list item change result</returns>
        GetListItemChangesResponseGetListItemChangesResult GetListItemChanges(string listName, CamlViewFields viewFields, string since, CamlContains contains);

        /// <summary>
        /// The GetAttachmentCollection operation is used to retrieve information about all the lists on the current site.
        /// </summary>
        /// <param name="listName">A parameter represents the list name or GUID for returning the result.</param>
        /// <param name="listItemId">A parameter represents the identifier of the content type which will be collected.</param>
        /// <returns>Return attachment collection result.</returns>
        GetAttachmentCollectionResponseGetAttachmentCollectionResult GetAttachmentCollection(string listName, string listItemId);

        /// <summary>
        /// AddDiscussionBoardItem operation is used to add new discussion items to a specified discussion board.
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
        /// <returns>A return value represents the list definition.</returns>
        GetListResponseGetListResult GetList(string listName);

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
        /// The UpdateListItems operation is used to insert, update, and delete to specified list items in a list.
        /// </summary>
        /// <param name="listName">The name of the list for which list items will be updated</param>
        /// <param name="updates">Specifies the operations to perform on a list item</param>
        /// <returns>return the updated list items</returns>
        UpdateListItemsResponseUpdateListItemsResult UpdateListItems(string listName, UpdateListItemsUpdates updates);

        /// <summary>
        ///  This operation used to get resource data Over HTTP protocol directly.
        /// </summary>
        /// <param name="requestResourceUrl">A parameter represents the resource where get data over HTTP protocol.</param>
        /// <param name="translateHeaderValue">A parameter represents the translate header which is used in HTTP request.</param>
        /// <returns>A return value represents the data get from the specified resource.</returns>
        byte[] HTTPGET(Uri requestResourceUrl, string translateHeaderValue);

        /// <summary>
        /// This operation used to put content data Over HTTP protocol directly.
        /// </summary>
        /// <param name="requestResourceUrl">A parameter represents the resource where put the data over HTTP protocol.</param>
        /// <param name="ifmatchHeader">>A parameter represents the IF-MATCH header which is used in HTTP request.</param>
        /// <param name="contentData">>A parameter represents the content data which is put to the SUT.</param>
        void HTTPPUT(Uri requestResourceUrl, string ifmatchHeader, byte[] contentData);

        /// <summary>
        /// A method used to update list properties and add, remove, or update fields.
        /// </summary>
        /// <param name="listName">A parameter represents the name of the list which will be updated.</param>
        /// <param name="listProperties">A parameter represents the properties of the specified list.</param>
        /// <param name="newFields">A parameter represents new fields which are added to the list.</param>
        /// <param name="updateFields">A parameter represents the fields which are updated in the list.</param>
        /// <param name="deleteFields">A parameter represents the fields which are deleted from the list.</param>
        /// <param name="listVersion">A parameter represents an integer format string that specifies the current version of the list.</param>
        /// <returns>A return value represents the actual update result.</returns>
        UpdateListResponseUpdateListResult UpdateList(string listName, UpdateListListProperties listProperties, UpdateListFieldsRequest newFields, UpdateListFieldsRequest updateFields, UpdateListFieldsRequest deleteFields, string listVersion);
    }
}