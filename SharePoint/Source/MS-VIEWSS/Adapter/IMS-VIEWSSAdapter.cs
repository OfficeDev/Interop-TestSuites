//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_VIEWSS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-VIEWSS protocol adapter.
    /// </summary>
    public interface IMS_VIEWSSAdapter : IAdapter
    {
        /// <summary>
        /// This operation is used to create a list view for the specified list.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        /// <param name="viewFields">Specify the fields included in a list view.</param>
        /// <param name="query">Include the information that affects how a list view displays the data.</param>
        /// <param name="rowLimit">Specify whether a list supports displaying items page-by-page, and the count of items a list view displays per page.</param>
        /// <param name="type">Specify the type of a list view.</param>
        /// <param name="makeViewDefault">Specify whether to make the list view the default list view for the specified list.</param>
        /// <returns>The result returns a list view that the type is BriefViewDefinition if the operation succeeds.</returns>
        AddViewResponseAddViewResult AddView(
            string listName,
            string viewName,
            AddViewViewFields viewFields,
            AddViewQuery query,
            AddViewRowLimit rowLimit,
            string type,
            bool makeViewDefault);

        /// <summary>
        /// This operation is used to delete the specified list view of the specified list.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        void DeleteView(string listName, string viewName);

        /// <summary>
        /// This operation is used to obtain details of a specified list view of the specified list.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        /// <returns>The result returns a list view that the type is BriefViewDefinition if the operation succeeds.</returns>
        GetViewResponseGetViewResult GetView(string listName, string viewName);

        /// <summary>
        /// This operation is used to retrieve the collection of list views of a specified list.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <returns>The result returns a collection of View elements of the specified list if the operation succeeds.</returns>
        GetViewCollectionResponseGetViewCollectionResult GetViewCollection(string listName);

        /// <summary>
        /// This operation is used to obtain details of a specified list view of the specified list, including display properties in CAML and HTML.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        /// <returns>The result returns the details of a specified list view of the specified list if the operation succeeds.</returns>
        GetViewHtmlResponseGetViewHtmlResult GetViewHtml(string listName, string viewName);

        /// <summary>
        /// This operation is used to update the specified list view, without the display properties.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        /// <param name="viewProperties">Specify the properties of a list view on the server.</param>
        /// <param name="query">Include the information that affects how a list view displays the data.</param>
        /// <param name="viewFields">Specify the fields included in a list view.</param>
        /// <param name="aggregations">The type of the aggregation.</param> 
        /// <param name="formats">Specify the row and column formatting of a list view.</param>
        /// <param name="rowLimit">Specify whether a list supports displaying items page-by-page, and the count of items a list view displays per page.</param>
        /// <returns>The result returns a list view that the type is BriefViewDefinition, if the operation succeeds.</returns>
        UpdateViewResponseUpdateViewResult UpdateView(
            string listName,
            string viewName,
            UpdateViewViewProperties viewProperties,
            UpdateViewQuery query,
            UpdateViewViewFields viewFields,
            UpdateViewAggregations aggregations,
            UpdateViewFormats formats,
            UpdateViewRowLimit rowLimit);

        /// <summary>
        /// This operation is used to update a list view for a specified list, including display properties in CAML and HTML.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        /// <param name="viewProperties">Specify the properties of a list view on the server.</param>
        /// <param name="toolbar">Specify the rendering of the toolbar of a list.</param>
        /// <param name="viewHeader">Specify the rendering of the header, or the top of a list view page.</param>
        /// <param name="viewBody">Specify the rendering of the main, or the middle portion of a list view page.</param>
        /// <param name="viewFooter">Specify the rendering of the footer, or the bottom of a list view page.</param>
        /// <param name="viewEmpty">Specify the message to be displayed when no items are in a list view.</param>
        /// <param name="rowLimitExceeded">Specify rendering of additional items when the count of items exceeds the value.</param>
        /// <param name="query">Include the information that affects how a list view displays the data.</param>
        /// <param name="viewFields">Specify the fields included in a list view.</param>
        /// <param name="aggregations">The type of the aggregation.</param> 
        /// <param name="formats">Specify the row and column formatting of a list view.</param>
        /// <param name="rowLimit">Specify whether a list supports displaying items page-by-page, and the count of items a list view displays per page.</param>
        /// <returns>The result returns a View that the type is ViewDefinition if the operation succeeds.</returns>
        UpdateViewHtmlResponseUpdateViewHtmlResult UpdateViewHtml(
            string listName, 
            string viewName, 
            UpdateViewHtmlViewProperties viewProperties, 
            UpdateViewHtmlToolbar toolbar, 
            UpdateViewHtmlViewHeader viewHeader, 
            UpdateViewHtmlViewBody viewBody, 
            UpdateViewHtmlViewFooter viewFooter, 
            UpdateViewHtmlViewEmpty viewEmpty, 
            UpdateViewHtmlRowLimitExceeded rowLimitExceeded, 
            UpdateViewHtmlQuery query, 
            UpdateViewHtmlViewFields viewFields, 
            UpdateViewHtmlAggregations aggregations, 
            UpdateViewHtmlFormats formats, 
            UpdateViewHtmlRowLimit rowLimit);

        /// <summary>
        /// This operation is used to obtain details of a specified list view of the specified list, including display properties in CAML and HTML.
        /// </summary>
        /// <param name="listName">Specify a list on the server.</param>
        /// <param name="viewName">Specify a list view on the server.</param>
        /// <param name="viewProperties">Specify the properties of a list view on the server.</param>
        /// <param name="toolbar">Specify the rendering of the toolbar of a list.</param>
        /// <param name="viewHeader">Specify the rendering of the header, or the top of a list view page.</param>
        /// <param name="viewBody">Specify the rendering of the main, or the middle portion of a list view page.</param>
        /// <param name="viewFooter">Specify the rendering of the footer, or the bottom of a list view page.</param>
        /// <param name="viewEmpty">Specify the message to be displayed when no items are in a list view.</param>
        /// <param name="rowLimitExceeded">Specify rendering of additional items when the number of items exceeds the value.</param>
        /// <param name="query">Include the information that affects how a list view displays the data.</param>
        /// <param name="viewFields">Specify the fields included in a list view.</param>
        /// <param name="aggregations">The type of the aggregation.</param> 
        /// <param name="formats">Specify the row and column formatting of a list view.</param>
        /// <param name="rowLimit">Specify whether a list supports displaying items page-by-page, and the count of items a list view displays per page.</param>
        /// <param name="openApplicationExtension">Specify what kind of application to use to edit the view.</param>
        /// <returns>The result returns a View that the type is ViewDefinition if the operation succeeds.</returns>
        UpdateViewHtml2ResponseUpdateViewHtml2Result UpdateViewHtml2(
            string listName, 
            string viewName, 
            UpdateViewHtml2ViewProperties viewProperties, 
            UpdateViewHtml2Toolbar toolbar, 
            UpdateViewHtml2ViewHeader viewHeader, 
            UpdateViewHtml2ViewBody viewBody, 
            UpdateViewHtml2ViewFooter viewFooter, 
            UpdateViewHtml2ViewEmpty viewEmpty, 
            UpdateViewHtml2RowLimitExceeded rowLimitExceeded, 
            UpdateViewHtml2Query query, 
            UpdateViewHtml2ViewFields viewFields, 
            UpdateViewHtml2Aggregations aggregations, 
            UpdateViewHtml2Formats formats, 
            UpdateViewHtml2RowLimit rowLimit, 
            string openApplicationExtension);
    }
}
