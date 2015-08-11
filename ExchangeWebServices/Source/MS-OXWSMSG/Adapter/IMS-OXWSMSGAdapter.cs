namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-OXWSMSG.
    /// </summary>
    public interface IMS_OXWSMSGAdapter : IAdapter
    {
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT
        /// </summary>
        IXPathNavigable LastRawRequestXml { get; }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT
        /// </summary>
        IXPathNavigable LastRawResponseXml { get; }

        /// <summary>
        /// Get message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to get message objects.</param>
        /// <returns>The response message returned by GetItem operation.</returns>
        GetItemResponseType GetItem(GetItemType request);

        /// <summary>
        /// Copy message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to copy message objects.</param>
        /// <returns>The response message returned by CopyItem operation.</returns>
        CopyItemResponseType CopyItem(CopyItemType request);

        /// <summary>
        /// Create message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to create message objects.</param>
        /// <returns>The response message returned by CreateItem operation.</returns>
        CreateItemResponseType CreateItem(CreateItemType request);

        /// <summary>
        /// Delete message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to delete message objects.</param>
        /// <returns>The response message returned by DeleteItem operation.</returns>
        DeleteItemResponseType DeleteItem(DeleteItemType request);

        /// <summary>
        /// Move message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to move message objects.</param>
        /// <returns>The response message returned by MoveItem operation.</returns>
        MoveItemResponseType MoveItem(MoveItemType request);

        /// <summary>
        /// Update message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to update message objects.</param>
        /// <returns>The response message returned by UpdateItem operation.</returns>
        UpdateItemResponseType UpdateItem(UpdateItemType request);

        /// <summary>
        /// Send message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to send message objects.</param>
        /// <returns>The response message returned by SendItem operation.</returns>
        SendItemResponseType SendItem(SendItemType request);
    }
}