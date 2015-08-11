namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides method FindItem defined in MS-OXWSSRCH.
    /// </summary>
    public interface IMS_OXWSSRCHAdapter : IAdapter
    {
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT.
        /// </summary>
        IXPathNavigable LastRawRequestXml { get; }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT.
        /// </summary>
        IXPathNavigable LastRawResponseXml { get; }

        /// <summary>
        /// Find item on the server.
        /// </summary>
        /// <param name="findItemRequest">Find item operation request type.</param>
        /// <returns>Find item operation response type.</returns>
        FindItemResponseType FindItem(FindItemType findItemRequest);
    }
}