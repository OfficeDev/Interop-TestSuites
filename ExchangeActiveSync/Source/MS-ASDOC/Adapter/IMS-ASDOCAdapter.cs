namespace Microsoft.Protocols.TestSuites.MS_ASDOC
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-ASDOC.
    /// </summary>
    public interface IMS_ASDOCAdapter : IAdapter
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
        /// Retrieves data from the server for one or more individual documents.
        /// </summary>
        /// <param name="itemOperationsRequest">ItemOperations command request.</param>
        /// <param name="deliverMethod">Deliver method parameter.</param>
        /// <returns>ItemOperations command response.</returns>
        ItemOperationsResponse ItemOperations(ItemOperationsRequest itemOperationsRequest, DeliveryMethodForFetch deliverMethod);

        /// <summary>
        /// Finds entries in document library (using Universal Naming Convention paths).
        /// </summary>
        /// <param name="searchRequest">Search command request.</param>
        /// <returns>Search command response.</returns>
        SearchResponse Search(SearchRequest searchRequest);
    }
}