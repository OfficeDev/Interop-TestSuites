namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-FSSHTTP-FSSHTTPB protocol adapter interface implementation.
    /// </summary>
    public interface IMS_FSSHTTP_FSSHTTPBAdapter : IAdapter
    {
        #region Protocol Interface Design

        /// <summary>
        /// Gets the raw XML request sent to protocol SUT
        /// </summary>
        XmlElement LastRawRequestXml { get; }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT
        /// </summary>
        XmlElement LastRawResponseXml { get; }

        /// <summary>
        /// This method is used to send the cell storage request to the server.
        /// </summary>
        /// <param name="url">Specifies the URL of the file to edit.</param>
        /// <param name="subRequests">Specifies the sub request array.</param>
        /// <param name="requestToken">Specifies a non-negative request token integer that uniquely identifies the Request <seealso cref="Request"/>.</param>
        /// <param name="version">Specifies the version number of the request, whose value should only be 2.</param>
        /// <param name="minorVersion">Specifies the minor version number of the request, whose value should only be 0 or 2.</param>
        /// <param name="interval">Specifies a nonnegative integer in seconds, which the protocol client will repeat this request, the default value is null.</param>
        /// <param name="metaData">Specifies a 32-bit value that specifies information about the scenario and urgency of the request, the default value is null.</param>
        /// <param name="lastModifiedTime">Specify the last modified time, which is expressed as a tick count.</param>
        /// <param name="parentFolderResourceID">If UseResourceID is true, this parameter tells the host to create a file in the given folder ResourceID, regardless of the request URL value.</param>
        /// <param name="shouldReturnDisambiguatedFileName">If an upload request fails with a coherency failure, this flag specifies whether the host should return a suggested/available file name that the client can try instead</param>
        /// <param name="resourceID">Specify the invariant ResourceID for a file that uniquely identifies the file whose response is being generated</param>
        /// <param name="useResourceID">Specify if the protocol server MAY perform ResourceID specific behavior for the file whose contents or metadata contents are requested for uploading to the server or downloading from the server. </param>
        /// <returns>Returns the CellStorageResponse message received from the server.</returns>
        CellStorageResponse CellStorageRequest(string url, SubRequestType[] subRequests, string requestToken = "1", ushort? version = 2, ushort? minorVersion = 2, uint? interval = null, int? metaData = null, string lastModifiedTime = null, string parentFolderResourceID = null, bool? shouldReturnDisambiguatedFileName = null, string resourceID = null, bool? useResourceID = null);
        #endregion
    }
}