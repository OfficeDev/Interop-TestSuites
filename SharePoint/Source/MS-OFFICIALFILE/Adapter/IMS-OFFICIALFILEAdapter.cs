namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-OFFICIALFILE adapter.
    /// </summary>
    public interface IMS_OFFICIALFILEAdapter : IAdapter
    {
        /// <summary>
        /// Initialize the services of OFFICIALFILE.
        /// </summary>
        /// <param name="paras">The TransportType object indicates which transport parameters are used.</param>
        void IntializeService(InitialPara paras);

        /// <summary>
        /// This operation is used to retrieves data about the type, version of the repository and whether the repository is configured for routing.
        /// </summary>
        /// <returns>Data about the type, version of the repository and whether the repository is configured for routing or SoapException thrown.</returns>
        ServerInfo GetServerInfo();

        /// <summary>
        /// This operation is used to submit a file and its associated properties to the repository.
        /// </summary>
        /// <param name="fileToSubmit">The contents of the file.</param>
        /// <param name="properties"> The properties of the file.</param>
        /// <param name="recordRouting">The file type</param>
        /// <param name="sourceUrl">The source URL of the file.</param>
        /// <param name="userName">The name of the user submitting the file.</param>
        /// <returns>The data of SubmitFileResult or SoapException thrown.</returns>
        SubmitFileResult SubmitFile([System.Xml.Serialization.XmlElementAttribute(DataType = "base64Binary")] byte[] fileToSubmit, [System.Xml.Serialization.XmlArrayItemAttribute(IsNullable = false)] RecordsRepositoryProperty[] properties, string recordRouting, string sourceUrl, string userName);

        /// <summary>
        /// This operation is called to determine the storage location for the submission based on the rules in the repository and a suggested save location chosen by a user.
        /// </summary>
        /// <param name="properties">The properties of the file.</param>
        /// <param name="contentTypeName">The file type.</param>
        /// <param name="originalSaveLocation">The suggested save location chosen by a user.</param>
        /// <returns>Data details about the result.</returns>
        DocumentRoutingResult GetFinalRoutingDestinationFolderUrl(RecordsRepositoryProperty[] properties, string contentTypeName, string originalSaveLocation);

        /// <summary>
        /// This operation is called to retrieve information about the legal holds in a repository.
        /// </summary>
        /// <returns>A list of legal holds.</returns>
        HoldInfo[] GetHoldsInfo();

        /// <summary>
        /// This method is used to retrieve the recording routing information.
        /// </summary>
        /// <param name="recordRouting">The file type.</param>
        /// <returns>Recording routing information.</returns>
        string GetRecordingRouting(string recordRouting);

        /// <summary>
        /// This method is used to retrieve the recording routing information.
        /// </summary>
        /// <returns>Implementation-specific result data</returns>
        string GetRecordRoutingCollection();
    }
}