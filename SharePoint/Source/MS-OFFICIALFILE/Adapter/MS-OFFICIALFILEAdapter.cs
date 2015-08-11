namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    using System;
    using System.Web.Services.Protocols;
    using System.Xml;
    using System.Xml.Serialization;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OFFICIALFILE protocol's adapter.
    /// </summary>
    public partial class MS_OFFICIALFILEAdapter : ManagedAdapterBase, IMS_OFFICIALFILEAdapter
    {
        #region variables
        /// <summary>
        /// Specify Enum Represent Result of Validation.
        /// </summary>
        private OfficialFileSoap officialfileService;
        #endregion variables

        /// <summary>
        /// Overrides IAdapter's Initialize(), to set testSite.DefaultProtocolDocShortName.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter,Make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OFFICIALFILE";

            // Get the name of common configuration file.
            string commonConfigFileName = Common.Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);

            // Merge the common configuration.
            Common.Common.MergeGlobalConfig(commonConfigFileName, this.Site);

            // Merge the Should/May configuration.
            Common.Common.MergeSHOULDMAYConfig(this.Site);

            // Initialize the OfficialFileSoap.
            this.officialfileService = Common.Proxy.CreateProxy<OfficialFileSoap>(testSite, true, true);

            AdapterHelper.Initialize(testSite);

            // Set the transportType.
            TransportType transportType = Common.Common.GetConfigurationPropertyValue<TransportType>("TransportType", this.Site);

            // Case request URl include HTTPS prefix, use this function to avoid closing base connection.
            // Local client will accept all certificate after execute this function. 
            if (transportType == TransportType.HTTPS)
            {
                AdapterHelper.AcceptAllCertificate();
            }

            // Set the version of the SOAP protocol used to make the SOAP request to the web service
            this.officialfileService.SoapVersion = Common.Common.GetConfigurationPropertyValue<SoapProtocolVersion>("SoapVersion", this.Site);
        }

        /// <summary>
        /// Initialize the services of MS-OFFICIALFILE.
        /// </summary>
        /// <param name="paras">The TransportType object indicates which transport parameters are used.</param>
        public void IntializeService(InitialPara paras)
        {
            // Initialize the url
            this.officialfileService.Url = paras.Url;

            // Set the security credential for Web service client authentication.
            this.officialfileService.Credentials = AdapterHelper.ConfigureCredential(paras.UserName, paras.Password, paras.Domain);
        }

        /// <summary>
        /// This operation is used to determine the storage location for the submission based on the rules in the repository and a suggested save location chosen by a user.
        /// </summary>
        /// <param name="properties">The properties of the file.</param>
        /// <param name="contentTypeName">The file type.</param>
        /// <param name="originalSaveLocation">The suggested save location chosen by a user.</param>
        /// <returns>An implementation specific URL to the storage location and some results for the submission or SoapException thrown.</returns>
        public DocumentRoutingResult GetFinalRoutingDestinationFolderUrl(RecordsRepositoryProperty[] properties, string contentTypeName, string originalSaveLocation)
        {
            DocumentRoutingResult documentRoutingResult = this.officialfileService.GetFinalRoutingDestinationFolderUrl(properties, contentTypeName, originalSaveLocation);

            // As response is returned successfully, the transport related requirements can be captured.
            this.VerifyTransportRelatedRequirments();

            this.VerifyGetFinalRoutingDestinationFolderUrl();

            return documentRoutingResult;
        }

        /// <summary>
        /// This operation is used to retrieves data about the type, version of the repository and whether the repository is configured for routing.
        /// </summary>
        /// <returns>Data about the type, version of the repository and whether the repository is configured for routing or SoapException thrown.</returns>
        public ServerInfo GetServerInfo()
        {
            string serverInfo = this.officialfileService.GetServerInfo();

            // As response is returned successfully, the transport related requirements can be captured.
            this.VerifyTransportRelatedRequirments();

            // Verify GetServerInfo operation.
            ServerInfo serverInfoClass = this.VerifyAndParseGetServerInfo(serverInfo);

            return serverInfoClass;
        }

        /// <summary>
        /// This operation is used to submit a file and its associated properties to the repository.
        /// </summary>
        /// <param name="fileToSubmit">The contents of the file.</param>
        /// <param name="properties"> The properties of the file.</param>
        /// <param name="recordRouting">The file type.</param>
        /// <param name="sourceUrl">The source URL of the file.</param>
        /// <param name="userName">The name of the user submitting the file.</param>
        /// <returns>The data of SubmitFileResult or SoapException thrown.</returns>
        public SubmitFileResult SubmitFile(
            [XmlElementAttribute(DataType = "base64Binary")] byte[] fileToSubmit,
            [XmlArrayItemAttribute(IsNullable = false)] RecordsRepositoryProperty[] properties,
            string recordRouting,
            string sourceUrl,
            string userName)
        {
            string submitFileResult = this.officialfileService.SubmitFile(
                fileToSubmit, properties, recordRouting, sourceUrl, userName);

            // As response is returned successfully, the transport related requirements can be captured.
            this.VerifyTransportRelatedRequirments();

            SubmitFileResult value = this.VerifyAndParseSubmitFile(submitFileResult);
            return value;
        }

        /// <summary>
        /// This operation is called to retrieve information about the legal holds in a repository.
        /// </summary>
        /// <returns>A list of legal holds.</returns>
        public HoldInfo[] GetHoldsInfo()
        {
            HoldInfo[] getHoldsInfoResult = this.officialfileService.GetHoldsInfo();

            // As response is returned successfully, the transport related requirements can be captured
            this.VerifyTransportRelatedRequirments();
            this.VerifyGetHoldsInfo();

            return getHoldsInfoResult;
        }

        /// <summary>
        /// This method is used to retrieve the recording routing information.
        /// </summary>
        /// <param name="recordRouting">The file type.</param>
        /// <returns>Recording routing information.</returns>
        public string GetRecordingRouting(string recordRouting)
        {
            string routingInfor = this.officialfileService.GetRecordRouting(recordRouting);

            // Verify GetRecordingRouting operation.
            this.VerifyGetRoutingInfo();

            return routingInfor;
        }

        /// <summary>
        /// This method is used to retrieve the recording routing collection information.
        /// </summary>
        /// <returns>Implementation-specific result data</returns>
        public string GetRecordRoutingCollection()
        {
            string recordingCollection = this.officialfileService.GetRecordRoutingCollection();

            // Verify GetRecordingRoutingCollection operation.
            this.VerifyGetRoutingCollectionInfo();

            return recordingCollection;
        }

        /// <summary>
        /// This method is used to parse the GetServerInfoResult from xml to class.
        /// </summary>
        /// <param name="serverInfo">A string indicates the serverInfo returned by the protocol server.</param>
        /// <returns>A serverInfo parsed from the response.</returns>
        private ServerInfo ParseGetServerInfoResult(string serverInfo)
        {
            ServerInfo info = new ServerInfo();

            // Read string in SubmitFileResult response
            XmlDocument document = new XmlDocument();
            document.LoadXml(serverInfo);

            XmlNodeList nodeList = null;
            XmlNode node = null;

            nodeList = document.GetElementsByTagName("ServerType");
            if (nodeList != null && nodeList.Count == 1)
            {
                node = nodeList[0];
                info.ServerType = node.InnerText;
            }

            nodeList = document.GetElementsByTagName("ServerVersion");
            if (nodeList != null && nodeList.Count == 1)
            {
                node = nodeList[0];
                info.ServerVersion = node.InnerText;
            }

            nodeList = document.GetElementsByTagName("RoutingWeb");
            if (nodeList != null && nodeList.Count == 1)
            {
                node = nodeList[0];
                info.RoutingWeb = node.InnerText;
            }

            return info;
        }

        /// <summary>
        /// This method is used to parse the SubmitFileResult from xml to class.
        /// </summary>
        /// <param name="response">The value response from server.</param>
        /// <returns>Enum value of SubmitFileResult.</returns>
        private SubmitFileResult ParseSubmitFileResult(string response)
        {
            SubmitFileResult result = new SubmitFileResult();

            // Read string in SubmitFileResult response
            XmlDocument document = new XmlDocument();
            document.LoadXml(response);

            XmlNodeList nodeList = null;
            XmlNode node = null;

            nodeList = document.GetElementsByTagName("ResultCode");
            if (nodeList != null && nodeList.Count == 1)
            {
                node = nodeList[0];
                switch (node.InnerText)
                {
                    case "Success":
                        result.ResultCode = SubmitFileResultCode.Success;
                        break;
                    case "MoreInformation":
                        result.ResultCode = SubmitFileResultCode.MoreInformation;
                        break;
                    case "InvalidRouterConfiguration":
                        result.ResultCode = SubmitFileResultCode.InvalidRouterConfiguration;
                        break;
                    case "InvalidArgument":
                        result.ResultCode = SubmitFileResultCode.InvalidArgument;
                        break;
                    case "InvalidUser":
                        result.ResultCode = SubmitFileResultCode.InvalidUser;
                        break;
                    case "NotFound":
                        result.ResultCode = SubmitFileResultCode.NotFound;
                        break;
                    case "FileRejected":
                        result.ResultCode = SubmitFileResultCode.FileRejected;
                        break;
                    case "UnknownError":
                        result.ResultCode = SubmitFileResultCode.UnknownError;
                        break;
                    default:
                        break;
                }
            }

            nodeList = document.GetElementsByTagName("CustomProcessingResult");
            if (nodeList != null && nodeList.Count == 1)
            {
                node = nodeList[0];
                result.CustomProcessingResult = new CustomProcessingResult();
                switch (node.InnerText)
                {
                    case "Success":
                        result.CustomProcessingResult.HoldProcessingResult = HoldProcessingResult.Success;
                        break;
                    case "Failure":
                        result.CustomProcessingResult.HoldProcessingResult = HoldProcessingResult.Failure;
                        break;
                    case "InDropOffZone":
                        result.CustomProcessingResult.HoldProcessingResult = HoldProcessingResult.InDropOffZone;
                        break;
                    default:
                        break;
                }
            }

            nodeList = document.GetElementsByTagName("ResultUrl");
            if (nodeList != null && nodeList.Count == 1)
            {
                node = nodeList[0];
                result.ResultUrl = node.InnerText;
            }

            nodeList = document.GetElementsByTagName("AdditionalInformation");
            if (nodeList != null && nodeList.Count == 1)
            {
                node = nodeList[0];
                result.AdditionalInformation = node.InnerText;
            }

            return result;
        }
    }
}