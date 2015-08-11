namespace Microsoft.Protocols.TestSuites.MS_VERSS
{
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-VERSS.
    /// </summary>
    public partial class MS_VERSSAdapter : ManagedAdapterBase, IMS_VERSSAdapter
    {
        #region Variable

        /// <summary>
        /// Instance of the Web Service
        /// </summary>
        private VersionsSoap service = null;

        /// <summary>
        /// The transport protocol that is used.
        /// </summary>
        private TransportProtocol transportProtocol;

        #endregion

        #region Initialize TestSuite

        /// <summary>
        /// Initialize the protocol adapter.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-VERSS";
            AdapterHelper.Initialize(testSite);

            if (this.service == null)
            {
                string globalConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", testSite);
                Common.MergeGlobalConfig(globalConfigFileName, testSite);
                Common.CheckCommonProperties(this.Site, true);
                Common.MergeSHOULDMAYConfig(testSite);

                this.service = Proxy.CreateProxy<VersionsSoap>(testSite);

                this.service.Url = Common.GetConfigurationPropertyValue("MSVERSSServiceUrl", testSite);

                SoapVersion soapVersion = Common.GetConfigurationPropertyValue<SoapVersion>("SoapVersion", this.Site);
                switch (soapVersion)
                {
                    case SoapVersion.SOAP11:
                        this.service.SoapVersion = SoapProtocolVersion.Soap11;
                        break;
                    default:
                        this.service.SoapVersion = SoapProtocolVersion.Soap12;
                        break;
                }

                TransportProtocol transportType = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
                switch (transportType)
                {
                    case TransportProtocol.HTTPS:
                        Common.AcceptServerCertificate();
                        this.transportProtocol = TransportProtocol.HTTPS;
                        break;
                    default:
                        this.transportProtocol = TransportProtocol.HTTP;
                        break;
                }
            }

            string domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            string userName = Common.GetConfigurationPropertyValue("UserName", testSite);
            string userPassword = Common.GetConfigurationPropertyValue("Password", testSite);
            this.service.Credentials = new NetworkCredential(userName, userPassword, domain);
        }
        #endregion

        #region Implement IMS_VERSSAdapter
        /// <summary>
        /// Initialize a protocol web service using incorrect authorization information.
        /// </summary>
        public void InitializeUnauthorizedService()
        {
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            string userPassword = Common.GenerateInvalidPassword(Common.GetConfigurationPropertyValue("Password", this.Site));
            this.service.Credentials = new NetworkCredential(userName, userPassword, domain);
        }

        /// <summary>
        /// This operation is used to get details about all versions of the specified file that the user can access.
        /// </summary>
        /// <param name="fileName">The site-relative path of a file on the protocol server.</param>
        /// <returns>The response message for getting all versions of the specified file that the user can access.</returns>
        public GetVersionsResponseGetVersionsResult GetVersions(string fileName)
        {
            try
            {
                GetVersionsResponseGetVersionsResult getVersionsResult = this.service.GetVersions(fileName);

                this.VerifyTransport();
                this.VerifySOAPVersion(this.service.SoapVersion);
                this.VerifyGetVersions(getVersionsResult, SchemaValidation.LastRawResponseXml.OuterXml);

                return getVersionsResult;
            }
            catch (SoapException soapException)
            {
                this.VerifySOAPFaultDetails(soapException, SchemaValidation.LastRawResponseXml.OuterXml);
                throw;
            }
            catch (WebException)
            {
                this.VerifyServerFaults();
                throw;
            }
        }

        /// <summary>
        /// This operation is used to restore the specified file to a specific version.
        /// </summary>
        /// <param name="fileName">The site-relative path of the file which will be restored.</param>
        /// <param name="fileVersion">The version number of the file which will be restored.</param>
        /// <returns>The response message for restoring the specified file to a specific version.</returns>
        public RestoreVersionResponseRestoreVersionResult RestoreVersion(string fileName, string fileVersion)
        {
            try
            {
                fileVersion = fileVersion.Replace("@", string.Empty);
                RestoreVersionResponseRestoreVersionResult restoreVersionResult = this.service.RestoreVersion(
                    fileName, 
                    fileVersion);

                this.VerifyTransport();
                this.VerifySOAPVersion(this.service.SoapVersion);
                this.VerifyRestoreVersion(restoreVersionResult, SchemaValidation.LastRawResponseXml.OuterXml);
                return restoreVersionResult;
            }
            catch (SoapException soapException)
            {
                this.VerifySOAPFaultDetails(soapException, SchemaValidation.LastRawResponseXml.OuterXml);
                throw;
            }
            catch (WebException)
            {
                this.VerifyServerFaults();
                throw;
            }
        }

        /// <summary>
        /// The DeleteVersion operation is used to delete a specific version of the specified file. 
        /// </summary>
        /// <param name="fileName">The site-relative path of the file name whose version is to be deleted.</param>
        /// <param name="fileVersion">The number of the file version to be deleted.</param>
        /// <returns>The response message for deleting a version of the specified file on the protocol server.</returns>
        public DeleteVersionResponseDeleteVersionResult DeleteVersion(string fileName, string fileVersion)
        {
            try
            {
                fileVersion = fileVersion.Replace("@", string.Empty);
                DeleteVersionResponseDeleteVersionResult deleteVersionResult = this.service.DeleteVersion(
                    fileName,
                    fileVersion);

                this.VerifyTransport();
                this.VerifySOAPVersion(this.service.SoapVersion);
                this.VerifyDeleteVersion(deleteVersionResult, SchemaValidation.LastRawResponseXml.OuterXml);
                return deleteVersionResult;
            }
            catch (SoapException soapException)
            {
                this.VerifySOAPFaultDetails(soapException, SchemaValidation.LastRawResponseXml.OuterXml);
                throw;
            }
            catch (WebException)
            {
                this.VerifyServerFaults();
                throw;
            }
        }

        /// <summary>
        /// This operation is used to delete all the previous versions of the specified file except 
        /// the published version and the current version.
        /// </summary>
        /// <param name="fileName">The site-relative path of the file which will be deleted.</param>
        /// <returns>The response message for deleting all previous versions of 
        /// the specified file on the protocol server.</returns>
        public DeleteAllVersionsResponseDeleteAllVersionsResult DeleteAllVersions(string fileName)
        {
            try
            {
                DeleteAllVersionsResponseDeleteAllVersionsResult deleteAllVersionsResult = 
                    this.service.DeleteAllVersions(fileName);

                this.VerifyTransport();
                this.VerifySOAPVersion(this.service.SoapVersion);
                this.VerifyDeleteAllVersions(deleteAllVersionsResult, SchemaValidation.LastRawResponseXml.OuterXml);
                return deleteAllVersionsResult;
            }
            catch (SoapException soapException)
            {
                this.VerifySOAPFaultDetails(soapException, SchemaValidation.LastRawResponseXml.OuterXml);
                throw;
            }
            catch (WebException)
            {
                this.VerifyServerFaults();
                throw;
            }
        }

        #endregion
    }
}