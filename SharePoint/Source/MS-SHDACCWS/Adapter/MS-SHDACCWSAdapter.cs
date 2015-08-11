namespace Microsoft.Protocols.TestSuites.MS_SHDACCWS
{
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This is the adapter implementation class of IMS_SHDACCWSAdapter 
    /// </summary>
    public partial class MS_SHDACCWSAdapter : ManagedAdapterBase, IMS_SHDACCWSAdapter
    {
        #region Variable

        /// <summary>
        /// Web service proxy generated from the full WSDL of MS-SHDACCWS protocol
        /// </summary>
         private SharedAccessSoap service;

        #endregion

        #region Initialize TestSuite

        /// <summary>
        /// Overrides IAdapter's Initialize method, to set default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">A parameter represents a ITestSite instance which is used to get/operate current test suite context.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            // Set the protocol name of current test suite
            testSite.DefaultProtocolDocShortName = "MS-SHDACCWS";

            // Merge the common configuration into local configuration
            string conmmonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);
            Common.MergeGlobalConfig(conmmonConfigFileName, this.Site);

            Common.CheckCommonProperties(this.Site, true);

            // Load SHOULDMAY configuration
            Common.MergeSHOULDMAYConfig(this.Site);

            // Initialize the proxy.
            this.service = Proxy.CreateProxy<SharedAccessSoap>(this.Site, true, true, false);

            this.service.Url = this.GetTargetServiceUrl();
            string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            string password = Common.GetConfigurationPropertyValue("Password", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.service.Credentials = new NetworkCredential(userName, password, domain);
            this.SetSoapVersion();
            if (TransportProtocol.HTTPS == Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site))
            {
                Common.AcceptServerCertificate();
            }

            // Configure the service timeout.
            int soapTimeOut = Common.GetConfigurationPropertyValue<int>("ServiceTimeOut", this.Site);

            // 60000 means the configure SOAP Timeout is in minute.
            this.service.Timeout = soapTimeOut * 60000;
        }

        #endregion

        #region Implement ISHDACCWSAdapter

        /// <summary>
        /// Specifies whether a co-authoring transition request was made for a document.
        /// </summary>
        /// <param name="id">The identifier(Guid) of the document in the server.</param>
        /// <returns>Whether a co-authoring transition request was made for a document.</returns>
        public bool IsOnlyClient(Guid id)
        {
            bool isOnlyClient = false;
            isOnlyClient = this.service.IsOnlyClient(id.ToString("B"));

            // Capture requirements about transport.
            this.VerifyTransportProtocol();

            // If no exception thrown, capture requirement about schema of IsOnlyClient operation.
            this.VerifySchemaOfIsOnlyClientOperation();
            return isOnlyClient;
        }

        /// <summary>
        /// Set the SOAP version according to the SoapVersion property.
        /// </summary>
        private void SetSoapVersion()
        {
            SoapVersion soapVersion = Common.GetConfigurationPropertyValue<SoapVersion>("SoapVersion", this.Site);

            switch (soapVersion)
            {
                case SoapVersion.SOAP11:
                    {
                        this.service.SoapVersion = SoapProtocolVersion.Soap11;
                        break;
                    }

                default:
                    {
                        this.service.SoapVersion = SoapProtocolVersion.Soap12;
                        break;
                    }
            }
        }

        /// <summary>
        /// A method used to get target service fully qualified URL, it indicates which site the test suite will run on.
        /// </summary>
        /// <returns>A return value represents the target service fully qualified URL</returns>
        private string GetTargetServiceUrl()
        {
            string fullyServiceURL = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);
            return fullyServiceURL;
        }
        #endregion
    }
}