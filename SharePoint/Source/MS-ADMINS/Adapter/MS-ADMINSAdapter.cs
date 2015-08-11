namespace Microsoft.Protocols.TestSuites.MS_ADMINS
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-ADMINS Adapter implementation.
    /// </summary>
    public partial class MS_ADMINSAdapter : ManagedAdapterBase, IMS_ADMINSAdapter
    {
        /// <summary>
        /// A list used to record all site collection created by test cases.
        /// </summary>
        private static List<string> listCreatedSiteCollection = new List<string>();

        /// <summary>
        /// An Administration Web Services instance.
        /// </summary>
        private AdminSoap adminService;

        /// <summary>
        /// Gets or sets the destination Url of web service operation.
        /// </summary>
        /// <value>Destination Url of web service operation.</value>
        public string Url
        {
            get
            {
                return this.adminService.Url;
            }

            set
            {
                this.adminService.Url = value;
            }
        }

        #region Test Suite Initialization and CleanUp

        /// <summary>
        /// Overrides IAdapter's Initialize().
        /// </summary>
        /// <param name="testSite">A parameter represents an ITestSite instance.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-ADMINS";

            // Initialize the AdminSoap.
            this.adminService = Proxy.CreateProxy<AdminSoap>(this.Site);

            // Merge the common configuration into local configuration
            string conmmonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);

            // Execute the merge the common configuration
            Common.MergeGlobalConfig(conmmonConfigFileName, this.Site);

            Common.CheckCommonProperties(this.Site, true);

            // Load SHOULDMAY configuration 
            Common.MergeSHOULDMAYConfig(this.Site);

            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            switch (transport)
            {
                case TransportProtocol.HTTP:
                    {
                        this.adminService.Url = Common.GetConfigurationPropertyValue("HTTPTargetServiceUrl", this.Site);
                        break;
                    }

                default:
                    {
                        this.adminService.Url = Common.GetConfigurationPropertyValue("HTTPSTargetServiceUrl", this.Site);

                        // When request Url include HTTPS prefix, avoid closing base connection.
                        // Local client will accept all certificates after executing this function. 
                        Common.AcceptServerCertificate();
                        break;
                    }
            }

            string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            string password = Common.GetConfigurationPropertyValue("Password", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.adminService.Credentials = new NetworkCredential(userName, password, domain);

            this.SetSoapVersion(this.adminService);

            // Configure the service timeout.
            string soapTimeOut = Common.GetConfigurationPropertyValue("ServiceTimeOut", this.Site);

            // 60000 means the configure SOAP Timeout is in milliseconds.
            this.adminService.Timeout = Convert.ToInt32(soapTimeOut) * 60000;
        }

        /// <summary>
        /// A method used to Clean up meetings added by test case.
        /// </summary>
        public override void Reset()
        {
            base.Reset();

            // Delete all sites created by test case and make sure current environment is clean.
            foreach (string urlCache in listCreatedSiteCollection.ToArray())
            {
                this.DeleteSite(urlCache);
            }
        }

        #endregion

        /// <summary>
        /// Creates a Site collection.
        /// </summary>
        /// <param name="url">The absolute URL of the site collection.</param>
        /// <param name="title">The display name of the site collection.</param>
        /// <param name="description">A description of the site collection.</param>
        /// <param name="lcid">The language that is used in the site collection.</param>
        /// <param name="webTemplate">The name of the site template which is used when creating the site collection.</param>
        /// <param name="ownerLogin">The user name of the site collection owner.</param>
        /// <param name="ownerName">The display name of the owner.</param>
        /// <param name="ownerEmail">The e-mail address of the owner.</param>
        /// <param name="portalUrl">The URL of the portal site for the site collection.</param>
        /// <param name="portalName">The name of the portal site for the site collection.</param>
        /// <returns>The CreateSite result.</returns>
        public string CreateSite(string url, string title, string description, int? lcid, string webTemplate, string ownerLogin, string ownerName, string ownerEmail, string portalUrl, string portalName)
        {
            string result = null;
            try
            {
                // Call CreateSite method.
                result = this.adminService.CreateSite(url, title, description, lcid ?? 0, lcid.HasValue, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);

                // Add the site's url to cache list.
                listCreatedSiteCollection.Add(result);

                // Capture the transport related requirements and verify CreateSite response requirements.
                this.VerifyTransportRelatedRequirements();
                this.ValidateCreateSiteResponseData(result);
            }
            catch (SoapException soapEx)
            {
                // Validate soap fault message structure and capture related requirements.
                this.VerifySoapFaultRequirements(soapEx);

                throw new SoapException(soapEx.Detail.InnerText, soapEx.Code, soapEx.Actor, soapEx.Detail, soapEx);
            }
            catch (WebException webEx)
            {
                Site.Assert.Fail("Server failed to create the site collection: {0}", webEx.Message);
                throw;
            }

            return result;
        }

        /// <summary>
        /// Deletes the specified Site collection.
        /// </summary>
        /// <param name="url">The absolute URL of the site collection which is to be deleted.</param>
        public void DeleteSite(string url)
        {
            try
            {
                // Call DeleteSite method.
                this.adminService.DeleteSite(url);

                // Capture the transport related requirements and verify DeleteSite response requirements.
                this.VerifyTransportRelatedRequirements();
                this.ValidateDeleteSiteResponse();
            }
            catch (SoapException ex)
            {
                Site.Log.Add(LogEntryKind.Debug, "Failed to Delete the site {0}, and the returned SOAP exception is: \r\n{1}!", url, ex.Detail.InnerXml);

                // Validate soap fault message structure and capture related requirements.
                this.VerifySoapFaultRequirements(ex);

                throw new SoapException(ex.Detail.InnerText, ex.Code, ex.Actor, ex.Detail, ex);
            }
            finally
            {
                // Remove the deleted url from cache list.
                listCreatedSiteCollection.Remove(url);
            }
        }

        /// <summary>
        /// Returns information about the languages which are used in the protocol server deployment.
        /// </summary>
        /// <returns>The GetLanguages result.</returns>
        public GetLanguagesResponseGetLanguagesResult GetLanguages()
        {
            GetLanguagesResponseGetLanguagesResult getLanguagesResult = null;
            try
            {
                getLanguagesResult = this.adminService.GetLanguages();

                // Capture the transport related requirements and verify GetLanguages response requirements.
                this.VerifyTransportRelatedRequirements();
                this.ValidateGetLanguagesResponseData(getLanguagesResult);
            }
            catch (SoapException ex)
            {
                Site.Log.Add(LogEntryKind.Debug, "Throw exceptions while get LCID values that specify the languages used in the protocol server deployment: {0}.", ex.Message);

                throw new SoapException(ex.Detail.InnerText, ex.Code, ex.Actor, ex.Detail, ex);
            }

            return getLanguagesResult;
        }

        /// <summary>
        /// Set the SOAP version according to the SoapVersion property.
        /// </summary>
        /// <param name="adminProxy">set admin proxy</param>
        private void SetSoapVersion(AdminSoap adminProxy)
        {
            SoapVersion soapVersion = Common.GetConfigurationPropertyValue<SoapVersion>("SoapVersion", this.Site);

            switch (soapVersion)
            {
                case SoapVersion.SOAP11:
                    {
                        adminProxy.SoapVersion = SoapProtocolVersion.Soap11;
                        break;
                    }

                default:
                    {
                        adminProxy.SoapVersion = SoapProtocolVersion.Soap12;
                        break;
                    }
            }
        }
    }
}