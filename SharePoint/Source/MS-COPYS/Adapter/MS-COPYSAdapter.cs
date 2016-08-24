namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This adapter class of MS-COPYS, it implements the IMS_COPYSAdapter interface.
    /// </summary>
    public partial class MS_COPYSAdapter : ManagedAdapterBase, IMS_COPYSAdapter
    {
        #region Variable

        /// <summary>
        /// An instance of CopySoap WebService.
        /// </summary>
        private CopySoap copySoapService;
 
        /// <summary>
        /// A ServiceLocation enum represents the current service the adapter use.
        /// </summary>
        private ServiceLocation currentServiceLocation = ServiceLocation.DestinationSUT;

        /// <summary>
        /// A string represents the password for the current user.
        /// </summary>
        private string passwordOfCurrentUser;

        /// <summary>
        /// A string represents the domain name for the current user.
        /// </summary>
        private string domainOfCurrentUser;

        /// <summary>
        /// A string represents the user name for the current user.
        /// </summary>
        private string currentUser;

        #endregion

        #region Initialize adapter

        /// <summary>
        /// Overrides the ManagedAdapterBase class' Initialize() method, it is used to initialize the adapter.
        /// </summary>
        /// <param name="testSite">A parameter represents the ITestSite instance, which is used to visit the test context.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            // Set the protocol name of current test suite
            Site.DefaultProtocolDocShortName = "MS-COPYS";

            TestSuiteManageHelper.Initialize(this.Site);

            // Merge the SHOULDMAY configuration file.
            Common.MergeSHOULDMAYConfig(this.Site);

            this.currentUser = TestSuiteManageHelper.DefaultUser;
            this.domainOfCurrentUser = TestSuiteManageHelper.DomainOfDefaultUser;
            this.passwordOfCurrentUser = TestSuiteManageHelper.PasswordOfDefaultUser;

            // Initialize the proxy class
            this.copySoapService = Proxy.CreateProxy<CopySoap>(this.Site, true, true, true);
            
            // Set service URL.
            this.copySoapService.Url = this.GetTargetServiceUrl(this.currentServiceLocation);

            // Set credential
            this.copySoapService.Credentials = TestSuiteManageHelper.DefaultUserCredential;

            // Set SOAP version
            this.copySoapService.SoapVersion = TestSuiteManageHelper.GetSoapProtoclVersionByCurrentSetting();

            // Accept Certificate
            TestSuiteManageHelper.AcceptServerCertificate();

            // set the service timeout.
            this.copySoapService.Timeout = TestSuiteManageHelper.CurrentSoapTimeOutValue;
        }

        #endregion Initialize adapter

        #region Implement IMS_COPYSAdapter

        /// <summary>
        /// Switch the current credentials of the protocol adapter by specified user. After perform this method, all protocol operations will be performed by specified user.
        /// </summary>
        /// <param name="userName">A parameter represents the user name.</param>
        /// <param name="password">A parameter represents the password of the user.</param>
        /// <param name="domain">A parameter represents the domain of the user.</param>
        public void SwitchUser(string userName, string password, string domain)
        {
            #region validate parameter
            if (string.IsNullOrEmpty(userName))
            {
                throw new ArgumentNullException("userName");
            }

            if (string.IsNullOrEmpty(password))
            {
                throw new ArgumentNullException("password");
            }

            if (string.IsNullOrEmpty(domain))
            {
                throw new ArgumentNullException("domain");
            }
            #endregion validate parameter

            if (null == this.copySoapService)
            {
                throw new InvalidOperationException(@"The adapter is not initialized, should call the ""IMS_COPYSAdapter.Initialize"" method to initialize adapter or call the ""ITestSite.GetAdapter"" method to get the initialized adapter instance.");
            }

            if (this.currentUser.Equals(userName, StringComparison.OrdinalIgnoreCase)
                && this.domainOfCurrentUser.Equals(domain, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            this.currentUser = userName;
            this.passwordOfCurrentUser = password;
            this.domainOfCurrentUser = domain;
            this.copySoapService.Credentials = new NetworkCredential(this.currentUser, this.passwordOfCurrentUser, this.domainOfCurrentUser);
        }

        /// <summary>
        /// Switch the target service location. The adapter will send the MS-COPYS message to specified service location.
        /// </summary>
        /// <param name="serviceLocation">A parameter represents the service location which host the MS-COPYS service.</param>
        public void SwitchTargetServiceLocation(ServiceLocation serviceLocation)
        {
            if (null == this.copySoapService)
            {
                throw new InvalidOperationException(@"The adapter is not initialized, should call the ""IMS_COPYSAdapter.Initialize"" method to initialize adapter or call the ""ITestSite.GetAdapter"" method to get the initialized adapter instance.");
            }

           if (serviceLocation == this.currentServiceLocation)
           {
               return;
           }

           this.currentServiceLocation = serviceLocation;
           this.copySoapService.Url = this.GetTargetServiceUrl(this.currentServiceLocation);
        }

        /// <summary>
        /// A method is used to copy a file when the destination of the operation is on the same protocol server as the source location.
        /// </summary>
        /// <param name="sourceUrl">A parameter represents the location of the file in the source location.</param>
        /// <param name="destinationUrls">A parameter represents a collection of destination location. The operation will try to copy files to that locations.</param>
        /// <returns>A return value represents the result of the operation. It includes status of the copy operation for a destination location.</returns>
        public CopyIntoItemsLocalResponse CopyIntoItemsLocal(string sourceUrl, string[] destinationUrls)
        {
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse;
            uint copyIntoItemsLocalResult;
            CopyResult[] results;

            try
            {
                copyIntoItemsLocalResult = this.copySoapService.CopyIntoItemsLocal(sourceUrl, destinationUrls, out results);
                copyIntoItemsLocalResponse = new CopyIntoItemsLocalResponse(copyIntoItemsLocalResult, results);
            }
            catch (SoapException soapEx)
            {
                throw;
            }

            this.VerifyCopyIntoItemsLocalOperationCapture(copyIntoItemsLocalResponse, destinationUrls);
            return copyIntoItemsLocalResponse;
        }

        /// <summary>
        /// A method is used to retrieve the contents and metadata for a file from the specified location.
        /// </summary>
        /// <param name="url">A parameter represents the location of the file.</param>
        /// <returns>A return value represents the file contents and metadata.</returns>
        public GetItemResponse GetItem(string url)
        {
            GetItemResponse getItemResponse;
            uint getItemResult;
            FieldInformation[] fields;
            byte[] rawStreamValues;

            try
            {
                getItemResult = this.copySoapService.GetItem(url, out fields, out rawStreamValues);
                getItemResponse = new GetItemResponse(getItemResult, fields, rawStreamValues);
            }
            catch (SoapException soapEx)
            {
                throw;
            }

            this.VerifyGetItemOperationCapture(getItemResponse);
            return getItemResponse;
        }

        /// <summary>
        /// A method used to copy a file to a destination server that is different from the source location.
        /// </summary>
        /// <param name="sourceUrl">A parameter represents the absolute URL of the file in the source location.</param>
        /// <param name="destinationUrls">A parameter represents a collection of locations on the destination server.</param>
        /// <param name="fields">A parameter represents a collection of the metadata for the file.</param>
        /// <param name="rawStreamValue">A parameter represents the contents of the file. The contents will be encoded in Base64 format and sent in request.</param>
        /// <returns>A return value represents the result of the operation.</returns>
        public CopyIntoItemsResponse CopyIntoItems(string sourceUrl, string[] destinationUrls, FieldInformation[] fields, byte[] rawStreamValue)
        {
            CopyIntoItemsResponse copyIntoItemsResponse;
            uint copyIntoItemsResult;
            CopyResult[] results;
            try
            {   
                copyIntoItemsResult = this.copySoapService.CopyIntoItems(sourceUrl, destinationUrls, fields, rawStreamValue, out results);
                copyIntoItemsResponse = new CopyIntoItemsResponse(copyIntoItemsResult, results);
            }
            catch (SoapException soapEx)
            {
                throw;
            }

            this.VerifyCopyIntoItemsOperationCapture(copyIntoItemsResponse, destinationUrls);
            return copyIntoItemsResponse;
        }

        #endregion

        #region private method
        /// <summary>
        /// A method used to get the target service url by specified service location.
        /// </summary>
        /// <param name="serviceLocation">A parameter represents the service location where host the MS-COPYS service.</param>
        /// <returns>A return value represents the service URL.</returns>
        private string GetTargetServiceUrl(ServiceLocation serviceLocation)
        {
            string targetServiceUrl;
            switch (serviceLocation)
            {
                case ServiceLocation.SourceSUT:
                    {
                        targetServiceUrl = Common.GetConfigurationPropertyValue("TargetServiceUrlOfSourceSUT", this.Site);
                        break;
                    }

                case ServiceLocation.DestinationSUT:
                    {
                        targetServiceUrl = Common.GetConfigurationPropertyValue("TargetServiceUrlOfDestinationSUT", this.Site);
                        break;
                    }

                default:
                    {
                        throw new InvalidOperationException("The test suite only support [SourceSUT] and [DestinationSUT] types ServiceLocation.");
                    }
            }

            return targetServiceUrl;
        }

        #endregion private method
    }
}