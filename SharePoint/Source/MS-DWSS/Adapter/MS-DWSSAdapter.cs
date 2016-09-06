namespace Microsoft.Protocols.TestSuites.MS_DWSS
{
    using System.Net;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-DWSS.
    /// </summary>
    public partial class MS_DWSSAdapter : ManagedAdapterBase, IMS_DWSSAdapter
    {
        #region Variables

        /// <summary>
        /// The Document Workspace Soap Service.
        /// </summary>
        private DwsSoap dwsService;

        /// <summary>
        /// The protocol transport type which is used to transfer messages between the client and SUT.
        /// </summary>
        private TransportProtocol transport;

        /// <summary>
        /// The SOAP version which is used to format the messages between the client and SUT.
        /// </summary>
        private SoapVersion soapVersion;

        /// <summary>
        /// Indicates weather the "Tasks", "Documents" and "Links" is added in the new created workspace.
        /// </summary>
        private bool isListAdded = false;

        #endregion Variables

        #region Properties

        /// <summary>
        /// Gets or sets the base URL of the Document Workspace Soap Service the client is requesting.
        /// </summary>
        public string ServiceUrl
        {
            get
            {
                return this.dwsService.Url;
            }

            set
            {
                this.dwsService.Url = value;
            }
        }

        /// <summary>
        /// Gets or sets the security credentials for Document Workspace Soap Service client authentication.
        /// </summary>
        public ICredentials Credentials
        {
            get
            {
                return this.dwsService.Credentials;
            }

            set
            {
                this.dwsService.Credentials = value;
            }
        }

        /// <summary>
        /// Gets or sets the private variable isListAdded.
        /// </summary>
        public bool IsListAdded
        {
            get
            {
                return this.isListAdded;
            }

            set
            {
                this.isListAdded = value;
            }
        }

        #endregion

        #region Initialization

        /// <summary>
        /// Overrides base Initialize method.
        /// </summary>
        /// <param name="testSite">An instance of the ITestSite.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-DWSS";

            this.LoadCommonConfiguration();

            // Load SHOULDMAY configuration 
            this.LoadCurrentSutSHOULDMAYConfiguration();

            this.dwsService = Proxy.CreateProxy<DwsSoap>(testSite);

            // Set default Dws service url to site collection.
            this.dwsService.Url = Common.GetConfigurationPropertyValue("SiteCollection", this.Site);

            // Set default Dws service credential to admin credential.
            string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            string password = Common.GetConfigurationPropertyValue("Password", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsService.Credentials = new NetworkCredential(userName, password, domain);

            this.SetSoapVersion(this.dwsService);

            // When request Url include HTTPS prefix, avoid closing base connection.
            // Local client will accept all certificates after executing this function. 
            this.transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            if (this.transport == TransportProtocol.HTTPS)
            {
                Common.AcceptServerCertificate();
            }
        }

        #endregion

        #region Dwss Protocol Operations

        /// <summary>
        /// The operation to determine whether an authenticated user has permission to create a Document Workspace at the specified URL.
        /// </summary>
        /// <param name="dwsUrl">Site-relative URL that specifies where to create the Document Workspace.</param>
        /// <param name="error">An error indication.</param>
        /// <returns>A site-relative URL that specifies where the Document Workspace is created.</returns>
        public string CanCreateDwsUrl(string dwsUrl, out Error error)
        {
            string respString = this.dwsService.CanCreateDwsUrl(dwsUrl);

            // Decode CanCreateDwsUrlResult element returned by server which is presented as a standalone xml string.
            string respXmlString = AdapterHelper.GenRespXmlString("CanCreateDwsUrlResult", respString);

            // Validate response xml schema and capture the related requirements.
            this.ValidateCanCreateDwsUrlResponseSchema(respXmlString);

            // Capture protocol transport related requirements.
            this.ValidateProtocolTransport();

            // Capture SOAP version related requirements.
            this.ValidateSoapVersion();

            // Deserialize CanCreateDwsUrlResult xml string to CanCreateDwsUrlResult object.
            CanCreateDwsUrlResult resp = AdapterHelper.XmlDeserialize<CanCreateDwsUrlResult>(respXmlString);

            error = resp.Item as Error;

            return resp.Item as string;
        }

        /// <summary>
        /// The operation to create a new Document Workspace.
        /// </summary>
        /// <param name="dwsName">Specifies the name of the Document Workspace site, this parameter can be empty.</param>
        /// <param name="users">Specifies the users to be added as contributors in the Document Workspace site, this parameter can be null.</param>
        /// <param name="dwsTitle">Specifies the title of the workspace, this parameter can be empty.</param>
        /// <param name="docs">Specifies information to be stored as a key-value pair in the site metadata, this parameter can be null.</param>
        /// <param name="error">An error indication.</param>
        /// <returns>CreateDws operation response.</returns>
        public CreateDwsResultResults CreateDws(string dwsName, UsersItem users, string dwsTitle, DocumentsItem docs, out Error error)
        {
            // Serialize users object to xml string.
            string usersString = users == null ? string.Empty : AdapterHelper.XmlSerialize(users);

            // Serialize docs object to xml string.
            string docsString = docs == null ? string.Empty : AdapterHelper.XmlSerialize(docs);

            string respString = this.dwsService.CreateDws(dwsName, usersString, dwsTitle, docsString);

            // Decode response standalone xml string.
            string respXmlString = AdapterHelper.GenRespXmlString("CreateDwsResult", respString);

            // Validate response xml schema and capture the related requirements.
            this.ValidateCreateDwsResponseSchema(respXmlString);

            // Capture protocol transport related requirements.
            this.ValidateProtocolTransport();

            // Capture SOAP version related requirements.
            this.ValidateSoapVersion();

            // Deserialize response xml string to CreateDwsResult object.
            CreateDwsResult resp = AdapterHelper.XmlDeserialize<CreateDwsResult>(respXmlString);

            error = resp.Item as Error;

            if (resp.Item is CreateDwsResultResults)
            {
                // Validate the requirements related to Results element in CreateDwsResult element.
                this.ValidateCreateDwsResultResults(usersString, resp.Item as CreateDwsResultResults);
            }

            return resp.Item as CreateDwsResultResults;
        }

        /// <summary>
        /// The operation to create a folder in the document library of the current Document Workspace site.
        /// </summary>
        /// <param name="folderUrl">Site-relative URL with the full path for the new folder.</param>
        /// <param name="error">An error indication.</param>
        public void CreateFolder(string folderUrl, out Error error)
        {
            string respString = this.dwsService.CreateFolder(folderUrl);

            // Decode response standalone xml string.
            string respXmlString = AdapterHelper.GenRespXmlString("CreateFolderResult", respString);

            // Validate response xml schema and capture the related requirements.
            this.ValidateCreateFolderResponseSchema(respXmlString);

            // Capture protocol transport related requirements.
            this.ValidateProtocolTransport();

            // Capture SOAP version related requirements.
            this.ValidateSoapVersion();

            // Deserialize response xml string to CanCreateDwsUrlResponse object.
            CreateFolderResult resp = AdapterHelper.XmlDeserialize<CreateFolderResult>(respXmlString);

            error = resp.Item as Error;
            if (error == null)
            {
                Site.Assert.AreEqual(
                    string.Empty,
                    resp.Item,
                    "An empty Result element (\"<Result/>\") should be returned if the call is successful, the actual returned Result element is '{0}'.",
                    resp.Item);

                this.ValidateCreateFolderResultResult();
            }
        }

        /// <summary>
        /// The operation to delete a Document Workspace from the protocol server.
        /// </summary>
        /// <param name="error">An error indication.</param>
        public void DeleteDws(out Error error)
        {
            string respString = this.dwsService.DeleteDws();

            // Decode response standalone xml string.
            string respXmlString = AdapterHelper.GenRespXmlString("DeleteDwsResult", respString);

            // Validate response xml schema and capture the related requirements.
            this.ValidateDeleteDwsResponseSchema(respXmlString);

            // Capture protocol transport related requirements.
            this.ValidateProtocolTransport();

            // Capture SOAP version related requirements.
            this.ValidateSoapVersion();

            // Deserialize response xml string to DeleteDwsResponse object.
            DeleteDwsResult resp = AdapterHelper.XmlDeserialize<DeleteDwsResult>(respXmlString);

            error = resp.Item as Error;
            if (error == null)
            {
                Site.Assert.AreEqual(
                    string.Empty,
                    resp.Item,
                    "An empty Result element (\"<Result/>\") should be returned if the call is successful, the actual returned Result element is '{0}'.",
                    resp.Item);

                this.ValidateDeleteDwsResultResult();
            }
        }

        /// <summary>
        /// The operation to delete a folder from a document library on the site.
        /// </summary>
        /// <param name="folderUrl">Site-relative URL specifying the folder to delete.</param>
        /// <param name="error">An error indication.</param>
        public void DeleteFolder(string folderUrl, out Error error)
        {
            string respString = this.dwsService.DeleteFolder(folderUrl);

            // Decode response standalone xml string.
            string respXmlString = AdapterHelper.GenRespXmlString("DeleteFolderResult", respString);

            // Validate response xml schema and capture the related requirements.
            this.ValidateDeleteFolderResponseSchema(respXmlString);

            // Capture protocol transport related requirements.
            this.ValidateProtocolTransport();

            // Capture SOAP version related requirements.
            this.ValidateSoapVersion();

            // Deserialize response xml string to DeleteFolderResponse object.
            DeleteFolderResult resp = AdapterHelper.XmlDeserialize<DeleteFolderResult>(respXmlString);

            error = resp.Item as Error;
            if (error == null)
            {
                Site.Assert.AreEqual(
                    string.Empty,
                    resp.Item,
                    "An empty Result element (\"<Result/>\") should be returned if the call is successful, the actual returned Result element is '{0}'.",
                    resp.Item);

                this.ValidateDeleteFolderResultResult();
            }
        }

        /// <summary>
        /// The operation to obtain a URL for a named document in a Document Workspace.
        /// </summary>
        /// <param name="docId">A unique string that represents a document key.</param>
        /// <param name="error">An error indication.</param>
        /// <returns>An absolute URL that refers to the requested document if the call is successful.</returns>
        public string FindDwsDoc(string docId, out Error error)
        {
            string respString = this.dwsService.FindDwsDoc(docId);

            // Decode response standalone xml string.
            string respXmlString = AdapterHelper.GenRespXmlString("FindDwsDocResult", respString);

            // Validate response xml schema and capture the related requirements.
            this.ValidateFindDwsDocResponseSchema(respXmlString);

            // Capture protocol transport related requirements.
            this.ValidateProtocolTransport();

            // Capture SOAP version related requirements.
            this.ValidateSoapVersion();

            // Deserialize response xml string to FindDwsDocResponse object.
            FindDwsDocResult resp = AdapterHelper.XmlDeserialize<FindDwsDocResult>(respXmlString);

            error = resp.Item as Error;

            return resp.Item as string;
        }

        /// <summary>
        /// The operation to return general information about the Document Workspace site, as well as its members, documents, links, and tasks.
        /// </summary>
        /// <param name="docUrl">A site-based URL of a document in the document library in the Document Workspace.</param>
        /// <param name="lastUpdate">Contains the lastUpdate value returned in the result of a previous GetDwsData or GetDwsMetaData operation, or an empty string.</param>
        /// <param name="error">An error indication.</param>
        /// <returns>GetDwsData operation response.</returns>
        public Results GetDwsData(string docUrl, string lastUpdate, out Error error)
        {
            string respString = this.dwsService.GetDwsData(docUrl, lastUpdate);

            // Decode response standalone xml string.
            string respXmlString = AdapterHelper.GenRespXmlString("GetDwsDataResult", respString);

            // Validate response xml schema and capture the related requirements.
            this.ValidateGetDwsDataResponseSchema(respXmlString);

            // Capture protocol transport related requirements.
            this.ValidateProtocolTransport();

            // Capture SOAP version related requirements.
            this.ValidateSoapVersion();

            // Deserialize response xml string to GetDwsDataResponse object.
            GetDwsDataResult resp = AdapterHelper.XmlDeserialize<GetDwsDataResult>(respXmlString);

            error = resp.Item as Error;

            if (error != null)
            {
                // Validate the requirements related to Error element in GetDwsDataResult element.
                this.ValidateGetDwsDataResultError();
            }

            if (resp.Item is Results)
            {
                // Validate the requirements related to Results element in GetDwsDataResult element.
                this.ValidateGetDwsDataResultResults(resp.Item as Results);
            }

            return resp.Item as Results;
        }

        /// <summary>
        /// The operation to return information about a Document Workspace site and the lists that it contains.
        /// </summary>
        /// <param name="docUrl">A site-relative URL that specifies the list or document to describe in the response.</param>
        /// <param name="docId">A unique string that represents a document key.</param>
        /// <param name="isMinimal">A Boolean value that specifies whether to return information.</param>
        /// <param name="error">An error indication.</param>
        /// <returns>GetDwsMetaData operation response.</returns>
        public GetDwsMetaDataResultTypeResults GetDwsMetaData(string docUrl, string docId, bool isMinimal, out Error error)
        {
            string respString = this.dwsService.GetDwsMetaData(docUrl, docId, isMinimal);

            // Decode response standalone xml string.
            string respXmlString = AdapterHelper.GenRespXmlString("GetDwsMetaDataResult", respString);

            // Validate response xml schema and capture the related requirements.
            this.ValidateGetDwsMetaDataResponseSchema(respXmlString);

            // Capture protocol transport related requirements.
            this.ValidateProtocolTransport();

            // Capture SOAP version related requirements.
            this.ValidateSoapVersion();

            // Deserialize response xml string to GetDwsMetaDataResponse object.
            GetDwsMetaDataResultType resp = AdapterHelper.XmlDeserialize<GetDwsMetaDataResultType>(respXmlString);

            error = resp.Item as Error;

            if (resp.Item is GetDwsMetaDataResultTypeResults)
            {
                // Validate the requirements related to Results element in GetDwsMetaDataResult element.
                this.ValidateGetDwsMetaDataResultResults(resp.Item as GetDwsMetaDataResultTypeResults, isMinimal);
            }

            return resp.Item as GetDwsMetaDataResultTypeResults;
        }

        /// <summary>
        /// The operation to delete a user from a Document Workspace.
        /// </summary>
        /// <param name="userId">The user identifier of the user to remove from the workspace. This positive integer MUST be in the range from zero through 2,147,483,647, inclusive.</param>
        /// <param name="error">An error indication.</param>
        public void RemoveDwsUser(int userId, out Error error)
        {
            string respString = this.dwsService.RemoveDwsUser(userId.ToString());

            // Decode response standalone xml string.
            string respXmlString = AdapterHelper.GenRespXmlString("RemoveDwsUserResult", respString);

            // Validate response xml schema and capture the related requirements.
            this.ValidateRemoveDwsUserResponseSchema(respXmlString);

            // Capture protocol transport related requirements.
            this.ValidateProtocolTransport();

            // Capture SOAP version related requirements.
            this.ValidateSoapVersion();

            // Deserialize response xml string to RemoveDwsUserResponse object.
            RemoveDwsUserResult resp = AdapterHelper.XmlDeserialize<RemoveDwsUserResult>(respXmlString);

            error = resp.Item as Error;
            if (error == null)
            {
                Site.Assert.AreEqual(
                    string.Empty,
                    resp.Item,
                    "An empty element should be returned if the call is successful, the actual returned Result element is '{0}'.",
                    resp.Item);

                this.ValidateRemoveDwsUserResultResults();
            }
        }

        /// <summary>
        /// The operation to change the title of a Document Workspace.
        /// </summary>
        /// <param name="dwsTitle">A string contains the new title of the workspace.</param>
        /// <param name="error">An error indication.</param>
        public void RenameDws(string dwsTitle, out Error error)
        {
            string respString = this.dwsService.RenameDws(dwsTitle);

            // Decode RenameDwsResult element returned by server which is presented as a standalone xml string.
            string respXmlString = AdapterHelper.GenRespXmlString("RenameDwsResult", respString);

            // Validate response xml schema and capture the related requirements.
            this.ValidateRenameDwsResponseSchema(respXmlString);

            // Capture protocol transport related requirements.
            this.ValidateProtocolTransport();

            // Capture SOAP version related requirements.
            this.ValidateSoapVersion();

            // Deserialize RenameDwsResult xml string to RenameDwsResult object.
            RenameDwsResult resp = AdapterHelper.XmlDeserialize<RenameDwsResult>(respXmlString);

            error = resp.Item as Error;
            if (error == null)
            {
                Site.Assert.AreEqual(
                    string.Empty,
                    resp.Item,
                    "An empty element should be returned if the call is successful, the actual returned Result element is '{0}'.",
                    resp.Item);

                this.ValidateRenameDwsResultResult();
            }
        }

        /// <summary>
        /// The operation to modify the metadata of a Document Workspace. This method is deprecated and should not be called by the protocol client.
        /// </summary>
        /// <param name="updates">A string that contains CAML instructions specifying how to update the workspace information.</param>
        /// <param name="meetingInstance">A string that contains the meeting information, this parameter can be empty.</param>
        /// <param name="error">An error indication.</param>
        /// <returns>UpdateDwsData operation response.</returns>
        public string UpdateDwsData(string updates, string meetingInstance, out Error error)
        {
            error = null;

            string respString = this.dwsService.UpdateDwsData(updates, meetingInstance);

            // Decode response standalone xml string.
            string respXmlString = AdapterHelper.GenRespXmlString(null, respString);

            // Validate response xml schema and capture the related requirements.
            this.ValidateUpdateDwsData();

            // Capture protocol transport related requirements.
            this.ValidateProtocolTransport();

            // Capture SOAP version related requirements.
            this.ValidateSoapVersion();

            // Check whether the return xml string is an Error object.
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(respXmlString);

            if (xmlDoc.DocumentElement.Name == "Error")
            {
                error = AdapterHelper.XmlDeserialize<Error>(respXmlString);
                return null;
            }

            return respXmlString;
        }

        #endregion

        #region Private methods

        /// <summary>
        /// A method used to load Common Configuration
        /// </summary>
        private void LoadCommonConfiguration()
        {
            // Merge the common configuration into local configuration
            string conmmonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);

            // Execute the merge the common configuration
            Common.MergeGlobalConfig(conmmonConfigFileName, this.Site);

            Common.CheckCommonProperties(this.Site, true);
        }

        /// <summary>
        /// A method used to load SHOULDMAY Configuration according to the current SUT version
        /// </summary>
        private void LoadCurrentSutSHOULDMAYConfiguration()
        {
            Common.MergeSHOULDMAYConfig(this.Site);
        }

        /// <summary>
        /// Set the SOAP version according to the SoapVersion property.
        /// </summary>
        /// <param name="dwsSoap">set meeting proxy</param>
        private void SetSoapVersion(DwsSoap dwsSoap)
        {
            this.soapVersion = Common.GetConfigurationPropertyValue<SoapVersion>("SoapVersion", this.Site);

            if (this.soapVersion == SoapVersion.SOAP11)
            {
                dwsSoap.SoapVersion = SoapProtocolVersion.Soap11;
            }
            else if (this.soapVersion == SoapVersion.SOAP12)
            {
                dwsSoap.SoapVersion = SoapProtocolVersion.Soap12;
            }
            else
            {
                Site.Assume.Fail(
                    "Property SoapVersion value must be {0} or {1} at the ptfconfig file.",
                    SoapVersion.SOAP11.ToString(),
                    SoapVersion.SOAP12.ToString());
            }
        }

        #endregion
    }
}