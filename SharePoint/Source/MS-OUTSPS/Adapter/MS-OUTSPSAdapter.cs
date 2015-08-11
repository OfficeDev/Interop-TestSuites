namespace Microsoft.Protocols.TestSuites.MS_OUTSPS
{
    using System;
    using System.IO;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This adapter class of MS-OUTSPS 
    /// </summary>
    public partial class MS_OUTSPSAdapter : ManagedAdapterBase, IMS_OUTSPSAdapter
    {
        #region Private member variables
        /// <summary>
        /// Web service proxy generated from the full WSDL of LISTSWS
        /// </summary>
        private ListsSoap listsProxy;

        #endregion
 
        #region MS-OUTSPS adapter operations

        /// <summary>
        /// Add a list in the current site based on the specified name, description, and list template identifier.
        /// </summary>
        /// <param name="listName">The title of the list which will be added.</param>
        /// <param name="description">Text which will be set as description of newly created list.</param>
        /// <param name="templateId">The template ID used to create this list.</param>
        /// <returns>Returns the AddList result.</returns>
        public AddListResponseAddListResult AddList(string listName, string description, int templateId)
        {
            if (null == this.listsProxy)
            {
                throw new InvalidOperationException("The Proxy instance is NULL, need to initialize the adapter");
            }

            AddListResponseAddListResult result = null;
            result = this.listsProxy.AddList(listName, description, templateId);
            this.VerifyTransportRequirement();

            return result;
        }

        /// <summary>
        /// The AddAttachment operation is used to add an attachment to the specified list item in the specified list.
        /// </summary>
        /// <param name="listName">The GUID or the list title of the list in which the list item to add attachment.</param>
        /// <param name="listItemId">The id of the list item in which the attachment will be added.</param>
        /// <param name="fileName">The name of the file being added as an attachment.</param>
        /// <param name="attachment">Content of the attachment file (byte array).</param>
        /// <returns>The URL of the newly added attachment.</returns>
        public string AddAttachment(string listName, string listItemId, string fileName, byte[] attachment)
        {
            if (null == this.listsProxy)
            {
                throw new InvalidOperationException("The Proxy instance is NULL, need to initialize the adapter");
            }

            string attachmentRelativeUrl = null;
            attachmentRelativeUrl = this.listsProxy.AddAttachment(listName, listItemId, fileName, attachment);
            this.VerifyTransportRequirement();

            return attachmentRelativeUrl;
        }

        /// <summary>
        /// The DeleteAttachment operation is used to remove the attachment from the specified list 
        /// item in the specified list.
        /// </summary>
        /// <param name="listName">The name of the list in which the list item to delete existing attachment.</param>
        /// <param name="listItemId">The id of the list item from which the attachment will be deleted.</param>
        /// <param name="url">Absolute URL of the attachment that should be deleted.</param>
        public void DeleteAttachment(string listName, string listItemId, string url)
        {
            if (null == this.listsProxy)
            {
                throw new InvalidOperationException("The Proxy instance is NULL, need to initialize the adapter");
            }

            this.listsProxy.DeleteAttachment(listName, listItemId, url);
            this.VerifyTransportRequirement();
        }

        /// <summary>
        /// The GetListItemChanges operation is used to retrieve the list items that have been inserted or updated
        /// since the specified date and time and matching the specified filter criteria.
        /// </summary>
        /// <param name="listName">The name of the list from which the list item changes will be got</param>
        /// <param name="viewFields">Indicates which fields of the list item SHOULD be returned</param>
        /// <param name="since">The date and time to start retrieving changes in the list
        /// If the parameter is null, protocol server should return all list items
        /// If the date that is passed in is not in UTC format, protocol server will use protocol server's local time zone and convert it to UTC time</param>
        /// <param name="contains">Restricts the results returned by giving a specific value to be searched for in the specified list item field</param>
        /// <returns>Return the list item change result</returns>
        public GetListItemChangesResponseGetListItemChangesResult GetListItemChanges(string listName, CamlViewFields viewFields, string since, CamlContains contains)
        {
            if (null == this.listsProxy)
            {
                throw new InvalidOperationException("The Proxy instance is NULL, need to initialize the adapter");
            }

            GetListItemChangesResponseGetListItemChangesResult result = null;
            result = this.listsProxy.GetListItemChanges(listName, viewFields, since, contains);
            this.VerifyTransportRequirement();

            return result;
        }

        /// <summary>
        /// The GetAttachmentCollection operation is used to retrieve information about all the lists on the current site.
        /// </summary>
        /// <param name="listName">A parameter represents the list name or GUID for returning the result.</param>
        /// <param name="listItemId">A parameter represents the identifier of the content type which will be collected.</param>
        /// <returns>Return attachment collection result.</returns>
        public GetAttachmentCollectionResponseGetAttachmentCollectionResult GetAttachmentCollection(string listName, string listItemId)
        {
            if (null == this.listsProxy)
            {
                throw new InvalidOperationException("The Proxy instance is NULL, need to initialize the adapter");
            }

            GetAttachmentCollectionResponseGetAttachmentCollectionResult result = null;

            result = this.listsProxy.GetAttachmentCollection(listName, listItemId);

            this.VerifyTransportRequirement();
            return result;
        }

        /// <summary>
        /// AddDiscussionBoardItem operation is used to add new discussion items to a specified discussion board.
        /// </summary>
        /// <param name="listName">The name of the discussion board in which the new item will be added</param>
        /// <param name="message">The message to be added to the discussion board. The message MUST be in MIME format and then Base64 encoded</param>
        /// <returns>AddDiscussionBoardItem Result</returns>
        public AddDiscussionBoardItemResponseAddDiscussionBoardItemResult AddDiscussionBoardItem(string listName, byte[] message)
        {
            if (null == this.listsProxy)
            {
                throw new InvalidOperationException("The Proxy instance is NULL, need to initialize the adapter");
            }

            AddDiscussionBoardItemResponseAddDiscussionBoardItemResult result = null;
            result = this.listsProxy.AddDiscussionBoardItem(listName, message);
            this.VerifyTransportRequirement();

            return result;
        }

        /// <summary>
        /// The DeleteList operation is used to delete the specified list from the specified site.
        /// </summary>
        /// <param name="listName">The name of the list which will be deleted</param>
        public void DeleteList(string listName)
        {
            if (null == this.listsProxy)
            {
                throw new InvalidOperationException("The Proxy instance is NULL, need to initialize the adapter");
            }

            this.listsProxy.DeleteList(listName);
            this.VerifyTransportRequirement();
        }

        /// <summary>
        /// The GetList operation is used to retrieve properties and fields for a specified list.
        /// </summary>
        /// <param name="listName">The name of the list from which information will be got</param>
        /// <returns>A return value represents the list definition.</returns>
        public GetListResponseGetListResult GetList(string listName)
        {
            if (null == this.listsProxy)
            {
                throw new InvalidOperationException("The Proxy instance is NULL, need to initialize the adapter");
            }

            GetListResponseGetListResult result = null;

            result = this.listsProxy.GetList(listName);
            this.VerifyTransportRequirement();
            this.VerifyCommonSchemaOfListDefinition(result);

            return result;
        }

        /// <summary>
        /// The GetListItemChangesSinceToken operation is used to return changes made to a specified list after the event
        /// expressed by the change token, if specified, or to return all the list items in the list.
        /// </summary>
        /// <param name="listName">The name of the list from which version collection will be got</param>
        /// <param name="viewName">The GUID refers to a view of the list</param>
        /// <param name="query">The query to determine which records from the list are to be 
        /// returned and the order in which they will be returned</param>
        /// <param name="viewFields">Specifies which fields of the list item will be returned</param>
        /// <param name="rowLimit">Indicate the maximum number of rows of data to return</param>
        /// <param name="queryOptions">Specifies various options for modifying the query</param>
        /// <param name="changeToken">Assigned a string comprising a token returned by a previous 
        /// call to this operation.</param>
        /// <param name="contains">Specifies a value to search for</param>
        /// <returns>A return value represent the list item changes since the specified token</returns>
        public GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult GetListItemChangesSinceToken(string listName, string viewName, GetListItemChangesSinceTokenQuery query, CamlViewFields viewFields, string rowLimit, CamlQueryOptions queryOptions, string changeToken, CamlContains contains)
        {
            if (null == this.listsProxy)
            {
                throw new InvalidOperationException("The Proxy instance is NULL, need to initialize the adapter");
            }

            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult result = null;
            result = this.listsProxy.GetListItemChangesSinceToken(listName, viewName, query, viewFields, rowLimit, queryOptions, changeToken, contains);
            this.VerifyTransportRequirement();
            this.VerifyGetListItemChangesSinceTokenResponse(result);

            return result;
        }

        /// <summary>
        /// The UpdateListItems operation is used to insert, update, and delete to specified list items in a list.
        /// </summary>
        /// <param name="listName">The name of the list for which list items will be updated</param>
        /// <param name="updates">Specifies the operations to perform on a list item</param>
        /// <returns>return the updated list items</returns>
        public UpdateListItemsResponseUpdateListItemsResult UpdateListItems(string listName, UpdateListItemsUpdates updates)
        {
            if (null == this.listsProxy)
            {
                throw new InvalidOperationException("The Proxy instance is NULL, need to initialize the adapter");
            }

            UpdateListItemsResponseUpdateListItemsResult result = null;
            result = this.listsProxy.UpdateListItems(listName, updates);
            this.VerifyTransportRequirement();

            return result;
        }

        /// <summary>
        ///  This operation used to get resource data Over HTTP protocol directly.
        /// </summary>
        /// <param name="requestResourceUrl">A parameter represents the resource where get data over HTTP protocol.</param>
        /// <param name="translateHeaderValue">A parameter represents the translate header which is used in HTTP request.</param>
        /// <returns>A return value represents the data get from the specified resource.</returns>
        public byte[] HTTPGET(Uri requestResourceUrl, string translateHeaderValue)
        {
            HttpWebRequest objRequest = null;
            HttpWebResponse objResponse = null;
            try
            {
                objRequest = (HttpWebRequest)HttpWebRequest.Create(requestResourceUrl);
                string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
                string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
                string userPassword = Common.GetConfigurationPropertyValue("Password", this.Site);
                objRequest.Credentials = new NetworkCredential(userName, userPassword, domain);
                objRequest.Method = WebRequestMethods.Http.Get;
                objRequest.Headers.Add("Translate", translateHeaderValue);
                objResponse = objRequest.GetResponse() as HttpWebResponse;
            }
            catch (WebException)
            {
                this.Site.Log.Add(LogEntryKind.Debug, "Get resource[{0}] with translateHeaderValue[Translate:{1}] fail.", requestResourceUrl, translateHeaderValue);
                throw;
            }

            long contentLength = objResponse.ContentLength; 

            // In MS-OUTSPS test scope, test suite only handle the data length in integer max value.
            // If the data is too large, it will impact the performance, and all test cases only use small attachment data.
            if (contentLength > int.MaxValue)
            {
                string errormessage = string.Format("Body length is too long, is larger than int MaxValue[{0}]", int.MaxValue.ToString());
                throw new InvalidOperationException(errormessage);
            }

            int readLength = (int)contentLength;
            byte[] dataTemp = new byte[readLength];

            using (BinaryReader binReader = new BinaryReader(objResponse.GetResponseStream()))
            {
                dataTemp = binReader.ReadBytes(readLength);
            }

            return dataTemp;
        }

        /// <summary>
        /// This operation used to put content data Over HTTP protocol directly.
        /// </summary>
        /// <param name="requestResourceUrl">A parameter represents the resource where put the data over HTTP protocol.</param>
        /// <param name="ifmatchHeader">>A parameter represents the IF-MATCH header which is used in HTTP request.</param>
        /// <param name="contentData">>A parameter represents the content data which is put to the SUT.</param>
        public void HTTPPUT(Uri requestResourceUrl, string ifmatchHeader, byte[] contentData)
        {
            HttpWebRequest objRequest = null;
            HttpWebResponse objResponse = null;
            try
            {
                objRequest = (HttpWebRequest)HttpWebRequest.Create(requestResourceUrl);
                string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
                string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
                string userPassword = Common.GetConfigurationPropertyValue("Password", this.Site);
                objRequest.Credentials = new NetworkCredential(userName, userPassword, domain);
                objRequest.Method = WebRequestMethods.Http.Put;
                objRequest.ContentLength = contentData.Length;
                objRequest.ContentType = "application/binary";
                objRequest.Headers.Add("IF-Match", ifmatchHeader);
                Stream stream = objRequest.GetRequestStream();
                stream.Write(contentData, 0, contentData.Length);
                objResponse = objRequest.GetResponse() as HttpWebResponse;
            }
            catch (WebException)
            {
                this.Site.Log.Add(LogEntryKind.Debug, @"Put content data to resource[{0}] with ""IF-Match"" header [IF-Match:{1}] fail.", requestResourceUrl, ifmatchHeader);
                throw;
            }

            StreamReader streamReader = new StreamReader(objResponse.GetResponseStream());
            string temp = string.Empty;
            while (streamReader.Peek() > -1)
            {
                string strInput = streamReader.ReadLine();
                temp += strInput;
            }

            streamReader.Close();
            this.Site.Log.Add(LogEntryKind.Debug, @"Put content data to resource[{0}] with ""IF-Match"" header [IF-Match:{1}] successfully.{2}", requestResourceUrl, ifmatchHeader, temp);
        }

        /// <summary>
        /// A method used to update list properties and add, remove, or update fields.
        /// </summary>
        /// <param name="listName">A parameter represents the name of the list which will be updated.</param>
        /// <param name="listProperties">A parameter represents the properties of the specified list.</param>
        /// <param name="newFields">A parameter represents new fields which are added to the list.</param>
        /// <param name="updateFields">A parameter represents the fields which are updated in the list.</param>
        /// <param name="deleteFields">A parameter represents the fields which are deleted from the list.</param>
        /// <param name="listVersion">A parameter represents an integer format string that specifies the current version of the list.</param>
        /// <returns>A return value represents the actual update result.</returns>
        public UpdateListResponseUpdateListResult UpdateList(string listName, UpdateListListProperties listProperties, UpdateListFieldsRequest newFields, UpdateListFieldsRequest updateFields, UpdateListFieldsRequest deleteFields, string listVersion)
        {
            if (null == this.listsProxy)
            {
                throw new InvalidOperationException("The Proxy instance is NULL, need to initialize the adapter");
            }

            UpdateListResponseUpdateListResult result = null;

            result = this.listsProxy.UpdateList(listName, listProperties, newFields, updateFields, deleteFields, listVersion);

            this.VerifyTransportRequirement();
            return result;
        }

        #endregion

        #region Override methods

        /// <summary>
        /// The Overridden Initialize method
        /// </summary>
        /// <param name="testSite">The ITestSite member of ManagedAdapterBase</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            // Set the protocol name of current test suite
            testSite.DefaultProtocolDocShortName = "MS-OUTSPS";

            // Initialize the ListsSoap proxy class instance without the schema validation.
            this.listsProxy = Proxy.CreateProxy<ListsSoap>(this.Site, false, false, true);

            // Merge the common configuration into local configuration
            string conmmonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);
            Common.MergeGlobalConfig(conmmonConfigFileName, this.Site);

            Common.CheckCommonProperties(this.Site, true);

            // Merge the SHOULDMAY configuration file.
            Common.MergeSHOULDMAYConfig(this.Site);

            this.listsProxy.Url = this.GetTargetServiceUrl();
            string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            string password = Common.GetConfigurationPropertyValue("Password", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.listsProxy.Credentials = new NetworkCredential(userName, password, domain);
            this.SetSoapVersion();
            if (TransportProtocol.HTTPS == Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site))
            {
               Common.AcceptServerCertificate();
            }

            // Configure the service timeout.
            int soapTimeOut = Common.GetConfigurationPropertyValue<int>("ServiceTimeOut", this.Site);

            // 60000 means the configure SOAP Timeout is in minute.
            this.listsProxy.Timeout = soapTimeOut * 60000;
        }

        #endregion

        #region Private helper methods

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
                        this.listsProxy.SoapVersion = SoapProtocolVersion.Soap11;
                        break;
                    }

                default:
                    {
                        this.listsProxy.SoapVersion = SoapProtocolVersion.Soap12;
                        break;
                    }
            }
        }

        /// <summary>
        /// A method used to Get target service fully qualified URL, it indicates which site the test suite will run on.
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