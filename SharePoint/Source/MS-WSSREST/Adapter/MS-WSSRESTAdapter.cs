namespace Microsoft.Protocols.TestSuites.MS_WSSREST
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Text;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-WSSREST.
    /// </summary>
    public partial class MS_WSSRESTAdapter : ManagedAdapterBase, IMS_WSSRESTAdapter
    {
        #region Variables

        /// <summary>
        /// Proxy class service.
        /// </summary>
        private MS_WSSREST service = new MS_WSSREST();

        /// <summary>
        /// The SUT control adapter.
        /// </summary>
        private IMS_WSSRESTSUTControlAdapter sutAdapter;

        #endregion Variables

        #region Initialize TestSuite

        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-WSSREST";

            // Load Common configuration
            Common.MergeGlobalConfig(Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site), this.Site);

            // Load SHOULDMAY configuration
            Common.MergeSHOULDMAYConfig(this.Site);

            this.service.Initialize(this.Site);
            this.sutAdapter = this.Site.GetAdapter<IMS_WSSRESTSUTControlAdapter>();
            this.sutAdapter.Initialize(testSite);
            AdapterHelper.Initialize(testSite);
        }

        #endregion Initialize TestSuite

        #region MS-WSSRESTAdapter Members

        /// <summary>
        /// Insert a list item.
        /// </summary>
        /// <param name="request">The content of the list item that be inserted.</param>
        /// <returns>The list item that be inserted.</returns>
        public Entry InsertListItem(Request request)
        {
            HttpWebResponse response = this.service.SendMessage(HttpMethod.POST, request);
            XmlDocument doc = AdapterHelper.GetXmlData(response);
            SchemaValidation.ValidateXml(this.Site, doc.OuterXml);
            this.CaptureTransportRelatedRequirements();
            this.ValidateAndCaptureSchemaValidation();
            List<Entry> result = AdapterHelper.AnalyseResponse(doc);
            Site.Assert.IsNotNull(result, "Verify the result is not null");
            Site.Assert.AreEqual<int>(result.Count, 1, "The response of insert list item should only include one entry result.");

            return result[0];
        }

        /// <summary>
        /// Update a list item.
        /// </summary>
        /// <param name="request">The content of the list item that be updated.</param>
        /// <returns>The ETag of this list item.</returns>
        public string UpdateListItem(Request request)
        {
            HttpMethod method;

            if (request.UpdateMethod == UpdateMethod.PUT)
            {
                method = HttpMethod.PUT;
            }
            else
            {
                method = HttpMethod.MERGE;
            }

            HttpWebResponse response = this.service.SendMessage(method, request);
            Site.Assert.IsNotNull(response, "Verify the response is not null");

            this.CaptureTransportRelatedRequirements();

            return response.Headers.Get("ETag");
        }

        /// <summary>
        /// Retrieve list item from server.
        /// </summary>
        /// <param name="request">The retrieve condition.</param>
        /// <returns>The response from server.</returns>
        public object RetrieveListItem(Request request)
        {
            HttpWebResponse response = this.service.SendMessage(HttpMethod.GET, request);

            if (request.Parameter.Contains("$count"))
            {
                string count = AdapterHelper.GetResponseContent(response.GetResponseStream());
                this.CaptureTransportRelatedRequirements();

                return count;
            }
            else if (request.Parameter.Contains("$metadata"))
            {
                XmlDocument doc = AdapterHelper.GetXmlData(response);
                SchemaValidation.ValidateXml(this.Site, doc.OuterXml);
                this.ValidateRetrieveCSDLDocument(doc);

                return doc;
            }
            else
            {
                XmlDocument doc = AdapterHelper.GetXmlData(response);
                SchemaValidation.ValidateXml(this.Site, doc.OuterXml);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSchemaValidation();
                return AdapterHelper.AnalyseResponse(doc);
            }
        }

        /// <summary>
        /// Delete a special list item.
        /// </summary>
        /// <param name="request">The special list item.</param>
        /// <returns>True if the list item be deleted success, otherwise false.</returns>
        public bool DeleteListItem(Request request)
        {
            try
            {
                this.service.SendMessage(HttpMethod.DELETE, request);
                this.CaptureTransportRelatedRequirements();

                return true;
            }
            catch (WebException webEx)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [DeleteListItem] method:\r\n{0}",
                                webEx.Message);

                return false;
            }
        }

        /// <summary>
        /// Package many requests(insert,update or delete request) in one batch request.
        /// </summary>
        /// <param name="requests">The multi requests.</param>
        /// <returns>The response from server.</returns>
        public string BatchRequests(List<BatchRequest> requests)
        {
            string batchID = Guid.NewGuid().ToString();
            string changesetID = Guid.NewGuid().ToString();
            int contentID = 1;

            Request request = new Request();
            request.Parameter = "$batch";
            request.ContentType = string.Format("multipart/mixed; boundary=batch_{0}", batchID);

            StringBuilder batchBody = new StringBuilder();
            batchBody.AppendLine(string.Format("--batch_{0}", batchID));
            batchBody.AppendLine(string.Format("Content-Type: multipart/mixed; boundary=changeset_{0}", changesetID));
            batchBody.AppendLine();

            foreach (BatchRequest br in requests)
            {
                string method = string.Empty;
                switch (br.OperationType)
                {
                    case OperationType.Insert:
                        method = HttpMethod.POST.ToString();
                        break;
                    case OperationType.Delete:
                        method = HttpMethod.DELETE.ToString();
                        break;
                    case OperationType.Update:
                        method = br.UpdateMethod.ToString();
                        break;
                }

                batchBody.AppendLine(string.Format("--changeset_{0}", changesetID));
                batchBody.AppendLine("Content-Type: application/http");
                batchBody.AppendLine("Content-Transfer-Encoding: binary");
                batchBody.AppendLine();
                batchBody.AppendLine(string.Format("{0} {1}/{2} HTTP/1.1", method, Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site), br.Parameter));
                batchBody.AppendLine(string.Format("Content-ID: {0}", contentID++));

                if (!string.IsNullOrEmpty(br.ContentType))
                {
                    batchBody.AppendLine(string.Format("Content-Type: {0}", br.ContentType));
                }

                if (br.OperationType == OperationType.Update || br.OperationType == OperationType.Delete)
                {
                    batchBody.AppendLine(string.Format("If-Match : {0}", br.ETag));
                }

                batchBody.AppendLine();

                if (br.OperationType == OperationType.Insert || br.OperationType == OperationType.Update)
                {
                    batchBody.AppendLine(br.Content);
                    batchBody.AppendLine();
                }
            }

            batchBody.AppendLine(string.Format("--changeset_{0}--", changesetID));
            batchBody.Append(string.Format("--batch_{0}--\r\n", batchID));

            request.Content = batchBody.ToString();

            HttpWebResponse response = this.service.SendMessage(HttpMethod.POST, request);
            Site.Assert.IsNotNull(response, "Verify the response is not null");

            string result = AdapterHelper.GetResponseContent(response.GetResponseStream());
            response.Close();

            this.CaptureTransportRelatedRequirements();

            return result;
        }

        #endregion MS-WSSRESTAdapter Members
    }
}