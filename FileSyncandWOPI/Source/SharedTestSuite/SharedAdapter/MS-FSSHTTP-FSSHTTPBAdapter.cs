//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.ServiceModel;
    using System.ServiceModel.Channels;
    using System.Xml;
    using System.Xml.Serialization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Class MS_FSSHTTP_FSSHTTPB_Adapter implement interface IFSSHTTPFSSHTTPB
    /// </summary>
    public partial class MS_FSSHTTP_FSSHTTPBAdapter : ManagedAdapterBase, IMS_FSSHTTP_FSSHTTPBAdapter
    {
        /// <summary>
        /// The constant string action URL.
        /// </summary>
        public const string ActionURL = "http://schemas.microsoft.com/sharepoint/soap/ICellStorages/ExecuteCellStorageRequest";

        #region Variable

        /// <summary>
        /// Specify the static channel manager.
        /// </summary>
        private static SharedChannelManager channelManager;

        /// <summary>
        /// The field of sub-response validation wrapper list.
        /// </summary>
        private List<SubResponseValidationWrapper> subResponseValidationWrappers = new List<SubResponseValidationWrapper>();

        /// <summary>
        /// An Instance of CellStorages Service Client.
        /// </summary>
        private CellStoragesClient cellStoragesProxy;

        /// <summary>
        /// The last request XML.
        /// </summary>
        private XmlElement lastRawRequestXml;

        /// <summary>
        /// The last response XML.
        /// </summary>
        private XmlElement lastRawResponseXml;
        #endregion 

        /// <summary>
        /// Initializes static members of the MS_FSSHTTP_FSSHTTPBAdapter class.
        /// </summary>
        static MS_FSSHTTP_FSSHTTPBAdapter()
        {
            channelManager = new SharedChannelManager();
        }

        /// <summary>
        /// Gets the raw XML request sent to protocol SUT
        /// </summary>
        public XmlElement LastRawRequestXml
        {
            get { return this.lastRawRequestXml; }
        }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT
        /// </summary>
        public XmlElement LastRawResponseXml
        {
            get { return this.lastRawResponseXml; }
        }

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
        /// <returns>Returns the CellStorageResponse message received from the server.</returns>
        public CellStorageResponse CellStorageRequest(
                string url, 
                SubRequestType[] subRequests, 
                string requestToken = "1", 
                ushort? version = 2, 
                ushort? minorVersion = 2, 
                uint? interval = null, 
                int? metaData = null)
        {
            // If the transport is HTTPS, then try to accept the certificate.
            if ("HTTPS".Equals(Common.GetConfigurationPropertyValue("TransportType", this.Site), StringComparison.OrdinalIgnoreCase))
            {
                Common.AcceptServerCertificate();
            }

            ICellStoragesChannel channel = channelManager.CreateChannel<ICellStoragesChannel>(SharedContext.Current);

            // Store the SubRequestToken, subRequest type pairs
            this.subResponseValidationWrappers.Clear();
            if (subRequests != null)
            {
                foreach (SubRequestType item in subRequests)
                {
                    this.subResponseValidationWrappers.Add(new SubResponseValidationWrapper { SubToken = item.SubRequestToken, SubRequestType = item.GetType().Name });
                }
            }

            // Create web request message.
            RequestMessageBodyWriter fsshttpBodyWriter = new RequestMessageBodyWriter(version, minorVersion);
            fsshttpBodyWriter.AddRequest(url, subRequests, requestToken, interval, metaData);
            this.lastRawRequestXml = fsshttpBodyWriter.MessageBodyXml;

            // Try to log the request body information
            this.Site.Log.Add(LogEntryKind.Debug, "The raw xml request message is:\r\n{0}", this.lastRawRequestXml.OuterXml);

            Message request = Message.CreateMessage(MessageVersion.Soap11, ActionURL, fsshttpBodyWriter);

            try
            {
                // Invoke the web service
                Message response = channel.ExecuteCellStorageRequest(request);

                // Extract and de-serialize the response.
                CellStorageResponse cellStorageResponseObjects = this.GetCellStorageResponseObject(response);

                // Restore the current version type
                SharedContext.Current.CellStorageVersionType = cellStorageResponseObjects.ResponseVersion;

                // Schema Validation for the response.
                this.ValidateGenericType(cellStorageResponseObjects, requestToken);
                this.ValidateSpecificType(cellStorageResponseObjects);

                request.Close();

                return cellStorageResponseObjects;
            }
            catch (EndpointNotFoundException ex)
            {
                // Here try to catch the EndpointNotFoundException due to X-WOPI-ServerError.
                if (ex.InnerException is WebException)
                {
                    if ((ex.InnerException as WebException).Response.Headers.AllKeys.Contains<string>("X-WOPI-ServerError"))
                    {
                        throw new WOPIServerErrorException(
                            (ex.InnerException as System.Net.WebException).Response.Headers["X-WOPI-ServerError"],
                            ex.InnerException);
                    }
                }

                request.Close();

                throw;
            }
        }

        /// <summary>
        /// Override the Reset function.
        /// </summary>
        public override void Reset()
        {
            base.Reset();

            // Clear the ms-fsshttpb sub request mapping
            MsfsshttpbSubRequestMapping.Clear();

            // Gracefully close the proxy.
            this.CloseProxy();
        }

        /// <summary>
        /// Used to parse the response message and retrieve the CellStorageResponse object
        /// </summary>
        /// <param name="responseMsg">Response message</param>
        /// <returns>CellStorageResponse object</returns>
        private CellStorageResponse GetCellStorageResponseObject(Message responseMsg)
        {
            CellStorageResponse cellStorageResponse = new CellStorageResponse();

            XmlDocument xmlDoc = new XmlDocument();
            MemoryStream ms = null;
            try
            {
                ms = new MemoryStream();
                using (XmlWriter xw = XmlWriter.Create(ms))
                {
                    responseMsg.WriteMessage(xw);
                    xw.Flush();
                    ms.Flush();
                    ms.Position = 0;
                    xmlDoc.Load(ms);
                }
            }
            finally
            {
                if (ms != null)
                {
                    ms.Dispose();
                }
            }

            this.lastRawResponseXml = xmlDoc.DocumentElement;

            // Try to log the response body information
            this.Site.Log.Add(LogEntryKind.Debug, "The raw xml response message is:\r\n{0}", this.lastRawResponseXml.OuterXml);

            // De-serialize the ResponseVersion node
            XmlNodeList responseVersionNodeList = xmlDoc.GetElementsByTagName("ResponseVersion");
            if (responseVersionNodeList.Count > 0)
            {
                XmlSerializer responseVersionNodeSerializer = new XmlSerializer(typeof(XmlNode));
                MemoryStream responseVersionMemStream = new MemoryStream();

                responseVersionNodeSerializer.Serialize(responseVersionMemStream, responseVersionNodeList[0]);
                responseVersionMemStream.Position = 0;

                XmlSerializer responseVersionSerializer = new XmlSerializer(typeof(ResponseVersion));
                cellStorageResponse.ResponseVersion = (ResponseVersion)responseVersionSerializer.Deserialize(responseVersionMemStream);

                responseVersionMemStream.Dispose();
            }

            // De-serialize the ResponseCollection node.
            XmlNodeList responseCollectionNodeList = xmlDoc.GetElementsByTagName("ResponseCollection");
            if (responseCollectionNodeList.Count > 0)
            {
                XmlSerializer responseCollectionNodeSerializer = new XmlSerializer(typeof(XmlNode));
                MemoryStream responseCollectionMemStream = new MemoryStream();

                responseCollectionNodeSerializer.Serialize(responseCollectionMemStream, responseCollectionNodeList[0]);
                responseCollectionMemStream.Position = 0;

                XmlSerializer responseCollectionSerializer = new XmlSerializer(typeof(ResponseCollection));
                cellStorageResponse.ResponseCollection = (ResponseCollection)responseCollectionSerializer.Deserialize(responseCollectionMemStream);

                responseCollectionMemStream.Dispose();
            }

            return cellStorageResponse;
        }
        
        /// <summary>
        /// Close and dispose the proxy.
        /// </summary>
        private void CloseProxy()
        {
            if (this.cellStoragesProxy != null && this.cellStoragesProxy.InnerChannel != null)
            {
                this.cellStoragesProxy.InnerChannel.Close();
                ((IDisposable)this.cellStoragesProxy).Dispose();
            }

            this.cellStoragesProxy = null;
        }

        /// <summary>
        /// Validate the generic response type from the server.
        /// </summary>
        /// <param name="cellStorageResponseObjects">The server returned CellStorageResponse instance.</param>
        /// <param name="requestToken">The expected RequestToken</param>
        private void ValidateGenericType(CellStorageResponse cellStorageResponseObjects, string requestToken)
        {
            // Do the generic schema validation based on the different operation type.
            if (SharedContext.Current.OperationType == OperationType.FSSHTTPCellStorageRequest)
            {
                SchemaValidation.ValidateXml(this.lastRawResponseXml.OuterXml);

                // Validate the generic part schema
                if (SchemaValidation.ValidationResult == ValidationResult.Success)
                {
                    MsfsshttpAdapterCapture.ValidateResponse(cellStorageResponseObjects, requestToken, this.Site);
                    MsfsshttpAdapterCapture.ValidateTransport(this.Site);
                }
                else
                {
                    this.Site.Assert.Fail("The schema validation fails, the reason is " + SchemaValidation.GenerateValidationResult());
                }
            }
            else
            {
                XmlNodeList responseVersionNodeList = this.lastRawResponseXml.GetElementsByTagName("ResponseVersion");
                if (responseVersionNodeList.Count > 0)
                {
                    SchemaValidation.ValidateXml(responseVersionNodeList[0].OuterXml);

                    if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                    {
                        if (SchemaValidation.ValidationResult == ValidationResult.Success)
                        {
                            MsfsshttpAdapterCapture.ValidateResponseVersion(cellStorageResponseObjects.ResponseVersion, this.Site);
                        }
                        else
                        {
                            this.Site.Assert.Fail("The schema validation fails, the reason is " + SchemaValidation.GenerateValidationResult());
                        }
                    }
                    else
                    {
                        if (SchemaValidation.ValidationResult != ValidationResult.Success)
                        {
                            this.Site.Assert.Fail("The schema validation fails, the reason is " + SchemaValidation.GenerateValidationResult());
                        }
                    }
                }

                // De-serialize the ResponseCollection node.
                XmlNodeList responseCollectionNodeList = this.lastRawResponseXml.GetElementsByTagName("ResponseCollection");
                if (responseCollectionNodeList.Count > 0)
                {
                    SchemaValidation.ValidateXml(responseCollectionNodeList[0].OuterXml);

                    if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                    {
                        if (SchemaValidation.ValidationResult == ValidationResult.Success)
                        {
                            MsfsshttpAdapterCapture.ValidateResponseCollection(cellStorageResponseObjects.ResponseCollection, requestToken, this.Site);
                        }
                        else
                        {
                            this.Site.Assert.Fail("The schema validation fails, the reason is " + SchemaValidation.GenerateValidationResult());
                        }
                    }
                    else
                    {
                        if (SchemaValidation.ValidationResult != ValidationResult.Success)
                        {
                            this.Site.Assert.Fail("The schema validation fails, the reason is " + SchemaValidation.GenerateValidationResult());
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Validate the specific response type from the server.
        /// </summary>
        /// <param name="cellStorageResponseObjects">The server returned CellStorageResponse instance.</param>
        private void ValidateSpecificType(CellStorageResponse cellStorageResponseObjects)
        {
            // Validate each specific sub response type.
            if (cellStorageResponseObjects.ResponseCollection != null && cellStorageResponseObjects.ResponseCollection.Response != null && cellStorageResponseObjects.ResponseCollection.Response.Length != 0)
            {
                foreach (SubResponseValidationWrapper subResponseValidationWrapper in this.subResponseValidationWrappers)
                {
                    // Try to validate sub response schema
                    subResponseValidationWrapper.Validate(this.lastRawResponseXml.OuterXml, this.Site);
                }
            }
        }
    }
}