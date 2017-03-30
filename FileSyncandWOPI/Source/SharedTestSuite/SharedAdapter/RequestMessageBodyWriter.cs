namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.ServiceModel.Channels;
    using System.Xml;
    using System.Xml.Serialization;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class is used to construct the body of the request message.
    /// </summary>
    public class RequestMessageBodyWriter : BodyWriter
    {
        /// <summary>
        /// Specify the request of envelop body which will be send to the server.
        /// </summary>
        private EnvelopeBody requestEnvelope;

        /// <summary>
        /// Initializes a new instance of the RequestMessageBodyWriter class with the specified version and minor version type number.
        /// </summary>
        /// <param name="version">Specify the version type number.</param>
        /// <param name="minorVersion">Specify the minor version type number.</param>
        public RequestMessageBodyWriter(ushort? version, ushort? minorVersion)
            : base(true)
        {
            this.CreateEnvelopeBody(version, minorVersion);
        }

        /// <summary>
        /// Gets the message body XML.
        /// </summary>
        public XmlElement MessageBodyXml
        {
            get
            {
                MemoryStream ms = new MemoryStream();
                XmlSerializer serializer = new XmlSerializer(typeof(EnvelopeBody));
                serializer.Serialize(ms, this.requestEnvelope);

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(System.Text.ASCIIEncoding.ASCII.GetString(ms.ToArray(), 0, ms.ToArray().Length));
                ms.Dispose();
                return doc.DocumentElement;
            }
        }

        /// <summary>
        /// This method is used to add a request to the request collection in the request envelope body.
        /// </summary>
        /// <param name="url">Specify the URL of the file to edit.</param>
        /// <param name="subRequests">Specify the sub request array.</param>
        /// <param name="requestToken">Specify the token which uniquely identify the request.</param>
        /// <param name="interval">Specify a nonnegative integer in seconds, which the protocol client will repeat this request.</param>
        /// <param name="metaData">Specify a 32-bit value that specifies information about the scenario and urgency of the request.</param>
        /// <param name="lastModifiedTime">Specify the last modified time, which is expressed as a tick count.</param>
        /// <param name="parentFolderResourceID">If UseResourceID is true, this parameter tells the host to create a file in the given folder ResourceID, regardless of the request URL value.</param>
        /// <param name="shouldReturnDisambiguatedFileName">If an upload request fails with a coherency failure, this flag specifies whether the host should return a suggested/available file name that the client can try instead</param>
        /// <param name="resourceID">Specify the invariant ResourceID for a file that uniquely identifies the file whose response is being generated</param>
        /// <param name="useResourceID">Specify if the protocol server MAY perform ResourceID specific behavior for the file whose contents or metadata contents are requested for uploading to the server or downloading from the server. </param>
        public void AddRequest(string url, SubRequestType[] subRequests, string requestToken, uint? interval, int? metaData,
            string lastModifiedTime = null, 
            string parentFolderResourceID = null,
            bool? shouldReturnDisambiguatedFileName = null,
            string resourceID = null,
            bool? useResourceID = null)
        {
            Request request = new Request();
            request.RequestToken = requestToken;
            request.Url = url;
            request.Interval = interval == null ? null : interval.Value.ToString();
            request.MetaData = metaData == null ? null : metaData.Value.ToString();
            request.ParentFolderResourceID = parentFolderResourceID;
            request.ResourceID = resourceID;
            if (shouldReturnDisambiguatedFileName != null)
            {
                request.ShouldReturnDisambiguatedFileName = (bool)shouldReturnDisambiguatedFileName;
                request.ShouldReturnDisambiguatedFileNameSpecified = true;
            }
            if (useResourceID != null)
            {
                request.UseResourceID = useResourceID.ToString();
            }

            int index = 0;
            if (subRequests != null)
            {
                request.SubRequest = new SubRequestElementGenericType[subRequests.Length];
                foreach (SubRequestType item in subRequests)
                {
                    if (item != null)
                    {
                        int temp = index;
                        request.SubRequest[temp] = FsshttpConverter.ConvertSubRequestToGenericType<SubRequestElementGenericType>(item);
                        if (!string.IsNullOrEmpty(lastModifiedTime))
                        {
                            if (request.SubRequest[temp].SubRequestData == null)
                            {
                                request.SubRequest[temp].SubRequestData = new SubRequestDataGenericType();
                            }
                            request.SubRequest[temp].SubRequestData.LastModifiedTime = lastModifiedTime;
                        }

                        index++;
                    }
                }
            }
            else
            {
                throw new System.ArgumentException("subRequests parameter in FSSHTTPMessageBodyWriter::AddRequest cannot be null.");
            }

            List<Request> requestList = this.requestEnvelope.RequestCollection.Request == null ? new List<Request>(1) : new List<Request>(this.requestEnvelope.RequestCollection.Request);
            requestList.Add(request);

            this.requestEnvelope.RequestCollection.Request = requestList.ToArray();
        }

        /// <summary>
        /// Override the method to write the content to the XML dictionary writer.
        /// </summary>
        /// <param name="writer">Specify the output destination of the content.</param>
        protected override void OnWriteBodyContents(XmlDictionaryWriter writer)
        {
            MemoryStream ms = new MemoryStream();
            XmlSerializer serializer = new XmlSerializer(typeof(EnvelopeBody));
            serializer.Serialize(ms, this.requestEnvelope);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(System.Text.ASCIIEncoding.ASCII.GetString(ms.ToArray(), 0, ms.ToArray().Length));
            ms.Dispose();
            foreach (XmlNode node in doc.LastChild.ChildNodes)
            {
                if (node.Name == "RequestVersion")
                {
                    writer.WriteRaw(node.OuterXml);
                }
                else if (node.Name == "RequestCollection")
                {
                    this.WriteNode(node, writer);
                }
                else
                {
                    throw new InvalidOperationException(string.Format("this element [{0}] is not expected element", node.Name));
                }
            }
        }

        /// <summary>
        /// This method is used to write a XML node to a XML writer.
        /// </summary>
        /// <param name="node">Specify the XML node.</param>
        /// <param name="writer">Specify the XML writer.</param>
        private void WriteNode(XmlNode node, XmlDictionaryWriter writer)
        {
            writer.WriteStartElement(node.Name);
            foreach (XmlAttribute xmlAttribute in node.Attributes)
            {
                writer.WriteAttributeString(xmlAttribute.Name, xmlAttribute.Value);
                if (!string.IsNullOrEmpty(xmlAttribute.Prefix))
                {
                    writer.WriteXmlnsAttribute(xmlAttribute.Prefix, xmlAttribute.NamespaceURI);
                }
            }

            if (node.Name == "SubRequestData")
            {
                string base64 = node.InnerText;
                byte[] bytes = Convert.FromBase64String(base64);
                writer.WriteBase64(bytes, 0, bytes.Length);
            }
            else
            {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    this.WriteNode(childNode, writer);
                }
            }

            writer.WriteEndElement();
        }

        /// <summary>
        /// This method is used to create request envelope message body, but not include the RequestCollection.
        /// </summary>
        /// <param name="versionTypeNumber">Specify the version type number.</param>
        /// <param name="minorVersionTypeNumber">Specify the minor version type number.</param>
        private void CreateEnvelopeBody(ushort? versionTypeNumber, ushort? minorVersionTypeNumber)
        {
            this.requestEnvelope = new EnvelopeBody();
            this.requestEnvelope.RequestVersion = this.CreateRequestVersion(versionTypeNumber, minorVersionTypeNumber);
            this.requestEnvelope.RequestCollection = new RequestCollection();
            this.requestEnvelope.RequestCollection.CorrelationId = Guid.NewGuid().ToString(); 
        }

        /// <summary>
        /// This method is used to create the VersionType instance using the specified version type number and minor version type number
        /// </summary>
        /// <param name="versionTypeNumber">Specify the version type number, which can only be 2.</param>
        /// <param name="minorVersionTypeNumber">Specify the minor version type number, which could be 0 or 2.</param>
        /// <returns>Returns the VersionType instance.</returns>
        private VersionType CreateRequestVersion(ushort? versionTypeNumber, ushort? minorVersionTypeNumber)
        {
            if (versionTypeNumber == null || minorVersionTypeNumber == null)
            {
                return null;
            }
            else
            {
                VersionType versionType = new VersionType();

                // The version number type can only be 2.
                versionType.Version = versionTypeNumber.Value;

                // The minor version type number can 0 or 2.
                versionType.MinorVersion = minorVersionTypeNumber.Value;
                return versionType;
            }
        }
    }
}