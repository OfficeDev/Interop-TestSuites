namespace Microsoft.Protocols.TestSuites.Common
{
    using System.IO;
    using System.ServiceModel.Channels;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// A class is used to construct the body of the response message.
    /// </summary>
    public class ResponseMessageBodyWriter : BodyWriter
    {
        /// <summary>
        ///  Specify the response body stream.
        /// </summary>
        private Stream stream;
        
        /// <summary>
        /// Specify the response body encoding type.
        /// </summary>
        private Encoding encoding;

        /// <summary>
        /// Initializes a new instance of the ResponseMessageBodyWriter class with the specified encoding type and body stream.
        /// </summary>
        /// <param name="encoding">Specify the encoding type.</param>
        /// <param name="bodyStream">Specify the body stream.</param>
        public ResponseMessageBodyWriter(Encoding encoding, Stream bodyStream) : base(true)
        {
            this.stream = bodyStream;
            this.encoding = encoding;
        }

        /// <summary>
        /// Override the method to write the content to the xml dictionary writer.
        /// </summary>
        /// <param name="writer">Specify the output destination of the content.</param>
        protected override void OnWriteBodyContents(System.Xml.XmlDictionaryWriter writer)
        {
            byte[] bytes = new byte[this.stream.Length];
            this.stream.Position = 0;
            this.stream.Read(bytes, 0, bytes.Length);
            string content = this.encoding.GetString(bytes);

            if (content.ToLowerInvariant().Contains("Content-Type: multipart/related".ToLowerInvariant()))
            {
                using (XmlDictionaryReader reader = XmlDictionaryReader.CreateMtomReader(bytes, 0, (int)bytes.Length, this.encoding, XmlDictionaryReaderQuotas.Max))
                {
                    XmlDocument msgDoc = new XmlDocument();
                    msgDoc.PreserveWhitespace = true;
                    msgDoc.Load(reader);
                    content = msgDoc.OuterXml;
                }   
            }
            writer.WriteRaw(content);
        }
    }
}