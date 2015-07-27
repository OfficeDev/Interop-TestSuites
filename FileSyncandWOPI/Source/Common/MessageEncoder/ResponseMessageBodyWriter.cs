//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System.IO;
    using System.ServiceModel.Channels;
    using System.Text;

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
            writer.WriteRaw(this.encoding.GetString(bytes));
        }
    }
}