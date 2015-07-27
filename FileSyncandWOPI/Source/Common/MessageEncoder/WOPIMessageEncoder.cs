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
    using System;
    using System.IO;
    using System.ServiceModel.Channels;
    using System.Text;

    /// <summary>
    /// WOPIMessageEncoder is the component that is used to write the XML corresponding with the MS-FSSHTTP type definition to a payload in a HTTP post request and to read the XML from a payload when invoking the ExecuteCellStorageRequest and ExecuteCellStorageRelativeRequest for MS-WOPI server.
    /// </summary>
    public class WOPIMessageEncoder : MessageEncoder
    {
        /// <summary>
        /// Specify the inner encoder for the text message type.
        /// </summary>
        private MessageEncoder innerEncoder;

        /// <summary>
        /// Specify the message encoder factory.
        /// </summary>
        private WOPIMessageEncoderFactory factory;

        /// <summary>
        /// Specify the content type for the message.
        /// </summary>
        private string contentType;

        /// <summary>
        /// Initializes a new instance of the WOPIMessageEncoder class with the specified message encoder factory.
        /// </summary>
        /// <param name="factory">Specify the message encoder factory.</param>
        public WOPIMessageEncoder(WOPIMessageEncoderFactory factory)
        {
            TextMessageEncodingBindingElement element = new TextMessageEncodingBindingElement();
            element.MessageVersion = factory.MessageVersion;
            element.WriteEncoding = Encoding.GetEncoding(factory.CharSet);

            this.innerEncoder = element.CreateMessageEncoderFactory().Encoder;
            this.factory = factory;
            this.contentType = string.Format("{0}; charset={1}", "text/xml", this.factory.CharSet);
        }

        /// <summary>
        /// Gets the content type.
        /// </summary>
        public override string ContentType
        {
            get { return this.contentType; }
        }

        /// <summary>
        /// Gets the media type.
        /// </summary>
        public override string MediaType
        {
            get { return "text/xml"; }
        }

        /// <summary>
        /// Gets the message version.
        /// </summary>
        public override MessageVersion MessageVersion
        {
            get { return this.factory.MessageVersion; }
        }

        /// <summary>
        /// This method is used to read a message from a specified buffer.
        /// </summary>
        /// <param name="buffer">Specify the buffer.</param>
        /// <param name="bufferManager">Specify the buffer manager. </param>
        /// <param name="messageContentType">Specify the content type.</param>
        /// <returns>Return the System.ServiceModel.Channels.Message that is read from the buffer specified.</returns>
        public override Message ReadMessage(ArraySegment<byte> buffer, BufferManager bufferManager, string messageContentType)
        {
            byte[] msgContents = new byte[buffer.Count];
            Array.Copy(buffer.Array, buffer.Offset, msgContents, 0, msgContents.Length);
            bufferManager.ReturnBuffer(buffer.Array);

            MemoryStream stream = new MemoryStream(msgContents);
            stream.Position = 0;
            return this.ReadMessage(stream, int.MaxValue);
        }

        /// <summary>
        /// This method is used to read a message from a specified stream.
        /// </summary>
        /// <param name="stream">Specify the stream.</param>
        /// <param name="maxSizeOfHeaders">Specify the maximum size of the headers that can be read from the message</param>
        /// <param name="messageContentType">Specify the content type.</param>
        /// <returns>Return the System.ServiceModel.Channels.Message that is read from the stream specified.</returns>
        public override Message ReadMessage(System.IO.Stream stream, int maxSizeOfHeaders, string messageContentType)
        {
            Encoding encoding = Encoding.GetEncoding(this.factory.CharSet);
            ResponseMessageBodyWriter writer = new ResponseMessageBodyWriter(encoding, stream);
            Message message = Message.CreateMessage(this.MessageVersion, @"http://schemas.microsoft.com/sharepoint/soap/ICellStorages/ExecuteCellStorageRequestResponse", writer);
            return message;
        }

        /// <summary>
        /// This method is used to write a message less than a specified size to a byte array buffer at the specified offset using the inner message encoder.
        /// </summary>
        /// <param name="message">Specify the message which is needed to be written into buffer.</param>
        /// <param name="maxMessageSize">Specify the max size of the message.</param>
        /// <param name="bufferManager">Specify the buffer manager.</param>
        /// <param name="messageOffset">Specify the offset of the segment that begins from the start of the byte array that provides the buffer.</param>
        /// <returns>Return a System.ArraySegment of type byte that provides the buffer to which the message is serialized</returns>
        public override ArraySegment<byte> WriteMessage(Message message, int maxMessageSize, BufferManager bufferManager, int messageOffset)
        {
            return this.innerEncoder.WriteMessage(message, maxMessageSize, bufferManager, messageOffset);
        }

        /// <summary>
        /// This method is used to write a message to a specified stream using the inner message encoder.
        /// </summary>
        /// <param name="message">Specify the message which is needed to be written into buffer.</param>
        /// <param name="stream">Specify the written stream.</param>
        public override void WriteMessage(Message message, System.IO.Stream stream)
        {
            this.innerEncoder.WriteMessage(message, stream);
        }

        /// <summary>
        /// This method is used to support the text/html content type.
        /// </summary>
        /// <param name="messageLevelContentType">Specify the message-level content-type being tested.</param>
        /// <returns>Return true if the message-level content-type specified is supported; otherwise false.</returns>
        public override bool IsContentTypeSupported(string messageLevelContentType)
        {
            if (string.Compare("text/html", messageLevelContentType, StringComparison.OrdinalIgnoreCase) == 0)
            {
                return true;
            }

            return this.innerEncoder.IsContentTypeSupported(messageLevelContentType);
        }
    }
}