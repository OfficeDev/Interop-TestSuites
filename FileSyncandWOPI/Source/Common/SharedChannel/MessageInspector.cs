namespace Microsoft.Protocols.TestSuites.Common
{
    using System.IO;
    using System.Linq;
    using System.ServiceModel.Channels;
    using System.ServiceModel.Description;
    using System.ServiceModel.Dispatcher;
    using System.Xml;

    /// <summary>
    /// This class which implements the IEndpointBehavior and IClientMessageInspector interface is used to add MS-WOPI request header and store all the requests and responses for both MS-FSSHTTP and MS-WOPI. 
    /// </summary>
    public class MessageInspector : IEndpointBehavior, IClientMessageInspector
    {
        /// <summary>
        /// Specify the shared context.
        /// </summary>
        private SharedContext context;

        /// <summary>
        /// Initializes a new instance of the MessageInspector class with the specified context.
        /// </summary>
        /// <param name="context">Specify the context which will give information for MS-WOPI headers.</param>
        public MessageInspector(SharedContext context)
        {
            this.context = context;
        }

        /// <summary>
        /// Implement the IEndpointBehavior interface to pass data at runtime to bindings to support custom behavior. In this implementation, this function will be kept empty.
        /// </summary>
        /// <param name="endpoint">The endpoint to modify</param>
        /// <param name="bindingParameters">The objects that binding elements require to support the behavior.</param>
        public void AddBindingParameters(ServiceEndpoint endpoint, System.ServiceModel.Channels.BindingParameterCollection bindingParameters)
        {
        }

        /// <summary>
        /// Implements a modification or extension of the client across an endpoint. In this implementation, this function will add a custom inspector to the client endpoint to add the MS-WOPI headers and do the schema validation.
        /// </summary>
        /// <param name="endpoint">The endpoint that is to be customized.</param>
        /// <param name="clientRuntime">The client runtime to be customized.</param>
        public void ApplyClientBehavior(ServiceEndpoint endpoint, ClientRuntime clientRuntime)
        {
            clientRuntime.MessageInspectors.Add(this);
        }

        /// <summary>
        /// Implements a modification or extension of the service across an endpoint. In this implementation, this function will be kept empty.
        /// </summary>
        /// <param name="endpoint">The endpoint that exposes the contract.</param>
        /// <param name="endpointDispatcher">The endpoint dispatcher to be modified or extended.</param>
        public void ApplyDispatchBehavior(ServiceEndpoint endpoint, EndpointDispatcher endpointDispatcher)
        {
        }

        /// <summary>
        /// Implement to confirm that the endpoint meets some intended criteria. In this implementation, this function will be kept empty.
        /// </summary>
        /// <param name="endpoint"> The endpoint to validate.</param>
        public void Validate(ServiceEndpoint endpoint)
        {
        }

        /// <summary>
        /// Implementation does schema validation for the MS-FSSHTTP defined element after a reply message is received. 
        /// </summary>
        /// <param name="reply">Specify the message to be transformed into types and handed back to the client application.</param>
        /// <param name="correlationState">Specify the correlation state data.</param>
        public void AfterReceiveReply(ref System.ServiceModel.Channels.Message reply, object correlationState)
        {
        }

        /// <summary>
        /// Implementation will add the headers for MS-WOPI before a request message is sent to a service.
        /// </summary>
        /// <param name="request">Specify the message to be sent to the service.</param>
        /// <param name="channel">Specify the client object channel.</param>
        /// <returns>Return null.</returns>
        public object BeforeSendRequest(ref System.ServiceModel.Channels.Message request, System.ServiceModel.IClientChannel channel)
        {
            // Restore the request.
            string requestString = request.ToString();
            XmlDocument requestXml = new XmlDocument();
            requestXml.LoadXml(requestString);
            SchemaValidation.LastRawRequestXml = requestXml.DocumentElement;

            // Create a clone message for the request.
            MessageBuffer messageBuffer = request.CreateBufferedCopy(int.MaxValue);
            request = messageBuffer.CreateMessage();

            if (this.context.OperationType == OperationType.WOPICellStorageRequest
                || this.context.OperationType == OperationType.WOPICellStorageRelativeRequest)
            {
                // Remove the SoapAction header
                request.Headers.RemoveAll("Action", "http://schemas.microsoft.com/ws/2005/05/addressing/none");

                // Add all the MUST header values
                this.AddHeader(ref request, "X-WOPI-Proof", this.context.XWOPIProof);
                this.AddHeader(ref request, "X-WOPI-TimeStamp", this.context.XWOPITimeStamp);
                this.AddHeader(ref request, "Authorization", this.context.XWOPIAuthorization);

                // Add all the optional value
                if (this.context.IsXWOPIOverrideSpecified)
                {
                    this.AddHeader(ref request, "X-WOPI-Override", this.context.GetValueOrDefault<string>("X-WOPI-Override", "COBALT"));
                }

                if (this.context.IsXWOPISizeSpecified)
                {
                    this.AddHeader(ref request, "X-WOPI-Size", this.context.GetValueOrDefault<int>("X-WOPI-Size", this.SizeOfWOPIMessage(messageBuffer.CreateMessage())).ToString());
                }

                if (this.context.OperationType == OperationType.WOPICellStorageRelativeRequest)
                {
                    if (this.context.IsXWOPIRelativeTargetSpecified)
                    {
                        this.AddHeader(ref request, "X-WOPI-RelativeTarget", this.context.XWOPIRelativeTarget);
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// This method is used to calculate the message size.
        /// </summary>
        /// <param name="message">Specify the request message.</param>
        /// <returns>Return the size of the request message.</returns>
        private int SizeOfWOPIMessage(Message message)
        {
            // Create an WOPI encoder
            WOPIMessageEncodingBindingElement element = new WOPIMessageEncodingBindingElement();
            element.MessageVersion = message.Version;
            element.Encoding = this.context.GetValueOrDefault<string>("Encoding", "utf-8");
            WOPIMessageEncoderFactory factory = (WOPIMessageEncoderFactory)element.CreateMessageEncoderFactory();
            MessageEncoder encoder = factory.Encoder;

            // Write the message and return its length
            int size;
            using (MemoryStream stream = new MemoryStream())
            {
                encoder.WriteMessage(message, stream);
                size = (int)stream.Length;
            }

            return size;
        }

        /// <summary>
        /// This method is used to add headers to the MS-WOPI requests.
        /// </summary>
        /// <param name="request">Specify the request message which will be modified.</param>
        /// <param name="headerKey">Specify the http header key.</param>
        /// <param name="value">Specify the http header value.</param>
        private void AddHeader(ref System.ServiceModel.Channels.Message request, string headerKey, string value)
        {
            HttpRequestMessageProperty property;
            object propertyObject;

            if (request.Properties.TryGetValue(HttpRequestMessageProperty.Name, out propertyObject))
            {
                property = propertyObject as HttpRequestMessageProperty;
                if (!property.Headers.AllKeys.Contains(headerKey))
                {
                    property.Headers.Add(headerKey, value);
                }
            }
            else
            {
                HttpRequestMessageProperty messageProperty = new HttpRequestMessageProperty();
                request.Properties.Add(HttpRequestMessageProperty.Name, messageProperty);
                if (!messageProperty.Headers.AllKeys.Contains(headerKey))
                {
                    messageProperty.Headers.Add(headerKey, value);
                }
            }
        }
    }
}