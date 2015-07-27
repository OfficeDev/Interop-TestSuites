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
    using System.ServiceModel;
    using System.ServiceModel.Description;
    using System.ServiceModel.Dispatcher;

    /// <summary>
    /// A class that is used to inject into the WCF response handle process to do schema validation. 
    /// </summary>
    public class ResponseSchemaValidationInspector : IEndpointBehavior, IClientMessageInspector
    {
        /// <summary>
        /// The last request message in XML format.
        /// </summary>
        private string lastRawRequestMessgae;

        /// <summary>
        /// The last response message in XML format.
        /// </summary>
        private string lastRawResponseMessage;

        /// <summary>
        /// The validation event.
        /// </summary>
        public event EventHandler<CustomerEventArgs> ValidationEvent;

        /// <summary>
        /// Pass the data at runtime to bindings to support custom behavior. The current implementation does nothing.
        /// </summary>
        /// <param name="endpoint">The endpoint to modify.</param>
        /// <param name="bindingParameters">The objects that binding elements require to support the behavior.</param>
        public void AddBindingParameters(ServiceEndpoint endpoint, System.ServiceModel.Channels.BindingParameterCollection bindingParameters)
        {
            // Leave empty, do nothing
        }

        /// <summary>
        /// Add a message inspector in the specified ClientRuntime to give the ability of inspecting the request and response message.
        /// </summary>
        /// <param name="endpoint">The endpoint that is to be customized.</param>
        /// <param name="clientRuntime">The client runtime to be added message inspector.</param>
        public void ApplyClientBehavior(ServiceEndpoint endpoint, System.ServiceModel.Dispatcher.ClientRuntime clientRuntime)
        {
            clientRuntime.MessageInspectors.Add(this);
        }

        /// <summary>
        /// Implements a modification or extension of the service across an endpoint. The current implementation does nothing.
        /// </summary>
        /// <param name="endpoint">The endpoint that is to be customized.</param>
        /// <param name="endpointDispatcher">The endpoint dispatcher to be modified or extended.</param>
        public void ApplyDispatchBehavior(ServiceEndpoint endpoint, System.ServiceModel.Dispatcher.EndpointDispatcher endpointDispatcher)
        {
            // Leave empty, do nothing
        }

        /// <summary>
        /// Implement to confirm that the endpoint meets some intended criteria. The current implementation does nothing.
        /// </summary>
        /// <param name="endpoint">The endpoint to validate.</param>
        public void Validate(ServiceEndpoint endpoint)
        {
            // Leave empty, do nothing
        }

        /// <summary>
        /// Inspect the response message and do the schema validation.
        /// </summary>
        /// <param name="reply">The message to be transformed into types and handed back to the client application.</param>
        /// <param name="correlationState">Correlation state data</param>
        public void AfterReceiveReply(ref System.ServiceModel.Channels.Message reply, object correlationState)
        {
            this.lastRawResponseMessage = reply.ToString();

            if (this.ValidationEvent != null)
            {
                CustomerEventArgs args = new CustomerEventArgs { RawRequestXml = this.lastRawRequestMessgae, RawResponseXml = this.lastRawResponseMessage };
                this.ValidationEvent(this, args);
            }
        }

        /// <summary>
        /// Inspect the request message and restore this message.
        /// </summary>
        /// <param name="request">The message to be sent to the service.</param>
        /// <param name="channel">The client object channel.</param>
        /// <returns>Return null to indicate that no correlation state is used.</returns>
        public object BeforeSendRequest(ref System.ServiceModel.Channels.Message request, IClientChannel channel)
        {
            this.lastRawRequestMessgae = request.ToString();
            return null;
        }
    }
}