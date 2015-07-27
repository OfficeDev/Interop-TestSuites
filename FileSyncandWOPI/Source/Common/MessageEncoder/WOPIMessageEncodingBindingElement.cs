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
    using System.ServiceModel.Channels;
    using System.Xml;

    /// <summary>
    /// This class represents the MS-WOPI binding element that specifies the message version used to encode messages.
    /// </summary>
    public class WOPIMessageEncodingBindingElement : MessageEncodingBindingElement
    {
        /// <summary>
        /// Specify the message version.
        /// </summary>
        private MessageVersion messageVersion;

        /// <summary>
        /// Initializes a new instance of the WOPIMessageEncodingBindingElement class.
        /// </summary>
        public WOPIMessageEncodingBindingElement()
        {
            this.ReaderQuotas = new XmlDictionaryReaderQuotas();
        }

        /// <summary>
        /// Initializes a new instance of the WOPIMessageEncodingBindingElement class. This is the copy constructor.
        /// </summary>
        /// <param name="binding">Specify the copy binding element.</param>
        private WOPIMessageEncodingBindingElement(WOPIMessageEncodingBindingElement binding)
            : this()
        {
            this.MessageVersion = binding.MessageVersion;
            this.Encoding = binding.Encoding;
            binding.ReaderQuotas.CopyTo(this.ReaderQuotas);
        }

        /// <summary>
        /// Gets or sets the encoding type.
        /// </summary>
        public string Encoding { get; set; }

        /// <summary>
        /// Gets the reader quotas.
        /// </summary>
        public XmlDictionaryReaderQuotas ReaderQuotas { get; private set; }

        /// <summary>
        /// Gets the message version.
        /// </summary>
        public override MessageVersion MessageVersion
        {
            get
            {
                return this.messageVersion;
            }

            set
            {
                this.messageVersion = value;
            }
        }

        /// <summary>
        /// This method is used to create the WOPIMessageEncoderFactory.
        /// </summary>
        /// <returns>Return the WOPIMessageEncoderFactory instance.</returns>
        public override MessageEncoderFactory CreateMessageEncoderFactory()
        {
            return new WOPIMessageEncoderFactory(this.messageVersion, this.Encoding);
        }

        /// <summary>
        /// This method is used to clone the WOPIMessageEncodingBindingElement.
        /// </summary>
        /// <returns>Return the clone of the WOPIMessageEncodingBindingElement.</returns>
        public override BindingElement Clone()
        {
            return new WOPIMessageEncodingBindingElement(this);
        }

        /// <summary>
        /// Override the BuildChannelFactory to add the WOPIMessageEncodingBindingElement instance to the BindingContext.
        /// </summary>
        /// <typeparam name="TChannel">Specify the type of channel the factory builds.</typeparam>
        /// <param name="context">Specify the binding context.</param>
        /// <returns>Return The System.ServiceModel.Channels.IChannelFactory of type TChannel initialized from the context.</returns>
        public override IChannelFactory<TChannel> BuildChannelFactory<TChannel>(BindingContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            context.BindingParameters.Add(this);
            return context.BuildInnerChannelFactory<TChannel>();
        }

        /// <summary>
        /// Override the BuildChannelListener to add the WOPIMessageEncodingBindingElement instance to the BindingContext.
        /// </summary>
        /// <typeparam name="TChannel">Specify the type of the listener that is built to accept.</typeparam>
        /// <param name="context">Specify the binding context.</param>
        /// <returns>Return the System.ServiceModel.Channels.IChannelListener of type System.ServiceModel.Channels.IChannel initialized from the context.</returns>
        public override IChannelListener<TChannel> BuildChannelListener<TChannel>(BindingContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            context.BindingParameters.Add(this);
            return context.BuildInnerChannelListener<TChannel>();
        }

        /// <summary>
        /// Override the CanBuildChannelFactory to add the WOPIMessageEncodingBindingElement instance to the BindingContext.
        /// </summary>
        /// <typeparam name="TChannel">Specify the type of channel the factory builds.</typeparam>
        /// <param name="context">Specify the binding context.</param>
        /// <returns>Return true if the System.ServiceModel.Channels.IChannelFactory of type TChannel can be built by the binding element; otherwise, false.</returns>
        public override bool CanBuildChannelFactory<TChannel>(BindingContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            context.BindingParameters.Add(this);
            return context.CanBuildInnerChannelFactory<TChannel>();
        }

        /// <summary>
        /// Override the CanBuildChannelListener to add the WOPIMessageEncodingBindingElement instance to the BindingContext.
        /// </summary>
        /// <typeparam name="TChannel">Specify the type of the listener that is built to accept.</typeparam>
        /// <param name="context">Specify the binding context.</param>
        /// <returns>Return true if the System.ServiceModel.Channels.IChannelListener of type System.ServiceModel.Channels.IChannel can be built by the binding element; otherwise, false.</returns>
        public override bool CanBuildChannelListener<TChannel>(BindingContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            context.BindingParameters.Add(this);
            return context.CanBuildInnerChannelListener<TChannel>();
        }

        /// <summary>
        /// Override this method to deal with the XmlDictionaryReaderQuotas type properties.
        /// </summary>
        /// <typeparam name="T">The typed object for which the method is querying.</typeparam>
        /// <param name="context">Specify the binding context.</param>
        /// <returns>Returns the typed object T requested if it is present or null if it is not present.</returns>
        public override T GetProperty<T>(BindingContext context)
        {
            if (typeof(T) == typeof(XmlDictionaryReaderQuotas))
            {
                return (T)(object)this.ReaderQuotas;
            }
            else
            {
                return base.GetProperty<T>(context);
            }
        }
    }
}