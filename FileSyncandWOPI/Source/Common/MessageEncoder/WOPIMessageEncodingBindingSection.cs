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
    using System.Configuration;
    using System.ServiceModel.Channels;
    using System.ServiceModel.Configuration;
    using System.Xml;

    /// <summary>
    /// This class is used to enable the usage of a custom WOPIMessageEncodingBindingElement implementation from a machine or application configuration file.
    /// </summary>
    public class WOPIMessageEncodingBindingSection : BindingElementExtensionElement
    {
        /// <summary>
        /// Gets or sets the configurable encoding property.
        /// </summary>
        [ConfigurationProperty("encoding", DefaultValue = "utf-8")]
        public string Encoding
        {
            get
            {
                return (string)base["encoding"];
            }

            set
            {
                base["encoding"] = value;
            }
        }

        /// <summary>
        /// Gets the configurable reader quotas property.
        /// </summary>
        [ConfigurationProperty("readerQuotas")]
        public XmlDictionaryReaderQuotasElement ReaderQuotasElement
        {
            get
            {
                return (XmlDictionaryReaderQuotasElement)base["readerQuotas"];
            }
        }

        /// <summary>
        /// Gets the binding element type.
        /// </summary>
        public override Type BindingElementType
        {
            get
            {
                return typeof(WOPIMessageEncodingBindingElement);
            }
        }

        /// <summary>
        /// This method is used to apply the configuration to WOPIMessageEncodingBindingElement.
        /// </summary>
        /// <param name="bindingElement">Specify the WOPIMessageEncodingBindingElement type instance.</param>
        public override void ApplyConfiguration(BindingElement bindingElement)
        {
            base.ApplyConfiguration(bindingElement);
            WOPIMessageEncodingBindingElement be = (WOPIMessageEncodingBindingElement)bindingElement;
            
            // Take the soap11 as the default message version.
            be.MessageVersion = MessageVersion.CreateVersion(System.ServiceModel.EnvelopeVersion.Soap11, AddressingVersion.None);
            be.Encoding = this.Encoding;
            this.ApplyReaderQuotas(be.ReaderQuotas);
        }

        /// <summary>
        /// This method is used to create the WOPIMessageEncodingBindingElement instance.
        /// </summary>
        /// <returns>Return the created WOPIMessageEncodingBindingElement instance.</returns>
        protected override System.ServiceModel.Channels.BindingElement CreateBindingElement()
        {
            WOPIMessageEncodingBindingElement bindingElement = new WOPIMessageEncodingBindingElement();
            this.ApplyConfiguration(bindingElement);
            return bindingElement;
        }

        /// <summary>
        /// This method is used to apply reader quotas configuration.
        /// </summary>
        /// <param name="readerQuotas">Specify the XmlDictionaryReaderQuotas type instance.</param>
        private void ApplyReaderQuotas(XmlDictionaryReaderQuotas readerQuotas)
        {
            if (readerQuotas == null)
            {
                throw new ArgumentNullException("readerQuotas");
            }

            if (this.ReaderQuotasElement.MaxDepth != 0)
            {
                readerQuotas.MaxDepth = this.ReaderQuotasElement.MaxDepth;
            }

            if (this.ReaderQuotasElement.MaxStringContentLength != 0)
            {
                readerQuotas.MaxStringContentLength = this.ReaderQuotasElement.MaxStringContentLength;
            }

            if (this.ReaderQuotasElement.MaxArrayLength != 0)
            {
                readerQuotas.MaxArrayLength = this.ReaderQuotasElement.MaxArrayLength;
            }

            if (this.ReaderQuotasElement.MaxBytesPerRead != 0)
            {
                readerQuotas.MaxBytesPerRead = this.ReaderQuotasElement.MaxBytesPerRead;
            }

            if (this.ReaderQuotasElement.MaxNameTableCharCount != 0)
            {
                readerQuotas.MaxNameTableCharCount = this.ReaderQuotasElement.MaxNameTableCharCount;
            }
        }
    }
}