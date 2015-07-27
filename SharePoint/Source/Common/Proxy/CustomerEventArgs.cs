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
    using System.Web.Services.Protocols;
    using System.Xml;
    
    /// <summary>
    /// The EventArgs that contains the even data.
    /// </summary>
    public class CustomerEventArgs : EventArgs
    {
        /// <summary>
        /// The response message from server.
        /// </summary>
        private SoapClientMessage responseMessage;

        /// <summary>
        /// The request message that will be sent.
        /// </summary>
        private XmlWriterInjector requestMessage;

        /// <summary>
        /// The XmlReader that contains response xml.
        /// </summary>
        private XmlReader validationXmlReaderOut;

        /// <summary>
        /// The Raw request string in xml format.
        /// </summary>
        private string rawRequestXml;

        /// <summary>
        /// The Raw request string in xml format.
        /// </summary>
        private string rawResponseXml;

        /// <summary>
        /// Gets or sets the XmlReader that contains response xml.
        /// </summary>
        public XmlReader ValidationXmlReaderOut
        {
            get
            {
                return this.validationXmlReaderOut;
            }

            set
            {
                this.validationXmlReaderOut = value;
            }
        }

        /// <summary>
        /// Gets or sets the response message from server.
        /// </summary>
        public SoapClientMessage ResponseMessage
        {
            get
            {
                return this.responseMessage;
            }

            set
            {
                this.responseMessage = value;
                this.rawResponseXml = null;
            }
        }

        /// <summary>
        /// Gets or sets the request message that will be sent.
        /// </summary>
        public XmlWriterInjector RequestMessage
        {
            get
            {
                return this.requestMessage;
            }

            set
            {
                this.requestMessage = value;
                this.rawRequestXml = null;
            }
        }

        /// <summary>
        /// Gets or sets the request string in XML format.
        /// </summary>
        public string RawRequestXml
        {
            get
            {
                if (this.rawRequestXml == null)
                {
                    this.rawRequestXml = this.requestMessage == null ? null : this.requestMessage.Xml;
                }

                return this.rawRequestXml;
            }

            set
            {
                this.rawRequestXml = value;
            }
        }

        /// <summary>
        /// Gets or sets the response string in XML format.
        /// </summary>
        public string RawResponseXml
        {
            get
            {
                if (this.rawResponseXml == null)
                {
                    if (this.responseMessage != null)
                    {
                        using (StreamReader sr = new StreamReader(this.responseMessage.Stream))
                        {
                            this.rawResponseXml = sr.ReadToEnd();
                        }
                    }
                }

                return this.rawResponseXml;
            }

            set
            {
                this.rawResponseXml = value;
            }
        }
    }
}