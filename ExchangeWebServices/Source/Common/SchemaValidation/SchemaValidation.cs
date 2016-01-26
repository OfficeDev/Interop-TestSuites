namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Reflection;
    using System.Text;
    using System.Web.Services.Protocols;
    using System.Xml;
    using System.Xml.Linq;
    using System.Xml.Schema;
    using System.Xml.Serialization;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Partial class of proxy class.
    /// </summary>
    public partial class ExchangeServiceBinding
    {
        /// <summary>
        /// The paths of xsd files.
        /// </summary>
        private static string[] xsdPaths = new string[] 
        { 
            "MS-OXWSDLGM-message.xsd",
            "MS-OXWSDLGM-types.xsd", 
            "MS-OXWSATT-messages.xsd",
            "MS-OXWSATT-types.xsd", 
            "MS-OXWSBTRF-messages.xsd",
            "MS-OXWSBTRF-types.xsd",
            "MS-OXWSCDATA-messages.xsd",
            "MS-OXWSCDATA-types.xsd",
            "MS-OXWSCEXT-messages.xsd",
            "MS-OXWSCEXT-types.xsd",
            "MS-OXWSCONT-types.xsd",
            "MS-OXWSCONV-messages.xsd",
            "MS-OXWSCONV-types.xsd",
            "MS-OXWSCORE-messages.xsd",
            "MS-OXWSCORE-types.xsd",
            "MS-OXWSCVTID-messages.xsd",
            "MS-OXWSCVTID-types.xsd",
            "MS-OXWSDLIST-messages.xsd",
            "MS-OXWSDLIST-types.xsd",
            "MS-OXWSEDISC-messages.xsd",
            "MS-OXWSEDISC-types.xsd",
            "MS-OXWSFOLD-messages.xsd",
            "MS-OXWSFOLD-types.xsd",
            "MS-OXWSGNI-messages.xsd", 
            "MS-OXWSGNI-types.xsd", 
            "MS-OXWSGTRM-messages.xsd",
            "MS-OXWSGTRM-types.xsd", 
            "MS-OXWSGTZ-messages.xsd",
            "MS-OXWSGTZ-types.xsd", 
            "MS-OXWSMSG-types.xsd", 
            "MS-OXWSMSHR-messages.xsd",
            "MS-OXWSMSHR-types.xsd",
            "MS-OXWSMTGS-types.xsd",
            "MS-OXWSNTIF-messages.xsd", 
            "MS-OXWSNTIF-types.xsd", 
            "MS-OXWSPED-messages.xsd",
            "MS-OXWSPERS-messages.xsd", 
            "MS-OXWSPERS-types.xsd", 
            "MS-OXWSPOST-types.xsd",
            "MS-OXWSPSNTIF-messages.xsd",
            "MS-OXWSPSNTIF-types.xsd",
            "MS-OXWSRSLNM-messages.xsd",
            "MS-OXWSRSLNM-types.xsd",
            "MS-OXWSSRCH-messages.xsd", 
            "MS-OXWSSRCH-types.xsd",
            "MS-OXWSSYNC-messages.xsd",
            "MS-OXWSSYNC-types.xsd",
            "MS-OXWSTASK-types.xsd",
            "MS-OXWSURPT-messages.xsd",
            "MS-OXWSURPT-types.xsd",
            "MS-OXWSXPROP-types.xsd",
            "MS-OXWSUSRCFG-messages.xsd",
            "MS-OXWSUSRCFG-types.xsd"
        };

        /// <summary>
        /// ITestSite object to use ITestSite's functions.
        /// </summary>
        private ITestSite site;

        /// <summary>
        /// An XmlReaderSettings that is used to validate the schema.
        /// </summary>
        private XmlReaderSettings xmlReaderSettings;

        /// <summary>
        /// Get Uri property of SoapHttpClientProtocol.
        /// </summary>
        private PropertyInfo baseURI = typeof(ExchangeServiceBinding).GetProperty("Uri", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.InvokeMethod);

        /// <summary>
        /// Get PendingSyncRequest property of SoapHttpClientProtocol.
        /// </summary>
        private PropertyInfo basePendingSyncRequest = typeof(ExchangeServiceBinding).GetProperty("PendingSyncRequest", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.InvokeMethod);

        /// <summary>
        /// Get SetStream method of SoapClientMessage.
        /// </summary>
        private MethodInfo setStream = typeof(SoapClientMessage).GetMethod("SetStream", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.InvokeMethod);

        /// <summary>
        /// Represents the callback method that will handle XML schema validation events and the System.Xml.Schema.ValidationEventArgs.
        /// </summary>
        private ValidationEventHandler itemStackValidationEventHandler;

        /// <summary>
        /// An xmlWriterHook that is used to hook actual request xml in the request stream.
        /// </summary>
        private XmlWriterInjector xmlWriterHookInstance;

        /// <summary>
        /// The raw XML request to server
        /// </summary>
        private IXPathNavigable rawRequestXml;

        /// <summary>
        /// The raw XML response from server
        /// </summary>
        private IXPathNavigable rawResponseXml;

        /// <summary>
        /// The warning collection of schema validation.
        /// </summary>
        private Collection<ValidationEventArgs> schemaValidationWarnings = new Collection<ValidationEventArgs>();

        /// <summary>
        /// The error collection of schema validation.
        /// </summary>
        private Collection<ValidationEventArgs> schemaValidationErrors = new Collection<ValidationEventArgs>();

        /// <summary>
        /// Initializes a new instance of the ExchangeServiceBinding class with the specified parameter.
        /// </summary>
        /// <param name="url">The base URL of the XML Web service the client is requesting.</param>
        /// <param name="userName">The user name associated with the credentials.</param>
        /// <param name="password">The password for the user name associated with the credentials.</param>
        /// <param name="domain">The domain associated with these credentials.</param>
        /// <param name="testSite">The test site instance.</param>
        public ExchangeServiceBinding(string url, string userName, string password, string domain, ITestSite testSite)
        {
            this.Url = url;
            this.Credentials = new NetworkCredential(userName, password, domain);
            this.itemStackValidationEventHandler = this.ValidationCallBack;
            this.InitialXmlReaderSettings();
            if (this.site == null)
            {
                this.site = testSite;
            }
        }

        /// <summary>
        /// Handle the server response.
        /// </summary>
        /// <param name="request">The request messages.</param>
        /// <param name="response">The response messages.</param>
        /// <param name="isSchemaValidated">The schema validation result.</param>
        public delegate void ServiceResponseDelegate(BaseRequestType request, BaseResponseMessageType response, bool isSchemaValidated);

        /// <summary>
        /// Handle the server response. Invoked when a response received from server.
        /// </summary>
        public static event ServiceResponseDelegate ServiceResponseEvent;

        /// <summary>
        /// Gets the warning collection of schema validation.
        /// </summary>
        public Collection<ValidationEventArgs> SchemaValidationWarnings
        {
            get
            {
                return this.schemaValidationWarnings;
            }
        }

        /// <summary>
        /// Gets the error collection of schema validation.
        /// </summary>
        public Collection<ValidationEventArgs> SchemaValidationErrors
        {
            get
            {
                return this.schemaValidationErrors;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the schema validation is successful.
        /// </summary>
        public bool IsSchemaValidated
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the raw XML request sent to protocol SUT
        /// </summary>
        public IXPathNavigable LastRawRequestXml
        {
            get
            {
                return this.rawRequestXml;
            }
        }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT
        /// </summary>
        public IXPathNavigable LastRawResponseXml
        {
            get
            {
                return this.rawResponseXml;
            }
        }

        /// <summary>
        /// Overload .NET framework Invoke method to provide extra XML schema validation function.
        /// </summary>
        /// <param name="methodName">The name of the XML Web service method.</param>
        /// <param name="parameters">An array of objects that contains the parameters to pass to the XML Web service. The order of the values in the array corresponds to the order of the parameters in the calling method of the derived class.</param>
        /// <returns>An array of objects that contains the return value and any reference or out parameters of the derived class method.</returns>
        public new object[] Invoke(string methodName, object[] parameters)
        {
            this.rawRequestXml = null;
            this.rawResponseXml = null;
            WebResponse webResponse = null;
            WebRequest webRequest = null;
            object[] objArray;
            StringBuilder sb = new StringBuilder();
            using (StringWriter sw = new StringWriter(sb))
            {
                Assembly assembly = Assembly.GetAssembly(typeof(ItemType));
                BaseRequestType requestInfo = (BaseRequestType)parameters[0];
                Type type = typeof(BaseRequestType);
                Type[] types = assembly.GetTypes();
                foreach (Type t in types)
                {
                    if (requestInfo.GetType().Equals(t))
                    {
                        type = t;
                    }
                }

                // Serialize the request.
                XmlSerializer xs = new XmlSerializer(type);
                xs.Serialize(sw, requestInfo);
            }

            try
            {
                // Creates a WebRequest for the specified uri.
                webRequest = this.GetWebRequest(this.baseURI.GetValue(this, null) as Uri);
                webRequest.PreAuthenticate = true;

                // Check PropertyInfo is null or not.
                Trace.Assert(this.basePendingSyncRequest != null, "PropertyInfo can not be NULL");

                // Sets the value of PendingSyncRequest.
                this.basePendingSyncRequest.SetValue(this, webRequest, null);

                // Invoke method of HttpWebClientProtocol.
                SoapClientMessage message = this.InstanceInvokeBase("BeforeSerialize", webRequest, methodName, parameters) as SoapClientMessage;
                Stream requestStream = webRequest.GetRequestStream();
                try
                {
                    Trace.Assert(this.setStream != null, "MethodInfo can not be NULL");
                    this.setStream.Invoke(message, new object[] { requestStream });
                    this.InstanceInvokeBase("Serialize", message);
                }
                finally
                {
                    requestStream.Close();
                }

                // Get the actual request xml by using xmlWriterHookInstance. The xmlWriterHookInstance is appended to the GetWriterForMessage method of the proxy class.
                string requestXmlString = this.xmlWriterHookInstance.Xml.ToString();

                // Load the actual request xml to an XmlElement
                if (!string.IsNullOrEmpty(requestXmlString))
                {
                    XmlDocument xmlDocOfReadRequest = new XmlDocument();
                    xmlDocOfReadRequest.LoadXml(requestXmlString);
                    this.rawRequestXml = xmlDocOfReadRequest.DocumentElement;
                    this.site.Log.Add(LogEntryKind.Debug, "The raw xml request message is:\r\n{0}", ((XmlElement)this.rawRequestXml).OuterXml);
                }

                webResponse = this.GetWebResponse(webRequest);
                HttpWebResponse httpWebResponse = (HttpWebResponse)webResponse;
                if (httpWebResponse.StatusCode != HttpStatusCode.OK)
                {
                    throw new WebException(httpWebResponse.StatusDescription);
                }

                Stream responseStream = null;
                try
                {
                    responseStream = webResponse.GetResponseStream();
                    string streamString = string.Empty;
                    using (StreamReader sr = new StreamReader(responseStream))
                    {
                        responseStream = null;

                        StringBuilder xmlString = new StringBuilder();
                        while (sr.Peek() > -1)
                        {
                            string strInput = sr.ReadLine();
                            xmlString.Append(strInput);
                        }

                        streamString = xmlString.ToString();
                        Trace.TraceInformation(streamString);
                    }

                    using (Stream streamObjRawXmlResponse = new MemoryStream(ASCIIEncoding.Default.GetBytes(streamString)))
                    {
                        XmlDocument responseXml = new XmlDocument();
                        responseXml.LoadXml(streamString);
                        this.rawResponseXml = responseXml.DocumentElement;
                        this.site.Log.Add(LogEntryKind.Debug, "The raw xml response message is:\r\n{0}", ((XmlElement)this.rawResponseXml).OuterXml);

                        objArray = this.InstanceInvoke("ReadResponse", message, webResponse, streamObjRawXmlResponse, false) as object[];

                        // Gets SOAP header from the response.
                        string soapHeader = this.GetSoapElement(responseXml, "Header");

                        // Gets SOAP body from the response.
                        string soapBody = this.GetSoapElement(responseXml, "Body");

                        this.XmlValidater(soapHeader, soapBody);

                        if (ExchangeServiceBinding.ServiceResponseEvent != null)
                        {
                            if (objArray[0] is BaseResponseMessageType)
                            {
                                ExchangeServiceBinding.ServiceResponseEvent(
                                    (BaseRequestType)parameters[0],
                                    (BaseResponseMessageType)objArray[0],
                                    this.IsSchemaValidated);
                            }
                        }
                    }
                }
                catch (XmlException exception)
                {
                    if (exception.Message.Contains("The following elements are not closed"))
                    {
                        throw new InvalidOperationException("The xml is not complete.", exception);
                    }
                    else
                    {
                        throw new InvalidOperationException("WebResponseBadXml", exception);
                    }
                }
                finally
                {
                    if (responseStream != null)
                    {
                        responseStream.Dispose();
                    }
                }
            }
            finally
            {
                if (webRequest == this.basePendingSyncRequest.GetValue(this, null) as WebRequest)
                {
                    if (this.basePendingSyncRequest.CanWrite)
                    {
                        this.basePendingSyncRequest.SetValue(this, null, null);
                    }
                }
            }

            return objArray;
        }

        #region Override implementation, append the xmlWriterHookInstance into the GetWriterForMessage method
        /// <summary>
        /// Override implementation, append the xmlWriterHookInstance into the GetWriterForMessage method
        /// </summary>
        /// <param name="message">The SOAP client message</param>
        /// <param name="bufferSize">The size of the buffer for the xml writer</param>
        /// <returns>An instance of the XmlWriterHook</returns>
        protected override XmlWriter GetWriterForMessage(SoapClientMessage message, int bufferSize)
        {
            XmlWriter originalXmlWriterImplementation = base.GetWriterForMessage(message, bufferSize);
            this.xmlWriterHookInstance = new XmlWriterInjector(originalXmlWriterImplementation);
            return this.xmlWriterHookInstance;
        }

        #endregion

        /// <summary>
        /// Initialize an instance of XmlReaderSettings.
        /// </summary>
        private void InitialXmlReaderSettings()
        {
            this.xmlReaderSettings = new XmlReaderSettings();
            foreach (string schemaFilePath in xsdPaths)
            {
                string xsd = this.ReadSchema(schemaFilePath);
                string targetNamespace = this.GetTargetNamespace(xsd);

                StringReader reader = null;
                try
                {
                    reader = new StringReader(xsd);
                    using (XmlReader schemaReader = XmlReader.Create(reader))
                    {
                        reader = null;
                        this.xmlReaderSettings.Schemas.Add(targetNamespace, schemaReader);
                    }
                }
                finally
                {
                    if (reader != null)
                    {
                        reader.Dispose();
                    }
                }
            }

            this.xmlReaderSettings.ValidationType = ValidationType.Schema;
            this.xmlReaderSettings.ConformanceLevel = ConformanceLevel.Document;
            this.xmlReaderSettings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
            this.xmlReaderSettings.ValidationEventHandler += new ValidationEventHandler(this.itemStackValidationEventHandler);
        }

        /// <summary>
        /// Gets the SOAP header or SOAP body from response.
        /// </summary>
        /// <param name="responseXml">The response xml</param>
        /// <param name="element">The element's name, Header or Body.</param>
        /// <returns>The SOAP element.</returns>
        private string GetSoapElement(XmlDocument responseXml, string element)
        {
            XmlNode nodesForSoapElement;
            XNamespace soapNamespace = @"http://schemas.xmlsoap.org/soap/envelope/";
            XName elementName = soapNamespace + element;
            XDocument doc;

            using (XmlNodeReader nodeReader = new XmlNodeReader(responseXml))
            {
                doc = XDocument.Load(nodeReader, LoadOptions.None);
            }

            // Select elements.
            IEnumerable<XElement> elements = from x in doc.Root.Elements(elementName)
                           select x;

            // If there is no SOAP header, return null.
            if (elements.Count() == 0 && element.Equals("Header"))
            {
                return null;
            }

            using (XmlReader reader = elements.First().CreateReader())
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(reader);
                nodesForSoapElement = xmlDoc;
            }

            // There is only one SOAP header or body element, return the first child's xml.
            return nodesForSoapElement.FirstChild.InnerXml;
        }

        /// <summary>
        /// Reads schema from an XSD file.
        /// </summary>
        /// <param name="xsdFile">An XSD file name.</param>
        /// <returns>The schema string used to validate server response.</returns>
        private string ReadSchema(string xsdFile)
        {
            using (FileStream stream = new FileStream(xsdFile, FileMode.Open, FileAccess.Read))
            {
                // Creates a new XDocument instance by using the specified stream.
                XDocument xsd = XDocument.Load(stream);
                XNamespace namespaceW3 = @"http://www.w3.org/2001/XMLSchema";
                XName includeElementName = namespaceW3 + "include";

                // Selects elements.
                IEnumerable<XElement> includeElements = from x in xsd.Root.Elements(includeElementName)
                                      select x;
                includeElements = new List<XElement>(includeElements);
                foreach (XElement includeElement in includeElements)
                {
                    // Removes this node from its parent.
                    includeElement.Remove();
                }

                XName importElementName = namespaceW3 + "import";
                XName schemaLocationName = "schemaLocation";

                // Selects attributes.
                IEnumerable<XAttribute> schemaLocationAttributes = from x in xsd.Root.Elements(importElementName)
                                               where x.Attribute(schemaLocationName) != null
                                               select x.Attribute(schemaLocationName);
                schemaLocationAttributes = new List<XAttribute>(schemaLocationAttributes);
                foreach (XAttribute schemaLocationAttribute in schemaLocationAttributes)
                {
                    // Removes this attribute from its parent element.
                    schemaLocationAttribute.Remove();
                }

                return xsd.ToString();
            }
        }

        /// <summary>
        /// Gets the target namespace.
        /// </summary>
        /// <param name="fullSchema">Full schema content.</param>
        /// <returns>The target namespace.</returns>
        private string GetTargetNamespace(string fullSchema)
        {
            XNamespace namespaceW3 = @"http://www.w3.org/2001/XMLSchema";
            XName schemaElementName = namespaceW3 + "schema";
            XDocument doc = XDocument.Parse(fullSchema);
            XElement schemaElement = doc.Element(schemaElementName);
            if (schemaElement != null)
            {
                XAttribute targetAttribute = schemaElement.Attribute("targetNamespace");
                if (targetAttribute != null)
                {
                    return targetAttribute.Value;
                }
            }

            return null;
        }

        /// <summary>
        /// Validate a piece of Xml fragment according the given Xml Schema.
        /// </summary>
        /// <param name="xmlHeader">The xml fragment of SOAP header.</param>
        /// <param name="xmlBody">The xml fragment of SOAP body.</param>
        /// <returns>A Boolean value of the validate result.</returns>
        private bool XmlValidater(string xmlHeader, string xmlBody)
        {
            this.IsSchemaValidated = false;
            if (xmlHeader == null)
            {
                return this.IsSchemaValidated = this.ElementValidater(xmlBody);
            }
            else
            {
                return this.IsSchemaValidated = this.ElementValidater(xmlHeader) && this.ElementValidater(xmlBody);
            }
        }

        /// <summary>
        /// Validate a piece of Xml fragment according the given Xml Schema.
        /// </summary>
        /// <param name="elementXml">The xml fragment.</param>
        /// <returns>A Boolean value of the validate result.</returns>
        private bool ElementValidater(string elementXml)
        {
            this.schemaValidationErrors.Clear();
            this.schemaValidationWarnings.Clear();
            bool isValidated = false;
            StringReader stringReader = null;
            try
            {
                stringReader = new StringReader(elementXml);
                using (XmlReader xmlReader = XmlReader.Create(stringReader, this.xmlReaderSettings))
                {
                    stringReader = null;
                    while (xmlReader.Read())
                    {
                    }

                    isValidated = this.schemaValidationErrors.Count == 0 && this.schemaValidationWarnings.Count == 0;
                }
            }
            finally
            {
                if (stringReader != null)
                {
                    stringReader.Dispose();
                }
            }

            return isValidated;
        }

        /// <summary>
        /// Display any warnings or errors.
        /// </summary>
        /// <param name="sender">The source of the event. Note Determine the type of a sender before using it in your code. You cannot assume that the sender is an instance of a particular type. The sender is also not guaranteed to not be null. Always surround your casts with failure handling logic.</param>
        /// <param name="args">The event data.</param>
        private void ValidationCallBack(object sender, ValidationEventArgs args)
        {
            if (args.Severity == XmlSeverityType.Warning)
            {
                this.schemaValidationWarnings.Add(args);
            }

            if (args.Severity == XmlSeverityType.Error)
            {
                this.schemaValidationErrors.Add(args);
            }
        }

        /// <summary>
        /// Invoke method of HttpWebClientProtocol.
        /// </summary>
        /// <param name="name">Invokes the method or constructor represented by the current instance, using the specified parameters.</param>
        /// <param name="parameters">An argument list for the invoked method or constructor.</param>
        /// <returns>An object containing the return value of the invoked method, or null in the case of a constructor.</returns>
        private object InstanceInvokeBase(string name, params object[] parameters)
        {
            try
            {
                BindingFlags bfs = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.InvokeMethod;
                MethodInfo mi = typeof(ExchangeServiceBinding).BaseType.GetMethod(name, bfs);
                Trace.Assert(mi != null, "MethodInfo can not be NULL");
                return mi.Invoke(this, parameters);
            }
            catch (TargetInvocationException ex)
            {
                if (ex.InnerException != null)
                {
                    throw ex.InnerException;
                }
                else
                {
                    throw;
                }
            }
        }

        /// <summary>
        /// Invoke method of HttpWebClientProtocol.
        /// </summary>
        /// <param name="name">Invokes the method or constructor represented by the current instance, using the specified parameters.</param>
        /// <param name="parameters">An argument list for the invoked method or constructor.</param>
        /// <returns>An object containing the return value of the invoked method, or null in the case of a constructor.</returns>
        private object InstanceInvoke(string name, params object[] parameters)
        {
            try
            {
                BindingFlags bfs = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.InvokeMethod;
                MethodInfo mi = typeof(SoapHttpClientProtocol).GetMethod(name, bfs);
                Trace.Assert(mi != null, "MethodInfo can not be NULL");
                return mi.Invoke(this, parameters);
            }
            catch (TargetInvocationException ex)
            {
                if (ex.InnerException != null)
                {
                    throw ex.InnerException;
                }
                else
                {
                    throw;
                }
            }
        }
    }
}