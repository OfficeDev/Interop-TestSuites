namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Web.Services.Protocols;
    using System.Xml;
    using System.Xml.Schema;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A Class represent the XML schema validation function.
    /// </summary>
    public static class SchemaValidation
    {
        /// <summary>
        /// The standard SOAP message namespace.
        /// </summary>
        public const string SoapNamespace = "http://www.w3.org/2003/05/soap-envelope";

        /// <summary>
        /// The result of schema validation.
        /// </summary>
        private static ValidationResult validationResult;

        /// <summary>
        /// ITestSite object, so that proxy can use ITestSite's functions.
        /// </summary>
        private static ITestSite site;

        /// <summary>
        /// A list of schema which is used to construct schema validation reader.
        /// </summary>
        private static List<string> schemaList = new List<string>();

        /// <summary>
        /// The raw XMl response from server.
        /// </summary>
        private static XmlElement rawResponseXml;

        /// <summary>
        /// The raw XMl request to server.
        /// </summary>
        private static XmlElement rawRequestXml;

        /// <summary>
        /// The warning results of XML schema validation.
        /// </summary>
        private static List<ValidationEventArgs> xmlValidationWarnings;

        /// <summary>
        /// The error results of XML schema validation.
        /// </summary>
        private static List<ValidationEventArgs> xmlValidationErrors;

        /// <summary>
        /// Gets the validationResult of schema.
        /// </summary>
        public static ValidationResult ValidationResult
        {
            get
            {
                return validationResult;
            }
        }

        /// <summary>
        /// Gets the error results of XML schema validation.
        /// </summary>
        public static List<ValidationEventArgs> XmlValidationErrors
        {
            get
            {
                return xmlValidationErrors;
            }
        }

        /// <summary>
        /// Gets the warnings results of XML schema validation.
        /// </summary>
        public static List<ValidationEventArgs> XmlValidationWarnings
        {
            get
            {
                return xmlValidationWarnings;
            }
        }

        /// <summary>
        /// Gets or sets the operate raw XMl response from protocol SUT.
        /// </summary>
        public static XmlElement LastRawResponseXml
        {
            get
            {
                return rawResponseXml;
            }

            set
            {
                rawResponseXml = value;
            }
        }

        /// <summary>
        /// Gets or sets the raw Request Xml send to protocol SUT.
        /// </summary>
        public static XmlElement LastRawRequestXml
        {
            get
            {
                return rawRequestXml;
            }

            set
            {
                rawRequestXml = value;
            }
        }

        /// <summary>
        /// Validate a piece of Xml fragment.
        /// </summary>
        /// <param name="testSite">Implements Microsoft.Protocols.TestTools.IAdapter.Site.</param>
        /// <param name="xmlValue">Xml fragment string.</param>
        /// <returns>A return represents the validation result of the specified Xml fragment string.</returns>
        public static ValidationResult ValidateXml(ITestSite testSite, string xmlValue)
        {
            site = testSite;
            InitSchemaList();
            InitValidateRecoder();
            ValidateSpecialXml(xmlValue);

            return validationResult;
        }

        /// <summary>
        /// This method is used to generate xml validation result information by using 
        /// the record errors and warnings.
        /// </summary>
        /// <returns>Return the xml validation result string which contains the errors and warning information.</returns>
        public static string GenerateValidationResult()
        {
            StringBuilder sb = new StringBuilder();

            if (XmlValidationWarnings.Count != 0)
            {
                sb.Append("The xml validation warnings:");
                sb.Append(Environment.NewLine);
                foreach (ValidationEventArgs warning in XmlValidationWarnings)
                {
                    sb.Append(warning.Message);
                    sb.Append(Environment.NewLine);
                }
            }

            if (XmlValidationErrors.Count != 0)
            {
                sb.Append("The xml validation errors:");
                foreach (ValidationEventArgs error in XmlValidationErrors)
                {
                    sb.Append(error.Message);
                    sb.Append(Environment.NewLine);
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// This method is used to get the soap fault detail.
        /// </summary>
        /// <param name="soapBody">Xml fragment string.</param>
        /// <returns>The xml string.</returns>
        public static string GetSoapFaultDetailBody(string soapBody)
        {
            string detailBody = null;
            string elementName = null;

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(soapBody);

            // Prepare XPath
            string soapVersion = Common.GetConfigurationPropertyValue("SoapVersion", site);

            if (string.Compare(soapVersion, "SOAP11", true) == 0)
            {
                elementName = "detail";
            }

            if (string.Compare(soapVersion, "SOAP12", true) == 0)
            {
                elementName = "Detail";
            }

            // [Note] According to SOAP1.1 and SOAP 1.2 definitions, SOAP1.1 defines subelement 'detail' in SOAP Fault element, SOAP1.2 defines subelement "Detail". But actually, whether the soap fault response message are formatted 
            //  with SOAP 1.1 or 1.2, the subelement name in O12, O14 and O15 response is still "detail".
            string path = string.Format("//{0}", elementName.ToLower());

            XmlNode detailNode = doc.SelectSingleNode(path);

            XmlElement detailElement = (XmlElement)detailNode;
            detailElement.SetAttribute("xmlns", detailNode.FirstChild.NamespaceURI);
            detailBody = detailElement.OuterXml;

            return detailBody;
        }

        /// <summary>
        /// Validate a piece of Xml fragment.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into this class, make this class can use ITestSite's function.</param>
        /// <param name="ignoreSoapFaultSchemaValidationForSoap12">Indicate that whether ignore schema validation for SOAP fault in SOAP1.2.</param>
        /// <returns>The result of validation.</returns>
        internal static ValidationResult ValidateXml(ITestSite testSite, bool ignoreSoapFaultSchemaValidationForSoap12)
        {
            XmlNodeList nodesForSoapFault = SchemaValidation.LastRawResponseXml.GetElementsByTagName("Fault", SoapNamespace);
            SoapProtocolVersion soapVersion = Common.GetConfigurationPropertyValue<SoapProtocolVersion>("SoapVersion", testSite);
            if (ignoreSoapFaultSchemaValidationForSoap12
                && (soapVersion == SoapProtocolVersion.Soap12) && (nodesForSoapFault.Count > 0))
            {
                site = testSite;
                validationResult = ValidationResult.Inconclusive;
                return validationResult;
            }
            else
            {
                return ValidateXml(testSite, rawResponseXml.OuterXml);
            }
        }

        /// <summary>
        /// Initialize the validate recorder.
        /// </summary>
        private static void InitValidateRecoder()
        {
            // Initialize the result as "Success", if there are any validation error, this value will be changed by call back method.
            validationResult = ValidationResult.Success;
            if (null == xmlValidationWarnings)
            {
                xmlValidationWarnings = new List<ValidationEventArgs>();
            }

            xmlValidationWarnings.Clear();

            if (null == xmlValidationErrors)
            {
                xmlValidationErrors = new List<ValidationEventArgs>();
            }

            xmlValidationErrors.Clear();
        }

        /// <summary>
        /// Validate special xml.
        /// </summary>
        /// <param name="xmlString">The xml string.</param>
        private static void ValidateSpecialXml(string xmlString)
        {
            XmlReaderSettings validationSettings = new XmlReaderSettings();
            foreach (string eachSchema in schemaList)
            {
                using (StringReader schemaReader = new StringReader(eachSchema))
                {
                    using (XmlReader xsdXmlReader = XmlReader.Create(schemaReader))
                    {
                        string targetNameSpace = GetTargetNamespace(eachSchema);
                        if (string.IsNullOrEmpty(targetNameSpace))
                        {
                            validationSettings.Schemas.Add(null, xsdXmlReader);
                        }
                        else
                        {
                            validationSettings.Schemas.Add(targetNameSpace, xsdXmlReader);
                        }
                    }
                }
            }

            validationSettings.ValidationType = ValidationType.Schema;
            validationSettings.ConformanceLevel = ConformanceLevel.Document;
            validationSettings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
            validationSettings.ValidationEventHandler += new ValidationEventHandler(ValidationCallBack);

            using (StringReader reader = new StringReader(xmlString))
            {
                using (XmlReader validationReader = XmlReader.Create(reader, validationSettings))
                {
                    while (validationReader.Read())
                    {
                    }
                }
            }
        }

        /// <summary>
        /// This method is a callback function which is used to record the schema validation errors and warnings.
        /// </summary>
        /// <param name="sender">An object instance which represents the sender of the events.</param>
        /// <param name="args">The ValidationEventArgs contains validation result.</param>
        private static void ValidationCallBack(object sender, ValidationEventArgs args)
        {
            if (args.Severity == XmlSeverityType.Warning)
            {
                validationResult = ValidationResult.Warning;
                xmlValidationWarnings.Add(args);
            }
            else
            {
                validationResult = ValidationResult.Error;
                xmlValidationErrors.Add(args);
            }
        }

        /// <summary>
        /// This method is used to get the target namespace from a schema string.
        /// </summary>
        /// <param name="fullSchema">The schema string</param>
        /// <returns>A return value represents the target namespace value</returns>
        private static string GetTargetNamespace(string fullSchema)
        {
            XmlDocument xd = new XmlDocument();
            xd.LoadXml(fullSchema);
            XmlAttributeCollection attributeList = xd.DocumentElement.Attributes;
            foreach (XmlAttribute attribute in attributeList)
            {
                if (attribute.Name.Equals("targetNamespace", StringComparison.OrdinalIgnoreCase))
                {
                    return attribute.Value;
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// This method is used to initialize the schema definitions list which is used to set the schema validation settings.
        /// </summary>
        private static void InitSchemaList()
        {
            schemaList.Clear();
            schemaList.AddRange(GetSoapStandardSchema());

            string schemaDefinitionFileName = string.Format("{0}.wsdl", site.DefaultProtocolDocShortName);
            schemaList.AddRange(GetSchemaStringFromWsdlFile(schemaDefinitionFileName));
        }

        /// <summary> 
        /// A method used to get the standard soap schema definitions from xsd file. 
        /// </summary> 
        /// <returns>A return value represents the soap schema definitions.</returns> 
        private static string[] GetSoapStandardSchema()
        {
            List<string> soapSchema = new List<string>();
            XmlDocument xsdFile = new XmlDocument();
            xsdFile.Load("SOAP11.xsd");
            soapSchema.Add(xsdFile.OuterXml);
            xsdFile.Load("SOAP12.xsd");
            soapSchema.Add(xsdFile.OuterXml);

            return soapSchema.ToArray();
        }

        /// <summary>
        /// A method used to get the schema definitions from full WSDL file. It is designed to read schema definition from full WSDL file.
        /// </summary>
        /// <param name="schemaLoadPath">A parameter represents a full WSDL file path where the schema definitions should be loaded.</param>
        /// <returns>A return value represents the schema definitions.</returns>
        private static string[] GetSchemaStringFromWsdlFile(string schemaLoadPath)
        {
            List<string> schemas = new List<string>();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(schemaLoadPath);
            XmlNodeList wsdlTypes = xmlDoc.GetElementsByTagName("wsdl:types");
            XmlNodeList wsdlDefinitions = xmlDoc.GetElementsByTagName("wsdl:definitions");

            if (wsdlTypes != null && wsdlTypes.Count == 1)
            {
                foreach (XmlNode schemaElement in wsdlTypes[0].ChildNodes)
                {
                    if (schemaElement.NodeType != XmlNodeType.Comment)
                    {
                        foreach (XmlAttribute attribute in wsdlDefinitions[0].Attributes)
                        {
                            if (attribute.Prefix == "xmlns" && schemaElement.Attributes.GetNamedItem(attribute.Name) == null)
                            {
                                XmlAttribute namespaceAttribute = (XmlAttribute)attribute.Clone();
                                schemaElement.Attributes.Append(namespaceAttribute);
                            }
                        }

                        schemas.Add(schemaElement.OuterXml);
                    }
                }
            }

            return schemas.ToArray();
        }
    }
}