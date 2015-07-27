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
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System.Xml;
    using System.Xml.Schema;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A Class represent the XML schema validation function.
    /// </summary>
    public static class SchemaValidation
    {
        /// <summary>
        /// The result of schema validation.
        /// </summary>
        private static ValidationResult validationResult;

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
        /// Gets or sets the operated raw XMl response from protocol SUT.
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
        /// <param name="xmlValue">Xml fragment string.</param>
        /// <returns>A return represents the validation result of the specified Xml fragment string.</returns>
        public static ValidationResult ValidateXml(string xmlValue)
        {
            InitSchemaList();
            InitValidateRecoder();
            DoValidation(xmlValue);
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
        private static void DoValidation(string xmlString)
        {
            XmlReaderSettings validationSettings = new XmlReaderSettings();
            foreach (string eachSchema in schemaList)
            {
                StringReader schemaReader = null;
                try
                {
                    schemaReader = new StringReader(eachSchema);
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
                finally
                {
                    if (schemaReader != null)
                    {
                        schemaReader.Dispose();
                    }
                }
            }

            validationSettings.ValidationType = ValidationType.Schema;
            validationSettings.ConformanceLevel = ConformanceLevel.Document;
            validationSettings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
            validationSettings.ValidationEventHandler += new ValidationEventHandler(ValidationCallBack);

            StringReader reader = null;
            try
            {
                reader = new StringReader(xmlString);
                using (XmlReader validationReader = XmlReader.Create(reader, validationSettings))
                {
                    while (validationReader.Read())
                    {
                    }
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

            // In the current stage, only one WSDL exists.
            schemaList.AddRange(GetSchemaStringFromWsdlFile("MS-FSSHTTP-FSSHTTPB.wsdl"));
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