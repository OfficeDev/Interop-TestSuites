namespace Microsoft.Protocols.TestSuites.MS_WWSP
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System.Xml;
    using System.Xml.Schema;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A Class represent the XML schema validation function.
    /// </summary>
    public class XmlValidator
    {   
        /// <summary>
        /// The result of schema validation.
        /// </summary>
        private static ValidationResult validationResult;

        /// <summary>
        /// The warning results of XML Schema validation.
        /// </summary>
        private static List<ValidationEventArgs> xmlValidationWarnings;

        /// <summary>
        /// The error results of XML Schema validation.
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
        /// Gets the error results of XML Schema validation.
        /// </summary>
        public static List<ValidationEventArgs> XmlValidationErrors
        {
            get
            {
                return xmlValidationErrors;
            }
        }

        /// <summary>
        /// Gets the warnings results of XML Schema validation.
        /// </summary>
        public static List<ValidationEventArgs> XmlValidationWarnings
        {
            get
            {
                return xmlValidationWarnings;
            }
        }

        /// <summary>
        /// Validate a piece of Xml fragment according the given Xml Schema.
        /// Both Xml Schema and Xml fragment are String type.
        /// </summary>
        /// <param name="schemaList">Xml Schema string.</param>
        /// <param name="xmlString">Xml fragment string.</param>
        /// <returns>A return represents the validation result of the specified Xml fragment string.</returns>
        public static ValidationResult ValidateXml(List<string> schemaList, string xmlString)
        {
            #region Initialize the validator's recorder.

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

            #endregion

            XmlReaderSettings validationSettings = new XmlReaderSettings();
            foreach (string eachSchema in schemaList)
            {
                XmlReader xsdXmlReader = XmlReader.Create(new StringReader(eachSchema));
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

            validationSettings.ValidationType = ValidationType.Schema;
            validationSettings.ConformanceLevel = ConformanceLevel.Document;
            validationSettings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
            validationSettings.ValidationEventHandler += new ValidationEventHandler(ValidationCallBack);
            XmlReader validationReader = XmlReader.Create(new StringReader(xmlString), validationSettings);

            while (validationReader.Read()) 
            {
            }

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
    }
}