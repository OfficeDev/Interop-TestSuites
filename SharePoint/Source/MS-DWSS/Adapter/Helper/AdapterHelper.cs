namespace Microsoft.Protocols.TestSuites.MS_DWSS
{
    using System.IO;
    using System.Net;
    using System.Xml;
    using System.Xml.Schema;
    using System.Xml.Serialization;

    /// <summary>
    /// Provide AdapterHelper Methods.
    /// </summary>
    public static class AdapterHelper
    {
        /// <summary>
        /// A const string that represents the schema definition file name of the response messages.
        /// </summary>
        private const string RespSchemaFileName = "DwsResponseResult.xsd";

        /// <summary>
        /// Converts a string that has been HTML-encoded for HTTP transmission into a decoded string, and add the root node using the name specified in parameter rootNodeName.
        /// </summary>
        /// <param name="rootNodeName">The name of root node will be added.</param>
        /// <param name="htmlEncodedString">HTML-encoded string.</param>
        /// <returns>A decoded string that added the root node using the name specified in parameter rootNodeName </returns>
        public static string GenRespXmlString(string rootNodeName, string htmlEncodedString)
        {
            string startRootTag = string.IsNullOrEmpty(rootNodeName) ? string.Empty : string.Format("&lt;{0}&gt;", rootNodeName);
            string endRootTag = string.IsNullOrEmpty(rootNodeName) ? string.Empty : string.Format("&lt;/{0}&gt;", rootNodeName);

            return WebUtility.HtmlDecode(startRootTag + htmlEncodedString + endRootTag);
        }

        /// <summary>
        /// Validate response string according to response schemas.
        /// </summary>
        /// <param name="respXmlString">Response string.</param>
        public static void ValidateRespSchema(string respXmlString)
        {
            XmlReaderSettings validationSettings = new XmlReaderSettings();
            XmlSchemaSet schemaSet = LoadSchemaSet();

            validationSettings.Schemas.Add(schemaSet);
            validationSettings.ValidationType = ValidationType.Schema;
            validationSettings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
            validationSettings.ValidationEventHandler += new ValidationEventHandler(SchemaValidationError);

            using (StringReader strReader = new StringReader(respXmlString))
            {
                XmlReader validationReader = XmlReader.Create(strReader, validationSettings);
                
                while (validationReader.Read())
                {
                }
            }
        }

        /// <summary>
        /// Deserializes the xml string.
        /// </summary>
        /// <typeparam name="T">The type of the object that XML string will be deserialized.</typeparam>
        /// <param name="xmlString">XML string.</param>
        /// <returns>The object deserialized from the XML string.</returns>
        public static T XmlDeserialize<T>(string xmlString)
            where T : class
        {
            // Create an instance of the XmlSerializer specifying type.
            XmlSerializer serializer = new XmlSerializer(typeof(T));

            using (StringReader strReader = new StringReader(xmlString))
            {
                XmlReader reader = XmlReader.Create(strReader);
                return serializer.Deserialize(reader) as T;
            }
        }

        /// <summary>
        /// Serialize the specified object to XML string.
        /// </summary>
        /// <param name="obj">The specified object need be serialized.</param>
        /// <returns>The XML string serialized from the specified object.</returns>
        public static string XmlSerialize(object obj)
        {
            XmlSerializer serializer = new XmlSerializer(obj.GetType());

            XmlWriterSettings serializeSettings = new XmlWriterSettings();

            serializeSettings.OmitXmlDeclaration = true;

            XmlSerializerNamespaces ns = new XmlSerializerNamespaces(new[] { XmlQualifiedName.Empty });

            XmlWriter xmlWriter = null;

            using (StringWriter strWriter = new StringWriter())
            {
                xmlWriter = XmlWriter.Create(strWriter, serializeSettings);
                serializer.Serialize(xmlWriter, obj, ns);

                return strWriter.ToString();
            }
        }

        /// <summary>
        /// Load XML schemas from specified .xsd file.
        /// </summary>
        /// <param name="xsdFileName">The file name of .xsd.</param>
        /// <returns>XmlSchema instance.</returns>
        private static XmlSchema LoadSchema(string xsdFileName)
        {
            using (FileStream fs = new FileStream(xsdFileName, FileMode.Open))
            {
                return XmlSchema.Read(fs, new ValidationEventHandler(SchemaCompileError));
            }
        }

        /// <summary>
        /// Load several XML schemas from specified .xsd files and return an instance of XmlSchemaSet.
        /// </summary>
        /// <returns>XmlSchemaSet instance</returns>
        private static XmlSchemaSet LoadSchemaSet()
        {
            XmlSchemaSet schemaSet = new XmlSchemaSet();
            schemaSet.ValidationEventHandler += new ValidationEventHandler(SchemaCompileError);
            schemaSet.Add(LoadSchema(RespSchemaFileName));
            schemaSet.Compile();

            return schemaSet;
        }

        /// <summary>
        /// Handle the schema compile errors.
        /// </summary>
        /// <param name="sender">sender object.</param>
        /// <param name="args">args object.</param>
        private static void SchemaCompileError(object sender, ValidationEventArgs args)
        {
            throw new XmlSchemaException(args.Message, args.Exception);
        }

        /// <summary>
        /// Handle the schema validation errors.
        /// </summary>
        /// <param name="sender">sender object.</param>
        /// <param name="args">args object.</param>
        private static void SchemaValidationError(object sender, ValidationEventArgs args)
        {
            throw new XmlSchemaValidationException(args.Message, args.Exception);
        }
    }
}