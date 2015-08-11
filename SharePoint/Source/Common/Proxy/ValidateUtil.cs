namespace Microsoft.Protocols.TestSuites.Common
{
    using System.IO;
    using System.Xml;
    using System.Xml.Schema;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class that use to validate schema.
    /// </summary>
    internal class ValidateUtil
    {
        /// <summary>
        /// Implements Microsoft.Protocols.TestTools.IAdapter.Site.
        /// </summary>
        private ITestSite site;

        /// <summary>
        /// A Boolean value indicate that whether throw exception when schema validation not success.
        /// </summary>
        private bool throwException;

        /// <summary>
        /// A Boolean value indicate that whether perform schema validation.
        /// </summary>
        private bool performSchemaValidation;

        /// <summary>
        /// A Boolean value indicate that ignore schema validation for SOAP fault in SOAP1.2.
        /// </summary>
        private bool ignoreSoapFaultSchemaValidationForSoap12;

        /// <summary>
        /// Initializes a new instance of the ValidateUtil class with specified parameters.
        /// </summary>
        /// <param name="testSite">Implements Microsoft.Protocols.TestTools.IAdapter.Site.</param>
        /// <param name="throwException">Indicate that whether throw exception when schema validation not success.</param>
        /// <param name="performSchemaValidation">Indicate that whether perform schema validation.</param>
        /// <param name="ignoreSoapFaultSchemaValidationForSoap12">Indicate that whether ignore schema validation for SOAP fault in SOAP1.2.</param>
        public ValidateUtil(ITestSite testSite, bool throwException, bool performSchemaValidation, bool ignoreSoapFaultSchemaValidationForSoap12)
        {
            this.site = testSite;
            this.throwException = throwException;
            this.performSchemaValidation = performSchemaValidation;
            this.ignoreSoapFaultSchemaValidationForSoap12 = ignoreSoapFaultSchemaValidationForSoap12;
        }

        /// <summary>
        /// Validate the response xml.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">An CustomerEventArgs that contains event data.</param>
        public void ValidateSchema(object sender, CustomerEventArgs e)
        {
            if (!string.IsNullOrEmpty(e.RawRequestXml))
            {
                this.site.Log.Add(LogEntryKind.Debug, "The raw xml request message is:\r\n{0}", e.RawRequestXml);
                XmlDocument requestXml = new XmlDocument();
                requestXml.LoadXml(e.RawRequestXml);
                SchemaValidation.LastRawRequestXml = requestXml.DocumentElement;
            }
            else
            {
                SchemaValidation.LastRawRequestXml = null;
            }

            if (!string.IsNullOrEmpty(e.RawResponseXml))
            {
                this.site.Log.Add(LogEntryKind.Debug, "The raw xml response message is:\r\n{0}", e.RawResponseXml);

                MemoryStream ms = new MemoryStream(System.Text.ASCIIEncoding.ASCII.GetBytes(e.RawResponseXml));
                XmlReader xmlReader = XmlReader.Create(ms);
                e.ValidationXmlReaderOut = xmlReader;

                XmlDocument responseDoc = new XmlDocument();
                responseDoc.LoadXml(e.RawResponseXml);
                SchemaValidation.LastRawResponseXml = responseDoc.DocumentElement;

                if (this.performSchemaValidation)
                {
                    SchemaValidation.ValidateXml(this.site, this.ignoreSoapFaultSchemaValidationForSoap12);

                    if (this.throwException 
                        && ((ValidationResult.Error == SchemaValidation.ValidationResult)
                           || (ValidationResult.Warning == SchemaValidation.ValidationResult)))
                    {
                        throw new XmlSchemaValidationException(SchemaValidation.GenerateValidationResult());
                    }
                }
            }
            else
            {
                SchemaValidation.LastRawResponseXml = null;
            }
        }
    }
}