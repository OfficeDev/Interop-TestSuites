namespace Microsoft.Protocols.TestSuites.MS_LISTSWS
{
    using System;
    using System.Data;
    using System.Linq;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The AdapterHelper class.
    /// </summary>
    public static class AdapterHelper
    {
        #region Constant Values

        /// <summary>
        /// HTTP transport.
        /// </summary>
        public const string TransportHttp = "HTTP";

        /// <summary>
        /// HTTPS transport.
        /// </summary>
        public const string TransportHttps = "HTTPS";

        /// <summary>
        /// The SOAP 1.1
        /// </summary>
        public const string SOAP11 = "SOAP11";

        /// <summary>
        /// The SOAP 1.2.
        /// </summary>
        public const string SOAP12 = "SOAP12";

        /// <summary>
        /// The ows_ prefix.
        /// </summary>
        public const string PrefixOws = "ows_";

        /// <summary>
        /// The field name of ID.
        /// </summary>
        public const string FieldIDName = "ID";

        /// <summary>
        /// The field name of owshiddenversion.
        /// </summary>
        public const string FieldOwshiddenversionName = "owshiddenversion";

        /// <summary>
        /// The field name of Author.
        /// </summary>
        public const string FieldAuthorName = "Author";

        /// <summary>
        /// The value of base content type for a document in a site. 
        /// </summary>
        public const string SiteDocumentContentTypeId = "0x0101";

        /// <summary>
        /// The default protocol doc short name.
        /// </summary>
        public const string DefaultProtocolDocShortName = "MS-LISTSWS";

        /// <summary>
        /// The error string of the SOAP fault.
        /// </summary>
        public const string OutOfRange = "Value does not fall within the expected range.";

        /// <summary>
        /// Custom Send To Destination Name for document library.
        /// </summary>
        public const string SendToDestinationName = "SendToDestinationName";

        /// <summary>
        /// Custom Send To Destination Url for document library.
        /// </summary>
        public const string SendToDestinationUrl = "http://SendToDestinationUrl";

        /// <summary>
        /// The value of MajorVersionLimit
        /// </summary>
        public const int MajorVersionLimitValue = 3;

        /// <summary>
        /// The value of MajorWithMinorVersionsLimit
        /// </summary>
        public const int MajorWithMinorVersionsLimitValue = 5;

        #endregion

        /// <summary>
        /// An ITestSite instance.
        /// </summary>
        private static ITestSite testsite;

        /// <summary>
        /// A method used to initialize this helper class with ITestSite.
        /// </summary>
        /// <param name="testSiteInstance">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void Initialize(ITestSite testSiteInstance)
        {
            testsite = testSiteInstance;
        }

        /// <summary>
        /// A method used to get error code from the specified Soap Exception.
        /// </summary>
        /// <param name="soapEx">A parameter represents the specified Soap Exception.</param>
        /// <returns>A return value represents the error code value in the Soap Exception.</returns>
        public static string GetErrorCodeFromSoapException(SoapException soapEx)
        {
            return Common.ExtractErrorCodeFromSoapFault(soapEx);
        }

        /// <summary>
        /// The helper method is used to retrieve property value from PTF configuration.
        /// </summary>
        /// <param name="propertyName">The name of property to be retrieved.</param>
        /// <returns>The value retrieved from the configuration.</returns>
        public static string GetValueFromConfig(string propertyName)
        {
            return Common.GetConfigurationPropertyValue(propertyName, testsite);
        }

        /// <summary>
        /// A method used to extract a row elements to a DataTable. It is used for "z:row" data.
        /// </summary>
        /// <param name="rowElements">A parameter represents the "z:row" data which need to be extracted.</param>
        /// <returns>A return value represents the extracted DataTable from "z:row" data.</returns>
        public static DataTable ExtractData(XmlNode[] rowElements)
        {
            // If there are rows in the response.
            if (rowElements != null && rowElements.Length != 0)
            {
                using (DataTable dataTable = new DataTable())
                {
                    foreach (XmlElement rowElement in rowElements)
                    {
                        if ("row".Equals(rowElement.LocalName, StringComparison.OrdinalIgnoreCase))
                        {
                            // Create a new row based on the column.
                            DataRow row = dataTable.NewRow();

                            foreach (XmlAttribute column in rowElement.Attributes)
                            {
                                if (!dataTable.Columns.Contains(column.Name))
                                {
                                    dataTable.Columns.Add(column.Name, typeof(string));
                                }

                                row[column.Name] = column.Value;
                            }

                            // Add the row into the table.
                            dataTable.Rows.Add(row);
                        }
                    }

                    return dataTable;
                }
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// A method used to get the value of Element from an XmlElement array by specified name.
        /// </summary>
        /// <param name="elements">A parameter represents the source elements where the method finds the value.</param>
        /// <param name="name">A parameter represents the specified element name which is used to find element's value.</param>
        /// <returns>A return value represents the value of the matched element.</returns>
        public static string GetElementValue(XmlElement[] elements, string name)
        {
            return elements.FirstOrDefault<XmlElement>(e => e.LocalName.Equals(name, StringComparison.OrdinalIgnoreCase)).InnerText;
        }
    }
}