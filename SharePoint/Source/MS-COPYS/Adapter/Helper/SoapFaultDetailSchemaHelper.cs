namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    using System;
    using System.Web.Services.Protocols;
    using System.Xml;
    using System.Xml.Schema;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This class is used to process the schema validation for the SOAP fault of detail element.
    /// </summary>
    public class SoapFaultDetailSchemaHelper
    {
        /// <summary>
        /// A method used to validate the schema definition for the detail element of a SOAP fault. If the schema validation is not successful, this method will raise an XmlSchemaValidationException.
        /// </summary>
        /// <param name="soapEx">A parameter represents the SoapException instance which will be recorded.</param>
        /// <param name="testSiteInstance">A parameter represents the ITestSite instance which contains test context.</param>
        public static void ValidateSoapFaultDetail(SoapException soapEx, ITestSite testSiteInstance)
        { 
           if (null == soapEx)
           {
               throw new ArgumentNullException("soapEx");
           }

          if (null == testSiteInstance)
          {
             throw new ArgumentNullException("testSiteInstance");
          }

          RecordSoapExceptionInfor(soapEx, testSiteInstance);

          string detailValue = LoadSoapFaultDetailXml(soapEx);
          SchemaValidation.ValidateXml(testSiteInstance, detailValue);
          if (SchemaValidation.XmlValidationErrors.Count != 0 || SchemaValidation.XmlValidationWarnings.Count != 0)
          {
              string errorString = string.Format("There are schema validation issues for detail element of SOAP fault.\r\n:{0}", SchemaValidation.GenerateValidationResult());
              throw new XmlSchemaValidationException(errorString);
          }
        }

        /// <summary>
        /// A method used to record a SOAP Exception information into the log.
        /// </summary>
        /// <param name="soapEx">A parameter represents the SoapException instance which will be record.</param>
        /// <param name="testSiteInstance">A parameter represents the ITestSite instance which contains test context.</param>
        private static void RecordSoapExceptionInfor(SoapException soapEx, ITestSite testSiteInstance)
        {
            string detailOutPut = string.Empty;
            if (null == soapEx.Detail || string.IsNullOrEmpty(soapEx.Detail.OuterXml))
            {
                detailOutPut = "None";
            }
            else
            {
                detailOutPut = soapEx.Detail.OuterXml;
            }

            TestSuiteManageHelper.Initialize(testSiteInstance);
            testSiteInstance.Log.Add(
                                    LogEntryKind.Debug,
                                    @"There is a SoapException generated. Information:\r\nMessage:[{0}]\r\nStackTrace:[{1}]\r\ndetail:[{2}]\r\nSoapVersion:[{3}]\r\nUrl:[{4}]",
                                    soapEx.Message,
                                    soapEx.StackTrace,
                                    detailOutPut,
                                    TestSuiteManageHelper.CurrentSoapVersion,
                                    string.IsNullOrEmpty(soapEx.Role) ? "None/Empty" : soapEx.Role);
        }

        /// <summary>
        /// A method used to load the XML string for the detail element in the SOAP exception.
        /// </summary>
        /// <param name="soapEx">A parameter represents SOAP exception.</param>
        /// <returns>A return value represents the xml string of the "Detail" property in SOAP exception.</returns>
        private static string LoadSoapFaultDetailXml(SoapException soapEx)
        {
            if (null == soapEx.Detail || string.IsNullOrEmpty(soapEx.Detail.OuterXml))
            {
                throw new ArgumentException("The SOAP exception should contain the detail value.");
            }

            // Load the original "detail" element.
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(soapEx.Detail.OuterXml);
            XmlElement originalDetailElement = xmlDoc.DocumentElement;

            if (!originalDetailElement.LocalName.Equals("detail", StringComparison.OrdinalIgnoreCase))
            {
                string errorMsg = string.Format("The root element should be \"Detail\" or \"detail\".Current XML fragment:\r\n[{0}]", soapEx.Detail.OuterXml);
                throw new XmlSchemaValidationException(errorMsg);
            }
 
            // [Note] According to SOAP1.1 and SOAP 1.2 definitions, SOAP1.1 defines subelement 'detail' in SOAP Fault element, SOAP1.2 defines subelement "Detail". But actually, whether the soap fault response message are formatted with SOAP 1.1 or 1.2, the subelement name in O12, O14 and O15 response is always "detail".
            // This test suite does not verify the SOAP fault element whether contain the correct "detail" or "Detail" element in sequence. The "detail" and "Detail" element is defined as extendable element in W3C standard, there are some different implementations for them.
            // This test suite only verify the "detail" or "Detail" element whether contain the correct extendable content which is defined in MS-COPYS.(child elements definitions of as described in section 2.2.2.1), so the test suite will construct a new "detail" element to store all the child elements in order to perform the schema validation. 
            XmlDocument newXmlfragment = new XmlDocument();

            // This test suite will build a new "detail" element which contains the actual detail information, and this new "detail" element always point to the "http://schemas.microsoft.com/sharepoint/soap/" name space.
            XmlElement newBuildDetailElement = newXmlfragment.CreateElement("detail");

            // Copy the attributes from original detail element.
            if (null != originalDetailElement.Attributes && 0 != originalDetailElement.Attributes.Count)
            {
                foreach (XmlAttribute attributeItem in originalDetailElement.Attributes)
                {
                    XmlAttribute newAddedAttribute = newXmlfragment.CreateAttribute(attributeItem.Name, attributeItem.NamespaceURI);
                    newAddedAttribute.Value = attributeItem.Value;
                    newBuildDetailElement.Attributes.Append(newAddedAttribute);
                }
            }

            // Copy the child nodes from original detail element. 
            foreach (XmlNode xmlNodeItem in originalDetailElement.ChildNodes)
            {
                newXmlfragment.LoadXml(xmlNodeItem.OuterXml);
                newBuildDetailElement.AppendChild(newXmlfragment.DocumentElement);
            }

            // Point to the "http://schemas.microsoft.com/sharepoint/soap/" namespace which matches the WSDL used by test suite.
            newBuildDetailElement.SetAttribute("xmlns", @"http://schemas.microsoft.com/sharepoint/soap/");
            return newBuildDetailElement.OuterXml;
         }
    }
}