namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    using System.Collections.Generic;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// This class provides the methods to verify requirements.
    /// </summary>
    public partial class MS_WEBSSAdapter
    {
        /// <summary>
        /// This method is used to capture requirements of detail complex type.
        /// </summary>
        /// <param name="detail">The detail complex type.</param>
        private void ValidateDetail(XmlNode detail)
        {
            Site.Assert.IsNotNull(detail, "The XmlNode cannot be NULL");

            // Verify MS-WEBSS requirement: MS-WEBSS_R38
            bool isDetail = SchemaValidation.ValidateXml(this.Site, SchemaValidation.GetSoapFaultDetailBody(detail.OuterXml)) == ValidationResult.Success;

            Site.CaptureRequirementIfIsTrue(
                isDetail,
                38,
                @"[In Message Processing Events and Sequencing Rules] The following schema specifies the structure of the detail element in the SOAP fault used by this protocol[MS-WEBSS].
 <s:element name=""detail"">
  <s:complexType>
    <s:sequence>
      <s:element name=""errorString"" type=""s:string"" minOccurs=""1"" maxOccurs=""1""/>
      <s:element name=""errorCode"" type=""s:string"" minOccurs=""0"" maxOccurs=""1""/>
    </s:sequence>
  </s:complexType>
</s:element>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R12
            Site.CaptureRequirementIfIsTrue(
                isDetail,
                12,
                @"[In Protocol Details] This protocol[MS-WEBSS] enables protocol servers to provide additional details for SOAP faults by including either a detail element as specified in [SOAP1.1] section 4.4 or a Detail element as specified in [SOAP1.2-1/2007] section 5.4.5, which conforms to the XML schema of the SOAPFaultDetails complex type specified in section 2.2.4.1.");

            // If MS-WEBSS_R38 is captured, the schema including the elements errorString and errorCode, therefore MS-WEBSS_R39 will be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R39
            Site.CaptureRequirementIfIsTrue(
                isDetail,
                39,
                @"[In Message Processing Events and Sequencing Rules] [SOAP fault] detail: A container for errorString and errorCode elements.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R686
            Site.CaptureRequirementIfIsTrue(
                isDetail,
                686,
                @"[In SOAPFaultDetails] This complex type[SOAPFaultDetails] is defined as follows:
    <s:schema xmlns:s=""http://www.w3.org/2001/XMLSchema"" targetNamespace="" http://schemas.microsoft.com/sharepoint/soap"">
    <s:complexType name=""detail"">
        <s:sequence>
            <s:element name=""errorstring"" type=""s:string""/>
            <s:element name=""errorcode"" type=""s:string"" minOccurs=""0""/>
        </s:sequence>
    </s:complexType>
</s:schema>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R1091
            if (Common.IsRequirementEnabled(1091, this.Site))
            {
                Site.CaptureRequirementIfIsTrue(
                 isDetail,
                 1091,
                @"[In Appendix B: Product Behavior] Implementation does support this method[UpdateContentType]. (<1> Section 2.2.4.1: This attribute is returned in Microsoft SharePoint Foundation 2010 and above follow this behavior).");
            }
        }

        /// <summary>
        /// Capture underlying transport protocol related requirements.
        /// </summary>
        private void CaptureTransportRelatedRequirements()
        {
            SoapProtocolVersion soapVersion = Common.GetConfigurationPropertyValue<SoapProtocolVersion>("SoapVersion", this.Site);
            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);

            if (transport == TransportProtocol.HTTP)
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R1
                // COMMENT: The transport protocol is HTTP, then having received the response 
                // successfully have proved the HTTP transport is supported. If the HTTP transport is 
                // not supported, the response can't be received successfully.
                this.Site.CaptureRequirement(
                    1,
                    @"[In Transport] Protocol servers MUST support SOAP over HTTP.");
            }
            else if (transport == TransportProtocol.HTTPS)
            {
                if (Common.IsRequirementEnabled(5, this.Site))
                {
                    // Verify requirement: MS-WEBSS_R5 derived from MS-WEBSS_R2
                    // COMMENT: The transport protocol is HTTPS, then having received the response 
                    // successfully have proved the HTTPS transport is supported. If the HTTPS transport is 
                    // not supported, the response can't be received successfully.
                    Site.CaptureRequirement(
                        5,
                        @"[In Appendix B: Product Behavior] [In Transport] Implementation does additionally support SOAP over HTTPS for securing communication with clients.(Windows SharePoint Services 3.0 and above products follow this behavior.)");
                }
            }

            if (soapVersion == SoapProtocolVersion.Soap11)
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R665
                // COMMENT: The SOAP version is SOAP11, then having received the response 
                // successfully have proved the protocol messages are formatted as specified in [SOAP1.1]. 
                // If the protocol messages are not formatted as specified in [SOAP1.1]. , the response 
                // can't be received successfully.
                Site.CaptureRequirement(
                    665,
                    @"[In Transport] Protocol messages MUST be formatted as specified in [SOAP1.1] section 4, SOAP Envelope.");
            }
            else if (soapVersion == SoapProtocolVersion.Soap12)
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R1040
                // COMMENT: The SOAP version is SOAP12, then having received the response 
                // successfully have proved the protocol messages are formatted as specified in [SOAP1.2/1]. 
                // If the protocol messages are not formatted as specified in [SOAP1.2/1], the response 
                // can't be received successfully.
                Site.CaptureRequirement(
                    1040,
                    @"[In Transport] Protocol messages MUST be formatted as specified in [SOAP1.2-1/2007] section 5, SOAP Message Construct.");
            }
        }

        /// <summary>
        /// This method is used to capture requirements of WebDefinition complex type.
        /// </summary>
        /// <param name="web">The WebDefinition complex type.</param>
        private void ValidateWebDefinition(WebDefinition web)
        {
            // Proxy handles operation's SoapIn and SoapOut, if the server didn't respond the SoapOut message for the operation, Proxy will fail. 
            // Proxy didn't fail now, that indicates server responds corresponding SoapOut message for the operation.So the requirement be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R691
            if (Common.IsRequirementEnabled(691, this.Site))
            {
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                    691,
                    @"[In WebDefinition] This type[WebDefinition] is defined as follows:
 <s:complexType name=""WebDefinition"">
  <s:attribute name=""Title"" type=""s:string"" use=""required"" />
  <s:attribute name=""Url"" type=""s:string"" use=""required"" />
  <s:attribute name=""Description"" type=""s:string"" />
  <s:attribute name=""Language"" type=""s:string"" />
  <s:attribute name=""Theme"" type=""s:string"" />
  <s:attribute name=""FarmId"" type=""core:UniqueIdentifierWithBraces"" />
  <s:attribute name=""Id"" type=""core:UniqueIdentifierWithBraces"" />
  <s:attribute name=""SiteId"" type =""core: UniqueIdentifierWithBraces"" />
  <s:attribute name = ""IsSPO"" type = ""core:TRUEFALSE"" />
  <s:attribute name=""ExcludeFromOfflineClient"" type=""core:TRUEFALSE"" />
  <s:attribute name=""CellStorageWebServiceEnabled"" type=""core:TRUEFALSE"" />
  <s:attribute name=""AlternateUrls"" type=""s:string"" />
</s:complexType>");
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R1092
            // COMMENT: When the SUT product is Windows SharePoint Services 3.0, if the 
            // Id attribute of the returned WebDefinition is null or empty, which means the Id 
            // attribute is not returned in Windows SharePoint Services 3.0, then the following 
            // requirement can be captured.
            if (Common.IsRequirementEnabled(1092, this.Site))
            {
                Site.CaptureRequirementIfIsTrue(
                !string.IsNullOrEmpty(web.FarmId),
                1092,
                @"[In Appendix B: Product Behavior]Implementation does support this method [GetWeb]. (<2> Section 2.2.4.2: This attribute is returned in Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R1093
            // COMMENT: When the SUT product is Windows SharePoint Services 3.0, if the 
            // ExcludeFromOfflineClient attribute of the returned WebDefinition is null or empty, 
            // which means the ExcludeFromOfflineClient attribute is not returned in Windows 
            // SharePoint Services 3.0, then the following requirement can be captured.
            if (Common.IsRequirementEnabled(1093, this.Site))
            {
                Site.CaptureRequirementIfIsTrue(
                !string.IsNullOrEmpty(web.Id),
                1093,
                @"[In Appendix B: Product Behavior]Implementation does support this method [GetWeb]. (<3> Section 2.2.4.2: This attribute is returned in Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(69700101, this.Site))
            {
                Site.CaptureRequirementIfIsTrue(
                !string.IsNullOrEmpty(web.SiteId),
                69700101,
                @"[In Appendix B: Product Behavior]Implementation does support this method [SiteId]. (<3> Section 2.2.4.2: This attribute is returned in Microsoft SharePoint Foundation 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(69700201, this.Site))
            {
                Site.CaptureRequirementIfIsTrue(
                !string.IsNullOrEmpty(web.IsSPO),
                69700201,
                @"[In Appendix B: Product Behavior]Implementation does support this method [IsSPO]. (<4> Section 2.2.4.2: This attribute is returned in Microsoft SharePoint Server 2016 and above follow this behavior.)");
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R1094
            // COMMENT: When the SUT product is Windows SharePoint Services 3.0, if the 
            // CellStorageWebServiceEnabled attribute of the returned WebDefinition is null or 
            // empty, which means the CellStorageWebServiceEnabled attribute is not returned 
            // in Windows SharePoint Services 3.0, then the following requirement can be 
            // captured.
            if (Common.IsRequirementEnabled(1094, this.Site))
            {
                Site.CaptureRequirementIfIsTrue(
                    !string.IsNullOrEmpty(web.ExcludeFromOfflineClient),
                    1094,
                    @"[In Appendix B: Product Behavior]Implementation does support this method [GetWeb]. (<6> Section 2.2.4.2: This attribute is returned in Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R1095
            // COMMENT: When the SUT product is Windows SharePoint Services 3.0, if the 
            // AlternateUrls attribute of the returned WebDefinition is null or empty, which means 
            // the AlternateUrls attribute is not returned in Windows SharePoint Services 3.0, 
            // then the following requirement can be captured.
            if (Common.IsRequirementEnabled(1095, this.Site))
            {
                Site.CaptureRequirementIfIsTrue(
                !string.IsNullOrEmpty(web.CellStorageWebServiceEnabled),
                1095,
                @"[In Appendix B: Product Behavior]Implementation does support this method [GetWeb]. (<7> Section 2.2.4.2: This attribute is returned in Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R1096
            // COMMENT: When the SUT product is Windows SharePoint Services 3.0, if the 
            // AlternateUrls attribute of the returned WebDefinition is null or empty, which means 
            // the AlternateUrls attribute is not returned in Windows SharePoint Services 3.0, 
            // then the following requirement can be captured.
            if (Common.IsRequirementEnabled(1096, this.Site))
            {
                Site.CaptureRequirementIfIsTrue(
                !string.IsNullOrEmpty(web.AlternateUrls),
                1096,
                    @"[In Appendix B: Product Behavior]Implementation does support this method [GetWeb]. (<8> Section 2.2.4.2: This attribute is returned in Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This method is used to capture requirements of WebDefinition complex type For Sub Web Collection.
        /// </summary>
        private void ValidateWebDefinitionForSubWebCollection()
        {
            // Proxy handles operation's SoapIn and SoapOut, if the server didn't respond the SoapOut message for the operation, Proxy will fail. 
            // Proxy didn't fail now, that indicates server responds corresponding SoapOut message for the operation.So the requirement be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R691
            if (Common.IsRequirementEnabled(691, this.Site))
            {
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                    691,
                    @"[In WebDefinition] This type[WebDefinition] is defined as follows:
 <s:complexType name=""WebDefinition"">
  <s:attribute name=""Title"" type=""s:string"" use=""required"" />
  <s:attribute name=""Url"" type=""s:string"" use=""required"" />
  <s:attribute name=""Description"" type=""s:string"" />
  <s:attribute name=""Language"" type=""s:string"" />
  <s:attribute name=""Theme"" type=""s:string"" />
  <s:attribute name=""FarmId"" type=""core:UniqueIdentifierWithBraces"" />
  <s:attribute name=""Id"" type=""core:UniqueIdentifierWithBraces"" />
  <s:attribute name=""ExcludeFromOfflineClient"" type=""core:TRUEFALSE"" />
  <s:attribute name=""CellStorageWebServiceEnabled"" type=""core:TRUEFALSE"" />
  <s:attribute name=""AlternateUrls"" type=""s:string"" />
</s:complexType>");
            }
        }

        /// <summary>
        /// Capture CreateContentType related requirements.
        /// </summary>
        private void ValidateCreateContentType()
        {
            // Proxy handles operation's SoapIn and SoapOut, if the server didn't respond the soapout message for the operation, Proxy will fail. 
            // Proxy didn't fail now, that indicates server responds corresponding soapout message for the operation.So the requirement be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R846
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                846,
                @"[In CreateContentType] This operation[CreateContentType] is defined as follows:
 <wsdl:operation name=""CreateContentType"">
    <wsdl:input message=""tns:CreateContentTypeSoapIn"" />
    <wsdl:output message=""tns:CreateContentTypeSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R44
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                44,
                @"[In CreateContentType] The protocol client sends a CreateContentTypeSoapIn request message, and the protocol server responds with a CreateContentTypeSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R77
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                77,
                @"[In CreateContentTypeResponse] This element[CreateContentTypeResponse] contains the response to a request to create a new content type on the context site.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R78
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                78,
                @"[In CreateContentTypeResponse] This element[CreateContentTypeResponse] is defined as follows:
<s:element name=""CreateContentTypeResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""CreateContentTypeResult"" type=""s:string"" minOccurs=""1""/>
    </s:sequence>
  </s:complexType>
</s:element>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R59
            bool isVerifiedR59 = this.HasElement(SchemaValidation.LastRawResponseXml, "CreateContentTypeResponse");
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR59,
                59,
                @"[In CreateContentTypeSoapOut] The SOAP body contains a CreateContentTypeResponse element.");
            //Verify MS-WEBSS requirement: MS-WEBSS_R712001
            string sPattern = "0x([0-9A-Fa-f][1-9A-Fa-f]|[1-9A-Fa-f][0-9A-Fa-f]|00[0-9A-Faf]{32})*";
            bool isVerifiedR712001 = false;
            bool isVerifiedR712001length = false;
            isVerifiedR712001 = System.Text.RegularExpressions.Regex.IsMatch(SchemaValidation.LastRawResponseXml.InnerText, sPattern);
            if (SchemaValidation.LastRawResponseXml.InnerText.Length <1027 &&SchemaValidation.LastRawResponseXml.InnerText.Length >1)
            {
                isVerifiedR712001length = true;
            }
            Site.CaptureRequirementIfIsTrue
            (isVerifiedR712001&&isVerifiedR712001length, 
                712001,
                @"[In CreateContentTypeResponse] CreateContentTypeResult: It MUST conform to the ContentTypeId type, as specified in [MS-WSSCAML] section 2.3.1.4."
            );
        }

        /// <summary>
        /// Capture DeleteContentType related requirements.
        /// </summary>
        private void ValidateDeleteContentType()
        {
            // Proxy handles operation's SoapIn and SoapOut, if the server didn't respond the soapout message for the operation, Proxy will fail. 
            // Proxy didn't fail now, that indicates server responds corresponding soapout message for the operation.So the requirement be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R104
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                104,
                @"[In DeleteContentType] This operation[DeleteContentType] is defined as follows:
<wsdl:operation name=""DeleteContentType"">
    <wsdl:input message=""tns:DeleteContentTypeSoapIn"" />
    <wsdl:output message=""tns:DeleteContentTypeSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R105
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                105,
                @"[In DeleteContentType] The protocol client sends a DeleteContentTypeSoapIn request message, and the protocol server responds with a DeleteContentTypeSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R113
            Site.CaptureRequirementIfIsTrue(
                this.HasElement(SchemaValidation.LastRawResponseXml, "DeleteContentTypeResponse"),
                113,
                @"[In DeleteContentTypeSoapOut] The SOAP body contains a DeleteContentTypeResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R118
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                118,
                @"[In DeleteContentTypeResponse] This element is defined as follows:
<s:element name=""DeleteContentTypeResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""DeleteContentTypeResult"" minOccurs=""0"">
        <s:complexType mixed=""true"">
          <s:sequence>
            <s:element name=""Success"">
               <s:complexType/>
            </s:element>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");
        }

        /// <summary>
        /// This method is used to capture requirements of GetActivatedFeatures operation.
        /// </summary>
        private void ValidateGetActivatedFeatures()
        {
            // Proxy handles operation's SoapIn and SoapOut, if the server didn't respond the SoapOut message for the operation, Proxy will fail. 
            // Proxy didn't fail now, that indicates server responds corresponding SoapOut message for the operation.So the requirement be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R122
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                122,
                @"[In GetActivatedFeatures] This operation[GetActivatedFeatures] is defined as follows:
<wsdl:operation name=""GetActivatedFeatures"">
    <wsdl:input message=""tns:GetActivatedFeaturesSoapIn"" />
    <wsdl:output message=""tns:GetActivatedFeaturesSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R123
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                123,
                @"[In GetActivatedFeatures] The protocol client sends a GetActivatedFeaturesSoapIn request message, and the protocol server responds with a GetActivatedFeaturesSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R124
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                124,
                @"[In GetActivatedFeatures] The GetActivatedFeaturesSoapOut message MUST contain a single GetActivatedFeaturesResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R134
            Site.CaptureRequirementIfIsTrue(
                this.HasElement(SchemaValidation.LastRawResponseXml, "GetActivatedFeaturesResponse"),
                134,
                @"[In GetActivatedFeaturesSoapOut] The SOAP body contains a GetActivatedFeaturesResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R136
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                136,
                @"[In GetActivatedFeaturesResponse] The SOAP body contains a GetActivatedFeaturesResponse element, which has the following definition:
<s:element name=""GetActivatedFeaturesResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""GetActivatedFeaturesResult"" type=""s:string"" minOccurs=""0"" maxOccurs=""1"" />
    </s:sequence>
  </s:complexType>
</s:element>");
        }

        /// <summary>
        /// This method is used to capture requirements of GetAllSubWebCollection operation.
        /// </summary>
        /// <param name="getAllSubWebCollectionResult">The result of the operation.</param>
        private void ValidateGetAllSubWebCollection(
            GetAllSubWebCollectionResponseGetAllSubWebCollectionResult getAllSubWebCollectionResult)
        {
            // Proxy handles operation's SoapIn and SoapOut, if the server didn't respond the SoapOut message for the operation, Proxy will fail. 
            // Proxy didn't fail now, that indicates server responds corresponding SoapOut message for the operation.So the requirement be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R142
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                142,
                @"[In GetAllSubWebCollection] This operation[GetAllSubWebCollection] is defined as follows:
<wsdl:operation name=""GetAllSubWebCollection"">
    <wsdl:input message=""tns:GetAllSubWebCollectionSoapIn"" />
    <wsdl:output message=""tns:GetAllSubWebCollectionSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R143
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                143,
                @"[In GetAllSubWebCollection] [The protocol client sends a GetAllSubWebCollectionSoapIn request message] The protocol server responds with a GetAllSubWebCollectionSoapOut response message.");

            // Ensure the SOAP result is returned successfully.
            Site.Assume.IsNotNull(
                getAllSubWebCollectionResult,
                "The result of GetAllSubWebCollection operation must not be null.");

            // COMMENT: There is at least one more site besides the one we added in the 
            // TestInitialize in the server. If the response contains at least 2 sites and one of them 
            // has the expected Title and the expected URL, then the following requirement can be 
            // captured.
            bool isVerifiedR148 = false;
            string subSiteUrl = Common.GetConfigurationPropertyValue("SubSiteUrl", this.Site);

            if (getAllSubWebCollectionResult.Webs.Length > 1)
            {
                foreach (WebDefinition w in getAllSubWebCollectionResult.Webs)
                {
                    string testSiteTitle = Common.GetConfigurationPropertyValue("TestSiteTitle", this.Site);
                    if (w.Title.Equals(testSiteTitle, System.StringComparison.OrdinalIgnoreCase) &&
                        w.Url.Equals(subSiteUrl, System.StringComparison.OrdinalIgnoreCase))
                    {
                        isVerifiedR148 = true;
                        break;
                    }
                }
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R148
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR148,
                148,
                @"[In GetAllSubWebCollectionSoapOut] This message[GetAllSubWebCollectionSoapOut] returns the title and URL of all sites in the site collection.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R150
            Site.CaptureRequirementIfIsTrue(
                this.HasElement(SchemaValidation.LastRawResponseXml, "GetAllSubWebCollectionResponse"),
                150,
                @"[In GetAllSubWebCollectionSoapOut] The SOAP body contains a GetAllSubWebCollectionResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R152
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                152,
                @"[In GetAllSubWebCollectionResponse] The SOAP body contains a GetAllSubWebCollectionResponse element, which has the following definition:
<s:element name=""GetAllSubWebCollectionResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""GetAllSubWebCollectionResult"" minOccurs=""0"" maxOccurs=""1"" >
        <s:complexType mixed=""true"">
          <s:sequence>
            <s:element name=""Webs"">
              <s:complexType>
                <s:sequence>
                  <s:element name=""Web"" type=""tns:WebDefinition"" minOccurs=""1"" maxOccurs=""unbounded"" />
                </s:sequence>
              </s:complexType>
            </s:element>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R153
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                153,
                @"[In GetAllSubWebCollectionResponse] GetAllSubWebCollectionResult: Contains a single Webs element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R154
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                154,
                @"[In GetAllSubWebCollectionResponse] Webs: This element contains one or more Web elements.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R155
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                155,
                @"[In GetAllSubWebCollectionResponse] Web: This element is of type WebDefinition, as specified in section 2.2.4.2.");

            // COMMENT: There is at least one more site besides the one we added in the 
            // TestInitialize in the server. If the response contains at least 2 sites and one of them 
            // has the expected Title and the expected URL, then the following requirement can be 
            // captured.
            bool isVerifiedR158 = false;
            if (getAllSubWebCollectionResult.Webs.Length > 1)
            {
                foreach (WebDefinition w in getAllSubWebCollectionResult.Webs)
                {
                    if (w.Title.Equals(Common.GetConfigurationPropertyValue("TestSiteTitle", this.Site), System.StringComparison.OrdinalIgnoreCase)
                        && w.Url.Equals(subSiteUrl, System.StringComparison.OrdinalIgnoreCase))
                    {
                        isVerifiedR158 = true;
                        break;
                    }
                }
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R158
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR158,
                158,
                @"[In GetAllSubWebCollectionResponse] The GetAllSubWebCollectionResult element MUST contain one child Webs element, which MUST contain a sequence of one or more child Web elements, one for each site in the current site collection.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R159
            // COMMENT: There is at least one more site besides the one we added in the 
            // TestInitialize in the server. If the response contains at least 2 sites and one of them 
            // has the expected Title, then the following requirement can be captured.
            bool isVerifiedR159 = false;
            if (getAllSubWebCollectionResult.Webs.Length > 1)
            {
                foreach (WebDefinition w in getAllSubWebCollectionResult.Webs)
                {
                    if (w.Title.Equals(Common.GetConfigurationPropertyValue("TestSiteTitle", this.Site), System.StringComparison.OrdinalIgnoreCase))
                    {
                        isVerifiedR159 = true;
                        break;
                    }
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR159,
                159,
                @"[In GetAllSubWebCollectionResponse] Each Web element MUST specify the title  of one site in the collection.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R713
            // COMMENT: There is at least one more site besides the one we added in the 
            // TestInitialize in the server. If the response contains at least 2 sites and one of them 
            // has the expected URL, then the following requirement can be captured.
            bool isVerifiedR713 = false;
            if (getAllSubWebCollectionResult.Webs.Length > 1)
            {
                foreach (WebDefinition w in getAllSubWebCollectionResult.Webs)
                {
                    if (w.Url.Equals(subSiteUrl, System.StringComparison.OrdinalIgnoreCase))
                    {
                        isVerifiedR713 = true;
                        break;
                    }
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR713,
                713,
                @"[In GetAllSubWebCollectionResponse] Each Web element MUST specify the URL of one site in the collection.");
        }

        /// <summary>
        /// Capture GetContentTypes related requirements.
        /// </summary>
        private void ValidateGetContentTypes()
        {
            // Proxy handles operation's SoapIn and SoapOut, if the server didn't respond the soapout message for the operation, Proxy will fail. 
            // Proxy didn't fail now, that indicates server responds corresponding soapout message for the operation.So the requirement be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R230
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                230,
                @"[In GetContentTypes] This operation[GetContentTypes]<16> retrieves all content types for a specified context site. This operation is defined as follows:
<wsdl:operation name=""GetContentTypes"">
    <wsdl:input message=""tns:GetContentTypesSoapIn"" />
    <wsdl:output message=""tns:GetContentTypesSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R227
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                227,
                @"[In GetContentTypes] [The protocol client sends a GetContentTypesSoapIn request message] the protocol server responds with a GetContentTypesSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R240
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                240,
                @"[In GetContentTypesResponse] This[GetContentTypesResponse] element is defined as follows:
<s:element name=""GetContentTypesResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""GetContentTypesResult"" minOccurs=""0"">
        <s:complexType>
          <s:sequence>
            <s:element name=""ContentTypes"" minOccurs=""1"" maxOccurs=""1"">
             <s:complexType>
              <s:sequence>
               <s:element name=""ContentType"" minOccurs=""0"" maxOccurs=""unbounded"">
                <s:complexType>
                 <s:sequence />
                 <s:attribute name=""Name"" type=""s:string"" use=""required""/>
                 <s:attribute name=""ID"" type=""core:ContentTypeId"" use=""required"" />
                 <s:attribute name=""Description"" type=""s:string"" use=""required"" />
                 <s:attribute name=""NewDocumentControl"" type=""s:string"" use=""required""/>
                  <s:attribute name=""Scope"" type=""s:string"" use=""required"" />
                  <s:attribute name=""Version"" type=""s:int"" use=""required"" />
                  <s:attribute name=""RequireClientRenderingOnNew"" type=""core:TRUEFALSE"" use=""required"" />
                 </s:complexType>
                </s:element>
              </s:sequence>
             </s:complexType>
            </s:element>
       </s:sequence>
      </s:complexType>
     </s:element>
    </s:sequence>
  </s:complexType>  
</s:element>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R236
            Site.CaptureRequirementIfIsTrue(
                this.HasElement(SchemaValidation.LastRawResponseXml, "GetContentTypesResponse"),
                236,
                @"[In GetContentTypesSoapOut] The SOAP body contains a GetContentTypesResponse element.");
        }

        /// <summary>
        /// Capture GetCustomizedPageStatus operation related requirements.
        /// </summary>
        /// <param name="customizedPageStatus">The return result of GetCustomizedPageStatus operation.</param>
        private void ValidateGetCustomizedPageStatus(CustomizedPageStatus customizedPageStatus)
        {
            // Verify MS-WEBSS requirement: MS-WEBSS_R257.
            // The GetCustomizedPageStatus operation is called successfully with right response returned,
            // so MS-WEBSS_R257 can be captured.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                257,
                @"[In GetCustomizedPageStatus]  This[GetCustomizedPageStatus] operation is defined as follows:
<wsdl:operation name=""GetCustomizedPageStatus"">
    <wsdl:input message=""tns:GetCustomizedPageStatusSoapIn"" />
    <wsdl:output message=""tns:GetCustomizedPageStatusSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R259
            Site.CaptureRequirementIfIsNotNull(
                customizedPageStatus,
                259,
                @"[In GetCustomizedPageStatus] The protocol server responds by sending a GetCustomizedPageStatusSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R260
            bool isContainInCustomizedPageStatus = (customizedPageStatus == CustomizedPageStatus.Customized)
                || (customizedPageStatus == CustomizedPageStatus.None)
                || (customizedPageStatus == CustomizedPageStatus.Uncustomized);
            Site.Assert.IsTrue(isContainInCustomizedPageStatus, "The value of the CustomizedPageStatus is {0}", customizedPageStatus);
            Site.CaptureRequirementIfIsTrue(
                isContainInCustomizedPageStatus,
                260,
                @"[In GetCustomizedPageStatus] The response specifies customization status of the page or file identified by the fileUrl, where the customization status MUST be one of the following: None, Customized, Uncustomized");

            // Verify MS-WEBSS requirement: MS-WEBSS_R276
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                276,
                @"[In GetCustomizedPageStatusSoapOut] The SOAP body contains a GetCustomizedPageStatusResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R715
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                715,
                @"[In GetCustomizedPageStatusResponse] The definition of the GetCustomizedPageStatusResponse element is as follows:
<s:element name=""GetCustomizedPageStatusResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""GetCustomizedPageStatusResult"" type=""tns:CustomizedPageStatus""/>
    </s:sequence>
  </s:complexType>
</s:element>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R279
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                279,
                @"[In GetCustomizedPageStatusResponse] GetCustomizedPageStatusResult: A single element of type only as defined in CustomizedPageStatus (section 3.1.4.9.4.1).");

            // Verify MS-WEBSS requirement: MS-WEBSS_R923
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                923,
                @"[In CustomizedPageStatus] The CustomizedPageStatus type is an enumeration of three possible values defined as follows:
<s:simpleType name=""CustomizedPageStatus"">
  <s:restriction base=""s:string"">
    <s:enumeration value=""None""/>
    <s:enumeration value=""Uncustomized""/>
    <s:enumeration value=""Customized""/>
  </s:restriction>
</s:simpleType>");
        }

        /// <summary>
        /// Capture GetListTemplates operation related requirements.
        /// </summary>
        private void ValidateGetListTemplates()
        {
            // The GetCustomizedPageStatus operation is called successfully with right response returned,
            // so MS-WEBSS_R298 can be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R298
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                298,
                @"[In GetListTemplates] This operation[GetListTemplates] is defined as follows: 
<wsdl:operation name=""GetListTemplates"">
    <wsdl:input message=""tns:GetListTemplatesSoapIn"" />
    <wsdl:output message=""tns:GetListTemplatesSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R299
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
               ValidationResult.Success,
               SchemaValidation.ValidationResult,
                299,
                @"[In GetListTemplates] [The protocol client sends a GetListTemplatesSoapIn request message] the protocol server responds with a GetListTemplatesSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R305
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
               ValidationResult.Success,
               SchemaValidation.ValidationResult,
                305,
                @"[In GetListTemplatesSoapOut] The SOAP body contains a GetListTemplatesResponse element as specified in section 3.1.4.10.2.2.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R307
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
               ValidationResult.Success,
               SchemaValidation.ValidationResult,
                307,
                @"[In GetListTemplatesResponse] The SOAP body contains a GetListTemplatesResponse element, which has the following definition:
<s:element name=""GetListTemplatesResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""GetListTemplatesResult"" minOccurs=""0"">
        <s:complexType mixed=""true"">
          <s:sequence>
            <s:element name=""ListTemplates"" type=""core:ListTemplateDefinitions"" minOccurs=""1"" >
            </s:element>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R308
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                308,
                @"[In GetListTemplatesResponse] GetListTemplatesResult: Contains a ListTemplates element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R309
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                309,
                @"[In GetListTemplatesResponse]  ListTemplates: An element of type ListTemplateDefinitions as specified in [MS-WSSCAML] section 2.3.2.12.[<xs:complexType name=""ListTemplateDefinitions"" mixed=""true"">
  <xs:sequence>
    <xs:element name=""ListTemplate"" type=""ListTemplateDefinition"" minOccurs=""0"" maxOccurs=""unbounded"" />
  </xs:sequence>
</xs:complexType> ].");
        }

        /// <summary>
        /// Capture GetObjectIdFromUrl operation related requirements.
        /// </summary>
        /// <param name="getObjectIdFromUrlResult">The Result of GetObjectIdFromUrl.</param>
        private void ValidateGetObjectIdFromUrl(GetObjectIdFromUrlResponseGetObjectIdFromUrlResult getObjectIdFromUrlResult)
        {
            if (!string.IsNullOrEmpty(getObjectIdFromUrlResult.ObjectId.ListId))
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R311
                // The operation is called successfully with the right response returned,
                // so the requirement can be captured.
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                    311,
                    @"[In GetObjectIdFromUrl] This[GetObjectIdFromUrl] operation is defined as follows:
<wsdl:operation name=""GetObjectIdFromUrl"">
    <wsdl:input message=""tns:GetObjectIdFromUrlSoapIn"" />
    <wsdl:output message=""tns:GetObjectIdFromUrlSoapOut"" />
</wsdl:operation>");

                // Verify MS-WEBSS requirement: MS-WEBSS_R312
                // The operation is called successfully with the right response returned,
                // so the requirement can be captured.
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                    312,
                    @"[In GetObjectIdFromUrl] [The protocol client sends a GetObjectIdFromUrlSoapIn request message] the protocol server responds with a GetObjectIdFromUrlSoapOut response message as follows:");

                // Verify MS-WEBSS requirement: MS-WEBSS_R320
                bool hasResponse = AdapterHelper.ElementExists(SchemaValidation.LastRawResponseXml, "GetObjectIdFromUrlResponse");
                Site.CaptureRequirementIfIsTrue(
                    hasResponse,
                    320,
                    @"[In GetObjectIdFromUrlSoapOut] The SOAP body contains a GetObjectIdFromUrlResponse element.");

                // Verify MS-WEBSS requirement: MS-WEBSS_R325
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                    325,
                    @"[In GetObjectIdFromUrlResponse] This element represents the result data of a GetObjectIdFromUrl operation. This element is defined as follows:
<s:element name=""GetObjectIdFromUrlResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""GetObjectIdFromUrlResult"" minOccurs=""1"" maxOccurs=""1"">
        <s:complexType mixed=""true"">
          <s:sequence>
            <s:element name=""ObjectId"" minOccurs=""1"" maxOccurs=""1"">
              <s:complexType>
                <s:attribute name=""ListId"" type=""core:UniqueIdentifierWithBracesOrEmpty"" />
                <s:attribute name=""ListServerTemplate"">
                  <s:simpleType>
                    <s:restriction base=""s:int"">
                      <s:enumeration value=""0""/>
                      <s:enumeration value=""100""/>
                      <s:enumeration value=""101""/>
                      <s:enumeration value=""102""/>
                      <s:enumeration value=""103""/>
                      <s:enumeration value=""104""/>
                      <s:enumeration value=""105""/>
                      <s:enumeration value=""106""/>
                      <s:enumeration value=""107""/>
                      <s:enumeration value=""108""/>
                      <s:enumeration value=""109""/>
                      <s:enumeration value=""110""/>
                      <s:enumeration value=""111""/>
                      <s:enumeration value=""112""/>
                      <s:enumeration value=""113""/>
                      <s:enumeration value=""114""/>
                      <s:enumeration value=""115""/>
                      <s:enumeration value=""116""/>
                      <s:enumeration value=""117""/>
                      <s:enumeration value=""118""/>
                      <s:enumeration value=""119""/>
                      <s:enumeration value=""120""/>
                      <s:enumeration value=""121""/>
                      <s:enumeration value=""122""/>
                      <s:enumeration value=""123""/>
                      <s:enumeration value=""130""/>
                      <s:enumeration value=""140""/>
                      <s:enumeration value=""150""/>
                      <s:enumeration value=""200""/>
                      <s:enumeration value=""201""/>
                      <s:enumeration value=""202""/>
                      <s:enumeration value=""204""/>
                      <s:enumeration value=""207""/>
                      <s:enumeration value=""210""/>
                      <s:enumeration value=""211""/>
                      <s:enumeration value=""212""/>
                      <s:enumeration value=""301""/>
                      <s:enumeration value=""302""/>
                      <s:enumeration value=""303""/>
                      <s:enumeration value=""402""/>
                      <s:enumeration value=""403""/>
                      <s:enumeration value=""404""/>
                      <s:enumeration value=""405""/>
                      <s:enumeration value=""420""/>
                      <s:enumeration value=""421""/>
                      <s:enumeration value=""499""/>
                      <s:enumeration value=""851""/>
                      <s:enumeration value=""1100""/>
                      <s:enumeration value=""1200""/>
                      <s:enumeration value=""1220""/>
                      <s:enumeration value=""1221""/>
                    </s:restriction>
                  </s:simpleType>
                </s:attribute>
                <s:attribute name=""ListBaseType"">
                  <s:simpleType>
                    <s:restriction base=""s:int"">
                      <s:enumeration value=""0""/>
                      <s:enumeration value=""1""/>
                      <s:enumeration value=""2""/>
                      <s:enumeration value=""3""/>
                      <s:enumeration value=""4""/>
                      <s:enumeration value=""5""/>
                    </s:restriction>
                  </s:simpleType>
                </s:attribute>
                <s:attribute name=""ListItem"" type=""core:TRUEFALSE"" />
                <s:attribute name=""ListItemId"" type=""s:string"" />
                <s:attribute name=""File"" type=""core:TRUEFALSE"" />
                <s:attribute name=""Folder"" type=""core:TRUEFALSE"" />
                <s:attribute name=""AlternateUrls"" type=""s:string"" />
              </s:complexType>
            </s:element>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");

                // Verify MS-WEBSS requirement: MS-WEBSS_R326
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                326,
                @"[In GetObjectIdFromUrlResponse] GetObjectIdFromUrlResult: If no error conditions as specified earlier cause the protocol server to return a SOAP exception, a GetObjectIdFromUrlResult MUST be returned.");

                // Verify MS-WEBSS requirement: MS-WEBSS_R327
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                327,
                @"[In GetObjectIdFromUrlResponse] ObjectId: The container element for the object properties.");

                // Verify MS-WEBSS requirement: MS-WEBSS_R1045
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                    1045,
                    @"[In GetObjectIdFromUrlResponse] ObjectId.ListServerTemplate: If the object is a list, the value of the attribute MUST be one of the list template types as specified in [MS-WSSFO2] section 2.2.3.12 [the values of the list template types are -1,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,130,140,150,200,201,202,204,207,210,211,212,301,302,303,402,403,404,405,420,421,499,1100,1200,1220,1221].");

                // Verify MS-WEBSS requirement: MS-WEBSS_R1045001002
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                    1045001002,
                    @"[In GetObjectIdFromUrlResponse] ObjectId.ListServerTemplate: If the object is a list item, the value of the attribute MUST be one of the list template types as specified in [MS-WSSFO2] section 2.2.3.12 [the values of the list template types are -1,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,130,140,150,200,201,202,204,207,210,211,212,301,302,303,402,403,404,405,420,421,499,1100,1200,1220,1221].");

                // Verify MS-WEBSS requirement: MS-WEBSS_R1046
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                    1046,
                    @"[In GetObjectIdFromUrlResponse] ObjectId.ListBaseType: If the object is a list, the value of the attribute MUST be one of the List Base Types as specified in [MS-WSSFO2] section 2.2.3.11 [the values of the List Base Types are 0,1,3,4,5].");
                
                // Verify MS-WEBSS requirement: MS-WEBSS_R1046
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                    1046001002,
                    @"[In GetObjectIdFromUrlResponse] ObjectId.ListBaseType: If the object is a list item, the value of the attribute MUST be one of the List Base Types as specified in [MS-WSSFO2] section 2.2.3.11 [the values of the List Base Types are 0,1,3,4,5].");
            }
        }

        /// <summary>
        /// This method is used to capture requirements of GetWeb operation.
        /// </summary>
        /// <param name="getWebResult">The result of the operation.</param>
        private void ValidateGetWeb(GetWebResponseGetWebResult getWebResult)
        {
            // Proxy handles operation's SoapIn and SoapOut, if the server didn't respond the SoapOut message for the operation, Proxy will fail. 
            // Proxy didn't fail now, that indicates server responds corresponding SoapOut message for the operation.So the requirement be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R344
            if (Common.IsRequirementEnabled(691, this.Site))
            {
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                    344,
                    @"[In GetWeb] This[GetWeb] operation is defined as follows:
<wsdl:operation name=""GetWeb"">
    <wsdl:input message=""tns:GetWebSoapIn"" />
    <wsdl:output message=""tns:GetWebSoapOut"" />
</wsdl:operation>");

                // Verify MS-WEBSS requirement: MS-WEBSS_R345
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                    345,
                    @"[In GetWeb] The protocol client sends a GetWebSoapIn request message, and the protocol server responds with a GetWebSoapOut response message.");
            }
            // Verify MS-WEBSS requirement: MS-WEBSS_R348
            // COMMENT: If the actual Language value is contained in the expected domain of allowed 
            // LCID values, then the requirement can be captured.
            // Ensure the SOAP result is returned successfully.
            Site.Assume.IsNotNull(getWebResult, "The result of GetWeb operation must not be null.");
            string[] lcidValues = 
            {
                                      ((uint)LCID_Values.Afrikaans).ToString(), 
                                      ((uint)LCID_Values.Albanian).ToString(), 
                                      ((uint)LCID_Values.Arabic_Algeria).ToString(), 
                                      ((uint)LCID_Values.Arabic_Bahrain).ToString(), 
                                      ((uint)LCID_Values.Arabic_Egypt).ToString(), 
                                      ((uint)LCID_Values.Arabic_Iraq).ToString(), 
                                      ((uint)LCID_Values.Arabic_Jordan).ToString(), 
                                      ((uint)LCID_Values.Arabic_Kuwait).ToString(), 
                                      ((uint)LCID_Values.Arabic_Lebanon).ToString(), 
                                      ((uint)LCID_Values.Arabic_Libya).ToString(), 
                                      ((uint)LCID_Values.Arabic_Morocco).ToString(), 
                                      ((uint)LCID_Values.Arabic_Oman).ToString(), 
                                      ((uint)LCID_Values.Arabic_Qatar).ToString(), 
                                      ((uint)LCID_Values.Arabic_Saudi_Arabia).ToString(), 
                                      ((uint)LCID_Values.Arabic_Syria).ToString(), 
                                      ((uint)LCID_Values.Arabic_Tunisia).ToString(), 
                                      ((uint)LCID_Values.Arabic_United_Arab_Emirates).ToString(), 
                                      ((uint)LCID_Values.Arabic_Yemen).ToString(), 
                                      ((uint)LCID_Values.Armenian).ToString(), 
                                      ((uint)LCID_Values.Azeri_Cyrillic).ToString(), 
                                      ((uint)LCID_Values.Azeri_Latin).ToString(), 
                                      ((uint)LCID_Values.Basque).ToString(), 
                                      ((uint)LCID_Values.Belarusian).ToString(), 
                                      ((uint)LCID_Values.Bulgarian).ToString(), 
                                      ((uint)LCID_Values.Catalan).ToString(), 
                                      ((uint)LCID_Values.Chinese_China).ToString(), 
                                      ((uint)LCID_Values.Chinese_Hong_Kong_SAR).ToString(), 
                                      ((uint)LCID_Values.Chinese_Macau_SAR).ToString(), 
                                      ((uint)LCID_Values.Chinese_Singapore).ToString(), 
                                      ((uint)LCID_Values.Chinese_Taiwan).ToString(), 
                                      ((uint)LCID_Values.Croatian).ToString(), 
                                      ((uint)LCID_Values.Czech).ToString(), 
                                      ((uint)LCID_Values.Danish).ToString(), 
                                      ((uint)LCID_Values.Dutch_Belgium).ToString(), 
                                      ((uint)LCID_Values.Dutch_Netherlands).ToString(), 
                                      ((uint)LCID_Values.English_Australia).ToString(), 
                                      ((uint)LCID_Values.English_Belize).ToString(), 
                                      ((uint)LCID_Values.English_Canada).ToString(), 
                                      ((uint)LCID_Values.English_Caribbean).ToString(), 
                                      ((uint)LCID_Values.English_Great_Britain).ToString(), 
                                      ((uint)LCID_Values.English_Ireland).ToString(), 
                                      ((uint)LCID_Values.English_Jamaica).ToString(), 
                                      ((uint)LCID_Values.English_New_Zealand).ToString(), 
                                      ((uint)LCID_Values.English_Philippines).ToString(), 
                                      ((uint)LCID_Values.English_Southern_Africa).ToString(), 
                                      ((uint)LCID_Values.English_Trinidad).ToString(), 
                                      ((uint)LCID_Values.English_United_States).ToString(), 
                                      ((uint)LCID_Values.Estonian).ToString(), 
                                      ((uint)LCID_Values.Faroese).ToString(), 
                                      ((uint)LCID_Values.Farsi).ToString(), 
                                      ((uint)LCID_Values.Finnish).ToString(), 
                                      ((uint)LCID_Values.French_Belgium).ToString(), 
                                      ((uint)LCID_Values.French_Canada).ToString(), 
                                      ((uint)LCID_Values.French_France).ToString(), 
                                      ((uint)LCID_Values.French_Luxembourg).ToString(), 
                                      ((uint)LCID_Values.French_Switzerland).ToString(), 
                                      ((uint)LCID_Values.FYRO_Macedonia).ToString(), 
                                      ((uint)LCID_Values.Gaelic_Ireland).ToString(), 
                                      ((uint)LCID_Values.Gaelic_Scotland).ToString(), 
                                      ((uint)LCID_Values.German_Austria).ToString(), 
                                      ((uint)LCID_Values.German_Germany).ToString(), 
                                      ((uint)LCID_Values.German_Liechtenstein).ToString(), 
                                      ((uint)LCID_Values.German_Luxembourg).ToString(), 
                                      ((uint)LCID_Values.German_Switzerland).ToString(), 
                                      ((uint)LCID_Values.Greek).ToString(), 
                                      ((uint)LCID_Values.Hebrew).ToString(), 
                                      ((uint)LCID_Values.Hindi).ToString(), 
                                      ((uint)LCID_Values.Hungarian).ToString(), 
                                      ((uint)LCID_Values.Icelandic).ToString(), 
                                      ((uint)LCID_Values.Indonesian).ToString(), 
                                      ((uint)LCID_Values.Italian_Italy).ToString(), 
                                      ((uint)LCID_Values.Italian_Switzerland).ToString(), 
                                      ((uint)LCID_Values.Japanese).ToString(), 
                                      ((uint)LCID_Values.Korean).ToString(), 
                                      ((uint)LCID_Values.Latvian).ToString(), 
                                      ((uint)LCID_Values.Lithuanian).ToString(), 
                                      ((uint)LCID_Values.Malay_Brunei).ToString(), 
                                      ((uint)LCID_Values.Malay_Malaysia).ToString(), 
                                      ((uint)LCID_Values.Maltese).ToString(), 
                                      ((uint)LCID_Values.Marathi).ToString(), 
                                      ((uint)LCID_Values.Norwegian_Bokml).ToString(), 
                                      ((uint)LCID_Values.Norwegian_Nynorsk).ToString(), 
                                      ((uint)LCID_Values.Polish).ToString(), 
                                      ((uint)LCID_Values.Portuguese_Brazil).ToString(), 
                                      ((uint)LCID_Values.Portuguese_Portugal).ToString(), 
                                      ((uint)LCID_Values.Raeto_Romance).ToString(), 
                                      ((uint)LCID_Values.Romanian_Republic_of_Moldova).ToString(), 
                                      ((uint)LCID_Values.Romanian_Romania).ToString(), 
                                      ((uint)LCID_Values.Russian).ToString(), 
                                      ((uint)LCID_Values.Russian_Republic_of_Moldova).ToString(), 
                                      ((uint)LCID_Values.Sanskrit).ToString(), 
                                      ((uint)LCID_Values.Serbian_Cyrillic).ToString(), 
                                      ((uint)LCID_Values.Serbian_Latin).ToString(), 
                                      ((uint)LCID_Values.Setswana).ToString(), 
                                      ((uint)LCID_Values.Slovak).ToString(), 
                                      ((uint)LCID_Values.Slovenian).ToString(), 
                                      ((uint)LCID_Values.Sorbian).ToString(), 
                                      ((uint)LCID_Values.Southern_Sotho).ToString(), 
                                      ((uint)LCID_Values.Spanish_Argentina).ToString(), 
                                      ((uint)LCID_Values.Spanish_Bolivia).ToString(), 
                                      ((uint)LCID_Values.Spanish_Chile).ToString(), 
                                      ((uint)LCID_Values.Spanish_Colombia).ToString(), 
                                      ((uint)LCID_Values.Spanish_Costa_Rica).ToString(), 
                                      ((uint)LCID_Values.Spanish_Dominican_Republic).ToString(), 
                                      ((uint)LCID_Values.Spanish_Ecuador).ToString(), 
                                      ((uint)LCID_Values.Spanish_El_Salvador).ToString(), 
                                      ((uint)LCID_Values.Spanish_Guatemala).ToString(), 
                                      ((uint)LCID_Values.Spanish_Honduras).ToString(), 
                                      ((uint)LCID_Values.Spanish_Mexico).ToString(), 
                                      ((uint)LCID_Values.Spanish_Nicaragua).ToString(), 
                                      ((uint)LCID_Values.Spanish_Panama).ToString(), 
                                      ((uint)LCID_Values.Spanish_Paraguay).ToString(), 
                                      ((uint)LCID_Values.Spanish_Peru).ToString(), 
                                      ((uint)LCID_Values.Spanish_Puerto_Rico).ToString(), 
                                      ((uint)LCID_Values.Spanish_Spain_Traditional).ToString(), 
                                      ((uint)LCID_Values.Spanish_Uruguay).ToString(), 
                                      ((uint)LCID_Values.Spanish_Venezuela).ToString(), 
                                      ((uint)LCID_Values.Swahili).ToString(), 
                                      ((uint)LCID_Values.Swedish_Finland).ToString(), 
                                      ((uint)LCID_Values.Swedish_Sweden).ToString(), 
                                      ((uint)LCID_Values.Tamil).ToString(), 
                                      ((uint)LCID_Values.Tatar).ToString(), 
                                      ((uint)LCID_Values.Thai).ToString(), 
                                      ((uint)LCID_Values.Tsonga).ToString(), 
                                      ((uint)LCID_Values.Turkish).ToString(), 
                                      ((uint)LCID_Values.Ukrainian).ToString(), 
                                      ((uint)LCID_Values.Urdu).ToString(), 
                                      ((uint)LCID_Values.Uzbek_Cyrillic).ToString(), 
                                      ((uint)LCID_Values.Uzbek_Latin).ToString(), 
                                      ((uint)LCID_Values.Vietnamese).ToString(), 
                                      ((uint)LCID_Values.Xhosa).ToString(), 
                                      ((uint)LCID_Values.Yiddish).ToString(), 
                                      ((uint)LCID_Values.Zulu).ToString() 
            };

            bool isVerifiedR348 = false;
            List<string> lcidDomain = new List<string>(lcidValues);
            if (lcidDomain.Contains(getWebResult.Web.Language))
            {
                isVerifiedR348 = true;
            }
            else
            {
                Site.Assert.Fail("lcidDomain doesn't contain web language {0}", getWebResult.Web.Language);
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR348,
                348,
                @"[In GetWeb] The Language property MUST include an LCID value.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R355
            Site.CaptureRequirementIfIsTrue(
                this.HasElement(SchemaValidation.LastRawResponseXml, "GetWebResponse"),
                355,
                @"[In GetWebSoapOut] The SOAP body contains a GetWebResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R360
            if (Common.IsRequirementEnabled(691, this.Site))
            {
                Site.CaptureRequirementIfAreEqual<ValidationResult>(
                    ValidationResult.Success,
                    SchemaValidation.ValidationResult,
                    360,
                    @"[In GetWebResponse] The SOAP body contains a GetWebResponse element, which has the following definition:
<s:element name=""GetWebResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""GetWebResult"" minOccurs=""0"">
        <s:complexType>
          <s:sequence>
            <s:element name=""Web"" type=""tns:WebDefinition""/>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");
            }
        }

        /// <summary>
        /// Capture GetWebCollection operation related requirements.
        /// </summary>
        private void ValidateGetWebCollection()
        {
            // The GetWebCollection operation is called successfully with right response returned,
            // so MS-WEBSS_R366 can be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R366
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                366,
                @"[In GetWebCollection] [This operation GetWebCollection  is defined as follows]<wsdl:operation name=""GetWebCollection"">
    <wsdl:input message=""tns:GetWebCollectionSoapIn"" />
    <wsdl:output message=""tns:GetWebCollectionSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R367
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                367,
                @"[In GetWebCollection] [The protocol client sends a GetWebCollectionSoapIn request message] the protocol server responds with a GetWebCollectionSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R375
            Site.CaptureRequirementIfIsTrue(
                this.HasElement(SchemaValidation.LastRawResponseXml, "GetWebCollectionResponse"),
                375,
                @"[In GetWebCollectionSoapOut] The SOAP body contains a GetWebCollectionResponse XML element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R377
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                377,
                @"[In GetWebCollectionResponse] The SOAP body contains a GetWebCollectionResponse element, which has the following definition:
<s:element name=""GetWebCollectionResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""GetWebCollectionResult"" minOccurs=""0"">
        <s:complexType>
          <s:sequence>
            <s:element name=""Webs"">
              <s:complexType>
                <s:sequence>
                  <s:element name=""Web"" type=""tns:WebDefinition"" minOccurs=""0"" maxOccurs=""unbounded"">
                  </s:element>
                </s:sequence>
              </s:complexType>
            </s:element>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R378
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                378,
                @"[In GetWebCollectionResponse] GetWebCollectionResult: This element contains Webs element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R379
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                379,
                @"[In GetWebCollectionResponse] Webs: This element is a collection of Web elements.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R380
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                380,
                @"[In GetWebCollectionResponse] Web: The structure of the Web element is defined by the WebDefinition complex type.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R381
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                381,
                @"[In GetWebCollectionResponse] The Web element contains only Title and Url attributes of WebDefinition.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R382
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                382,
                @"[In GetWebCollectionResponse] The collection of Web elements included in a Webs element contains all immediate child sites of the context site.");
        }

        /// <summary>
        /// Capture RemoveContentTypeXmlDocument related requirements.
        /// </summary>
        private void ValidateRemoveContentTypeXmlDocument()
        {
            // Proxy handles operation's SoapIn and SoapOut, if the server didn't respond the soapout message for the operation, Proxy will fail. 
            // Proxy didn't fail now, that indicates server responds corresponding soapout message for the operation.So the requirement be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R384
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                384,
                @"[In RemoveContentTypeXmlDocument] This operation[RemoveContentTypeXmlDocument] is defined as follows:
<wsdl:operation name=""RemoveContentTypeXmlDocument"">
    <wsdl:input message=""tns:RemoveContentTypeXmlDocumentSoapIn"" />
    <wsdl:output message=""tns:RemoveContentTypeXmlDocumentSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R385
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                385,
                @"[In RemoveContentTypeXmlDocument] [The protocol client sends a RemoveContentTypeXmlDocumentSoapIn request message] the protocol server responds with a RemoveContentTypeXmlDocumentSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R394
            Site.CaptureRequirementIfIsTrue(
                this.HasElement(SchemaValidation.LastRawResponseXml, "RemoveContentTypeXmlDocumentResponse"),
                394,
                @"[In RemoveContentTypeXmlDocumentSoapOut] The SOAP body contains a RemoveContentTypeXmlDocumentResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R400
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                400,
                @"[In RemoveContentTypeXmlDocumentResponse] This[RemoveContentTypeXmlDocumentResponse] element is defined as follows:
<s:element name=""RemoveContentTypeXmlDocumentResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""RemoveContentTypeXmlDocumentResult"" minOccurs=""0"">
        <s:complexType>
          <s:sequence>
            <s:element name=""Success"" minOccurs=""1"" maxOccurs=""1"">
              <s:complexType />
            </s:element>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");
        }

        /// <summary>
        /// Capture RevertAllFileContentStreams operation related requirements.
        /// </summary>
        private void ValidateRevertAllFileContentStreams()
        {
            // Verify MS-WEBSS requirement: MS-WEBSS_R404.
            // The operation is called successfully with the right response returned,
            // so the requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                404,
                @"[In RevertAllFileContentStreams] This operation[RevertAllFileContentStreams] is defined as follows:
<wsdl:operation name=""RevertAllFileContentStreams"">
    <wsdl:input message=""tns:RevertAllFileContentStreamsSoapIn"" />
    <wsdl:output message=""tns:RevertAllFileContentStreamsSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R406
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                406,
                @"[In RevertAllFileContentStreams] If the operation succeeds, the protocol server MUST return a RevertAllFileContentStreamsReponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R413
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.ElementExists(SchemaValidation.LastRawResponseXml, "RevertAllFileContentStreamsResponse"),
                413,
                "[In RevertAllFileContentStreamsSoapOut] The SOAP body contains a RevertAllFileContentStreamsResponse element.");

            // The operation is called successfully with the right response returned,
            // so the requirement can be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R405
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                405,
                @"[In RevertAllFileContentStreams] [The protocol client sends a RevertAllFileContentStreamsSoapIn request message] the protocol server responds with a RevertAllFileContentStreamsSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R415
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                415,
                @"[In RevertAllFileContentStreamsResponse] The SOAP body contains a RevertAllFileContentStreamsResponse element, which has the following definition:
<s:element name=""RevertAllFileContentStreamsResponse"">
  <s:complexType/>
</s:element>");
        }

        /// <summary>
        /// This method is used to capture requirements of RevertCss operation.
        /// </summary>
        private void ValidateRevertCss()
        {
            // Proxy handles operation's SoapIn and SoapOut, if the server didn't respond the SoapOut message for the operation, Proxy will fail. 
            // Proxy didn't fail now, that indicates server responds corresponding SoapOut message for the operation.So the requirement be captured.

            // Verify MS-WEBSS requirement: MS-WEBSS_R417
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                417,
                 @"[In RevertCss] This operation[RevertCss] is defined as follows:
<wsdl:operation name=""RevertCss"">
    <wsdl:input message=""tns:RevertCssSoapIn"" />
    <wsdl:output message=""tns:RevertCssSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R425
            Site.CaptureRequirementIfIsTrue(
                this.HasElement(SchemaValidation.LastRawResponseXml, "RevertCssResponse"),
                425,
                @"[In RevertCssSoapOut] The SOAP body contains a RevertCssResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R430
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                430,
                @"[In RevertCssResponse] The definition of the RevertCssResponse element is as follows:
<s:element name=""RevertCssResponse"">
  <s:complexType/>
</s:element>");
        }

        /// <summary>
        /// Capture RevertFileContentStreams operation related requirements.
        /// </summary>
        private void ValidateRevertFileContentStream()
        {
            // Verify MS-WEBSS requirement: MS-WEBSS_R435.
            // The operation is called successfully with the right response returned,
            // so the requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                435,
                @"[In RevertFileContentStream] This operation[RevertFileContentStream] is defined as follows:
<wsdl:operation name=""RevertFileContentStream"">
    <wsdl:input message=""tns:RevertFileContentStreamSoapIn"" />
    <wsdl:output message=""tns:RevertFileContentStreamSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R436
            // The operation is called successfully with the right response returned,
            // so the requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                436,
                @"[In RevertFileContentStream] [The protocol client sends a RevertFileContentStreamSoapIn request message] the protocol server responds with a RevertFileContentStreamSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R438
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.ElementExists(SchemaValidation.LastRawResponseXml, "RevertFileContentStreamResponse"),
                438,
                @"[In RevertFileContentStream] If the operation succeeds, it MUST return a RevertFileContentStreamResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R446
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.ElementExists(SchemaValidation.LastRawResponseXml, "RevertFileContentStreamResponse"),
                446,
                "[In RevertFileContentStreamSoapOut] The SOAP body contains a RevertFileContentStreamResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R450
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                450,
                @"[In RevertFileContentStreamResponse] The SOAP body contains a RevertFileContentStreamResponse element, which has the following definition:
<s:element name=""RevertFileContentStreamResponse"">
  <s:complexType/>
</s:element>");
        }

        /// <summary>
        /// Capture UpdateColumns operation related requirements.
        /// </summary>
        private void ValidateUpdateColumns()
        {
            // Verify MS-WEBSS requirement: MS-WEBSS_R463.
            // The operation is called successfully with the right response returned,
            // so the requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                463,
                @"[In UpdateColumns] This operation[UpdateColumns] is defined as follows:
<wsdl:operation name=""UpdateColumns"">
    <wsdl:input message=""tns:UpdateColumnsSoapIn"" />
    <wsdl:output message=""tns:UpdateColumnsSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R728
            // The operation is called successfully with the right response returned,
            // so the requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                728,
                @"[In UpdateColumns] [The protocol client sends an UpdateColumnsSoapIn request message] the protocol server responds with an UpdateColumnsSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R477
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.ElementExists(SchemaValidation.LastRawResponseXml, "UpdateColumnsResponse"),
                477,
                @"[In UpdateColumnsSoapOut] The SOAP body contains an UpdateColumnsResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R500
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                500,
                @"[In UpdateColumns] It [Method: An element that represents a field to be used in the new, update, or delete operation. ] MUST contain an ID attribute.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R501
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                501,
                @"[In UpdateColumnsResponse] The SOAP body contains an UpdateColumnsResponse element, which has the following definition:
<s:element name=""UpdateColumnsResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""UpdateColumnsResult"" minOccurs=""0"">
        <s:complexType>
          <s:sequence>
            <s:element name=""Results"">
              <s:complexType>
                <s:sequence>
                  <s:element name=""NewFields"">
                    <s:complexType>
                      <s:sequence>
                        <s:element name=""Method"" minOccurs=""0"" maxOccurs=""unbounded"" >
                          <s:attribute name=""ID"" type=""s:string"" use=""required""/> 
                          <s:complexType>
                            <s:sequence>
                              <s:element name=""ErrorCode"" type=""s:string"" minOccurs=""1"" maxOccurs=""1""/>
                              <s:element name=""ErrorText"" type=""s:string"" minOccurs=""0"" maxOccurs=""1""/>
                              <s:element name=""Field"" minOccurs=""0"" maxOccurs=""1"" type=""core:FieldDefinition""/>
                            </s:sequence>
                          </s:complexType>
                        </s:element>        
                      </s:sequence>
                    </s:complexType>
                  </s:element>
                  <s:element name=""UpdateFields"">
                    <s:complexType>
                      <s:sequence>
                        <s:element name=""Method"" minOccurs=""0"" maxOccurs=""unbounded"" >
                          <s:attribute name=""ID"" type=""s:string"" use=""required""/> 
                          <s:complexType>
                            <s:sequence>
                              <s:element name=""ErrorCode"" type=""s:string"" minOccurs=""1"" maxOccurs=""1""/>
                              <s:element name=""ErrorText"" type=""s:string"" minOccurs=""0"" maxOccurs=""1""/>
                              <s:element name=""Field"" minOccurs=""0"" maxOccurs=""1"" type=""core:FieldDefinition""/>          
                            </s:sequence>
                          </s:complexType>
                        </s:element>        
                      </s:sequence>
                    </s:complexType>
                  </s:element>
                  <s:element name=""DeleteFields"">
                    <s:complexType>
                      <s:sequence>
                        <s:element name=""Method"" minOccurs=""0"" maxOccurs=""unbounded"" >
                          <s:attribute name=""ID"" type=""s:string"" use=""required""/> 
                          <s:complexType>
                            <s:sequence>
                              <s:element name=""ErrorCode"" type=""s:string"" minOccurs=""1"" maxOccurs=""1""/>
                              <s:element name=""ErrorText"" type=""s:string"" minOccurs=""0"" maxOccurs=""1""/>
                              <s:element name=""Field"" minOccurs=""0"" maxOccurs=""1"" type=""core:FieldDefinition""/>       
                            </s:sequence>
                          </s:complexType>
                        </s:element>        
                      </s:sequence>
                    </s:complexType>
                  </s:element>
                </s:sequence>
              </s:complexType>
            </s:element>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R502
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                502,
                @"[In UpdateColumnsResponse] UpdateColumnsResult: This element contains a Results element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R503
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
               ValidationResult.Success,
               SchemaValidation.ValidationResult,
                503,
                @"[In UpdateColumnsResponse] Results: This element contains NewFields, UpdateFields, and DeleteFields elements.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R507
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
               ValidationResult.Success,
               SchemaValidation.ValidationResult,
                507,
                @"[In UpdateColumnsResponse] Method: This element contains an ErrorText, an ErrorCode, and a Field element.");
        }

        /// <summary>
        /// Capture UpdateContentType related requirements.
        /// </summary>
        private void ValidateUpdateContentType()
        {
            // Proxy handles operation's SoapIn and SoapOut, if the server didn't respond the soapout message for the operation, Proxy will fail. 
            // Proxy didn't fail now, that indicates server responds corresponding soapout message for the operation.So the requirement be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R524
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                524,
                @"[In UpdateContentType] This operation[UpdateContentType] is defined as follows:
<wsdl:operation name=""UpdateContentType"">
    <wsdl:input message=""tns:UpdateContentTypeSoapIn"" />
    <wsdl:output message=""tns:UpdateContentTypeSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R525
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                525,
                @"[In UpdateContentType] [The protocol client sends an UpdateContentTypeSoapIn request message] the protocol server responds with an UpdateContentTypeSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R540
            Site.CaptureRequirementIfIsTrue(
                this.HasElement(SchemaValidation.LastRawResponseXml, "UpdateContentTypeResponse"),
                540,
                @"[In UpdateContentTypeSoapOut] The SOAP body contains an UpdateContentTypeResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R554
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                554,
                @"[In UpdateContentTypeResponse] This element[UpdateContentTypeResponse] is defined as follows:
<s:element name=""UpdateContentTypeResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""UpdateContentTypeResult"" minOccurs=""0"">
        <s:complexType>
          <s:sequence>
            <s:element name=""Results"" minOccurs=""1"" maxOccurs=""1"">
              <s:complexType>
                <s:sequence>
                  <s:element name=""Method"" minOccurs=""0"" maxOccurs=""unbounded"">
                    <s:complexType>
                      <s:sequence>
                        <s:element name=""ErrorCode"" type=""s:string"" minOccurs=""1"" maxOccurs=""1""/>
                        <s:element name=""FieldRef"" type=""FieldReferenceDefinitionCT"" minOccurs=""0"" maxOccurs=""1""/>
                        <s:element name=""ErrorText"" type=""s:string"" minOccurs=""0"" maxOccurs=""1""/>
                      </s:sequence>
                      <s:attribute name=""ID"" type=""s:string"" use=""required""/>
                    </s:complexType>
                  </s:element>
                  <s:element name=""ListProperties"" minOccurs=""1"" maxOccurs=""1"">
                    <s:complexType>
                      <s:sequence />
                      <s:attribute name=""Description"" type=""s:string"" use=""optional"" />
                      <s:attribute name=""FeatureId"" type=""core: UniqueIdentifierWithOrWithoutBraces"" use=""optional"" />
                      <s:attribute name=""Group"" type=""s:string"" use=""optional"" />
                      <s:attribute name=""Hidden"" type=""TRUEONLY"" use=""optional"" />
                      <s:attribute name=""ID"" type=""core:ContentTypeId"" use=""required"" />
                      <s:attribute name=""Locs"" type=""ONEONLY"" use=""optional"" />
                      <s:attribute name=""Name"" type=""s:string"" use=""optional"" />
                      <s:attribute name=""NewDocumentControl"" type=""s:string"" use=""optional"" />
                      <s:attribute name=""ReadOnly"" type=""TRUEONLY"" use=""optional"" />
                      <s:attribute name=""RequireClientRenderingOnNew"" type=""FALSEONLY"" use=""optional"" />
                      <s:attribute name=""Sealed"" type=""TRUEONLY"" use=""optional"" />
                      <s:attribute name=""Version"" type=""s:long"" use=""required"" />
                    </s:complexType>
                  </s:element>
                </s:sequence>
              </s:complexType>
            </s:element>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");
        }

        /// <summary>
        /// Capture UpdateContentTypeXmlDocument related requirements.
        /// </summary>
        private void ValidateUpdateContentTypeXmlDocument()
        {
            // Proxy handles operation's SoapIn and SoapOut, if the server didn't respond the soapout message for the operation, Proxy will fail. 
            // Proxy didn't fail now, that indicates server responds corresponding soapout message for the operation.So the requirement be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R586
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                586,
                @"[In UpdateContentTypeXmlDocument] This operation[UpdateContentTypeXmlDocument] is defined as follows:
<wsdl:operation name=""UpdateContentTypeXmlDocument"">
    <wsdl:input message=""tns:UpdateContentTypeXmlDocumentSoapIn"" />
    <wsdl:output message=""tns:UpdateContentTypeXmlDocumentSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R587
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                587,
                @"[In UpdateContentTypeXmlDocument] [The protocol client sends an UpdateContentTypeXmlDocumentSoapIn request message] the protocol server responds with an UpdateContentTypeXmlDocumentSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R597
            Site.CaptureRequirementIfIsTrue(
                this.HasElement(SchemaValidation.LastRawResponseXml, "UpdateContentTypeXmlDocumentResponse"),
                597,
                @"[In UpdateContentTypeXmlDocumentSoapOut] The SOAP body contains an UpdateContentTypeXmlDocumentResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R614
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                614,
                @"[In UpdateContentTypeXmlDocumentResponse] This element is defined as follows:
<s:element name=""UpdateContentTypeXmlDocumentResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""UpdateContentTypeXmlDocumentResult"" minOccurs=""0"" maxOccurs=""1"">
        <s:complexType mixed=""true"">
          <s:sequence>            
            <s:element name=""Success"" minOccurs=""0"" maxOccurs=""1"">
              <s:complexType />
            </s:element>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");
        }

        /// <summary>
        /// Capture WebUrlFromPageUrl operation related requirements.
        /// </summary>
        private void ValidateWebUrlFromPageUrl()
        {
            // Verify MS-WEBSS requirement: MS-WEBSS_R618
            // The operation is called successfully with the right response returned,
            // so the requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                618,
                @"[In WebUrlFromPageUrl] This operation[WebUrlFromPageUrl] is defined as follows:
<wsdl:operation name=""WebUrlFromPageUrl"">
    <wsdl:input message=""tns:WebUrlFromPageUrlSoapIn"" />
    <wsdl:output message=""tns:WebUrlFromPageUrlSoapOut"" />
</wsdl:operation>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R619
            // The operation is called successfully with the right response returned,
            // so the requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                619,
                @"[In WebUrlFromPageUrl] The protocol client sends a WebUrlFromPageUrlSoapIn request message, and the protocol server responds with a WebUrlFromPageUrlSoapOut response message.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R632
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                632,
                @"[In WebUrlFromPageUrlResponse]  This element[WebUrlFromPageUrlResponse] is defined as follows:
 <s:element name=""WebUrlFromPageUrlResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""WebUrlFromPageUrlResult"" type=""s:string"" minOccurs=""0""/>
    </s:sequence>
  </s:complexType>
</s:element>");

            // Verify MS-WEBSS requirement: MS-WEBSS_R627
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.ElementExists(SchemaValidation.LastRawResponseXml, "WebUrlFromPageUrlResponse"),
                627,
                @"[In WebUrlFromPageUrlSoapOut] The SOAP body contains a WebUrlFromPageUrlResponse element.");
        }

        private void ValidateCustomizeCss()
        {
            // Verify MS-WEBSS requirement: MS-WEBSS_R80
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                80,
                @"[In CustomizeCss] This operation[CustomizeCss] is defined as follows:
 <wsdl:operation name=""CustomizeCss"">
      < wsdl:input message = ""tns:CustomizeCssSoapIn"" />
      < wsdl:output message = ""tns:CustomizeCssSoapOut"" />
</ wsdl:operation > ");

            // Verify MS-WEBSS requirement: MS-WEBSS_R88
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.ElementExists(SchemaValidation.LastRawResponseXml, "CustomizeCssResponse"),
                88,
                @"[In CustomizeCssSoapOut] The SOAP body contains a CustomizeCssResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R94
            // The operation is called successfully with the right response returned,
            // so the requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                94,
                @"[In CustomizeCssResponse] On successful completion, the response SOAP body contains only the CustomizeCssResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R95
            // The operation is called successfully with the right response returned,
            // so the requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                95,
                @"[In CustomizeCssResponse] Other than the namespace attribute, the CustomizeCssResponse element contains no other attributes or child elements.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R96
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                96,
                @"[In CustomizeCssResponse] <s:element name=""CustomizeCssResponse"">
  < s:complexType />
</ s:element > ");
        }
        #region Help Method
        /// <summary>
        /// Verify if elementName is existed in the specified XmlElement..
        /// </summary>
        /// <param name="xmlElement">A stream of xml element.</param>
        /// <param name="elementName">The element name which need to check whether it is existed.</param>
        /// <returns>If the XML response has contain element, true means include, otherwise false.</returns>
        private bool HasElement(XmlElement xmlElement, string elementName)
        {
            // Verify whether elementName is existed.
            // If server response XML contains elementName, true will be returned. otherwise false will be returned.
            if (xmlElement.GetElementsByTagName(elementName).Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion
    }
}