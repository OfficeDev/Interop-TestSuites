namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    using System;
    using System.Text.RegularExpressions;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This partial class provides the methods to verify requirements.
    /// </summary>
    public partial class MS_SITESSAdapter
    {
        /// <summary>
        /// Verify whether the transport used by this test suite is HTTP or HTTPS, and the soap version is soap 1.1 or soap 1.2.
        /// </summary>
        private void VerifyCommonReqs()
        {
            switch (this.transportProtocol)
            {
                case TransportType.HTTP:

                    // If can get a successful response by using SOAP over HTTP. It means the server supports SOAP over HTTP.
                    // Verify requirement: MS-SITESS_R3
                    Site.CaptureRequirement(3, "[In Transport] Protocol servers MUST support SOAP over HTTP.");
                    break;
                case TransportType.HTTPS:

                    if (Common.IsRequirementEnabled(541, this.Site))
                    {
                        // Microsoft Office SharePoint Server 2007 and above Microsoft products do additionally support SOAP over HTTPS for securing communication with clients.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R541");

                        // If can get a successful response by using SOAP over HTTPS. It means the server supports SOAP over HTTPS.
                        // Verify MS-SITESS requirement: MS-SITESS_R541
                        Site.CaptureRequirement(
                            541,
                            @"[In Transport] Implementation does additionally support SOAP over HTTPS for securing communication with clients. (Microsoft Office SharePoint Server 2007 and above follow this behavior.)");
                    }

                    break;
                default:
                    Site.Assert.Fail("Transport: {0} is not Http or Https", this.transportProtocol);
                    break;
            }

            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R5, the SoapVersion is {0}.", this.service.SoapVersion);

            // Verify MS-SITESS requirement: MS-SITESS_R5
            bool isVerifyR5 = this.service.SoapVersion == SoapProtocolVersion.Soap11 || this.service.SoapVersion == SoapProtocolVersion.Soap12;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR5,
                5,
                @"[In Transport] Protocol messages MUST be formatted as specified either in [SOAP1.1] section 4, or in [SOAP1.2/1] section 5.");
        }

        /// <summary>
        /// Verify CreateWeb related requirements.
        /// </summary>
        /// <param name="createWebResult">The result of CreateWeb.</param>
        private void VerifyCreateWeb(CreateWebResponseCreateWebResult createWebResult)
        {
            Site.Log.Add(LogEntryKind.Comment, "Verify common requirements in CreateWeb operation.");
            this.VerifyCommonReqs();

            // When code can run to this line, it indicates the soap out message for this operation is received, else the operation will throw exception above.
            // So this operation's description is consistent with server.
            Site.CaptureRequirement(
                278,
                @"[In CreateWeb] [The CreateWeb operation is defined as:] <wsdl:operation name=""CreateWeb"">
              <wsdl:input message=""tns:CreateWebSoapIn"" />
              <wsdl:output message=""tns:CreateWebSoapOut"" />
              </wsdl:operation>");

            bool isSchemaRight = SchemaValidation.XmlValidationErrors.Count == 0 && SchemaValidation.XmlValidationWarnings.Count == 0;

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R280.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R280");

            // Verify MS-SITESS requirement: MS-SITESS_R280
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                280,
                @"[In CreateWeb] [The client sends a CreateWebSoapIn request message and] the server responds with a CreateWebSoapOut response message upon successful completion of creating the subsite.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R289.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R289");

            // Verify MS-SITESS requirement: MS-SITESS_R289
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                289,
                @"[In CreateWebSoapOut] The SOAP body contains a CreateWebResponse element.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R310.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R310");

            // Verify MS-SITESS requirement: MS-SITESS_R310
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                310,
                @"[In CreateWebResponse] [The CreateWebResponse element is defined as:] <s:element name=""CreateWebResponse"">
              <s:complexType>
              <s:sequence>
              <s:element minOccurs=""1"" maxOccurs=""1"" name=""CreateWebResult"">
              <s:complexType>
              <s:sequence>
              <s:element minOccurs=""1"" maxOccurs=""1"" name=""CreateWeb"">
              <s:complexType>
              <s:attribute name=""Url"" type=""s:string""/>
              </s:complexType>
              </s:element>
              </s:sequence>
              </s:complexType>
              </s:element>
              </s:sequence>
              </s:complexType>
              </s:element>");

            // Specifies whether the format of the Url is a fully qualified URL.
            bool isURI = false;
            isURI = Uri.IsWellFormedUriString(createWebResult.CreateWeb.Url, UriKind.RelativeOrAbsolute);
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R377, the value of isURI is {0}, the Url is {1}.", isURI, createWebResult.CreateWeb.Url);

            // Verify MS-SITESS requirement: MS-SITESS_R377
            Site.CaptureRequirementIfIsTrue(
                isURI,
                377,
                @"[In CreateWebResponse] Url: The fully qualified URL to the subsite which was successfully created.");
        }

        /// <summary>
        /// Verify ExportSolution related requirements.
        /// </summary>
        /// <param name="exportSolutionResultURL">The result of the export operation.</param>
        private void VerifyExportSolution(string exportSolutionResultURL)
        {
            Site.Log.Add(LogEntryKind.Comment, "Verify common requirements in ExportSolution operation.");
            this.VerifyCommonReqs();

            // When code can run to this line, it indicates the soap out message for this operation is received, else the operation will throw exception above.
            // So this operation's description is consistent with server.
            Site.CaptureRequirement(
                25,
                @"[In ExportSolution] [The ExportSolution operation is defined as:]
              <wsdl:operation name=""ExportSolution"">
              <wsdl:input message=""tns:ExportSolutionSoapIn"" />
              <wsdl:output message=""tns:ExportSolutionSoapOut"" />
              </wsdl:operation>");
            bool isSchemaRight = SchemaValidation.XmlValidationErrors.Count == 0 && SchemaValidation.XmlValidationWarnings.Count == 0;

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R27.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R27");

            // Verify MS-SITESS requirement: MS-SITESS_R27
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                27,
                @"[In ExportSolution] [The client sends an ExportSolutionSoapIn request message and] the server responds with an ExportSolutionSoapOut response message.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R33.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R33");

            // Verify MS-SITESS requirement: MS-SITESS_R33
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                33,
                @"[In ExportSolutionSoapOut] The SOAP body contains an ExportSolutionResponse element.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R50.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R50");

            // Verify MS-SITESS requirement: MS-SITESS_R50
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                50,
                @"[In ExportSolutionResponse] [The ExportSolutionResponse element is defined as:] <s:element name=""ExportSolutionResponse"">
              <s:complexType>
              <s:sequence>
              <s:element name=""ExportSolutionResult"" minOccurs=""1"" maxOccurs=""1"" type=""s:string"" />
              </s:sequence>
              </s:complexType>
              </s:element>");

            // If ExportSolution operation succeeds, we can verify whether ExportSolutionResult is a site-collection relative URL.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R432, the exportSolutionResultURL is {0}.", exportSolutionResultURL);

            // Verify MS-SITESS requirement: MS-SITESS_R432
            bool isVerifyR432 = Uri.IsWellFormedUriString(exportSolutionResultURL, UriKind.Relative) && exportSolutionResultURL.Substring(0, 1) != "/";

            Site.CaptureRequirementIfIsTrue(
                isVerifyR432,
                432,
                @"[In ExportSolutionResponse] ExportSolutionResult: It MUST be the site-collection relative URL of the created solution file.");
        }

        /// <summary>
        /// Verify ExportWeb related requirements.
        /// </summary>
        /// <param name="exportWebResult">The result of ExportWeb.</param>
        private void VerifyExportWeb(int exportWebResult)
        {
            Site.Log.Add(LogEntryKind.Comment, "Verify common requirements in ExportWeb operation.");
            this.VerifyCommonReqs();

            // When code can run to this line, it indicates the soap out message for this operation is received, else the operation will throw exception above.
            // So this operation's description is consistent with server.
            Site.CaptureRequirement(
                54,
                @"[In ExportWeb] [The Exportweb operation is defined as:] <wsdl:operation name=""ExportWeb"">
              <wsdl:input message=""tns:ExportWebSoapIn"" />
              <wsdl:output message=""tns:ExportWebSoapOut"" />
              </wsdl:operation>");
            bool isSchemaRight = SchemaValidation.XmlValidationErrors.Count == 0 && SchemaValidation.XmlValidationWarnings.Count == 0;

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R56, R66, R87.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R56");

            // Verify MS-SITESS requirement: MS-SITESS_R56
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                56,
                @"[In ExportWeb] [The client sends an ExportWebSoapIn request message and] the server responds with an ExportWebSoapOut response message.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R66.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R66");

            // Verify MS-SITESS requirement: MS-SITESS_R66
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                66,
                @"[In ExportWebSoapOut] The SOAP body contains an ExportWebResponse element.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R87.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R87");

            // Verify MS-SITESS requirement: MS-SITESS_R87
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                87,
                @"[In ExportWebResponse] [The ExportWebResponse element is defined as:] <s:element name=""ExportWebResponse"">
              <s:complexType>
              <s:sequence>
              <s:element minOccurs=""1"" maxOccurs=""1"" name=""ExportWebResult"" type=""s:int"" />
              </s:sequence>
              </s:complexType>
              </s:element>");

            // Specifies whether the ExportWebResult is one of the following values:1, 4, 5, 6, 7, 8.
            bool isResultInRange = (exportWebResult == 1) ||
                                      (exportWebResult == 4) ||
                                      (exportWebResult == 5) ||
                                      (exportWebResult == 6) ||
                                      (exportWebResult == 7) ||
                                      (exportWebResult == 8);
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R88, the result of ExportWeb is {0}", exportWebResult);

            // Verify MS-SITESS requirement: MS-SITESS_R88
            Site.CaptureRequirementIfIsTrue(
                isResultInRange,
                88,
                @"[In ExportWebResponse] ExportWebResult: The result of the operation is in the range [1, 4, 5, 6, 7, 8] as  specified in the table in section 3.1.4.2.2.");
        }

        /// <summary>
        /// Verify ExportWorkflowTemplate related requirements.
        /// </summary>
        /// <param name="exportWorkflowTemplateResult">The result of ExportWorkflowTemplate.</param>
        private void VerifyExportWorkflowTemplate(string exportWorkflowTemplateResult)
        {
            Site.Log.Add(LogEntryKind.Comment, "Verify common requirements in ExportWorkflowTemplate operation.");
            this.VerifyCommonReqs();

            // When code can run to this line, it indicates the soap out message for this operation is received, else the operation will throw exception above.
            // So this operation's description is consistent with server.
            Site.CaptureRequirement(
                96,
                @"[In ExportWorkflowTemplate] [The ExportWorkflowTemplate operation is defined as:] <wsdl:operation name=""ExportWorkflowTemplate"">
              <wsdl:input message=""tns:ExportWorkflowTemplateSoapIn"" />
              <wsdl:output message=""tns:ExportWorkflowTemplateSoapOut"" />
              </wsdl:operation>");
            bool isSchemaRight = SchemaValidation.XmlValidationErrors.Count == 0 && SchemaValidation.XmlValidationWarnings.Count == 0;

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R98,R104,R117.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R98");

            // Verify MS-SITESS requirement: MS-SITESS_R98
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                98,
                @"[In ExportWorkflowTemplate] [The client sends an ExportWorkflowTemplateSoapIn request message] and the server responds with an ExportWorkflowTemplateSoapOut response message.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R104.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R104");

            // Verify MS-SITESS requirement: MS-SITESS_R104
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                104,
                @"[In ExportWorkflowTemplateSoapOut] The SOAP body contains an ExportWorkflowTemplateResponse element.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R117.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R117");

            // Verify MS-SITESS requirement: MS-SITESS_R117
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                117,
                @"[In ExportWorkflowTemplateResponse] [The ExportWorkflowTemplateResponse element is defined as:] <s:element name=""ExportWorkflowTemplateResponse"">
              <s:complexType>
              <s:sequence>
              <s:element name=""ExportWorkflowTemplateResult"" minOccurs=""1"" maxOccurs=""1"" type=""s:string"" />
              </s:sequence>
              </s:complexType>
              </s:element>");

            // If ExportWorkflowTemplate operation succeeds, we can verify whether ExportWorkflowTemplateResult is a site-collection relative URL.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R450, the ExportWorkflowTemplateResult is {0}", exportWorkflowTemplateResult);

            bool isVerifyR450 = Uri.IsWellFormedUriString(exportWorkflowTemplateResult, UriKind.Relative) && exportWorkflowTemplateResult.Substring(0, 1) != "/";

            // Verify MS-SITESS requirement: MS-SITESS_R450
            Site.CaptureRequirementIfIsTrue(
                isVerifyR450,
                450,
                @"[In ExportWorkflowTemplateResponse] ExportWorkflowTemplateResult: It MUST be the site-relative URL of the created solution file..");
        }

        /// <summary>
        /// Verify GetSite related requirements.
        /// </summary>
        /// <param name="getSiteResult">The result of GetSite.</param>
        private void VerifyGetSite(string getSiteResult)
        {
            Site.Log.Add(LogEntryKind.Comment, "Verify common requirements in GetSite operation.");
            this.VerifyCommonReqs();

            // When code can run to this line, it indicates the soap out message for this operation is received, else the operation will throw exception above.
            // So this operation's description is consistent with server.
            Site.CaptureRequirement(
                120,
                @"[In GetSite] [The GetSite operation is defined as:] <wsdl:operation name=""GetSite"">
              <wsdl:input message=""tns:GetSiteSoapIn"" />
              <wsdl:output message=""tns:GetSiteSoapOut"" />
              </wsdl:operation>");
            bool isSchemaRight = SchemaValidation.XmlValidationErrors.Count == 0 && SchemaValidation.XmlValidationWarnings.Count == 0;

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R122.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R122");

            // Verify MS-SITESS requirement: MS-SITESS_R122
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                122,
                @"[In GetSite] [The client sends a GetSiteSoapIn request message and] the server responds with a GetSiteSoapOut response message.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R128.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R128");

            // Verify MS-SITESS requirement: MS-SITESS_R128
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                128,
                @"[In GetSiteSoapOut] The SOAP body contains a GetSiteResponse element.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R134.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R134");

            // Verify MS-SITESS requirement: MS-SITESS_R134
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                134,
                @"[In GetSiteResponse] [The GetSiteResponse element is defined as:] <s:element name=""GetSiteResponse"">
              <s:complexType>
              <s:sequence>
              <s:element minOccurs=""0"" maxOccurs=""1"" name=""GetSiteResult""
              type=""s:string"" />
              </s:sequence>
              </s:complexType>
              </s:element>");

            // Specifies whether the format of GetSiteResult is consistent with "<Site Url=""<Url>"" Id=""<Id>"" UserCodeEnabled=""<UserCodeEnabled>".
            string getSiteResultSchema = @"<s:schema xmlns:s=""http://www.w3.org/2001/XMLSchema"">
<s:element name=""Site"">
<s:complexType>
<s:attribute name=""Url"" type=""s:string"" use=""required""/>
<s:attribute name=""Id"" type=""s:string"" use=""required""/>
<s:attribute name=""UserCodeEnabled"" type=""s:string"" use=""required""/>
</s:complexType>
</s:element>
</s:schema>";

            AdapterHelper.MessageValidation(getSiteResult, getSiteResultSchema);

            if (!string.IsNullOrEmpty(getSiteResult))
            {
                bool isGetSiteResuletSchemaRight = AdapterHelper.ValidationInfo.Count == 0;

                // If the Schema of GetSiteResult element is right, R136 can be captured.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R136, the count of ValidationInfos is {0}", AdapterHelper.ValidationInfo.Count);

                // Verify MS-SITESS requirement: MS-SITESS_R136
                Site.CaptureRequirementIfIsTrue(
                    isGetSiteResuletSchemaRight,
                    136,
                    @"[In GetSiteResponse] [GetSiteResult:] It is in the following form: <Site Url=UrlString Id=IdString UserCodeEnabled=UserCodeEnabledString />");

                // Deserialize the returned string to a Site object.
                Site result;
                result = AdapterHelper.SiteResultDeserialize(getSiteResult);

                // Specifies whether the Url contained in GetSiteResult is the absolute URL of the site collection.
                string expectSiteUrl = Common.GetConfigurationPropertyValue(Constants.SiteCollectionUrl, this.Site).TrimEnd('/');
                string actualSiteUrl = result.Url.TrimEnd('/');
                bool isSiteUrl = expectSiteUrl.Equals(actualSiteUrl, StringComparison.CurrentCultureIgnoreCase);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R367, the actual SiteUrl is {0}", actualSiteUrl);

                // Verify MS-SITESS requirement: MS-SITESS_R367
                Site.CaptureRequirementIfIsTrue(
                    isSiteUrl,
                    367,
                    @"[In GetSiteResponse] [GetSiteResult:] Where UrlString is a quoted string that is the absolute URL of the site collection.");
            }
        }

        /// <summary>
        /// Verify GetSiteTemplates related requirements.
        /// </summary>
        /// <param name="templateList">SiteTemplates list.</param>
        private void VerifyGetSiteTemplates(Template[] templateList)
        {
            Site.Log.Add(LogEntryKind.Comment, "Verify common requirements in GetSiteTemplates operation.");
            this.VerifyCommonReqs();

            // When code can run to this line, it indicates the soap out message for this operation is received, else the operation will throw exception above.
            // So this operation's description is consistent with server.
            Site.CaptureRequirement(
                138,
                @"[In GetSiteTemplates] [The GetSiteTemplates operation is defined as:] <wsdl:operation name=""GetSiteTemplates"">
              <wsdl:input message=""tns:GetSiteTemplatesSoapIn"" />
              <wsdl:output message=""tns:GetSiteTemplatesSoapOut"" />
              </wsdl:operation>");
            bool isSchemaRight = SchemaValidation.XmlValidationErrors.Count == 0 && SchemaValidation.XmlValidationWarnings.Count == 0;

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R140.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R140");

            // Verify MS-SITESS requirement: MS-SITESS_R140
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                140,
                @"[In GetSiteTemplates] [The client sends a GetSiteTemplatesSoapIn request message and] the server responds with a GetSiteTemplatesSoapOut response message.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R146.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R146");

            // Verify MS-SITESS requirement: MS-SITESS_R146
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                146,
                @"[In GetSiteTemplatesSoapOut] The SOAP body contains a GetSiteTemplatesResponse element.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R151.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R151");

            // Verify MS-SITESS requirement: MS-SITESS_R151
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                151,
                @"[In GetSiteTemplatesResponse] [The GetSiteTemplatesResponse element is defined as:] <s:element name=""GetSiteTemplatesResponse"">
              <s:complexType>
              <s:sequence>
              <s:element minOccurs=""1"" maxOccurs=""1"" name=""GetSiteTemplatesResult""
              type=""s:unsignedInt"" />
              <s:element minOccurs=""0"" maxOccurs=""1"" name=""TemplateList""
              type=""tns:ArrayOfTemplate"" />
              </s:sequence>
              </s:complexType>
              </s:element>");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R155.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R155");

            // Verify MS-SITESS requirement: MS-SITESS_R155
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                155,
                @"[In GetSiteTemplatesResponse] [TemplateList:] The type of TemplateList is specified in section 3.1.4.5.3.1.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R159.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R159");

            // Verify MS-SITESS requirement: MS-SITESS_R159
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                159,
                @"[In ArrayOfTemplate] It [the ArrayOfTemplate complex type] contains the element Template, which is defined in section 3.1.4.5.3.2.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R160.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R160");

            // Verify MS-SITESS requirement: MS-SITESS_R160
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                160,
                @"[In ArrayOfTemplate] [The ArrayOfTemplate complex type is defined as:] <s:complexType name=""ArrayOfTemplate"">
              <s:sequence>
              <s:element minOccurs=""0"" maxOccurs=""unbounded"" name=""Template"" nillable=""true""
              type=""tns:Template"" />
              </s:sequence>
              </s:complexType>");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R163.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R163");

            // Verify MS-SITESS requirement: MS-SITESS_R163
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                163,
                @"[In Template] [The Template complex type is defined as:] <s:complexType name=""Template"">
              <s:attribute name=""ID"" type=""s:int"" use=""required"" />
              <s:attribute name=""Title"" type=""s:string"" use=""required"" />
              <s:attribute name=""Name"" type=""s:string"" use=""required"" />
              <s:attribute name=""IsUnique"" type=""s:boolean"" use=""required"" />
              <s:attribute name=""IsHidden"" type=""s:boolean"" use=""required"" />
              <s:attribute name=""Description"" type=""s:string"" />
              <s:attribute name=""ImageUrl"" type=""s:string"" use=""required"" />
              <s:attribute name=""IsCustom"" type=""s:boolean"" use=""required"" />
              <s:attribute name=""IsSubWebOnly"" type=""s:boolean"" use=""required"" />
              <s:attribute name=""IsRootWebOnly"" type=""s:boolean"" use=""required"" />
              <s:attribute name=""DisplayCategory"" type=""s:string"" />
              <s:attribute name=""FilterCategories"" type=""s:string"" />
              <s:attribute name=""HasProvisionClass"" type=""s:boolean"" use=""required"" />
              </s:complexType>");

            if (templateList != null)
            {
                for (int i = 0; i < templateList.Length; i++)
                {
                    // Check whether the "Name" consists of string, "#" and number.
                    bool isValidName = Regex.IsMatch(templateList[i].Name, @"^(\w+)#(\d+)$");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R168, the name of template is {0}", templateList[i].Name);

                    // If "Name" consists of string, "#" and number, then R168 can be verified.
                    // Verify MS-SITESS requirement: MS-SITESS_R168
                    Site.CaptureRequirementIfIsTrue(
                        isValidName,
                        168,
                        @"[In Template] [Name:] It contains the name of the site definition followed by a number sign (#), and then the site definition configuration number.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R170, the Template {0} unique.", templateList[i].IsUnique ? "is" : "is not");

                    // Verify MS-SITESS requirement: MS-SITESS_R170
                    Site.CaptureRequirementIfIsFalse(
                        templateList[i].IsUnique,
                        170,
                        @"[In Template] [IsUnique:] It MUST be false when sending[ and ignored on receipt].");

                    // Specifies whether the ImageUrl is a relative URL.
                    bool isURI = Uri.IsWellFormedUriString(templateList[i].ImageUrl, UriKind.Relative);

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R175, the ImageUrl is {0}", templateList[i].ImageUrl);

                    // Verify MS-SITESS requirement: MS-SITESS_R175
                    Site.CaptureRequirementIfIsTrue(
                        isURI,
                        175,
                        @"[In Template] [ImageUrl:] It MUST be the URL in relative to the URL of the site.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R177, the value of templateList[{0}].IsCustom is {1}", i, templateList[i].IsCustom);

                    // Verify MS-SITESS requirement: MS-SITESS_R177
                    Site.CaptureRequirementIfIsTrue(
                        templateList[i].IsCustom,
                        177,
                        @"[In Template] [IsCustom:] It MUST be true.");
                }
            }
        }

        /// <summary>
        /// Verify GetUpdatedFormDigest related requirements.
        /// </summary>
        /// <param name="getUpdateFormDigestResult">The result of GetUpdatedFormDigest.</param>
        private void VerifyGetUpdatedFormDigest(string getUpdateFormDigestResult)
        {
            Site.Log.Add(LogEntryKind.Comment, "Verify common requirements in GetUpdatedFormDigest operation.");
            this.VerifyCommonReqs();

            // When code can run to this line, it indicates the soap out message for this operation is received, else the operation will throw exception above.
            // So this operation's description is consistent with server.
            Site.CaptureRequirement(
                199,
                @"[In GetUpdatedFormDigest] [The GetUpdatedFormDigest operation is defined as:] <wsdl:operation name=""GetUpdatedFormDigest"">
              <wsdl:input message=""tns:GetUpdatedFormDigestSoapIn"" />
              <wsdl:output message=""tns:GetUpdatedFormDigestSoapOut"" />
              </wsdl:operation>");
            bool isSchemaRight = SchemaValidation.XmlValidationErrors.Count == 0 && SchemaValidation.XmlValidationWarnings.Count == 0;

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R201.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R201");

            // Verify MS-SITESS requirement: MS-SITESS_R201
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                201,
                @"[In GetUpdatedFormDigest] [The client sends a GetUpdatedFormDigestSoapIn request message and] the server responds with a GetUpdatedFormDigestSoapOut response message.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R207.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R207");

            // Verify MS-SITESS requirement: MS-SITESS_R207
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                207,
                @"[In GetUpdatedFormDigestSoapOut] The SOAP body contains a GetUpdatedFormDigestResponse element.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R211.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R211");

            // Verify MS-SITESS requirement: MS-SITESS_R211
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                211,
                @"[In GetUpdatedFormDigestResponse] [The GetUpdatedFormDigestResponse element is defined as:] <s:element name=""GetUpdatedFormDigestResponse"">
              <s:complexType>
              <s:sequence>
              <s:element minOccurs=""1"" maxOccurs=""1"" name=""GetUpdatedFormDigestResult""
              type=""s:string"" />
              </s:sequence>
              </s:complexType>
              </s:element>");

            if (Common.IsRequirementEnabled(332, this.Site))
            {
                string[] result = getUpdateFormDigestResult.Split(',');

                // If result contains two values, then verify R332.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R332, getUpdateFormDigestResult is {0}", getUpdateFormDigestResult);

                // Verify MS-SITESS requirement: MS-SITESS_R332
                Site.CaptureRequirementIfIsTrue(
                    result.Length == 2,
                    332,
                    @"[In Appendix B: Product Behavior] <11> Section 3.1.4.6.2.2: The Windows SharePoint Services implementation of the security validation consists of two values separated by a comma. (Microsoft Office SharePoint Server 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(368, this.Site))
            {
                string[] result = getUpdateFormDigestResult.Split(',');

                // Specifies whether the format of the second value is UTC format.
                bool isUTCFormat = false;
                bool isDateTime = false;
                DateTime currentTime = DateTime.UtcNow;
                DateTime timeStamp;
                string timeStampStr = result[1];
                timeStampStr = timeStampStr.Substring(0, 20);
                isDateTime = DateTime.TryParse(timeStampStr, out timeStamp);
                if (isDateTime)
                {
                    TimeSpan span = currentTime - timeStamp;

                    // Due to network reasons, allow 1 minute deviation here.
                    if (span.TotalMinutes > -1 && span.TotalMinutes < 1)
                    {
                        isUTCFormat = true;
                    }
                    else
                    {
                        // Log time span.
                        Site.Log.Add(LogEntryKind.Debug, "The time stamp {0} is considered incorrect. Because the time span between it and the time client validated this value exceeded 1 minute which is the defined deviation limit based on network reasons. This will cause verification failure of MS-SITESS_R368.", timeStamp);
                        isUTCFormat = false;
                    }
                }
                else
                {
                    Site.Log.Add(LogEntryKind.Debug, "The second value of the security validation {0} is not a correctly formatted time stamp.", timeStampStr);
                    isUTCFormat = false;
                }

                // Verify MS-SITESS requirement: MS-SITESS_R368
                Site.CaptureRequirementIfIsTrue(
                    isUTCFormat,
                    368,
                    @"[In Appendix B: Product Behavior][<11> Section 3.1.4.6.2.2:] the second value [of the security validation] is a time stamp in Coordinated Universal Time (UTC) format. (Microsoft Office SharePoint Server 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// Verify GetUpdatedFormDigest related requirements.
        /// </summary>
        /// <param name="getUpdatedFormDigestInfoResult">The result of GetUpdatedFormDigestInformation.</param>
        private void VerifyGetUpdatedFormDigestInformation(FormDigestInformation getUpdatedFormDigestInfoResult)
        {
            Site.Log.Add(LogEntryKind.Comment, "Verify common requirements in GetUpdatedFormDigestInformation operation.");
            this.VerifyCommonReqs();

            // If TimeoutSeconds>0 and DigestValue is not empty, then we can be sure that the server returns the two variables.
            bool isRS215Satisfied = getUpdatedFormDigestInfoResult.TimeoutSeconds > 0
                && !string.IsNullOrEmpty(getUpdatedFormDigestInfoResult.DigestValue);

            // If TimeoutSeconds>0 and DigestValue is not empty, R215 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R215, the TimeoutSeconds of getUpdatedFormDigestInfoResult is {0}, the DigestValue of getUpdatedFormDigestInfoResult is {1}", getUpdatedFormDigestInfoResult.TimeoutSeconds, getUpdatedFormDigestInfoResult.DigestValue);

            // Verify MS-SITESS requirement: MS-SITESS_R215
            Site.CaptureRequirementIfIsTrue(
                isRS215Satisfied,
                215,
                @"[In GetUpdatedFormDigestInformation] In this operation [GetUpdatedFormDigestInformation], the protocol server MUST return the security validation token's expiration time in addition to the security validation token.");

            // When code can run to this line, it indicates the soap out message for this operation is received, else the operation will throw exception above.
            // So this operation's description is consistent with server.
            Site.CaptureRequirement(
                216,
                @"[In GetUpdatedFormDigestInformation] [The GetUpdatedFormDigestInformation operation is defined as:] <wsdl:operation name=""GetUpdatedFormDigestInformation"">
              <wsdl:input message=""tns:GetUpdatedFormDigestInformationSoapIn"" />
              <wsdl:output message=""tns:GetUpdatedFormDigestInformationSoapOut"" />
              </wsdl:operation>");

            bool isSchemaRight = SchemaValidation.XmlValidationErrors.Count == 0 && SchemaValidation.XmlValidationWarnings.Count == 0;

            // If we can get the element name of GetUpdatedFormDigestInformationResponse, then we can get the element.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R224");

            // Verify MS-SITESS requirement: MS-SITESS_R224
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                224,
                @"[In GetUpdatedFormDigestInformationSoapOut] The SOAP body contains a GetUpdatedFormDigestInformationResponse element.");

            // If the response contains the element GetUpdatedFormDigestInformationResult,then we can be sure the format of GetUpdatedFormDigestInformationResponse.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R228");

            // Verify MS-SITESS requirement: MS-SITESS_R228
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                228,
                @"[In GetUpdatedFormDigestInformationResponse] [The GetUpdatedFormDigestInformationResponse element is defined as:] <s:element name=""GetUpdatedFormDigestInformationResponse"">
              <s:complexType>
              <s:sequence>
              <s:element minOccurs=""1"" maxOccurs=""1""
              name=""GetUpdatedFormDigestInformationResult""
              type=""tns:FormDigestInformation"" />
              </s:sequence>
              </s:complexType>
              </s:element>");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R218.
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                218,
                @"[In GetUpdatedFormDigestInformation] [The protocol client sends a GetUpdatedFormDigestInformationSoapIn request message and] the protocol server responds with a GetUpdatedFormDigestInformationSoapOut response message.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R231.
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                231,
                @"[In FormDigestInformation] [The FormDigestInformation complex type is defined as:] <s:complexType name=""FormDigestInformation"">
  <s:sequence>
    <s:element minOccurs=""0"" maxOccurs=""1"" name=""DigestValue"" type=""s:string"" /> 
    <s:element minOccurs=""1"" maxOccurs=""1"" name=""TimeoutSeconds"" type=""s:int"" /> 
    <s:element minOccurs=""0"" maxOccurs=""1"" name=""WebFullUrl"" type=""s:string"" />
    <s:element minOccurs=""0"" maxOccurs=""1"" name=""LibraryVersion"" type=""s:string"" />
    <s:element minOccurs=""0"" maxOccurs=""1"" name=""SupportedSchemaVersions"" type=""s:string"" />
  </s:sequence>
</s:complexType>");

            // If the library version is right, MS-CSOM_R1 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-CSOM_R1, the LibraryVersion of getUpdatedFormDigestInfoResult is {0}", getUpdatedFormDigestInfoResult.LibraryVersion);

            // Verify MS-SITESS requirement: MS-CSOM_R1
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                "MS-CSOM",
                1,
                @"[In VersionStringType] [The VersionStringType is defined as:] <xs:simpleType name=""VersionStringType"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:restriction base=""xs:string"">
    <xs:pattern value=""[0-9]{1,8}\.[0-9]{1,8}\.[0-9]{1,8}\.[0-9]{1,8}""/>
  </xs:restriction>
</xs:simpleType>");

            // If the SupportedSchemaVersions is comma-separated list of ""14.0.0.0"" or ""15.0.0.0"", MS-CSOM_R2 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-CSOM_R2", getUpdatedFormDigestInfoResult.SupportedSchemaVersions);

            bool isVerifyR2 = true;

            char[] separator = new char[] { ',' };
            string[] splitResult = { };
            splitResult = getUpdatedFormDigestInfoResult.SupportedSchemaVersions.Split(separator);
            for (int i = 0; i < splitResult.Length; i++)
            {
                if (splitResult[i] != "14.0.0.0" && splitResult[i] != "15.0.0.0")
                {
                    isVerifyR2 = false;
                    break;
                }
            }

            // Verify MS-SITESS requirement: MS-CSOM_R2
            Site.CaptureRequirementIfIsTrue(
                isVerifyR2,
                "MS-CSOM",
                2,
                @"[In Attributes] [SchemaVersion:] This value MUST be ""14.0.0.0"" or ""15.0.0.0"".");
        }

        /// <summary>
        /// Verify ImportWeb related requirements.
        /// </summary>
        /// <param name="importWebResult">The result of ImportWeb.</param>
        private void VerifyImportWeb(int importWebResult)
        {
            Site.Log.Add(LogEntryKind.Comment, "Verify common requirements in ImportWeb operation.");
            this.VerifyCommonReqs();

            // When code can run to this line, it indicates the soap out message for this operation is received, else the operation will throw exception above.
            // So this operation's description is consistent with server.
            Site.CaptureRequirement(
                235,
                @"[In ImportWeb] [The ImportWeb operation is defined as:] <wsdl:operation name=""ImportWeb"">
              <wsdl:input message=""tns:ImportWebSoapIn"" />
              <wsdl:output message=""tns:ImportWebSoapOut"" />
              </wsdl:operation>");

            bool isSchemaRight = SchemaValidation.XmlValidationErrors.Count == 0 && SchemaValidation.XmlValidationWarnings.Count == 0;

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R237.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R237");

            // Verify MS-SITESS requirement: MS-SITESS_R237
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                237,
                @"[In ImportWeb] [The client sends an ImportWebSoapIn request message and] the server responds with an ImportWebSoapOut response message");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R247.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R247");

            // Verify MS-SITESS requirement: MS-SITESS_R247
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                247,
                @"[In ImportWebSoapOut] The SOAP body contains an ImportWebResponse element.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R265.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R265");

            // Verify MS-SITESS requirement: MS-SITESS_R265
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                265,
                @"[In ImportWebResponse] [The ImportWebResponse element is defined as:] <s:element name=""ImportWebResponse"">
              <s:complexType>
              <s:sequence>
              <s:element minOccurs=""1"" maxOccurs=""1"" name=""ImportWebResult"" type=""s:int"" />
              </s:sequence>
              </s:complexType>
              </s:element>");

            // Specifies whether the ImportWebResult is one of the following values:1,2,4,5,6,8,11.
            bool isRS266Satisfied = importWebResult == 1 ||
                        importWebResult == 2 || importWebResult == 4 ||
                        importWebResult == 5 || importWebResult == 6 ||
                        importWebResult == 8 || importWebResult == 11;
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R266, the value of isRS266Satisfied is:{0}, importWebResult = {1}", isRS266Satisfied, importWebResult);

            // Verify MS-SITESS requirement: MS-SITESS_R266
            Site.CaptureRequirementIfIsTrue(
                isRS266Satisfied,
                266,
                @"[In ImportWebResponse] ImportWebResult: The result of the operation is in the range [1, 2, 4 , 5, 6, 8, 11]  as specified in the table in section 3.1.4.8.2.2.");
        }

        /// <summary>
        /// Verify VerifyDeleteWeb related requirements.
        /// </summary>
        private void VerifyDeleteWeb()
        {
            Site.Log.Add(LogEntryKind.Comment, "Verify common requirements in DeleteWeb operation.");
            this.VerifyCommonReqs();

            // When code can run to this line, it indicates the soap out message for this operation is received, else the operation will throw exception above.
            // So this operation's description is consistent with server.
            Site.CaptureRequirement(
                313,
                @"[In DeleteWeb] [The DeleteWeb operation is defined as:] <wsdl:operation name=""DeleteWeb"">
              <wsdl:input message=""tns:DeleteWebSoapIn"" />
              <wsdl:output message=""tns:DeleteWebSoapOut"" />
              </wsdl:operation>");

            bool isSchemaRight = SchemaValidation.XmlValidationErrors.Count == 0 && SchemaValidation.XmlValidationWarnings.Count == 0;

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R315.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R315");

            // Verify MS-SITESS requirement: MS-SITESS_R315
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                315,
                @"[In DeleteWeb] [The client sends a DeleteWebSoapIn request message and] the server responds with a DeleteWebSoapOut response message upon successful completion of deleting the subsite.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R321.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R321");

            // Verify MS-SITESS requirement: MS-SITESS_R321
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                321,
                @"[In DeleteWebSoapOut] The SOAP body contains a DeleteWebResponse element.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R326.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R326");

            // Verify MS-SITESS requirement: MS-SITESS_R326
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                326,
                @"[In DeleteWebResponse] [The DeleteWebResponse element is defined as:] <s:element name=""DeleteWebResponse"">
              <s:complexType />
              </s:element>");
        }

        /// <summary>
        /// A method used to validate the ArrayOfString complex type. 
        /// </summary>        
        private void ValidArrayOfStringComplexType()
        {
            Site.Log.Add(LogEntryKind.Comment, "Verify common requirements in DeleteWeb operation.");
            this.VerifyCommonReqs();
            bool isSchemaRight = SchemaValidation.XmlValidationErrors.Count == 0 && SchemaValidation.XmlValidationWarnings.Count == 0;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                422002,
                @" [In ArrayOfString] [The ArrayOfString is defined as:]<s:complexType name=""ArrayOfString"">
   <s:sequence>
     <s:element minOccurs=""1"" maxOccurs=""unbounded"" name=""string"" nillable=""true""
                  type=""s:string"" />
   </s:sequence>
 </s:complexType>
");
        }

        /// <summary>
        /// Verify IsScriptSafeUrl related requirements.
        /// </summary>
        /// <param name="isScriptSafeUrlResult">The result of IsScriptSafeUrl.</param>
        private void VerifyIsScriptSafeUrl(Boolean[] isScriptSafeUrlResult)
        {
            Site.Log.Add(LogEntryKind.Comment, "Verify common requirements in IsScriptSafeUrl operation.");
            this.VerifyCommonReqs();

            // If code can run to here, it means Microsoft SharePoint Foundation 2013 supports operation IsScriptSafeUrl.
            Site.CaptureRequirement(
                326001002,
                @"[In Appendix B: Product Behavior] Implementation does support this method [IsScriptSafeUrl]. <19> Section 3.1.4.11:  Only SharePoint Foundation 2013 supports this method.");

            // When code can run to this line, it indicates the soap out message for this operation is received, else the operation will throw exception above.
            // So this operation's description is consistent with server.
            Site.CaptureRequirement(
                326003,
                @" [In IsScriptSafeUrl] [The IsScriptSafeUrl operation is defined as:]<wsdl:operation name=""IsScriptSafeUrl"">
   <wsdl:input message=""tns:IsScriptSafeUrlSoapIn"" />
   <wsdl:output message=""tns:IsScriptSafeUrlSoapOut"" />
 </wsdl:operation>
");

            bool isSchemaRight = SchemaValidation.XmlValidationErrors.Count == 0 && SchemaValidation.XmlValidationWarnings.Count == 0;

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R326005.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R326005");

            // Verify MS-SITESS requirement: MS-SITESS_R326005
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                326005,
                @"[In IsScriptSafeUrl] [The client sends an IsScriptSafeUrlSoapIn request message, and] the server responds with an IsScriptSafeUrlSoapOut response message.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R326011.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R326011");

            // Verify MS-SITESS requirement: MS-SITESS_R326011
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                326011,
                @"[In IsScriptSafeUrlSoapOut] The SOAP body contains an IsScriptSafeUrlResponse element.");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R326020.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R326020");

            // Verify MS-SITESS requirement: MS-SITESS_R326020
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                326020,
                @" [In IsScriptSafeUrlResponse] [The IsScriptSafeUrlResponse element is defined as:]<s:element name=""IsScriptSafeUrlResponse"">
   <s:complexType>
     <s:sequence>
       <s:element minOccurs=""1"" maxOccurs=""1"" name=""IsScriptSafeUrlResult"" type=""tns:ArrayOfBoolean"" />
     </s:sequence>
   </s:complexType>
 </s:element>
");

            // When code can run to this line, it indicates the soap out message for this operation is received, else the operation will throw exception above.
            // So this operation's description is consistent with server.
            Site.CaptureRequirement(
                326001021,
                @" [In IsScriptSafeUrlResponse] IsScriptSafeUrlResult: An ArrayOfBoolean as defined in section 3.1.4.11.3.1, ");

            // When the variable isSchemaRight is true, it exposes that the message's format described in the Open Specification is consistent with server. So we can verify R326026.
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                326026,
                @"[InArrayOfBoolean]  [The ArrayOfBoolean complexType is defined as:]<s:complexType name=""ArrayOfBoolean"">
    <s:sequence>
       <s:element minOccurs=""0"" maxOccurs=""unbounded"" name=""boolean"" type=""s:boolean"" />
    </s:sequence>
 </s:complexType>
");

            if (isScriptSafeUrlResult != null)
            {
                this.ValidArrayOfStringComplexType();
            }
        }

    }
}