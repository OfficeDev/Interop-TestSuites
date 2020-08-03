namespace Microsoft.Protocols.TestSuites.MS_VERSS
{
    using System.Text.RegularExpressions;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The capture requirement part of adapter's implementation.
    /// </summary>
    public partial class MS_VERSSAdapter 
    {
        /// <summary>
        /// Verify the transport of protocol.
        /// </summary>
        private void VerifyTransport()
        {
            if (this.transportProtocol == TransportProtocol.HTTP)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R3");

                // Verify MS-VERSS requirement: MS-VERSS_R3
                // Because Adapter uses SOAP and HTTP to communicate with server, 
                // if server returned data without exception, this requirement has been captured.
                Site.CaptureRequirement(
                    3,
                    @"[In Transport] Protocol servers MUST support SOAP over HTTP.");
            }

            if (this.transportProtocol == TransportProtocol.HTTPS)
            {
                if (Common.IsRequirementEnabled(5, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R5");

                    // Verify MS-VERSS requirement: MS-VERSS_R5
                    Site.CaptureRequirement(
                        5,
                        @"[In Transport] Implementation does additionally support SOAP over HTTPS for enhancing the security of communication with protocol clients. (Windows SharePoint Services 3.0 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// Verify the GetVersions operation related requirements.
        /// </summary>
        /// <param name="getVersionResult">The GetVersionsResponseGetVersionsResult object
        /// indicates GetVersions operation response.</param>
        /// <param name="soapBody">The string value indicates the SOAP body in GetVersions operation response.</param>
        private void VerifyGetVersions(GetVersionsResponseGetVersionsResult getVersionResult, string soapBody)
        {
            bool isSchemaVerified = SchemaValidation.ValidationResult.Equals(ValidationResult.Success);

            #region Verify MS-VERSS_R119
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R119");

            // Verify MS-VERSS requirement: MS-VERSS_R119
            Site.CaptureRequirementIfIsTrue(
                isSchemaVerified,
                119,
                @"[In GetVersions][The schema of GetVersions is defined as:]
<wsdl:operation name=""GetVersions"">
    <wsdl:input message=""tns:GetVersionsSoapIn"" />
    <wsdl:output message=""tns:GetVersionsSoapOut"" />
</wsdl:operation>");
            #endregion

            #region Verify MS-VERSS_R121
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R121");

            // Verify MS-VERSS requirement: MS-VERSS_R121
            Site.CaptureRequirementIfIsNotNull(
                getVersionResult,
                121,
                @"[In GetVersions] [The protocol client sends a GetVersionsSoapIn request message,] and the protocol server responds with a GetVersionsSoapOut response message.");
            #endregion

            #region Verify MS-VERSS_R127
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R127");

            bool isR127Verified = AdapterHelper.IsExistElementInSoapBody(soapBody, "GetVersionsResponse");

            // Verify MS-VERSS requirement: MS-VERSS_R127
            Site.CaptureRequirementIfIsTrue(
                isR127Verified,
                127,
                @"[In GetVersionsSoapOut] The SOAP body contains a GetVersionsResponse element.");
            #endregion

            #region Verify MS-VERSS_R130
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R130");

            // Verify MS-VERSS requirement: MS-VERSS_R130
            Site.CaptureRequirementIfIsTrue(
                isSchemaVerified,
                130,
                @"[In GetVersionsResponse][The schema of GetVersionsResponse element is defined as:]  
<s:element name=""GetVersionsResponse"">
  <s:complexType>
    <s:sequence>
      <s:element minOccurs=""1"" maxOccurs=""1"" name=""GetVersionsResult"">
        <s:complexType>
          <s:sequence>
            <s:element name=""results"" minOccurs=""1"" maxOccurs=""1"" type=""tns:Results"" />
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");
            #endregion

            this.VerifyResultsComplexType(getVersionResult.results, isSchemaVerified);
        }

        /// <summary>
        /// Verify the RestoreVersion operation related requirements.
        /// </summary>
        /// <param name="restoreVersionResult">The RestoreVersionResponseRestoreVersionResult object indicates
        /// RestoreVersion operation response.</param>
        /// <param name="soapBody">The string value indicates the SOAP body in RestoreVersion operation response.</param>
        private void VerifyRestoreVersion(RestoreVersionResponseRestoreVersionResult restoreVersionResult, string soapBody)
        {
            bool isSchemaVerified = SchemaValidation.ValidationResult.Equals(ValidationResult.Success);

            #region Verify MS-VERSS_R18701
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R18701");

            // Verify MS-VERSS requirement: MS-VERSS_R18701
            Site.CaptureRequirementIfIsTrue(
                isSchemaVerified,
                18701,
                @"[In RestoreVersion][The schema of GetVersions is defined as:] 
<wsdl:operation name=""RestoreVersion"">
    <wsdl:input message=""tns:RestoreVersionSoapIn"" />
    <wsdl:output message=""tns:RestoreVersionSoapOut"" />
</wsdl:operation>");
            #endregion

            #region Verify MS-VERSS_R141
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R141");

            // Verify MS-VERSS requirement: MS-VERSS_R141
            Site.CaptureRequirementIfIsNotNull(
                restoreVersionResult,
                141,
                @"[In RestoreVersion] [The protocol client sends a RestoreVersionSoapIn request message,] and the protocol server responds with a RestoreVersionSoapOut response message.");
            #endregion

            #region Verify MS-VERSS_R147
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R147");

            bool isR147Verified = AdapterHelper.IsExistElementInSoapBody(soapBody, "RestoreVersionResponse");

            // Verify MS-VERSS requirement: MS-VERSS_R147
            Site.CaptureRequirementIfIsTrue(
                isR147Verified,
                147,
                @"[In RestoreVersionSoapOut] The SOAP body contains a RestoreVersionResponse element.");
            #endregion

            #region Verify MS-VERSS_R151
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R151");

            // Verify MS-VERSS requirement: MS-VERSS_R151
            Site.CaptureRequirementIfIsTrue(
                isSchemaVerified,
                151,
                @"[In RestoreVersionResponse][The schema of RestoreVersionResponse element is defined as:]  
<s:element name=""RestoreVersionResponse"">
  <s:complexType>
    <s:sequence>
      <s:element minOccurs=""1"" maxOccurs=""1"" name=""RestoreVersionResult"">
        <s:complexType>
          <s:sequence>
            <s:element name=""results"" minOccurs=""1"" maxOccurs=""1"" type=""tns:Results"" />
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");
            #endregion

            this.VerifyResultsComplexType(restoreVersionResult.results, isSchemaVerified);
        }

        /// <summary>
        /// Verify the DeleteVersion operation related requirements.
        /// </summary>
        /// <param name="deleteVersionResult">The DeleteVersionResponseDeleteVersionResult object indicates 
        /// DeleteVersion operation response.</param>
        /// <param name="soapBody">The string value indicates the SOAP body in DeleteVersion operation response.</param>
        private void VerifyDeleteVersion(DeleteVersionResponseDeleteVersionResult deleteVersionResult, string soapBody)
        {
            bool isSchemaVerified = SchemaValidation.ValidationResult.Equals(ValidationResult.Success);

            #region Verify MS-VERSS_R99
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R99");

            // Verify MS-VERSS requirement: MS-VERSS_R99
            Site.CaptureRequirementIfIsTrue(
                isSchemaVerified,
                99,
                @"[In DeleteVersion][The schema of DeleteVersion is defined as:]
<wsdl:operation name=""DeleteVersion"">
    <wsdl:input message=""tns:DeleteVersionSoapIn"" />
    <wsdl:output message=""tns:DeleteVersionSoapOut"" />
</wsdl:operation>");
            #endregion

            #region Verify MS-VERSS_R101
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R101");

            // Verify MS-VERSS requirement: MS-VERSS_R101
            Site.CaptureRequirementIfIsNotNull(
                deleteVersionResult,
                101,
                @"[In DeleteVersion] [The protocol client sends a DeleteVersionSoapIn request message,] and the protocol server responds with a DeleteVersionSoapOut response message.");
            #endregion

            #region Verify MS-VERSS_R107
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R107");

            bool isR107Verified = AdapterHelper.IsExistElementInSoapBody(soapBody, "DeleteVersionResponse");

            // Verify MS-VERSS requirement: MS-VERSS_R107
            Site.CaptureRequirementIfIsTrue(
                isR107Verified,
                107,
                @"[In DeleteVersionSoapOut] The SOAP body contains a DeleteVersionResponse element.");
            #endregion

            #region Verify MS-VERSS_R111
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R111");

            // Verify MS-VERSS requirement: MS-VERSS_R111
            Site.CaptureRequirementIfIsTrue(
                isSchemaVerified,
                111,
                @"[In DeleteVersionResponse][The schema of DeleteVersionResponse element is defined as:]  
<s:element name=""DeleteVersionResponse"">
  <s:complexType>
    <s:sequence>
      <s:element minOccurs=""1"" maxOccurs=""1"" name=""DeleteVersionResult"">
        <s:complexType>
          <s:sequence>
            <s:element name=""results"" minOccurs=""1"" maxOccurs=""1"" type=""tns:Results"" />
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");
            #endregion

            this.VerifyResultsComplexType(deleteVersionResult.results, isSchemaVerified);
        }

        /// <summary>
        /// Verify the DeleteAllVersions operation related requirements.
        /// </summary>
        /// <param name="deleteAllversionsResult">The DeleteAllVersionsResponseDeleteAllVersionsResult object
        /// indicates DeleteAllVersions operation response.</param>
        /// <param name="soapBody">The string value indicates the SOAP body in DeleteAllVersions operation response.</param>
        private void VerifyDeleteAllVersions(
            DeleteAllVersionsResponseDeleteAllVersionsResult deleteAllversionsResult,
            string soapBody)
        {
            bool isSchemaVerified = SchemaValidation.ValidationResult.Equals(ValidationResult.Success);

            #region Verify MS-VERSS_R80
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R80");

            // Verify MS-VERSS requirement: MS-VERSS_R80
            Site.CaptureRequirementIfIsTrue(
                isSchemaVerified,
                80,
                @"[In DeleteAllVersions operation][The schema of DeleteAllVersions is defined as:]
<wsdl:operation name=""DeleteAllVersions"">
    <wsdl:input message=""tns:DeleteAllVersionsSoapIn"" />
    <wsdl:output message=""tns:DeleteAllVersionsSoapOut"" />
</wsdl:operation>");
            #endregion

            #region Verify MS-VERSS_R82
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R82");

            // Verify MS-VERSS requirement: MS-VERSS_R82
            Site.CaptureRequirementIfIsNotNull(
                deleteAllversionsResult,
                82,
                @"[In DeleteAllVersions operation] [The protocol client sends a DeleteAllVersionsSoapIn request message], and the protocol server responds with a DeleteAllVersionsSoapOut response message.");
            #endregion

            #region Verify MS-VERSS_R88
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R88");

            bool isR88Verified = AdapterHelper.IsExistElementInSoapBody(soapBody, "DeleteAllVersionsResponse");

            // Verify MS-VERSS requirement: MS-VERSS_R88
            Site.CaptureRequirementIfIsTrue(
                isR88Verified,
                88,
                @"[In DeleteAllVersionsSoapOut] The SOAP body contains a DeleteAllVersionsResponse element.");
            #endregion

            #region Verify MS-VERSS_R91
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R91");

            // Verify MS-VERSS requirement: MS-VERSS_R91
            Site.CaptureRequirementIfIsTrue(
                isSchemaVerified,
                91,
                @"[In DeleteAllVersionsResponse][The schema of DeleteAllVersionsResponse element is defined as:]  
<s:element name=""DeleteAllVersionsResponse"">
  <s:complexType>
    <s:sequence>
      <s:element minOccurs=""1"" maxOccurs=""1"" name=""DeleteAllVersionsResult"">
        <s:complexType>
          <s:sequence>
            <s:element name=""results"" minOccurs=""1"" maxOccurs=""1"" type=""tns:Results"" />
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");
            #endregion

            this.VerifyResultsComplexType(deleteAllversionsResult.results, isSchemaVerified);
        }

        /// <summary>
        /// Verify the Results complex type related requirements
        /// </summary>
        /// <param name="result">The Results object indicates the Results complex type in response.</param>
        /// <param name="isSchemaVerified">A Boolean value indicates whether the schema has been verified.</param>
        private void VerifyResultsComplexType(Results result, bool isSchemaVerified)
        {
            #region Verify MS-VERSS_R38
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R38");

            // Verify MS-VERSS requirement: MS-VERSS_R38
            Site.CaptureRequirementIfIsTrue(
                isSchemaVerified,
                38,
                @"[In Results] The DeleteAllVersions, DeleteVersion, GetVersions, and RestoreVersion methods return the Results complex type.
<s:complexType name=""Results"">
  <s:sequence>
    <s:element name=""list"" maxOccurs=""1"" minOccurs=""1"">
      <s:complexType>
        <s:attribute name=""id"" type=""s:string"" use=""required"" />
      </s:complexType>
    </s:element>
    <s:element name=""versioning"" maxOccurs=""1"" minOccurs=""1"">
      <s:complexType>
        <s:attribute name=""enabled"" type=""s:unsignedByte"" use=""required"" />
      </s:complexType>
    </s:element>
    <s:element name=""settings"" maxOccurs=""1"" minOccurs=""1"">
      <s:complexType>
        <s:attribute name=""url"" type=""s:string"" use=""required"" />
      </s:complexType>
    </s:element>
    <s:element name=""result"" maxOccurs=""unbounded"" minOccurs=""1"" type=""tns:VersionData""/>
  </s:sequence>
</s:complexType>");
            #endregion

            #region Verify MS-VERSS_R42
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-VERSS_R42, the value of attribute versioning.enabled in Results is {0}", 
                result.versioning.enabled);

            bool isR42Verified = result.versioning.enabled == 0 || result.versioning.enabled == 1;

            // Verify MS-VERSS requirement: MS-VERSS_R42
            Site.CaptureRequirementIfIsTrue(
                isR42Verified,
                42,
                @"[In Results] versioning.enabled: The value of this attribute [versioning.enabled] MUST be ""0"" or ""1"".");
            #endregion

            #region Verify MS-VERSS_R45
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-VERSS_R45, the value of attribute settings.url in Results is {0}",
                result.settings.url);

            System.Uri settingsUrl = new System.Uri(result.settings.url);
            bool isR45Verified = AdapterHelper.ValidateAbsoluteUrlFormat(settingsUrl);

            // Verify MS-VERSS requirement: MS-VERSS_R45
            Site.CaptureRequirementIfIsTrue(
                isR45Verified,
                45,
                @"[In Results] settings.url: Specifies the URL to the webpage of versioning-related settings for the document library in which the file resides.");
            #endregion

            this.VerifyVersionDataComplexType(result.result, isSchemaVerified);
        }

        /// <summary>
        /// Verify the VersionData complex type related requirements
        /// </summary>
        /// <param name="versionDataArray">An array of VersionData object.</param>
        /// <param name="isSchemaVerified">A Boolean value indicates whether the schema has been verified.</param>
        private void VerifyVersionDataComplexType(VersionData[] versionDataArray, bool isSchemaVerified)
        {
            #region Verify MS-VERSS_R56
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R56");

            // Verify MS-VERSS requirement: MS-VERSS_R56
            Site.CaptureRequirementIfIsTrue(
                isSchemaVerified,
                56,
                @"[In VersionData] The VersionData complex type specifies the details about a single version of a file.
<s:complexType name=""VersionData"">
  <s:attribute name=""version"" type=""s:string"" use=""required"" />
  <s:attribute name=""url"" type=""s:string"" use=""required"" />
  <s:attribute name=""created"" type=""s:string"" use=""required"" />
  <s:attribute name=""createdRaw"" type=""s:string"" use=""required"" />  
  <s:attribute name=""createdBy"" type=""s:string"" use=""required"" />
  <s:attribute name=""createdByName"" type=""s:string"" use=""optional"" />
  <s:attribute name=""size"" type=""s:unsignedLong"" use=""required"" />
  <s:attribute name=""comments"" type=""s:string"" use=""required"" />
</s:complexType>");
            #endregion

            #region Verify MS-VERSS_R58
            // According to MS-OFCGLOS,the current version is the latest version of a document.
            // Then current version is the most recent version of the file.
            // If the current version is preceded with an at sign (@),then R58 will be verified.
            string currentVersion = AdapterHelper.GetCurrentVersion(versionDataArray);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R58, the value of current version is {0}", currentVersion);

            bool isR58Verified = currentVersion.StartsWith("@", System.StringComparison.CurrentCulture);

            // Verify MS-VERSS requirement: MS-VERSS_R58
            Site.CaptureRequirementIfIsTrue(
                isR58Verified,
                58,
                @"[In VersionData] version: The most recent version of the file MUST be preceded with an at sign (@).");
            #endregion

            #region Verify MS-VERSS_R59
            foreach (VersionData versionData in versionDataArray)
            {
                if (versionData.version != currentVersion)
                {
                    float versionNumber;
                    bool isR59Verified = float.TryParse(versionData.version, out versionNumber);

                    // Add the debug information
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "Verify MS-VERSS_R59, the value of version is {0}",
                        versionData.version);

                    // Verify MS-VERSS requirement: MS-VERSS_R59
                    Site.CaptureRequirementIfIsTrue(
                        isR59Verified,
                        59,
                        @"[In VersionData] version: All the other versions MUST exist without any prefix. ");
                }
            }
            #endregion

            #region Verify MS-VERSS_R60101
            foreach (VersionData versionData in versionDataArray)
            {
                System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"^@?\d+\.\d+$");

                bool isR60101Verified = regex.IsMatch(versionData.version);

                if (Common.IsRequirementEnabled(60101, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "Verify MS-VERSS_R60101, the value of attribute version in VersionData is {0}",
                        versionData.version);

                    // Verify MS-VERSS requirement: MS-VERSS_R60101
                    Site.CaptureRequirementIfIsTrue(
                        isR60101Verified,
                        60101,
                        @"[In Appendix B: Product Behavior] Implementation does contain the version of the file, including the major version and minor version numbers connected by period, for example, ""1.0"". (Windows SharePoint Services 3.0 and above follow this behavior.)");
                }
            }
            #endregion

            #region Verify MS-VERSS_R61
            foreach (VersionData versionData in versionDataArray)
            {
                System.Uri versionDataUrl = new System.Uri(versionData.url);
                bool isR61Verified = AdapterHelper.ValidateAbsoluteUrlFormat(versionDataUrl);

                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug, 
                    "Verify MS-VERSS_R61, the value of attribute URL in VersionData is {0}",
                    versionData.url);

                // Verify MS-VERSS requirement: MS-VERSS_R61
                Site.CaptureRequirementIfIsTrue(
                    isR61Verified,
                    61,
                    @"[In VersionData] url: The complete URL of the version of the file.");
            }
            #endregion

            #region Verify MS-VERSS_R164
            foreach (VersionData versionData in versionDataArray)
            {
                bool isR16401Enabled = Common.IsRequirementEnabled(16401, this.Site);
                bool isR16402Enabled = Common.IsRequirementEnabled(16402, this.Site);

                if (isR16401Enabled)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R16401");

                    // Verify MS-VERSS requirement: MS-VERSS_R16401
                    Site.CaptureRequirementIfIsNotNull(
                        versionData.createdRaw,
                        16401,
                        @"[In Appendix B: Product Behavior] Implementation does return this attribute. [In VersionData] createdRaw: The creation date and time for the version of the file in Datetime format, as specified in [ISO-8601]. (Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
                }

                if (isR16402Enabled)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R16402");

                    // Verify MS-VERSS requirement: MS-VERSS_R16402
                    Site.CaptureRequirementIfIsNull(
                        versionData.createdRaw,
                        16402,
                        @"[In Appendix B: Product Behavior] Implementation does not return this attribute. [In VersionData] createdRaw: The creation date and time for the version of the file in Datetime format, as specified in [ISO-8601]. (<1> Section 2.2.4.3: Windows SharePoint Services 3.0 does not return this attribute.)");
                }
            }
            #endregion

            #region Verify MS-VERSS_R65
            foreach (VersionData versionData in versionDataArray)
            {
                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-VERSS_R65, the value of attribute comments in VersionData is {0}",
                    versionData.comments);

                // Verify MS-VERSS requirement: MS-VERSS_R65
                Site.CaptureRequirementIfIsNotNull(
                    versionData.comments,
                    65,
                    @"[In VersionData] comments: The comment entered when the version of the file was replaced on the protocol server during check in.");
            }
            #endregion
        }

        /// <summary>
        /// Verify the SOAPFaultDetails complex type related requirements.
        /// </summary>
        /// <param name="soapExp">A SoapException object indicates the soap fault.</param>
        /// <param name="responseString">A string value indicates the raw xml of response.</param>
        private void VerifySOAPFaultDetails(SoapException soapExp, string responseString)
        {
            if (this.service.SoapVersion == SoapProtocolVersion.Soap11)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R1102");

                // Verify MS-VERSS requirement: MS-VERSS_R1102
                // According to the implementation of adapter, 
                // this requirement is directly verified when SoapException is received. 
                Site.CaptureRequirement(
                    1102,
                    @"[In Transport] Protocol server faults can be returned via SOAP faults, as specified in [SOAP1.1] section 4.4.");
            }

            if (this.service.SoapVersion == SoapProtocolVersion.Soap12)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R1103");

                // Verify MS-VERSS requirement: MS-VERSS_R1103
                // According to the implementation of adapter,
                // this requirement is directly verified when SoapException is received. 
                Site.CaptureRequirement(
                    1103,
                    @"[In Transport] Protocol server faults can be returned via SOAP faults, as specified in [SOAP1.2/-1/2007] section 5.4.");
            }

            string detailBody = SchemaValidation.GetSoapFaultDetailBody(responseString);
            ValidationResult detailResult = SchemaValidation.ValidateXml(Site, detailBody);
            bool isSchemaVerified = detailResult.Equals(ValidationResult.Success);

            #region Verify MS-VERSS_R51
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug, 
                "Verify MS-VERSS_R51, the value of SOAPFaultDetails complex type in SOAP fault is {0}",
                soapExp.Detail.InnerXml);

            // Verify MS-VERSS requirement: MS-VERSS_R51
            Site.CaptureRequirementIfIsTrue(
                isSchemaVerified,
                51,
                @"[In SOAPFaultDetails] The SOAPFaultDetails complex type specifies the details about a SOAP fault.
<s:schema xmlns:s=""http://www.w3.org/2001/XMLSchema"" targetNamespace="" http://schemas.microsoft.com/sharepoint/soap"">
   <s:complexType name=""SOAPFaultDetails"">
      <s:sequence>
         <s:element name=""errorstring"" type=""s:string""/>
         <s:element name=""errorcode"" type=""s:string"" minOccurs=""0""/>
      </s:sequence>
   </s:complexType>
</s:schema>");
            #endregion

            #region Verify MS-VERSS_R72
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R72");

            // Verify MS-VERSS requirement: MS-VERSS_R72
            // If R51 has been verified, then R72 will be verified.
            Site.CaptureRequirementIfIsTrue(
                isSchemaVerified,
                72,
                @"[In Protocol Details] These elements conform to the XSD of the SOAPFaultDetails complex type specified in section 2.2.4.2. ");
            #endregion

            #region Verify MS-VERSS_R53
            string errorCode = Common.ExtractErrorCodeFromSoapFault(soapExp);

            if (!string.IsNullOrEmpty(errorCode))
            {
                // Verify whether ErrorCode is a 4-byte hexadecimal.
                Regex reg = new Regex(@"^0x[0-9A-F]{8}$", RegexOptions.IgnoreCase);
                bool isR53Verified = reg.IsMatch(errorCode);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R53, the value of errorcode element is {0}", errorCode);

                // Verify MS-VERSS requirement: MS-VERSS_R53
                Site.CaptureRequirementIfIsTrue(
                    isR53Verified,
                    53,
                    @"[In SOAPFaultDetails] errorcode: The hexadecimal representation of a 4-byte result code.");
            }
            #endregion
        }

        /// <summary>
        /// Verify the requirements related to SOAP version.
        /// </summary>
        /// <param name="soapVersion">The soap protocol version.</param>
        private void VerifySOAPVersion(SoapProtocolVersion soapVersion)
        {
            if (soapVersion == SoapProtocolVersion.Soap11)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R701");

                // Verify MS-VERSS requirement: MS-VERSS_R701
                // According to the implementation of adapter, the message is formatted as SOAP1.1.
                // If this operation is invoked successfully, then this requirement can be verified.
                Site.CaptureRequirement(
                    701,
                    @"[In Transport] Protocol messages can be formatted as specified in [SOAP1.1] section 4.");
            }

            if (soapVersion == SoapProtocolVersion.Soap12)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R702");

                // Verify MS-VERSS requirement: MS-VERSS_R702
                // According to the implementation of adapter, the message is formatted as SOAP1.2.
                // If this operation is invoked successfully, then this requirement can be verified.
                Site.CaptureRequirement(
                    702,
                    @"[In Transport] Protocol messages can be formatted as specified in [SOAP1.2/-1/2007] section 5.");
            }
        }

        /// <summary>
        /// Verify the requirement related to server faults.
        /// </summary>
        private void VerifyServerFaults()
        {
             // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-VERSS_R1101");

            // Verify MS-VERSS requirement: MS-VERSS_R1101
            // According to the implementation of adapter, 
            // this requirement is directly verified when WebException is received. 
            Site.CaptureRequirement(
                1101,
                @"[In Transport] Protocol server faults can be returned via HTTP status codes, as specified in [RFC2616] section 10.");
        }
    }
}