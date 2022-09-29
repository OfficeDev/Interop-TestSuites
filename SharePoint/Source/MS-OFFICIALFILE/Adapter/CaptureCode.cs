namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    using System;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Class MS_OFFICIALFILEAdapter is used to verify adapter requirements.
    /// </summary>
    public partial class MS_OFFICIALFILEAdapter
    {
        /// <summary>
        /// Verify underlying transport protocol related requirements. These requirements can be captured directly 
        /// when the server returns a SOAP response successfully, which includes web service output message and soap fault.
        /// </summary>
        private void VerifyTransportRelatedRequirments()
        {
            TransportType transportType = Common.Common.GetConfigurationPropertyValue<TransportType>("TransportType", this.Site);

            // As SOAP response successfully returned, the following requirements can be captured directly
            switch (transportType)
            {
                case TransportType.HTTP:
                    // Transport soap over HTTP
                    // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R1
                    Site.CaptureRequirement(
                             "MS-OFFICIALFILE",
                             1,
                             @"[In Transport] Protocol servers MUST support SOAP over HTTP.");
                    break;

                case TransportType.HTTPS:                    
                    if (Common.Common.IsRequirementEnabled(301, this.Site))
                    {
                        // Having received the response successfully has proved the HTTPS transport is supported.
                        // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R301
                        Site.CaptureRequirement(
                                 301,
                                 @"[In Appendix B: Product Behavior] Implementation does additionally support SOAP over HTTPS for securing communication with protocol clients. (Microsoft Office SharePoint Server 2007 and above products follow this behavior.)");
                    }
                    else
                    {
                        this.Site.Assume.Inconclusive("Implementation does not additionally support SOAP over HTTPS for securing communication with protocol clients.");
                    }

                    break;

                default:
                    Site.Assert.Fail("Transport can only either HTTP or HTTPS.");
                    break;
            }

            SoapProtocolVersion soapVersion = Common.Common.GetConfigurationPropertyValue<SoapProtocolVersion>("SoapVersion", this.Site);

            // Add the log information for current SoapVersion.
            Site.Log.Add(Microsoft.Protocols.TestTools.LogEntryKind.Comment, "Current SoapVersion is {0}.", soapVersion.ToString());

            // As the transport is successful, the following requirements can be captured directly.
            if (soapVersion == SoapProtocolVersion.Soap11)
            {
                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R6
                Site.CaptureRequirement(
                         "MS-OFFICIALFILE",
                         6,
                         @"[In Transport] Protocol messages MUST be formatted as specified either in [SOAP1.1] (Section 4, SOAP Envelope).");
            }
            else if (soapVersion == SoapProtocolVersion.Soap12)
            {
                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R6001
                Site.CaptureRequirement(
                         "MS-OFFICIALFILE",
                         6001,
                         @"[In Transport] Protocol messages MUST be formatted as specified either in [SOAP1.1] (Section 4, SOAP Envelope) or in [SOAP1.2-1/2007] (Section 5, SOAP Message Construct).");
            }
            else
            {
                Site.Assume.Fail("The SOAP type is not recognized.");
            }
        }

        /// <summary>
        /// Verify the requirements of FinalRoutingDestinationFolderUrl in Adapter.
        /// When the server returns a GetFinalRoutingDestinationFolderUrl response successfully, 
        /// which includes web service output message and soap fault.
        /// </summary>
        private void VerifyGetFinalRoutingDestinationFolderUrl()
        {
            // Schema verification has been made by Proxy class before the response is returned, then all schema
            // related requirements can be directly verified, if no exception is thrown.
            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R61
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     61,
                     @"[In GetFinalRoutingDestinationFolderUrl] 
                     <wsdl:operation name=""GetFinalRoutingDestinationFolderUrl"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
                       <wsdl:input message=""tns:GetFinalRoutingDestinationFolderUrlSoapIn""/>
                       <wsdl:output message=""tns:GetFinalRoutingDestinationFolderUrlSoapOut""/>
                     </wsdl:operation>");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R63
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     63,
                     @"[In GetFinalRoutingDestinationFolderUrl] The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R106
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     106,
                     @"[In GetFinalRoutingDestinationFolderUrlSoapOut] The SOAP body contains the GetFinalRoutingDestinationFolderUrlResponse element.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R112
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     112,
                     @"[In GetFinalRoutingDestinationFolderUrlResponse] The GetFinalRoutingDestinationFolderUrlResponse element specifies the result data for the GetFinalRoutingDestinationFolderUrl WSDL operation.
                     <xs:element name=""GetFinalRoutingDestinationFolderUrlResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:complexType>
                         <xs:sequence>
                           <xs:element minOccurs=""0"" maxOccurs=""1"" name=""GetFinalRoutingDestinationFolderUrlResult"" type=""tns:DocumentRoutingResult""/>
                         </xs:sequence>
                       </xs:complexType>
                     </xs:element>");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R113
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     113,
                     @"[In GetFinalRoutingDestinationFolderUrlResponse] GetFinalRoutingDestinationFolderUrlResult: Data details about the result, which is an XML fragment that MUST conform to the XML schema of the DocumentRoutingResult complex type.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R114
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     114,
                     @"[DocumentRoutingResult] Result data details for a GetFinalRoutingDestinationFolderUrl WSDL operation.
                     <xs:complexType name=""DocumentRoutingResult"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:sequence>
                         <xs:element minOccurs=""0"" maxOccurs=""1"" name=""Url"" type=""xs:string""/>
                         <xs:element minOccurs=""1"" maxOccurs=""1"" name=""ResultType"" type=""tns:DocumentRoutingResultType""/>
                         <xs:element minOccurs=""1"" maxOccurs=""1"" name=""CollisionSetting"" type=""tns:DocumentRoutingCollisionSetting""/>
                       </xs:sequence>
                     </xs:complexType>");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R118
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     118,
                     @"[In DocumentRoutingResultType] Type of result.
                     <xs:simpleType name=""DocumentRoutingResultType"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:restriction base=""xs:string"">
                         <xs:enumeration value=""Success""/>
                         <xs:enumeration value=""SuccessToDropOffLibrary""/>
                         <xs:enumeration value=""MissingRequiredProperties""/>
                         <xs:enumeration value=""NoMatchingRules""/>
                         <xs:enumeration value=""DocumentRoutingDisabled""/>
                         <xs:enumeration value=""PermissionDeniedAtDestination""/>
                         <xs:enumeration value=""ParsingDisabledAtDestination""/>
                         <xs:enumeration value=""OriginalSaveLocationIsDocumentSet""/>
                         <xs:enumeration value=""NoEnforcementAtOriginalSaveLocation""/>
                         <xs:enumeration value=""UnknownFailure""/>
                       </xs:restriction>
                     </xs:simpleType>");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R130
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     130,
                     @"[In DocumentRoutingCollisionSetting]
                     <xs:simpleType name=""DocumentRoutingCollisionSetting"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:restriction base=""xs:string"">
                         <xs:enumeration value=""NoCollision""/>
                         <xs:enumeration value=""UseSharePointVersioning""/>
                         <xs:enumeration value=""AppendUniqueSuffixes""/>
                       </xs:restriction>
                     </xs:simpleType>");
        }

        /// <summary>
        /// Verify the requirements of GetRoutingInfo in Adapter When the server returns a GetRoutingInfo response successfully, 
        /// which includes web service output message and soap fault.
        /// </summary>
        private void VerifyGetRoutingInfo()
        {
            // Schema verification has been made by Proxy class before the response is returned, then all schema
            // related requirements can be directly verified, if no exception is thrown.
            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R160
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     160,
                     @"[In GetRecordRouting] [The following is the WSDL port type specification of the GetRecordRouting WSDL operation.]
<wsdl:operation name=""GetRecordRouting"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input message=""tns:GetRecordRoutingSoapIn""/>
  <wsdl:output message=""tns:GetRecordRoutingSoapOut""/>
</wsdl:operation>");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R161
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     161,
                     @"[In GetRecprdRouting] The protocol client sends a GetRecordRoutingSoapIn request WSDL message and the protocol server MUST respond with a GetRecordRoutingSoapOut response WSDL message. [as follows: The protocol server returns an implementation-specific value in the GetRecordRoutingResult element that MUST be ignored.]");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R168
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     168,
                     @"[In GetRecordRoutingSoapOut] The SOAP body contains the GetRecordRoutingResponse element.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R171
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     171,
                     @"[In GetRecordRoutingResponse] The GetRecordRoutingResponse element specifies the result data for the GetRecordRouting WSDL operation.
                     <xs:element name=""GetRecordRoutingResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:complexType>
                         <xs:sequence>
                           <xs:element minOccurs=""0"" maxOccurs=""1"" name=""GetRecordRoutingResult"" type=""xs:string""/>
                         </xs:sequence>
                       </xs:complexType>
                     </xs:element>");
        }

        /// <summary>
        /// Verify the requirements of GetRoutingInfo When the server returns a GetRoutingInfo response successfully, 
        /// which includes web service output message and soap fault.
        /// </summary>
        private void VerifyGetRoutingCollectionInfo()
        {
            // Schema verification has been made by Proxy class before the response is returned, then all schema
            // related requirements can be directly verified, if no exception is thrown.
            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R173
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     173,
                     @"[In GetRecordRoutingCollection] [The following is the WSDL port type specification of the GetRecordRoutingCollection WSDL operation.]
<wsdl:operation name=""GetRecordRoutingCollection"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input message=""tns:GetRecordRoutingCollectionSoapIn""/>
  <wsdl:output message=""tns:GetRecordRoutingCollectionSoapOut""/>
</wsdl:operation>");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R174
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     174,
                     @"[In GetRecordRoutingCollection] The protocol client sends a GetRecordRoutingCollectionSoapIn request WSDL message, and the protocol server MUST respond with a GetRecordRoutingCollectionSoapOut response WSDL message.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R181
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     181,
                     @"[In GetRecordRoutingCollectionSoapOut] The SOAP body contains the GetRecordRoutingCollectionResponse element.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R183
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     183,
                     @"[In GetRecordRoutingCollectionResponse] The GetRecordRoutingCollectionResponse element specifies the result data for the GetRecordRoutingCollection WSDL operation.
                     <xs:element name=""GetRecordRoutingCollectionResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:complexType>
                         <xs:sequence>
                           <xs:element minOccurs=""0"" maxOccurs=""1"" name=""GetRecordRoutingCollectionResult"" type=""xs:string""/>
                         </xs:sequence>
                       </xs:complexType>
                     </xs:element>");
        }

        /// <summary>
        /// Verify the requirements of GetHoldsInfo When the server returns a GetHoldsInfo response successfully, 
        /// which includes web service output message and soap fault.
        /// </summary>
        private void VerifyGetHoldsInfo()
        {
            // Schema verification has been made by Proxy class before the response is returned, then all schema
            // related requirements can be directly verified.
            #region Requirments of schema are verified directly according to the comment above.

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R135
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     135,
                     @"[In GetHoldsInfo] [The following is the WSDL port type specification of the GetHoldsInfo WSDL operation.]
<wsdl:operation name=""GetHoldsInfo"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input message=""tns:GetHoldsInfoSoapIn""/>
  <wsdl:output message=""tns:GetHoldsInfoSoapOut""/>
</wsdl:operation>");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R136
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     136,
                     @"[In GetHoldsInfo] The protocol client sends a GetHoldsInfoSoapIn request WSDL message and the protocol server MUST respond with a GetHoldsInfoSoapOut response WSDL message as follows.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R137
            // This requirement is verified directly after every items in ArrayOfHoldInfo is verified.
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     137,
                     @"[In GetHoldsInfo] [The protocol client sends a GetHoldsInfoSoapIn request WSDL message and the protocol server MUST respond with a GetHoldsInfoSoapOut response WSDL message as follows.] The protocol server returns the data associated with the legal holds in the repository in the GetHoldsInfoResult element.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R143
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     143,
                     @"[In GetHoldsInfoSoapOut] The SOAP body contains the GetHoldsInfoResponse element.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R145
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     145,
                     @"[In GetHoldsInfoResponse] The GetHoldsInfoResponse element specifies the result data for the GetHoldsInfo WSDL operation.
                     <xs:element name=""GetHoldsInfoResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:complexType>
                         <xs:sequence>
                           <xs:element minOccurs=""0"" maxOccurs=""1"" name=""GetHoldsInfoResult"" type=""tns:ArrayOfHoldInfo""/>
                         </xs:sequence>
                       </xs:complexType>
                     </xs:element>");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R147
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     147,
                     @"[In ArrayOfHoldInfo] A list of legal holds.
                     <xs:complexType name=""ArrayOfHoldInfo"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:sequence>
                         <xs:element minOccurs=""0"" maxOccurs=""unbounded"" name=""HoldInfo"" nillable=""true"" type=""tns:HoldInfo""/>
                       </xs:sequence>
                     </xs:complexType>");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R149
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     149,
                     @"[In HoldInfo] A legal hold.
                     <xs:complexType name=""HoldInfo"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:sequence>
                         <xs:element minOccurs=""0"" maxOccurs=""1"" name=""Url"" type=""xs:string""/>
                         <xs:element minOccurs=""0"" maxOccurs=""1"" name=""Title"" type=""xs:string""/>
                         <xs:element minOccurs=""0"" maxOccurs=""1"" name=""Description"" type=""xs:string""/>
                         <xs:element minOccurs=""0"" maxOccurs=""1"" name=""ManagedBy"" type=""xs:string""/>
                         <xs:element minOccurs=""0"" maxOccurs=""1"" name=""RepositoryName"" type=""xs:string""/>
                         <xs:element minOccurs=""0"" maxOccurs=""1"" name=""DiscoveryQueries"" type=""xs:string""/>
                         <xs:element minOccurs=""1"" maxOccurs=""1"" name=""Id"" type=""xs:int""/>
                         <xs:element minOccurs=""1"" maxOccurs=""1"" name=""ListId"" xmlns:s1=""http://microsoft.com/wsdl/types/"" type=""s1:guid""/>
                         <xs:element minOccurs=""1"" maxOccurs=""1"" name=""WebId"" xmlns:s1=""http://microsoft.com/wsdl/types/"" type=""s1:guid""/>
                       </xs:sequence>
                     </xs:complexType>");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R159
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     159,
                     @"[In guid] A GUID.
                     <xs:simpleType name=""guid"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:restriction base=""xs:string"">
                         <xs:pattern value=""[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}""/>
                       </xs:restriction>
                     </xs:simpleType>");

            #endregion  verify schema of response result of GetHoldsInfo
        }

        /// <summary>
        /// Verify the requirements of GetServerInfo and parse the GetServerInfoResult from xml to class When the server returns a GetServerInfo response successfully. 
        /// </summary>
        /// <param name="serverInfo">The response of GetServerInfo.</param>
        /// <returns>Return the GetServerInfoResult corresponding class instance.</returns>
        private ServerInfo VerifyAndParseGetServerInfo(string serverInfo)
        {
            #region Requirments of schema are verified directly according to the comment above.

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R185
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     185,
                     @"[In GetServerInfo] [The following is the WSDL port type specification of the GetServerInfo WSDL operation.]
<wsdl:operation name=""GetServerInfo"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input message=""tns:GetServerInfoSoapIn""/>
  <wsdl:output message=""tns:GetServerInfoSoapOut""/>
</wsdl:operation>");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R193
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     193,
                     @"[In GetServerInfoSoapOut] The SOAP body contains the GetServerInfoResponse element.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R195
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     195,
                     @"[In GetServerInfoResponse] The GetServerInfoResponse element specifies the result data for the GetServerInfo WSDL operation.
                     <xs:element name=""GetServerInfoResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:complexType>
                         <xs:sequence>
                           <xs:element minOccurs=""0"" maxOccurs=""1"" name=""GetServerInfoResult"" type=""xs:string""/>
                         </xs:sequence>
                       </xs:complexType>
                     </xs:element>");

            #endregion

            // Validate the GetServerInfoResult string which confirm to the specified XML definition.
            string result = @"<ServerInfoResult xmlns=""http://schemas.microsoft.com/sharepoint/soap/recordsrepository/"">" + serverInfo + "</ServerInfoResult>";
            Common.ValidationResult validationResult = Common.SchemaValidation.ValidateXml(this.Site, result);

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R196
            Site.CaptureRequirementIfAreEqual<Common.ValidationResult>(
                     Common.ValidationResult.Success,
                     validationResult,
                     "MS-OFFICIALFILE",
                     196,
                     @"[In GetServerInfoResponse] GetServerInfoResult: Type and version information for a protocol server, which MUST be an XML fragment that conforms to the XML schema of the ServerInfo complex type.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R34
            Site.CaptureRequirementIfAreEqual<Common.ValidationResult>(
                     Common.ValidationResult.Success,
                     validationResult,
                     "MS-OFFICIALFILE",
                     34,
                     @"[In ServerInfo] Server information. 
<xs:complexType name=""ServerInfo"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element name=""ServerType"" type=""xs:string""/>
    <xs:element name=""ServerVersion"" type=""xs:string""/>
    <xs:element minOccurs=""0"" name=""RoutingWeb"" type=""xs:string""/>
  </xs:sequence>
</xs:complexType>");

            // Parse the GetServerInfoResult from xml to class
            ServerInfo serverInfoInstance = this.ParseGetServerInfoResult(serverInfo);
           
            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R1025
            Site.CaptureRequirementIfAreNotEqual(
                     string.Empty,
                     serverInfoInstance.ServerVersion,
                     "MS-OFFICIALFILE",
                     1025,
                     @"[In ServerInfo] ServerVersion: MUST be non-empty.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R1023
            Site.CaptureRequirementIfAreNotEqual(
                     string.Empty,
                     serverInfoInstance.ServerType,
                     "MS-OFFICIALFILE",
                     1023,
                     @"[In ServerInfo] ServerType: MUST be non-empty.");

            // Add log information for the Requirement R1024.
            bool isR1024Verified = serverInfoInstance.ServerType.Length <= 256;
            Site.Log.Add(
                LogEntryKind.Debug,
                "For R1024 the ServerType length MUST be less than or equal to 256 characters, but actual length {0}.",
                serverInfoInstance.ServerType.Length);

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R1024
            Site.CaptureRequirementIfIsTrue(
                     isR1024Verified,
                     "MS-OFFICIALFILE",
                     1024,
                     @"[In ServerInfo] ServerType: MUST be less than or equal to 256 characters in length.");

            // Add log information for the Requirement R1026.
            bool isR1026Verified = serverInfoInstance.ServerVersion.Length <= 256;
            Site.Log.Add(
                LogEntryKind.Debug,
                "For R1024 the ServerVersion length MUST be less than or equal to 256 characters, but actual length {0}.",
                serverInfoInstance.ServerVersion.Length);

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R1026
            Site.CaptureRequirementIfIsTrue(
                     isR1026Verified,
                     "MS-OFFICIALFILE",
                     1026,
                     @"[In ServerInfo] ServerVersion: MUST be less than or equal to 256 characters in length.");

            if (serverInfoInstance.RoutingWeb != null)
            {
                bool routingWebResult;
                bool isR1032Verified = bool.TryParse(serverInfoInstance.RoutingWeb, out routingWebResult);

                Site.Log.Add(
                    LogEntryKind.Debug,
                    "For R1032 the RoutingWeb length MUST conform to the XML schema of the boolean simple type, but actual value {0}.",
                    serverInfoInstance.RoutingWeb);

                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R1032
                Site.CaptureRequirementIfIsTrue(
                    isR1032Verified,
                    "MS-OFFICIALFILE",
                    1032,
                    @"[In ServerInfo] RoutingWeb: If present, it MUST conform to the XML schema of the boolean simple type.");
            }

            return serverInfoInstance;
        }

        /// <summary>
        /// Verify the requirements of SubmitFile in Adapter.
        /// When the server returns a SubmitFile response successfully, 
        /// which includes web service output message and soap fault.
        /// </summary>
        /// <param name="result">The response of SubmitFile.</param>
        /// <returns>Return the SubmitFileResult corresponding class instance.</returns>
        private SubmitFileResult VerifyAndParseSubmitFile(string result)
        {
            #region verify schema of SubmitFile
 
            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R197
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     197,
                     @"[In SubmitFile] [The following is the WSDL port type specification of the SubmitFile WSDL operation.]
<wsdl:operation name=""SubmitFile"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input message=""tns:SubmitFileSoapIn""/>
  <wsdl:output message=""tns:SubmitFileSoapOut""/>
</wsdl:operation>");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R224
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     224,
                     @"[In SubmitFileSoapOut] The SOAP body contains the SubmitFileResponse element.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R198
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     198,
                     @"[In SubmitFile] The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R237
            Site.CaptureRequirement(
                     "MS-OFFICIALFILE",
                     237,
                     @"[In SubmitFileResponse] The SubmitFileResponse element specifies the result data for the SubmitFile WSDL operation.
                     <xs:element name=""SubmitFileResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:complexType>
                         <xs:sequence>
                           <xs:element minOccurs=""0"" maxOccurs=""1"" name=""SubmitFileResult"" type=""xs:string""/>
                         </xs:sequence>
                       </xs:complexType>
                     </xs:element>");

            #endregion verify of SubmitFile

            #region schema of response result of SubmitFile

            // Validate the SubmitFileResult string which confirm to the specified XML definition.
            result = @"<SubmitFileResult xmlns=""http://schemas.microsoft.com/sharepoint/soap/recordsrepository/"">" + result + "</SubmitFileResult>";
            Common.ValidationResult validationResult = Common.SchemaValidation.ValidateXml(this.Site, result);

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R238
            Site.CaptureRequirementIfAreEqual<Common.ValidationResult>(
                Common.ValidationResult.Success,
                validationResult,
                "MS-OFFICIALFILE",
                238,
                @"[In SubmitFileResponse] SubmitFileResult: Data details about the result, which is a string of an encoded XML fragment that MUST conform to the XML schema of the SubmitFileResult complex type.");

            // Parse the SubmitFileResult from xml to class.
            SubmitFileResult value = this.ParseSubmitFileResult(result);

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R239
            Site.CaptureRequirementIfAreEqual<Common.ValidationResult>(
                    Common.ValidationResult.Success,
                    validationResult,
                    "MS-OFFICIALFILE",
                     239,
                     @"[In SubmitFileResult] The detailed data result for the SubmitFile WSDL operation.
                     <xs:complexType name=""SubmitFileResult"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:sequence>
                         <xs:element name=""ResultCode"" type=""tns:SubmitFileResultCode""/>
                         <xs:choice>
                           <xs:element minOccurs=""0"" name=""ResultUrl"" type=""xs:anyURI""/>
                           <xs:element minOccurs=""0"" name=""AdditionalInformation"" type=""xs:string""/>
                         </xs:choice>
                         <xs:element minOccurs=""0"" name=""CustomProcessingResult"" type=""tns:CustomProcessingResult""/>
                       </xs:sequence>
                     </xs:complexType>");

            if (value.CustomProcessingResult != null)
            {
                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R244
                Site.CaptureRequirementIfAreEqual<Common.ValidationResult>(
                         Common.ValidationResult.Success,
                         validationResult,
                         "MS-OFFICIALFILE",
                         244,
                         @"[In CustomProcessingResult] The result of custom processing of a legal hold.
                     <xs:complexType name=""CustomProcessingResult"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:sequence>
                         <xs:element minOccurs=""0"" name=""HoldsProcessingResult"" type=""tns:HoldProcessingResult""/>
                       </xs:sequence>
                     </xs:complexType");

                    // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R256
                Site.CaptureRequirementIfAreEqual<Common.ValidationResult>(
                     Common.ValidationResult.Success,
                     validationResult,
                     "MS-OFFICIALFILE",
                     256,
                     @"[In HoldProcessingResult] The result of processing a legal hold.
                     <xs:simpleType name=""HoldProcessingResult"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:restriction base=""xs:string"">
                         <xs:enumeration value=""Success""/>
                         <xs:enumeration value=""Failure""/>
                         <xs:enumeration value=""InDropOffZone""/>
                       </xs:restriction>
                     </xs:simpleType>");
            }

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R246
            Site.CaptureRequirementIfAreEqual<Common.ValidationResult>(
                     Common.ValidationResult.Success,
                     validationResult,
                     "MS-OFFICIALFILE",
                     246,
                     @"[In SubmitFileResultCode] The result status code of a SubmitFile WSDL operation.
                     <xs:simpleType name=""SubmitFileResultCode"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                       <xs:restriction base=""xs:string"">
                         <xs:enumeration value=""Success""/>
                         <xs:enumeration value=""MoreInformation""/>
                         <xs:enumeration value=""InvalidRouterConfiguration""/>
                         <xs:enumeration value=""InvalidArgument""/>
                         <xs:enumeration value=""InvalidUser""/>
                         <xs:enumeration value=""NotFound""/>
                         <xs:enumeration value=""FileRejected""/>
                         <xs:enumeration value=""UnknownError""/>
                       </xs:restriction>
                     </xs:simpleType>");
            #endregion schema of response result of SubmitFile
            return value;
        }
    }
}