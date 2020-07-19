namespace Microsoft.Protocols.TestSuites.MS_CPSWS
{
    using System;
    using System.ServiceModel;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// This class use to verify every request with the adapter when get a response. 
    /// </summary>
    public partial class MS_CPSWSAdapter : ManagedAdapterBase, IMS_CPSWSAdapter
    {
        /// <summary>
        /// A method used to validate the ArrayOfString complex type. 
        /// </summary>        
        private void ValidArrayOfStringComplexType()
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                45,
                @"[In ArrayOfString] ArrayOfString: An array of elements of type string.
<xs:complexType name=""ArrayOfString"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""unbounded"" name=""string"" nillable=""true"" type=""xs:string""/>
  </xs:sequence>
</xs:complexType>");
        }

        /// <summary>
        /// A method used to validate the ClaimTypes response. 
        /// </summary>
        /// <param name="claimTypesResult">A parameter represents the ClaimTypes result.</param> 
        private void ValidateClaimTypesResponseData(ArrayOfString claimTypesResult)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                141,
                @"[In ClaimTypes] The following is the WSDL port type specification of the ClaimTypes WSDL operation.
<wsdl:operation name=""ClaimTypes"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input wsaw:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/ClaimTypes"" message=""tns:IClaimProviderWebService_ClaimTypes_InputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
  <wsdl:output wsaw:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/ClaimTypesResponse"" message=""tns:IClaimProviderWebService_ClaimTypes_OutputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
</wsdl:operation>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                161,
                @"[In ClaimTypes] The protocol client sends an IClaimProviderWebService_ClaimTypes_InputMessage (section 3.1.4.1.1.1) request WSDL message and the protocol server responds with an IClaimProviderWebService_ClaimTypes_OutputMessage (section 3.1.4.1.1.2) response message.");

            // If the server response is validated successfully, and the ClaimTypesResponse has returned, MS-CPSWS_R152 can be verified.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(xmlResponse, "ClaimTypesResponse"),
                152,
                @"[In IClaimProviderWebService_ClaimTypes_OutputMessage] The [ClaimTypes] SOAP body contains the ClaimTypesResponse element.");

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                156,
                @"[In ClaimTypesResponse] The ClaimTypesResponse element specifies the result data for the ClaimTypes WSDL operation.
<xs:element name=""ClaimTypesResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexType>
    <xs:sequence>
      <xs:element minOccurs=""0"" maxOccurs=""1"" name=""ClaimTypesResult"" type=""tns:ArrayOfString""/>
    </xs:sequence>
  </xs:complexType>
</xs:element>");

            if (claimTypesResult != null)
            {
                this.ValidArrayOfStringComplexType();
            }
            //Verify MS-CPSWS_R157001
            foreach (string claim in claimTypesResult)
            {
                Site.CaptureRequirementIfIsTrue(
                    Uri.IsWellFormedUriString(claim, UriKind.Absolute),
                    157001,
                    @"[In ClaimTypesResponse] The claim type SHOULD format as a URI.");
            }
        }

        /// <summary>
        /// A method used to validate the ClaimValueTypes response. 
        /// </summary>
        /// <param name="claimValueTypesResult">A parameter represents the ClaimValueTypes result.</param> 
        private void ValidateClaimValueTypesResponseData(ArrayOfString claimValueTypesResult)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                159,
                @"[In ClaimValueTypes] The following is the WSDL port type specification of the ClaimValueTypes WSDL operation.
<wsdl:operation name=""ClaimValueTypes"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input wsaw:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/ClaimValueTypes"" message=""tns:IClaimProviderWebService_ClaimValueTypes_InputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
  <wsdl:output wsaw:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/ClaimValueTypesResponse"" message=""tns:IClaimProviderWebService_ClaimValueTypes_OutputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
</wsdl:operation>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                162,
                @"[In ClaimValueTypes] The protocol client sends an IClaimProviderWebService_ClaimValueTypes_InputMessage (section 3.1.4.2.1.1) request WSDL message and the protocol server responds with an IClaimProviderWebService_ClaimValueTypes_OutputMessage (section 3.1.4.2.1.2) response message.");

            // If the server response is validated successfully, and the ClaimValueTypesResponse has returned, MS-CPSWS_R171 can be verified.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(xmlResponse, "ClaimValueTypesResponse"),
                171,
                @"[In IClaimProviderWebService_ClaimValueTypes_OutputMessage] The [ClaimValueTypes] SOAP body contains the ClaimValueTypesResponse element.");

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                173,
                @"[In ClaimValueTypesResponse] The ClaimValueTypesResponse element specifies the result data for the ClaimValueTypes WSDL operation.
<xs:element name=""ClaimValueTypesResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexType>
    <xs:sequence>
      <xs:element minOccurs=""0"" maxOccurs=""1"" name=""ClaimValueTypesResult"" type=""tns:ArrayOfString""/>
    </xs:sequence>
  </xs:complexType>
</xs:element>");

            if (claimValueTypesResult != null)
            {
                this.ValidArrayOfStringComplexType();
            }
            //Verify MS-CPSWS_R174001
            foreach (string claim in claimValueTypesResult)
            {
                Site.CaptureRequirementIfIsTrue(
                    Uri.IsWellFormedUriString(claim, UriKind.Absolute),
                    174001,
                    @"[In ClaimValueTypesResponse] The claim value type SHOULD format as a URI. ");
            }
        }

        /// <summary>
        /// A method used to validate the EntityTypes response. 
        /// </summary>
        /// <param name="entityTypesResult">A parameter represents the EntityTypes result.</param> 
        private void ValidateEntityTypesResponseData(ArrayOfString entityTypesResult)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                176,
                @"[In EntityTypes] The following is the WSDL port type specification of the EntityTypes WSDL operation.
<wsdl:operation name=""EntityTypes"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input wsaw:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/EntityTypes"" message=""tns:IClaimProviderWebService_EntityTypes_InputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
  <wsdl:output wsaw:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/EntityTypesResponse"" message=""tns:IClaimProviderWebService_EntityTypes_OutputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
</wsdl:operation>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                177,
                @"[In EntityTypes] The protocol client sends an IClaimProviderWebService_EntityTypes_InputMessage (section 3.1.4.3.1.1) request WSDL message and the protocol server responds with an IClaimProviderWebService_EntityTypes_OutputMessage (section 3.1.4.3.1.2) response message.");

            // If the server response is validated successfully, and the EntityTypesResponse has returned, MS-CPSWS_R186 can be verified.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(xmlResponse, "EntityTypesResponse"),
                186,
                @"[In IClaimProviderWebService_EntityTypes_OutputMessage] The [EntityTypes] SOAP body contains the EntityTypesResponse element.");

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                189,
                @"[In EntityTypesResponse] The EntityTypesResponse element specifies the result data for the EntityTypes WSDL operation.
<xs:element name=""EntityTypesResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexType>
    <xs:sequence>
      <xs:element minOccurs=""0"" maxOccurs=""1"" name=""EntityTypesResult"" type=""tns:ArrayOfString""/>
    </xs:sequence>
  </xs:complexType>
</xs:element>");

            if (entityTypesResult != null)
            {
                this.ValidArrayOfStringComplexType();
            }
        }

        /// <summary>
        /// A method used to validate the GetHierarchy response. 
        /// </summary>
        /// <param name="getHierarchyResult">A parameter represents the GetHierarchy result.</param> 
        private void ValidateGetHierarchyResponseData(SPProviderHierarchyTree getHierarchyResult)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                192,
                @"[In GetHierarchy] The following is the WSDL port type specification of the GetHierarchy WSDL operation.
<wsdl:operation name=""GetHierarchy"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input wsaw:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/GetHierarchy"" message=""tns:IClaimProviderWebService_GetHierarchy_InputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
  <wsdl:output wsaw:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/GetHierarchyResponse"" message=""tns:IClaimProviderWebService_GetHierarchy_OutputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
</wsdl:operation>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                193,
                @"[In GetHierarchy] The protocol client sends an IClaimProviderWebService_GetHierarchy_InputMessage (section 3.1.4.4.1.1) request WSDL message and the protocol server responds with an IClaimProviderWebService_GetHierarchy_OutputMessage (section 3.1.4.4.1.2) response message.");

            // If the server response is validated successfully, and the GetHierarchyResponse has returned, MS-CPSWS_R202 can be verified.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(xmlResponse, "GetHierarchyResponse"),
                202,
                @"[In IClaimProviderWebService_GetHierarchy_OutputMessage] The [GetHierarchy] SOAP body contains the GetHierarchyResponse element.");

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                213,
                @"[In GetHierarchyResponse] The GetHierarchyResponse element specifies the result data for the GetHierarchy WSDL operation.
<xs:element name=""GetHierarchyResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexType>
    <xs:sequence>
      <xs:element minOccurs=""0"" maxOccurs=""1"" name=""GetHierarchyResult"" type=""tns:SPProviderHierarchyTree""/>
    </xs:sequence>
  </xs:complexType>
</xs:element>");

            if (getHierarchyResult != null)
            {
                this.VerifySPProviderHierarchyTreeComplexType(getHierarchyResult);
            }
        }

        /// <summary>
        /// A method used to validate the GetHierarchyAll response. 
        /// </summary>
        /// <param name="getHierarchyAllResult">A parameter represents the GetHierarchyAll result.</param> 
        private void ValidateGetHierarchyAllResponseData(SPProviderHierarchyTree[] getHierarchyAllResult)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;
            bool isPrefix = false;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                217,
                @"[In GetHierarchyAll] The following is the WSDL port type specification of the GetHierarchyAll WSDL operation.
<wsdl:operation name=""GetHierarchyAll"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input wsaw:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/GetHierarchyAll"" message=""tns:IClaimProviderWebService_GetHierarchyAll_InputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
  <wsdl:output wsaw:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/GetHierarchyAllResponse"" message=""tns:IClaimProviderWebService_GetHierarchyAll_OutputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
</wsdl:operation>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                218,
                @"[In GetHierarchyAll] The protocol client sends an IClaimProviderWebService_GetHierarchyAll_InputMessage (section 3.1.4.5.1.1) request WSDL message and the protocol server responds with an IClaimProviderWebService_GetHierarchyAll_OutputMessage (section 3.1.4.5.1.2) response message.");

            // If the server response is validated successfully, and the GetHierarchyAllResponse has returned, MS-CPSWS_R228 can be verified.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(xmlResponse, "GetHierarchyAllResponse"),
                228,
                @"[In IClaimProviderWebService_GetHierarchyAll_OutputMessage] The [GetHierarchyAll]  SOAP body contains the GetHierarchyAllResponse element.");

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                237,
                @"[In GetHierarchyAllResponse] The GetHierarchyAllResponse element specifies the result data for the GetHierarchyAll WSDL operation.
<xs:element name=""GetHierarchyAllResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexType>
    <xs:sequence>
      <xs:element minOccurs=""0"" maxOccurs=""1"" name=""GetHierarchyAllResult"" type=""tns:ArrayOfSPProviderHierarchyTree""/>
    </xs:sequence>
  </xs:complexType>
</xs:element>");

            if (getHierarchyAllResult.Length != 0)
            {
                this.VerifyArrayOfSPProviderHierarchyTreeComplexType(getHierarchyAllResult);
            }
            string namePrefixed = Common.GetConfigurationPropertyValue("HierarchyProviderPrefix", this.Site);
            foreach (SPProviderHierarchyTree element in getHierarchyAllResult)
            {
                if (element.ProviderName.StartsWith(namePrefixed))
                {
                    isPrefix = true;
                    break;
                }
            }
           Site.CaptureRequirementIfIsTrue(
                isPrefix,
                584001,
                @"[In GetHierarchyAll] The name of the hierarchy provider is prefixed with ""_HierarchyProvider_"".");
        }

        /// <summary>
        /// A method used to validate the resolve operation response. 
        /// </summary>
        /// <param name="resolveResult">A parameter represents the Resolve operation result.</param>       
        private void ValidateResolveResponseData(PickerEntity[] resolveResult)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                278,
                @"[In Resolve] The following is the WSDL port type specification of the Resolve WSDL operation.
<wsdl:operation name=""Resolve"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/Resolve"" message=""tns:IClaimProviderWebService_Resolve_InputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
  <wsdl:output wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/ResolveResponse"" message=""tns:IClaimProviderWebService_Resolve_OutputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
</wsdl:operation>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                279,
                @"[In Resolve] The protocol client sends an IClaimProviderWebService_Resolve_InputMessage (section 3.1.4.8.1.1) request WSDL message and the protocol server responds with an IClaimProviderWebService_Resolve_OutputMessage (section 3.1.4.8.1.2) response message.");

            // If the server response is validated successfully, and the ResolveResponse has returned, MS-CPSWS_R289 can be verified.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(xmlResponse, "ResolveResponse"),
                289,
                @"[In IClaimProviderWebService_Resolve_OutputMessage] The SOAP body contains the ResolveResponse element.");

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                298,
                @"[In ResolveResponse] The ResolveResponse element specifies the result data for the Resolve WSDL operation.
<xs:element name=""ResolveResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexType>
    <xs:sequence>
      <xs:element minOccurs=""0"" maxOccurs=""1"" name=""ResolveResult"" type=""tns:ArrayOfPickerEntity""/>
    </xs:sequence>
  </xs:complexType>
</xs:element>");

            if (resolveResult.Length != 0)
            {
                this.ValidateArrayOfPickerEntityComplexType();
            }
        }

        /// <summary>
        /// A method used to validate the resolve claim response. 
        /// </summary>
        /// <param name="resolveClaimResult">A parameter represents the resolve claim result.</param>       
        private void ValidateResolveClaimResponseData(PickerEntity[] resolveClaimResult)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                301,
                @"[In ResolveClaim] The following is the WSDL port type specification of the ResolveClaim WSDL operation.
<wsdl:operation name=""ResolveClaim"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/ResolveClaim"" message=""tns:IClaimProviderWebService_ResolveClaim_InputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
  <wsdl:output wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/ResolveClaimResponse"" message=""tns:IClaimProviderWebService_ResolveClaim_OutputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
</wsdl:operation>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                302,
                @"[In ResolveClaim] The protocol client sends an IClaimProviderWebService_ResolveClaim_InputMessage (section 3.1.4.9.1.1) request WSDL message and the protocol server responds with an IClaimProviderWebService_ResolveClaim_OutputMessage (section 3.1.4.9.1.2) response message.");

            // If the server response is validated successfully, and the ResolveClaimResponse has returned, MS-CPSWS_R312 can be verified.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(xmlResponse, "ResolveClaimResponse"),
                312,
                @"[In IClaimProviderWebService_ResolveClaim_OutputMessage] The SOAP body contains the ResolveClaimResponse element.");

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                320,
                @"[In ResolveClaimResponse] The ResolveClaimResponse element specifies the result data for the ResolveClaim WSDL operation.
<xs:element name=""ResolveClaimResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexType>
    <xs:sequence>
      <xs:element minOccurs=""0"" maxOccurs=""1"" name=""ResolveClaimResult"" type=""tns:ArrayOfPickerEntity""/>
    </xs:sequence>
  </xs:complexType>
</xs:element>");

            if (resolveClaimResult.Length != 0)
            {
                this.ValidateArrayOfPickerEntityComplexType();
            }
        }

        /// <summary>
        /// A method used to validate the resolve multiple response. 
        /// </summary>
        /// <param name="resolveMultipleResult">A parameter represents the resolve multiple result.</param>       
        private void ValidateResolveMultipleResponseData(PickerEntity[] resolveMultipleResult)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                323,
                @"[In ResolveMultiple] The following is the WSDL port type specification of the ResolveMultiple WSDL operation.
<wsdl:operation name=""ResolveMultiple"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/ResolveMultiple"" message=""tns:IClaimProviderWebService_ResolveMultiple_InputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
  <wsdl:output wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/ResolveMultipleResponse"" message=""tns:IClaimProviderWebService_ResolveMultiple_OutputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
</wsdl:operation>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                324,
                @"[In ResolveMultiple] The protocol client sends an IClaimProviderWebService_ResolveMultiple_InputMessage (section 3.1.4.10.1.1) request WSDL message and the protocol server responds with an IClaimProviderWebService_ResolveMultiple_OutputMessage (section 3.1.4.10.1.2) response message.");

            // If the server response is validated successfully, and the ResolveMultipleResponse has returned, MS-CPSWS_R334 can be verified.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(xmlResponse, "ResolveMultipleResponse"),
                334,
                @"[In IClaimProviderWebService_ResolveMultiple_OutputMessage] The SOAP body contains the ResolveMultipleResponse element.");

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                343,
                @"[In ResolveMultipleResponse] The ResolveMultipleResponse element specifies the result data for the ResolveMultiple WSDL operation.
<xs:element name=""ResolveMultipleResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexType>
    <xs:sequence>
      <xs:element minOccurs=""0"" maxOccurs=""1"" name=""ResolveMultipleResult"" type=""tns:ArrayOfPickerEntity""/>
    </xs:sequence>
  </xs:complexType>
</xs:element>");

            if (resolveMultipleResult.Length != 0)
            {
                this.ValidateArrayOfPickerEntityComplexType();
            }
        }

        /// <summary>
        /// A method used to validate the resolve multiple claim response. 
        /// </summary>
        /// <param name="resolveMultipleClaimResult">A parameter represents the resolve multiple claim result.</param>       
        private void ValidateResolveMultipleClaimResponseData(PickerEntity[] resolveMultipleClaimResult)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                348,
                @"[In ResolveMultipleClaim] The following is the WSDL port type specification of the ResolveMultipleClaim WSDL operation.
<wsdl:operation name=""ResolveMultipleClaim"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/ResolveMultipleClaim"" message=""tns:IClaimProviderWebService_ResolveMultipleClaim_InputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
  <wsdl:output wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/ResolveMultipleClaimResponse"" message=""tns:IClaimProviderWebService_ResolveMultipleClaim_OutputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
</wsdl:operation>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                349,
                @"[In ResolveMultipleClaim] The protocol client sends an IClaimProviderWebService_ResolveMultipleClaim_InputMessage (section 3.1.4.11.1.1) request WSDL message and the protocol server responds with an IClaimProviderWebService_ResolveMultipleClaim_OutputMessage (section 3.1.4.11.1.2) response message.");

            // If the server response is validated successfully, and the ResolveMultipleClaimResponse has returned, MS-CPSWS_R359 can be verified.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(xmlResponse, "ResolveMultipleClaimResponse"),
                359,
                @"[In IClaimProviderWebService_ResolveMultipleClaim_OutputMessage] The SOAP body contains the ResolveMultipleClaimResponse element.");

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                367,
                @"[In ResolveMultipleClaimResponse] The ResolveMultipleClaimResponse element specifies the result data for the ResolveMultipleClaim WSDL operation.
<xs:element name=""ResolveMultipleClaimResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexType>
    <xs:sequence>
      <xs:element minOccurs=""0"" maxOccurs=""1"" name=""ResolveMultipleClaimResult"" type=""tns:ArrayOfPickerEntity""/>
    </xs:sequence>
  </xs:complexType>
</xs:element>");

            if (resolveMultipleClaimResult.Length != 0)
            {
                this.ValidateArrayOfPickerEntityComplexType();
            }
        }

        /// <summary>
        /// A method used to validate ArrayOfAnyType complex type.
        /// </summary>        
        private void ValidateArrayOfAnyTypeComplexType()
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                29,
                @"[In ArrayOfAnyType] ArrayOfAnyType: An array of elements of any type.
                    <xs:complexType name=""ArrayOfAnyType"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""unbounded"" name=""anyType"" nillable=""true""/>
  </xs:sequence>
</xs:complexType>
");
        }

        /// <summary>
        /// A method used to validate ArrayOfPair complex type.
        /// </summary>        
        private void ValidateArrayOfPairComplexType()
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                33,
                @"[In ArrayOfPair] ArrayOfPair: An array of elements of type Pair (section 2.2.4.8).
<xs:complexType name=""ArrayOfPair"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""unbounded"" name=""Pair"" nillable=""true"" type=""tns:Pair""/>
  </xs:sequence>
</xs:complexType>");

            this.ValidatePairComplexType();
        }

        /// <summary>
        /// A method used to validate ArrayOfPickerEntity complex type.
        /// </summary>        
        private void ValidateArrayOfPickerEntityComplexType()
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                36,
                @"[In ArrayOfPickerEntity] ArrayOfPickerEntity: An array of elements of type PickerEntity (section 2.2.4.9).
<xs:complexType name=""ArrayOfPickerEntity"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""unbounded"" name=""PickerEntity"" nillable=""true"" type=""tns:PickerEntity""/>
  </xs:sequence>
</xs:complexType>");

            this.ValidatePickerEntityComplexType();
        }

        /// <summary>
        /// A method used to validate Pair complex type.
        /// </summary>        
        private void ValidatePairComplexType()
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                48,
                @"[In Pair] Pair: A collection of two named elements of any type in a sequence.
<xs:complexType name=""Pair"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""First""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""Second""/>
  </xs:sequence>
</xs:complexType>");
        }

        /// <summary>
        /// A method used to validate PickerEntity complex type.
        /// </summary>        
        private void ValidatePickerEntityComplexType()
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // The response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                52,
                @"[In PickerEntity] PickerEntity: An object containing basic information about a picker entity.
<xs:complexType name=""PickerEntity"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""Key"" type=""xs:string""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""DisplayText"" type=""xs:string""/>
    <xs:element minOccurs=""1"" maxOccurs=""1"" name=""IsResolved"" type=""xs:boolean""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""Description"" type=""xs:string""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""EntityType"" type=""xs:string""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""EntityGroupName"" type=""xs:string""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""HierarchyIdentifier""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""EntityDataElements"" type=""tns:ArrayOfPair""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""MultipleMatches"" type=""tns:ArrayOfAnyType""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""ProviderName"" type=""xs:string""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""ProviderDisplayName"" type=""xs:string""/>
  </xs:sequence>
</xs:complexType>");

            this.ValidateArrayOfPairComplexType();
            this.ValidateArrayOfAnyTypeComplexType();
        }

        /// <summary>
        /// A method used to verify the SPProviderHierarchyTree complex type.
        /// </summary>
        /// <param name="providerHierarchyTree">A parameter represents an instance of SPProviderHierarchyTree.</param>
        private void VerifySPProviderHierarchyTreeComplexType(SPProviderHierarchyTree providerHierarchyTree)
        {
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;
            if (providerHierarchyTree != null)
            {
                // If the response passed XML validation, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    71,
                    @"[In SPProviderHierarchyTree] A claim provider hierarchy tree.
<xs:complexType name=""SPProviderHierarchyTree"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexContent mixed=""false"">
    <xs:extension base=""tns:SPProviderHierarchyElement"">
      <xs:sequence>
        <xs:element minOccurs=""1"" maxOccurs=""1"" name=""IsRoot"" type=""xs:boolean""/>
      </xs:sequence>
    </xs:extension>
  </xs:complexContent>
</xs:complexType>");

                // If the response passed XML validation, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    537,
                    @"[In SPProviderHierarchyElement] SPProviderHierarchyElement: Defines the base type for SPProviderHierarchyNode and SPProviderHierarchyTree.
<xs:complexType name=""SPProviderHierarchyElement"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""Nm"" type=""xs:string""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""ProviderName"" type=""xs:string""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""HierarchyNodeID"" type=""xs:string""/>
    <xs:element minOccurs=""1"" maxOccurs=""1"" name=""IsLeaf"" type=""xs:boolean""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""Children"" type=""tns:ArrayOfSPProviderHierarchyNode""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""EntityData"" type=""tns:ArrayOfPickerEntity""/>
    <xs:element minOccurs=""1"" maxOccurs=""1"" name=""Count"" type=""xs:int""/>
  </xs:sequence>
</xs:complexType>");
                if (providerHierarchyTree.Children != null)
                {
                    // If the response passed XML validation, then the following requirement can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        516,
                        @"[In ArrayOfSPProviderHierarchyNode] An array of elements of type SPProviderHierarchyNode (section 2.2.4.12).
 <xs:complexType name=""ArrayOfSPProviderHierarchyNode"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""unbounded"" name=""SPProviderHierarchyNode"" nillable=""true"" type=""tns:SPProviderHierarchyNode""/>
  </xs:sequence>
</xs:complexType>");

                    // If the response passed XML validation, then the following requirement can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        516,
                        @"[In ArrayOfSPProviderHierarchyNode] An array of elements of type SPProviderHierarchyNode (section 2.2.4.12).
 <xs:complexType name=""ArrayOfSPProviderHierarchyNode"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""unbounded"" name=""SPProviderHierarchyNode"" nillable=""true"" type=""tns:SPProviderHierarchyNode""/>
  </xs:sequence>
</xs:complexType>");
                    if (providerHierarchyTree.Children.Length > 0)
                    {
                        // If the response passed XML validation, then the following requirement can be captured.
                        Site.CaptureRequirementIfIsTrue(
                            isResponseValid,
                            549,
                            @"[In SPProviderHierarchyNode] A claims provider hierarchy node.
<xs:complexType name=""SPProviderHierarchyNode"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexContent mixed=""false"">
    <xs:extension base=""tns:SPProviderHierarchyElement""/>
  </xs:complexContent>
</xs:complexType>");
                    }
                }
            }
        }

        /// <summary>
        /// A method used to verify the ArrayOfSPProviderHierarchyTree complex type.
        /// </summary>
        /// <param name="providerHierarchyTreeArray">A parameter represents a list of the SPProviderHierarchyTree.</param>
        private void VerifyArrayOfSPProviderHierarchyTreeComplexType(SPProviderHierarchyTree[] providerHierarchyTreeArray)
        {
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;
            if (providerHierarchyTreeArray != null)
            {
                // If the response passed XML validation, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    39,
                    @"[In ArrayOfSPProviderHierarchyTree] ArrayOfSPProviderHierarchyTree: An array of elements of type SPProviderHierarchyTree (section 2.2.4.13).
<xs:complexType name=""ArrayOfSPProviderHierarchyTree"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""unbounded"" name=""SPProviderHierarchyTree"" nillable=""true"" type=""tns:SPProviderHierarchyTree""/>
  </xs:sequence>
</xs:complexType>");

                if (providerHierarchyTreeArray.Length > 0)
                {
                    foreach (SPProviderHierarchyTree providerHierarchyTree in providerHierarchyTreeArray)
                    {
                        this.VerifySPProviderHierarchyTreeComplexType(providerHierarchyTree);
                    }
                }
            }
        }

        /// <summary>
        /// A method used to verify the SPProviderSchema complex type.
        /// </summary>
        /// <param name="providerSchema">A parameter represents an instance of the SPProviderSchema complex type.</param>
        private void VerifySPProviderSchemaComplexType(SPProviderSchema providerSchema)
        {
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            if (providerSchema != null)
            {
                // If the response passed XML validation, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    74,
                    @"[In SPProviderSchema] The user interface display characteristics of a claims provider.
<xs:complexType name=""SPProviderSchema"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""DisplayName"" type=""xs:string""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""ProviderName"" type=""xs:string""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""ProviderSchema"" type=""tns:ArrayOfSPSchemaElement""/>
    <xs:element minOccurs=""1"" maxOccurs=""1"" name=""SupportsHierarchy"" type=""xs:boolean""/>
  </xs:sequence>
</xs:complexType>");

                if (providerSchema.ProviderSchema != null)
                {
                    // If the response passed XML validation, then the following requirement can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        99,
                        @"[In SPSchemaElementType] The display type of a field.
<xs:simpleType name=""SPSchemaElementType"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:restriction base=""xs:string"">
    <xs:enumeration value=""None""/>
    <xs:enumeration value=""TableViewOnly""/>
    <xs:enumeration value=""DetailViewOnly""/>
    <xs:enumeration value=""Both""/>
  </xs:restriction>
</xs:simpleType>");

                    // If the response passed XML validation, then the following requirement can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        42,
                        @"[In ArrayOfSPSchemaElement] ArrayOfSPSchemaElement: An array of SPSchemaElement (section 2.2.4.15) elements.
<xs:complexType name=""ArrayOfSPSchemaElement"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""unbounded"" name=""SPSchemaElement"" nillable=""true"" type=""tns:SPSchemaElement""/>
  </xs:sequence>
</xs:complexType>");

                    if (providerSchema.ProviderSchema.Length > 0)
                    {
                        this.VerifySPSchemaElementComplexType(isResponseValid);
                    }
                }
            }
        }

        /// <summary>
        /// A method used to verify the SPSchemaElement complex type.
        /// </summary>
        /// <param name="isResponseValid">A parameter represents whether the server response passed XML validation.</param>
        private void VerifySPSchemaElementComplexType(bool isResponseValid)
        {
            // If the response passed XML validation, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                83,
                @"[In SPSchemaElement] The user interface display characteristics of a field in a picker entity.
<xs:complexType name=""SPSchemaElement"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""Name"" type=""xs:string""/>
    <xs:element minOccurs=""0"" maxOccurs=""1"" name=""DisplayName"" type=""xs:string""/>
    <xs:element minOccurs=""1"" maxOccurs=""1"" name=""Type"" type=""tns:SPSchemaElementType""/>
  </xs:sequence>
</xs:complexType>");

            // If the response passed XML validation, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                99,
                @"[In SPSchemaElementType] The display type of a field.
<xs:simpleType name=""SPSchemaElementType"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:restriction base=""xs:string"">
    <xs:enumeration value=""None""/>
    <xs:enumeration value=""TableViewOnly""/>
    <xs:enumeration value=""DetailViewOnly""/>
    <xs:enumeration value=""Both""/>
  </xs:restriction>
</xs:simpleType>");
        }

        /// <summary>
        /// A method used to verify the HierarchyProviderSchema response. 
        /// </summary>
        /// <param name="hierarchyProviderSchemaResult">A parameter represents the HierarchyProviderSchema result.</param> 
        private void VerifyHierarchyProviderSchema(SPProviderSchema hierarchyProviderSchemaResult)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            if (this.ResponseExists(xmlResponse, "HierarchyProviderSchemaResponse"))
            {
                // The response have been received successfully, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    242,
                    @"[In HierarchyProviderSchema] The following is the WSDL port type specification of the HierarchyProviderSchema WSDL operation.

<wsdl:operation name=""HierarchyProviderSchema"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/HierarchyProviderSchema"" message=""tns:IClaimProviderWebService_HierarchyProviderSchema_InputMessage""/>
  <wsdl:output wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/HierarchyProviderSchemaResponse"" message=""tns:IClaimProviderWebService_HierarchyProviderSchema_OutputMessage""/>
</wsdl:operation>");

                // The response have been received successfully, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    243,
                    @"[In HierarchyProviderSchema] The protocol client sends an IClaimProviderWebService_HierarchyProviderSchema_InputMessage (section 3.1.4.6.1.1) request WSDL message and the protocol server responds with an IClaimProviderWebService_HierarchyProviderSchema_OutputMessage (section 3.1.4.6.1.2) response message.");

                // If the response is contained in the response raw XML, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    this.ResponseExists(xmlResponse, "HierarchyProviderSchemaResponse"),
                    253,
                    @"[In IClaimProviderWebService_HierarchyProviderSchema_OutputMessage] The SOAP body contains the HierarchyProviderSchemaResponse element.");

                if (hierarchyProviderSchemaResult != null)
                {
                    // If the response passed the XML validation, then the following requirement can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        255,
                        @"[In HierarchyProviderSchema] The HierarchyProviderSchemaResponse element specifies the result data for the HierarchyProviderSchema WSDL operation.
<xs:element name=""HierarchyProviderSchemaResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexType>
    <xs:sequence>
      <xs:element minOccurs=""0"" maxOccurs=""1"" name=""HierarchyProviderSchemaResult"" type=""tns:SPProviderSchema""/>
    </xs:sequence>
  </xs:complexType>
</xs:element>");

                    this.VerifySPProviderSchemaComplexType(hierarchyProviderSchemaResult);
                }
            }
        }

        /// <summary>
        /// A method used to verify the ProviderSchemas response. 
        /// </summary>
        /// <param name="providerSchemasResult">A parameter represents the ProviderSchemas result.</param> 
        private void VerifyProviderSchemas(SPProviderSchema[] providerSchemasResult)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            if (this.ResponseExists(xmlResponse, "ProviderSchemasResponse"))
            {
                // The response have been received successfully, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    258,
                    @"[In ProviderSchemas] The following is the WSDL port type specification of the ProviderSchemas WSDL operation.
<wsdl:operation name=""ProviderSchemas""  xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/ProviderSchemas"" message=""tns:IClaimProviderWebService_ProviderSchemas_InputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
  <wsdl:output wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/ProviderSchemasResponse"" message=""tns:IClaimProviderWebService_ProviderSchemas_OutputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
</wsdl:operation>");

                // The response have been received successfully, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    259,
                    @"[In ProviderSchemas] The protocol client sends an IClaimProviderWebService_ProviderSchemas_InputMessage (section 3.1.4.7.1.1) request WSDL message and the protocol server responds with an IClaimProviderWebService_ProviderSchemas_OutputMessage (section 3.1.4.7.1.2) response message.");
                
                // If the response is contained in the response raw XML, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    this.ResponseExists(xmlResponse, "ProviderSchemasResponse"),
                    269,
                    @"[In IClaimProviderWebService_ProviderSchemas_OutputMessage] The SOAP body contains the ProviderSchemasResponse element.");

                if (providerSchemasResult != null)
                {
                    // If the response passed the XML validation, then the following requirement can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        272,
                        @"[In ProviderSchemasResponse] The ProviderSchemasResponse element specifies the result data for the ProviderSchemas WSDL operation.
<xs:element name=""ProviderSchemasResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexType>
    <xs:sequence>
      <xs:element minOccurs=""0"" maxOccurs=""1"" name=""ProviderSchemasResult"" type=""tns:ArrayOfSPProviderSchema""/>
    </xs:sequence>
  </xs:complexType>
</xs:element>");

                    // If the response passed the XML validation, then the following requirement can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        275,
                        @"[In ArrayOfSPProviderSchema] An array of SPProviderSchema (section 2.2.4.14) elements.
<xs:complexType name=""ArrayOfSPProviderSchema"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:sequence>
    <xs:element minOccurs=""0"" maxOccurs=""unbounded"" name=""SPProviderSchema"" nillable=""true"" type=""tns:SPProviderSchema""/>
  </xs:sequence>
</xs:complexType>");

                    if (providerSchemasResult.Length > 0)
                    {
                        foreach (SPProviderSchema providerSchema in providerSchemasResult)
                        {
                            this.VerifySPProviderSchemaComplexType(providerSchema);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// A method used to verify the Search response. 
        /// </summary>
        /// <param name="searchResult">A parameter represents the Search result.</param> 
        private void VerifySearch(SPProviderHierarchyTree[] searchResult)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            if (this.ResponseExists(xmlResponse, "SearchResponse"))
            {
                // The response have been received successfully, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    376,
                    @"[In Search] The following is the WSDL port type specification of the Search WSDL operation.
<wsdl:operation name=""Search"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/Search"" message=""tns:IClaimProviderWebService_Search_InputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
  <wsdl:output wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/SearchResponse"" message=""tns:IClaimProviderWebService_Search_OutputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
</wsdl:operation>");

                // The response have been received successfully, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    377,
                    @"[In Search] The protocol client sends an IClaimProviderWebService_Search_InputMessage (section 3.1.4.12.1.1) request WSDL message and the protocol server responds with an IClaimProviderWebService_Search_OutputMessage (section 3.1.4.12.1.2) response message.");

                // If the response is contained in the response raw XML, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    this.ResponseExists(xmlResponse, "SearchResponse"),
                    387,
                    @"[In IClaimProviderWebService_Search_OutputMessage] The SOAP body contains the SearchResponse element.");

                // The response have been received successfully, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    399,
                    @"[In SearchResponse] The SearchResponse element specifies the result data for the Search WSDL operation.
<xs:element name=""SearchResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexType>
    <xs:sequence>
      <xs:element minOccurs=""0"" maxOccurs=""1"" name=""SearchResult"" type=""tns:ArrayOfSPProviderHierarchyTree""/>
    </xs:sequence>
  </xs:complexType>
</xs:element>");

                if (searchResult != null)
                {
                    this.VerifyArrayOfSPProviderHierarchyTreeComplexType(searchResult);
                }
            }
        }

        /// <summary>
        /// A method used to verify the SearchAll response. 
        /// </summary>
        /// <param name="searchAllResult">A parameter represents the SearchAll result.</param> 
        private void VerifySearchAll(SPProviderHierarchyTree[] searchAllResult)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            if (this.ResponseExists(xmlResponse, "SearchAllResponse"))
            {
                // The response have been received successfully, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    415,
                    @"[In SearchAll] The following is the WSDL port type specification of the SearchAll WSDL operation.
<wsdl:operation name=""SearchAll"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
  <wsdl:input wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/SearchAll"" message=""tns:IClaimProviderWebService_SearchAll_InputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
  <wsdl:output wsam:Action=""http://schemas.microsoft.com/sharepoint/claims/IClaimProviderWebService/SearchAllResponse"" message=""tns:IClaimProviderWebService_SearchAll_OutputMessage"" xmlns:wsaw=""http://www.w3.org/2006/05/addressing/wsdl""/>
</wsdl:operation>");

                // The response have been received successfully, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    416,
                    @"[In SearchAll] The protocol client sends an IClaimProviderWebService_SearchAll_InputMessage (section 3.1.4.13.1.1) request WSDL message and the protocol server responds with an IClaimProviderWebService_SearchAll_OutputMessage (section 3.1.4.13.1.2) response message.");

                // If the response is contained in the response raw XML, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    this.ResponseExists(xmlResponse, "SearchAllResponse"),
                    426,
                    @"[In IClaimProviderWebService_SearchAll_OutputMessage] The SOAP body contains the SearchAllResponse element.");

                // The response have been received successfully, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    436,
                    @"[In SearchAllResponse] The SearchAllResponse element specifies the result data for the SearchAll WSDL operation.
<xs:element name=""SearchAllResponse"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:complexType>
    <xs:sequence>
      <xs:element minOccurs=""0"" maxOccurs=""1"" name=""SearchAllResult"" type=""tns:ArrayOfSPProviderHierarchyTree""/>
    </xs:sequence>
  </xs:complexType>
</xs:element>");

                if (searchAllResult != null)
                {
                    this.VerifyArrayOfSPProviderHierarchyTreeComplexType(searchAllResult);
                }
            }
        }

        /// <summary>
        /// This method is used to validate whether there is an element named as responseName in xmlElement.
        /// </summary>
        /// <param name="xmlElement">A XmlElement object which is the raw XML response from server.</param>
        /// <param name="responseName">The name of the response element.</param>
        /// <returns>A Boolean value indicates whether there is an element named as responseName.</returns>
        private bool ResponseExists(XmlElement xmlElement, string responseName)
        {
            // The first child is the response element.
            XmlNode firstChildNode = xmlElement.ChildNodes[1].FirstChild;
            return firstChildNode.Name == responseName;
        }

        /// <summary>
        /// Capture underlying transport protocol related requirements.
        /// </summary>
        private void CaptureTransportRelatedRequirements()
        {
            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            switch (transport)
            {
                case TransportProtocol.HTTP:

                    // As response successfully returned, the transport related requirements can be captured.
                    Site.CaptureRequirement(
                        1,
                        @"[In Transport]Protocol servers MUST support SOAP over HTTP.");
                    break;

                case TransportProtocol.HTTPS:

                    if (Common.IsRequirementEnabled(595, this.Site))
                    {
                        // Having received the response successfully have proved the HTTPS 
                        // transport is supported. If the HTTPS transport is not supported, the 
                        // response can't be received successfully.
                        Site.CaptureRequirement(
                        595,
                        @"[In Appendix C: Product Behavior] Implementation does additionally support SOAP over HTTPS for securing communication with protocol clients.(Windows SharePoint Foundation 2010 and above products follow this behavior.)");
                    }

                    break;

                default:
                    Site.Debug.Fail("Unknown transport type " + transport);
                    break;
            }

            SoapProtocolVersion soapVersion = Common.GetConfigurationPropertyValue<SoapProtocolVersion>("SoapVersion", this.Site);

            // Verifies MS-CPSWS requirement: MS-CPSWS_R7.
            bool isR7Verified = soapVersion == SoapProtocolVersion.Soap11 || soapVersion == SoapProtocolVersion.Soap12;
            Site.CaptureRequirementIfIsTrue(
                isR7Verified,
                7,
                @"[In Transport] Protocol messages MUST be formatted as specified in [SOAP1.1] , section 4.");
        }

        /// <summary>
        /// Validate server fault returned by FaultException and capture related requirements.
        /// </summary>
        /// <param name="exception">The FaultException thrown which contains the SOAP fault.</param>
        private void ValidateAndCaptureSOAPFaultRequirement(FaultException exception)
        {
            Site.CaptureRequirementIfIsNotNull(
                exception,
                8,
                @"[In Transport] Protocol server faults MUST be returned [using HTTP Status-Codes as specified in [RFC2616] , section 10 or] using SOAP faults as specified in [SOAP1.1] , section 4.4 Scope[[#Headers],[Description]]or [SOAP1.2/1] , section 5.4.");
        }
    }
}