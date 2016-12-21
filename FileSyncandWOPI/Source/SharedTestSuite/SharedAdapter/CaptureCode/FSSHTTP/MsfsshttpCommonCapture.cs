namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class contains methods which capture common requirements.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related with UserNameTypes.
        /// </summary>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateUserNameTypes(ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R793 and MS-FSSHTTP_R1826
            // The UserNameTypes is derived from NCName, so if the NCName type validation passes, then MS-FSSHTTP_R793 and MS-FSSHTTP_R1826 can be captured.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     793,
                     @"[In UserNameType] The UserNameType simple type specifies a representation of a user name value as specified in [RFC2822].");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1826
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1826,
                     @"[In UserNameType][The schema of UserNameType is] 
                     <xs:simpleType name=""UserNameType"">
                         <xs:restriction base=""xs:string"">
                         </xs:restriction>
                     </xs:simpleType> ");
        }

        /// <summary>
        /// Capture underlying transport protocol related requirements. They can be captured directly when the server returns a SOAP response successfully.
        /// </summary>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateTransport(ITestSite site)
        {
            if (Common.GetConfigurationPropertyValue("TransportType", site).Equals("HTTP", StringComparison.CurrentCultureIgnoreCase))
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1
                // Since all the messages are transported following Http configured in BasicHttpBinding_ICellStorages, so MS-FSSHTTP_R1 can be captured.
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         1,
                         @"[In Transport] Protocol servers MUST support SOAP over HTTP, as specified in [RFC2616] [,or HTTPS, as specified in [RFC2818]].");
            }

            if (Common.GetConfigurationPropertyValue("TransportType", site).Equals("HTTPS", StringComparison.CurrentCultureIgnoreCase))
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11000
                // Since all the messages are transported following Https configured in BasicHttpBinding_ICellStorages, so MS-FSSHTTP_R11000 can be captured.
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         11000,
                         @"[In Transport] Protocol servers MUST support [SOAP over HTTP, as specified in [RFC2616], or] HTTPS, as specified in [RFC2818].");
            }

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2
            // Since all the messages are SOAP1.1 message configured in BasicHttpBinding_ICellStorages, so MS-FSSHTTP_R2 can be captured.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2,
                     @"[In Transport] Protocol messages MUST be formatted as specified in [SOAP1.1] section 4.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3
            // Since all the messages are SOAP1.1 message configured in BasicHttpBinding_ICellStorages, so MS-FSSHTTP_R3 can be captured.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     3,
                     @"[In Transport] Protocol server MUST use MTOM encoding as specified in [SOAP1.2-MTOM].");
        }

        /// <summary>
        /// Capture requirements related with RequestToken within Response element
        /// </summary>
        /// <param name="response">The Response information</param>
        /// <param name="expectedToken">The expected RequestToken</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateResponseToken(Response response, string expectedToken, ITestSite site)
        {
            if (expectedToken == null)
            {
                // When the expected token is null, then indicating there is no expected token value returned by server.
                return;
            }

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R65
            site.CaptureRequirementIfAreEqual<string>(
                     expectedToken,
                     response.RequestToken,
                     "MS-FSSHTTP",
                     65,
                     @"[In Request] The one-to-one mapping between the Response element and the Request element MUST be maintained by using RequestToken.");

            // Directly capture requirement MS-FSSHTTPB_R70, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     70,
                     @"[In Request] Depending on the other types of errors[GenericErrorCodeTypes, CellRequestErrorCodeTypes, DependencyCheckRelatedErrorCodeTypes, LockAndCoauthRelatedErrorCodeTypes and NewEditorsTableCategoryErrorCodeTypes], the error code for that type MUST be returned by the protocol server.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R929
            site.CaptureRequirementIfAreEqual<string>(
                     expectedToken,
                     response.RequestToken,
                     "MS-FSSHTTP",
                     929,
                     @"[In Common Message Processing Rules and Events][The protocol server MUST follow the following common processing rules for all types of subrequests] The protocol server sends a Response element for each Request element.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R95
            site.CaptureRequirementIfAreEqual<string>(
                     expectedToken,
                     response.RequestToken,
                     "MS-FSSHTTP",
                     95,
                     @"[In Response] For each Request element that is part of a cell storage service request, there MUST be a corresponding Response element in a cell storage service response.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R105
            site.CaptureRequirementIfAreEqual<string>(
                     expectedToken,
                     response.RequestToken,
                     "MS-FSSHTTP",
                     105,
                     @"[In Response] RequestToken: A nonnegative integer that specifies the request token that uniquely identifies the Request element whose response is being generated.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R108
            site.CaptureRequirementIfAreEqual<string>(
                     expectedToken,
                     response.RequestToken,
                     "MS-FSSHTTP",
                     108,
                     @"[In Response] The one-to-one mapping between the Response element and the Request element MUST be maintained by using the request token.");
        }

        /// <summary>
        /// Capture requirements related with SubRequestToken within SubResponse element
        /// </summary>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateSubResponseToken(ITestSite site)
        {
            // All there requirements can be directly captured when the code runs here successfully.
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R143
            // If the SubRequestToken values in Request element equals the values in Response element, then capture MS-FSSHTTP_R143.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     143,
                     @"[In SubResponse] Within a Response element for each SubRequest element that is in a Request element of a cell storage service request message, there MUST be a corresponding SubResponse element.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R930
            // If the protocol server returns subResponse for each subRequest, then capture MS-FSSHTTP_R930.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     930,
                     @"[In Common Message Processing Rules and Events][The protocol server MUST follow the following common processing rules for all types of subrequests] and [The protocol server sends] a SubResponse element corresponding to each SubRequest element contained within a Request element.");
        }

        /// <summary>
        /// Capture requirements related with ResponseVersion element
        /// </summary>
        /// <param name="version">The ResponseVersion information</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateResponseVersion(ResponseVersion version, ITestSite site)
        {
            // Verify requirements related with its base type: VersionType
            ValidateVersionType(version as VersionType, site);

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R123
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     123,
                     @"[In ResponseVersion][ResponseVersion element schema is:]
                      <xs:element name=""ResponseVersion"">
                       < xs:complexType >
                        < xs:complexContent >
                         < xs:extension base = ""tns:VersionType"" >
                          < xs:attribute name = ""ErrorCode"" type = ""tns:GenericErrorCodeTypes"" use = ""optional"" />
                          < xs:attribute name = ""ErrorMessage"" type = ""xs:string"" use = ""optional"" />
                           </ xs:extension >
                        </ xs:complexContent >
                       </ xs:complexType >
                      </ xs:element > ");

            if (version.ErrorCodeSpecified)
            {
                // Verify requirements related with GenericErrorCodeTypes
                ValidateGenericErrorCodeTypes(site);
            }
        }

        /// <summary>
        /// Capture requirements related with ResponseCollection
        /// </summary>
        /// <param name="responseCollection">The ResponseCollection information</param>
        /// <param name="requestToken">The expected RequestToken</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateResponseCollection(ResponseCollection responseCollection, string requestToken, ITestSite site)
        {
            // Verify MS-FSSHTTP_R19
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     19,
                     @"[In Response] [The Body element of each SOAP response message MUST contain] zero or more ResponseCollection elements.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R121
            // If WebUrl is not null, WebUrl attribute is specified.
            site.Log.Add(
                LogEntryKind.Debug,
                "For requirement MS-FSSHTTP_R121, the WebUrl attribute should be specified, the actual WebUrl value is: {0}",
                responseCollection.WebUrl != null ? responseCollection.WebUrl : "NULL");

            site.CaptureRequirementIfIsNotNull(
                     responseCollection.WebUrl,
                     "MS-FSSHTTP",
                     121,
                     @"[In ResponseCollection] The WebUrl attribute MUST be specified for each ResponseCollection element.");

            // Now here only supported one request.
            if (responseCollection.Response != null && responseCollection.Response.Length >= 1)
            {
                MsfsshttpAdapterCapture.ValidateResponseElement(responseCollection.Response[0], site);
                MsfsshttpAdapterCapture.ValidateResponseToken(responseCollection.Response[0], requestToken, site);
            }
        }

        /// <summary>
        /// Capture requirements related with Response message
        /// </summary>
        /// <param name="storageResponse">The storage response information</param>
        /// <param name="requestToken">The expected RequestToken</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateResponse(CellStorageResponse storageResponse, string requestToken, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R17
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     17,
                     @"[In Response] The protocol response schema is specified by the following:
                     <?xml version=""1.0"" encoding=""utf-8""?>
                     <xs:schema xmlns:tns=""http://schemas.microsoft.com/sharepoint/soap/"" attributeFormDefault=""unqualified"" elementFormDefault=""qualified"" 
                     targetNamespace=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"" xmlns:i=""http://www.w3.org/2004/08/xop/include"">
                     <xs:import namespace=""http://www.w3.org/2004/08/xop/include"" />
                     
                     <xs:element name=""Envelope"">
                      <xs:complexType>
                       <xs:sequence>
                        <xs:element name=""Body"">
                         <xs:complexType>
                          <xs:sequence>
                           <xs:element ref=""tns:ResponseVersion"" minOccurs=""1"" maxOccurs=""1"" />
                           <xs:element ref=""tns:ResponseCollection"" minOccurs=""0"" maxOccurs=""1""/>
                          </xs:sequence>
                         </xs:complexType>
                        </xs:element>
                       </xs:sequence>
                      </xs:complexType>
                     </xs:element>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R18
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     18,
                     @"[In Response] The Body element of each SOAP response message MUST contain a ResponseVersion element.");
            
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R19
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     19,
                     @"[In Response] [The Body element of each SOAP response message MUST contain] zero or more ResponseCollection elements.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1613
            bool isVerifiedR1613 = storageResponse.ResponseVersion != null;
            site.Log.Add(
               LogEntryKind.Debug,
               "For requirement MS-FSSHTTP_R1613, [In Messages] Response: The detail element of the protocol response contains a ResponseVersion element, the actual ResponseVersion elements is: {0}",
               storageResponse.ResponseVersion != null ? storageResponse.ResponseVersion.ToString() : "NULL");

            site.CaptureRequirementIfIsTrue(
                     isVerifiedR1613,
                     "MS-FSSHTTP",
                     1613,
                     @"[In Messages] Response: The detail element of the protocol response contains a ResponseVersion element and zero or one ResponseCollection elements.");

            // Verify requirements related with ResponseVersion
            ValidateResponseVersion(storageResponse.ResponseVersion, site);

            // Verify requirements related with ResponseCollection
            if (storageResponse.ResponseCollection != null)
            {
                ValidateResponseCollection(storageResponse.ResponseCollection, requestToken, site);
            }
        }

        /// <summary>
        /// Capture requirements related with SubResponseType
        /// </summary>
        /// <param name="subResponse">The SubResponseType information</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateSubResponseType(SubResponseType subResponse, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R281
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     281,
                     @"[In SubResponseType][SubResponseType schema is:]
                      <xs:complexType name=""SubResponseType"">
                         < xs:attribute name = ""SubRequestToken"" type = ""xs:nonNegativeInteger"" use = ""required"" />
                         < xs:attribute name = ""ServerCorrelationId"" type = ""tns:guid"" use = ""optional"" />
                         < xs:attribute name = ""ErrorCode"" type = ""tns:ErrorCodeTypes"" use = ""required"" />
                         < xs:attribute name = ""HResult"" type = ""xs:integer"" use = ""required"" />
                         < xs:attribute name = ""ErrorMessage"" type = ""xs:string"" use = ""optional"" />
                      </ xs:complexType > ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R284
            // If the SubRequestToken attribute isn't null, then capture MS-FSSHTTP_R284.  
            site.Log.Add(
                LogEntryKind.Debug,
                "For requirement MS-FSSHTTP_R284, the SubRequestToken should be specified, the actual SubRequestToken value is: {0}",
                subResponse.SubRequestToken != null ? subResponse.SubRequestToken : "NULL");

            site.CaptureRequirementIfIsNotNull(
                     subResponse.SubRequestToken,
                     "MS-FSSHTTP",
                     284,
                     @"[In SubResponseType] The SubRequestToken attribute MUST be specified for a SubResponse element.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R290
            // If the ErrorCode attribute is not null, capture R290.
            site.Log.Add(
                LogEntryKind.Debug,
                "For requirement MS-FSSHTTP_R290, the ErrorCode attribute should be specified, the actual ErrorCode value is: {0}",
                subResponse.ErrorCode != null ? subResponse.ErrorCode : "NULL");

            site.CaptureRequirementIfIsNotNull(
                     subResponse.ErrorCode,
                     "MS-FSSHTTP",
                     290,
                     @"[In SubResponseType] The ErrorCode attribute MUST be specified for a SubResponse element.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1003
            site.CaptureRequirementIfIsNotNull(
                     subResponse.ErrorCode,
                     "MS-FSSHTTP",
                     1003,
                     @"[In Coauth Subrequest][The protocol server returns results based on the following conditions:]  Depending on the type of error, ErrorCode is returned as an attribute of the SubResponse element.");

            if (subResponse.ErrorCode != null)
            {
                ErrorCodeType errorCode;
                site.Assert.IsTrue(Enum.TryParse<ErrorCodeType>(subResponse.ErrorCode, true, out errorCode), "Fail to convert the error code string {0} to the Enum type ErrorCodeType", subResponse.ErrorCode);

                // Verify requirements related with ErrorCodeTypes
                ValidateErrorCodeTypes(errorCode, site);
            }
        }

        /// <summary>
        /// Capture requirements related with ErrorCodeTypes.
        /// </summary>
        /// <param name="errorCode">The error code returned by the protocol server</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateErrorCodeTypes(ErrorCodeType errorCode, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R341
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     341,
                     @"[In ErrorCodeTypes][ErrorCodeTypes schema is:]
                      <xs:simpleType name=""ErrorCodeTypes"">
                       < xs:union memberTypes = ""tns:GenericErrorCodeTypes tns:CellRequestErrorCodeTypes tns:DependencyCheckRelatedErrorCodeTypes tns:LockAndCoauthRelatedErrorCodeTypes tns:NewEditorsTableCategoryErrorCodeTypes"" tns: VersioningRelatedErrorCodeTypes >
                      </ xs:simpleType > ");

            switch (errorCode)
            {
                // Verify DependencyCheckRelatedErrorCodeTypes
                case ErrorCodeType.DependentRequestNotExecuted:
                case ErrorCodeType.DependentOnlyOnSuccessRequestFailed:
                case ErrorCodeType.DependentOnlyOnFailRequestSucceeded:
                case ErrorCodeType.DependentOnlyOnNotSupportedRequestGetSupported:
                case ErrorCodeType.InvalidRequestDependencyType:
                    {
                        ValidateDependencyCheckRelatedErrorCodeTypes(site);
                        break;
                    }

                // Verify LockAndCoauthRelatedErrorCodeTypes
                case ErrorCodeType.LockRequestFail:
                case ErrorCodeType.FileAlreadyLockedOnServer:
                case ErrorCodeType.FileNotLockedOnServer:
                case ErrorCodeType.FileNotLockedOnServerAsCoauthDisabled:
                case ErrorCodeType.LockNotConvertedAsCoauthDisabled:
                case ErrorCodeType.FileAlreadyCheckedOutOnServer:
                case ErrorCodeType.ConvertToSchemaFailedFileCheckedOutByCurrentUser:
                case ErrorCodeType.CoauthRefblobConcurrencyViolation:
                case ErrorCodeType.MultipleClientsInCoauthSession:
                case ErrorCodeType.InvalidCoauthSession:
                case ErrorCodeType.NumberOfCoauthorsReachedMax:
                case ErrorCodeType.ExitCoauthSessionAsConvertToExclusiveFailed:
                    {
                        ValidateLockAndCoauthRelatedErrorCodeTypes(site);
                        break;
                    }

                // Verify GenericErrorCodeTypes
                case ErrorCodeType.Success:
                case ErrorCodeType.IncompatibleVersion:
                case ErrorCodeType.FileNotExistsOrCannotBeCreated:
                case ErrorCodeType.FileUnauthorizedAccess:
                case ErrorCodeType.InvalidSubRequest:
                case ErrorCodeType.SubRequestFail:
                case ErrorCodeType.BlockedFileType:
                case ErrorCodeType.DocumentCheckoutRequired:
                case ErrorCodeType.InvalidArgument:
                case ErrorCodeType.RequestNotSupported:
                case ErrorCodeType.InvalidWebUrl:
                case ErrorCodeType.WebServiceTurnedOff:
                case ErrorCodeType.ColdStoreConcurrencyViolation:
                case ErrorCodeType.Unknown:
                case ErrorCodeType.EditorClientIdNotFound:
                case ErrorCodeType.EditorMetadataQuotaReached:
                case ErrorCodeType.PathNotFound:
                case ErrorCodeType.EditorMetadataStringExceedsLengthLimit:
                    {
                        ValidateGenericErrorCodeTypes(site);
                        break;
                    }

                // Verify CellRequestErrorCodeTypes
                case ErrorCodeType.CellRequestFail:
                case ErrorCodeType.IRMDocLibarysOnlySupportWebDAV:
                    {
                        ValidateCellRequestErrorCodeTypes(errorCode, site);
                        break;
                    }

                default:
                    site.Assert.Fail(string.Format("Unknown ErrorCodeType: {0}", errorCode.ToString()));
                    break;
            }
        }

        /// <summary>
        /// Capture requirements related with DependencyCheckRelatedErrorCodeTypes
        /// </summary>
        /// <param name="site">ITestSite site</param>
        private static void ValidateDependencyCheckRelatedErrorCodeTypes(ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R319
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     319,
                     @"[In DependencyCheckRelatedErrorCodeTypes][DependencyCheckRelatedErrorCodeTypes schema is:]
                     <xs:simpleType name=""DependencyCheckRelatedErrorCodeTypes"">
                         <xs:restriction base=""xs:string"">
                           <xs:enumeration value=""DependentRequestNotExecuted""/>
                           <xs:enumeration value=""DependentOnlyOnSuccessRequestFailed""/>
                           <xs:enumeration value=""DependentOnlyOnFailRequestSucceeded""/>
                           <xs:enumeration value=""DependentOnlyOnNotSupportedRequestGetSupported""/>
                           <xs:enumeration value=""InvalidRequestDependencyType""/>
                         </xs:restriction>
                     </xs:simpleType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R320
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     320,
                     @"[In DependencyCheckRelatedErrorCodeTypes] The value of DependencyCheckRelatedErrorCodeTypes MUST be one of the following:
                     
                     [DependentRequestNotExecuted, DependentOnlyOnSuccessRequestFailed, DependentOnlyOnFailRequestSucceeded, DependentOnlyOnNotSupportedRequestGetSupported, InvalidRequestDependencyType]");
        }

        /// <summary>
        /// Capture requirements related with LockAndCoauthRelatedErrorCodeTypes
        /// </summary>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateLockAndCoauthRelatedErrorCodeTypes(ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R375
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     375,
                     @"[In LockAndCoauthRelatedErrorCodeTypes][The LockCoauthRelatedErrorCodeTypes schema is:]
                     <xs:simpleType name=""LockAndCoauthRelatedErrorCodeTypes"">
                        <xs:restriction base=""xs:string"">
                           <xs:enumeration value=""LockRequestFail""/>
                           <xs:enumeration value=""FileAlreadyLockedOnServer""/>
                           <xs:enumeration value=""FileNotLockedOnServer""/>
                           <xs:enumeration value=""FileNotLockedOnServerAsCoauthDisabled""/>
                           <xs:enumeration value=""LockNotConvertedAsCoauthDisabled""/>
                           <xs:enumeration value=""FileAlreadyCheckedOutOnServer""/>
                           <xs:enumeration value=""ConvertToSchemaFailedFileCheckedOutByCurrentUser""/>
                           <xs:enumeration value=""CoauthRefblobConcurrencyViolation""/>
                           <xs:enumeration value=""MultipleClientsInCoauthSession""/>
                           <xs:enumeration value=""InvalidCoauthSession""/>
                           <xs:enumeration value=""NumberOfCoauthorsReachedMax""/>
                           <xs:enumeration value=""ExitCoauthSessionAsConvertToExclusiveFailed""/>
                        </xs:restriction>
                     </xs:simpleType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R376
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     376,
                     @"[In LockAndCoauthRelatedErrorCodeTypes] The value of LockAndCoauthRelatedErrorCodeTypes MUST be one of the following:
                     
                     [LockRequestFail, FileAlreadyLockedOnServer, FileNotLockedOnServer, FileNotLockedOnServerAsCoauthDisabled, LockNotConvertedAsCoauthDisabled, FileAlreadyCheckedOutOnServer, ConvertToSchemaFailedFileCheckedOutByCurrentUser, CoauthRefblobConcurrencyViolation, MultipleClientsInCoauthSession, InvalidCoauthSession, NumberOfCoauthorsReachedMax, ExitCoauthSessionAsConvertToExclusiveFailed]");
        }

        /// <summary>
        /// Capture requirements related with VersionType
        /// </summary>
        /// <param name="version">The VersionType information</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateVersionType(VersionType version, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R295
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     295,
                     @"[In VersionType][VersionType schema is:]
                     <xs:complexType name=""VersionType"">
                        <xs:attribute name=""Version"" type=""tns:VersionNumberType"" use=""required"" />
                        <xs:attribute name=""MinorVersion"" type=""tns:MinorVersionNumberType"" use=""required"" />
                     </xs:complexType>");

            // Verify requirements related with VersionNumberType
            ValidateVersionNumberType(version.Version, site);

            // Verify requirements related with MinorVersionNumberType
            ValidateMinorVersionNumberType(version.MinorVersion, site);
        }

        /// <summary>
        /// Capture requirements related with VersionNumberType.
        /// </summary>
        /// <param name="versionNumber">A version number</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateVersionNumberType(ushort versionNumber, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R423
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     423,
                     @"[In VersionNumberType][VersionNumberType schema is:]
                     <xs:simpleType name=""VersionNumberType"">
                         <xs:restriction base=""xs:unsignedShort"">
                           <xs:minInclusive value=""2""/>
                     <xs:maxInclusive value=""2""/>
                         </xs:restriction>
                     </xs:simpleType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R424
            site.CaptureRequirementIfAreEqual<ushort>(
                     2,
                     versionNumber,
                     "MS-FSSHTTP",
                     424,
                     @"[In VersionNumberType] The value of VersionNumberType MUST be the value that is listed in the following: [2].");

            site.CaptureRequirementIfAreEqual<ushort>(
                     2,
                     versionNumber,
                     "MS-FSSHTTP",
                     425,
                     @"[In VersionNumberType] 2 [means] a version number of 2.");
        }

        /// <summary>
        /// Capture requirements related with MinorVersionNumberType.
        /// </summary>
        /// <param name="minorVersionNumber">The minor version number</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateMinorVersionNumberType(ushort minorVersionNumber, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R409
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     409,
                     @"[In MinorVersionNumberType][MinorVersionNumberType schema is:]
                     <xs:simpleType name=""MinorVersionNumberType"">
                         < xs:restriction base = ""xs:unsignedShort"" >
                            < xs:minInclusive value = ""0"" />
                     < xs:maxInclusive value = ""3"" />
                         </ xs:restriction >
                     </ xs:simpleType > ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R410
            bool isVerifiedR410 = minorVersionNumber == 0 || minorVersionNumber == 2 || minorVersionNumber == 3;
            site.Log.Add(
                LogEntryKind.Debug,
                "For requirement MS-FSSHTTP_R410, the MinorVersionNumberType value should be 0 or 2, the actual MinorVersionNumberType value is: {0}",
                minorVersionNumber.ToString());

            site.CaptureRequirementIfIsTrue(
                     isVerifiedR410,
                     "MS-FSSHTTP",
                     410,
                     @"[In MinorVersionNumberType] The value of MinorVersionNumberType MUST be the value [0, 2, 3] that is listed in the following table.");
        }

        /// <summary>
        /// Capture requirements related with GenericErrorCodeTypes.
        /// </summary>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateGenericErrorCodeTypes(ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R353
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     353,
                     @"[In GenericErrorCodeTypes][The GenericErrorodeTypes schema is:]
                      <xs:simpleType name=""GenericErrorCodeTypes"">
                          < xs:restriction base = ""xs:string"" >
                            < xs:enumeration value = ""Success"" />
                            < xs:enumeration value = ""IncompatibleVersion"" />
                            < xs:enumeration value = ""InvalidUrl"" />
                            < xs:enumeration value = ""FileNotExistsOrCannotBeCreated"" />
                            < xs:enumeration value = ""FileUnauthorizedAccess"" />
                            < xs:enumeration value = ""PathNotFound"" />
                            < xs:enumeration value = ""InvalidSubRequest"" />
                            < xs:enumeration value = ""SubRequestFail"" />
                            < xs:enumeration value = ""BlockedFileType"" />
                            < xs:enumeration value = ""DocumentCheckoutRequired"" />
                            < xs:enumeration value = ""InvalidArgument"" />
                            < xs:enumeration value = ""RequestNotSupported"" />
                            < xs:enumeration value = ""InvalidWebUrl"" />
                            < xs:enumeration value = ""WebServiceTurnedOff"" />
                            < xs:enumeration value = ""ColdStoreConcurrencyViolation"" />
                            < xs:enumeration value = ""HighLevelExceptionThrown"" />
                            < xs:enumeration value = ""Unknown"" />
                          </ xs:restriction >
                      </ xs:simpleType > ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R354
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     354,
                     @"[In GenericErrorCodeTypes] The value of GenericErrorCodeTypes MUST be one of the following:
                     [Success, IncompatibleVersion, InvalidUrl, FileNotExistsOrCannotBeCreated, FileUnauthorizedAccess, PathNotFound, InvalidSubRequest, SubRequestFail, BlockedFileType, DocumentCheckoutRequired, InvalidArgument, RequestNotSupported, InvalidWebUrl, WebServiceTurnedOff, ColdStoreConcurrencyViolation, Unknown]");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3329
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     3329,
                     @"[In TRUEFALSE][The schema of TRUEFALSE is:]
                     <xs:simpleType name=""TRUEFALSE"">
                         <xs:restriction base=""xs:string"">
                           <xs:pattern value=""[Tt][Rr][Uu][Ee]|[Ff][Aa][Ll][Ss][Ee]""/>
                         </xs:restriction>
                       </xs:simpleType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3033
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     3033,
                     @"[In NewEditorsTableCategoryErrorCodeTypes][The schema of NewEditorsTableCategoryErrorCodeTypes is:]
                     <xs:simpleType name=""NewEditorsTableCategoryErrorCodeTypes"">
                         <xs:restriction base=""xs:string"">
                            <xs:enumeration value=""EditorMetadataQuotaReached""/>
                            <xs:enumeration value=""EditorMetadataStringExceedsLengthLimit""/>
                            <xs:enumeration value=""EditorClientIdNotFound""/>
                         </xs:restriction>
                     </xs:simpleType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3034
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     3034,
                     @"[In NewEditorsTableCategoryErrorCodeTypes] The value of NewEditorsTableCategoryErrorCodeTypes MUST be one of the values in the following table. [""EditorMetadataQuotaReached"", ""EditorMetadataStringExceedsLengthLimit"", ""EditorClientIdNotFound""]");
        }

        /// <summary>
        /// Capture requirements related with Response element.
        /// </summary>
        /// <param name="response">The Response information</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateResponseElement(Response response, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R97
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     97,
                     @"[In Response][Response element schema is:]
                       <xs:element name=""Response"">
                           < !--Allows for the numbers to be displayed between the SubResponse elements-- >
                           < xs:complexType mixed = ""true"" >
                               < xs:sequence minOccurs = ""1"" maxOccurs = ""unbounded"" >
                                   < xs:element name = ""SubResponse"" type = ""tns:SubResponseElementGenericType"" />
                               </ xs:sequence >
                               < xs:attribute name = ""Url"" type = ""xs:string"" use = ""required"" />
                               < xs:attribute name = ""RequestToken"" type = ""xs:nonNegativeInteger"" use = ""required"" />
                               < xs:attribute name = ""HealthScore"" type = ""xs:integer"" use = ""required"" />
                               < xs:attribute name = ""ErrorCode"" type = ""tns:GenericErrorCodeTypes"" use = ""optional"" />
                               < xs:attribute name = ""ErrorMessage"" type = ""xs:string"" use = ""optional"" />
                               < xs:attribute name = ""SuggestedFileName"" type = ""xs:string"" use = ""optional"" />
                               < xs:attribute name = ""ResourceID"" type = ""xs:string"" use = ""optional"" />
                           </ xs:complexType >
                       </ xs:element > ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R104
            site.Log.Add(
                LogEntryKind.Debug,
                "For requirement MS-FSSHTTP_R104, the Url attribute value should be specified, the actual Url value is: {0}",
                response.Url != null ? response.Url : "NULL");

            site.CaptureRequirementIfIsNotNull(
                     response.Url,
                     "MS-FSSHTTP",
                     104,
                     @"[In Response] The Url attribute MUST be specified for each Response element.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1482
            // The responseElement.SubResponse.Length specifies the number of SubResponse element in Response.
            bool isVerifiedR96 = response.SubResponse.Length >= 1;
            site.Log.Add(
                LogEntryKind.Debug,
                "For requirement MS-FSSHTTP_R1482, the Response element should contain one or more SubResponse elements, the actual SubResponse elements number is: {0}",
                response.SubResponse.Length.ToString());

            site.CaptureRequirementIfIsTrue(
                     isVerifiedR96,
                     "MS-FSSHTTP",
                     96,
                     @"[In Response] Each Response element MUST contain one or more SubResponse elements.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R107
            site.Log.Add(
                LogEntryKind.Debug,
                "For requirement MS-FSSHTTP_R263, the RequestToken should be specified, the actual RequestToken value is: {0}",
                response.RequestToken != null ? response.RequestToken : "NULL");

            site.CaptureRequirementIfIsNotNull(
                     response.RequestToken,
                     "MS-FSSHTTP",
                     107,
                     @"[In Response] The RequestToken MUST be specified for each Response element.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2076
            bool isVerifyR2076 = int.Parse(response.HealthScore) <= 10 && int.Parse(response.HealthScore) >= 0;
            site.Log.Add(
                LogEntryKind.Debug,
                "For requirement MS-FSSHTTP_R2076, the HealthScore value should be between 0 and 10, the actual HealthScore value is: {0}",
                response.HealthScore);

            site.CaptureRequirementIfIsTrue(
                     isVerifyR2076,
                     "MS-FSSHTTP",
                     2076,
                     @"[In Response] HealthScore: An integer which value is between 0 and 10.");

            // Verify requirements related with SubResponse element
            if (response.SubResponse != null)
            {
                foreach (SubResponseElementGenericType subResponse in response.SubResponse)
                {
                    ValidateSubResponseElement(subResponse, site);
                }
            }

            // Verify requirements related with GenericErrorCodeTypes
            if (response.ErrorCodeSpecified)
            {
                ValidateGenericErrorCodeTypes(site);
            }
        }

        /// <summary>
        /// Capture requirements related with SubResponse element
        /// </summary>
        /// <param name="subResponse">The SubResponse element</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateSubResponseElement(SubResponseElementGenericType subResponse, ITestSite site)
        {
            // Verify requirements related with SubResponseElementGenericType
            ValidateSubResponseElementGenericType(subResponse, site);

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R151
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     151,
                     @"[In SubResponse][SubResponse element schema is:]
                     <xs:element name=""SubResponse"" type=""tns:SubResponseElementGenericType"" />");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R149
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     149,
                     @"[In SubResponse] Each SubResponse element MUST have zero or one SubResponseData elements.");
        }

        /// <summary>
        /// Capture requirements related with SubResponseElementGenericType
        /// </summary>
        /// <param name="subResponse">The SubResponse information</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateSubResponseElementGenericType(SubResponseElementGenericType subResponse, ITestSite site)
        {
            // Verify requirements related with its base type: SubResponseType
            ValidateSubResponseType(subResponse as SubResponseType, site);

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R253
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     253,
                     @"[In SubResponseElementGenericType][SubResponseElementGenericType schema is:]
                      <xs:complexType name=""SubResponseElementGenericType"">
                       < xs:complexContent >
                        < xs:extension base = ""tns:SubResponseType"" >
                         < xs:sequence >
                          < xs:element name = ""SubResponseData"" minOccurs = ""0"" maxOccurs = ""1"" type = ""tns:SubResponseDataGenericType"" />
                          < xs:element name = ""SubResponseStreamInvalid"" minOccurs = ""0"" maxOccurs = ""1"" />
                          < xs:element ref= ""GetVersionsResponse"" minOccurs = ""0"" maxOccurs = ""1"" />
                         </ xs:sequence >
                        </ xs:extension >
                       </ xs:complexContent >
                      </ xs:complexType > ");

            // Verify requirements related with SubResponseData
            if (subResponse.SubResponseData != null)
            {
                ValidateSubResponseData(subResponse.SubResponseData, site);
            }
        }

        /// <summary>
        /// Capture requirements related with SubResponseData element
        /// </summary>
        /// <param name="subResponseData">The SubResponseData element</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateSubResponseData(SubResponseDataGenericType subResponseData, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R155
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     155,
                     @"[In SubResponseData][SubResponseData element schema is:]
                     <xs:element name=""SubResponseData"" minOccurs=""0"" maxOccurs=""1"" type=""tns:SubResponseDataGenericType"" />");

            // Verify requirements related with SubResponseDataGenericType
            ValidateSubResponseDataGenericType(subResponseData, site);
        }

        /// <summary>
        /// Capture requirements related with SubResponseDataGenericType
        /// </summary>
        /// <param name="subResponseData">The SubResponseDataGenericType information</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateSubResponseDataGenericType(SubResponseDataGenericType subResponseData, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R230
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     230,
                     @"[In SubResponseDataGenericType][SubResponseDataGenericType schema is:]
                      <xs:complexType name=""SubResponseDataGenericType"" mixed=""true"">
                        < xs:all >
                          < xs:element ref= ""i:Include"" minOccurs = ""0"" maxOccurs = ""1"" />
                          < xs:element name = ""DocProps"" minOccurs = ""0"" maxOccurs = ""1"" type = ""tns:GetDocMetaInfoPropertySetType"" />
                          < xs:element name = ""FolderProps"" minOccurs = ""0"" maxOccurs = ""1"" type = ""tns:GetDocMetaInfoPropertySetType"" />
                          < xs:element name = ""UserTable"" type = ""tns:VersioningUserTableType"" />
                          < xs:element name = ""Versions"" type = ""tns:VersioningVersionListType"" />
                        </ xs:all >
                        < xs:attributeGroup ref= ""tns:SubResponseDataOptionalAttributes"" />
                      </ xs:complexType > ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R543
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     543,
                     @"[In CellSubResponseDataType][CellSubResponseDataType schema is:]
                     <xs:complexType name=""CellSubResponseDataType"" mixed=""true"">
                        <xs:all>
                          <xs:element ref=""i:Include"" minOccurs=""0"" maxOccurs=""1"" />
                        </xs:all>
                        <xs:attributeGroup ref=""tns:CellSubResponseDataOptionalAttributes"" />
                        <xs:attribute name=""LockType"" type=""tns:LockTypes"" use=""optional"" />
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R544
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     544,
                     @"[In CellSubResponseDataType] Include: A complex type, as specified in [XOP10] section 2.1.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R231
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     231,
                     @"[In SubResponseDataGenericType] Include: A complex type, as specified in  [XOP10] section 2.1.");

            // Verify requirements related with SubResponseDataOptionalAttributes
            ValidateSubResponseDataOptionalAttributes(subResponseData, site);
        }

        /// <summary>
        /// Capture requirements related with SubResponseDataOptionalAttributes type
        /// </summary>
        /// <param name="subResponseData">The SubResponseData information</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateSubResponseDataOptionalAttributes(SubResponseDataGenericType subResponseData, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R458
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     458,
                     @"[In SubResponseDataOptionalAttributes] The definition of the SubResponseDataOptionalAttributes attribute group is as follows:
                     <xs:attributeGroup name=""SubResponseDataOptionalAttributes"">
                         <xs:attributeGroup ref=""tns:CellSubResponseDataOptionalAttributes""/>
                         <xs:attributeGroup ref=""tns:WhoAmISubResponseDataOptionalAttributes""/>
                         <xs:attribute name=""ServerTime"" type=""xs:positiveInteger"" use=""optional""/>
                         <xs:attribute name=""LockType"" type=""tns:LockTypes"" use=""optional"" />
                         <xs:attribute name=""CoauthStatus"" type=""tns:CoauthStatusType"" use=""optional""/>
                         <xs:attribute name=""TransitionID"" type=""tns:guid"" use=""optional""/>
                         <xs:attribute name=""ExclusiveLockReturnReason"" type=""tns:ExclusiveLockReturnReasonTypes"" use=""optional"" /> 
                     </xs:attributeGroup>");

            // Verify requirements related with CellSubResponseDataOptionalAttributes
            if (subResponseData.Etag != null
                || subResponseData.CreateTime != null
                || subResponseData.ModifiedBy != null
                || subResponseData.LastModifiedTime != null
                || subResponseData.CoalesceErrorMessage != null)
            {
                CellSubResponseDataType cellSubResponseData = new CellSubResponseDataType();
                cellSubResponseData.Etag = subResponseData.Etag;
                cellSubResponseData.CreateTime = subResponseData.CreateTime;
                cellSubResponseData.ModifiedBy = subResponseData.ModifiedBy;
                cellSubResponseData.LastModifiedTime = subResponseData.LastModifiedTime;
                cellSubResponseData.CoalesceErrorMessage = subResponseData.CoalesceErrorMessage;
                ValidateCellSubResponseDataOptionalAttributes(cellSubResponseData, site);
            }

            // Verify requirements related with WhoAmISubResponseDataOptionalAttributes
            if (subResponseData.UserName != null
                || subResponseData.UserEmailAddress != null
                || subResponseData.UserSIPAddress != null)
            {
                WhoAmISubResponseDataType whoamiSubResponseData = new WhoAmISubResponseDataType();
                whoamiSubResponseData.UserName = subResponseData.UserName;
                whoamiSubResponseData.UserEmailAddress = subResponseData.UserEmailAddress;
                whoamiSubResponseData.UserSIPAddress = subResponseData.UserSIPAddress;
                ValidateWhoAmISubResponseDataOptionalAttributes(whoamiSubResponseData, site);
            }

            if (subResponseData.ServerTime != null)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R463
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         463,
                         @"[In SubResponseDataOptionalAttributes] ServerTime: A positive integer that specifies the server time, which is expressed as a tick count.");
            }

            if (subResponseData.DocProps != null)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11050
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         11050,
                         @"[In SubResponseDataGenericType] DocProps: An element of type GetDocMetaInfoPropertySetType (section 2.3.1.28) that specifies metadata properties pertaining to the server file.");
            }

            if (subResponseData.FolderProps != null)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11051
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         11051,
                         @"[In SubResponseDataGenericType] FolderProps: An element of type GetDocMetaInfoPropertySetType (section 2.3.1.28) that specifies metadata properties pertaining to the parent directory of the server file.");
            }

            if (subResponseData.LockTypeSpecified)
            {
                // Verify LockTypes
                ValidateLockTypes(site);
            }

            if (subResponseData.CoauthStatusSpecified)
            {
                // Verify CoauthStatusType
                ValidateCoauthStatusType(site);
            }

            if (subResponseData.TransitionID != null)
            {
                // Verify GUID type
                ValidateGUID(site);
            }

            if (subResponseData.ExclusiveLockReturnReasonSpecified)
            {
                // Verify ExclusiveLockReturnReasonTypes
                ValidateExclusiveLockReturnReasonTypes(site);
            }
        }

        #region ValidateCoauthStatusType

        /// <summary>
        /// Capture the CoauthStatusType schema related requirements. 
        /// </summary>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        private static void ValidateCoauthStatusType(ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R314
            // If LockType exists,its schema will have been validated before invoking the method CaptureJoinCoauthoringSessionRelatedRequirements.
            // So MS-FSSHTTP_R314 can be captured.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     314,
                     @"[In CoauthStatusType][CoauthStatusType schema is:]
                     <xs:simpleType name=""CoauthStatusType"">
                         <xs:restriction base=""xs:string"">
                           <!--None-->
                           <xs:enumeration value=""None""/>
                           
                           <!--Alone-->
                           <xs:enumeration value=""Alone""/>
                           
                           <!--Coauthoring-->
                           <xs:enumeration value=""Coauthoring""/>
                         </xs:restriction>
                     </xs:simpleType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R315
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     315,
                     @"[In CoauthStatusType] The value of CoauthStatusType MUST be one of the following: [None, Alone, Coauthoring]");
        }

        #endregion 

        #region ValidateExclusiveLockReturnReasonTypes

        /// <summary>
        /// Capture the ExclusiveLockReturnReasonTypes schema related requirements. 
        /// </summary>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        private static void ValidateExclusiveLockReturnReasonTypes(ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R347
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     347,
                     @"[In ExclusiveLockReturnReasonTypes][The schema of ExclusiveLockReturnReasonTypes is] <xs:simpleType name=""ExclusiveLockReturnReasonTypes"">
                        <xs:restriction base=""xs:string"">
                           <xs:enumeration value=""CoauthoringDisabled"" />
                           <xs:enumeration value=""CheckedOutByCurrentUser"" />
                           <xs:enumeration value=""CurrentUserHasExclusiveLock"" />
                        </xs:restriction>
                     </xs:simpleType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R348
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     348,
                     @"[In ExclusiveLockReturnReasonTypes] The value of ExclusiveLockReturnReasonTypes MUST be one of the following:
                     [CoauthoringDisabled, CheckedOutByCurrentUser, CurrentUserHasExclusiveLock]");
        }

        #endregion 

        #region ValidateGUID

        /// <summary>
        /// Capture the GUID schema related requirements. 
        /// </summary>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        private static void ValidateGUID(ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R373
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     373,
                     @"[In GUID][The GUID schema is:]
                     <xs:simpleType name=""guid"">
                        <xs:restriction base=""xs:string"">
                          <xs:pattern value=""[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"" />
                        </xs:restriction>
                     </xs:simpleType>");
        }

        #endregion 

        #region ValidateLockTypes

        /// <summary>
        /// Capture the soap version related requirements. They can be captured directly when the server returns a SOAP response successfully.
        /// </summary>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        private static void ValidateLockTypes(ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R398
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     398,
                     @"[In LockTypes][The Locktypes schema is:]
                     
                     <xs:simpleType name=""LockTypes"">
                        <xs:restriction base=""xs:string"">
                           <xs:enumeration value=""None"" />
                           <xs:enumeration value=""SchemaLock"" />
                           <xs:enumeration value=""ExclusiveLock"" />
                        </xs:restriction>
                     </xs:simpleType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R399
            // There is only 3 values in the enumeration LockTypes:None, SchemaLock, ExclusiveLock.So if lockType can be a non-null value,
            // the value must be one of them.So MS-FSSHTTP_R399 can be captured as the following:
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     399,
                     @"[In LockTypes] The value of LockTypes MUST be one of the following:
                     [None, SchemaLock, ExclusiveLock].");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R467
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     467,
                     @"[In SubResponseDataOptionalAttributes] LockType: A LockTypes that specifies the type of lock granted in a coauthoring subresponse.");
        }

        #endregion 
    }
}