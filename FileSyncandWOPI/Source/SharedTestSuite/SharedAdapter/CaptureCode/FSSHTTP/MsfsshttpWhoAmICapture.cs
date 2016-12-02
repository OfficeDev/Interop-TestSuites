namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class contains methods which capture requirements related with WhoAmI Sub-request.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related with WhoAmI Sub-request.
        /// </summary>
        /// <param name="whoamiSubResponse">Containing the WhoAmISubResponse information</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateWhoAmISubResponse(WhoAmISubResponseType whoamiSubResponse, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R765
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     765,
                     @"[In WhoAmISubResponseType][WhoAmISubResponseType schema is:]
                     <xs:complexType name=""WhoAmISubResponseType"">
                       <xs:complexContent>
                         <xs:extension base=""tns:SubResponseType"">
                           <xs:sequence minOccurs=""0"" maxOccurs=""1"">
                              <xs:element name=""SubResponseData"" type=""tns:WhoAmISubResponseDataType""/>
                           </xs:sequence>
                         </xs:extension>
                       </xs:complexContent>
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1316
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(WhoAmISubResponseType),
                     whoamiSubResponse.GetType(),
                     "MS-FSSHTTP",
                     1316,
                     @"[In WhoAmI Subrequest][The protocol client sends a WhoAmI SubRequest message, which is of type WhoAmISubRequestType] The protocol server responds with a WhoAmI SubResponse message, which is of type WhoAmISubResponseType as specified in section 2.3.1.22.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R4692
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(WhoAmISubResponseType),
                     whoamiSubResponse.GetType(),
                     "MS-FSSHTTP",
                     4692,
                     @"[In SubResponseElementGenericType] Depending on the Type attribute specified in the SubRequest element, the SubResponseElementGenericType MUST take one of the forms: WhoAmISubResponseType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R5746
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(WhoAmISubResponseType),
                     whoamiSubResponse.GetType(),
                     "MS-FSSHTTP",
                     5746,
                     @"[In SubResponseType] The SubResponseElementGenericType takes one of the following forms: WhoAmISubResponseType.");

            ErrorCodeType errorCode;
            site.Assert.IsTrue(Enum.TryParse<ErrorCodeType>(whoamiSubResponse.ErrorCode, true, out errorCode), "Fail to convert the error code string {0} to the Enum type ErrorCodeType", whoamiSubResponse.ErrorCode);
            if (errorCode == ErrorCodeType.Success)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R268
                site.CaptureRequirementIfIsNotNull(
                         whoamiSubResponse.SubResponseData,
                         "MS-FSSHTTP",
                         268,
                         @"[In SubResponseElementGenericType][The SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"" and one of the following conditions is true:] The Type attribute that is specified in the SubRequest element is set to a value of ""WhoAmI"".");
            }

            // Verify requirements related with its base type: SubResponseType
            ValidateSubResponseType(whoamiSubResponse as SubResponseType, site);

            // Verify requirements related with SubResponseDataType
            if (whoamiSubResponse.SubResponseData != null)
            {
                ValidateWhoAmISubResponseDataType(whoamiSubResponse.SubResponseData, site);
            }
        }

        /// <summary>
        /// Capture requirements related with WhoAmISubResponseDataType
        /// </summary>
        /// <param name="whoamiSubResponseData">The WhoAmISubResponseData information</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateWhoAmISubResponseDataType(WhoAmISubResponseDataType whoamiSubResponseData, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R757
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     757,
                     @"[In WhoAmISubResponseDataType][WhoAmISubResponseDataType schema is:]
                     <xs:complexType name=""WhoAmISubResponseDataType"">
                     <xs:attributeGroup ref=""tns:WhoAmISubResponseDataOptionalAttributes""/>
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1396
            // The SubResponseData of WhoamiSubResponse is of type WhoAmISubResponseDataType, so if whoamiSubResponse.SubResponseData is not null, then MS-FSSHTTP_R1396 can be captured.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(WhoAmISubResponseDataType),
                     whoamiSubResponseData.GetType(),
                     "MS-FSSHTTP",
                     1396,
                     @"[In SubResponseDataGenericType][SubResponseDataGenericType MUST take one of the forms described in the following table] WhoAmISubResponseDataType: Type definition for Who Am I subresponse data.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1322
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(WhoAmISubResponseDataType),
                     whoamiSubResponseData.GetType(),
                     "MS-FSSHTTP",
                     1322,
                     @"[In WhoAmI Subrequest] The WhoAmISubResponseDataType defines the type of the SubResponseData element inside the WhoAmI SubResponse element.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1543
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1543,
                     @"[In WhoAmISubResponseDataOptionalAttributes][UserLoginType] The UserLogin attribute MUST be specified in a WhoAmI subresponse that is generated in response to a WhoAmI subrequest.");

            if (whoamiSubResponseData.UserName != null
                || whoamiSubResponseData.UserEmailAddress != null
                || whoamiSubResponseData.UserSIPAddress != null)
            {
                // Verify requirements related with WhoAmISubResponseDataOptionalAttributes
                ValidateWhoAmISubResponseDataOptionalAttributes(whoamiSubResponseData, site);
            }
        }

        /// <summary>
        /// Capture requirements related with WhoAmISubResponseDataOptionalAttributes.
        /// </summary>
        /// <param name="whoamiSubResponseData">The WhoAmISubResponseData</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateWhoAmISubResponseDataOptionalAttributes(WhoAmISubResponseDataType whoamiSubResponseData, ITestSite site)
        {
            if (whoamiSubResponseData.UserName != null)
            {
                // Verify requirements related with UserNameTypes
                ValidateUserNameTypes(site);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R794
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         794,
                         @"[In UserNameType] UserNameType is the type definition of the UserName attribute, which is part of the subresponse for a Who Am I subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2130
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         2130,
                         @" [In WhoAmISubResponseDataOptionalAttributes] UserName: [is] A UserNameType [that specifies the user name for the client.]");
            }

            if (!string.IsNullOrEmpty(whoamiSubResponseData.UserEmailAddress))
            {
                ValidateUserEmailAddress(whoamiSubResponseData.UserEmailAddress, site);
            }

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R903
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     903,
                     @"[In WhoAmISubResponseDataOptionalAttributes] The schema definition of the WhoAmISubResponseDataOptionalAttributes attribute group is as follows:
                     <xs:attributeGroup name=""WhoAmISubResponseDataOptionalAttributes"">
                         < xs:attribute name = ""UserName"" type = ""tns:UserNameType"" use = ""optional"" />
                         < xs:attribute name = ""UserEmailAddress"" type = ""xs:string"" use = ""optional"" />
                         < xs:attribute name = ""UserSIPAddress"" type = ""xs:string"" use = ""optional"" />
                         < xs:attribute name = ""UserIsAnonymous"" type = ""xs:boolean"" use = ""optional"" />
                         < xs:attribute name = ""UserLogin"" type = ""xs:UserLoginType"" use = ""required"" />
                     </ xs:attributeGroup > ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1466
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1466,
                     @"[In WhoAmISubResponseDataOptionalAttributes] The WhoAmISubResponseDataOptionalAttributes attribute group contains attributes that MUST be used in SubResponseData elements associated with a subresponse for a Who Am I subrequest.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1477
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1477,
                     @"[In SubResponseDataOptionalAttributes] WhoAmISubResponseDataOptionalAttributes: An attribute group that specifies attributes that MUST be used for SubResponseData elements associated with a subresponse for a WhoAmI subrequest.");
        }

        /// <summary>
        /// Capture requirements related with UserEmailAddress.
        /// </summary>
        /// <param name="userEmailAddress">The UserEmailAddress</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateUserEmailAddress(string userEmailAddress, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R909
            bool isVerifiedR909 = AdapterHelper.IsValidEmailAddr(userEmailAddress);
            site.Log.Add(
                LogEntryKind.Debug,
                "For requirement MS-FSSHTTP_R909, the format of the e-mail addresses should be as specified in [RFC2822] section 3.4.1, the actual e-mail addresses value is: {0}",
                userEmailAddress);

            site.CaptureRequirementIfIsTrue(
                     isVerifiedR909,
                     "MS-FSSHTTP",
                     909,
                     @"[In WhoAmISubResponseDataOptionalAttributes][UserEmailAddress] The format of the email address MUST be as specified in [RFC2822] section 3.4.1.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1467
            bool isVerifiedR1467 = AdapterHelper.IsValidEmailAddr(userEmailAddress);
            site.Log.Add(
               LogEntryKind.Debug,
               "For requirement MS-FSSHTTP_R1467, the value format of the userEmailAddress attribute should be as specified in [RFC2822], the actual userEmailAddress value is: {0}",
               userEmailAddress);
            site.CaptureRequirementIfIsTrue(
                     isVerifiedR1467,
                     "MS-FSSHTTP",
                     1467,
                     @"[In WhoAmISubResponseDataOptionalAttributes][UserEmailAddress] Format of the e-mail addresses MUST be:
                     
                     addr-spec       =       local-part ""@"" domain
                     
                     
                     
                     local-part      =       dot-atom / quoted-string / obs-local-part
                     
                     domain          =       dot-atom / domain-literal / obs-domain
                     
                     domain-literal  =       [CFWS] ""["" *([FWS] dcontent) [FWS] ""]"" [CFWS]
                     
                     dcontent        =       dtext / quoted-pair
                     
                     dtext           =       NO-WS-CTL /     ; Non white space controls
                     
                                             %d33-90 /       ; The rest of the US-ASCII
                     
                                             %d94-126        ;  characters not including ""["",
                     
                                                             ;  ""]"", or ""\""
                     
                     quoted-pair     =       (""\"" text) / obs-qp
                     
                     text            =       %d1-9 /         ; Characters excluding CR and LF
                     
                                             %d11 /
                     
                                             %d12 /
                     
                                             %d14-127 /
                     
                                             obs-text
                     
                     obs-text        =       *LF *CR *(obs-char *LF *CR)
                     
                     obs-char        =       %d0-9 / %d11 /          ; %d0-127 except CR and
                     
                                             %d12 / %d14-127         ;  LF
                     
                     obs-domain      =       atom *(""."" atom)
                     
                     atom            =       [CFWS] 1*atext [CFWS]
                     
                     atext           =       ALPHA / DIGIT / ; Any character except controls,
                     
                                             ""!"" / ""#"" /     ;  SP, and specials.
                     
                                             ""$"" / ""%"" /     ;  Used for atoms
                     
                                             ""&"" / ""'"" /
                     
                                             ""*"" / ""+"" /
                     
                                             ""-"" / ""/"" /
                     
                                             ""="" / ""?"" /
                     
                                             ""^"" / ""_"" /
                     
                                             ""`"" / ""{"" /
                     
                                             ""|"" / ""}"" /
                     
                                             ""~""
                     
                     dot-atom        =       [CFWS] dot-atom-text [CFWS]
                     
                     dot-atom-text   =       1*atext *(""."" 1*atext)
                     
                     NO-WS-CTL       =       %d1-8 /         ; US-ASCII control characters
                     
                                             %d11 /          ;  that do not include the
                     
                                             %d12 /          ;  carriage return, line feed,
                     
                                             %d14-31 /       ;  and white space characters
                     
                                             %d127");
        }
    }
}