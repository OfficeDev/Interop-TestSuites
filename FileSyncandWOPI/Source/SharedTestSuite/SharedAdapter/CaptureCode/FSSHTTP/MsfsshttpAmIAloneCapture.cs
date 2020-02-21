namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class contains methods which capture requirements related with AmIAlone Sub-request.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related with AmIAlone Sub-request.
        /// </summary>
        /// <param name="amIAloneSubResponse">Containing the AmIAloneSubResponse information</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateAmIAloneSubResponse(AmIAloneSubResponseType amIAloneSubResponse, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R22555
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     22555,
                     @"[In AmIAloneSubResponseType]	
  <xs:complexType name=""AmIAloneSubResponseType"" >
   < xs:complexContent >
    < xs:extension base = ""tns:SubResponseType"" >
      < xs:sequence minOccurs = ""0"" maxOccurs = ""1"" >
         < xs:element name = ""SubResponseData"" type = ""tns:AmIAloneSubResponseDataType"" />
      </ xs:sequence >
    </ xs:extension >
  </ xs:complexContent >
</ xs:complexType > ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2256
            site.CaptureRequirement(
                "MS-FSSHTTP",
                2256,
                @"[In AmIAloneSubResponseType]SubResponseData: An AmIAloneSubResponseDataType that specifies the information about whether the user is alone that was requested as part of the AmIAlone subrequest.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2363
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(AmIAloneSubResponseType),
                     amIAloneSubResponse.GetType(),
                     "MS-FSSHTTP",
                     2363,
                     @"[AmIAlone Subrequest]The protocol server responds with an AmIAlone SubResponse message, which is of type AmIAloneSubResponseType as specified in section 2.3.1.48. ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2146
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(AmIAloneSubResponseType),
                     amIAloneSubResponse.GetType(),
                     "MS-FSSHTTP",
                     2146,
                     @"[In SubResponseElementGenericType] Depending on the Type attribute specified in the SubRequest element, the SubResponseElementGenericType MUST take one of the forms: AmIAloneSubResponseType");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2164
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(AmIAloneSubResponseType),
                     amIAloneSubResponse.GetType(),
                     "MS-FSSHTTP",
                     2164,
                     @"[In SubResponseType] The SubResponseElementGenericType takes one of the following forms: AmIAloneSubResponseType.");

            ErrorCodeType errorCode;
            site.Assert.IsTrue(Enum.TryParse<ErrorCodeType>(amIAloneSubResponse.ErrorCode, true, out errorCode), "Fail to convert the error code string {0} to the Enum type ErrorCodeType", amIAloneSubResponse.ErrorCode);
            if (errorCode == ErrorCodeType.Success)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2158
                site.CaptureRequirementIfIsNotNull(
                         amIAloneSubResponse.SubResponseData,
                         "MS-FSSHTTP",
                         2158,
                         @"[In SubResponseElementGenericType][The SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"" and one of the following conditions is true:] The Type attribute that is specified in the SubRequest element is set to a value of ""AmIAlone"".");
            }

            // Verify requirements related with its base type: SubResponseType
            ValidateSubResponseType(amIAloneSubResponse as SubResponseType, site);

            // Verify requirements related with SubResponseDataType
            if (amIAloneSubResponse.SubResponseData != null)
            {
                ValidateAmIAloneSubResponseDataType(amIAloneSubResponse.SubResponseData, site);
            }
        }

        /// <summary>
        /// Capture requirements related with AmIAloneSubResponseDataType
        /// </summary>
        /// <param name="amIAloneSubResponseData">The AmIAloneSubResponseData information</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateAmIAloneSubResponseDataType(AmIAloneSubResponseDataType amIAloneSubResponseData, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2369
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2369,
                     @"[AmIAlone Subrequest]The AmIAloneSubResponseDataType defines the type of the SubResponseData element inside the AmIAloneSubResponse element. ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R224811
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     224811,
                     @"[In AmIAloneSubResponseDataType]	
<xs:complexType name=""AmIAloneSubResponseDataType"" >
    < xs:attribute name = ""AmIAlone"" type = ""tns:TRUEFALSE"" use = ""optional"" />
</ xs:complexType > ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2134
            // The SubResponseData of AmIAloneSubResponse is of type AmIAloneSubResponseDataType, so if amIAloneSubResponse.SubResponseData is not null, then MS-FSSHTTP_R2134 can be captured.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(AmIAloneSubResponseDataType),
                     amIAloneSubResponseData.GetType(),
                     "MS-FSSHTTP",
                     2134,
                     @"[In SubResponseDataGenericType][SubResponseDataGenericType MUST take one of the forms described in the following table]AmIAloneSubResponseDataType:Type definition for Am I Alone subresponse data.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2371
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(AmIAloneSubResponseDataType),
                     amIAloneSubResponseData.GetType(),
                     "MS-FSSHTTP",
                     2371,
                     @"[AmIAlone Subrequest]The protocol server sends the requested information as AmIAlone attribute in the AmIAlone SubResponseData element. ");
        }
    }
}