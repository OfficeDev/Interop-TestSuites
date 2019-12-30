namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class contains methods which capture requirements related with Properties Sub-request.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related with Properties Sub-request.
        /// </summary>
        /// <param name="propertiesSubResponse">Containing the PropertiesSubResponse information</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidatePropertiesSubResponse(PropertiesSubResponseType propertiesSubResponse, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2305
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2305,
                     @"[PropertiesSubResponseType]
	<xs:complexType name=""PropertiesSubResponseType"">
	  <xs:complexContent>
	    <xs:extension base=""tns:SubResponseType"">
	      <xs:sequence minOccurs=""0"" maxOccurs=""1"">
	         <xs:element name=""SubResponseData"" type=""tns:PropertiesSubResponseDataType"" />
	      </xs:sequence>
	    </xs:extension>
	  </xs:complexContent>
	</xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2406
            site.CaptureRequirement(
                "MS-FSSHTTP",
                2406,
                @"[PropertiesSubResponseType]SubResponseData: A PropertiesSubResponseDataType that specifies the information about the properties for the resource that was requested as part of the Properties subrequest. ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2393
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(PropertiesSubResponseType),
                     propertiesSubResponse.GetType(),
                     "MS-FSSHTTP",
                     2393,
                     @"[Properties Subrequest]The protocol server responds with a Properties SubResponse message, which is of type PropertiesSubResponseType as specified in section 2.3.1.55. ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2148
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(PropertiesSubResponseType),
                     propertiesSubResponse.GetType(),
                     "MS-FSSHTTP",
                     2148,
                     @"[In SubResponseElementGenericType] Depending on the Type attribute specified in the SubRequest element, the SubResponseElementGenericType MUST take one of the forms: PropertiesSubResponseType");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2166
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(PropertiesSubResponseType),
                     propertiesSubResponse.GetType(),
                     "MS-FSSHTTP",
                     2166,
                     @"[In SubResponseType] The SubResponseElementGenericType takes one of the following forms: PropertiesSubResponseType.");

            ErrorCodeType errorCode;
            site.Assert.IsTrue(Enum.TryParse<ErrorCodeType>(propertiesSubResponse.ErrorCode, true, out errorCode), "Fail to convert the error code string {0} to the Enum type ErrorCodeType", propertiesSubResponse.ErrorCode);
            if (errorCode == ErrorCodeType.Success)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2160
                site.CaptureRequirementIfIsNotNull(
                         propertiesSubResponse.SubResponseData,
                         "MS-FSSHTTP",
                         2160,
                         @"[In SubResponseElementGenericType][The SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"" and one of the following conditions is true:] The Type attribute that is specified in the SubRequest element is set to a value of ""Properties"".");
            }

            // Verify requirements related with its base type: SubResponseType
            ValidateSubResponseType(propertiesSubResponse as SubResponseType, site);

            // Verify requirements related with SubResponseDataType
            if (propertiesSubResponse.SubResponseData != null)
            {
                ValidatePropertiesSubResponseDataType(propertiesSubResponse.SubResponseData, site);
            }
        }

        /// <summary>
        /// Capture requirements related with AmIAloneSubResponseDataType
        /// </summary>
        /// <param name="amIAloneSubResponseData">The AmIAloneSubResponseData information</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidatePropertiesSubResponseDataType(PropertiesSubResponseDataType propertiesSubResponseData, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2295
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2295,
                     @"[PropertiesSubResponseDataType]
	<xs:complexType name=""PropertiesSubResponseDataType"">
	  <xs:sequence>
	    <xs:element name=""PropertyIds"" minOccurs=""0"" maxOccurs=""1"" type=""tns:PropertyIdsType""/>
	    <xs:element name=""PropertyValues"" minOccurs=""0"" maxOccurs=""1"" type=""tns:PropertyValuesType""/>
	  </xs:sequence>
	</xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2136
            // The SubResponseData of PropertiesSubResponse is of type PropertiesSubResponseDataType, so if propertiesSubResponse.SubResponseData is not null, then MS-FSSHTTP_R2136 can be captured.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(PropertiesSubResponseDataType),
                     propertiesSubResponseData.GetType(),
                     "MS-FSSHTTP",
                     2136,
                     @"[In SubResponseDataGenericType][SubResponseDataGenericType MUST take one of the forms described in the following table]PropertiesSubResponseDataType:Type definition for Properties subresponse data.");
        }
    }
}