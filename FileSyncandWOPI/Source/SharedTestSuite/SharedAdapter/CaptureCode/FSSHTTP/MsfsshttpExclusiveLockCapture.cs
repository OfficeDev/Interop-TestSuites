namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This partial part of MsfsshttpAdapterCapture class is used to test Exclusive lock response related requirements.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related to exclusive lock sub request of type, "Convert to schema lock with co-authoring transition tracked".
        /// </summary>
        /// <param name="exclusiveLockSubResponse">Containing the subResponse information.</param>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void ValidateExclusiveLockSubResponse(ExclusiveLockSubResponseType exclusiveLockSubResponse, ITestSite site)
        {
            ValidateSubResponseType(exclusiveLockSubResponse as SubResponseType, site);

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R659
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     659,
                     @"[In ExclusiveLockSubResponseType][ExclusiveLockSubResponseType schema is:]
                     <xs:complexType name=""ExclusiveLockSubResponseType"">
                      <xs:complexContent>
                        <xs:extension base=""tns:SubResponseType"">
                         <xs:sequence minOccurs=""1"" maxOccurs=""1"">
                          <xs:element name=""SubResponseData"" type=""tns:ExclusiveLockSubResponseDataType"" />
                         </xs:sequence>
                        </xs:extension>
                       </xs:complexContent>
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1216
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1216,
                     @"[In ExclusiveLock Subrequest][The protocol client sends an exclusive lock SubRequest message] The protocol server responds with an exclusive lock SubResponse message, which is of type ExclusiveLockSubResponseType as specified in section 2.3.1.12.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R4689
            // if can launch this method, the schema matches.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(ExclusiveLockSubResponseType),
                     exclusiveLockSubResponse.GetType(),
                     "MS-FSSHTTP",
                     4689,
                     @"[In SubResponseElementGenericType] Depending on the Type attribute specified in the SubRequest element, the SubResponseElementGenericType MUST take one of the forms: ExclusiveLockSubResponseType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R5744
            // if can launch this method, the schema matches.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(ExclusiveLockSubResponseType),
                     exclusiveLockSubResponse.GetType(),
                     "MS-FSSHTTP",
                     5744,
                     @"[In SubResponseType] The SubResponseElementGenericType takes one of the following forms: ExclusiveLockSubResponseType.");

            if (string.Compare("Success", exclusiveLockSubResponse.ErrorCode, StringComparison.OrdinalIgnoreCase) == 0)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R264
                // if can launch this method, the schema matches
                bool isVerifyR264 = exclusiveLockSubResponse.SubResponseData != null;

                // If popup the assert, the case run fail. 
                site.Assert.IsTrue(
                            isVerifyR264,
                            "For requirement MS-FSSHTTPB_R264, the SubResponseData should not be NULL.");

                // If the above logic is right, MS-FSSHTTP_R264 can be captured.
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         264,
                         @"[In SubResponseElementGenericType][The SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"" and one of the following conditions is true:] The Type attribute that is specified in the SubRequest element is set to a value of ""ExclusiveLock"".");
            }

            if (exclusiveLockSubResponse.SubResponseData != null)
            {
                ValidateExclusiveLockSubResponseDataType(exclusiveLockSubResponse.SubResponseData, site);
            }
        }

        /// <summary>
        /// Capture requirements related with ExclusiveLockSubResponseDataType.
        /// </summary>
        /// <param name="exclusiveLockSubResponseDataType">The ExclusiveLockSubResponseDataType</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateExclusiveLockSubResponseDataType(ExclusiveLockSubResponseDataType exclusiveLockSubResponseDataType, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R645
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     645,
                     @"[In ExclusiveLockSubResponseDataType][ExclusiveLockSubResponseDataType schema is:]
                     <xs:complexType name=""ExclusiveLockSubResponseDataType"">
                         <xs:attribute name=""CoauthStatus"" type=""tns:CoauthStatusType"" use=""optional""/>
                         <xs:attribute name=""TransitionID"" type=""tns:guid"" use=""optional""/>
                     </xs:complexType>");
            if (exclusiveLockSubResponseDataType.CoauthStatusSpecified)
            {
                // Verify the CoauthStatusType schema related requirements. 
                ValidateCoauthStatusType(site);
            }

            if (exclusiveLockSubResponseDataType.TransitionID != null)
            {
                // Verify the GUID schema related requirements. 
                ValidateGUID(site);
            }

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1393
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1393,
                     @"[In SubResponseDataGenericType][SubResponseDataGenericType MUST take one of the forms described in the following table] ExclusiveLockSubResponseDataType: Type definition for exclusive lock subresponse data.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R660
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     660,
                     @"[In ExclusiveLockSubResponseType] SubResponseData: A ExclusiveLockSubResponseDataType that specifies exclusive lock-related information provided by the protocol server that was requested as part of the exclusive lock subrequest.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1225
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1225,
                     @"[In ExclusiveLock Subrequest] The ExclusiveLockSubResponseDataType defines the type of the SubResponseData element inside the exclusive lock SubResponse element.");
        }
    }
}