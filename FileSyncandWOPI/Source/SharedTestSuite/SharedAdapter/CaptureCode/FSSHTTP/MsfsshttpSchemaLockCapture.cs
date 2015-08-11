namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This partial part of MsfsshttpAdapterCapture class is used to test schema lock response related requirements.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related to SchemaLock sub response.
        /// </summary>
        /// <param name="schemaLockSubResponse">The schemaLock response.</param>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void ValidateSchemaLockSubResponse(SchemaLockSubResponseType schemaLockSubResponse, ITestSite site)
        {
            ValidateSubResponseType(schemaLockSubResponse as SubResponseType, site);

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1120
            // if can launch this method, the schema matches
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1120,
                     @"[In SchemaLock Subrequest][The protocol client sends a schema lock SubRequest message, which is of type SchemaLockSubRequestType,] The protocol server responds with a schema lock SubResponse message, which is of type SchemaLockSubResponseType as specified in section 2.3.1.16.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R4690
            // if can launch this method, the schema matches
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(SchemaLockSubResponseType),
                     schemaLockSubResponse.GetType(),
                     "MS-FSSHTTP",
                     4690,
                     @"[In SubResponseElementGenericType] Depending on the Type attribute specified in the SubRequest element, the SubResponseElementGenericType MUST take one of the forms:  SchemaLockSubResponseType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R5743
            // if can launch this method, the schema matches
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(SchemaLockSubResponseType),
                     schemaLockSubResponse.GetType(),
                     "MS-FSSHTTP",
                     5743,
                     @"[In SubResponseType] The SubResponseElementGenericType takes one of the following forms: SchemaLockSubResponseType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R725
            // if can launch this method, the schema matches
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     725,
                     @"[In SchemaLockSubResponseType][SchemaLockSubResponseType schema is:]
                     <xs:complexType name=""SchemaLockSubResponseType"">
                      <xs:complexContent>
                       <xs:extension base=""tns:SubResponseType"">
                        <xs:sequence minOccurs=""1"" maxOccurs=""1"">
                          <xs:element name=""SubResponseData"" type=""tns:SchemaLockSubResponseDataType"" />
                        </xs:sequence>
                       </xs:extension>
                      </xs:complexContent>
                     </xs:complexType>");

            if (string.Compare("Success", schemaLockSubResponse.ErrorCode, StringComparison.OrdinalIgnoreCase) == 0)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R265
                // if can launch this method, the schema matches
                bool isVerifyR265 = schemaLockSubResponse.SubResponseData != null;

                // If popup the assert, the case run fail. 
                site.Assert.IsTrue(
                            isVerifyR265,
                            "For requirement MS-FSSHTTPB_R265, the SubResponseData should not be NULL.");

                // If the above logic is right, MS-FSSHTTP_R265 can be captured.
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         265,
                         @"[In SubResponseElementGenericType][The SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"" and one of the following conditions is true:] The Type attribute that is specified in the SubRequest element is set to a value of ""SchemaLock"".");
            }

            if (schemaLockSubResponse.SubResponseData != null)
            {
                ValidateSchemaLockSubResponseDataType(schemaLockSubResponse.SubResponseData, site);
            }
        }

        /// <summary>
        /// Capture requirements related with SchemaLockSubResponseDataType.
        /// </summary>
        /// <param name="schemaLockSubResponseData">The SchemaLockSubResponseDataType</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateSchemaLockSubResponseDataType(SchemaLockSubResponseDataType schemaLockSubResponseData, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1136
            // if can launch this method, the schema matches
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(SchemaLockSubResponseDataType),
                     schemaLockSubResponseData.GetType(),
                     "MS-FSSHTTP",
                     1136,
                     @"[In SchemaLock Subrequest] The SubResponseData element returned for a schema lock subrequest is of type SchemaLockSubResponseDataType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1394
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1394,
                     @"[In SubResponseDataGenericType][SubResponseDataGenericType MUST take one of the forms described in the following table] SchemaLockSubResponseDataType: Type definition for schema lock subresponse data.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1132
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1132,
                     @"[In SchemaLock Subrequest] The SchemaLockSubResponseDataType defines the type of the SubResponseData element inside the schema lock SubResponse element.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R726
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     726,
                     @"[In SchemaLockSubResponseType] SubResponseData: A SchemaLockSubResponseDataType that specifies schema lock-related information provided by the protocol server that was requested as part of the schema lock subrequest.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R709
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     709,
                     @"[In SchemaLockSubResponseDataType][SchemaLockSubResponseDataType schema is:]
                     <xs:complexType name=""SchemaLockSubResponseDataType"">
                         <xs:attribute name=""LockType"" type=""tns:LockTypes"" use=""optional"" />
                         <xs:attribute name=""ExclusiveLockReturnReason"" type=""tns:ExclusiveLockReturnReasonTypes"" use=""optional"" />
                     </xs:complexType>");

            if (schemaLockSubResponseData.LockTypeSpecified)
            {
                ValidateLockTypes(site);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1135
                // if can launch this method, the schema matches.
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         1135,
                         @"[In SchemaLock Subrequest] The lock type is sent as the LockType attribute in the schema lock SubResponseData element.");
            }

            if (schemaLockSubResponseData.ExclusiveLockReturnReasonSpecified)
            {
                ValidateExclusiveLockReturnReasonTypes(site);
            }
        }
    }
}