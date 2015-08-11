namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This partial part of MsfsshttpAdapterCapture class is used to test coauthoring response related requirements.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Prevents a default instance of the MsfsshttpAdapterCapture class from being created
        /// </summary>
        private MsfsshttpAdapterCapture()
        { 
        }

        /// <summary>
        /// Capture requirements related to coauthoring sub response.
        /// </summary>
        /// <param name="coauthSubResponse">The coauthoring response.</param>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void ValidateCoauthSubResponse(CoauthSubResponseType coauthSubResponse, ITestSite site)
        {
            ValidateSubResponseType(coauthSubResponse as SubResponseType, site);

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R4687
            // if can launch this method, the schema matches.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(CoauthSubResponseType),
                     coauthSubResponse.GetType(),
                     "MS-FSSHTTP",
                     4687,
                     @"[In SubResponseElementGenericType] Depending on the Type attribute specified in the SubRequest element, the SubResponseElementGenericType MUST take one of the forms: CoauthSubResponseType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R5742
            // if can launch this method, the schema matches.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(CoauthSubResponseType),
                     coauthSubResponse.GetType(),
                     "MS-FSSHTTP",
                     5742,
                     @"[In SubResponseType] The SubResponseElementGenericType takes one of the following forms: CoauthSubResponseType.");

            if (string.Compare("Success", coauthSubResponse.ErrorCode, StringComparison.OrdinalIgnoreCase) == 0)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1496
                bool isVerifyR1496 = coauthSubResponse.SubResponseData != null;

                // If the coauthSubResponse.SubResponseData is null, the case run fail.
                site.Assert.IsTrue(
                            isVerifyR1496,
                            "For requirement MS-FSSHTTP_R1496, the SubResponseData should not be NULL.");

                // if coauthSubResponse.SubResponseData is not null, make sure sent as part of the SubResponse element in a cell storage service response message,
                // so this requirement can be captured.
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         1496,
                         @"[In CoauthSubResponseType] As part of processing the coauthoring subrequest, the SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the following condition is true:
                         The ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"".");

                // Add the log information.
                site.Log.Add(LogEntryKind.Debug, "Verify MS-FSSHTTP_R267:the returned SubResponseData is {0}", coauthSubResponse.SubResponseData.ToString());

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R267
                // if coauthSubResponse.SubResponseData is not null, make sure sent as part of the SubResponse element in a cell storage service response message,
                // so this requirement can be captured.
                bool isVerifyR267 = coauthSubResponse.SubResponseData != null;
                site.CaptureRequirementIfIsTrue(
                         isVerifyR267,
                         "MS-FSSHTTP",
                         267,
                         @"[In SubResponseElementGenericType][The SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"" and one of the following conditions is true:] The Type attribute that is specified in the SubRequest element is set to a value of ""Coauth"".");
            }

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R624
            // If the coauthoring subResponse exists,its schema will have been validated before invoking the method CaptureJoinCoauthoringSessionRelatedRequirements.
            // So MS-FSSHTTP_R624 can be captured.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     624,
                     @"[In CoauthSubResponseType][CoauthSubResponseType schema is:]
                     <xs:complexType name=""CoauthSubResponseType"">
                       <xs:complexContent>
                         <xs:extension base=""tns:SubResponseType"">
                           <xs:sequence minOccurs=""0"" maxOccurs=""1"">
                              <xs:element name=""SubResponseData"" type=""tns:CoauthSubResponseDataType"" />
                           </xs:sequence>
                         </xs:extension>
                       </xs:complexContent>
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R984
            // Since the subResponse is of type CoauthSubResponseType, so this requirement can be captured.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(CoauthSubResponseType),
                     coauthSubResponse.GetType(),
                     "MS-FSSHTTP",
                     984,
                     @"[In Coauth Subrequest][The protocol client sends a coauthoring SubRequest message, which is of type CoauthSubRequestType,] The protocol server responds with a coauthoring SubResponse message, which is of type CoauthSubResponseType as specified in section 2.3.1.8.");

            if (coauthSubResponse.SubResponseData != null)
            {
                ValidateCoauthSubResponseDataType(coauthSubResponse.SubResponseData, site);
            }
        }

        /// <summary>
        /// Capture requirements related with CoauthSubResponseDataType.
        /// </summary>
        /// <param name="coauthSubResponseData">The CoauthSubResponseDataType</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateCoauthSubResponseDataType(CoauthSubResponseDataType coauthSubResponseData, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R994
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     994,
                     @"[In Coauth Subrequest] CoauthSubResponseDataType defines the type of the SubResponseData element inside the coauthoring SubResponse element.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R602
            // If SubResponseData exists,its schema will have been validated before invoking the method CaptureJoinCoauthoringSessionRelatedRequirements.
            // So MS-FSSHTTP_R602 can be captured.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     602,
                     @"[In CoauthSubResponseDataType][CoauthSubResponseDataType schema is:]
                     <xs:complexType name=""CoauthSubResponseDataType"">
                         <xs:attribute name=""LockType"" type=""tns:LockTypes"" use=""optional"" />
                         <xs:attribute name=""CoauthStatus"" type=""tns:CoauthStatusType"" use=""optional""/>
                         <xs:attribute name=""TransitionID"" type=""tns:guid"" use=""optional""/>
                         <xs:attribute name=""ExclusiveLockReturnReason"" type=""tns:ExclusiveLockReturnReasonTypes"" use=""optional"" />
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R999
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     999,
                     @"[In Coauth Subrequest] The SubResponseData element returned for a coauthoring subrequest is of type CoauthSubResponseDataType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1392
            // if can launch this method, the schema matches
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1392,
                     @"[In SubResponseDataGenericType][SubResponseDataGenericType MUST take one of the forms described in the following table] CoauthSubResponseDataType: Type definition for coauthoring subresponse data.");

            if (coauthSubResponseData.LockTypeSpecified)
            {
                ValidateLockTypes(site);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R998
                // if can launch this method, the schema matches.
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         998,
                         @"[In Coauth Subrequest] The lock type is specified in the LockType attribute in the coauthoring SubResponseData element.");
            }

            if (coauthSubResponseData.CoauthStatusSpecified)
            {
                // Verify the CoauthStatusType schema related requirements. 
                ValidateCoauthStatusType(site);
            }

            if (coauthSubResponseData.TransitionID != null)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2075, since TransitionID is marked as "tns:guid" type, the following 
                // 2 requirements can be capture if the schema is valid.
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         2075,
                         @"[In SubResponseDataOptionalAttributes] TransitionID: A guid that specifies the file identifier stored for that file on the protocol server.");

                // Verify the GUID schema related requirements. 
                ValidateGUID(site);
            }

            if (coauthSubResponseData.ExclusiveLockReturnReasonSpecified)
            {
                // Verify the ExclusiveLockReturnReasonTypes schema related requirements.
                ValidateExclusiveLockReturnReasonTypes(site);
            }
        }
    }
}