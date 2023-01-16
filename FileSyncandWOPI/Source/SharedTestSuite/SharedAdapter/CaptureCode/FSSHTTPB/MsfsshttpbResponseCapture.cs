namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Globalization;
    using Microsoft.Protocols.TestTools;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// This is the partial part of the class MsfsshttpbAdapterCapture for MS-FSSHTTPB response part.
    /// </summary>
    public partial class MsfsshttpbAdapterCapture
    {
        /// <summary>
        /// This method is used to test Response Message Syntax related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyFsshttpbResponse(FsshttpbResponse instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Response Message Syntax related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type FsshttpbResponse is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R534, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     534,
                     @"[In Response Message Syntax]  Protocol Version (2bytes): An unsigned integer that specifies the protocol schema version number used in this request.");

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4124, site))
            {
                bool isVerified4124 = instance.ProtocolVersion == 12 || instance.ProtocolVersion == 13 || instance.ProtocolVersion == 14;
                // Capture requirement MS-FSSHTTPB_R4124, if the protocol version number equals to 12, 13 or 14. 
                site.CaptureRequirementIfIsTrue(
                    isVerified4124,
                    "MS-FSSHTTPB",
                    4124,
                    @"[In Appendix B: Product Behavior] The valid values for this field [Protocol Version] are 12, 13 and 14. ( SharePoint Server 2010 and SharePoint Workspace 2010 follow this behavior.)");
            }

            // Directly capture requirement MS-FSSHTTPB_R536, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     536,
                     @"[In Response Message Syntax] Minimum Version (2 bytes): An unsigned integer that specifies the oldest version of the protocol schema that this schema is compatible with.");

            // Capture requirement MS-FSSHTTPB_R537, if the minimum version number equals to 11. 
            site.CaptureRequirementIfAreEqual<int>(
                     11,
                     instance.MinimumVersion,
                     "MS-FSSHTTPB",
                     537,
                     @"[In Response Message Syntax] Minimum Version (2 bytes): This value[Minimum Version] MUST be 11.");

            // Directly capture requirement MS-FSSHTTPB_R538, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     538,
                     @"[In Response Message Syntax] Signature (8 bytes): An unsigned integer that specifies a constant signature, to identify this as a response.");

            // Capture requirement MS-FSSHTTPB_R539, if the signature number equals to 0x9B069439F329CF9D.
            site.CaptureRequirementIfAreEqual<ulong>(
                     0x9B069439F329CF9D,
                     instance.Signature,
                     "MS-FSSHTTPB",
                     539,
                     @"[In Response Message Syntax] Signature (8 bytes): This[Signature] MUST be set to 0x9B069439F329CF9D.");

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.ResponseStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R540, if the response header equals to StreamObjectHeaderStart32bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.ResponseStart.GetType(),
                     "MS-FSSHTTPB",
                     540,
                     @"[In Response Message Syntax] Response Start (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies a response start.");

            if (instance.Status == true)
            {
                site.Log.Add(
                        LogEntryKind.Debug,
                        "When the status is set, the response error value {0}.",
                        instance.ResponseError.ErrorData.ErrorDetail);
            }
            else
            {
                if (instance.DataElementPackage != null)
                {
                    // Directly capture requirement MS-FSSHTTPB_R2191, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             2191,
                             @"[In Response Message Syntax] Data Element Package (variable): An optional Data Element Package structure (section 2.2.1.12) that specifies data elements corresponding to the sub-responses (section 2.2.3.1).");
                }

                if (instance.CellSubResponses != null && instance.CellSubResponses.Count != 0)
                {
                    // Directly capture requirement MS-FSSHTTPB_R2192, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             2192,
                             @"[In Response Message Syntax] Sub-response (variable): Specifies an array of sub-responses corresponding to the sub-requests as specified in section 2.2.2.1.");

                    // Directly capture requirement MS-FSSHTTPB_R851, if there are no parsing errors.
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             851,
                             @"[In Message Processing Events and Sequencing Rules] The server MUST reply to a well-formed request with a response, as specified in section 2.2.3, which includes a sub-response for each sub-request.");
                }

                // Directly capture requirement MS-FSSHTTPB_R2190, if there are no parsing errors.
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         2190,
                         @"[In Response Message Syntax] Response Data (variable): If the request did not fail, the response data is specified by the following structure [Data Element Package, Sub Responses].");
            }

            // Capture requirement MS-FSSHTTPB_R543, if the reserved value equals to 0.
            site.CaptureRequirementIfAreEqual<int>(
                     0,
                     instance.Reserved,
                     "MS-FSSHTTPB",
                     543,
                     @"[In Response Message Syntax] Reserved (7 bits): A 7-bit reserved field that MUST be set to zero.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.ResponseEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.ResponseStart, site);

            // Capture requirement MS-FSSHTTPB_R546, if the response header equals to StreamObjectHeaderEnd16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd16bit),
                     instance.ResponseEnd.GetType(),
                     "MS-FSSHTTPB",
                     546,
                     @"[In Query Access] Response Error (variable): If the Put Changes operation will succeed, the Response Error will have an error type of HRESULT error.");
        }

        /// <summary>
        /// This method is used to test Sub-Responses related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyFsshttpbSubResponse(FsshttpbSubResponse instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Sub-Responses related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type FsshttpbSubResponse is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R548, if the header start type is StreamObjectHeaderStart32bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     548,
                     @"[In Sub-Responses] Sub-response Start (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies a sub-response start.");

            RequestTypes expectTypeValue = MsfsshttpbSubRequestMapping.GetSubRequestType((int)instance.RequestID.DecodedValue, site);

            // Capture requirement MS-FSSHTTPB_R549, if the request type is consistent with the expected value.
            site.CaptureRequirementIfAreEqual<RequestTypes>(
                     expectTypeValue,
                     (RequestTypes)instance.RequestType.DecodedValue,
                     "MS-FSSHTTPB",
                     549,
                     @"[In Sub-Responses] Request ID (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the request number this sub-response is for.");

            // Capture requirement MS-FSSHTTPB_R550, if the request type is consistent with the expected value.
            site.CaptureRequirementIfAreEqual<RequestTypes>(
                     expectTypeValue,
                     (RequestTypes)instance.RequestType.DecodedValue,
                     "MS-FSSHTTPB",
                     550,
                     @"[In Sub-Responses] Request Type (variable): A compact unsigned 64-bit integer that specifies the request type (section 2.2.1.6) matching the request.");

            // Verify the request types related requirements.
            this.VerifyRequestTypes((RequestTypes)instance.RequestType.DecodedValue, site);

            if (instance.Status == true)
            {
                site.Log.Add(
                        LogEntryKind.Debug,
                        "When the status is set, the response error should exist.");

                // Capture requirement MS-FSSHTTPB_R552, if the response error is exist when the status is set.
                site.CaptureRequirementIfIsNotNull(
                         instance.ResponseError,
                         "MS-FSSHTTPB",
                         552,
                         @"[In Sub-Responses] A - Status (1 bit): If set, A Response Error (section 2.2.3.2) MUST follow.");

                // Capture requirement MS-FSSHTTPB_R551, if the response error is exist when the status is set.
                site.CaptureRequirementIfIsNotNull(
                         instance.ResponseError,
                         "MS-FSSHTTPB",
                         551,
                         @"[In Sub-Responses] A - Status (1 bit): If set, a bit that specifies the sub-request has failed.");

                // Directly capture requirement MS-FSSHTTPB_R555, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         555,
                         @"[In Sub-Responses] Sub-response data (variable): A Response Error that specifies the error information about failure of the sub-request.");
            }
            else
            {
                site.Assert.IsNotNull(
                            instance.SubResponseData,
                            "When the status is not set, the response data cannot be null.");

                switch ((int)instance.RequestType.DecodedValue)
                {
                    case 1:
                        site.Assert.AreEqual<Type>(
                                    typeof(QueryAccessSubResponseData),
                                    instance.SubResponseData.GetType(),
                                    "When the request type value equals to 1, then the response data MUST be the type QueryAccessSubResponseData");

                        // Directly capture requirement MS-FSSHTTPB_R578, if the above assertion was validated.
                        site.CaptureRequirement(
                                 "MS-FSSHTTPB",
                                 578,
                                 @"[In Sub-Responses][Request Type is set to ]1 [specifies the Sub-response data is  for the ]Query Access (section 2.2.3.1.1)[operation].");
                        break;

                    case 2:
                        site.Assert.AreEqual<Type>(
                                    typeof(QueryChangesSubResponseData),
                                    instance.SubResponseData.GetType(),
                                    "When the request type value equals to 2, then the response data MUST be the type QueryChangesSubResponseData");

                        // Directly capture requirement MS-FSSHTTPB_R579, if the above assertion was validated.
                        site.CaptureRequirement(
                                 "MS-FSSHTTPB",
                                 579,
                                 @"[In Sub-Responses][Request Type is set to ]2 [specifies the Sub-response data is  for the ]Query Changes (section 2.2.3.1.2)[operation].");
                        break;

                    case 5:
                        site.Assert.AreEqual<Type>(
                                    typeof(PutChangesSubResponseData),
                                    instance.SubResponseData.GetType(),
                                    "When the request type value equals to 5, then the response data MUST be the type PutChangesSubResponseData");

                        // Directly capture requirement MS-FSSHTTPB_R582, if the above assertion was validated.
                        site.CaptureRequirement(
                                 "MS-FSSHTTPB",
                                 582,
                                 @"[In Sub-Responses][Request Type is set to ]5 [specifies the Sub-response data is  for the ]Put Changes (section 2.2.3.1.3)[operation].");
                        break;

                    case 11:
                        site.Assert.AreEqual<Type>(
                                   typeof(AllocateExtendedGuidRangeSubResponseData),
                                   instance.SubResponseData.GetType(),
                                   "When the request type value equals to 11, then the response data MUST be the type AllocateExtendedGuidRangeSubResponseData");
                        break;

                    default:
                        site.Assert.Fail("Unsupported sub request type " + (int)instance.RequestType.DecodedValue);
                        break;
                }

                // If the above asserts are validate, the requirement MS-FSSHTTPB_R5551 can be captured.
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         5551,
                         @"[In Sub-Responses] depending on the request type (section 2.2.1.6), [Sub-response data (variable)] specifies additional data. See the following table for details:[of request type number value and operation]");
            }

            // Capture requirement MS-FSSHTTPB_R553, if the reserved value equals to 0. 
            site.CaptureRequirementIfAreEqual<byte>(
                     0,
                     instance.Reserved,
                     "MS-FSSHTTPB",
                     553,
                     @"[In Sub-Responses] Reserved (7 bits): A 7 bit reserved field that MUST be set to zero.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Capture requirement MS-FSSHTTPB_R586, if the end type is StreamObjectHeaderEnd16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd16bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     586,
                     @"[In Sub-Responses] Sub-response End (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.4) that specifies a sub-response end.");
        }

        /// <summary>
        /// This method is used to test Query Access related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyQueryAccessSubResponseData(QueryAccessSubResponseData instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Query Access related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type QueryAccessSubResponseData is null due to parsing error or type casting error.");
            }

            site.Assert.IsNotNull(instance.ReadAccessResponse, "The QueryAccessSubResponseData::ReadAccessResponse cannot be null.");
            site.Assert.IsNotNull(instance.ReadAccessResponse, "The QueryAccessSubResponseData::WriteAccessResponse cannot be null.");
        }

        /// <summary>
        /// This method is used to test ReadAccessResponse related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyReadAccessResponse(ReadAccessResponse instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the ReadAccessResponse related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ReadAccessResponse is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R588, if the ReadAccessResponse stream object start header is StreamObjectHeaderStart32bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     588,
                     @"[In Query Access] Read Access Response Start (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies a read access response start.");

            // Directly capture requirement MS-FSSHTTPB_R589, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     589,
                     @"[In Query Access] Response Error (variable): A Response Error (section 2.2.3.2) that specifies read access permission.");

            // Directly capture requirement MS-FSSHTTPB_R943, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     943,
                     @"[In Query Access] Response Error (variable): This error[Response Error] is received in response to any read request made by the client.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Directly capture requirement MS-FSSHTTPB_R590, if ReadAccessResponse stream object end header is StreamObjectHeaderEnd16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd16bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     590,
                     @"[In Query Access] Read Access Response End (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.4) that specifies a read access response end.");
        }

        /// <summary>
        /// This method is used to test WriteAccessResponse related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyWriteAccessResponse(WriteAccessResponse instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the WriteAccessResponse related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type WriteAccessResponse is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R591, if the WriteAccessResponse stream object start header is StreamObjectHeaderStart32bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     591,
                     @"[In Query Access] Write Access Response Start (4 bytes): A 32-bit Stream Object Header that specifies a write access response start.");

            // Directly capture requirement MS-FSSHTTPB_R592, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     592,
                     @"[In Query Access] Response Error (variable): A Response Error that specifies write access permission.");

            // Directly capture requirement MS-FSSHTTPB_R945, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     945,
                     @"[In Query Access] Response Error (variable): This error[Response Error] is received in response to a Put Changes request made by the client.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Capture requirement MS-FSSHTTPB_R593, if WriteAccessResponse stream object end header is StreamObjectHeaderEnd16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd16bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     593,
                     @"[In Query Access] Write Access Response End (2 bytes): A 16-bit Stream Object Header that specifies a write access response end.");
        }

        /// <summary>
        /// This method is used to test Query Changes related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyQueryChangesSubResponseData(QueryChangesSubResponseData instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Query Changes related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type QueryChangesSubResponseData is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.QueryChangesResponseStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R595, if the stream object header is StreamObjectHeaderStart32bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.QueryChangesResponseStart.GetType(),
                     "MS-FSSHTTPB",
                     595,
                     @"[In Query Changes] Query Changes Response (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies a Query Changes response.");

            // Directly capture requirement MS-FSSHTTPB_R596, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     596,
                     @"[In Query Changes] Storage Index Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies Storage Index.");

            // Directly capture requirement MS-FSSHTTPB_R598, if the reserved value equals to 1. 
            site.CaptureRequirementIfAreEqual<int>(
                     0,
                     instance.ReservedQueryChanges,
                     "MS-FSSHTTPB",
                     598,
                     @"[In Query Changes] Reserved (6 bits): A 6-bit reserved field that MUST be set to zero.");

            // Directly capture requirement MS-FSSHTTPB_R601, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     601,
                     @"[In Query Changes] Knowledge (variable): A Knowledge (section 2.2.1.13) that specifies the current state of the file on the server.");

            // Directly capture requirement MS-FSSHTTPB_R1341, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     1341,
                     @"[In Query Changes]File Hash (4 bytes): An optional 32-bit Stream Object Header that specifies the beginning of File Hash.");

            // Verify the compound related requirements.
            this.ExpectSingleObject(instance.QueryChangesResponseStart, site);
        }

        /// <summary>
        /// This method is used to test Put Changes related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyPutChangesSubResponseData(PutChangesSubResponseData instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Put Changes related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type PutChangesSubResponseData is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R610, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     610,
                     @"[In Put Changes] Resultant Knowledge (variable): A Knowledge (section 2.2.1.13) that specifies the current state of the file on the server after the changes is merged.");

            if (instance.DiagnosticRequestOptionOutput != null)
            {
                // Directly capture requirement MS-FSSHTTPB_R4096, if the reserved value equals to 0. 
                site.CaptureRequirementIfAreEqual<int>(
                         0,
                         instance.DiagnosticRequestOptionOutput.Reserved,
                         "MS-FSSHTTPB",
                         4096,
                         @"[In Put Changes] Reserved (7 bits): A 7-bit reserved field that MUST be set to zero [and MUST be ignored].");
            }
        }

        /// <summary>
        /// This method is used to test Allocate ExtendedGuid Range related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyAllocateExtendedGuidRangeSubResponseData(AllocateExtendedGuidRangeSubResponseData instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Allocate ExtendedGuid Range related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type AllocateExtendedGuidRangeSubResponseData is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.AllocateExtendedGUIDRangeResponse, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R2203, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     2203,
                     @"[In Allocate ExtendedGuid Range] Allocate ExtendedGuid Range Response (4 bytes): A stream object header (section 2.2.1.5) that specifies an allocate extendedGUID range response.");

            // Directly capture requirement MS-FSSHTTPB_R2204, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     2204,
                     @"[In Allocate ExtendedGuid Range] GUID Component (16 bytes): A GUID that specifies the GUID portion of the reserved extended GUIDs (section 2.2.1.7).");

            // Directly capture requirement MS-FSSHTTPB_R2205, if there are no parsing errors. 
            // This requirement is partially captured, only the type "A compact unsigned 64-bit integer" can be captured.
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     2205,
                     @"[In Allocate ExtendedGuid Range] Integer Range Min (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the first integer element in the range of extended GUIDs.");

            // Directly capture requirement MS-FSSHTTPB_R2206, if there are no parsing errors. 
            // This requirement is partially captured, only the type "A compact unsigned 64-bit integer" can be captured.
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     2206,
                     @"[In Allocate ExtendedGuid Range] [Integer Range Max (variable)] specifies the last + 1 integer element in the range of extended GUIDs.");

            bool isVerifiedR99062 = instance.IntegerRangeMax.DecodedValue >= 1000 && instance.IntegerRangeMax.DecodedValue <= 100000;
            site.Log.Add(
                LogEntryKind.Debug,
                "The Integer Range Max should be in the range 1000 and 100000, its actual value is {0}",
                instance.IntegerRangeMax.DecodedValue);

            site.CaptureRequirementIfIsTrue(
                     isVerifiedR99062,
                     "MS-FSSHTTPB",
                     99062,
                     @"[In Allocate ExtendedGuid Range] Integer Range Max (variable): A compact unsigned 64-bit integer with a minimum allowed value of 1,000 and maximum allowed value of 100,000.");

            // Verify the compound related requirements.
            this.ExpectSingleObject(instance.AllocateExtendedGUIDRangeResponse, site);
        }

        /// <summary>
        /// This method is used to test Response Error related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyResponseError(ResponseError instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Response Error related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ResponseError is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R617, if the stream object header is StreamObjectHeaderStart32bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     617,
                     @"[In Response Error] Error Start (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies an error start.");

            // Directly capture requirement MS-FSSHTTPB_R618, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     618,
                     @"[In Response Error] Error Type GUID (16 bytes): A GUID that specifies the error type.");

            bool responseErrorFlag = instance.ErrorTypeGUID == new Guid(ResponseError.CellErrorGuid)
                                   || instance.ErrorTypeGUID == new Guid(ResponseError.ProtocolErrorGuid)
                                   || instance.ErrorTypeGUID == new Guid(ResponseError.Win32ErrorGuid)
                                   || instance.ErrorTypeGUID == new Guid(ResponseError.HresultErrorGuid);

            site.Assert.IsTrue(
                        responseErrorFlag,
                        "Actual the error type guid {0}, which should be either 5A66A756-87CE-4290-A38B-C61C5BA05A67,7AFEAEBF-033D-4828-9C31-3977AFE58249, 32C39011-6E39-46C4-AB78-DB41929D679E or 8454C8F2-E401-405A-A198-A10B6991B56E for requirement MS-FSSHTTPB_R619",
                        instance.ErrorTypeGUID.ToString());

            // Directly capture requirement MS-FSSHTTPB_R619, if the responseErrorFlag is true;
            site.CaptureRequirementIfIsTrue(
                     responseErrorFlag,
                     "MS-FSSHTTPB",
                     619,
                     @"[In Response Error] Error Type GUID (16 bytes): The following table contains the possible values for the error type: [the value of the Error Type GUID field must be {5A66A756-87CE-4290-A38B-C61C5BA05A67},{7AFEAEBF-033D-4828-9C31-3977AFE58249}, {32C39011-6E39-46C4-AB78-DB41929D679E}, {8454C8F2-E401-405A-A198-A10B6991B56E}.");

            switch (instance.ErrorTypeGUID.ToString("D").ToUpper(CultureInfo.CurrentCulture))
            {
                case ResponseError.CellErrorGuid:

                    // Capture requirement MS-FSSHTTPB_R620, if the error data type is CellError. 
                    site.CaptureRequirementIfAreEqual<Type>(
                             typeof(CellError),
                             instance.ErrorData.GetType(),
                             "MS-FSSHTTPB",
                             620,
                             @"[In Response Error] Error Type GUID field is set to {5A66A756-87CE-4290-A38B-C61C5BA05A67}[ specifies the error type is ]Cell Error (section 2.2.3.2.1).");

                    break;

                case ResponseError.ProtocolErrorGuid:

                    // All the serial number null values related requirements can be located here.
                    site.Log.Add(LogEntryKind.Debug, "Runs for ProtocolErrorGuid verification logic with the error code {0}.", instance.ErrorData.ErrorDetail);
                    break;

                case ResponseError.Win32ErrorGuid:

                    // Capture requirement MS-FSSHTTPB_R622, if the error data type is Win32 error. 
                    site.CaptureRequirementIfAreEqual<Type>(
                             typeof(Win32Error),
                             instance.ErrorData.GetType(),
                             "MS-FSSHTTPB",
                             622,
                             @"[In Response Error] Error Type GUID field is set to {32C39011-6E39-46C4-AB78-DB41929D679E}[ specifies the error type is ]Win32 Error (section 2.2.3.2.3).");

                    break;

                case ResponseError.HresultErrorGuid:

                    // Capture requirement MS-FSSHTTPB_R623, if the error data type is HRESULTError.  
                    site.CaptureRequirementIfAreEqual<Type>(
                             typeof(HRESULTError),
                             instance.ErrorData.GetType(),
                             "MS-FSSHTTPB",
                             623,
                             @"[In Response Error] Error Type GUID field is set to {8454C8F2-E401-405A-A198-A10B6991B56E}[ specifies the error type is ]HRESULT Error (section 2.2.3.2.4).");

                    break;

                default:
                    site.Assert.Fail("Unsupported error type GUID " + instance.ErrorTypeGUID.ToString());
                    break;
            }

            // Directly capture requirement MS-FSSHTTPB_R625, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     625,
                     @"[In Response Error] Error Data (variable): A structure that specifies the error data based on the Error Type GUID.");

            if (instance.ChainedError != null)
            {
                // Directly capture requirement MS-FSSHTTPB_R631, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         631,
                         @"[In Response Error] Chained Error (variable): An optional Response Error that specifies the chained error information.");
            }

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Directly capture requirement MS-FSSHTTPB_R632, if the stream object end is StreamObjectHeaderEnd16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd16bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     632,
                     @"[In Response Error] Error End (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.4) that specifies an error end.");
        }

        /// <summary>
        /// This method is used to test Response Error related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyErrorStringSupplementalInfo(ErrorStringSupplementalInfo instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the ErrorStringSupplementalInfo related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ErrorStringSupplementalInfo is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R99063, if the header type is StreamObjectHeaderStart32bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     99063,
                     @"[In Response Error] Error String Supplemental Info Start (4 bytes, optional): Zero or one 32-bit Stream Object Header (section 2.2.1.5.2) that specifies an error string supplemental info start. ");

            // Directly capture requirement MS-FSSHTTPB_R99064, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     99064,
                     @"[In Response Error] Error String Supplemental Info (variable, optional): A string item (section 2.2.1.4) that specifies the supplemental information of the error string for the error string supplemental info start.");

            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Cell Error related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyCellError(CellError instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Cell Error related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type CellError is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R634, if the header type is StreamObjectHeaderStart32bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     634,
                     @"[In Cell Error] Error Cell (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies an error cell.");

            // Directly capture requirement MS-FSSHTTPB_R635, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     635,
                     @"[In Cell Error] Error Code (4 bytes): An unsigned integer that specifies the error code.");

            bool cellErrorFlag = (CellErrorCode)instance.ErrorCode == CellErrorCode.Unknownerror
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.InvalidObject
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Invalidpartition
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Requestnotsupported
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Storagereadonly
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.RevisionIDnotfound
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Badtoken
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Requestnotfinished
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Incompatibletoken
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Scopedcellstorage
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Coherencyfailure
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Cellstoragestatedeserializationfailure
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Incompatibleprotocolversion
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Referenceddataelementnotfound
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Requeststreamschemaerror
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Responsestreamschemaerror
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Unknownrequest
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Storagefailure
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Storagewriteonly
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Invalidserialization
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Dataelementnotfound
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Invalidimplementation
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Incompatibleoldstorage
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Incompatiblenewstorage
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.IncorrectcontextfordataelementID
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Objectgroupduplicateobjects
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Objectreferencenotfoundinrevision
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Mergecellstoragestateconflict
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Unknownquerychangesfilter
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Unsupportedquerychangesfilter
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Unabletoprovideknowledge
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.DataelementmissingID
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Dataelementmissingserialnumber
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Requestargumentinvalid
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Partialchangesnotsupported
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Storebusyretrylater
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.GUIDIDtablenotsupported
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Dataelementcycle
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Fragmentknowledgeerror
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Fragmentsizemismatch
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Fragmentsincomplete
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Fragmentinvalid
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.Abortedafterfailedputchanges
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.FailedNoUpgradeableContents
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.UnableAllocateAdditionalExtendedGuids
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.SiteReadonlyMode
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.MultiRequestPartitionReachQutoa
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.ExtendedGuidCollision
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.InsufficientPermisssions
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.ServerThrottling
                              || (CellErrorCode)instance.ErrorCode == CellErrorCode.FileTooLarge;

            site.Assert.IsTrue(
                        cellErrorFlag,
                        "The error code value for the cell error MUST be 1~47, 79, 106, 108, 111, 112, 113, 114 and 115, but except 10, 14, 17, 30");

            // Directly capture requirement MS-FSSHTTPB_R636, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     636,
                     @"[In Cell Error] Error Code (4 bytes): The following table contains the possible error codes: [the value of Error Code must be (1~47,except 10, 14, 17, 30) and 79, 106, 108, 111~115].");

            // Verify the compound related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Protocol Error related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyProtocolError(ProtocolError instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Protocol Error related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ProtocolError is null due to parsing error or type casting error.");
            }

            // All the serial number null values related requirements can be located here.
            site.Log.Add(LogEntryKind.Debug, "Runs for ProtocolErrorGuid verification logic with the error code {0}.", instance.ErrorCode);
        }

        /// <summary>
        /// This method aims to test DiagnosticRequestOptionOutput related adapter requirements, but // No source code is needed for this method and the method is needed for reflection.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyDiagnosticRequestOptionOutput(DiagnosticRequestOptionOutput instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Win32 Error related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type Win32Error is null due to parsing error or type casting error.");
            }

            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R4094, if the header is StreamObjectHeaderStart32bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     4094,
                     @"[In Put Changes] Diagnostic Request Option Output Header (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies a Diagnostic Request Option Output.");
        }

        /// <summary>
        /// This method is used to test AppliedStorageIndex related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyPutChangesResponse(PutChangesResponse instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Win32 Error related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type Win32Error is null due to parsing error or type casting error.");
            }

            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R99059, if the header is StreamObjectHeaderStart32bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     99059,
                     @"[In Put Changes] Put Changes Response (4 bytes):  An optional 32-bit Stream Object Header (section 2.2.1.5.2) that specifies a Put Changes response.");

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4126, site))
            {
                // Directly capture requirement MS-FSSHTTPB_R4126, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         4126,
                         @"[In Appendix B: Product Behavior] Implementation does support the Applied Storage Index Id field. (Microsoft Office 2013 and Microsoft SharePoint 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4128, site))
            {
                // Directly capture requirement MS-FSSHTTPB_R4128, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         4128,
                         @"[In Appendix B: Product Behavior] Implementation does support the Data Elements Added field. (Microsoft Office 2013 and Microsoft SharePoint 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4130, site))
            {
                // Directly capture requirement MS-FSSHTTPB_R4130, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         4130,
                         @"[In Appendix B: Product Behavior] Implementation does support the Diagnostic Request Option Output field. (Microsoft Office 2013 and Microsoft SharePoint 2013 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This method is used to test Win32 Error related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyWin32Error(Win32Error instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Win32 Error related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type Win32Error is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R769, if the header is StreamObjectHeaderStart32bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     769,
                     @"[In Win32 Error] Error Win32 (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies an error win32.");

            // Directly capture requirement MS-FSSHTTPB_R770, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     770,
                     @"[In Win32 Error] Error Code (4 bytes): An unsigned integer that specifies the Win32 Error code.");

            // Verify the compound related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test HRESULT Error related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyHRESULTError(HRESULTError instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the HRESULT Error related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type HRESULTError is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R772, if the header is StreamObjectHeaderStart32bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     772,
                     @"[In HRESULT Error] Error HRESULT (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies an Error HRESULT.");

            // Directly capture requirement MS-FSSHTTPB_R773, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     773,
                     @"[In HRESULT Error] Error Code (4 bytes): An unsigned integer that specifies the HRESULT Error code.");

            // Verify the compound related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Request Type Enumeration related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyRequestTypes(RequestTypes instance, ITestSite site)
        {
            // Directly capture requirement MS-FSSHTTPB_R176, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     176,
                     @"[In Request Type Enumeration][Request Type Enumeration] Specifies the sub-request type (section 2.2.2.1) to indicate the operation being requested, or the sub-response type (section 2.2.3.1) to indicate the response per the request.");

            // Directly capture requirement MS-FSSHTTPB_R177, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     177,
                     @"[In Request Type Enumeration] The following table enumerates the values for each operation:[Its value must be one of 1,2,5,11].");

            switch ((int)instance)
            {
                case 1:

                    // Directly capture requirement MS-FSSHTTPB_R178, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             178,
                             @"[In Request Type Enumeration][Request Type Enumeration is set to ]1 [specifies a request or response for the ]Query Access[operation].");

                    // Directly capture requirement MS-FSSHTTPB_R863, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             863,
                             @"[In Query Access Sub-Request Processing] The server MUST reply back to the client with a Query Access sub-response, as specified in section 2.2.3.1.1.");

                    break;

                case 2:

                    // Directly capture requirement MS-FSSHTTPB_R179, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             179,
                             @"[In Request Type Enumeration][Request Type Enumeration is set to ]2 [specifies a request or response for the ]Query Changes[operation].");

                    // Directly capture requirement MS-FSSHTTPB_R865, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             865,
                             @"[In Query Changes Sub-Request Processing] The server MUST reply back to the client with a Query Changes sub-response, as specified in section 2.2.3.1.2. [<28>]");

                    break;

                case 5:

                    // Directly capture requirement MS-FSSHTTPB_R182, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             182,
                             @"[In Request Type Enumeration][Request Type Enumeration is set to ]5 [specifies a request or response for the]Put Changes[operation].");

                    // Directly capture requirement MS-FSSHTTPB_R872, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             872,
                             @"[In Put Changes Sub-Request Processing] The server MUST reply back to the client with a Put Changes sub-response, as specified in section 2.2.3.1.3.");

                    break;

                case 11:

                    // Directly capture requirement MS-FSSHTTPB_R2097, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             2097,
                             @"[In Request Type Enumeration][Request Type Enumeration is set to ]11 [specifies a request or response for the] Allocate Extended Guid Range[operation].");

                    break;

                default:
                    site.Assert.Fail("Unsupported request types " + (int)instance);
                    break;
            }
        }
    }
}