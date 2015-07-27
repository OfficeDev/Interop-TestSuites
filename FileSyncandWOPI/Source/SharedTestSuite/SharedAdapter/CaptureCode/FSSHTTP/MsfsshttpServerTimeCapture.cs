//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class contains methods which capture requirements related with ServerTime Sub-request.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related with ServerTime Sub-request.
        /// </summary>
        /// <param name="servertimeSubResponse">The SubResponse information.</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateServerTimeSubResponse(ServerTimeSubResponseType servertimeSubResponse, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R4691
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(ServerTimeSubResponseType),
                     servertimeSubResponse.GetType(),
                     "MS-FSSHTTP",
                     4691,
                     @"[In SubResponseElementGenericType] Depending on the Type attribute specified in the SubRequest element, the SubResponseElementGenericType MUST take one of the forms: ServerTimeSubResponseType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R5745
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(ServerTimeSubResponseType),
                     servertimeSubResponse.GetType(),
                     "MS-FSSHTTP",
                     5745,
                     @"[In SubResponseType] The SubResponseElementGenericType takes one of the following forms: ServerTimeSubResponseType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R746
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     746,
                     @"[In ServerTimeSubResponseType][ServerTimeSubResponseType schema is:]
                     <xs:complexType name=""ServerTimeSubResponseType"">
                       <xs:complexContent>
                         <xs:extension base=""tns:SubResponseType"">
                          <xs:sequence minOccurs=""1"" maxOccurs=""1"">
                            <xs:element name=""SubResponseData"" type=""tns:ServerTimeSubResponseDataType""/>
                          </xs:sequence>
                         </xs:extension>
                       </xs:complexContent>
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1331
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(ServerTimeSubResponseType),
                     servertimeSubResponse.GetType(),
                     "MS-FSSHTTP",
                     1331,
                     @"[In ServerTime Subrequest][The protocol client sends a ServerTime SubRequest message, which is of ServerTimeSubRequestType] The protocol server responds with a ServerTime SubResponse message, which is of type ServerTimeSubResponseType as specified in section 2.3.1.19.");

            ErrorCodeType errorCode;
            site.Assert.IsTrue(Enum.TryParse<ErrorCodeType>(servertimeSubResponse.ErrorCode, true, out errorCode), "Fail to convert the error code string {0} to the Enum type ErrorCodeType", servertimeSubResponse.ErrorCode);
            if (errorCode == ErrorCodeType.Success)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R266
                // If servertimeSubResponse.SubResponseData is not null, SubResponseData element is sent.
                site.Log.Add(
                    LogEntryKind.Debug,
                    "For requirement MS-FSSHTTP_R266, the SubResponseData element should be sent as part of the SubResponse element in a cell storage service response message, the actual SubResponseData value is: {0}",
                    servertimeSubResponse.SubResponseData != null ? servertimeSubResponse.SubResponseData.ToString() : "NULL");

                site.CaptureRequirementIfIsNotNull(
                         servertimeSubResponse.SubResponseData,
                         "MS-FSSHTTP",
                         266,
                         @"[In SubResponseElementGenericType][The SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"" and one of the following conditions is true:] The Type attribute that is specified in the SubRequest element is set to a value of ""ServerTime"".");
            }

            // Verify requirements related with its base type: SubResponseType 
            ValidateSubResponseType(servertimeSubResponse as SubResponseType, site);

            // Verify requirements related with ServerTimeSubResponseDataType
            if (servertimeSubResponse.SubResponseData != null)
            {
                ValidateServerTimeSubResponseDataType(servertimeSubResponse.SubResponseData, site);
            }
        }

        /// <summary>
        /// Capture requirements related with ServerTimeSubResponseDataType.
        /// </summary>
        /// <param name="serverSubResponseData">The ServerTimeSubResponseData</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateServerTimeSubResponseDataType(ServerTimeSubResponseDataType serverSubResponseData, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R736
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     736,
                     @"[In ServerTimeSubResponseDataType][ServerTimeSubResponseDataType schema is:]
                     <xs:complexType name=""ServerTimeSubResponseDataType"">
                         <xs:attribute name=""ServerTime"" type=""xs:positiveInteger"" use=""required""/>
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R740
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     740,
                     @"[In ServerTimeSubResponseDataType] The ServerTime attribute MUST be specified in a server time subresponse that is generated in response to a server time subrequest.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R466
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     466,
                     @"[In SubResponseDataOptionalAttributes] The ServerTime attribute MUST be specified in a server time subresponse that is generated in response to a server time subrequest.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1337
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(ServerTimeSubResponseDataType),
                     serverSubResponseData.GetType(),
                     "MS-FSSHTTP",
                     1337,
                     @"[In ServerTime Subrequest] The ServerTimeSubResponseDataType defines the type of the SubResponseData element that is sent in a server time SubResponse element.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1395
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(ServerTimeSubResponseDataType),
                     serverSubResponseData.GetType(),
                     "MS-FSSHTTP",
                     1395,
                     @"[In SubResponseDataGenericType][SubResponseDataGenericType MUST take one of the forms described in the following table] ServerTimeSubResponseDataType: Type definition for server time subresponse data.");
        }
    }
}