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
    /// This partial part of MsfsshttpAdapterCapture class is used to test Get docMetaInfo response related requirements.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related to GetDocMetaInfo sub response.
        /// </summary>
        /// <param name="getDocMetaInfoSubResponse">Containing the GetDocMetaInfoSubResponse information.</param>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void ValidateGetDocMetaInfoSubResponse(GetDocMetaInfoSubResponseType getDocMetaInfoSubResponse, ITestSite site)
        {
            ValidateSubResponseType(getDocMetaInfoSubResponse as SubResponseType, site);

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1797
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1797,
                     @"[In GetDocMetaInfoSubResponseType][The schema of GetDocMetaInfoSubResponseType is] <xs:complexType name=""GetDocMetaInfoSubResponseType"">
                       <xs:complexContent>
                         <xs:extension base=""tns:SubResponseType"">
                           <xs:sequence minOccurs=""0"" maxOccurs=""1"">
                     <xs:element name=""SubResponseData"" type=""tns:GetDocMetaInfoSubResponseDataType""/>
                     </xs:sequence>
                         </xs:extension>
                       </xs:complexContent>
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2003
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2003,
                     @"[In GetDocMetaInfo Subrequest][The protocol client sends a GetDocMetaInfo SubRequest message, which is of type GetDocMetaInfoSubRequestType] The protocol server responds with a GetDocMetaInfo SubResponse message, which is of type GetDocMetaInfoSubResponseType as specified in section 2.3.1.30.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R4694
            // if can launch this method, the schema matches.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(GetDocMetaInfoSubResponseType),
                     getDocMetaInfoSubResponse.GetType(),
                     "MS-FSSHTTP",
                     4694,
                     @"[In SubResponseElementGenericType] Depending on the Type attribute specified in the SubRequest element, the SubResponseElementGenericType MUST take one of the forms: GetDocMetaInfoSubResponseType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R5748
            // if can launch this method, the schema matches.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(GetDocMetaInfoSubResponseType),
                     getDocMetaInfoSubResponse.GetType(),
                     "MS-FSSHTTP",
                     5748,
                     @"[In SubResponseType] The SubResponseElementGenericType takes one of the following forms: GetDocMetaInfoSubResponseType.");

            if (getDocMetaInfoSubResponse.SubResponseData != null)
            {
                ValidateGetDocMetaInfoSubResponseData(site);
            }

            if (string.Compare("Success", getDocMetaInfoSubResponse.ErrorCode, StringComparison.OrdinalIgnoreCase) == 0)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1693
                bool isVerifyR1693 = getDocMetaInfoSubResponse.SubResponseData != null;

                // If popup the assert, the case run fail. 
                site.Assert.IsTrue(
                            isVerifyR1693,
                            "For requirement MS-FSSHTTPB_R1693, the SubResponseData should not be NULL.");

                // If the above logic is right, MS-FSSHTTP_R1693 can be captured.
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         1693,
                         @"[In SubResponseElementGenericType][The SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"" and one of the following conditions is true:] The Type attribute that is specified in the SubRequest element is set to a value of ""GetDocMetaInfo"".");

                // If the value of element "SubResponseData" in the sub-response is not null, then capture MS-FSSHTTP_R1800.
                site.CaptureRequirementIfIsNotNull(
                         getDocMetaInfoSubResponse.SubResponseData,
                         "MS-FSSHTTP",
                         1800,
                         @"[In GetDocMetaInfoSubResponseType] [In GetDocMetaInfoSubResponseType] As part of processing the GetDocMetaInfo subrequest, the SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the following condition is true:
                         The ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"".");
            }

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2013
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2013,
                     @"[In GetDocMetaInfo Subrequest] The GetDocMetaInfoPropertySetType elements have one Property element per metadata item of type GetDocMetaInfoPropertyType as defined in section 2.3.1.29.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1685
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1685,
                     @"[In SubResponseDataGenericType][SubResponseDataGenericType MUST take one of the forms described in the following table] GetDocMetaInfoSubResponseDataType: Type definition for Get Doc Meta Info subresponse data.");
        }

        /// <summary>
        /// Capture requirements related with GetDocMetaInfoSubResponseType.
        /// </summary>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateGetDocMetaInfoSubResponseData(ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1785
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1785,
                     @"[In GetDocMetaInfoPropertySetType][The schema of GetDocMetaInfoPropertySetType is] <xs:complexType name=""GetDocMetaInfoPropertySetType"">
                         <xs:sequence minOccurs=""0"" maxOccurs=""unbounded"">
                           <xs:element name=""Property"" type=""tns:GetDocMetaInfoPropertyType""/>
                         </xs:sequence>
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1789
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1789,
                     @"[In GetDocMetaInfoPropertyType][The schema of GetDocMetaInfoPropertyType is] <xs:complexType name=""GetDocMetaInfoPropertyType"">
                         <xs:attribute name=""Key"" type=""xs:string"" use=""required""/>
                         <xs:attribute name=""Value"" type=""xs:string"" use=""required""/>
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2009
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2009,
                     @"[In GetDocMetaInfo Subrequest] GetDocMetaInfoSubResponseDataType defines the type of the SubResponseData element inside the GetDocMetaInfo SubResponse element.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1782
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1782,
                     @"[In GetDocMetaInfoSubResponseDataType][The schema of GetDocMetaInfoSubResponseDataType is] 
                     <xs:complexType name=""GetDocMetaInfoSubResponseDataType"">
                         <xs:sequence>
                           <xs:element name=""DocProps"" type=""tns:GetDocMetaInfoPropertySetType""/>
                           <xs:element name=""FolderProps"" type=""tns:GetDocMetaInfoPropertySetType""/>
                         </xs:sequence>
                     </xs:complexType>");
        }
    }
}