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
    /// This partial part of MsFsshttpAdapterCapture class is used to test Editor table response related requirements.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related to EditorsTable sub response.
        /// </summary>
        /// <param name="editorsTableSubResponse">Containing the EditorsTableSubResponse information.</param>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void ValidateEditorsTableSubResponse(EditorsTableSubResponseType editorsTableSubResponse, ITestSite site)
        {
            ValidateSubResponseType(editorsTableSubResponse as SubResponseType, site);

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R4693
            // if can launch this method, the schema matches.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(EditorsTableSubResponseType),
                     editorsTableSubResponse.GetType(),
                     "MS-FSSHTTP",
                     4693,
                     @"[In SubResponseElementGenericType] Depending on the Type attribute specified in the SubRequest element, the SubResponseElementGenericType MUST take one of the forms: EditorsTableSubResponseType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R5747
            // if can launch this method, the schema matches.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(EditorsTableSubResponseType),
                     editorsTableSubResponse.GetType(),
                     "MS-FSSHTTP",
                     5747,
                     @"[In SubResponseType] The SubResponseElementGenericType takes one of the following forms: EditorsTableSubResponseType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1769
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1769,
                     @"[In EditorsTableSubResponseType][The schema of EditorsTableSubResponseType is] <xs:complexType name=""EditorsTableSubResponseType"">
                       <xs:complexContent>
                         <xs:extension base=""tns:SubResponseType"">
                            <xs:sequence minOccurs=""0"" maxOccurs=""1"">
                               <xs:element name=""SubResponseData"">
                                 <xs:complexType>
                                   <xs:complexContent>
                                     <xs:restriction base=""xs:anyType""/>
                                   </xs:complexContent>
                                 </xs:complexType>
                               </xs:element>
                             </xs:sequence>
                         </xs:extension>
                       </xs:complexContent>
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3079
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     3079,
                     @"[In EditorsTableSubResponseType] SubResponseData: It MUST be an empty element without any attributes.");
        }
    }
}