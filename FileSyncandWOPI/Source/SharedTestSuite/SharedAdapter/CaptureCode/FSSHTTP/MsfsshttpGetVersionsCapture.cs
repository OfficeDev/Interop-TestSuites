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
    /// This partial part of MsfsshttpAdapterCapture class is used to test Get versions response related requirements.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related to GetVersions sub response.
        /// </summary>
        /// <param name="getVersionsSubResponse">Containing the getVersionsSubResponse information.</param>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void ValidateGetVersionsSubResponse(GetVersionsSubResponseType getVersionsSubResponse, ITestSite site)
        {
            ValidateSubResponseType(getVersionsSubResponse as SubResponseType, site);

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R4694
            // if can launch this method, the schema matches
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(GetVersionsSubResponseType),
                     getVersionsSubResponse.GetType(),
                     "MS-FSSHTTP",
                     4695,
                     @"[In SubResponseElementGenericType] Depending on the Type attribute specified in the SubRequest element, the SubResponseElementGenericType MUST take one of the forms: GetVersionsSubResponseType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1816
            // if can launch this method, the GetVersionsSubResponseType schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1816,
                     @"[In GetVersionsSubResponseType][The schema of GetVersionsSubResponseType is] <xs:complexType name=""GetVersionsSubResponseType"">
                       <xs:complexContent>
                         <xs:extension base=""tns:SubResponseType"">
                           <xs:sequence minOccurs=""0"" maxOccurs=""1"">
                              <xs:element ref=""tns:GetVersionsResponse""/>
                           </xs:sequence>
                         </xs:extension>
                       </xs:complexContent>
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2020
            // if can launch this method, the GetVersionsSubResponseType schema matches.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(GetVersionsSubResponseType),
                     getVersionsSubResponse.GetType(),
                     "MS-FSSHTTP",
                     2020,
                     @"[In GetVersions Subrequest][The protocol client sends a GetVersions SubRequest message, which is of type GetVersionsSubRequestType] The protocol server responds with a GetVersions SubResponse message, which is of type GetVersionsSubResponseType as specified in section 2.3.1.32.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2026
            // if can launch this method, the GetVersionsSubResponseType schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2026,
                     @"[In GetVersions Subrequest] The Results element, as specified in [MS-VERSS] section 2.2.4.1, is a complex type that specifies information about the file's versions.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2301
            // If isSchemaValid is true, the GetVersionsSubResponseType schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2301,
                     @"[In GetVersionsSubResponseType][In Results] The DeleteAllVersions, DeleteVersion, GetVersions, and RestoreVersion methods return the Results complex type.
                     <s:complexType name=""Results"">
                       <s:sequence>
                         <s:element name=""list"" maxOccurs=""1"" minOccurs=""1"">
                           <s:complexType>
                             <s:attribute name=""id"" type=""s:string"" use=""required"" />
                           </s:complexType>
                         </s:element>
                         <s:element name=""versioning"" maxOccurs=""1"" minOccurs=""1"">
                           <s:complexType>
                             <s:attribute name=""enabled"" type=""s:unsignedByte"" use=""required"" />
                           </s:complexType>
                         </s:element>
                         <s:element name=""settings"" maxOccurs=""1"" minOccurs=""1"">
                           <s:complexType>
                             <s:attribute name=""url"" type=""s:string"" use=""required"" />
                           </s:complexType>
                         </s:element>
                         <s:element name=""result"" maxOccurs=""unbounded"" minOccurs=""1"" type=""tns:VersionData""/>
                       </s:sequence>
                     </s:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2303
            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "For requirement MS-FSSHTTP_R2303, the versioning.enabled MUST be '0' or '1', the versioning.enabled value is : {0}", getVersionsSubResponse.GetVersionsResponse.GetVersionsResult.results.versioning.enabled.ToString());

            // if can launch this method and the versioning.enabled schema matches and value must be 0 or 1.
            bool isVerifyR2303 = getVersionsSubResponse.GetVersionsResponse.GetVersionsResult.results.versioning.enabled == 0 || getVersionsSubResponse.GetVersionsResponse.GetVersionsResult.results.versioning.enabled == 1;
            site.CaptureRequirementIfIsTrue(
                     isVerifyR2303,
                     "MS-FSSHTTP",
                     2303,
                     @"[In GetVersionsSubResponseType][Results complex type] versioning.enabled: The value of this attribute [versioning.enabled] MUST be ""0"" or ""1"".");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2308
            // if can launch this method, the versioning.enabled schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2308,
                     @"[In GetVersionsSubResponseType][In VersionData] The VersionData complex type specifies the details about a single version of a file.
                     <s:complexType name=""VersionData"">
                       <s:attribute name=""version"" type=""s:string"" use=""required"" />
                       <s:attribute name=""url"" type=""s:string"" use=""required"" />
                       <s:attribute name=""created"" type=""s:string"" use=""required"" />
                       <s:attribute name=""createdRaw"" type=""s:string"" use=""required"" />  
                       <s:attribute name=""createdBy"" type=""s:string"" use=""required"" />
                       <s:attribute name=""createdByName"" type=""s:string"" use=""optional"" />
                       <s:attribute name=""size"" type=""s:unsignedLong"" use=""required"" />
                       <s:attribute name=""comments"" type=""s:string"" use=""required"" />
                     </s:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3083
            // if can launch this method, the versioning.enabled schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     3083,
                     @"[In GetVersionsResponse][The schema of GetVersionsResponse element is defined as:] 
                     <s:element name=""GetVersionsResponse"">
                       <s:complexType>
                         <s:sequence>
                           <s:element minOccurs=""1"" maxOccurs=""1"" name=""GetVersionsResult"">
                             <s:complexType>
                               <s:sequence>
                                 <s:element name=""results"" minOccurs=""1"" maxOccurs=""1"" type=""tns:Results"" />
                               </s:sequence>
                             </s:complexType>
                           </s:element>
                         </s:sequence>
                       </s:complexType>
                     </s:element>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3084
            // if can launch this method, the versioning.enabled schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     3084,
                     @"[In GetVersionsResponse] GetVersionsResult: An XML node that conforms to the structure specified in section 2.2.4.1. ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2306
            // if can launch this method, the versioning.enabled schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2306,
                     @"[In GetVersionsSubResponseType][Results complex type] settings.url: Specifies the URL to the webpage of versioning-related settings for the document library in which the file resides. ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R60101
            // if can launch this method, the versioning.enabled schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     60101,
                     @"[In GetVersionsSubResponseType][VersionData] Implementation does contain the version of the file, including the major version and minor version numbers connected by period, for example, ""1.0"". (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 and above follow this behavior.)");
        }
    }
}