//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_DWSS
{
    using System;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter requirements capture code for MS-DWSS server role. 
    /// </summary>
    public partial class MS_DWSSAdapter
    {
        /// <summary>
        /// Validate the requirements related to schema for CanCreateDwsUrl operation.
        /// </summary>
        /// <param name="respXmlString">CanCreateDwsUrlResult decoded xml string returned by server.</param>
        private void ValidateCanCreateDwsUrlResponseSchema(string respXmlString)
        {
            // Validate CanCreateDwsUrlResult element schema.
            AdapterHelper.ValidateRespSchema(respXmlString);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R108");

            // Verify MS-DWSS requirement: MS-DWSS_R108. If there is any validation error or warning occurred in previous step, an exception will be thrown.
            Site.CaptureRequirement(
                108,
                @"[In CanCreateDwsUrlResponse] The string[CanCreateDwsUrlResult] MUST be a standalone XML element specified in the following format:
                <s:complexType>
                  <s:choice>
                    <s:element name=""Error"" type=""Error""/>
                    <s:element name=""Result"" type=""s:string""/>
                  </s:choice>
                </s:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R443");

            // Verify MS-DWSS requirement: MS-DWSS_R443
            // The schema of CanCreateDwsUrl operation has been validated by full WSDL. If it returns success, the schema of CanCreateDwsUrl operation is valid, capture related requirements.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                443,
                @"[In CanCreateDwsUrlResponse] This structure is defined as follows:
                <s:element name=""CanCreateDwsUrlResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""CanCreateDwsUrlResult"" type=""s:string"" minOccurs=""0"" maxOccurs=""1""/>
                    </s:sequence>
                  </s:complexType>
                </s:element>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R100");

            // Verify MS-DWSS requirement: MS-DWSS_R100
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                100,
                @"[In CanCreateDwsUrl] The protocol client sends a CanCreateDwsUrlSoapIn request message, and the protocol server responds with the following CanCreateDwsUrlSoapOut response message:
                <wsdl:operation name=""CanCreateDwsUrl"">
                  <wsdl:input message=""tns:CanCreateDwsUrlSoapIn"" />
                  <wsdl:output message=""tns:CanCreateDwsUrlSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R438");

            // Verify MS-DWSS requirement: MS-DWSS_R438
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                438,
                @"[In CanCreateDwsUrlSoapOut] The SOAP body contains a CanCreateDwsUrlResponse element.");
        }

        /// <summary>
        /// Validate the requirements related to schema for operation CreateDws.
        /// </summary>
        /// <param name="respXmlString">CreateDwsResult decoded xml string returned by server.</param>
        private void ValidateCreateDwsResponseSchema(string respXmlString)
        {
            // Validate response xml schema.
            AdapterHelper.ValidateRespSchema(respXmlString);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R138");

            // Verify MS-DWSS requirement: MS-DWSS_R138. If there is any validation error or warning occurred in previous step, an exception will be thrown.
            Site.CaptureRequirement(
                138,
                @"[In CreateDwsResponse] The string [CreateDwsResult] MUST conform to the following XSD:
                <s:complexType>
                  <s:choice>
                    <s:element name=""Error"" type=""Error""/>
                    <s:element name=""Results"">
                      <s:complexType>
                        <s:sequence>
                          <s:element name=""Url"" type=""s:string"" />
                          <s:element name=""DoclibUrl"" type=""s:string"" />
                          <s:element name=""ParentWeb"" type=""s:string""/>
                          <s:element name=""FailedUsers"" type=""tns:UserType"" minOccurs=""0"" maxOccurs=""unbounded""/>
                          <s:element name=""AddUsersUrl"" type=""s:string""/>
                          <s:element name=""AddUsersRole"" type=""s:string""/>
                        </s:sequence>
                      </s:complexType>
                    </s:element>
                  </s:choice>
                </s:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R118");

            // Verify MS-DWSS requirement: MS-DWSS_R118
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                118,
                @"[In CreateDws] The protocol client sends a CreateDwsSoapIn request message, and the protocol server responds with a CreateDwsSoapOut response message, as follows:
                <wsdl:operation name=""CreateDws"">
                  <wsdl:input message=""tns:CreateDwsSoapIn"" />
                  <wsdl:output message=""tns:CreateDwsSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R472");

            // Verify MS-DWSS requirement: MS-DWSS_R472
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                472,
                @"[In CreateDwsResponse] This element is defined as follows:
                    <s:element name=""CreateDwsResponse"">
                      <s:complexType>
                        <s:sequence>
                          <s:element name=""CreateDwsResult"" type=""s:string"" minOccurs=""0"" maxOccurs=""1"" />
                        </s:sequence>
                      </s:complexType>
                    </s:element>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R125, the response message is {0}.", respXmlString);

            // Verify MS-DWSS requirement: MS-DWSS_R125
            Site.CaptureRequirementIfAreNotEqual<string>(
                string.Empty,
                respXmlString,
                125,
                @"[In CreateDws] The CreateDwsResponse response message MUST NOT be empty.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R454");

            // Verify MS-DWSS requirement: MS-DWSS_R454
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                454,
                @"[In CreateDwsSoapOut] The SOAP body contains a CreateDwsResponse element.");
        }

        /// <summary>
        /// Validate the requirements related to Results element in CreateDwsResult element.
        /// </summary>
        /// <param name="usersString">'users' parameter pass in request.</param>
        /// <param name="respResults">Results element in CreateDwsResult element.</param>
        private void ValidateCreateDwsResultResults(string usersString, CreateDwsResultResults respResults)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R139, the Url is {0}.", respResults.Url);

            // Only if FailedUsers element presents, then capture the requirement related to schema validation.
            if (respResults.FailedUsers != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R140");

                // Verify MS-DWSS requirement: MS-DWSS_R140. If there is any validation error or warning occurred in previous step, an exception will be thrown.
                Site.CaptureRequirement(
                    140,
                    @"[In CreateDwsResponse] It[FailedUsers] is standalone XML that MUST use the following definition:
                <s:complexType name=""UserType"">
                  <s:sequence>
                    <s:element name=""User"" minOccurs=""0"">
                      <s:complexType>
                      <s:attribute name=""Email"" type=""s:string""/>
                      </s:complexType>
                    </s:element>
                  </s:sequence>
                </s:complexType>");
            }

            Uri workspaceUri = new Uri(respResults.Url);

            // Verify MS-DWSS requirement: MS-DWSS_R139
            Site.CaptureRequirementIfIsTrue(
                workspaceUri.IsAbsoluteUri,
                139,
                @"[In CreateDwsResponse] This [Url] MUST be an absolute URL.");

            if (!string.IsNullOrEmpty(usersString))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R481");

                // Verify MS-DWSS requirement: MS-DWSS_R481
                Site.CaptureRequirementIfAreEqual<string>(
                    "Microsoft.SharePoint.SPRoleDefinition",
                    respResults.AddUsersRole,
                    481,
                    @"[In CreateDwsResponse] If the users parameter to CreateDws (section 3.1.4.2.2.1) was not empty, the value [of AddUsersRole] MUST be the string ""Microsoft.SharePoint.SPRoleDefinition"".");
            }
        }

        /// <summary>
        /// Validate the requirements related to schema for CreateFolder operation.
        /// </summary>
        /// <param name="respXmlString">CreateFolderResult decoded xml string returned by server.</param>
        private void ValidateCreateFolderResponseSchema(string respXmlString)
        {
            // Validate response xml schema.
            AdapterHelper.ValidateRespSchema(respXmlString);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R156");

            // Verify MS-DWSS requirement: MS-DWSS_R156. If there is any validation error or warning occurred in previous step, an exception will be thrown.
            Site.CaptureRequirement(
                156,
                @"[In CreateFolderResult] The element[CreateFolderResult] MUST contain either an Error element or a Result element as follows:
                <s:complexType>
                  <s:choice>
                    <s:element name=""Result""/>
                    <s:element ref=""tns:Error""/>
                  </s:choice>
                </s:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R497");

            // Verify MS-DWSS requirement: MS-DWSS_R497
            // The schema of CreateFolder operation has been validated by full WSDL. If it returns success, the schema of CreateFolder operation is valid, capture related requirements.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                497,
                @"[In CreateFolderResponse] This element is defined as follows:
                <s:element name=""CreateFolderResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""CreateFolderResult"" type=""s:string"" minOccurs=""0"" maxOccurs=""1""/>
                    </s:sequence>
                  </s:complexType>
                </s:element>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R148");

            // Verify MS-DWSS requirement: MS-DWSS_R148
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                148,
                @"[In CreateFolder] The protocol client sends a CreateFolderSoapIn request message, and the protocol server responds with a CreateFolderSoapOut response message, as follows:
                <wsdl:operation name=""CreateFolder"">
                    <wsdl:input message=""tns:CreateFolderSoapIn"" />
                    <wsdl:output message=""tns:CreateFolderSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R492");

            // Verify MS-DWSS requirement: MS-DWSS_R492
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                492,
                @"[In CreateFolderSoapOut] The SOAP body contains a CreateFolderResponse element.");
        }

        /// <summary>
        /// Validate the requirements related to schema for DeleteDws operation.
        /// </summary>
        /// <param name="respXmlString">DeleteDwsResult decoded xml string returned by server.</param>
        private void ValidateDeleteDwsResponseSchema(string respXmlString)
        {
            // Validate response xml schema.
            AdapterHelper.ValidateRespSchema(respXmlString);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R167");

            // Verify MS-DWSS requirement: MS-DWSS_R167. If there is any validation error or warning occurred in previous step, an exception will be thrown.
            Site.CaptureRequirement(
                167,
                @"[In DeleteDwsResponse] The string[DeleteDwsResult] MUST conform to either the Error element or a Result element as follows:
                <s:complexType>
                  <s:choice>
                    <s:element ref=""tns:Error""/>
                    <s:element name=""Result""/>
                  </s:choice>
                </s:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R161");

            // Verify MS-DWSS requirement: MS-DWSS_R161
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                161,
                @"[In DeleteDws] The protocol client sends a DeleteDwsSoapIn request message, and the protocol server responds with a DeleteDwsSoapOut response message, as follows:
                <wsdl:operation name=""DeleteDws"">
                   <wsdl:input message=""tns:DeleteDwsSoapIn"" />
                   <wsdl:output message=""tns:DeleteDwsSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R509");

            // Verify MS-DWSS requirement: MS-DWSS_R509
            // The schema of RenameDws operation has been validated by full WSDL. If it returns success, the schema of RenameDws operation is valid, capture related requirements.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                509,
                @"[In DeleteDwsResponse] [Element DeleteDwsResponse is defined as follows:]
                <s:element name=""DeleteDwsResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""DeleteDwsResult"" type=""s:string"" minOccurs=""0"" maxOccurs=""1""/>
                    </s:sequence>
                  </s:complexType>
                </s:element>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R505");

            // Verify MS-DWSS requirement: MS-DWSS_R505
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                505,
                @"[In DeleteDwsSoapOut] The SOAP body contains a DeleteDwsResponse element.");
        }

        /// <summary>
        /// Validate the requirements related to Result element in DeleteDws operation.
        /// </summary>
        private void ValidateDeleteDwsResultResult()
        {
            // Verify MS-DWSS requirement: MS-DWSS_R1672
            Site.CaptureRequirement(
                1672,
                "[In DeleteDwsResponse] Result: An empty Result element (\"<Result/>\") if the call is successful.");
        }

        /// <summary>
        /// Validate the requirements related to Result element in DeleteDws operation.
        /// </summary>
        private void ValidateCreateFolderResultResult()
        {
            // Verify MS-DWSS requirement: MS-DWSS_R1159
            Site.CaptureRequirement(
                1159,
                "[In CreateFolderResponse] Result: An empty Result element (\"<Result/>\") if the call is successful.");
        }

        /// <summary>
        /// Validate the requirements related to schema for the DeleteFolder operation.
        /// </summary>
        /// <param name="respXmlString">DeleteFolderResult decoded xml string returned by server.</param>
        private void ValidateDeleteFolderResponseSchema(string respXmlString)
        {
            // Validate response xml schema.
            AdapterHelper.ValidateRespSchema(respXmlString);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R178");

            // Verify MS-DWSS requirement: MS-DWSS_R178. If there is any validation error or warning occurred in previous step, an exception will be thrown.
            Site.CaptureRequirement(
                178,
                @"[In DeleteFolderResponse] The XML MUST conform to either the Error element or a Result element as follows:
                <s:complexType>
                  <s:choice>
                    <s:element ref=""tns:Error""/>
                    <s:element name=""Result""/>
                  </s:choice>
                </s:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R171");

            // Verify MS-DWSS requirement: MS-DWSS_R171
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                171,
                @"[In DeleteFolder] The protocol client sends a DeleteFolderSoapIn request message, and the protocol server responds with a DeleteFolderSoapOut response message, as follows:
                <wsdl:operation name=""DeleteFolder"">
                    <wsdl:input message=""tns:DeleteFolderSoapIn"" />
                    <wsdl:output message=""tns:DeleteFolderSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R522");

            // Verify MS-DWSS requirement: MS-DWSS_R522
            // The schema of DeleteFolder operation has been validated by full WSDL. If it returns success, the schema of DeleteFolder operation is valid, capture related requirements.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                522,
                @"[In DeleteFolderResponse] This element is defined as follows:
                <s:element name=""DeleteFolderResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""DeleteFolderResult"" type=""s:string"" minOccurs=""0""/>
                    </s:sequence>
                  </s:complexType>
                </s:element>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R517");

            // Verify MS-DWSS requirement: MS-DWSS_R517
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                517,
                @"[In DeleteFolderSoapOut] The SOAP body contains a DeleteFolderResponse element.");
        }

        /// <summary>
        /// Validate the requirements related to result element in the DeleteFolder operation.
        /// </summary>
        private void ValidateDeleteFolderResultResult()
        {
            // Verify MS-DWSS requirement: MS-DWSS_R1524
            Site.CaptureRequirement(
                1524,
                "[In DeleteFolderResponse] Result: An empty Result element (\"<Result/>\") if the call is successful.");
        }

        /// <summary>
        /// Validate the requirements related to schema for the FindDwsDoc operation.
        /// </summary>
        /// <param name="respXmlString">FindDwsDocResult decoded xml string returned by server.</param>
        private void ValidateFindDwsDocResponseSchema(string respXmlString)
        {
            // Validate response xml schema.
            AdapterHelper.ValidateRespSchema(respXmlString);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R188");

            // Verify MS-DWSS requirement: MS-DWSS_R188. If there is any validation error or warning occurred in previous step, an exception will be thrown.
            Site.CaptureRequirement(
                188,
                @"[In FindDwsDocResponse] FindDwsDocResult: This element contains a string that is stand–alone XML encoded either as an Error element as specified in 2.2.3.2 or a Result element defined as follows:
                <s:complexType>
                  <s:choice>
                    <s:element ref=""tns:Error""/>
                    <s:element name=""Result"" type=""s:string""/>
                  </s:choice>
                </s:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R182");

            // Verify MS-DWSS requirement: MS-DWSS_R182
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                182,
                @"[In FindDwsDoc] The protocol client sends a FindDwsDocSoapIn request message, and the protocol server responds with a FindDwsDocSoapOut response message, as follows:
                <wsdl:operation name=""FindDwsDoc"">
                    <wsdl:input message=""tns:FindDwsDocSoapIn"" />
                    <wsdl:output message=""tns:FindDwsDocSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R534");

            // Verify MS-DWSS requirement: MS-DWSS_R534
            // The schema of FindDwsDoc operation has been validated by full WSDL. If it returns success, the schema of FindDwsDoc operation is valid, capture related requirements.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                534,
                @"[In FindDwsDocResponse] This element[FindDwsDocResponse] is defined as follows:
                <s:element name=""FindDwsDocResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""FindDwsDocResult"" type=""s:string"" minOccurs=""0""/>
                    </s:sequence>
                  </s:complexType>
                </s:element>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R529");

            // Verify MS-DWSS requirement: MS-DWSS_R529
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                529,
                @"[In FindDwsDocSoapOut] The SOAP body contains a FindDwsDocResponse element.");
        }

        /// <summary>
        /// Validate the requirements related to schema for GetDwsData operation.
        /// </summary>
        /// <param name="respXmlString">GetDwsDataResult decoded xml string returned by server.</param>
        private void ValidateGetDwsDataResponseSchema(string respXmlString)
        {
            // Validate response xml schema.
            AdapterHelper.ValidateRespSchema(respXmlString);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R207");

            // Verify MS-DWSS requirement: MS-DWSS_R207. If there is any validation error or warning occurred in previous step, an exception will be thrown.
            Site.CaptureRequirement(
                207,
                @"[In GetDwsDataResponse] The XML MUST conform to the following schema:
                <s:complexType name=""GetDwsDataResultType"">
                  <s:choice>
                    <s:element ref=""tns:Error""/>
                    <s:element name=""Results"">
                      <xs:sequence>
                        <xs:element name=""Title"" type=""xs:string""/>
                        <xs:element name=""LastUpdate"" type=""xs:integer""/>
                        <xs:element name=""User"">
                          <xs:complexType>
                            <xs:sequence>
                              <xs:element name=""ID"" type=""xs:string""/>
                              <xs:element name=""Name"" type=""xs:string""/>
                              <xs:element name=""LoginName"" type=""xs:string""/>
                              <xs:element name=""Email"" type=""xs:string""/>
                              <xs:element name=""IsDomainGroup""/>
                                <s:simpleType>
                                  <s:restriction base=""s:string"">
                                    <s:enumeration value=""True"" />
                                    <s:enumeration value=""False"" />
                                  </s:restriction>
                                </s:simpleType>
                              <xs:element name=""IsSiteAdmin""/>
                                <s:simpleType>
                                  <s:restriction base=""s:string"">
                                    <s:enumeration value=""True"" />
                                    <s:enumeration value=""False"" />
                                  </s:restriction>
                                </s:simpleType>
                            </xs:sequence>
                          </xs:complexType>
                        </xs:element>
                        <xs:element name=""Members"" type=""tns:MemberData""/>
                        <xs:sequence minOccurs=""0"">
                          <xs:element ref=""tns:Assignees""/>
                          <xs:element ref=""tns:List""/>
                          <xs:element ref=""tns:List""/>
                          <xs:element ref=""tns:List""/>
                        </xs:sequence>
                      </xs:sequence>
                    </s:element>
                  </s:choice>
                </s:complexType>");

            // The schema of GetDwsDataResult operation has been validated by full WSDL. If it returns success, the schema of GetDwsDataResult operation is valid, capture related requirements.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R194");

            // Verify MS-DWSS requirement: MS-DWSS_R194
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                194,
                @"[In GetDwsData] The protocol client sends a GetDwsDataSoapIn request message, and the protocol server responds with a GetDwsDataSoapOut response message, as follows:
                <wsdl:operation name=""GetDwsData"">
                    <wsdl:input message=""tns:GetDwsDataSoapIn"" />
                    <wsdl:output message=""tns:GetDwsDataSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R552");

            // Verify MS-DWSS requirement: MS-DWSS_R552
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                552,
                @"[In GetDwsDataResponse] This element is defined as follows:
                <s:element name=""GetDwsDataResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""GetDwsDataResult"" type=""s:string"" minOccurs=""0""/>
                    </s:sequence>
                  </s:complexType>
                </s:element>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R542");

            // Verify MS-DWSS requirement: MS-DWSS_R542
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                542,
                @"[In GetDwsDataSoapOut] The SOAP body contains a GetDwsDataResponse element.");
        }

        /// <summary>
        /// Validate the requirements related to schema for GetDwsData results element.
        /// </summary>
        /// <param name="respResults">Results element in GetDwsDataResult element.</param>
        private void ValidateGetDwsDataResultResults(Results respResults)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R343");

            // Verify MS-DWSS requirement: MS-DWSS_R343. If GetDwsDataResultResults element returned and there is no schema validation error or warning, capture the following requirement.
            Site.CaptureRequirement(
                343,
                @"[In Assignees] This element[Assignees] is defined as follows:
                  <xs:element name=""Assignees"">
                    <xs:complexType>
                      <xs:sequence>
                        <xs:element ref=""tns:Member"" maxOccurs=""unbounded""/>
                      </xs:sequence>
                    </xs:complexType>
                  </xs:element>");

            // If MS-DWSS_R343 is verified correctly, it means that Assignees which type is Assignees is also verified.
            Site.CaptureRequirement(
                216,
                @"[In GetDwsDataResponse] This element [Assignees] MUST conform to the Assignees element schema specified in section 2.2.3.1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R350");

            // Verify MS-DWSS requirement: MS-DWSS_R350
            Site.CaptureRequirement(
                350,
                @"[In ID] This element is defined as follows:
                  <xs:element name=""ID"">
                    <xs:complexType>
                      <xs:simpleContent>
                        <xs:extension base=""xs:string"">
                          <xs:attribute name=""DefaultUrl"" type=""xs:string""/>
                        </xs:extension>
                      </xs:simpleContent>
                    </xs:complexType>
                  </xs:element>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R366");

            // Verify MS-DWSS requirement: MS-DWSS_R366
            Site.CaptureRequirement(
                366,
                @"[In Member] This element[Member] is defined as follows:
                    <s:element name=""Member"">
                      <s:complexType>
                        <s:all>
                          <s:element name=""ID"" type=""s:integer""/>
                          <s:element name=""Name"" type=""s:string""/>
                          <s:element name=""LoginName"" type=""s:string""/>
                          <s:element name=""Email"" type=""s:string"" minOccurs=""0""/>
                          <s:element name=""IsDomainGroup"" minOccurs=""0"">
                            <s:simpleType>
                              <s:restriction base=""s:string"">
                                <s:enumeration value=""True"" />
                                <s:enumeration value=""False"" />
                                </s:restriction>
                            </s:simpleType>
                          </s:element>
                        </s:all>
                      </s:complexType>
                     </s:element>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R385");

            // Verify MS-DWSS requirement: MS-DWSS_R385
            Site.CaptureRequirement(
                385,
                @"[In MemberData] This type[MemberData] is defined as follows:
                <xs:complexType name=""MemberData"">
                  <xs:choice>
                    <xs:element ref=""tns:Error""/>
                    <xs:sequence>
                      <xs:element name=""DefaultUrl"" type=""xs:string""/>
                      <xs:element name=""AlternateUrl"" type=""xs:string""/>
                      <xs:element ref=""tns:Error""/>
                    </xs:sequence>
                    <xs:sequence>
                      <xs:element ref=""tns:Member"" minOccurs=""0"" maxOccurs=""unbounded""/>
                    </xs:sequence>
                  </xs:choice>
                </xs:complexType>");

            // If MS-DWSS_R385 is verified correctly, it means that Members which type is MemberData is also verified.
            Site.CaptureRequirement(
                213,
                @"[In GetDwsDataResponse] Members: This element MUST conform to the schema for the complex data type MemberData.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R394");

            // Verify MS-DWSS requirement: MS-DWSS_R394
            Site.CaptureRequirement(
                394,
                @"[In ListType] This type is defined as follows:
                <xs:simpleType name=""ListType"">
                  <xs:restriction base=""xs:string"">
                    <xs:enumeration value=""Tasks""/>
                    <xs:enumeration value=""Documents""/>
                    <xs:enumeration value=""Links""/>
                  </xs:restriction>
                </xs:simpleType>");

            this.ValidateListElement(respResults.List);
            this.ValidateListElement(respResults.List1);
            this.ValidateListElement(respResults.List2);
        }

        /// <summary>
        /// Validate the requirements related to schema for GetDwsData error element.
        /// </summary>
        private void ValidateGetDwsDataResultError()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R344");

            // Verify MS-DWSS requirement: MS-DWSS_R344
            Site.CaptureRequirement(
                344,
                @"[In Error] This element[Error] is defined as follows:
                <xs:simpleType name=""ErrorTypes"">
                  <xs:restriction base=""xs:string"">
                    <xs:enumeration value=""ServerFailure""/>
                    <xs:enumeration value=""Failed""/>
                    <xs:enumeration value=""NoAccess""/>
                    <xs:enumeration value=""Conflict""/>
                    <xs:enumeration value=""ItemNotFound""/>
                    <xs:enumeration value=""MemberNotFound""/>
                    <xs:enumeration value=""ListNotFound""/>
                    <xs:enumeration value=""TooManyItems""/>
                    <xs:enumeration value=""DocumentNotFound""/>
                    <xs:enumeration value=""FolderNotFound""/>
                    <xs:enumeration value=""WebContainsSubwebs""/>
                    <xs:enumeration value=""ADMode""/>
                    <xs:enumeration value=""AlreadyExists""/>
                    <xs:enumeration value=""QuotaExceeded""/>
                  </xs:restriction>
                </xs:simpleType>
                <xs:element name=""Error"">
                  <xs:complexType>
                    <xs:simpleContent>
                      <xs:extension base=""tns:ErrorTypes"">
                        <xs:attribute name=""ID"">
                          <xs:simpleType>
                            <xs:restriction base=""xs:integer"">
                              <xs:minInclusive value=""1""/>
                              <xs:maxInclusive value=""14""/>
                            </xs:restriction>
                          </xs:simpleType>
                        </xs:attribute>
                        <xs:attribute name=""AccessUrl"" type=""xs:string""/>
                      </xs:extension>
                    </xs:simpleContent>
                  </xs:complexType>
                </xs:element>");
        }

        /// <summary>
        /// Validate the requirements related to schema for GetDwsMetaData operation.
        /// </summary>
        /// <param name="respXmlString">GetDwsMetaDataResult decoded xml string returned by server.</param>
        private void ValidateGetDwsMetaDataResponseSchema(string respXmlString)
        {
            // Validate response xml schema.
            AdapterHelper.ValidateRespSchema(respXmlString);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R238");

            // Verify MS-DWSS requirement: MS-DWSS_R238. If there is any validation error or warning occurred in previous step, an exception will be thrown.
            Site.CaptureRequirement(
                238,
                @"[In GetDwsMetaDataResponse] GetDwsMetaDataResult: This element contains a string that is stand-alone XML, encoded either as an Error element as specified in section 2.2.3.2 or a Result element defined as follows:
                   <s:complexType>
                      <s:choice>
                         <s:element ref=""tns:Error""/>
                         <s:element name=""Results"" type=""s:string""/>
                      </s:choice>
                   </s:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R595");

            // Verify MS-DWSS requirement: MS-DWSS_R595
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                595,
                @"[In GetDwsMetaDataResponse] The element[GetDwsMetaDataResponse] contains a string that is stand-alone XML, defined as follows:
                  <s:element name=""GetDwsMetaDataResponse"">
                    <s:complexType>
                      <s:sequence>
                        <s:element minOccurs=""0"" maxOccurs=""1""
                        name=""GetDwsMetaDataResult"" type=""s:string"" />
                      </s:sequence>
                    </s:complexType>
                  </s:element>");

            // The schema of GetDwsMetaData operation has been validated by full WSDL. If it returns success, the schema of GetDwsMetaData operation is valid, capture related requirements.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R230");

            // Verify MS-DWSS requirement: MS-DWSS_R230
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                230,
                @"[In GetDwsMetaData] The protocol client sends a GetDwsMetaDataSoapIn request message, and the protocol server responds with the following GetDwsMetaDataSoapOut response message:
                <wsdl:operation name=""GetDwsMetaData"">
                    <wsdl:input message=""tns:GetDwsMetaDataSoapIn"" />
                    <wsdl:output message=""tns:GetDwsMetaDataSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R572");

            // Verify MS-DWSS requirement: MS-DWSS_R572
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                572,
                @"[In GetDwsMetaDataSoapOut] The SOAP body contains a GetDwsMetaDataResponse element.");
        }

        /// <summary>
        /// Validate the requirements related to schema for GetDwsMetaData results element.
        /// </summary>
        /// <param name="respResults">Results element in GetDwsMetaDataResult element.</param>
        private void ValidateGetDwsMetaDataResultResults(GetDwsMetaDataResultTypeResults respResults)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R373");

            // Verify MS-DWSS requirement: MS-DWSS_R373
            Site.CaptureRequirement(
                373,
                @"[In Roles] This element[Roles] is defined as follows:
                <xs:element name=""Roles"">
                    <xs:complexType>
                      <xs:sequence>
                        <xs:choice>
                          <xs:element ref=""tns:Error""/>
                          <xs:sequence>
                            <xs:element name=""Role"" maxOccurs=""unbounded"">
                              <xs:complexType>
                                <xs:attribute name=""Name"" type=""xs:string""
                                              use=""required""/>
                                <xs:attribute name=""Type"" type=""tns:RoleType""
                                              use=""required""/>
                                <xs:attribute name=""Description"" type=""xs:string""
                                              use=""required""/>
                              </xs:complexType>
                            </xs:element>
                          </xs:sequence>
                        </xs:choice>
                      </xs:sequence>
                    </xs:complexType>
                  </xs:element>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R395");

            // Verify MS-DWSS requirement: MS-DWSS_R395
            Site.CaptureRequirement(
                395,
                @"[In RoleType] This type is defined as follows:
                  <xs:simpleType name=""RoleType"">
                    <xs:restriction base=""xs:string"">
                      <xs:enumeration value=""None""/>
                      <xs:enumeration value=""Reader""/>
                      <xs:enumeration value=""Contributor""/>
                      <xs:enumeration value=""WebDesigner""/>
                      <xs:enumeration value=""Administrator""/>
                    </xs:restriction>
                  </xs:simpleType>");

            // If MS-DWSS_R373 is verified correctly, it means this requirement can be verified.
            Site.CaptureRequirement(
                252,
                @"[In GetDwsMetaDataResponse] This element [Roles] MUST conform to the element specification in 2.2.3.6.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R246");

            // Verify MS-DWSS requirement: MS-DWSS_R246
            Site.CaptureRequirement(
                246,
                @"[In GetDwsMetaDataResponse] It[Results] MUST conform to the following schema:
                <xs:element name=""Results"">
                  <xs:complexType>
                    <xs:sequence>
                       <xs:element name=""SubscribeUrl"" type=""xs:string"" minOccurs=""0""/>
                       <xs:element name=""MtgInstance"" type=""xs:string""/>
                       <xs:element name=""SettingUrl"" type=""xs:string""/>
                   <xs:element name=""PermsUrl"" type=""xs:string""/>
                       <xs:element name=""UserInfoUrl"" type=""xs:string""/>
                       <xs:element ref=""tns:Roles""/>
                   <xs:element ref=""Schema"" type=""xs:string""/>
                       <xs:element ref=""Schema"" type=""xs:string""/>
                       <xs:element ref=""Schema"" type=""xs:string""/>
                       <xs:element ref=""tns:ListInfo""/>
                       <xs:element ref=""tns:ListInfo""/>
                       <xs:element ref=""tns:ListInfo""/>
                   <xs:element name=""Permissions"">
                        <xs:complexType>
                           <xs:choice>
                            <xs:element ref=""tns:Error""/>
                              <xs:sequence>
                                <xs:element name=""ManageSubwebs"" minOccurs=""0""/>
                                <xs:element name=""ManageWeb"" minOccurs=""0""/>
                                <xs:element name=""ManageRoles"" minOccurs=""0""/>
                                <xs:element name=""ManageLists"" minOccurs=""0""/>
                                <xs:element name=""InsertListItems"" minOccurs=""0""/>
                                <xs:element name=""EditListItems"" minOccurs=""0""/>
                                <xs:element name=""DeleteListItems"" minOccurs=""0""/>
                              </xs:sequence>
                            </xs:choice>
                          </xs:complexType>
                      </xs:element>
                      <xs:element name=""HasUniquePerm""/>
                      <xs:element name=""WorkspaceType""/>
                      <xs:element name=""IsADMode""/>
                      <xs:element name=""DocUrl""/>
                      <xs:element name=""Minimal""/>
                    <s:element name=""GetDwsDataResult"" type=""tns:GetDwsDataResultType""/>
                    </xs:sequence>
                  </xs:complexType>
                </xs:element>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R620");

            // Verify MS-DWSS requirement: MS-DWSS_R620
            Site.CaptureRequirement(
                620,
                @"[In ListInfo] This element[ListInfo] is defined as follows:
                <xs:element name=""ListInfo"">
                  <xs:complexType>
                    <xs:choice>
                      <xs:element ref=""tns:Error"" minOccurs=""0""/>
                      <xs:sequence>
                        <xs:element name=""Moderated"" type=""xs:boolean""/>
                        <xs:element name=""ListPermissions"">
                          <xs:complexType>
                              <xs:sequence>
                                <xs:element name=""InsertListItems"" minOccurs=""0""/>
                                <xs:element name=""EditListItems"" minOccurs=""0""/>
                                <xs:element name=""DeleteListItems"" minOccurs=""0""/>
                                <xs:element name=""ManageLists"" minOccurs=""0""/>
                                <xs:element ref=""tns:Error"" minOccurs=""0""/>
                              </xs:sequence>
                          </xs:complexType>
                        </xs:element>
                      </xs:sequence>
                    </xs:choice>
                    <xs:attribute name=""Name"" type=""xs:string"" use=""required""/>
                  </xs:complexType>
                </xs:element>");

            // If MS-DWSS_R620 is verified correctly, this requirements also can be verified directly.
            Site.CaptureRequirement(
                260,
                @"[In GetDwsMetaDataResponse] This element [ListInfo] MUST conform to the ListInfo element specified in 3.1.4.8.2.3.");

            // If MS-DWSS_R620 is verified correctly, this requirements also can be verified directly.
            Site.CaptureRequirement(
                262,
                @"[In GetDwsMetaDataResponse] This element [ListInfo] MUST conform to the ListInfo element specified in 3.1.4.8.2.3.");

            // If MS-DWSS_R620 is verified correctly, this requirements also can be verified directly.
            Site.CaptureRequirement(
                264,
                @"[In GetDwsMetaDataResponse] This element [ListInfo] MUST conform to the ListInfo element specified in 3.1.4.8.2.3.");

            Results results = respResults.Results;

            if (results != null)
            {
                if (results.List != null)
                {
                    this.ValidateListElement(results.List);
                }

                if (results.List1 != null)
                {
                    this.ValidateListElement(results.List1);
                }

                if (results.List2 != null)
                {
                    this.ValidateListElement(results.List2);
                }
            }

            if (respResults.Schema != null)
            {
                // If the Schema is not null, it indicate that the server does return a Schema element as is specified.
                this.Site.CaptureRequirement(
                    861,
                    @"[In Schema] The Schema element is specified in [MS-PRSTFR] section 2.3.1.2. The following XML schema defines the Schema element:
                  <xs:element  name='Schema'>
                    <xs:complexType >
                      <xs:choice  minOccurs='0' maxOccurs='unbounded'>
                        <xs:element  name= 'AttributeType' type='xdr:AttributeType'/>
                        <xs:element  name= 'ElementType' type='xdr:ElementType'/>
                        <xs:element  ref='xdr:description'/>
                        <xs:any  namespace='##other' processContents='skip'/>
                      </xs:choice>
                      <xs:attribute  name='name' type='xs:string/>
                      <xs:attribute  name='id' type='xs:ID/>
                      <xs:anyAttribute  namespace='##other'
                                        processContents='skip'/>
                    </xs:complexType>
                  </xs:element>");

                // If MS-DWSS_R861 is verified correctly, this requirement also can be verified directly.
                this.Site.CaptureRequirement(
                    862,
                    @"[In Schema] AttributeType: Specified in [MS-PRSTFR] section 2.3.1.2.");

                // If MS-DWSS_R861 is verified correctly, this requirement also can be verified directly.
                this.Site.CaptureRequirement(
                    863,
                    @"[In Schema] ElementType: Specified in [MS-PRSTFR] section 2.3.1.2.");

                // If MS-DWSS_R861 is verified correctly, this requirement also can be verified directly.
                this.Site.CaptureRequirement(
                    864,
                    @"[In Schema] description: A description for the Schema element.");

                // If MS-DWSS_R861 is verified correctly, this requirement also can be verified directly.
                this.Site.CaptureRequirement(
                    865,
                    @"[In Schema] name: The name of the schema.");

                // If MS-DWSS_R861 is verified correctly, this requirement also can be verified directly.
                this.Site.CaptureRequirement(
                    866,
                    @"[In Schema] id: The identifier of the schema. Specified in [MS-PRSTFR] section 2.3.1.2.");

                // If MS-DWSS_R861 is verified correctly, this requirement also can be verified directly.
                Site.CaptureRequirement(
                    254,
                    @"[In GetDwsMetaDataResponse] The element[Schema] MUST conform to the Schema element in the xml element specified in [MS-PRSTFR] section 2.3.1.2.");

                // If MS-DWSS_R861 is verified correctly, this requirement also can be verified directly.
                Site.CaptureRequirement(
                    256,
                    @"[In GetDwsMetaDataResponse] The element[Schema] MUST conform to the Schema element in the xml element specified in [MS-PRSTFR] section 2.3.1.2.");

                // If MS-DWSS_R861 is verified correctly, this requirement also can be verified directly.
                Site.CaptureRequirement(
                    258,
                    @"[In GetDwsMetaDataResponse] This element [Schema] MUST conform to the Schema element in the xml element specified in [MS-PRSTFR] section 2.3.1.2.");
            }
        }

        /// <summary>
        /// Validate the requirements related to List element.
        /// </summary>
        /// <param name="list">List element.</param>
        private void ValidateListElement(List list)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R351");

            // Verify MS-DWSS requirement: MS-DWSS_R351
            Site.CaptureRequirement(
                351,
                @"[In List] It[List] is defined as follows:
                  <xs:element name=""List"">
                    <xs:complexType>
                      <xs:choice>
                        <xs:element ref=""tns:Error""/>
                        <xs:choice>
                          <xs:element name=""NoChanges"" type=""xs:string""/>
                          <xs:sequence>
                            <xs:sequence>
                              <xs:element ref=""tns:ID""/>
                              <xs:choice>
                                <xs:element ref=""tns:Error"" minOccurs=""0""/>
                                <xs:sequence>
                                  <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##other""/>
                                </xs:sequence>
                              </xs:choice>
                            </xs:sequence>
                          </xs:sequence>
                        </xs:choice>
                      </xs:choice>
                      <xs:attribute name=""Name"" type=""tns:ListType"" use=""required""/>
                    </xs:complexType>
                  </xs:element>");

            XmlLinkedNode anyNode = null;

            for (int i = 0; i < list.Items.Length; i++)
            {
                anyNode = list.Items[i] as XmlLinkedNode;

                if (anyNode != null && anyNode.Name == "z:row")
                {
                    Site.Assert.AreEqual<string>(
                        "#RowsetSchema",
                        anyNode.NamespaceURI,
                        "In a z:row element, the z should be equal to #RowsetSchema in the ActiveX Data Objects (ADO) XML Persistence format.");
                }
            }

            // Verify MS-DWSS requirement: MS-DWSS_R728
            Site.CaptureRequirement(
                728,
                "[In List] The List element can contain an array of z:row elements in any of its elements.");

            // Verify MS-DWSS requirement: MS-DWSS_R362
            Site.CaptureRequirement(
                362,
                "[In List] In a z:row element, the z is equal to #RowsetSchema in the ActiveX Data Objects (ADO) XML Persistence format (see [MS-PRSTFR]).");
        }

        /// <summary>
        /// Validate the requirements related to schema for RemoveDwsUser operation.
        /// </summary>
        /// <param name="respXmlString">RemoveDwsUserResult decoded xml string returned by server.</param>
        private void ValidateRemoveDwsUserResponseSchema(string respXmlString)
        {
            // Validate response xml schema.
            AdapterHelper.ValidateRespSchema(respXmlString);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R288");

            // Verify MS-DWSS requirement: MS-DWSS_R288. If there is any validation error or warning occurred in previous step, an exception will be thrown.
            Site.CaptureRequirement(
                288,
                @"[In RemoveDwsUserResponse] RemoveDwsUserResult: This element contains a string that is stand–alone XML as follows:
                <s:complexType>
                  <xs:choice>
                      <xs:element ref=""tns:Error""/>
                      <xs:element name=""Results""/>
                  </xs:choice>
                </s:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R284");

            // Verify MS-DWSS requirement: MS-DWSS_R284
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                284,
                @"[In RemoveDwsUser] The protocol client sends a RemoveDwsUserSoapIn request message, and the protocol server responds with a RemoveDwsUserSoapOut response message, as follows:
                <wsdl:operation name=""RemoveDwsUser"">
                    <wsdl:input message=""tns:RemoveDwsUserSoapIn"" />
                    <wsdl:output message=""tns:RemoveDwsUserSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R645");

            // Verify MS-DWSS requirement: MS-DWSS_R645
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                645,
                @"[In RemoveDwsUserResponse] This element[RemoveDwsUserResponse] is defined as follows:
                <s:element name=""RemoveDwsUserResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""RemoveDwsUserResult"" type=""s:string"" minOccurs=""0"" maxOccurs=""1""/>
                    </s:sequence>
                  </s:complexType>
                </s:element>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R640");

            // Verify MS-DWSS requirement: MS-DWSS_R640
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                640,
                @"[In RemoveDwsUserSoapOut] The SOAP body contains a RemoveDwsUserResponse element.");
        }

        /// <summary>
        /// Validate the requirements related to Results element in RemoveDwsUser operation.
        /// </summary>
        private void ValidateRemoveDwsUserResultResults()
        {
            // Verify MS-DWSS requirement: MS-DWSS_R285
            Site.CaptureRequirement(
                285,
                @"[In RemoveDwsUser] If the protocol server successfully deletes the specified user from the workspace members list, the protocol server MUST return a string containing an empty Result element as follows:
                        <s:element name=""Results"">
                          <s:complexType/>
                        </s:element>");

            // Verify MS-DWSS requirement: MS-DWSS_R290
            Site.CaptureRequirement(
                290,
                "[In RemoveDwsUserResponse] Results: This is an empty element that indicates success.");
        }

        /// <summary>
        /// Validate the requirements related to schema for RenameDws operation.
        /// </summary>
        /// <param name="respXmlString">RenameDwsResult decoded xml string returned by server.</param>
        private void ValidateRenameDwsResponseSchema(string respXmlString)
        {
            // Validate RenameDwsResult xml schema.
            AdapterHelper.ValidateRespSchema(respXmlString);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R298");

            // Verify MS-DWSS requirement: MS-DWSS_R298. If there is any validation error or warning occurred in previous step, an exception will be thrown.
            Site.CaptureRequirement(
                298,
                @"[In RenameDwsResponse] RenameDwsResult: This element contains a string that is standalone XML as follows:
                <xs:complexType>
                  <xs:choice>
                    <xs:element ref=""tns:Error""/>
                    <xs:element name=""Result""/>
                  </xs:choice>
                </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R292");

            // Verify MS-DWSS requirement: MS-DWSS_R292
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                292,
                @"[In RenameDws] The protocol client sends a RenameDwsSoapIn request message, and the protocol server responds with a RenameDwsSoapOut response message, as follows:
                <wsdl:operation name=""RenameDws"">
                    <wsdl:input message=""tns:RenameDwsSoapIn"" />
                    <wsdl:output message=""tns:RenameDwsSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R656");

            // Verify MS-DWSS requirement: MS-DWSS_R656
            // The schema of RenameDws operation has been validated by full WSDL. If it returns success, the schema of RenameDws operation is valid, capture related requirements.
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                656,
                @"[In RenameDwsResponse] This element[RenameDwsResponse] is defined as follows:
                <s:element name=""RenameDwsResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""RenameDwsResult"" type=""s:string"" minOccurs=""0"" maxOccurs=""1""/>
                    </s:sequence>
                  </s:complexType>
                </s:element>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R651");

            // Verify MS-DWSS requirement: MS-DWSS_R651
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                651,
                @"[In RenameDwsSoapOut] The SOAP body contains a RenameDwsResponse element.");
        }

        /// <summary>
        /// Validate the requirements related to Result element in RenameDws operation.
        /// </summary>
        private void ValidateRenameDwsResultResult()
        {
            // Verify MS-DWSS requirement: MS-DWSS_R297
            Site.CaptureRequirement(
                297,
                "[In RenameDws] If the protocol server successfully changes the title of the workspace, it MUST return an empty Result element.");

            // Verify MS-DWSS requirement: MS-DWSS_R300
            Site.CaptureRequirement(
                300,
                "[In RenameDwsResponse] Result: This is an empty element that indicates success.");
        }

        /// <summary>
        /// Verify the UpdateDwsData operation related requirements.
        /// </summary>
        private void ValidateUpdateDwsData()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R302");

            // Verify MS-DWSS requirement: MS-DWSS_R302
            Site.CaptureRequirementIfAreEqual<ValidationResult>(
                ValidationResult.Success,
                SchemaValidation.ValidationResult,
                302,
                @"[In UpdateDwsData] This operation[UpdateDwsData] is defined as follows:
                <wsdl:operation name=""UpdateDwsData"">
                    <wsdl:input message=""tns:UpdateDwsDataSoapIn"" />
                    <wsdl:output message=""tns:UpdateDwsDataSoapOut"" />
                </wsdl:operation>
                The protocol client sends an UpdateDwsDataSoapIn request message, and the protocol server responds with an UpdateDwsDataSoapOut response message.");
        }

        /// <summary>
        /// Validates the protocol transport-related requirements.
        /// </summary>
        private void ValidateProtocolTransport()
        {
            switch (this.transport)
            {
                case TransportProtocol.HTTP:
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R1");

                        // Verify MS-DWSS requirement: MS-DWSS_R1
                        // Because Adapter uses SOAP and HTTP to communicate with server, if server returned data without exception, this requirement has been captured.
                        Site.CaptureRequirement(
                            1,
                            @"[In Transport] Protocol servers MUST support SOAP over HTTP.");
                    }

                    break;
                case TransportProtocol.HTTPS:
                    {
                        if (Common.IsRequirementEnabled(686, Site))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R686");

                            // Verify MS-DWSS requirement: MS-DWSS_R686
                            Site.CaptureRequirement(
                                686,
                                @"[In Appendix B: Product Behavior] Implementation does additionally support SOAP over HTTPS for securing communication with protocol clients. (Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
                        }
                    }

                    break;
                default:
                    Site.Assume.Fail("Un-recognized transport type: {0}.", this.transport.ToString());
                    break;
            }
        }

        /// <summary>
        /// Validates the protocol soap-related requirements for the CanCreateDwsUrl operation.
        /// </summary>
        private void ValidateSoapVersion()
        {
            // According to the implementation of adapter, the message is formatted as SOAP1.1 or soap 1.2. If this operation is invoked successfully, then this requirement can be verified.
            switch (this.soapVersion)
            {
                case SoapVersion.SOAP11:
                case SoapVersion.SOAP12:
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R4, the current soap version is {0}", this.soapVersion.ToString());

                        // Verify MS-DWSS requirement: MS-DWSS_R4
                        Site.CaptureRequirement(
                            4,
                            @"[In Transport] Protocol messages MUST be formatted as specified in [SOAP1.1] SOAP Envelope section 4, or in [SOAP1.2/1] SOAP Message Construct section 5. ");
                    }

                    break;
                default:
                    Site.Assume.Fail("Un-recognized soap version: {0}.", this.soapVersion.ToString());
                    break;
            }
        }
    }
}