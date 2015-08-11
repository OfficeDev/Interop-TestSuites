namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class contains methods which capture requirements related with Cell Sub-request.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related with Cell Sub-request.
        /// </summary>
        /// <param name="cellSubResponse">Containing the CellSubResponse information</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateCellSubResponse(CellSubResponseType cellSubResponse, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R553
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     553,
                     @"[In CellSubResponseType][CellSubResponseType schema is:]
                     <xs:complexType name=""CellSubResponseType"">
                        <xs:complexContent>
                          <xs:extension base=""tns:SubResponseType"">
                            <xs:sequence>
                               <xs:element name=""SubResponseData"" type=""tns:CellSubResponseDataType"" minOccurs=""0"" maxOccurs=""1"" />
                               <xs:element name=""SubResponseStreamInvalid"" minOccurs=""0"" maxOccurs=""1"" />
                            </xs:sequence>
                          </xs:extension>
                        </xs:complexContent>
                     </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R951
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(CellSubResponseType),
                     cellSubResponse.GetType(),
                     "MS-FSSHTTP",
                     951,
                     @"[In Cell Subrequest][The protocol client sends a cell SubRequest message, which is of type CellSubRequestType,] The protocol server responds with a cell SubResponse message, which is of type CellSubResponseType as specified in section 2.3.1.4.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1687
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(CellSubResponseType),
                     cellSubResponse.GetType(),
                     "MS-FSSHTTP",
                     1687,
                     @"[In SubResponseElementGenericType] Depending on the Type attribute specified in the SubRequest element, the SubResponseElementGenericType MUST take one of the forms: CellSubResponseType.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R274
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(CellSubResponseType),
                     cellSubResponse.GetType(),
                     "MS-FSSHTTP",
                     274,
                     @"[In SubResponseType] The SubResponseElementGenericType takes one of the following forms: CellSubResponseType.");

            ErrorCodeType errorCode;
            site.Assert.IsTrue(Enum.TryParse<ErrorCodeType>(cellSubResponse.ErrorCode, true, out errorCode), "Fail to convert the error code string {0} to the Enum type ErrorCodeType", cellSubResponse.ErrorCode);

            if (cellSubResponse.ErrorCode != null)
            {
                ValidateCellRequestErrorCodeTypes(errorCode, site);
            }

            if (errorCode == ErrorCodeType.Success)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R263
                site.Log.Add(
                    LogEntryKind.Debug,
                    "For requirement MS-FSSHTTP_R263, the SubResponseData value should not be NULL when the cell sub-request succeeds, the actual SubResponseData value is: {0}",
                    cellSubResponse.SubResponseData != null ? cellSubResponse.SubResponseData.ToString() : "NULL");

                site.CaptureRequirementIfIsNotNull(
                         cellSubResponse.SubResponseData,
                         "MS-FSSHTTP",
                         263,
                         @"[In SubResponseElementGenericType][The SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"" and one of the following conditions is true:] The Type attribute that is specified in the SubRequest element is set to a value of ""Cell"".");
            }

            // Verify requirements related with its base type : SubResponseType
            ValidateSubResponseType(cellSubResponse as SubResponseType, site);

            // Verify requirements related with CellSubResponseDataType
            if (cellSubResponse.SubResponseData != null)
            {
                ValidateCellSubResponseDataType(cellSubResponse.SubResponseData, site);
            }
        }

        /// <summary>
        /// Capture requirements related with CellSubResponseDataType.
        /// </summary>
        /// <param name="cellSubResponseData">The cellSubResponseData</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateCellSubResponseDataType(CellSubResponseDataType cellSubResponseData, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1391
            // The SubResponseData of CellSubResponse is of type CellSubResponseDataType, so if cellSubResponse.SubResponseData is not null, then MS-FSSHTTP_R1391 can be captured.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(CellSubResponseDataType),
                     cellSubResponseData.GetType(),
                     "MS-FSSHTTP",
                     1391,
                     @"[In SubResponseDataGenericType][SubResponseDataGenericType MUST take one of the forms described in the following table] CellSubResponseDataType: Type definition for cell subresponse data.");

            if (cellSubResponseData.LockTypeSpecified)
            {
                // Verified LockTypes
                ValidateLockTypes(site);
            }

            // Verify requirements related with CellSubResponseDataOptionalAttributes
            if (cellSubResponseData.Etag != null
                || cellSubResponseData.LastModifiedTime != null
                || cellSubResponseData.CreateTime != null
                || cellSubResponseData.ModifiedBy != null
                || cellSubResponseData.CoalesceErrorMessage != null)
            {
                ValidateCellSubResponseDataOptionalAttributes(cellSubResponseData, site);
            }
        }

        /// <summary>
        /// Capture requirements related with CellSubResponseDataOptionalAttributes.
        /// </summary>
        /// <param name="cellSubResponseData">The cellSubResponseData</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateCellSubResponseDataOptionalAttributes(CellSubResponseDataType cellSubResponseData, ITestSite site)
        {
            if (cellSubResponseData.ModifiedBy != null)
            {
                // Verify requirements related with UserNameTypes
                ValidateUserNameTypes(site);
            }

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1465
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1465,
                     @"[In CellSubResponseDataOptionalAttributes] The CellSubResponseDataOptionalAttributes attribute group contains attributes that is used in SubResponseData elements associated with a SubResponse for a cell subrequest.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1497
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     1497,
                     @"[In SubResponseDataOptionalAttributes] CellSubResponseDataOptionalAttributes: An attribute group that specifies attributes that MUST be used for SubResponseData elements associated with a subresponse for a cell subrequest.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R839
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     839,
                     @"[In CellSubResponseDataOptionalAttributes] The definition of the CellSubResponseDataOptionalAttributes attribute group is as follows:
                     <xs:attributeGroup name=""CellSubResponseDataOptionalAttributes"">
                         <xs:attribute name=""Etag"" type=""xs:string"" use=""optional"" />
                         <xs:attribute name=""CreateTime"" type=""xs:integer"" use=""optional""/>
                         <xs:attribute name=""LastModifiedTime"" type=""xs:integer"" use=""optional""/>
                         <xs:attribute name=""ModifiedBy"" type=""tns:UserNameType"" use=""optional"" />
                         <xs:attribute name=""CoalesceErrorMessage"" type=""xs:string"" use=""optional""/>
                         <xs:attribute name=""CoalesceHResult"" type=""xs:integer"" use=""optional""/>
                         <xs:attribute name=""ContainsHotboxData"" type=""tns:TRUEFALSE"" use=""optional""/>
                         <xs:attribute name=""HaveOnlyDemotionChanges"" type=""tns:TRUEFALSE"" use=""optional""/>
                     </xs:attributeGroup>");
        }

        /// <summary>
        /// Capture requirements related with CellRequestErrorCodeTypes.
        /// </summary>
        /// <param name="cellRequestErrorCode">A cellRequestErrorCode</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateCellRequestErrorCodeTypes(ErrorCodeType cellRequestErrorCode, ITestSite site)
        {
            if (cellRequestErrorCode == ErrorCodeType.CellRequestFail || cellRequestErrorCode == ErrorCodeType.IRMDocLibarysOnlySupportWebDAV)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R772
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         772,
                         @"[In CellRequestErrorCodeTypes][CellRequestErrorCodeTypes schema is:]
                         <xs:simpleType name=""CellRequestErrorCodeTypes"">
                           <xs:restriction base=""xs:string"">
                             <!--cell request fail-->
                             <xs:enumeration value=""CellRequestFail""/>
                                <!--cell request etag not matching-->
                             <xs:enumeration value=""IRMDocLibarysOnlySupportWebDAV""/>
                           </xs:restriction>
                         </xs:simpleType>");

                // If the validation succeed, then the requirement MS-FSSHTTP_R773 can be captured.
                site.CaptureRequirement(
                         "MS-FSSHTTP",
                         773,
                         @"[In CellRequestErrorCodeTypes] The value of CellRequestErrorCodeTypes MUST be one of the following:
                         [CellRequestFail, IRMDocLibarysOnlySupportWebDAV]");
            }
        }
    }
}