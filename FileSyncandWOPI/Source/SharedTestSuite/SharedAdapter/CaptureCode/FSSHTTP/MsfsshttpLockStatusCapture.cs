namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class contains methods which capture requirements related with LockStatus Sub-request.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related with LockStatus Sub-request.
        /// </summary>
        /// <param name="lockStatusSubResponse">Containing the LockStatusSubResponse information</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateLockStatusSubResponse(LockStatusSubResponseType lockStatusSubResponse, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2276
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2276,
                     @"[LockStatusSubResponseType]
	<xs:complexType name=""LockStatusSubResponseType"">
	  <xs:complexContent>
	    <xs:extension base=""tns:SubResponseType"">
	      <xs:sequence minOccurs=""0"" maxOccurs=""1"">
	         <xs:element name=""SubResponseData"" type=""tns:LockStatusSubResponseDataType"" />
	      </xs:sequence>
	    </xs:extension>
	  </xs:complexContent>
	</xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2277
            site.CaptureRequirement(
                "MS-FSSHTTP",
                2277,
                @"[LockStatusSubResponseType]SubResponseData: A LockStatusSubResponseDataType that specifies the information about the lock status of a file that was requested as part of the LockStatus subrequest. ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2377
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(LockStatusSubResponseType),
                     lockStatusSubResponse.GetType(),
                     "MS-FSSHTTP",
                     2377,
                     @"[LockStatus Subrequest]The protocol server responds with a LockStatus SubResponse message, which is of type LockStatusSubResponseType as specified in section 2.3.1.51. ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2147
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(LockStatusSubResponseType),
                     lockStatusSubResponse.GetType(),
                     "MS-FSSHTTP",
                     2147,
                     @"[In SubResponseElementGenericType] Depending on the Type attribute specified in the SubRequest element, the SubResponseElementGenericType MUST take one of the forms: LockStatusSubResponseType");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2165
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(LockStatusSubResponseType),
                     lockStatusSubResponse.GetType(),
                     "MS-FSSHTTP",
                     2165,
                     @"[In SubResponseType] The SubResponseElementGenericType takes one of the following forms: LockStatusSubResponseType.");

            ErrorCodeType errorCode;
            site.Assert.IsTrue(Enum.TryParse<ErrorCodeType>(lockStatusSubResponse.ErrorCode, true, out errorCode), "Fail to convert the error code string {0} to the Enum type ErrorCodeType", lockStatusSubResponse.ErrorCode);
            if (errorCode == ErrorCodeType.Success)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2159
                site.CaptureRequirementIfIsNotNull(
                         lockStatusSubResponse.SubResponseData,
                         "MS-FSSHTTP",
                         2159,
                         @"[In SubResponseElementGenericType][The SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"" and one of the following conditions is true:] The Type attribute that is specified in the SubRequest element is set to a value of ""LockStatus"".");
            }

            // Verify requirements related with its base type: SubResponseType
            ValidateSubResponseType(lockStatusSubResponse as SubResponseType, site);

            // Verify requirements related with SubResponseDataType
            if (lockStatusSubResponse.SubResponseData != null)
            {
                ValidateLockStatusSubResponseDataType(lockStatusSubResponse.SubResponseData, site);
            }
        }

        /// <summary>
        /// Capture requirements related with LockStatusSubResponseDataType
        /// </summary>
        /// <param name="lockStatusSubResponseData">The LockStatusSubResponseData information</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateLockStatusSubResponseDataType(LockStatusSubResponseDataType lockStatusSubResponseData, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2382
            // if can launch this method, the schema matches.
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2382,
                     @"[LockStatus Subrequest]The LockStatusSubResponseDataType defines the type of the SubResponseData element inside the LockStatusSubResponse element. ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2266
            site.CaptureRequirement(
                     "MS-FSSHTTP",
                     2266,
                     @"[In LockStatusSubResponseDataType]
	<xs:complexType name=""LockStatusSubResponseDataType"">
	    <xs:attribute name=""LockType"" type=""tns:LockTypes"" use=""optional"" />
	    <xs:attribute name=""LockID"" type=""tns:guid"" use=""optional"" />
	    <xs:attribute name=""LockedBy"" type=""xs:string"" use=""optional"" />
	</xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2135
            // The SubResponseData of LockStatusSubResponse is of type LockStatusSubResponseDataType, so if lockStatusSubResponse.SubResponseData is not null, then MS-FSSHTTP_R2135 can be captured.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(LockStatusSubResponseDataType),
                     lockStatusSubResponseData.GetType(),
                     "MS-FSSHTTP",
                     2135,
                     @"[In SubResponseDataGenericType][SubResponseDataGenericType MUST take one of the forms described in the following table]LockStatusSubResponseDataType:Type definition for Lock Status subresponse data.");

            if (lockStatusSubResponseData.LockID != null)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2182
                site.CaptureRequirementIfIsNotNull(
                    lockStatusSubResponseData.LockID,
                    "MS-FSSHTTP",
                    2182,
                    @"[In SubResponseDataOptionalAttributes]LockedID: A guid that specifies the id of the lock.");

                // MS-FSSHTTP_R2182 is verified,so MS-FSSHTTP_R22699 can be verified directly
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    22699,
                    @"[In LockStatusSubResponseDataType]LockedID: A guid that specifies the id of the lock.");
            }

            if (lockStatusSubResponseData.LockedBy != null)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2183
                site.CaptureRequirementIfIsNotNull(
                    lockStatusSubResponseData.LockedBy,
                    "MS-FSSHTTP",
                    2183,
                    @"[In SubResponseDataOptionalAttributes]LockedBy: A string that specifies the user that has the file locked, if any.");

                // MS-FSSHTTP_R2183 is verified,so MS-FSSHTTP_R2270 can be verified directly
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    2270,
                    @"[In LockStatusSubResponseDataType]LockedBy: A string that specifies the user that has the file locked, if any.");
            }

            if(lockStatusSubResponseData.LockTypeSpecified)
            {
                ValidateLockTypes(site);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2267
                site.CaptureRequirementIfIsNotNull(
                    lockStatusSubResponseData.LockType,
                    "MS-FSSHTTP",
                    2267,
                    @"[In LockStatusSubResponseDataType]LockType: A LockTypes that specifies the type of lock granted in a coauthoring subresponse. ");
            }
        }
    }
}