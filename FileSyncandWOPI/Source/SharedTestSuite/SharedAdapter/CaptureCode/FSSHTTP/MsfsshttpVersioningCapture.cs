namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class contains methods which capture requirements related with WhoAmI Sub-request.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related with Versioning Sub-request.
        /// </summary>
        /// <param name="versioningSubResponse">Containing the VersioningSubResponse information</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateVersioningSubResponse(VersioningSubResponseType versioningSubResponse, ITestSite site)
        {
            ErrorCodeType errorCode = (ErrorCodeType)Enum.Parse(typeof(ErrorCodeType), versioningSubResponse.ErrorCode);
            if (errorCode == ErrorCodeType.VersionNotFound)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11081
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11081,
                    @"[In VersioningRelatedErrorCodeTypes] [The schema of VersioningRelatedErrorCodeTypes is:]
                     <xs:simpleType name=""VersioningRelatedErrorCodeTypes"">
                        < xs:restriction base = ""xs:string"" >
                           < xs:enumeration value = ""VersionNotFound"" />
                        </ xs:restriction >
                     </ xs:simpleType > ");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11082
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11082,
                    @"[In VersioningRelatedErrorCodeTypes] The value of VersioningRelatedErrorCodeTypes MUST be one of value [VersionNotFound] in the following table.");
            }

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11236
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11236,
                @"[In Versioning Subrequest] The protocol server responds with a versioning SubResponse message, which is of type VersioningSubResponseType as specified in section 2.3.1.39.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11242
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11242,
                @"[In Versioning Subrequest] The VersioningSubResponseDataType defines the type of the SubResponseData element inside the versioning SubResponse element.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11154
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11154,
                @"[In VersioningSubResponseType] 
                 <xs:complexType name=""VersioningSubResponseType"">
                   < xs:complexContent >
                     < xs:extension base = ""tns:SubResponseType"" >
                       < xs:sequence minOccurs = ""0"" maxOccurs = ""1"" >
                          < xs:element name = ""SubResponseData"" type = ""tns:VersioningSubResponseDataType"" />
                       </ xs:sequence >
                     </ xs:extension >
                   </ xs:complexContent >
                 </ xs:complexType > ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11155
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11155,
                @"[In VersioningSubResponseType] SubResponseData: A VersioningSubResponseDataType that specifies versioning related information provided by the protocol server that was requested as part of the versioning subrequest.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11144
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11144,
                @"[In VersioningSubResponseDataType] 
                <xs:complexType name=""VersioningSubResponseDataType"">
                    <xs:sequence minOccurs=""0"" maxOccurs=""1"">
                        <xs:sequence minOccurs=""0"" maxOccurs=""1"">
                            <xs:element name=""UserTable"" type=""tns:VersioningUserTableType""/>
                        </xs:sequence>
                        <xs:element name=""Versions"" type=""tns:VersioningVersionListType""/>
                    </xs:sequence>
                </xs:complexType>");

            if (versioningSubResponse.SubResponseData.UserTable != null)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11145
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11145,
                    @"[In VersioningSubResponseDataType] UserTable: An element of type VersioningUserTableType (section 2.3.1.40) that specifies data for the users represented in the version list.");

                ValidateVersioningUserTableType(versioningSubResponse.SubResponseData.UserTable, site);


            }

            if (versioningSubResponse.SubResponseData.Versions != null)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11147
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11147,
                    @"[In VersioningSubResponseDataType] Versions: An element of type VersioningVersionListType (section 2.3.1.41) that specifies the list of versions of this file that exist on the server.");

                ValidateVersioningVersionListType(versioningSubResponse.SubResponseData.Versions, site);
            }
        }

        /// <summary>
        /// Capture requirements related with VersioningUserTableType.
        /// </summary>
        /// <param name="userTable">The VersioningUserTableType</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateVersioningUserTableType(VersioningUserTableType userTable, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11158
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11158,
                @"[In VersioningUserTableType] 
 <xs:complexType name=""VersioningUserTableType"">
       <xs:sequence>
          <xs:element name=""User"" maxOccurs=""unbounded"" minOccurs=""1"" type=""tns:UserDataType"" />
       </xs:sequence>
 </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11159
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11159,
                @"[In VersioningUserTableType] User: An element of type UserDataType (section 2.3.1.42) which describes a single user.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11164
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11164,
                @"[In UserDataType] 
                     <xs:complexType name=""UserDataType"">
                        < xs:attribute name = ""UserId"" type = ""xs:integer"" use = ""required"" />
                        < xs:attribute name = ""UserLogin"" type = ""xs:UserLoginType"" use = ""required"" />
                        < xs:attribute name = ""UserName"" type = ""xs:UserNameType"" use = ""optional"" />
                        < xs:attribute name = ""UserEmailAddress"" type = ""s:string"" use = ""optional"" />
                     </ xs:complexType > ");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11165
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11165,
                @"[In UserDataType] UserId: An integer that uniquely specifies the user in this user table.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11167
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11167,
                @"[In UserDataType] UserLogin: A UserLoginType that specifies the user login alias of  the protocol client.");


            if (!string.IsNullOrEmpty(userTable.User[0].UserName))
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11169
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11169,
                    @"[In UserDataType] UserName: A UserNameType that specifies the user name for the protocol client.");
            }

            if (!string.IsNullOrEmpty(userTable.User[0].UserEmailAddress))
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11171
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11171,
                    @"[In UserDataType] UserEmailAddress: A string that specifies the email address associated with the protocol client.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11172
                site.CaptureRequirementIfIsTrue(
                    AdapterHelper.IsValidEmailAddr(userTable.User[0].UserEmailAddress),
                    "MS-FSSHTTP",
                    11172,
                    @"[In UserDataType] The format of the email address MUST be as specified in [RFC2822] section 3.4.1.");
            }
        }

        /// <summary>
        /// Capture requirements related with VersioningVersionListType.
        /// </summary>
        /// <param name="versionList">The VersioningVersionListType</param>
        /// <param name="site">Instance of ITestSite</param>
        private static void ValidateVersioningVersionListType(VersioningVersionListType versionList, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11161
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11161,
                @"[In VersioningVersionListType] 
 <xs:complexType name=""VersioningVersionListType"">
       <xs:sequence>
          <xs:element name=""Version"" maxOccurs=""unbounded"" minOccurs=""1"" type=""tns:FileVersionDataType"" />
       <xs:sequence>
 </xs:complexType>");

            if (versionList.Version != null)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11162
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11162,
                    @"[In VersioningVersionListType] Version: An element of type FileVersionDataType (section 2.3.1.43) which describes a single version of the file on the server.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11174
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11174,
                    @"[In FileVersionDataType] 
 <xs:complexType name=""FileVersionDataType"">
 <xs:sequence>
    <xs:element name=""Events"" minOccurs=""0"" maxOccurs=""1"" >
      <xs:complexType>
        <xs:sequence>
          <xs:element name=""Event"" minOccurs=""1"" maxOccurs=""unbounded"" type=""tns:FileVersionEventDataType"" />
        </xs:sequence>
      </xs:complexType>
    </xs:element>
  </xs:sequence>
  <xs:attribute name=""IsCurrent"" type=""tns:TRUEFALSE"" use=""optional"" />
  <xs:attribute name=""Number"" type=""tns:FileVersionNumberType"" use=""required"" />
  <xs:attribute name=""LastModifiedTime"" type=""xs:positiveInteger"" use=""optional"" />
  <xs:attribute name=""UserId"" type=""xs:integer"" use=""optional"" />
 </xs:complexType>");
            }

            if (!string.IsNullOrEmpty(versionList.Version[0].Number))
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11176
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11176,
                    @"[In FileVersionDataType] Number: A FileVersionNumberType (section 2.2.5.15) that specifies the unique version number of the version of the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11078
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11078,
                    @"[In FileVersionNumberType] [The schema of FileVersionNumberType is:] 
<xs:simpleType name=""FileVersionNumberType"">
    <xs:restriction base=""xs:string"">
      <xs:pattern value=""[0-9]+[.][0-9]+"" />
    </xs:restriction>
 </xs:simpleType>");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11079
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11079,
                    @"[In FileVersionNumberType] The value of a FileVersionNumberType is typically in the form major.minor (for example, 1.0).");
            }

            if (!string.IsNullOrEmpty(versionList.Version[0].LastModifiedTime))
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11177
                site.CaptureRequirement(
                "MS-FSSHTTP",
                11177,
                @"[In FileVersionDataType] LastModifiedTime: A positive integer that specifies the last modified time of the version of the file, which is expressed as a tick count.");
            }

            if (versionList.Version[0].Events != null && versionList.Version[0].Events != null)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11182
                site.CaptureRequirement(
                "MS-FSSHTTP",
                11182,
                @"[In FileVersionDataType] Events: An array of events of type FileVersionEventDataType that represents an event that happened to the version of the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11184
                site.CaptureRequirement(
                "MS-FSSHTTP",
                11184,
                @"[In FileVersionEventDataType] 
                 <xs:complexType name=""FileVersionEventDataType"">
                    < xs:attribute name = ""Id"" type = ""xs:integer"" use = ""required"" />
                    < xs:attribute name = ""Type"" type = ""xs:integer"" use = ""required"" />
                    < xs:attribute name = ""CreateTime"" type = ""xs:positiveInteger"" use = ""optional"" />
                    < xs:attribute name = ""UserId"" type = ""xs:integer"" use = ""optional"" />
                 </ xs:complexType > ");

                System.Collections.Generic.List<string> ids = new System.Collections.Generic.List<string>();

                bool isR11185Verified = true;
                foreach (FileVersionDataType version in versionList.Version)
                {
                    if (version.Events != null)
                    {
                        foreach (FileVersionEventDataType eventData in version.Events)
                        {
                            if (!ids.Contains(eventData.Id))
                            {
                                ids.Add(eventData.Id);
                            }
                            else
                            {
                                isR11185Verified = false;
                                break;
                            }
                        }
                    }
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11185
                site.CaptureRequirementIfIsTrue(
                    isR11185Verified,
                    "MS-FSSHTTP",
                    11185,
                    @"[In FileVersionEventDataType] Id: An integer that uniquely identifies an event among all events to all versions of the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11186
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11186,
                    @"[In FileVersionEventDataType] Type: An integer that identifies the type of event that occurred to the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11191
                site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11191,
                    @"[In FileVersionEventDataType] CreateTime: A positive integer that specifies the time when this event occurred, which is expressed as a tick count.");

                foreach (FileVersionDataType version in versionList.Version)
                {
                    if (version.Events != null)
                    {
                        foreach (FileVersionEventDataType eventData in version.Events)
                        {
                            bool isR11187Verified = versionList.Version[0].Events[0].Type.Equals("1", StringComparison.CurrentCultureIgnoreCase)
                                || versionList.Version[0].Events[0].Type.Equals("2", StringComparison.CurrentCultureIgnoreCase)
                                || versionList.Version[0].Events[0].Type.Equals("3", StringComparison.CurrentCultureIgnoreCase);

                            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11187
                            site.CaptureRequirementIfIsTrue(
                                isR11187Verified,
                                "MS-FSSHTTP",
                                11187,
                                @"[In FileVersionEventDataType] The value MUST be one of the values [1, 2, 3] in the following table.");
                        }
                    }
                }
            }
        }
    }
}