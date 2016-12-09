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
                        < s:attribute name = ""UserId"" type = ""xs:integer"" use = ""required"" />
                        < s:attribute name = ""UserLogin"" type = ""xs:UserLoginType"" use = ""required"" />
                        < s:attribute name = ""UserName"" type = ""xs:UserNameType"" use = ""optional"" />
                        < s:attribute name = ""UserEmailAddress"" type = ""s:string"" use = ""optional"" />
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
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11162
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11162,
                @"[In VersioningVersionListType] Version: An element of type FileVersionDataType (section 2.3.1.43) which describes a single version of the file on the server.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11176
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11176,
                @"[In FileVersionDataType] Number: A FileVersionNumberType (section 2.2.5.15) that specifies the unique version number of the version of the file.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11177
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11177,
                @"[In FileVersionDataType] LastModifiedTime: A positive integer that specifies the last modified time of the version of the file, which is expressed as a tick count.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11182
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11182,
                @"[In FileVersionDataType] Event: A FileVersionEventDataType that represents an event that happened to the version of the file.");
        }
    }
}