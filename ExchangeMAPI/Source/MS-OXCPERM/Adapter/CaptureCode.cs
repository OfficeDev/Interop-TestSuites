namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter capture code for MS_OXCPERMAdapter
    /// </summary>
    public partial class MS_OXCPERMAdapter : ManagedAdapterBase
    {
        /// <summary>
        /// Verify MAPIHTTP transport.
        /// </summary>
        private void VerifyMAPITransport()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http" && Common.IsRequirementEnabled(1184, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1184");

                // Verify requirement MS-OXCPERM_R1184
                // If the transport sequence is MAPIHTTP and the code can reach here, it means that the implementation does support MAPIHTTP transport.
                Site.CaptureRequirement(
                        1184,
                        @"[In Appendix A: Product Behavior] Implementation does support this specification [MS-OXCMAPIHTTP]. (Exchange Server 2013 Service Pack 1 (SP1) follows this behavior.)");
            }
        }

        #region Message Syntax

        /// <summary>
        /// Verify the message syntax.
        /// </summary>
        private void VerifyMessageSyntax()
        {
            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                3,
                @"[In Transport] The ROP request buffers and ROP response buffers specified in this protocol [Exchange Access and Operation Permissions Protocol] are sent to and received from the server respectively by using the underlying protocol specified by [MS-OXCROPS] section 2.1.");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                6,
                @"[In Message Syntax] Unless otherwise noted, the fields specified in this section, which are larger than a single byte, MUST be converted to little-endian order when packed in buffers.");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                7,
                @"[In Message Syntax] Unless otherwise noted, the fields specified in this section, which are larger than a single byte, MUST be converted from little-endian order when unpacked.");
        }

        #endregion

        #region Verify Properties
        /// <summary>
        /// Verify properties' value
        /// </summary>
        /// <param name="succeedToParse">Specify whether the data can parse successfully or not. True indicate success, else false</param>
        /// <param name="permissionUserList">The permission list of all users</param>
        private void VerifyProperties(bool succeedToParse, List<PermissionUserInfo> permissionUserList)
        {
            Site.Assert.IsTrue(succeedToParse, "True indicates the permissions list is parsed successfully.");
            if (permissionUserList == null || permissionUserList.Count == 0)
            {
                Site.Log.Add(LogEntryKind.Comment, "The permissions list is empty. The data related to the permission can't be verified here.");
                return;
            }

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1068
            // If the returned data can be parsed successfully, the format is following the defined structure.
            bool isVerifyR1068 = succeedToParse;
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1068: The data is{0} same as defined.", isVerifyR1068 ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                isVerifyR1068,
                1068,
                @"[In PidTagEntryId Property] The first two bytes of this property specify the number of bytes that follow.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1069
            bool isVerifyR1069 = succeedToParse;
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1069: The data is{0} same as defined.", isVerifyR1069 ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                isVerifyR1069,
                1069,
                @"[In PidTagEntryId Property] [The first two bytes of this property specify the number of bytes that follow.] The remaining bytes constitute the PermanentEntryID structure ([MS-OXNSPI] section 2.3.9.3).");

            bool memberIdIsUnique = true;
            for (int i = 0; i < permissionUserList.Count; i++)
            {
                PermissionUserInfo permissionInfo = permissionUserList[i];
                string memberName = permissionInfo.PidTagMemberName;
                if (memberName == this.defaultUser || memberName == this.anonymousUser)
                {
                    // Verify MS-OXCPERM requirement: MS-OXCPERM_R1070
                    Site.CaptureRequirementIfAreEqual<int>(
                        0,
                        permissionInfo.PidTagEntryId.Length,
                        1070,
                        @"[In PidTagEntryId Property] If the PidTagMemberId property (section 2.2.5) is set to one of the two reserved values, the first two bytes of this property MUST be 0x0000, indicating that zero bytes follow (that is, no PermanentEntryID structure follows the first two bytes).");
                }

                if (permissionInfo.PidTagMemberId == 0xFFFFFFFFFFFFFFFF)
                {
                    // Verify MS-OXCPERM requirement: MS-OXCPERM_R1079
                    // That PidTagMemberId is 0xFFFFFFFFFFFFFFFF and the name is "Anonymous" must be matched.
                    bool isVerifyR1079 = permissionInfo.PidTagMemberId == 0xFFFFFFFFFFFFFFFF
                                          && permissionInfo.PidTagMemberName == this.anonymousUser;
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1079: the PidTagMemberId:{0:X} and PidTagMemberName:{1} are{2} matched.", permissionInfo.PidTagMemberId, permissionInfo.PidTagMemberName, isVerifyR1079 ? string.Empty : " not");
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR1079,
                        1079,
                        @"[In PidTagMemberName Property] Anonymous user: Value of the PidTagMemberId is 0xFFFFFFFFFFFFFFFF and the name is ""Anonymous"".");
                }

                if (permissionInfo.PidTagMemberId == 0x0000000000000000)
                {
                    // Verify MS-OXCPERM requirement: MS-OXCPERM_R1080
                    // That PidTagMemberId is 0x0000000000000000 and name is "" (empty string) must be matched.
                    bool isVerifyR1080 = permissionInfo.PidTagMemberId == 0x0000000000000000
                                          && permissionInfo.PidTagMemberName == this.defaultUser;
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1080: the PidTagMemberId:{0:X} and PidTagMemberName:{1} are{2} matched.", permissionInfo.PidTagMemberId, permissionInfo.PidTagMemberName, isVerifyR1080 ? string.Empty : " not");
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR1080,
                        1080,
                        @"[In PidTagMemberName Property] Default user: Value of the PidTagMemberId is 0x0000000000000000 and name is """" (empty string).");
                }

                for (int j = i + 1; j < permissionUserList.Count; j++)
                {
                    if (permissionInfo.PidTagMemberId == permissionUserList[j].PidTagMemberId)
                    {
                        memberIdIsUnique = false;
                    }
                }
            }

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R74
            bool isVerifyR74 = memberIdIsUnique;
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R74: the all PidTagMemberId are{0} unique.", isVerifyR74 ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                isVerifyR74,
                74,
                @"[In PidTagMemberId Property] The PidTagMemberId property ([MS-OXPROPS] section 2.773) specifies the unique identifier that the server generates for each user.");

            // The PidTagMemberName is parsed as string, this requirement can be captured directly.
            Site.CaptureRequirement(
                1081,
                @"[In PidTagMemberName Property] The server provides the user-readable name for all entries in the permissions list.");

            // The parser has ensured the field satisfies the format, otherwise can't get the response.
            Site.CaptureRequirement(
                154,
                @"[In Retrieving Folder Permissions] If all three of the ROP requests [RopGetPermissionsTable, RopSetColumns, RopQueryRows] succeed, the permissions list is returned in the RowData field of the RopQueryRows ROP response buffer.");

            // The parser has ensured the field satisfies the format, otherwise can't get the response.
            Site.CaptureRequirement(
                155,
                @"[In Retrieving Folder Permissions] The RowData field contains one PropertyRow structure ([MS-OXCDATA] section 2.8.1) for each entry in the permissions list.");
            
            // The parser has ensured the field satisfies the format, otherwise can't get the response.
            Site.CaptureRequirement(
                1164,
                @"[In Processing a RopGetPermissionsTable ROP Request] The server responds with a RopGetPermissionsTable ROP response buffer.");

            // The parser has ensured the field satisfies the format, otherwise can't get the response.
            Site.CaptureRequirement(
                179,
                @"[In Processing a RopGetPermissionsTable ROP Request] If the user does have permission to view the folder, the server MUST return a Server object handle to a Table object, which can be used to retrieve the permissions list of the folder, as specified in section 3.1.4.1.");

            if (this.CheckUserPermissionContained(this.anonymousUser, permissionUserList))
            {
                // Verify MS-OXCPERM requirement: MS-OXCPERM_R189
                ulong anonymousMemberID = this.GetMemberIdByName(this.anonymousUser, permissionUserList);
                    Site.CaptureRequirementIfAreEqual<ulong>(
                        0xFFFFFFFFFFFFFFFF,
                        anonymousMemberID,
                        189,
                        @"[In PidTagMemberId Property] 0xFFFFFFFFFFFFFFFF: Identifier for the anonymous user entry in the permissions list.");
            }

            if (this.CheckUserPermissionContained(this.defaultUser, permissionUserList))
            {
                // Verify MS-OXCPERM requirement: MS-OXCPERM_R190
                ulong defaultUserMemberID = this.GetMemberIdByName(this.defaultUser, permissionUserList);
                Site.CaptureRequirementIfAreEqual<ulong>(
                    0x0000000000000000,
                    defaultUserMemberID,
                    190,
                    @"[In PidTagMemberId Property] 0x0000000000000000: Identifier for the default user entry in the permissions list.");
            }
        }

        #endregion

        #region Verify the return value.
        /// <summary>
        /// Verify the return value when calling RopGetPermissionsTable is successful.
        /// </summary>
        /// <param name="returnValue">The return value from the server</param>
        private void VerifyReturnValueSuccessForGetPermission(uint returnValue)
        {
            // Verify MS-OXCPERM requirement: MS-OXCPERM_R28
            Site.CaptureRequirementIfAreEqual<uint>(
                UINT32SUCCESS,
                returnValue,
                28,
                @"[In RopGetPermissionsTable ROP Response Buffer] ReturnValue (4 bytes): The value 0x00000000 indicates success.");
        }

        /// <summary>
        /// Verify the return value is 4 bytes.
        /// </summary>
        private void VerifyReturnValueForGetPermission()
        {
            // The return value is parsed as 4 bytes.
            Site.CaptureRequirement(
                27,
                @"[In RopGetPermissionsTable ROP Response Buffer] ReturnValue (4 bytes): An integer that indicates the result of the operation.");
        }

        /// <summary>
        /// Verify the return value when calling RopModifyPermissions is successful.
        /// </summary>
        /// <param name="returnValue">The return value from the server</param>
        private void VerifyReturnValueSuccessForModifyPermission(uint returnValue)
        {
            // Verify MS-OXCPERM requirement: MS-OXCPERM_R66
            Site.CaptureRequirementIfAreEqual<uint>(
                UINT32SUCCESS,
                returnValue,
                66,
                @"[In RopModifyPermissions ROP Response Buffer] ReturnValue (4 bytes): The value 0x00000000 indicates success.");
        }

        /// <summary>
        /// Verify the return value is 4 bytes.
        /// </summary>
        private void VerifyReturnValueForModifyPermission()
        {
            // The return value is parsed as 4 bytes.
            Site.CaptureRequirement(
                65,
                @"[In RopModifyPermissions ROP Response Buffer] ReturnValue (4 bytes): An integer that indicates the result of the operation.");
        }
        #endregion

        #region Capture requirements related to the data structure of MS-OXCPERM
        /// <summary>
        /// Capture Requirements related to the data structures for the MS-OXCPERM
        /// </summary>
        /// <param name="succeedToParse">Specify whether the data can be parsed successfully or not. True indicates success, else false.</param>
        /// <param name="permissionUserList">The PermissionList of all users</param>
        private void VerifyPropertiesOfDataStructure(bool succeedToParse, List<PermissionUserInfo> permissionUserList)
        {
            Site.Assert.IsTrue(succeedToParse, "Succeed to parse the permissions list.");
            const string OXCDATA = "MS-OXCDATA";
            const string OXPROPS = "MS-OXPROPS";

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                OXCDATA,
                65,
                @"[In PropertyRow Structures] For the RopFindRow, RopGetReceiveFolderTable, and RopQueryRows ROPs, property values are returned in the order of the properties in the table, set by a prior call to a RopSetColumns ROP request ([MS-OXCROPS] section 2.2.5.1).");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                OXCDATA,
                72,
                @"[In StandardPropertyRow Structure] Flag (1 byte): An unsigned integer.");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                OXCDATA,
                73,
                @"[In StandardPropertyRow Structure] Flag (1 byte): This value MUST be set to 0x00 to indicate that all property values are present and without error.");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                OXPROPS,
                6100,
                @"[In PidTagEntryId] Property ID: 0x0FFF.");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                OXPROPS,
                6101,
                @"[In PidTagEntryId] Data type: PtypBinary, 0x0102.");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                OXCDATA,
                2707,
                @"[In Property Data Types] PtypBinary (PT_BINARY) is that variable size; a COUNT field followed by that many bytes with Property Type Value 0x0102,%x02.01.");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                1066,
                @"[In PidTagEntryId Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1)");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                OXCDATA,
                454,
                @"[In PropertyValue Structure] PropertyValue (variable):  For multivalue types, the first element in the ROP buffer is a 16-bit integer specifying the number of entries.");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                OXPROPS,
                6904,
                @"[In PidTagMemberId] Property ID: 0x6671.");

            // PtypInteger64 is 8 bytes, so if the length of memberRight is 8 bytes, the data type of PidTagMemberRights is PtypInteger64.
            List<byte[]> memberIDs = this.GetMemberIDs(permissionUserList);
            foreach (byte[] memberID in memberIDs)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6905");

                // Verify MS-OXPROPS requirement: MS-OXPROPS_R6905
                Site.CaptureRequirementIfAreEqual<int>(
                    8,
                    memberID.Length,
                    OXPROPS,
                    6905,
                    @"[In PidTagMemberId] Data type: PtypInteger64, 0x0014.");

                // Verify MS-OXCPERM requirement: MS-OXCPERM_R1071
                Site.CaptureRequirementIfAreEqual<int>(
                    8,
                    memberID.Length,
                    1071,
                @"[In PidTagMemberId Property] Type: PtypInteger64 ([MS-OXCDATA] section 2.11.1)");
            }

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(OXPROPS, 6911, @"[In PidTagMemberName] Property ID: 0x6672.");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(OXPROPS, 6912, @"[In PidTagMemberName] Data type: PtypString, 0x001F.");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                OXCDATA,
                2700,
                @"[In Property Data Types] PtypString (PT_UNICODE, string) is that Variable size; a string of Unicode characters in UTF-16LE format encoding with terminating null character (0x0000). with Property Type Value  0x001F,%x1F.00.");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                1076,
                @"[In PidTagMemberName Property] Type: PtypString ([MS-OXCDATA] section 2.11.1)");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                OXCDATA,
                455,
                @"[In PropertyValue Structure] PropertyValue (variable): If the property value being passed is a string, the data includes the terminating null characters.");

            // The parser has ensured the field satisfied the format, otherwise can't get the response.
            Site.CaptureRequirement(
                OXPROPS,
                6918,
                @"[In PidTagMemberRights] Property ID: 0x6673.");

            // PtypInteger32 is 4 bytes, so if the length of memberRight is 4 bytes, the data type of PidTagMemberRights is PtypInteger32.
            List<byte[]> memberRights = this.GetMemberRights(permissionUserList);
            foreach (byte[] memberRight in memberRights)
            {
                // Verify MS-OXPROPS requirement: MS-OXPROPS_R6919
                Site.CaptureRequirementIfAreEqual<int>(
                    4,
                    memberRight.Length,
                    OXPROPS,
                    6919,
                    @"[In PidTagMemberRights] Data type: PtypInteger32, 0x0003.");

                // Verify MS-OXPROPS requirement: MS-OXCDATA_R2691
                Site.CaptureRequirementIfAreEqual<int>(
                    4,
                    memberRight.Length,
                    OXCDATA,
                    2691,
                    @"[In Property Data Types] PtypInteger32 (PT_LONG, PT_I4, int, ui4) is that 4 bytes; a 32-bit integer [MS-DTYP]: INT32 with Property Type Value 0x0003,%x03.00.");

                // Verify MS-OXCPERM requirement: MS-OXCPERM_R1082
                Site.CaptureRequirementIfAreEqual<int>(
                    4,
                    memberRight.Length,
                    1082,
                    @"[In PidTagMemberRights Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");
            }
        }

        #endregion

        /// <summary>
        /// Verify the server object handle to a Table object
        /// </summary>
        /// <param name="responseValue">The return value from the server</param>
        private void VerifyGetPermissionHandle(uint responseValue)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R2006");
        
            // Verify MS-OXCPERM requirement: MS-OXCPERM_R2006
            this.Site.CaptureRequirementIfAreEqual<uint>(
                UINT32SUCCESS,
                responseValue,
                2006,
                @"[In Processing a RopGetPermissionsTable ROP Request] The server MUST return a Server object handle to a Table object, which the client uses to retrieve the permissions list of the folder, as specified in section 3.1.4.1.");

            // That the server return Success indicates the server object handle to the Table object is returned successfully.
            Site.CaptureRequirementIfAreEqual<uint>(
                UINT32SUCCESS,
                responseValue,
                9,
                "[In RopGetPermissionsTable ROP] The RopGetPermissionsTable ROP ([MS-OXCROPS] section 2.2.10.2) retrieves a Server object handle to a Table object, which is then used in other ROP requests to retrieve the current permissions list on a folder.");
        }

        /// <summary>
        /// Verify the RopModifyPermissions ROP response buffer
        /// </summary>
        private void VerifyModifyPermissionsResponse()
        {
            // If the RopModifyPermissions ROP response buffer can be parsed successfully, it indicates the server responds with a RopModifyPermissions ROP response buffer.
            Site.CaptureRequirement(
                1170,
                @"[In Processing a RopModifyPermissions ROP Request] The server responds with a RopModifyPermissions ROP response buffer.");
        }

        #region Verify the right flags specified by PidTagMemberRights property
        /// <summary>
        /// Verify the Create flag value
        /// </summary>
        private void VerifyCreateFlagValue()
        {
            // The parser parse the Create flag in the PidTagMemberRights through the flag definition.
            Site.CaptureRequirement(
                139,
                @"[In PidTagMemberRights Property] Create: Its value is 0x00000002.");
        }

        /// <summary>
        /// Verify the CreateSubFolder flag value
        /// </summary>
        private void VerifyCreateSubFolderFlagValue()
        {
            // The parser parse the CreateSubFolder flag in the PidTagMemberRights through the flag definition.
            Site.CaptureRequirement(
                117,
                @"[In PidTagMemberRights Property] CreateSubFolder: Its value is 0x00000080.");
        }

        /// <summary>
        /// Verify the DeleteAny flag value
        /// </summary>
        private void VerifyDeleteAnyFlagValue()
        {
            // The parser parse the DeleteAny flag in the PidTagMemberRights through the flag definition.
            Site.CaptureRequirement(
                        120,
                        @"[In PidTagMemberRights Property] DeleteAny: Its value is 0x00000040.");
        }

        /// <summary>
        /// Verify the DeleteOwned flag value
        /// </summary>
        private void VerifyDeleteOwnedFlagValue()
        {
            // The parser parse the DeleteOwned flag in the PidTagMemberRights through the flag definition.
            Site.CaptureRequirement(
                        126,
                        @"[In PidTagMemberRights Property] DeleteOwned: Its value is 0x00000010.");
        }

        /// <summary>
        /// Verify the EditAny flag value
        /// </summary>
        private void VerifyEditAnyFlagValue()
        {
            // The parser parse the EditAny flag in the PidTagMemberRights through the flag definition.
            Site.CaptureRequirement(
                        123,
                        @"[In PidTagMemberRights Property] EditAny: Its value is 0x00000020.");
        }

        /// <summary>
        /// Verify the EditOwned flag value
        /// </summary>
        private void VerifyEditOwnedFlagValue()
        {
            // The parser parse the EditOwned flag in the PidTagMemberRights through the flag definition.
            Site.CaptureRequirement(
                        130,
                        @"[In PidTagMemberRights Property] EditOwned: Its value is 0x00000008.");
        }

        /// <summary>
        /// Verify the FolderOwner flag value
        /// </summary>
        private void VerifyFolderOwnerFlagValue()
        {
            // The parser parse the FolderOwner flag in the PidTagMemberRights through the flag definition.
            Site.CaptureRequirement(
                        114,
                        @"[In PidTagMemberRights Property] FolderOwner: Its value is 0x00000100.");
        }

        /// <summary>
        /// Verify the FolderVisible flag value
        /// </summary>
        private void VerifyFolderVisibleFlagValue()
        {
            // The parser parse the FolderVisible flag in the PidTagMemberRights through the flag definition.
            Site.CaptureRequirement(
                        102,
                        @"[In PidTagMemberRights Property] FolderVisible: Its value is 0x00000400.");
        }

        /// <summary>
        /// Verify the FreeBusyDetailed flag value
        /// </summary>
        private void VerifyFreeBusyDetailedFlagValue()
        {
            // The parser parse the FreeBusyDetailed flag in the PidTagMemberRights through the flag definition.
            Site.CaptureRequirement(
                        96,
                        @"[In PidTagMemberRights Property] FreeBusySimple: Its value is 0x00000800.");
        }

        /// <summary>
        /// Verify the FreeBusySimple flag value
        /// </summary>
        private void VerifyFreeBusySimpleFlagValue()
        {
            // The parser parse the FreeBusySimple flag in the PidTagMemberRights through the flag definition.
            Site.CaptureRequirement(
                        90,
                        @"[In PidTagMemberRights Property] FreeBusyDetailed: Its value is 0x00001000.");
        }

        /// <summary>
        /// Verify the ReadAny flag value
        /// </summary>
        private void VerifyReadAnyFlagValue()
        {
            // The parser parse the ReadAny flag in the PidTagMemberRights through the flag definition.
            Site.CaptureRequirement(
                        142,
                        @"[In PidTagMemberRights Property] ReadAny: Its value is 0x00000001.");
        }

        #endregion

        /// <summary>
        /// Check the permissions list contains the user's permissions.
        /// </summary>
        /// <param name="user">The user name</param>
        /// <param name="permissionsList">User permission list</param> 
        /// <returns>True, if the permissions list contains the user's permissions, else false.</returns>
        private bool CheckUserPermissionContained(string user, List<PermissionUserInfo> permissionsList)
        {
            if (permissionsList.Count > 0)
            {
                foreach (PermissionUserInfo pui in permissionsList)
                {
                    if (string.Equals(pui.PidTagMemberName, user, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Verify the credential is provided to log on the server.
        /// </summary>
        /// <param name="user">The name of the user.</param>
        /// <param name="password">The password of the user.</param>
        /// <param name="logonReturnValue">The return value of the RopLogon.</param>
        private void VerifyCredential(string user, string password, uint logonReturnValue)
        {
            bool userIsEmpty = string.IsNullOrEmpty(user);
            bool passwordIsEmpty = string.IsNullOrEmpty(password);
            Site.Assert.IsFalse(userIsEmpty, "False indicates the user name is not null or empty. The user name:{0}", user);
            Site.Assert.IsFalse(passwordIsEmpty, "False indicates the password is not null or empty. The password:{0}", password);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R2005");
        
            // Verify MS-OXCPERM requirement: MS-OXCPERM_R2005
            this.Site.CaptureRequirementIfAreEqual<uint>(
                UINT32SUCCESS,
                logonReturnValue,
                2005,
                @"[In Accessing a Folder] Anonymous user permissions: the server requires that the client provide user credentials.");
        }
    }
}