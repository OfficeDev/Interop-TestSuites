namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    using System;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Capture code for MS-OXCFOLD adapter.
    /// </summary>
    public partial class MS_OXCFOLDAdapter : ManagedAdapterBase
    {
        #region Verify the requirements about transport

        /// <summary>
        /// Verify MAPIHTTP transport.
        /// </summary>
        private void VerifyTransport()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http" && Common.IsRequirementEnabled(1340, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1340");

                // Verify requirement MS-OXCFOLD_R1340
                // If the transport sequence is MAPIHTTP and the code can reach here, it means that the implementation does support MAPIHTTP transport.
                Site.CaptureRequirement(
                    1340,
                    @"[In Appendix A: Product Behavior] Implementation does support this specification [MS-OXCMAPIHTTP]. (Exchange Server 2013 Service Pack 1 (SP1) follows this behavior.)");
            }
            else if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "ncacn_ip_tcp" && Common.IsRequirementEnabled(99999, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R99999");

                // Verify requirement MS-OXCFOLD_R99999
                // If the transport sequence is ncacn_ip_tcp and the code can reach here, it means that the implementation does support ncacn_ip_tcp transport.
                Site.CaptureRequirement(
                    99999,
                    @"[In Appendix B: Product Behavior] Implementation does support this given protocol sequence [ncacn_ip_tcp]. ( Microsoft Exchange Server 2007 and 2010 follow this behavior.)");
            }
        }
        #endregion

        /// <summary>
        /// Verify RPC layer requirement
        /// </summary>
        private void VerifyRPCLayerRequirement()
        {
            // Since the request and response can be parsed correctly, the following requirement can be captured directly.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R2");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R2
            Site.CaptureRequirement(
                2,
                @"[In Transport] The ROP request buffers and ROP response buffers specified by this protocol [MS-OXCFOLD] are sent to and received by the server by using the underlying Remote Operations (ROP) List and Encoding Protocol, as specified in [MS-OXCROPS].");
        }

        #region Message Syntax

        #region Verify RopOpenFolder
        /// <summary>
        /// Verify the response of RopOpenFolder ROP operation.
        /// </summary>
        /// <param name="openFolderResponse">The response of RopOpenFolder operation</param>
        private void VerifyRopOpenFolder(RopOpenFolderResponse openFolderResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R24");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R24
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                openFolderResponse.ReturnValue,
                24,
                @"[In RopOpenFolder ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");
            
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R9");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R9.
            // The server returns a successful RopOpenFolder response, it indicates that the server opened an existing folder, MS-OXCFOLD_R9 can be captured directly.
            Site.CaptureRequirement(
                9,
                @"[In RopOpenFolder ROP] The RopOpenFolder ROP ([MS-OXCROPS] section 2.2.4.1) opens an existing folder.");

            if (openFolderResponse.IsGhosted != 0)
            {
                Site.Assert.IsNotNull(openFolderResponse.ServerCount, "[In RopOpenFolder Rop response] The ServerCount field should be present when the IsGhosted field is set to a nonzero (TRUE) value.");
                Site.Assert.IsNotNull(openFolderResponse.CheapServerCount, "[In RopOpenFolder Rop response] The CheapServerCount field should be present when the IsGhosted field is set to a nonzero (TRUE) value.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R911, the cheap server count is {0}, the server count is {1}.", openFolderResponse.CheapServerCount, openFolderResponse.ServerCount);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R911
                Site.CaptureRequirementIfIsTrue(
                    openFolderResponse.CheapServerCount <= openFolderResponse.ServerCount,
                    911,
                    @"[In RopOpenFolder ROP Response Buffer] The value of this field [CheapServerCount] MUST be less than or equal to the value of the ServerCount field.");
                
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R914, the count of strings contained in field Servers is {0}, the ServerCount is {1}.", openFolderResponse.Servers.Length, openFolderResponse.ServerCount);

                bool isVerifiedR914 = false;
                if (openFolderResponse.Servers == null)
                {
                    isVerifiedR914 = openFolderResponse.ServerCount == 0;
                }
                else
                {
                    isVerifiedR914 = openFolderResponse.ServerCount == (ushort)openFolderResponse.Servers.Length;
                }

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R914
                Site.CaptureRequirementIfIsTrue(
                    isVerifiedR914,
                    914,
                    @"[In RopOpenFolder ROP Response Buffer] The number of strings contained in this field [Servers] is specified by the ServerCount field.");
            }
        }
        #endregion

        #region Verfiy RopCreateFolder
        /// <summary>
        /// Verify the response of RopCreateFolder ROP operation. 
        /// </summary>
        /// <param name="createFolderResponse">The response of RopCreateFolder operation</param>
        private void VerifyRopCreateFolder(RopCreateFolderResponse createFolderResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R58");

            // If the RopCreateFolder operation returns successfully, it indicates that the server creates a new folder.
            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R58
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                createFolderResponse.ReturnValue,
                58,
                @"[RopCreateFolder ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R37");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R37
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                createFolderResponse.ReturnValue,
                37,
                @"[In RopCreateFolder ROP] The RopCreateFolder ROP ([MS-OXCROPS] section 2.2.4.2) creates a new folder.");

            if (createFolderResponse.IsExistingFolder != 0)
            {
                Site.Assert.IsNotNull(createFolderResponse.IsGhosted, "[In RopCreateFolder Rop response] The IsGhosted field should be present when the IsExistingFolder field is set to a nonzero (TRUE) value.");

                if (createFolderResponse.IsGhosted != 0)
                {
                    Site.Assert.IsNotNull(createFolderResponse.ServerCount, "[In RopCreateFolder Rop response] The ServerCount field should be present when the IsGhosted field is set to a nonzero (TRUE) value.");
                    Site.Assert.IsNotNull(createFolderResponse.CheapServerCount, "[In RopCreateFolder Rop response] The CheapServerCount field should be present when the IsGhosted field is set to a nonzero (TRUE) value.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R936");
                    
                    // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R936
                    Site.CaptureRequirementIfIsTrue(
                        createFolderResponse.CheapServerCount <= createFolderResponse.ServerCount,
                        936,
                        @"[In RopCreateFolder ROP Response Buffer] The value of this field [CheapServerCount] MUST be less than or equal to the value of the ServerCount field.");

                    if (createFolderResponse.ServerCount > 0)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R937");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R937
                        Site.CaptureRequirementIfIsTrue(
                            createFolderResponse.CheapServerCount > 0,
                            937,
                            @"[In RopCreateFolder ROP Response Buffer] And [the value of this field ""CheapServerCount""] MUST be greater than zero when the value of the ServerCount field is greater than zero.");
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R940, the number of strings contained in field Servers is {0}, the ServerCount is {1}.", createFolderResponse.Servers.Length, createFolderResponse.ServerCount);

                    bool isVerifiedR940 = false;
                    if (createFolderResponse.Servers == null)
                    {
                        isVerifiedR940 = createFolderResponse.ServerCount == 0;
                    }
                    else
                    {
                        isVerifiedR940 = createFolderResponse.ServerCount == (ushort)createFolderResponse.Servers.Length;
                    }

                    // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R940.
                    Site.CaptureRequirementIfIsTrue(
                        isVerifiedR940,
                        940,
                        @"[In RopCreateFolder ROP Response Buffer] The number of strings contained in this field [Servers] is specified by the ServerCount field.");
                }
            }
        }
        #endregion

        #region Verfiy RopDeleteFolder
        /// <summary>
        /// Verify the response of RopDeleteFolder ROP operation.
        /// </summary>
        /// <param name="deleteFolderResponse">The response of RopDeleteFolder operation</param>
        private void VerifyRopDeleteFolder(RopDeleteFolderResponse deleteFolderResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R95");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R95
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                deleteFolderResponse.ReturnValue,
                95,
                @"[In RopDeleteFolder ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R75");

            // If the RopDeleteFolder operation returns successfully, it indicates that the server removed a folder.
            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R75
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                deleteFolderResponse.ReturnValue,
                75,
                @"[In RopDeleteFolder ROP] The RopDeleteFolder ROP ([MS-OXCROPS] section 2.2.4.3) removes a folder.");
        }
        #endregion

        #region Verfiy RopSetSearchCriteria
        /// <summary>
        /// Verify the response of RopSetSearchCriteria ROP operation.
        /// </summary>
        /// <param name="setSearchCriteriaResponse">The response of RopSetSearchCriteria operation</param>
        private void VerifyRopSetSearchCriteria(RopSetSearchCriteriaResponse setSearchCriteriaResponse)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R119");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R119.
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                setSearchCriteriaResponse.ReturnValue,
                119,
                @"[In RopSetSearchCriteria ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");
        }
        #endregion

        #region Verfiy RopGetSearchCriteria
        /// <summary>
        /// Verify the response of RopGetSearchCriteria ROP operation.
        /// </summary>
        /// <param name="getSearchCriteriaResponse">The response of RopGetSearchCriteria operation</param>
        private void VerifyRopGetSearchCriteria(RopGetSearchCriteriaResponse getSearchCriteriaResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R134");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R134
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                getSearchCriteriaResponse.ReturnValue,
                134,
                @"[In RopGetSearchCriteria ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, @"Verify MS-OXCFOLD_R120");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R120
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                getSearchCriteriaResponse.ReturnValue,
                120,
                @"[In RopGetSearchCriteria ROP] The RopGetSearchCriteria ROP ([MS-OXCROPS] section 2.2.4.5) obtains the search criteria and the status of a search for a search folder.");

            if (getSearchCriteriaResponse.RestrictionData != null)
            {
                RestrictsFactory.Deserialize(getSearchCriteriaResponse.RestrictionData);

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R137");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R137.
                // The RestrictionData in RopGetSearchCriteria ROP response were deserialized successfully, MS-OXCFOLD_R137 can be verified directly.
                Site.CaptureRequirement(
                    137,
                    @"[In RopGetSearchCriteria ROP Response Buffer] RestrictionData (variable): A packet of structures that specify restrictions for the search folder.");

                byte restrictions = getSearchCriteriaResponse.RestrictionData[0];
                bool isVerifiedR510 = restrictions == 0
                    || restrictions == (byte)0x01
                    || restrictions == (byte)0x02
                    || restrictions == (byte)0x03
                    || restrictions == (byte)0x04
                    || restrictions == (byte)0x05
                    || restrictions == (byte)0x06
                    || restrictions == (byte)0x07
                    || restrictions == (byte)0x08
                    || restrictions == (byte)0x09
                    || restrictions == (byte)0x0A
                    || restrictions == (byte)0x0B;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R510,  RestrictType expected value is: 0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A and 0x0B, actual value is {0}", restrictions.ToString());

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R510
                Site.CaptureRequirementIfIsTrue(
                    isVerifiedR510,
                    Constants.MSOXCDATA,
                    510,
                    @"[In Restrictions] Although the packet formats differ, the first 8 bits always store RestrictType, an unsigned byte value specifying the type of restriction.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R958");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R958
                Site.CaptureRequirementIfAreEqual<int>(
                    getSearchCriteriaResponse.RestrictionData.Length,
                    getSearchCriteriaResponse.RestrictionDataSize,
                    958,
                    @"[In RopGetSearchCriteria ROP Response Buffer] The size of this field [RestrictionData] is specified by the RestrictionDataSize field.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R957");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R957
                Site.CaptureRequirementIfAreEqual<int>(
                    getSearchCriteriaResponse.RestrictionData.Length,
                    getSearchCriteriaResponse.RestrictionDataSize,
                    957,
                    @"[In RopGetSearchCriteria ROP Response Buffer] RestrictionDataSize (2 bytes): An integer that specifies the size, in bytes, of the RestrictionData field.");
            }

            if (getSearchCriteriaResponse.FolderIds != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R138");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R138
                Site.CaptureRequirementIfAreEqual<int>(
                    getSearchCriteriaResponse.FolderIds.Length,
                    getSearchCriteriaResponse.FolderIdCount,
                    138,
                    @"[InRopGetSearchCriteria ROP Response Buffer] FolderIdCount (2 bytes): An integer that specifies the number of structures contained in the FolderIds field.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R2140");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R2140
                Site.CaptureRequirementIfAreEqual<int>(
                    getSearchCriteriaResponse.FolderIdCount,
                    getSearchCriteriaResponse.FolderIds.Length,
                    2140,
                    @"[InRopGetSearchCriteria ROP Response Buffer] The number of structures contained in the array is specified by the value of the FolderIdCount field. ");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R140");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R140
                // The RopGetSearchCriteriaResponse ROP response was parsed successfully and the FolderIdCount field was verified, MS-OXCFOLD_R140 can be captured directly.
                Site.CaptureRequirement(
                    140,
                    "[In RopGetSearchCriteria ROP Response Buffer] FolderIds (variable): An array of FID structures ([MS-OXCDATA] section 2.2.1.1), each of which specifies a folder that is being searched.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R121");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R121
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                getSearchCriteriaResponse.ReturnValue,
                121,
                @"[In RopGetSearchCriteria ROP] The search criteria are created by using RopSetSearchCriteria (section 2.2.1.4).");
        }
        #endregion

        #region Verfiy RopMoveCopyMessages
        /// <summary>
        /// Verify the response of RopMoveCopyMessages ROP operation.
        /// </summary>
        /// <param name="moveCopyMessagesResponse">The response of RopMoveCopyMessages operation</param>
        private void VerifyRopMoveCopyMessages(RopMoveCopyMessagesResponse moveCopyMessagesResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R161");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R161
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                moveCopyMessagesResponse.ReturnValue,
                161,
                @"[In RopMoveCopyMessages ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R144");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R144
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                moveCopyMessagesResponse.ReturnValue,
                144,
                @"[In RopMoveCopyMessages ROP] The RopMoveCopyMessages ROP ([MS-OXCROPS] section 2.2.4.6) moves or copies messages from a source folder to a destination folder.");
        }
        #endregion

        #region Verfiy RopMoveFolder
        /// <summary>
        /// Verify the response of RopMoveFolder ROP operation.
        /// </summary>
        /// <param name="moveFolderResponse">The response package of RopMoveFolder operation</param>
        private void VerifyRopMoveFolder(RopMoveFolderResponse moveFolderResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R188");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R188
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                moveFolderResponse.ReturnValue,
                188,
                @"[In RopMoveFolder ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R169");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R169
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                moveFolderResponse.ReturnValue,
                169,
                @"[In RopMoveFolder ROP] The RopMoveFolder ROP ([MS-OXCROPS] section 2.2.4.7) moves a folder from one parent folder to another parent folder.");
        }
        #endregion

        #region Verfiy RopCopyFolder
        /// <summary>
        /// Verify the response of RopCopyFolder ROP operation.
        /// </summary>
        /// <param name="copyFolderResponse">The response of RopCopyFolder operation</param>
        private void VerifyRopCopyFolder(RopCopyFolderResponse copyFolderResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R217");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R217
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                copyFolderResponse.ReturnValue,
                217,
                @"[In RopCopyFolder ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R196");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R196
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                copyFolderResponse.ReturnValue,
                196,
                @"[In RopCopyFolder ROP] The RopCopyFolder ROP ([MS-OXCROPS] section 2.2.4.8) copies a folder from one parent folder to another parent folder.");
        }
        #endregion

        #region Verfiy RopEmptyFolder
        /// <summary>
        /// Verify the response of RopEmptyFolder ROP operation.
        /// </summary>
        /// <param name="emptyFolderResponse">The response of RopEmptyFolder operation</param>
        private void VerifyRopEmptyFolder(RopEmptyFolderResponse emptyFolderResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R241");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R241
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                emptyFolderResponse.ReturnValue,
                241,
                @"[In RopEmptyFolder ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");
        }
        #endregion

        #region Verfiy RopHardDeleteMessagesAndSubfolders
        /// <summary>
        /// Verify the response of RopHardDeleteMessagesAndSubfolders ROP operation.
        /// </summary>
        /// <param name="hardDeleteMessagesAndSubfoldersResponse">The response of RopHardDeleteMessagesAndSubfolders operation</param>
        private void VerifyRopHardDeleteMessagesAndSubfolders(RopHardDeleteMessagesAndSubfoldersResponse hardDeleteMessagesAndSubfoldersResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R259");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R259
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                hardDeleteMessagesAndSubfoldersResponse.ReturnValue,
                259,
                @"[In RopHardDeleteMessagesAndSubfolders ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R244");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R244
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                hardDeleteMessagesAndSubfoldersResponse.ReturnValue,
                244,
                @"[In RopHardDeleteMessagesAndSubfolders ROP] The RopHardDeleteMessagesAndSubfolders ROP ([MS-OXCROPS] section 2.2.4.10) is used to hard delete all messages and subfolders from a folder without deleting the folder itself.");
        }
        #endregion

        #region Verfiy RopDeleteMessages
        /// <summary>
        /// Verify the response of RopDeleteMessages ROP operation.
        /// </summary>
        /// <param name="deleteMessagesResponse">The response of RopDeleteMessages operation</param>
        private void VerifyRopDeleteMessages(RopDeleteMessagesResponse deleteMessagesResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R280");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R280
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                deleteMessagesResponse.ReturnValue,
                280,
                @"[In RopDeleteMessages ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");
        }
        #endregion

        #region Verfiy RopHardDeleteMessages
        /// <summary>
        /// Verify the response of RopHardDeleteMessages ROP operation.
        /// </summary>
        /// <param name="hardDeleteMessages">The response of RopHardDeleteMessages operation</param>
        private void VerifyRopHardDeleteMessages(RopHardDeleteMessagesResponse hardDeleteMessages)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R302");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R302
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                hardDeleteMessages.ReturnValue,
                302,
                @"[In RopHardDeleteMessages ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R283");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R283
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                hardDeleteMessages.ReturnValue,
                283,
                @"[In RopHardDeleteMessages ROP] The RopHardDeleteMessages ROP ([MS-OXCROPS] section 2.2.4.12) is used to hard delete one or more messages from a folder.");
        }
        #endregion

        #region Verfiy RopGetHierarchyTable
        /// <summary>
        /// Verify the response of RopGetHierarchyTable ROP operation. 
        /// </summary>
        /// <param name="getHierarchyTableResponse">The response of RopGetHierarchyTable operation </param>
        private void VerifyRopGetHierarchyTable(RopGetHierarchyTableResponse getHierarchyTableResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R317");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R317
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                getHierarchyTableResponse.ReturnValue,
                317,
                @"[In RopGetHierarchyTable ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R305");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R305
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                getHierarchyTableResponse.ReturnValue,
                305,
                @"[In RopGetHierarchyTable ROP] The RopGetHierarchyTable ROP ([MS-OXCROPS] section 2.2.4.13) is used to retrieve the hierarchy table for a folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R306");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R306
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                getHierarchyTableResponse.ReturnValue,
                306,
                @"[In RopGetHierarchyTable ROP] This ROP [RopGetHierarchyTable] returns a Table object on which table operations can be performed.");
        }
        #endregion

        #region Verfiy RopGetContentsTable
        /// <summary>
        /// Verify the response of RopGetContentsTable ROP operation. 
        /// </summary>
        /// <param name="getContentsTableResponse">The response of RopGetContentsTable operation</param>
        private void VerifyRopGetContentsTable(RopGetContentsTableResponse getContentsTableResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R332");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R332
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                getContentsTableResponse.ReturnValue,
                332,
                @"[In RopGetContentsTable ROP Response Buffer] ReturnValue (4 bytes): The server returns 0x00000000 to indicate success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R320");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R320
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                getContentsTableResponse.ReturnValue,
                320,
                @"[In RopGetContentsTable ROP] The RopGetContentsTable ROP ([MS-OXCROPS] section 2.2.4.14) is used to retrieve the contents table for a folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R321");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R321
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                getContentsTableResponse.ReturnValue,
                321,
                @"[In RopGetContentsTable ROP] This ROP [RopGetContentsTable] returns a Table object on which table operations can be performed.");
        }
        #endregion

        #endregion

        #region Folder properties validation.

        /// <summary>
        /// Verify the response of RopGetPropertiesAll ROP operation.
        /// </summary>
        /// <param name="response">The response of RopGetPropertiesAll operation.</param>
        private void VerifyGetFolderPropertiesAll(RopGetPropertiesAllResponse response)
        {
            foreach (TaggedPropertyValue taggedPropertyValue in response.PropertyValues)
            {
                PropertyTag tag = taggedPropertyValue.PropertyTag;
                switch (taggedPropertyValue.PropertyTag.PropertyId)
                {
                    case (ushort)FolderPropertyId.PidTagChangeKey:

                        #region PidTagChangeKey
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5687");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5687
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypBinary,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            5687,
                            "[In PidTagChangeKey] Data type: PtypBinary, 0x0102.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5686");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5686
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            5686,
                            "[In PidTagChangeKey] Property ID: 0x65E2.");

                        this.VerifyPtypBinary(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagContentUnreadCount:

                        #region PidTagContentUnreadCount
                        if (tag.PropertyType != (ushort)PropertyType.PtypErrorCode)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5841");

                            // Verify MS-OXPROPS requirement: MS-OXPROPS_R5841
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                (ushort)PropertyType.PtypInteger32,
                                tag.PropertyType,
                                Constants.MSOXPROPS,
                                5841,
                                "[In PidTagContentUnreadCount] Data type: PtypInteger32, 0x0003.");

                            this.VerifyPtypInteger32(taggedPropertyValue.Value);
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R347");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R347.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            347,
                            "[In PidTagContentUnreadCount Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5840");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5840
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            5840,
                            "[In PidTagContentUnreadCount] Property ID: 0x3603.");
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagHierarchyChangeNumber:

                        #region PidTagHierarchyChangeNumber
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6350");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6350
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypInteger32,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            6350,
                            "[In PidTagHierarchyChangeNumber] Data type: PtypInteger32, 0x0003.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1028");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1028.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            1028,
                            "[In PidTagHierarchyChangeNumber Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6349");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6349
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            6349,
                            "[In PidTagHierarchyChangeNumber] Property ID: 0x663E.");

                        this.VerifyPtypInteger32(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagAccess:

                        #region PidTagAccess
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R4928");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R4928
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypInteger32,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            4928,
                            "[In PidTagAccess] Data type: PtypInteger32, 0x0003.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R4927");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R4927
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            4927,
                            "[In PidTagAccess] Property ID: 0x0FF4.");

                        this.VerifyPtypInteger32(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagAddressBookEntryId:

                        #region PidTagAddressBookEntryId
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5024");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5024
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypBinary,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            5024,
                            "[In PidTagAddressBookEntryId] Data type: PtypBinary, 0x0102.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R349");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R349.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            349,
                            "[In PidTagAddressBookEntryId Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5023");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5023
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            5023,
                            "[In PidTagAddressBookEntryId] Property ID: 0x663B.");

                        this.VerifyPtypBinary(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagAttributeHidden:

                        #region PidTagAttributeHidden
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5556");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5556
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypBoolean,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            5556,
                            "[In PidTagAttributeHidden] Data type: PtypBoolean, 0x000B.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R356");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R356.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            356,
                            "[In PidTagAttributeHidden Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5555");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5555
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            5555,
                            "[In PidTagAttributeHidden] Property ID: 0x10F4.");

                        this.VerifyPtypBoolean(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagComment:

                        #region PidTagComment
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5753");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5753
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypString,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            5753,
                            "[In PidTagComment] Data type: PtypString, 0x001F.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R359");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R359.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            359,
                            "[In PidTagComment Property] Type: PtypString ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5752");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5752
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            5752,
                            "[In PidTagComment] Property ID: 0x3004.");

                        this.VerifyPtypString(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagContentCount:

                        #region PidTagContentCount
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5821");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5821
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypInteger32,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            5821,
                            "[In PidTagContentCount] Data type: PtypInteger32, 0x0003.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R345");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R345.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            345,
                            "[In PidTagContentCount Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5820");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5820.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            5820,
                            "[In PidTagContentCount] Property ID: 0x3602.");

                        this.VerifyPtypInteger32(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagCreationTime:

                        #region PidTagCreationTime
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5881");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5881
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypTime,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            5881,
                            "[In PidTagCreationTime] Data type: PtypTime, 0x0040.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5880");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5880
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            5880,
                            "[In PidTagCreationTime] Property ID: 0x3007.");

                        this.VerifyPtypTime(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagDisplayName:

                        #region PidTagDisplayName
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6015");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6015
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypString,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            6015,
                            "[In PidTagDisplayName] Data type: PtypString, 0x001F.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R360");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R360.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            360,
                            "[In PidTagDisplayName Property] Type: PtypString ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6014");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6014
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            6014,
                            "[In PidTagDisplayName] Property ID: 0x3001.");

                        this.VerifyPtypString(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagFolderId:

                        #region PidTagFolderId
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6213");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6213
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypInteger64,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            6213,
                            "[In PidTagFolderId] Data type: PtypInteger64, 0x0014.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R351");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R351.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            351,
                            "[In PidTagFolderId Property] Type: PtypInteger64 ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6212");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6212
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            6212,
                            "[In PidTagFolderId] Property ID: 0x6748.");

                        this.VerifyPtypInteger64(taggedPropertyValue.Value);
                        this.VerifyFolderIDStructure(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagFolderType:

                        #region PidTagFolderType
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6220");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6220
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypInteger32,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            6220,
                            "[In PidTagFolderType] Data type: PtypInteger32, 0x0003.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R362");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R362.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            362,
                            "[In PidTagFolderType Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6219");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6219
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            6219,
                            "[In PidTagFolderType] Property ID: 0x3601.");

                        this.VerifyPtypInteger32(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagLastModificationTime:

                        #region PidTagLastModificationTime
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6784");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6784
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypTime,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            6784,
                            "[In PidTagLastModificationTime] Data type: PtypTime, 0x0040.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6783");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6783
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            6783,
                            "[In PidTagLastModificationTime] Property ID: 0x3008.");

                        this.VerifyPtypTime(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagRights:

                        #region PidTagRights
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7982");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R7982
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypInteger32,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            7982,
                            "[In PidTagRights] Data type: PtypInteger32, 0x0003.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R367");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R367.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            367,
                            "[In PidTagRights Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7981");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R7981
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            7981,
                            "[In PidTagRights] Property ID: 0x6639.");

                        this.VerifyPtypInteger32(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagSubfolders:

                        #region PidTagSubfolders
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R8691");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R8691
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypBoolean,
                            tag.PropertyType,
                            Constants.MSOXPROPS,
                            8691,
                            "[In PidTagSubfolders] Data type: PtypBoolean, 0x000B.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R355");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R355.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            355,
                            "[In PidTagSubfolders Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R8690");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R8690
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            8690,
                            "[In PidTagSubfolders] Property ID: 0x360A.");

                        this.VerifyPtypBoolean(taggedPropertyValue.Value);
                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagMessageSize:

                        #region PidTagMessageSize
                        if (tag.PropertyType == (ushort)PropertyType.PtypInteger32)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7002");

                            // Verify MS-OXPROPS requirement: MS-OXPROPS_R7002
                            // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                Constants.MSOXPROPS,
                                7002,
                                "[In PidTagMessageSize] Data type: PtypInteger32, 0x0003.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R353");

                            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R353.
                            // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                353,
                                "[In PidTagMessageSize Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7001");

                            // Verify MS-OXPROPS requirement: MS-OXPROPS_R7001
                            // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                            Site.CaptureRequirement(
                                Constants.MSOXPROPS,
                                7001,
                                "[In PidTagMessageSize] Property ID: 0x0E08.");

                            this.VerifyPtypInteger32(taggedPropertyValue.Value);
                        }

                        break;
                        #endregion

                    case (ushort)FolderPropertyId.PidTagAccessControlListData:

                        #region PidTagAccessControlListData

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1042");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1042.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypBinary,
                            tag.PropertyType,
                            1042,
                            "[In PidTagAccessControlListData Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R10102");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R10102.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            10102,
                            "[In PidTagAccessControlListData] Property ID: 0x3FE0");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R10103");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R10103.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            10103,
                            "[In PidTagAccessControlListData] Data type: PtypBinary, 0x0102");

                        this.VerifyPtypBinary(taggedPropertyValue.Value);
                        #endregion
                        break;

                    case (ushort)FolderPropertyId.PidTagLocalCommitTime:

                        #region PidTagLocalCommitTime

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R3000");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R3000.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypTime,
                            tag.PropertyType,
                            3000,
                            "[In PidTagLocalCommitTime Property] Type: PtypTime ([MS-OXCDATA] section 2.11.1).");
                        #endregion
                        break;

                    case (ushort)FolderPropertyId.PidTagLocalCommitTimeMax:

                        #region PidTagLocalCommitTimeMax

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R3002");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R3002.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypTime,
                            tag.PropertyType,
                            3002,
                            "[In PidTagLocalCommitTimeMax Property] Type: PtypTime ([MS-OXCDATA] section 2.11.1).");
                        #endregion
                        break;

                    case (ushort)FolderPropertyId.PidTagDeletedCountTotal:

                        #region PidTagDeletedCountTotal

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R3005");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R3005.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypInteger32,
                            tag.PropertyType,
                            3005,
                            "[In PidTagDeletedCountTotal Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1).");
                        #endregion
                        break;
                    case (ushort)FolderPropertyId.PidTagFolderFlags:

                        #region PidTagFolderFlags
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R35101");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R35101
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypInteger32,
                            tag.PropertyType,
                            35101,
                            "[In PidTagFolderFlags Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");
                        break;
                        #endregion
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Verify the specific property.
        /// </summary>
        /// <param name="propertyTags">PropertyTag array.</param>
        private void VerifyGetFolderSpecificProperties(PropertyTag[] propertyTags)
        {
            foreach (PropertyTag propertyTag in propertyTags)
            {
                switch (propertyTag.PropertyId)
                {
                    case (ushort)FolderPropertyId.PidTagDeletedOn:
                        #region PidTagDeletedOn

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R348");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R348.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypTime,
                            propertyTag.PropertyType,
                            348,
                            "[In PidTagDeletedOn Property] Type: PtypTime ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5979");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5979.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            5979,
                            "[In PidTagDeletedOn] Property ID: 0x668F.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5980");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5980.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            5980,
                            "[In PidTagDeletedOn] Data type: PtypTime, 0x0040.");
                        #endregion
                        break;
                    case (ushort)FolderPropertyId.PidTagMessageSizeExtended:
                        #region PidTagMessageSizeExtended

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7008");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R7008.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            7008,
                            "[In PidTagMessageSizeExtended] Property ID: 0x0E08.");

                        #endregion
                        break;
                    case (ushort)FolderPropertyId.PidTagContainerClass:
                        #region PidTagContainerClass

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1036");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1036.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypString,
                            propertyTag.PropertyType,
                            1036,
                            "[In PidTagContainerClass Property] Type: PtypString ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5792");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5792.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            5792,
                            "[In PidTagContainerClass] Property ID: 0x3613.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5793");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5793.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            5793,
                            "[In PidTagContainerClass] Data type: PtypString, 0x001F.");
                        #endregion
                        break;
                    case (ushort)FolderPropertyId.PidTagParentEntryId:
                        #region PidTagParentEntryId

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1027");

                        // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1027.
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            (ushort)PropertyType.PtypBinary,
                            propertyTag.PropertyType,
                            1027,
                            "[In PidTagParentEntryId Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1)");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7456");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R7456.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            7456,
                            "[In PidTagParentEntryId] Property ID: 0x0E09.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7457");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R7457.
                        // The property ID and property type match the description in [MS-OXPROPS] and the property type was verified, this requirement can be captured directly.
                        Site.CaptureRequirement(
                            Constants.MSOXPROPS,
                            7457,
                            "[In PidTagParentEntryId] Data type: PtypBinary, 0x0102.");
                        #endregion
                        break;
                    default:
                        break;
                }
            }
        }
        #endregion

        #region Data type validation.
        /// <summary>
        /// Verify the Folder ID Structure.
        /// </summary>
        /// <param name="folderIdStructure">The Folder ID Structure instance in bytes.</param>
        private void VerifyFolderIDStructure(byte[] folderIdStructure)
        {
            // Add debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2175.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2175
            Site.CaptureRequirementIfAreEqual<int>(
                8,
                folderIdStructure.Length,
                Constants.MSOXCDATA,
                2175,
                @"[In Folder ID Structure] It [Folder ID] is an 8-byte structure.");
        }

        /// <summary>
        /// Verify the type of PtypInteger32.
        /// </summary>
        /// <param name="ptypInteger32">The PtypInteger32 instance in bytes.</param>
        private void VerifyPtypInteger32(byte[] ptypInteger32)
        {
            // Add debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2691.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2691
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                ptypInteger32.Length,
                Constants.MSOXCDATA,
                2691,
                @"[In Property Data Types] PtypInteger32 (PT_LONG, PT_I4, int, ui4) is that 4 bytes; a 32-bit integer [MS-DTYP]: INT32 with Property Type Value 0x0003,%x03.00.");
        }

        /// <summary>
        /// Verify the type of PtypBoolean
        /// </summary>
        /// <param name="ptypBoolean">The PtypBoolean instance in bytes.</param>
        private void VerifyPtypBoolean(byte[] ptypBoolean)
        {
            // Add debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2698.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2698
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                ptypBoolean.Length,
                Constants.MSOXCDATA,
                2698,
                @"[In Property Data Types] PtypBoolean (PT_BOOLEAN. bool) is that 1 byte, restricted to 1 or 0 [MS-DTYP]: BOOLEAN with Property Type Value 0x000B, %x0B.00.");
        }

        /// <summary>
        /// Verify the type of PtypInteger64
        /// </summary>
        /// <param name="ptypInteger64">The PtypInteger64 instance in bytes.</param>
        private void VerifyPtypInteger64(byte[] ptypInteger64)
        {
            // Add debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2699.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2699
            Site.CaptureRequirementIfAreEqual<int>(        
                8,
                ptypInteger64.Length,
                Constants.MSOXCDATA,
                2699,
                @"[In Property Data Types] PtypInteger64 (PT_LONGLONG, PT_I8, i8, ui8) is that 8 bytes; a 64-bit integer [MS-DTYP]: LONGLONG with Property Type Value 0x0014,%x14.00.");
        }

        /// <summary>
        /// Verify the type of PtypTime
        /// </summary>
        /// <param name="ptypTime">The PtypTime instance in bytes.</param>
        private void VerifyPtypTime(byte[] ptypTime)
        {
            // Add debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2702.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2702
            Site.CaptureRequirementIfAreEqual<int>(     
                8,     
                ptypTime.Length,
                Constants.MSOXCDATA,
                2702,
                @"[In Property Data Types] PtypTime (PT_SYSTIME, time, datetime, datetime.tz, datetime.rfc1123, Date, time, time.tz) is that 8 bytes; a 64-bit integer representing the number of 100-nanosecond intervals since January 1, 1601 [MS-DTYP]: FILETIME with Property Type Value 0x0040,%x40.00.");
        }

        /// <summary>
        /// Verify the type of PtypBinary
        /// </summary>
        /// <param name="ptypBinary">The PtypBinary instance in bytes.</param>
        private void VerifyPtypBinary(byte[] ptypBinary)
        {
            byte[] length = new byte[2];
            Array.Copy(ptypBinary, 0, length, 0, 2);
            short len = BitConverter.ToInt16(length, 0);
            short followedBytesCount = (short)(ptypBinary.Length - 2);

            // Add debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2707.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2707
            Site.CaptureRequirementIfAreEqual<short>(
                len,
                followedBytesCount,
                Constants.MSOXCDATA,
                2707,
                @"[In Property Data Types] PtypBinary (PT_BINARY) is that variable size; a COUNT field followed by that many bytes with Property Type Value 0x0102,%x02.01.");
        }

        /// <summary>
        /// Verify if the bytes is PtypString type
        /// </summary>
        /// <param name="bytes">The PtypString instance in bytes.</param>
        private void VerifyPtypString(byte[] bytes)
        {
            byte[] length = new byte[2];
            Array.Copy(bytes, bytes.Length - 2, length, 0, 1);
            short len = BitConverter.ToInt16(length, 0);

            // Add debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2700.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2700
            Site.CaptureRequirementIfAreEqual<short>(
                0,
                len,
                Constants.MSOXCDATA,
                2700,
                @"[In Property Data Types] PtypString (PT_UNICODE, string) is that Variable size; a string of Unicode characters in UTF-16LE format encoding with terminating null character (0x0000). with Property Type Value  0x001F,%x1F.00.");
        }
        #endregion
    }
}