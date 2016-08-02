namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Runtime.InteropServices;
    using Microsoft.Modeling;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides the methods to write capture code.
    /// </summary>
    public partial class MS_OXCFXICSAdapter : ManagedAdapterBase, IMS_OXCFXICSAdapter
    {
        #region MAPIHTTP transport

        /// <summary>
        /// Verify MAPIHTTP transport.
        /// </summary>
        private void VerifyMAPITransport()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http" && Common.IsRequirementEnabled(5100, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R5100");

                // Verify requirement MS-MS-OXCFXICS_R5100
                // If the transport sequence is MAPIHTTP and the code can reach here, it means that the implementation does support MAPIHTTP transport.
                Site.CaptureRequirement(
                        5100,
                        @"[In Appendix A: Product Behavior] Implementation does support this specification [MS-OXCMAPIHTTP]. (Exchange Server 2013 Service Pack 1 (SP1) follows this behavior.)");
            }
        }
        #endregion

        /// <summary>
        /// Verify the FolderReplicaInfo structure.
        /// </summary>
        /// <param name="folderContent">A structure.</param>
        private void VerifyFolderReplicaInfoStructure(FolderContent folderContent)
        {
            if (folderContent.NewFXFolderList.Count != 0)
            {
                foreach (Tuple<List<uint>, FolderReplicaInfo> tuple in folderContent.NewFXFolderList)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3081");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3081
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x00000000,
                        tuple.Item2.Flags,
                        3081,
                        @"[In FolderReplicaInfo Structure] Flags (4 bytes): This value MUST be set to 0x00000000.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3082");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3082
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x00000000,
                        tuple.Item2.Depth,
                        3082,
                        @"[In FolderReplicaInfo Structure] Depth (4 bytes): This value MUST be set to 0x00000000.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3083");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3083
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(LongTermId),
                        tuple.Item2.FolderLongTermId.GetType(),
                        3083,
                        @"[In FolderReplicaInfo Structure] FolderLongTermId (24 bytes): A LongTermID structure ([MS-OXCDATA] section 2.2.1.3.1).");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3084");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3084
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(uint),
                        tuple.Item2.ServerDNCount.GetType(),
                        3084,
                        @"[In FolderReplicaInfo Structure] ServerDNCount (4 bytes): An unsigned integer value.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2189");

                    // Verify MS-OXCFXICS requirement: MS-OXCDATA_R2189
                    this.Site.CaptureRequirementIfAreEqual<int>(
                        24,
                        tuple.Item2.FolderLongTermId.Size(),
                        "MS-OXCDATA",
                        2189,
                        @"[In LongTermID Structure] The total length of the LongTermID structure is 24 bytes.");

                    // The structure LongTermID deserialize according to the specification, so this requirement can be verified directly.
                    this.Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2188,
                        @"[In LongTermID Structure] A LongTermID structure is a Global Identifier structure, as specified in section 2.2.1.3, plus a 2-byte Pad field that has the value 0x0000.");

                    // If the program has executed here, this requirement can be verified directly.
                    this.Site.CaptureRequirement(
                        273,
                        @"[In FolderReplicaInfo Structure] [ServerDNCount (4 bytes)] determines how many elements exist in the ServerDNArray field.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3085");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3085
                    this.Site.CaptureRequirementIfAreNotEqual<uint>(
                        0x00000000,
                        tuple.Item2.ServerDNCount,
                        3085,
                        @"[In FolderReplicaInfo Structure] [ServerDNCount (4 bytes)] This value MUST NOT be 0x00000000.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3086");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3086
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(uint),
                        tuple.Item2.CheapServerDNCount.GetType(),
                        3086,
                        @"[In FolderReplicaInfo Structure] CheapServerDNCount (4 bytes): An unsigned integer value.");

                    // If the program has executed here, this requirement can be verified directly.
                    this.Site.CaptureRequirement(
                        276,
                        @"[In FolderReplicaInfo Structure] [CheapServerDNCount (4 bytes)] determines how many of the leading elements in the ServerDNArray field have the same, lowest, network access cost.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3087");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3087
                    bool isVerifiedR3087 = tuple.Item2.CheapServerDNCount <= tuple.Item2.ServerDNCount;

                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR3087,
                        3087,
                        @"[In FolderReplicaInfo Structure] The value of the CheapServerDNCount field MUST be less than or equal to value of the ServerDNCount field.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3089");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3089
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        tuple.Item2.ServerDNCount,
                        (uint)tuple.Item2.ServerDNArray.Length,
                        3089,
                        @"[In FolderReplicaInfo Structure] The number of ServerDNCount strings in this field is specified by the ServerDNCount field.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3088");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3088
                    bool isVerifiedR3088 = true;

                    // Since the folderContent.NewFXFolderList.Count must not be null when precondition met, so if any string in ServerDNArray does not end with null, the requirement will not be captured, else it will be captured.
                    foreach (string serverDN in tuple.Item2.ServerDNArray)
                    {
                        if (!serverDN.EndsWith("\0"))
                        {
                            isVerifiedR3088 = false;
                            break;
                        }
                    }

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3088
                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR3088,
                        3088,
                        @"[In FolderReplicaInfo Structure] ServerDNArray (variable): An array of ASCII-encoded NULL-terminated strings. ");
                }

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1037
                // When NewFXFolderList is not null, which means the MetaTagNewFXFolder has parsed correctly according to the property ID and data type. 
                // So this requirement can be captured directly.
                this.Site.CaptureRequirement(
                        1037,
                    @"[In MetaTagNewFXFolder Meta-Property] Property ID: 0x4011 Data type: PtypBinary, 0x0102 ([MS-OXCDATA] section 2.11.1)");
            }
        }

        /// <summary>
        /// Verify the Folder change number related requirements.
        /// </summary>
        /// <param name="propTag">Property Tag</param>
        /// <param name="propValue">Property Value</param>
        private void VerifyPidTagChangeNumberValueRelated(string propTag, PropertyValue propValue)
        {
            uint propValues;

            if (propTag == "PidTagDeletedCountTotal")
            {
                propValues = Common.ConvertByteArrayToUint(propValue.Value);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R6004");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R6004
                Site.CaptureRequirementIfAreEqual<uint>(
                            1,
                            propValues,
                            6004,
                            @"[In Identifying Objects and Maintaining Change Numbers] [Also, if a message is deleted within the folder, the value of the folder change number does not change] But the aggregated PidTagDeletedCountTotal property ([MS-OXCFOLD] section 2.2.2.2.1.15) is updated to reflect the change.");
            }
            else if (propTag == "PidTagChangeNumber" && Common.IsRequirementEnabled(6003, this.Site))
            {
                propValues = Common.ConvertByteArrayToUint(propValue.Value);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R6003");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R6003
                Site.CaptureRequirementIfAreEqual<uint>(
                            0x8004010f,
                            propValues,
                            6003,
                            @"[In Identifying Objects and Maintaining Change Numbers] Also, if a message is deleted within the folder, the 
                            value of the folder change number does not change.");
            }

            if (propTag == "PidTagLocalCommitTimeMax")
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R6002");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R6002
                Site.CaptureRequirementIfIsNotNull(
                           propValue.Value,
                           6002,
                           @"[In Identifying Objects and Maintaining Change Numbers] [If a message within a folder changes, the change number is not updated] However, the aggregated property PidTagLocalCommitTimeMax property ([MS-OXCFOLD] section 2.2.2.2.1.14) is modified to reflect that something within the folder has been changed.");
            }
            else if (propTag == "PidTagChangeNumber" && Common.IsRequirementEnabled(6001, this.Site))
            {
                propValues = Common.ConvertByteArrayToUint(propValue.Value);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R6001");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R6001
                Site.CaptureRequirementIfAreEqual<uint>(
                            0x8004010f,
                            propValues,
                            6001,
                            @"[In Identifying Objects and Maintaining Change Numbers] If a message within a folder
                             changes, the change number is not updated.");
            }
        }

        /// <summary>
        /// Verify whether the read status changed.
        /// </summary>
        /// <param name="propValue">The value of the property.</param>
        private void VerifySyncImportReadStateChanged(PropertyValue propValue)
        {
            if (Common.IsRequirementEnabled(3513001, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3513001");

                // If the code reached here, which means the RopSynchronizationImportReadStateChanges failed in previous step and the read state doesn't been changed, 
                // which is same as the last step: mfRead(0x00000001) and mfUnsent(0x00000008).
                bool isVerifiedR3513001 = Common.ConvertByteArrayToUint(propValue.Value) == 0x09;

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3513001
                Site.CaptureRequirementIfIsTrue(
                            isVerifiedR3513001,
                            3513001,
                            @"[In Receiving a RopSynchronizationImportReadStateChanges Request] Implementation does support fail the ROP before performing any read state changes, as opposed to partially completing the ROP if a read state change will not succeed. (Exchange 2007 and above products follow this behavior.)");
            }
        }

        /// <summary>
        /// Verify Rop FastTransferDestinationPutBuffer.
        /// </summary>
        /// <param name="fastTransferDestinationPutBufferRequest">The request was sent to server.</param>
        /// <param name="response">The response was gotten from server.</param>
        private void VerifyRopFastTransferDestinationPutBuffer(RopFastTransferDestinationPutBufferRequest fastTransferDestinationPutBufferRequest, object response)
        {
            RopFastTransferDestinationPutBufferResponse fastTransferDestinationPutBufferResponse = (RopFastTransferDestinationPutBufferResponse)response;

            if (Common.IsRequirementEnabled(3178001, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3178001");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3178001
                this.Site.CaptureRequirementIfAreEqual<ushort>(
                    0x0001,
                    fastTransferDestinationPutBufferResponse.TotalStepCount,
                    3178001,
                    @"[In Appendix A: Product Behavior] Implementation does set this value to 0x0000. [<11> Section 2.2.3.1.2.2.2: Exchange 2010, Exchange 2013 and Exchange 2016 set the value of the TotalStepCount field to 0x0001.]");
            }
            else if (Common.IsRequirementEnabled(317800201, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R317800201");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3178002001
                this.Site.CaptureRequirementIfAreEqual<ushort>(
                    0x0000,
                    fastTransferDestinationPutBufferResponse.TotalStepCount,
                    317800201,
                    @"[In Appendix A: Product Behavior] Implementation does set this value to 0x0001. [<11> Section 2.2.3.1.2.2.2: Exchange 2007 set the value of the TotalStepCount field to 0x0000.]");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2156");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2156
            // When RopId is equal to 0x54 which means the ROP is RopFastTransferDestinationPutBuffer, so this requirement can be captured.
            bool isVerifiedR2156 = fastTransferDestinationPutBufferResponse.RopId == 0x54;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR2156,
                2156,
                @"[In Receiving a RopFastTransferDestinationPutBuffer Request] The server MUST respond with a RopFastTransferDestinationPutBuffer ROP response, as specified in [MS-OXCROPS] section 2.2.12.2.2 and section 2.2.3.1.2.2 of this specification.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3171");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3171
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                fastTransferDestinationPutBufferResponse.ReturnValue.GetType(),
                3171,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3172");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3172
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                fastTransferDestinationPutBufferResponse.TransferStatus.GetType(),
                3172,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer] TransferStatus (2 bytes): A 16-bit enumeration.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3173");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3173
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                fastTransferDestinationPutBufferResponse.InProgressCount.GetType(),
                3173,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer] InProgressCount (2 bytes): An unsigned 16-bit integer. ");

            if (Common.IsRequirementEnabled(31701, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R31701");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R31701
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0x0000,
                    fastTransferDestinationPutBufferResponse.InProgressCount,
                    31701,
                    @"[In Appendix A: Product Behavior] Implementation does set the InProgressCount field to 0x0000 in RopFastTransferDestinationPutBuffer ROP Response Buffer.  (Section 2.2.3.1.2.2.2: Microsoft Exchange Server 2007 and above follow this behavior.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3176");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3176
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                fastTransferDestinationPutBufferResponse.TotalStepCount.GetType(),
                3176,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer] TotalStepCount (2 bytes): An unsigned 16-bit integer. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3179");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3179
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                fastTransferDestinationPutBufferResponse.Reserved,
                3179,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer] Reserved (1 byte): The field MUST be set to 0x00 when sending.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3181");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3181
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                fastTransferDestinationPutBufferResponse.BufferUsedSize.GetType(),
                3181,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer] BufferUsedSize (2 bytes): An unsigned 16-bit integer. ");

            if (fastTransferDestinationPutBufferResponse.ReturnValue == (uint)RopResult.Success)
            {
                if (Common.IsRequirementEnabled(318200101, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R318200101");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R318200101
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        fastTransferDestinationPutBufferRequest.TransferDataSize,
                        fastTransferDestinationPutBufferResponse.BufferUsedSize,
                        318200101,
                        @"[In Appendix A: Product Behavior] Implementation does return a value for the BufferSizeUsed field that is equal to the value of the TransferDataSize field. ( <12> Section 2.2.3.1.2.2.2: Exchange 2007, Exchange 2010, and Exchange 2013 always return a value for the BufferSizeUsed field that is equal to the value of the TransferDataSize field, regardless of whether the value of the ReturnValue field is Success (0x00000000).)");
                }
            }
        }

        /// <summary>
        /// Verify Rop FastTransferSourceCopyTo
        /// </summary>
        /// <param name="fastTransferSourceCopyToRequest">The request was sent to server.</param>
        /// <param name="fastTransferSourceCopyToResponse">The response was gotten from server.</param>
        private void VerifyRopFastTransferSourceCopyTo(RopFastTransferSourceCopyToRequest fastTransferSourceCopyToRequest, RopFastTransferSourceCopyToResponse fastTransferSourceCopyToResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R382");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R382
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                fastTransferSourceCopyToResponse.ReturnValue.GetType(),
                382,
                @"[In RopFastTransferSourceCopyTo ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            if (fastTransferSourceCopyToResponse.ReturnValue == 0x00000000)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R385");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R385
                // When deserialize the response get from the server,
                // if outputServerObject not present it will be set to 0xff
                bool isVerifyR385 = this.responseSOHs[fastTransferSourceCopyToResponse.OutputHandleIndex] != 0xff;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR385,
                    385,
                    @"[In RopFastTransferSourceCopyTo ROP Response Buffer] [OutputServerObject] This value MUST be present if and only if the value of the ReturnValue field equals Success (0x00000000).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2090");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2090
                // RopFastTransferSourceCopyToresponse.RopId of this Rop is 0x4D according [MS-OXCROPS]
                // if the requirements above are verified and this RopId == 0x4D, this requirement can be captured
                bool isVerifyR2090 = fastTransferSourceCopyToResponse.RopId == 0x4D;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2090,
                    2090,
                    @"[In Receiving a RopFastTransferSourceCopyTo  ROP Request]The server MUST respond with a RopFastTransferSourceCopyTo ROP response, as 
                    specified in [MS-OXCROPS] section 2.2.13.6.2 and section 2.2.3.1.1.1 of this specification.");
            }
            else
            {
                if (fastTransferSourceCopyToRequest.CopyFlags == (uint)CopyToCopyFlags.Invalid)
                {
                    if (Common.IsRequirementEnabled(3445, this.Site))
                    {
                        // If the program has executed here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3445,
                            @"[In Receiving a RopFastTransferSourceCopyTo Request] Implementation does fail the ROP if unknown CopyFlags flag bits are set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                    }
                }

                if (fastTransferSourceCopyToRequest.SendOptions == (uint)SendOptionAlls.Invalid)
                {
                    if (Common.IsRequirementEnabled(3463, this.Site))
                    {
                        // If the program has executed here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3463,
                            @"[In Receiving a RopFastTransferSourceCopyTo Request] Implementation does fail the ROP in receiving a RopFastTransferSourceCopyTo request if any unknown SendOptions flag bits are set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                    }
                }
            }
        }

        /// <summary>
        /// Verify Rop FastTransferSourceCopyProperties
        /// </summary>
        /// <param name="fastTransferSourceCopyPropertiesRequest">The request was sent to server.</param>
        /// <param name="fastTransferSourceCopyPropertiesResponse">The response was gotten from server.</param>
        private void VerifyRopFastTransferSourceCopyProperties(RopFastTransferSourceCopyPropertiesRequest fastTransferSourceCopyPropertiesRequest, RopFastTransferSourceCopyPropertiesResponse fastTransferSourceCopyPropertiesResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R453");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R453
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                fastTransferSourceCopyPropertiesResponse.ReturnValue.GetType(),
                453,
                @"[In RopFastTransferSourceCopyProperties ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            if (fastTransferSourceCopyPropertiesResponse.ReturnValue == 0x00000000)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R456");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R456
                // When deserialize the response get from the server,
                // if outputServerObject not present it will be set to 0xff
                bool isVerifyR456 = this.responseSOHs[fastTransferSourceCopyPropertiesResponse.OutputHandleIndex] != 0xFF;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR456,
                    456,
                    @"[In RopFastTransferSourceCopyProperties ROP Response Buffer] [OutputServerObject] This value MUST be present if and only if the value of the ReturnValue field equals Success (0x00000000).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2103");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2103
                // RopFastTransferSourceCopyPropertiesresponse.RopId of this Rop is 0x69 according [MS-OXCROPS]
                // if the requirements above are verified and this RopId == 0x69, this requirement can be captured
                bool isVerifyR2103 = fastTransferSourceCopyPropertiesResponse.RopId == 0x69;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2103,
                    2103,
                    @"[In Receiving a RopFastTransferSourceCopyProperties Request] The server MUST respond with a RopFastTransferSourceCopyProperties ROP response, as specified in [MS-OXCROPS] section 2.2.12.7.2 and section 2.2.3.1.1.2 of this specification.");
            }
            else
            {
                if (fastTransferSourceCopyPropertiesRequest.CopyFlags == (byte)CopyPropertiesCopyFlags.Invalid)
                {
                    if (Common.IsRequirementEnabled(3466, this.Site))
                    {
                        // If the program has executed here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3466,
                            @"[In Receiving a RopFastTransferSourceCopyProperties Request] Implementation does fail the ROP in receiving a RopFastTransferSourceCopyProperties request if any unknown CopyFlag flag bits are set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                    }
                }

                if (fastTransferSourceCopyPropertiesRequest.SendOptions == (uint)SendOptionAlls.Invalid)
                {
                    if (Common.IsRequirementEnabled(3470, this.Site))
                    {
                        // If the program has executed here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3470,
                            @"[In Receiving a RopFastTransferSourceCopyProperties Request] Implementation does fail the ROP in receiving a RopFastTransferSourceCopyProperties request if any unknown SendOptions flag bits are set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                    }
                }
            }
        }

        /// <summary>
        /// Verify Rop FastTransferSourceCopyMessages
        /// </summary>
        /// <param name="fastTransferSourceCopyMessagesRequest">The request was sent to server.</param>
        /// <param name="fastTransferSourceCopyMessagesResponse">The response was gotten from server.</param>
        private void VerifyRopFastTransferSourceCopyMessages(RopFastTransferSourceCopyMessagesRequest fastTransferSourceCopyMessagesRequest, RopFastTransferSourceCopyMessagesResponse fastTransferSourceCopyMessagesResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R485");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R485
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                fastTransferSourceCopyMessagesResponse.ReturnValue.GetType(),
                485,
                @"[In RopFastTransferSourceCopyMessages ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            if (fastTransferSourceCopyMessagesResponse.ReturnValue == 0x00000000)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R488");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R488
                // When deserialize the response get from the server,
                // if outputServerObject not present it will be set to 0xff
                bool isVerifyR488 = this.responseSOHs[fastTransferSourceCopyMessagesResponse.OutputHandleIndex] != 0xff;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR488,
                    488,
                    @"[In RopFastTransferSourceCopyMessages ROP Response Buffer] [OutputServerObject] This value MUST be present if and only if the value of the ReturnValue field equals Success (0x00000000).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2113");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2113
                // RopFastTransferSourceCopyMessagesresponse.RopId of this Rop is 0x4B according [MS-OXCROPS]
                // if the requirements above are verified and this RopId == 0x4B, this requirement can be captured
                bool isVerifyR2113 = fastTransferSourceCopyMessagesResponse.RopId == 0x4B;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2113,
                    2113,
                    @"[In Receiving a RopFastTransferSourceCopyMessages ROP Request] The server MUST respond with a RopFastTransferSourceCopyMessages ROP response, as specified in [MS-OXCROPS] section 2.2.12.5.2 and section 2.2.3.1.1.3 of this specification.");
            }
            else
            {
                if (fastTransferSourceCopyMessagesRequest.SendOptions == (byte)SendOptionAlls.Invalid)
                {
                    if (Common.IsRequirementEnabled(3479, this.Site))
                    {
                        // If the program has executed here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3479,
                            @"[In Receiving a RopFastTransferSourceCopyMessages ROP Request] Implementation does fail the ROP in receiving a RopFastTransferSourceCopyMessage request if any unknown SendOptions flag bits are set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                    }
                }
            }
        }

        /// <summary>
        /// Verify Rop FastTransferSourceCopyFolder
        /// </summary>
        /// <param name="fastTransferSourceCopyFolderRequest">The request was sent to server.</param>
        /// <param name="fastTransferSourceCopyFolderResponse">The response was gotten from server.</param>
        private void VerifyRopFastTransferSourceCopyFolder(RopFastTransferSourceCopyFolderRequest fastTransferSourceCopyFolderRequest, RopFastTransferSourceCopyFolderResponse fastTransferSourceCopyFolderResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R509");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R509
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                fastTransferSourceCopyFolderResponse.ReturnValue.GetType(),
                509,
                @"[In RopFastTransferSourceCopyFolder ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            if (fastTransferSourceCopyFolderResponse.ReturnValue == 0x00000000)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R512");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R512
                // When deserialize the response get from the server,
                // if outputServerObject not present it will be set to 0xff
                bool isVerifyR512 = this.responseSOHs[fastTransferSourceCopyFolderResponse.OutputHandleIndex] != 0xff;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR512,
                    512,
                    @"[In RopFastTransferSourceCopyFolder ROP Response Buffer] [OutputServerObject] This value MUST be present if and only if the value of the ReturnValue field equals Success (0x00000000).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2122");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2122
                // RopFastTransferSourceCopyFolderresponse.RopId of this Rop is 0x4C according [MS-OXCROPS]
                // if the requirements above are verified and this RopId == 0x4C, this requirement can be captured
                bool isVerifyR2122 = fastTransferSourceCopyFolderResponse.RopId == 0x4C;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2122,
                    2122,
                    @"[In Receiving a RopFastTransferSourceCopyFolder ROP Request] The server MUST respond with a RopFastTransferSourceCopyFolder ROP response, as specified in [MS-OXCROPS] section 2.2.12.4.2 and section 2.2.3.1.1.4 of this specification.");
            }
            else
            {
                if (fastTransferSourceCopyFolderRequest.CopyFlags == (byte)CopyFolderCopyFlags.Invalid)
                {
                    if (Common.IsRequirementEnabled(3483, this.Site))
                    {
                        // If the program has executed here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3483,
                            @"[In Receiving a RopFastTransferSourceCopyFolder ROP Request] Implementation does fail the ROP in receiving a RopFastTransferSourceCopyFolder request if any unknown CopyFlags flag bits are set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                    }
                }

                if (fastTransferSourceCopyFolderRequest.SendOptions == (byte)SendOptionAlls.Invalid)
                {
                    if (Common.IsRequirementEnabled(3487, this.Site))
                    {
                        // If the program has executed here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3487,
                            @"[In Receiving a RopFastTransferSourceCopyFolder ROP Request] Implementation does fail the ROP in receiving a RopFastTransferSourceCopyFolder request if any unknown SendOptions flag bits are set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                    }
                }
            }
        }

        /// <summary>
        /// Verify Rop FastTransferSourceGetBuffer
        /// </summary>
        /// <param name="fastTransferSourceGetBufferResponse">The response was gotten from server.</param>
        private void VerifyRopFastTransferSourceGetBuffer(RopFastTransferSourceGetBufferResponse fastTransferSourceGetBufferResponse)
        {
            // If the program can execute here, this requirement can be verified, because the server must has serialized each portion in specified format.
            this.Site.CaptureRequirement(
                3490,
                @"[In Receiving a RopFastTransferSourceGetBuffer Request] The server MUST serialize each portion of the FastTransfer stream using the syntax specified in section 2.2.4.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3491");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3491
            // If the TransferBuffer field is not null, it can explain that this field has been used.
            this.Site.CaptureRequirementIfIsNotNull(
                fastTransferSourceGetBufferResponse.TransferBuffer,
                3491,
                @"[In Receiving a RopFastTransferSourceGetBuffer Request] The server MUST output it[each portion of the FastTransfer stream] using the TransferBuffer field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R540");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R540
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                fastTransferSourceGetBufferResponse.ReturnValue.GetType(),
                540,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R543");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R543
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                fastTransferSourceGetBufferResponse.TransferStatus.GetType(),
                543,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer]TransferStatus (2 bytes): A 16-bit enumeration.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R545");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R545
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                fastTransferSourceGetBufferResponse.InProgressCount.GetType(),
                545,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer]InProgressCount (2 bytes): An unsigned 16-bit integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R548");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R548
            // "contains the approximate total number of steps to be completed in the current operation" is informative
            // needn't test it
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                fastTransferSourceGetBufferResponse.TotalStepCount.GetType(),
                548,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer] TotalStepCount (2 bytes): An unsigned 16-bit integer that contains the approximate total number of steps to be completed in the current operation.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R552");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R552
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                fastTransferSourceGetBufferResponse.TransferBufferSize.GetType(),
                552,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer]TransferBufferSize (2 bytes): An unsigned 16-bit integer.");

            // When the RopFastTransferSourceGetBuffer deserialize successfully indicate the structure described in the protocol is valid, then this requirement can be captured.
            this.Site.CaptureRequirement(
                3293,
                @"[In Syntactical Structure] The syntactical structure of a FastTransfer stream is as follows:
                root  = contentsSync
                  / hierarchySync
                  / state
                  / folderContent
                  / messageContent
                  / attachmentContent
                  / messageList
                  / topFolder

                propValue = <see lexical structure in 2.2.4.1>
                errorInfo = FXErrorInfo propList
                propList = *propValue

                subFolder = StartSubFld folderContent EndFolder
                subFolderNoDelProps = StartSubFld folderContentNoDelProps EndFolder
                topFolder = StartTopFld folderContentNoDelProps EndFolder
                folderContent = propList
                  ( MetaTagNewFXFolder / folderMessages ) 
                  [ MetaTagFXDelProp *subFolder ] 
                folderContentNoDelProps = propList
                  ( MetaTagNewFXFolder / folderMessagesNoDelProps ) 
                  [ *subFolderNoDelProps ] 
                folderMessages = *2( MetaTagFXDelProp messageList )
                folderMessagesNoDelProps = *2( messageList )
                message = ( StartMessage / StartFAIMsg ) 
                  messageContent 
                  EndMessage
                messageChildren = [ MetaTagFXDelProp ] [ *recipient ] 
                  [ MetaTagFXDelProp ] [ *attachment ] 
                messageContent = propList messageChildren
                messageList = *( [MetaTagEcWarning] message )
                recipient = StartRecip propList EndToRecip

                attachment = NewAttach PidTagAttachNumber attachmentContent EndAttach
                attachmentContent = propList [embeddedMessage] 
                embeddedMessage = StartEmbed messageContent EndEmbed

                contentsSync = [progressTotal] 
                  *( [progressPerMessage] messageChange )
                  [deletions] 
                  [readStateChanges] 
                  state
                  IncrSyncEnd
                hierarchySync = *folderChange 
                  [deletions] 
                  state 
                  IncrSyncEnd
                deletions = IncrSyncDel propList
                folderChange = IncrSyncChg propList
                groupInfo = IncrSyncGroupInfo propList
                messageChange = messageChangeFull / messageChangePartial
                messageChangeFull = IncrSyncChg messageChangeHeader 
                  IncrSyncMessage propList 
                  messageChildren
                messageChangeHeader = propList
                messageChangePartial = groupInfo MetaTagIncrSyncGroupId
                  IncrSyncChgPartial messageChangeHeader
                  *( MetaTagIncrementalSyncMessagePartial propList )
                  messageChildren
                progressPerMessage = IncrSyncProgressPerMsg propList
                progressTotal = IncrSyncProgressMode propList
                readStateChanges = IncrSyncRead propList
                state = IncrSyncStateBegin propList IncrSyncStateEnd
                ");

            if (Common.IsRequirementEnabled(3150001, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3150001");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3150001
                // The parameter isVerifiedR3150001 is true means the TransferStatus in Exchange 2010 and above do not contain NoRoom[0x0002].
                bool isVerifiedR3150001 = (fastTransferSourceGetBufferResponse.TransferStatus == 0x0000)
                                     || (fastTransferSourceGetBufferResponse.TransferStatus == 0x0001)
                                     || (fastTransferSourceGetBufferResponse.TransferStatus == 0x0003);

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR3150001,
                    3150001,
                    @"[In Appendix A: Product Behavior] Implementation does not support. [<8> Section 2.2.3.1.1.5.2: The NoRoom value is not returned by Exchange 2010 or Exchange 2013 or Exchange 2016.]");
            }
            else if (Common.IsRequirementEnabled(315000201, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R315000201");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3150002001
                bool isVerifiedR3150002001 = (fastTransferSourceGetBufferResponse.TransferStatus == 0x0000)
                                      || (fastTransferSourceGetBufferResponse.TransferStatus == 0x0001)
                                      || (fastTransferSourceGetBufferResponse.TransferStatus == 0x0003)
                                      || (fastTransferSourceGetBufferResponse.TransferStatus == 0x0002);

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR3150002001,
                    315000201,
                    @"[In Appendix A: Product Behavior] Implementation does support. [<8> Section 2.2.3.1.1.5.2: The NoRoom value is supported by Exchange 2007.]");
            }

            // If the requirement can execute here, this requirement can be verified directly.
            Site.CaptureRequirement(
                551,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer]Reserved (1 byte): The value of this field MUST be set to 0x00 when sending and ignored on receipt.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2136");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2136
            // RopFastTransferSourceGetBufferresponse.RopId of this Rop is 0x4E according [MS-OXCROPS]
            // if the requirements above are verified and this RopId == 0x4E, this requirement can be captured
            bool isVerifyR2136 = fastTransferSourceGetBufferResponse.RopId == 0x4E;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2136,
                2136,
                @"[In Receiving a RopFastTransferSourceGetBuffer Request] The server MUST respond with a RopFastTransferSourceGetBuffer ROP response, as specified in [MS-OXCROPS] section 2.2.12.6.2 and section 2.2.3.1.1.1 of this specification.");

            if (fastTransferSourceGetBufferResponse.ReturnValue == (uint)RopResult.Success)
            {
                if (Common.IsRequirementEnabled(5561, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R5561");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R5561
                    Site.CaptureRequirementIfIsNotNull(
                        fastTransferSourceGetBufferResponse.TransferBuffer,
                        5561,
                        @"[In RopFastTransferSourceGetBuffer ROP Response Buffer] Implementation does only have TransferBuffer field in RopFastTransferSourceGetBuffer ROP Response Buffer if the value of the ReturnValue field is Success (0x00000000). (Microsoft Exchange Server 2007 and above follow this behavior.)");
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R553");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R553
                Site.CaptureRequirementIfAreEqual<int>(
                    fastTransferSourceGetBufferResponse.TransferBuffer.Length,
                    fastTransferSourceGetBufferResponse.TransferBufferSize,
                    553,
                    @"[In RopFastTransferSourceGetBuffer ROP Response Buffer] [TransferBufferSize (2 bytes)] This value specifies the size of the TransferBuffer field.");
            }
        }

        /// <summary>
        /// Verify Rop TellVersion
        /// </summary>
        /// <param name="tellVersionResponse">The response was gotten from server.</param>
        private void VerifyRopTellVersion(RopTellVersionResponse tellVersionResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R577");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R577
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                tellVersionResponse.ReturnValue.GetType(),
                577,
                @"[In RopTellVersion ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2149");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2149
            // tellVersionResponse.RopId of this Rop is 0x86 according [MS-OXCROPS]
            // if the requirements above are verified and this RopId == 0x86, this requirement can be captured
            bool isVerifyR2149 = tellVersionResponse.RopId == 0x86;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2149,
                2149,
                @"[In  Receiving a RopTellVersion Request]The server MUST respond with a RopTellVersion ROP response, as specified in [MS-OXCROPS] section 
                2.2.12.8.2 and section 2.2.3.1.1.6 of this specification.");
        }

        /// <summary>
        /// Verify Rop FastTransferDestinationConfigure
        /// </summary>
        /// <param name="fastTransferDestinationConfigureResponse">The response was gotten from server.</param>
        private void VerifyRopFastTransferDestinationConfigure(RopFastTransferDestinationConfigureResponse fastTransferDestinationConfigureResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R588");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R588
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                fastTransferDestinationConfigureResponse.ReturnValue.GetType(),
                588,
                @"[In RopFastTransferDestinationConfigure ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R590");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R590
            // When deserialize the response get from the server,
            // if outputServerObject not present it will be set to 0xff
            bool isVerifyR590 = this.responseSOHs[fastTransferDestinationConfigureResponse.OutputHandleIndex] != 0xff;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR590,
                590,
                @"[In RopFastTransferDestinationConfigure ROP Response Buffer]OutputServerObject: The value of this field MUST be the FastTransfer upload context.");

            if (fastTransferDestinationConfigureResponse.ReturnValue == 0x00000000)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R591");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R591
                // not null means be present
                Site.CaptureRequirementIfIsNotNull(
                    fastTransferDestinationConfigureResponse.OutputHandleIndex,
                    591,
                    @"[In RopFastTransferDestinationConfigure ROP Response Buffer] [OutputServerObject] This field MUST be present if and only if the value of the ReturnValue field equals Success (0x00000000).");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2151");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2151
            // fastTransferDestinationConfigureResponse.RopId of this Rop is 0x53 according [MS-OXCROPS]
            // if the requirements above are verified and this RopId == 0x53, this requirement can be captured
            bool isVerifyR2151 = fastTransferDestinationConfigureResponse.RopId == 0x53;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2151,
                2151,
                @"[In Receiving a RopFastTransferDestinationConfigure Request]The server MUST respond with a RopFastTransferDestinationConfigure ROP response, 
                as specified in [MS-OXCROPS] section 2.2.12.1.2 and section 2.2.3.1.2.1 of this specification.");
        }

        /// <summary>
        /// Verify Rop SynchronizationConfigure
        /// </summary>
        /// <param name="synchronizationConfigureRequest">The request was sent from client</param>
        /// <param name="synchronizationConfigureResponse">The response was gotten from server.</param>
        private void VerifyRopSynchronizationConfigure(RopSynchronizationConfigureRequest synchronizationConfigureRequest, RopSynchronizationConfigureResponse synchronizationConfigureResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R667");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R667
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                synchronizationConfigureResponse.ReturnValue.GetType(),
                667,
                @"[In RopSynchronizationConfigure ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            if (synchronizationConfigureResponse.ReturnValue == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R670");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R670
                // When deserialize the response get from the server,
                // if outputServerObject not present it will be set to 0xff
                bool isVerifyR670 = this.responseSOHs[synchronizationConfigureResponse.OutputHandleIndex] != 0xff;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR670,
                    670,
                    @"[In RopSynchronizationConfigure ROP Response Buffer] [OutputServerObject] This value MUST be present if and only if ReturnValue is Success (0x00000000).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2158");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2158
                // synchronizationConfigureResponse.RopId of this Rop is 0x70 according [MS-OXCROPS]
                // if the requirements above are verified and this RopId == 0x70, this requirement can be captured
                bool isVerifyR2158 = synchronizationConfigureResponse.RopId == 0x70;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2158,
                    2158,
                    @"[In Receiving a RopSynchronizationConfigure ROP Request] The server MUST respond with a RopSynchronizationConfigure ROP response, as specified in [MS-OXCROPS] section 2.2.13.1.2 and section 2.2.3.2.1.1 of this specification.");

                if (synchronizationConfigureRequest.SendOptions == (byte)SendOptionAlls.UseCpid)
                {
                    if (Common.IsRequirementEnabled(34991, this.Site))
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R34991");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R34991
                        this.Site.CaptureRequirementIfAreEqual<uint>(
                            0x00000000,
                            synchronizationConfigureResponse.ReturnValue,
                            34991,
                            @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] Implementation does not fail the ROP if additional flags, used only in server-to-server communications are set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                    }
                }
            }
            else
            {
                if (synchronizationConfigureRequest.SendOptions == (byte)SendOptionAlls.Invalid)
                {
                    if (Common.IsRequirementEnabled(3498, this.Site))
                    {
                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3498,
                            @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] Implementation does fail the ROP [RopSynchronizationConfigure] if unknown flag bits are set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                    }
                }
            }
        }

        /// <summary>
        /// Verify Rop SynchronizationUploadStateStreamBegin
        /// </summary>
        /// <param name="synchronizationUploadStateStreamBeginResponse">The response was gotten from server.</param>
        private void VerifyRopSynchronizationUploadStateStreamBegin(RopSynchronizationUploadStateStreamBeginResponse synchronizationUploadStateStreamBeginResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R736");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R736
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                synchronizationUploadStateStreamBeginResponse.ReturnValue.GetType(),
                736,
                @"[In RopSynchronizationUploadStateStreamBegin ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2208");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2208
            // synchronizationUploadStateStreamBeginResponse.RopId of this Rop is 0x75 according [MS-OXCROPS]
            // if the requirements above are verified and this RopId == 0x75, this requirement can be captured
            bool isVerifyR2208 = synchronizationUploadStateStreamBeginResponse.RopId == 0x75;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2208,
                2208,
                @"[In Receiving a RopSynchronizationUploadStateStreamBegin Request]The server MUST respond with a RopSynchronizationUploadStateStreamBegin ROP 
                response, as specified in [MS-OXCROPS] section 2.2.13.9.2 and section 2.2.3.2.2.1 of this specification.");

            // Because when the SynchronizationUploadStateStreamBegin success indicate this requirement can be captured.
            Site.CaptureRequirement(
                729,
                @"[In RopSynchronizationUploadStateStreamBegin ROP] The RopSynchronizationUploadStateStreamBegin ROP ([MS-OXCROPS] section 2.2.13.9) initiates the upload of an ICS state property into the synchronization context.");
        }

        /// <summary>
        /// Verify Rop SynchronizationUploadStateStreamContinue
        /// </summary>
        /// <param name="synchronizationUploadStateStreamContinueResponse">The response was gotten from server.</param>
        private void VerifyRopSynchronizationUploadStateStreamContinue(RopSynchronizationUploadStateStreamContinueResponse synchronizationUploadStateStreamContinueResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R747");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R747
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                synchronizationUploadStateStreamContinueResponse.ReturnValue.GetType(),
                747,
                @"[In RopSynchronizationUploadStateStreamContinue ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2210");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2210
            // synchronizationUploadStateStreamContinueResponse.RopId of this Rop is 0x76 according [MS-OXCROPS]
            // if the requirements above are verified and this RopId == 0x76, this requirement can be captured
            bool isVerifyR2210 = synchronizationUploadStateStreamContinueResponse.RopId == 0x76;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2210,
                2210,
                @"[In Receiving a RopSynchronizationUploadStateStreamContinue Request]The server MUST respond with a RopSynchronizationUploadStateStreamContinue 
                ROP response, as specified in [MS-OXCROPS] section 2.2.13.10.2 and section 2.2.3.2.2.2 of this specification.");

            // Because successful RopSynchronizationUploadStateStreamContinue indicates this requirement can be captured.
            Site.CaptureRequirement(
                3194,
                @"[In RopSynchronizationUploadStateStreamContinue ROP] The RopSynchronizationUploadStateStreamContinue ROP ([MS-OXCROPS] section 2.2.13.10) continues to upload an ICS state property value into the synchronization context. ");
        }

        /// <summary>
        /// Verify Rop SynchronizationUploadStateStreamEnd
        /// </summary>
        /// <param name="synchronizationUploadStateStreamEndResponse">The response was gotten from server.</param>
        private void VerifyRopSynchronizationUploadStateStreamEnd(RopSynchronizationUploadStateStreamEndResponse synchronizationUploadStateStreamEndResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R754");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R754
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                synchronizationUploadStateStreamEndResponse.ReturnValue.GetType(),
                754,
                @"[In RopSynchronizationUploadStateStreamEnd ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2213");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2213
            // synchronizationUploadStateStreamEndResponse.RopId of this Rop is 0x77 according [MS-OXCROPS]
            // if the requirements above are verified and this RopId == 0x77, this requirement can be captured
            bool isVerifyR2213 = synchronizationUploadStateStreamEndResponse.RopId == 0x77;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2213,
                2213,
                @"[In Receiving a RopSynchronizationUploadStateStreamEnd Request]. The server MUST respond with a RopSynchronizationUploadStateStreamEnd ROP 
                response, as specified in [MS-OXCROPS] section 2.2.13.11.2 and section 2.2.3.2.2.3 of this specification.");

            // Since successful RopSynchronizationUploadStateStreamEnd response can be returned from server indicates this requirement can be captured.
            Site.CaptureRequirement(
                751,
                @"[In RopSynchronizationUploadStateStreamEnd ROP] The RopSynchronizationUploadStateStreamEnd ROP ([MS-OXCROPS] section 2.2.13.11) concludes the upload of an ICS state property value into the synchronization context.");
        }

        /// <summary>
        /// Verify Rop SynchronizationGetTransferState
        /// </summary>
        /// <param name="synchronizationGetTransferStateResponse">The response was gotten from server.</param>
        private void VerifyRopSynchronizationGetTransferState(RopSynchronizationGetTransferStateResponse synchronizationGetTransferStateResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R763");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R763
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                synchronizationGetTransferStateResponse.ReturnValue.GetType(),
                763,
                @"[In RopSynchronizationGetTransferState ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            if (synchronizationGetTransferStateResponse.ReturnValue == 0x00000000)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R766");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R766
                // When deserialize the response get from the server,
                // if outputServerObject not present it will be set to 0xff
                bool isVerifyR766 = this.responseSOHs[synchronizationGetTransferStateResponse.OutputHandleIndex] != 0xff;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR766,
                    766,
                    @"[In RopSynchronizationGetTransferState ROP Response Buffer] [OutputServerObject] This value MUST be present if and only if the value of the ReturnValue field equals Success (0x00000000).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2215");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2215
                // synchronizationGetTransferStateResponse.RopId of this Rop is 0x82 according [MS-OXCROPS]
                // if the requirements above are verified and this RopId == 0x82, this requirement can be captured
                bool isVerifyR2215 = synchronizationGetTransferStateResponse.RopId == 0x82;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2215,
                    2215,
                    @"[In Receiving a RopSynchronizationGetTransferState Request]The server MUST respond with a RopSynchronizationGetTransferState ROP response, 
                    as specified in [MS-OXCROPS] section 2.2.13.8.2 and section 2.2.3.2.3.1 of this specification.");
            }
        }

        /// <summary>
        /// Verify Rop SynchronizationOpenCollector
        /// </summary>
        /// <param name="synchronizationOpenCollectorResponse">The response was gotten from server.</param>
        private void VerifyRopSynchronizationOpenCollector(RopSynchronizationOpenCollectorResponse synchronizationOpenCollectorResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R776");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R776
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                synchronizationOpenCollectorResponse.ReturnValue.GetType(),
                776,
                @"[In RopSynchronizationOpenCollector ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            if (synchronizationOpenCollectorResponse.ReturnValue == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R779");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R779
                // When deserialize the response get from the server,
                // if outputServerObject not present it will be set to 0xff
                bool isVerifyR779 = this.responseSOHs[synchronizationOpenCollectorResponse.OutputHandleIndex] != 0xff;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR779,
                    779,
                    @"[In RopSynchronizationOpenCollector ROP Response Buffer] [OutputServerObject] This value MUST be present if and only if the value of the ReturnValue field equals Success.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2218");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2218
                // synchronizationOpenCollectorResponse.RopId of this Rop is 0x7E according [MS-OXCROPS]
                // if the requirements above are verified and this RopId == 0x7E, this requirement can be captured
                bool isVerifyR2218 = synchronizationOpenCollectorResponse.RopId == 0x7E;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2218,
                    2218,
                    @"[In Receiving a RopSynchronizationOpenCollector Request]The server MUST respond with a RopSynchronizationOpenCollector ROP response, as 
                    specified in [MS-OXCROPS] section 2.2.13.7.2 and section 2.2.3.2.4.1 of this specification.");
            }
        }

        /// <summary>
        /// Verify Rop SynchronizationImportMessageChange
        /// </summary>
        /// <param name="synchronizationImportMessageChangeResponse">The response was gotten from server.</param>
        private void VerifyRopSynchronizationImportMessageChange(RopSynchronizationImportMessageChangeResponse synchronizationImportMessageChangeResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R800");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R800
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                synchronizationImportMessageChangeResponse.ReturnValue.GetType(),
                800,
                @"[In RopSynchronizationImportMessageChange ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R807");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R807
            // ulong equal to 8 bytes
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ulong),
                synchronizationImportMessageChangeResponse.MessageId.GetType(),
                807,
                @"[In RopSynchronizationImportMessageChange ROP Response Buffer]MessageId (8 bytes): A 64-bit identifier that specifies the MID structure ([MS-OXCDATA] section 2.2.1.2) of the message that was imported.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R808");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R808
            Site.CaptureRequirementIfAreEqual<ulong?>(
                0x0000000000000000,
                synchronizationImportMessageChangeResponse.MessageId,
                808,
                @"[In RopSynchronizationImportMessageChange ROP Response Buffer] [ MessageId (8 bytes)] This value MUST be set to 0x0000000000000000.");

            #region For successful Response Buffer
            if (synchronizationImportMessageChangeResponse.ReturnValue == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R806");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R806
                // When deserialize the response get from the server,
                // if outputServerObject not present it will be set to 0xff
                bool isVerifyR806 = this.responseSOHs[synchronizationImportMessageChangeResponse.OutputHandleIndex] != 0xff;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR806,
                    806,
                    @"[In RopSynchronizationImportMessageChange ROP Response Buffer] [ OutputServerObject] This value MUST be present if and only if ReturnValue equals Success (0x00000000).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R809");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R809
                bool isVerifyR809 = synchronizationImportMessageChangeResponse.MessageId != null;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR809,
                    809,
                    @"[In RopSynchronizationImportMessageChange ROP Response Buffer] [ MessageId (8 bytes)] This value MUST be present if and only if ReturnValue equals Success (0x00000000).");
            }
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2220");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2220
            // synchronizationImportMessageChangeResponse.RopId of this Rop is 0x72 according [MS-OXCROPS]
            // if the requirements above are verified and this RopId == 0x72, this requirement can be captured
            bool isVerifyR2220 = synchronizationImportMessageChangeResponse.RopId == 0x72;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2220,
                2220,
                @"[In Receiving a RopSynchronizationImportMessageChange Request]The server MUST respond with a RopSynchronizationImportMessageChange ROP 
                response, as specified in [MS-OXCROPS] section 2.2.13.2.2 or 2.2.13.2.3, and in section 2.2.3.2.4.2 of this specification.");
        }

        /// <summary>
        /// Verify Rop SynchronizationImportHierarchyChange
        /// </summary>
        /// <param name="synchronizationImportHierarchyChangeResponse">The response was gotten from server.</param>
        private void VerifyRopSynchronizationImportHierarchyChange(RopSynchronizationImportHierarchyChangeResponse synchronizationImportHierarchyChangeResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R830");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R830
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                synchronizationImportHierarchyChangeResponse.ReturnValue.GetType(),
                830,
                @"[In RopSynchronizationImportHierarchyChange ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R833");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R833
            // ulong is 8 bytes
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ulong),
                synchronizationImportHierarchyChangeResponse.FolderId.GetType(),
                833,
                @"[In RopSynchronizationImportHierarchyChange ROP Response Buffer]FolderId (8 bytes): A 64-bit identifier.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R834");

            FolderId folderId = new FolderId((ulong)synchronizationImportHierarchyChangeResponse.FolderId);

            // Because the GlobalCounter of folderId is GLOBCNT struct, so this requirement can be captured.
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3059");

            Site.CaptureRequirementIfAreEqual(
                6,
                folderId.GlobalCounter.Length,
                3059,
                @"[In GLOBCNT Structure] A GLOBCNT structure is a 6-byte global namespace counter. ");

            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R834");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R834
            // This requirement can be captured since the FolderId can be parsed successfully, which means the FolderId is Folder ID structure, so this requirement can be captured directly.
            this.Site.CaptureRequirement(
                834,
                @"[In RopSynchronizationImportHierarchyChange ROP Response Buffer] [FolderId (8 bytes)] that contains the Folder ID structure ([MS-OXCDATA] section 2.2.1.1) of the folder that was imported.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R835");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R835
            Site.CaptureRequirementIfAreEqual<ulong>(
                0x0000000000000000,
                synchronizationImportHierarchyChangeResponse.FolderId,
                835,
                @"[In RopSynchronizationImportHierarchyChange ROP Response Buffer] [FolderId (8 bytes)] This value of this field MUST be set to 0x0000000000000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2175");

            // Verify MS-OXCFXICS requirement: MS-OXCDATA_R2175
            Site.CaptureRequirementIfAreEqual<ulong?>(
                0x0000000000000000,
                synchronizationImportHierarchyChangeResponse.FolderId,
                "MS-OXCDATA",
                2175,
                @"[In Folder ID Structure] It [Folder ID] is an 8-byte structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2176");

            // Verify MS-OXCFXICS requirement: MS-OXCDATA_R2176
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                folderId.ReplicaId.Length,
                "MS-OXCDATA",
                2176,
                @"[In Folder ID Structure] ReplicaId (2 bytes): An unsigned integer identifying a Store object.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2177");

            // Verify MS-OXCFXICS requirement: MS-OXCDATA_R2177
            Site.CaptureRequirementIfAreEqual<int>(
                6,
                folderId.GlobalCounter.Length,
                "MS-OXCDATA",
                2177,
                @"[In Folder ID Structure] GlobalCounter (6 bytes): An unsigned integer identifying the folder within its Store object.");

            #region For successful Response Buffer
            if (synchronizationImportHierarchyChangeResponse.ReturnValue == 0x00000000)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R836");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R836
                // When deserialize the response get from the server,
                // if outputServerObject not present it will be set to 0xff
                bool isVerifyR836 = synchronizationImportHierarchyChangeResponse.FolderId != 0xff;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR836,
                    836,
                    @"[In RopSynchronizationImportHierarchyChange ROP Response Buffer] [FolderId (8 bytes)] This value MUST be present if and only if the value of the ReturnValue field equals Success (0x00000000).");
            }

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2234");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2234
            // synchronizationImportHierarchyChangeResponse.RopId of this Rop is 0x73 according [MS-OXCROPS]
            // if the requirements above are verified and this RopId == 0x73, this requirement can be captured
            bool isVerifyR2234 = synchronizationImportHierarchyChangeResponse.RopId == 0x73;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2234,
                2234,
                @"[In Receiving a RopSynchronizationImportHierarchyChange Request]The server MUST respond with a RopSynchronizationImportHierarchyChange ROP 
                response, as specified in [MS-OXCROPS] section 2.2.13.4.2 or 2.2.13.4.3, and section 2.2.3.2.4.3 of this specification.");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R936
            // If the successful response can be returned from server, which indicate the serverReplicaGuid and localId got from RopGetLocalReplicaIdsResponse is parsed according to the rule specified in requirement. So it can be captured directly.
            Site.CaptureRequirement(
                936,
                @"[In RopGetLocalReplicaIds ROP Response Buffer] [GlobalCount (6 bytes)] This array and the value of the REPLGUID field are combined to produce the first GID structure ([MS-OXCDATA] section 2.2.1.3) in the allocated range of IDs, which is defined as [GlobalCount, GlobalCount + IdCount - 1] .");
        }

        /// <summary>
        /// Verify Rop SynchronizationImportMessageMove
        /// </summary>
        /// <param name="synchronizationImportMessageMoveResponse">The response was gotten from server.</param>
        private void VerifyRopSynchronizationImportMessageMove(RopSynchronizationImportMessageMoveResponse synchronizationImportMessageMoveResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R870");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R870
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                synchronizationImportMessageMoveResponse.ReturnValue.GetType(),
                870,
                @"[In RopSynchronizationImportMessageMove ROP Response Buffer]Return value (4 bytes): An unsigned 32-bit integer.");

            if (synchronizationImportMessageMoveResponse.MessageId != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R878");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R878
                // ulong is 8 bytes
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ulong),
                    synchronizationImportMessageMoveResponse.MessageId.GetType(),
                    878,
                    @"[In RopSynchronizationImportMessageMove ROP Response Buffer]MessageId (8 bytes): A 64-bit identifier.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R879001");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R879001
                Site.CaptureRequirementIfAreEqual<ulong?>(
                    0x0000000000000000,
                    synchronizationImportMessageMoveResponse.MessageId,
                    879001,
                    @"[In RopSynchronizationImportMessageMove ROP Response Buffer] [MessageId (8 bytes)] [MID structure] This value MUST be set to 0x0000000000000000.");

                MessageId messageId = new MessageId((ulong)synchronizationImportMessageMoveResponse.MessageId);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2180");

                // Verify MS-OXCFXICS requirement: MS-OXCDATA_R2180
                Site.CaptureRequirementIfAreEqual<ulong?>(
                    0x0000000000000000,
                    synchronizationImportMessageMoveResponse.MessageId,
                    "MS-OXCDATA",
                    2180,
                    @"[In Message ID Structure] It [Message ID] is an 8-byte structure.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2181");

                // Verify MS-OXCFXICS requirement: MS-OXCDATA_R2181
                Site.CaptureRequirementIfAreEqual<int>(
                    2,
                    messageId.ReplicaId.Length,
                    "MS-OXCDATA",
                    2181,
                    @"[In Message ID Structure] ReplicaId (2 bytes): An unsigned integer identifying a Store object.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2182");

                // Verify MS-OXCFXICS requirement: MS-OXCDATA_R2182
                Site.CaptureRequirementIfAreEqual<int>(
                    6,
                    messageId.GlobalCounter.Length,
                    "MS-OXCDATA",
                    2182,
                    @"[In Message ID Structure] GlobalCounter (6 bytes): An unsigned integer identifying the message within its Store object.");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R879
                // Since the MID structure can be verified according to MS-OXCDATA, so this requirement can be verified directly.
                Site.CaptureRequirement(
                    879,
                    @"[In RopSynchronizationImportMessageMove ROP Response Buffer] [MessageId (8 bytes)] The MID structure ([MS-OXCDATA] section 2.2.1.2) of the moved message in a destination folder.");
            }

            #region For successful response
            if (synchronizationImportMessageMoveResponse.ReturnValue == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R880");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R880
                bool isVerifyR880 = synchronizationImportMessageMoveResponse.MessageId != null;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR880,
                    880,
                    @"[In RopSynchronizationImportMessageMove ROP Response Buffer] [MessageId (8 bytes)] This value MUST be present if and only if ReturnValue equals Success (0x00000000).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2245");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2245
                // synchronizationImportMessageMoveResponse.RopId of this Rop is 0x78 according [MS-OXCROPS]
                // if the requirements above are verified and this RopId == 0x78, this requirement can be captured
                bool isVerifyR2245 = synchronizationImportMessageMoveResponse.RopId == 0x78;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2245,
                    2245,
                    @"[In Receiving a RopSynchronizationImportMessageMove Request]The server MUST respond with a RopSynchronizationImportMessageMove ROP 
                    response, as specified in [MS-OXCROPS] section 2.2.13.6.2 or 2.2.13.6.3, and section 2.2.3.2.4.4 of this specification.");
            }
            #endregion
        }

        /// <summary>
        /// Verify Rop SynchronizationImportDeletes
        /// </summary>
        /// <param name="synchronizationImportDeletesRequest">The request was sent to server.</param>
        /// <param name="synchronizationImportDeletesResponse">The response was gotten from server.</param>
        private void VerifyRopSynchronizationImportDeletes(RopSynchronizationImportDeletesRequest synchronizationImportDeletesRequest, RopSynchronizationImportDeletesResponse synchronizationImportDeletesResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R896");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R896
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                synchronizationImportDeletesResponse.ReturnValue.GetType(),
                896,
                @"[In RopSynchronizationImportDeletes ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2249");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2249
            // synchronizationImportDeletesResponse.RopId of this Rop is 0x74 according [MS-OXCROPS]
            // if the requirements above are verified and this RopId == 0x74, this requirement can be captured
            bool isVerifyR2249 = synchronizationImportDeletesResponse.RopId == 0x74;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2249,
                2249,
                @"[In Receiving a RopSynchronizationImportDeletes Request]The server MUST respond with a RopSynchronizationImportDeletes ROP response, as 
                specified in [MS-OXCROPS] section 2.2.13.5.2 and section 2.2.3.2.4.5 of this specification.");

            if (synchronizationImportDeletesResponse.ReturnValue != 0)
            {
                if (synchronizationImportDeletesRequest.IsHierarchy == 0x10)
                {
                    if (Common.IsRequirementEnabled(2254001, this.Site))
                    {
                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            2254001,
                            @"[In Receiving a RopSynchronizationImportDeletes Request] Implementation does fail the ROP if unknown ImportDeleteFlags flag bits are set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                    }
                }
            }
            else
            {
                if (synchronizationImportDeletesRequest.IsHierarchy == (byte)ImportDeleteFlags.Hierarchy)
                {
                    // If the ImportDeleteFlag is Hierarchy and the ROP responses the successful result, which means the flag can work. So this requirement will be verified.
                    this.Site.CaptureRequirement(
                            901,
                            @"[In RopSynchronizationImportDeletes ROP Request Buffer] [ImportDeleteFlags, When the Name is Hierarchy, the Value of Hierarchy is 0x01] If this flag is set, folder deletions are being imported.");
                }
                else if (synchronizationImportDeletesRequest.IsHierarchy == (byte)ImportDeleteFlags.delete)
                {
                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R902
                    // If the ImportDeleteFlag is not Hierarchy and the ROP responses the successful result, which means the flag can work. So this requirement will be verified.
                    this.Site.CaptureRequirement(
                           902,
                           @"[In RopSynchronizationImportDeletes ROP Request Buffer] [ImportDeleteFlags, When the Name is Hierarchy, the Value of Hierarchy is 0x01] If this flag is not set, message deletions are being imported.");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R904
                    // if the HardDelete flag is not set and the ROP response the successful result, which indicates hard deletions are not being imported. 
                    this.Site.CaptureRequirement(
                        904,
                        @"[In RopSynchronizationImportDeletes ROP Request Buffer] [ImportDeleteFlags,When the Name is HardDelete, the Value of HardDelete is 0x02] If this flag is not set, hard deletions are not being imported.");
                }
                else if (synchronizationImportDeletesRequest.IsHierarchy == (byte)ImportDeleteFlags.HardDelete)
                {
                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R903,MS-OXCFXICS_R90205002
                    // If the ImportDeleteFlag is HardDelete and the ROP responses the successful result, which means the flag can work. So this requirement will be verified.
                    this.Site.CaptureRequirement(
                          903,
                          @"[In RopSynchronizationImportDeletes ROP Request Buffer] [ImportDeleteFlags,When the Name is HardDelete, the Value of HardDelete is 0x02] If this flag is set, hard deletions are being imported.");

                    if (Common.IsRequirementEnabled(90205002, this.Site))
                    {
                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            90205002,
                            @"[In Appendix A: Product Behavior] Implementation does support. (<16> Section 2.2.3.2.4.5.1: Microsoft Exchange Server 2010 and Microsoft Exchange Server 2013 follow this behavior.)");
                    }
                }
            }
        }

        /// <summary>
        /// Verify Rop SynchronizationImportReadStateChanges
        /// </summary>
        /// <param name="synchronizationImportReadStateChangesResponse">The response was gotten from server.</param>
        private void VerifyRopSynchronizationImportReadStateChanges(RopSynchronizationImportReadStateChangesResponse synchronizationImportReadStateChangesResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R919");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R919
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                synchronizationImportReadStateChangesResponse.ReturnValue.GetType(),
                919,
                @"[In RopSynchronizationImportReadStateChanges ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2256");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2256
            // synchronizationImportReadStateChangesResponse.RopId of this Rop is 0x80 according [MS-OXCROPS]
            // if the requirements above are verified and this RopId == 0x80, this requirement can be captured
            bool isVerifyR2256 = synchronizationImportReadStateChangesResponse.RopId == 0x80;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2256,
                2256,
                @"[In Receiving a RopSynchronizationImportReadStateChanges Request]The server MUST respond with a RopSynchronizationImportReadStateChanges 
                response, as specified in [MS-OXCROPS] section 2.2.13.3.2 and section 2.2.3.2.4.6 of this specification.");
        }

        /// <summary>
        /// Verify Rop GetLocalReplicaIds
        /// </summary>
        /// <param name="getLocalReplicaIdsResponse">The response was gotten from server.</param>
        private void VerifyRopGetLocalReplicaIds(RopGetLocalReplicaIdsResponse getLocalReplicaIdsResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R931");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R931
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                getLocalReplicaIdsResponse.ReturnValue.GetType(),
                931,
                @"[In RopGetLocalReplicaIds ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R933");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R933
            Site.CaptureRequirementIfAreEqual<int>(
                16,
                getLocalReplicaIdsResponse.ReplGuid.Length,
                933,
                @"[In RopGetLocalReplicaIds ROP Response Buffer] REPLGUID (16 bytes): A GUID value.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R935");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R935
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte[]),
                getLocalReplicaIdsResponse.GlobalCount.GetType(),
                935,
                @"[In RopGetLocalReplicaIds ROP Response Buffer] GlobalCount (6 bytes): An array of bytes.");

            if (getLocalReplicaIdsResponse.ReturnValue == 0x00000000)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R934");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R934
                bool isVerifyR934 = getLocalReplicaIdsResponse.ReplGuid != null;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR934,
                    934,
                    @"[In RopGetLocalReplicaIds ROP Response Buffer] [REPLGUID (16 bytes)] This value MUST be present if and only if value of the ReturnValue field equals Success (0x00000000).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R937");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R937
                bool isVerifyR937 = getLocalReplicaIdsResponse.GlobalCount != null;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR937,
                    937,
                    @"[In RopGetLocalReplicaIds ROP Response Buffer] [GlobalCount (6 bytes)] This value MUST be present if and only if the value of the ReturnValue field equals Success (0x00000000).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2264");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2264
                // getLocalReplicaIdsResponse.RopId of this Rop is 0x7F according [MS-OXCROPS]
                // if the requirements above are verified and this RopId == 0x7F, this requirement can be captured
                bool isVerifyR2264 = getLocalReplicaIdsResponse.RopId == 0x7F;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2264,
                    2264,
                    @"[In Receiving a RopGetLocalReplicaIds Request]The server MUST respond with a RopGetLocalReplicaIds ROP response, as specified in 
                    [MS-OXCROPS] section 2.2.13.13.2 or 2.2.13.13.3, and section 2.2.3.2.4.7 of this specification.");
            }
        }

        /// <summary>
        /// Verify Rop SetLocalReplicaMidsetDeleted
        /// </summary>
        /// <param name="setLocalReplicaMidsetDeletedResponse">The response was gotten from server.</param>
        private void VerifyRopSetLocalReplicaMidsetDeleted(RopSetLocalReplicaMidsetDeletedResponse setLocalReplicaMidsetDeletedResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R954");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R954
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                setLocalReplicaMidsetDeletedResponse.ReturnValue.GetType(),
                954,
                @"[In RopSetLocalReplicaMidsetDeleted ROP Response Buffer] ReturnValue (4 bytes): An unsigned 32-bit integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2268");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2268
            // setLocalReplicaMidsetDeletedResponse.RopId of this Rop is 0x93 according [MS-OXCROPS]
            // if the requirements above are verified and this RopId == 0x93, this requirement can be captured
            bool isVerifyR2268 = setLocalReplicaMidsetDeletedResponse.RopId == 0x93;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2268,
                2268,
                @"[In Receiving a RopSetLocalReplicaMidsetDeleted Request]The server MUST respond with a RopSetLocalReplicaMidsetDeleted ROP response, as 
                specified in [MS-OXCROPS] section 2.2.13.12.2 and section 2.2.3.2.4.8 of this specification.");
        }

        /// <summary>
        /// Verify server accept PidTagIdsetGiven (with the PtypBinary type)
        /// </summary>
        /// <param name="serverId">The id for server.</param>
        /// <param name="handleIndex">The index for getProperties handle.</param>
        private void VerifyServerAcceptPidTagIdsetGivenPtypBinary(int serverId, int handleIndex)
        {
            // The server accepts PidTagIdsetGivenPtypBinary with Ptypinary type if ROP GetPropertiesSpecific is returned successfully.
            Sequence<string> propertyTag = new Sequence<string>("PidTagIdsetGivenPtypBinary");
            RopResult result = this.GetPropertiesSpecific(serverId, handleIndex, propertyTag);

            if (Common.IsRequirementEnabled(33551, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R33551");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R33551
                this.Site.CaptureRequirementIfAreEqual<RopResult>(
                    RopResult.Success,
                    result,
                    33551,
                    @"[In Sending and Receiving the PidTagIdsetGiven ICS State Property] Implementation does accept this PidTagIdsetGiven property when the property tag identifies it as PtypInteger32 or PtypBinary. (Microsoft Exchange Server 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(2658, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2658");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2658
                this.Site.CaptureRequirementIfAreEqual<RopResult>(
                    RopResult.Success,
                    result,
                    2658,
                    @"[In Receiving the MetaTagIdsetGiven ICS State Property] Implementation does accept this MetaTagIdsetGiven property when the property tag identifies it as PtypBinary. (Microsoft Exchange Server 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// Verify server accept PidTagIdsetGiven (with the PtypInteger32 type).
        /// </summary>
        /// <param name="ropResult">The result of the ROP operation.</param>
        private void VerifyServerAcceptPidTagIdsetGivenPtypInteger32(RopResult ropResult)
        {
            if (Common.IsRequirementEnabled(33551, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R33551");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R33551
                this.Site.CaptureRequirementIfAreEqual<RopResult>(
                    RopResult.Success,
                    ropResult,
                    33551,
                    @"[In Sending and Receiving the PidTagIdsetGiven ICS State Property] Implementation does accept this PidTagIdsetGiven property when the property tag identifies it as PtypInteger32 or PtypBinary. (Microsoft Exchange Server 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(2656, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2656");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2656
                this.Site.CaptureRequirementIfAreEqual<RopResult>(
                    RopResult.Success,
                    ropResult,
                    2656,
                    @"[In Receiving the MetaTagIdsetGiven ICS State Property] Implementation does send the MetaTagIdsetGiven property (section 2.2.1.1.1) with a property tag that defines it as PtypInteger32. (Microsoft Exchange Server 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// Verify Whether PCL Equal PidTagChangeKey
        /// </summary>
        /// <param name="serverId">the id for server</param>
        /// <param name="handleIndex">The index for getProperties handle</param>
        private void VerifyValuePidTagPredecessorChangeListAndPidTagChangeKey(int serverId, int handleIndex)
        {
            // Get the value of PidTagChangeKey and PidTagPredecessorChangeList by use ROP GetPropertiesSpecific
            Sequence<string> propertyTag = new Sequence<string>("PidTagChangeKey", "PidTagPredecessorChangeList");
            this.GetPropertiesSpecific(serverId, handleIndex, propertyTag);
            byte[] pidTagChangeKeyValue = this.propertyValuesSpecific["PidTagChangeKey"];
            byte[] pidTagPredecessorChangeListValue = this.propertyValuesSpecific["PidTagPredecessorChangeList"];

            if (this.lastConflictInfo.PCLA != null)
            {
                this.lastConflictInfo.PCLXFromServer = pidTagPredecessorChangeListValue;
                this.VerifyPidTagPredecessorChangeList();
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1911");

            // If messageID is less than 0 means this case is used for verify conflict status.
            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1911
            bool isVerifyR1911 = false;
            string changeKeyValue1, changeKeyValue2 = string.Empty;
            if (pidTagChangeKeyValue.Length == 24)
            {
                changeKeyValue1 = BitConverter.ToString(pidTagChangeKeyValue, 2);
            }
            else
            {
                changeKeyValue1 = BitConverter.ToString(pidTagChangeKeyValue, 0);
            }

            for (int i = 0; i < (pidTagPredecessorChangeListValue.Length - 2) / 23; i++)
            {
                changeKeyValue2 += BitConverter.ToString(pidTagPredecessorChangeListValue, 2 + (23 * i) + 1);
            }

            isVerifyR1911 = changeKeyValue2.Contains(changeKeyValue1);

            if (isVerifyR1911 == false)
            {
                this.Site.Log.Add(LogEntryKind.Comment, "PidTagChangeKey value is:" + changeKeyValue1 + ".\r\n");
                this.Site.Log.Add(LogEntryKind.Comment, "PidTagPredecessorChangeList value is:" + changeKeyValue2 + ".\r\n");
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1911,
                1911,
                @"[In Identifying Objects and Maintaining Change Numbers]Upon successful import of a new or changed object using ICS upload, the server MUST do 
                the following when receiving RopSaveChangesMessage ROP:	Assign the object the given PidTagChangeKey property value (section 2.2.1.2.7) and 
                PidTagPredecessorChangeList (section 2.2.1.2.8) that equals PCL {PidTagChangeKey}.");
        }

        /// <summary>
        /// Verify FastTransfer Stream get from server
        /// </summary>
        /// <param name="fastTransferStream">The fastTransfer Stream get from server</param>
        /// <param name="streamType">The type of the ROP for Initiate Operation</param>
        private void VerifyFastTransferStream(FastTransferStream fastTransferStream, FastTransferStreamType streamType)
        {
            switch (streamType)
            {
                // Verify AttachmentContent
                case FastTransferStreamType.attachmentContent:
                    this.VerifyAttachmentContent(fastTransferStream);
                    break;

                // Verify ContentsSync
                case FastTransferStreamType.contentsSync:
                    this.VerifyContentsSync(fastTransferStream);
                    break;

                // Verify HierarchySync
                case FastTransferStreamType.hierarchySync:
                    this.VerifyHierarchySync(fastTransferStream);
                    break;

                // Verify FolderContent
                case FastTransferStreamType.folderContent:
                    this.VerifyFolderContent(fastTransferStream);
                    break;

                // Verify MessageContent
                case FastTransferStreamType.MessageContent:
                    this.VerifyMessageContent(fastTransferStream);
                    break;

                // Verify State
                case FastTransferStreamType.state:
                    this.VerifyState(fastTransferStream);
                    break;

                // Verify TopFolder
                case FastTransferStreamType.TopFolder:
                    this.VerifyTopFolder(fastTransferStream);
                    break;

                // Verify MessageList
                case FastTransferStreamType.MessageList:
                    this.VerifyMessageList(fastTransferStream);
                    break;

                default:
                    break;
            }

            // When the stream verified means that it is in line with the Lexical structure so these requirements are captured. 
            this.Site.CaptureRequirement(
                2740,
                @"[In Lexical structure] The lexical structure of a FastTransfer stream is                 as follows:
                stream = 1*element
                element = marker / propValue
                marker = PtypInteger32 <from the table in 2.2.4.1.4>
                propValue = fixedPropType propInfo fixedSizeValue
                propValue =/ varPropType propInfo length varSizeValue
                propValue =/ mvPropType 
                  propInfo 
                  length 
                  *( fixedSizeValue / length varSizeValue )
                propInfo = taggedPropId / ( namedPropId namedPropInfo )
                fixedSizeValue = PtypInteger16 / PtypInteger32 / PtypFloating32
                  / PtypFloating64 / PtypCurrency / PtypFloatingTime
                  / PtypBoolean / PtypInteger64 / PtypTime
                  / PtypGuid 
                varSizeValue = PtypString / PtypString8 / PtypServerId 
                  / PtypBinary / PtypObject


                namedPropInfo = propertySet 
                  ((%x00 dispid) 
                  / (%x01 name))
                propertySet = PtypGuid
                dispid = PtypInteger32
                name = PtypString
                namedPropId = propertyId 
                  <Greater or equal to 0x8000>
                propertyId = PtypInteger16
                taggedPropId = propertyId 
                  <less than 0x8000>
                length = PtypInteger32 <MUST be greater than 0>
                propType = fixedPropType / varPropType / mvPropType
                fixedPropType = PtypInteger16
                varPropType = PtypInteger16
                mvPropType = PtypInteger16
                ");

            // When the code can reached here, which means that it is in line with the Lexical structure so these requirements are captured. 
            this.Site.CaptureRequirement(
                981,
                "[In Lexical structure] If a split is required, the stream MUST be split either between two atoms or at any point inside a varSizeValue lexeme.");

            // When the code can reached here, which means that it is in line with the Lexical structure so these requirements are captured. 
            this.Site.CaptureRequirement(
                983,
                @"[In Lexical structure] The lexical structure of an atom is as follows:
                atom = marker 
                    / propDef 
                    / fixedSizeValue 
                    / length
                propDef = ( propType propInfo )");
        }

        /// <summary>
        /// Verify the AttachmentContent
        /// </summary>
        /// <param name="fastTransferStream">The fastTransfer Stream get from server</param>
        private void VerifyAttachmentContent(FastTransferStream fastTransferStream)
        {
            AttachmentContent attachContent = new AttachmentContent(fastTransferStream);

            // Indicate whether the simple property type PtypBoolean or PtypString exist
            bool simplePropertyTypeAppear = false;

            #region Verify string and String8
            foreach (PropValue prop in attachContent.PropList.PropValues)
            {
                // PropType equal to 0x001F means type is PtypString
                if (prop.PropType == 0x001F)
                {
                    VarPropTypePropValue varPropTypePropValue = prop as VarPropTypePropValue;
                    if (varPropTypePropValue != null && varPropTypePropValue.ValueArray != null && varPropTypePropValue.ValueArray.Length > 1)
                    {
                        simplePropertyTypeAppear = true;

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2700");

                        // Verify MS-OXCFXICS requirement: MS-OXCDATA_R2700
                        // A string property  with terminating null character (2 bytes of zero)
                        bool isVerifyR2700 = (prop as VarPropTypePropValue).ValueArray[(prop as VarPropTypePropValue).ValueArray.Length - 1] == 0
                                            && (prop as VarPropTypePropValue).ValueArray[(prop as VarPropTypePropValue).ValueArray.Length - 2] == 0;
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR2700,
                            "MS-OXCDATA",
                            2700,
                            @"[In Property Data Types] PtypString (PT_UNICODE, string) is that Variable size; a string of Unicode characters in UTF-16LE format encoding with terminating null character (0x0000). with Property Type Value  0x001F,%x1F.00.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R992");

                        // String type is Serialized the as specified in [MS-OXCDATA]
                        // if this element can be gotten, this requirement can be covered .
                        Site.CaptureRequirement(
                             992,
                             @"[In Serialization of Simple Types] Serialization of simple types in FastTransfer streams is identical to serialization of property values as specified [MS-OXCDATA] , with the following exceptions: For PtypString and PtypString8 ([MS-OXCDATA] section 2.11.1) type, serialization MUST be performed, as specified in [MS-OXCDATA] .");
                    }
                }
                
                // PropType equal to 0x001E means type is PtypString8
                if (prop.PropType == 0x001E)
                {
                    VarPropTypePropValue varPropTypePropValue = prop as VarPropTypePropValue;
                    if (varPropTypePropValue != null && varPropTypePropValue.ValueArray != null && varPropTypePropValue.ValueArray.Length > 1)
                    {
                        simplePropertyTypeAppear = true;

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2701");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2701
                        // A string property  with terminating null character (1 bytes of zero)
                        bool isVerifyR2701 = (prop as VarPropTypePropValue).ValueArray[(prop as VarPropTypePropValue).ValueArray.Length - 1] == 0
                                            && (prop as VarPropTypePropValue).ValueArray[(prop as VarPropTypePropValue).ValueArray.Length - 2] != 0;

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR2701,
                            "MS-OXCDATA",
                            2701,
                            @"[In Property Data Types]PtypString8 is that variable size, a string of multi-byte characters in externally specified 
                                encoding with terminating null character (single 0 byte).PT_STRING8 with Property Type Value 0x001E,%z1E.00.");

                        // String type is Serialized the as specified in [MS-OXCDATA]
                        // if this element can be gotten, this requirement can be covered .
                        Site.CaptureRequirement(
                             992,
                            @"[In Serialization of Simple Types] Serialization of simple types in FastTransfer streams is identical to serialization of property values as specified [MS-OXCDATA] , with the following exceptions: For PtypString and PtypString8 ([MS-OXCDATA] section 2.11.1) type, serialization MUST be performed, as specified in [MS-OXCDATA] .");
                    }
                }
            }

            if (simplePropertyTypeAppear == true)
            {
                // This capture code can be directly verified since the adapter parser analyzes the data just in little-endian byte ordering only if the PtypString or PtypString8 type property exists. 
                Site.CaptureRequirement(
                        996,
                        "[In Serialization of Simple Types] Note that little-endian byte ordering MUST be used.");
            }

            #endregion
        }

        /// <summary>
        /// Verify the ContentsSync element of fastTransfer Stream
        /// </summary>
        /// <param name="fastTransferStream">The fastTransfer Stream get from server</param>
        private void VerifyContentsSync(FastTransferStream fastTransferStream)
        {
            ContentsSync conSync = new ContentsSync(fastTransferStream);
            this.previousStepVerifiedResult = true;

            #region Verify different subelements of contentsSync in FastTransfer stream for different configurations.
            if (((this.synchroniztionFlag & SynchronizationFlag.Normal) == SynchronizationFlag.Normal) && ((this.synchroniztionFlag & SynchronizationFlag.FAI) == SynchronizationFlag.FAI))
            {
                if (conSync.MessageChangeTuples.Count != 0)
                {
                    // The first changed message id FAI message and the second is normal. For FAI message is verified in other test suite so only verify normal message and all message here.
                    Tuple<ProgressPerMessage, MessageChange> tuple = conSync.MessageChangeTuples[0];
                    Tuple<ProgressPerMessage, MessageChange> tuple1 = conSync.MessageChangeTuples[1];

                    if (tuple.Item2 is MessageChangeFull && tuple1.Item2 is MessageChangeFull)
                    {
                        MessageChangeFull item1 = (MessageChangeFull)tuple.Item2;
                        MessageChangeFull item2 = (MessageChangeFull)tuple1.Item2;

                        if (this.synchroniztionFlag == (SynchronizationFlag.FAI | SynchronizationFlag.Normal | SynchronizationFlag.IgnoreSpecifiedOnFAI))
                        {
                            bool isPropertyExist = false;

                            if (this.propertyTagForConfigure.PropertyId == item2.PropList.PropValues[0].PropInfo.PropID)
                            {
                                isPropertyExist = true;
                            }

                            // If the OnlySpecifiedProperties flag is not set, server must ignore the IgnoreSpecifiedOnFAI flag.
                            // So Server must write all properties and subobjects for non-FAI messages to the FastTransfer stream.
                            this.Site.CaptureRequirementIfIsTrue(
                                        isPropertyExist==false && item2.PropList.PropValues.Count>1,
                                        218300301,
                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the OnlySpecifiedProperties flag is not set, the server MUST return same value wether the IgnoreSpecifiedOnFAI flag is set or not.");

                            this.messageChangeFull = item2;
                        }
                        else if (this.synchroniztionFlag == (SynchronizationFlag.FAI | SynchronizationFlag.Normal))
                        {
                            this.messageChangeFull = item2;
                        }

                        if (this.synchroniztionFlag == (SynchronizationFlag.FAI | SynchronizationFlag.Normal | SynchronizationFlag.OnlySpecifiedProperties))
                        {
                            bool isPropertyExist = false;
                            if (this.propertyTagForConfigure.PropertyId == item1.PropList.PropValues[0].PropInfo.PropID || this.propertyTagForConfigure.PropertyId == item2.PropList.PropValues[0].PropInfo.PropID)
                            {
                                isPropertyExist = true;
                            }

                            this.Site.CaptureRequirementIfIsTrue(
                                isPropertyExist,
                                2183004,
                                @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the OnlySpecifiedProperties flag is set and the IgnoreSpecifiedOnFAI flag is not set, the server writes only the specified properties for all messages to the FastTransfer stream.");

                            this.Site.CaptureRequirementIfIsTrue(
                                isPropertyExist,
                                21830041,
                                @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the OnlySpecifiedProperties flag is set and the IgnoreSpecifiedOnFAI flag is not set, the server writes only the subobjects for all messages to the FastTransfer stream.");
                        }
                    }
                }
            }

            if (this.synchroniztionFlag == (SynchronizationFlag.Normal | SynchronizationFlag.OnlySpecifiedProperties | SynchronizationFlag.IgnoreSpecifiedOnFAI))
            {
                Tuple<ProgressPerMessage, MessageChange> tuple = conSync.MessageChangeTuples[0];

                if (tuple.Item2 is MessageChangeFull && tuple.Item2 is MessageChangeFull)
                {
                    MessageChangeFull item2 = (MessageChangeFull)tuple.Item2;
                    bool isPropertyExist = false;

                    if (this.propertyTagForConfigure.PropertyId == item2.PropList.PropValues[0].PropInfo.PropID)
                    {
                        isPropertyExist = true;
                    }

                    this.Site.CaptureRequirementIfIsTrue(
                        isPropertyExist,
                        2183005,
                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the OnlySpecifiedProperties flag is set and the IgnoreSpecifiedOnFAI flag is also set, the server MUST write only the specified properties for all non-FAI messages to the FastTransfer stream.");

                    this.Site.CaptureRequirementIfIsTrue(
                        isPropertyExist,
                        21830051,
                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the OnlySpecifiedProperties flag is set and the IgnoreSpecifiedOnFAI flag is also set, the server MUST write only the subobjects for all non-FAI messages to the FastTransfer stream.");
                }
            }

            if ((this.synchroniztionFlag & SynchronizationFlag.FAI) == SynchronizationFlag.FAI)
            {
                if (conSync.MessageChangeTuples.Count != 0)
                {
                    Tuple<ProgressPerMessage, MessageChange> tuple = conSync.MessageChangeTuples[0];
                    MessageChangeFull item2 = tuple.Item2 as MessageChangeFull;
                    if (item2 != null)
                    {
                        if (this.synchroniztionFlag == SynchronizationFlag.FAI)
                        {
                            this.messageChangeFull = item2;
                        }

                        if (this.synchroniztionFlag == (SynchronizationFlag.FAI | SynchronizationFlag.IgnoreSpecifiedOnFAI | SynchronizationFlag.OnlySpecifiedProperties))
                        {
                            bool isSameProperties = true;
                            if (item2.PropList.PropValues.Count != this.messageChangeFull.PropList.PropValues.Count)
                            {
                                isSameProperties = false;
                            }
                            else
                            {
                                foreach (PropValue proValue1 in item2.PropList.PropValues)
                                {
                                    bool isPropertyExist = false;
                                    foreach (PropValue proValue2 in this.messageChangeFull.PropList.PropValues)
                                    {
                                        if (proValue1.PropInfo.PropID == proValue2.PropInfo.PropID)
                                        {
                                            isPropertyExist = true;
                                            break;
                                        }
                                    }

                                    if (!isPropertyExist)
                                    {
                                        isSameProperties = false;
                                        break;
                                    }
                                }
                            }

                            this.Site.CaptureRequirementIfIsTrue(
                                isSameProperties,
                                2183006,
                                @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the OnlySpecifiedProperties flag is set and the IgnoreSpecifiedOnFAI flag is also set, the server MUST write all properties to the FastTransfer stream for FAI messages.");

                            this.Site.CaptureRequirementIfIsTrue(
                                isSameProperties,
                                21830061,
                                @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the OnlySpecifiedProperties flag is set and the IgnoreSpecifiedOnFAI flag is also set, the server MUST write subobjects to the FastTransfer stream for FAI messages.");
                        }
                    }
                }
            }

            if (conSync.MessageChangeTuples.Count != 0)
            {
                foreach (Tuple<ProgressPerMessage, MessageChange> tuple in conSync.MessageChangeTuples)
                {
                    if (tuple.Item2 is MessageChangeFull)
                    {
                        MessageChangeFull item2 = (MessageChangeFull)tuple.Item2;
                        if (item2.MessageChildren.FXDelProps != null)
                        {
                            // If the program can execute here, this requirement can be verified directly.
                            this.Site.CaptureRequirement(
                                3285,
                                @"[In MetaTagFXDelProp Meta-Property] Property ID: 0x4016 Data type: PtypInteger32, 0x0003 ([MS-OXCDATA] section 2.11.1)");

                            foreach (uint property in item2.MessageChildren.FXDelProps)
                            {
                                switch (property)
                                {
                                    case (uint)MessageProperties.PidTagMessageRecipients:

                                        // If the program can execute here, this requirement can be verified directly.
                                        Site.CaptureRequirement(
                                            "MS-OXPROPS",
                                            6994,
                                            @"[In PidTagMessageRecipients] Property ID: 0x0E12.");

                                        // If the program can execute here, this requirement can be verified directly.
                                        Site.CaptureRequirement(
                                            "MS-OXPROPS",
                                            6995,
                                            @"[In PidTagMessageRecipients] Data type: PtypObject, 0x000D");
                                        break;
                                    case (uint)MessageProperties.PidTagMessageAttachments:

                                        // If the program can execute here, this requirement can be verified directly.
                                        Site.CaptureRequirement(
                                            "MS-OXPROPS",
                                            6925,
                                            @"[In PidTagMessageAttachments] Property ID: 0x0E13.");

                                        // If the program can execute here, this requirement can be verified directly.
                                        Site.CaptureRequirement(
                                            "MS-OXPROPS",
                                            6926,
                                            @"[In PidTagMessageAttachments] Data type: PtypObject, 0x000D");
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }

                        if ((this.synchroniztionFlag & SynchronizationFlag.IgnoreSpecifiedOnFAI) != SynchronizationFlag.IgnoreSpecifiedOnFAI)
                        {
                            bool isPidTagCreationTimeExist = false;
                            foreach (PropValue pro in item2.PropList.PropValues)
                            {
                                if (pro.PropInfo.PropID == this.propertyTagForConfigure.PropertyId)
                                {
                                    isPidTagCreationTimeExist = true;
                                    break;
                                }
                            }

                            if ((this.synchroniztionFlag & SynchronizationFlag.OnlySpecifiedProperties) == SynchronizationFlag.OnlySpecifiedProperties)
                            {
                                if (Common.IsRequirementEnabled(2646, this.Site))
                                {
                                    Site.CaptureRequirementIfIsTrue(
                                        isPidTagCreationTimeExist,
                                        2646,
                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [PropertyTags Constraints] [This field has different semantics, depending on the value of the SynchronizationFlag OnlySpecifiedProperties, as follows:] Implementation does limit properties and subobjects written to the FastTransfer stream for top-level messages to properties listed in the PropertyTags field, if the OnlySpecifiedProperties flag is set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                }

                                if (Common.IsRequirementEnabled(2647, this.Site))
                                {
                                    Site.CaptureRequirementIfIsTrue(
                                        isPidTagCreationTimeExist,
                                        2647,
                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [PropertyTags Constraints] Implementation does include these special parts from output for top-level messages, while inclusion of properties that denote message subobjects in the PropertyTags field. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                }
                            }
                            else
                            {
                                if (Common.IsRequirementEnabled(2642, this.Site))
                                {
                                    Site.CaptureRequirementIfIsFalse(
                                        isPidTagCreationTimeExist,
                                        2642,
                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] Implementation does exclude properties and subobjects output for folders and top-level messages, if they are listed in PropertyTags, if the SynchronizationFlag OnlySpecifiedProperties flag is not set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                }

                                if (Common.IsRequirementEnabled(2645, this.Site))
                                {
                                    Site.CaptureRequirementIfIsFalse(
                                        isPidTagCreationTimeExist,
                                        2645,
                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [PropertyTags Constraints] [This field has different semantics, depending on the value of the SynchronizationFlag OnlySpecifiedProperties, as follows:] Implementation does exclude properties and subobjects from the FastTransfer stream for folders and top-level messages, if the property is listed in the PropertyTags field, and if the OnlySpecifiedProperties flag is not set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                }

                                if (Common.IsRequirementEnabled(2648, this.Site))
                                {
                                    Site.CaptureRequirementIfIsFalse(
                                        isPidTagCreationTimeExist,
                                        2648,
                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [PropertyTags Constraints] Implementation does exclude these special parts from output for top-level messages, while inclusion of properties that denote message subobjects in the PropertyTags field. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                }
                            }
                        }
                    }
                }
            }
            #endregion Verify the expected FastTransfer stream returned from server for different configuration.

            #region IDSET UID GUID Command

            if (conSync.State.PropList.HasPropertyID(0x4017))
            {
                PropValue propValue = conSync.State.PropList.PropValues.Find(p => p.PropInfo.PropID == 0x4017);
                VarPropTypePropValue varPropTypePropValue = propValue as VarPropTypePropValue;
                if (varPropTypePropValue != null && varPropTypePropValue.ValueArray.Length > 0)
                {
                    using (MemoryStream memoryStream = new MemoryStream(((VarPropTypePropValue)propValue).ValueArray))
                    {
                        REPLGUID_IDSET repLid = new REPLGUID_IDSET();
                        repLid.Deserialize(memoryStream, -1);
                        List<GLOBCNT> serverGLOBCNTs = repLid.IDSetList[repLid.IDSetList.Count - 1].GLOBSET.GLOBCNTList;
                        GLOBCNT localGLOBCNT = StructureSerializer.Deserialize<GLOBCNT>(this.localId);
                        if (serverGLOBCNTs.Contains(localGLOBCNT))
                        {
                            // Because the GLOBCNT is paired with a REPLID structure it forms a Message ID, so it can be recognized as messageID. So when the GLOBCNT set from server contains the GLOBCNT from local, this requirement captured.
                            this.Site.CaptureRequirement(
                                3658,
                                @"[In MetaTagIdsetGiven ICS State Property] The IDSETs contain Message ID structures for content synchronization operations, that exist in the local replica of the client.");
                        }
                    }
                }
            }

            if ((null != conSync.State) && (null != conSync.State.PropList) && (null != conSync.State.PropList.PropValues))
            {
                foreach (PropValue prop in conSync.State.PropList.PropValues)
                {
                    // 0x4017 is the ID of PidTagIdsetGiven, 0x6796 is the ID of PidTagCnsetSeen
                    // 0x67DA is the ID of PidTagCnsetSeenFAI, 0x67D2 is the ID of PidTagCnsetRead
                    if ((prop.PropInfo.PropID == 0x4017) || (prop.PropInfo.PropID == 0x6796)
                        || (prop.PropInfo.PropID == 0x67DA) || (prop.PropInfo.PropID == 0x67D2))
                    {
                        VarPropTypePropValue varProp = (VarPropTypePropValue)prop;

                        // Transform the value of this property to a memory stream, the property value is a serialized IDSET.
                        using (MemoryStream memoryStream = new MemoryStream(varProp.ValueArray))
                        {
                            REPLGUID_IDSET idset = new REPLGUID_IDSET();

                            // Deserialize the above stream to get an IDSET.
                            // "-1" means no limitation except stream length.
                            idset.Deserialize(memoryStream, -1);

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1934");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1934
                            // The structure is deserialized as this requirement, if Deserialize success, set IsAllGLOBCNTRanged as true
                            bool isVerifyR1934 = idset.IsAllGLOBCNTRanged;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1934,
                                1934,
                                @"[In Formatted IDSET Structures]. The ID values MUST be arranged by REPLID structure value");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1935");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1935
                            // The structure is deserialized as this requirement, if Deserialize success, set IsAllGLOBCNTInGLOBSET as true
                            bool isVerifyR1935 = idset.IsAllGLOBCNTInGLOBSET;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1935,
                                1935,
                                @"[In Formatted IDSET Structures]all IDs for each REPLID MUST be reduced into a GLOBSET, as specified in section 2.2.2.6,
                            of GLOBCNTs, as specified in section 2.2.2.5.");

                            // Add the debug information
                            // The structure is deserialized as this requirement, if Deserialize success, set IsAllGLOBCNTInGLOBSET as true
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1936");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1936
                            bool isVerifyR1936 = idset.HasAllDuplicateGLOBCNTRemoved;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1936,
                                1936,
                                @"[In Formatted IDSET Structures]Each GLOBSET MUST be arranged from lowest to highest GLOBCNT value where all duplicate GLOBCNT values are removed.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1937");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1937
                            // The structure is deserialized as this requirement, if Deserialize success, set HasGLOBCNTGroupedIntoRanges as true
                            bool isVerifyR1937 = idset.HasGLOBCNTGroupedIntoRanges;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1937,
                                1937,
                                @"[In Formatted IDSET Structures]The remaining GLOBCNT values MUST be grouped into consecutive ranges with a low GLOBCNT value and a high GLOBCNT value.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1938");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1938
                            // The structure is deserialized as this requirement, if Deserialize success, set IsDisjointGLOBCNTMadeIntoSingleton as true
                            bool isVerifyR1938 = idset.IsDisjointGLOBCNTMadeIntoSingleton;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1938,
                                1938,
                                @"[In Formatted IDSET Structures]If a GLOBCNT value is disjoint it MUST be made into a singleton range with the low and high GLOBCNT values being the same.");

                            foreach (REPLGUID_IDSETElement replguid in idset.IDSetList)
                            {
                                // The response from the server is serialized as this requirement is related to server, 
                                // if de-serialized the response (as the rule for server serializes the response) successfully, this requirement could be verified
                                Site.CaptureRequirement(
                                    186,
                                    @"[In Serialized IDSET Structure Containing a REPLGUID Structure] REPLGUID-GLOBSET structure pairs MUST be serialized by the value of the REPLGUID in the ascending order, using byte-to-byte comparison.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R187");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R187
                                Site.CaptureRequirementIfAreEqual<Type>(
                                    typeof(Guid),
                                    replguid.REPLGUID.GetType(),
                                    187,
                                    @"[In Serialized IDSET Structure Containing a REPLGUID Structure]REPLGUID (16 bytes): A GUID that identifies a REPLGUID structure.");

                                for (int i = 0; i < replguid.GLOBSET.DeserializedCommandList.Count; i++)
                                {
                                    if (replguid.GLOBSET.DeserializedCommandList[i] is PushCommand)
                                    {
                                        // comByteStack is the common byte stack
                                        CommonByteStack comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, new CommonByteStack());

                                        // The deserialization do as:
                                        // The Push command SHOULD be used when multiple GLOBCNT values share the same high-order values.
                                        // When push command is encountered, multiple GLOBCNT values share the same high-order values in Exchange
                                        bool pushHasSameHighOrderValue = this.WhetherPushHasSameHighOrderValue(replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTList);
                                        if (pushHasSameHighOrderValue)
                                        {
                                            // Verify requirement MS-OXCFXICS_R2660
                                            bool isR2660Satisfied = true;
                                            if (null != replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTList)
                                            {
                                                foreach (GLOBCNT globcnt in replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTList)
                                                {
                                                    for (int index = 0; index < comByteStack.Bytes; index++)
                                                    {
                                                        if (globcnt[index] != comByteStack.GetCommonBytes()[index])
                                                        {
                                                            // The multiple GLOBCNT structure values share the same higher-order values but Push command is not used.
                                                            isR2660Satisfied = false;
                                                            break;
                                                        }
                                                    }
                                                }

                                                if (Common.IsRequirementEnabled(2660, this.Site))
                                                {
                                                    // Add the debug information
                                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2660");

                                                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2660
                                                    Site.CaptureRequirementIfIsTrue(
                                                        isR2660Satisfied,
                                                        2660,
                                                        @"[In Push Command (0x01 - 0x06)] Implementation does use the Push command when multiple GLOBCNT structure values share the same high-order values. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                                }
                                            }
                                        }

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R202");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R202
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            (replguid.GLOBSET.DeserializedCommandList[i] as PushCommand).CommandByte,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as PushCommand).CommonBytes.Length,
                                            202,
                                            @"[In Push Command (0x01 - 0x06)] [CommonBytes (variable):] The length of this field is specified by value of the Command field (0x01 through 0x06), in bytes.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3064");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3064
                                        Site.CaptureRequirementIfAreEqual<Type>(
                                            typeof(byte),
                                            (replguid.GLOBSET.DeserializedCommandList[i] as PushCommand).CommandByte.GetType(),
                                            3064,
                                            @"[In Push Command (0x01 - 0x06)] Command (1 byte): An integer that specifies the number of high-order bytes that the GLOBCNT structures, as specified in section 2.2.2.5, share. ");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3065");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3065
                                        short command = Convert.ToInt16((replguid.GLOBSET.DeserializedCommandList[i] as PushCommand).CommandByte);
                                        bool isVerifyR3065 = command >= 0x01 && command <= 0x06;
                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR3065,
                                            3065,
                                            @"[In Push Command (0x01 - 0x06)] Command: This value MUST be in the range 0x01 through 0x06.");
                                    }

                                    if (replguid.GLOBSET.DeserializedCommandList[i] is PopCommand)
                                    {
                                        // comByteStack is the common byte stack
                                        CommonByteStack comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, new CommonByteStack());

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1966");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1966
                                        // from this requirement, can know if "Pop Command" is used, there are some bytes on the common byte stack. 
                                        // So if "Pop Command" is used and the common byte stack has some bytes, then this requirement can be covered.
                                        bool isVerifyR1966 = comByteStack.Bytes > 0;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR1966,
                                            1966,
                                            @"[In Pop Command (0x50)]The Pop command MUST NOT be used if no bytes are currently on the common byte stack.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3068");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3068
                                        Site.CaptureRequirementIfAreEqual<short>(
                                            0x50,
                                            Convert.ToInt16((replguid.GLOBSET.DeserializedCommandList[i] as PopCommand).CommandByte),
                                            3068,
                                            @"[In Pop Command (0x50)] Command (1 byte): This value MUST be set to 0x50.");
                                    }

                                    if (replguid.GLOBSET.DeserializedCommandList[i] is RangeCommand)
                                    {
                                        bool rangHasSameHighOrderValue = this.WhetherRangHasSameHighOrderValue(replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTRangeList);

                                        // comByteStack is the common byte stack
                                        CommonByteStack comByteStack = new CommonByteStack();
                                        if (rangHasSameHighOrderValue)
                                        {
                                            // Check the commands before current one in DeserializedCommandList,
                                            comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, comByteStack);

                                            if (Common.IsRequirementEnabled(2661, this.Site))
                                            {
                                                // Verify requirement MS-OXCFXICS_R2661
                                                bool isR2661Satisfied = true;

                                                for (int index = 0; index < comByteStack.Bytes; index++)
                                                {
                                                    if (replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTRangeList[0].StartGLOBCNT[index] != comByteStack.GetCommonBytes()[index] || replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTRangeList[0].EndGLOBCNT[index] != comByteStack.GetCommonBytes()[index])
                                                    {
                                                        isR2661Satisfied = false;
                                                        break;
                                                    }
                                                }

                                                // Add the debug information
                                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2661");

                                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2661
                                                // StartGLOBCNT and EndGLOBCNT are the low and high GLOBCNT values. 
                                                // They will share the high-order bytes which are in the common byte stack.
                                                Site.CaptureRequirementIfIsTrue(
                                                    isR2661Satisfied,
                                                    2661,
                                                    @"[In Range Command (0x52)] Implementation does push these common high-order bytes, that are shared by the low and high GLOBCNT values, onto the common byte stack by using the Push command, as specified in section 2.2.2.6.1, prior to using the Range command. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                            }

                                            // Add the debug information
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3074");

                                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3074
                                            Site.CaptureRequirementIfAreEqual<short>(
                                                0x52,
                                                Convert.ToInt16((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).CommandByte),
                                                3074,
                                                @"[In Range Command (0x52)] Command (1 byte): This value MUST be set to 0x52.");
                                        }

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R213");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R213
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            6 - comByteStack.Bytes,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue.Length,
                                            213,
                                            @"[In Range Command (0x52)] [LowValue (variable):] The number of bytes in this field [LowValue] is equal to 6 minus the number of high-order bytes in the common byte stack.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R216");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R216
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            6 - comByteStack.Bytes,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue.Length,
                                            216,
                                            @"[In Range Command (0x52)] [HighValue (variable):] The number of bytes in this field [HighValue] is equal to 6 minus the number of high-order bytes in the common byte stack.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R214");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R214
                                        bool isVerifyR214 = true;
                                        for (int j = 0; j < 6 - comByteStack.Bytes; j++)
                                        {
                                            if ((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue[j] > (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue[j])
                                            {
                                                break;
                                            }
                                            else if ((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue[j] == (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue[j])
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                isVerifyR214 = false;
                                                break;
                                            }
                                        }

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR214,
                                            214,
                                            @"[In Range Command (0x52)] [LowValue (variable):] This value MUST be less than or equal to the value of the HighValue field.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R217");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R217
                                        bool isVerifyR217 = isVerifyR214;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR217,
                                            217,
                                            @"[In Range Command (0x52)] [HighValue (variable):] This value MUST be greater than or equal to the value of the LowValue field.");

                                        // When deserialize the FastTransfer Stream from the server response will as this requirement say,
                                        // if deserialize successfully, this method will be called , this requirement will be covered.
                                        Site.CaptureRequirement(
                                            1980,
                                            @"[In Range Command (0x52)]If the low and high value of the GLOBCNT range are not the same, the Range command MUST be used.");

                                        // When deserialize the FastTransfer Stream from the server response will as this requirement say,
                                        // if deserialize successfully, this method will be called , this requirement will be covered.
                                        Site.CaptureRequirement(
                                            1981,
                                            @"[In Range Command (0x52)]If the range has values that are more than 8 bytes from each other, the Range command MUST be used.");

                                        // When deserialize the FastTransfer Stream from the server response will as this requirement say,
                                        // if deserialize successfully, this method will be called , this requirement will be covered.
                                        Site.CaptureRequirement(
                                            1982,
                                            @"[In Range Command (0x52)]If the low and high value do not share five high-order bytes in common, the Range command MUST be used.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2009");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2009
                                        bool isVerifyR2009 = true;
                                        foreach (GLOBCNTRange gcntR in (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).CorrespondingGLOBCNTRangeList)
                                        {
                                            if (!replguid.GLOBSET.GLOBCNTRangeList.Contains(gcntR))
                                            {
                                                isVerifyR2009 = false;
                                                break;
                                            }
                                        }

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR2009,
                                            2009,
                                            @"[In Range Command (0x52)]The GLOBCNT structure range MUST be added to the GLOBSET structure.");

                                        byte[] lowValueOfRange = StructureSerializer.Serialize((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).CorrespondingGLOBCNTRangeList[0].StartGLOBCNT);
                                        byte[] highValueOfRange = StructureSerializer.Serialize((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).CorrespondingGLOBCNTRangeList[0].EndGLOBCNT);

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2012");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2012
                                        bool isVerifyR2012 = true;
                                        for (int index = 1; index <= (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue.Length; index++)
                                        {
                                            if ((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue[(replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue.Length - index] != lowValueOfRange[lowValueOfRange.Length - index]
                                             || (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue[(replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue.Length - index] != highValueOfRange[highValueOfRange.Length - index])
                                            {
                                                isVerifyR2012 = false;
                                                break;
                                            }
                                        }

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR2012,
                                            2012,
                                            @"[In Range Command (0x52)]Each of these fields MUST be combined with any high-order bytes in the common byte stack to produce a 6-byte GLOBCNT structure.");
                                    }

                                    if (replguid.GLOBSET.DeserializedCommandList[i] is BitmaskCommand)
                                    {
                                        // comByteStack is the common byte stack
                                        CommonByteStack comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, new CommonByteStack());

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1970");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1970
                                        bool isVerifyR1970 = comByteStack.Bytes == 5;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR1970,
                                            1970,
                                            @"[In Bitmask Command (0x42)]There MUST already be five high-order bytes in the common byte stack to use this command.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2007");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2007
                                        // "globcnt" is a variable, it is constituted by the "low-order" and "five high-order bytes",
                                        // so if GLOBSET contains "globcnt", then this requirement can be covered.
                                        GLOBCNT globcnt = new GLOBCNT
                                        {
                                            Byte1 = comByteStack.GetCommonBytes()[0],
                                            Byte2 = comByteStack.GetCommonBytes()[1],
                                            Byte3 = comByteStack.GetCommonBytes()[2],
                                            Byte4 = comByteStack.GetCommonBytes()[3],
                                            Byte5 = comByteStack.GetCommonBytes()[4],
                                            Byte6 = (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue
                                        };

                                        bool isVerifyR2007 = false;
                                        byte start_value = (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue;
                                        byte bitmask = (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).Bitmask;

                                        // Start value combined with the required five high-order bytes
                                        // on the common byte stack must be included in the set.
                                        if (replguid.GLOBSET.GLOBCNTList.Contains(globcnt))
                                        {
                                            int j;

                                            // Generates low-order byte, combined with the required five high-order bytes
                                            // on the common byte stack
                                            for (j = 0; j < 8; j++)
                                            {
                                                if ((bitmask & (1 << j)) != 0)
                                                {
                                                    globcnt.Byte6 = (byte)(start_value + j + 1);
                                                    if (!replguid.GLOBSET.GLOBCNTList.Contains(globcnt))
                                                    {
                                                        break;
                                                    }
                                                }
                                            }

                                            // no break j equals to 8.if break one low-order byte MUST be combined with the required five high-order bytes
                                            // on the common byte stack is not in the set.
                                            isVerifyR2007 = j == 8;
                                        }

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR2007,
                                            2007,
                                            @"[In Bitmask Command (0x42)]Each low-order byte MUST be combined with the required five high-order bytes on the common
                                        byte stack to form a complete 6-byte GLOBCNT structure, as specified in section 2.2.2.5, which MUST be added to the
                                        GLOBSET structure.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1973");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1973
                                        Site.CaptureRequirementIfAreEqual<byte>(
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).CorrespondingGLOBCNTList[0].Byte6,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue,
                                            1973,
                                            @"[In Bitmask Command (0x42)]The StartingValue field, as specified in section 2.2.2.6.3, MUST be set to the low-order byte of the low value of the first GLOBCNT range.");

                                        byte byt = 0;
                                        foreach (GLOBCNT gcnt in (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).CorrespondingGLOBCNTList)
                                        {
                                            byt |= (byte)(1 << ((int)gcnt.Byte6 - (int)(replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue - 1));
                                        }

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1974");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1974
                                        Site.CaptureRequirementIfAreEqual<byte>(
                                            byt,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).Bitmask,
                                            1974,
                                            @"[In Bitmask Command (0x42)]The Bitmask field, specified in section 2.2.2.6.3, MUST have one bit set for each value within a range, excluding the low value of the first GLOBCNT range.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1978");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1978
                                        Site.CaptureRequirementIfAreEqual<byte>(
                                            byt,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).Bitmask,
                                            1978,
                                            @"[In Bitmask Command (0x42)] For all GLOBCNTs between ranges, the bit associated with the value is not set in the bitmask.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3071");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3071
                                        Site.CaptureRequirementIfAreEqual<short>(
                                            0x42,
                                            Convert.ToInt16((replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).CommandByte),
                                            3071,
                                            @"[In Bitmask Command (0x42)] Command (1 byte): This value MUST be set to 0x42.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R207");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R207
                                        Site.CaptureRequirementIfAreEqual<Type>(
                                            typeof(byte),
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue.GetType(),
                                            207,
                                            @"[In Bitmask Command (0x42)] StartingValue (1 byte): The low-order byte of the first GLOBCNT structure, as specified in section 2.2.2.5.");
                                    }

                                    if (replguid.GLOBSET.DeserializedCommandList[i] is EndCommand)
                                    {
                                        CommonByteStack comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, new CommonByteStack());

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1992");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1992
                                        bool isVerifyR1992 = comByteStack.Bytes == 0;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR1992,
                                            1992,
                                            @"[In End Command (0x00)]The End command can only be used if the common byte stack is empty.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1989");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1989
                                        // When de-serializing the FastTransfer Stream from the server response according to this requirement description,
                                        // so that if the de-serialization is successful, only need to verify that there is no byte on the common byte stack when "End command" is used.
                                        bool isVerifyR1989 = comByteStack.Bytes == 0;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR1989,
                                            1989,
                                            @"[In End Command (0x00)]If after all GLOBCNT ranges have been encoded, there are still bytes on the common byte stack, 
                                        they MUST be removed with one or more Pop commands, as specified in section 2.2.2.6.1, before the End command can be
                                        used.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2014");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2014
                                        // EndCommand must be the last Command in GLOBSET.
                                        bool isVerifyR2014 = i == replguid.GLOBSET.DeserializedCommandList.Count - 1;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR2014,
                                            2014,
                                            @"[In End Command (0x00)]When the End command, as specified in section 2.2.2.6.5, is encountered, the GLOBSET structure
                                        MUST be complete based on the GLOBCNT structure values, as specified in section 2.2.2.5, generated from any previous
                                        encoding commands.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1987");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1987
                                        // EndCommand must be the last Command in GLOBSET.
                                        bool isVerifyR1987 = i == replguid.GLOBSET.DeserializedCommandList.Count - 1;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR1987,
                                            1987,
                                            @"[In End Command (0x00)]This command MUST be added after all GLOBCNT structure ranges, as specified in section 2.2.2.5,
                                        within the GLOBSET have been encoded.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3077");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3077
                                        Site.CaptureRequirementIfAreEqual<short>(
                                            0x00,
                                            Convert.ToInt16((replguid.GLOBSET.DeserializedCommandList[i] as EndCommand).CommandByte),
                                            3077,
                                            @"[In End Command (0x00)] Command (1 byte): This value MUST be set to 0x00.");
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // The FastTransfer Stream from the server response will be deserialized as this requirement,
            // if deserialize successfully this requirement will be covered
            Site.CaptureRequirement(
                2026,
                @"[In  Creating Compact IDSET Structures]Note that this MUST NOT be done for IDSET structures that are ever iterated over, such as the
                MetaTagIdsetGiven property (section 2.2.1.1.1), as it will change the outcome.");
            #endregion

            if (conSync.ProgressTotal != null)
            {
                this.VerifyProgressTotal(conSync);
            }

            if (conSync.Deletions != null)
            {
                this.VerifyDeletions(conSync);
            }

            if (conSync.MessageChangeTuples != null && conSync.MessageChangeTuples.Count > 0)
            {
                this.VerifyMessageChangeAndProgressPerMessage(conSync);
            }

            if (conSync.ReadStateChanges != null)
            {
                this.VerifyReadStateChanges(conSync);
            }
        }

        /// <summary>
        /// Verify the subelelment progressTotal under contentsSync element. 
        /// </summary>
        /// <param name="contentSync">The contentsSync element of fastTransfer Stream.</param>
        private void VerifyProgressTotal(ContentsSync contentSync)
        {            
                foreach (PropValue prop in contentSync.ProgressTotal.PropList.PropValues)
                {
                    VarPropTypePropValue varProp = (VarPropTypePropValue)prop;
                    using (MemoryStream memoryStream = new MemoryStream(varProp.ValueArray))
                    {
                        ProgressInformation progressInformation = new ProgressInformation();
                        progressInformation.Deserialize(memoryStream, -1);

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1388");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1388
                        Site.CaptureRequirementIfIsNotNull(
                            progressInformation,
                            1388,
                            @"[In progressTotal Element] [[PtypBinary] 0x00000102] [PtypBinary] 0x00000102 is a Serialized ProgressInformation structure.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R220");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R220
                        // ushort is unsigned 16-bit
                        Site.CaptureRequirementIfAreEqual<Type>(
                            typeof(ushort),
                            progressInformation.Version.GetType(),
                            220,
                            @"[In ProgressInformation Structure]Version (2 bytes):An unsigned 16-bit value.");

                        // Verify requirement MS-OXCFXICS_R224.
                        bool isR224Satisfied = progressInformation.Padding1 == 0;

                        if (Common.IsRequirementEnabled(224, this.Site))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R224");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R224
                            Site.CaptureRequirementIfIsTrue(
                                isR224Satisfied,
                                224,
                                @"[In ProgressInformation Structure] Implementation does set the padding (2 bytes) value to 0x0000 in ProgressInformation Structure. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R226");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R226
                        // uint is an unsigned 32-bit integer
                        Site.CaptureRequirementIfAreEqual<Type>(
                            typeof(uint),
                            progressInformation.FAIMessageCount.GetType(),
                            226,
                            @"[In ProgressInformation Structure]FAIMessageCount (4 bytes): An unsigned 32-bit integer value.");

                        if (this.importFlag == ImportFlag.Associated)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R227");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R227
                            // ulong is an unsigned 64-bit integer
                            Site.CaptureRequirementIfAreEqual<uint>(
                                1,
                                progressInformation.FAIMessageCount,
                                227,
                                @"[In ProgressInformation Structure] [FAIMessageCount (4 bytes)] contains the total number of changes to FAI messages that are scheduled for download during the current synchronization operation.");
                        }
                        else if (this.importFlag == ImportFlag.Normal)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R231");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R231
                            Site.CaptureRequirementIfAreEqual<uint>(
                                1,
                                progressInformation.NormalMessageCount,
                                231,
                                @"[In ProgressInformation Structure] [NormalMessageCount (4 bytes)] contains the total number of changes to normal messages 
                        that are scheduled for download during the current synchronization operation.");
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R228");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R228
                        // ulong is an unsigned 64-bit integer
                        Site.CaptureRequirementIfAreEqual<Type>(
                            typeof(ulong),
                            progressInformation.FAIMessageTotalSize.GetType(),
                            228,
                            @"[In ProgressInformation Structure]FAIMessageTotalSize (8 bytes):An unsigned 64-bit integer value.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R230");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R230
                        // uint is an unsigned 32-bit integer
                        Site.CaptureRequirementIfAreEqual<Type>(
                            typeof(uint),
                            progressInformation.NormalMessageCount.GetType(),
                            230,
                            @"[In ProgressInformation Structure]NormalMessageCount (4 bytes):An unsigned 32-bit integer value.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R236");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R236
                        // ulong is an unsigned 64-bit integer
                        Site.CaptureRequirementIfAreEqual<Type>(
                            typeof(ulong),
                            progressInformation.NormalMessageTotalSize.GetType(),
                            236,
                            @"[In ProgressInformation Structure]NormalMessageTotalSize (8 bytes):An unsigned 64-bit integer value.");

                        // Verify requirement MS-OXCFXICS_R233
                        bool isR233Satisfied = progressInformation.Padding2 == 0;

                        if (Common.IsRequirementEnabled(233, this.Site))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R233");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R233
                            Site.CaptureRequirementIfIsTrue(
                                isR233Satisfied,
                                233,
                                @"[In ProgressInformation Structure] Implementation does set the padding (4 bytes) value to 0x00000000 in ProgressInformation Structure. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                        }

                        if (prop.PropInfo.PropID == 0x0000 && prop.PropType == 0x0102)
                        {
                            // If prop.PropInfo.PropID == 0x0000 and prop.PropType==0x0102), this PropValue is [PtypBinary] 0x00000102
                            Site.CaptureRequirement(
                                1386,
                                @"[In progressTotal Element] [[PtypBinary] 0x00000102 ,Required] The property [PtypBinary] 0x00000102 MUST be present in the propList.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1387");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1387
                            // The position of the property [PtypBinary] 0x00000102 is 1
                            bool isVerifyR1387 = prop == contentSync.ProgressTotal.PropList.PropValues[0];

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1387,
                                1387,
                                @"[In progressTotal Element] [[PtypBinary] 0x00000102 ,Fixed position] The position[1] of the property [PtypBinary] 0x00000102 within the array propList is fixed and MUST correspond to the position of the corresponding restriction in the property list restriction table.");
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R222");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R222
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0000,
                            progressInformation.Version,
                            222,
                            @"[In ProgressInformation Structure] [Version (2 bytes)] The preceding packet diagram specifies the format for version 0x0000, which is the only version of this structure defined for this protocol.");
                    }
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1390");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1390
                bool isVerifyR1390 = true;
                foreach (PropValue prop in contentSync.ProgressTotal.PropList.PropValues)
                {
                    if (prop.PropInfo.PropID != 0x0000 || prop.PropType != 0x0102)
                    {
                        isVerifyR1390 = false;
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1390,
                    1390,
                    @"[In progressTotal Element,other properties, Prohibited] The other properties [the properties except [PtypBinary] 0x00000102] MUST NOT be present in the propList.");
        }

         /// <summary>
        /// Verify the subelelments messageChange and ProgressPerMessage under contentsSync element. 
        /// </summary>
        /// <param name="contentSync">The contentsSync element of fastTransfer Stream.</param>
        private void VerifyMessageChangeAndProgressPerMessage(ContentsSync contentSync)
        {
            #region Verify ProgressPerMessage
                for (int i = 0; i < contentSync.MessageChangeTuples.Count; i++)
                {
                    if (null != contentSync.MessageChangeTuples[i].Item1)
                    {
                        ProgressPerMessage progPerMessage = (ProgressPerMessage)contentSync.MessageChangeTuples[0].Item1;
                        foreach (PropValue prop in progPerMessage.PropList.PropValues)
                        {
                            if (prop.PropInfo.PropID == 0x0000 && prop.PropType == 0x0003)
                            {
                                // If prop.PropInfo.PropID is 0x0000 and prop.PropType is 0x0003, this property value is [PtypInteger32] 0x00000003
                                Site.CaptureRequirement(
                                    1376,
                                    @"[In progressPerMessage Element] [[PtypInteger32] 0x00000003,Required] The property [PtypInteger32] 0x00000003 MUST be present in the propList.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1377");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1377
                                // The position of the property [PtypInteger32] 0x00000003 is 1
                                Site.CaptureRequirementIfAreEqual<PropValue>(
                                    progPerMessage.PropList.PropValues[0],
                                    prop,
                                    1377,
                                    @"[In progressPerMessage Element] [[PtypInteger32] 0x00000003 ,Fixed position] The position[1] of the property [PtypInteger32] 0x00000003 within the array propList is fixed and MUST correspond to the position of the corresponding restriction in the property list restriction table.");
                            }

                            if (prop.PropInfo.PropID == 0x0000 && prop.PropType == 0x000B)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R989");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R989
                                bool isVerifyR989 = (prop as FixedPropTypePropValue) != null;

                                // PtypBoolean value is fixedSizeValue, PtypBoolean value  is serialized as this requirement.
                                // if this value can be retrieved, this requirement can be covered.
                                Site.CaptureRequirementIfIsTrue(
                                    isVerifyR989,
                                    989,
                                    @"[In propValue Lexical Element] The fixedSizeValue or varSizeValue lexemes contained in a propValue element represent the value of the property and MUST be serializations of a base property type for a property type specified with contained fixedPropType, varPropType, or mvPropType property type values.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R990");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R990
                                Site.CaptureRequirementIfAreEqual<int>(
                                    2,
                                    Marshal.SizeOf((prop as FixedPropTypePropValue).FixedValue),
                                    990,
                                    @"[In Serialization of Simple Types] Serialization of simple types in FastTransfer streams is identical to serialization of property values as specified [MS-OXCDATA] , with the following exceptions:The PtypBoolean ([MS-OXCDATA] section 2.11.1) type is 2-bytes in FastTransfer streams, instead of 1-byte as specified in [MS-OXCDATA] .");

                                // If prop.PropInfo.PropID == 0x0000 and prop.PropType == 0x0003, this property value is [PtypBoolean] 0x0000000B
                                Site.CaptureRequirement(
                                    1380,
                                    @"[In progressPerMessage Element] [[PtypBoolean] 0x0000000B,Required] The property [PtypBoolean] 0x0000000B MUST be present in the propList.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1381");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1381
                                // The position of the property [PtypBoolean] 0x0000000B 2
                                Site.CaptureRequirementIfAreEqual<PropValue>(
                                    progPerMessage.PropList.PropValues[1],
                                    prop,
                                    1381,
                                    @"[In progressPerMessage Element] [[PtypBoolean] 0x0000000B,Fixed position] The position[2] of the property [PtypBoolean] 0x0000000B within the array propList is fixed and MUST correspond to the position of the corresponding restriction in the property list restriction table.");
                            }
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1385");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1385
                        bool isVerifyR1385 = true;
                        foreach (PropValue prop in progPerMessage.PropList.PropValues)
                        {
                            if (prop.PropInfo.PropID != 0x0000)
                            {
                                isVerifyR1385 = false;
                            }
                        }

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR1385,
                            1385,
                            @"[In progressPerMessage Element] [other properties,Prohibited] The other properties [the properties except [PtypInteger32] 0x00000003 and [PtypBoolean] 0x0000000B] MUST NOT be present in the propList.");
                    }
                }
                #endregion Verify ProgressPerMessage

            // Verify MessageChangePartial
            if (contentSync.MessageChangeTuples[0].Item2 is MessageChangePartial)
            {
                MessageChangePartial msgChangePartial = (MessageChangePartial)contentSync.MessageChangeTuples[0].Item2;

                if (Common.IsRequirementEnabled(3305001, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3305001");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3305001
                    Site.CaptureRequirementIfIsNotNull(
                        msgChangePartial,
                        3305001,
                        @"[In messageChangePartial Element] Implementation does support partial item downloads. (<20> Section 2.2.4.3.15: Microsoft Exchange Server 2007 and above follow this behavior.)");
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_1155");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_1155
                // MUST be used means not null
                Site.CaptureRequirementIfIsNotNull(
                    msgChangePartial.IncrSyncGroupId,
                    1155,
                    @"[In messageChangePartial Element]The value of the last encountered MetaTagIncrSyncGroupId meta-property (section 2.2.4.1.5.4) in 
                    the stream determines which property group mapping MUST be used to make partial updates to the messaging object.");

                if (msgChangePartial.IncrSyncGroupId != null)
                {
                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_3289
                    // If IncrSyncGroupId of msgChangePartial is not null, which means the MetaTagIncrSyncGroupId property is deserialized successfully according to MetaTagIncrSyncGroupId(0x407c0003) is true. 
                    // So this requirement can be captured directly.
                    Site.CaptureRequirement(
                        3289,
                        @"[In MetaTagIncrSyncGroupId Meta-Property] Property ID: 0x407C Data type: PtypInteger32, 0x0003 ([MS-OXCDATA] section 2.11.1)");
                }

                if (msgChangePartial.IncrementalSyncMessagePartial != null)
                {
                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_3290
                    // If IncrementalSyncMessagePartial of msgChangePartial is not null, which means the MetaTagIncrementalSyncMessagePartial property is deserialized successfully according to MetaTagIncrSyncGroupId(0x407a0003) is true. 
                    // So this requirement can be captured directly.
                    Site.CaptureRequirement(
                        3290,
                        @"[In MetaTagIncrementalSyncMessagePartial Meta-Property] Property ID: 0x407A, Data type: PtypInteger32, 0x0003 ([MS-OXCDATA] section 2.11.1)");
                }

                #region Verify GroupInfo
                if (msgChangePartial.GroupInfo != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1123");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1123
                    bool isVerifyR1123 = true;
                    bool isVerifyR1122 = true;

                    foreach (PropValue prop in msgChangePartial.GroupInfo.PropList.PropValues)
                    {
                        if (prop.PropType != 0x0102)
                        {
                            isVerifyR1122 = false;
                        }

                        VarPropTypePropValue varProp = (VarPropTypePropValue)prop;
                        using (MemoryStream memoryStream = new MemoryStream(varProp.ValueArray))
                        {
                            PropertyGroupInfo propertyGroupInfo = new PropertyGroupInfo();
                            propertyGroupInfo.Deserialize(memoryStream, -1);

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R246");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R246
                            Site.CaptureRequirement(
                                246,
                                @"[In PropertyGroupInfo Structure] Groups (variable): An array of PropertyGroup structures, as specified in section 2.2.2.8.1.");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R244");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R244
                            this.Site.CaptureRequirementIfAreEqual<uint>(
                                propertyGroupInfo.GroupCount,
                                (uint)propertyGroupInfo.Groups.Length,
                                244,
                                @"[In PropertyGroupInfo Structure] [GroupCount (4 bytes)] specifies how many PropertyGroup structures, as specified in section 2.2.2.8.1, are present in the Groups field.");

                            for (int i = 0; i < propertyGroupInfo.Groups.Length; i++)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R249");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R249
                                Site.CaptureRequirementIfAreEqual<uint>(
                                    (uint)propertyGroupInfo.Groups[i].PropertyTags.Length,
                                    propertyGroupInfo.Groups[i].PropertyTagCount,
                                    249,
                                    @"[In PropertyGroup Structure] [PropertyTagCount (4 bytes)] specifies how many PropertyTag structures are present in the PropertyTags field.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R251");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R251
                                bool isVerifyR251 = true;
                                for (int j = 0; j < propertyGroupInfo.Groups[i].PropertyTags.Length; j++)
                                {
                                    if (propertyGroupInfo.Groups[i].PropertyTags[j].Item1.GetType() != typeof(PropertyTag))
                                    {
                                        isVerifyR251 = false;
                                    }
                                }

                                Site.CaptureRequirementIfIsTrue(
                                    isVerifyR251,
                                    251,
                                    @"[In PropertyGroup Structure] PropertyTags (variable): An array of PropertyTag structures ([MS-OXCDATA] section 2.9).");

                                for (int j = 0; j < propertyGroupInfo.Groups[i].PropertyTags.Length; j++)
                                {
                                    if (propertyGroupInfo.Groups[i].PropertyTags[j].Item2 != null)
                                    {
                                        if (propertyGroupInfo.Groups[i].PropertyTags[j].Item2.NameSize != null)
                                        {
                                            // Add the debug information
                                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R253");

                                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R253
                                            bool isVerifiedR253 = propertyGroupInfo.Groups[i].PropertyTags[j].Item1.PropertyId >= 0x8000;

                                            this.Site.CaptureRequirementIfIsTrue(
                                                isVerifiedR253,
                                                253,
                                                @"[In PropertyGroup Structure] [PropertyTags (variable)] Named properties are identified by a PropertyId structure ([MS-OXCDATA] section 2.9) with a value greater than or equal to 0x8000.");

                                            // Add the debug information
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R263");

                                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R263
                                            // Open Specification specified that Name (optional, variable): A Unicode (UTF-16) string that identifies the property within the property set, but not UTF-8.
                                            // So here using the length of Name field multiple 2 to compare the length. If they are equal, the requirement can be captured.
                                            Site.CaptureRequirementIfAreEqual<uint>(
                                                (uint)(propertyGroupInfo.Groups[i].PropertyTags[j].Item2.Name.Length * 2),
                                                (uint)propertyGroupInfo.Groups[i].PropertyTags[j].Item2.NameSize,
                                                263,
                                                @"[In GroupPropertyName Structure] NameSize (optional) (4 bytes): A value that specifies the length of the Name field, in bytes.");

                                            // Add the debug information
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3670");

                                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3670
                                            // Open Specification specified that Name (optional, variable): A Unicode (UTF-16) string that identifies the property within the property set, but not UTF-8.
                                            // So here using the length of Name field multiple 2 to compare the length. If they are equal, the requirement can be captured. 
                                            Site.CaptureRequirementIfAreEqual<uint>(
                                                (uint)propertyGroupInfo.Groups[i].PropertyTags[j].Item2.NameSize,
                                                (uint)(propertyGroupInfo.Groups[i].PropertyTags[j].Item2.Name.Length * 2),
                                                3670,
                                                @"[In GroupPropertyName Structure] Name (optional, variable): The length of this field is specified by the value of the NameSize field, in bytes.");

                                            // Add the debug information
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R26201");

                                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R26201
                                            // Capture this requirement directly since the Name had been deserialized as char array
                                            Site.CaptureRequirement(
                                                26201,
                                                @"[In GroupPropertyName Structure] Name (optional) (4 bytes): A Unicode (UTF-16) string. ");
                                        }

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R256");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R256
                                        bool isVerifyR256 = propertyGroupInfo.Groups[i].PropertyTags[j].Item2.Kind == 0x00000000 || propertyGroupInfo.Groups[i].PropertyTags[j].Item2.Kind == 0x00000001;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR256,
                                            256,
                                            @"[In GroupPropertyName Structure] Kind (4 bytes): The following table lists the possible values [0x00000000 and 0x00000001] for the Kind field.");

                                        if (propertyGroupInfo.Groups[i].PropertyTags[j].Item2.Kind == 0x00000000)
                                        {
                                            // Add the debug information
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R259");

                                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R259
                                            Site.CaptureRequirementIfIsNotNull(
                                                propertyGroupInfo.Groups[i].PropertyTags[j].Item2.LID,
                                                259,
                                                @"[In GroupPropertyName Structure]LID (optional) (4 bytes): This value is present only if the Kind field is set to 0x00000000.");

                                            // Add the debug information
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R257");

                                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R257
                                            Site.CaptureRequirementIfIsNotNull(
                                                propertyGroupInfo.Groups[i].PropertyTags[j].Item2.LID,
                                                257,
                                                @"[In GroupPropertyName Structure] [Kind (4 bytes):] The following are possible values for the Kind field] When the Kind value is 0x00000000, the property is identified by the LID field.");
                                        }
                                        else
                                        {
                                            Site.Assert.IsTrue(
                                                propertyGroupInfo.Groups[i].PropertyTags[j].Item2.LID == null,
                                                @"[In GroupPropertyName Structure]LID (optional) (4 bytes): This value is present only if the Kind field is set to 0x00000000.");
                                        }

                                        if (propertyGroupInfo.Groups[i].PropertyTags[j].Item2.Kind == 0x00000001)
                                        {
                                            // Add the debug information
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R262");

                                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R262
                                            Site.CaptureRequirementIfIsNotNull(
                                                propertyGroupInfo.Groups[i].PropertyTags[j].Item2.NameSize,
                                                262,
                                                @"[In GroupPropertyName Structure]NameSize (optional) (4 bytes): This value is present only if the Kind field is set to 0x00000001.");

                                            // Add the debug information
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R264");

                                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R264
                                            Site.CaptureRequirementIfIsNotNull(
                                                propertyGroupInfo.Groups[i].PropertyTags[j].Item2.Name,
                                                264,
                                                @"[In GroupPropertyName Structure] Name (optional, variable): This value is present only if the Kind field is set to 0x00000001.");

                                            // Add the debug information
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R258");

                                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R258
                                            Site.CaptureRequirementIfIsNotNull(
                                                propertyGroupInfo.Groups[i].PropertyTags[j].Item2.Name,
                                                258,
                                                @"[In GroupPropertyName Structure] [Kind (4 bytes):] The following are possible values for the Kind field] When the Kind value is 0x00000001, the property is identified by the Name field.");
                                        }
                                        else
                                        {
                                            Site.Assert.IsTrue(
                                                propertyGroupInfo.Groups[i].PropertyTags[j].Item2.Name == null,
                                                @"[In GroupPropertyName Structure]NameSize (optional) (4 bytes): This value is present only if the Kind field is set to 0x00000001.");

                                            Site.Assert.IsTrue(
                                                propertyGroupInfo.Groups[i].PropertyTags[j].Item2.NameSize == null,
                                                @"[In GroupPropertyName Structure]Name (optional, variable): This value is present only if the Kind field is set to 0x00000001");
                                        }
                                    }
                                }
                            }

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1122");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1122
                            // uint is an unsigned 32-bit integer
                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1122,
                                1122,
                                @"[In groupInfo Element,[PtypBinary] 0x00000102] [PtypBinary] 0x00000102 is a serialized PropertyGroupInfo structure, as specified in section 2.2.2.8.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R240");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R240
                            // uint is an unsigned 32-bit integer
                            Site.CaptureRequirementIfAreEqual<Type>(
                                typeof(uint),
                                propertyGroupInfo.GroupId.GetType(),
                                240,
                                @"[In PropertyGroupInfo Structure]GroupId (4 bytes):An unsigned 32-bit integer value.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R242");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R242
                            Site.CaptureRequirementIfAreEqual<uint>(
                                0x00000000,
                                propertyGroupInfo.Reserved,
                                242,
                                @"[In PropertyGroupInfo Structure]Reserved (4 bytes):  This value MUST be set to 0x00000000.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R243");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R243
                            // uint is an unsigned 32-bit integer
                            Site.CaptureRequirementIfAreEqual<Type>(
                                typeof(uint),
                                propertyGroupInfo.GroupCount.GetType(),
                                243,
                                @"[In PropertyGroupInfo Structure]GroupCount (4 bytes):An unsigned 32-bit integer value.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R245");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R245
                            Site.CaptureRequirementIfAreNotEqual<uint>(
                                0x00000000,
                                propertyGroupInfo.GroupCount,
                                245,
                                @"[In PropertyGroupInfo Structure] [GroupCount (4 bytes)] This field MUST NOT be set to 0x00000000.");

                            // When deserialize the response from the server will as this requirement say, if deserialize successfully, this requirement can be covered
                            Site.CaptureRequirement(
                                247,
                                @"[In PropertyGroupInfo Structure] [Groups (variable)] The number of PropertyGroup structures in this value is specified by the value of the GroupCount field.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R248");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R248
                            // uint is an unsigned 32-bit integer
                            Site.CaptureRequirementIfAreEqual<Type>(
                                typeof(uint),
                                propertyGroupInfo.Groups[0].PropertyTagCount.GetType(),
                                248,
                                @"[In PropertyGroup Structure]PropertyTagCount (4 bytes):An unsigned 32-bit integer value.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R250");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R250
                            Site.CaptureRequirementIfAreNotEqual<uint>(
                                0,
                                propertyGroupInfo.Groups[0].PropertyTagCount,
                                250,
                                @"[In PropertyGroup Structure] [PropertyTagCount (4 bytes)] This value MUST NOT be set to 0x00000000.");

                            // When deserialize the response from the server will as this requirement say, if deserialize successfully, this requirement can be covered
                            Site.CaptureRequirement(
                                254,
                                @"[In PropertyGroup Structure] [PropertyTags (variable)] The number of PropertyTag structures in this field is specified by the value of the PropertyTagCount field.");

                            if (prop.PropInfo.PropID == 0x00)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1120");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1120
                                Site.CaptureRequirementIfAreEqual<ushort>(
                                    0x0102,
                                    prop.PropType,
                                    1120,
                                    @"[In groupInfo Element,[PtypBinary]0x00000102, Required]The property [PtypBinary] 0x00000102 MUST be present in the propList.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1121");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1121
                                // The position of the property [PtypBinary] 0x00000102 is 1
                                Site.CaptureRequirementIfAreEqual<PropValue>(
                                    msgChangePartial.GroupInfo.PropList.PropValues[0],
                                    prop,
                                    1121,
                                    @"[In groupInfo Element,[PtypBinary] 0x00000102 ,Fixed position] The position[1] of the property [PtypBinary] 0x00000102 within the array propList is fixed and MUST correspond to the position of the corresponding restriction in the property list restriction table.");
                            }

                            if (prop.PropType != 0x0102)
                            {
                                isVerifyR1122 = false;
                            }

                            if (prop.PropInfo.PropID != 0x00)
                            {
                                isVerifyR1123 = false;
                            }
                        }
                    }

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR1123,
                        1123,
                        @"[In groupInfo Element,other properties,Prohibited] The other properties [the properties except [PtypBinary] 0x00000102] MUST NOT be present in the propList.");
                }
            }
                #endregion

            #region Verify messageChangeHeader

            // Verify messageChangeHeader
            if (contentSync.MessageChangeTuples[0].Item2 is MessageChangeFull)
            {
                MessageChangeFull msgChangeFull = (MessageChangeFull)contentSync.MessageChangeTuples[0].Item2;

                if (msgChangeFull.MessageChildren.Recipients != null)
                {
                    foreach (Recipient recipient in msgChangeFull.MessageChildren.Recipients)
                    {
                        bool isVerifyR1391 = false;
                        foreach (PropValue propValue in recipient.PropList.PropValues)
                        {
                            // 0x3000 is the property id of PidTagRowid according [MS-OXPROPS]
                            if (propValue.PropInfo.PropID == 0x3000)
                            {
                                isVerifyR1391 = true;
                            }
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1391");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1391                        
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR1391,
                            1391,
                            @"[In recipient Element] [PidTagRowid,Required] [The following table lists the restrictions[Required,Fixed position] that exist on the contained propList element] The property PidTagRowid ([MS-OXPROPS] section 2.930) MUST be present in the propList.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1392");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1392
                        bool isVerifyR1392 = isVerifyR1391;
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR1392,
                            1392,
                            @"[In recipient Element] [PidTagRowid ,Fixed position] The position[1] of the property PidTagRowid ([MS-OXPROPS] section 2.1002) within the array propList is fixed and MUST correspond to the position of the corresponding restriction in the property list restriction table.");
                    }
                }

                foreach (PropValue prop in msgChangeFull.MessageChangeHeader.PropList.PropValues)
                {
                    if (prop.PropInfo.PropID == 0x65E0)
                    {
                        VarPropTypePropValue varProp = (VarPropTypePropValue)prop;
                        GID gid = new GID();
                        gid = StructureSerializer.Deserialize<GID>(varProp.ValueArray);

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2186
                        Site.CaptureRequirementIfAreEqual<int>(
                            16,
                            gid.DatabaseGuid.ToByteArray().Length,
                            "MS-OXCDATA",
                            2186,
                            @"[In GID] DatabaseGuid (16 bytes):  A 128-bit unsigned integer identifying a Store object.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2187");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2187
                        Site.CaptureRequirementIfAreEqual<int>(
                            6,
                            gid.GlobalCounter.Length,
                            "MS-OXCDATA",
                            2187,
                            @"[In GID] GlobalCounter (6 bytes):  An unsigned 48-bit integer identifying the folder within its Store object.");

                        // The property id of PidTagSourceKey is 0x65E0,
                        // if 0x65E0 is present, this requirement can be verified
                        Site.CaptureRequirement(
                            1352,
                            @"[In messageChangeHeader Element] [PidTagSourceKey ,Required] The property PidTagSourceKey MUST be present in the propList.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1353");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1353
                        // The position of the property PidTagSourceKey is 1
                        Site.CaptureRequirementIfAreEqual<PropValue>(
                            msgChangeFull.MessageChangeHeader.PropList.PropValues[0],
                            prop,
                            1353,
                            @"[In messageChangeHeader Element] [PidTagSourceKey,Fixed position] The position[1] of the property PidTagSourceKey within the array propList is fixed and MUST correspond to the position of the corresponding restriction in the property list restriction table.");
                    }

                    if (prop.PropInfo.PropID == 0x3008)
                    {
                        // The property id of PidTagLastModificationTime is 0x3008,
                        // if 0x3008 is present, this requirement can be verified
                        Site.CaptureRequirement(
                            1354,
                            @"[In messageChangeHeader Element] [PidTagLastModificationTime,Required] The property PidTagLastModificationTime MUST be present in the propList.");

                        // The property id of PidTagLastModificationTime is 0x3008,
                        Site.CaptureRequirement(
                            "MS-OXPROPS",
                            6783,
                            @"[In PidTagLastModificationTime] Property ID: 0x3008.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6784");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R6784
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0040,
                            prop.PropType,
                            "MS-OXPROPS",
                            6784,
                            @"[In PidTagLastModificationTime] Data type: PtypTime, 0x0040.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1355");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1355
                        // The position of the property PidTagLastModificationTime is 2
                        Site.CaptureRequirementIfAreEqual<PropValue>(
                            msgChangeFull.MessageChangeHeader.PropList.PropValues[1],
                            prop,
                            1355,
                            @"[In messageChangeHeader Element] [PidTagLastModificationTime,Fixed position] The position[2] of the property PidTagLastModificationTime within the array propList is fixed and MUST correspond to the position of the corresponding restriction in the property list restriction table.");
                    }

                    if (prop.PropInfo.PropID == 0x65E2)
                    {
                        VarPropTypePropValue varProp = (VarPropTypePropValue)prop;
                        using (MemoryStream memoryStream = new MemoryStream(varProp.ValueArray))
                        {
                            XID xid = new XID();
                            xid.Deserialize(memoryStream, -1);

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R153");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R153
                            Site.CaptureRequirementIfAreEqual<Type>(
                                typeof(Guid),
                                xid.NamespaceGuid.GetType(),
                                153,
                                @"[In XID Structure] NamespaceGuid (16 bytes): A 128-bit GUID.");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R155                     
                            Site.CaptureRequirementIfAreEqual<Type>(
                                typeof(byte[]),
                                xid.LocalId.GetType(),
                                155,
                                @"[In XID Structure] LocalId (variable): A variable binary value.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R157");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R157
                            bool isVerifyR157 = xid.LocalId.Length >= 1 && xid.LocalId.Length <= 239;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR157,
                                157,
                                @"[In XID Structure] LocalId (variable): This field has a minimum length of 1 byte and a maximum length of 8 bytes.");

                            // The property id of PidTagChangeKey is 0x65E2,
                            // if 0x65E2 is present, this requirement can be verified
                            Site.CaptureRequirement(
                                1356,
                                @"[In messageChangeHeader Element] [PidTagChangeKey,Required] The property PidTagChangeKey MUST be present in the propList.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1357");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1357
                            // The position of the property PidTagChangeKey is 3
                            Site.CaptureRequirementIfAreEqual<PropValue>(
                                msgChangeFull.MessageChangeHeader.PropList.PropValues[2],
                                prop,
                                1357,
                                @"[In messageChangeHeader Element] [PidTagChangeKey,Fixed position] The position[3] of the property PidTagChangeKey within the array propList is fixed and MUST correspond to the position of the corresponding restriction in the property list restriction table.");
                        }
                    }

                    if (prop.PropInfo.PropID == 0x65E3)
                    {
                        VarPropTypePropValue varProp = (VarPropTypePropValue)prop;
                        using (MemoryStream memoryStream = new MemoryStream(varProp.ValueArray))
                        {
                            PredecessorChangeList predecessorChangeList = new PredecessorChangeList();
                            predecessorChangeList.Deserialize(memoryStream, -1);

                            // When deserialize the response from the server will as this requirement say, if deserialize successfully, this requirement can be covered
                            Site.CaptureRequirement(
                                169,
                                @"[In PredecessorChangeList Structure]The set of XIDs MUST be serialized without padding as an array of SizedXid structures 
                            binary-sorted by the value of NamespaceGuid field of the XID structure in the ascending order.");

                            // If deserialize successfully, this requirement can be covered
                            Site.CaptureRequirement(
                                165,
                                @"[In PredecessorChangeList Structure] The PredecessorChangeList structure contains a set of XID structures, as specified in section 2.2.2.2, that identify change numbers of messaging objects in different replicas (1).");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2737
                            // When deserialize buffer this requirement will be verified and get a boolean value to indicates whether this requirement be test.
                            bool isVerifyR2737 = predecessorChangeList.IsSameGuidHasSameLengthXID;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR2737,
                                2737,
                                @"[In XID Structure] XIDs in PredecessorChangeList with the same NamespaceGuid MUST have the same length of LocalId fields.");

                            if (null != predecessorChangeList.SizedXidList)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R170");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R170
                                Site.CaptureRequirementIfAreEqual<Type>(
                                    typeof(byte),
                                    predecessorChangeList.SizedXidList[0].XidSize.GetType(),
                                    170,
                                    @"[In SizedXid Structure] XidSize (1 byte): An unsigned 8-bit integer.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R171");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R171
                                // the XID field contain a 16 byte NamespaceGuid and a variable byte [] LocalId
                                bool isVerifyR171 = predecessorChangeList.SizedXidList[0].XidSize == (predecessorChangeList.SizedXidList[0].XID.LocalId.Length + 16);

                                Site.CaptureRequirementIfIsTrue(
                                    isVerifyR171,
                                    171,
                                    @"[In SizedXid Structure] XidSize (1 byte): that specifies the length of the XID field, in bytes.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R173");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R173
                                // this requirement is same as R171
                                bool isVerifyR173 = isVerifyR171;

                                Site.CaptureRequirementIfIsTrue(
                                    isVerifyR173,
                                    173,
                                    @"[In SizedXid Structure] XID (variable): The length of this field is specified by the XidSize field, in bytes.");
                            }
                        }

                        // The property id of PidTagPredecessorChangeList is 0x65E3,
                        Site.CaptureRequirement(
                            "MS-OXPROPS",
                            7518,
                            @"[In PidTagPredecessorChangeList] Property ID: 0x65E3.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7519");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R7519
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0102,
                            prop.PropType,
                            "MS-OXPROPS",
                            7519,
                            @"[In PidTagPredecessorChangeList] Data type: PtypBinary, 0x0102.");

                        // The property id of PidTagPredecessorChangeList is 0x65E3,
                        // if 0x65E3 is present, this requirement can be verified
                        Site.CaptureRequirement(
                            1358,
                            @"[In messageChangeHeader Element] [PidTagPredecessorChangeList,Required] The property PidTagPredecessorChangeList MUST be present in the propList.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1359");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1359
                        // The position of the property PidTagPredecessorChangeList is 4
                        Site.CaptureRequirementIfAreEqual<PropValue>(
                            msgChangeFull.MessageChangeHeader.PropList.PropValues[3],
                            prop,
                            1359,
                            @"[In messageChangeHeader Element] [PidTagPredecessorChangeList,Fixed position] The position[4] of the property PidTagPredecessorChangeList within the array propList is fixed and MUST correspond to the position of the corresponding restriction in the property list restriction table.");
                    }

                    if (prop.PropInfo.PropID == 0x67AA)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3047");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3047
                        this.Site.CaptureRequirementIfAreEqual<ushort>(
                            0x000B,
                            prop.PropType,
                            3047,
                            @"[In PidTagAssociated Property] Data type: PtypBoolean ([MS-OXCDATA] section 2.11.1)");

                        // The property id of PidTagAssociated is  0x67AA,
                        Site.CaptureRequirement(
                            "MS-OXPROPS",
                            5368,
                            @"[In PidTagAssociated] Property ID: 0x67AA.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5369");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5369
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            0x000B,
                            prop.PropType,
                            "MS-OXPROPS",
                            5369,
                            @"[In PidTagAssociated] Data type: PtypBoolean, 0x000B.");

                        // The property id of PidTagAssociated is  0x67AA,
                        // if  0x67AA is present, this requirement can be verified
                        Site.CaptureRequirement(
                            1360,
                            @"[In messageChangeHeader Element] [PidTagAssociated,Required] The property PidTagAssociated MUST be present in the propList.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1361");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1361
                        // The position of the property PidTagAssociated is 5
                        Site.CaptureRequirementIfAreEqual<PropValue>(
                            msgChangeFull.MessageChangeHeader.PropList.PropValues[4],
                            prop,
                            1361,
                            @"[In messageChangeHeader Element] [PidTagAssociated ,Fixed position] The position[5] of the property PidTagAssociated within the array propList is fixed and MUST correspond to the position of the corresponding restriction in the property list restriction table.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2698");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2698
                        // PidTagAssociated is a PtypBoolean property, if its type is 0x000B, this requirement can be verified
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            0x000B,
                            prop.PropType,
                            "MS-OXCDATA",
                            2698,
                            @"[In Property Data Types] PtypBoolean (PT_BOOLEAN. bool) is that 1 byte, restricted to 1 or 0 [MS-DTYP]: BOOLEAN with Property Type Value 0x000B, %x0B.00.");
                    }

                    if (prop.PropInfo.PropID == 0x0E08)
                    {
                        // The property id of PidTagMessageSize is  0x0E08,
                        Site.CaptureRequirement(
                            "MS-OXPROPS",
                            7001,
                            @"[In PidTagMessageSize] Property ID: 0x0E08.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1284");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1284
                        // The property id of PidTagMessageSize is  0x0E08,
                        // 0x0003 is PtypInteger32
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0003,
                            prop.PropType,
                            1284,
                            @"[In PidTagMessageSize Property] Data type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");

                        // If the program can execute here, this requirement can be verified.
                        this.Site.CaptureRequirement(
                            3049,
                            @"[In PidTagMessageSize Property] The PidTagMessageSize property ([MS-OXPROPS] section 2.787) identifies the size of the message in bytes.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2691");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2691
                        // PidTagMessageSize is a PtypInteger32 property, if its type is 0x0003, this requirement can be verified
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0003,
                            prop.PropType,
                            "MS-OXCDATA",
                            2691,
                            @"[In Property Data Types] PtypInteger32 (PT_LONG, PT_I4, int, ui4) is that 4 bytes; a 32-bit integer [MS-DTYP]: INT32 with Property Type Value 0x0003,%x03.00.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7002");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R7002
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0003,
                            prop.PropType,
                            "MS-OXPROPS",
                            7002,
                            @"[In PidTagMessageSize] Data type: PtypInteger32, 0x0003.");
                    }

                    if (prop.PropInfo.PropID == 0x67A4)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3020");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3020
                        this.Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0014,
                            prop.PropType,
                            3020,
                            @"[In PidTagChangeNumber Property] [PidTagChangeNumber Property] Data type: PtypInteger64 ([MS-OXCDATA] section 2.11.1)");

                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            1269,
                            @"[In PidTagChangeNumber Property] The PidTagChangeNumber property ([MS-OXPROPS] section 2.623) contains the CN structure, as specified in section 2.2.2.1, that identifies the last change to the message or folder that is currently being synchronized.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2699");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2699
                        // PidTagMid is a PtypInteger64 property, if its type is 0x0014, this requirement can be verified
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0014,
                            prop.PropType,
                            "MS-OXCDATA",
                            2699,
                            @"[In Property Data Types] PtypInteger64 (PT_LONGLONG, PT_I8, i8, ui8) is that 8 bytes; a 64-bit integer [MS-DTYP]: LONGLONG with Property Type Value 0x0014,%x14.00.");

                        // The property id of PidTagMessageSize is  0x0E08,
                        Site.CaptureRequirement(
                            "MS-OXPROPS",
                            5699,
                            @"[In PidTagChangeNumber] Property ID: 0x67A4.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5700");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R5700
                        Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0014,
                            prop.PropType,
                            "MS-OXPROPS",
                            5700,
                            @"[In PidTagChangeNumber] Data type: PtypInteger64, 0x0014.");

                        FixedPropTypePropValue fixedProp = (FixedPropTypePropValue)prop;
                        byte[] bytesCN = BitConverter.GetBytes((long)fixedProp.FixedValue);
                        using (MemoryStream memoryStream = new MemoryStream(bytesCN))
                        {
                            CN cn = new CN();
                            cn.Deserialize(memoryStream, -1);

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R3053");

                            // Verify MS-OXPROPS requirement: MS-OXPROPS_R3053
                            Site.CaptureRequirementIfAreEqual<Type>(
                                typeof(uint),
                                cn.ReplicaId.GetType(),
                                3053,
                                @"[In CN Structure] ReplicaId (2 bytes): A 16-bit unsigned integer identifying the server replica in which the messaging object was last changed.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R3054");

                            // Verify MS-OXPROPS requirement: MS-OXPROPS_R3054
                            Site.CaptureRequirementIfAreEqual<Type>(
                                typeof(ulong),
                                cn.GlobalCounter.GetType(),
                                3054,
                                @"[In CN Structure] GlobalCounter (6 bytes): An unsigned 48-bit integer identifying the change to the messaging object.");
                        }
                    }

                    if (prop.PropInfo.PropID == 0x674A)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3016");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3016
                        this.Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0014,
                            prop.PropType,
                            3016,
                            @"[In PidTagMid Property] [PidTagMid Property] Data type: PtypInteger64 ([MS-OXCDATA] section 2.11.1)");

                        // If the program can execute here, this requirement can be verified.
                        this.Site.CaptureRequirement(
                            3017,
                            @"[In PidTagMid Property] The PidTagMid property ([MS-OXPROPS] section 2.792) contains the MID structure ([MS-OXCDATA] section 2.2.1.2) of the message currently being synchronized.");

                        // If the program can execute here, this requirement can be verified.
                        this.Site.CaptureRequirement(
                            "MS-OXPROPS",
                            7035,
                            @"[In PidTagMid] Property ID: 0x674A.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7036");

                        // Verify MS-OXCFXICS requirement: MS-OXPROPS_R7036
                        this.Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0014,
                            prop.PropType,
                            "MS-OXPROPS",
                            7036,
                            @"[In PidTagMid] Data type: PtypInteger64, 0x0014.");
                    }
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1368");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1368
                bool isVerifyR1368 = true;
                foreach (PropValue prop in msgChangeFull.MessageChangeHeader.PropList.PropValues)
                {
                    bool isExist = (prop.PropInfo.PropID != 0x65E0) // PidTagSourceKey
                                    && (prop.PropInfo.PropID != 0x3008) // PidTagLastModificationTime
                                    && (prop.PropInfo.PropID != 0x67AA) // PidTagAssociated
                                    && (prop.PropInfo.PropID != 0x65E3) // PidTagPredecessorChangeList
                                    && (prop.PropInfo.PropID != 0x65E2) // PidTagChangeKey
                                    && (prop.PropInfo.PropID != 0x674A) // PidTagMid
                                    && (prop.PropInfo.PropID != 0x0E08) // PidTagMessageSize
                                    && (prop.PropInfo.PropID != 0x67A4); // PidTagChangeNumber
                    if (isExist)
                    {
                        isVerifyR1368 = false;
                    }
                }

                byte[] changeKey = msgChangeFull.MessageChangeHeader.PidTagChangeKey;
                if (!Common.CompareByteArray(this.importPidTagChangeKeyValue, new byte[22]) && !this.lastChangeMadeByClient && !this.isNonImportMessageChangeOperation && this.importMessageTimes == 1)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R61");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R61
                    bool isVerifyR61 = Common.CompareByteArray(this.importPidTagChangeKeyValue, changeKey);
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR61,
                        61,
                        @"[In PidTagChangeKey Property] If the last change to the messaging object was imported from the client by using the RopSynchronizationImportMessageChange ROP (section 2.2.3.2.4.2), the value of the PidTagChangeKey property that is saved to the message store by the server contains the value for the PidTagChangeKey property that was passed in the PropertyValues field of the RopSynchronizationImportMessageChange ROP request buffer.");

                    this.importPidTagChangeKeyValue = new byte[22];
                }

                XID xidValue = new XID();
                using (MemoryStream streamChangeKey = new MemoryStream(changeKey))
                {
                    xidValue.Deserialize(streamChangeKey, -1);
                }

                byte[] changeKeyValue = xidValue.LocalId;

                if (this.lastChangeMadeByClient && changeKey != null)
                {
                    long changeNumber = msgChangeFull.MessageChangeHeader.PidTagChangeNumber;
                    CN changeNumberValue = new CN();
                    using (MemoryStream streamBytes = new MemoryStream(BitConverter.GetBytes(changeNumber)))
                    {
                        changeNumberValue.Deserialize(streamBytes, -1);
                    }

                    byte[] changeNumberByte = BitConverter.GetBytes(changeNumberValue.GlobalCounter);
                    byte[] newChangeNumber = new byte[6];
                    Array.Copy(changeNumberByte, 0, newChangeNumber, 0, 6);

                    if (changeNumber != 0)
                    {
                        bool isVerifyR62 = Common.CompareByteArray(changeKeyValue, newChangeNumber);
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR62,
                            62,
                            @"[In PidTagChangeKey Property] If the last change to a messaging object was made by the server, the value of the PidTagChangeKey property that is saved to the message store by the server contains an XID generated from the PidTagChangeNumber property (section 2.2.1.2.3).");
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1368,
                    1368,
                    @"[In messageChangeHeader Element] [other properties,prohibited] The other properties [the properties except PidTagSourceKey, PidTagLastModificationTime, PidTagChangeKey, PidTagPredecessorChangeList, PidTagAssociated, PidTagMid, PidTagMessageSize and PidTagChangeNumber] MUST NOT be present in the propList.");
            }
            #endregion Verify messageChangeHeader
        }

        /// <summary>
        /// Verify the subelelment deletions under contentsSync element. 
        /// </summary>
        /// <param name="contentSync">The contentsSync element of fastTransfer Stream.</param>
        private void VerifyDeletions(ContentsSync contentSync)
        {
            bool isVerifyR133703 = true;
            foreach (PropValue prop in contentSync.Deletions.PropList.PropValues)
            {
                if ((prop.PropInfo.PropID != 0x67E5) && (prop.PropInfo.PropID != 0x4021) && (prop.PropInfo.PropID != 0x6793))
                {
                    isVerifyR133703 = false;
                }

                if (prop.PropInfo.PropID == 0x67E5)
                {
                    // If the program can execute here, this requirement can be verified directly.
                    this.Site.CaptureRequirement(
                        3030,
                        @"[In MetaTagIdsetDeleted Meta-Property] [MetaTagIdsetDeleted property] Property ID: 0x67E5");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3662");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3662
                    this.Site.CaptureRequirementIfAreEqual<ushort>(
                        0x0102,
                        prop.PropType,
                        3662,
                        @"[In MetaTagIdsetDeleted Meta-Property] [MetaTagIdsetDeleted property]: Data type: PtypBinary, 0x0102 ([MS-OXCDATA] section 2.11.1)");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1276");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1276
                    // 0x67E5 is the property id of PidTagIdsetDeleted according [MS-OXPROPS]
                    // 0x0102 is PtypBinary
                    this.Site.CaptureRequirement(
                        1276,
                        @"[In MetaTagIdsetDeleted Meta-Property] The MetaTagIdsetDeleted property contains a serialization of a REPLID-based IDSET structures, as specified in section 2.2.2.4.1.");

                    VarPropTypePropValue varProp = (VarPropTypePropValue)prop;
                    using (MemoryStream memoryStream = new MemoryStream(varProp.ValueArray))
                    {
                        REPLID_IDSETElement repLid = new REPLID_IDSETElement();
                        repLid.Deserialize(memoryStream, -1);

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R182");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R182
                        Site.CaptureRequirement(
                            182,
                            @"[In Serialized IDSET Structure Containing a REPLID Structure] REPLID (2 bytes): A REPLID structure that when combined with all GLOBCNT structures contained in the GLOBSET field, produces a set of IDs.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R183");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R183
                        Site.CaptureRequirementIfAreEqual<Type>(
                            typeof(GLOBSET),
                            repLid.GLOBSET.GetType(),
                            183,
                            @"[In Serialized IDSET Structure Containing a REPLID Structure] GLOBSET (variable): A serialized GLOBSET structure.");
                    }
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R133703");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R133703
            Site.CaptureRequirementIfIsTrue(
                isVerifyR133703,
                133703,
                @"[In deletions Element] [The following restrictions exist on the contained propList element, as specified in section 2.2.4.3.20:] MUST adhere to the following restrictions: ] < other properties > is Prohibited.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1063");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1063
            bool isVerifyR1063 = contentSync.Deletions.PropList.PropValues.Count > 0;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1063,
                1063,
                @"[In deletions Element]The following restrictions exist on the contained propList element, as specified in section 2.2.4.3.20:
                MUST contain at least one property[PidTagIdsetDeleted, PidTagIdsetNoLongerInScope,PidTagIdsetExpired].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1337");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1337
            // 0x0102 0x4021  0x6793 are propertyIDs of PidTagIdsetDeleted,PidTagIdsetNoLongerInScope and PidTagIdsetExpired according [MS-OXPROPS]
            bool isVerifyR1337 = true;
            foreach (PropValue prop in contentSync.Deletions.PropList.PropValues)
            {
                bool isExist = (prop.PropInfo.PropID != 0x67E5) // PidTagIdsetDeleted
                            && (prop.PropInfo.PropID != 0x4021) // PidTagIdsetNoLongerInScope
                            && (prop.PropInfo.PropID != 0x6793); // PidTagIdsetExpired
                if (isExist)
                {
                    isVerifyR1337 = false;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1337,
                1337,
                @"[In deletions Element] [The following restrictions exist on the contained propList element, as specified in section 2.2.4.3.20:] MUST adhere to the following restrictions: ] The other properties[the properties except MetaTagIdsetDeleted,MetaTagIdsetNoLongerInScope and MetaTagIdsetExpired] MUST NOT be present in the contained propList.");
    }

        /// <summary>
        /// Verify the subelelment readStateChanges under contentsSync element. 
        /// </summary>
        /// <param name="contentSync">The contentsSync element of fastTransfer Stream.</param>
        private void VerifyReadStateChanges(ContentsSync contentSync)
        {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R3316");

                // Verify MS-OXPROPS requirement: MS-OXPROPS_R3316
                Site.CaptureRequirementIfIsTrue(
                     contentSync.ReadStateChanges.PropList.PropValues.Count > 0,
                     3316,
                     @"[In readStateChanges Element] The following restrictions exist on the contained propList element, as specified in section 2.2.4.3.20:
                        MUST contain at least one property.");

                bool isVerifyR3319 = true;
                foreach (PropValue prop in contentSync.ReadStateChanges.PropList.PropValues)
                {
                    if (prop.PropInfo.PropID == 0x402d)
                    {
                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3035,
                            @"[In MetaTagIdsetRead Meta-Property] [MetaTagIdsetRead property] Property ID: 0x402D");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3665");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3665
                        this.Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0102,
                            prop.PropType,
                            3665,
                            @"[In MetaTagIdsetRead Meta-Property] [MetaTagIdsetRead property]: Data type: PtypBinary, 0x0102 ([MS-OXCDATA] section 2.11.1)");

                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3036,
                            @"[In MetaTagIdsetRead Meta-Property] The MetaTagIdsetRead property contains a serialization of REPLID-based IDSET structures. ");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R3317");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R3317
                        this.Site.CaptureRequirement(
                             3317,
                             @"[In readStateChanges Element] The following restrictions exist on the contained propList element, as specified in section 2.2.4.3.20:
                        MUST adhere to the following restrictions:
                        Property name is MetaTagIdsetRead (section 2.2.1.3.4)");
                    }

                    if (prop.PropInfo.PropID == 0x402e)
                    {
                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3037,
                            @"[In MetaTagIdsetUnread Meta-Property] [MetaTagIdsetUnread property] Property ID: 0x402E");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3666");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3666
                        this.Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0102,
                            prop.PropType,
                            3666,
                            @"[In MetaTagIdsetUnread Meta-Property] [MetaTagIdsetUnread property]: Data type: PtypBinary, 0x0102 ([MS-OXCDATA] section 2.11.1)");

                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3038,
                            @"[In MetaTagIdsetUnread Meta-Property] The MetaTagIdsetUnread property contains a serialization of REPLID-based IDSET structures. ");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R3318");

                        // Verify MS-OXPROPS requirement: MS-OXPROPS_R3318
                        this.Site.CaptureRequirement(
                             3318,
                             @"[In readStateChanges Element] The following restrictions exist on the contained propList element, as specified in section 2.2.4.3.20:
                        MUST adhere to the following restrictions:
                        Property name is MetaTagIdsetUnread (section 2.2.1.3.5)");
                    }

                    if (prop.PropInfo.PropID != 0x402D && prop.PropInfo.PropID != 0x402E)
                    {
                        isVerifyR3319 = false;
                    }
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R3319");

                // Verify MS-OXPROPS requirement: MS-OXPROPS_R3319
                Site.CaptureRequirementIfIsTrue(
                     isVerifyR3319,
                     3319,
                     @"[In readStateChanges Element] The following restrictions exist on the contained propList element, as specified in section 2.2.4.3.20:
                        MUST adhere to the following restrictions:
                        <other properties> are prohibited.");
        }

        /// <summary>
        /// Verify the HierarchySync element of fastTransfer Stream
        /// </summary>
        /// <param name="fastTransferStream">The fastTransfer Stream get from server</param>
        private void VerifyHierarchySync(FastTransferStream fastTransferStream)
        {
            HierarchySync hrchSync = new HierarchySync(fastTransferStream);

            // When deserialize the FastTransfer Stream from the server response will as this requirement say,
            // if deserialize successfully, this requirement will be covered
            Site.CaptureRequirement(
                2026,
                @"[In  Creating Compact IDSET Structures]Note that this MUST NOT be done for IDSET structures that are ever iterated over, such as the
                MetaTagIdsetGiven property (section 2.2.1.1.1), as it will change the outcome.");

            #region IDSET UID GUID Command
            if ((null != hrchSync.State) && (null != hrchSync.State.PropList) && (null != hrchSync.State.PropList.PropValues))
            {
                foreach (PropValue prop in hrchSync.State.PropList.PropValues)
                {
                    // 0x4017 is the ID of PidTagIdsetGiven, 0x6796 is the ID of PidTagCnsetSeen
                    // 0x67DA is the ID of PidTagCnsetSeenFAI, 0x67D2 is the ID of PidTagCnsetRead
                    if ((prop.PropInfo.PropID == 0x4017) || (prop.PropInfo.PropID == 0x6796)
                        || (prop.PropInfo.PropID == 0x67DA) || (prop.PropInfo.PropID == 0x67D2))
                    {
                        VarPropTypePropValue varProp = (VarPropTypePropValue)prop;

                        // Transform the value of this property to a memory stream, the property value is a serialized IDSET.
                        using (MemoryStream memoryStream = new MemoryStream(varProp.ValueArray))
                        {
                            REPLGUID_IDSET idset = new REPLGUID_IDSET();

                            // Deserialize the above stream to get an IDSET.
                            // "-1" means no limitation except stream length.
                            idset.Deserialize(memoryStream, -1);

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R191");

                            Site.CaptureRequirementIfIsTrue(
                                idset.IsAllGLOBCNTInGLOBSET,
                                191,
                                @"[In Serialized IDSET Structure Containing a REPLGUID Structure] GLOBSET (variable): A serialized GLOBSET structure.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1934");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1934
                            // The structure is deserialized as this requirement, if Deserialize success, set IsAllGLOBCNTRanged as true
                            bool isVerifyR1934 = idset.IsAllGLOBCNTRanged;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1934,
                                1934,
                                @"[In Formatted IDSET Structures]. The ID values MUST be arranged by REPLID structure value");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1935");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1935
                            // The structure is deserialized as this requirement, if Deserialize success, set IsAllGLOBCNTInGLOBSET as true
                            bool isVerifyR1935 = idset.IsAllGLOBCNTInGLOBSET;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1935,
                                1935,
                                @"[In Formatted IDSET Structures]all IDs for each REPLID MUST be reduced into a GLOBSET, as specified in section 2.2.2.6,
                            of GLOBCNTs, as specified in section 2.2.2.5.");

                            // Add the debug information
                            // The structure is deserialized as this requirement, if Deserialize success, set IsAllGLOBCNTInGLOBSET as true
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1936");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1936
                            bool isVerifyR1936 = idset.HasAllDuplicateGLOBCNTRemoved;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1936,
                                1936,
                                @"[In Formatted IDSET Structures]Each GLOBSET MUST be arranged from lowest to highest GLOBCNT value where all duplicate GLOBCNT values are removed.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1937");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1937
                            // The structure is deserialized as this requirement, if Deserialize success, set HasGLOBCNTGroupedIntoRanges as true
                            bool isVerifyR1937 = idset.HasGLOBCNTGroupedIntoRanges;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1937,
                                1937,
                                @"[In Formatted IDSET Structures]The remaining GLOBCNT values MUST be grouped into consecutive ranges with a low GLOBCNT value and a high GLOBCNT value.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1938");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1938
                            // The structure is deserialized as this requirement, if Deserialize success, set IsDisjointGLOBCNTMadeIntoSingleton as true
                            bool isVerifyR1938 = idset.IsDisjointGLOBCNTMadeIntoSingleton;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1938,
                                1938,
                                @"[In Formatted IDSET Structures]If a GLOBCNT value is disjoint it MUST be made into a singleton range with the low and high GLOBCNT values being the same.");

                            foreach (REPLGUID_IDSETElement replguid in idset.IDSetList)
                            {
                                // The response from the server is serialized as this requirement is related to server, 
                                // if de-serialized the response (as the rule for server serializes the response) successfully, this requirement could be verified
                                Site.CaptureRequirement(
                                    186,
                                    @"[In Serialized IDSET Structure Containing a REPLGUID Structure] REPLGUID-GLOBSET structure pairs MUST be serialized by the value of the REPLGUID in the ascending order, using byte-to-byte comparison.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R187");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R187
                                Site.CaptureRequirementIfAreEqual<Type>(
                                    typeof(Guid),
                                    replguid.REPLGUID.GetType(),
                                    187,
                                    @"[In Serialized IDSET Structure Containing a REPLGUID Structure]REPLGUID (16 bytes): A GUID that identifies a REPLGUID structure.");

                                for (int i = 0; i < replguid.GLOBSET.DeserializedCommandList.Count; i++)
                                {
                                    if (replguid.GLOBSET.DeserializedCommandList[i] is PushCommand)
                                    {
                                        // comByteStack is the common byte stack
                                        CommonByteStack comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, new CommonByteStack());

                                        // The deserialization do as:
                                        // The Push command SHOULD be used when multiple GLOBCNT values share the same high-order values.
                                        // When push command is encountered, multiple GLOBCNT values share the same high-order values in Exchange
                                        bool pushHasSameHighOrderValue = this.WhetherPushHasSameHighOrderValue(replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTList);
                                        if (pushHasSameHighOrderValue)
                                        {
                                            // Verify requirement MS-OXCFXICS_R2660
                                            bool isR2660Satisfied = true;
                                            if (null != replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTList)
                                            {
                                                foreach (GLOBCNT globcnt in replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTList)
                                                {
                                                    for (int index = 0; index < comByteStack.Bytes; index++)
                                                    {
                                                        if (globcnt[index] != comByteStack.GetCommonBytes()[index])
                                                        {
                                                            // The multiple GLOBCNT structure values share the same higher-order values but Push command is not used.
                                                            isR2660Satisfied = false;
                                                            break;
                                                        }
                                                    }
                                                }

                                                if (Common.IsRequirementEnabled(2660, this.Site))
                                                {
                                                    // Add the debug information
                                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2660");

                                                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2660
                                                    Site.CaptureRequirementIfIsTrue(
                                                        isR2660Satisfied,
                                                        2660,
                                                        @"[In Push Command (0x01 - 0x06)] Implementation does use the Push command when multiple GLOBCNT structure values share the same high-order values. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                                }
                                            }
                                        }

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R202");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R202
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            (replguid.GLOBSET.DeserializedCommandList[i] as PushCommand).CommandByte,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as PushCommand).CommonBytes.Length,
                                            202,
                                            @"[In Push Command (0x01 - 0x06)] [CommonBytes (variable):] The length of this field is specified by value of the Command field (0x01 through 0x06), in bytes.");
                                    }

                                    if (replguid.GLOBSET.DeserializedCommandList[i] is PopCommand)
                                    {
                                        // comByteStack is the common byte stack
                                        CommonByteStack comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, new CommonByteStack());

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1966");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1966
                                        // from this requirement, can know if Pop Command is used, there are some bytes on the common byte stack. 
                                        // So if "Pop Command" is used and the common byte stack has some bytes, then this requirement can be covered.
                                        bool isVerifyR1966 = comByteStack.Bytes > 0;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR1966,
                                            1966,
                                            @"[In Pop Command (0x50)]The Pop command MUST NOT be used if no bytes are currently on the common byte stack.");
                                    }

                                    if (replguid.GLOBSET.DeserializedCommandList[i] is RangeCommand)
                                    {
                                        bool rangHasSameHighOrderValue = this.WhetherRangHasSameHighOrderValue(replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTRangeList);

                                        // comByteStack is the common byte stack
                                        CommonByteStack comByteStack = new CommonByteStack();
                                        if (rangHasSameHighOrderValue)
                                        {
                                            // Check the commands before current one in DeserializedCommandList,
                                            comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, comByteStack);

                                            if (Common.IsRequirementEnabled(2661, this.Site))
                                            {
                                                // Verify requirement MS-OXCFXICS_R2661
                                                bool isR2661Satisfied = true;

                                                for (int index = 0; index < comByteStack.Bytes; index++)
                                                {
                                                    if (replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTRangeList[0].StartGLOBCNT[index] != comByteStack.GetCommonBytes()[index] || replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTRangeList[0].EndGLOBCNT[index] != comByteStack.GetCommonBytes()[index])
                                                    {
                                                        isR2661Satisfied = false;
                                                        break;
                                                    }
                                                }

                                                // Add the debug information
                                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2661");

                                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2661
                                                // StartGLOBCNT and EndGLOBCNT are the low and high GLOBCNT values. 
                                                // They will share the high-order bytes which are in the common byte stack.
                                                Site.CaptureRequirementIfIsTrue(
                                                    isR2661Satisfied,
                                                    2661,
                                                    @"[In Range Command (0x52)] Implementation does push these common high-order bytes, that are shared by the low and high GLOBCNT values, onto the common byte stack by using the Push command, as specified in section 2.2.2.6.1, prior to using the Range command. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                            }
                                        }

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R213");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R213
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            6 - comByteStack.Bytes,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue.Length,
                                            213,
                                            @"[In Range Command (0x52)] [LowValue (variable):] The number of bytes in this field [LowValue] is equal to 6 minus the number of high-order bytes in the common byte stack.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R216");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R216
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            6 - comByteStack.Bytes,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue.Length,
                                            216,
                                            @"[In Range Command (0x52)] [HighValue (variable):] The number of bytes in this field [HighValue] is equal to 6 minus the number of high-order bytes in the common byte stack.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R214");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R214
                                        bool isVerifyR214 = true;
                                        for (int j = 0; j < 6 - comByteStack.Bytes; j++)
                                        {
                                            if ((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue[j] > (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue[j])
                                            {
                                                break;
                                            }
                                            else if ((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue[j] == (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue[j])
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                isVerifyR214 = false;
                                                break;
                                            }
                                        }

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR214,
                                            214,
                                            @"[In Range Command (0x52)] [LowValue (variable):] This value MUST be less than or equal to the value of the HighValue field.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R217");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R217
                                        bool isVerifyR217 = isVerifyR214;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR217,
                                            217,
                                            @"[In Range Command (0x52)] [HighValue (variable):] This value MUST be greater than or equal to the value of the LowValue field.");

                                        // When deserialize the FastTransfer Stream from the server response will as this requirement say,
                                        // if deserialize successfully, this method will be called , this requirement will be covered.
                                        Site.CaptureRequirement(
                                            1980,
                                            @"[In Range Command (0x52)]If the low and high value of the GLOBCNT range are not the same, the Range command MUST be used.");

                                        // When deserialize the FastTransfer Stream from the server response will as this requirement say,
                                        // if deserialize successfully, this method will be called , this requirement will be covered.
                                        Site.CaptureRequirement(
                                            1981,
                                            @"[In Range Command (0x52)]If the range has values that are more than 8 bytes from each other, the Range command MUST be used.");

                                        // When deserialize the FastTransfer Stream from the server response will as this requirement say,
                                        // if deserialize successfully, this method will be called , this requirement will be covered.
                                        Site.CaptureRequirement(
                                            1982,
                                            @"[In Range Command (0x52)]If the low and high value do not share five high-order bytes in common, the Range command MUST be used.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2009");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2009
                                        bool isVerifyR2009 = true;
                                        foreach (GLOBCNTRange gcntR in (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).CorrespondingGLOBCNTRangeList)
                                        {
                                            if (!replguid.GLOBSET.GLOBCNTRangeList.Contains(gcntR))
                                            {
                                                isVerifyR2009 = false;
                                                break;
                                            }
                                        }

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR2009,
                                            2009,
                                            @"[In Range Command (0x52)]The GLOBCNT structure range MUST be added to the GLOBSET structure.");

                                        byte[] lowValueOfRange = StructureSerializer.Serialize((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).CorrespondingGLOBCNTRangeList[0].StartGLOBCNT);
                                        byte[] highValueOfRange = StructureSerializer.Serialize((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).CorrespondingGLOBCNTRangeList[0].EndGLOBCNT);

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2012");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2012
                                        bool isVerifyR2012 = true;
                                        for (int index = 1; index <= (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue.Length; index++)
                                        {
                                            if ((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue[(replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue.Length - index] != lowValueOfRange[lowValueOfRange.Length - index]
                                             || (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue[(replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue.Length - index] != highValueOfRange[highValueOfRange.Length - index])
                                            {
                                                isVerifyR2012 = false;
                                                break;
                                            }
                                        }

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR2012,
                                            2012,
                                            @"[In Range Command (0x52)]Each of these fields MUST be combined with any high-order bytes in the common byte stack to produce a 6-byte GLOBCNT structure.");
                                    }

                                    if (replguid.GLOBSET.DeserializedCommandList[i] is BitmaskCommand)
                                    {
                                        // comByteStack is the common byte stack
                                        CommonByteStack comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, new CommonByteStack());

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1970");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1970
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            5,
                                            comByteStack.Bytes,
                                            1970,
                                            @"[In Bitmask Command (0x42)]There MUST already be five high-order bytes in the common byte stack to use this command.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2007");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2007
                                        // "globcnt" is a variable, it is constituted by the "low-order" and "five high-order bytes",
                                        // so if GLOBSET contains "globcnt", then this requirement can be covered.
                                        GLOBCNT globcnt = new GLOBCNT
                                        {
                                            Byte1 = comByteStack.GetCommonBytes()[0],
                                            Byte2 = comByteStack.GetCommonBytes()[1],
                                            Byte3 = comByteStack.GetCommonBytes()[2],
                                            Byte4 = comByteStack.GetCommonBytes()[3],
                                            Byte5 = comByteStack.GetCommonBytes()[4],
                                            Byte6 = (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue
                                        };

                                        bool isVerifyR2007 = false;
                                        byte start_value = (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue;
                                        byte bitmask = (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).Bitmask;

                                        // Start value combined with the required five high-order bytes
                                        // on the common byte stack must be included in the set.
                                        if (replguid.GLOBSET.GLOBCNTList.Contains(globcnt))
                                        {
                                            int j;

                                            // Generates low-order byte, combined with the required five high-order bytes
                                            // on the common byte stack
                                            for (j = 0; j < 8; j++)
                                            {
                                                if ((bitmask & (1 << j)) != 0)
                                                {
                                                    globcnt.Byte6 = (byte)(start_value + j + 1);
                                                    if (!replguid.GLOBSET.GLOBCNTList.Contains(globcnt))
                                                    {
                                                        break;
                                                    }
                                                }
                                            }

                                            // no break j equals to 8.if break one low-order byte MUST be combined with the required five high-order bytes
                                            // on the common byte stack is not in the set.
                                            isVerifyR2007 = j == 8;
                                        }

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR2007,
                                            2007,
                                            @"[In Bitmask Command (0x42)]Each low-order byte MUST be combined with the required five high-order bytes on the common
                                        byte stack to form a complete 6-byte GLOBCNT structure, as specified in section 2.2.2.5, which MUST be added to the
                                        GLOBSET structure.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1973");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1973
                                        Site.CaptureRequirementIfAreEqual<byte>(
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).CorrespondingGLOBCNTList[0].Byte6,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue,
                                            1973,
                                            @"[In Bitmask Command (0x42)]The StartingValue field, as specified in section 2.2.2.6.3, MUST be set to the low-order byte of the low value of the first GLOBCNT range.");

                                        byte byt = 0;
                                        foreach (GLOBCNT gcnt in (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).CorrespondingGLOBCNTList)
                                        {
                                            byt |= (byte)(1 << ((int)gcnt.Byte6 - (int)(replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue - 1));
                                        }

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1974");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1974
                                        Site.CaptureRequirementIfAreEqual<byte>(
                                            byt,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).Bitmask,
                                            1974,
                                            @"[In Bitmask Command (0x42)]The Bitmask field, specified in section 2.2.2.6.3, MUST have one bit set for each value within a range, excluding the low value of the first GLOBCNT range.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1978");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1978
                                        Site.CaptureRequirementIfAreEqual<byte>(
                                            byt,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).Bitmask,
                                            1978,
                                            @"[In Bitmask Command (0x42)] For all GLOBCNTs between ranges, the bit associated with the value is not set in the bitmask.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3071");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3071
                                        Site.CaptureRequirementIfAreEqual<short>(
                                            0x42,
                                            Convert.ToInt16((replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).CommandByte),
                                            3071,
                                            @"[In Bitmask Command (0x42)] Command (1 byte): This value MUST be set to 0x42.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R207");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R207
                                        Site.CaptureRequirementIfAreEqual<Type>(
                                            typeof(byte),
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue.GetType(),
                                            207,
                                            @"[In Bitmask Command (0x42)] StartingValue (1 byte): The low-order byte of the first GLOBCNT structure, as specified in section 2.2.2.5.");

                                        if ((replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).Bitmask > 0)
                                        {
                                            if (replguid.GLOBSET.GLOBCNTRangeList[0].StartGLOBCNT.Byte6 == (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue)
                                            {
                                                // If the range of GLOBCNT starts the StartingValue field, specified in Bitmask, this requirement can be captured.
                                                Site.CaptureRequirement(208, "[In Bitmask Command (0x42)] Bitmask (1 byte): A flag that identifies whether the GLOBCNT structure that exists in the GLOBCNT range that starts with the value of the StartingValue field.");
                                            }
                                        }
                                    }

                                    if (replguid.GLOBSET.DeserializedCommandList[i] is EndCommand)
                                    {
                                        CommonByteStack comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, new CommonByteStack());

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1992");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1992
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            0,
                                            comByteStack.Bytes,
                                            1992,
                                            @"[In End Command (0x00)]The End command can only be used if the common byte stack is empty.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1989");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1989
                                        // When de-serializing the FastTransfer Stream from the server response according to this requirement description,
                                        // so that if the de-serialization is successful, only need to verify that there is no byte on the common byte stack when "End command" is used.
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            0,
                                            comByteStack.Bytes,
                                            1989,
                                            @"[In End Command (0x00)]If after all GLOBCNT ranges have been encoded, there are still bytes on the common byte stack, 
                                        they MUST be removed with one or more Pop commands, as specified in section 2.2.2.6.1, before the End command can be
                                        used.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2014");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2014
                                        // EndCommand must be the last Command in GLOBSET.
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            replguid.GLOBSET.DeserializedCommandList.Count - 1,
                                            i,
                                            2014,
                                            @"[In End Command (0x00)]When the End command, as specified in section 2.2.2.6.5, is encountered, the GLOBSET structure
                                        MUST be complete based on the GLOBCNT structure values, as specified in section 2.2.2.5, generated from any previous
                                        encoding commands.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1987");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1987
                                        // EndCommand must be the last Command in GLOBSET.
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            replguid.GLOBSET.DeserializedCommandList.Count - 1,
                                            i,
                                            1987,
                                            @"[In End Command (0x00)]This command MUST be added after all GLOBCNT structure ranges, as specified in section 2.2.2.5,
                                        within the GLOBSET have been encoded.");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            #endregion

            if (null != hrchSync.FolderChangeList && hrchSync.FolderChangeList.Count > 0)
            {
                this.VerifyFolderChange(hrchSync);
            }
        }
        
        /// <summary>
        /// Verify the folderChange element.
        /// </summary>
        /// <param name="hrchSync">The HierarchySync element in fastTransfer Stream.</param>
        private void VerifyFolderChange(HierarchySync hrchSync)
        {
            foreach (PropValue prop in hrchSync.FolderChangeList[0].PropList.PropValues)
            {
                #region PidTagFolderId (Property ID = 0x6748) PropertyTag.

                if (prop.PropInfo.PropID == 0x6748)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3018");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3018
                    this.Site.CaptureRequirementIfAreEqual<ushort>(
                        0x0014,
                        prop.PropType,
                        3018,
                        @"[In PidTagFolderId Property] [PidTagFolderId Property] Data type: PtypInteger64 ([MS-OXCDATA] section 2.11.1)");

                    // If the program can execute here, this requirement can be verified directly.
                    this.Site.CaptureRequirement(
                        3019,
                        @"[In PidTagFolderId Property] The PidTagFolderId property ([MS-OXPROPS] section 2.691) contains the Folder ID structure ([MS-OXCDATA] section 2.2.1.1) of the folder currently being synchronized.");

                    // If the program can execute here, this requirement can be verified directly.
                    this.Site.CaptureRequirement(
                        "MS-OXPROPS",
                        6212,
                        @"[In PidTagFolderId] Property ID: 0x6748.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6213");

                    // Verify MS-OXCFXICS requirement: MS-OXPROPS_R6213
                    this.Site.CaptureRequirementIfAreEqual<ushort>(
                        0x0014,
                        prop.PropType,
                        "MS-OXPROPS",
                        6213,
                        @"[In PidTagFolderId] Data type: PtypInteger64, 0x0014.");
                }

                #endregion

                #region PidTagParentFolderId PropertyTag
                if (prop.PropInfo.PropID == 0x6749)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R181");

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R181
                    // ushort is 16-bit unsigned integer
                    Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(ushort),
                        prop.PropType.GetType(),
                        "MS-OXCDATA",
                        181,
                        @"[In PropertyTag Structure] PropertyType (2 bytes): An unsigned integer that identifies the data type of the property value, as specified by the table in section 2.11.1.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R182");

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R182
                    // ushort is 16-bit unsigned integer
                    // type of PropID is PropertyId 
                    Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(ushort),
                        prop.PropInfo.PropID.GetType(),
                        "MS-OXCDATA",
                        182,
                        @"[In PropertyTag Structure] PropertyId (2 bytes): An unsigned integer that identifies the property.");

                    // 0x6749 is the property id of PidTagParentFolderId
                    Site.CaptureRequirement(
                        "MS-OXPROPS",
                        7463,
                        @"[In PidTagParentFolderId] Property ID: 0x6749.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7464");

                    // Verify MS-OXPROPS requirement: MS-OXPROPS_R7464
                    Site.CaptureRequirementIfAreEqual<ushort>(
                        0x0014,
                        prop.PropType,
                        "MS-OXPROPS",
                        7464,
                        @"[In PidTagParentFolderId] Data type: PtypInteger64, 0x0014.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3021");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3021
                    this.Site.CaptureRequirementIfAreEqual<ushort>(
                        0x0014,
                        prop.PropType,
                        3021,
                        @"[In PidTagParentFolderId Property] [PidTagParentFolderId Property] Data type: PtypInteger64 ([MS-OXCDATA] section 2.11.1)");

                    // If the program execute here, this requirement can be verified directly.
                    this.Site.CaptureRequirement(
                        1270,
                        @"[In PidTagParentFolderId Property] The PidTagParentFolderId property ([MS-OXPROPS] section 2.850) contains the Folder ID structure ([MS-OXCDATA] section 2.2.1.1) that identifies the parent folder of the messaging object being synchronized.");
                }
                #endregion
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1088 and MS-OXCFXICS_R1127");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1088
            // 0x65E1 is the property id of PidTagParentSourceKey according [MS-OXPROPS]
            bool isVerifyR1088 = false;
            bool isVerifyR1127 = false;
            foreach (PropValue prop in hrchSync.FolderChangeList[0].PropList.PropValues)
            {
                if (prop.PropInfo.PropID == 0x65E1)
                {
                    isVerifyR1088 = true;

                    VarPropTypePropValue varProp = (VarPropTypePropValue)prop;
                    if (varProp.ValueArray.Length == 0)
                    {
                        isVerifyR1127 = true;
                    }
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1088,
                1088,
                @"[In folderChange Element] [PidTagParentSourceKey ,Required] The property PidTagParentSourceKey MUST be present in the propList.");

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1127,
                1127,
                @"[In hierarchySync Element] The folderChange elements that have a PidTagParentSourceKey property with a zero-length value are children of the root of the synchronization operation.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7477");

            // Verify MS-OXPROPS requirement: MS-OXPROPS_R7477
            // if isVerifyR1088 is true, the id of PidTagParentSourceKey is 0x65E1
            bool isVerifyR7477 = isVerifyR1088;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR7477,
                "MS-OXPROPS",
                7477,
                @"[In PidTagParentSourceKey] Property ID: 0x65E1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1272");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1272
            // 0x0102 is PtypBinary
            bool isVerifiedR3025 = false;
            foreach (PropValue prop in hrchSync.FolderChangeList[0].PropList.PropValues)
            {
                if (prop.PropInfo.PropID == 0x65E1)
                {
                    if (prop.PropType == 0x0102)
                    {
                        isVerifiedR3025 = true;
                    }
                }
            }

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3025
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR3025,
                3025,
                @"[In PidTagParentSourceKey Property] Data type: PtypBinary ([MS-OXCDATA] section 2.11.1)");

            // If the program can execute here, this requirement can be verified directly.
            this.Site.CaptureRequirement(
                1272,
                @"[In PidTagParentSourceKey Property] The PidTagParentSourceKey property ([MS-OXPROPS] section 2.852) specifies the PidTagSourceKey property (section 2.2.1.2.5) of the current folder's parent folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7478");

            // Verify MS-OXPROPS requirement: MS-OXPROPS_R7478
            // if isVerifyR1272 is true, the type of PidTagParentSourceKey is 0x0102
            bool isVerifyR7478 = isVerifiedR3025;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR7478,
                "MS-OXPROPS",
                7478,
                @"[In PidTagParentSourceKey] Data type: PtypBinary, 0x0102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1089");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1089
            //  0x65E0 is the property id of PidTagSourceKey according [MS-OXPROPS]
            bool isVerifyR1089 = false;
            foreach (PropValue prop in hrchSync.FolderChangeList[0].PropList.PropValues)
            {
                if (prop.PropInfo.PropID == 0x65E0)
                {
                    isVerifyR1089 = true;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1089,
                1089,
                @"[In folderChange Element] [PidTagSourceKey ,Required] The property PidTagSourceKey MUST be present in the propList.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R8582");

            // Verify MS-OXPROPS requirement: MS-OXPROPS_R8582
            // if isVerifyR1089 is true, the id of PidTagSourceKey is 0x65E0
            bool isVerifyR8582 = isVerifyR1089;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR8582,
                "MS-OXPROPS",
                8582,
                @"[In PidTagSourceKey] Property ID: 0x65E0.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3022");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3022
            // 0x0102 is PtypBinary
            bool isVerifiedR3022 = false;
            foreach (PropValue prop in hrchSync.FolderChangeList[0].PropList.PropValues)
            {
                if (prop.PropInfo.PropID == 0x65E0)
                {
                    if (prop.PropType == 0x0102)
                    {
                        isVerifiedR3022 = true;
                        break;
                    }
                }
            }

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3022
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR3022,
                3022,
                @"[In PidTagSourceKey Property] Data type: PtypBinary ([MS-OXCDATA] section 2.11.1)");

            // If the program can execute here, this requirement can be verified.
            this.Site.CaptureRequirement(
                1271,
                @"[In PidTagSourceKey Property] The PidTagSourceKey property ([MS-OXPROPS] section 2.1012) contains a serialized XID structure, as specified in section 2.2.2.2, that specifies the internal identifier (2) for the folder or message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R8583");

            // Verify MS-OXPROPS requirement: MS-OXPROPS_R8583
            // if isVerifyR1271 is true, the type of PidTagSourceKey is 0x0102.
            bool isVerifyR8583 = isVerifiedR3022;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR8583,
                "MS-OXPROPS",
                8583,
                @"[In PidTagSourceKey] Data type: PtypBinary, 0x0102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1090");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1090
            // 0x3008 is the property id of PidTagLastModificationTime according [MS-OXPROPS]
            bool isVerifyR1090 = false;
            foreach (PropValue prop in hrchSync.FolderChangeList[0].PropList.PropValues)
            {
                if (prop.PropInfo.PropID == 0x3008)
                {
                    isVerifyR1090 = true;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1090,
                1090,
                @"[In folderChange Element] [PidTagLastModificationTime,Required] The property PidTagLastModificationTime MUST be present in the propList.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1091");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1091
            // 0x65E2 is the property id of PidTagChangeKey according [MS-OXPROPS]
            bool isVerifyR1091 = false;
            foreach (PropValue prop in hrchSync.FolderChangeList[0].PropList.PropValues)
            {
                if (prop.PropInfo.PropID == 0x65E2)
                {
                    isVerifyR1091 = true;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1091,
                1091,
                @"[In folderChange Element] [PidTagChangeKey,Required] The property PidTagChangeKey MUST be present in the propList.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5686");

            // Verify MS-OXPROPS requirement: MS-OXPROPS_R5686
            // if isVerifyR1090 is true, the id of PidTagChangeKey is 0x65E2
            bool isVerifyR5686 = isVerifyR1090;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR5686,
                "MS-OXPROPS",
                5686,
                @"[In PidTagChangeKey] Property ID: 0x65E2.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3026");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3026
            // 0x0102 is PtypBinary
            bool isVerifiedR3026 = false;
            foreach (PropValue prop in hrchSync.FolderChangeList[0].PropList.PropValues)
            {
                if (prop.PropInfo.PropID == 0x65E2)
                {
                    if (prop.PropType == 0x0102)
                    {
                        isVerifiedR3026 = true;
                        break;
                    }
                }
            }

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3026
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR3026,
                3026,
                @"[In PidTagChangeKey Property] Data type: PtypBinary ([MS-OXCDATA] section 2.11.1)");

            // If the program can execute here, this requirement can be verified directly.
            this.Site.CaptureRequirement(
                1273,
                @"[In PidTagChangeKey Property] The PidTagChangeKey property ([MS-OXPROPS] section 2.622) contains a serialized XID structure, as specified in section 2.2.2.2, that identifies the last change to the messaging object.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R5687");

            // Verify MS-OXPROPS requirement: MS-OXPROPS_R5687
            // if isVerifyR1273 is true, the type of PidTagChangeKey is 0x0102
            bool isVerifyR5687 = isVerifiedR3026;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR5687,
                "MS-OXPROPS",
                5687,
                @"[In PidTagChangeKey] Data type: PtypBinary, 0x0102.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1092");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1092
            // 0x65E3 is the property id of PidTagPredecessorChangeList according [MS-OXPROPS]
            bool isVerifyR1092 = false;
            foreach (PropValue prop in hrchSync.FolderChangeList[0].PropList.PropValues)
            {
                if (prop.PropInfo.PropID == 0x65E3)
                {
                    isVerifyR1092 = true;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1092,
                1092,
                @"[In folderChange Element] [PidTagPredecessorChangeList,Required] The property PidTagPredecessorChangeList MUST be present in the propList.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3027");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3027
            // 0x0102 is PtypBinary
            bool isVerifiedR3027 = false;
            foreach (PropValue prop in hrchSync.FolderChangeList[0].PropList.PropValues)
            {
                if (prop.PropInfo.PropID == 0x65E3)
                {
                    if (prop.PropType == 0x0102)
                    {
                        isVerifiedR3027 = true;
                        break;
                    }
                }
            }

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3027
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR3027,
                3027,
                @"[In PidTagPredecessorChangeList Property] Data type: PtypBinary ([MS-OXCDATA] section 2.11.1)");

            // If the program can execute here, this requirement can be verified directly.
            this.Site.CaptureRequirement(
                1274,
                @"[In PidTagPredecessorChangeList Property] The PidTagPredecessorChangeList property ([MS-OXPROPS] section 2.858) contains PredecessorChangeList structures, as specified in section 2.2.2.3.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1093");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1093
            // 0x3001 is the property id of PidTagDisplayName according [MS-OXPROPS]
            bool isVerifyR1093 = false;
            foreach (PropValue prop in hrchSync.FolderChangeList[0].PropList.PropValues)
            {
                if (prop.PropInfo.PropID == 0x3001)
                {
                    isVerifyR1093 = true;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1093,
                1093,
                @"[In folderChange Element] [PidTagDisplayName,Required] The property PidTagDisplayName MUST be present in the propList.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6015");

            // Verify MS-OXPROPS requirement: MS-OXPROPS_R6015
            bool isVerifyR6015 = false;
            foreach (PropValue prop in hrchSync.FolderChangeList[0].PropList.PropValues)
            {
                if (prop.PropInfo.PropID == 0x3001 && prop.PropType == 0x001F)
                {
                    isVerifyR6015 = true;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR6015,
                "MS-OXPROPS",
                6015,
                @"[In PidTagDisplayName] Data type: PtypString, 0x001F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R6014");

            // Verify MS-OXPROPS requirement: MS-OXPROPS_R6014
            // if isVerifyR1093 is true, the id of PidTagDisplayName is 0x3001
            bool isVerifyR6014 = isVerifyR1093;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR6014,
                "MS-OXPROPS",
                6014,
                @"[In PidTagDisplayName] Property ID: 0x3001.");
        }
        
        /// <summary>
        /// Verify the FolderContent
        /// </summary>
        /// <param name="fastTransferStream">The fastTransfer Stream get from server</param>
        private void VerifyFolderContent(FastTransferStream fastTransferStream)
        {
            #region Verify folderContent
            FolderContent fldContent = new FolderContent(fastTransferStream);

            if (this.propertyTags != null)
            {
                if (this.previousOperation == EnumFastTransferOperation.FastTransferSourceCopyProperties)
                {
                    bool isVerifyR450 = true;
                    foreach (PropValue prop in fldContent.PropList.PropValues)
                    {
                        // The MetaTagDnPrefix meta-property MUST be ignored when received
                        if (prop.PropInfo.PropID == 0x4008)
                        {
                            continue;
                        }

                        bool isPropertyExist = false;
                        foreach (PropertyTag propertyTag in this.propertyTags)
                        {
                            if (propertyTag.PropertyId == prop.PropInfo.PropID)
                            {
                                isPropertyExist = true;
                                break;
                            }
                        }

                        if (!isPropertyExist)
                        {
                            isVerifyR450 = false;
                            break;
                        }
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R450");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R450
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR450,
                        450,
                        @"[In RopFastTransferSourceCopyProperties ROP Request Buffer] [PropertyTags (variable)] This array specifies the properties and subobjects, as specified in section 2.2.1.7, to copy from the messaging object pointed to by the InputServerObject field.");

                    if (fldContent.SubFolders != null)
                    {
                        bool isVerifyR451 = fldContent.SubFolders[0].FolderContent.PropList.PropValues.Count > fldContent.PropList.PropValues.Count;

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R451");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R451
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR451,
                            451,
                            @"[In RopFastTransferSourceCopyProperties ROP Request Buffer] [PropertyTags (variable)] Note that this field MUST NOT be considered when determining what properties and subobjects to copy for descendant subobjects of the InputServerObject field.");
                    }
                }

                if (this.previousOperation == EnumFastTransferOperation.FastTransferSourceCopyTo)
                {
                    bool isVerifyR379 = true;
                    foreach (PropertyTag propertyTag in this.propertyTags)
                    {
                        bool isPropertyExist = false;
                        foreach (PropValue prop in fldContent.PropList.PropValues)
                        {
                            if (propertyTag.PropertyId == prop.PropInfo.PropID)
                            {
                                isPropertyExist = true;
                                break;
                            }
                        }

                        if (isPropertyExist)
                        {
                            isVerifyR379 = false;
                            break;
                        }
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R379");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R379
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR379,
                        379,
                        @"[In RopFastTransferSourceCopyTo ROP Request Buffer] [PropertyTags (variable)] Specifies properties and subobjects, as specified in section 2.2.1.7, to be excluded when copying a messaging object pointed to by the InputServerObject field.");
                    if (this.currentPermission == PermissionLevels.FolderVisible && ((this.currentCopyFlag & CopyToCopyFlags.Move) == CopyToCopyFlags.Move))
                    {
                        if (fldContent.SubFolders.Count > 0)
                        {
                            bool isVerifiedR2813 = false;
                            if (Common.IsRequirementEnabled(2813, this.Site) == false)
                            {
                                isVerifiedR2813 = true;
                                foreach (SubFolder subFolder in fldContent.SubFolders)
                                {
                                    if (subFolder.FolderContent.PropList.PropValues.Count > 0)
                                    {
                                        if (subFolder.FolderContent.PropList.PropValues[3] != null)
                                        {
                                            isVerifiedR2813 = false;
                                            break;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                foreach (SubFolder subFolder in fldContent.SubFolders)
                                {
                                    if (subFolder.FolderContent.PropList.PropValues.Count > 0)
                                    {
                                        if (subFolder.FolderContent.PropList.PropValues[3].PropInfo.PropID == 0x400F)
                                        {
                                            isVerifiedR2813 = true;
                                            break;
                                        }
                                    }
                                }
                            }

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3299");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3299
                            bool isVerifiedR3299 = false;
                            if (Common.IsRequirementEnabled(118201, this.Site))
                            {
                                foreach (SubFolder subFolder in fldContent.SubFolders)
                                {
                                    if (subFolder.FolderContent.PropList.PropValues.Count > 0)
                                    {
                                        if (subFolder.FolderContent.PropList.PropValues[3].PropInfo.PropID == 0x400F)
                                        {
                                            isVerifiedR3299 = true;
                                            break;
                                        }
                                    }
                                }

                                this.Site.CaptureRequirementIfIsTrue(
                                    isVerifiedR3299,
                                    3299,
                                    @"[In folderContent Element] If the MetaTagEcWarning (section 2.2.4.1.5.2) is present,  it will be in the fixed the position [4] of the array list.");

                                // Add the debug information
                                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R118201");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R118201
                                this.Site.CaptureRequirementIfIsTrue(
                                    isVerifiedR3299,
                                    118201,
                                    @"[In Appendix A: Product Behavior] Implementation does support include the MetaTagEcWarning meta-property (section 2.2.4.1.5.2) in the propList of the folderContent element. (<21> Section 2.2.4.3.20: Exchange 2007 follows this behavior.)");
                            }

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2813");

                            this.Site.CaptureRequirementIfIsTrue(
                                   isVerifiedR2813,
                                   2813,
                                   @"[In Appendix A: Product Behavior] If the client set the Move flag of the CopyFlags field and the user does not have permissions to delete the source folder, implementation does not output.  <19> Section 2.2.4.3.6: Exchange 2010, Exchange 2013 and Exchange 2016 do not include the MetaTagEcWarning meta-property (section 2.2.4.1.5.2) in the propList element as Exchange 2010 , Exchange 2013 and Exchange 2016 do not check permissions on move operations.");
                        }                        
                    }
                }
            }

            if (fldContent.FolderMessages == null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3301");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3301
                // 0x4011 is the property id of MetaTagNewFXFolder according [MS-OXPROPS]
                bool isVerifyR3301 = true;
                foreach (PropValue prop in fldContent.PropList.PropValues)
                {
                    if (prop.PropInfo.PropID == 0x4011)
                    {
                        isVerifyR3301 = false;
                        break;
                    }
                }

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R3301");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3301
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR3301,
                    3301,
                    @"[In folderContent Element] The MetaTagNewFXFolder meta-property (section 2.2.4.1.5.3) MUST be output instead of message elements when outputting a public folder whose contents do not exist on the server because the content is ghosted. ");
            }

            if (null != fldContent.PropList)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3786");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3786
                // ID 0x6600 to 0x67FF means Provider-defined internal non-transmittable property 
                bool isVerifyR3786 = true;
                foreach (PropValue prop in fldContent.PropList.PropValues)
                {
                    if (prop.PropInfo.PropID >= 0x6600 && prop.PropInfo.PropID <= 0x67FF)
                    {
                        isVerifyR3786 = false;
                        break;
                    }
                }

                // If the property id is under the scope (0x6600-0x67ff), which means it belongs to the provider-defined internal nontransmittable range.
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR3786,
                    3786,
                    @"[In Properties to Ignore on Download] Unless specified otherwise in property list restriction tables, propValue elements of FastTransfer streams, as specified in section 2.2.4.3.21, that belong to the provider-defined internal nontransmittable range [0x6600-0x67FF] , as specified in [MS-OXPROPS] section 1.3.3, MUST be excluded from download.");

                if (fldContent.SubFolders != null)
                {
                    foreach (SubFolder folder in fldContent.SubFolders)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1344");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1344
                        bool isVerifyR1344 = false;
                        foreach (PropValue prop in folder.FolderContent.PropList.PropValues)
                        {
                            if (prop.PropInfo.PropID == 0x6748)
                            {
                                isVerifyR1344 = true;
                                break;
                            }
                        }

                        if (isVerifyR1344)
                        {
                            // 0x6748 is the property id of PidTagFolderId according [MS-OXPROPS]
                            // Capture this requirement here since the folder content only contains subfolders.
                            Site.CaptureRequirement(
                                1344,
                                @"[In folderContent Element] if the PidTagFolderId (section 2.2.1.2.2) is present, it will fix the position of the array list.");
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1345");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1345
                        bool isVerifyR1345 = isVerifyR1344;

                        if (isVerifyR1345)
                        {
                            // 0x6748 is the property id of PidTagFolderId according [MS-OXPROPS]
                            // Capture this requirement here since the folder content only contains subfolders
                            Site.CaptureRequirement(
                                1345,
                                @"[In folderContent Element] Whether the PidTagFolderId (section 2.2.1.2.2) is present that is optional and it MUST be present if and only if the first marker in the folder element is not the StartTopFld marker, as specified in section 2.2.4.1.4. ");
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1347");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1347
                        bool isVerifyR1347 = false;
                        foreach (PropValue prop in fldContent.PropList.PropValues)
                        {
                            if (prop.PropInfo.PropID == 0x3001)
                            {
                                isVerifyR1347 = true;
                                break;
                            }
                        }

                        if (isVerifyR1347)
                        {
                            // The 0x3001 is the property id of PidTagDisplayName according to MS-OXPROPS.
                            // Capture this requirement here since the folder content only contains subfolders.
                            Site.CaptureRequirement(
                                1347,
                                @"[In folderContent Element] If the PidTagDisplayName ([MS-OXCFOLD] section 2.2.2.2.2.5) is present, it will fix the position of the array list.");
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1348");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1348
                        bool isVerifyR1348 = isVerifyR1347;

                        if (isVerifyR1348)
                        {
                            // The 0x3001 is the property id of PidTagDisplayName according to MS-OXPROPS.
                            // Capture this requirement here since the folder content only contains subfolders.
                            Site.CaptureRequirement(
                                1348,
                                @"[In folderContent Element] Whether the PidTagDisplayName ([MS-OXCFOLD] section 2.2.2.2.2.5) is present is optional, and it MUST be present if and only if the first marker in the folder element is not the StartTopFld marker.");
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1350");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1350
                        bool isVerifyR1350 = false;
                        foreach (PropValue prop in fldContent.PropList.PropValues)
                        {
                            if (prop.PropInfo.PropID == 0x3004)
                            {
                                isVerifyR1350 = true;
                                break;
                            }
                        }

                        if (isVerifyR1350)
                        {
                            // The 0x3004 is the property id of PidTagComment according to MS-OXPROPS.
                            // Capture this requirement here since the folder content only contains subfolders
                            Site.CaptureRequirement(
                                1350,
                                @"[In folderContent Element] If the PidTagComment ([MS-OXCFOLD] section 2.2.2.2.2.2) is present, it will fix the position of the array list.");
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1351");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1351
                        bool isVerifyR1351 = isVerifyR1350;

                        if (isVerifyR1351)
                        {
                            // The 0x3004 is the property id of PidTagComment according to MS-OXPROPS.
                            // Capture this requirement here since the folder content only contains subfolders
                            Site.CaptureRequirement(
                                1351,
                                @"[In folderContent Element] Whether the PidTagComment ([MS-OXCFOLD] section 2.2.2.2.2.2) is present that is optional, it MUST be present if and only if the first marker in the folder element is not the StartTopFld marker.");
                        }
                    }
                }
            }

            #endregion Verify folderContent
        }

        /// <summary>
        /// Verify the MessageContent
        /// </summary>
        /// <param name="fastTransferStream">The fastTransfer Stream get from server</param>
        private void VerifyMessageContent(FastTransferStream fastTransferStream)
        {
            MessageContent msgContent = new MessageContent(fastTransferStream);

            #region Verify string and String8
            foreach (PropValue prop in msgContent.PropList.PropValues)
            {
                // PropType equal to 0x001F means type is PtypString
                if (prop.PropType == 0x001F)
                {
                    VarPropTypePropValue varPropTypePropValue = prop as VarPropTypePropValue;
                    if (varPropTypePropValue != null && varPropTypePropValue.ValueArray != null && varPropTypePropValue.ValueArray.Length > 1)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2700");

                        // Verify MS-OXCFXICS requirement: MS-OXCDATA_R2700
                        // A string property  with terminating null character (2 bytes of zero)
                        bool isVerifyR2700 = (prop as VarPropTypePropValue).ValueArray[(prop as VarPropTypePropValue).ValueArray.Length - 1] == 0
                                            && (prop as VarPropTypePropValue).ValueArray[(prop as VarPropTypePropValue).ValueArray.Length - 2] == 0;
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR2700,
                            "MS-OXCDATA",
                            2700,
                            @"[In Property Data Types] PtypString (PT_UNICODE, string) is that Variable size; a string of Unicode characters in UTF-16LE format encoding with terminating null character (0x0000). with Property Type Value  0x001F,%x1F.00.");

                        if ((prop as VarPropTypePropValue).ValueArray.Length < 32 * 1024 && Common.IsRequirementEnabled(325901, this.Site))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R325901");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R325901
                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR2700,
                                325901,
                                @"[In Appendix A: Product Behavior] Serialization of simple types in FastTransfer streams is identical to serialization of property values as specified [MS-OXCDATA] , with the following exceptions: For PtypString and PtypString8 ([MS-OXCDATA] section 2.11.1) type, Implementation does output string values with the terminating nulls. (<17> Section 2.2.4.1.3: Microsoft Exchange Server 2007 (only when when string values are not larger than 32 KB) and above follow this behavior.)");
                        }

                        // String type is Serialized the as specified in [MS-OXCDATA]
                        // if this element can be gotten, this requirement can be covered.
                        Site.CaptureRequirement(
                             992,
                            @"[In Serialization of Simple Types] Serialization of simple types in FastTransfer streams is identical to serialization of property values as specified [MS-OXCDATA] , with the following exceptions: For PtypString and PtypString8 ([MS-OXCDATA] section 2.11.1) type, serialization MUST be performed, as specified in [MS-OXCDATA] .");
                    }
                }

                // PropType equal to 0x001E means type is PtypString8
                if (prop.PropType == 0x001E)
                {
                    VarPropTypePropValue varPropTypePropValue = prop as VarPropTypePropValue;

                    if (varPropTypePropValue != null && varPropTypePropValue.ValueArray != null && varPropTypePropValue.ValueArray.Length > 1)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2701");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R2701
                        // A string property  with terminating null character (1 bytes of zero)
                        bool isVerifyR2701 = (prop as VarPropTypePropValue).ValueArray[(prop as VarPropTypePropValue).ValueArray.Length - 1] == 0
                                            && (prop as VarPropTypePropValue).ValueArray[(prop as VarPropTypePropValue).ValueArray.Length - 2] != 0;

                        Site.CaptureRequirementIfIsTrue(
                            isVerifyR2701,
                            "MS-OXCDATA",
                            2701,
                            @"[In Property Data Types]PtypString8 is that variable size, a string of multi-byte characters in externally specified 
                                encoding with terminating null character (single 0 byte).PT_STRING8 with Property Type Value 0x001E,%z1E.00.");

                        if ((prop as VarPropTypePropValue).ValueArray.Length < 32 * 1024 && Common.IsRequirementEnabled(325901, this.Site))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R325901");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R325901
                            Site.CaptureRequirementIfIsTrue(
                               isVerifyR2701,
                               325901,
                               @"[In Appendix A: Product Behavior] Serialization of simple types in FastTransfer streams is identical to serialization of property values as specified [MS-OXCDATA] , with the following exceptions: For PtypString and PtypString8 ([MS-OXCDATA] section 2.11.1) type, Implementation does output string values with the terminating nulls. (<17> Section 2.2.4.1.3: Microsoft Exchange Server 2007 (only when when string values are not larger than 32 KB) and above follow this behavior.)");
                        }

                        // String type is Serialized the as specified in [MS-OXCDATA]
                        // if this element can be gotten, this requirement can be covered.
                        Site.CaptureRequirement(
                             992,
                            @"[In Serialization of Simple Types] Serialization of simple types in FastTransfer streams is identical to serialization of property values as specified [MS-OXCDATA] , with the following exceptions: For PtypString and PtypString8 ([MS-OXCDATA] section 2.11.1) type, serialization MUST be performed, as specified in [MS-OXCDATA] .");
                    }
                }
            }
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3311002");

            for (int i = 0; i < msgContent.PropList.PropValues.Count; i++)
            {
                if (msgContent.PropList.PropValues[i].PropInfo.PropID == 0x674A && i == 0)
                {
                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3311002
                    // 0x674A is the property id of PidTagMid according [MS-OXPROPS]
                    Site.CaptureRequirement(
                           3311002,
                           @"[In messageContent Element] [PidTagMid] Whether the property PidTagMid is present that is optional, it will Fix position of the array list.");
                }
            }

            if (this.previousOperation == EnumFastTransferOperation.FastTransferSourceCopyProperties)
            {
                if (this.copySubObjects && msgContent.MessageChildren.FXDelProps.Count > 0)
                {
                    // When the level is false, sub-objects are copied in FastTransferSourceCopyProperties operation
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2280");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2280
                    Site.CaptureRequirement(
                        2280,
                        @"[In Effect of Property and Subobject Filters on Download] Subobjects of a particular type are [either filtered out, in which case the PidTagFXDelProp meta-property MUST NOT be output;] or are filtered in, that is, they MUST be output one after another, prefixed by the PidTagFXDelProp meta-property.");
                }

                if (!this.copySubObjects && msgContent.MessageChildren.FXDelProps.Count == 0)
                {
                    // When the level is true, sub-objects are not copied in FastTransferSourceCopyProperties operation
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2278");

                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2278
                    Site.CaptureRequirement(
                        2278,
                        @"[In Effect of Property and Subobject Filters on Download] Subobjects of a particular type are either filtered out, in which case the MetaTagFXDelProp meta-property MUST NOT be output [or are filtered in] .");
                }
            }
        }

        /// <summary>
        /// Verify the TopFolder
        /// </summary>
        /// <param name="fastTransferStream">The fastTransfer Stream get from server</param>
        private void VerifyTopFolder(FastTransferStream fastTransferStream)
        {
            TopFolder topFolder = new TopFolder(fastTransferStream);
            if (this.folderCopyFlag == CopyFolderCopyFlags.CopySubfolders)
            {
                if (topFolder.FolderContent.SubFolders.Count > 0)
                {
                    // The count of subfolders is greater than zero, which means the server does include the subfolders of the folder. So this requirement can be verified.
                    Site.CaptureRequirement(
                        2124,
                        @"[In Receiving a RopFastTransferSourceCopyFolder ROP Request] If the CopySubfolders flag of the CopyFlags field is set, the server MUST recursively include the subfolders of the folder specified in the InputServerObject field in the scope.");
                }
            }
            else if (this.folderCopyFlag == CopyFolderCopyFlags.None)
            {
                if (topFolder.FolderContent.SubFolders.Count == 0)
                {
                    // The count of subfolders is zero, which means the server doesn't include the subfolders of the folder. So this requirement can be verified.
                    Site.CaptureRequirement(
                        3482,
                        @"[In Receiving a RopFastTransferSourceCopyFolder ROP Request] If the Move flag of the CopyFlags field is not set and the CopySubfolders flag is not set, the server MUST NOT recursively include the subfolders of the folder specified in the InputServerObject field in the scope.");
                }
            }

            #region PidTagMid
            bool isR3311001Verified = false;
            if (topFolder.FolderContent.FolderMessages.MessageLists.Count > 0 && topFolder.FolderContent.FolderMessages.MessageLists[0].Messages.Count > 0)
            {
                foreach (PropValue prop in topFolder.FolderContent.FolderMessages.MessageLists[0].Messages[0].Content.PropList.PropValues)
                {
                    // Property PidTagMid must be present in the property list of message content
                    // 0x674A is the property id of PidTagMid
                    if (prop.PropInfo.PropID == 0x674A)
                    {
                        isR3311001Verified = true;
                        break;
                    }
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3311001");

                Site.CaptureRequirementIfIsTrue(
                    isR3311001Verified,
                    3311001,
                    @"[In messageContent Element] [PidTagMid] Whether the property PidTagMid is present that is optional, it MUST be present in FastTransfer streams created by the RopFastTransferSourceCopyFolder (section 2.2.3.1.1.4) ROP.");
            }
            #endregion
        }

        /// <summary>
        /// Verify the MessageList
        /// </summary>
        /// <param name="fastTransferStream">The fastTransfer Stream get from server</param>
        private void VerifyMessageList(FastTransferStream fastTransferStream)
        {
            MessageList msgList = new MessageList(fastTransferStream);

            if (msgList.Messages.Count == 0)
            {
                return;
            }

            #region PidTagMid
            bool isR479Verified = false;
            bool isR3311Verified = false;
            foreach (PropValue prop in msgList.Messages[0].Content.PropList.PropValues)
            {
                if (prop.PropInfo.PropID == 0x674A)
                {
                    isR3311Verified = true;

                    FixedPropTypePropValue fixedProp = (FixedPropTypePropValue)prop;
                    if ((long)fixedProp.FixedValue == this.messageIdForFastTransferSourceCopyMessages)
                    {
                        isR479Verified = true;
                    }
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3311");

            Site.CaptureRequirementIfIsTrue(
                isR3311Verified,
                3311,
                @"[In messageContent Element] [PidTagMid] Whether the property PidTagMid is present that is optional, it MUST be present in FastTransfer streams created by the RopFastTransferSourceCopyMessages (section 2.2.3.1.1.3) ROP.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R479");

            Site.CaptureRequirementIfIsTrue(
                isR479Verified,
                479,
                @"[In RopFastTransferSourceCopyMessages ROP Request Buffer] [MessageIds (variable)] This list specifies the MID structures ([MS-OXCDATA] section 2.2.1.2) of the messages to be copied.");

            #endregion

            if (this.copyFlagForFastTransferSourceCopyMessages == RopFastTransferSourceCopyMessagesCopyFlags.SendEntryId)
            {
                bool isR3132Verified = false;
                bool isPidTagSourceKeyExist = false;
                bool isPidTagChangeKeyExist = false;
                bool isPidTagLastModificationTimeExist = false;
                bool isPidTagPredecessorChangeListExist = false;
                foreach (PropValue prop in msgList.Messages[0].Content.PropList.PropValues)
                {
                    if (prop.PropInfo.PropID == 0x65E0)
                    {
                        isPidTagSourceKeyExist = true;
                    }

                    if (prop.PropInfo.PropID == 0x65E2)
                    {
                        isPidTagChangeKeyExist = true;
                    }

                    if (prop.PropInfo.PropID == 0x3008)
                    {
                        isPidTagLastModificationTimeExist = true;
                    }

                    if (prop.PropInfo.PropID == 0x65E3)
                    {
                        isPidTagPredecessorChangeListExist = true;
                    }
                }

                isR3132Verified = isPidTagSourceKeyExist && isPidTagChangeKeyExist && isPidTagLastModificationTimeExist && isPidTagPredecessorChangeListExist;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3132");

                Site.CaptureRequirementIfIsTrue(
                    isR3132Verified,
                    3132,
                    @"[In RopFastTransferSourceCopyMessages ROP Request Buffer] [CopyFlags] [When the name is SendEntryId(0x20)] If this bit flag is set, the PidTagSourceKey (section 2.2.1.2.5), PidTagChangeKey (section 2.2.1.2.7), PidTagLastModificationTime ([MS-OXCMSG] section 2.2.2.2) and PidTagPredecessorChangeList (section 2.2.1.2.8) properties are included in the FastTransfer stream. ");
            }
            else
            {
                bool isPidTagSourceKeyExist = false;
                bool isPidTagChangeKeyExist = false;
                bool isPidTagLastModificationTimeExist = false;
                bool isPidTagPredecessorChangeListExist = false;
                bool isPidTagOriginalEntryIdExist = false;
                foreach (PropValue prop in msgList.Messages[0].Content.PropList.PropValues)
                {
                    if (prop.PropInfo.PropID == 0x65E0)
                    {
                        isPidTagSourceKeyExist = true;
                    }

                    if (prop.PropInfo.PropID == 0x65E2)
                    {
                        isPidTagChangeKeyExist = true;
                    }

                    if (prop.PropInfo.PropID == 0x3008)
                    {
                        isPidTagLastModificationTimeExist = true;
                    }

                    if (prop.PropInfo.PropID == 0x65E3)
                    {
                        isPidTagPredecessorChangeListExist = true;
                    }

                    if (prop.PropInfo.PropID == 0x3A12)
                    {
                        isPidTagOriginalEntryIdExist = true;
                    }
                }

                if (Common.IsRequirementEnabled(3134, this.Site))
                {
                    bool isR3134Verified = isPidTagSourceKeyExist == isPidTagChangeKeyExist == isPidTagLastModificationTimeExist == isPidTagPredecessorChangeListExist == isPidTagOriginalEntryIdExist == false;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3134");

                    Site.CaptureRequirementIfIsTrue(
                        isR3134Verified,
                        3134,
                        @"[In RopFastTransferSourceCopyMessages ROP Request Buffer] [CopyFlags] [When the name is SendEntryId(0x20)] If this bit flag is not set, the PidTagSourceKey, PidTagChangeKey, PidTagLastModificationTime, PidTagPredecessorChangeList, and PidTagOriginalEntryId properties are not included in the FastTransfer stream.");
                }
            }
        }

        /// <summary>
        /// Verify the State 
        /// </summary>
        /// <param name="fastTransferStream">The fastTransfer Stream get from server</param>
        private void VerifyState(FastTransferStream fastTransferStream)
        {
            #region Verify state
            State transferState = new State(fastTransferStream);

            // When deserialize the FastTransfer Stream from the server response will as this requirement say,
            // if deserialize successfully this requirement will be covered
            Site.CaptureRequirement(
                2026,
                @"[In  Creating Compact IDSET Structures]Note that this MUST NOT be done for IDSET structures that are ever iterated over, such as the
                MetaTagIdsetGiven property (section 2.2.1.1.1), as it will change the outcome.");

            // When deserialize the FastTransfer Stream from the server response will as this requirement say,
            // if deserialize successfully, this requirement will be covered
            Site.CaptureRequirement(
                1945,
                @"[In IDSET Serialization]The context in which the serialized IDSET structure is being used on the wire dictates which format MUST be used: if 
                an IDSET was persisted or is intended to be persisted across sessions, such as when it represents a portion of an ICS state, as specified in 
                section 2.2.1.1, it MUST be transmitted in the REPLGUID-based form.");

            // When deserialize the FastTransfer Stream from the server response will as this requirement say,
            // if deserialize successfully this requirement will be covered
            Site.CaptureRequirement(
                1946,
                @"[In IDSET Serialization]If it's only a part of a transient set of data, like IDs of items that were deleted since the last synchronization, as 
                specified in section 2.2.1.3.1, it MUST be transmitted in a REPLID-based form.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R982");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R982
            // When deserialize buffer this requirement will be verified and get a boolean value to indicates whether this requirement be test.
            bool isVerifyR982 = transferState.IsNotSplitedInSingleItem;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR982,
                982,
                @"[In Lexical structure]A stream MUST NOT be split within a single atom.");

            // The stream is deserialized as this requirement, if deserialize successfully, this requirement can be covered.
            Site.CaptureRequirement(
                1203,
                @"[In root Element]Producers of the FastTransfer stream MUST choose a contained element to generate depending on the Bulk Data Transfer 
                operation in effect.");

            if (transferState != null && transferState.PropList != null && transferState.PropList.PropValues != null)
            {
                foreach (PropValue prop in transferState.PropList.PropValues)
                {
                    #region MetaTagIdsetRead (Property ID = 0x402D) Meta-Property.

                    if (prop.PropInfo.PropID == 0x402d)
                    {
                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3035,
                            @"[In MetaTagIdsetRead Meta-Property] [MetaTagIdsetRead property] Property ID: 0x402D");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3665");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3665
                        this.Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0102,
                            prop.PropType,
                            3665,
                            @"[In MetaTagIdsetRead Meta-Property] [MetaTagIdsetRead property]: Data type: PtypBinary, 0x0102 ([MS-OXCDATA] section 2.11.1)");

                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3036,
                            @"[In MetaTagIdsetRead Meta-Property] The MetaTagIdsetRead property contains a serialization of REPLID-based IDSET structures. ");
                    }

                    #endregion

                    #region MetaTagIdsetUnread (Property ID = 0x402E) Meta-Property.

                    if (prop.PropInfo.PropID == 0x402e)
                    {
                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3037,
                            @"[In MetaTagIdsetUnread Meta-Property] [MetaTagIdsetUnread property] Property ID: 0x402E");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3666");

                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3666
                        this.Site.CaptureRequirementIfAreEqual<ushort>(
                            0x0102,
                            prop.PropType,
                            3666,
                            @"[In MetaTagIdsetUnread Meta-Property] [MetaTagIdsetUnread property]: Data type: PtypBinary, 0x0102 ([MS-OXCDATA] section 2.11.1)");

                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3038,
                            @"[In MetaTagIdsetUnread Meta-Property] The MetaTagIdsetUnread property contains a serialization of REPLID-based IDSET structures. ");
                    }

                    #endregion

                    // 0x4017 is the ID of PidTagIdsetGiven, 0x6796 is the ID of PidTagCnsetSeen
                    // 0x67DA is the ID of PidTagCnsetSeenFAI, 0x67D2 is the ID of PidTagCnsetRead
                    if ((prop.PropInfo.PropID == 0x4017) || (prop.PropInfo.PropID == 0x6796)
                        || (prop.PropInfo.PropID == 0x67DA) || (prop.PropInfo.PropID == 0x67D2))
                    {
                        #region IDSET UID GUID Command
                        VarPropTypePropValue varProp = (VarPropTypePropValue)prop;

                        // Transform the value of this property to a memory stream, the property value is a serialized IDSET.
                        using (MemoryStream memoryStream = new MemoryStream(varProp.ValueArray))
                        {
                            REPLGUID_IDSET idset = new REPLGUID_IDSET();

                            // Deserialize the above stream to get an IDSET.
                            // "-1" means no limitation except stream length.
                            idset.Deserialize(memoryStream, -1);

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1934");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1934
                            // The structure is deserialized as this requirement, if Deserialize success, set IsAllGLOBCNTRanged as true
                            bool isVerifyR1934 = idset.IsAllGLOBCNTRanged;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1934,
                                1934,
                                @"[In Formatted IDSET Structures]. The ID values MUST be arranged by REPLID structure value");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1935");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1935
                            // The structure is deserialized as this requirement, if Deserialize success, set IsAllGLOBCNTInGLOBSET as true
                            bool isVerifyR1935 = idset.IsAllGLOBCNTInGLOBSET;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1935,
                                1935,
                                @"[In Formatted IDSET Structures]all IDs for each REPLID MUST be reduced into a GLOBSET, as specified in section 2.2.2.6,
                            of GLOBCNTs, as specified in section 2.2.2.5.");

                            // Add the debug information
                            // The structure is deserialized as this requirement, if Deserialize success, set IsAllGLOBCNTInGLOBSET as true
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1936");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1936
                            bool isVerifyR1936 = idset.HasAllDuplicateGLOBCNTRemoved;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1936,
                                1936,
                                @"[In Formatted IDSET Structures]Each GLOBSET MUST be arranged from lowest to highest GLOBCNT value where all duplicate GLOBCNT values are removed.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1937");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1937
                            // The structure is deserialized as this requirement, if Deserialize success, set HasGLOBCNTGroupedIntoRanges as true
                            bool isVerifyR1937 = idset.HasGLOBCNTGroupedIntoRanges;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1937,
                                1937,
                                @"[In Formatted IDSET Structures]The remaining GLOBCNT values MUST be grouped into consecutive ranges with a low GLOBCNT value and a high GLOBCNT value.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1938");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1938
                            // The structure is deserialized as this requirement, if Deserialize success, set IsDisjointGLOBCNTMadeIntoSingleton as true
                            bool isVerifyR1938 = idset.IsDisjointGLOBCNTMadeIntoSingleton;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR1938,
                                1938,
                                @"[In Formatted IDSET Structures]If a GLOBCNT value is disjoint it MUST be made into a singleton range with the low and high GLOBCNT values being the same.");

                            foreach (REPLGUID_IDSETElement replguid in idset.IDSetList)
                            {
                                // The response from the server is serialized as this requirement is related to server, 
                                // if de-serialized the response (as the rule for server serializes the response) successfully, this requirement could be verified
                                Site.CaptureRequirement(
                                    186,
                                    @"[In Serialized IDSET Structure Containing a REPLGUID Structure] REPLGUID-GLOBSET structure pairs MUST be serialized by the value of the REPLGUID in the ascending order, using byte-to-byte comparison.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R187");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R187
                                Site.CaptureRequirementIfAreEqual<Type>(
                                    typeof(Guid),
                                    replguid.REPLGUID.GetType(),
                                    187,
                                    @"[In Serialized IDSET Structure Containing a REPLGUID Structure]REPLGUID (16 bytes): A GUID that identifies a REPLGUID structure.");

                                for (int i = 0; i < replguid.GLOBSET.DeserializedCommandList.Count; i++)
                                {
                                    if (replguid.GLOBSET.DeserializedCommandList[i] is PushCommand)
                                    {
                                        // comByteStack is the common byte stack
                                        CommonByteStack comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, new CommonByteStack());

                                        // The deserialization do as:
                                        // The Push command SHOULD be used when multiple GLOBCNT values share the same high-order values.
                                        // When push command is encountered, multiple GLOBCNT values share the same high-order values in Exchange
                                        bool pushHasSameHighOrderValue = this.WhetherPushHasSameHighOrderValue(replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTList);
                                        if (pushHasSameHighOrderValue)
                                        {
                                            // Verify requirement MS-OXCFXICS_R2660
                                            bool isR2660Satisfied = true;
                                            if (null != replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTList)
                                            {
                                                foreach (GLOBCNT globcnt in replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTList)
                                                {
                                                    for (int index = 0; index < comByteStack.Bytes; index++)
                                                    {
                                                        if (globcnt[index] != comByteStack.GetCommonBytes()[index])
                                                        {
                                                            // The multiple GLOBCNT structure values share the same higher-order values but Push command is not used.
                                                            isR2660Satisfied = false;
                                                            break;
                                                        }
                                                    }
                                                }

                                                if (Common.IsRequirementEnabled(2660, this.Site))
                                                {
                                                    // Add the debug information
                                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2660");

                                                    // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2660
                                                    Site.CaptureRequirementIfIsTrue(
                                                        isR2660Satisfied,
                                                        2660,
                                                        @"[In Push Command (0x01 - 0x06)] Implementation does use the Push command when multiple GLOBCNT structure values share the same high-order values. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                                }
                                            }
                                        }

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R202");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R202
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            (replguid.GLOBSET.DeserializedCommandList[i] as PushCommand).CommandByte,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as PushCommand).CommonBytes.Length,
                                            202,
                                            @"[In Push Command (0x01 - 0x06)] [CommonBytes (variable):] The length of this field is specified by value of the Command field (0x01 through 0x06), in bytes.");
                                    }

                                    if (replguid.GLOBSET.DeserializedCommandList[i] is PopCommand)
                                    {
                                        // comByteStack is the common byte stack
                                        CommonByteStack comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, new CommonByteStack());

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1966");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1966
                                        // from this requirement, can know if "Pop Command" is used, there are some bytes on the common byte stack. 
                                        // So if "Pop Command" is used and the common byte stack has some bytes, then this requirement can be covered.
                                        bool isVerifyR1966 = comByteStack.Bytes > 0;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR1966,
                                            1966,
                                            @"[In Pop Command (0x50)]The Pop command MUST NOT be used if no bytes are currently on the common byte stack.");
                                    }

                                    if (replguid.GLOBSET.DeserializedCommandList[i] is RangeCommand)
                                    {
                                        bool rangHasSameHighOrderValue = this.WhetherRangHasSameHighOrderValue(replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTRangeList);

                                        // comByteStack is the common byte stack
                                        CommonByteStack comByteStack = new CommonByteStack();
                                        if (rangHasSameHighOrderValue)
                                        {
                                            // Check the commands before current one in DeserializedCommandList,
                                            comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, comByteStack);

                                            if (Common.IsRequirementEnabled(2661, this.Site))
                                            {
                                                // Verify requirement MS-OXCFXICS_R2661
                                                bool isR2661Satisfied = true;

                                                for (int index = 0; index < comByteStack.Bytes; index++)
                                                {
                                                    if (replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTRangeList[0].StartGLOBCNT[index] != comByteStack.GetCommonBytes()[index] || replguid.GLOBSET.DeserializedCommandList[i].CorrespondingGLOBCNTRangeList[0].EndGLOBCNT[index] != comByteStack.GetCommonBytes()[index])
                                                    {
                                                        isR2661Satisfied = false;
                                                        break;
                                                    }
                                                }

                                                // Add the debug information
                                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2661");

                                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2661
                                                // StartGLOBCNT and EndGLOBCNT are the low and high GLOBCNT values. 
                                                // They will share the high-order bytes which are in the common byte stack.
                                                Site.CaptureRequirementIfIsTrue(
                                                     isR2661Satisfied,
                                                     2661,
                                                     @"[In Range Command (0x52)] Implementation does push these common high-order bytes, that are shared by the low and high GLOBCNT values, onto the common byte stack by using the Push command, as specified in section 2.2.2.6.1, prior to using the Range command. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                            }
                                        }

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R213");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R213
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            6 - comByteStack.Bytes,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue.Length,
                                            213,
                                            @"[In Range Command (0x52)] [LowValue (variable):] The number of bytes in this field [LowValue] is equal to 6 minus the number of high-order bytes in the common byte stack.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R216");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R216
                                        Site.CaptureRequirementIfAreEqual<int>(
                                            6 - comByteStack.Bytes,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue.Length,
                                            216,
                                            @"[In Range Command (0x52)] [HighValue (variable):] The number of bytes in this field [HighValue] is equal to 6 minus the number of high-order bytes in the common byte stack.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R214");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R214
                                        bool isVerifyR214 = true;
                                        for (int j = 0; j < 6 - comByteStack.Bytes; j++)
                                        {
                                            if ((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue[j] > (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue[j])
                                            {
                                                break;
                                            }
                                            else if ((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue[j] == (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue[j])
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                isVerifyR214 = false;
                                                break;
                                            }
                                        }

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR214,
                                            214,
                                            @"[In Range Command (0x52)] [LowValue (variable):] This value MUST be less than or equal to the value of the HighValue field.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R217");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R217
                                        bool isVerifyR217 = isVerifyR214;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR217,
                                            217,
                                            @"[In Range Command (0x52)] [HighValue (variable):] This value MUST be greater than or equal to the value of the LowValue field.");

                                        // When deserialize the FastTransfer Stream from the server response will as this requirement say,
                                        // if deserialize successfully, this method will be called , this requirement will be covered.
                                        Site.CaptureRequirement(
                                            1980,
                                            @"[In Range Command (0x52)]If the low and high value of the GLOBCNT range are not the same, the Range command MUST be used.");

                                        // When deserialize the FastTransfer Stream from the server response will as this requirement say,
                                        // if deserialize successfully, this method will be called , this requirement will be covered.
                                        Site.CaptureRequirement(
                                            1981,
                                            @"[In Range Command (0x52)]If the range has values that are more than 8 bytes from each other, the Range command MUST be used.");

                                        // When deserialize the FastTransfer Stream from the server response will as this requirement say,
                                        // if deserialize successfully, this method will be called , this requirement will be covered.
                                        Site.CaptureRequirement(
                                            1982,
                                            @"[In Range Command (0x52)]If the low and high value do not share five high-order bytes in common, the Range command MUST be used.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2009");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2009
                                        bool isVerifyR2009 = true;
                                        foreach (GLOBCNTRange gcntR in (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).CorrespondingGLOBCNTRangeList)
                                        {
                                            if (!replguid.GLOBSET.GLOBCNTRangeList.Contains(gcntR))
                                            {
                                                isVerifyR2009 = false;
                                                break;
                                            }
                                        }

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR2009,
                                            2009,
                                            @"[In Range Command (0x52)]The GLOBCNT structure range MUST be added to the GLOBSET structure.");

                                        byte[] lowValueOfRange = StructureSerializer.Serialize((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).CorrespondingGLOBCNTRangeList[0].StartGLOBCNT);
                                        byte[] highValueOfRange = StructureSerializer.Serialize((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).CorrespondingGLOBCNTRangeList[0].EndGLOBCNT);

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2012");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2012
                                        bool isVerifyR2012 = true;
                                        for (int index = 1; index <= (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue.Length; index++)
                                        {
                                            if ((replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue[(replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).LowValue.Length - index] != lowValueOfRange[lowValueOfRange.Length - index]
                                             || (replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue[(replguid.GLOBSET.DeserializedCommandList[i] as RangeCommand).HighValue.Length - index] != highValueOfRange[highValueOfRange.Length - index])
                                            {
                                                isVerifyR2012 = false;
                                                break;
                                            }
                                        }

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR2012,
                                            2012,
                                            @"[In Range Command (0x52)]Each of these fields MUST be combined with any high-order bytes in the common byte stack to produce a 6-byte GLOBCNT structure.");
                                    }

                                    if (replguid.GLOBSET.DeserializedCommandList[i] is BitmaskCommand)
                                    {
                                        // comByteStack is the common byte stack
                                        CommonByteStack comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, new CommonByteStack());

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1970");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1970
                                        bool isVerifyR1970 = comByteStack.Bytes == 5;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR1970,
                                            1970,
                                            @"[In Bitmask Command (0x42)]There MUST already be five high-order bytes in the common byte stack to use this command.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2007");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2007
                                        // "globcnt" is a variable, it is constituted by the "low-order" and "five high-order bytes",
                                        // so if GLOBSET contains "globcnt", then this requirement can be covered.
                                        GLOBCNT globcnt = new GLOBCNT
                                        {
                                            Byte1 = comByteStack.GetCommonBytes()[0],
                                            Byte2 = comByteStack.GetCommonBytes()[1],
                                            Byte3 = comByteStack.GetCommonBytes()[2],
                                            Byte4 = comByteStack.GetCommonBytes()[3],
                                            Byte5 = comByteStack.GetCommonBytes()[4],
                                            Byte6 = (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue
                                        };

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3059");

                                        // The code can reach here, which means the GLOBCNT has been parsed successfully according to 6 byte counter.
                                        Site.CaptureRequirement(
                                            3059,
                                            @"[In GLOBCNT Structure] A GLOBCNT structure is a 6-byte global namespace counter. ");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2007");

                                        bool isVerifyR2007 = false;
                                        byte start_value = (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue;
                                        byte bitmask = (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).Bitmask;

                                        // Start value combined with the required five high-order bytes
                                        // on the common byte stack must be included in the set.
                                        if (replguid.GLOBSET.GLOBCNTList.Contains(globcnt))
                                        {
                                            int j;

                                            // Generates low-order byte, combined with the required five high-order bytes
                                            // on the common byte stack
                                            for (j = 0; j < 8; j++)
                                            {
                                                if ((bitmask & (1 << j)) != 0)
                                                {
                                                    globcnt.Byte6 = (byte)(start_value + j + 1);
                                                    if (!replguid.GLOBSET.GLOBCNTList.Contains(globcnt))
                                                    {
                                                        break;
                                                    }
                                                }
                                            }

                                            // no break j equals to 8.if break one low-order byte MUST be combined with the required five high-order bytes
                                            // on the common byte stack is not in the set.
                                            isVerifyR2007 = j == 8;
                                        }

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR2007,
                                            2007,
                                            @"[In Bitmask Command (0x42)]Each low-order byte MUST be combined with the required five high-order bytes on the common
                                        byte stack to form a complete 6-byte GLOBCNT structure, as specified in section 2.2.2.5, which MUST be added to the
                                        GLOBSET structure.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1973");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1973
                                        Site.CaptureRequirementIfAreEqual<byte>(
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).CorrespondingGLOBCNTList[0].Byte6,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue,
                                            1973,
                                            @"[In Bitmask Command (0x42)]The StartingValue field, as specified in section 2.2.2.6.3, MUST be set to the low-order byte of the low value of the first GLOBCNT range.");

                                        byte byt = 0;
                                        foreach (GLOBCNT gcnt in (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).CorrespondingGLOBCNTList)
                                        {
                                            byt |= (byte)(1 << ((int)gcnt.Byte6 - (int)(replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).StartValue - 1));
                                        }

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1974");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1974
                                        Site.CaptureRequirementIfAreEqual<byte>(
                                            byt,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).Bitmask,
                                            1974,
                                            @"[In Bitmask Command (0x42)]The Bitmask field, specified in section 2.2.2.6.3, MUST have one bit set for each value within a range, excluding the low value of the first GLOBCNT range.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1978");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1978
                                        Site.CaptureRequirementIfAreEqual<byte>(
                                            byt,
                                            (replguid.GLOBSET.DeserializedCommandList[i] as BitmaskCommand).Bitmask,
                                            1978,
                                            @"[In Bitmask Command (0x42)] For all GLOBCNTs between ranges, the bit associated with the value is not set in the bitmask.");
                                    }

                                    if (replguid.GLOBSET.DeserializedCommandList[i] is EndCommand)
                                    {
                                        CommonByteStack comByteStack = this.GetCommonByteStack(i, replguid.GLOBSET.DeserializedCommandList, new CommonByteStack());

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1992");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1992
                                        bool isVerifyR1992 = comByteStack.Bytes == 0;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR1992,
                                            1992,
                                            @"[In End Command (0x00)]The End command can only be used if the common byte stack is empty.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1989");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1989
                                        // When de-serializing the FastTransfer Stream from the server response according to this requirement description,
                                        // so that if the de-serialization is successful, only need to verify that there is no byte on the common byte stack when "End command" is used.
                                        bool isVerifyR1989 = comByteStack.Bytes == 0;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR1989,
                                            1989,
                                            @"[In End Command (0x00)]If after all GLOBCNT ranges have been encoded, there are still bytes on the common byte stack, 
                                        they MUST be removed with one or more Pop commands, as specified in section 2.2.2.6.1, before the End command can be
                                        used.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2014");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2014
                                        // EndCommand must be the last Command in GLOBSET.
                                        bool isVerifyR2014 = i == replguid.GLOBSET.DeserializedCommandList.Count - 1;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR2014,
                                            2014,
                                            @"[In End Command (0x00)]When the End command, as specified in section 2.2.2.6.5, is encountered, the GLOBSET structure
                                        MUST be complete based on the GLOBCNT structure values, as specified in section 2.2.2.5, generated from any previous
                                        encoding commands.");

                                        // Add the debug information
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1987");

                                        // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1987
                                        // EndCommand must be the last Command in GLOBSET.
                                        bool isVerifyR1987 = i == replguid.GLOBSET.DeserializedCommandList.Count - 1;

                                        Site.CaptureRequirementIfIsTrue(
                                            isVerifyR1987,
                                            1987,
                                            @"[In End Command (0x00)]This command MUST be added after all GLOBCNT structure ranges, as specified in section 2.2.2.5,
                                        within the GLOBSET have been encoded.");
                                    }
                                }
                            }
                        }
                        #endregion

                        #region PidTagIdsetGiven
                        if (prop.PropInfo.PropID == 0x4017)
                        {
                            if (Common.IsRequirementEnabled(33541, this.Site))
                            {
                                // Add the debug information
                                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R33541");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R33541
                                this.Site.CaptureRequirementIfAreEqual<ushort>(
                                    0x0003,
                                    prop.PropType,
                                    33541,
                                    @"[In Sending and Receiving the PidTagIdsetGiven ICS State Property] Implementation does send the PidTagIdsetGiven property (section 2.2.1.1.1) with a property tag that defines it as PtypInteger32. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                            }

                            // If the program can execute here, this requirement can be verified directly.
                            this.Site.CaptureRequirement(
                                3006,
                                @"[In MetaTagIdsetGiven ICS State Property] [MetaTagIdsetGiven property] Property ID: 0x4017");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3657");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3657
                            this.Site.CaptureRequirementIfAreEqual<ushort>(
                                0x0003,
                                prop.PropType,
                                3657,
                                @"[In MetaTagIdsetGiven ICS State Property] [MetaTagIdsetGiven property] Data type: PtypInteger32, 0x0003 ([MS-OXCDATA] section 2.11.1)");

                            // If the program can execute here, this requirement can be verified directly.
                            this.Site.CaptureRequirement(
                                3007,
                                @"[In MetaTagIdsetGiven ICS State Property] The MetaTagIdsetGiven property contains a serialization of REPLGUID-based IDSET structures, as specified in section 2.2.2.4.2. ");

                            // Verify requirement MS-OXCFXICS_R2656
                            // 0x0003 is PtypInteger32
                            bool isR2656Satisfied = prop.PropType == 0x0003;

                            if (Common.IsRequirementEnabled(2656, this.Site))
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2656");

                                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2656
                                Site.CaptureRequirementIfIsTrue(
                                    isR2656Satisfied,
                                    2656,
                                    @"[In Receiving the MetaTagIdsetGiven ICS State Property] Implementation does send the MetaTagIdsetGiven property (section 2.2.1.1.1) with a property tag that defines it as PtypInteger32. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                            }
                        }
                        #endregion

                        #region PidTagCnsetSeen
                        if (prop.PropInfo.PropID == 0x6796)
                        {
                            // If the program can execute here, this requirement can be verified directly.
                            this.Site.CaptureRequirement(
                                3009,
                                @"[In MetaTagCnsetSeen ICS State Property] [MetaTagCnsetSeen property] Property ID: 0x6796");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3659");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3659
                            this.Site.CaptureRequirementIfAreEqual<ushort>(
                                0x0102,
                                prop.PropType,
                                3659,
                                @"[In MetaTagCnsetSeen ICS State Property] [MetaTagCnsetSeen property] Data type: PtypBinary, 0x0102 ([MS-OXCDATA] section 2.11.1)");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1264");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1264
                            // 0x6796 is the property id of PidTagCnsetSeen according [MS-OXPROPS]
                            // 0x0102 is PtypBinary
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                 0x0102,
                                 prop.PropType,
                                 1264,
                                 @"[In MetaTagCnsetSeen ICS State Property]The MetaTagCnsetSeen property contains a serialization of REPLGUID-based CNSET structures, as specified in section 2.2.2.4.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2707");

                            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2707
                            // PidTagCnsetSeen is a Binary property, if its type is 0x0102, this requirement can be verified
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                        0x0102,
                                        prop.PropType,
                                        "MS-OXCDATA",
                                        2707,
                                        @"[In Property Data Types] PtypBinary (PT_BINARY) is that variable size; a COUNT field followed by that many bytes with Property Type Value 0x0102,%x02.01.");
                        }
                        #endregion

                        #region PidTagCnsetSeenFAI
                        if (prop.PropInfo.PropID == 0x67DA)
                        {
                            // If the program can execute here, this requirement can be verified directly.
                            this.Site.CaptureRequirement(
                                3660,
                                @"[In MetaTagCnsetSeenFAI ICS State Property] [MetaTagCnsetSeenFAI property] Property ID: 0x67DA");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1265");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1265
                            // 0x67DA is the property id of PidTagCnsetSeenFAI according [MS-OXPROPS]
                            // 0x0102 is PtypBinary
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                 0x0102,
                                 prop.PropType,
                                 1265,
                                 @"[In MetaTagCnsetSeenFAI ICS State Property]Data type: PtypBinary, 0x0102 ([MS-OXCDATA] section 2.11.1)");

                            // If the program can execute here, this requirement can be verified directly.
                            this.Site.CaptureRequirement(
                                3010,
                                @"[In MetaTagCnsetSeenFAI ICS State Property] The MetaTagCnsetSeenFAI property contains a serialization of REPLGUID-based IDSET structures, as specified in section 2.2.2.4. ");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1265");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1265
                            // 0x67DA is the property id of PidTagCnsetSeenFAI according [MS-OXPROPS]
                            // 0x0102 is PtypBinary
                            Site.CaptureRequirementIfAreEqual<ushort>(
                                 0x0102,
                                 prop.PropType,
                                 1265,
                                 @"[In MetaTagCnsetSeenFAI ICS State Property]Data type: PtypBinary, 0x0102 ([MS-OXCDATA] section 2.11.1)");
                        }
                        #endregion

                        #region PidTagCnsetRead
                        if (prop.PropInfo.PropID == 0x67D2)
                        {
                            // If the program can execute here, this requirement can be verified directly.
                            this.Site.CaptureRequirement(
                                3012,
                                @"[In MetaTagCnsetRead ICS State Property] [MetaTagCnsetRead property] Property ID: 0x67D2");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3661");

                            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3661
                            this.Site.CaptureRequirementIfAreEqual<ushort>(
                                0x0102,
                                prop.PropType,
                                3661,
                                @"[In MetaTagCnsetRead ICS State Property] [MetaTagCnsetRead property] Data type: PtypBinary, 0x0102 ([MS-OXCDATA] section 2.11.1)");

                            // If the program can execute here, this requirement can be verified directly.
                            this.Site.CaptureRequirement(
                                3013,
                                @"[In MetaTagCnsetRead ICS State Property] The MetaTagCnsetRead property contains a serialization of REPLGUID-based CNSET structures, as specified in section 2.2.2.4. ");
                        }
                        #endregion
                    }
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1212");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R1212
            bool isVerifyR1212 = true;
            if (transferState != null && transferState.PropList != null && transferState.PropList.PropValues != null)
            {
                foreach (PropValue prop in transferState.PropList.PropValues)
                {
                    if (prop.PropInfo.PropID != 0x4017 && prop.PropInfo.PropID != 0x6796 && prop.PropInfo.PropID != 0x67DA && prop.PropInfo.PropID != 0x67D2)
                    {
                        isVerifyR1212 = false;
                        break;
                    }
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1212,
                1212,
                @"[In state Element] [other properties,Prohibited] The other properties[the properties except MetaTagIdsetGiven, MetaTagCnsetSeen, MetaTagCnsetSeenFAI, MetaTagCnsetRead] MUST NOT be present in the propList.");

            #endregion Verify state
        }

        /// <summary>
        /// Verify the Markers.
        /// </summary>
        /// <param name="obj">An object on the fast transfer stream.</param>
        private void VerifyMarkers(object obj)
        {
            // Because the length is a const value 4, so when the deserialize success this requirement can be captured.
            this.Site.CaptureRequirement(
                1050,
                "[In Syntactical Structure] Therefore, wherever a marker exists, it is serialized as 4 bytes.");

            // If the object has a State type, the requirements with the markers (IncrSyncStateBegin, IncrSyncStateBegin) can be verified.
            if (obj is State)
            {
                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1021,
                    @"[In Markers] [Synchronization download] The numeric value 0x403A0003 means IncrSyncStateBegin marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1022,
                    @"[In Markers] [Synchronization download] The numeric value 0x403B0003 means IncrSyncStateEnd marker.");
            }
            else if (obj is AttachmentContent)
            {
                // If the object has an AttachmentContent type, the requirements with the markers (StartEmbed, EndEmbed) can be verified.
                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1010,
                    @"[In Markers] [Messages and their parts] The numeric value 0x40010003 means StartEmbed marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1011,
                    @"[In Markers] [Messages and their parts] The numeric value 0x40020003 means EndEmbed[the end marker of StartEmbed] marker.");
            }
            else if (obj is ContentsSync)
            {
                // If the object has a ContentsSync type, the requirements with the markers 
                // (IncrSyncEnd, IncrSyncProgressMode, IncrSyncProgressPerMsg, IncrSyncChgPartial, IncrSyncGroupInfo, IncrSyncMessage, IncrSyncRead) can be verified.
                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1017,
                    @"[In Markers] [Synchronization download] The numeric value 0x407D0003 means IncrSyncChgPartial marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1019,
                    @"[In Markers] [Synchronization download] The numeric value 0x40140003 means IncrSyncEnd marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1020,
                    @"[In Markers] [Synchronization download] The numeric value 0x402F0003 means IncrSyncRead marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1023,
                    @"[In Markers] [Synchronization download] The numeric value 0x4074000B means IncrSyncProgressMode marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1024,
                    @"[In Markers] [Synchronization download] The numeric value 0x4075000B means IncrSyncProgressPerMsg marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1025,
                    @"[In Markers] [Synchronization download] The numeric value 0x40150003 means IncrSyncMessage marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1026,
                    @"[In Markers] [Synchronization download] The numeric value 0x407B0102 means IncrSyncGroupInfo marker.");
            }
            else if (obj is HierarchySync)
            {
                // If the object has a HierarchySync type, the requirements with the markers (IncrSyncDel, IncrSyncChg) can be verified.
                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1016,
                    @"[In Markers] [Synchronization download] The numeric value 0x40120003 means IncrSyncChg marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1018,
                    @"[In Markers] [Synchronization download] The numeric value 0x40130003 means IncrSyncDel marker.");
            }
            else if (obj is MessageContent)
            {
                // If the object has a MessageContent type, the requirements with the markers (StartRecip, EndToRecip, NewAttach, EndAttach) can be verified.
                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1012,
                    @"[In Markers] [Messages and their parts] The numeric value 0x40030003 means StartRecip marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1013,
                    @"[In Markers] [Messages and their parts] The numeric value 0x40040003 means EndToRecip[the end marker of StartRecip] marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1014,
                    @"[In Markers] [Messages and their parts] The numeric value 0x40000003 means NewAttach marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1015,
                    @"[In Markers] [Messages and their parts] The numeric value 0x400E0003 means EndAttach[the end marker of NewAttach] marker.");
            }
            else if (obj is FolderContent)
            {
                // If the object has a FolderContent type, the requirements with the markers (StartMessage, EndMessage, StartFAIMsg, StartSubFld) can be verified.
                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1004,
                    @"[In Markers] [Folders] The numeric value 0x400A0003 means StartSubFld marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1006,
                    @"[In Markers] [Messages and their parts] The numeric value 0x400C0003 means StartMessage marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    2681,
                    @"[In Markers] [Messages and their parts] The numeric value 0x400D0003 means EndMessage[the end marker of StartMessage, StartFAIMsg] marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1008,
                    @"[In Markers] [Messages and their parts] The numeric value 0x40100003 means StartFAIMsg marker.");
            }
            else if (obj is TopFolder)
            {
                // If the object has a TopFolder type, the requirements with the markers (StartTopFld, EndFolder) can be verified.
                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1002,
                    @"[In Markers] [Folders] The numeric value 0x40090003 means StartTopFld marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1003,
                    @"[In Markers] [Folders] The numeric value 0x400B0003 means EndFolder[the end marker of StartTopFld, StartSubFld] marker.");
            }
            else if (obj is MessageList)
            {
                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1132,
                    @"[In message Element] Normal messages use the StartMessage marker.");

                // If the program can execute here, this requirement can be verified directly.
                this.Site.CaptureRequirement(
                    1133,
                    @"[In message Element] FAI messages use the StartFAIMsg marker.");

                MessageList messageList = obj as MessageList;
                if (messageList.Messages.Count > 0)
                {
                    Message message = messageList.Messages[0];
                    PropValue prop = message.Content.PropList.PropValues.Find(p => p.PropInfo.PropID == 0x674A);
                    if (prop != null)
                    {
                        // Since the the PidTagMid (0x674A) property can be found in the array list of messageContent element, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3311002,
                            "[In messageContent Element] [PidTagMid] Whether the property PidTagMid is present that is optional, it will Fix position of the array list.");
                    }
                }
            }
            else
            {
                return;
            }
        }

        /// <summary>
        /// Verify the Meta-Property.
        /// </summary>
        /// <param name="afts">A parameter with the AbstractFastTransferStream type.</param>
        private void VerifyMetaProperty(AbstractFastTransferStream afts)
        {
            switch (afts.StreamType)
            {
                case FastTransferStreamType.MessageList:
                    if (afts.AbstractMessageList.IsPidTagEcWarningOut == true)
                    {
                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3286,
                            @"[In MetaTagEcWarning Meta-Property] Property ID: 0x400F Data type: PtypInteger32, 0x0003 ([MS-OXCDATA] section 2.11.1)");
                    }

                    break;
                case FastTransferStreamType.folderContent:
                    if (afts.AbstractFolderContent.IsSubFolderPrecededByPidTagFXDelProp == true)
                    {
                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            3285,
                            @"[In MetaTagFXDelProp Meta-Property] Property ID: 0x4016 Data type: PtypInteger32, 0x0003 ([MS-OXCDATA] section 2.11.1)");
                    }

                    break;
                case FastTransferStreamType.TopFolder:
                    if (afts.AbstractFolderContent.IsPidTagNewFXFolderOut == true)
                    {
                        // If the program can execute here, this requirement can be verified directly.
                        this.Site.CaptureRequirement(
                            1037,
                            @"[In MetaTagNewFXFolder Meta-Property] Property ID: 0x4011 Data type: PtypBinary, 0x0102 ([MS-OXCDATA] section 2.11.1)");
                    }

                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// Verify the code page property.
        /// </summary>
        /// <param name="codePage">A code page property.</param>
        private void VerifyCodePageProperty(CodePage codePage)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3256");

            // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3256
            this.Site.CaptureRequirementIfAreEqual<ushort>(
                0x8000,
                codePage.A,
                3256,
                @"[Code Page Property Types] A (1 bit): 1-bit flag (mask 0x8000).");

            // If the program can execute here, this requirement can be verified directly.
            this.Site.CaptureRequirement(
                3257,
                @"[Code Page Property Types] [A bit] This bit MUST be set to 1 to indicate the property is an internal code page string.");
        }

        #region Verify the requirements which is related to MetaTagIdsetDeleted property.

        /// <summary>
        /// Verify the MetaTagIdsetDeleted property.
        /// </summary>
        /// <param name="hasExecuteSynchronizationImportDeletes">indicates whether the SynchronizationImportDelete ROP has been excuted.</param>
        /// <param name="importDelFlag">The import delete flag.</param>
        /// <param name="delObjIdsResponse">The deleted Message ID in response.</param>
        /// <param name="delObjIdRequest">The deleted Message ID in request.</param>
        private void VerifyMetaTagIdsetDeletedProperty(bool hasExecuteSynchronizationImportDeletes, byte importDelFlag, List<ulong> delObjIdsResponse, ulong delObjIdRequest)
        {
            if (hasExecuteSynchronizationImportDeletes == true)
            {
                bool isVerifyRequirement = false;

                for (int i = 0; i < delObjIdsResponse.Count; i++)
                {
                    if (delObjIdsResponse[i] == delObjIdRequest)
                    {
                        isVerifyRequirement = true;
                        break;
                    }
                }

                if ((byte)importDelFlag == (byte)ImportDeleteFlags.Hierarchy)
                {
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyRequirement,
                        72,
                        @"[In MetaTagIdsetDeleted Meta-Property] The IDSETs contain the IDs of folders (for hierarchy synchronization operations) 
                          that were hard deleted since the last synchronization identified by the initial ICS state.");
                }

                if ((byte)importDelFlag == ((byte)ImportDeleteFlags.Hierarchy | (byte)ImportDeleteFlags.HardDelete))
                {
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyRequirement,
                        7201,
                        @"[In MetaTagIdsetDeleted Meta-Property] The IDSETs contain the IDs of folders (for hierarchy synchronization operations) 
                        that were soft deleted since the last synchronization identified by the initial ICS state.");
                }

                if ((byte)importDelFlag == (byte)ImportDeleteFlags.delete)
                {
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyRequirement,
                        72001,
                        @"[In MetaTagIdsetDeleted Meta-Property] The IDSETs contain the IDs of messages (for content synchronization operations) 
                        that were hard deleted since the last synchronization identified by the initial ICS state.");
                }

                if ((byte)importDelFlag == (byte)ImportDeleteFlags.HardDelete)
                {
                    if (!Common.IsRequirementEnabled(2593, this.Site))
                    {
                        Site.CaptureRequirementIfIsTrue(
                            isVerifyRequirement,
                            7200101,
                            @"[In MetaTagIdsetDeleted Meta-Property] The IDSETs contain the IDs of messages (for content synchronization operations) 
                            that were soft deleted since the last synchronization identified by the initial ICS state.");
                    }
                }
            }

            return;
        }

        #endregion

        /// <summary>
        /// Verify push whether Has Same High Order Value
        /// </summary>
        /// <param name="cntList">GLOBCNT list </param>
        /// <returns>If has Same High Order Value return true else return false</returns>
        private bool WhetherPushHasSameHighOrderValue(List<GLOBCNT> cntList)
        {
            // Calculate whether Corresponding GLOBCNTs have same high-order value.
            List<GLOBCNT> cntlist = cntList;
            if (null != cntList && cntList.Count > 0)
            {
                int j = 0;
                byte tmp = cntlist[0].Byte1;
                for (j = 1; j < cntlist.Count; j++)
                {
                    if (tmp != cntlist[j].Byte1)
                    {
                        break;
                    }
                }

                return j == cntlist.Count;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// verify whether Rang Has Same High Order Value
        /// </summary>
        /// <param name="cntRangeList">GLOBCNTRange list</param>
        /// <returns>If has Same High Order Value return true else return false</returns>
        private bool WhetherRangHasSameHighOrderValue(List<GLOBCNTRange> cntRangeList)
        {
            // Calculate whether Corresponding GLOBCNTs have same high-order value.
            List<GLOBCNTRange> cntrangList = cntRangeList;
            if (null != cntrangList && cntrangList.Count > 0)
            {
                int j = 0;
                if (cntrangList[0].GetSameHighOrderValues().Length == 0)
                {
                    return false;
                }

                byte tmp = cntrangList[0].GetSameHighOrderValues()[0];
                for (j = 1; j < cntrangList.Count; j++)
                {
                    if (cntrangList[j].GetSameHighOrderValues().Length == 0)
                    {
                        return false;
                    }

                    if (tmp != cntrangList[j].GetSameHighOrderValues()[0])
                    {
                        break;
                    }
                }

                return j == cntrangList.Count;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// This function called after calling GetPropertiesSpecific ROP to get PidTagResolveMethod property value, if the propertyValuesSpecific field of ROP response contains these properties, it means the propery can be verified.
        /// </summary>
        private void VerifyPidTagResolveMethodValue()
        {
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7942");

            // In order to get the PidTagResolveMethod value, client sends a request with its data type and propID. when the ROP is successful, it means the server can recognize it and return a valid value. So the requirement about property ID can be verified.
            Site.CaptureRequirement(
                       "MS-OXPROPS",
                       7942,
                       @"[In PidTagResolveMethod] Property ID: 0x3FE7.");

            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXPROPS_R7943");

            // In order to get the PidTagResolveMethod value, client sends a request with its data type and propID. when the ROP is successful, it means the server can recognize it and return a valid value. So the requirement about property ID can be verified.
            Site.CaptureRequirement(
                       "MS-OXPROPS",
                       7943,
                       @"[In PidTagResolveMethod] Data type: PtypInteger32, 0x0003.");

            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1285");

            // In order to get the PidTagResolveMethod value, client sends a request with its data type and propID. when the ROP is successful, it means the server can recognize it and return a valid value. So the requirement about property ID can be verified.
            Site.CaptureRequirement(
                       1285,
                       @"[In PidTagResolveMethod Property] Data type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");
        }

        /// <summary>
        /// Verify the conflict handler result by checking the PCL in the response and the change saved in the local.
        /// </summary>
        private void VerifyPidTagPredecessorChangeList()
        {
            if (this.lastConflictInfo.CheckResult() == true)
            {
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1783");

                // If the checking result is true, it means server implements conflict detection using an algorithm compatible with the one described in this protocol. 
                // Since this method is only called after conflict detected and if the successful ROP returned, it indicates the server process it successfully.
                Site.CaptureRequirement(
                           1783,
                           @"[In Detection] Servers MUST implement conflict detection using an algorithm compatible with the one described in this section.");

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1784");

                // If the checking result is true, it means server performs conflict detection on ICS uploads for versions of messaging objects stored in a server replica 
                // and passed by the client. And if the successful ROP is returned, it indicates the server process it successfully. 
                Site.CaptureRequirement(
                           1784,
                           @"[In Detection] Servers MUST perform conflict detection on ICS uploads for versions of messaging objects stored in a server replica and passed by the client through the RopSynchronizationImport* ROPs.");
            }

            if (this.lastConflictInfo.DetectedResult == ConflictTypes.AINCLUDEB && (this.lastConflictInfo.CheckResult() == true))
            {
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1791");

                // When the conflict type is AINCLUDEB, it means PCLA including PCLB.
                // If the checking result is true, it means a change to a messaging object is being synchronized from replica A to replica B, use the following statements to identify the conflict and the version to replicate, so this requirement can be captured.
                Site.CaptureRequirement(
                           1791,
                           @"[In Detection] Conflict Detection Algorithm: If a change to a messaging object is being synchronized from replica A to replica B, use the following statements to identify the conflict and the version to replicate: 1.If PCLA includes PCLB, then the version from replica A is newer and replaces the version in replica B.");
            }

            if (this.lastConflictInfo.DetectedResult == ConflictTypes.BINCLUDEA && (this.lastConflictInfo.CheckResult() == true))
            {
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1792001");

                // When the conflict type is BINCLUDEA, it means PCLB including PCLA. 
                // If the checking result is true, it means the version in replica B remains intact, so this requirement can be captured.
                Site.CaptureRequirement(
                           1792001,
                           @"[In Detection,Conflict Detection Algorithm:If a change to a messaging object is being synchronized from replica A to replica B, use the following statements to identify the conflict and the version to replicate:] [2.If PCLB includes or is equal to PCLA, then the version from replica A is older, and is ignored] The version in replica B remains intact.");

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1792");

                // When the conflict type is BINCLUDEA, it means that PCLB includes PCLA. 
                // If the checking result is true, it means the version from replica A is older and is ignored. So this requirement can be captured.
                Site.CaptureRequirement(
                           1792,
                           @"[In Detection,Conflict Detection Algorithm:If a change to a messaging object is being synchronized from replica A to replica B, use the following statements to identify the conflict and the version to replicate:] 2.If PCLB includes or is equal to PCLA, then the version from replica A is older, and is ignored.");
            }

            if (this.lastConflictInfo.DetectedResult != ConflictTypes.CONFLICT && (this.lastConflictInfo.CheckResult() == true))
            {
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1806");

                // When the conflict type is not conflict, it means its PCLA includes PCLB or PCLB includes PCLA. 
                // If the checking result shows that the expected and actual result are the same, PCLX is a successor of versions A and B if and only if the conflict detection algorithm in section 3.1.5.6.1 determines that X is not in conflict.
                Site.CaptureRequirement(
                           1806,
                           @"[In Resolution] Version X is a successor of versions A and B if and only if the conflict detection algorithm in section 3.1.5.6.1 would determine that X is not in conflict.");
            }

            // When the conflict type is conflict, it indicates that PCLA does not include PCLB and PCLB does not include PCLA either. 
            // If the checking result is true, it means the expected and actual result are the same.
            if (this.lastConflictInfo.DetectedResult == ConflictTypes.CONFLICT && (this.lastConflictInfo.CheckResult() == true))
            {
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1793");

                // When the detection conflict type is conflict, it means neither 1[PCLA includes PCLB] nor 2[PCLB includes or is equal to PCLA] is true, then versions from replicas A and B are in conflict. So this requirement can be captured. 
                Site.CaptureRequirement(
                           1793,
                           @"[In Detection,Conflict Detection Algorithm:If a change to a messaging object is being synchronized from replica A to replica B, use the following statements to identify the conflict and the version to replicate:] 3.If neither 1[If PCLA includes PCLB, then the version from replica A is newer and replaces the version in replica B] nor 2[If PCLB includes or is equal to PCLA, then the version from replica A is older, and is ignored. The version in replica B remains intact] is true, then versions from replicas A and B are in conflict.");

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1804");

                // When the conflict type is PCL, it's a successor of all the conflicting versions. If the checking result is true, it means that the expected result is the same as server. So this requirement can be captured.
                Site.CaptureRequirement(
                           1804,
                           @"[In Resolution] A version that results from conflict resolution MUST have a PCL that makes it a successor of all conflicting versions.");

                bool isRS2734Enabled = false;
                this.CheckRequirementEnabled(2734, out isRS2734Enabled);
                if (isRS2734Enabled == true)
                {
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2734");

                    // When the conflict type is PCL, it's a successor of all the conflicting versions. If the result is successful, it means the expected result is same as the server. So this requirement can be captured.
                    Site.CaptureRequirement(
                               2734,
                               @"[In Resolution] Implementation does assign the successor a PCL created by merging the PCLs of all conflicting versions. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                }

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1807");

                // When the conflict type is PCL, it's a successor of all the conflicting versions and is newer than A and B. If the checking result is true, it means that the expected result is same as the server. So this requirement can be captured.
                Site.CaptureRequirement(
                           1807,
                           @"[In Resolution] Version X is a successor of versions A and B if and only if the conflict detection algorithm in section 3.1.5.6.1 would determine that X is newer than both A and B.");

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1808");

                // When the conflict type is PCL, it's a successor of all conflicting versions and contains A and B. If the checking result is true, it means that the expected result is same as the server. So this requirement can be captured.
                Site.CaptureRequirement(
                           1808,
                           @"[In Resolution] PCLX is a merge of PCLA and PCLB if and only if all of the following statements are true: PCLX ⊆ (PCLA ∪ PCLB)");

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1809");

                // When the conflict type is PCL, it's a successor of all conflicting versions and includes A, If the checking result is true, it means that the expected result is the same as the server. So this requirement can be captured.
                Site.CaptureRequirement(
                           1809,
                           @"[In Resolution] PCLX is a merge of PCLA and PCLB if and only if all of the following statements are true: PCLX ≥ PCLA");

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R1810");

                // When the conflict type is PCL, it's a successor of all conflicting versions and includes B. If the checking result returns success, it means that the expected result is the same as the server. So this requirement can be captured.
                Site.CaptureRequirement(
                           1810,
                           @"[In Resolution] PCLX is a merge of PCLA and PCLB if and only if all of the following statements are true:PCLX ≥ PCLB");

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2237");

                // When the conflict type is conflict, it means a conflict is detected. If the expected and actual result are both success, it means that the server resolves it as the protocol. If this method is called after the ROP success, it indicates the requirement is verified.
                Site.CaptureRequirement(
                           2237,
                           @"[In Receiving a RopSynchronizationImportHierarchyChange Request] If a conflict is detected, the server MUST resolve it as specified in section 3.1.5.6.2 and return Success.");

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2242");

                // When the conflict type is conflict, it means a conflict is detected. If the expected and actual result are both success, it means that the server resolves it as the protocol. If this method is called after the ROP success, it indicates the requirement is verified.
                Site.CaptureRequirement(
                           2242,
                           @"[In Receiving a RopSynchronizationImportHierarchyChange Request] If a conflict has occurred, the server: MUST return a value of Success in the ReturnValue field.");
            }
        }

        /// <summary>
        /// Verify the status for FastTransferSourceGetBuffer
        /// </summary>
        /// <param name="isRunIntoPartial">Indicates if run into Partial status</param>
        /// <param name="isRunIntoNoRoom">Indicates if run into NoRoom status</param>
        /// <param name="isReachedBufferTooSmall">Indicates the ROP response is ecBufferTooSmall</param>
        private void VerifyTransferStatus(bool isRunIntoPartial, bool isRunIntoNoRoom, bool isReachedBufferTooSmall)
        {
            if (isRunIntoPartial)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R568");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R568
                Site.CaptureRequirement(
                    568,
                    @"[In RopFastTransferSourceGetBuffer ROP Response Buffer] [TransferStatus] Partial (0x0001): The FastTransfer stream was split, and more data is available.");
            }

            if (isRunIntoNoRoom)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R3149");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R3149
                Site.CaptureRequirement(
                    3149,
                    @"[In RopFastTransferSourceGetBuffer ROP Response Buffer] [TransferStatus] NoRoom(0x0002): The FastTransfer stream was split, more data is available.");
            }

            if (isReachedBufferTooSmall)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R2625");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R2625
                Site.CaptureRequirement(
                    2625,
                    @"[In propValue Element] Servers MUST fail the operation if the size of data being sent is larger than the maximum size imposed by an implementation, rather than truncating the data.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R262501");

                // Verify MS-OXCFXICS requirement: MS-OXCFXICS_R262501
                Site.CaptureRequirement(
                    262501,
                    @"[In propValue Element] Servers MUST fail the operation if the size of data being received is larger than the maximum size imposed by an implementation, rather than truncating the data.");

                if (this.previousGetBufferResult == RopResult.BufferTooSmall && Common.IsRequirementEnabled(348901, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFXICS_R348901");

                    // This requirement can be verified since the first call to the RopFastTransferSourceGetBuffer ROP is BufferTooSmall error, the second is same error. So verify it.
                    Site.CaptureRequirement(
                        348901,
                        @"[In Receiving a RopFastTransferSourceGetBuffer Request] [Servers SHOULD fail any successive calls to the RopFastTransferSourceGetBuffer ROP if the previous iteration returns a value other than of Success or 0x00000480 in the ReturnValue field] Implementation does fail the successive calls with the same error as previous failed operations. (Section 3.2.5.8.1.5: Microsoft Exchange Server 2007 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// Verify the Internal identifier for RopSaveChangesMessage
        /// </summary>
        /// <param name="messageID">The message id returned from server.</param>
        private void VerifyInternalIdentifier(ulong messageID)
        {
            bool isVerifiedR1909 = !this.objectIdContainer.ContainsValue(messageID);
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR1909,
                1909,
                @"[In Identifying Objects and Maintaining Change Numbers] [Upon successful import of a new or changed object using ICS upload, the server MUST do the following when receiving the RopSaveChangesMessage ROP:] Assign the object an internal identifier.");
        }
    }
}