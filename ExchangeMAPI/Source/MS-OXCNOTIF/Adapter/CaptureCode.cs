namespace Microsoft.Protocols.TestSuites.MS_OXCNOTIF
{
    using System.Runtime.InteropServices;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter requirements capture code for MS-OXCNOTIF server role.
    /// </summary>
    public partial class MS_OXCNOTIFAdapter
    {
        #region MAPIHTTP transport

        /// <summary>
        /// Verify MAPIHTTP transport.
        /// </summary>
        private void VerifyMAPITransport()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http" && Common.IsRequirementEnabled(1340, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R1340");

                // Verify requirement MS-OXCNOTIF_R1340
                // If the transport sequence is MAPIHTTP and the code can reach here, it means that the implementation does support MAPIHTTP transport.
                Site.CaptureRequirement(
                        1340,
                        @"[In Appendix A: Product Behavior] Implementation does support this specification [MS-OXCMAPIHTTP]. (Exchange Server 2013 Service Pack 1 (SP1) follows this behavior.)");
            }
        }
        #endregion

        /// <summary>
        /// Verify Syntax about Transport.
        /// </summary>
        private void VerifyAsyncCallOnRPCTransport()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R2");

            // When the client receive asynchronous RPC notification in the response through test suites by using
            // underlying networking protocols, this method will be invoked, and this requirement will be verified directly.
            Site.CaptureRequirement(
                2,
                @"[In Transport] Asynchronous calls are made on the server by using RPC transport, as specified in [MS-OXCRPC].");
        }

        /// <summary>
        /// Verify Syntax about ROP Transport
        /// </summary>
        private void VerifyROPTransport()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http")
            {
                if (Common.IsRequirementEnabled(475, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R475");

                    // When the client receive asynchronous MAPI notification in the response through test suites by using
                    // underlying networking protocols, this method will be invoked, and this requirement will be verified directly.
                    Site.CaptureRequirement(
                        475,
                        @"[In Appendix A: Product Behavior] Asynchronous calls are made on the server by using the MAPI extensions to HTTP. (Exchange 2013 SP1 follows this behavior.)");
                }

                if (Common.IsRequirementEnabled(517, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R517");

                    // When the client receive asynchronous MAPI notification in the response through test suites by using
                    // underlying networking protocols, this method will be invoked, and this requirement will be verified directly.
                    Site.CaptureRequirement(
                        517,
                        @"[In Appendix A: Product Behavior] Implementation does support the session context cookie. (<11> Section 3.1.1:  The session context cookie was introduced in Exchange 2013 SP1.)");
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R1");

            // If the client send request in ROP request buffers and receive response in ROP response buffers through test suites,
            // this method will be invoked, and this requirement will be verified directly.
            Site.CaptureRequirement(
                1,
                @"[In Transport] The commands specified by this protocol [MS-OXCNOTIF] are sent to and received from the server by using the underlying ROP request buffers and ROP response buffers, respectively, as specified in [MS-OXCROPS].");
        }

        /// <summary>
        /// Verify RopNotify response.
        /// </summary>
        /// <param name="ropNotifyResponse">The response of ropNotify</param>
        private void VerifyRopNotifyResponse(RopNotifyResponse ropNotifyResponse)
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http")
            {
                if (Common.IsRequirementEnabled(498, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R498");

                    // When the client receive asynchronous MAPI notification in the response through test suites by using
                    // underlying networking protocols, this method will be invoked, and this requirement will be verified directly.
                    Site.CaptureRequirement(
                        498,
                        @"[In Appendix A: Product Behavior] This ROP [RopNotify] MUST appear in the Execute request type success response body. (Exchange 2013 SP1 and above follow this behavior.)");
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R240");
        
            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R240
            // If server response a RopNotify, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                ropNotifyResponse,
                240,
                @"[In Sending Notification Details] The server sends notification details to the client by sending the RopNotify ROP response (section 2.2.1.4.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R89");

            // If RopNotify is returned in the response buffer through the test suite, 
            // this method will be invoked, and this requirement will be verified directly.
            Site.CaptureRequirement(
                89,
                @"[In RopNotify ROP] This ROP [RopNotify] MUST appear in response buffers of the EcDoRpcExt2 method, as specified in [MS-OXCRPC] section 3.1.4.2.");

            if (Common.IsRequirementEnabled(346, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R346: the response type is {0}", ropNotifyResponse.GetType().Name);

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R346
                this.Site.CaptureRequirementIfAreEqual<string>(
                    "RopNotifyResponse",
                    ropNotifyResponse.GetType().Name,
                    346,
                    @"[In Appendix A: Product Behavior] Implementation does send a RopNotify ROP response (section 2.2.1.4.1) to the client for each pending notification on the session context that is associated with the client. (Exchange 2007 and above follow this behavior.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R136");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R136
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                Marshal.SizeOf(ropNotifyResponse.NotificationHandle),
                136,
                @"[In RopNotify ROP Response Buffer] It [NotificationHandle] is 4 bytes.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R137");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R137
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                Marshal.SizeOf(ropNotifyResponse.NotificationData.NotificationFlags),
                137,
                @"[In NotificationData Structure] It [NotificationFlags] is 2 bytes.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R97002");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R97002
            Site.CaptureRequirementIfIsInstanceOfType(
                ropNotifyResponse.LogonId,
                typeof(byte),
                97002,
                @"[In RopNotify ROP Response Buffer] [LogonId ] is 1 byte.");

            // Verify NotificationFlags of RopNotify response
            this.VerifyNotificationFlagsOfRopNotify(ropNotifyResponse);

            // Verify RopNotify response for NewMail events
            this.VerifyRopNotifyForNewMail(ropNotifyResponse);

            // Verify RopNotify response for TableModified events
            this.VerifyRopNotifyForTableModified(ropNotifyResponse);

            // Verify RopNotify response for other events(ObjectCreated, ObjectDeleted, ObjectModified, ObjectMoved, ObjectCopied and SearchResult)
            this.VerifyRopNotifyForOtherEvents(ropNotifyResponse);
        }

        /// <summary>
        /// Verify RopNotify response for TableModified events.
        /// </summary>
        /// <param name="ropNotifyResponse">The response of ropNotify</param>
        private void VerifyRopNotifyForTableModified(RopNotifyResponse ropNotifyResponse)
        {
            if (ropNotifyResponse.NotificationData.NotificationType == NotificationType.TableModified)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R139");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R139
                this.Site.CaptureRequirementIfAreEqual<int>(
                    2,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.TableEventType),
                    139,
                    @"[In NotificationData Structure] It [TableEventType] is 2 bytes.");
            }

            // Refer to MS-OXCNOTIF section 2.2.1.4.1.1: TableRowFolderID is available only if the TableEventType field is available and is 0x0003(TableRowAdded), 0x0004(TableRowDeleted), or 0x0005(TableRowModified).
            bool isTableRowFolderIDAvailable = ropNotifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.TableRowAdded ||
                            ropNotifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.TableRowDeleted ||
                            ropNotifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.TableRowModified;
            if (isTableRowFolderIDAvailable)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R140");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R140
                this.Site.CaptureRequirementIfAreEqual<int>(
                    8,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.TableRowFolderID),
                    140,
                    @"[In NotificationData Structure] It [TableRowFolderID] is 8 bytes.");
            }

            // Refer to MS-OXCNOTIF section 2.2.1.4.1.1: TableRowMessageID is available only if bit 0x8000(M bit) is set in the NotificationFlags field and if the TableEventType field is available
            // and is 0x0003(TableRowAdded), 0x0004(TableRowDeleted), or 0x0005(TableRowModified).
            bool isTableRowMessageIDAvailable = (ropNotifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.TableRowAdded ||
                           ropNotifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.TableRowDeleted ||
                           ropNotifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.TableRowModified) &&
                           ((ropNotifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) == (ushort)FlagsBit.M);

            if (isTableRowMessageIDAvailable)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R144");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R144
                this.Site.CaptureRequirementIfAreEqual<int>(
                    8,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.TableRowMessageID),
                    144,
                    @"[In NotificationData Structure] It [TableRowMessageID] is 8 bytes.");
            }

            // Refer to MS-OXCNOTIF section 2.2.1.4.1.1: TableRowInstance is available only if bit 0x8000(M bit) is set in the NotificationFlags field and if the TableEventType field is available
            // and is 0x0003(TableRowAdded), 0x0004(TableRowDeleted), or 0x0005(TableRowModified).
            bool isTableRowInstanceAvalible = (ropNotifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.TableRowAdded ||
                            ropNotifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.TableRowDeleted ||
                            ropNotifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.TableRowModified) &&
                            (ropNotifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) == (ushort)FlagsBit.M;

            if (isTableRowInstanceAvalible)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R148");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R148
                this.Site.CaptureRequirementIfAreEqual<int>(
                    4,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.TableRowInstance),
                    148,
                    @"[In NotificationData Structure] It [TableRowInstance] is 4 bytes.");
            }

            if (ropNotifyResponse.NotificationData.InsertAfterTableRowFolderID != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R152");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R152
                this.Site.CaptureRequirementIfAreEqual<int>(
                    8,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.InsertAfterTableRowFolderID),
                    152,
                    @"[In NotificationData Structure] It [InsertAfterTableRowFolderID] is 8 bytes.");
            }

            // Refer to MS-OXCNOTIF section 2.2.1.4.1.1: InsertAfterTableRowID is available only if bit 0x8000(M bit) is set in the NotificationFlags field
            // and if the TableEventType field is available and is 0x0003(TableRowAdded)or 0x0005(TableRowModified).
            bool isInsertAfterTableRowIDAvailable = (ropNotifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.TableRowAdded ||
              ropNotifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.TableRowModified) &&
              (ropNotifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) == (ushort)FlagsBit.M;

            if (isInsertAfterTableRowIDAvailable)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R156");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R156
                this.Site.CaptureRequirementIfAreEqual<int>(
                    8,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.InsertAfterTableRowID),
                    156,
                    @"[In NotificationData Structure] It [InsertAfterTableRowID] is 8 bytes.");
            }

            if (ropNotifyResponse.NotificationData.InsertAfterTableRowInstance != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R160");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R160
                this.Site.CaptureRequirementIfAreEqual<int>(
                    4,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.InsertAfterTableRowInstance),
                    160,
                    @"[In NotificationData Structure] It [InsertAfterTableRowInstance] is 4 bytes.");
            }

            // Refer to MS-OXCNOTIF section 2.2.1.4.1.1: TableRowDataSize is available only if the TableEventType field is available and is 0x0003(TableRowAdded)or 0x0005(TableRowModified).
            bool isTableRowDataSizeAvailable = ropNotifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.TableRowAdded ||
                          ropNotifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.TableRowModified;

            if (isTableRowDataSizeAvailable)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R162");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R162
                this.Site.CaptureRequirementIfAreEqual<int>(
                    2,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.TableRowDataSize),
                    162,
                    @"[In NotificationData Structure] It [TableRowDataSize] is 2 bytes.");
            }
        }

        /// <summary>
        /// Verify RopNotify response for NewMail events
        /// </summary>
        /// <param name="ropNotifyResponse">The response of ropNotify</param>
        private void VerifyRopNotifyForNewMail(RopNotifyResponse ropNotifyResponse)
        {
            if (ropNotifyResponse.NotificationData.NotificationType == NotificationType.NewMail)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R213");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R213
                this.Site.CaptureRequirementIfAreEqual<int>(
                    4,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.MessageFlags),
                    213,
                    @"[In NotificationData Structure] It [MessageFlags] is 4 bytes.");
            
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R217");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R217
                this.Site.CaptureRequirementIfAreEqual<int>(
                    1,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.UnicodeFlag),
                    217,
                    @"[In NotificationData Structure] It [UnicodeFlag]  is 1 byte.");
            }

            if (ropNotifyResponse.NotificationData.MessageClass != null)
            {
                if (ropNotifyResponse.NotificationData.UnicodeFlag == 0x00)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R219");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R219
                    bool isVerifiedR219 = this.IsNullTerminatedASCIIStr(ropNotifyResponse.NotificationData.MessageClass);

                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR219,
                        219,
                        @"[In NotificationData Structure]  otherwise, [the value of UnicodeFlag is]FALSE (0x00) indicates the value of the MessageClass is in ASCII.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R225");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R225
                    bool isVerifiedR225 = this.IsNullTerminatedASCIIStr(ropNotifyResponse.NotificationData.MessageClass);

                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR225,
                        225,
                        @"[In NotificationData Structure] The string [MessageClass] is in ASCII if UnicodeFlag is set to FALSE (0x00).");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R223");

                    // The requirements R225 have already verify the string type of MessageClass, 
                    // so capture it directly.
                    this.Site.CaptureRequirement(
                        223,
                        @"[In NotificationData Structure] MessageClass (variable): A null-terminated string containing the message class of the new mail.");
                }
            }
        }

        /// <summary>
        /// Verify RopNotify response for other events(ObjectCreated,ObjectDeleted,ObjectModified,ObjectMoved,ObjectCopied,SearchResult)
        /// </summary>
        /// <param name="ropNotifyResponse">The response of ropNotify</param>
        private void VerifyRopNotifyForOtherEvents(RopNotifyResponse ropNotifyResponse)
        {
            if (ropNotifyResponse.NotificationData.FolderId != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R169");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R169
                this.Site.CaptureRequirementIfAreEqual<int>(
                    8,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.FolderId),
                    169,
                    @"[In NotificationData Structure] It [FolderId] is 8 bytes.");
            }

            if (ropNotifyResponse.NotificationData.MessageId != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R173");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R173
                this.Site.CaptureRequirementIfAreEqual<int>(
                    8,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.MessageId),
                    173,
                    @"[In NotificationData Structure] It [MessageId] is 8 bytes.");
            }

            bool isTotalMessageCountAvailable = (ropNotifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.T) == (ushort)FlagsBit.T;

            if (isTotalMessageCountAvailable)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R205");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R205
                this.Site.CaptureRequirementIfAreEqual<int>(
                    4,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.TotalMessageCount),
                    205,
                    @"[In NotificationData Structure] It [TotalMessageCount]  is 4 bytes.");
            }

            bool isUnreadMessageCountAvailable = (ropNotifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.U) == (ushort)FlagsBit.U;

            if (isUnreadMessageCountAvailable)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R209");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R209
                this.Site.CaptureRequirementIfAreEqual<int>(
                    4,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.UnreadMessageCount),
                    209,
                    @"[In NotificationData Structure] It [UnreadMessageCount]  is 4 bytes.");
            }

            if (ropNotifyResponse.NotificationData.ParentFolderId != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R177");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R177
                this.Site.CaptureRequirementIfAreEqual<int>(
                    8,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.ParentFolderId),
                    177,
                    @"[In NotificationData Structure] It [ParentFolderId]  is 8 bytes.");
            }

            // Refer to MS-OXCNOTIF section 2.2.1.4.1.1: OldFolderId is available only if the NotificationType value in the NotificationFlags field is 0x0020(ObjectMoved) or 0x0040(ObjectCopied).
            bool isOldFolderIdAvailable = ropNotifyResponse.NotificationData.NotificationType == NotificationType.ObjectMoved ||
                        ropNotifyResponse.NotificationData.NotificationType == NotificationType.ObjectCopied;

            if (isOldFolderIdAvailable)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R181");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R181
                this.Site.CaptureRequirementIfAreEqual<int>(
                    8,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.OldFolderId),
                    181,
                    @"[In NotificationData Structure] It [OldFolderId] is 8 bytes.");
            }

            // Refer to MS-OXCNOTIF section 2.2.1.4.1.1: OldMessageId is available only if the value of the NotificationType field in the NotificationFlags field is
            // 0x0020(ObjectMoved) or 0x0040(ObjectCopied) and bit 0x8000(M bit) is set in the NotificationFlags field.
            bool isOldMessageIdAvailable = (ropNotifyResponse.NotificationData.NotificationType == NotificationType.ObjectMoved ||
                ropNotifyResponse.NotificationData.NotificationType == NotificationType.ObjectCopied) &&
                (ropNotifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) == (ushort)FlagsBit.M;
            if (isOldMessageIdAvailable)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R185");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R185
                this.Site.CaptureRequirementIfAreEqual<int>(
                    8,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.OldMessageId),
                    185,
                    @"[In NotificationData Structure] It [OldMessageId] is 8 bytes.");
            }

            // Refer to MS-OXCNOTIF section 2.2.1.4.1.1: OldParentFolderId is available only if the value of the NotificationType field in the NotificationFlags field is 0x0020(ObjectMoved) or 0x0040(ObjectCopied)
            // and bit 0x8000(M bit) is not set in the NotificationFlags field.
            bool isOldParentFolderIdAvailable = (ropNotifyResponse.NotificationData.NotificationType == NotificationType.ObjectMoved ||
                ropNotifyResponse.NotificationData.NotificationType == NotificationType.ObjectCopied) &&
                (ropNotifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) != (ushort)FlagsBit.M;

            if (isOldParentFolderIdAvailable)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R189");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R189
                this.Site.CaptureRequirementIfAreEqual<int>(
                    8,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.OldParentFolderId),
                    189,
                    @"[In NotificationData Structure] It [OldParentFolderId] is 8 bytes.");
            }

            // Refer to MS-OXCNOTIF section 2.2.1.4.1.1: TagCount is available only if the value of the NotificationType field in the NotificationFlags field is 0x0004(ObjectCreated) or 0x0010(ObjectModified).
            bool isTagCountAvailable = ropNotifyResponse.NotificationData.NotificationType == NotificationType.ObjectCreated ||
                    ropNotifyResponse.NotificationData.NotificationType == NotificationType.ObjectModified;
            if (isTagCountAvailable)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R193");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R193
                this.Site.CaptureRequirementIfAreEqual<int>(
                    2,
                    Marshal.SizeOf(ropNotifyResponse.NotificationData.TagCount),
                    193,
                    @"[In NotificationData Structure] It [TagCount]  is 2 bytes.");
            }
        }

        /// <summary>
        /// Verify NotificationFlags of RopNotify response.
        /// </summary>
        /// <param name="ropNotifyResponse">The response of ropNotify</param>
        private void VerifyNotificationFlagsOfRopNotify(RopNotifyResponse ropNotifyResponse)
        {
            // FlagsBit.T is 0x1000. T bit is set in NotificationFlags.
            bool isTBitSet = (ropNotifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.T) == (ushort)FlagsBit.T;
            if (isTBitSet)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R112");

                // Since the RopNotify response is de-serialized as this requirement's description, so if the T bit can get successfully, 
                // this requirement can be verified directly.
                Site.CaptureRequirement(
                    112,
                    @"[In NotificationData Structure] 0x1000: specify flag T.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R114");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R114
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x0010,
                    (int)ropNotifyResponse.NotificationData.NotificationType,
                    114,
                    @"[In NotificationData Structure] If this bit [0x1000] is set, the NotificationType MUST be 0x0010.");
            }

            // FlagsBit.U is 0x2000. U bit is set.
            bool isUBitSet = (ropNotifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.U) == (ushort)FlagsBit.U;
            if (isUBitSet)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R115");

                // Since the RopNotify response is de-serialized as this requirement's description, so if the U bit can get successfully, 
                // this requirement can be verified directly.
                Site.CaptureRequirement(
                    115,
                    @"[In NotificationData Structure] 0x2000: specify flag U.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R117");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R117
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x0010,
                    (int)ropNotifyResponse.NotificationData.NotificationType,
                    117,
                    @"[In NotificationData Structure] If this bit [0x2000] is set, the NotificationType MUST be 0x0010.");
            }

            // FlagsBit.S is 0x4000. S bit is set.
            bool isSBitSet = (ropNotifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.S) == (ushort)FlagsBit.S;
            if (isSBitSet)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R118");

                // Since the RopNotify response is de-serialized as this requirement's description, so if the S bit can get successfully, 
                // this requirement can be verified directly.
                Site.CaptureRequirement(
                    118,
                    @"[In NotificationData Structure] 0x4000: specify flag S.");

                int actualBitSet = ropNotifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M;

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R120");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R120
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x8000,
                    actualBitSet,
                    120,
                    @"[In NotificationData Structure] If this bit [0x4000] is set, bit 0x8000 MUST be set.");
            }

            // FlagsBit.M is 0x8000. M bit is set.
            bool isMBitSet = (ropNotifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) == (ushort)FlagsBit.M;
            if (isMBitSet)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R121");

                // Since the RopNotify response is de-serialized as this requirement's description, so if the M bit can get successfully, 
                // this requirement can be verified directly.
                Site.CaptureRequirement(
                    121,
                    @"[In NotificationData Structure] 0x8000: specify flag M.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R138");

            int notificationFlagsLength = Marshal.SizeOf(ropNotifyResponse.NotificationData.NotificationFlags) * 8;
            int notificationTypeLength = notificationFlagsLength - 4;

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R138
            // NotificationFlags (2 bytes) is a combination of NotificationType and flags (T bit,U bit,S bit,M bit), so the size of the NotificationType is the size of NotificationFlags subtract the four bit (T bit,U bit,S bit,M bit). 
            this.Site.CaptureRequirementIfAreEqual<int>(
                12,
                notificationTypeLength,
                138,
                @"[In NotificationData Structure] It [NotificationType] is 12 bits.");
        }

        /// <summary>
        /// Verify register Notification response handle.
        /// </summary>
        /// <param name="registerNotificationResponseHandle">The response handle of RopRegisterNotification</param>
        private void VerifyRopRegisterNotificationResponseHandle(uint registerNotificationResponseHandle)
        {
            // If the returned Handle is not null, this requirement can be verified.
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R35");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R35 
            this.Site.CaptureRequirementIfIsNotNull(
                registerNotificationResponseHandle,
                35,
                @"[In RopRegisterNotification ROP] The RopRegisterNotification ROP ([MS-OXCROPS] section 2.2.14.1) returns a handle of the subscription to the client.");
        }

        /// <summary>
        /// Verify Pending response.
        /// </summary>
        private void VerifyRopPendingResponse()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R83");

            // If RopPending is returned in the response buffer through the test suite, 
            // this method will be invoked, and this requirement will be verified directly.
            Site.CaptureRequirement(
                83,
                @"[In EcDoRpcExt] This ROP [RopPending] MUST appear in response buffers of the EcDoRpcExt2 method, as specified in [MS-OXCRPC] section 3.1.4.2.");            
        }

        /// <summary>
        /// Verify Syntax about UDP Transport.
        /// </summary>
        private void VerifyUDPTransport()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R3");

            // When the client receives a datagram from the server at the callback address through test suites by using underlying
            // networking protocols, this method will be invoked, and this requirement will be verified directly.
            Site.CaptureRequirement(
                3,
                @"[In Transport] UDP datagrams are sent from server to client by using the User Datagram Protocol (UDP), as specified in [RFC768].");
        }

        /// <summary>
        /// Verify callback address is used to receive UDP datagrams.
        /// </summary>
        private void VerifyCallbackAddressForUDPDatagrams()
        {
            // Callback address has been registered in EcRRegisterPushNotification method, so after a successful UDP transport, 
            // this requirement will be verified directly.
            Site.CaptureRequirement(
                78,
                @"[In EcRRegisterPushNotification Method] The callback address is required in order to receive UDP datagrams from the server.");
        }
    }
}