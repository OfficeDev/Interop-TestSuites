namespace Microsoft.Protocols.TestSuites.MS_ASEMAIL
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Text;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.Response;
    using Microsoft.Protocols.TestTools;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-ASEMAIL.
    /// </summary>
    public partial class MS_ASEMAILAdapter
    {
        /// <summary>
        /// A boolean indicate whether the Location tag exists in code page 17.
        /// </summary>
        private bool isLocationExistInCodePage17 = false;

        /// <summary>
        /// A boolean indicate whether the Location tag exists in code page 2.
        /// </summary>
        private bool isLocationExistInCodePage2 = false;

        /// <summary>
        /// A boolean indicate whether the UID tag exists in code page 4.
        /// </summary>
        private bool isUIDExistInCodePage4 = false;

        /// <summary>
        /// A boolean indicate whether the GlobalObjId tag exists in code page 2.
        /// </summary>
        private bool isGlobalObjIdExistInCodePage2 = false;

        #region Verify message syntax
        /// <summary>
        /// This method is used to verify the Message Syntax related requirements.
        /// </summary>
        private void VerifyMessageSyntax()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R109");

            // If the server returns response successfully, the MS-ASEMAIL_R109 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R109
            Site.CaptureRequirement(
                109,
                @"[In Message Syntax] The markup that is used by this protocol [MS-ASEMAIL] MUST be well-formed XML, as specified in [XML].");
        }
        #endregion

        #region Verify abstract data model
        /// <summary>
        /// This method is used to verify the Abstract Data Model related requirements.
        /// </summary>
        private void VerifyAbstractDataModel()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R51");

            // If server returns response successfully, the MS-ASEMAIL_R51 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R51
            Site.CaptureRequirement(
                51,
                @"[In Abstract Data Model] E-mail class data is returned by the server to the client as part of the full XML response to the client requests that are specified in section 3.1.5.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R53");

            // If the schema validation is successful, then MS-ASEMAIL_R53 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R53
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                53,
                @"[In Abstract Data Model] Command response: A WBXML-formatted message that adheres to the command schemas specified in [MS-ASCMD].");
        }
        #endregion

        #region Verify transport
        /// <summary>
        /// This method is used to verify transport related requirement.
        /// </summary>
        private void VerifyTransport()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R108");

            // ActiveSyncClient encodes XML request into WBXML and decodes WBXML to XML response, capture it directly if server responses succeed.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R108
            Site.CaptureRequirement(
                108,
                @"[In Transport] The XML markup that constitutes the request body or the response body that is transmitted between the client and server uses Wireless Application Protocol (WAP) Binary XML (WBXML) as specified in [MS-ASWBXML].");
        }
        #endregion

        #region Verify Sync command response
        /// <summary>
        /// This method is used to verify the Sync Command related requirements.
        /// </summary>
        /// <param name="syncStore">Server response store of a Sync command request to synchronize its E-mail class items.</param>
        private void VerifySyncCommand(DataStructures.SyncStore syncStore)
        {
            if (syncStore.AddElements.Count != 0 || syncStore.ChangeElements.Count != 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R55");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R55
                Site.CaptureRequirementIfAreEqual<byte>(
                    1,
                    syncStore.CollectionStatus,
                    55,
                    @"[In Synchronizing E-Mail Data Between Client and Server] The server responds with a Sync command response ([MS-ASCMD] section 2.2.1.21), as specified in section 3.2.5.4.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R70");

                // If the server responds with a Sync command response and the status is 1, then MS-ASEMAIL_R70 can be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R70
                Site.CaptureRequirementIfAreEqual<byte>(
                    1,
                    syncStore.CollectionStatus,
                    70,
                    @"[In Sync Command Response] When a client uses the Sync command request ([MS-ASCMD] section 2.2.1.21), as specified in section 3.1.5.4, to synchronize its E-mail class items for a specified user with the e-mail items that are currently stored by the server, the server responds with a Sync command response.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R71");

                // If the schema validation is successful, then MS-ASEMAIL_R71 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R71
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    71,
                    @"[In Sync Command Response] Any of the elements that belong to the E-mail class, as specified in section 2.2.2, can be included in a Sync command response.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R72");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R72
                Site.CaptureRequirement(
                    72,
                    @"[In Sync Command Response] E-mail class elements MUST be returned as child elements of the airsync:ApplicationData element ([MS-ASCMD] section 2.2.3.11) within either an airsync:Add element ([MS-ASCMD] section 2.2.3.7.2) or an airsync:Change element ([MS-ASCMD] section 2.2.3.24) in the Sync command response.");
            }

            this.VerifyMessageSyntax();
            this.VerifyAbstractDataModel();

            // Verify E-Mail Class elements in Sync command Add response
            if (syncStore.AddElements.Count != 0)
            {
                foreach (DataStructures.Sync item in syncStore.AddElements)
                {
                    if (item.Email != null)
                    {
                        if (item.Email.Body != null)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R246");

                            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R246
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                246,
                                @"[In Body (Airsyncbase Namespace)] When[airsyncbase:Body] included in a Sync command response ([MS-ASCMD] section 2.2.1.21), the airsyncbase:Body element can contain the following child elements: [airsyncbase:Type, airsyncbase:EstimatedDataSize, airsyncbase:Truncated and airsyncbase:Data]");
                        }

                        this.VerifyEmailClassElements(item.Email);
                    }
                }
            }

            // Verify E-Mail Class elements in Sync command Change response
            if (syncStore.ChangeElements.Count != 0)
            {
                foreach (DataStructures.Sync item in syncStore.ChangeElements)
                {
                    if (item.Email != null)
                    {
                        this.VerifyEmailClassElements(item.Email);
                    }
                }
            }
        }
        #endregion

        #region Verify ItemOperations response
        /// <summary>
        /// This method is used to verify the ItemOperations command related requirements.
        /// </summary>
        /// <param name="itemOperations">Server response store for the ItemOperations command request to retrieve data from the server for one or more specific e-mail items. </param>
        private void VerifyItemOperations(DataStructures.ItemOperationsStore itemOperations)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R59");

            // If the response is not null, then requirement MS-ASEMAIL_R59 can be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R59
            Site.CaptureRequirementIfIsNotNull(
                itemOperations,
                59,
                @"[In Retrieving Data for One or More E-Mail Items] The server responds with an ItemOperations command response ([MS-ASCMD] section 2.2.1.10), as specified in section 3.2.5.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R62");

            // If the schema validation is successful, then MS-ASEMAIL_R62 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R62
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                62,
                @"[In ItemOperations Command Response] Any of the elements that belong to the E-mail class, as specified in section 2.2.2, can be included in an ItemOperations command response.");

            if (itemOperations.Items != null)
            {
                foreach (DataStructures.ItemOperations itemOperationsItem in itemOperations.Items)
                {
                    if (itemOperationsItem.Email.Body != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R248");

                        // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R248
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            248,
                            @"[In Body (Airsyncbase Namespace)] When[airsyncbase:Body] included in an ItemOperations command response ([MS-ASCMD] section 2.2.1.10), the airsyncbase:Body element can contain the following child elements: [airsyncbase:Type, airsyncbase:EstimatedDataSize, airsyncbase:Truncated and airsyncbase:Data]");
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R64");

                    // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R64
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        64,
                        @"[In ItemOperations Command Response] E-mail class elements MUST be returned as child elements of the itemoperations:Properties element ([MS-ASCMD] section 2.2.3.139.2) in the ItemOperations command response.");

                    this.VerifyEmailClassElements(itemOperationsItem.Email);
                }
            }

            this.VerifyMessageSyntax();
            this.VerifyAbstractDataModel();
        }
        #endregion

        #region Verify Search command response
        /// <summary>
        /// This method is used to verify the Search Command related requirements.
        /// </summary>
        /// <param name="store">A SearchStore object for the Search command request to retrieve e-mail class items from the server that match the criteria specified by the client. </param>
        private void VerifySearchCommand(DataStructures.SearchStore store)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R57");

            // If the response of Search command is not null, then requirements MS-ASEMAIL_R57 and MS-ASEMAIL_R66  can be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R57
            Site.CaptureRequirementIfIsNotNull(
                store,
                57,
                @"[In Searching for E-Mail Data] The server responds with a Search command response ([MS-ASCMD] section 2.2.1.16), as specified in section 3.2.5.3.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R66");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R66
            Site.CaptureRequirementIfIsNotNull(
                store,
                66,
                @"[In Search Command Response] When a client uses the Search command request ([MS-ASCMD] section 2.2.1.16), as specified in section 3.1.5.3, to retrieve E-mail class items from the server that match the criteria specified by the client, the server responds with a Search command response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R67");

            // If the schema validation is successful, then MS-ASEMAIL_R67 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R67
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                67,
                @"[In Search Command Response] Any of the elements that belong to the E-mail class, as specified in section 2.2.2, can be included in a Search command response as child elements of the search:Properties element ([MS-ASCMD] section 2.2.3.139.3).");

            foreach (DataStructures.Search item in store.Results)
            {
                if (item.Email != null)
                {
                    if (item.Email.Body != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R247");

                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            247,
                            @"[In Body (Airsyncbase Namespace)] When[airsyncbase:Body] included in a Search command response ([MS-ASCMD] section 2.2.1.16), the airsyncbase:Body element can contain the following child elements: [airsyncbase:Type, airsyncbase:EstimatedDataSize, airsyncbase:Truncated and airsyncbase:Data]");
                    }

                    this.VerifyEmailClassElements(item.Email);
                }
            }

            this.VerifyMessageSyntax();
            this.VerifyAbstractDataModel();
        }
        #endregion

        #region Verify Find command response
        /// <summary>
        /// This method is used to verify the Find Command related requirements.
        /// </summary>
        private void VerifyFindCommand(Microsoft.Protocols.TestSuites.Common.FindResponse findResponse)
        {

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R6000");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R6000
            Site.CaptureRequirementIfIsNotNull(
                findResponse,
                6000,
                @"[In Find Command Response]When a client uses the Find command request ([MS-ASCMD] section 2.2.1.2), as specified in section 3.1.5.1, to retrieve E-mail class items from the server that match the criteria specified by the client, the server responds with a Find command response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R6001");

            // If the schema validation is successful, then MS-ASEMAIL_R6001 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R6001
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                6001,
                @"[In Find Command Response]Any of the elements that belong to the E-mail class, as specified in section 2.2.2, can be included in a Find command response as child elements of the find:Properties element ([MS-ASCMD] section 2.2.3.139.1).");

            this.VerifyMessageSyntax();
            this.VerifyAbstractDataModel();
        }
        #endregion

        #region Verify E-Mail Class elements
        /// <summary>
        /// Verify E-Mail Class elements.
        /// </summary>
        /// <param name="email">The email message synchronized from server.</param>
        private void VerifyEmailClassElements(DataStructures.Email email)
        {
            this.VerifyBody(email.Body);

            this.VerifyMeetingRequest(email.MeetingRequest);

            this.VerifyTo(email.To);

            this.VerifySender(email.Sender);

            this.VerifyThreadTopic(email.ThreadTopic);

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                this.VerifyReceivedAsBcc(email.ReceivedAsBcc);

                this.VerifyLastVerbExecutionTime(email.LastVerbExecutionTime);

                this.VerifyLastVerbExecuted(email.LastVerbExecuted);

                this.VerifyConversationId(email.ConversationId);

                this.VerifyConversationIndex(email.ConversationIndex);

                this.VerifyCategories(email.Categories);

                this.VerifyUmCallerID(email.UmCallerID, email.MessageClass);

                this.VerifyUmUserNotes(email.UmUserNotes, email.MessageClass);

                if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
                {
                    this.VerifyBodyPart(email.BodyPart);
                }
            }

            this.VerifyFrom(email.From);

            this.VerifyEmailSubject(email.Subject);

            this.VerifyCc(email.CC);

            this.VerifyReplyTo(email.ReplyTo);

            this.VerifyDisplayTo(email.DisplayTo);

            this.VerifyImportance(email.Importance);

            this.VerifyRead(email.Read);

            this.VerifyBcc(email.Bcc);

            this.VerifyIsDraft(email.IsDraft);

            this.VerifyAttachments(email.Attachments);

            this.VerifyMessageClass(email.MessageClass);

            this.VerifyInternetCPID();

            this.VerifyFlag(email.Flag);

            this.VerifyContentClass(email.ContentClass);

            this.VerifyDateReceived(email.DateReceived);

            this.VerifyNativeBodyType(email.NativeBodyType);
        }

        /// <summary>
        /// This method is used to verify the NativeBodyType related requirements.
        /// </summary>
        /// <param name="nativeBodyType">Specifies how the e-mail message is stored on the server.</param>
        private void VerifyNativeBodyType(byte? nativeBodyType)
        {
            if (nativeBodyType != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R623");

                // If the schema validation is successful, then MS-ASEMAIL_R623 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R623
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    623,
                    @"[In NativeBodyType] The value of this element[NativeBodyType] is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.8.");

                this.VerifyUnsignedByteStructure(nativeBodyType);
            }
        }

        /// <summary>
        /// This method is used to verify the Organizer related requirements.
        /// </summary>
        /// <param name="from">Specifies the coordinator of the meeting</param>
        private void VerifyOrganizer(string organizer)
        {
            if (organizer != null)
            {
                this.VerifyStringStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R637");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R637
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    637,
                    @"[In Organizer] The value of this element[Organizer] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");
            }
        }

        /// <summary>
        /// This method is used to verify the AllDayEvent related requirements.
        /// </summary>
        /// <param name="allDayEvent">Specifies whether the meeting request is for an all-day event. </param>
        private void VerifyAllDayEvent(byte allDayEvent)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R234");

            // If the schema validation is successful and the element AllDayEvent appear, then MS-ASEMAIL_R234 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R234
            Site.CaptureRequirementIfIsTrue(
               this.activeSyncClient.ValidationResult,
                234,
                @"[In AllDayEvent] The value of this element[AllDayEvent] is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.8.");

            this.VerifyUnsignedByteStructure(allDayEvent);
        }

        /// <summary>
        /// This method is used to verify the Attachment related requirements.
        /// </summary>
        /// <param name="attachment">Specifies an e-mail attachment. </param>
        private void VerifyAttachment(Response.AttachmentsAttachment attachment)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1122");

            // If the schema validation is successful, then MS-ASEMAIL_R1122 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1122
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                1122,
                @"[In Attachments(Airsyncbase Namespace)] The airsyncbase:Attachments element is a container data type, as specified in [MS-ASDTYPE] section 2.2.");

            this.VerifyContainerStructure();

            if (!string.IsNullOrEmpty(attachment.DisplayName))
            {
                this.VerifyDisplayName();
            }

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                this.VerifyUmAttDuration();
                this.VerifyUmAttOrder();
            }
        }

        /// <summary>
        /// This method is used to verify the Attachments related requirements.
        /// </summary>
        /// <param name="attachments">The attachments get from server response.</param>
        private void VerifyAttachments(Response.Attachments attachments)
        {
            if (attachments != null)
            {
                this.VerifyContainerStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1122");

                // If the schema validation is successful, then MS-ASEMAIL_R1122 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1122
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    1122,
                    @"[In Attachments(Airsyncbase Namespace)] The airsyncbase:Attachments element is a container data type, as specified in [MS-ASDTYPE] section 2.2.");

                foreach (object attachment in attachments.Items)
                {
                    if (attachment is AttachmentsAttachment)
                    {
                        this.VerifyAttachment((AttachmentsAttachment)attachment);
                    }
                }
            }
        }

        /// <summary>
        /// This method is used to verify the Body related requirements.
        /// </summary>
        /// <param name="body">The body of the email item.</param>
        private void VerifyBody(Body body)
        {
            if (body != null)
            {
                this.VerifyContainerStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R250");

                // If the schema validation is successful, then MS-ASEMAIL_R250 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R250
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    250,
                    @"[In Body (Airsyncbase Namespace)] [When[airsyncbase:Body] included in a Sync command response ([MS-ASCMD] section 2.2.1.21), a Search command response ([MS-ASCMD] section 2.2.1.16), or an ItemOperations command response ([MS-ASCMD] section 2.2.1.10), the airsyncbase:Body element can contain the following child element:] airsyncbase:Type ([MS-ASAIRS] section 2.2.2.41.1): This element [airsyncbase:Type] is required.");
            }
        }

        /// <summary>
        /// This method is used to verify the BodyPart related requirements.
        /// </summary>
        /// <param name="bodyPart">The BodyPart of the email item.</param>
        private void VerifyBodyPart(BodyPart bodyPart)
        {
            if (bodyPart != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R259");

                // If the schema validation is successful, then MS-ASEMAIL_R259 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R259
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    259,
                    @"[In BodyPart] The airsyncbase:BodyPart element is an optional container ([MS-ASDTYPE] section 2.2) element that specifies details about the message part of an e-mail message.");

                this.VerifyContainerStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the BusyStatus related requirements.
        /// </summary>
        /// <param name="busyStatus">Specifies the busy status of the recipient for the meeting. </param>
        private void VerifyBusyStatus(string busyStatus)
        {
            if (busyStatus != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R264");

                // If the schema validation is successful, then MS-ASEMAIL_R264 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R264
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    264,
                    @"[In BusyStatus] The value of this element[BusyStatus] is an integer data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyIntegerStructure();

                string[] expecedValues = new string[] { "0", "1", "2", "3", "4" };
                Common.VerifyActualValues("BusyStatus", expecedValues, busyStatus, this.Site);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R265");

                // If the verification of actual values success, then requirement MS-ASEMAIL_R265 can be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R265
                Site.CaptureRequirement(
                    265,
                    @"[In BusyStatus] The value of this element[BusyStatus] MUST be one of the values[0, 1, 2, 3, 4] listed in the following table.");
            }
        }

        /// <summary>
        /// This method is used to verify the CalendarType related requirements.
        /// </summary>
        /// <param name="calendarType">Specifies the type of calendar associated with the recurrence. </param>
        private void VerifyCalendarType(string calendarType)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R294");

            // If the schema validation is successful, then MS-ASEMAIL_R294 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R294
            Site.CaptureRequirementIfIsTrue(
                 this.activeSyncClient.ValidationResult,
                 294,
                 @"[In CalendarType] The value of this element[email2:CalendarType<5>] is an integer data type, as specified in [MS-ASDTYPE] section 2.6.");

            this.VerifyIntegerStructure();

            string[] expecedValues = new string[] { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "14", "15", "20" };
            Common.VerifyActualValues("CalendarType", expecedValues, calendarType, this.Site);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R295");

            // If the verification of actual values success, then requirement MS-ASEMAIL_R295 can be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R295
            Site.CaptureRequirement(
                295,
                @"[In CalendarType] The following table lists valid values [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 15, 20] for the email2:CalendarType element.");
        }

        /// <summary>
        /// This method is used to verify the Categories related requirements.
        /// </summary>
        /// <param name="categories">Specifies a collection of user-selected categories assigned to the e-mail message. </param>
        private void VerifyCategories(Response.Categories2 categories)
        {
            if (categories != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R313");

                // If the schema validation is successful, then MS-ASEMAIL_R313 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R313
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    313,
                    @"[In Categories] The Categories element is an optional container ([MS-ASDTYPE] section 2.2) element that specifies a collection of user-selected categories assigned to the e-mail message.");

                this.VerifyContainerStructure();

                if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R321");

                    // If the schema validation is successful, then MS-ASEMAIL_R321 could be captured.
                    // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R321
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        321,
                        @"[In Category] The Category element is an optional child element of the Categories element (section 2.2.2.16) that specifies a category that is assigned to the e-mail item.");

                    this.VerifyCategory(categories.Category);
                }
            }
        }

        /// <summary>
        /// This method is used to verify the Category related requirements.
        /// </summary>
        /// <param name="categories">The categories of the item.</param>
        private void VerifyCategory(string[] categories)
        {
            if (categories != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R323");

                // If the schema validation is successful, then MS-ASEMAIL_R323 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R323
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    323,
                    @"[In Category] The value of this element[Category] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                this.VerifyStringStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the Bcc related requirements.
        /// </summary>
        /// <param name="bcc">Specifies the blind carbon copy (Bcc) recipients of an email.</param>
        private void VerifyBcc(string bcc)
        {
            if (!string.IsNullOrEmpty(bcc))
            {
                this.Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    1131,
                    @"[In Bcc] This [email2:Bcc element is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                this.VerifyStringStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the IsDraft related requirements.
        /// </summary>
        /// <param name="bcc">The value of IsDraft element.</param>
        private void VerifyIsDraft(bool? isDraft)
        {
            if (isDraft != null)
            {
                this.Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    1277,
                    @"[In IsDraft] This element [email2:IsDraft] is a boolean data type, as specified in [MS-ASDTYPE] section 2.1. ");
            }
        }

        /// <summary>
        /// This method is used to verify the Cc related requirements.
        /// </summary>
        /// <param name="cc">Specifies the list of secondary recipients (1) of a message. </param>
        private void VerifyCc(string cc)
        {
            if (cc != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R329");

                // If the schema validation is successful, then MS-ASEMAIL_R329 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R329
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    329,
                    @"[In Cc] The value of this element[Cc] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                this.VerifyStringStructure();

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R330
                string[] emailAddresses = cc.Split(',');
                bool isValidEmailAddress = false;
                foreach (string emailAddress in emailAddresses)
                {
                    if (RFC822AddressParser.IsValidAddress(emailAddress))
                    {
                        isValidEmailAddress = true;
                    }
                    else
                    {
                        isValidEmailAddress = false;
                        break;
                    }
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R330");

                Site.CaptureRequirementIfIsTrue(
                    isValidEmailAddress,
                    330,
                    @"[In Cc] The value of this element[Cc] contains one or more e-mail addresses.");
            }
        }

        /// <summary>
        /// This method is used to verify the CompleteTime related requirements.
        /// </summary>
        private void VerifyCompleteTime()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R337");

            // If the schema validation is successful, then MS-ASEMAIL_R337 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R337
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                337,
                @"[In CompleteTime] The value of this element[CompleteTime] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

            this.VerifyDateTimeStructure();
        }

        /// <summary>
        /// This method is used to verify the ContentClass related requirements.
        /// </summary>
        /// <param name="contentClass">The content class of the item.</param>
        private void VerifyContentClass(string contentClass)
        {
            if (contentClass != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R341");

                // If the schema validation is successful, then MS-ASEMAIL_R341 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R341
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    341,
                    @"[In ContentClass] The value of this element[ContentClass] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                this.VerifyStringStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R342");

                // If the schema validation is successful, then MS-ASEMAIL_R342 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R342
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    342,
                    @"[In ContentClass] For e-mail messages, the value of this element[ContentClass] MUST be set to ""urn:content-classes:message"".");
            }
        }

        /// <summary>
        /// Verify the ConversationIndex element relative requirements.
        /// </summary>
        /// <param name="conversationIndex">The conversation index of the item.</param>
        private void VerifyConversationIndex(string conversationIndex)
        {
            if (conversationIndex != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R354");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R354
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    354,
                    @"[In ConversationIndex] The value of this element[email2:ConversationIndex] is a byte array data type, as specified in [MS-ASDTYPE] section 2.7.1.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1003");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1003
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    1003,
                    @"[In ConversationIndex] The email2:ConversationId content is transferred as an opaque binary large object (BLOB) within the WBXML tags.");
            }
        }

        /// <summary>
        /// This method is used to verify the ConversationId related requirements.
        /// </summary>
        /// <param name="conversationId">The unique identifier for a conversation.</param>
        private void VerifyConversationId(string conversationId)
        {
            if (conversationId != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R343");

                // If the schema validation is successful, then MS-ASEMAIL_R343 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R343
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    343,
                    @"[In ConversationId] The email2:ConversationId element is a required element in server responses that specifies a unique identifier for a conversation.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R345");

                // If the schema validation is successful, then MS-ASEMAIL_R345 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R345
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    345,
                    @"[In ConversationId] The value of this element[email2:ConversationId] is a byte array data type, as specified in [MS-ASDTYPE] section 2.7.1.");

                this.VerifyStringStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R348");

                // The content of ConversationId can be transferred successfully, then MS-ASEMAIL_R348 can be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R348
                Site.CaptureRequirement(
                    348,
                    @"[In ConversationId] The email2:ConversationId content is transferred as an opaque binary large object (BLOB) within the WBXML tags.");
            }
        }

        /// <summary>
        /// This method is used to verify the DateCompleted related requirements.
        /// </summary>
        /// <param name="dateCompleted">Specifies the date on which a flagged item was completed. </param>
        private void VerifyDateCompleted(DateTime dateCompleted)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R358");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R358
            // If dateCompleted is not null means server set the DateCompleted element value then MS-ASEMAIL_R358 is verified.
            Site.CaptureRequirementIfIsNotNull(
                dateCompleted,
                358,
                @"[In DateCompleted] The tasks:DateCompleted element is required to mark a flagged item as complete.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R360");

            // If the schema validation is successful, then MS-ASEMAIL_R360 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R360
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                360,
                @"[In DateCompleted] The value of this element[tasks:DateCompleted] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

            this.VerifyDateTimeStructure();
        }

        /// <summary>
        /// This method is used to verify the DateReceived related requirements.
        /// </summary>
        /// <param name="dateReceived">The date and time the message was received by the current recipient.</param>
        private void VerifyDateReceived(DateTime? dateReceived)
        {
            if (dateReceived != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R364");

                // If the schema validation is successful, then MS-ASEMAIL_R364 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R364
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    364,
                    @"[In DateReceived] The value of this element[DateReceived] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

                this.VerifyDateTimeStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the DayOfMonth related requirements.
        /// </summary>
        private void VerifyDayOfMonth()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R367");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R367
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                367,
                @"[In DayOfMonth] The value of this element[DayOfMonth] is an integer data type, as specified in [MS-ASDTYPE] section 2.6.");

            this.VerifyIntegerStructure();
        }

        /// <summary>
        /// This method is used to verify the DayOfWeek related requirements.
        /// </summary>
        /// <param name="dayOfWeek">Specifies the day of the week on which this meeting recurs. </param>
        private void VerifyDayOfWeek(string dayOfWeek)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R374");

            // If the schema validation is successful, then MS-ASEMAIL_R374 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R374
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                374,
                @"[In DayOfWeek] The value of this element[DayOfWeek] is an integer data type, as specified in [MS-ASDTYPE] section 2.6.");

            this.VerifyIntegerStructure();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R377");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R377
            Site.CaptureRequirementIfIsTrue(
                int.Parse(dayOfWeek) >= 1 && int.Parse(dayOfWeek) <= 127,
                377,
                @"[In DayOfWeek] The value of this element[DayOfWeek] MUST be the sum of a minimum of one and a maximum of seven independent values[1, 2, 4, 8, 16, 32, 64] from the following table.");
        }

        /// <summary>
        /// This method is used to verify the DisallowNewTimeProposal related requirements.
        /// </summary>
        /// <param name="disallowNewTimeProposal">Specifies whether recipients (1) can propose a new meeting time for the meeting. </param>
        private void VerifyDisallowNewTimeProposal(byte disallowNewTimeProposal)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R391");

            // If the schema validation is successful, then MS-ASEMAIL_R391 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R391
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                391,
                @"[In DisallowNewTimeProposal] The value of this element[DisallowNewTimeProposal] is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.8.");

            this.VerifyUnsignedByteStructure(disallowNewTimeProposal);
        }

        /// <summary>
        /// This method is used to verify the DisplayName related requirements.
        /// </summary>
        private void VerifyDisplayName()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R394");

            // If the schema validation is successful, then MS-ASEMAIL_R394 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R394
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                394,
                @"[In DisplayName] The value of this element[DisplayName] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

            this.VerifyStringStructure();
        }

        /// <summary>
        /// This method is used to verify the DisplayTo related requirements.
        /// </summary>
        /// <param name="displayTo">Specifies the e-mail addresses of the primary recipients (1) of this message. </param>
        private void VerifyDisplayTo(string displayTo)
        {
            if (displayTo != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R398");

                // If the schema validation is successful, then MS-ASEMAIL_R398 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R398
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    398,
                    @"[In DisplayTo] The value of this element[DisplayTo] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                this.VerifyStringStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R399");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R399
                Site.CaptureRequirementIfIsNotNull(
                    displayTo,
                    399,
                    @"[In DisplayTo] The value of this element[DisplayTo] contains one or more display names.");
            }
        }

        /// <summary>
        /// This method is used to verify the DtStamp related requirements.
        /// </summary>
        private void VerifyDtStamp()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R403");

            // If the schema validation is successful, then MS-ASEMAIL_R403 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R403
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                403,
                @"[In DtStamp] The value of this element[DtStamp] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

            this.VerifyDateTimeStructure();
        }

        /// <summary>
        /// This method is used to verify the DueDate related requirements.
        /// </summary>
        private void VerifyDueDate()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R410");

            // If the schema validation is successful, then MS-ASEMAIL_R410 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R410
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                410,
                @"[In DueDate] The value of this element[tasks:DueDate] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

            this.VerifyDateTimeStructure();
        }

        /// <summary>
        /// This method is used to verify the EndTime related requirements.
        /// </summary>
        /// <param name="endTime">The date and time when the meeting ends.</param>
        private void VerifyEndTime(DateTime endTime)
        {
            if (endTime != new DateTime())
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R416");

                // If the schema validation is successful, then MS-ASEMAIL_R416 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R416
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    416,
                    @"[In EndTime] The value of this element[EndTime] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

                this.VerifyDateTimeStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the FirstDayOfWeek related requirements.
        /// </summary>
        /// <param name="firstDayOfWeek">Specifies which day is considered the first day of the calendar week for the recurrence. </param>
        private void VerifyFirstDayOfWeek(byte firstDayOfWeek)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R423");

            // If the schema validation is successful, then MS-ASEMAIL_R423 and MS-ASEMAIL_R420 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R423
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                423,
                @"[In FirstDayOfWeek] The value of this element[email2:FirstDayOfWeek] is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.8.");

            this.VerifyUnsignedByteStructure(firstDayOfWeek);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R420");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R420
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                420,
                @"[In FirstDayOfWeek] A command response has a maximum of one email2:FirstDayOfWeek child element per Recurrence element.");

            string[] expecedValues = new string[] { "0", "1", "2", "3", "4", "5", "6" };
            Common.VerifyActualValues("FirstDayOfWeek", expecedValues, firstDayOfWeek.ToString(), this.Site);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R424");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R424
            Site.CaptureRequirement(
                424,
                @"[In FirstDayOfWeek] The value of the email2:FirstDayOfWeek element MUST be one of the values[0, 1, 2, 3, 4, 5, 6] listed in the following table.");
        }

        /// <summary>
        /// This method is used to verify the Flag related requirements.
        /// </summary>
        /// <param name="flag">Specifies the flag associated with the item and indicates the item's current status. </param>
        private void VerifyFlag(Response.Flag flag)
        {
            if (flag == null || ((XmlElement)this.LastRawResponseXml).OuterXml.Contains(@"<Flag xmlns=""Email"" />"))
            {
                return;
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R434");

            // If the schema validation is successful, then MS-ASEMAIL_R434 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R434
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                434,
                @"[In Flag] The Flag element is an optional container ([MS-ASDTYPE] section 2.2) element that defines the flag associated with the item [and indicates the item's current status].");

            this.VerifyContainerStructure();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1069");

            // If the flag is not null, then MS-ASEMAIL_R1069 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1069
            Site.CaptureRequirement(
                1069,
                @"[In Flag] The Flag element is an optional container ([MS-ASDTYPE] section 2.2) element that [defines the flag associated with the item and] indicates the item's current status.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R436");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R436
            Site.CaptureRequirementIfIsTrue(
                flag.Subject != null || flag.Status != null || flag.FlagType != null || flag.DateCompletedSpecified,
                436,
                @"[In Flag] If flags are present on the e-mail item, the Flag element contains one or more child elements that define the flag.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R440");

            // If the schema validation is successful, then MS-ASEMAIL_R440 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R440
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                440,
                @"[In Flag] The Flag element can contain the following child elements:[tasks:Subject, Status, FlagType, tasks:DateCompleted, CompleteTime, tasks:StartDate, tasks:DueDate, tasks:UtcStartDate, tasks:UtcDueDate, tasks:ReminderSet, tasks:ReminderTime, tasks:OrdinalDate, tasks:SubOrdinalDate]");

            this.VerifyTaskSubject(flag.Subject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R758");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R758
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                758,
                @"[In Status] A maximum of one Status element is allowed per Flag.");

            this.VerifyStatus(flag.Status);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R472");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R472
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                472,
                @"[In FlagType] A maximum of one FlagType child element is allowed per Flag.");

            this.VerifyFlagType(flag.FlagType);

            if (flag.DateCompletedSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R359");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R359
                // If flag.CompleteTime is not null means server set the CompleteTime element value then MS-ASEMAIL_R359 is verified.
                Site.CaptureRequirementIfIsNotNull(
                    flag.CompleteTime,
                    359,
                    @"[In DateCompleted] If a message includes a value for the tasks:DateCompleted element, the CompleteTime element (section 2.2.2.19) is also required.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R361");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R361
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    361,
                    @"[In DateCompleted] A maximum of one tasks:DateCompleted child element is allowed per Flag element.");

                this.VerifyDateCompleted(flag.DateCompleted);
            }

            if (flag.CompleteTimeSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R338");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R338
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    338,
                    @"[In CompleteTime] A maximum of one CompleteTime child element is allowed per Flag element.");

                this.VerifyCompleteTime();
            }

            if (flag.StartDateSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R743");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R743
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    743,
                    @"[In StartDate] A maximum of one tasks:StartDate child element is allowed per Flag element.");

                this.VerifyStartDate();
            }

            if (flag.DueDateSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R411");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R411
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    411,
                    @"[In DueDate] A maximum of one tasks:DueDate child element is allowed per Flag element.");

                this.VerifyDueDate();
            }

            if (flag.UtcStartDateSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R866");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R866
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    866,
                    @"[In UtcStartDate] A maximum of one tasks:UtcStartDate child element is allowed per Flag element.");

                this.VerifyUtcStartDate();
            }

            if (flag.UtcDueDateSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R855");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R855
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    855,
                    @"[In UtcDueDate] A maximum of one tasks:UtcDueDate child element is allowed per Flag element.");

                this.VerifyUtcDueDate();
            }

            if (flag.ReminderSetSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R694");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R694
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    694,
                    @"[In ReminderSet] A maximum of one tasks:ReminderSet child element is allowed per Flag element.");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R701
                if (flag.ReminderSet.Equals(1))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R701");

                    Site.CaptureRequirementIfIsTrue(
                        flag.ReminderTimeSpecified,
                        701,
                        @"[In ReminderTime] The tasks:ReminderTime element MUST be set if the tasks:ReminderSet element value is set to 1 (TRUE).");
                }

                this.VerifyReminderSet(flag.ReminderSet);
            }

            if (flag.ReminderTimeSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R703");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R703
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    703,
                    @"[In ReminderTime] A maximum of one tasks:ReminderTime child element is allowed per Flag element.");

                this.VerifyReminderTime();
            }

            if (flag.OrdinalDateSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R631");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R631
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    631,
                    @"[In OrdinalDate] A maximum of one tasks:OrdinalDate child element is allowed per Flag element.");

                this.VerifyOrdinalDate();
            }

            this.VerifySubOrdinalDate(flag.SubOrdinalDate);
        }

        /// <summary>
        /// This method is used to verify the FlagType related requirements.
        /// </summary>
        /// <param name="flagType">The type of the flag.</param>
        private void VerifyFlagType(string flagType)
        {
            if (flagType != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R469");

                // If the schema validation is successful, then MS-ASEMAIL_R469 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R469
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    469,
                    @"[In FlagType] The value of this element[FlagType] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                this.VerifyStringStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the From related requirements.
        /// </summary>
        /// <param name="from">Specifies the e-mail address of the message sender. </param>
        private void VerifyFrom(string from)
        {
            if (from != null)
            {
                // As specified in MS-ASDTYPE, An e-mail address is an unconstrained value of an element of the string type.
                this.VerifyStringStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R476");

                // If the schema validation is successful, then MS-ASEMAIL_R476 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R476
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    476,
                    @"[In From] The value of the From element is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                Site.Log.Add(LogEntryKind.Debug, "The value of From element is {0}.", from);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R477");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R477
                Site.CaptureRequirementIfIsTrue(
                    from.Length <= 32768,
                    477,
                    @"[In From] and [The value of the From element] has a maximum length of 32,768 characters.");
            }
        }

        /// <summary>
        /// This method is used to verify the Importance related requirements.
        /// </summary>
        /// <param name="importance">Specifies the importance of the message, as assigned by the sender. </param>
        private void VerifyImportance(byte? importance)
        {
            if (importance != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R482");

                // If the schema validation is successful, then MS-ASEMAIL_R482 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R482
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    482,
                    @"[In Importance] The value of this element[Importance] is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.8.");

                this.VerifyUnsignedByteStructure(importance);

                string[] expecedValues = new string[] { "0", "1", "2" };
                Common.VerifyActualValues("Importance", expecedValues, importance.ToString(), this.Site);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R483");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R483
                Site.CaptureRequirement(
                    483,
                    @"[In Importance] The value of this element[Importance] MUST be one of the values[0, 1, 2] listed in the following table.");
            }
        }

        /// <summary>
        /// This method is used to verify the InstanceType related requirements.
        /// </summary>
        /// <param name="instanceType">Specifies whether the calendar item is a single or recurring appointment. </param>
        private void VerifyInstanceType(byte instanceType)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R490");

            // If the schema validation is successful, then MS-ASEMAIL_R490 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R490
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                490,
                @"[In InstanceType] The value of this element[InstanceType] is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.8.");

            this.VerifyUnsignedByteStructure(instanceType);

            string[] expecedValues = new string[] { "0", "1", "2", "3", "4" };
            Common.VerifyActualValues("InstanceType", expecedValues, instanceType.ToString(), this.Site);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R491");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R491
            Site.CaptureRequirement(
                491,
                @"[In InstanceType] The value of this element[InstanceType] MUST be one of the values[0, 1, 2, 3, 4] listed in the following table.");
        }

        /// <summary>
        /// This method is used to verify the InternetCPID related requirements.
        /// </summary>
        private void VerifyInternetCPID()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R496");

            // If the schema validation is successful, then MS-ASEMAIL_R496 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R496
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                496,
                @"[In InternetCPID] The InternetCPID element is a required element that contains the original code page ID from the MIME message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R498");

            // If the schema validation is successful, then MS-ASEMAIL_R498 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R498
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                498,
                @"[In InternetCPID] The value of this element[InternetCPID] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

            this.VerifyStringStructure();
        }

        /// <summary>
        /// This method is used to verify the Interval related requirements.
        /// </summary>
        private void VerifyInterval()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R501");

            // If the schema validation is successful, then MS-ASEMAIL_R501 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R501
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                501,
                @"[In Interval] The value of this element[Interval] is an integer data type, as specified in [MS-ASDTYPE] section 2.6.");

            this.VerifyIntegerStructure();
        }

        /// <summary>
        /// This method is used to verify the IsLeapMonth related requirements.
        /// </summary>
        /// <param name="isLeapMonth">Specifies whether the recurrence takes place in the leap month of the given year. </param>
        private void VerifyIsLeapMonth(byte isLeapMonth)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R507");

            // If the schema validation is successful, then MS-ASEMAIL_R507 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R507
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                507,
                @"[In IsLeapMonth] The value of this element[email2:IsLeapMonth] is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.8.");

            this.VerifyUnsignedByteStructure(isLeapMonth);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R508");

            // If the schema validation is successful, then MS-ASEMAIL_R508 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R508
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                508,
                @"[In IsLeapMonth] This element[email2:IsLeapMonth] is required in server responses [and is optional in client requests].");
        }

        /// <summary>
        /// This method is used to verify the LastVerbExecuted related requirements.
        /// </summary>
        /// <param name="lastVerbExecuted">Specifies the last action, such as reply or forward, that was taken on the message. </param>
        private void VerifyLastVerbExecuted(int? lastVerbExecuted)
        {
            if (lastVerbExecuted != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R513");

                // If the schema validation is successful, then MS-ASEMAIL_R513 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R513
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    513,
                    @"[In LastVerbExecuted] The value of this element[email2:LastVerbExecuted] is an integer data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyIntegerStructure();

                string[] expecedValues = new string[] { "0", "1", "2", "3" };
                Common.VerifyActualValues("LastVerbExecuted", expecedValues, lastVerbExecuted.ToString(), this.Site);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R514");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R514
                Site.CaptureRequirement(
                    514,
                    @"[In LastVerbExecuted] The following table lists the valid values[0, 1, 2, 3] for this element[email2:LastVerbExecuted].");
            }
        }

        /// <summary>
        /// This method is used to verify the LastVerbExecutionTime related requirements.
        /// </summary>
        /// <param name="lastVerbExecutionTime">The date and time when the action was performed on the message.</param>
        private void VerifyLastVerbExecutionTime(DateTime? lastVerbExecutionTime)
        {
            if (lastVerbExecutionTime != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R522");

                // If the schema validation is successful, then MS-ASEMAIL_R522 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R522
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    522,
                    @"[In LastVerbExecutionTime] The value of this element[email2:LastVerbExecutionTime] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

                this.VerifyDateTimeStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the Location related requirements.
        /// </summary>
        /// <param name="location">Specifies where the meeting will occur. </param>
        private void VerifyLocation(string location)
        {
            if (location != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R526");

                // If the schema validation is successful, then MS-ASEMAIL_R526 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R526
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    526,
                    @"[In Location] The value of the Location element is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R527");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R527
                Site.CaptureRequirementIfIsTrue(
                    location.Length <= 32768,
                    527,
                    @"[In Location] and [The value of the Location element is a string data type, as specified in [MS-ASDTYPE] section 2.6,] has a maximum length of 32,768 characters.");

                this.VerifyStringStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the MeetingMessageType related requirements.
        /// </summary>
        /// <param name="meetingMessageType">Specifies the type of meeting message. </param>
        private void VerifyMeetingMessageType(byte meetingMessageType)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R530");

            // If the schema validation is successful, then MS-ASEMAIL_R530 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R530
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                530,
                @"[In MeetingMessageType] The value of this element[email2:MeetingMessageType] is an unsignedByte value, as specified in [MS-ASDTYPE] section 2.8.");

            this.VerifyUnsignedByteStructure(meetingMessageType);

            string[] expecedValues = new string[] { "0", "1", "2", "3", "4", "5", "6" };
            Common.VerifyActualValues("MeetingMessageType", expecedValues, meetingMessageType.ToString(), this.Site);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R532");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R532
            Site.CaptureRequirement(
                532,
                @"[In MeetingMessageType] The value of this element[email2:MeetingMessageType] MUST be one of the values[0, 1, 2, 3, 4, 5, 6] listed in the following table.");
        }

        /// <summary>
        /// This method is used to verify the MeetingRequest related requirements.
        /// </summary>
        /// <param name="meetingRequest">Specifies information about the meeting request. </param>
        private void VerifyMeetingRequest(Response.MeetingRequest meetingRequest)
        {
            if (meetingRequest != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R546");

                // If the schema validation is successful, then MS-ASEMAIL_R546 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R546
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    546,
                    @"[In MeetingRequest] The MeetingRequest element can contain the following child elements in a command response:[AllDayEvent, StartTime, DtStamp, EndTime, InstanceType, Location, Organizer, RecurrenceId, Reminder, ResponseRequested, Recurrences, Sensitivity, BusyStatus, TimeZone, GlobalObjId, DisallowNewTimeProposal, MeetingMessageType]");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R541");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R541
                Site.CaptureRequirement(
                    541,
                    @"[In MeetingRequest] The MeetingRequest element is an optional container ([MS-ASDTYPE] section 2.2) element that contains information about the meeting.");

                this.VerifyRecurrences(meetingRequest.Recurrences);

                this.VerifyOrganizer(meetingRequest.Organizer);

                if (meetingRequest.AllDayEventSpecified)
                {
                    this.VerifyAllDayEvent(meetingRequest.AllDayEvent);
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R896");

                // If the schema validation is successful, then MS-ASEMAIL_R896 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R896
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    896,
                    @"[In MeetingRequest] StartTime (section 2.2.2.73): This element is optional.");

                this.VerifyStartTime(meetingRequest.StartTime);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R897");

                // If the schema validation is successful, then MS-ASEMAIL_R897 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R897
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    897,
                    @"[In MeetingRequest] DtStamp (section 2.2.2.30): One instance of this element is required.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R401");

                // If the schema validation is successful, then MS-ASEMAIL_R401 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R401
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    401,
                    @"[In DtStamp] The DtStamp element is a required child element of the MeetingRequest element (section 2.2.2.48) that specifies the date and time the calendar item was created.");

                this.VerifyDtStamp();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R898");

                // If the schema validation is successful, then MS-ASEMAIL_R898 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R898
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    898,
                    @"[In MeetingRequest] EndTime (section 2.2.2.32): This element is optional.");

                this.VerifyEndTime(meetingRequest.EndTime);

                this.VerifyInstanceType(meetingRequest.InstanceType);

                if (meetingRequest.RecurrenceIdSpecified)
                {
                    this.VerifyRecurrenceId();
                }

                if (meetingRequest.ResponseRequestedSpecified)
                {
                    this.VerifyResponseRequested(meetingRequest.ResponseRequested);
                }

                this.VerifySensitivity(meetingRequest.Sensitivity);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R682");

                // If the schema validation is successful, then MS-ASEMAIL_R682 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R682
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    682,
                    @"[In Recurrences] It[Recurrences] is a child element of the MeetingRequest element (section 2.2.2.48).");

                this.VerifyRecurrences(meetingRequest.Recurrences);

                this.VerifyBusyStatus(meetingRequest.BusyStatus);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R908");

                // If the schema validation is successful, then MS-ASEMAIL_R908 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R908
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    908,
                    @"[In MeetingRequest] TimeZone (section 2.2.2.78): One instance of this element is required.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R781");

                // If the schema validation is successful, then MS-ASEMAIL_R781 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R781
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    781,
                    @"[In TimeZone] The TimeZone element is a required child element of the MeetingRequest element (section 2.2.2.48) that specifies the time zone specified when the calendar item was created.");

                this.VerifyTimeZone();

                if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1")
                    || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0")
                    || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.1"))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R909");

                    // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R909
                    Site.CaptureRequirementIfIsTrue(
                        !string.IsNullOrEmpty(meetingRequest.GlobalObjId),
                        909,
                        @"[In MeetingRequest] GlobalObjId (section 2.2.2.37): One instance of this element is required.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R478");

                    // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R478
                    Site.CaptureRequirementIfIsTrue(
                       !string.IsNullOrEmpty(meetingRequest.GlobalObjId),
                        478,
                        @"[In GlobalObjId] The GlobalObjId element is a required child element of the MeetingRequest element (section 2.2.2.48) that contains a hexadecimal ID generated by the server for the meeting request.");

                    this.VerifyLocation(meetingRequest.Location);
                }
                else
                {
                    if (meetingRequest.Location1 != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1311");

                        // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1311
                        Site.CaptureRequirementIfIsTrue(
                            string.IsNullOrEmpty(meetingRequest.Location),
                            1311,
                            @"[In MeetingRequest] In protocol version 16.0: The [calendar:UID element is used instead of the email:GlobalObjId element; the] airsyncbase:Location element is used instead of the email:Location element.");
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R13111");

                        // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R13111
                        Site.CaptureRequirementIfIsTrue(
                            string.IsNullOrEmpty(meetingRequest.Location),
                            13111,
                            @"[In MeetingRequest] In protocol version 16.1: The [calendar:UID element is used instead of the email:GlobalObjId element; the] airsyncbase:Location element is used instead of the email:Location element.");
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1305");

                    // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1305
                    Site.CaptureRequirementIfIsTrue(
                        !string.IsNullOrEmpty(meetingRequest.UID),
                        1305,
                        @"[In MeetingRequest] calendar:UID ([MS-ASCAL] section 2.2.2.46): One instance of this element is required.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R13110");

                    // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R13110
                    Site.CaptureRequirementIfIsTrue(
                        !string.IsNullOrEmpty(meetingRequest.UID),
                        1310,
                        @"[In MeetingRequest] In protocol version 16.0: The calendar:UID element is used instead of the email:GlobalObjId element[; the airsyncbase:Location element is used instead of the email:Location element].");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1310");

                    // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1310
                    Site.CaptureRequirementIfIsTrue(
                        !string.IsNullOrEmpty(meetingRequest.UID),
                        13110,
                        @"[In MeetingRequest] In protocol version 16.1: The calendar:UID element is used instead of the email:GlobalObjId element[; the airsyncbase:Location element is used instead of the email:Location element].");
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1253");

                    // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1253
                    Site.CaptureRequirementIfIsTrue(
                        !string.IsNullOrEmpty(meetingRequest.UID),
                        1253,
                        @"[In GlobalObjId] The server will return the calendar:UID element ([MS-ASCAL] section 2.2.2.46) instead of the GlobalObjId element when protocol version 16.0 is used.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R12540");

                    // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R12540
                    Site.CaptureRequirementIfIsTrue(
                        !string.IsNullOrEmpty(meetingRequest.UID),
                        12540,
                        @"[In GlobalObjId] The server will return the calendar:UID element ([MS-ASCAL] section 2.2.2.46) instead of the GlobalObjId element when protocol version 16.1 is used.");
                }

                if (meetingRequest.DisallowNewTimeProposalSpecified && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
                {
                    this.VerifyDisallowNewTimeProposal(meetingRequest.DisallowNewTimeProposal);
                }

                if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1") || !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R911");

                    // If the schema validation is successful, then MS-ASEMAIL_R911 could be captured.
                    // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R911
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        911,
                        @"[In MeetingRequest] MeetingMessageType (section 2.2.2.47): This element is required.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R528");

                    // If the schema validation is successful, then MS-ASEMAIL_R528 could be captured.
                    // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R528
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        528,
                        @"[In MeetingMessageType] The email2:MeetingMessageType element is a required child element of the MeetingRequest element (section 2.2.2.48) that specifies the type of meeting message.");

                    this.VerifyMeetingMessageType(meetingRequest.MeetingMessageType);
                }
            }
        }

        /// <summary>
        /// Verify the Recurrences element.
        /// </summary>
        /// <param name="recurrences">The Recurrences element returned from server.</param>
        private void VerifyRecurrences(Response.MeetingRequestRecurrences recurrences)
        {
            if (recurrences != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R681");

                // If the schema validation is successful, then MS-ASEMAIL_R681 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R681
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    681,
                    @"[In Recurrences] The Recurrences element is an optional container ([MS-ASDTYPE] section 2.2) element that contains details about the recurrence pattern of the meeting.");

                this.VerifyContainerStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R684");

                // If the schema validation is successful, then MS-ASEMAIL_R684 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R684
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    684,
                    @"[In Recurrences] The Recurrences element MUST contain the following child element: Recurrence (section 2.2.2.60): This element is required.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R652");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R652
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    652,
                    @"[In Recurrence] The Recurrence element is a required child element of the Recurrences element (section 2.2.2.62).");

                this.VerifyRecurrence(recurrences.Recurrence);
            }
        }

        /// <summary>
        /// This method is used to verify the MessageClass related requirements.
        /// </summary>
        /// <param name="messageClass">Specifies the message class of this e-mail message. </param>
        private void VerifyMessageClass(string messageClass)
        {
            if (messageClass != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R584");

                // If the schema validation is successful, then MS-ASEMAIL_R584 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R584
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    584,
                    @"[In MessageClass] The value of this element[MessageClass] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                this.VerifyStringStructure();

                if (messageClass.Contains("REPORT"))
                {
                    this.VerifyAdministrativeMessageClass(messageClass);
                }

                if (Common.IsRequirementEnabled(588, this.Site))
                {
                    Site.Log.Add(LogEntryKind.Debug, "The value of MessageClass element is {0}.", messageClass);

                    bool isR588Satisfied = messageClass.Contains("IPM.Note") || messageClass.Contains("IPM.Note.SMIME")
                        || messageClass.Contains("IPM.Note.SMIME.MultipartSigned") || messageClass.Contains("IPM.Note.Receipt.SMIME")
                        || messageClass.Contains("IPM.InfoPathForm") || messageClass.Contains("IPM.Schedule.Meeting")
                        || messageClass.Contains("IPM.Notification.Meeting") || messageClass.Contains("IPM.Post")
                        || messageClass.Contains("IPM.Octel.Voice") || messageClass.Contains("IPM.Voicenotes")
                        || messageClass.Contains("IPM.Sharing");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R588");

                    // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R588
                    Site.CaptureRequirementIfIsTrue(
                        isR588Satisfied,
                        588,
                        @"[In Appendix B: Product Behavior] The value of the MessageClass element is one of the values listed in the following table or derive from one of the values[IPM.Note, IPM.Note.SMIME, IPM.Note.SMIME.MultipartSigned, IPM.Note.Receipt.SMIME, IPM.InfoPathForm, IPM.Schedule.Meeting, IPM.Notification.Meeting, IPM.Post, IPM.Octel.Voice, IPM.Voicenotes, IPM.Sharing] listed in the following table. (Exchange Server 2007 SP1 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// This method is used to verify the Administrative MessageClass related requirements.
        /// </summary>
        /// <param name="administrativeMessageClass">Specifies the message class of this e-mail message.</param>
        private void VerifyAdministrativeMessageClass(string administrativeMessageClass)
        {
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R601
            bool isVerifyR601 = administrativeMessageClass.Contains("IPM.Note") || administrativeMessageClass.Contains("IPM.Note.SMIME")
                || administrativeMessageClass.Contains("IPM.Note.SMIME.MultipartSigned") || administrativeMessageClass.Contains("IPM.Note.Receipt.SMIME")
                || administrativeMessageClass.Contains("IPM.InfoPathForm") || administrativeMessageClass.Contains("IPM.Schedule.Meeting")
                || administrativeMessageClass.Contains("IPM.Notification.Meeting") || administrativeMessageClass.Contains("IPM.Post")
                || administrativeMessageClass.Contains("IPM.Octel.Voice") || administrativeMessageClass.Contains("IPM.Voicenotes")
                || administrativeMessageClass.Contains("IPM.Sharing");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R601");

            Site.CaptureRequirementIfIsTrue(
                isVerifyR601,
                601,
                @"[In MessageClass] In addition, certain administrative messages, such as read receipts and non-delivery reports that are generated by the server, have a message class that is derived from one of the message classes [IPM.Note, IPM.Note.SMIME, IPM.Note.SMIME.MultipartSigned, IPM.Note.Receipt.SMIME, IPM.InfoPathForm, IPM.Schedule.Meeting, IPM.Notification.Meeting, IPM.Post, IPM.Octel.Voice, IPM.Voicenotes, IPM.Sharing] listed in the preceding table.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R602");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R602
            Site.CaptureRequirementIfIsTrue(
                administrativeMessageClass.StartsWith("REPORT", StringComparison.CurrentCulture),
                602,
                @"[In MessageClass] The format of this value is a prefix of ""REPORT"" and a suffix that indicates the type of report.");

            // The values are case insensitive.
            bool isVerifyR603 = administrativeMessageClass.Equals("REPORT.IPM.NOTE.NDR", StringComparison.CurrentCultureIgnoreCase)
                || administrativeMessageClass.Equals("REPORT.IPM.NOTE.DR", StringComparison.CurrentCultureIgnoreCase)
                || administrativeMessageClass.Equals("REPORT.IPM.NOTE.DELAYED", StringComparison.CurrentCultureIgnoreCase)
                || administrativeMessageClass.EndsWith("REPORT.IPM.NOTE.IPNRN", StringComparison.CurrentCultureIgnoreCase)
                || administrativeMessageClass.EndsWith("REPORT.IPM.NOTE.IPNNRN", StringComparison.CurrentCultureIgnoreCase)
                || administrativeMessageClass.Equals("REPORT.IPM.SCHEDULE. MEETING.REQUEST.NDR", StringComparison.CurrentCultureIgnoreCase)
                || administrativeMessageClass.Equals("REPORT.IPM.SCHEDULE.MEETING.RESP.POS.NDR", StringComparison.CurrentCultureIgnoreCase)
                || administrativeMessageClass.Equals("REPORT.IPM.SCHEDULE.MEETING.RESP.TENT.NDR", StringComparison.CurrentCultureIgnoreCase)
                || administrativeMessageClass.Equals("REPORT.IPM.SCHEDULE.MEETING.CANCELED.NDR", StringComparison.CurrentCultureIgnoreCase)
                || administrativeMessageClass.Equals("REPORT.IPM.NOTE.SMIME.NDR", StringComparison.CurrentCultureIgnoreCase)
                || administrativeMessageClass.EndsWith("REPORT.IPM.NOTE.SMIME.DR", StringComparison.CurrentCultureIgnoreCase)
                || administrativeMessageClass.EndsWith("REPORT.IPM.NOTE.SMIME.MULTIPARTSIGNED.NDR", StringComparison.CurrentCultureIgnoreCase)
                || administrativeMessageClass.EndsWith("REPORT.IPM.NOTE.SMIME.MULTIPARTSIGNED.DR", StringComparison.CurrentCultureIgnoreCase);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R603");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R603
            Site.CaptureRequirementIfIsTrue(
                isVerifyR603,
                603,
                @"[In MessageClass] For these administrative messages, the value of the MessageClass element MUST be one of the following values[REPORT.IPM.NOTE.NDR,
                  REPORT.IPM.NOTE.DR, REPORT.IPM.NOTE.DELAYED, *REPORT.IPM.NOTE.IPNRN, *REPORT.IPM.NOTE.IPNNRN, REPORT.IPM.SCHEDULE. MEETING.REQUEST.NDR,
                  REPORT.IPM.SCHEDULE.MEETING.RESP.POS.NDR, REPORT.IPM.SCHEDULE.MEETING.RESP.TENT.NDR, REPORT.IPM.SCHEDULE.MEETING.CANCELED.NDR, REPORT.IPM.NOTE.SMIME.NDR,
                  *REPORT.IPM.NOTE.SMIME.DR, *REPORT.IPM.NOTE.SMIME.MULTIPARTSIGNED.NDR, *REPORT.IPM.NOTE.SMIME.MULTIPARTSIGNED.DR].");
        }

        /// <summary>
        /// This method is used to verify the MonthOfYear related requirements.
        /// </summary>
        private void VerifyMonthOfYear()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R619");

            // If the schema validation is successful, then MS-ASEMAIL_R619 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R619
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                619,
                @"[In MonthOfYear] The value of this element[MonthOfYear] is an integer data type, as specified in [MS-ASDTYPE] section 2.6.");

            this.VerifyIntegerStructure();
        }

        /// <summary>
        /// This method is used to verify the Occurrences related requirements.
        /// </summary>
        private void VerifyOccurrences()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R627");

            // If the schema validation is successful, then MS-ASEMAIL_R627 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R627
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                627,
                @"[In Occurrences] The value of this element[Occurrences] is an integer data type, as specified in [MS-ASDTYPE] section 2.6.");

            this.VerifyStringStructure();
        }

        /// <summary>
        /// This method is used to verify the OrdinalDate related requirements.
        /// </summary>
        private void VerifyOrdinalDate()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R630");

            // If the schema validation is successful, then MS-ASEMAIL_R630 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R630
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                630,
                @"[In OrdinalDate] The value of this element[tasks:OrdinalDate] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

            this.VerifyDateTimeStructure();
        }

        /// <summary>
        /// This method is used to verify the Read related requirements.
        /// </summary>
        /// <param name="read">True indicates the message has been read, False value indicates the message has not been read </param>
        private void VerifyRead(bool? read)
        {
            if (read != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R642");

                // If the schema validation is successful, then MS-ASEMAIL_R642 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R642
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    642,
                    @"[In Read] The value of this element is a boolean data type, as specified in [MS-ASDTYPE] section 2.1.");

                this.VerifyBooleanStructure(read);
            }
        }

        /// <summary>
        /// This method is used to verify the ReceivedAsBcc element.
        /// </summary>
        /// <param name="receivedAsBcc">The boolean value of ReceivedAsBcc.</param>
        private void VerifyReceivedAsBcc(bool? receivedAsBcc)
        {
            if (receivedAsBcc != null)
            {
                this.VerifyBooleanStructure(receivedAsBcc);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R646");

                // If the schema validation is successful, then MS-ASEMAIL_R646 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R646
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    646,
                    @"[In ReceivedAsBcc] The value of this element[email2:ReceivedAsBcc] is a boolean data type, as specified in [MS-ASDTYPE] section 2.1.");
            }
        }

        /// <summary>
        /// This method is used to verify the Recurrence related requirements.
        /// </summary>
        /// <param name="recurrence">Specifies when and how often the meeting recurs. </param>
        private void VerifyRecurrence(Response.MeetingRequestRecurrencesRecurrence recurrence)
        {
            this.VerifyContainerStructure();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R653");

            // If the schema validation is successful, then MS-ASEMAIL_R653 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R653
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                653,
                @"[In Recurrence] The Recurrence element can contain the following child elements: [Type, Interval, Until, Occurrences, WeekOfMonth, DayOfMonth, DayOfWeek, MonthOfYear, email2:CalendarType, email2:IsLeapMonth, email2:FirstDayOfWeek]");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R880");

            // If the schema validation is successful, then MS-ASEMAIL_R880 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R880
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                880,
                @"[In Recurrence] Type (section 2.2.2.80): One instance of this element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R789");

            // If the schema validation is successful, then MS-ASEMAIL_R789 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R789
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                789,
                @"[In Type] The Type element is a required child element of the Recurrence element (section 2.2.2.60) that specifies how the meeting recurs.");

            this.VerifyType(recurrence.Type);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R912");

            // If the schema validation is successful, then MS-ASEMAIL_R912 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R912
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                912,
                @"[In Recurrence] Interval (section 2.2.2.41): One instance of this element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R499");

            // If the schema validation is successful, then MS-ASEMAIL_R499 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R499
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                499,
                @"[In Interval] The Interval element is a required child element of the Recurrence element (section 2.2.2.60) that specifies the interval between meeting recurrences.");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R369
            if (recurrence.Type.Equals(2))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R369");

                // If the DayOfMonth element is included in Sync command response, then requirement MS-ASEAMIL_R369 can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    recurrence.DayOfMonth,
                    369,
                    @"[In DayOfMonth] This element[DayOfMonth] is required when the Type element (section 2.2.2.80) is set to a value of 2 (that is, the meeting recurs monthly on the Nth day of the month),");
            }

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R370
            if (recurrence.Type.Equals(5))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R370");

                // If the DayOfMonth element is included in Sync command response, then requirement MS-ASEAMIL_R370 can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    recurrence.DayOfMonth,
                    370,
                    @"[In DayOfMonth] [This element[DayOfMonth] is required when the Type element (section 2.2.2.80) is set to] a value of 5 (that is, the meeting recurs yearly on the Nth day of the Nth month each year).");
            }

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R375
            if (recurrence.Type.Equals(1))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R375");

                // If the DayOfWeek element is included in Sync command response, then requirement MS-ASEAMIL_R375 can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    recurrence.DayOfWeek,
                    375,
                    @"[In DayOfWeek] This element[DayOfWeek] is required when the Type element (section 2.2.2.80) is set to a value of 1 (that is, the meeting recurs weekly),");
            }

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R376
            if (recurrence.Type.Equals(6))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R376");

                // If the DayOfWeek element is included in Sync command response, then requirement MS-ASEAMIL_R376 can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    recurrence.DayOfWeek,
                    376,
                    @"[In DayOfWeek] [This element[DayOfWeek] is required when the Type element (section 2.2.2.80) is set to]a value of 6 (that is, the meeting recurs yearly on the Nth day of the week of the Nth month each year).");
            }

            this.VerifyInterval();

            if (recurrence.Until != null)
            {
                this.VerifyUntil();
            }

            if (!string.IsNullOrEmpty(recurrence.Occurrences))
            {
                this.VerifyOccurrences();
            }

            if (!string.IsNullOrEmpty(recurrence.WeekOfMonth))
            {
                this.VerifyWeekOfMonth();
            }

            if (!string.IsNullOrEmpty(recurrence.DayOfMonth))
            {
                this.VerifyDayOfMonth();
            }

            if (!string.IsNullOrEmpty(recurrence.DayOfWeek))
            {
                this.VerifyDayOfWeek(recurrence.DayOfWeek);
            }

            if (!string.IsNullOrEmpty(recurrence.MonthOfYear))
            {
                this.VerifyMonthOfYear();
            }

            if (recurrence.CalendarType != null && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                this.VerifyCalendarType(recurrence.CalendarType);
            }

            if (recurrence.IsLeapMonthSpecified && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                this.VerifyIsLeapMonth(recurrence.IsLeapMonth);
            }

            if (recurrence.FirstDayOfWeekSpecified && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1")
                && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R420");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R420
                Site.CaptureRequirementIfIsTrue(
                    recurrence.FirstDayOfWeek.ToString().Length <= 1,
                    420,
                    @"[In FirstDayOfWeek] A command response has a maximum of one email2:FirstDayOfWeek child element per Recurrence element.");

                this.VerifyFirstDayOfWeek(recurrence.FirstDayOfWeek);
            }
        }

        /// <summary>
        /// This method is used to verify the RecurrenceId related requirements.
        /// </summary>
        private void VerifyRecurrenceId()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R680");

            // If the schema validation is successful, then MS-ASEMAIL_R680 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R680
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                680,
                @"[In RecurrenceId] The value of this element[RecurrenceId] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

            this.VerifyDateTimeStructure();
        }

        /// <summary>
        /// This method is used to verify the ReminderSet related requirements.
        /// </summary>
        /// <param name="reminderSet">Specifies whether a reminder has been set for the task. </param>
        private void VerifyReminderSet(byte reminderSet)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R691");

            // If the schema validation is successful, then MS-ASEMAIL_R691 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R691
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                691,
                @"[In ReminderSet] The value of this element[tasks:ReminderSet] is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.8.");

            this.VerifyUnsignedByteStructure(reminderSet);
        }

        /// <summary>
        /// This method is used to verify the ReminderTime related requirements.
        /// </summary>
        private void VerifyReminderTime()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R700");

            // If the schema validation is successful, then MS-ASEMAIL_R700 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R700
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                700,
                @"[In ReminderTime] The value of this element[tasks:ReminderTime] is a dateTime value, as specified in [MS-ASDTYPE] section 2.3.");

            this.VerifyDateTimeStructure();
        }

        /// <summary>
        /// This method is used to verify the ReplyTo related requirements.
        /// </summary>
        /// <param name="replyTo">Specifies the e-mail address(es) to which replies will be addressed by default. </param>
        private void VerifyReplyTo(string replyTo)
        {
            if (replyTo != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R709");

                // If the schema validation is successful, then MS-ASEMAIL_R709 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R709
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    709,
                    @"[In ReplyTo] The value of this element[ReplyTo] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                this.VerifyStringStructure();

                Site.Log.Add(LogEntryKind.Debug, "The value of RelpyTo element is {0}", replyTo);

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R710
                string[] emailAddresses = replyTo.Split(';');
                bool isValidEmailAddress = false;
                foreach (string emailAddress in emailAddresses)
                {
                    if (RFC822AddressParser.IsValidAddress(emailAddress))
                    {
                        isValidEmailAddress = true;
                    }
                    else
                    {
                        isValidEmailAddress = false;
                        break;
                    }
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R710");

                Site.CaptureRequirementIfIsTrue(
                    isValidEmailAddress,
                    710,
                    @"[In ReplyTo] The value of this element[ReplyTo] contains one or more e-mail addresses.");
            }
        }

        /// <summary>
        /// This method is used to verify the ResponseRequested related requirements.
        /// </summary>
        /// <param name="responseRequested">Specifies whether the organizer has requested a response to the meeting request. </param>
        private void VerifyResponseRequested(byte responseRequested)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R714");

            // If the schema validation is successful, then MS-ASEMAIL_R714 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R714
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                714,
                @"[In ResponseRequested] The value of this element[ResponseRequested] is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.8.");

            this.VerifyUnsignedByteStructure(responseRequested);
        }

        /// <summary>
        /// This method is used to verify the Sender element.
        /// </summary>
        /// <param name="sender">Identifies the user that actually sent the message.</param>
        private void VerifySender(string sender)
        {
            if (sender != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R721");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R721
                Site.CaptureRequirementIfIsTrue(
                    RFC822AddressParser.IsValidAddress(sender),
                    721,
                    @"[In Sender] The value of the Sender element is an e-mail address, as specified in [MS-ASDTYPE] section 2.7.3.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R100");

                // Verify MS-ASEMAIL requirement: MS-ASDTYPE_R100
                Site.CaptureRequirementIfIsTrue(
                    RFC822AddressParser.IsValidAddress(sender),
                    "MS-ASDTYPE",
                    100,
                    @"[In E-Mail Address] However, a valid individual e-mail address MUST have the following format: ""local-part@domain"".");
            }
        }

        /// <summary>
        /// This method is used to verify the Sensitivity related requirements.
        /// </summary>
        /// <param name="sensitivity">Specifies the confidentiality level of the meeting request. </param>
        private void VerifySensitivity(string sensitivity)
        {
            if (sensitivity != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R729");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R729
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    729,
                    @"[In Sensitivity] The value of this element[Sensitivity] is an integer data type, as specified in [MS-ASDTYPE] section 2.6.");

                string[] expecedValues = new string[] { "0", "1", "2", "3" };
                Common.VerifyActualValues("Sensitivity", expecedValues, sensitivity.ToString(), this.Site);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R730");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R730
                Site.CaptureRequirement(
                    730,
                    @"[In Sensitivity] The value of this element[Sensitivity] MUST be one of the values[0, 1, 2, 3] listed in the following table.");
            }
        }

        /// <summary>
        /// This method is used to verify the StartDate related requirements.
        /// </summary>
        private void VerifyStartDate()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R742");

            // If the schema validation is successful, then MS-ASEMAIL_R742 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R742
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                742,
                @"[In StartDate] The value of this element[tasks:StartDate] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

            this.VerifyDateTimeStructure();
        }

        /// <summary>
        /// This method is used to verify the StartTime related requirements.
        /// </summary>
        /// <param name="startTime">Specifies when the meeting begin.</param>
        private void VerifyStartTime(DateTime startTime)
        {
            if (startTime != new DateTime())
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R749");

                // If the schema validation is successful, then MS-ASEMAIL_R749 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R749
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    749,
                    @"[In StartTime] The value of this element[StartTime] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

                this.VerifyDateTimeStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the Status related requirements.
        /// </summary>
        /// <param name="status">Specifies the current status of the flag. </param>
        private void VerifyStatus(string status)
        {
            if (status != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R752");

                // If the schema validation is successful, then MS-ASEMAIL_R752 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R752
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    752,
                    @"[In Status] The value of this element[Status] is an integer data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyIntegerStructure();

                string[] expecedValues = new string[] { "0", "1", "2", "3" };
                Common.VerifyActualValues("Status", expecedValues, status.ToString(), this.Site);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R753");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R753
                Site.CaptureRequirement(
                    753,
                    @"[In Status] The value of this element[Status] MUST be one of the values[0, 1, 2] in the following table.");
            }
        }

        /// <summary>
        /// This method is used to verify the email message subject related requirements.
        /// </summary>
        /// <param name="subject">The subject of the e-mail message.</param>
        private void VerifyEmailSubject(string subject)
        {
            if (subject != null)
            {
                this.VerifyStringStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1416");

                // If the schema validation is successful, then MS-ASEMAIL_R1416 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1416
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    1416,
                    @"[In Subject (Email Namespace)] The value of this element [Subject (Email Namespace)] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");
            }
        }

        /// <summary>
        /// This method is used to verify the task message subject related requirements.
        /// </summary>
        /// <param name="subject">The subject of the flag.</param>
        private void VerifyTaskSubject(string subject)
        {
            if (subject != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R769");

                // If the schema validation is successful, then MS-ASEMAIL_R769 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R769
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    769,
                    @"[In Subject (Tasks Namespace)] The value of this element [Tasks:Subject] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R768");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R768
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    768,
                    @"[In Subject (Tasks Namespace)] A maximum of one tasks:Subject child element is allowed per Flag element.");

                this.VerifyStringStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the SubOrdinalDate related requirements.
        /// </summary>
        /// <param name="subOrdinalDate">Specifies a value that should be used for sorting.</param>
        private void VerifySubOrdinalDate(string subOrdinalDate)
        {
            if (subOrdinalDate != null)
            {
                if (Common.IsRequirementEnabled(1006, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1006");

                    // If the schema validation is successful, then MS-ASEMAIL_R1006 could be captured.
                    // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1006
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        1006,
                        @"[In Appendix B: Product Behavior] The tasks:SubOrdinalDate element is an optional child element of the Flag element (section 2.2.2.27) that specifies a value that is used for sorting.(Exchange Server 2007 SP1 and above follow this behavior.)");
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R774");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R774
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    774,
                    @"[In SubOrdinalDate] A maximum of one tasks:SubOrdinalDate child element is allowed per Flag element.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R772");

                // If the schema validation is successful, then MS-ASEMAIL_R772 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R772
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    772,
                    @"[In SubOrdinalDate] The value of this element[tasks:SubOrdinalDate] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                this.VerifyStringStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the ThreadTopic related requirements.
        /// </summary>
        /// <param name="threadTopic">The topic used for conversation threading.</param>
        private void VerifyThreadTopic(string threadTopic)
        {
            if (threadTopic != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R780");

                // If the schema validation is successful, then MS-ASEMAIL_R780 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R780
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    780,
                    @"[In ThreadTopic] The value of this element[ThreadTopic] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                this.VerifyStringStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the TimeZone related requirements.
        /// </summary>
        private void VerifyTimeZone()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R783");

            // If the schema validation is successful, then MS-ASEMAIL_R783 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R783
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                783,
                @"[In TimeZone] The value of this element[TimeZone] is a string data type ([MS-ASDTYPE] section 2.7) in TimeZone format, as specified in [MS-ASDTYPE] section 2.7.6.");

            this.VerifyTimeZoneStructure();
        }

        /// <summary>
        /// This method is used to verify the To related requirements.
        /// </summary>
        /// <param name="to">Specifies the list of primary recipients (1) of a message. </param>
        private void VerifyTo(string to)
        {
            if (to != null)
            {
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R786
                string[] emailAddresses = to.Split(',');
                bool isValidEmailAddress = false;
                foreach (string emailAddress in emailAddresses)
                {
                    if (RFC822AddressParser.IsValidAddress(emailAddress))
                    {
                        isValidEmailAddress = true;
                    }
                    else
                    {
                        isValidEmailAddress = false;
                        break;
                    }
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R786");

                Site.CaptureRequirementIfIsTrue(
                    isValidEmailAddress,
                    786,
                    @"[In To] The value of this element[To] contains one or more e-mail addresses.");

                Site.Log.Add(LogEntryKind.Debug, "The value of To element is {0}.", to);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R788");

                // If the schema validation is successful and the length of To element less than 32,768 characters, then MS-ASEMAIL_R788 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R788
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult && to.Length <= 32768,
                    788,
                    @"[In To] The value of this element[To] is a string data type, as specified in [MS-ASDTYPE] section 2.7, and has a maximum length of 32,768 characters.");

                this.VerifyStringStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the Type related requirements.
        /// </summary>
        /// <param name="type">Specifies how the meeting recurs. </param>
        private void VerifyType(byte type)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R791");

            // If the schema validation is successful, then MS-ASEMAIL_R791 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R791
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                791,
                @"[In Type] The value of this element[Type] is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.8.");

            this.VerifyUnsignedByteStructure(type);

            string[] expecedValues = new string[] { "0", "1", "2", "3", "5", "6" };
            Common.VerifyActualValues("Type", expecedValues, type.ToString(), this.Site);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R792");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R792
            Site.CaptureRequirement(
                792,
                @"[In Type] The value of this element[Type] MUST be one of the values[0, 1, 2, 3, 5, 6] in the following table.");
        }

        /// <summary>
        /// Verify the element UmAttDuration relative requirements
        /// </summary>
        private void VerifyUmAttDuration()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R801");

            // If the schema validation is successful, then MS-ASEMAIL_R801 and MS-ASEMAIL_R802 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R801
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                801,
                @"[In UmAttDuration] The value of this element[email2:UmAttDuration] is an integer data type, as specified in [MS-ASDTYPE] section 2.6.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R802");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R802
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                802,
                @"[In UmAttDuration] This element[email2:UmAttDuration] MUST only be used for electronic voice message attachments.");

            this.VerifyIntegerStructure();
        }

        /// <summary>
        ///  Verify the element UmAttOrder relative requirements
        /// </summary>
        private void VerifyUmAttOrder()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R807");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R807
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                807,
                @"[In UmAttOrder] The value of this element[email2:UmAttOrder ] is an integer data type, as specified in [MS-ASDTYPE] section 2.6.");

            this.VerifyIntegerStructure();
        }

        /// <summary>
        /// This method is used to verify the UmCallerID related requirements.
        /// </summary>
        /// <param name="callerID">Specifies the callback telephone number of the person who called or left an electronic voice message. </param>
        /// <param name="messageClass">Specifies the message class of this e-mail message.</param>
        private void VerifyUmCallerID(string callerID, string messageClass)
        {
            if (callerID != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R827");

                // If the schema validation is successful, then MS-ASEMAIL_R827 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R827
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    827,
                    @"[In UmCallerID] Only one email2:UmCallerID element is allowed per message.");

                string[] expecedValues = new string[] { "IPM.Note.Microsoft.Voicemail", "IPM.Note.Microsoft.Voicemail.UM", "IPM.Note.Microsoft.Voicemail.UM.CA", "IPM.Note.RPMSG.Microsoft.Voicemail", "IPM.Note.RPMSG.Microsoft.Voicemail.UM", "IPM.Note.RPMSG.Microsoft.Voicemail.UM.CA", "IPM.Note.Microsoft.Missed.Voice" };
                Common.VerifyActualValues("MessageClass", expecedValues, messageClass, this.Site);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R878");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R878
                // Firstly, if routine can reach here, it means the MessageClass element value meet the R878 description.
                // So if the callerID element is not null means server set the UmCallerID element value then MS-ASEMAIL_R878 is verified.
                Site.CaptureRequirementIfIsNotNull(
                    callerID,
                    878,
                    @"[In UmCallerID] This element MUST only be included for messages with one of the following MessageClass values:IPM.Note.Microsoft.Voicemail, IPM.Note.Microsoft.Voicemail.UM, IPM.Note.Microsoft.Voicemail.UM.CA, IPM.Note.RPMSG.Microsoft.Voicemail, IPM.Note.RPMSG.Microsoft.Voicemail.UM, IPM.Note.RPMSG.Microsoft.Voicemail.UM.CA, IPM.Note.Microsoft.Missed.Voice.");

                this.VerifyStringStructure();
            }
        }

        /// <summary>
        /// This method is used to verify the UmUserNotes related requirements.
        /// </summary>
        /// <param name="userNotes">Contains user notes related to an electronic voice message. </param>
        /// <param name="messageClass">Specifies the message class of this e-mail message.</param>
        private void VerifyUmUserNotes(string userNotes, string messageClass)
        {
            if (userNotes != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R832");

                // If the schema validation is successful, then MS-ASEMAIL_R832 could be captured.
                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R832
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    832,
                    @"[In UmUserNotes] The value of this element[email2:UmUserNotes] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                this.VerifyStringStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R833");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R833
                Site.CaptureRequirementIfIsTrue(
                    Encoding.Default.GetBytes(userNotes).Length <= 16374,
                    833,
                    @"[In UmUserNotes] The server truncates notes larger than 16,374 bytes, to 16,374 bytes.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R844");

                // If the schema validation is successful, then MS-ASEMAIL_R844 could be captured.
                // Verify MS-ASEMAIL requirement: MMS-ASEMAIL_R844
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    844,
                    @"[In UmUserNotes] Only one email2:UmUserNotes element is allowed per message.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R937");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R937
                // If userNotes element is not null means server set the UmUserNotes element value then MS-ASEMAIL_R937 is verified.
                Site.CaptureRequirementIfIsNotNull(
                    userNotes,
                    937,
                    @"[In UmUserNotes] This element[UmUserNotes] is sent from the server to the client.");

                string[] expecedValues = new string[] { "IPM.Note.Microsoft.Voicemail", "IPM.Note.Microsoft.Voicemail.UM", "IPM.Note.Microsoft.Voicemail.UM.CA", "IPM.Note.RPMSG.Microsoft.Voicemail", "IPM.Note.RPMSG.Microsoft.Voicemail.UM", "IPM.Note.RPMSG.Microsoft.Voicemail.UM.CA", "IPM.Note.Microsoft.Missed.Voice" };
                Common.VerifyActualValues("MessageClass", expecedValues, messageClass, this.Site);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R879");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R879
                // Firstly, if routine can reach here, it means the MessageClass element value meet the R879 description.
                // So if the userNotes element is not null means server set the UmUserNotes element value then MS-ASEMAIL_R879 is verified.
                Site.CaptureRequirementIfIsNotNull(
                    userNotes,
                    879,
                    @"[In UmUserNotes] This element MUST only be included for electronic voice messages with one of the following MessageClass values: IPM.Note.Microsoft.Voicemail, IPM.Note.Microsoft.Voicemail.UM, IPM.Note.Microsoft.Voicemail.UM.CA, IPM.Note.RPMSG.Microsoft.Voicemail, IPM.Note.RPMSG.Microsoft.Voicemail.UM, IPM.Note.RPMSG.Microsoft.Voicemail.UM.CA, IPM.Note.Microsoft.Missed.Voice.");
            }
        }

        /// <summary>
        /// This method is used to verify the Until related requirements.
        /// </summary>
        private void VerifyUntil()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R847");

            // If the schema validation is successful, then MS-ASEMAIL_R847 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R847
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                847,
                @"[In Until] The value of this element[Until] is a string value, as specified in [MS-ASDTYPE] section 2.7.");

            this.VerifyStringStructure();
        }

        /// <summary>
        /// This method is used to verify the UtcDueDate related requirements.
        /// </summary>
        private void VerifyUtcDueDate()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R854");

            // If the schema validation is successful, then MS-ASEMAIL_R854 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R854
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                854,
                @"[In UtcDueDate] The value of this element[tasks:UtcDueDate] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

            this.VerifyDateTimeStructure();
        }

        /// <summary>
        /// This method is used to verify the UtcStartDate related requirements.
        /// </summary>
        private void VerifyUtcStartDate()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R865");

            // If the schema validation is successful, then MS-ASEMAIL_R865 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R865
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                865,
                @"[In UtcStartDate] The value of this element[tasks:UtcStartDate] is a dateTime data type, as specified in [MS-ASDTYPE] section 2.3.");

            this.VerifyDateTimeStructure();
        }

        /// <summary>
        /// This method is used to verify the WeekOfMonth related requirements.
        /// </summary>
        private void VerifyWeekOfMonth()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R872");

            // If the schema validation is successful, then MS-ASEMAIL_R872 could be captured.
            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R872
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                872,
                @"[In WeekOfMonth] The value of this element[WeekOfMonth] is an integer data type, as specified in [MS-ASDTYPE] section 2.6.");

            this.VerifyIntegerStructure();
        }
        #endregion

        #region Verify Data type
        /// <summary>
        /// This method is used to verify the boolean related requirements.
        /// </summary>
        /// <param name="boolValue">A bool value</param>
        private void VerifyBooleanStructure(bool? boolValue)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R4");

            // If the schema validation is successful, then MS-ASDTYPE_R4 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R4
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                4,
                @"[In boolean Data Type] It [a boolean] is declared as an element with a type attribute of ""boolean"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R5");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R5
            Site.CaptureRequirementIfIsTrue(
                boolValue.Equals(true) || boolValue.Equals(false),
                "MS-ASDTYPE",
                5,
                @"[In boolean Data Type] The value of a boolean element is an integer whose only valid values are 1 (TRUE) or 0 (FALSE).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R7");

            // ActiveSyncClient encoded boolean data as inline strings, so if response is successfully returned MS-ASDTYPE_R7 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R7
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                7,
                @"[In boolean Data Type] Elements with a boolean data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");
        }

        /// <summary>
        /// This method is used to verify the container related requirements.
        /// </summary>
        private void VerifyContainerStructure()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R9");

            // If the schema validation is successful, then MS-ASDTYPE_R9 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R9
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                9,
                @"[In container Data Type] It [container] is a complex type with complex content, as specified in [XMLSCHEMA1/2] section 3.4.2.");
        }

        /// <summary>
        /// This method is used to verify the dateTime related requirements.
        /// </summary>
        private void VerifyDateTimeStructure()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R12");

            // If the schema validation is successful, then MS-ASDTYPE_R12 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R12
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                12,
                @"[In dateTime Data Type] It [dateTime]is declared as an element whose type attribute is set to ""dateTime"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R20");

            // ActiveSyncClient encoded dateTime data as inline strings, so if response is successfully returned this requirement can be verified.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R20
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                20,
                @"[In dateTime Data Type] Elements with a dateTime data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R15");

            // If the schema validation is successful, then MS-ASDTYPE_R15 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R15
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                15,
                @"[In dateTime Data Type] All dates are given in Coordinated Universal Time (UTC) and are represented as a string in the following format.
YYYY-MM-DDTHH:MM:SS.MSSZ where
YYYY = Year (Gregorian calendar year)
MM = Month (01 - 12)
DD = Day (01 - 31)
HH = Number of complete hours since midnight (00 - 24)
MM = Number of complete minutes since start of hour (00 - 59)
SS = Number of seconds since start of minute (00 - 59)
MSS = Number of milliseconds");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R18");

            // If the schema validation is successful, then MS-ASDTYPE_R18 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R18
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                18,
                @"[In dateTime Data Type] Note: Dates and times in calendar items (as specified in [MS-ASCAL]) MUST NOT include punctuation separators.");
        }

        /// <summary>
        /// This method is used to verify the integer related requirements.
        /// </summary>
        private void VerifyIntegerStructure()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R86");

            // If the schema validation is successful, then MS-ASDTYPE_R86 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R86
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                86,
                @"[In integer Data Type] It [an integer] is an XML schema primitive data type, as specified in [XMLSCHEMA2/2] section 3.3.13.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R87");

            // ActiveSyncClient encoded integer data as inline strings, so if response is successfully returned this requirement can be covered.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R87
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                87,
                @"[In integer Data Type] Elements with an integer data type MUST be encoded and transmitted as WBXML inline strings, as specified in [WBXML1.2].");
        }

        /// <summary>
        /// This method is used to verify the string related requirements.
        /// </summary>
        private void VerifyStringStructure()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R88");

            // If the schema validation is successful, then MS-ASDTYPE_R88 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R88
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                88,
                @"[In string Data Type] A string is a chunk of Unicode text.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R90");

            // If the schema validation is successful, then MS-ASDTYPE_R90 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R90
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                90,
                @"[In string Data Type] An element of this [string] type is declared as an element with a type attribute of ""string"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R91");

            // ActiveSyncClient encoded string data as inline strings, so if response is successfully returned this requirement can be covered.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R91
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                91,
                @"[In string Data Type] Elements with a string data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R94");

            // If the schema validation is successful, then MS-ASDTYPE_R94 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R94
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                94,
                @"[In string Data Type] Elements of these types [ActiveSync defines several conventions for strings that adhere to commonly used formats]are defined as string types in XML schemas.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R97");

            // If the schema validation is successful, then MS-ASDTYPE_R97 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R97
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                97,
                @"[In Byte Array] The structure is comprised of a length, which is expressed as a multi-byte integer, as specified in [WBXML1.2], followed by that many bytes of data.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R98");

            // If the schema validation is successful, then MS-ASDTYPE_R98 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R98
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                98,
                @"[In Byte Array] Elements with a byte array structure MUST be encoded and transmitted as [WBXML1.2] opaque data.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R99");

            // If the schema validation is successful, then MS-ASDTYPE_R99 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R99
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                99,
                @"[In E-Mail Address] An e-mail address is an unconstrained value of an element of the string type (section 2.6).");
        }

        /// <summary>
        /// This method is used to verify the TimeZone related requirements.
        /// </summary>
        private void VerifyTimeZoneStructure()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R103");

            // If the schema validation is successful, then MS-ASDTYPE_R103 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R103
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                103,
                @"[In TimeZone] The TimeZone structure is a structure inside of an element of the string type (section 2.6).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R104");

            // If the schema validation is successful, then MS-ASDTYPE_R104 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R104
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                104,
                @"[In TimeZone] Bias (4 bytes): The value of this [Bias]field is a LONG, as specified in [MS-DTYP].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R107");

            // If the schema validation is successful, then MS-ASDTYPE_R107 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R107
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                107,
                @"[In TimeZone] StandardName (64 bytes): The value of this field is an array of 32 WCHARs, as specified in [MS-DTYP].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R108");

            // If the schema validation is successful, then MS-ASDTYPE_R108 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R108
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                108,
                @"[In TimeZone] It [TimeZone]contains an optional description for standard time.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R109");

            // If the schema validation is successful, then MS-ASDTYPE_R109 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R109
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                109,
                @"[In TimeZone] Any unused WCHARs in the array MUST be set to 0x0000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R110");

            // If the schema validation is successful, then MS-ASDTYPE_R110 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R110
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                110,
                @"[In TimeZone] StandardDate (16 bytes): The value of this field is a SYSTEMTIME structure, as specified in [MS-DTYP].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R111");

            // If the schema validation is successful, then MS-ASDTYPE_R111 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R111
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                111,
                @"[In TimeZone] It [TimeZone]contains the date and time when the transition from DST to standard time occurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R112");

            // If the schema validation is successful, then MS-ASDTYPE_R112 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R112
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                112,
                @"[In TimeZone] StandardBias (4 bytes): The value of this field is a LONG.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R113");

            // If the schema validation is successful, then MS-ASDTYPE_R113 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R113
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                113,
                @"[In TimeZone] It[TimeZone] contains the number of minutes to add to the value of the Bias field during standard time.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R114");

            // If the schema validation is successful, then MS-ASDTYPE_R114 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R114
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                114,
                @"[In TimeZone] DaylightName (64 bytes): The value of this field is an array of 32 WCHARs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R115");

            // If the schema validation is successful, then MS-ASDTYPE_R115 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R115
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                115,
                @"[In TimeZone] It [TimeZone] contains an optional description for DST.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R116");

            // If the schema validation is successful, then MS-ASDTYPE_R116 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R116
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                116,
                @"[In TimeZone] Any unused WCHARs in the array MUST be set to 0x0000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R117");

            // If the schema validation is successful, then MS-ASDTYPE_R117 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R117
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                117,
                @"[In TimeZone] DaylightDate (16 bytes): The value of this field is a SYSTEMTIME structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R118");

            // If the schema validation is successful, then MS-ASDTYPE_R118 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R118
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                118,
                @"[In TimeZone] It [TimeZone]contains the date and time when the transition from standard time to DST occurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R119");

            // If the schema validation is successful, then MS-ASDTYPE_R119 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R119
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                119,
                @"[In TimeZone] DaylightBias (4 bytes): The value of this field is a LONG.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R120");

            // If the schema validation is successful, then MS-ASDTYPE_R120 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R120
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                120,
                @"[In TimeZone] It [TimeZone]contains the number of minutes to add to the value of the Bias field during DST.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R121");

            // If the schema validation is successful, then MS-ASDTYPE_R121 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R121
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                121,
                @"[In TimeZone] The TimeZone structure is encoded using base64 encoding prior to being inserted in an XML element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R122");

            // If the schema validation is successful, then MS-ASDTYPE_R122 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R122
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                122,
                @"[In TimeZone] Elements with a TimeZone structure MUST be encoded and transmitted as [WBXML1.2] inline strings.");
        }

        /// <summary>
        /// This method is used to verify the unsignedByte related requirements.
        /// </summary>
        /// <param name="byteValue">A byte value.</param>
        private void VerifyUnsignedByteStructure(byte? byteValue)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R123");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R123
            Site.CaptureRequirementIfIsTrue(
                (byteValue >= 0) && (byteValue <= 255),
                "MS-ASDTYPE",
                123,
                @"[In unsignedByte Data Type] The unsignedByte data type is an integer value between 0 and 255, inclusive.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R125");

            // If the schema validation is successful, then MS-ASDTYPE_R125 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R125
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                125,
                @"[In unsignedByte Data Type] Elements of this type [unsignedByte type] are declared with an element whose type attribute is set to ""unsignedByte"".");
        }
        #endregion

        #region Verify requirements of [MS-ASWBXML] for code page 2, 9, 17 and 22
        /// <summary>
        /// Verify WBXML Capture for WBXML process.
        /// </summary>
        private void VerifyWBXMLCapture()
        {
            // Get decode data and capture requirement for decode processing
            Dictionary<string, int> decodedData = this.activeSyncClient.GetMSASWBXMLImplementationInstance().DecodeDataCollection;

            if (decodedData != null)
            {
                // check out all tag-token
                foreach (KeyValuePair<string, int> decodeDataItem in decodedData)
                {
                    byte token;
                    string tagName = Common.GetTagName(decodeDataItem.Key, out token);
                    int codepage = decodeDataItem.Value;
                    string codePageName = Common.GetCodePageName(decodeDataItem.Key);
                    bool isValidCodePage = codepage >= 0 && codepage <= 25;
                    Site.Assert.IsTrue(isValidCodePage, "Code page value should between 0-24, the actual value is :{0}", codepage);

                    // Begin to capture requirement of Email namespace
                    if (2 == codepage)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R12");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R12
                        Site.CaptureRequirementIfAreEqual<string>(
                            "email",
                            codePageName.ToLower(CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            12,
                            @"[In Code Pages] [This algorithm supports] [Code page] 2 [that indicates] [XML namespace] Email.");

                        this.VerifyRequirementsRelateToCodePage2(codepage, tagName, token);
                    }

                    // Begin to capture requirement of Task namespace
                    if (9 == codepage)
                    {
                        this.VerifyRequirementsRelateToCodePage9(tagName, token);
                    }

                    // Begin to capture requirement of AirSyncBase namespace
                    if (17 == codepage)
                    {
                        this.VerifyRequirementsRelateToCodePage17(tagName, token);
                    }

                    // Begin to capture requirement of Email2 namespace
                    if (22 == codepage)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R32");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R32
                        Site.CaptureRequirementIfAreEqual<string>(
                            "email2",
                            codePageName.ToLower(CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            32,
                            @"[In Code Pages] [This algorithm supports] [Code page] 22[that indicates] [XML namespace] Email2");

                        this.VerifyRequirementsRelateToCodePage22(codepage, tagName, token);
                    }

                    if (4 == codepage)
                    {
                        if (tagName == "UID")
                        {
                            this.isUIDExistInCodePage4 = true;
                        }
                    }

                    if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0"))
                    {
                        if (tagName.Equals("Location"))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R821");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R821
                            Site.CaptureRequirementIfIsTrue(
                                this.isLocationExistInCodePage17 == true && this.isLocationExistInCodePage2 == false,
                                "MS-ASWBXML",
                                821,
                                @"[In Code Page 2: Email] Note 3: The Location tag in WBXML code page 17 (AirSyncBase) is used instead of the Location tag in WBXML code page 2 with protocol wersion 16.0 and 16.1.");
                        }

                        if (tagName.Equals("UID"))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R822");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R822
                            Site.CaptureRequirementIfIsTrue(
                                this.isGlobalObjIdExistInCodePage2 == false && this.isUIDExistInCodePage4 == true,
                                "MS-ASWBXML",
                                822,
                                @"[In Code Page 2: Email] Note 4: The UID tag in WBXML code page 4 (Calendar) is used instead of the GlobalObjId tag in WBXML code page 2 with protocol version 16.0 and 16.1.");
                        }
                    }
                }
            }
        }

        #region Tag and token mapping captures.
        /// <summary>
        /// Verify the tags and tokens in WBXML code page 2.
        /// </summary>
        /// <param name="codePageNumber">code page number</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified. </param>
        private void VerifyRequirementsRelateToCodePage2(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "DateReceived":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R129");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R554
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0F,
                            token,
                            "MS-ASWBXML",
                            129,
                            @"[In Code Page 2: Email] [Tag name] DateReceived [Token] 0x0F [supports protocol versions] All");

                        break;
                    }

                case "DisplayTo":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R130");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R130
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x11,
                            token,
                            "MS-ASWBXML",
                            130,
                            @"[In Code Page 2: Email] [Tag name] DisplayTo [Token] 0x11 [supports protocol versions] All");

                        break;
                    }

                case "Importance":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R131");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R131
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x12,
                            token,
                            "MS-ASWBXML",
                            131,
                            @"[In Code Page 2: Email] [Tag name] Importance [Token] 0x12 [supports protocol versions] All");

                        break;
                    }

                case "MessageClass":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R132");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R132
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x13,
                            token,
                            "MS-ASWBXML",
                            132,
                            @"[In Code Page 2: Email] [Tag name] MessageClass [Token] 0x13 [supports protocol versions] All");

                        break;
                    }

                case "Subject":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R133");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R133
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x14,
                            token,
                            "MS-ASWBXML",
                            133,
                            @"[In Code Page 2: Email] [Tag name] Subject [Token] 0x14 [supports protocol versions] All");

                        break;
                    }

                case "Read":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R134");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R134
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x15,
                            token,
                            "MS-ASWBXML",
                            134,
                            @"[In Code Page 2: Email] [Tag name] Read [Token] 0x15 [supports protocol versions] All");

                        break;
                    }

                case "To":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R135");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R135
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x16,
                            token,
                            "MS-ASWBXML",
                            135,
                            @"[In Code Page 2: Email] [Tag name] To [Token] 0x16 [supports protocol versions] All");

                        break;
                    }

                case "Cc":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R136");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R136
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x17,
                            token,
                            "MS-ASWBXML",
                            136,
                            @"[In Code Page 2: Email] [Tag name] Cc [Token] 0x17 [supports protocol versions] All");

                        break;
                    }

                case "From":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R137");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R137
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x18,
                            token,
                            "MS-ASWBXML",
                            137,
                            @"[In Code Page 2: Email] [Tag name] From [Token] 0x18 [supports protocol versions] All");

                        break;
                    }

                case "ReplyTo":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R138");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R138
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x19,
                            token,
                            "MS-ASWBXML",
                            138,
                            @"[In Code Page 2: Email] [Tag name] ReplyTo [Token] 0x19 [supports protocol versions] All");

                        break;
                    }

                case "AllDayEvent":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R139");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R139
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1A,
                            token,
                            "MS-ASWBXML",
                            139,
                            @"[In Code Page 2: Email] [Tag name] AllDayEvent [Token] 0x1A [supports protocol versions] All");

                        break;
                    }

                case "Categories":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R140");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R140
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1B,
                            token,
                            "MS-ASWBXML",
                            140,
                            @"[In Code Page 2: Email] [Tag name] Categories [Token] 0x1B [supports protocol versions] 14.0, 14.1, 16.0");

                        break;
                    }

                case "Category":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R143");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R143
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1C,
                            token,
                            "MS-ASWBXML",
                            143,
                            @"[In Code Page 2: Email] [Tag name] Category [Token] 0x1C [supports protocol versions] 14.0, 14.1, 16.0");

                        break;
                    }

                case "DtStamp":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R145");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R145
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1D,
                            token,
                            "MS-ASWBXML",
                            145,
                            @"[In Code Page 2: Email] [Tag name] DtStamp [Token] 0x1D [supports protocol versions] All");

                        break;
                    }

                case "EndTime":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R146");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R146
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1E,
                            token,
                            "MS-ASWBXML",
                            146,
                            @"[In Code Page 2: Email] [Tag name] EndTime [Token] 0x1E [supports protocol versions] All");

                        break;
                    }

                case "InstanceType":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R147");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R147
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1F,
                            token,
                            "MS-ASWBXML",
                            147,
                            @"[In Code Page 2: Email] [Tag name] InstanceType [Token] 0x1F [supports protocol versions] All");

                        break;
                    }

                case "BusyStatus":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R148");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R148
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x20,
                            token,
                            "MS-ASWBXML",
                            148,
                            @"[In Code Page 2: Email] [Tag name] BusyStatus [Token] 0x20 [supports protocol versions] All");

                        break;
                    }

                case "Location":
                    {
                        if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0"))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R149");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R149
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x21,
                                token,
                                "MS-ASWBXML",
                                149,
                                @"[In Code Page 2: Email] [Tag name] Location - see note 3 following this table [Token] 0x21 [supports protocol versions] 2.5, 12.0, 12.1, 14.0, 14.1");
                        }

                        this.isLocationExistInCodePage2 = true;
                        break;
                    }

                case "MeetingRequest":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R150");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R150
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x22,
                            token,
                            "MS-ASWBXML",
                            150,
                            @"[In Code Page 2: Email] [Tag name] MeetingRequest [Token] 0x22 [supports protocol versions] All");

                        break;
                    }

                case "Organizer":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R151");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R151
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x23,
                            token,
                            "MS-ASWBXML",
                            151,
                            @"[In Code Page 2: Email] [Tag name] Organizer [Token] 0x23 [supports protocol versions] All");

                        break;
                    }

                case "RecurrenceId":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R152");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R152
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x24,
                            token,
                            "MS-ASWBXML",
                            152,
                            @"[In Code Page 2: Email] [Tag name] RecurrenceId [Token] 0x24 [supports protocol versions] All");

                        break;
                    }

                case "Reminder":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R153");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R153
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x25,
                            token,
                            "MS-ASWBXML",
                            153,
                            @"[In Code Page 2: Email] [Tag name] Reminder [Token] 0x25 [supports protocol versions] All");

                        break;
                    }

                case "ResponseRequested":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R154");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R154
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x26,
                            token,
                            "MS-ASWBXML",
                            154,
                            @"[In Code Page 2: Email] [Tag name] ResponseRequested [Token] 0x26 [supports protocol versions] All");

                        break;
                    }

                case "Recurrences":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R155");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R155
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x27,
                            token,
                            "MS-ASWBXML",
                            155,
                            @"[In Code Page 2: Email] [Tag name] Recurrences [Token] 0x27 [supports protocol versions] All");

                        break;
                    }

                case "Recurrence":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R156");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R156
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x28,
                            token,
                            "MS-ASWBXML",
                            156,
                            @"[In Code Page 2: Email] [Tag name] Recurrence [Token] 0x28 [supports protocol versions] All");

                        break;
                    }

                case "Type":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R157");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R157
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x29,
                            token,
                            "MS-ASWBXML",
                            157,
                            @"[In Code Page 2: Email] [Tag name] Type [Token] 0x29 [supports protocol versions] All");

                        break;
                    }

                case "Until":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R158");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R158
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x2A,
                            token,
                            "MS-ASWBXML",
                            158,
                            @"[In Code Page 2: Email] [Tag name] Until [Token] 0x2A [supports protocol versions] All");

                        break;
                    }

                case "Occurrences":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R159");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R159
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x2B,
                            token,
                            "MS-ASWBXML",
                            159,
                            @"[In Code Page 2: Email] [Tag name] Occurrences [Token] 0x2B [supports protocol versions] All");

                        break;
                    }

                case "Interval":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R160");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R160
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x2C,
                            token,
                            "MS-ASWBXML",
                            160,
                            @"[In Code Page 2: Email] [Tag name] Interval [Token] 0x2C [supports protocol versions] All");

                        break;
                    }

                case "DayOfWeek":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R161");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R161
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x2D,
                            token,
                            "MS-ASWBXML",
                            161,
                            @"[In Code Page 2: Email] [Tag name] DayOfWeek [Token] 0x2D [supports protocol versions] All");

                        break;
                    }

                case "DayOfMonth":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R162");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R162
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x2E,
                            token,
                            "MS-ASWBXML",
                            162,
                            @"[In Code Page 2: Email] [Tag name] DayOfMonth [Token] 0x2E [supports protocol versions] All");

                        break;
                    }

                case "WeekOfMonth":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R163");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R163
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x2F,
                            token,
                            "MS-ASWBXML",
                            163,
                            @"[In Code Page 2: Email] [Tag name] WeekOfMonth [Token] 0x2F [supports protocol versions] All");

                        break;
                    }

                case "MonthOfYear":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R164");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R164
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x30,
                            token,
                            "MS-ASWBXML",
                            164,
                            @"[In Code Page 2: Email] [Tag name] MonthOfYear [Token] 0x30 [supports protocol versions] All");

                        break;
                    }

                case "StartTime":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R165");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R165
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x31,
                            token,
                            "MS-ASWBXML",
                            165,
                            @"[In Code Page 2: Email] [Tag name] StartTime [Token] 0x31 [supports protocol versions] All");

                        break;
                    }

                case "Sensitivity":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R166");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R166
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x32,
                            token,
                            "MS-ASWBXML",
                            166,
                            @"[In Code Page 2: Email] [Tag name] Sensitivity [Token] 0x32 [supports protocol versions] All");

                        break;
                    }

                case "TimeZone":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R167");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R167
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x33,
                            token,
                            "MS-ASWBXML",
                            167,
                            @"[In Code Page 2: Email] [Tag name] TimeZone [Token] 0x33 [supports protocol versions] All");

                        break;
                    }

                case "GlobalObjId":
                    {
                        if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0"))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R168");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R168
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x34,
                                token,
                                "MS-ASWBXML",
                                168,
                                @"[In Code Page 2: Email] [Tag name] GlobalObjId - see note 4 following this table [Token] 0x34 [supports protocol versions] 2.5, 12.0, 12.1, 14.0, 14.1");
                        }

                        this.isGlobalObjIdExistInCodePage2 = true;
                        break;
                    }

                case "ThreadTopic":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R169");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R169
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x35,
                            token,
                            "MS-ASWBXML",
                            169,
                            @"[In Code Page 2: Email] [Tag name] ThreadTopic [Token] 0x35 [supports protocol versions] All");

                        break;
                    }

                case "InternetCPID":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R170");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R170
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x39,
                            token,
                            "MS-ASWBXML",
                            170,
                            @"[In Code Page 2: Email] [Tag name] InternetCPID [Token] 0x39 [supports protocol versions] All");

                        break;
                    }

                case "Flag":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R171");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R171
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x3A,
                            token,
                            "MS-ASWBXML",
                            171,
                            @"[In Code Page 2: Email] [Tag name] Flag [Token] 0x3A [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "Status":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R172");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R172
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x3B,
                            token,
                            "MS-ASWBXML",
                            172,
                            @"[In Code Page 2: Email] [Tag name] Status [Token] 0x3B [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "ContentClass":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R173");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R173
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x3C,
                            token,
                            "MS-ASWBXML",
                            173,
                            @"[In Code Page 2: Email] [Tag name] ContentClass [Token] 0x3C [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "FlagType":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R174");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R174
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x3D,
                            token,
                            "MS-ASWBXML",
                            174,
                            @"[In Code Page 2: Email] [Tag name] FlagType [Token] 0x3D [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "CompleteTime":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R175");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R175
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x3E,
                            token,
                            "MS-ASWBXML",
                            175,
                            @"[In Code Page 2: Email] [Tag name] CompleteTime [Token] 0x3E [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "DisallowNewTimeProposal":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R176");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R176
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x3F,
                            token,
                            "MS-ASWBXML",
                            176,
                            @"[In Code Page 2: Email] [Tag name] DisallowNewTimeProposal [Token] 0x3F [supports protocol versions] 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 9.
        /// </summary>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified. </param>
        private void VerifyRequirementsRelateToCodePage9(string tagName, byte token)
        {
            switch (tagName)
            {
                case "DateCompleted":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R285");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R285
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0B,
                            token,
                            "MS-ASWBXML",
                            285,
                            @"[In Code Page 9: Tasks] [Tag name] DateCompleted [Token] 0x0B [supports protocol versions] All");

                        break;
                    }

                case "DueDate":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R286");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R286
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            286,
                            @"[In Code Page 9: Tasks] [Tag name] DueDate [Token] 0x0C [supports protocol versions] All");

                        break;
                    }

                case "UtcDueDate":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R287");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R287
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0D,
                            token,
                            "MS-ASWBXML",
                            287,
                            @"[In Code Page 9: Tasks] [Tag name] UtcDueDate [Token] 0x0D [supports protocol versions] All");

                        break;
                    }

                case "ReminderSet":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R301");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R301
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1B,
                            token,
                            "MS-ASWBXML",
                            301,
                            @"[In Code Page 9: Tasks] [Tag name] ReminderSet [Token] 0x1B [supports protocol versions] All");

                        break;
                    }

                case "ReminderTime":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R302");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R302
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1C,
                            token,
                            "MS-ASWBXML",
                            302,
                            @"[In Code Page 9: Tasks] [Tag name] ReminderTime [Token] 0x1C [supports protocol versions] All");

                        break;
                    }

                case "StartDate":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R304");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R304
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1E,
                            token,
                            "MS-ASWBXML",
                            304,
                            @"[In Code Page 9: Tasks] [Tag name] StartDate [Token] 0x1E [supports protocol versions] All");

                        break;
                    }

                case "UtcStartDate":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R305");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R305
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1F,
                            token,
                            "MS-ASWBXML",
                            305,
                            @"[In Code Page 9: Tasks] [Tag name] UtcStartDate [Token] 0x1F [supports protocol versions] All");

                        break;
                    }

                case "OrdinalDate":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R307");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R307
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x22,
                            token,
                            "MS-ASWBXML",
                            307,
                            @"[In Code Page 9: Tasks] [Tag name] OrdinalDate [Token] 0x22 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "SubOrdinalDate":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R308");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R308
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x23,
                            token,
                            "MS-ASWBXML",
                            308,
                            @"[In Code Page 9: Tasks] [Tag name] SubOrdinalDate [Token] 0x23 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 17.
        /// </summary>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified. </param>
        private void VerifyRequirementsRelateToCodePage17(string tagName, byte token)
        {
            switch (tagName)
            {
                case "Body":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R452");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R452
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            452,
                            @"[In Code Page 17: AirSyncBase] [Tag name] Body [Token] 0x0A [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "Attachments":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R456");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R456
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0E,
                            token,
                            "MS-ASWBXML",
                            456,
                            @"[In Code Page 17: AirSyncBase] [Tag name] Attachments [Token] 0x0E [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "Attachment":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R457");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R457
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0F,
                            token,
                            "MS-ASWBXML",
                            457,
                            @"[In Code Page 17: AirSyncBase] [Tag name] Attachment [Token] 0x0F [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "DisplayName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R458");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R458
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x10,
                            token,
                            "MS-ASWBXML",
                            458,
                            @"[In Code Page 17: AirSyncBase] [Tag name] DisplayName [Token] 0x10 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "NativeBodyType":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R464");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R464
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x16,
                            token,
                            "MS-ASWBXML",
                            464,
                            @"[In Code Page 17: AirSyncBase] [Tag name] NativeBodyType [Token] 0x16 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "BodyPart":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R468");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R468
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1A,
                            token,
                            "MS-ASWBXML",
                            468,
                            @"[In Code Page 17: AirSyncBase] [Tag name] BodyPart [Token] 0x1A [supports protocol versions] 14.1, 16.0, 16.1");

                        break;
                    }

                case "Location":
                    {
                        this.isLocationExistInCodePage17 = true;
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 22.
        /// </summary>
        /// <param name="codePageNumber">code page number</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified. </param>
        private void VerifyRequirementsRelateToCodePage22(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "UmCallerID":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R606");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R606
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            606,
                            @"[In Code Page 22: Email2] [Tag name] UmCallerID [Token] 0x05 [supports protocol versions] 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "UmUserNotes":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R607");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R607
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x06,
                            token,
                            "MS-ASWBXML",
                            607,
                            @"[In Code Page 22: Email2] [Tag name] UmUserNotes [Token] 0x06 [supports protocol versions] 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "UmAttDuration":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R608");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R608
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            608,
                            @"[In Code Page 22: Email2] [Tag name] UmAttDuration [Token] 0x07 [supports protocol versions] 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "UmAttOrder":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R609");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R609
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x08,
                            token,
                            "MS-ASWBXML",
                            609,
                            @"[In Code Page 22: Email2] [Tag name] UmAttOrder [Token] 0x08 [supports protocol versions] 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "ConversationId":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R610");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R610
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x09,
                            token,
                            "MS-ASWBXML",
                            610,
                            @"[In Code Page 22: Email2] [Tag name] ConversationId [Token] 0x09 [supports protocol versions] 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "ConversationIndex":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R611");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R611
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            611,
                            @"[In Code Page 22: Email2] [Tag name] ConversationIndex [Token] 0x0A [supports protocol versions] 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "LastVerbExecuted":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R612");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R612
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0B,
                            token,
                            "MS-ASWBXML",
                            612,
                            @"[In Code Page 22: Email2] [Tag name] LastVerbExecuted [Token] 0x0B [supports protocol versions] 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "LastVerbExecutionTime":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R613");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R613
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            613,
                            @"[In Code Page 22: Email2] [Tag name] LastVerbExecutionTime [Token] 0x0C [supports protocol versions] 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "ReceivedAsBcc":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R614");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R614
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0D,
                            token,
                            "MS-ASWBXML",
                            614,
                            @"[In Code Page 22: Email2] [Tag name] ReceivedAsBcc [Token] 0x0D [supports protocol versions] 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "Sender":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R615");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R615
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0E,
                            token,
                            "MS-ASWBXML",
                            615,
                            @"[In Code Page 22: Email2] [Tag name] Sender [Token] 0x0E [supports protocol versions] 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "CalendarType":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R616");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R616
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0F,
                            token,
                            "MS-ASWBXML",
                            616,
                            @"[In Code Page 22: Email2] [Tag name] CalendarType [Token] 0x0F [supports protocol versions] 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "IsLeapMonth":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R617");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R617
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x10,
                            token,
                            "MS-ASWBXML",
                            617,
                            @"[In Code Page 22: Email2] [Tag name] IsLeapMonth [Token] 0x10 [supports protocol versions] 14.0, 14.1, 16.0, 16.1");

                        break;
                    }

                case "AccountId":
                    break;

                case "FirstDayOfWeek":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R619");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R619
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x12,
                            token,
                            "MS-ASWBXML",
                            619,
                            @"[In Code Page 22: Email2] [Tag name] FirstDayOfWeek [Token] 0x12 [supports protocol versions] 14.1, 16.0, 16.1");

                        break;
                    }

                case "MeetingMessageType":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R620");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R620
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x13,
                            token,
                            "MS-ASWBXML",
                            620,
                            @"[In Code Page 22: Email2] [Tag name] MeetingMessageType [Token] 0x13 [supports protocol versions] 14.1, 16.0, 16.1");

                        break;
                    }

                case "IsDraft":
                    {
                        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0") || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R851");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R851
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x15,
                                token,
                                "MS-ASWBXML",
                                851,
                                @"[In Code Page 22: Email2] [Tag name] IsDraft [Token] 0x15 [supports protocol versions] 16.0, 16.1");
                        }

                        break;
                    }

                case "Bcc":
                    {
                        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0") || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R852");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R852
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x16,
                                token,
                                "MS-ASWBXML",
                                852,
                                @"[In Code Page 22: Email2] [Tag name] Bcc [Token] 0x16 [supports protocol versions] 16.0, 16.1");
                        }

                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }
        #endregion
        #endregion
    }
}