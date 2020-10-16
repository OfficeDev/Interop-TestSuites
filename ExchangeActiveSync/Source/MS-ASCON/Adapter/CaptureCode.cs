namespace Microsoft.Protocols.TestSuites.MS_ASCON
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-ASCON.
    /// </summary>
    public partial class MS_ASCONAdapter
    {
        #region Verify common requirements
        /// <summary>
        /// This method is used to verify common requirements.
        /// </summary>
        private void VerifyCommonRequirements()
        {
            // If the schema validation is successful, then following requirements can be captured.
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "Schema validation should be successful.");
            
            bool testIsSupported = bool.Parse(Common.GetConfigurationPropertyValue(string.Format("{0}_Supported", this.Site.DefaultProtocolDocShortName), this.Site));
            if (testIsSupported)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "The implementation does support this specification [MS-ASCON]. (Exchange 2010 and above follow this behavior.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R5");

            // Verify MS-ASCON requirement: MS-ASCON_R5
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
                5,
                @"[In Transport] The XML markup that constitutes the request body or the response body that is transmitted between the client and the server uses Wireless Application Protocol (WAP) Binary XML (WBXML), as specified in [MS-ASWBXML].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R7");

            // Verify MS-ASCON requirement: MS-ASCON_R7
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
                7,
                @"[In Message Syntax] The XML markup that is used by this protocol [MS-ASCON] MUST be well-formed XML as specified in [XML].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R188");

            // Verify MS-ASCON requirement: MS-ASCON_R188
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
                188,
                @"[In Abstract Data Model] Command response: A WBXML-formatted message that adheres to the command schemas specified in [MS-ASCMD].");
        }
        #endregion

        #region Verify requirements of Sync command
        /// <summary>
        /// This method is used to verify requirements related to Sync command.
        /// </summary>
        /// <param name="syncResult">The Sync result returned from the server.</param>
        private void VerifySyncCommandResponse(Sync syncResult)
        {
            // If the schema validation is successful, then following requirements can be captured.
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "Schema validation should be successful.");

            if (syncResult.Email.Subject != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R74");

                // Verify MS-ASCON requirement: MS-ASCON_R74
                // The ConversationId element is not null, so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    syncResult.Email.ConversationId,
                    74,
                    @"[In ConversationId (Sync)] The email2:ConversationId element ([MS-ASEMAIL] section 2.2.2.21) is a required child element of the airsync:ApplicationData element ([MS-ASCMD] section 2.2.3.11) in a Sync command response ([MS-ASCMD] section 2.2.1.21) that specifies the conversation ID of a conversation.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R76");

                // Verify MS-ASCON requirement: MS-ASCON_R76
                // The ConversationId element is not null, so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    syncResult.Email.ConversationId,
                    76,
                    @"[In ConversationId (Sync)] The value of this [email2:ConversationId] element is a byte array, as specified in [MS-ASDTYPE] section 2.7.1.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R195");

                // Verify MS-ASCON requirement: MS-ASCON_R195
                // The ConversationId element is not null, so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    syncResult.Email.ConversationId,
                    195,
                    @"[In Higher-Layer Triggered Events] The conversation ID is specified by the email2:ConversationId element ([MS-ASEMAIL] section 2.2.2.21) that is included in the Email class.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R79");

                // Verify MS-ASCON requirement: MS-ASCON_R79
                // The ConversationIndex element is not null, so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    syncResult.Email.ConversationIndex,
                    79,
                    @"[In ConversationIndex] The email2:ConversationIndex element ([MS-ASEMAIL] section 2.2.2.22) is a required child element of the airsync:ApplicationData element ([MS-ASCMD] section 2.2.3.11) in a Sync command response ([MS-ASCMD] section 2.2.1.21) that specifies the conversation index for an e-mail message.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R196");

                // Verify MS-ASCON requirement: MS-ASCON_R196
                // The ConversationIndex element is not null, so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    syncResult.Email.ConversationIndex,
                    196,
                    @"[In Higher-Layer Triggered Events] The conversation index is specified by the email2:ConversationIndex element ([MS-ASEMAIL] section 2.2.2.22) that is included in the Email class.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R81");

                // Verify MS-ASCON requirement: MS-ASCON_R81
                // The ConversationIndex element is not null, so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    syncResult.Email.ConversationIndex,
                    81,
                    @"[In ConversationIndex] The value of this element [email2:ConversationIndex] is a byte array, as specified in [MS-ASDTYPE] section 2.7.1.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R93");

                // Verify MS-ASCON requirement: MS-ASCON_R93
                // The ConversationIndex element is not null, so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    syncResult.Email.ConversationIndex,
                    93,
                    @"[In ConversationIndex] The content of the email2:ConversationIndex element is transferred as a byte array within the WBXML tags.");

                this.VerifyStringDataType();

                if (syncResult.Email.BodyPart != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R46");

                    // Verify MS-ASCON requirement: MS-ASCON_R46
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        46,
                        @"[In BodyPart] The airsyncbase:BodyPart element is a container ([MS-ASDTYPE] section 2.2) element.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R47");

                    // Verify MS-ASCON requirement: MS-ASCON_R47
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        47,
                        @"[In BodyPart] It [airsyncbase:BodyPart element] has the following child elements:
airsyncbase:Status ([MS-ASAIRS] section 2.2.2.37)
airsyncbase:Type ([MS-ASAIRS] section 2.2.2.41.2)
airsyncbase:EstimatedDataSize ([MS-ASAIRS] section 2.2.2.23.3)
airsyncbase:Truncated ([MS-ASAIRS] section 2.2.2.39.2)
airsyncbase:Data ([MS-ASAIRS] section 2.2.2.20.2)
airsyncbase:Preview ([MS-ASAIRS] section 2.2.2.35.2).");

                    this.VerifyContainerDataType();
                }
            }
        }
        #endregion

        #region Verify requirements related to ItemOperations command
        /// <summary>
        /// This method is used to verify requirements related to ItemOperations command.
        /// </summary>
        /// <param name="itemOperationsResponse">The ItemOperations command response returned from the server.</param>
        private void VerifyItemOperationsCommandResponse(ItemOperationsResponse itemOperationsResponse)
        {
            // If the schema validation is successful, then following requirements can be captured.
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "Schema validation should be successful.");

            if (itemOperationsResponse.ResponseData.Response.Move != null && itemOperationsResponse.ResponseData.Response.Move.Length != 0)
            {
                foreach (Response.ItemOperationsResponseMove move in itemOperationsResponse.ResponseData.Response.Move)
                {
                    bool isVerifiedR126 = move.ConversationId != null && move.Status != null;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R126");
                    Site.Log.Add(LogEntryKind.Debug, "The ConversationId element is: {0}.", move.ConversationId);
                    Site.Log.Add(LogEntryKind.Debug, "The Status element is: {0}.", move.Status);

                    // Verify MS-ASCON requirement: MS-ASCON_R126
                    // The ConversationId element and Status element are not null, so this requirement can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        isVerifiedR126,
                        126,
                        @"[In Move] It [itemoperations:Move element ] has the following child elements:
itemoperations:ConversationId (section 2.2.2.3.1)
itemoperations:Status (section 2.2.2.10).");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R69");

                    // Verify MS-ASCON requirement: MS-ASCON_R69
                    // The ConversationId element is not null, so this requirement can be captured.
                    Site.CaptureRequirementIfIsNotNull(
                        move.ConversationId,
                        69,
                        @"[In ConversationId (ItemOperations)] In an ItemOperations command response, the itemoperations:ConversationId is a required child element of the itemoperations:Move element that specifies the conversation ID of the conversation that was moved.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R70");

                    // Verify MS-ASCON requirement: MS-ASCON_R70
                    // The ConversationId element is not null and schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirementIfIsNotNull(
                        move.ConversationId,
                        70,
                        @"[In ConversationId (ItemOperations)] The value of this element [itemoperations:ConversationId] is a byte array, as specified in [MS-ASDTYPE] section 2.7.1.");

                    this.VerifyStringDataType();

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R139");

                    // Verify MS-ASCON requirement: MS-ASCON_R139
                    // The Status element is not null, so this requirement can be captured.
                    Site.CaptureRequirementIfIsNotNull(
                        move.Status,
                        139,
                        @"[In Status] The itemoperations:Status element ([MS-ASCMD] section 2.2.3.177.8) is a required child element of the itemoperations:Move element ([MS-ASCMD] section 2.2.3.117.1) in the ItemOperations command response ([MS-ASCMD] section 2.2.1.10) that specifies the status of the move operation.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R142");

                    // Verify MS-ASCON requirement: MS-ASCON_R142
                    // The Status element is not null and schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirementIfIsNotNull(
                        move.Status,
                        142,
                        @"[In Status] The value of this [itemoperations:Status] element is an integer, as specified in [MS-ASDTYPE] section 2.6.");

                    this.VerifyIntegerDataType();
                }
            }
        }
        #endregion

        #region Verify reuqirements of [MS-ASDTYPE]
        /// <summary>
        /// This method is used to verify the container data type related requirements.
        /// </summary>
        private void VerifyContainerDataType()
        {
            // If the schema validation is successful, then following requirements can be captured.
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "Schema validation should be successful.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R8");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R8
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                8,
                @"[In container Data Type] A container is an XML element that encloses other elements but has no value of its own.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R9");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R9
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                9,
                @"[In container Data Type] It [container] is a complex type with complex content, as specified in [XMLSCHEMA1/2] section 3.4.2.");
        }

        /// <summary>
        /// This method is used to verify the integer data type related requirements.
        /// </summary>
        private void VerifyIntegerDataType()
        {
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
        /// This method is used to verify the string data type related requirements.
        /// </summary>
        private void VerifyStringDataType()
        {
            // If the schema validation is successful, then following requirements can be captured.
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "Schema validation should be successful.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R88");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R88
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                88,
                @"[In string Data Type] A string is a chunk of Unicode text.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R90");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R90
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
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

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R94
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                94,
                @"[In string Data Type] Elements of these types [ActiveSync defines several conventions for strings that adhere to commonly used formats]are defined as string types in XML schemas.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R97");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R97
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                97,
                @"[In Byte Array] The structure is comprised of a length, which is expressed as a multi-byte integer, as specified in [WBXML1.2], followed by that many bytes of data.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R98");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R98
            // ActiveSyncClient encoded byte array as opaque data, so if response is successfully returned this requirement can be captured.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                98,
                @"[In Byte Array] Elements with a byte array structure MUST be encoded and transmitted as [WBXML1.2] opaque data.");
        }
        #endregion

        #region Verify requirements of [MS-ASWBXML] for code page 5, 17, 20 and 22
        /// <summary>
        /// Verify WBXML Capture for WBXML process.
        /// </summary>
        private void VerifyWBXMLCapture()
        {
            // Get decode data and capture requirement for decode processing.
            Dictionary<string, int> decodedData = this.activeSyncClient.GetMSASWBXMLImplementationInstance().DecodeDataCollection;

            if (decodedData != null)
            {
                // Check out all tag-token
                foreach (KeyValuePair<string, int> decodeDataItem in decodedData)
                {
                    byte token;
                    string tagName = Common.GetTagName(decodeDataItem.Key, out token);
                    int codepage = decodeDataItem.Value;
                    bool isValidCodePage = codepage >= 0 && codepage <= 24;
                    Site.Assert.IsTrue(isValidCodePage, "Code page value should between 0-24, the actual value is :{0}", codepage);

                    // Begin to capture requirements of code page 5.
                    if (5 == codepage)
                    {
                        switch (tagName)
                        {
                            case "MoveItems":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R234");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R234
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x05,
                                        token,
                                        "MS-ASWBXML",
                                        234,
                                        @"[In Code Page 5: Move] [Tag name] MoveItems [Token] 0x05 [supports protocol versions] All");

                                    break;
                                }

                            case "Response":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R239");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R239
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x0A,
                                        token,
                                        "MS-ASWBXML",
                                        239,
                                        @"[In Code Page 5: Move] [Tag name] Response [Token] 0x0A [supports protocol versions] All");

                                    break;
                                }

                            case "Status":
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R240");

                                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R240
                                    Site.CaptureRequirementIfAreEqual<byte>(
                                        0x0B,
                                        token,
                                        "MS-ASWBXML",
                                        240,
                                        @"[In Code Page 5: Move] [Tag name] Status [Token] 0x0B [supports protocol versions] All");

                                    break;
                                 }
                        }
                    }

                    // Begin to capture requirements of code page 17.
                    if (17 == codepage && tagName == "BodyPart")
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
                    }

                    // Begin to capture requirements of code page 20.
                    if (20 == codepage)
                    {
                        if (tagName == "Status")
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R571");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R571
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x0D,
                                token,
                                "MS-ASWBXML",
                                571,
                                @"[In Code Page 20: ItemOperations] [Tag name] Status [Token] 0x0D [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");
                        }
                        else if (tagName == "ConversationId")
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R582");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R582
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x18,
                                token,
                                "MS-ASWBXML",
                                582,
                                @"[In Code Page 20: ItemOperations] [Tag name] ConversationId [Token] 0x18 [supports protocol versions] 14.0, 14.1, 16.0, 16.1");
                        }
                        else if (tagName == "Move")
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R580");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R580
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x16,
                                token,
                                "MS-ASWBXML",
                                580,
                                @"[In Code Page 20: ItemOperations] [Tag name] Move [Token] 0x16 [supports protocol versions] 14.0, 14.1, 16.0, 16.1");
                        }
                    }

                    // Begin to capture requirements of code page 22.
                    if (22 == codepage)
                    {
                        if (tagName == "ConversationId")
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
                        }
                        else if (tagName == "ConversationIndex")
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
                        }
                    }
                }
            }
        }
        #endregion
    }
}