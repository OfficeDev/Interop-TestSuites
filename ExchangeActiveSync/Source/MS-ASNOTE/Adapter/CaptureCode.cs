namespace Microsoft.Protocols.TestSuites.MS_ASNOTE
{
    using System.Collections.Generic;
    using System.Text.RegularExpressions;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides the methods to write capture code.
    /// </summary>
    public partial class MS_ASNOTEAdapter
    {
        /// <summary>
        /// This method is used to verify transport related requirement.
        /// </summary>
        private void VerifyTransport()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R2");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R2
            // ActiveSyncClient encodes XML request into WBXML and decodes WBXML to XML response, capture it directly if server responses succeed.
            Site.CaptureRequirement(
                2,
                @"[In Transport] The XML markup that constitutes the request body or the response body that is transmitted between the client and the server uses Wireless Application Protocol (WAP) Binary XML (WBXML), as specified in [MS-ASWBXML].");
        }

        /// <summary>
        /// This method is used to verify the Message Syntax related requirements.
        /// </summary>
        private void VerifyMessageSyntax()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R3");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R3
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                3,
                @"[In Message Syntax] The markup that is used by this protocol MUST be well-formed XML, as specified in [XML].");
        }

        /// <summary>
        /// This method is used to verify the Note related requirements
        /// </summary>
        /// <param name="note">note information</param>
        /// <param name="isSchemaRestricted">Specified whether the note elements in response are restricted by schema element</param>
        private void VerifyNote(Note note, bool isSchemaRestricted)
        {
            if (note.Body != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R35");

                // Verify MS-ASNOTE requirement: MS-ASNOTE_R35
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    35,
                    @"[In Body] The airsyncbase:Body element is a container ([MS-ASDTYPE] section 2.2) element that specifies the text of the note.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R37, the actual value of Type element is {0}", note.Body.Type);

                // Verify MS-ASNOTE requirement: MS-ASNOTE_R37
                bool isVerifyR37 = note.Body.Type == 1 || note.Body.Type == 2 || note.Body.Type == 3;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR37,
                    37,
                    @"[In Body] The airsyncbase:Type element ([MS-ASAIRS] section 2.2.2.22.1) (a required child element of the airsyncbase:Body element) MUST be set to one of the following values[1, 2, 3].");
                this.VerifyContainerDataType();
            }

            if (note.Categories != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R44");

                // Verify MS-ASNOTE requirement: MS-ASNOTE_R44
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    44,
                    @"[In Categories] The Categories element is a container ([MS-ASDTYPE] section 2.2) element that specifies a collection of labels assigned to the note.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R47");

                // Verify MS-ASNOTE requirement: MS-ASNOTE_R47
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    47,
                    @"[In Categories] [The Categories element has the following child element:] Category (section 2.2.2.3): Zero or more instances of this element are allowed.");

                if (note.Categories.Category.Length != 0)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R50");

                    // Verify MS-ASNOTE requirement: MS-ASNOTE_R50
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        50,
                        @"[In Category] The value of this element[Category] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                    this.VerifyStringDataType();
                }

                this.VerifyContainerDataType();
            }

            if (note.Subject != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R60");

                // Verify MS-ASNOTE requirement: MS-ASNOTE_R60
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    60,
                    @"[In Subject] The value of this element[Subject] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");
                
                this.VerifyStringDataType();
            }

            // MessageClass is a required element if not restricted by schema element in ItemOperations command request.
            if (!isSchemaRestricted)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R54");

                // Verify MS-ASNOTE requirement: MS-ASNOTE_R54
                Site.CaptureRequirementIfIsNotNull(
                    note.MessageClass,
                    54,
                    @"[In MessageClass] The MessageClass element is a required element that specifies the Inter-Personal Mail (IPM) type of the note.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R56");

                // Verify MS-ASNOTE requirement: MS-ASNOTE_R56
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    56,
                    @"[In MessageClass] The value of this element[MessageClass] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R57, the actual value of MessageClass element is {0}", note.MessageClass);

                // Verify MS-ASNOTE requirement: MS-ASNOTE_R57
                bool isVerifyR57 = note.MessageClass.Equals("IPM.StickyNote", System.StringComparison.CurrentCulture) || note.MessageClass.StartsWith("IPM.StickyNote.", System.StringComparison.CurrentCulture);
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR57,
                    57,
                    @"[In MessageClass] The value of the MessageClass element MUST be either ""IPM.StickyNote"" or ""IPM.StickyNote.*"", where ""*"" represents an arbitrary string chosen by the client or server.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R53");

                // Verify MS-ASNOTE requirement: MS-ASNOTE_R53
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    53,
                    @"[In LastModifiedDate] The value of the LastModifiedDate element is a string data type represented as a Compact DateTime ([MS-ASDTYPE] section 2.6.5).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R125");

                // Verify MS-ASNOTE requirement: MS-ASNOTE_R125
                Site.CaptureRequirementIfIsTrue(
                    note.IsLastModifiedDateSpecified,
                    125,
                    @"[In LastModifiedDate Element] but it[The LastModifiedDate element (section 2.2.2.4) ] is required in the Sync command response ([MS-ASCMD] section 2.2.2.19).");

                Regex regex = new Regex(@"\d{4}(0\d|1[0-2])([0-2]\d|3[0-1])T([0-1]\d|2[0-3])[0-5]\d[0-5]\d?Z", RegexOptions.Singleline);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R12213");

                // Verify MS-ASNOTE requirement: MS-ASDTYPE_R12213
                Site.CaptureRequirementIfIsTrue(
                    regex.IsMatch(note.LastModifiedDateString),
                    "MS-ASDTYPE",
                    12213,
                    @"[In Compact DateTime] [The format of a Compact DateTime value is specified by the following Augmented Backus-Naur Form (ABNF) notation. ]
date_string   = year month day ""T"" hour minute seconds ""Z""
year          = 4*DIGIT
month         = (""0"" DIGIT) / ""10"" / ""11"" / ""12""
day           = (""0"" DIGIT) / (""1"" DIGIT) / (""2"" DIGIT) / ""30"" / ""31""
hour          = (""0"" DIGIT) / (""1"" DIGIT) / ""20"" / ""21"" / ""22"" / ""23""
minute        = (""0"" DIGIT) / (""1"" DIGIT) / (""2"" DIGIT) / (""3"" DIGIT) / (""4"" DIGIT) / (""5"" DIGIT)
seconds       = (""0"" DIGIT) / (""1"" DIGIT) / (""2"" DIGIT) / (""3"" DIGIT) / (""4"" DIGIT) / (""5"" DIGIT)");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R12211");

                // Verify MS-ASNOTE requirement: MS-ASDTYPE_R12211
                Site.CaptureRequirementIfIsTrue(
                    regex.IsMatch(note.LastModifiedDateString),
                    "MS-ASDTYPE",
                    12211,
                    @"[In Compact DateTime] A Compact DateTime value is a representation of a UTC date and time within an element of type xs:string, as specified in [XMLSCHEMA2/2] section 3.2.1.");

                this.VerifyStringDataType();
            }
        }

        /// <summary>
        /// This method is used to verify the ItemOperations response related requirements.
        /// </summary>
        /// <param name="itemOperationsResponse">Specified the ItemOperations Command Response returned from the server</param>
        private void VerifyItemOperationsResponse(ItemOperationsResponse itemOperationsResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R100");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R100
            Site.CaptureRequirementIfIsNotNull(
                itemOperationsResponse,
                100,
                @"[In Retrieving Details for One or More Notes] [If the client sends a ItemOperations command request to the server] The server responds with an ItemOperations command response ([MS-ASCMD] section 2.2.2.8).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R101");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R101
            Site.CaptureRequirementIfIsNotNull(
                itemOperationsResponse,
                101,
                @"[In ItemOperations Command Response] When a client uses an ItemOperations command request ([MS-ASCMD] section 2.2.2.8) to retrieve data from the server for one or more notes items, as specified in section 3.1.5.1, the server responds with an ItemOperations command response.");
        }

        /// <summary>
        /// This method is used to verify the ItemOperation response related requirements.
        /// </summary>
        /// <param name="itemResponse">Specified ItemStore result returned from the server</param>
        /// <param name="isSchemaRestricted">Specified whether the note elements in response are restricted by schema element</param>
        private void VerifyItemOperationResult(ItemOperationsStore itemResponse, bool isSchemaRestricted)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R104");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R104
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                104,
                @"[In ItemOperations Command Response] Notes class elements[airsyncbase:Body, Subject,  MessageClass, LastModifiedDate, Categories or Category] are returned as child elements of the itemoperations:Properties element ([MS-ASCMD] section 2.2.3.128) in the ItemOperations command response.");
            
            foreach (ItemOperations itemOperations in itemResponse.Items)
            {
                if (itemOperations.Note != null)
                {
                    this.VerifyNote(itemOperations.Note, isSchemaRestricted);
                }
            }

            this.VerifyMessageSyntax();
        }

        /// <summary>
        /// This method is used to verify the Sync response related requirements.
        /// </summary>
        /// <param name="syncResponse">Specified the Sync Command Response returned from the server</param>
        private void VerifySyncResponse(SyncResponse syncResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R96");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R96
            Site.CaptureRequirementIfIsNotNull(
                syncResponse,
                96,
                @"[In Synchronizing Notes Data Between Client and Server] [If the client sends a Sync command request to the server] The server responds with a Sync command response ([MS-ASCMD] section 2.2.2.19).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R110");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R110
            Site.CaptureRequirementIfIsNotNull(
                syncResponse,
                110,
                @"[In Sync Command Response] When a client uses the Sync command request ([MS-ASCMD] section 2.2.2.19) to synchronize its Notes class items for a specified user with the notes currently stored by the server, as specified in section 3.1.5.3, the server responds with a Sync command response.");
        }

        /// <summary>
        /// This method is used to verify the Sync response related requirements.
        /// </summary>
        /// <param name="syncResponse">Specified the SyncStore result returned from the server</param>
        private void VerifySyncResult(SyncStore syncResponse)
        {
            if (syncResponse.AddElements != null)
            {
                foreach (Sync sync in syncResponse.AddElements)
                {
                    this.VerifyNote(sync.Note, false);
                }
            }

            if (syncResponse.ChangeElements != null)
            {
                foreach (Sync sync in syncResponse.ChangeElements)
                {
                    this.VerifyNote(sync.Note, false);
                }
            }            

            this.VerifyMessageSyntax();
        }

        /// <summary>
        /// This method is used to verify the Search response related requirements.
        /// </summary>
        /// <param name="searchResponse">Specified the Search Command Response returned from the server</param>
        private void VerifySearchResponse(SearchResponse searchResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R98");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R98
            Site.CaptureRequirementIfIsNotNull(
                searchResponse,
                98,
                @"[In Searching for Notes Data] [If the client sends a search command request to the server] The server responds with a Search command response ([MS-ASCMD] section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R106");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R106
            Site.CaptureRequirementIfIsNotNull(
                searchResponse,
                106,
                @"[In Search Command Response] When a client uses the Search command request ([MS-ASCMD] section 2.2.2.14) to retrieve Notes class items that match the criteria specified by the client, as specified in section 3.1.5.2, the server responds with a Search command response.");
        }

        /// <summary>
        /// This method is used to verify the Search Command response related requirements.
        /// </summary>
        /// <param name="searchResponse">Specified the SearchStore result from the server</param>
        private void VerifySearchResult(SearchStore searchResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASNOTE_R108");

            // Verify MS-ASNOTE requirement: MS-ASNOTE_R108
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                108,
                @"[In Search Command Response] Notes class elements[airsyncbase:Body, Subject, MessageClass, LastModifiedDate, Categories or Category] are returned as child elements of the search:Properties element ([MS-ASCMD] section 2.2.3.128) in the Search command response.");

            foreach (Search search in searchResponse.Results)
            {
                this.VerifyNote(search.Note, false);
            }

            this.VerifyMessageSyntax();
        }

        /// <summary>
        /// This method is used to verify the container data type related requirements.
        /// </summary>
        private void VerifyContainerDataType()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R8");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R8
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                8,
                @"[In container Data Type] A container is an XML element that encloses other elements but has no value of its own.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R9");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R9
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                9,
                @"[In container Data Type] It [container] is a complex type with complex content, as specified in [XMLSCHEMA1/2] section 3.4.2.");
        }

        /// <summary>
        /// This method is used to verify the string data type related requirements.
        /// </summary>
        private void VerifyStringDataType()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R88");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R88
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                88,
                @"[In string Data Type] A string is a chunk of Unicode text.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R90");

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

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R94
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                94,
                @"[In string Data Type] Elements of these types [ActiveSync defines several conventions for strings that adhere to commonly used formats]are defined as string types in XML schemas.");
        }

        #region Capture requirements of [MS-ASWBXML] code page 23

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
                    string codePageName = Common.GetCodePageName(decodeDataItem.Key);
                    int codepage = decodeDataItem.Value;
                    bool isValidCodePage = codepage >= 0 && codepage <= 24;
                    Site.Assert.IsTrue(isValidCodePage, "Code page value should between 0-24,actual value is :{0}", codepage);

                    // begin to capture requirement
                    if (23 == codepage)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R33");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R33
                        Site.CaptureRequirementIfAreEqual<string>(
                            "notes",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            33,
                            @"[In Code Pages] [This algorithm supports] [Code page] 23[that indicates] [XML namespace] Notes");

                        this.CaptureRequirementsRelateToCodePage23(codepage, tagName, token);
                    }
                }
            }
        }

        #region Tag and token mapping captures.

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 23
        /// </summary>
        /// <param name="codePageNumber">code page number</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified. </param>
        private void CaptureRequirementsRelateToCodePage23(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "Subject":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R626");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R626
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            626,
                            @"[In Code Page 23: Notes] [Tag name] Subject  [Token] 0x05");

                        break;
                    }

                case "MessageClass":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R627");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R627
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x06,
                            token,
                            "MS-ASWBXML",
                            627,
                            @"[In Code Page 23: Notes] [Tag name] MessageClass  [Token] 0x06");

                        break;
                    }

                case "LastModifiedDate":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R628");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R628
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            628,
                            @"[In Code Page 23: Notes] [Tag name] LastModifiedDate [Token] 0x07");

                        break;
                    }

                case "Categories":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R629");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R629
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x08,
                            token,
                            "MS-ASWBXML",
                            629,
                            @"[In Code Page 23: Notes] [Tag name] Categories [Token] 0x08");

                        break;
                    }

                case "Category":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R630");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R630
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x09,
                            token,
                            "MS-ASWBXML",
                            630,
                            @"[In Code Page 23: Notes] [Tag name] Category  [Token] 0x09");

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