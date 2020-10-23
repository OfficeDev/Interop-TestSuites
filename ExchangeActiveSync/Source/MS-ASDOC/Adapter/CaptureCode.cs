namespace Microsoft.Protocols.TestSuites.MS_ASDOC
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-ASDOC. 
    /// </summary>
    public partial class MS_ASDOCAdapter
    {
        /// <summary>
        /// This method is used to verify transport related requirement.
        /// </summary>
        private void VerifyTransport()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R4");

            // Verify MS-ASCAL requirement: MS-ASDOC_R4
            // ActiveSyncClient encodes XML request into WBXML and decodes WBXML to XML response, capture it directly if server responses succeed.
            Site.CaptureRequirement(
                4,
                @"[In Transport] The XML markup that constitutes the request body or the response body that is transmitted between the client and the server uses Wireless Application Protocol (WAP) Binary XML (WBXML), as specified in [MS-ASWBXML].");
        }

        /// <summary>
        /// Verify the Search operation.
        /// </summary>
        /// <param name="response">The response of the Search operation.</param>
        private void VerifySearch(SearchResponse response)
        {
            // Only verify the adapter capture code when response is successful.
            if (string.Equals(response.ResponseData.Response.Store.Status, "1"))
            {
                Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "Schema should be verified successfully.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R110");

                // Verify MS-ASDOC requirement: MS-ASDOC_R110
                Site.CaptureRequirementIfIsNotNull(
                    response.ResponseData,
                    110,
                    @"[In Search Command Response] When a client uses the Search command request ([MS-ASCMD] section 2.2.1.16) to retrieve Document class items that match the criteria specified by the client, as specified in section 3.1.5.2, the server responds with a Search command response.");

                foreach (Microsoft.Protocols.TestSuites.Common.Response.SearchResponseStoreResult result in response.ResponseData.Response.Store.Result)
                {
                    string linkId = null;
                    DateTime? lastModifiedDate = null;
                    DateTime? creationDate = null;
                    string displayName = null;
                    byte? isFolder = null;
                    byte? isHidden = null;
                    string contentLength = null;
                    string contentType = null;

                    for (int i = 0; i < result.Properties.ItemsElementName.Length; i++)
                    {
                        if (result.Properties.ItemsElementName[i] == Response.ItemsChoiceType6.LinkId)
                        {
                            linkId = (string)result.Properties.Items[i];
                        }

                        if (result.Properties.ItemsElementName[i] == Response.ItemsChoiceType6.CreationDate)
                        {
                            creationDate = (DateTime)result.Properties.Items[i];
                        }

                        if (result.Properties.ItemsElementName[i] == Response.ItemsChoiceType6.LastModifiedDate)
                        {
                            lastModifiedDate = (DateTime)result.Properties.Items[i];
                        }

                        if (result.Properties.ItemsElementName[i] == Response.ItemsChoiceType6.DisplayName)
                        {
                            displayName = (string)result.Properties.Items[i];
                        }

                        if (result.Properties.ItemsElementName[i] == Response.ItemsChoiceType6.IsHidden)
                        {
                            isHidden = (byte)result.Properties.Items[i];
                        }

                        if (result.Properties.ItemsElementName[i] == Response.ItemsChoiceType6.IsFolder)
                        {
                            isFolder = (byte)result.Properties.Items[i];
                        }

                        if (result.Properties.ItemsElementName[i] == Response.ItemsChoiceType6.ContentLength)
                        {
                            contentLength = (string)result.Properties.Items[i];
                        }

                        if (result.Properties.ItemsElementName[i] == Response.ItemsChoiceType6.ContentType)
                        {
                            contentType = (string)result.Properties.Items[i];
                        }
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R15");

                    // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R15
                    bool isVerifyR15 = null != lastModifiedDate && null != creationDate;

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR15,
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
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R16");

                    // Since schema is validated, this requirement can be captured directly.
                    Site.CaptureRequirement(
                        "MS-ASDTYPE",
                        16,
                        @"[In dateTime Data Type][in YYYY-MM-DDTHH:MM:SS.MSSZ ] The T serves as a separator, and the Z indicates that this time is in UTC.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R36");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R36
                    Site.CaptureRequirementIfIsNotNull(
                        creationDate,
                        36,
                        @"[In CreationDate] The CreationDate element is a required child element of the search:Properties element ([MS-ASCMD] section 2.2.3.139.3) for Document class document items in a Search command response ([MS-ASCMD] section 2.2.1.16) that specifies the date and time when the document or folder was created.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R39");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R39
                    Site.CaptureRequirementIfIsNotNull(
                        displayName,
                        39,
                        @"[In DisplayName] The DisplayName element is a required child element of the search:Properties element ([MS-ASCMD] section 2.2.3.139.3) for Document class items in a Search command response ([MS-ASCMD] section 2.2.1.16) that specifies the name of the document or folder as it is displayed to the user.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R41");

                    // If the element is returned and schema validation is successful, this requirement can be verified.
                    Site.CaptureRequirement(
                        41,
                        @"[In DisplayName] The value of this element[DisplayName] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R94");

                    // Since schema is validated, this requirement can be captured directly.
                    Site.CaptureRequirement(
                        "MS-ASDTYPE",
                        94,
                        @"[In string Data Type] Elements of these types [ActiveSync defines several conventions for strings that adhere to commonly used formats] Are defined as string types in XML schemas.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R48");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R48
                    Site.CaptureRequirementIfIsNotNull(
                        isHidden,
                        48,
                        @"[In IsHidden] The IsHidden element is a required child element of the search:Properties element ([MS-ASCMD] section 2.2.3.139.3) for Document class items in a Search command response ([MS-ASCMD] section 2.2.1.16) that specifies whether the document or folder is a hidden object.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R50");

                    // If the element is returned and schema validation is successful, this requirement can be verified.
                    Site.CaptureRequirement(
                        50,
                        @"[In IsHidden] The value of this element[IsHidden] is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.8.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R45, the value of IsFolder is {0}", isFolder);

                    // Verify MS-ASDOC requirement: MS-ASDOC_R45
                    bool isVerifyR45 = (byte)isFolder == 0 || (byte)isFolder == 1;

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR45,
                        45,
                        @"[In IsFolder] Valid values for this element are as follows[0,1].");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R42");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R42
                    Site.CaptureRequirementIfIsNotNull(
                        isFolder,
                        42,
                        @"[In IsFolder] The IsFolder element is a required child element of the search:Properties element ([MS-ASCMD] section 2.2.3.139.3) for Document class items in a Search command response ([MS-ASCMD] section 2.2.1.16) that specifies whether the item is a folder.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R44");

                    // If the element is returned and schema validation is successful, this requirement can be verified.
                    Site.CaptureRequirement(
                        44,
                        @"[In IsFolder] The value of this element[IsFolder] is an unsignedByte data type, as specified in [MS-ASDTYPE] section 2.8.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R51, the value of IsHidden is {0}", isHidden);

                    // Verify MS-ASDOC requirement: MS-ASDOC_R51
                    bool isVerifyR51 = (byte)isHidden == 0 || (byte)isHidden == 1;

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR51,
                        51,
                        @"[In IsHidden] The value of the IsHidden element MUST be one of the following values[0,1].");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R54");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R54
                    Site.CaptureRequirementIfIsNotNull(
                        lastModifiedDate,
                        54,
                        @"[In LastModifiedDate] The LastModifiedDate element is a required child element of the search:Properties element ([MS-ASCMD] section 2.2.3.139.3) for Document class items in a Search command response ([MS-ASCMD] section 2.2.1.16) that specifies the date and time that the document, or the folder, or the properties of either the document or folder were was last modified.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R123");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R123
                    Site.CaptureRequirementIfIsNotNull(
                        linkId,
                        123,
                        @"[In LinkId] It[LinkId] is a required child element of the search:Properties element ([MS-ASCMD] section 2.2.3.139.3) for Document class items in a Search command response.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R114");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R114
                    Site.CaptureRequirementIfIsNotNull(
                        result.Properties,
                        114,
                        @"[In Search Command Response] Document class elements are returned as child elements of the search:Properties element ([MS-ASCMD] section 2.2.3.139.3) in a Search command response.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R38,the value of CreationDate is:{0}", creationDate);

                    // Verify MS-ASDOC requirement: MS-ASDOC_R38
                    Site.CaptureRequirementIfIsNotNull(
                        creationDate,
                        38,
                        @"[In CreationDate] The value of this element[CreationDate] is a datetime data type in Coordinated Universal Time (UTC) format, as specified in [MS-ASDTYPE] section 2.3.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R56, the value of LastModifiedDate is:{0}", lastModifiedDate);

                    // Verify MS-ASDOC requirement: MS-ASDOC_R56
                    Site.CaptureRequirementIfIsNotNull(
                        lastModifiedDate,
                        56,
                        @"[In LastModifiedDate] The value of this element[LastModifiedDate] is a datetime data type in UTC format, as specified in [MS-ASDTYPE] section 2.3.");

                    if (null != contentLength)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R31");

                        // If the schema is verified successfully, this requirement can be covered.
                        Site.CaptureRequirement(
                            31,
                            @"[In ContentLength] The value of this element[ContentLength] is an integer data type, as specified in[MS-ASDTYPE] section 2.6.");
                    }

                    if (null != contentType)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R35");

                        // If the schema is verified successfully, this requirement can be covered.
                        Site.CaptureRequirement(
                            35,
                            @"[In ContentType] The value of this element[ContentType] is a string data type, as specified in [MS-ASDTYPE] section 2.7.");
                    }
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R12");

                // If the schema is verified successfully, this requirement can be covered.
                Site.CaptureRequirement(
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
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R87");

                // ActiveSyncClient encoded integer data as inline strings, so if response is successfully returned this requirement can be covered.
                // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R87
                Site.CaptureRequirement(
                    "MS-ASDTYPE",
                    87,
                    @"[In integer Data Type] Elements with an integer data type MUST be encoded and transmitted as WBXML inline strings, as specified in [WBXML1.2].");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R88");

                // If the schema is verified successfully, this requirement can be covered.
                Site.CaptureRequirement(
                    "MS-ASDTYPE",
                    88,
                    @"[In string Data Type] A string is a chunk of Unicode text.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R90");

                // If the schema is verified successfully, this requirement can be covered.
                Site.CaptureRequirement(
                    "MS-ASDTYPE",
                    90,
                    @"[In string Data Type] An element of this[string] type is declared as an element with a type attribute of ""string"".");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R91");

                // ActiveSyncClient encoded string data as inline strings, so if response is successfully returned this requirement can be covered.
                // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R91
                Site.CaptureRequirement(
                    "MS-ASDTYPE",
                    91,
                    @"[In string Data Type] Elements with a string data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R97");

                // If the schema is verified successfully, this requirement can be covered.
                Site.CaptureRequirement(
                    "MS-ASDTYPE",
                    97,
                    @"[In Byte Array] The structure is comprised of a length, which is expressed as a multi-byte integer, as specified in [WBXML1.2], followed by that many bytes of data.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R98");

                // ActiveSyncClient encoded byte array structure as opaque data, so if response is successfully returned this requirement can be covered.
                Site.CaptureRequirement(
                    "MS-ASDTYPE",
                    98,
                    @"[In Byte Array] Elements with a byte array structure MUST be encoded and transmitted as [WBXML1.2] opaque data.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R123");

                // If the schema is verified successfully, this requirement can be covered.
                Site.CaptureRequirement(
                    "MS-ASDTYPE",
                    123,
                    @"[In unsignedByte Data Type] The unsignedByte data type is an integer value between 0 and 255, inclusive.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R125");

                // If the schema is verified successfully, this requirement can be covered.
                Site.CaptureRequirement(
                    "MS-ASDTYPE",
                    125,
                    @"[In unsignedByte Data Type] Elements of this type[unsignedByte type] are declared with an element whose type attribute is set to ""unsignedByte"".");
                
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R5");

                // If the schema is verified successfully, this requirement can be covered.
                Site.CaptureRequirement(
                    5,
                    @"[In Message Syntax] The markup that is used by this protocol[MS-ASDOC] MUST be well-formed XML, as specified in [XML].");
            }
        }

        /// <summary>
        /// Verify the ItemOperations operation.
        /// </summary>
        /// <param name="response">The response of the ItemOperations command.</param>
        private void VerifyItemOperations(ItemOperationsResponse response)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "Schema should be verified successfully.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R101");

            // Verify MS-ASDOC requirement: MS-ASDOC_R101
            Site.CaptureRequirementIfIsNotNull(
                 response.ResponseData,
                101,
                @"[In ItemOperations Command Response] When a client uses an ItemOperations command request ([MS-ASCMD] section 2.2.1.10) to retrieve data from the server for one or more individual documents, as specified in section 3.1.5.1, the server responds with an ItemOperations command response.");

            foreach (Microsoft.Protocols.TestSuites.Common.Response.ItemOperationsResponseFetch fetch in response.ResponseData.Response.Fetch)
            {
                string statusString = fetch.Status;
                if (statusString == "1")
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R121");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R121  
                    Site.CaptureRequirementIfIsNotNull(
                        fetch.LinkId,
                        121,
                        @"[In LinkId] It[LinkId] is a required child element of the itemoperations:Fetch element ([MS-ASCMD] section 2.2.3.63.1) for Document class items in an ItemOperations command response ([MS-ASCMD] section 2.2.1.10).");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R104");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R104
                    // LinkId is the only Document class element that defined in ItemOperation,if schema is verified this requirement can be captured.
                    Site.CaptureRequirementIfIsNotNull(
                        fetch.LinkId,
                        104,
                        @"[In ItemOperations Command Response] The LinkId element (section 2.2.2.8) is the only Document class element returned in an ItemOperations command response.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDOC_R105");

                    // Verify MS-ASDOC requirement: MS-ASDOC_R105
                    Site.CaptureRequirementIfIsNotNull(
                        fetch.LinkId,
                        105,
                        @"[In ItemOperations Command Response] The LinkId element is transmitted as a child element of the Itemoperations:Fetch element ([MS-ASCMD] section 2.2.3.67.1).");
                }
            }
        }

        #region Capture requirements of [MS-ASWBXML] code page 19

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
                    if (19 == codepage)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R29");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R29
                        Site.CaptureRequirementIfAreEqual<string>(
                            "documentlibrary",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            29,
                            @"[In Code Pages] [This algorithm supports] [Code page] 19[that indicates] [XML namespace] DocumentLibrary");

                        this.CaptureRequirementsRelateToCodePage19(codepage, tagName, token);
                    }
                }
            }
        }

        #region Tag and token mapping captures.

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 19.
        /// </summary>
        /// <param name="codePageNumber">code page number</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified. </param>
        private void CaptureRequirementsRelateToCodePage19(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "LinkId":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R554");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R554
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            554,
                            @"[In Code Page 19: DocumentLibrary] [Tag name] LinkId [Token] 0x05");

                        break;
                    }

                case "DisplayName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R555");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R555
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x06,
                            token,
                            "MS-ASWBXML",
                            555,
                            @"[In Code Page 19: DocumentLibrary] [Tag name] DisplayName [Token] 0x06");

                        break;
                    }

                case "IsFolder":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R556");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R556
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            556,
                            @"[In Code Page 19: DocumentLibrary] [Tag name] IsFolder [Token] 0x07");

                        break;
                    }

                case "CreationDate":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R557");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R557
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x08,
                            token,
                            "MS-ASWBXML",
                            557,
                            @"[In Code Page 19: DocumentLibrary] [Tag name] CreationDate [Token] 0x08");

                        break;
                    }

                case "LastModifiedDate":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R558");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R558
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x09,
                            token,
                            "MS-ASWBXML",
                            558,
                            @"[In Code Page 19: DocumentLibrary] [Tag name] LastModifiedDate [Token] 0x09");

                        break;
                    }

                case "IsHidden":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R559");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R559
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            559,
                            @"[In Code Page 19: DocumentLibrary] [Tag name] IsHidden [Token] 0x0A");

                        break;
                    }

                case "ContentLength":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R560");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R560
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0B,
                            token,
                            "MS-ASWBXML",
                            560,
                            @"[In Code Page 19: DocumentLibrary] [Tag name] ContentLength [Token] 0x0B");

                        break;
                    }

                case "ContentType":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R561");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R561
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            561,
                            @"[In Code Page 19: DocumentLibrary] [Tag name] ContentType [Token] 0x0C");

                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in WBXML processing\r\n Code Page[{0}]:Tag Name[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        #endregion

        #endregion
    }
}