namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Server role and both role Adapter requirements capture code for MS-OXNSPI.
    /// </summary>
    public partial class MS_OXNSPIAdapter : ManagedAdapterBase, IMS_OXNSPIAdapter
    {
        #region Transport

        /// <summary>
        /// Verify the transport related requirements.
        /// </summary>
        private void VerifyTransport()
        {
            string transportSequence = Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower(System.Globalization.CultureInfo.CurrentCulture);
            switch (transportSequence)
            {
                case "ncacn_http":
                    if (Boolean.Parse(Common.GetConfigurationPropertyValue("RpchUseSsl", this.Site).ToLower(System.Globalization.CultureInfo.CurrentCulture)))
                    {
                        // If the transport sequence is RPC over HTTP and the code can reach here, it means that the implementation does support RPC over HTTP transport.
                        this.Site.CaptureRequirement(
                            1502,
                            @"[In Transport] The Exchange Server NSPI Protocol uses the following RPC protocol sequences: RPC over HTTPS.");
                    }
                    break;
                case "ncacn_ip_tcp":
                    if (Common.IsRequirementEnabled(1877, this.Site))
                    {
                        // If the transport sequence is RPC over TCP and the code can reach here, it means that the implementation does support RPC over TCP transport.
                        this.Site.CaptureRequirement(
                            1877,
                            @"[In Appendix B: Product Behavior] Implementation does support RPC over TCP. <1> Section 2.1: Exchange 2010 NSPI server using RPC over TCP.");
                    }

                    break;
                case "mapi_http":
                    if (Common.IsRequirementEnabled(2003, this.Site))
                    {
                        // If the transport sequence is MAPIHTTP and the code can reach here, it means that the implementation does support MAPIHTTP transport.
                        this.Site.CaptureRequirement(
                            2003,
                            @"[In Appendix A: Product Behavior] Implementation does support this specification [MS-OXCMAPIHTTP]. (Exchange Server 2013 Service Pack 1 (SP1) follows this behavior.)");
                    }

                    break;
                default:
                    this.Site.Debug.Fail("Transport \"{0}\" is not supported", transportSequence);
                    break;
            }

            if (transportSequence != "mapi_http")
            {
                // Verify MS-OXNSPI requirement: MS-OXNSPI_R8
                // UUID F5CC5A18-4264-101A-8C59-08002B2F8426 is a constant in the IDL. 
                // So if code can reach here, it means that the server uses the UUID.
                this.Site.CaptureRequirement(
                    8,
                    @"[In Transport] This protocol MUST use the UUID F5CC5A18-4264-101A-8C59-08002B2F8426.");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R9
                // The IDL encapsulates the package based on RPC version 56.0. 
                // So if code can reach here, it means that the server uses the version 56.0.
                this.Site.CaptureRequirement(
                    9,
                    @"[In Transport] The protocol MUST use the RPC version number 56.0.");
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R2
            // Because the underlying parser code has parsed out all numeric values in the structures correctly, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                2,
                @"[In Messages] Unless otherwise specified, all numeric values in this specification are in little-endian format.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R3
            // Because the underlying parser code has parsed out all Unicode string values in the structures correctly, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                3,
                @"[In Messages] Unless otherwise specified, all Unicode string representations are in UTF-16LE format.");
        }

        #endregion

        #region Property Tag
        /// <summary>
        /// Verify property value Types.
        /// </summary>
        /// <param name="propertyTag">The PropertyValue_r value which contains a property flag.</param>
        private void VerifyPropertyTag(uint propertyTag)
        {
            bool permittedPropValue = false;
            bool otherPropFlag = false;
            bool isPtypUnspecifiedReturned = false;

            switch (propertyTag & 0x0000ffff)
            {
                #region PtypInteger32
                case (uint)PropertyTypeValue.PtypInteger32:
                    permittedPropValue = true;

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1505
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        1505,
                        @"[In Permitted Property Type Values] PtypInteger32 defined in [MS-OXCDATA] section 2.11.1 is 0x0003 but 0x00000003 is used in the Exchange Server NSPI Protocol.");

                    this.VerifyMustReturnPtyp();

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2691
                    // // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2691,
                        @"[In Property Data Types] PtypInteger32(PT_LONG, PT_I4, int, ui4) is 
                        that 4 bytes; a 32-bit integer [MS-DTYP]: INT32 with Property Type Value 0x0003,%x03.00.");
                    break;
                #endregion

                #region PtypBoolean
                case (uint)PropertyTypeValue.PtypBoolean:
                    permittedPropValue = true;

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1506
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        1506,
                        @"[In Permitted Property Type Values] PtypBoolean defined in [MS-OXCDATA] section 2.11.1 is 0x000B but 0x0000000B is used in the Exchange Server NSPI Protocol.");

                    this.VerifyMustReturnPtyp();

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2698
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2698,
                        @"[In Property Data Types] PtypBoolean(PT_BOOLEAN. bool) is that 1 byte, restricted to 1 or 0 [MS-DTYP]: BOOLEAN with Property Type Value 0x000B, %x0B.00.");
                    break;
                #endregion

                #region PtypString8
                case (uint)PropertyTypeValue.PtypString8:
                    permittedPropValue = true;

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1507
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        1507,
                        @"[In Permitted Property Type Values] PtypString8 defined in [MS-OXCDATA] section 2.11.1 is 0x001E but 0x0000001E is used in the Exchange Server NSPI Protocol.");

                    this.VerifyMustReturnPtyp();

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2701.
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2701,
                        @"[In Property Data Types] PtypString8 is that variable size, a string of multi-byte characters in externally specified encoding with terminating null character (single 0 byte).PT_STRING8 with Property Type Value 0x001E, %z1E.00.");
                    break;
                #endregion

                #region PtypBinary
                case (uint)PropertyTypeValue.PtypBinary:
                    permittedPropValue = true;

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1508
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        1508,
                        @"[In Permitted Property Type Values] PtypBinary defined in [MS-OXCDATA] section 2.11.1 is 0x0102 but 0x00000102 is used in the Exchange Server NSPI Protocol.");

                    this.VerifyMustReturnPtyp();

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2707
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2707,
                        @"[In Property Data Types] PtypBinary 
                        (PT_BINARY) is that variable size; a COUNT field followed by that many bytes with Property Type Value 0x0102,%x02.01.");
                    break;
                #endregion

                #region PtypString
                case (uint)PropertyTypeValue.PtypString:
                    permittedPropValue = true;

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1509
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        1509,
                        @"[In Permitted Property Type Values] PtypString defined in [MS-OXCDATA] section 2.11.1 is 0x001F but 0x0000001F is used in the Exchange Server NSPI Protocol.");

                    this.VerifyMustReturnPtyp();

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2700
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2700,
                        @"[In Property Data Types] PtypString (PT_UNICODE, string) is that Variable size; a string of Unicode characters in UTF-16LE format encoding with terminating null character (0x0000). with Property Type Value  0x001F,%x1F.00.");
                    break;
                #endregion

                #region PtypTime
                case (uint)PropertyTypeValue.PtypTime:
                    permittedPropValue = true;

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1511
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        1511,
                        @"[In Permitted Property Type Values] PtypTime defined in [MS-OXCDATA] section 2.11.1 is 0x0040 but 0x00000040 is used in the Exchange Server NSPI Protocol.");

                    this.VerifyMustReturnPtyp();

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2702
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2702,
                        @"[In Property Data Types] PtypTime (PT_SYSTIME, time, datetime, datetime.tz, datetime.rfc1123, Date, time, time.tz) is that 
                        8 bytes; a 64-bit integer representing the number of 100-nanosecond intervals since January 1, 1601 [MS-DTYP]: FILETIME 
                        with Property Type Value 0x0040,%x40.00.");
                    break;
                #endregion

                #region PtypErrorCode
                case (uint)PropertyTypeValue.PtypErrorCode:
                    permittedPropValue = true;

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1512
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        1512,
                        @"[In Permitted Property Type Values] PtypErrorCode defined in [MS-OXCDATA] section 2.11.1 is 0x000A but 0x0000000A is used in the Exchange Server NSPI Protocol.");

                    this.VerifyMustReturnPtyp();

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2697
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2697,
                        @"[In Property Value Types] PtypErrorCode(PT_ERROR) is that 
                        4 bytes; a 32-bit integer encoding error information as specified in section 2.4.1.
                        with Property Type Value  0x000A,%x0A.00.");
                    break;
                #endregion

                #region PtypMultipleString8
                case (uint)PropertyTypeValue.PtypMultipleString8:
                    permittedPropValue = true;

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1515
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        1515,
                        @"[In Permitted Property Type Values] PtypMultipleString8 defined in [MS-OXCDATA] section 2.11.1 is 0x101E but 0x0000101E is used in the Exchange Server NSPI Protocol.");

                    this.VerifyMustReturnPtyp();

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2716
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2716,
                        @"[In Property Value Types] PtypMultiplestring8(PT_MV_STRING8, mv.string) is that 
                        variable size; a COUNT field followed by that many Ptypstring8 values with Property Type Value 0x101E,%x1E.10.");
                    break;
                #endregion

                #region PtypMultipleBinary
                case (uint)PropertyTypeValue.PtypMultipleBinary:
                    permittedPropValue = true;

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1516
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        1516,
                        @"[In Permitted Property Type Values] PtypMultipleBinary defined in [MS-OXCDATA] section 2.11.1 is 0x1102 but 0x00001102 is used in the Exchange Server NSPI Protocol.");

                    this.VerifyMustReturnPtyp();

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2719
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2719,
                        @"[In Property Value Types] PtypMultipleBinary(PT_MV_BINARY, mv.bin.hex) is that 
                        variable size; a COUNT field followed by that many PtypBinary values 
                        with Property Type Value 0x1102,%x02.11.");
                    break;
                #endregion

                #region PtypMultipleString
                case (uint)PropertyTypeValue.PtypMultipleString:
                    permittedPropValue = true;

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1517
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        1517,
                        @"[In Permitted Property Type Values] PtypMultipleString defined in [MS-OXCDATA] section 2.11.1 is 0x101F but 0x0000101F is used in the Exchange Server NSPI Protocol.");

                    this.VerifyMustReturnPtyp();

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2715
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        "MS-OXCDATA",
                        2715,
                        @"[In Property Value Types] PtypMultiplestring(PT_MV_UNICODE) is  that variable size; a COUNT field followed by that 
                        Ptypstring values with Property Type Value  0x101F,%x1F.10.");
                    break;
                #endregion

                #region PtypEmbeddedTable
                case (uint)PropertyTypeValue.PtypEmbeddedTable:
                    permittedPropValue = true;

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R31
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                            31,
                            @"[In Permitted Property Type Values] PtypEmbeddedTable ([MS-OXCDATA] section 2.11.1.5) 
                            property type with 0x0000000D value means Single 32-bit value, referencing an address list.");

                    this.VerifyMustReturnPtypAddition();
                    break;
                #endregion

                #region PtypNull
                case (uint)PropertyTypeValue.PtypNull:
                    permittedPropValue = true;
                    this.VerifyMustReturnPtypAddition();
                    break;
                #endregion

                #region PtypUnspecified
                case (uint)PropertyTypeValue.PtypUnspecified:
                    permittedPropValue = false;
                    isPtypUnspecifiedReturned = true;
                    break;
                #endregion

                default:
                    otherPropFlag = true;
                    break;
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R37. Property type {0} is used here.", (PropertyTypeValue)(propertyTag & 0x0000ffff));

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R37
            this.Site.CaptureRequirementIfIsFalse(
                isPtypUnspecifiedReturned,
                37,
                @"[In Permitted Property Type Values] [PtypUnspecified ([MS-OXCDATA] section 2.11.1) property type with 0x00000000 value] Servers MUST NOT specify this property type in any method's output parameters.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1656. The server under test uses property type {0} here.", (PropertyTypeValue)(propertyTag & 0x0000ffff));

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1656
            this.Site.CaptureRequirementIfIsTrue(
                permittedPropValue,
                1656,
                @"[In Permitted Property Type Values] The Exchange NSPI server under test MUST be capable of returning these property types [PtypInteger32, PtypBoolean, PtypString8, PtypBinary, PtypString, PtypTime, PtypErrorCode, PtypMultipleString8, PtypMultipleBinary and PtypMultipleString].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1658. The server under test uses property type {0} here.", (PropertyTypeValue)(propertyTag & 0x0000ffff));

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1658
            this.Site.CaptureRequirementIfIsFalse(
                otherPropFlag,
                1658,
                @"[In Permitted Property Type Values] The server under test MUST use the following property types [PtypInteger32, PtypBoolean, PtypString8, PtypBinary, PtypString, PtypTime, PtypErrorCode, PtypMultipleString8, PtypMultipleBinary and PtypMultipleString].");
        }

        /// <summary>
        /// Verify the property type related requirements.
        /// </summary>
        private void VerifyMustReturnPtyp()
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R26
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                26,
                @"[In Permitted Property Type Values] The Exchange Server NSPI Protocol uses the same numeric values [The values specified in [MS-OXCDATA] are 16-bit integers] but expressed as 32-bit integers.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R27
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                27,
                @"[In Permitted Property Type Values] The high-order 16 bits of the 32-bit representation that is used 
                by the Exchange Server NSPI Protocol are always 0x0000.");
        }

        /// <summary>
        /// Verify the additional property type related requirements.
        /// </summary>
        private void VerifyMustReturnPtypAddition()
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1657
            // This requirements will be captured every time when the server returns one of the following property types. If all the properties are captured, this requirement will be verified.
            this.Site.CaptureRequirement(
                1657,
                @"[In Permitted Property Type Values] In addition to the property types that are defined in [MS-OXCDATA], the Exchange NSPI server under test MUST be capable of returning the property types [PtypEmbeddedTable, PtypNull and PtypUnspecified] that are listed in the following table.");
        }
        #endregion

        #region Return values of the NSPI operation the server used
        /// <summary>
        /// Verify that the return values of the NSPI operation used by the server are valid.
        /// </summary>
        /// <param name="errorCode">The errorCode to be verified.</param>
        private void VerifyReturnValues(ErrorCodeValue errorCode)
        {
            bool permittedErrorCode = false;
            if (errorCode == ErrorCodeValue.AccessDenied
                || errorCode == ErrorCodeValue.AmbiguousRecipient
                || errorCode == ErrorCodeValue.ErrorsReturned
                || errorCode == ErrorCodeValue.GeneralFailure
                || errorCode == ErrorCodeValue.InvalidBookmark
                || errorCode == ErrorCodeValue.InvalidCodepage
                || errorCode == ErrorCodeValue.InvalidLocale
                || errorCode == ErrorCodeValue.InvalidObject
                || errorCode == ErrorCodeValue.InvalidParameter
                || errorCode == ErrorCodeValue.LogonFailed
                || errorCode == ErrorCodeValue.NotEnoughMemory
                || errorCode == ErrorCodeValue.NotFound
                || errorCode == ErrorCodeValue.NotSupported
                || errorCode == ErrorCodeValue.OutOfResources
                || errorCode == ErrorCodeValue.Success
                || errorCode == ErrorCodeValue.TableTooBig
                || errorCode == ErrorCodeValue.TooComplex)
            {
                permittedErrorCode = true;
            }

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                @"Verify MS-OXNSPI_R1661: The server uses {0} in the return value of the operation here.",
                errorCode);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1661
            this.Site.CaptureRequirementIfIsTrue(
                permittedErrorCode,
                1661,
                @"[In Permitted Error Code Values] The server under test MUST use the following error codes [Success, UnbindSuccess, UnbindFailure, ErrorsReturned, GeneralFailure, NotSupported, InvalidObject, OutOfResources, NotFound, LogonFailed, TooComplex, InvalidCodepage, InvalidLocale, TableTooBig, InvalidBookmark, AccessDenied, NotEnoughMemory and InvalidParameter] in the return value of the operation.");
        }
        #endregion

        #region Verify display type
        /// <summary>
        /// Verify the display type related requirements.
        /// </summary>
        /// <param name="displayTypeValue">The DisplayTypeValues to be verified.</param>
        private void VerifyDisplayType(DisplayTypeValue displayTypeValue)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R59");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R59
            this.Site.CaptureRequirementIfAreNotEqual(
                DisplayTypeValue.DT_TEMPLATE,
                displayTypeValue,
                59,
                @"[In Display Type Values] [DT_TEMPLATE display type] An Exchange NSPI server MUST NOT return this display type.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R63");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R63
            this.Site.CaptureRequirementIfAreNotEqual(
                DisplayTypeValue.DT_SEARCH,
                displayTypeValue,
                63,
                @"[In Display Type Values] [DT_SEARCH display type] An Exchange NSPI server MUST NOT return this display type.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R53001");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R53001
            this.Site.CaptureRequirementIfAreNotEqual(
                DisplayTypeValue.DT_ORGANIZATION,
                displayTypeValue,
                53001,
                @"[In Display Type Values] [DT_ORGANIZATION] An Exchange NSPI server MUST NOT return display type.");

            bool validDisplayType = false;
            if (displayTypeValue == DisplayTypeValue.DT_ADDRESS_TEMPLATE
                || displayTypeValue == DisplayTypeValue.DT_AGENT
                || displayTypeValue == DisplayTypeValue.DT_CONTAINER
                || displayTypeValue == DisplayTypeValue.DT_DISTLIST
                || displayTypeValue == DisplayTypeValue.DT_FORUM
                || displayTypeValue == DisplayTypeValue.DT_MAILUSER
                || displayTypeValue == DisplayTypeValue.DT_ORGANIZATION
                || displayTypeValue == DisplayTypeValue.DT_PRIVATE_DISTLIST
                || displayTypeValue == DisplayTypeValue.DT_REMOTE_MAILUSER
                || displayTypeValue == DisplayTypeValue.DT_SEARCH
                || displayTypeValue == DisplayTypeValue.DT_TEMPLATE)
            {
                validDisplayType = true;
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1663: The server under test use the display type {0}.", displayTypeValue);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1663
            this.Site.CaptureRequirementIfIsTrue(
                validDisplayType,
                1663,
                @"[In Display Type Values] The server under test MUST use the following display types [DT_MAILUSER, DT_DISTLIST, DT_FORUM, DT_AGENT, DT_ORGANIZATION, DT_PRIVATE_DISTLIST, DT_REMOTE_MAILUSER, DT_CONTAINER, DT_TEMPLATE, DT_ADDRESS_TEMPLATE and DT_SEARCH].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1662: The server under test use the display type {0}.", displayTypeValue);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1662
            this.Site.CaptureRequirementIfIsTrue(
                validDisplayType,
                1662,
                @"[In Display Type Values] Except where otherwise specified in the following table, the Exchange NSPI server under test MUST be capable of returning these display types [DT_MAILUSER, DT_DISTLIST, DT_FORUM, DT_AGENT, DT_ORGANIZATION, DT_PRIVATE_DISTLIST, DT_REMOTE_MAILUSER, DT_CONTAINER, DT_TEMPLATE, DT_ADDRESS_TEMPLATE, DT_SEARCH].");
        }
        #endregion

        #region Structures

        /// <summary>
        /// Property values' structure verification.
        /// </summary>
        private void VerifyPropertyStructures()
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R134
            // This is a general description. This requirement will be captured every time when the detail requirement is captured.
            this.Site.CaptureRequirement(
                134,
                @"[In Property Values] The following structures [FlatUID_r, PropertyTagArray_r, Binary_r, ShortArray_r, LongArray_r, StringArray_r, BinaryArray_r, FlatUIDArray_r, WStringArray_r, DateTimeArray_r, PROP_VAL_UNION and PropertyValue_r] are used to represent specific property values.");
        }

        #region FlatUID_r
        /// <summary>
        /// Verify FlatUID_r structure.
        /// </summary>
        private void VerifyFlatUID_r()
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1562
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1562,
                @"[In FlatUID_r Structure] [The type is defined as following:]  typedef struct { BYTE ab[16];} FlatUID_r");

            // This test suite defines FlatUID_r based on the meaning of FlatUID data structure. So if the codes can reach here, it means that the semantic meaning is unchanged.
            this.Site.CaptureRequirement(
                136,
                @"[In FlatUID_r Structure] The semantic meaning is unchanged from the FlatUID data structure.");

            // Verify MS-OXNSPI requirement: MS-OXCDATA_R6
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                6,
                @"[In FlatUID Structure] A FlatUID structure is a GUID structure put into little-endian byte order.");

            // Verify MS-OXNSPI requirement: MS-OXCDATA_R9
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                9,
                @"[In FlatUID Structure] FlatUID (16 bytes):  A flat little-endian sequence used as a unique identifier in various structures.");

            // Verify MS-OXNSPI requirement: MS-OXCDATA_R10
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                10,
                @"[In FlatUID_r Structure] FlatUID (16 bytes): A flat little-endian sequence used as a unique identifier in various structures.");
        }

        #endregion

        #region PropertyTagArray_r
        /// <summary>
        /// Verify PropertyTagArray_r structure.
        /// </summary>
        private void VerifyPropertyTagArray_r()
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1563.
            // Because the underlying parser code has parsed out the structure which contains it, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1563,
                @"[In PropertyTagArray_r Structure] [The type is defined as following:]  typedef struct PropertyTagArray_r {
                    DWORD cValues;
                     size_is(cValues+1), length_is(cValues)] 
                        DWORD aulPropTag[];
                    } PropertyTagArray_r;");

            // This test suite defines PropertyTagArray_r structure based on the meaning of PropertyTagArray_r data structure. So if the codes can reach here, it means that the semantic meaning is unchanged.
            this.Site.CaptureRequirement(
                140,
                @"[In PropertyTagArray_r Structure] The semantic meaning is otherwise unchanged from the PropertyTagArray_r data structure.");

            this.VerifyPropertyStructures();

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R187
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                187,
                @"[In PropertyTagArray Structure] Count (2 bytes): An unsigned integer, specifying the number of property tags to follow.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R189.
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                189,
                @"[In PropertyTagArray_r Structure] Count (2 bytes): Encodes the Count field in the PropTagArray structure, as specified in section 2.10.1.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R191.
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                191,
                @"[In PropertyTagArray_r Structure] PropertyTags (variable): Encodes the PropertyTags field of the PropTagArray structure.");
        }
        #endregion

        #region Binary_r
        /// <summary>
        /// Verify Binary_r structure.
        /// </summary>
        /// <param name="binary">A Binary_r value to be verified.</param>
        private void VerifyBinary_r(Binary_r binary)
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R144
            // The parser method parses code according to this definition, so MS-OXNSPI_R144 can be captured directly.
            this.Site.CaptureRequirement(
                144,
                @"[In Binary_r Structure] The Binary_r structure encodes an array of uninterpreted bytes.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R147
            // The parser method parses code according to this definition, so MS-OXNSPI_R147 can be captured directly.
            this.Site.CaptureRequirement(
                147,
                @"[In Binary_r Structure] lpb: The uninterpreted bytes.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R145");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R145
            this.Site.CaptureRequirementIfAreEqual<int>(
                binary.Lpb.Length,
                (int)binary.Cb,
                145,
                @"[In Binary_r Structure] cb: The number of uninterpreted bytes represented in this structure.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1564
            // Because the underlying parser code has parsed out the structure which contains it, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1564,
                @"[In Binary_r Structure] [The type is defined as following:]  typedef struct Binary_r {
                    [range(0,2097152)] DWORD cb;
                    [size_is(cb)] BYTE* lpb;
                } Binary_r;");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R146: The value of cb in Binary_r Structure is {0} here.", binary.Cb);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R146
            this.Site.CaptureRequirementIfIsTrue(
                binary.Cb <= 2097152,
                146,
                @"[In Binary_r Structure] [cb] This value MUST NOT exceed 2,097,152.");

            this.VerifyPropertyStructures();
        }

        #endregion

        #region stringArray_r
        /// <summary>
        /// Verify StringArray_r structure.
        /// </summary>
        /// <param name="stringArray">A StringArray_r value to be verified.</param>
        private void VerifystringArray_r(StringArray_r stringArray)
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1567
            // Because the underlying parser code has parsed out the structure which contains it, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1567,
                @"[In stringArray_r Structure] [The type is defined as following:] typedef struct _stringArray_r {
                    [range(0,100000)] DWORD cValues;
                    [string, size_is(cValues)] char** lppszA;
                } stringArray_r;");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R158: The value of cValues in StringArray_r Structure is {0} here.",
                stringArray.CValues);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R158
            this.Site.CaptureRequirementIfIsTrue(
                stringArray.CValues <= 100000,
                158,
                @"[In StringArray_r Structure] [cValues] This value MUST NOT exceed 100,000.");

            this.VerifyPropertyStructures();

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R156
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                156,
                @"[In StringArray_r Structure] The StringArray_r structure encodes an array of references to 8-bit character strings.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R157
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                157,
                @"[In StringArray_r Structure] cValues: The number of 8-bit character string references represented in the StringArray_r structure.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R159
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                159,
                @"[In StringArray_r Structure] lppszA: The 8-bit character string references.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R160
            // The parser code parses the strings based on the string that is ended with null. If all the codes can reach here, it illustrates that it is parsed correctly.
            this.Site.CaptureRequirement(
                160,
                @"[In StringArray_r Structure] [lppszA] The strings referred to are NULL-terminated.");
        }

        #endregion

        #region BinaryArray_r
        /// <summary>
        /// Verify BinaryArray_r structure.
        /// </summary>
        /// <param name="binaryArray">A BinaryArray_r value to be verified.</param>
        private void VerifyBinaryArray_r(BinaryArray_r binaryArray)
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1568
            // Because the underlying parser code has parsed out the structure which contains it, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1568,
                @"[In BinaryArray_r Structure] [The type is defined as following:] typedef struct _BinaryArray_r {
                    [range(0,100000)] DWORD cValues;
                    [size_is(cValues)] Binary_r* lpbin;
                } BinaryArray_r;");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R163: the cValues in BinaryArray_r Structure is {0}.", binaryArray.CValues);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R163
            this.Site.CaptureRequirementIfIsTrue(
                binaryArray.CValues <= 100000,
                163,
                @"[In BinaryArray_r Structure] [cValues] This value MUST NOT exceed 100,000.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R161
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                161,
                @"[In BinaryArray_r Structure] The BinaryArray_r structure is an array of Binary_r data structures.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R162");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R162
            this.Site.CaptureRequirementIfAreEqual<int>(
                binaryArray.Lpbin.Length,
                (int)binaryArray.CValues,
                162,
                @"[In BinaryArray_r Structure] cValues: The number of Binary_r data structures represented in the BinaryArray_r structure.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R164
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                164,
                @"[In BinaryArray_r Structure] lpbin: The Binary_r data structures.");

            this.VerifyPropertyStructures();

            foreach (Binary_r binary in binaryArray.Lpbin)
            {
                this.VerifyBinary_r(binary);
            }
        }

        #endregion

        #region WstringArray_r
        /// <summary>
        /// Verify WStringArray_r structure.
        /// </summary>
        /// <param name="wstringArray">A WStringArray_r value to be verified.</param>
        private void VerifyWstringArray_r(WStringArray_r wstringArray)
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1570
            // Because the underlying parser code has parsed out the structure which contains it, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1570,
                @"[In WstringArray_r Structure] [The type is defined as following:] typedef struct _WstringArray_r {
                    [range(0,100000)] DWORD cValues;
                    [string, size_is(cValues)] wchar_t** lppszW;
                } WstringArray_r;");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R171: the cValues in WStringArray_r Structure is {0}.", wstringArray.CValues);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R171
            this.Site.CaptureRequirementIfIsTrue(
                wstringArray.CValues <= 100000,
                171,
                @"[In WStringArray_r Structure] [cValues]This value MUST NOT exceed 100,000.");

            this.VerifyPropertyStructures();

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R169
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                169,
                @"[In WStringArray_r Structure] The WStringArray_r structure encodes an array of references to Unicode strings.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R170");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R170
            this.Site.CaptureRequirementIfAreEqual<int>(
                wstringArray.LppszW.Length,
                (int)wstringArray.CValues,
                170,
                @"[In WStringArray_r Structure] cValues: The number of Unicode character string references represented in the WStringArray_r structure.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R172
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                172,
                @"[In WStringArray_r Structure] lppszW: The Unicode character string references.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R173
            // The parser code parses the strings based on the string that is ended with null. If all the codes can reach here, it illustrates that it is parsed correctly.
            this.Site.CaptureRequirement(
                173,
                @"[In WStringArray_r Structure] [ppszW] The strings referred to are NULL-terminated.");
        }

        #endregion

        #region PROP_VAL_UNION
        /// <summary>
        /// Verify PropertyValue_r structure.
        /// </summary>
        /// <param name="propertyValue">A PropertyValue_r value to be verified.</param>
        private void VerifyPROP_VAL_UNION(PropertyValue_r propertyValue)
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1572
            // Because the underlying parser code has parsed out the structure which contains it, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1572,
                @"[In PROP_VAL_UNION Structure] [The type is defined as following:] typedef 
                    [switch_type(long)] 
                    union _PV_r {
                    [case(0x00000002)] 
                    short int i;
                    [case(0x00000003)] 
                    long l;
                    [case(0x0000000B)] 
                    unsigned short int b;
                    [case(0x0000001E)] 
                    [string] char* lpszA;
                    [case(0x00000102)] 
                    Binary_r bin;
                    [case(0x0000001F)] 
                    [string] wchar_t* lpszW;
                    [case(0x00000048)] 
                    FlatUID_r* lpguid;
                    [case(0x00000040)] 
                    FILETIME ft;
                    [case(0x0000000A)] 
                    long err;
                    [case(0x00001002)] 
                    ShortArray_r MVi;
                    [case(0x00001003)] 
                    LongArray_r MVl;
                    [case(0x0000101E)] 
                    stringArray_r MVszA;
                    [case(0x00001102)] 
                    BinaryArray_r MVbin;
                    [case(0x00001048)] 
                    FlatUIDArray_r MVguid;
                    [case(0x0000101F)] 
                    WstringArray_r MVszW;
                    [case(0x00001040)] 
                    DateTimeArray_r MVft;
                    [case(0x00000001, 0x0000000D)] 
                    long lReserved;
                } PROP_VAL_UNION;");

            this.VerifyPropertyStructures();

            switch (propertyValue.PropTag & 0x0000FFFF)
            {
                case 0x00000003:

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R181
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        181,
                        @"[In PROP_VAL_UNION Structure] l: PROP_VAL_UNION contains an encoding of the value of a property that can contain a single 32-bit integer value.");
                    break;

                case 0x0000000B:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R183: b: PROP_VAL_UNION, the value is {0}.", propertyValue.Value.B);

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R183
                    this.Site.CaptureRequirementIfIsTrue(
                          (propertyValue.Value.B == 1) || (propertyValue.Value.B == 0),
                          183,
                          @"[In PROP_VAL_UNION Structure] [b] The client and server MUST NOT set this to values other than 1 or 0.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R182. b: PROP_VAL_UNION, the value is {0}.", propertyValue.Value.B);

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R182
                    this.Site.CaptureRequirementIfIsTrue(
                        (propertyValue.Value.B == 1) || (propertyValue.Value.B == 0),
                        182,
                        @"[In PROP_VAL_UNION Structure] b: PROP_VAL_UNION contains an encoding of the value of a property that can contain a single Boolean value.");

                    break;
                case 0x0000001E:

                    if (propertyValue.Value.LpszA != null)
                    {
                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R184
                        // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                        this.Site.CaptureRequirement(
                            184,
                            @"[In PROP_VAL_UNION Structure] lpszA: PROP_VAL_UNION contains an encoding of the value of a property that can contain a single 8-bit character string value.");

                        // The parser code parses the strings based on the string that is ended with null If all the codes can reach here, it illustrates that it is parsed correctly.
                        this.Site.CaptureRequirement(
                            185,
                            @"[In PROP_VAL_UNION Structure] [lpszA] This value is NULL-terminated.");
                    }

                    break;

                case 0x00000102:

                    if (propertyValue.Value.Bin.Lpb != null)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R187: the current number of bytes that be encoded is {0}.", propertyValue.Value.Bin.Lpb.Length);

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R187
                        this.Site.CaptureRequirementIfIsTrue(
                            propertyValue.Value.Bin.Lpb.Length <= 2097152,
                            187,
                            @"[In PROP_VAL_UNION Structure] [bin] The number of bytes that can be encoded in this structure is 2,097,152.");

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R186
                        // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                        this.Site.CaptureRequirement(
                            186,
                            @"[In PROP_VAL_UNION Structure] bin: PROP_VAL_UNION contains an encoding of the value of a property that can contain a single binary data value.");

                        this.VerifyBinary_r(propertyValue.Value.Bin);
                    }

                    break;

                case 0x0000001F:

                    if (propertyValue.Value.LpszW != null && propertyValue.Value.LpszW.Length != 0)
                    {
                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R188
                        // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                        this.Site.CaptureRequirement(
                            188,
                            @"[In PROP_VAL_UNION Structure] lpszW: PROP_VAL_UNION contains an encoding of the value of a property that can contain a single Unicode string value.");

                        // The parser code parses the strings based on the string that is ended with null If all the codes can reach here, it illustrates that it is parsed correctly.
                        this.Site.CaptureRequirement(
                            189,
                            @"[In PROP_VAL_UNION Structure] [lpszW] This value is NULL-terminated.");
                    }

                    break;

                case 0x00000048:
                    this.VerifyFlatUID_r();
                    break;

                case 0x00000040:
                    if (propertyValue.Value.Ft.ToString() != null)
                    {
                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R193
                        // Because the underlying parser code has parsed out the structure which contains it, this requirement can be captured directly.
                        this.Site.CaptureRequirement(
                            193,
                            @"[In PROP_VAL_UNION Structure] [ft] The value is encoded as a FILETIME structure.");

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R192
                        // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                        this.Site.CaptureRequirement(
                            192,
                            @"[In PROP_VAL_UNION Structure] ft: PROP_VAL_UNION contains an encoding of the value of a property that can contain a single 64-bit integer value.");
                    }

                    break;

                case 0x0000000A:

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R194
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        194,
                        @"[In PROP_VAL_UNION Structure] err: PROP_VAL_UNION contains an encoding of the value of a property that can contain a single PtypErrorCode value.");

                    break;

                case 0x0000101E:

                    if (propertyValue.Value.MVszA.LppszA != null)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R201: the current number of values that be encoded is {0}.", propertyValue.Value.MVszA.LppszA.Length);

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R201
                        this.Site.CaptureRequirementIfIsTrue(
                            propertyValue.Value.MVszA.LppszA.Length <= 100000,
                            201,
                            @"[In PROP_VAL_UNION Structure] [MVszA] The number of values that can be encoded in this structure is 100,000.");

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R199
                        // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                        this.Site.CaptureRequirement(
                            199,
                            @"[In PROP_VAL_UNION Structure] MVszA: PROP_VAL_UNION contains an encoding of the values of a property that can contain multiple 8-bit character string values.");

                        // The parser code parses the strings based on the string that is ended with null. If all the codes can reach here, it illustrates that it is parsed correctly.
                        this.Site.CaptureRequirement(
                            200,
                            @"[In PROP_VAL_UNION Structure] [MVszA] These string values [values of a property] are NULL-terminated.");

                        this.VerifystringArray_r(propertyValue.Value.MVszA);
                    }

                    break;

                case 0x00001102:

                    if (propertyValue.Value.MVbin.Lpbin != null)
                    {
                        for (int i = 0; i < propertyValue.Value.MVbin.Lpbin.Length; i++)
                        {
                            if (propertyValue.Value.MVbin.Lpbin.Length > 2097152)
                            {
                                this.Site.Assert.Fail("The number of bytes that be encoded is {0}, which is large than 2,097,152.", propertyValue.Value.MVbin.Lpbin.Length);
                            }
                        }

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R203
                        this.Site.CaptureRequirement(
                            203,
                            @"[In PROP_VAL_UNION Structure] [MVbin] The number of bytes that can be encoded in each value of this structure is 2,097,152.");

                        // Add the debug information.
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R204: [In PROP_VAL_UNION Structure] [MVbin] The number of values that be encoded in structure PROP_VAL_UNION is {0}", propertyValue.Value.MVbin.CValues);

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R204
                        this.Site.CaptureRequirementIfIsTrue(
                            propertyValue.Value.MVbin.CValues <= 100000,
                            204,
                            @"[In PROP_VAL_UNION Structure] [MVbin] The number of values that can be encoded in this structure is 100,000.");

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R202
                        // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                        this.Site.CaptureRequirement(
                            202,
                            @"[In PROP_VAL_UNION Structure] MVbin: PROP_VAL_UNION contains an encoding of the values of a property that can contain multiple binary data values.");

                        this.VerifyBinaryArray_r(propertyValue.Value.MVbin);
                    }

                    break;

                case 0x0000101F:
                    if (propertyValue.Value.MVszW.LppszW != null)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R210: [In PROP_VAL_UNION Structure] [MVszW] The number of values that be encoded in this structure is {0}", propertyValue.Value.MVszW.CValues);

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R210
                        this.Site.CaptureRequirementIfIsTrue(
                            propertyValue.Value.MVszW.CValues <= 100000,
                            210,
                            @"[In PROP_VAL_UNION Structure] [MVszW] The number of values that can be encoded in this structure is 100,000.");

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R208
                        // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                        this.Site.CaptureRequirement(
                            208,
                            @"[In PROP_VAL_UNION Structure] MVszW: PROP_VAL_UNION contains an encoding of the values of a property that can contain multiple Unicode string values.");

                        // The parser code parses the strings based on the string that is ended with null. If all the codes can reach here, it illustrates that it is parsed correctly.
                        this.Site.CaptureRequirement(
                            209,
                            @"[In PROP_VAL_UNION Structure] [MVszW] These string values [values of a property] are NULL-terminated.");

                        this.VerifyWstringArray_r(propertyValue.Value.MVszW);
                    }

                    break;

                // As described in [MS-OXNSPI], these two cases have the same behavior.
                case 0x00000001:
                case 0x0000000D:

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1707");

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1707
                    this.Site.CaptureRequirementIfAreEqual<int>(
                        0x0,
                        propertyValue.Value.Reserved,
                        1707,
                        @"[In PROP_VAL_UNION Structure] The server under test MUST set this value[lReserved] to the constant 0x00000000.");
                    break;

                default:
                    this.Site.Log.Add(LogEntryKind.Debug, "Property tag {0} is returned by the server but is not covered by the current test suite.", propertyValue.PropTag & 0x0000FFFF);
                    break;
            }
        }

        #endregion

        #region PropertyValue_r
        /// <summary>
        /// Verify PropertyValue_r structure.
        /// </summary>
        /// <param name="propertyValue">A PropertyValue_r value to be verified.</param>
        private void VerifyPropertyValueStructure(PropertyValue_r propertyValue)
        {
            this.VerifyPropertyTag(propertyValue.PropTag);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1573
            // Because the underlying parser code has parsed out the structure which contains it, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1573,
                @"[In PropertyValue_r Structure] [The PropertyValue_r is defined as following:] typedef struct _PropertyValue_r {
                    DWORD ulPropTag;
                    DWORD ulReserved;
                    [switch_is((long)(ulPropTag & 0x0000FFFF))] 
                        PROP_VAL_UNION Value;
                } PropertyValue_r;");

            // This test suite defines PropertyValue_r structure based on the meaning of PropertyValue data structure. So if the codes can reach here, it means that the semantic meaning is unchanged.
            this.Site.CaptureRequirement(
                219,
                @"[In PropertyValue_r Structure] The semantic meaning is otherwise unchanged from the PropertyValue data structure.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1708");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1708
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x0,
                propertyValue.Reserved,
                1708,
                @"[In PropertyValue_r Structure] The server under test MUST set this value[ulReserved] to the constant 0x00000000.");

            this.VerifyPropertyStructures();

            if (propertyValue.PropTag == (uint)AulProp.PidTagEntryId)
            {
                if (propertyValue.Value.Bin.Lpb[0] == 0x00)
                {
                    // 0x00 means a PermanentEntryID.
                    PermanentEntryID permanentEntryID = AdapterHelper.ParsePermanentEntryIDFromBytes(propertyValue.Value.Bin.Lpb);
                    this.VerifyPermanentEntryID(permanentEntryID);
                    this.VerifyDisplayType(permanentEntryID.DisplayTypeString);
                }
                else if (propertyValue.Value.Bin.Lpb[0] == 0x87)
                {
                    // 0x87 means an EphemeralEntryID.
                    EphemeralEntryID ephemeralEntryID = AdapterHelper.ParseEphemeralEntryIDFromBytes(propertyValue.Value.Bin.Lpb);
                    this.VerifyEphemeralEntryID(ephemeralEntryID);
                    this.VerifyDisplayType(ephemeralEntryID.DisplayType);
                }
            }

            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                224,
                @"[In PropertyValue_r Structure] [Value] The type value held is specified by the property type of the proptag in the ulPropTag field.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R456
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                456,
                @"[In PropertyValue_r Structure] PropertyTag (4 bytes): This value encodes the property tag with the value represented by the PropertyValue_r structure.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R457
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                457,
                @"[In PropertyValue_r Structure] Reserved (4 bytes):  All clients and servers MUST set this value to 0x00000000.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R458
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                458,
                @"[In PropertyValue_r Structure] PropertyValue (variable): This value encodes the PropertyValue field of the PropertyValue structure, as specified in section 2.11.2.1.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R75
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                75,
                @"[In StandardPropertyRow Structure] ValueArray (variable): At each position of the array, the structure will either be a PropertyValue structure, as specified in section 2.11.2.1, if the type of the corresponding property tag was specified, or a TypedPropertyValue structure, as specified in section 2.11.3, if the type of the corresponding property tag was PtypUnspecified (section 2.11.1).");

            this.VerifyPROP_VAL_UNION(propertyValue);
        }

        #endregion

        #region PropertyRow_r
        /// <summary>
        /// Verify PropertyRow_r structure.
        /// </summary>
        /// <param name="propertyRow">A PropertyRow_r value to be verified.</param>
        private void VerifyPropertyRowStructure(PropertyRow_r propertyRow)
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1574
            // Because the underlying parser code has parsed out the structure which contains it, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1574,
                @"[In PropertyRow_r Structure] [The type is defined as following:] typedef struct _PropertyRow_r {
                    DWORD Reserved;
                    [range(0,100000)] DWORD cValues;
                    [size_is(cValues)] PropertyValue_r* lpProps;
                } PropertyRow_r;");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1709");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1709
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x0,
                propertyRow.Reserved,
                1709,
                @"[In PropertyRow_r Structure] [Reserved] The server under test MUST set this value to the constant 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R72");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R72
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                72,
                @"[In StandardPropertyRow Structure] Flag (1 byte): An unsigned integer.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R73");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R73
            // Consider the Flag in StandardPropertyRow structure has the same value with propertyRow.Reserved
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x0,
                propertyRow.Reserved,
                "MS-OXCDATA",
                73,
                @"[In StandardPropertyRow Structure] Flag (1 byte): This value MUST be set to 0x00 to indicate that all property values are present and without error.");

            // Check whether cValues represents the number of PropertyValue_r structures.
            bool iscValuesRepresented = false;
            if (propertyRow.LpProps == null)
            {
                if (propertyRow.CValues == 0)
                {
                    iscValuesRepresented = true;
                }
                else
                {
                    iscValuesRepresented = false;
                }
            }
            else
            {
                if (propertyRow.LpProps.Length == (int)propertyRow.CValues)
                {
                    iscValuesRepresented = true;
                }
                else
                {
                    iscValuesRepresented = false;
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R229. cValues is {0}", propertyRow.CValues);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R229
            this.Site.CaptureRequirementIfIsTrue(
                iscValuesRepresented,
                229,
                @"[In PropertyRow_r Structure] cValues: The number of PropertyValue_r structures represented in the PropertyRow_r structure.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R230: [In PropertyRow_r Structure] [cValues] This value is {0}.", propertyRow.CValues);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R230
            this.Site.CaptureRequirementIfIsTrue(
                propertyRow.CValues <= 100000,
                230,
                @"[In PropertyRow_r Structure] [cValues] This value MUST NOT exceed 100,000.");

            if (propertyRow.LpProps != null)
            {
                foreach (PropertyValue_r propertyValue in propertyRow.LpProps)
                {
                    this.VerifyPropertyValueStructure(propertyValue);
                }

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R231");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R231
                // The PropertyValue_r should be the encoding of the ValueArray field of the StandardPropertyRow data structure
                this.Site.CaptureRequirement(
                    231,
                    @"[In PropertyRow_r Structure] lpProps: Encodes the ValueArray field of the StandardPropertyRow data structure.");
            }

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R82
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                82,
                @"[In PropertyRow_r Structure] Reserved (4 bytes): Servers MUST set this value to 0x00000000.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R83
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                83,
                @"[In PropertyRow_r Structure] ValueCount (4 bytes):  The number of property values represented in the ValueArray field.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R84
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                84,
                @"[In PropertyRow_r Structure] ValueCount (4 bytes):This value[ValueCount (4 bytes)] MUST NOT exceed 100,000.");
        }
        #endregion

        #region PropertyRowSet_r
        /// <summary>
        /// Verify PropertyRowSet_r structure.
        /// </summary>
        /// <param name="propertyRowSet">A PropertyRowSet_r value to be verified.</param>
        private void VerifyPropertyRowSetStructure(PropertyRowSet_r propertyRowSet)
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1575
            // Because the underlying parser code has parsed out the structure which contains it, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1575,
                @"[In PropertyRowSet_r Structure] [The type is defined as following:] typedef struct _PropertyRowSet_r {
                [range(0,100000)] DWORD cRows;
                [size_is(cRows)] PropertyRow_r aRow[];
            } PropertyRowSet_r;");

            // This test suite defines PropertyRowSet_r structure based on the meaning of PropertyRowSet data structure. So if the codes can reach here, it means that the semantic meaning is unchanged.
            this.Site.CaptureRequirement(
                234,
                @"[In PropertyRowSet_r Structure] The semantic meaning is otherwise unchanged from the PropertyRowSet data structure.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R236: [In PropertyRowSet_r Structure] [cRows] This value is {0}", propertyRowSet.CRows);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R236
            this.Site.CaptureRequirementIfIsTrue(
                propertyRowSet.CRows <= 100000,
                236,
                @"[In PropertyRowSet_r Structure] [cRows] This value MUST NOT exceed 100,000.");

            if (propertyRowSet.ARow != null)
            {
                foreach (PropertyRow_r propertyRow in propertyRowSet.ARow)
                {
                    if (propertyRow.CValues != 0)
                    {
                        this.VerifyPropertyRowStructure(propertyRow);
                    }
                }
            }

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R95
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                95,
                @"[In PropertyRowSet Structure] RowCount (2 bytes): RowCount (2 bytes): An unsigned integer specifying the number of PropertyRow structures in the Rows field.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R97.
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                97,
                @"[In PropertyRowSet_r Structure] RowCount (2 bytes):This value encodes the RowCount field of the PropertyRowSet structure, as specified in section 2.8.2.1.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R99.
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                "MS-OXCDATA",
                99,
                @"[In PropertyRowSet_r Structure] Rows (variable): This value encodes the rows field of the PropertyRowSet structure.");
        }

        #endregion

        #region STAT
        /// <summary>
        /// Verify The STAT structure.
        /// </summary>
        /// <param name="returnValue">The return value of the operation that uses the STAT structure.</param>
        /// <param name="inputStat">The inputStat parameter of the operation that uses the STAT structure.</param>
        /// <param name="outputStat">The outputStat parameter of the operation that uses the STAT structure.</param>
        /// <param name="isNspiGetMatches">Indicates whether the method is NspiGetMatches.</param>
        private void VerifySTATStructure(ErrorCodeValue returnValue, STAT inputStat, STAT outputStat, bool isNspiGetMatches)
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1587
            // Because the underlying parser code has parsed out the structure which contains it, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1587,
                @"[In STAT] [The type is defined as following:] typedef struct {
                    DWORD SortType;
                    DWORD ContainerID;
                    DWORD CurrentRec;
                    long Delta;
                    DWORD NumPos;
                    DWORD TotalRecs;
                    DWORD CodePage;
                    DWORD TemplateLocale;
                    DWORD SortLocale;
                } STAT;");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R314");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R314
            this.Site.CaptureRequirementIfAreEqual<uint>(
                inputStat.SortType,
                outputStat.SortType,
                314,
                @"[In STAT] [SortType] The server MUST NOT modify this field.");

            if (returnValue == ErrorCodeValue.Success)
            {
                // According to the description in section 2.3.7 and 3.1.4.1.10 of the MS-OXNSPI specification, there are several exceptions when returning the STAT structure for the NspiGetMatches method.
                if (!isNspiGetMatches)
                {
                    bool isSortTypeSetCorrectValue = outputStat.SortType == (uint)TableSortOrder.SortTypeDisplayName ||
                                                     outputStat.SortType == (uint)TableSortOrder.SortTypePhoneticDisplayName ||
                                                     outputStat.SortType == (uint)TableSortOrder.SortTypeDisplayName_RO ||
                                                     outputStat.SortType == (uint)TableSortOrder.SortTypeDisplayName_W;

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R92, the sort type of state is {0}", outputStat.SortType);

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R92
                    this.Site.CaptureRequirementIfIsTrue(
                        isSortTypeSetCorrectValue,
                        92,
                        @"[In Table Sort Orders] These values [SortTypeDisplayName, SortTypePhoneticDisplayName, SortTypeDisplayName_RO and SortTypeDisplayName_W] appear in the SortType field of the STAT data structure, as specified in section 2.2.8.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R311");

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R311
                    this.Site.CaptureRequirementIfIsTrue(
                        isSortTypeSetCorrectValue,
                        311,
                        @"[In STAT] SortType: This field contains a DWORD [MS-DTYP] value that represents a sort order.");

                    // According to the description in section 2.2.8 and 2.3.8.1, if the returned ContainerID field is set to MID_BEGINNING_OF_TABLE or 
                    // MID_END_OF_TABLE, or its value is greater than or equal to 0x10, it is a valid Minimal Entry ID.
                    bool isContainerIDSetCorrectValue = outputStat.ContainerID >= 0x10
                                                     || outputStat.ContainerID == (uint)MinimalEntryID.MID_BEGINNING_OF_TABLE
                                                     || outputStat.ContainerID == (uint)MinimalEntryID.MID_END_OF_TABLE;

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R315. The value of ContainerID is {0}.", outputStat.ContainerID);

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R315
                    this.Site.CaptureRequirementIfIsTrue(
                        isContainerIDSetCorrectValue,
                        315,
                        @"[In STAT] ContainerID: This field contains a Minimal Entry ID.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R324");

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R324
                    this.Site.CaptureRequirementIfAreEqual<int>(
                        0,
                        outputStat.Delta,
                        324,
                        @"[In STAT] [Delta] If the NSPI method returns a success value, the server MUST set this field to 0.");
                }

                // According to the description in section 2.2.8 and 2.3.8.1, if the returned CurrentRec field is set to MID_BEGINNING_OF_TABLE or 
                // MID_END_OF_TABLE, or its value is greater than or equal to 0x10, it is a valid Minimal Entry ID.
                bool isCurrentRecSetCorrectValue = outputStat.CurrentRec >= 0x10
                                                || outputStat.CurrentRec == (uint)MinimalEntryID.MID_BEGINNING_OF_TABLE
                                                || outputStat.CurrentRec == (uint)MinimalEntryID.MID_END_OF_TABLE;

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R319. The value of CurrentRec is {0}.", outputStat.CurrentRec);

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R319
                this.Site.CaptureRequirementIfIsTrue(
                    isCurrentRecSetCorrectValue,
                    319,
                    @"[In STAT] CurrentRec: This field contains a Minimal Entry ID.");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R322
                // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    322,
                    @"[In STAT] Delta: This field contains a long value.");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R325
                // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    325,
                    @"[In STAT] NumPos: This field contains a DWORD value that specifies a position in the table.");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R330
                // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    330,
                    @"[In STAT] TotalRecs: This field contains a DWORD value that specifies the number of rows in the table.");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R334
                // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    334,
                    @"[In STAT] CodePage: This field contains a DWORD value that represents a code page.");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R338
                // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    338,
                    @"[In STAT] TemplateLocale: This field contains a DWORD value that represents a language code identifier (LCID).");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R341
                // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    341,
                    @"[In STAT] SortLocale: This field contains a DWORD value that represents an LCID.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R337");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R337
            this.Site.CaptureRequirementIfAreEqual<uint>(
                inputStat.CodePage,
                outputStat.CodePage,
                337,
                @"[In STAT] [CodePage] The server MUST NOT modify this field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R340");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R340
            this.Site.CaptureRequirementIfAreEqual<uint>(
                inputStat.TemplateLocale,
                outputStat.TemplateLocale,
                340,
                @"[In STAT] [TemplateLocale] The server MUST NOT modify this field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R344");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R344
            this.Site.CaptureRequirementIfAreEqual<uint>(
                inputStat.SortLocale,
                outputStat.SortLocale,
                344,
                @"[In STAT] [SortLocale] The server MUST NOT modify this field.");

            if (inputStat.TemplateLocale == (uint)DefaultLCID.NSPI_DEFAULT_LOCALE)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R480");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R480
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)DefaultLCID.NSPI_DEFAULT_LOCALE,
                    outputStat.TemplateLocale,
                    480,
                    @"[In Unicode String Comparison] The server MUST minimally support the LCID NSPI_DEFAULT_LOCALE flag, as specified in section 2.2.1.4.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R496");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R496
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)DefaultLCID.NSPI_DEFAULT_LOCALE,
                    outputStat.TemplateLocale,
                    496,
                    @"[In string Sorting] The server MUST minimally support the LCID NSPI_DEFAULT_LOCALE flag.");

                if (inputStat.SortType == (uint)TableSortOrder.SortTypeDisplayName && inputStat.TemplateLocale == (uint)DefaultLCID.NSPI_DEFAULT_LOCALE)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1749");

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1749
                    // The sort order has been verified by MS-OXNSPI_R93. So here just verify the SortType field in the returned STAT structure.
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        (uint)TableSortOrder.SortTypeDisplayName,
                        outputStat.SortType,
                        1749,
                        @"[In Table Sort Orders] The Exchange NSPI server under test MUST support this sort order [ascending on the PidTagDisplayName property] for the default LCID NSPI_DEFAULT_LOCALE.");
                }
            }

            this.VerifyMinimalEntryIDStructure();
        }

        /// <summary>
        /// Verify that ContainerID in STAT cannot be modified.
        /// </summary>
        /// <param name="inputStat">The inputStat parameter of the operation that uses the STAT structure.</param>
        /// <param name="outputStat">The outputStat parameter of the operation that uses the STAT structure.</param>
        private void VerifyContainerIDCanNotBeModified(STAT inputStat, STAT outputStat)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R318");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R318
            this.Site.CaptureRequirementIfAreEqual<uint>(
                inputStat.ContainerID,
                outputStat.ContainerID,
                318,
                @"[In STAT] [ContainerID] The server MUST NOT modify this field in any NSPI method except the NspiGetMatches method.");
        }

        #endregion

        #region MinimalEntryID
        /// <summary>
        /// Verify the MinimalEntryID structure.
        /// </summary>
        private void VerifyMinimalEntryIDStructure()
        {
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1591
            // Because the underlying parser code has parsed out the structure which contains it, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1591,
                @"[In MinimalEntryID] This type [MinimalEntryID] is declared as follows:
                typedef DWORD MinEntryID;");
        }

        #endregion

        #region EphemeralEntryID
        /// <summary>
        /// Verify the EphemeralEntryID structure.
        /// </summary>
        /// <param name="ephemeralEntryID">A EphemeralEntryID structure to be verified.</param>
        private void VerifyEphemeralEntryID(EphemeralEntryID ephemeralEntryID)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R361");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R361
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                Marshal.SizeOf(ephemeralEntryID.IDType),
                361,
                @"[In EphemeralEntryID] ID Type (1 byte): The type of this ID.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R362");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R362
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x87,
                ephemeralEntryID.IDType,
                362,
                @"[In EphemeralEntryID] [IDType] The value is the constant 0x87.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R364");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R364
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                Marshal.SizeOf(ephemeralEntryID.R1),
                364,
                @"[In EphemeralEntryID] R1 (1 byte): [means this field is] Reserved.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1713");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1713
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                ephemeralEntryID.R1,
                1713,
                @"[In EphemeralEntryID] Server MUST set this value [R1] to the constant 0x00.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R366");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R366
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                Marshal.SizeOf(ephemeralEntryID.R2),
                366,
                @"[In EphemeralEntryID] R2 (1 byte): [means this field is] Reserved.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1714");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1714
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00,
                ephemeralEntryID.R2,
                1714,
                @"[In EphemeralEntryID] Server MUST set this value [R2] to the constant 0x00.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R368");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R368
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                Marshal.SizeOf(ephemeralEntryID.R3),
                368,
                @"[In EphemeralEntryID] R3 (1 byte): [means this field is] Reserved.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1715");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1715
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                ephemeralEntryID.R3,
                1715,
                @"[In EphemeralEntryID] Server MUST set this value [R3] to the constant 0x00.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R370");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R370
            this.Site.CaptureRequirementIfAreEqual<int>(
                16,
                Marshal.SizeOf(ephemeralEntryID.ProviderUID),
                370,
                @"[In EphemeralEntryID] ProviderUID (16 bytes): A FlatUID_r value, as specified in section 2.2.2.1, that contains the GUID of the server that issued this Ephemeral Entry ID.");

            this.VerifyFlatUID_r();

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R372");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R372
            this.Site.CaptureRequirementIfAreEqual<int>(
                4,
                Marshal.SizeOf(ephemeralEntryID.R4),
                372,
                @"[In EphemeralEntryID] R4 (4 bytes): [means this field is] Reserved.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1716");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1716
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x1,
                ephemeralEntryID.R4,
                1716,
                @"[In EphemeralEntryID] Server MUST set this value [R4] to the constant 0x00000001.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R374
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                374,
                @"[In EphemeralEntryID] Display Type (4 bytes): The display type of the object specified by this Ephemeral Entry ID.");

            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                606,
                @"[In Object Identity] [Ephemeral Identifier] The display type of the object is included in the Ephemeral Identifier.");

            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                375,
                @"[In EphemeralEntryID] [Display Type] This value is expressed in little-endian format.");

            bool isValidDisplayType = false;
            if (ephemeralEntryID.DisplayType == DisplayTypeValue.DT_ADDRESS_TEMPLATE
                || ephemeralEntryID.DisplayType == DisplayTypeValue.DT_AGENT
                || ephemeralEntryID.DisplayType == DisplayTypeValue.DT_CONTAINER
                || ephemeralEntryID.DisplayType == DisplayTypeValue.DT_DISTLIST
                || ephemeralEntryID.DisplayType == DisplayTypeValue.DT_FORUM
                || ephemeralEntryID.DisplayType == DisplayTypeValue.DT_MAILUSER
                || ephemeralEntryID.DisplayType == DisplayTypeValue.DT_ORGANIZATION
                || ephemeralEntryID.DisplayType == DisplayTypeValue.DT_PRIVATE_DISTLIST
                || ephemeralEntryID.DisplayType == DisplayTypeValue.DT_REMOTE_MAILUSER
                || ephemeralEntryID.DisplayType == DisplayTypeValue.DT_SEARCH
                || ephemeralEntryID.DisplayType == DisplayTypeValue.DT_TEMPLATE)
            {
                isValidDisplayType = true;
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R378: [In EphemeralEntryID] [Display Type] The server MUST set this field {0}.", ephemeralEntryID.DisplayType);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R378
            this.Site.CaptureRequirementIfIsTrue(
                isValidDisplayType,
                378,
                @"[In EphemeralEntryID] [Display Type] The server MUST set this field when this data structure is returned in an output parameter.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R380");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R380
            this.Site.CaptureRequirementIfAreEqual<int>(
                4,
                Marshal.SizeOf(ephemeralEntryID.Mid),
                380,
                @"[In EphemeralEntryID] MId (4 bytes): The Minimal Entry ID of this object, as specified in section 2.2.9.1");

            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                381,
                @"[In EphemeralEntryID] [MId] This value is expressed in little-endian format.");
        }

        #endregion

        #region PermanentEntryID
        /// <summary>
        /// Verify the PermanentEntryID structure.
        /// </summary>
        /// <param name="permanentEntryID">A PermanentEntryID to be verified.</param>
        private void VerifyPermanentEntryID(PermanentEntryID permanentEntryID)
        {
            // The Permanent Entry ID is parsed from the PtypBinary property type. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                384,
                @"[In PermanentEntryID] Permanent Entry IDs are transmitted in the protocol as values with the PtypBinary property type.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R392");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R392
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                Marshal.SizeOf(permanentEntryID.IDType),
                392,
                @"[In PermanentEntryID] ID Type (1 byte): The type of this ID.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R393");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R393
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                permanentEntryID.IDType,
                393,
                @"[In PermanentEntryID] [ID Type] The value is the constant 0x00.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R395");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R395
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                Marshal.SizeOf(permanentEntryID.R1),
                395,
                @"[In PermanentEntryID] R1 (1 byte): [means that this field is] Reserved.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1717");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1717
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                permanentEntryID.R1,
                1717,
                @"[In PermanentEntryID] Server MUST set this value[R1] to the constant 0x00.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R397");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R397
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                Marshal.SizeOf(permanentEntryID.R2),
                397,
                @"[In PermanentEntryID] R2 (1 byte): [means that this field is] Reserved.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1718");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1718
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                permanentEntryID.R2,
                1718,
                @"[In PermanentEntryID] [R2] All servers MUST set this value to the constant 0x00.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R399");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R399
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                Marshal.SizeOf(permanentEntryID.R3),
                399,
                @"[In PermanentEntryID] R3 (1 byte): [means that this field is] Reserved.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1719");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1719
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                permanentEntryID.R3,
                1719,
                @"[In PermanentEntryID] Server MUST set this value [R3] to the constant 0x00.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R401.");

            byte[] guidNSPI = { 0xDC, 0xA7, 0x40, 0xC8, 0xC0, 0x42, 0x10, 0x1A, 0xB4, 0xB9, 0x08, 0x00, 0x2B, 0x2F, 0xE1, 0x82 };
            bool isGUIDContained = true;
            if (guidNSPI.Length != permanentEntryID.ProviderUID.Ab.Length)
            {
                isGUIDContained = false;
            }
            else
            {
                for (int i = 0; i < guidNSPI.Length; i++)
                {
                    if (guidNSPI[i] != permanentEntryID.ProviderUID.Ab[i])
                    {
                        isGUIDContained = false;
                        break;
                    }
                }
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R401
            this.Site.CaptureRequirementIfIsTrue(
                isGUIDContained,
                401,
                @"[In PermanentEntryID] ProviderUID (16 bytes): A FlatUID_r value that contains the constant GUID specified in Permanent Entry ID GUID, as specified in section 2.2.1.7.");

            this.VerifyFlatUID_r();

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1664
            // If MS-OXNSPI_R401 has been verified, MS-OXNSPI_R1664 can be verified directly.
            this.Site.CaptureRequirement(
                1664,
                @"[In Permanent Entry ID GUID] GUID_NSPI ({0xDC, 0xA7, 0x40, 0xC8, 0xC0, 0x42, 0x10, 0x1A, 0xB4, 0xB9, 0x08, 0x00, 0x2B, 0x2F, 0xE1, 0x82}): Represents the Exchange Server NSPI Protocol in Permanent Entry IDs.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R403");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R403
            this.Site.CaptureRequirementIfAreEqual<int>(
                4,
                Marshal.SizeOf(permanentEntryID.R4),
                403,
                @"[In PermanentEntryID] R4 (4 bytes): [means that this field is] Reserved.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1720");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1720
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x1,
                permanentEntryID.R4,
                1720,
                @"[In PermanentEntryID] Server MUST set this value[R4] to the constant 0x00000001.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R405
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                405,
                @"[In PermanentEntryID] Display Type string (4 bytes): The display type of the object specified by this Permanent Entry ID.");

            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                604,
                @"[In Object Identity] [Permanent Identifier] The display type of the object is included in the Permanent Identifier.");

            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                406,
                @"[In PermanentEntryID] [Display Type String] This value is expressed in little-endian format.");

            bool isValidDisplayType = false;
            if (permanentEntryID.DisplayTypeString == DisplayTypeValue.DT_ADDRESS_TEMPLATE
                || permanentEntryID.DisplayTypeString == DisplayTypeValue.DT_AGENT
                || permanentEntryID.DisplayTypeString == DisplayTypeValue.DT_CONTAINER
                || permanentEntryID.DisplayTypeString == DisplayTypeValue.DT_DISTLIST
                || permanentEntryID.DisplayTypeString == DisplayTypeValue.DT_FORUM
                || permanentEntryID.DisplayTypeString == DisplayTypeValue.DT_MAILUSER
                || permanentEntryID.DisplayTypeString == DisplayTypeValue.DT_ORGANIZATION
                || permanentEntryID.DisplayTypeString == DisplayTypeValue.DT_PRIVATE_DISTLIST
                || permanentEntryID.DisplayTypeString == DisplayTypeValue.DT_REMOTE_MAILUSER
                || permanentEntryID.DisplayTypeString == DisplayTypeValue.DT_SEARCH
                || permanentEntryID.DisplayTypeString == DisplayTypeValue.DT_TEMPLATE)
            {
                isValidDisplayType = true;
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R409: [In PermanentEntryID] [Display Type String] the value the server set in this field is {0}.", permanentEntryID.DisplayTypeString);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R409
            this.Site.CaptureRequirementIfIsTrue(
                isValidDisplayType,
                409,
                @"[In PermanentEntryID] [Display Type String] A server MUST set this field when this data structure is returned in an output parameter.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R411.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R411
            this.Site.CaptureRequirementIfIsNotNull(
                permanentEntryID.DistinguishedName,
                411,
                @"[In PermanentEntryID] Distinguished Name (variable): The DN of the object specified by this Permanent Entry ID.");

        }

        #endregion

        #region Nspihandle
        /// <summary>
        /// Verify The NSPI_HANDLE definition.
        /// </summary>
        private void VerifyNspihandle()
        {
            // Because the underlying parser code has parsed out the structure which contains it, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1592,
                @"[In NSPI_HANDLE] This type [NSPI_HANDLE] is declared as follows:
                typedef [context_handle] void* NSPI_HANDLE;");
        }

        #endregion

        #endregion

        #region NspiBind
        /// <summary>
        /// The capture code in NspiBind.
        /// </summary>
        private void VerifyNspiBind()
        {
            this.VerifyFlatUID_r();

            // The IDL code parses the return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                676,
                @"[In NspiBind] Return Values: The server returns a LONG [MS-DTYP] value that specifies the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1669,
                    @"[In NspiBind] 
                The interface definition is:

                long NspiBind(
                [in] handle_t hRpc,
                [in] DWORD dwFlags,
                [in] STAT* pStat,
                [in, out, unique] FlatUID_r* pServerGuid,
                [out, ref] NSPI_HANDLE* contextHandle
                );");
            }

            this.VerifyNspihandle();
        }

        #endregion

        #region NspiUnbind
        /// <summary>
        /// The capture code in NspiUnbind.
        /// </summary>
        /// <param name="contextHandle">The contextHandle indicating which handle will be released.</param>
        private void VerifyNspiUnbind(IntPtr contextHandle)
        {
            // The IDL code parses return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                706,
                @"[In NspiUnbind] Return Values: The server returns a DWORD value that specifies the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1672,
                    @"[In NspiUnbind] 
                The interface definition is:
                DWORD NspiUnbind(
                [in, out] NSPI_HANDLE* contextHandle,
                [in] DWORD Reserved
                );");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R716");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R716
                // IntPtr.Zero indicates that the contextHandle is null.
                this.Site.CaptureRequirementIfAreEqual<IntPtr>(
                    IntPtr.Zero,
                    contextHandle,
                    716,
                    @"[In NspiUnbind] [Server Processing Rules: Upon receiving message NspiUnbind, the server MUST process the data from the message subject to the following constraints:] [constraint 3] The server MUST set the output parameter contextHandle to NULL.");
            }
        }

        #endregion

        #region NspiGetSpecialTable
        /// <summary>
        /// The capture code in NspiGetSpecialTable.
        /// </summary>
        /// <param name="returnValue">The return value of NspiGetSpecialTable.</param>
        /// <param name="rows">The search result of NspiGetSpecialTable.</param>
        private void VerifyNspiGetSpecialTable(ErrorCodeValue returnValue, PropertyRowSet_r? rows)
        {
            // The IDL code parses the return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                728,
                @"[In NspiGetSpecialTable] Return Values: The server returns a LONG [MS-DTYP] value that specifies the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1674,
                    @"[In NspiGetSpecialTable] The interface definition is:
                long NspiGetSpecialTable(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD dwFlags,
                    [in] STAT* pStat,
                    [in, out] DWORD* lpVersion,
                    [out] PropertyRowSet_r** ppRows
                );");
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R726
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                726,
                @"[In NspiGetSpecialTable] ppRows: A PropertyRowSet_r structure.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R724
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                724,
                @"[In NspiGetSpecialTable] lpVersion: A reference to a DWORD.");

            this.VerifyReturnValues(returnValue);
            if (rows != null && rows.Value.ARow != null)
            {
                this.VerifyPropertyRowSetStructure(rows.Value);

                foreach (PropertyRow_r propertyRow in rows.Value.ARow)
                {
                    this.VerifyStringOrString8Value(propertyRow);
                }
            }
        }

        #endregion

        #region NspiUpdateStat
        /// <summary>
        /// The capture code in NspiUpdateStat.
        /// </summary>
        /// <param name="returnValue">The return value of NspiUpdateStat.</param>
        /// <param name="inputStat">The input parameter STAT.</param>
        /// <param name="outputStat">The output parameter STAT.</param>
        private void VerifyNspiUpdateStat(ErrorCodeValue returnValue, STAT inputStat, STAT outputStat)
        {
            // The IDL code parses the return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                776,
                @"[In NspiUpdateStat] Return Values: The server returns a long value specifying the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // If the operation can be called successfully, the interface must be correct.
                this.Site.CaptureRequirement(
                    1675,
                    @"[In NspiUpdateStat] The interface definition is:
                long NspiUpdateStat(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD Reserved,
                    [in, out] STAT* pStat,
                    [in, out, unique] long* plDelta
                );");
            }

            this.VerifyReturnValues(returnValue);

            if (returnValue == ErrorCodeValue.Success)
            {
                // According to the description in section 2.2.8 and 2.3.8.1, if the returned CurrentRec field is set to MID_BEGINNING_OF_TABLE or 
                // MID_END_OF_TABLE, or its value is greater than or equal to 0x10, it is a valid Minimal Entry ID.
                bool isCurrentRecSetCorrectValue = outputStat.CurrentRec >= 0x10
                                                || outputStat.CurrentRec == (uint)MinimalEntryID.MID_BEGINNING_OF_TABLE
                                                || outputStat.CurrentRec == (uint)MinimalEntryID.MID_END_OF_TABLE;

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R796, the value of CurrentRec is {0}", outputStat.CurrentRec);

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R796
                this.Site.CaptureRequirementIfIsTrue(
                    isCurrentRecSetCorrectValue,
                    796,
                    @"[In NspiUpdateStat] [Server Processing Rules: Upon receiving message NspiUpdateStat, the server MUST process the data from the message subject to the following constraints:] [Constraint 8] The server MUST set the CurrentRec field of the parameter pStat to the Minimal Entry ID of the current row in the current address book container.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R79, the value of CurrentRec is {0}", outputStat.CurrentRec);

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R79
                this.Site.CaptureRequirementIfIsTrue(
                    isCurrentRecSetCorrectValue,
                    79,
                    @"[In Positioning Minimal Entry IDs] The following table lists the possible values [MID_BEGINNING_OF_TABLE, MID_END_OF_TABLE and MID_CURRENT].");
            }

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R799. The value of CodePage in the output stat is {0}. The value of ContainerID in the output stat is {1}. The value of SortLocale in the output stat is {2}. The value of SortType in the output stat is {3}. The value of TemplateLocale in the output stat is {4}.",
                outputStat.CodePage,
                outputStat.ContainerID,
                outputStat.SortLocale,
                outputStat.SortType,
                outputStat.TemplateLocale);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R799
            bool isVerifyR799 = (outputStat.CodePage == inputStat.CodePage)
                                && (outputStat.ContainerID == inputStat.ContainerID)
                                && (outputStat.SortLocale == inputStat.SortLocale)
                                && (outputStat.SortType == inputStat.SortType)
                                && (outputStat.TemplateLocale == inputStat.TemplateLocale);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifyR799,
                799,
                @"[In NspiUpdateStat] [Server Processing Rules: Upon receiving message NspiUpdateStat, the server MUST process the data from the message subject to the following constraints:] [Constraint 11] The server MUST leave all other fields [other than CurrentRec, NumPos, TotalRecs and Delta] of the parameter pStat unchanged.");

            this.VerifySTATStructure(returnValue, inputStat, outputStat, false);
            this.VerifyContainerIDCanNotBeModified(inputStat, outputStat);
        }

        #endregion

        #region NspiQueryColumns
        /// <summary>
        /// The capture code in NspiQueryColumns.
        /// </summary>
        /// <param name="returnValue">The return value of NspiQueryColumns.</param>
        /// <param name="columns">The parameter indicates which columns to be returned.</param>
        private void VerifyNspiQueryColumns(ErrorCodeValue returnValue, PropertyTagArray_r? columns)
        {
            // The IDL code parses the return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                810,
                @"[In NspiQueryColumns] Return Values: The server returns a LONG [MS-DTYP] value that specifies the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // If the operation can be called successfully, the interface must be correct.
                this.Site.CaptureRequirement(
                    1677,
                    @"[In NspiQueryColumns] The interface definition is:
                long NspiQueryColumns(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD Reserved,
                    [in] DWORD dwFlags,
                    [out] PropertyTagArray_r** ppColumns
                );");
            }

            this.VerifyReturnValues(returnValue);

            // Because the underlying code parses ppColumns based on this structure, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                809,
                @"[In NspiQueryColumns] [ppColumns] On return, contains a list of proptags.");

            // Because the underlying code parses ppColumns based on this structure, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                824,
                @"[In NspiQueryColumns] [Server Processing Rules: Upon receiving message NspiQueryColumns, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] Subject to the prior constraints, the server MUST construct a list of all the properties it [server] is aware of and return that list as a PropertyTagArray_r structure, as specified in section 2.2.2.2, in the output parameter ppColumns.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R808
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                808,
                @"[In NspiQueryColumns] ppColumns: A reference to a PropertyTagArray_r structure.");

            if (columns != null)
            {
                this.VerifyPropertyTagArray_r();
                foreach (uint propTag in columns.Value.AulPropTag)
                {
                    this.VerifyPropertyTag(propTag);
                }
            }
        }

        #endregion

        #region NspiGetPropList
        /// <summary>
        /// The capture code in NspiGetPropList.
        /// </summary>
        /// <param name="returnValue">The return value of NspiGetPropList.</param>
        /// <param name="propTags">The parameter indicates which columns to be returned.</param>
        private void VerifyNspiGetPropList(ErrorCodeValue returnValue, PropertyTagArray_r? propTags)
        {
            // The IDL code parses the return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                836,
                @"[In NspiGetPropList] Return Values: The server returns a long value that specifies the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1678,
                    @"[In NspiGetPropList] The interface definition is:
                long NspiGetPropList(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD dwFlags,
                    [in] DWORD dwMId,
                    [in] DWORD CodePage,
                    [out] PropertyTagArray_r** ppPropTags
                );");
            }

            this.VerifyReturnValues(returnValue);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R834
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                834,
                @"[In NspiGetPropList] ppPropTags: A PropertyTagArray_r value.");

            if (propTags != null)
            {
                this.VerifyPropertyTagArray_r();
            }
        }

        #endregion

        #region NspiGetProps
        /// <summary>
        /// The capture code in NspiGetProps.
        /// </summary>
        /// <param name="returnValue">The return value of NspiGetProps.</param>
        /// <param name="propertyRow">The row that the server returns.</param>
        private void VerifyNspiGetProps(ErrorCodeValue returnValue, PropertyRow_r? propertyRow)
        {
            // The IDL code parses the return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                869,
                @"[In NspiGetProps] Return Values: The server returns a long value specifying the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1679,
                    @"[In NspiGetProps] The interface definition is:
                long NspiGetProps(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD dwFlags,
                    [in] STAT* pStat,
                    [in, unique] PropertyTagArray_r* pPropTags,
                    [out] PropertyRow_r** ppRows
                );");
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R867
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                867,
                @"[In NspiGetProps] ppRows: A reference to a PropertyRow_r value.");

            this.VerifyReturnValues(returnValue);
            if (propertyRow != null)
            {
                this.VerifyPropertyRowStructure(propertyRow.Value);
                this.VerifyStringOrString8Value(propertyRow.Value);
            }
        }

        #endregion

        #region NspiQueryRows
        /// <summary>
        /// The capture code in NspiQueryRows.
        /// </summary>
        /// <param name="returnValue">The return value of NspiQueryRows.</param>
        /// <param name="rows">The rows that the server returns.</param>
        /// <param name="inputStat">The input parameter STAT.</param>
        /// <param name="outputStat">The output parameter STAT.</param>
        private void VerifyNspiQueryRows(ErrorCodeValue returnValue, PropertyRowSet_r? rows, STAT inputStat, STAT outputStat)
        {
            // The IDL code parses the return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                931,
                @"[In NspiQueryRows] Return Values: The server returns a long value specifying the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1680,
                    @"[In NspiQueryRows] The interface definition is:
                long NspiQueryRows(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD dwFlags,
                    [in, out] STAT* pStat,
                    [in, range(0,100000)] DWORD dwETableCount,
                    [in, unique, size_is(dwETableCount)] 
                    DWORD* lpETable,
                    [in] DWORD Count,
                    [in, unique] PropertyTagArray_r* pPropTags,
                    [out] PropertyRowSet_r** ppRows
                );");
            }

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R929
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                929,
                @"[In NspiQueryRows] ppRows: A reference to a PropertyRowSet_r value.");

            this.VerifyReturnValues(returnValue);
            if (rows != null && rows.Value.ARow != null)
            {
                this.VerifyPropertyRowSetStructure(rows.Value);

                foreach (PropertyRow_r propertyRow in rows.Value.ARow)
                {
                    this.VerifyStringOrString8Value(propertyRow);
                }
            }

            this.VerifySTATStructure(returnValue, inputStat, outputStat, false);
            this.VerifyContainerIDCanNotBeModified(inputStat, outputStat);
        }

        #endregion

        #region NspiSeekEntries
        /// <summary>
        /// The capture code in NspiSeekEntries.
        /// </summary>
        /// <param name="returnValue">The return value of NspiSeekEntries.</param>
        /// <param name="rows">The rows that the server returns.</param>
        /// <param name="inputStat">The input parameter STAT.</param>
        /// <param name="outputStat">The output parameter STAT.</param>
        private void VerifyNspiSeekEntries(ErrorCodeValue returnValue, PropertyRowSet_r? rows, STAT inputStat, STAT outputStat)
        {
            // The IDL code parses the return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                1012,
                @"[In NspiSeekEntries] Return Values: The server returns a long value specifying the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1681,
                    @"[In NspiSeekEntries] The interface definition is:
                long NspiSeekEntries(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD Reserved,
                    [in, out] STAT* pStat,
                    [in] PropertyValue_r* pTarget,
                    [in, unique] PropertyTagArray_r* lpETable,
                    [in, unique] PropertyTagArray_r* pPropTags,
                    [out] PropertyRowSet_r** ppRows
                );");
            }

            if (rows != null)
            {
                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1010
                // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1010,
                    @"[In NspiSeekEntries] ppRows: A reference to a PropertyRowSet_r value.");

                this.VerifyPropertyRowSetStructure(rows.Value);

                foreach (PropertyRow_r propertyRow in rows.Value.ARow)
                {
                    this.VerifyStringOrString8Value(propertyRow);
                }
            }

            this.VerifyReturnValues(returnValue);
            this.VerifySTATStructure(returnValue, inputStat, outputStat, false);
            this.VerifyContainerIDCanNotBeModified(inputStat, outputStat);
        }

        #endregion

        #region NspiGetMatches
        /// <summary>
        /// The capture code in NspiGetMatches.
        /// </summary>
        /// <param name="returnValue">The return value of NspiGetMatches.</param>
        /// <param name="rows">The rows that the server returns.</param>
        /// <param name="outMIds">A PropertyTagArray_r value which contains a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="inputStat">The input parameter STAT.</param>
        /// <param name="outputStat">The output parameter STAT.</param>
        private void VerifyNspiGetMatches(ErrorCodeValue returnValue, PropertyRowSet_r? rows, PropertyTagArray_r? outMIds, STAT inputStat, STAT outputStat)
        {
            if (outMIds != null)
            {
                this.VerifyPropertyTagArray_r();

                // According to the description in the Open Specification, outMIds holds a list of Minimal Entry IDs that comprise a restricted address book container.
                // So if outMIds is not null, it specifies that the list is instantiated as a PropertyTagArray_r structure.
                this.Site.CaptureRequirement(
                    507,
                    @"[In Explicit Tables] The list [a list of Minimal Entry IDs] is instantiated in the protocol either as an array of DWORDs or as a PropertyTagArray_r structure.");
            }

            // The IDL code parses the returned value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                1097,
                @"[In NspiGetMatches] Return Values: The server returns a long value specifying the return status of the method.");

            // This parameter is defined as PropertyTagArray_r structure. If the parser code can parse the structure successfully, it illustrates that the server returns the result correctly.
            this.Site.CaptureRequirement(
                1091,
                @"[In NspiGetMatches] ppOutMIds: A PropertyTagArray_r value.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1682,
                    @"[In NspiGetMatches] The interface definition is:

                long NspiGetMatches(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD Reserved1,
                    [in, out] STAT* pStat,
                    [in, unique] PropertyTagArray_r* pReserved,
                    [in] DWORD Reserved2,
                    [in, unique] Restriction_r* Filter,
                    [in, unique] PropertyName_r* lpPropName,
                    [in] DWORD ulRequested,
                    [out] PropertyTagArray_r** ppOutMIds,
                    [in, unique] PropertyTagArray_r* pPropTags,
                    [out] PropertyRowSet_r** ppRows
                );");
            }

            if (returnValue == ErrorCodeValue.Success)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1166");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1166
                // The parameter rows is returned from the server. If it is not null, it illustrates that the 
                // server has returned the PropertyRowSet_r structure in the output parameter pProws.
                this.Site.CaptureRequirementIfIsNotNull(
                    rows,
                    1166,
                    @"[In NspiGetMatches] The server MUST return the constructed PropertyRowSet_r in the output parameter ppRows.");
            }

            this.VerifyReturnValues(returnValue);

            if (rows != null)
            {
                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1095
                // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1095,
                    @"[In NspiGetMatches] ppRows: A reference to a PropertyRowSet_r value.");

                foreach (PropertyRow_r propertyRow in rows.Value.ARow)
                {
                    this.VerifyStringOrString8Value(propertyRow);

                    // Parse the Minimal Entry ID if it exists.
                    if (propertyRow.LpProps[0].Value.Bin.Cb > 0)
                    {
                        EphemeralEntryID ephemeralEntryID = AdapterHelper.ParseEphemeralEntryIDFromBytes(propertyRow.LpProps[0].Value.Bin.Lpb);

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R506");

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R506
                        // The NspiGetMatches method returns an Explicit Table. So if the parsed ephemeralEntryID is not null, it means that the Explicit Table is implemented as a list of Minimal Entry IDs.
                        this.Site.CaptureRequirementIfIsNotNull(
                            ephemeralEntryID,
                            506,
                            @"[In Explicit Tables] This table [Explicit Table] is implemented as a list of Minimal Entry IDs.");
                    }
                }

                this.VerifyPropertyRowSetStructure(rows.Value);
            }

            this.VerifySTATStructure(returnValue, inputStat, outputStat, true);
            this.VerifyMinimalEntryIDStructure();
        }

        #endregion

        #region NspiResortRestriction
        /// <summary>
        /// The capture code in NspiResortRestriction.
        /// </summary>
        /// <param name="returnValue">The return value of NspiResortRestriction.</param>
        /// <param name="outMIds">The outMIds parameter of NspiResortRestriction</param>
        /// <param name="inputStat">The input parameter STAT.</param>
        /// <param name="outputStat">The output parameter STAT.</param>
        private void VerifyNspiResortRestriction(ErrorCodeValue returnValue, PropertyTagArray_r? outMIds, STAT inputStat, STAT outputStat)
        {
            // The IDL code parses return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                1177,
                @"[In NspiResortRestriction] Return Values: The server returns a long value that specifies the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1683,
                    @"[In NspiResortRestriction] The interface definition is:

                long NspiResortRestriction(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD Reserved,
                    [in, out] STAT* pStat,
                    [in] PropertyTagArray_r* pInMIds,
                    [in, out] PropertyTagArray_r** ppOutMIds
                );");
            }

            this.VerifyReturnValues(returnValue);

            if (returnValue == ErrorCodeValue.Success)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1201");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1201
                // The outMIds, which is returned from the server, contains a Explicit Table. If it is not null, it 
                // illustrates that the server has returned the Explicit Table in the output parameter pProws.
                this.Site.CaptureRequirementIfIsNotNull(
                    outMIds,
                    1201,
                    @"[In NspiResortRestriction] [Server Processing Rules: Upon receiving message NspiResortRestriction, the server MUST process the data from the message subject to the following constraints:] [Constraint 7] The server MUST return the constructed Explicit Table in the output parameter ppOutMIds.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1204");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1204
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    outMIds.Value.CValues,
                    outputStat.TotalRecs,
                    1204,
                    @"[In NspiResortRestriction] [Server Processing Rules: Upon receiving message NspiResortRestriction, the server MUST process the data from the message subject to the following constraints:] [Constraint 8] The Server MUST update the output parameter pStat as follows: The TotalRecs field is set to the number of objects in the constructed Explicit Table.");
            }

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1207. The value of CodePage in the output stat is {0}. The value of ContainerID in the output stat is {1}. The value of Delta in the output stat is {2}. The value of SortLocale in the output stat is {3}. The value of SortType in the output stat is {4}. The value of TemplateLocale in the output stat is {5}.",
                outputStat.CodePage,
                outputStat.ContainerID,
                outputStat.Delta,
                outputStat.SortLocale,
                outputStat.SortType,
                outputStat.TemplateLocale);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1207
            bool isVerifyR1207 = (outputStat.CodePage == inputStat.CodePage)
                                && (outputStat.ContainerID == inputStat.ContainerID)
                                && (outputStat.Delta == inputStat.Delta)
                                && (outputStat.SortLocale == inputStat.SortLocale)
                                && (outputStat.SortType == inputStat.SortType)
                                && (outputStat.TemplateLocale == inputStat.TemplateLocale);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifyR1207,
                1207,
                @"[In NspiResortRestriction] [Server Processing Rules: Upon receiving message NspiResortRestriction, the server MUST process the data from the message subject to the following constraints:] [Constraint 8 The Server MUST update the output parameter pStat as follows:] The server MUST NOT modify any other fields of the output parameter pStat.");

            if (outMIds != null)
            {
                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1175
                // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1175,
                    @"[In NspiResortRestriction] ppOutMIds: A PropertyTagArray_r value.");
            }

            this.VerifySTATStructure(returnValue, inputStat, outputStat, false);
            this.VerifyMinimalEntryIDStructure();
            this.VerifyContainerIDCanNotBeModified(inputStat, outputStat);
        }

        #endregion

        #region NspiCompareMIds
        /// <summary>
        /// The capture code in NspiCompareMIds.
        /// </summary>
        /// <param name="returnValue">The NspiCompareMIds return value.</param>
        private void VerifyNspiCompareMIds(ErrorCodeValue returnValue)
        {
            // The IDL code parses return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                1219,
                @"[In NspiCompareMIds] Return Values: The server returns a long value that specifies the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1684,
                    @"[In NspiCompareMIds] The interface definition is:
                long NspiCompareMIds(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD Reserved,
                    [in] STAT* pStat,
                    [in] DWORD MId1,
                    [in] DWORD MId2,
                    [out] long* plResult
                );");
            }

            this.VerifyReturnValues(returnValue);

            if (returnValue == ErrorCodeValue.Success)
            {
                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1217
                // This test suite parses code according to this definition, so if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1217,
                    @"[In NspiCompareMIds] plResult: A pointer to a long value [which specifies the compare result of the NspiCompareMids method. ]");
            }
        }

        #endregion

        #region NspiDNToMId
        /// <summary>
        /// The capture code in NspiDNToMId.
        /// </summary>
        /// <param name="returnValue">The NspiDNToMId return value.</param>
        /// <param name="mids">The mids parameter of NspiDNToMId.</param>
        private void VerifyNspiDNToMId(ErrorCodeValue returnValue, PropertyTagArray_r? mids)
        {
            // The IDL code parses the return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                1251,
                @"[In NspiDNToMId] Return Values: The server returns a long value that specifies the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1685,
                    @"[In NspiDNToMId] The interface definition is:
                long NspiDNToMId(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD Reserved,
                    [in] stringsArray_r* pNames,
                    [out] PropertyTagArray_r** ppMIds
                );");
            }

            this.VerifyReturnValues(returnValue);

            if (returnValue == ErrorCodeValue.Success)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1262");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1262
                // This structure, which is returned from the server, contains a list of Minimal Entry IDs. 
                // If it's not null, it illustrates that the server must have constructed the list.
                this.Site.CaptureRequirementIfIsNotNull(
                    mids,
                    1262,
                    @"[In NspiDNToMId] [Server Processing Rules: Upon receiving message NspiDNToMId, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] The server constructs a list of Minimal Entry IDs to return to the client, encoding the mappings.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1264");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1264
            // This structure, which is returned from the server, contains a list of Minimal Entry IDs. If it's not null, 
            // it illustrates that the server has returned the list of Minimal Entry IDs in the output parameter ppMIds.
            this.Site.CaptureRequirementIfIsNotNull(
                mids,
                1264,
                @"[In NspiDNToMId] [Server Processing Rules: Upon receiving message NspiDNToMId, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] The server MUST return the list in the output parameter ppMIds.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1249
            // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1249,
                @"[In NspiDNToMId] ppMIds: A PropertyTagArray_r value.");

            if (mids != null)
            {
                this.VerifyPropertyTagArray_r();
            }

            this.VerifyMinimalEntryIDStructure();
        }

        #endregion

        #region NspiModProps
        /// <summary>
        /// The capture code in NspiModProps.
        /// </summary>
        /// <param name="returnValue">The return value of NspiModProps.</param>
        private void VerifyNspiModProps(ErrorCodeValue returnValue)
        {
            // The IDL code parses the return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                1276,
                @"[In NspiModProps] Return Values: The server returns a long value that specifies the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1686,
                    @"[In NspiModProps] The interface definition is:

                long NspiModProps(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD Reserved,
                    [in] STAT* pStat,
                    [in, unique] PropertyTagArray_r* pPropTags,
                    [in] PropertyRow_r* pRow
                );");
            }

            this.VerifyReturnValues(returnValue);
        }

        #endregion

        #region NspiModLinkAtt
        /// <summary>
        /// The capture code in NspiModLinkAtt.
        /// </summary>
        /// <param name="returnValue">The return value of NspiModLinkAtt.</param>
        private void VerifyNspiModLinkAtt(ErrorCodeValue returnValue)
        {
            // The IDL code parses return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                1317,
                @"[In NspiModLinkAtt] Return Values: The server returns a long value that specifies the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1687,
                    @"[In NspiModLinkAtt] The interface definition is:

                long NspiModLinkAtt(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD dwFlags,
                    [in] DWORD ulPropTag,
                    [in] DWORD dwMId,
                    [in] BinaryArray_r* lpEntryIds
                );");
            }

            this.VerifyReturnValues(returnValue);
        }

        #endregion

        #region NspiResolveNames
        /// <summary>
        /// The capture code in NspiResolveNames.
        /// </summary>
        /// <param name="returnValue">The return value of NspiResolveNames.</param>
        /// <param name="mids">The mids parameter of NspiResolveNames</param>
        /// <param name="rows">The PropertyRowSet_r structure the server returns.</param>
        private void VerifyNspiResolveNames(ErrorCodeValue returnValue, PropertyTagArray_r? mids, PropertyRowSet_r? rows)
        {
            // The IDL code parses the return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                1355,
                @"[In NspiResolveNames] Return Values: The server returns a long value that specifies the return status of the method.");

            // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
            this.Site.CaptureRequirement(
                1688,
                @"[In NspiResolveNames] The interface definition is:

                long NspiResolveNames(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD Reserved,
                    [in] STAT* pStat,
                    [in, unique] PropertyTagArray_r* pPropTags,
                    [in] stringsArray_r* paStr,
                    [out] PropertyTagArray_r** ppMIds,
                    [out] PropertyRowSet_r** ppRows
                );");

            if (returnValue == ErrorCodeValue.Success)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1374");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1374
                // This structure, which is returned from the server, contains a list of Minimal Entry IDs. If it's not null, it illustrates that the 
                // server must have constructed a list of the Minimal Entry IDs.
                this.Site.CaptureRequirementIfIsNotNull(
                    mids,
                    1374,
                    @"[In NspiResolveNames] [Server Processing Rules: Upon receiving message NspiResolveNames, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] The server constructs a list of the Minimal Entry IDs specified in section 2.2.1.9 to return to the client.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1377");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1377
                // This structure, which is returned from the server, contains a list of Minimal Entry IDs. If it's not null, it illustrates that the 
                // server has returned a list of Minimal Entry IDs in the output parameter ppMIds.
                this.Site.CaptureRequirementIfIsNotNull(
                    mids,
                    1377,
                    @"[In NspiResolveNames] [Server Processing Rules: Upon receiving message NspiResolveNames, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] The server MUST return this list of Minimal Entry IDs [Minimal Entry IDs resulted from applying the ANR process] in the output parameter ppMIds.");
            }

            if (mids != null)
            {
                this.VerifyPropertyTagArray_r();

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1351
                // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1351,
                    @"[In NspiResolveNames] ppMIds: A PropertyTagArray_r value.");
            }

            if (rows != null)
            {
                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1353
                // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1353,
                    @"[In NspiResolveNames] ppRows: A reference to a PropertyRowSet_r structure (section 2.2.4), [which contains the address book container rows that the server returns in response to the request.]");

                this.VerifyPropertyRowSetStructure(rows.Value);

                foreach (PropertyRow_r propertyRow in rows.Value.ARow)
                {
                    this.VerifyStringOrString8Value(propertyRow);
                }
            }

            this.VerifyThreePossbleOutcomesOfANRProcess(mids, rows);

            this.VerifyReturnValues(returnValue);

            this.VerifyMinimalEntryIDStructure();
        }

        #endregion

        #region NspiResolveNamesW
        /// <summary>
        /// Verify there are three possible outcomes to the ANR process.
        /// </summary>
        /// <param name="mids">The mids parameter of NspiResolveNames or NspiResolveNamesW.</param>
        /// <param name="rows">The PropertyRowSet_r structure the server returns.</param>
        private void VerifyThreePossbleOutcomesOfANRProcess(PropertyTagArray_r? mids, PropertyRowSet_r? rows)
        {
            if (rows != null)
            {
                foreach (PropertyRow_r propertyRow in rows.Value.ARow)
                {
                    this.VerifyStringOrString8Value(propertyRow);
                }
            }

            if (mids != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1611.");

                bool validAnrResult = true;
                this.Site.Assert.IsTrue(mids.Value.CValues > 0, "There should be values returned from server.");
                foreach (uint minimalEntryID in mids.Value.AulPropTag)
                {
                    if (!(minimalEntryID == (uint)ANRMinEntryID.MID_AMBIGUOUS
                        || minimalEntryID == (uint)ANRMinEntryID.MID_RESOLVED
                        || minimalEntryID == (uint)ANRMinEntryID.MID_UNRESOLVED))
                    {
                        validAnrResult = false;
                        break;
                    }
                }

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1611
                this.Site.CaptureRequirementIfIsTrue(
                    validAnrResult,
                    1611,
                    @"[In Ambiguous Name Resolution] There are three possible outcomes [MID_UNRESOLVED, MID_AMBIGUOUS and MID_RESOLVED] to the ANR process.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R87.");

                this.Site.CaptureRequirementIfIsTrue(
                    validAnrResult,
                    87,
                    @"[In Ambiguous Name Resolution Minimal Entry IDs] The following table lists the possible values [MID_UNRESOLVED, MID_AMBIGUOUS and MID_RESOLVED].");
            }
        }

        /// <summary>
        /// The capture code in NspiResolveNamesW.
        /// </summary>
        /// <param name="returnValue">The return value of NspiResolveNamesW.</param>
        /// <param name="mids">The mids parameter of NspiResolveNamesW.</param>
        /// <param name="rows">The PropertyRowSet_r structure the server returns.</param>
        private void VerifyNspiResolveNamesW(ErrorCodeValue returnValue, PropertyTagArray_r? mids, PropertyRowSet_r? rows)
        {
            // The IDL code parses the return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                1406,
                @"[In NspiResolveNamesW] Return Values: The server returns a long value that specifies the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1689,
                    @"[In NspiResolveNamesW] The interface definition is:

                long NspiResolveNamesW(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD Reserved,
                    [in] STAT* pStat,
                    [in, unique] PropertyTagArray_r* pPropTags,
                    [in] WstringsArray_r* paWStr,
                    [out] PropertyTagArray_r** ppMIds,
                    [out] PropertyRowSet_r** ppRows
                );");
            }

            if ((returnValue == ErrorCodeValue.Success) || (returnValue == ErrorCodeValue.ErrorsReturned))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1425");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1425
                // This structure, which is returned from the server, contains a list of Minimal Entry IDs. 
                // If it's not null, it illustrates that the server must have constructed the list.
                this.Site.CaptureRequirementIfIsNotNull(
                    mids,
                    1425,
                    @"[In NspiResolveNamesW] [Server Processing Rules: Upon receiving message NspiResolveNamesW, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] The server constructs a list of the Minimal Entry IDs specified in section 2.2.1.9 to return to the client.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1428");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1428
                // This structure, which is returned from the server, contains a list of Minimal Entry IDs. If it's not null, it illustrates that the 
                // server has returned the list of Minimal Entry IDs in the output parameter ppMIds.
                this.Site.CaptureRequirementIfIsNotNull(
                    mids,
                    1428,
                    @"[In NspiResolveNamesW] [Server Processing Rules: Upon receiving message NspiResolveNamesW, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] The server MUST return this list of Minimal Entry IDs [Minimal Entry IDs resulted from the ANR process] in the output parameter ppMIds.");
            }

            if (mids != null)
            {
                this.VerifyPropertyTagArray_r();

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1403
                // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1403,
                    @"[In NspiResolveNamesW] ppMIds: A PropertyTagArray_r value.");
            }

            if (rows != null)
            {
                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1404
                // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1404,
                    @"[In NspiResolveNamesW] ppRows: A reference to a PropertyRowSet_r structure (section 2.2.4), [which contains the address book container rows that the server returns in response to the request.] ");

                this.VerifyPropertyRowSetStructure(rows.Value);
            }

            this.VerifyReturnValues(returnValue);
            this.VerifyThreePossbleOutcomesOfANRProcess(mids, rows);
            this.VerifyMinimalEntryIDStructure();
        }

        #endregion

        #region NspiGetTemplateInfo
        /// <summary>
        /// The capture code in NspiGetTemplateInfo.
        /// </summary>
        /// <param name="returnValue">The return value of NspiGetTemplateInfo.</param>
        /// <param name="flags">The flags parameter of NspiGetTemplateInfo.</param>
        /// <param name="data">The mids parameter of NspiGetTemplateInfo.</param>
        private void VerifyNspiGetTemplateInfo(ErrorCodeValue returnValue, uint flags, PropertyRow_r? data)
        {
            // The IDL code parses the return value. If the IDL code can return success, the server must return the right result.
            this.Site.CaptureRequirement(
                1458,
                @"[In NspiGetTemplateInfo] Return Values: The server returns a long value that specifies the return status of the method.");

            if (this.transport == "ncacn_http" || this.transport == "ncacn_ip_tcp")
            {
                // The following IDL definition is used to generate the stub class, which is used to communicate with server. So if the operation can be called successfully, the following requirement can be captured directly.
                this.Site.CaptureRequirement(
                    1690,
                    @"[In NspiGetTemplateInfo] The interface definition is:

                long NspiGetTemplateInfo(
                    [in] NSPI_HANDLE hRpc,
                    [in] DWORD dwFlags,
                    [in] DWORD ulType,
                    [string, in, unique] char* pDN,
                    [in] DWORD dwCodePage,
                    [in] DWORD dwLocaleID,
                    [out] PropertyRow_r** ppData
                );");
            }

            this.VerifyReturnValues(returnValue);

            if (returnValue == ErrorCodeValue.Success)
            {
                if (data != null)
                {
                    this.VerifyStringOrString8Value(data.Value);

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1456
                    // This test suite parses code according to this definition. So if the codes can reach here, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        1456,
                        @"[In NspiGetTemplateInfo] ppData: A reference to a PropertyRow_r value.");

                    this.VerifyPropertyRowStructure(data.Value);
                }

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1478");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1478
                // This structure is returned from the server. If it's not null, 
                // it illustrates that the server must have constructed a PropertyRow_r value.
                this.Site.CaptureRequirementIfIsNotNull(
                    data,
                    1478,
                    @"[In NspiGetTemplateInfo] [Server Processing Rules: Upon receiving message NspiGetTemplateInfo, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] The server constructs a PropertyRow_r value.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1490");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1490
                // This structure is returned from the server. If it is not null, it illustrates that the 
                // server has returned the PropertyRow_r structure in the output parameter ppData.
                this.Site.CaptureRequirementIfIsNotNull(
                    data,
                    1490,
                    @"[In NspiGetTemplateInfo] [Server Processing Rules: Upon receiving message NspiGetTemplateInfo, the server MUST process the data from the message subject to the following constraints:] [Constraint 7] The server MUST return the constructed PropertyRow_r structure in the output parameter ppData.");

                if ((flags & 0x00000004) == 0x00000004)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXOABKT_R303");

                    bool isPidTagScriptDataAdded = false;
                    foreach (PropertyValue_r propertyValue in data.Value.LpProps)
                    {
                        if (propertyValue.PropTag == 0x00040102)
                        {
                            isPidTagScriptDataAdded = true;
                            this.Site.Log.Add(LogEntryKind.Debug, "The property tag is {0}.", propertyValue.PropTag);
                            break;
                        }
                    }

                    // Verify MS-OXOABKT requirement: MS-OXOABKT_R303
                    this.Site.CaptureRequirementIfIsTrue(
                        isPidTagScriptDataAdded,
                        "MS-OXOABKT",
                        303,
                        @"[In NspiGetTemplateInfo PropertyRow_r format] The Property Added to PropertyRow_r of the flag TI_SCRIPT 0x00000004 is PidTagScriptData 0x00040102.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1484");

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1484
                    this.Site.CaptureRequirementIfIsTrue(
                        isPidTagScriptDataAdded,
                        1484,
                        @"[In NspiGetTemplateInfo] [Server Processing Rules: Upon receiving message NspiGetTemplateInfo, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] The server MUST place this data [the script data for the template] into the PropertyRow_r structure.");
                }
                else if ((flags & 0x00000001) == 0x00000001)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXOABKT_R302");

                    bool isPidTagTemplateDataAdded = false;
                    foreach (PropertyValue_r propertyValue in data.Value.LpProps)
                    {
                        if (propertyValue.PropTag == 0x00010102)
                        {
                            isPidTagTemplateDataAdded = true;
                            this.Site.Log.Add(LogEntryKind.Debug, "The property tag is {0}.", propertyValue.PropTag);
                            break;
                        }
                    }

                    // Verify MS-OXOABKT requirement: MS-OXOABKT_R302
                    this.Site.CaptureRequirementIfIsTrue(
                        isPidTagTemplateDataAdded,
                        "MS-OXOABKT",
                        302,
                        @"[In NspiGetTemplateInfo PropertyRow_r format] The Property Added to PropertyRow_r of the flag TI_TEMPLATE 0x00000001 is PidTagTemplateData 0x00010102.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1486");

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1486
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        (uint)AulProp.PidTagTemplateData,
                        data.Value.LpProps[0].PropTag,
                        1486,
                        @"[In NspiGetTemplateInfo] [Server Processing Rules: Upon receiving message NspiGetTemplateInfo, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] [If the input parameter dwFlags has the TI_TEMPLATE bit set] The server MUST place this data [the user interface data for the template] into the PropertyRow_r structure.");
                }
            }
        }

        #endregion

        #region Verify Unicode string or 8-bit character string value
        /// <summary>
        /// Verify that the string values can be represented as Unicode string or 8-bit character string values.
        /// </summary>
        /// <param name="returnedRows">The returned rows.</param>
        private void VerifyStringOrString8Value(PropertyRow_r returnedRows)
        {
            bool isString8ValueAllowed = false;
            bool isStringValueAllowed = false;
            for (int i = 0; i < returnedRows.LpProps.Length; i++)
            {
                // The first four bytes are the property ID and the last four bytes are the property type.
                // According to MS-OXPROPS, the native type of property PidTagDisplayName is Unicode string.
                // According to MS-OXCDATA, property errors appear in two different contexts. When an error occurs in getting a property of an object or a column of a table 
                // from the server, the type of the returned property value is ErrorCode (0x000A).
                // So the Unicode string or 8-bit character string values can only be verified when the property is returned successfully. 
                if ((returnedRows.LpProps[i].PropTag & 0xFFFF0000) == 0x30010000 && (returnedRows.LpProps[i].PropTag & 0x0000FFFF) != 0x0000000A)
                {
                    // 0x0000001E means that it is a 8-bit character string value.
                    if ((returnedRows.LpProps[i].PropTag & 0x0000FFFF) == 0x0000001E)
                    {
                        isString8ValueAllowed = true;
                    }

                    // 0x0000001F means that it is a Unicode string value.
                    if ((returnedRows.LpProps[i].PropTag & 0x0000FFFF) == 0x0000001F)
                    {
                        isStringValueAllowed = true;
                    }

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R433. isString8ValueAllowed is {0}, isStringValueAllowed is {1}.", isString8ValueAllowed, isStringValueAllowed);

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R433
                    this.Site.CaptureRequirementIfIsTrue(
                        isString8ValueAllowed || isStringValueAllowed,
                        433,
                        @"[In String Handling] The Exchange Server NSPI Protocol allows string values to be represented as 8-bit character strings or Unicode strings.");
                }
            }
        }
        #endregion
    }
}