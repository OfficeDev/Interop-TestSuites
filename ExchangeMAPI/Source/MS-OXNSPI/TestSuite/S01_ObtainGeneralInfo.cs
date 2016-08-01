namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class contains all the test cases designed to test the server behaviors for the NSPI calls 
    /// related to obtaining general information of Address Book Object.
    /// </summary>
    [TestClass]
    public class S01_ObtainGeneralInfo : TestSuiteBase
    {
        #region Test Suite Initialization

        /// <summary>
        /// Initialize the test suite.
        /// </summary>
        /// <param name="testContext">The test context instance.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Reset the test environment.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiBind operation returning success.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S01_TC01_BindSuccess()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind and set dwFlags to 0 to initiate a session between the client and the server.
            STAT stat = new STAT();
            stat.InitiateStat();

            // Set value for serverGuid
            FlatUID_r guid = new FlatUID_r
            {
                Ab = new byte[16]
            };
            FlatUID_r? serverGuid = guid;

            uint flags = 0x0;
            ErrorCodeValue result1 = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result1, "NspiBind should return Success!");
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion

            #region Call NspiBind and set dwFlags to 1 to initiate a session between the client and the server.
            // Set dwFlags to another value rather than fAnonymousLogin (0x20) to check that server ignores dwFlags when it is not fAnonymousLogin.
            flags = 0x1;
            ErrorCodeValue result2 = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result2, "NspiBind should return Success!");

            #region Verify the requirements about NspiBind operation.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1668");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1668
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                result1,
                result2,
                1668,
                @"[In NspiBind] [dwFlags] If the bits are set to different values [other than the bit flag fAnonymousLogin], the server will return the same result.");

            // Convert GUID bytes to GUID string.
            string guidString = new Guid(serverGuid.Value.Ab).ToString("B");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R697: The GUID string is {0}", guidString);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R697
            // If guidString is not null, it means that the byte array in serverGuid can be converted to a GUID string.
            Site.CaptureRequirementIfIsNotNull(
                guidString,
                697,
                @"[In NspiBind] [Server Processing Rules: Upon receiving message NspiBind, the server MUST process the data from the message subject to the following constraints:] [constraint 6] Subject to the prior constraints, if the input parameter pServerGuid is not NULL, the server MUST set the output parameter pServerGuid to a GUID associated with the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R674", "The value of pServerGuid is {0}.", serverGuid);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R674
            this.Site.CaptureRequirementIfIsTrue(
                serverGuid == null || guidString != null,
                674,
                @"[In NspiBind] pServerGuid: The value NULL or a pointer to a GUID value that is associated with the specific server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R700");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R700
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                result1,
                700,
                @"[In NspiBind] [Server Processing Rules: Upon receiving message NspiBind, the server MUST process the data from the message subject to the following constraints:] [constraint 7] If no other return code has been set, the server MUST return the value ""Success"".");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the related requirements about NspiUnbind with different reserved values.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S01_TC02_UnbindSuccess()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between the client and the server.
            STAT stat = new STAT();
            stat.InitiateStat();

            // Set value for serverGuid.
            FlatUID_r guid = new FlatUID_r
            {
                Ab = new byte[16]
            };
            FlatUID_r? serverGuid = guid;

            uint flags = 0x0;
            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");
            #endregion

            #region Call NspiUnbind with the Reserved field set to value 0.
            uint reserved = 0;
            uint returnValue1 = this.ProtocolAdatper.NspiUnbind(reserved);
            #endregion

            #region Call NspiBind to initiate a session between the client and the server.
            stat = new STAT();
            stat.InitiateStat();

            // Set value for serverGuid.
            guid = new FlatUID_r
            {
                Ab = new byte[16]
            };
            serverGuid = guid;

            flags = 0x0;
            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");
            #endregion

            #region Call NspiUnbind with the Reserved field set to a different value.
            reserved = 5;
            uint returnValue2 = this.ProtocolAdatper.NspiUnbind(reserved);

            #region Verify the requirements about NspiUnbind operation.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1673");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1673
            Site.CaptureRequirementIfAreEqual<uint>(
                returnValue1,
                returnValue2,
                1673,
                @"[In NspiUnbind] If this field [Reserved] is set to different values, the server will return the same result.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R712");

            // Whether the context handle is really destroyed is not directly tested by calling a method after the NspiUnbind method is called,
            // since it is not allowed by this protocol and it may cause the instability of the server. Because Windows Global Catalog services have a limitation of 50 concurrent 
            // NSPI connections per user, the context handle is believed to be destroyed successfully if all test cases have passed.
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R712
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                returnValue1,
                712,
                @"[In NspiUnbind] [Server Processing Rules: Upon receiving message NspiUnbind, the server MUST process the data from the message subject to the following constraints:] [constraint 1] If the server successfully destroys the context handle, the server MUST return the value ""UnbindSuccess"", as specified in section 2.2.1.2.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R701");

            // Whether the context handle is really destroyed is not directly tested by calling a method after the NspiUnbind method is called,
            // since it is not allowed by this protocol and it may the instability of the server. Because Windows Global Catalog services have a limitation of 50 concurrent 
            // NSPI connections per user, the context handle is believed to be destroyed successfully if all test cases have passed.
            // Verify MS-OXNSPI requirement: MS-OXNSPI_R701
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                returnValue1,
                701,
                @"[In NspiUnbind] The NspiUnbind method destroys the context handle.");

            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetSpecialTable operation returning success.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S01_TC03_GetSpecialTableSuccess()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between the client and the server.
            uint flags = 0;
            STAT stat = new STAT();
            stat.InitiateStat();
            FlatUID_r? serverGuid = null;

            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind method should return Success.");
            #endregion

            #region Call NspiGetSpecialTable method with dwFlags set to "NspiAddressCreationTemplates".
            uint flagsOfGetSpecialTable = (uint)NspiGetSpecialTableFlags.NspiAddressCreationTemplates;
            uint version = 0;
            stat.InitiateStat();
            stat.CodePage = (uint)RequiredCodePage.CP_TELETEX;
            PropertyRowSet_r? rows;
            this.Result = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetSpecialTable method should return Success.");

            #region Capture code
            // According to section 2.2.1 in [MS-OXOABKT], the table of the available address creation templates consists of these tags:
            // PidTagAddressType, PidTagDisplayName, PidTagDisplayType, PidTagEntryId, PidTagDepth, PidTagSelectabl and PidTagInstanceKey. 
            bool isPidTagAddressTypeExist = false;
            bool isPidTagDisplayNameExist = false;
            bool isPidTagDisplayTypeExist = false;
            bool isPidTagEntryIdExist = false;
            bool isPidTagDepthExist = false;
            bool isPidTagSelectableExist = false;
            bool isPidTagInstanceKeyExist = false;
            bool isAddressCreationTemplatesTable = false;

            // To inquiry each tags in rows
            foreach (PropertyValue_r propertyValue in rows.Value.ARow[0].LpProps)
            {
                Site.Log.Add(LogEntryKind.Debug, "The PropTag value is {0}.", (AulProp)propertyValue.PropTag);
                switch ((AulProp)propertyValue.PropTag)
                {
                    case AulProp.PidTagAddressType:
                        isPidTagAddressTypeExist = true;
                        break;

                    case AulProp.PidTagDisplayName:
                        isPidTagDisplayNameExist = true;
                        break;

                    case AulProp.PidTagDisplayType:
                        isPidTagDisplayTypeExist = true;
                        break;

                    case AulProp.PidTagEntryId:
                        isPidTagEntryIdExist = true;
                        break;

                    case AulProp.PidTagDepth:
                        isPidTagDepthExist = true;
                        break;

                    case AulProp.PidTagSelectable:
                        isPidTagSelectableExist = true;
                        break;

                    case AulProp.PidTagInstanceKey:
                        isPidTagInstanceKeyExist = true;
                        break;

                    default:
                        break;
                }
            }

            // If all these properties exist, the table is an address creation templates table.
            isAddressCreationTemplatesTable = isPidTagAddressTypeExist && isPidTagDisplayNameExist
                                           && isPidTagDisplayTypeExist && isPidTagEntryIdExist
                                           && isPidTagDepthExist && isPidTagSelectableExist
                                           && isPidTagInstanceKeyExist;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R109");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R109
            Site.CaptureRequirementIfIsTrue(
                isAddressCreationTemplatesTable,
                109,
                @"[In NspiGetSpecialTable Flags] NspiAddressCreationTemplates (0x00000002): Specifies that the server MUST return the table of the available address creation templates.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R752", "The rows returned from server is {0}.", rows.Value.CRows);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R752
            // If all the properties listed above are returned, the table is an address creation templates table.
            // If there are rows returned from server, the client is requesting the rows of an address creation table.
            this.Site.CaptureRequirementIfIsTrue(
                isAddressCreationTemplatesTable && rows.Value.CRows != 0,
                752,
                @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 9] If the input parameter dwFlags contains the value ""NspiAddressCreationTemplates"", the client is requesting the rows of an address creation table, as specified in section 3.1.4.4.3.2.");

            bool isDT_ADDRESS_TEMPLATETypeCorrect = false;
            PermanentEntryID perEntryId = new PermanentEntryID();

            foreach (PropertyRow_r row in rows.Value.ARow)
            {
                foreach (PropertyValue_r val in row.LpProps)
                {
                    if (val.PropTag == (uint)AulProp.PidTagEntryId)
                    {
                        perEntryId = AdapterHelper.ParsePermanentEntryIDFromBytes(val.Value.Bin.Lpb);
                        break;
                    }
                }

                if (perEntryId.DisplayTypeString == DisplayTypeValue.DT_ADDRESS_TEMPLATE)
                {
                    isDT_ADDRESS_TEMPLATETypeCorrect = true;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R60");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R60
            Site.CaptureRequirementIfIsTrue(
                isDT_ADDRESS_TEMPLATETypeCorrect,
                60,
                @"[In Display Type Values] DT_ADDRESS_TEMPLATE display type with 0x00000102 value means An address creation template.");
            #endregion
            #endregion

            #region Call NspiGetSpecialTable method with TemplateLocale field of stat set to a value which does not maintain an address creation table.
            stat.TemplateLocale = 0;
            flagsOfGetSpecialTable = (uint)NspiGetSpecialTableFlags.NspiAddressCreationTemplates;
            this.Result = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows);

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R757: The value of CRows is {0}.", rows.Value.CRows);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R757
            Site.CaptureRequirementIfIsTrue(
                (rows.Value.CRows == 0) && (rows.Value.ARow == null),
                757,
                @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 11] If the server does not maintain an address creation table for that LCID [the LCID specified by TemplateLocale field of the input parameter pStat], the server MUST proceed as if it [server] maintained an address creation table with no rows for that LCID [the LCID specified by TemplateLocale field of the input parameter pStat].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R758");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R758
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                758,
                @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 11] That is, the server MUST NOT return an error code if it [server] does not maintain an address creation table for that LCID.");
            #endregion
            #endregion

            #region Call NspiGetSpecialTable method with dwFlags set to "NspiUnicodeStrings".
            flagsOfGetSpecialTable = (uint)NspiGetSpecialTableFlags.NspiUnicodeStrings;
            stat.TemplateLocale = (uint)DefaultLCID.NSPI_DEFAULT_LOCALE;
            stat.CodePage = (uint)RequiredCodePage.CP_TELETEX;
            this.Result = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetSpecialTable should return Success.");

            #region Capture code

            for (int i = 0; i < rows.Value.CRows; i++)
            {
                PropertyValue_r[] propertyValue = rows.Value.ARow[i].LpProps;
                bool isStringValueFound = false;

                // The first four bytes are the property ID and the last four bytes are the property type.
                // Property ID 0x3001 indicates that it is a property PidTagDisplayName, and property type 0x0000001F indicates the property is a string.
                // According to MS-OXPROPS, the native type of property PidTagDisplayName is a string.
                for (int j = 0; j < propertyValue.Length; j++)
                {
                    if ((propertyValue[j].PropTag & 0xFFFF0000) == 0x30010000)
                    {
                        isStringValueFound = true;

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1966");

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R1966
                        // 0x0000001F means that the property type is string.
                        this.Site.CaptureRequirementIfAreEqual<uint>(
                            0x0000001F,
                            propertyValue[j].PropTag & 0x0000FFFF,
                            1966,
                            @"[In Conversion Rules for String Values Specified by the Server to the Client] In NspiGetSpecialTable method, String values can be returned in Unicode representation in the output parameter ppRows.");

                        break;
                    }
                }

                if (!isStringValueFound)
                {
                    Site.Assert.Fail("Property PidTagDisplayName with display type string is not found.");
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R743");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R743
            // In this step, dwFlags is set to NspiUnicodeStrings to request the rows of the server's address book hierarchy table.
            // So if the returned rows are not null, it specifies that server returns the rows of the address book hierarchy table.
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                rows.Value.CRows,
                743,
                @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] If the input parameter dwFlags does not contain the value ""NspiAddressCreationTemplates"", the client is requesting the rows of the server's address book hierarchy table, as specified in section 3.1.4.4.3.1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R111");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R111
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.IsPtypString(rows.Value),
                111,
                @"[In NspiGetSpecialTable Flags] NspiUnicodeStrings (0x00000004): Specifies that the server MUST return all strings as Unicode representations rather than as multibyte strings in the client's code page.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R56");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R56
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.IsDT_CONTAINERTypeCorrect(rows),
                56,
                @"[In Display Type Values] DT_CONTAINER display type with 0x00000100 value means An address book hierarchy table container.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R760");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R760
            // If isVerifyR760 is true, the property type of the string type is PtypString other than PtypString8.
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.IsPtypString(rows.Value),
                760,
                @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 12] If the input parameter dwFlags contains the value ""NspiUnicodeStrings"" and the client is requesting the rows of the server's address book hierarchy table, the server MUST express string-valued properties in the table as Unicode values, as specified in section 3.1.4.3.");
            #endregion
            #endregion

            #region Call NspiGetSpecialTable method with dwFlags that does not contain the value "NspiUnicodeStrings (0x04)" and CodePage field of the stat that does not contain the value "CP_WINUNICODE (0x04B0)".
            flagsOfGetSpecialTable = 0;
            stat.TemplateLocale = (uint)DefaultLCID.NSPI_DEFAULT_LOCALE;
            stat.CodePage = (uint)RequiredCodePage.CP_TELETEX;
            this.Result = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows);

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R762");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R762
            // If isVerifyR762 is true, all property type of string is PtypString8 other than PtypString.
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.IsPtypString8(rows.Value),
                762,
                @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 13] If the input parameter dwFlags does not contain the value ""NspiUnicodeStrings"" and the client is requesting the rows of the server's hierarchy table, and the CodePage field of the input parameter pStat does not contain the value CP_WINUNICODE, the server MUST express string-valued properties as 8-bit strings in the code page specified by the field CodePage in the input parameter pStat.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R765");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R765
            // The return rows is not null means that field rows contains the rows of the table
            Site.CaptureRequirementIfIsNotNull(
                rows,
                765,
                @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 17] Subject to the prior constraints, the server returns the rows of the table requested by the client in the output parameter ppRows.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R767");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R767
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                767,
                @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 18] If no error condition has been specified by the previous constraints, the server MUST return the value ""Success"".");
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint reserved = 0;
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiQueryColumns operation returning success. 
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S01_TC04_QueryColumnsSuccess()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between the client and the server.
            uint flags = 0;
            STAT stat = new STAT();
            stat.InitiateStat();

            FlatUID_r guid = new FlatUID_r
            {
                Ab = new byte[16]
            };

            FlatUID_r? serverGuid = guid;

            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");
            #endregion

            #region Call NspiQueryColumns method with dwFlags that does not contain the value "NspiUnicodeProptypes (0x80000000)".
            uint reservedOfQueryColumns = 0;
            uint flagsOfQueryColumns = 0;
            PropertyTagArray_r? columns;
            this.Result = this.ProtocolAdatper.NspiQueryColumns(reservedOfQueryColumns, flagsOfQueryColumns, out columns);

            #region Capture code
            bool isString8Type = false;
            foreach (uint tag in columns.Value.AulPropTag)
            {
                if ((PropertyTypeValue)(tag & 0x0000ffff) != PropertyTypeValue.PtypMultipleString && ((PropertyTypeValue)(tag & 0x0000ffff) != PropertyTypeValue.PtypString))
                {
                    isString8Type = true;
                }
                else
                {
                    isString8Type = false;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R822");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R822
            Site.CaptureRequirementIfIsTrue(
                isString8Type,
                822,
                @"[In NspiQueryColumns] [Server Processing Rules: Upon receiving message NspiQueryColumns, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] 
                If the input parameter dwFlags does not contain the bit flag NspiUnicodeProptypes, the server MUST report the property type of all string valued properties as PtypString8.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R116");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R116
            Site.CaptureRequirementIfIsTrue(
                isString8Type,
                116,
                @"[In NspiQueryColumns Flag] If the NspiUnicodeProptypes flag is not set, the server MUST return all proptags specifying values with string representations as having the PtypString8 property type. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R827");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R827
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                827,
                @"[In NspiQueryColumns] [Server Processing Rules: Upon receiving message NspiQueryColumns, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] If no error condition has been specified by the previous constraints, the server MUST return the value ""Success"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R802");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R802
            this.Site.CaptureRequirementIfIsInstanceOfType(
                columns,
                typeof(PropertyTagArray_r),
                802,
                @"[In NspiQueryColumns] It [NspiQueryColumns] returns this list [a list of all the properties that the server is aware of] as an array of proptags.");

            #endregion
            #endregion

            #region Call NspiGetPropList to get all the properties.
            // The Minimal Entry ID 0 specifies the default global address book object.
            uint mid = 0;
            uint flagsOfGetPropList = (uint)RetrievePropertyFlag.fEphID;
            PropertyTagArray_r? propTagsOfGetPropList;
            uint codePage = (uint)RequiredCodePage.CP_TELETEX;

            this.Result = this.ProtocolAdatper.NspiGetPropList(flagsOfGetPropList, mid, codePage, out propTagsOfGetPropList);
            Site.Assert.IsNotNull(propTagsOfGetPropList.Value.AulPropTag, "The output property list should not be empty. The property number is {0}.", propTagsOfGetPropList == null ? 0 : propTagsOfGetPropList.Value.CValues);

            #region Capture code

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R801");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R801
            // According to the definition of NspiGetPropList, NspiGetPropList will return all properties that the server is aware of. 
            // So if the property number returned is equal to the number of method NspiGetPropList, it specifies method NspiQueryColumns returns a list of all properties that the server is aware of.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                propTagsOfGetPropList.Value.CValues,
                columns.Value.CValues,
                801,
                @"[In NspiQueryColumns] The NspiQueryColumns method returns a list of all the properties that the server is aware of.");
            #endregion

            #endregion

            #region Call NspiQueryColumns method with dwFlags that contains the value "NspiUnicodeProptypes".
            reservedOfQueryColumns = 0;
            flagsOfQueryColumns = (uint)NspiQueryColumnsFlag.NspiUnicodeProptypes;
            this.Result = this.ProtocolAdatper.NspiQueryColumns(reservedOfQueryColumns, flagsOfQueryColumns, out columns);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryColumns should return Success.");

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R820");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R820
            bool isStringType = false;
            foreach (uint tag in columns.Value.AulPropTag)
            {
                if ((PropertyTypeValue)(tag & 0x0000ffff) != PropertyTypeValue.PtypMultipleString8 || (PropertyTypeValue)(tag & 0x0000ffff) != PropertyTypeValue.PtypString8)
                {
                    isStringType = true;
                }
                else
                {
                    isStringType = false;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isStringType,
                820,
                @"[In NspiQueryColumns] [Server Processing Rules: Upon receiving message NspiQueryColumns, the server MUST process the data from the message subject to the following constraints:]
                [Constraint 3] If the input parameter dwFlags contains the bit flag NspiUnicodeProptypes, then the server MUST report the property type of all string valued properties as PtypString.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R115");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R115
            Site.CaptureRequirementIfIsTrue(
                isStringType,
                115,
                @"[In NspiQueryColumns Flag] NspiUnicodeProptypes (0x80000000): Specifies that the server MUST return all proptags that specify values with string representations as having the PtypString property type.");
            #endregion

            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint reserved = 0;
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetPropsList operation returning success. 
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S01_TC05_GetPropsListSuccess()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between the client and the server.
            uint flags = 0;
            STAT stat = new STAT();
            stat.InitiateStat();

            FlatUID_r guid = new FlatUID_r
            {
                Ab = new byte[16]
            };

            FlatUID_r? serverGuid = guid;

            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");
            #endregion

            #region Call NspiUpdateStat to update the STAT block to make CurrentRec point to the first row of the table.
            uint reserved = 0;
            int? delta = 1;
            this.Result = this.ProtocolAdatper.NspiUpdateStat(reserved, ref stat, ref delta);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUpdateStat should return Success!");
            #endregion

            #region Call NspiGetPropList method with dwFlags set to fEphID and CodePage field of stat that is not CP_WINUNICODE (0x000004B0).
            uint flagsOfGetPropList = (uint)RetrievePropertyFlag.fEphID;
            PropertyTagArray_r? propTagsOfGetPropList;
            uint codePage = (uint)RequiredCodePage.CP_TELETEX;

            // The Minimal Entry ID 0 specifies the default global address book object.
            uint mid = stat.CurrentRec;

            this.Result = this.ProtocolAdatper.NspiGetPropList(flagsOfGetPropList, mid, codePage, out propTagsOfGetPropList);
            Site.Assert.IsNotNull(propTagsOfGetPropList.Value.AulPropTag, "The output property list should not be empty. The property number is {0}.", propTagsOfGetPropList == null ? 0 : propTagsOfGetPropList.Value.CValues);

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R858");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R858
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                858,
                @"[In NspiGetPropList] [Server Processing Rules: Upon receiving message NspiGetPropList, the server MUST process the data from the message subject to the following constraints:] [Constraint 6] If no error condition has been specified by the previous constraints, the server MUST return the value ""Success"".");

            bool isString8Type = false;
            foreach (uint propTag in propTagsOfGetPropList.Value.AulPropTag)
            {
                if ((PropertyTypeValue)(propTag & 0x0000ffff) != PropertyTypeValue.PtypMultipleString || (PropertyTypeValue)(propTag & 0x0000ffff) != PropertyTypeValue.PtypString)
                {
                    isString8Type = true;
                }
                else
                {
                    isString8Type = false;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R852");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R852
            Site.CaptureRequirementIfIsTrue(
                isString8Type,
                852,
                @"[In NspiGetPropList] [Server Processing Rules: Upon receiving message NspiGetPropList, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] The server MUST return all string valued properties as having the PtypString8 property type.");
            #endregion
            #endregion

            #region Call NspiGetPropList method with dwFlags set to another value rather than fEphID (0x2) and fSkipObjects (0x1).
            flagsOfGetPropList = 4;
            PropertyTagArray_r? propTagsOfGetPropList1;
            this.Result = this.ProtocolAdatper.NspiGetPropList(flagsOfGetPropList, mid, codePage, out propTagsOfGetPropList1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetPropList should return Success.");

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1924");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1924
            this.Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyTagArrayEqual(propTagsOfGetPropList, propTagsOfGetPropList1),
                1924,
                @"[In NspiGetPropList] If dwFlags is set to different values other than the bit flag fSkipObjects, server will return the same result.");
            #endregion
            #endregion

            #region Call NspiGetPropList method with dwFlags set to fSkipObjects.
            flagsOfGetPropList = (uint)RetrievePropertyFlag.fSkipObjects;
            PropertyTagArray_r? propTagsOfGetPropList2;
            this.Result = this.ProtocolAdatper.NspiGetPropList(flagsOfGetPropList, stat.CurrentRec, stat.CodePage, out propTagsOfGetPropList2);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetPropList should return Success!");

            #region Capture code
            bool isNoPtypEmbeddedTableExist = true;

            // To find PtypEmbeddedTable type in propTagsOfGetPropList.
            if (propTagsOfGetPropList2 != null && propTagsOfGetPropList2.Value.CValues != 0)
            {
                foreach (uint tag in propTagsOfGetPropList2.Value.AulPropTag)
                {
                    if ((PropertyTypeValue)(tag & 0xffff) == PropertyTypeValue.PtypEmbeddedTable)
                    {
                        isNoPtypEmbeddedTableExist = false;
                        break;
                    }
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R848");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R848
            Site.CaptureRequirementIfIsTrue(
                isNoPtypEmbeddedTableExist,
                848,
                @"[In NspiGetPropList] [Server Processing Rules: Upon receiving message NspiGetPropList, the server MUST process the data from the message subject to the following constraints:] [Constraint 3] If the input parameter dwFlags contains the bit flag fSkipObjects, the server MUST NOT return any proptags with the PtypEmbeddedTable property type in the output parameter ppPropTags.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R104");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R104
            Site.CaptureRequirementIfIsTrue(
                isNoPtypEmbeddedTableExist,
                104,
                @"[In Retrieve Property Flags] fSkipObjects (0x00000001): Client requires that the server MUST NOT include proptags with the PtypEmbeddedTable property type in any lists of proptags that the server creates on behalf of the client.");
            #endregion
            #endregion

            #region Call NspiQueryColumns to get a list of all the properties that the server is aware of.
            uint reservedOfQueryColumns = 0;
            uint flagsOfQueryColumns = 0;
            PropertyTagArray_r? columns;
            this.Result = this.ProtocolAdatper.NspiQueryColumns(reservedOfQueryColumns, flagsOfQueryColumns, out columns);
            #endregion

            #region Call NspiGetPropList method with dwMId set to 0 to get all the properties.
            // The Minimal Entry ID 0 specifies the default global address book object.
            mid = 0;
            this.Result = this.ProtocolAdatper.NspiGetPropList(flagsOfGetPropList, mid, codePage, out propTagsOfGetPropList);
            Site.Assert.IsNotNull(propTagsOfGetPropList.Value.AulPropTag, "The output property list should not be empty. The property number is {0}.", propTagsOfGetPropList == null ? 0 : propTagsOfGetPropList.Value.CValues);

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R828");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R828
            // According to the definition of NspiQueryColumns, NspiQueryColumns will return all properties that the server is aware of. 
            // So if the property number returned is equal to the number of method NspiQueryColumns, it specifies method NspiGetPropList returns a list of all properties that the server is aware of.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                columns.Value.CValues,
                propTagsOfGetPropList.Value.CValues,
                828,
                @"[In NspiGetPropList] The NspiGetPropList method returns a list of all the properties that have values on a specified object.");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R854. The CurrentRec value is {0}. The mid value is {1}. The returned properties number when dwMId field set to 0 is {2}. The returned properties number when dwMId field set to stat.CurrentRec is {3}.",
                stat.CurrentRec,
                mid,
                propTagsOfGetPropList.Value.CValues,
                propTagsOfGetPropList2.Value.CValues);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R854
            // Since the returned properties number is different for different dwMId values, so MS-OXNSPI_R854 can be verified.
            this.Site.CaptureRequirementIfIsTrue(
                stat.CurrentRec != mid && propTagsOfGetPropList.Value.CValues != propTagsOfGetPropList2.Value.CValues,
                854,
                @"[In NspiGetPropList] [Server Processing Rules: Upon receiving message NspiGetPropList, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] Subject to the previous constraints, the server constructs a list of all proptags that correspond to values on the object specified in the input parameter dwMId.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R855.");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R855
            // Since the returned properties number is different for different dwMId values, so MS-OXNSPI_R855 can be verified.
            this.Site.CaptureRequirementIfIsTrue(
                stat.CurrentRec != mid && propTagsOfGetPropList.Value.CValues != propTagsOfGetPropList2.Value.CValues,
                855,
                @"[In NspiGetPropList] [Server Processing Rules: Upon receiving message NspiGetPropList, the server MUST process the data from the message subject to the following constraints:] [Constraint 5] The server MUST return this list [a list of all proptags that correspond to values on the object specified in the input parameter dwMId] in the output parameter ppPropTags.");
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to NspiGetTemplateInfo operation with different template types.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S01_TC06_GetTemplateInfoSuccessWithNullDN()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between the client and the server.
            uint flags = 0;

            STAT stat = new STAT();
            stat.InitiateStat();

            FlatUID_r guid = new FlatUID_r
            {
                Ab = new byte[16]
            };

            FlatUID_r? serverGuid = guid;

            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");
            #endregion

            #region Call NspiGetTemplateInfo method with dwFlags set to TI_TEMPLATE.
            uint flagsOfGetTemplateInfo = (uint)NspiGetTemplateInfoFlag.TI_TEMPLATE;
            uint type = (uint)DisplayTypeValue.DT_MAILUSER;
            string dn = null;
            uint codePage = stat.CodePage;
            uint locateID = stat.TemplateLocale;
            PropertyRow_r? data;

            this.Result = this.ProtocolAdatper.NspiGetTemplateInfo(flagsOfGetTemplateInfo, type, dn, codePage, locateID, out data);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1492");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1492
            Site.CaptureRequirementIfAreEqual<ErrorCodeValue>(
                ErrorCodeValue.Success,
                this.Result,
                1492,
                @"[In NspiGetTemplateInfo] [Server Processing Rules: Upon receiving message NspiGetTemplateInfo, the server MUST process the data from the message subject to the following constraints:] [Constraint 8] If no other return values have been specified by these constraints [constraints 1-7], the server MUST return the return value ""Success"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R120");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R120
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagTemplateData,
                data.Value.LpProps[0].PropTag,
                120,
                @"[In NspiGetTemplateInfo Flags] TI_TEMPLATE (0x00000001): Specifies that the server is to return the value that represents a template.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1444");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1444
            // If data.Value.LpProps[0].PropTag is PidTagTemplateData, it specifies that server returns information about template.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagTemplateData,
                data.Value.LpProps[0].PropTag,
                1444,
                @"[In NspiGetTemplateInfo] The NspiGetTemplateInfo method returns information about template objects in the address book.");

            #endregion Capture

            #endregion

            #region Call NspiGetTemplateInfo method with dwFlags set to TI_SCRIPT.
            flagsOfGetTemplateInfo = (uint)NspiGetTemplateInfoFlag.TI_SCRIPT;
            this.Result = this.ProtocolAdatper.NspiGetTemplateInfo(flagsOfGetTemplateInfo, type, dn, codePage, locateID, out data);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetTemplateInfo method should return Success.");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R121");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R121
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)AulProp.PidTagScriptData,
                data.Value.LpProps[0].PropTag,
                121,
                @"[In NspiGetTemplateInfo Flags] TI_SCRIPT (0x00000004): Specifies that the server is to return the value of the script that is associated with a template.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the related requirements about NspiGetTemplateInfo operation with non-null DN. 
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S01_TC07_GetTemplateInfoSuccessWithNotNullDN()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between the client and the server.
            uint flags = 0;

            STAT stat = new STAT();
            stat.InitiateStat();

            FlatUID_r guid = new FlatUID_r
            {
                Ab = new byte[16]
            };

            FlatUID_r? serverGuid = guid;

            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");
            #endregion

            #region Call NspiGetSpecialTable to get the template DN.
            uint flagsOfGetSpecialTable = (uint)NspiGetSpecialTableFlags.NspiAddressCreationTemplates;
            uint version = 0;
            stat.InitiateStat();
            stat.CodePage = (uint)RequiredCodePage.CP_TELETEX;
            PropertyRowSet_r? rows;
            this.Result = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetSpecialTable method should return Success.");

            // Parse and record the template DN.
            string dn = string.Empty;

            foreach (PropertyRow_r row in rows.Value.ARow)
            {
                PropertyValue_r[] propertyValue = row.LpProps;
                for (int i = 0; i < propertyValue.Length; i++)
                {
                    if (propertyValue[i].PropTag == 0x0FFF0102)
                    {
                        PermanentEntryID permanentEntryID = AdapterHelper.ParsePermanentEntryIDFromBytes(propertyValue[i].Value.Bin.Lpb);
                        dn = permanentEntryID.DistinguishedName;

                        break;
                    }
                }

                if (!string.IsNullOrEmpty(dn))
                {
                    break;
                }
            }
            #endregion

            #region Capture
            if (Common.IsRequirementEnabled(2003004, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R2003004");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R2003004
                Site.CaptureRequirementIfIsTrue(
                    Common.IsDNMatchABNF(dn, DNFormat.Dn),
                    2003004,
                    @"[In Appendix A: Product Behavior] Implementation does follow the ABNF format. (Microsoft Exchange Server 2010 Service Pack 3 (SP3) follows this behavior).");
            }
            #endregion Capture

            #region Call NspiGetTemplateInfo method with DN set to a non-null value.
            uint flagsOfGetTemplateInfo = (uint)NspiGetTemplateInfoFlag.TI_SCRIPT;
            uint type = (uint)DisplayTypeValue.DT_MAILUSER;
            uint codePage = (uint)RequiredCodePage.CP_TELETEX;
            uint locateID = stat.TemplateLocale;
            PropertyRow_r? data;

            this.Result = this.ProtocolAdatper.NspiGetTemplateInfo(flagsOfGetTemplateInfo, type, dn, codePage, locateID, out data);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetTemplateInfo method should return Success.");
            #endregion

            if (this.Transport == "ncacn_http" || this.Transport == "ncacn_ip_tcp")
            {
                #region Call NspiGetTemplateInfo method with locateID set to a different value with the above step.
                PropertyRow_r? data1;
                type = (uint)DisplayTypeValue.DT_CONTAINER;
                locateID = 0x0000040c;
                this.Result = this.ProtocolAdatper.NspiGetTemplateInfo(flagsOfGetTemplateInfo, type, dn, codePage, locateID, out data1);
                Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetTemplateInfo method should return Success.");

                #region Capture

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1983");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1983
                this.Site.CaptureRequirementIfIsTrue(
                    AdapterHelper.AreTwoPropertyRowEqual(data, data1),
                    1983,
                    @"[In NspiGetTemplateInfo] If the input parameter pDN is not NULL and input parameter dwLocaleID is set to different values, the server will return the same result.");

                #endregion Capture
                #endregion

                #region Call NspiGetTemplateInfo method with ulType set to a different value with the above step.
                type = (uint)DisplayTypeValue.DT_ADDRESS_TEMPLATE;
                PropertyRow_r? data2;
                this.Result = this.ProtocolAdatper.NspiGetTemplateInfo(flagsOfGetTemplateInfo, type, dn, codePage, locateID, out data2);
                Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetTemplateInfo method should return Success.");

                #region Capture code
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1982");

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1982
                this.Site.CaptureRequirementIfIsTrue(
                    AdapterHelper.AreTwoPropertyRowEqual(data1, data2),
                    1982,
                    @"[In NspiGetTemplateInfo] If the input parameter pDN is not NULL and input parameter ulType is set to different values, the server will return the same result.");
                #endregion
                #endregion
            }

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify if the bits other than the bit flags TI_HELPFILE_NAME, TI_HELPFILE_CONTENTS, TI_SCRIPT, 
        /// TI_TEMPLATE, and TI_EMT are set in different values, the server will return the same value.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S01_TC08_GetTemplateInfoIgnoreSomeFlags()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between the client and the server.
            uint flags = 0;

            STAT stat = new STAT();
            stat.InitiateStat();

            FlatUID_r guid = new FlatUID_r
            {
                Ab = new byte[16]
            };

            FlatUID_r? serverGuid = guid;

            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");
            #endregion

            #region Call NspiGetTemplateInfo method with dwFlags containing TI_HELPFILE_NAME, TI_HELPFILE_CONTENTS, TI_SCRIPT, TI_TEMPLATE and TI_EMT.
            uint flagsOfGetTemplateInfo = (uint)(NspiGetTemplateInfoFlag.TI_TEMPLATE
                | NspiGetTemplateInfoFlag.TI_EMT
                | NspiGetTemplateInfoFlag.TI_HELPFILE_CONTENTS
                | NspiGetTemplateInfoFlag.TI_HELPFILE_NAME
                | NspiGetTemplateInfoFlag.TI_SCRIPT);
            uint type = (uint)DisplayTypeValue.DT_MAILUSER;
            string dn = null;
            uint codePage = stat.CodePage;
            uint locateID = stat.TemplateLocale;
            PropertyRow_r? data;

            this.Result = this.ProtocolAdatper.NspiGetTemplateInfo(flagsOfGetTemplateInfo, type, dn, codePage, locateID, out data);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetTemplateInfo should return Success!");
            #endregion

            #region Call NspiGetTemplateInfo method with dwFlags set to the value other than TI_HELPFILE_NAME (0x20), TI_HELPFILE_CONTENTS (0x40), TI_SCRIPT (0x4), TI_TEMPLATE (0x1) and TI_EMT (0x10).
            PropertyRow_r? data1;
            flagsOfGetTemplateInfo = uint.Parse(Constants.UnrecognizedMID);
            this.Result = this.ProtocolAdatper.NspiGetTemplateInfo(flagsOfGetTemplateInfo, type, dn, codePage, locateID, out data1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetTemplateInfo should return Success!");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1692");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1692
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyRowEqual(data, data1),
                1692,
                @"[In NspiGetTemplateInfo] [dwFlags] If the bits are set to different values other than the bit flags TI_SCRIPT and TI_TEMPLATE, the server will return the same value.");

            #endregion Capture
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify that when calling NspiQueryColumns method, the values of some flags are ignored by server. 
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S01_TC09_QueryColumnsIgnoreSomeFlags()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between the client and the server.
            uint flags = 0;

            STAT stat = new STAT();
            stat.InitiateStat();

            FlatUID_r guid = new FlatUID_r
            {
                Ab = new byte[16]
            };

            FlatUID_r? serverGuid = guid;

            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind should return Success!");
            #endregion

            #region Call NspiQueryColumns method with Reserved set to 0.
            uint reservedOfQueryColumns = 0;
            uint flagsOfQueryColumns = 0;
            PropertyTagArray_r? columns1;
            this.Result = this.ProtocolAdatper.NspiQueryColumns(reservedOfQueryColumns, flagsOfQueryColumns, out columns1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryColumns should return Success.");
            #endregion

            #region Call NspiQueryColumns method again with Reserved set to 1.
            reservedOfQueryColumns = 1;
            PropertyTagArray_r? columns2;
            this.Result = this.ProtocolAdatper.NspiQueryColumns(reservedOfQueryColumns, flagsOfQueryColumns, out columns2);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryColumns should return Success.");

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1922");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1922
            this.Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyTagArrayEqual(columns1, columns2),
                1922,
                @"[In NspiQueryColumns] If this field [Reserved] is set to different values, the server will return the same result.");
            #endregion
            #endregion

            #region Call NspiQueryColumns method with dwFlags set to another value rather than NspiUnicodeProptypes (0x80000000) and 0 (which is used in step 2).
            flagsOfQueryColumns = 1;
            PropertyTagArray_r? columns3;
            this.Result = this.ProtocolAdatper.NspiQueryColumns(reservedOfQueryColumns, flagsOfQueryColumns, out columns3);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiQueryColumns should return Success.");

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1923");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1923
            this.Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyTagArrayEqual(columns1, columns3),
                1923,
                @"[In NspiQueryColumns] If dwFlags is set to different values other than the bit flag NspiUnicodeProptypes, server will return the same result.");
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint returnValue = this.ProtocolAdatper.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify string conversion rules for method NspiGetSpecialTable. 
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S01_TC10_ConvertStringForGetSpecialTable()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between the client and the server.
            uint flags = 0;
            STAT stat = new STAT();
            stat.InitiateStat();
            FlatUID_r? serverGuid = null;

            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiUnbind method should return Success.");
            #endregion

            #region Call NspiGetSpecialTable method with dwFlags set to NspiUnicodeStrings and CodePage set to CP_WINUNICODE to require the string type property to be returned as string type.
            uint flagsOfGetSpecialTable = (uint)NspiGetSpecialTableFlags.NspiUnicodeStrings;
            uint version = 0;
            PropertyRowSet_r? rows;
            stat.TemplateLocale = (uint)DefaultLCID.NSPI_DEFAULT_LOCALE;
            stat.CodePage = (uint)RequiredCodePage.CP_WINUNICODE;
            this.Result = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetSpecialTable should return value Success.");

            #region Capture code
            for (int i = 0; i < rows.Value.CRows; i++)
            {
                PropertyValue_r[] propertyValue = rows.Value.ARow[i].LpProps;

                // The first four bytes are the property ID and the last four bytes are the property type.
                // Property ID 0x3001 indicates it is property PidTagDisplayName, and property type 0x0000001F indicates string type.
                // According to MS-OXPROPS, the native type of property PidTagDisplayName is string.
                for (int j = 0; j < propertyValue.Length; j++)
                {
                    if ((propertyValue[j].PropTag & 0xffff0000) == 0x30010000)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1941");

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R1941
                        this.Site.CaptureRequirementIfAreEqual<uint>(
                            0x0000001F,
                            propertyValue[j].PropTag & 0x0000FFFF,
                            1941,
                            @"[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiGetSpecialTable] If the native type of a property is PtypString and the client has requested that property with the type PtypString, the server MUST return the Unicode representation unmodified.");
                        break;
                    }
                }
            }
            #endregion
            #endregion

            #region Call NspiGetSpecialTable method with dwFlags that doesn't contain NspiUnicodeStrings (0x4) and CodePage set to CP_TELETEX to require the string type property to be returned as PtypString8 type.
            flagsOfGetSpecialTable = (uint)NspiGetSpecialTableFlags.None;
            stat.CodePage = (uint)RequiredCodePage.CP_TELETEX;
            this.Result = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetSpecialTable should return value Success.");

            #region Capture code
            for (int i = 0; i < rows.Value.CRows; i++)
            {
                PropertyValue_r[] propertyValue = rows.Value.ARow[i].LpProps;

                // The first four bytes are the property ID and the last four bytes are the property type.
                // Property ID 0x3001 indicates it is property PidTagDisplayName, and property type 0x0000001E indicates PtypString8 type.
                // According to MS-OXPROPS, the native type of property PidTagDisplayName is string.
                for (int j = 0; j < propertyValue.Length; j++)
                {
                    if ((propertyValue[j].PropTag & 0xFFFF0000) == 0x30010000)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1933");

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R1933
                        this.Site.CaptureRequirementIfAreEqual<uint>(
                            0x0000001E,
                            propertyValue[j].PropTag & 0x0000FFFF,
                            1933,
                            @"[In Conversion Rules for String Values Specified by the Server to the Client] [For method NspiGetSpecialTable] If the native type of a property is PtypString and the client has requested that property with the type PtypString8, the server MUST convert the Unicode representation to an 8-bit character representation in the code page specified by the CodePage field of the pStat parameter prior to returning the value.");
                        break;
                    }
                }
            }
            #endregion
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint reserved = 0;
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the requirements related to different display types of NspiGetSpecialTable operation.
        /// </summary>
        [TestCategory("MSOXNSPI"), TestMethod()]
        public void MSOXNSPI_S01_TC11_GetSpecialTableToGetDifferentDisplayType()
        {
            this.CheckProductSupported();
            this.CheckMAPIHTTPTransportSupported();

            #region Call NspiBind to initiate a session between the client and the server.
            uint flags = 0;
            STAT stat = new STAT();
            stat.InitiateStat();
            FlatUID_r? serverGuid = null;

            this.Result = this.ProtocolAdatper.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiBind method should return Success.");
            #endregion

            #region Call NspiGetSpecialTable method with dwFlags set to a value which doesn't contain the flag NspiUnicodeStrings (0x04) and NspiAddressCreationTemplates (0x02).
            uint version = 0;
            PropertyRowSet_r? rows;
            uint flagsOfGetSpecialTable = 0;

            ErrorCodeValue result1 = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result1, "NspiGetSpecialTable method should return Success.");

            #region Capture code
            // Check whether server returns the display type DT_ADDRESS_TEMPLATE as part of the EntryID of an object.
            bool addressTemplateReturnedInPidTagEntryId = AdapterHelper.CheckIfSpecificDisplayTypeExists(rows.Value.ARow, DisplayTypeValue.DT_ADDRESS_TEMPLATE);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1997",
                "Display type DT_ADDRESS_TEMPLATE {0} returned as part of an EntryID of an object if the table is not the Address Creation Table.",
                addressTemplateReturnedInPidTagEntryId ? "is" : "is not");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1997
            // NspiAddressCreationTemplates isn't contained in dwFlags according to step 3, so MS-OXNSPI_R1997 can be verified as follows.
            Site.CaptureRequirementIfIsFalse(
                addressTemplateReturnedInPidTagEntryId,
                1997,
                @"[In Display Type Values] Exchange NSPI server will not return display type DT_ADDRESS_TEMPLATE as part of an EntryID of an object if the table is not the Address Creation Table.");

            // Check whether server returns the display type DT_CONTAINER as part of the EntryID of an object.
            bool containerReturnedInPidTagEntryId = AdapterHelper.CheckIfSpecificDisplayTypeExists(rows.Value.ARow, DisplayTypeValue.DT_CONTAINER);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1998", "Display type DT_CONTAINER {0} returned as part of an EntryID of an object in the address book hierarchy table.", containerReturnedInPidTagEntryId ? "is" : "is not");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1998
            Site.CaptureRequirementIfIsTrue(
                containerReturnedInPidTagEntryId,
                1998,
                @"[In Display Type Values] Exchange NSPI server will return display type DT_CONTAINER as part of an EntryID of an object in the address book hierarchy table.");
            #endregion
            #endregion

            #region Call NspiGetSpecialTable method with dwFlags set to another value which doesn't contain the flag NspiUnicodeStrings (0x04) and NspiAddressCreationTemplates (0x02).
            PropertyRowSet_r? rows1;
            flagsOfGetSpecialTable = 1;
            ErrorCodeValue result2 = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows1);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result2, "NspiGetSpecialTable method should return Success.");

            #region Capture code

            for (int i = 0; i < rows.Value.CRows; i++)
            {
                PropertyValue_r[] propertyValue = rows.Value.ARow[i].LpProps;
                bool isString8ValueFound = false;

                // The first four bytes are the property ID and the last four bytes are the property type.
                // Property ID 0x3001 indicates that it is a PidTagDisplayName property, and property type 0x0000001E indicates that the property type is PtypString8.
                // According to MS-OXPROPS, the native type of property PidTagDisplayName is a string.
                for (int j = 0; j < propertyValue.Length; j++)
                {
                    if ((propertyValue[j].PropTag & 0xFFFF0000) == 0x30010000)
                    {
                        isString8ValueFound = true;

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1967");

                        // Verify MS-OXNSPI requirement: MS-OXNSPI_R1967
                        // 0x0000001E means that the property type is PtypString8.
                        this.Site.CaptureRequirementIfAreEqual<uint>(
                            0x0000001E,
                            propertyValue[j].PropTag & 0x0000FFFF,
                            1967,
                            @"[In Conversion Rules for String Values Specified by the Server to the Client] In NspiGetSpecialTable method, String values can be returned in 8-bit character representation in the output parameter ppRows.");

                        break;
                    }
                }

                if (!isString8ValueFound)
                {
                    Site.Assert.Fail("Property PidTagDisplayName with property type PtypString8 is not found.");
                }

                // Check the property sequence.
                Site.Assert.AreEqual<uint>((uint)AulProp.PidTagEntryId, propertyValue[0].PropTag, "The first property should be PidTagEntryId.");
                Site.Assert.AreEqual<uint>((uint)AulProp.PidTagContainerFlags, propertyValue[1].PropTag, "The second property should be PidTagContainerFlags.");
                Site.Assert.AreEqual<uint>((uint)AulProp.PidTagDepth, propertyValue[2].PropTag, "The third property should be PidTagDepth.");
                Site.Assert.AreEqual<uint>((uint)AulProp.PidTagAddressBookContainerId, propertyValue[3].PropTag, "The fourth property should be PidTagAddressBookContainerId.");
                Site.Assert.AreEqual<uint>((uint)AulProp.PidTagDisplayName, propertyValue[4].PropTag, "The fifth property should be PidTagDisplayName.");
                Site.Assert.AreEqual<uint>((uint)AulProp.PidTagAddressBookIsMaster, propertyValue[5].PropTag, "The sixth property should be PidTagAddressBookIsMaster.");

                // Check the property must have a value under its definition.
                Site.Assert.IsNotNull(Convert.ToString(propertyValue[0].Value), "PidTagEntryId must have a value if it exists.");
                Site.Assert.IsNotNull(Convert.ToString(propertyValue[1].Value), "PidTagContainerFlags must have a value if it exists.");
                Site.Assert.IsNotNull(Convert.ToString(propertyValue[2].Value), "PidTagDepth must have a value if it exists.");
                Site.Assert.IsNotNull(Convert.ToString(propertyValue[3].Value), "PidTagAddressBookContainerId must have a value if it exists.");
                Site.Assert.IsNotNull(Convert.ToString(propertyValue[4].Value), "PidTagDisplayName must have a value if it exists.");
                Site.Assert.IsNotNull(Convert.ToString(propertyValue[5].Value), "PidTagAddressBookIsMaster must have a value if it exists.");

                if (propertyValue.Length == 7)
                {
                    Site.Assert.AreEqual<uint>((uint)AulProp.PidTagAddressBookParentEntryId, propertyValue[6].PropTag, "PidTagAddressBookParentEntryId must be the seventh property if it exists.");
                    Site.Assert.IsNotNull(Convert.ToString(propertyValue[6].Value), "PidTagAddressBookParentEntryId must have a value if it exists.");
                }

                // The property sequence is ensured by the above asserts, so MS-OXNSPI_R1863 can be captured directly.
                this.Site.CaptureRequirement(
                    1863,
                    @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 14] The server MUST return the following properties for each container in the hierarchy, in the order listed: PidTagEntryId ([MS-OXPROPS] section 2.674)
                    PidTagContainerFlags ([MS-OXPROPS] section 2.635)
                    PidTagDepth ([MS-OXPROPS] section 2.664)
                    PidTagAddressBookContainerId ([MS-OXPROPS] section 2.503)
                    PidTagDisplayName ([MS-OXPROPS] section 2.667)
                    PidTagAddressBookIsMaster ([MS-OXPROPS] section 2.536)
                    PidTagAddressBookParentEntryId ([MS-OXPROPS] section 2.550) (optional, and MUST be the seventh column if it [property PidTagAddressBookParentEntryId] is included)");

                // The property must have a value according to the above asserts, so MS-OXNSPI_R1864 can be captured directly.
                this.Site.CaptureRequirement(
                    1864,
                    @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 15] For every row returned, all of these properties [PidTagEntryId, PidTagContainerFlags, PidTagDepth, PidTagAddressBookContainerId, PidTagDisplayName, PidTagAddressBookIsMaster and PidTagAddressBookParentEntryId] except PidTagAddressBookParentEntryId MUST be present, and each of them MUST have a value prescribed under its definition.");

                PermanentEntryID permanentEntryID = AdapterHelper.ParsePermanentEntryIDFromBytes(propertyValue[0].Value.Bin.Lpb);

                bool isHasContainer = false;

                // Check whether the display type is DT_CONTAINER.
                // The PidTagEntryId property is in the form of a PermanentEntryID structure. According to the definition of PermanentEntryID structure, the display type is defined from the 25th byte. 
                DisplayTypeValue displayType = (DisplayTypeValue)BitConverter.ToInt32(propertyValue[0].Value.Bin.Lpb, 24);
                if (displayType == DisplayTypeValue.DT_CONTAINER)
                {
                    isHasContainer = true;
                }

                // Check whether the DN follows the addresslist-dn format.
                bool isAddresslistDNFormat = Common.IsDNMatchABNF(permanentEntryID.DistinguishedName, DNFormat.AddressListDn);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1865. The ID of the PermanentEntryID is {0}, the display type of PidTagEntryId is {1}, the DN is {2}.", permanentEntryID.IDType, displayType, permanentEntryID.DistinguishedName);

                // Verify MS-OXNSPI requirement: MS-OXNSPI_R1865
                this.Site.CaptureRequirementIfIsTrue(
                    permanentEntryID.ToString() != null && isHasContainer && isAddresslistDNFormat,
                    1865,
                    @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 16] The PidTagEntryId property MUST be in the form of a PermanentEntryID structure, as section 2.2.9.3, with its PidTagDisplayType property having the value DT_CONTAINER, as specified in section 2.2.1.3, and its DN following the addresslist-dn format specification, as specified in [MS-OXOABK] section 2.2.1.1.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXOABK_R29.");

                // Verify MS-OXOABK requirement: MS-OXOABK_R29
                this.Site.CaptureRequirementIfIsTrue(
                    isAddresslistDNFormat,
                    "MS-OXOABK",
                    29,
                    @"[In Distinguished Names for Objects] [DNs(1) for specific objects have a strict format, as shown in the following table] When the object type is Address book container, the dn formats is addresslist-dn.");

                // According to MS-OXOABK, that the value of PidTagAddressBookContainerId is 0 (zero) represents the Global Address List (GAL).
                if (propertyValue[3].Value.L == 0)
                {
                    permanentEntryID = AdapterHelper.ParsePermanentEntryIDFromBytes(propertyValue[0].Value.Bin.Lpb);

                    // Check whether the DN follows the gal-addrlist-dn format.
                    isAddresslistDNFormat = Common.IsDNMatchABNF(permanentEntryID.DistinguishedName, DNFormat.GalAddrlistDn);

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1866. The distinguished name is {0}.", permanentEntryID.DistinguishedName);

                    // Verify MS-OXNSPI requirement: MS-OXNSPI_R1866
                    this.Site.CaptureRequirementIfIsTrue(
                        isAddresslistDNFormat,
                        1866,
                        @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 16] When the object is the Global Address List (GAL) container, its DN MUST follow the gal-addrlist-dn format specification.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXOABK_R30.");

                    // Verify MS-OXOABK requirement: MS-OXOABK_R30
                    this.Site.CaptureRequirementIfIsTrue(
                        isAddresslistDNFormat,
                        "MS-OXOABK",
                        30,
                        @"[In Distinguished Names for Objects] [DNs(1) for specific objects have a strict format, as shown in the following table] When the object type is Global Address List container, the dn formats is gal-addrlist-dn.");
                }
            }

            // The handle is used as the input parameter by all other NSPI methods. So MS-OXNSPI_R413 can be verified if all cases are passed.
            this.Site.CaptureRequirement(
                413,
                @"[In NSPI_HANDLE] The NSPI_HANDLE handle is an RPC context handle that is used to share a session between method calls.");

            // The handle is used as the input parameter by all other NSPI methods. So MS-OXNSPI_R675 can be verified if all cases are passed.
            this.Site.CaptureRequirement(
                675,
                @"[In NspiBind] contextHandle: An RPC context handle, as specified in section 2.2.10.");

            // The handle is used as the input parameter by all other NSPI methods. So MS-OXNSPI_R668 can be verified if all cases are passed.
            this.Site.CaptureRequirement(
                668,
                @"[In NspiBind] The NspiBind method initiates a session between a client and the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R1740", "The number of rows in the first call is {0}, the number of rows in the second call is {1}.", rows.Value.CRows, rows1.Value.CRows);

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1740
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoPropertyRowSetEqual(rows, rows1),
                1740,
                @"[In NspiGetSpecialTable] If the bit flags [dwFlags] are set to different values other than NspiAddressCreationTemplates and NspiUnicodeStrings, the server will return the same value.");

            #endregion
            #endregion

            #region Call NspiGetSpecialTable method with dwFlags containing both "NspiAddressCreationTemplates" and "NspiUnicodeStrings".
            PropertyRowSet_r? rows2;
            flagsOfGetSpecialTable = (uint)(NspiGetSpecialTableFlags.NspiUnicodeStrings | NspiGetSpecialTableFlags.NspiAddressCreationTemplates);
            this.Result = this.ProtocolAdatper.NspiGetSpecialTable(flagsOfGetSpecialTable, ref stat, ref version, out rows2);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, this.Result, "NspiGetSpecialTable method should return Success.");
            Site.Assert.IsNotNull(rows2.Value.ARow, "The returned rows should not be empty. The row number is {0}.", rows2 == null ? 0 : rows2.Value.CRows);

            #region Capture code
            // Check whether server returns the display type DT_ADDRESS_TEMPLATE as part of the EntryID of an object.
            addressTemplateReturnedInPidTagEntryId = AdapterHelper.CheckIfSpecificDisplayTypeExists(rows2.Value.ARow, DisplayTypeValue.DT_ADDRESS_TEMPLATE);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1996",
                "Display type DT_ADDRESS_TEMPLATE {0} returned as part of an EntryID of an object in the Address Creation Table.",
                addressTemplateReturnedInPidTagEntryId ? "is" : "is not");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1996
            Site.CaptureRequirementIfIsTrue(
                addressTemplateReturnedInPidTagEntryId,
                1996,
                @"[In Display Type Values] Exchange NSPI server will return display type DT_ADDRESS_TEMPLATE as part of an EntryID of an object in the Address Creation Table.");

            // Check whether server returns the display type DT_CONTAINER as part of an EntryID of an object.
            containerReturnedInPidTagEntryId = AdapterHelper.CheckIfSpecificDisplayTypeExists(rows2.Value.ARow, DisplayTypeValue.DT_CONTAINER);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXNSPI_R1999",
                "Display type DT_CONTAINER {0} returned if the table is not an address book hierarchy table.",
                containerReturnedInPidTagEntryId ? "is" : "is not");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R1999
            Site.CaptureRequirementIfIsFalse(
                containerReturnedInPidTagEntryId,
                1999,
                @"[In Display Type Values] Exchange NSPI server will not return display type DT_CONTAINER as part of an EntryID of an object if the table is not the address book hierarchy table.");

            PropertyValue_r? property = null;

            // PidTagInstanceKey exists in the address creation templates table. 
            // So the table is an address creation table if PidTagInstanceKey is found successfully.
            bool isAddressCreationTemplatesTable = AdapterHelper.FindFirstSpecifiedPropTagValueInRowSet(rows2.Value, (uint)AulProp.PidTagInstanceKey, out property);
            bool isUnicodeStringsReturned = AdapterHelper.IsPtypString(rows2.Value);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R741", "The specified property {0} found.", isAddressCreationTemplatesTable ? "is" : "is not");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R741
            // The flagsOfGetSpecialTable contains both "NspiAddressCreationTemplates" and "NspiUnicodeStrings" 
            // and the server returns an address creation table. So the server ignores the "NspiUnicodeStrings".
            Site.CaptureRequirementIfIsTrue(
                isAddressCreationTemplatesTable && !isUnicodeStringsReturned,
                741,
                @"[In NspiGetSpecialTable] [Server Processing Rules: Upon receiving message NspiGetSpecialTable, the server MUST process the data from the message subject to the following constraints:] [Constraint 4] If the input parameter dwFlags contains both the value ""NspiAddressCreationTemplates"" and the value ""NspiUnicodeStrings"", the server MUST ignore the value ""NspiUnicodeStrings"" and proceed as if the parameter dwFlags contained only the value ""NspiAddressCreationTemplates"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXNSPI_R110", "The specified property {0} found.", isAddressCreationTemplatesTable ? "is" : "is not");

            // Verify MS-OXNSPI requirement: MS-OXNSPI_R110
            // The flagsOfGetSpecialTable contains both "NspiAddressCreationTemplates" and "NspiUnicodeStrings" 
            // and the server returns an address creation table. So the server ignores the "NspiUnicodeStrings".
            Site.CaptureRequirementIfIsTrue(
                isAddressCreationTemplatesTable && !isUnicodeStringsReturned,
                110,
                @"[In NspiGetSpecialTable Flags] NspiAddressCreationTemplates (0x00000002): Specifying this flag causes the server to ignore the NspiUnicodeStrings flag.");

            #endregion
            #endregion

            #region Call NspiUnbind to destroy the context handle.
            uint reserved = 0;
            uint returnValue = this.ProtocolAdatper.NspiUnbind(reserved);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            #endregion
        }
    }
}